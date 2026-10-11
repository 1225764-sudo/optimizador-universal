import math
from datetime import datetime, time

import pandas as pd
from ortools.sat.python import cp_model


TRUE_VALUES = {"si", "sí", "yes", "true", "1"}


def _is_true(value):
    return str(value).strip().lower() in TRUE_VALUES


def _clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def _time_to_minutes(value):
    """Convierte horas de Excel, datetime.time o texto HH:MM[:SS] a minutos."""
    if pd.isna(value):
        raise ValueError("hora vacía")
    if isinstance(value, time):
        return value.hour * 60 + value.minute
    if isinstance(value, datetime):
        return value.hour * 60 + value.minute
    # Excel puede entregar fracción de día.
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if 0 <= float(value) < 1:
            return int(round(float(value) * 24 * 60))
    text = str(value).strip()
    if " " in text and ":" in text:  # p.ej. 1900-01-01 08:00:00
        text = text.split()[-1]
    parts = text.split(":")
    if len(parts) < 2:
        raise ValueError(f"formato de hora no reconocido: {value}")
    return int(parts[0]) * 60 + int(parts[1])


def _minutes_to_hhmm(minutes):
    minutes = int(minutes)
    return f"{minutes // 60:02d}:{minutes % 60:02d}"


def _config_int(config, key, default):
    value = config.get(key, default)
    if pd.isna(value) or str(value).strip() == "":
        return int(default)
    return int(float(value))


def optimizar_desde_excel(archivo_excel):
    """
    Genera un horario factible y, entre las soluciones factibles, optimiza de forma
    jerárquica: huecos internos -> permanencia/compactación -> horario tardío
    configurable -> equilibrio de carga docente.

    Mantiene la interfaz de retorno de la versión anterior y añade métricas/estado_solver.
    """

    # ============================================================
    # 1. LEER ARCHIVO EXCEL
    # ============================================================
    try:
        df_grupos = pd.read_excel(archivo_excel, sheet_name="Grupos")
        df_personal = pd.read_excel(archivo_excel, sheet_name="Personal")
        df_espacios = pd.read_excel(archivo_excel, sheet_name="Espacios")
        df_horarios = pd.read_excel(archivo_excel, sheet_name="Horarios")
        df_disponibilidad = pd.read_excel(archivo_excel, sheet_name="Disponibilidad")
        df_configuracion = pd.read_excel(archivo_excel, sheet_name="Configuracion")
    except Exception as error:
        return {"estado": "ERROR", "mensaje": f"No fue posible leer el Excel: {error}"}

    # ============================================================
    # 2. LIMPIEZA Y VALIDACIÓN BÁSICA
    # ============================================================
    for df_temp in [df_grupos, df_personal, df_espacios, df_horarios,
                    df_disponibilidad, df_configuracion]:
        df_temp.columns = df_temp.columns.astype(str).str.strip()

    required = {
        "Grupos": {"Grupo", "Materia", "Personas", "Sesiones_semana", "Tipo_recurso"},
        "Personal": {"Persona", "Materias_habilitadas", "Max_clases_semana", "Max_clases_dia"},
        "Espacios": {"Espacio", "Capacidad", "Tipo"},
        "Horarios": {"Dia", "Hora_inicio", "Hora_fin", "Disponible"},
        "Disponibilidad": {"Persona", "Dia", "Hora_inicio", "Hora_fin", "Disponible"},
        "Configuracion": {"Parametro", "Valor"},
    }
    frames = {
        "Grupos": df_grupos, "Personal": df_personal, "Espacios": df_espacios,
        "Horarios": df_horarios, "Disponibilidad": df_disponibilidad,
        "Configuracion": df_configuracion,
    }
    for sheet, cols in required.items():
        missing = cols - set(frames[sheet].columns)
        if missing:
            return {"estado": "ERROR", "mensaje": f"Faltan columnas en '{sheet}': {', '.join(sorted(missing))}."}

    # ============================================================
    # 3. CONFIGURACIÓN
    # ============================================================
    config = {
        _clean_text(k): v
        for k, v in zip(df_configuracion["Parametro"], df_configuracion["Valor"])
        if _clean_text(k)
    }
    mantener_mismo_profesor = _is_true(config.get("Mantener_mismo_profesor", "Si"))
    permitir_sobrecupo = _is_true(config.get("Permitir_sobrecupo", "No"))
    max_clases_dia_global = _config_int(config, "Max_clases_dia", 10**6) if "Max_clases_dia" in config else 10**6

    # Pesos: 0 desactiva ese criterio. La jerarquía se garantiza dinámicamente.
    peso_huecos = max(0, _config_int(config, "Peso_huecos", 1))
    peso_permanencia = max(0, _config_int(config, "Peso_permanencia", 1))
    peso_tardio = max(0, _config_int(config, "Peso_horario_tardio", 1))
    peso_equilibrio = max(0, _config_int(config, "Peso_equilibrio_docente", 1))

    # Compatibilidad con nombres usados en plantillas anteriores del proyecto.
    if "Penalizar_huecos" in config and not _is_true(config["Penalizar_huecos"]):
        peso_huecos = 0
    if "Compactar_jornada" in config and not _is_true(config["Compactar_jornada"]):
        peso_permanencia = 0
    if "Penalizar_salida_tardia" in config and not _is_true(config["Penalizar_salida_tardia"]):
        peso_tardio = 0

    hora_preferida_raw = config.get("Hora_preferida_salida", config.get("Hora_limite_salida", None))
    hora_preferida_salida = None
    if hora_preferida_raw is not None and not pd.isna(hora_preferida_raw) and str(hora_preferida_raw).strip():
        try:
            hora_preferida_salida = _time_to_minutes(hora_preferida_raw)
        except Exception as exc:
            return {"estado": "ERROR", "mensaje": f"Hora_preferida_salida inválida: {exc}"}

    # ============================================================
    # 4. DATOS PRINCIPALES
    # ============================================================
    grupos = df_grupos["Grupo"].map(_clean_text).drop_duplicates().tolist()
    profesores = df_personal["Persona"].map(_clean_text).drop_duplicates().tolist()
    espacios = df_espacios["Espacio"].map(_clean_text).drop_duplicates().tolist()

    if not grupos or not profesores or not espacios:
        return {"estado": "ERROR", "mensaje": "Grupos, Personal y Espacios deben contener al menos un registro."}

    # ============================================================
    # 5. EVENTOS (cada sesión requerida se vuelve un evento obligatorio)
    # ============================================================
    eventos = []
    try:
        for _, fila in df_grupos.iterrows():
            grupo = _clean_text(fila["Grupo"])
            materia = _clean_text(fila["Materia"])
            personas = int(fila["Personas"])
            sesiones = int(fila["Sesiones_semana"])
            tipo_recurso = _clean_text(fila["Tipo_recurso"])
            if not grupo or not materia or sesiones <= 0 or personas < 0:
                raise ValueError(f"registro inválido en Grupos: {grupo} - {materia}")
            for sesion in range(1, sesiones + 1):
                eventos.append({"grupo": grupo, "materia": materia, "sesion": sesion,
                                "personas": personas, "tipo": tipo_recurso})
    except Exception as exc:
        return {"estado": "ERROR", "mensaje": f"Datos inválidos en 'Grupos': {exc}"}

    # ============================================================
    # 6. PERIODOS INSTITUCIONALES DISPONIBLES
    # ============================================================
    slots = []
    try:
        for idx, fila in df_horarios.iterrows():
            if not _is_true(fila["Disponible"]):
                continue
            dia = _clean_text(fila["Dia"])
            ini = _time_to_minutes(fila["Hora_inicio"])
            fin = _time_to_minutes(fila["Hora_fin"])
            if not dia or fin <= ini:
                raise ValueError(f"periodo inválido en fila {idx + 2}")
            slots.append({"id": len(slots), "dia": dia, "inicio": ini, "fin": fin,
                          "label": f"{dia} {_minutes_to_hhmm(ini)}-{_minutes_to_hhmm(fin)}"})
    except Exception as exc:
        return {"estado": "ERROR", "mensaje": f"Datos inválidos en 'Horarios': {exc}"}

    if not slots:
        return {"estado": "ERROR", "mensaje": "No existen periodos institucionales marcados como disponibles."}

    # Evitar ambigüedad: un mismo día no debe contener periodos que se traslapen.
    slots_by_day = {}
    for s in slots:
        slots_by_day.setdefault(s["dia"], []).append(s)
    for dia, day_slots in slots_by_day.items():
        day_slots.sort(key=lambda z: (z["inicio"], z["fin"]))
        for a, b in zip(day_slots, day_slots[1:]):
            if b["inicio"] < a["fin"]:
                return {"estado": "ERROR", "mensaje": f"Hay periodos traslapados en Horarios para {dia}: {a['label']} y {b['label']}."}

    # ============================================================
    # 7. ESPACIOS
    # ============================================================
    info_espacios = {}
    try:
        for _, fila in df_espacios.iterrows():
            nombre = _clean_text(fila["Espacio"])
            info_espacios[nombre] = {
                "capacidad": int(fila["Capacidad"]),
                "tipo": _clean_text(fila["Tipo"]),
                # Se conservan como información. Solo se vuelven restricciones si
                # Grupos incluye requerimientos explícitos equivalentes.
                "computadoras": int(fila["Computadoras"]) if "Computadoras" in df_espacios.columns and not pd.isna(fila.get("Computadoras")) else 0,
                "proyector": _is_true(fila.get("Proyector", "No")),
            }
    except Exception as exc:
        return {"estado": "ERROR", "mensaje": f"Datos inválidos en 'Espacios': {exc}"}

    # Requerimientos opcionales, retrocompatibles.
    req_computadoras_col = "Computadoras_requeridas" if "Computadoras_requeridas" in df_grupos.columns else None
    req_proyector_col = "Requiere_proyector" if "Requiere_proyector" in df_grupos.columns else None
    if req_computadoras_col or req_proyector_col:
        # Propagar requisitos a los eventos en el mismo orden de df_grupos/sesiones.
        k = 0
        for _, fila in df_grupos.iterrows():
            sesiones = int(fila["Sesiones_semana"])
            req_pc = int(fila[req_computadoras_col]) if req_computadoras_col and not pd.isna(fila[req_computadoras_col]) else 0
            req_proj = _is_true(fila[req_proyector_col]) if req_proyector_col else False
            for _ in range(sesiones):
                eventos[k]["computadoras_requeridas"] = req_pc
                eventos[k]["requiere_proyector"] = req_proj
                k += 1

    def espacio_compatible(evento, espacio):
        info = info_espacios[espacio]
        if not permitir_sobrecupo and info["capacidad"] < evento["personas"]:
            return False
        # Tipo vacío o "Cualquiera" permite cualquier espacio; si se especifica, exige coincidencia.
        tipo_req = evento["tipo"].strip().lower()
        tipo_esp = info["tipo"].strip().lower()
        if tipo_req and tipo_req not in {"cualquiera", "any", "todos"} and tipo_req != tipo_esp:
            return False
        if evento.get("computadoras_requeridas", 0) > info["computadoras"]:
            return False
        if evento.get("requiere_proyector", False) and not info["proyector"]:
            return False
        return True

    # ============================================================
    # 8. PERSONAL Y DISPONIBILIDAD
    # ============================================================
    info_personal = {}
    try:
        for _, fila in df_personal.iterrows():
            persona = _clean_text(fila["Persona"])
            materias = [_clean_text(x) for x in str(fila["Materias_habilitadas"]).split(",") if _clean_text(x)]
            info_personal[persona] = {
                "materias": materias,
                "max_semana": int(fila["Max_clases_semana"]),
                "max_dia": int(fila["Max_clases_dia"]),
            }
    except Exception as exc:
        return {"estado": "ERROR", "mensaje": f"Datos inválidos en 'Personal': {exc}"}

    prof_ranges = {}
    try:
        for _, fila in df_disponibilidad.iterrows():
            if not _is_true(fila["Disponible"]):
                continue
            persona = _clean_text(fila["Persona"])
            dia = _clean_text(fila["Dia"])
            ini = _time_to_minutes(fila["Hora_inicio"])
            fin = _time_to_minutes(fila["Hora_fin"])
            if fin <= ini:
                raise ValueError(f"rango inválido para {persona} en {dia}")
            prof_ranges.setdefault(persona, {}).setdefault(dia, []).append((ini, fin))
    except Exception as exc:
        return {"estado": "ERROR", "mensaje": f"Datos inválidos en 'Disponibilidad': {exc}"}

    disponibilidad_personal = {}
    for profesor in profesores:
        for s in slots:
            ranges = prof_ranges.get(profesor, {}).get(s["dia"], [])
            disponibilidad_personal[(profesor, s["id"])] = any(
                s["inicio"] >= ini and s["fin"] <= fin for ini, fin in ranges
            )

    # ============================================================
    # 9. MODELO Y VARIABLES DE DECISIÓN
    # ============================================================
    model = cp_model.CpModel()
    x = {}
    for i, evento in enumerate(eventos):
        for profesor in profesores:
            if evento["materia"] not in info_personal[profesor]["materias"]:
                continue
            for espacio in espacios:
                if not espacio_compatible(evento, espacio):
                    continue
                for s in slots:
                    if not disponibilidad_personal.get((profesor, s["id"]), False):
                        continue
                    x[(i, profesor, espacio, s["id"])] = model.NewBoolVar(
                        f"x_{i}_{profesor}_{espacio}_{s['id']}"
                    )

    # ============================================================
    # 10. RESTRICCIONES DURAS
    # ============================================================
    # Cada evento exactamente una vez.
    for i, evento in enumerate(eventos):
        vars_i = [v for (eid, _, _, _), v in x.items() if eid == i]
        if not vars_i:
            return {
                "estado": "ERROR",
                "mensaje": (f"El evento {i + 1}: {evento['grupo']} - {evento['materia']} "
                            "no tiene ninguna combinación profesor/espacio/horario compatible.")
            }
        model.Add(sum(vars_i) == 1)

    # No colisiones de grupo, espacio ni profesor.
    for s in slots:
        sid = s["id"]
        for grupo in grupos:
            vv = [v for (eid, _, _, h), v in x.items() if h == sid and eventos[eid]["grupo"] == grupo]
            if vv:
                model.Add(sum(vv) <= 1)
        for espacio in espacios:
            vv = [v for (_, _, e, h), v in x.items() if h == sid and e == espacio]
            if vv:
                model.Add(sum(vv) <= 1)
        for profesor in profesores:
            vv = [v for (_, p, _, h), v in x.items() if h == sid and p == profesor]
            if vv:
                model.Add(sum(vv) <= 1)

    # Carga semanal y diaria docente.
    for profesor in profesores:
        vv = [v for (_, p, _, _), v in x.items() if p == profesor]
        if vv:
            model.Add(sum(vv) <= info_personal[profesor]["max_semana"])
        for dia, day_slots in slots_by_day.items():
            ids = {s["id"] for s in day_slots}
            vd = [v for (_, p, _, h), v in x.items() if p == profesor and h in ids]
            if vd:
                limite_diario = min(info_personal[profesor]["max_dia"], max_clases_dia_global)
                model.Add(sum(vd) <= limite_diario)

    # Mantener el mismo profesor para todas las sesiones de grupo-materia.
    if mantener_mismo_profesor:
        pares = {(e["grupo"], e["materia"]) for e in eventos}
        for grupo, materia in pares:
            ids_eventos = [i for i, e in enumerate(eventos) if e["grupo"] == grupo and e["materia"] == materia]
            if len(ids_eventos) <= 1:
                continue
            for profesor in profesores:
                sums = []
                for i in ids_eventos:
                    vv = [v for (eid, p, _, _), v in x.items() if eid == i and p == profesor]
                    sums.append(sum(vv) if vv else 0)
                for other in sums[1:]:
                    model.Add(sums[0] == other)

    # Romper simetría segura entre sesiones intercambiables de una misma
    # combinación grupo-materia. Las etiquetas Sesion 1, 2, ... no tienen
    # significado académico distinto, por lo que podemos imponer que sus
    # periodos aparezcan en orden cronológico sin eliminar horarios reales.
    # Esto reduce soluciones equivalentes y deja más tiempo al solver para
    # mejorar la calidad temporal.
    slot_rank = {s["id"]: rank for rank, s in enumerate(sorted(slots, key=lambda z: (list(slots_by_day).index(z["dia"]), z["inicio"], z["fin"]))) }
    for grupo, materia in {(e["grupo"], e["materia"]) for e in eventos}:
        ids_eventos = sorted(
            [i for i, e in enumerate(eventos) if e["grupo"] == grupo and e["materia"] == materia],
            key=lambda i: eventos[i]["sesion"],
        )
        slot_exprs = []
        for i in ids_eventos:
            terms = [slot_rank[h] * v for (eid, _, _, h), v in x.items() if eid == i]
            if terms:
                slot_exprs.append(sum(terms))
        for a, b in zip(slot_exprs, slot_exprs[1:]):
            model.Add(a < b)

    # ============================================================
    # 11. VARIABLES DE PRESENCIA DE GRUPO POR PERIODO
    # ============================================================
    y = {}
    for grupo in grupos:
        for s in slots:
            sid = s["id"]
            vv = [v for (eid, _, _, h), v in x.items() if h == sid and eventos[eid]["grupo"] == grupo]
            y[(grupo, sid)] = model.NewBoolVar(f"pres_{grupo}_{sid}")
            # Crear presencia incluso cuando no existe combinación posible en ese periodo:
            # así un periodo libre entre dos clases sí cuenta como hueco.
            if vv:
                model.Add(y[(grupo, sid)] == sum(vv))
            else:
                model.Add(y[(grupo, sid)] == 0)

    # ============================================================
    # 12. OBJETIVOS TEMPORALES
    # ============================================================
    gap_vars = []
    span_vars = []
    span_upper_bounds = []
    late_terms = []

    for grupo in grupos:
        for dia, day_slots in slots_by_day.items():
            # Solo periodos donde el grupo tiene al menos una asignación posible.
            pvars = [(s, y[(grupo, s["id"])]) for s in day_slots]
            if not pvars:
                continue

            # has_before/has_after permiten identificar huecos estrictamente internos.
            for j, (s, cur) in enumerate(pvars):
                before = [v for _, v in pvars[:j]]
                after = [v for _, v in pvars[j + 1:]]
                if not before or not after:
                    continue
                has_before = model.NewBoolVar(f"before_{grupo}_{s['id']}")
                has_after = model.NewBoolVar(f"after_{grupo}_{s['id']}")
                model.AddMaxEquality(has_before, before)
                model.AddMaxEquality(has_after, after)
                gap = model.NewBoolVar(f"gap_{grupo}_{s['id']}")
                # gap = has_before AND has_after AND NOT cur
                model.Add(gap <= has_before)
                model.Add(gap <= has_after)
                model.Add(gap + cur <= 1)
                model.Add(gap >= has_before + has_after - cur - 1)
                gap_vars.append(gap)

            # Permanencia: última salida - primera entrada. Si no hay clase ese día, 0.
            day_presence = [v for _, v in pvars]
            active = model.NewBoolVar(f"active_{grupo}_{dia}")
            model.AddMaxEquality(active, day_presence)

            starts = []
            ends = []
            for s, v in pvars:
                st = model.NewIntVar(0, s["inicio"], f"st_{grupo}_{s['id']}")
                en = model.NewIntVar(0, s["fin"], f"en_{grupo}_{s['id']}")
                model.Add(st == s["inicio"]).OnlyEnforceIf(v)
                model.Add(st == 0).OnlyEnforceIf(v.Not())
                model.Add(en == s["fin"]).OnlyEnforceIf(v)
                model.Add(en == 0).OnlyEnforceIf(v.Not())
                starts.append((s, v))
                ends.append(en)

            # Primera hora mediante selección booleana first; última mediante last.
            first_flags, last_flags = [], []
            for j, (s, v) in enumerate(pvars):
                f = model.NewBoolVar(f"first_{grupo}_{s['id']}")
                l = model.NewBoolVar(f"last_{grupo}_{s['id']}")
                # first iff v y no presencia anterior
                prev = [pv for _, pv in pvars[:j]]
                nxt = [pv for _, pv in pvars[j + 1:]]
                model.Add(f <= v)
                for pv in prev:
                    model.Add(f + pv <= 1)
                if prev:
                    model.Add(f >= v - sum(prev))
                else:
                    model.Add(f == v)
                model.Add(l <= v)
                for pv in nxt:
                    model.Add(l + pv <= 1)
                if nxt:
                    model.Add(l >= v - sum(nxt))
                else:
                    model.Add(l == v)
                first_flags.append((s, f))
                last_flags.append((s, l))

            model.Add(sum(f for _, f in first_flags) == active)
            model.Add(sum(l for _, l in last_flags) == active)
            max_span = max(s["fin"] for s, _ in pvars) - min(s["inicio"] for s, _ in pvars)
            span = model.NewIntVar(0, max(0, max_span), f"span_{grupo}_{dia}")
            first_time = sum(s["inicio"] * f for s, f in first_flags)
            last_time = sum(s["fin"] * l for s, l in last_flags)
            model.Add(span == last_time - first_time)
            span_vars.append(span)
            span_upper_bounds.append(max(0, max_span))

    # Penalización tardía progresiva: minutos que el periodo termina después del umbral.
    if hora_preferida_salida is not None:
        for grupo in grupos:
            for s in slots:
                v = y.get((grupo, s["id"]))
                if v is None:
                    continue
                exceso = max(0, s["fin"] - hora_preferida_salida)
                if exceso:
                    late_terms.append(exceso * v)

    # Equilibrio docente: diferencia entre carga máxima y mínima entre docentes activos.
    balance_var = None
    active_profesores = []
    load_vars = []
    for profesor in profesores:
        vv = [v for (_, p, _, _), v in x.items() if p == profesor]
        if not vv:
            continue
        active_profesores.append(profesor)
        max_load = min(info_personal[profesor]["max_semana"], len(eventos))
        load = model.NewIntVar(0, max_load, f"load_{profesor}")
        model.Add(load == sum(vv))
        load_vars.append(load)
    if len(load_vars) >= 2:
        max_load_var = model.NewIntVar(0, len(eventos), "max_teacher_load")
        min_load_var = model.NewIntVar(0, len(eventos), "min_teacher_load")
        model.AddMaxEquality(max_load_var, load_vars)
        model.AddMinEquality(min_load_var, load_vars)
        balance_var = model.NewIntVar(0, len(eventos), "teacher_imbalance")
        model.Add(balance_var == max_load_var - min_load_var)

    # ============================================================
    # 13. FUNCIÓN OBJETIVO JERÁRQUICA EN UNA SOLA RESOLUCIÓN
    # ============================================================
    gap_expr = sum(gap_vars) if gap_vars else 0
    span_expr = sum(span_vars) if span_vars else 0
    late_expr = sum(late_terms) if late_terms else 0
    balance_expr = balance_var if balance_var is not None else 0

    # Cotas superiores seguras para construir coeficientes de dominancia.
    max_gap = len(gap_vars) * max(1, peso_huecos)
    max_span = sum(span_upper_bounds) * max(1, peso_permanencia) if span_vars else 0
    max_late = sum(max(0, s["fin"] - (hora_preferida_salida or s["fin"])) for s in slots) * len(grupos) * max(1, peso_tardio) if hora_preferida_salida is not None else 0
    max_balance = len(eventos) * max(1, peso_equilibrio)

    # De abajo hacia arriba: cada nivel superior domina toda combinación de niveles inferiores.
    c_balance = 1
    c_late = max_balance + 1
    c_span = max_late * c_late + max_balance + 1
    c_gap = max_span * c_span + max_late * c_late + max_balance + 1

    objective_terms = []
    if peso_huecos > 0 and gap_vars:
        objective_terms.append(c_gap * peso_huecos * gap_expr)
    if peso_permanencia > 0 and span_vars:
        objective_terms.append(c_span * peso_permanencia * span_expr)
    if peso_tardio > 0 and late_terms:
        objective_terms.append(c_late * peso_tardio * late_expr)
    if peso_equilibrio > 0 and balance_var is not None:
        objective_terms.append(c_balance * peso_equilibrio * balance_expr)
    # CP-SAT utiliza enteros de 64 bits. En instituciones muy grandes, una
    # codificación jerárquica por coeficientes puede crecer demasiado.
    # Rechazamos de forma explícita una escala insegura en lugar de permitir
    # overflow silencioso o alterar prioridades.
    INT64_SAFE_OBJECTIVE = 8_000_000_000_000_000_000
    objective_upper_bound = 0
    if peso_huecos > 0 and gap_vars:
        objective_upper_bound += c_gap * peso_huecos * len(gap_vars)
    if peso_permanencia > 0 and span_vars:
        objective_upper_bound += c_span * peso_permanencia * sum(span_upper_bounds)
    if peso_tardio > 0 and late_terms:
        raw_late_bound = sum(max(0, s["fin"] - (hora_preferida_salida or s["fin"])) for s in slots) * len(grupos)
        objective_upper_bound += c_late * peso_tardio * raw_late_bound
    if peso_equilibrio > 0 and balance_var is not None:
        objective_upper_bound += c_balance * peso_equilibrio * len(eventos)
    if objective_upper_bound > INT64_SAFE_OBJECTIVE:
        return {
            "estado": "ERROR",
            "mensaje": ("La escala de la función objetivo supera un rango entero seguro para CP-SAT. "
                        "El escenario es demasiado grande para la codificación jerárquica actual; "
                        "reduzca pesos o utilice una estrategia lexicográfica por fases."),
        }
    if objective_terms:
        model.Minimize(sum(objective_terms))

    # ============================================================
    # 14. RESOLVER
    # ============================================================
    solver = cp_model.CpSolver()
    # La optimización temporal es considerablemente más difícil que hallar
    # cualquier solución factible. Los valores siguen siendo configurables.
    tiempo_solver = float(config.get("Tiempo_max_solver_segundos", config.get("Tiempo_max_solver", 60)) or 60)
    workers_solver = max(1, _config_int(config, "Num_search_workers", 8))
    semilla_solver = _config_int(config, "Semilla_solver", 42)
    solver.parameters.max_time_in_seconds = max(1.0, tiempo_solver)
    solver.parameters.num_search_workers = workers_solver
    solver.parameters.random_seed = semilla_solver
    status = solver.Solve(model)

    status_names = {
        cp_model.OPTIMAL: "OPTIMAL",
        cp_model.FEASIBLE: "FEASIBLE",
        cp_model.INFEASIBLE: "INFEASIBLE",
        cp_model.MODEL_INVALID: "MODEL_INVALID",
        cp_model.UNKNOWN: "UNKNOWN",
    }
    status_name = status_names.get(status, str(status))
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        messages = {
            "INFEASIBLE": "No existe una combinación que satisfaga todas las restricciones obligatorias con los datos actuales.",
            "MODEL_INVALID": "El modelo de optimización fue considerado inválido por CP-SAT.",
            "UNKNOWN": "CP-SAT no pudo determinar una solución dentro de las condiciones de búsqueda establecidas.",
        }
        return {"estado": "SIN_SOLUCION", "estado_solver": status_name,
                "mensaje": messages.get(status_name, "No se encontró una solución factible.")}

    # ============================================================
    # 15. EXTRAER Y VALIDAR SOLUCIÓN
    # ============================================================
    resultados = []
    slot_by_id = {s["id"]: s for s in slots}
    for (evento_id, profesor, espacio, sid), variable in x.items():
        if solver.Value(variable) == 1:
            evento = eventos[evento_id]
            s = slot_by_id[sid]
            resultados.append({
                "Grupo": evento["grupo"], "Materia": evento["materia"],
                "Sesion": evento["sesion"], "Profesor": profesor,
                "Espacio": espacio, "Horario": s["label"]
            })

    df_resultado = pd.DataFrame(resultados)
    if df_resultado.empty:
        return {"estado": "ERROR", "mensaje": "El modelo encontró una solución, pero no se generaron asignaciones."}

    # ------------------------------------------------------------
    # Validación independiente post-solver. No se confía únicamente
    # en la formulación CP-SAT: comprobamos de nuevo la solución extraída.
    # ------------------------------------------------------------
    violaciones = []
    if len(df_resultado) != len(eventos):
        violaciones.append(f"sesiones asignadas {len(df_resultado)}/{len(eventos)}")

    # Duplicidad/ausencia de identificadores de evento lógico.
    counts = df_resultado.groupby(["Grupo", "Materia", "Sesion"]).size()
    if (counts != 1).any() or len(counts) != len(eventos):
        violaciones.append("sesiones faltantes o duplicadas")

    # Reconstruir slot desde la etiqueta, que procede del catálogo interno.
    label_to_slot = {s["label"]: s for s in slots}
    for horario, sub in df_resultado.groupby("Horario"):
        if sub["Grupo"].duplicated().any():
            violaciones.append(f"colisión de grupo en {horario}")
        if sub["Profesor"].duplicated().any():
            violaciones.append(f"colisión docente en {horario}")
        if sub["Espacio"].duplicated().any():
            violaciones.append(f"colisión de espacio en {horario}")

    for _, r in df_resultado.iterrows():
        s = label_to_slot.get(r["Horario"])
        if s is None:
            violaciones.append(f"horario desconocido: {r['Horario']}")
            continue
        prof = r["Profesor"]
        if r["Materia"] not in info_personal.get(prof, {}).get("materias", []):
            violaciones.append(f"profesor no habilitado: {prof} - {r['Materia']}")
        if not disponibilidad_personal.get((prof, s["id"]), False):
            violaciones.append(f"profesor no disponible: {prof} - {r['Horario']}")
        # Recuperar requisitos del evento exacto para validar el espacio.
        matches = [e for e in eventos if e["grupo"] == r["Grupo"] and e["materia"] == r["Materia"] and e["sesion"] == int(r["Sesion"])]
        if not matches or not espacio_compatible(matches[0], r["Espacio"]):
            violaciones.append(f"espacio incompatible: {r['Espacio']} - {r['Grupo']} - {r['Materia']}")

    if violaciones:
        return {
            "estado": "ERROR_VALIDACION",
            "estado_solver": status_name,
            "mensaje": "La solución del solver no superó la validación independiente.",
            "violaciones": sorted(set(violaciones)),
            "solucion": df_resultado,
        }

    # Métricas calculadas directamente de la solución.
    total_huecos = sum(solver.Value(v) for v in gap_vars) if gap_vars else 0
    permanencia_total_min = sum(solver.Value(v) for v in span_vars) if span_vars else 0
    dias_activos = 0
    for grupo in grupos:
        for dia, day_slots in slots_by_day.items():
            if any((grupo, s["id"]) in y and solver.Value(y[(grupo, s["id"])]) for s in day_slots):
                dias_activos += 1
    permanencia_media = permanencia_total_min / dias_activos if dias_activos else 0
    asignaciones_tardias = 0
    if hora_preferida_salida is not None:
        for grupo in grupos:
            for s in slots:
                v = y.get((grupo, s["id"]))
                if v is not None and s["fin"] > hora_preferida_salida and solver.Value(v):
                    asignaciones_tardias += 1

    mensaje = (
        "Se encontró una solución óptima según el modelo y los criterios establecidos."
        if status == cp_model.OPTIMAL
        else "Se encontró una solución factible dentro del tiempo de búsqueda establecido; no se demostró optimalidad."
    )

    return {
        "estado": "OK",
        "estado_solver": status_name,
        "mensaje": mensaje,
        "grupos": len(grupos),
        "personal": len(profesores),
        "espacios": len(espacios),
        "horarios": len(slots),
        "disponibilidad": len(df_disponibilidad),
        "configuracion": len(df_configuracion),
        "eventos": len(eventos),
        "metricas": {
            "huecos_internos": int(total_huecos),
            "permanencia_media_minutos": round(permanencia_media, 2),
            "asignaciones_tardias": int(asignaciones_tardias),
            "hora_preferida_salida": _minutes_to_hhmm(hora_preferida_salida) if hora_preferida_salida is not None else None,
            "desequilibrio_docente": int(solver.Value(balance_var)) if balance_var is not None else 0,
            "sesiones_asignadas": int(len(df_resultado)),
            "sesiones_requeridas": int(len(eventos)),
        },
        "solver": {
            "status": status_name,
            "objective_value": float(solver.ObjectiveValue()) if objective_terms else None,
            "best_objective_bound": float(solver.BestObjectiveBound()) if objective_terms else None,
            "wall_time_segundos": round(float(solver.WallTime()), 4),
            "num_conflicts": int(solver.NumConflicts()),
            "num_branches": int(solver.NumBranches()),
            "max_time_segundos": float(tiempo_solver),
            "num_search_workers": int(workers_solver),
            "semilla_solver": int(semilla_solver),
        },
        "validacion": {"ok": True, "violaciones": []},
        "solucion": df_resultado,
    }
