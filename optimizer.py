import pandas as pd
from ortools.sat.python import cp_model


def optimizar_desde_excel(archivo_excel):

    # ============================================================
    # 1. LEER ARCHIVO EXCEL
    # ============================================================

    try:

        df_grupos = pd.read_excel(
            archivo_excel,
            sheet_name="Grupos"
        )

        df_personal = pd.read_excel(
            archivo_excel,
            sheet_name="Personal"
        )

        df_espacios = pd.read_excel(
            archivo_excel,
            sheet_name="Espacios"
        )

        df_horarios = pd.read_excel(
            archivo_excel,
            sheet_name="Horarios"
        )

        df_disponibilidad = pd.read_excel(
            archivo_excel,
            sheet_name="Disponibilidad"
        )

        df_configuracion = pd.read_excel(
            archivo_excel,
            sheet_name="Configuracion"
        )

    except Exception as error:

        return {
            "estado": "ERROR",
            "mensaje": f"No fue posible leer el Excel: {error}"
        }

    # ============================================================
    # 2. LIMPIAR NOMBRES DE COLUMNAS
    # ============================================================

    for df_temp in [
        df_grupos,
        df_personal,
        df_espacios,
        df_horarios,
        df_disponibilidad,
        df_configuracion
    ]:

        df_temp.columns = df_temp.columns.str.strip()

        # ============================================================
    # 3. CONFIGURACIÓN
    # ============================================================

    config = dict(
        zip(
            df_configuracion["Parametro"],
            df_configuracion["Valor"]
        )
    )

    max_clases_dia = int(
        config.get("Max_clases_dia", 3)
    )

    mantener_mismo_profesor = (
        str(
            config.get(
                "Mantener_mismo_profesor",
                "Si"
            )
        ).strip().lower()
        in ["si", "sí", "yes", "true", "1"]
    )

    permitir_sobrecupo = (
        str(
            config.get(
                "Permitir_sobrecupo",
                "No"
            )
        ).strip().lower()
        in ["si", "sí", "yes", "true", "1"]
    )

    # ============================================================
    # 4. LISTAS PRINCIPALES
    # ============================================================

    grupos = (
        df_grupos["Grupo"]
        .astype(str)
        .unique()
        .tolist()
    )

    profesores = (
        df_personal["Persona"]
        .astype(str)
        .unique()
        .tolist()
    )

    espacios = (
        df_espacios["Espacio"]
        .astype(str)
        .unique()
        .tolist()
    )

    # ============================================================
    # 5. CREAR EVENTOS
    # ============================================================

    eventos = []

    for _, fila in df_grupos.iterrows():

        grupo = str(fila["Grupo"])
        materia = str(fila["Materia"])
        personas = int(fila["Personas"])
        sesiones = int(fila["Sesiones_semana"])
        tipo_recurso = str(fila["Tipo_recurso"])

        for sesion in range(1, sesiones + 1):

            eventos.append(
                {
                    "grupo": grupo,
                    "materia": materia,
                    "sesion": sesion,
                    "personas": personas,
                    "tipo": tipo_recurso
                }
            )

    # ============================================================
    # 6. HORARIOS DISPONIBLES
    # ============================================================

    horarios = []

    for _, fila in df_horarios.iterrows():

        disponible = str(
            fila["Disponible"]
        ).strip().lower()

        if disponible not in [
            "si",
            "sí",
            "yes",
            "true",
            "1"
        ]:
            continue

        dia = str(fila["Dia"])
        inicio = str(fila["Hora_inicio"])
        fin = str(fila["Hora_fin"])

        horarios.append(
            f"{dia} {inicio}-{fin}"
        )

    # ============================================================
    # 7. INFORMACIÓN DE ESPACIOS
    # ============================================================

    info_espacios = {}

    for _, fila in df_espacios.iterrows():

        nombre = str(fila["Espacio"])

        info_espacios[nombre] = {
            "capacidad": int(fila["Capacidad"]),
            "tipo": str(fila["Tipo"]),
            "computadoras": int(fila["Computadoras"]),
            "proyector": (
                str(
                    fila["Proyector"]
                ).strip().lower()
                in ["si", "sí", "yes", "true", "1"]
            )
        }

    # ============================================================
    # 8. INFORMACIÓN DEL PERSONAL
    # ============================================================

    info_personal = {}

    for _, fila in df_personal.iterrows():

        persona = str(fila["Persona"])

        materias = [
            x.strip()
            for x in str(
                fila["Materias_habilitadas"]
            ).split(",")
        ]

        info_personal[persona] = {
            "materias": materias,
            "max_semana": int(
                fila["Max_clases_semana"]
            ),
            "max_dia": int(
                fila["Max_clases_dia"]
            )
        }
    # ============================================================
    # 9. DISPONIBILIDAD DEL PERSONAL
    # ============================================================

    def parse_time(time_str):
        h, m = map(int, str(time_str).split(":"))
        return h * 60 + m

    disponibilidad_personal = {}

    prof_ranges = {}

    for _, fila in df_disponibilidad.iterrows():

        persona = str(fila["Persona"])
        dia = str(fila["Dia"])
        inicio_str = str(fila["Hora_inicio"])
        fin_str = str(fila["Hora_fin"])

        disponible = str(
            fila["Disponible"]
        ).strip().lower()

        if disponible in [
            "si",
            "sí",
            "yes",
            "true",
            "1"
        ]:

            if persona not in prof_ranges:
                prof_ranges[persona] = {}

            if dia not in prof_ranges[persona]:
                prof_ranges[persona][dia] = []

            prof_ranges[persona][dia].append(
                (
                    parse_time(inicio_str),
                    parse_time(fin_str)
                )
            )

    for profesor in profesores:

        for horario_slot in horarios:

            dia_slot = horario_slot.split(" ")[0]

            rango_hora = horario_slot.split(" ")[1]

            inicio_slot_str, fin_slot_str = (
                rango_hora.split("-")
            )

            inicio_slot = parse_time(
                inicio_slot_str
            )

            fin_slot = parse_time(
                fin_slot_str
            )

            esta_disponible = False

            if (
                profesor in prof_ranges
                and dia_slot in prof_ranges[profesor]
            ):

                for (
                    prof_inicio,
                    prof_fin
                ) in prof_ranges[profesor][dia_slot]:

                    if (
                        inicio_slot >= prof_inicio
                        and fin_slot <= prof_fin
                    ):

                        esta_disponible = True
                        break

            disponibilidad_personal[
                (profesor, horario_slot)
            ] = esta_disponible



                # ============================================================
    # 10. COMPATIBILIDAD DE ESPACIOS
    # ============================================================

    def espacio_compatible(evento, espacio):

        info = info_espacios[espacio]

        # Capacidad
        if not permitir_sobrecupo:

            if info["capacidad"] < evento["personas"]:
                return False

        # Tipo de recurso
        if (
            evento["tipo"].strip().lower()
            != info["tipo"].strip().lower()
        ):
            return False

        return True

    


     # ==========================================================
    # 11. CREAR MODELO DE OPTIMIZACIÓN
    # ==========================================================

    model = cp_model.CpModel()




    # ============================================================
    # 11. CREAR MODELO DE OPTIMIZACIÓN
    # ============================================================

    model = cp_model.CpModel()

    # ============================================================
    # 12. VARIABLES DE DECISIÓN
    # ============================================================

    x = {}

    for i, evento in enumerate(eventos):
        materia = evento["materia"]

        for profesor in profesores:

            if materia not in info_personal[profesor]["materias"]:
                continue

            for espacio in espacios:

                if not espacio_compatible(evento, espacio):
                    continue

                for horario in horarios:

                    if not disponibilidad_personal.get(
                        (profesor, horario),
                        False
                    ):
                        continue

                    x[
                        (i, profesor, espacio, horario)
                    ] = model.NewBoolVar(
                        f"x_{i}_{profesor}_{espacio}_{horario}"
                    )
    # ============================================================
    # 13. CADA EVENTO DEBE PROGRAMARSE EXACTAMENTE UNA VEZ
    # ============================================================

    for i in range(len(eventos)):

        variables_evento = [
            variable
            for (
                evento_id,
                profesor,
                espacio,
                horario
            ), variable in x.items()
            if evento_id == i
        ]

        if not variables_evento:
            return {
                "estado": "ERROR",
                "mensaje": (
                    f"El evento {i + 1}: "
                    f"{eventos[i]['grupo']} - "
                    f"{eventos[i]['materia']} "
                    "no tiene ninguna combinación posible."
                )
            }

        model.Add(sum(variables_evento) == 1)
            # ============================================================
    # 14. NO DOS CLASES DEL MISMO GRUPO AL MISMO TIEMPO
    # ============================================================

    for grupo in grupos:

        for horario in horarios:

            variables_grupo = []

            for (
                evento_id,
                profesor,
                espacio,
                h
            ), variable in x.items():

                if (
                    eventos[evento_id]["grupo"] == grupo
                    and h == horario
                ):
                    variables_grupo.append(variable)

            if variables_grupo:
                model.Add(
                    sum(variables_grupo) <= 1
                )
    # ============================================================
    # 15. NO DOS CLASES EN EL MISMO ESPACIO AL MISMO TIEMPO
    # ============================================================

    for espacio in espacios:

        for horario in horarios:

            variables_espacio = []

            for (
                evento_id,
                profesor,
                e,
                h
            ), variable in x.items():

                if (
                    e == espacio
                    and h == horario
                ):
                    variables_espacio.append(variable)

            if variables_espacio:
                model.Add(
                    sum(variables_espacio) <= 1
                )
                    # ============================================================
    # 16. NO DOS CLASES DEL MISMO PROFESOR AL MISMO TIEMPO
    # ============================================================

    for profesor in profesores:

        for horario in horarios:

            variables_profesor = []

            for (
                evento_id,
                p,
                espacio,
                h
            ), variable in x.items():

                if (
                    p == profesor
                    and h == horario
                ):
                    variables_profesor.append(variable)

            if variables_profesor:
                model.Add(
                    sum(variables_profesor) <= 1
                )
                    # ============================================================
    # 17. MÁXIMO DE CLASES SEMANALES POR PROFESOR
    # ============================================================

    for profesor in profesores:

        variables_profesor = [
            variable
            for (
                evento_id,
                p,
                espacio,
                horario
            ), variable in x.items()
            if p == profesor
        ]

        if variables_profesor:
            model.Add(
                sum(variables_profesor)
                <= info_personal[profesor]["max_semana"]
            )
                # ============================================================
    # 18. MÁXIMO DE CLASES POR PROFESOR EN UN MISMO DÍA
    # ============================================================

    dias = (
        df_horarios["Dia"]
        .astype(str)
        .unique()
        .tolist()
    )

    for profesor in profesores:

        limite_diario = info_personal[profesor]["max_dia"]

        for dia in dias:

            variables_dia = []

            for (
                evento_id,
                p,
                espacio,
                horario
            ), variable in x.items():

                if (
                    p == profesor
                    and horario.startswith(dia + " ")
                ):
                    variables_dia.append(variable)

            if variables_dia:
                model.Add(
                    sum(variables_dia) <= limite_diario
                )
                    # ============================================================
    # 19. MANTENER MISMO PROFESOR POR GRUPO Y MATERIA
    # ============================================================

    if mantener_mismo_profesor:

        pares_grupo_materia = set(
            (evento["grupo"], evento["materia"])
            for evento in eventos
        )

        for grupo, materia in pares_grupo_materia:

            eventos_relacionados = [
                i
                for i, evento in enumerate(eventos)
                if (
                    evento["grupo"] == grupo
                    and evento["materia"] == materia
                )
            ]

            if len(eventos_relacionados) <= 1:
                continue

            for profesor in profesores:

                variables_por_evento = {}

                for i in eventos_relacionados:

                    variables_por_evento[i] = [
                        variable
                        for (
                            evento_id,
                            p,
                            espacio,
                            horario
                        ), variable in x.items()
                        if (
                            evento_id == i
                            and p == profesor
                        )
                    ]

                evento_base = eventos_relacionados[0]

                for otro_evento in eventos_relacionados[1:]:

                    model.Add(
                        sum(variables_por_evento[evento_base])
                        ==
                        sum(variables_por_evento[otro_evento])
                    )
                        # ============================================================
    # 20. RESOLVER EL MODELO
    # ============================================================

    solver = cp_model.CpSolver()

    solver.parameters.max_time_in_seconds = 30.0
    solver.parameters.num_search_workers = 1

    status = solver.Solve(model)

    if status not in (
        cp_model.OPTIMAL,
        cp_model.FEASIBLE
    ):
        return {
            "estado": "SIN_SOLUCION",
            "mensaje": (
                "No se encontró una solución factible "
                "con las restricciones actuales."
            )
        }
        # ============================================================
    # 21. EXTRAER LA SOLUCIÓN
    # ============================================================

    resultados = []

    for (
        evento_id,
        profesor,
        espacio,
        horario
    ), variable in x.items():

        if solver.Value(variable) == 1:

            evento = eventos[evento_id]

            resultados.append(
                {
                    "Grupo": evento["grupo"],
                    "Materia": evento["materia"],
                    "Sesion": evento["sesion"],
                    "Profesor": profesor,
                    "Espacio": espacio,
                    "Horario": horario
                }
            )

    df_resultado = pd.DataFrame(resultados)

    if df_resultado.empty:
        return {
            "estado": "ERROR",
            "mensaje": (
                "El modelo encontró una solución, "
                "pero no se generaron asignaciones."
            )
        }

    return {
        "estado": "OK",
        "mensaje": "Horario optimizado correctamente.",
        "grupos": len(grupos),
        "personal": len(profesores),
        "espacios": len(espacios),
        "horarios": len(horarios),
        "disponibilidad": len(df_disponibilidad),
        "configuracion": len(df_configuracion),
        "eventos": len(eventos),
        "solucion": df_resultado
    }
