Attribute VB_Name = "Módulo11"
Sub Worksheet_Change()
Dim xRg As Range
Worksheets("Excel_Licitacion_Publicada").Activate
Application.ScreenUpdating = False
    For Each xRg In Range("G8:G18080")
        If InStr(1, xRg, "ospital", vbTextCompare) Then
            xRg.EntireRow.Hidden = True
        'ElseIf xRg.Value = "unicipalidad" Then
        ElseIf InStr(1, xRg, "unicipalidad", vbTextCompare) And InStr(1, xRg, "ondes", vbTextCompare) Then
            xRg.EntireRow.Hidden = False
        ElseIf InStr(1, xRg, "unicipalidad", vbTextCompare) And InStr(1, xRg, "itacura", vbTextCompare) Then
            xRg.EntireRow.Hidden = False
        ElseIf InStr(1, xRg, "unicipalidad", vbTextCompare) Then
            xRg.EntireRow.Hidden = True
        ElseIf InStr(1, xRg, "sistencial", vbTextCompare) Then
            xRg.EntireRow.Hidden = True
        ElseIf InStr(1, xRg, "endarme", vbTextCompare) Then
            xRg.EntireRow.Hidden = True
        ElseIf InStr(1, xRg, "arabineros", vbTextCompare) Then
            xRg.EntireRow.Hidden = True
        ElseIf InStr(1, xRg, "onsultorio", vbTextCompare) Then
            xRg.EntireRow.Hidden = True
        ElseIf InStr(1, xRg, "raumatológico", vbTextCompare) Then
            xRg.EntireRow.Hidden = True
        ElseIf InStr(1, xRg, "alud", vbTextCompare) Then
            xRg.EntireRow.Hidden = True
        ElseIf InStr(1, xRg, "SERVIU", vbTextCompare) Then
            xRg.EntireRow.Hidden = True
        ElseIf InStr(1, xRg, "Gobernación", vbTextCompare) Then
            xRg.EntireRow.Hidden = True
        ElseIf InStr(1, xRg, "Tribunal", vbTextCompare) Then
            xRg.EntireRow.Hidden = True
        ElseIf InStr(1, xRg, "CONAF", vbTextCompare) Then
            xRg.EntireRow.Hidden = True
        ElseIf InStr(1, xRg, "MINISTERIO DE LAS CULTURAS Y LAS ARTES", vbTextCompare) Then
            xRg.EntireRow.Hidden = True
        ElseIf InStr(1, xRg, "PREVISION SOCIAL", vbTextCompare) Then
            xRg.EntireRow.Hidden = True
        ElseIf InStr(1, xRg, "JUNJI", vbTextCompare) Then
            xRg.EntireRow.Hidden = True
        ElseIf InStr(1, xRg, "INDAP", vbTextCompare) Then
            xRg.EntireRow.Hidden = True
        ElseIf InStr(1, xRg, "CONADI", vbTextCompare) Then
            xRg.EntireRow.Hidden = True
        ElseIf InStr(1, xRg, "CENABAST", vbTextCompare) Then
            xRg.EntireRow.Hidden = True
        ElseIf InStr(1, xRg, "Poder Judicial", vbTextCompare) Then
            xRg.EntireRow.Hidden = True
        ElseIf InStr(1, xRg, "Intendencia", vbTextCompare) Then
            xRg.EntireRow.Hidden = True
        ElseIf InStr(1, xRg, "Senado de la República de Chile", vbTextCompare) Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "JUNAEB" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "ARMADA DE CHILE" Then
            xRg.EntireRow.Hidden = True '
        ElseIf xRg.Value = "MINISTERIO DE VIVIENDA Y URBANISMO" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Fuerza Aérea de Chile" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Dirección del Trabajo" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Policía de Investigaciones de Chile" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Contraloría General de la República" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Servicio Nacional del Patrimonio Cultural" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Servicio Agrícola y Ganadero" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Instituto de Neurocirugía" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Defensoría Penal Pública" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Servicio Nacional de Capacitación y Empleo" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Instituto Nacional del Cancer" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Presidencia de la República" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Consejo de Defensa del Estado" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Instituto Nacional del Torax" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Ministerio de Educación" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Instituto Psiquiátrico Dr. José Horwitz Barak" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Servicio Nacional de Menores" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Subsecretaría de R.R.E.E" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Ministerio Secretaría General de Gobierno" Then
            xRg.EntireRow.Hidden = True
        Else
            xRg.EntireRow.Hidden = False
        End If
    Next xRg
Application.ScreenUpdating = False
End Sub



'Instituto de Seguridad Laboral - ISL
'Servicio Nacional del Patrimonio Cultural
'Servicio Agrícola y Ganadero
'Servicio Nacional de Aduanas
'Servicio Nacional de Pesca y Acuicultura
'Dirección de Relaciones Económicas
'Ministerio de Educación
'Servicio Nacional de la Mujer
'Fundación de las Familias
'Superintendencia de Pensiones
'Defensoría Penal Pública
'Gobierno Regional Metropolitano
'Corporacion Administrativa del Poder Judicial
'Servicio Nacional de Geología y Minería
'Ministerio de Educación
'Servicio de Impuestos Internos
'Servicio Médico Legal
'Instituto Nacional de Estadísticas
'Ministerio Secretaría General de Gobierno
'Agencia de Calidad de la Educación

















