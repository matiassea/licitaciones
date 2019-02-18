Attribute VB_Name = "M�dulo11"
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
        ElseIf InStr(1, xRg, "raumatol�gico", vbTextCompare) Then
            xRg.EntireRow.Hidden = True
        ElseIf InStr(1, xRg, "alud", vbTextCompare) Then
            xRg.EntireRow.Hidden = True
        ElseIf InStr(1, xRg, "SERVIU", vbTextCompare) Then
            xRg.EntireRow.Hidden = True
        ElseIf InStr(1, xRg, "Gobernaci�n", vbTextCompare) Then
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
        ElseIf InStr(1, xRg, "Senado de la Rep�blica de Chile", vbTextCompare) Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "JUNAEB" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "ARMADA DE CHILE" Then
            xRg.EntireRow.Hidden = True '
        ElseIf xRg.Value = "MINISTERIO DE VIVIENDA Y URBANISMO" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Fuerza A�rea de Chile" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Direcci�n del Trabajo" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Polic�a de Investigaciones de Chile" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Contralor�a General de la Rep�blica" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Servicio Nacional del Patrimonio Cultural" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Servicio Agr�cola y Ganadero" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Instituto de Neurocirug�a" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Defensor�a Penal P�blica" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Servicio Nacional de Capacitaci�n y Empleo" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Instituto Nacional del Cancer" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Presidencia de la Rep�blica" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Consejo de Defensa del Estado" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Instituto Nacional del Torax" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Ministerio de Educaci�n" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Instituto Psiqui�trico Dr. Jos� Horwitz Barak" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Servicio Nacional de Menores" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Subsecretar�a de R.R.E.E" Then
            xRg.EntireRow.Hidden = True
        ElseIf xRg.Value = "Ministerio Secretar�a General de Gobierno" Then
            xRg.EntireRow.Hidden = True
        Else
            xRg.EntireRow.Hidden = False
        End If
    Next xRg
Application.ScreenUpdating = False
End Sub



'Instituto de Seguridad Laboral - ISL
'Servicio Nacional del Patrimonio Cultural
'Servicio Agr�cola y Ganadero
'Servicio Nacional de Aduanas
'Servicio Nacional de Pesca y Acuicultura
'Direcci�n de Relaciones Econ�micas
'Ministerio de Educaci�n
'Servicio Nacional de la Mujer
'Fundaci�n de las Familias
'Superintendencia de Pensiones
'Defensor�a Penal P�blica
'Gobierno Regional Metropolitano
'Corporacion Administrativa del Poder Judicial
'Servicio Nacional de Geolog�a y Miner�a
'Ministerio de Educaci�n
'Servicio de Impuestos Internos
'Servicio M�dico Legal
'Instituto Nacional de Estad�sticas
'Ministerio Secretar�a General de Gobierno
'Agencia de Calidad de la Educaci�n

















