Attribute VB_Name = "Congfi"
Option Explicit

' ==============================================================================
' GESTIÓN Y SINCRONIZACIÓN DE PLANTILLA DE PROYECTOS
' ==============================================================================

Sub GenerarPlantillaProyectos()
    Dim wsActiva As Worksheet
    Dim wsInicio As Worksheet
    Dim tblConfig As ListObject, tblRamalConfig As ListObject
    Dim rngCant As Range
    Dim numProyectos As Long, currentCount As Long
    Dim i As Long, j As Long, r As Long
    Dim colNo As Long, colAbrev As Long, colAbrevAnt As Long
    Dim strAbrev As String, numX As String, nomDefecto As String
    Dim baseSheets As Variant, ramalBaseSheets As Variant, baseName As Variant
    Dim wsBase As Worksheet, wsNew As Worksheet
    Dim newSheetName As String, newTblName As String
    Dim tblBase As ListObject
    Dim expectedName As String
    Dim newRow As ListRow
    Dim rowIdx As Long
    Dim strRamalAbrev As String, strRamalAbrevAnt As String

    Set wsActiva = ActiveSheet

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo ErrorHandler

    Set wsInicio = ThisWorkbook.Sheets("Inicio")

    On Error Resume Next
    Set rngCant = wsInicio.Range("C_Coumindad")
    On Error GoTo ErrorHandler

    If Not rngCant Is Nothing Then
        numProyectos = val(rngCant.Value)
    Else
        numProyectos = val(wsInicio.Range("C7").Value)
    End If

    If numProyectos < 1 Then numProyectos = 1

    Set tblConfig = wsInicio.ListObjects("Configuration")

    colNo = ObtenerColumna(tblConfig, Array("No.", "No", "Número"))
    colAbrev = ObtenerColumna(tblConfig, Array("Abreviación", "Abreviacion", "Proyecto"))
    colAbrevAnt = ObtenerColumna(tblConfig, Array("Abrevi_Ant", "Abrevi_Anterior", "Abreviacion_Ant"))

    If colAbrev = 0 Then
        MsgBox "No se encontró la columna 'Abreviación' en la tabla Configuration.", vbExclamation, "Error"
        GoTo CleanUp
    End If

    baseSheets = Array("ANE. - COMUNIDAD", "ANE. - VAR. - COMUNIDAD", _
                       "VARIACIÓN - COMUNIDAD", "OC - COMUNIDAD", _
                       "OTS - COMUNIDAD", "ATE - COMUNIDAD")

    For Each baseName In baseSheets
        If SheetExists(CStr(baseName)) Then
            ThisWorkbook.Sheets(CStr(baseName)).Visible = xlSheetVisible
        End If
    Next baseName

    currentCount = tblConfig.ListRows.Count

    If currentCount = 1 Then
        If Trim(CStr(tblConfig.DataBodyRange(1, colAbrev).Value)) = "" Then
            currentCount = 0
        End If
    End If

    ' --- ELIMINAR COMUNIDADES SOBRANTES ---
    If currentCount > numProyectos Then
        For r = currentCount To numProyectos + 1 Step -1
            strAbrev = Trim(CStr(tblConfig.DataBodyRange(r, colAbrev).Value))
            If strAbrev = "" And colAbrevAnt > 0 Then
                strAbrev = Trim(CStr(tblConfig.DataBodyRange(r, colAbrevAnt).Value))
            End If

            If strAbrev <> "" Then
                For Each baseName In baseSheets
                    expectedName = Replace(CStr(baseName), "COMUNIDAD", strAbrev, , , vbTextCompare)
                    If SheetExists(expectedName) Then
                        ThisWorkbook.Sheets(expectedName).Delete
                    End If
                Next baseName
            End If

            tblConfig.ListRows(r).Delete
        Next r
    End If

    ' --- AGREGAR REGISTROS NUEVOS ---
    If currentCount < numProyectos Then
        For i = currentCount + 1 To numProyectos
            If i = 1 And tblConfig.ListRows.Count >= 1 Then
                Set newRow = tblConfig.ListRows(1)
            Else
                Set newRow = tblConfig.ListRows.Add
            End If

            rowIdx = newRow.Index

            If colNo > 0 Then tblConfig.DataBodyRange(rowIdx, colNo).Value = i

            If colNo > 0 Then
                numX = Trim(CStr(tblConfig.DataBodyRange(rowIdx, colNo).Value))
            Else
                numX = CStr(i)
            End If

            nomDefecto = "COMUNIDAD_" & numX

            tblConfig.DataBodyRange(rowIdx, colAbrev).Value = nomDefecto
            If colAbrevAnt > 0 Then tblConfig.DataBodyRange(rowIdx, colAbrevAnt).Value = nomDefecto
        Next i
    End If

    For i = 1 To numProyectos
        If colNo > 0 Then
            If Trim(CStr(tblConfig.DataBodyRange(i, colNo).Value)) = "" Then
                tblConfig.DataBodyRange(i, colNo).Value = i
            End If
        End If
    Next i

    ' --- CREAR HOJAS DE COMUNIDADES ---
    For i = 1 To numProyectos
        strAbrev = Trim(CStr(tblConfig.DataBodyRange(i, colAbrev).Value))

        If strAbrev <> "" Then
            If InStr(1, strAbrev, " ") > 0 Then
                strAbrev = Replace(strAbrev, " ", "_")
                tblConfig.DataBodyRange(i, colAbrev).Value = strAbrev
                If colAbrevAnt > 0 Then tblConfig.DataBodyRange(i, colAbrevAnt).Value = strAbrev
            End If

            For Each baseName In baseSheets
                If SheetExists(CStr(baseName)) Then
                    Set wsBase = ThisWorkbook.Sheets(CStr(baseName))
                    newSheetName = Replace(CStr(baseName), "COMUNIDAD", strAbrev, , , vbTextCompare)

                    If Not SheetExists(newSheetName) Then
                        wsBase.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                        Set wsNew = ActiveSheet
                        wsNew.Name = newSheetName

                        On Error Resume Next
                        wsNew.Cells.Replace What:="COMUNIDAD", Replacement:=strAbrev, LookAt:=xlPart, MatchCase:=False
                        On Error GoTo ErrorHandler

                        For j = 1 To wsBase.ListObjects.Count
                            Set tblBase = wsBase.ListObjects(j)
                            newTblName = Replace(tblBase.Name, "COMUNIDAD", strAbrev, , , vbTextCompare)

                            On Error Resume Next
                            wsNew.ListObjects(j).Name = newTblName
                            On Error GoTo ErrorHandler
                        Next j

                        wsNew.Visible = xlSheetHidden
                    End If
                End If
            Next baseName
        End If
    Next i

    ' --- GESTIÓN DE HOJAS DE RAMAL ---
    ramalBaseSheets = Array("ANE. - RAMAL", "ANE. - VAR. - RAMAL", _
                            "VARIACIÓN - RAMAL", "OC - RAMAL", _
                            "OTS - RAMAL", "ATE - RAMAL")

    On Error Resume Next
    Set tblRamalConfig = wsInicio.ListObjects("Configuration_Ramal")
    On Error GoTo ErrorHandler

    If Not tblRamalConfig Is Nothing Then
        colAbrev = ObtenerColumna(tblRamalConfig, Array("Abreviación", "Abreviacion", "Ramal", "Abrev"))
        colAbrevAnt = ObtenerColumna(tblRamalConfig, Array("Abrevi_Ant", "Abrevi_Anterior", "Abreviacion_Ant"))

        If colAbrev > 0 Then
            If numProyectos <= 1 Then
                For i = 1 To tblRamalConfig.ListRows.Count
                    strRamalAbrev = Trim(CStr(tblRamalConfig.DataBodyRange(i, colAbrev).Value))
                    If strRamalAbrev = "" Then strRamalAbrev = "RAMAL_" & i

                    For Each baseName In ramalBaseSheets
                        expectedName = Replace(CStr(baseName), "RAMAL", strRamalAbrev, , , vbTextCompare)
                        If SheetExists(expectedName) Then
                            ThisWorkbook.Sheets(expectedName).Delete
                        End If
                    Next baseName

                    If colAbrevAnt > 0 Then
                        strRamalAbrevAnt = Trim(CStr(tblRamalConfig.DataBodyRange(i, colAbrevAnt).Value))
                        If strRamalAbrevAnt <> "" And strRamalAbrevAnt <> strRamalAbrev Then
                            For Each baseName In ramalBaseSheets
                                expectedName = Replace(CStr(baseName), "RAMAL", strRamalAbrevAnt, , , vbTextCompare)
                                If SheetExists(expectedName) Then
                                    ThisWorkbook.Sheets(expectedName).Delete
                                End If
                            Next baseName
                        End If
                    End If
                Next i
            Else
                For Each baseName In ramalBaseSheets
                    If SheetExists(CStr(baseName)) Then
                        ThisWorkbook.Sheets(CStr(baseName)).Visible = xlSheetVisible
                    End If
                Next baseName

                For i = 1 To tblRamalConfig.ListRows.Count
                    strRamalAbrev = Trim(CStr(tblRamalConfig.DataBodyRange(i, colAbrev).Value))

                    If strRamalAbrev = "" Then strRamalAbrev = "RAMAL_" & i

                    If InStr(1, strRamalAbrev, " ") > 0 Then
                        strRamalAbrev = Replace(strRamalAbrev, " ", "_")
                        tblRamalConfig.DataBodyRange(i, colAbrev).Value = strRamalAbrev
                    End If

                    If colAbrevAnt > 0 Then
                        If Trim(CStr(tblRamalConfig.DataBodyRange(i, colAbrevAnt).Value)) = "" Then
                            tblRamalConfig.DataBodyRange(i, colAbrevAnt).Value = strRamalAbrev
                        End If
                    End If

                    For Each baseName In ramalBaseSheets
                        If SheetExists(CStr(baseName)) Then
                            Set wsBase = ThisWorkbook.Sheets(CStr(baseName))
                            newSheetName = Replace(CStr(baseName), "RAMAL", strRamalAbrev, , , vbTextCompare)

                            If Not SheetExists(newSheetName) Then
                                wsBase.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                                Set wsNew = ActiveSheet
                                wsNew.Name = newSheetName

                                On Error Resume Next
                                wsNew.Cells.Replace What:="RAMAL", Replacement:=strRamalAbrev, LookAt:=xlPart, MatchCase:=False
                                On Error GoTo ErrorHandler

                                For j = 1 To wsBase.ListObjects.Count
                                    Set tblBase = wsBase.ListObjects(j)
                                    newTblName = Replace(tblBase.Name, "RAMAL", strRamalAbrev, , , vbTextCompare)

                                    On Error Resume Next
                                    wsNew.ListObjects(j).Name = newTblName
                                    On Error GoTo ErrorHandler
                                Next j

                                wsNew.Visible = xlSheetHidden
                            End If
                        End If
                    Next baseName
                Next i

                For Each baseName In ramalBaseSheets
                    If SheetExists(CStr(baseName)) Then
                        ThisWorkbook.Sheets(CStr(baseName)).Visible = xlSheetHidden
                    End If
                Next baseName
            End If
        End If
    End If

    Call ActualizarFormulasRamal

    MsgBox "¡Plantilla sincronizada con éxito!" & vbCrLf & _
           "Proyectos activos: " & numProyectos, vbInformation, "Proceso Completado"

CleanUp:
    On Error Resume Next
    For Each baseName In baseSheets
        If SheetExists(CStr(baseName)) Then
            ThisWorkbook.Sheets(CStr(baseName)).Visible = xlSheetHidden
        End If
    Next baseName
    For Each baseName In Array("ANE. - RAMAL", "ANE. - VAR. - RAMAL", "VARIACIÓN - RAMAL", "OC - RAMAL", "OTS - RAMAL", "ATE - RAMAL")
        If SheetExists(CStr(baseName)) Then
            ThisWorkbook.Sheets(CStr(baseName)).Visible = xlSheetHidden
        End If
    Next baseName
    On Error GoTo 0

    wsActiva.Activate
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "Ocurrió un error al generar: " & Err.Description, vbCritical, "Error en Macro"
    Resume CleanUp
End Sub

Public Sub ActualizarFormulasRamal()
    Dim wsInicio As Worksheet
    Dim tblConfig As ListObject, tblRamalConfig As ListObject
    Dim colAbrev As Long, colRamalAbrev As Long
    Dim i As Long, p As Long, r As Long
    Dim strAbrev As String, strRamalAbrev As String
    Dim proyectos() As String, ramales() As String
    Dim projCount As Long, ramalCount As Long
    Dim wsRamal As Worksheet
    Dim tbl As ListObject
    Dim colIdx As Long
    Dim basePrefix As String
    Dim formulaStr As String
    Dim shName As String, targetSheets As Variant, targetSh As Variant
    
    On Error GoTo ErrorHandler
    
    Set wsInicio = ThisWorkbook.Sheets("Inicio")
    Set tblConfig = wsInicio.ListObjects("Configuration")
    
    colAbrev = ObtenerColumna(tblConfig, Array("Abreviación", "Abreviacion", "Abrev", "Proyecto"))
    If colAbrev = 0 Then Exit Sub
    
    projCount = 0
    For i = 1 To tblConfig.ListRows.Count
        strAbrev = Trim(CStr(tblConfig.DataBodyRange(i, colAbrev).Value))
        If strAbrev <> "" Then
            projCount = projCount + 1
            ReDim Preserve proyectos(1 To projCount)
            proyectos(projCount) = strAbrev
        End If
    Next i
    
    If projCount <= 1 Then Exit Sub
    
    On Error Resume Next
    Set tblRamalConfig = wsInicio.ListObjects("Configuration_Ramal")
    On Error GoTo ErrorHandler
    
    ramalCount = 0
    If Not tblRamalConfig Is Nothing Then
        colRamalAbrev = ObtenerColumna(tblRamalConfig, Array("Abreviación", "Abreviacion", "Abrev"))
        If colRamalAbrev > 0 Then
            For r = 1 To tblRamalConfig.ListRows.Count
                strRamalAbrev = Trim(CStr(tblRamalConfig.DataBodyRange(r, colRamalAbrev).Value))
                If strRamalAbrev <> "" Then
                    ramalCount = ramalCount + 1
                    ReDim Preserve ramales(1 To ramalCount)
                    ramales(ramalCount) = strRamalAbrev
                End If
            Next r
        End If
    End If
    
    If ramalCount = 0 Then
        ramalCount = 1
        ReDim ramales(1 To 1)
        ramales(1) = "RAMAL"
    End If
    
    For r = 1 To ramalCount
        strRamalAbrev = ramales(r)
        targetSheets = Array("ANE. - " & strRamalAbrev, "ANE. - VAR. - " & strRamalAbrev)
        
        For Each targetSh In targetSheets
            shName = CStr(targetSh)
            If SheetExists(shName) Then
                Set wsRamal = ThisWorkbook.Sheets(shName)
                
                For Each tbl In wsRamal.ListObjects
                    colIdx = ObtenerColumnaTabla(tbl, "Cantidad a ejecutar")
                    
                    If colIdx > 0 Then
                        basePrefix = Replace(Replace(tbl.Name, strRamalAbrev, ""), "COMUNIDAD", "")
                        basePrefix = Replace(basePrefix, "RAMAL", "")
                        
                        If Right(basePrefix, 1) <> "_" Then basePrefix = basePrefix & "_"
                        
                        formulaStr = "=SUMA("
                        For p = 1 To projCount
                            formulaStr = formulaStr & basePrefix & proyectos(p) & "[@[Cantidad a ejecutar]]"
                            If p < projCount Then formulaStr = formulaStr & ","
                        Next p
                        formulaStr = formulaStr & ")"
                        
                        On Error Resume Next
                        If Not tbl.DataBodyRange Is Nothing Then
                            tbl.ListColumns(colIdx).DataBodyRange.FormulaLocal = formulaStr
                        End If
                        On Error GoTo ErrorHandler
                    End If
                Next tbl
            End If
        Next targetSh
    Next r

ErrorHandler:
    On Error GoTo 0
End Sub





