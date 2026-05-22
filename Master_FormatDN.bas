' =============================================================================
' Module: Master_FormatDN_And_Save_Ultimate
' Version: 2.0
' Changes:
'   [Req 1] Added RangeToHTML helper function + full Outlook draft logic
'   [Req 2] Implemented On Error GoTo ErrorHandler with guaranteed cleanup
'   [Req 3] All Integer declarations changed to Long
'   [Req 4] ActiveSheet replaced with explicit ThisWorkbook.Sheets("YourSheetName")
'   [Req 5] All business placeholders ([Company_Name], emails, paths) preserved
' =============================================================================

Option Explicit

' =============================================================================
' MAIN SUBROUTINE
' =============================================================================
Sub Master_FormatDN_And_Save_Ultimate()

    ' ------------------------------------------
    ' [Req 2] Error handler label declared at top
    ' ------------------------------------------
    On Error GoTo ErrorHandler

    ' ------------------------------------------
    ' [Req 4] Explicit worksheet reference instead of ActiveSheet
    ' ------------------------------------------
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("YourSheetName") ' <-- Change "YourSheetName" to your actual sheet name

    ' ==========================================
    ' VARIABLES DECLARATION
    ' ==========================================
    Dim lastRowVal As Long
    Dim i As Long          ' [Req 3] Changed from Integer to Long
    Dim r As Long
    Dim hasError As Boolean, rowError As String
    Dim expectedHeaders As Variant, headerError As String
    Dim validCustomers As String, firstCustomerCode As String
    Dim currentShipCode As String, docDateCheck As Variant
    Dim docYear As Long    ' [Req 3] Changed from Integer to Long
    Dim qty As Variant

    Dim lastRowData As Long, lastRow As Long, splitRow As Long
    Dim docDate As Date
    Dim dateStr1 As String, dateStr2 As String
    Dim tCode As String, poNumber As String
    Dim headerColor As Long, headerColor2 As Long
    Dim txt As String, txt2 As String
    Dim isMergeIJK As Boolean, isSepDate As Boolean, isTwoRows As Boolean
    Dim targetRange As Range
    Dim invNumber As String, engMonths As Variant

    ' Date Variables
    Dim yyyy As String, yy As String, m As String, mm As String, mmm As String, dd As String
    Dim dd_mmm_yy As String, dd_mmm_yyyy As String

    ' Path and Email Variables
    Dim basePath As String, fileName As String
    Dim currentPath As String, currentFile As String
    Dim pathArr() As String, buildPath As String
    Dim fullSavePath As String
    Dim copyNum As Long    ' [Req 3] Changed from Integer to Long
    Dim attachBasePath As String, attachPath As String
    Dim isInvEmpty As Boolean
    Dim timeSelection As String

    ' Variable for paragraph indentation (Space 7 times)
    Dim sp7 As String: sp7 = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"

    Application.ScreenUpdating = False

    ' =========================================================
    ' PART 1: DATA VALIDATION
    ' =========================================================
    hasError = False
    ws.Columns("L:N").Clear ' Clear old errors and hyperlinks

    expectedHeaders = Array("Delivery Note NO", "INVOICE NO", "Document Date", "Ship to Code", "Ship to Name", _
                            "DN item", "PURCHASE ORDER", "SO", "Material Code", "Material Description", "QTY")

    For i = 0 To UBound(expectedHeaders)
        If Trim(UCase(ws.Cells(1, i + 1).Value)) <> Trim(UCase(expectedHeaders(i))) Then
            headerError = headerError & "Col " & (i + 1) & " must be '" & expectedHeaders(i) & "' | "
            hasError = True
        End If
    Next i

    If ws.Cells(1, 12).Value <> "" And ws.Cells(1, 12).Value <> "Error Description" Then
        headerError = headerError & "Extra columns found "
        hasError = True
    End If

    If hasError Then
        ws.Cells(1, 12).Value = "HEADER ERROR: " & headerError
        Call FormatErrorCell(ws.Cells(1, 12))
        MsgBox "Invalid column headers! Process aborted.", vbCritical, "Validation Error"
        GoTo Cleanup
    Else
        ws.Cells(1, 12).Value = "Error Description"
    End If

    lastRowVal = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRowVal < 2 Then
        MsgBox "No data found in the file!", vbExclamation, "No Data"
        GoTo Cleanup
    End If

    validCustomers = "|CUST-01|CUST-02|CUST-03|CUST-04|CUST-05|CUST-06|CUST-07|CUST-08|CUST-09|CUST-10|CUST-11|CUST-12|CUST-13|CUST-14|"
    firstCustomerCode = ""

    For r = 2 To lastRowVal
        rowError = ""
        currentShipCode = Trim(ws.Cells(r, 4).Value)

        If currentShipCode = "" Then
            rowError = rowError & "Missing Ship to Code, "
        Else
            If InStr(1, validCustomers, "|" & currentShipCode & "|", vbTextCompare) = 0 Then
                rowError = rowError & "Customer code (" & currentShipCode & ") not in system, "
            End If

            If firstCustomerCode = "" Then
                firstCustomerCode = currentShipCode
            ElseIf currentShipCode <> firstCustomerCode Then
                rowError = rowError & "Mixed customer codes, "
            End If
        End If

        docDateCheck = ws.Cells(r, 3).Value
        If Not IsDate(docDateCheck) And docDateCheck <> "" Then
            rowError = rowError & "Invalid date, "
        ElseIf IsDate(docDateCheck) Then
            docYear = Year(docDateCheck)
            If docYear < 2020 Or docYear > 2030 Then
                rowError = rowError & "Abnormal year, "
            End If
        End If

        qty = ws.Cells(r, 11).Value
        If Not IsNumeric(qty) And qty <> "" Then
            rowError = rowError & "QTY is not a number, "
        ElseIf IsNumeric(qty) Then
            If qty <= 0 Then
                rowError = rowError & "QTY <= 0, "
            End If
        End If

        If Trim(ws.Cells(r, 1).Value) = "" Then
            rowError = rowError & "Empty DN, "
        End If

        If rowError <> "" Then
            ws.Cells(r, 12).Value = Left(rowError, Len(rowError) - 2)
            Call FormatErrorCell(ws.Cells(r, 12))
            hasError = True
        End If
    Next r

    If hasError Then
        MsgBox "Errors found! Please check column L.", vbCritical, "Validation Failed"
        GoTo Cleanup
    End If

    ' =========================================================
    ' PART 2: FORMATTING
    ' =========================================================
    ws.Columns("L").Clear
    engMonths = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

    ws.Rows("1:2").Insert Shift:=xlDown
    lastRowData = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If ws.Cells(ws.Rows.Count, "K").End(xlUp).Row > lastRowData Then
        lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
    Else
        lastRow = lastRowData + 1
    End If

    ' Check if Invoice NO column is empty
    isInvEmpty = True
    For r = 4 To (lastRow - 1)
        If Trim(ws.Cells(r, 2).Value) <> "" Then
            isInvEmpty = False
            Exit For
        End If
    Next r

    ws.Range("A" & lastRow & ":I" & lastRow).Clear
    ws.Range("J" & lastRow).Value = "Total"
    ws.Range("J" & lastRow).Font.Bold = True
    ws.Range("K" & lastRow).Formula = "=SUM(K4:K" & (lastRow - 1) & ")"
    ws.Range("K" & lastRow).Font.Bold = True

    On Error Resume Next
    docDate = CDate(ws.Range("C4").Value)
    On Error GoTo ErrorHandler ' [Req 2] Resume the main error handler after inline resume

    dateStr1 = Format(docDate, "dd-Mmm-yy")
    dateStr2 = Format(docDate, "dd-Mmm-yyyy")
    tCode = Trim(ws.Range("D4").Value)
    poNumber = Trim(ws.Range("G4").Value)
    invNumber = Trim(ws.Range("B4").Value)

    yyyy = Format(docDate, "yyyy")
    yy = Format(docDate, "yy")
    m = Format(docDate, "m")
    mm = Format(docDate, "mm")
    mmm = engMonths(Month(docDate) - 1)
    dd = Format(docDate, "dd")
    dd_mmm_yy = dd & "-" & mmm & "-" & yy
    dd_mmm_yyyy = dd & "-" & mmm & "-" & yyyy

    isMergeIJK = False: isSepDate = False: isTwoRows = False
    basePath = "": fileName = "": attachBasePath = ""

    ' =========================================================
    ' Update Save As Path and File Name Formatting
    ' =========================================================
    Select Case tCode
        Case "CUST-02"
            If InStr(1, poNumber, "/", vbTextCompare) > 0 Then
                txt = "[Company_Name] [Location_1] time : 13.30 PM (SFT)": headerColor = RGB(181, 230, 162): isSepDate = True
                basePath = "C:\Users\[Username]\Documents\[Shipment_Folder]\[YYYY]\[Location_1]\[M].[MMM]"
                fileName = "Shipment [Company_Name] [Location_1] [DD-MMM-YY] Time 13.30 PM. Rev.00"
                attachBasePath = "C:\Users\[Username]\Documents\[DN_Folder]\[YYYY]\DN\[Location_1]\[M].[MMM]\[DD]-[MMM]"
            Else
                Dim userChoice As VbMsgBoxResult
                userChoice = MsgBox("Please select a time period for [Company_Name] [Location_2]" & vbCrLf & vbCrLf & _
                                    "Click [Yes] = 11.00 AM" & vbCrLf & _
                                    "Click [No] = 23.00 PM", _
                                    vbYesNo + vbQuestion, "Choose a time period")

                If userChoice = vbYes Then
                    timeSelection = "11.00 AM"
                Else
                    timeSelection = "23.00 PM"
                End If

                txt = "[Company_Name]-[Location_2] time " & timeSelection & ".(SFT)": headerColor = RGB(255, 255, 0): isSepDate = True
                basePath = "C:\Users\[Username]\Documents\[Shipment_Folder]\[YYYY]\[Location_2]\[M].[MMM]"
                fileName = "[DD-MMM-YY] time " & timeSelection
                attachBasePath = "C:\Users\[Username]\Documents\[DN_Folder]\[YYYY]\DN\[Location_2]\[M].[MMM]\[DD]-[MMM]"
            End If

        Case "CUST-13"
            txt = "[Company_Code] " & dateStr1 & " 09.30 AM.(Milk-run)": headerColor = RGB(202, 237, 251)
            txt2 = "[Company_Code] " & dateStr1 & " 21.30 PM.(Milk-run)": headerColor2 = RGB(218, 242, 208): isTwoRows = True
            basePath = "C:\Users\[Username]\Documents\[Shipment_Folder]\[YYYY]\[M].[MMM]'[YY]"
            fileName = "Shipment [Company_Name] [DD-MMM-YY]"
            attachBasePath = "C:\Users\[Username]\Documents\[DN_Folder]\[YYYY]\[M].[MMM]'[YY]\[DD]-[MMM]'[YY]"

        Case "CUST-01"
            txt = "[Company_Name] ([Location_Code]) time: 14.30 PM.(SFT)": headerColor = RGB(68, 179, 225): isSepDate = True
            basePath = "C:\Users\[Username]\Documents\[Shipment_Folder]\[YYYY]\[M].[MMM]"
            fileName = "Shipment [DD-MMM-YY] [Company_Name] Time 14.30 PM Rev.00"
            attachBasePath = "C:\Users\[Username]\Documents\[DN_Folder]\[YYYY]\[M].[MMM]\[DD]-[MMM]"

        Case "CUST-03"
            txt = "[Company_Name] [Location_Code] ship by Air on ETD " & dateStr1: headerColor = RGB(68, 179, 225): isMergeIJK = True
            basePath = "C:\Users\[Username]\Documents\[Shipment_Folder]\[YYYY]\[MM].[MMM]'[YY]"
            fileName = "[DD-MMM-YY] ship by Air SO -- Inv[INV]"
            attachBasePath = "C:\Users\[Username]\Documents\[DN_Folder]\[YYYY]\[MM].[MMM]'[YY]"

        Case "CUST-04"
            txt = "[Company_Name] [Location_Code] shipment  by Air on ETD " & dateStr2 & ".": headerColor = RGB(77, 147, 217): isMergeIJK = True
            basePath = "C:\Users\[Username]\Documents\[Shipment_Folder]\[YYYY]\[M].[MMM]"
            fileName = "Shipment [Company_Name] [Location_Code] [DD-MMM-YY] (Air) SO -- Inv.[INV]"
            attachBasePath = "C:\Users\[Username]\Documents\[DN_Folder]\[YYYY]\[M].[MMM]"

        Case "CUST-05"
            txt = "[Company_Name] [Location_Code] line 69 on ETD " & dateStr2: headerColor = RGB(15, 158, 213): isMergeIJK = True
            basePath = "C:\Users\[Username]\Documents\[Shipment_Folder]\[YYYY]\[M].[MMM]"
            fileName = "Shipment [Company_Name] [Location_Code] [DD-MMM-YY] Line -- Rev.00"
            attachBasePath = "C:\Users\[Username]\Documents\[DN_Folder]\[YYYY]\[M].[MMM]"

        Case "CUST-06"
            txt = "[Company_Name] [Country_Name] ETD " & dateStr2: headerColor = RGB(202, 237, 251): isMergeIJK = True
            basePath = "C:\Users\[Username]\Documents\[Shipment_Folder]\[YYYY]"
            fileName = "Shipment [Company_Name] [Country_Code] [DD-MMM-YY]"
            attachBasePath = "C:\Users\[Username]\Documents\[DN_Folder]\[YYYY]"

        Case "CUST-07"
            txt = "[Company_Code] [Location_1] " & dateStr1 & " (Milk-run) 13.50 PM.": headerColor = RGB(202, 237, 251)
            basePath = "C:\Users\[Username]\Documents\[Shipment_Folder]\[YYYY]\[Location_1]\[MM].[MMM]'[YY]"
            fileName = "[DD-MMM-YYYY] [Location_1] Rev.00"
            attachBasePath = "C:\Users\[Username]\Documents\[DN_Folder]\[YYYY]\2.[Location_1]\[MM]. [MMM]'[YY]"

        Case "CUST-08"
            txt = "[Company_Code] " & dateStr1 & " Time 14.30 PM.(Milk-run)": headerColor = RGB(192, 230, 245)
            basePath = "C:\Users\[Username]\Documents\[Shipment_Folder]\[YYYY]\[MM].[MMM]"
            fileName = "[DD-MMM-YYYY]_Rev.00"
            attachBasePath = "C:\Users\[Username]\Documents\[DN_Folder]\[YYYY]\[MM]. [MMM]'[YY]\[DD]-[MMM]"

        Case "CUST-09"
            txt = "[Company_Name] " & dateStr1 & " (By SEA)": headerColor = RGB(202, 237, 251)
            basePath = "C:\Users\[Username]\Documents\[Shipment_Folder]\[YYYY]\[MM].[MMM]"
            fileName = "Shipment [Company_Name] [DD-MMM-YY] (Sea)_Rev.00"
            attachBasePath = "C:\Users\[Username]\Documents\[DN_Folder]\[YYYY]\[MM].[MMM]"

        Case "CUST-10"
            txt = "[Company_Name]-Assy " & dateStr1 & " (By SEA)": headerColor = RGB(202, 237, 251)
            basePath = "C:\Users\[Username]\Documents\[Shipment_Folder]\[YYYY]\[MM].[MMM]'[YY]"
            fileName = "Shipment [Company_Name] [DD-MMM-YY] (Sea_TICA)"
            attachBasePath = "C:\Users\[Username]\Documents\[DN_Folder]\[YYYY]\[MM].[MMM]'[YY]"

        Case "CUST-11"
            txt = "[Company_Name]-[Location_Code] by air (DHL) shipment on ETD " & dateStr1: headerColor = RGB(242, 206, 239): isMergeIJK = True
            basePath = "C:\Users\[Username]\Documents\[Shipment_Folder]\[YYYY]\[M].[MMM]"
            fileName = "Shipment [DD-MMM-YY] Rev.00"
            attachBasePath = "C:\Users\[Username]\Documents\[DN_Folder]\[YYYY]\[M].[MMM]-[YY]\[DD]-[MMM]"

        Case "CUST-12"
            txt = "[Company_Name]-[Location_Code] by Air on ETD " & dateStr2: headerColor = RGB(68, 179, 225): isMergeIJK = True
            basePath = "C:\Users\[Username]\Documents\[Shipment_Folder]\[YYYY]\[M].[MMM]"
            fileName = "Shipment [DD-MMM-YY] (Air)"
            attachBasePath = "C:\Users\[Username]\Documents\[DN_Folder]\FY[YYYY]\[M].[MMM]\[DD]-[MMM]"

        Case "CUST-14"
            txt = "[Company_Code] [Location_2]  " & dateStr1 & " 10.00 AM.(Milk-run)": headerColor = RGB(202, 237, 251)
            basePath = "C:\Users\[Username]\Documents\[Shipment_Folder]\[YYYY]\[Location_2]\[MM].[MMM]'[YY]"
            fileName = "[DD-MMM-YYYY] [Location_2] Rev.00"
            attachBasePath = "C:\Users\[Username]\Documents\[DN_Folder]\[YYYY]\1. [Location_2]\[MM].[MMM]'[YY]"

        Case Else
            txt = "T-Code Not Found": headerColor = RGB(255, 255, 255)
    End Select

    ' =========================================================
    ' Special formatting for split table (CUST-13)
    ' =========================================================
    If isTwoRows Then
        Dim lastDN As String
        lastDN = Trim(ws.Cells(lastRow - 1, 1).Value)
        splitRow = 4
        For r = 4 To lastRow - 1
            If Trim(ws.Cells(r, 1).Value) = lastDN Then
                splitRow = r
                Exit For
            End If
        Next r

        If splitRow > 4 Then
            ws.Rows(splitRow & ":" & splitRow + 4).Insert Shift:=xlDown
            ws.Rows(splitRow & ":" & splitRow + 4).Clear

            ' --- Format top table (09.30) ---
            ws.Range("J2:K2").Merge: ws.Range("J2").Value = txt
            ws.Range("J2:K2").Interior.Color = headerColor
            ws.Range("J2:K2").Font.Bold = True
            ws.Range("J2:K2").HorizontalAlignment = xlCenter
            ws.Range("J2:K2").Borders.LineStyle = xlContinuous

            ws.Range("J" & splitRow).Value = "Total"
            ws.Range("J" & splitRow).Font.Bold = True
            ws.Range("J" & splitRow).HorizontalAlignment = xlCenter
            ws.Range("K" & splitRow).Formula = "=SUM(K4:K" & (splitRow - 1) & ")"
            ws.Range("K" & splitRow).Font.Bold = True
            ws.Range("J" & splitRow & ":K" & splitRow).Interior.Color = headerColor
            ws.Range("J" & splitRow & ":K" & splitRow).Borders.LineStyle = xlContinuous

            ws.Range("A3:K" & (splitRow - 1)).Borders.LineStyle = xlContinuous

            ' --- Format bottom table (21.30) ---
            ws.Range("J" & (splitRow + 3) & ":K" & (splitRow + 3)).Merge
            ws.Range("J" & (splitRow + 3)).Value = txt2
            ws.Range("J" & (splitRow + 3) & ":K" & (splitRow + 3)).Interior.Color = headerColor2
            ws.Range("J" & (splitRow + 3) & ":K" & (splitRow + 3)).Font.Bold = True
            ws.Range("J" & (splitRow + 3) & ":K" & (splitRow + 3)).HorizontalAlignment = xlCenter
            ws.Range("J" & (splitRow + 3) & ":K" & (splitRow + 3)).Borders.LineStyle = xlContinuous

            ws.Range("A3:K3").Copy Destination:=ws.Range("A" & (splitRow + 4))

            Dim newLastRow As Long
            newLastRow = lastRow + 5

            ws.Range("A" & newLastRow & ":K" & newLastRow).Clear

            ws.Range("J" & newLastRow).Value = "Total"
            ws.Range("J" & newLastRow).Font.Bold = True
            ws.Range("J" & newLastRow).HorizontalAlignment = xlCenter
            ws.Range("K" & newLastRow).Formula = "=SUM(K" & (splitRow + 5) & ":K" & (newLastRow - 1) & ")"
            ws.Range("K" & newLastRow).Font.Bold = True
            ws.Range("J" & newLastRow & ":K" & newLastRow).Interior.Color = headerColor2
            ws.Range("J" & newLastRow & ":K" & newLastRow).Borders.LineStyle = xlContinuous

            ws.Range("A" & (splitRow + 4) & ":K" & (newLastRow - 1)).Borders.LineStyle = xlContinuous

            ws.Rows((newLastRow + 1) & ":" & (newLastRow + 10)).Clear
            lastRow = newLastRow
        Else
            ws.Range("J2:K2").Merge: ws.Range("J2").Value = txt
            ws.Range("J2:K2").Interior.Color = headerColor
            ws.Range("J" & lastRow & ":K" & lastRow).Interior.Color = headerColor
            ws.Range("A3:K" & (lastRow - 1)).Borders.LineStyle = xlContinuous
            ws.Range("J" & lastRow & ":K" & lastRow).Borders.LineStyle = xlContinuous
        End If

    ' =========================================================
    ' Standard Formatting
    ' =========================================================
    ElseIf isSepDate Then
        ws.Range("I2").Value = dateStr1
        ws.Range("J2:K2").Merge: ws.Range("J2").Value = txt: ws.Range("I2:K2").Interior.Color = headerColor
        ws.Range("J" & lastRow & ":K" & lastRow).Interior.Color = headerColor
        Set targetRange = ws.Range("I2:K2")
    ElseIf isMergeIJK Then
        ws.Range("I2:K2").Merge: ws.Range("I2").Value = txt: ws.Range("I2:K2").Interior.Color = headerColor
        ws.Range("J" & lastRow & ":K" & lastRow).Interior.Color = headerColor
        Set targetRange = ws.Range("I2:K2")
        targetRange.Font.Name = "Calibri": ActiveWindow.DisplayGridlines = False
    Else
        ws.Range("J2:K2").Merge: ws.Range("J2").Value = txt: ws.Range("J2:K2").Interior.Color = headerColor
        ws.Range("J" & lastRow & ":K" & lastRow).Interior.Color = headerColor
        Set targetRange = ws.Range("J2:K2")
    End If

    If Not isTwoRows Then
        If Not targetRange Is Nothing Then
            With targetRange
                .Font.Bold = True: .HorizontalAlignment = xlCenter: .Borders.LineStyle = xlContinuous
            End With
        End If
        ws.Range("A3:K" & (lastRow - 1)).Borders.LineStyle = xlContinuous
        ws.Range("J" & lastRow & ":K" & lastRow).Borders.LineStyle = xlContinuous
    End If

    With ws.Range("A3:K" & lastRow)
        .Font.Name = "Calibri"
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
    End With

    If isTwoRows And splitRow > 4 Then
        ws.Range("C4:C" & (splitRow - 1)).NumberFormat = "[$-en-US]dd-Mmm-yy"
        ws.Range("C" & (splitRow + 5) & ":C" & (lastRow - 1)).NumberFormat = "[$-en-US]dd-Mmm-yy"
    Else
        ws.Range("C4:C" & (lastRow - 1)).NumberFormat = "[$-en-US]dd-Mmm-yy"
    End If

    ws.Columns("A:K").EntireColumn.AutoFit

    If isInvEmpty Then
        ws.Columns("B").Delete Shift:=xlToLeft
    End If

    ' =========================================================
    ' PART 3: Create Folder and Save As
    ' =========================================================
    If basePath <> "" And fileName <> "" Then
        currentPath = Replace(basePath, "[YYYY]", yyyy)
        currentPath = Replace(currentPath, "[YY]", yy)
        currentPath = Replace(currentPath, "[MM]", mm)
        currentPath = Replace(currentPath, "[M]", m)
        currentPath = Replace(currentPath, "[MMM]", mmm)

        currentFile = Replace(fileName, "[DD-MMM-YYYY]", dd_mmm_yyyy)
        currentFile = Replace(currentFile, "[DD-MMM-YY]", dd_mmm_yy)
        currentFile = Replace(currentFile, "[INV]", invNumber)

        If Right(currentPath, 1) = "\" Then currentPath = Left(currentPath, Len(currentPath) - 1)

        pathArr = Split(currentPath, "\"): buildPath = pathArr(0)
        On Error Resume Next
        For i = 1 To UBound(pathArr)
            buildPath = buildPath & "\" & pathArr(i)
            If Dir(buildPath, vbDirectory) = "" Then MkDir buildPath
        Next i
        On Error GoTo ErrorHandler ' [Req 2] Resume main error handler

        fullSavePath = currentPath & "\" & currentFile & ".xlsx"
        copyNum = 0
        While Dir(fullSavePath) <> ""
            fullSavePath = currentPath & "\" & currentFile & " - Copy" & Format(copyNum, "00") & ".xlsx"
            copyNum = copyNum + 1
        Wend

        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs fileName:=fullSavePath, FileFormat:=xlOpenXMLWorkbook
        Application.DisplayAlerts = True

        ws.Range("M1").Value = "[1] Data Folder"
        ws.Hyperlinks.Add Anchor:=ws.Range("M1"), Address:=currentPath

        If attachBasePath <> "" Then
            attachPath = Replace(attachBasePath, "[YYYY]", yyyy)
            attachPath = Replace(attachPath, "[YY]", yy)
            attachPath = Replace(attachPath, "[MM]", mm)
            attachPath = Replace(attachPath, "[M]", m)
            attachPath = Replace(attachPath, "[MMM]", mmm)
            attachPath = Replace(attachPath, "[DD]", dd)

            Dim isFolderExist As Boolean
            isFolderExist = False
            On Error Resume Next
            If Dir(attachPath, vbDirectory) <> "" Then isFolderExist = True
            On Error GoTo ErrorHandler ' [Req 2] Resume main error handler

            If isFolderExist Then
                ws.Range("M2").Value = "[2] Attachment Folder"
                ws.Range("M2").Font.Color = RGB(0, 112, 192)
                ws.Hyperlinks.Add Anchor:=ws.Range("M2"), Address:=attachPath
            Else
                ws.Range("M2").Value = "Attachment folder not found! (Click to force open)"
                ws.Range("M2").Font.Color = RGB(255, 0, 0)
                ws.Hyperlinks.Add Anchor:=ws.Range("M2"), Address:=attachPath

                Dim parentPath As String
                parentPath = Left(attachPath, InStrRev(attachPath, "\") - 1)
                ws.Range("N2").Value = "Backup folder (Go back 1 level)"
                ws.Range("N2").Font.Color = RGB(255, 165, 0)
                ws.Hyperlinks.Add Anchor:=ws.Range("N2"), Address:=parentPath
            End If
        End If
        ws.Columns("M:N").AutoFit

        ' =========================================================
        ' PART 4: Create Draft Email
        ' [Req 1] Full Outlook draft logic with RangeToHTML helper
        ' =========================================================
        Dim isDraftMail As Boolean: isDraftMail = False
        Dim mTo As String, mCc As String, mSub As String, mBody As String

        ' --------- CUST-13 ---------
        If tCode = "CUST-13" Then
            isDraftMail = True
            mTo = "[recipient1@example.com]; [recipient2@example.com]"
            mCc = "[cc_person1@example.com]; [cc_person2@example.com]"
            mSub = "Delivery note [Company_Name] on " & dd_mmm_yy & " Time 09.30 AM., 21.30 PM. (Milk Run) Rev.00"

            Dim dictAM As Object: Set dictAM = CreateObject("Scripting.Dictionary")
            Dim dictPM As Object: Set dictPM = CreateObject("Scripting.Dictionary")

            If splitRow > 4 Then
                For r = 4 To splitRow - 1
                    If Trim(ws.Cells(r, 1).Value) <> "" Then dictAM(Trim(ws.Cells(r, 1).Value)) = 1
                Next r
                For r = splitRow + 5 To lastRow - 1
                    If Trim(ws.Cells(r, 1).Value) <> "" Then dictPM(Trim(ws.Cells(r, 1).Value)) = 1
                Next r
            Else
                For r = 4 To lastRow - 1
                    If Trim(ws.Cells(r, 1).Value) <> "" Then dictAM(Trim(ws.Cells(r, 1).Value)) = 1
                Next r
            End If

            Dim dnListAM As String, dnListPM As String
            Dim dKey As Variant
            For Each dKey In dictAM.keys
                dnListAM = dnListAM & sp7 & sp7 & "DN " & dKey & "<br>"
            Next dKey
            For Each dKey In dictPM.keys
                dnListPM = dnListPM & sp7 & sp7 & "DN " & dKey & "<br>"
            Next dKey

            mBody = "<span style='font-family:Calibri; font-size:10.5pt;'><b>Dear</b>" & _
                    "<b> [Contact_Name_1] & [Client_Team]</b><br>" & _
                    "<b>CC. [Internal_Team],</b><br>" & _
                    sp7 & "We would like to confirm the delivery plan for <b>" & dd_mmm_yy & "</b> at 09:30 AM and 21:30 PM. Please refer to the delivery note below for shipment details; [Our_Company_Name] can support this plan according to the PO balance in the attached file. Rev.00<br><br>" & _
                    sp7 & "<u>Time 09.30 AM.&nbsp;&nbsp;&nbsp;MilkRun</u><br>" & dnListAM & _
                    sp7 & "<u>Time 21.30 PM.&nbsp;&nbsp;&nbsp;MilkRun</u><br>" & dnListPM & "<br>[[TABLE_HERE]]<br><br></span>"

        ' --------- CUST-02 Location 2 ---------
        ElseIf tCode = "CUST-02" And InStr(1, poNumber, "/", vbTextCompare) = 0 Then
            isDraftMail = True
            mTo = "[recipient_to@example.com]"
            mCc = "[recipient_cc@example.com]"
            mSub = "Delivery note on " & dd_mmm_yyyy & "  [Company_Name] [Location_2] Time " & timeSelection & " Rev.00."
            mBody = "<span style='font-family:Calibri; font-size:10.5pt;'><b>Dear [Contact_Name_2] and [Internal_Team].</b><br>" & _
                    sp7 & "Please see Delivery note on <b>" & dd_mmm_yyyy & "</b> for [Company_Name] [Location_2] Time " & timeSelection & " Rev.00 as attached file.<br><br>[[TABLE_HERE]]<br><br></span>"

        ' --------- CUST-02 Location 1 ---------
        ElseIf tCode = "CUST-02" And InStr(1, poNumber, "/", vbTextCompare) > 0 Then
            isDraftMail = True
            mTo = "[recipient_to@example.com]"
            mCc = "[recipient_cc@example.com]"
            mSub = "Delivery note on " & dd_mmm_yyyy & "  [Company_Name] [Location_1] Time 13:30 PM Rev.00."
            mBody = "<span style='font-family:Calibri; font-size:10.5pt;'><b>Dear [Contact_Name_3]</b><br>" & _
                    sp7 & "Please see Delivery note on <b>" & dd_mmm_yyyy & "</b> for [Company_Name] [Location_1] ETD 13:30 PM Rev.00 as attached file.<br><br>[[TABLE_HERE]]<br><br></span>"

        ' --------- CUST-10 ---------
        ElseIf tCode = "CUST-10" Then
            isDraftMail = True
            mTo = "[recipient_to@example.com]"
            mCc = "[recipient_cc@example.com]"
            mSub = "Shipment [Company_Name] ASSY on " & dd_mmm_yy
            mBody = "<span style='font-family:Calibri; font-size:10.5pt;'><b>Dear [Contact_Name_1]</b><br><b>CC, [Internal_Team]</b><br>" & _
                    sp7 & "Please see Delivery Note shipment for [Company_Name] on <b>" & dd_mmm_yy & "</b> Rev.00 as attached file,<br><br>[[TABLE_HERE]]<br><br></span>"

        ' --------- CUST-09 ---------
        ElseIf tCode = "CUST-09" Then
            isDraftMail = True
            mTo = "[recipient_to@example.com]"
            mCc = "[recipient_cc@example.com]"
            mSub = "Shipment [Company_Name]  on " & dd_mmm_yy & "  by  Sea"
            mBody = "<span style='font-family:Calibri; font-size:10.5pt;'><b>Dear [Contact_Name_1]</b><br><b>CC, [Internal_Team]</b><br>" & _
                    sp7 & "Please see Delivery Note shipment for [Company_Name] on <b>" & dd_mmm_yy & "</b> Rev.00 as attached file,<br><br>[[TABLE_HERE]]<br><br></span>"

        ' --------- CUST-08 ---------
        ElseIf tCode = "CUST-08" Then
            isDraftMail = True
            mTo = "[recipient_to@example.com]"
            mCc = "[recipient_cc@example.com]"
            mSub = "Delivery note on " & dd_mmm_yyyy & " [Company_Name] [Location_1] Time 14.30 PM. Rev.00"
            mBody = "<span style='font-family:Calibri; font-size:10.5pt;'><b>Dear [Contact_Name_4]</b><br><b>CC, [Internal_Team]</b><br>" & _
                    sp7 & "Delivery Note for [Company_Name] [Location_1] on <b>" & dd_mmm_yyyy & "</b> Time 14.30 PM. Rev.00 per attached.<br><br>[[TABLE_HERE]]<br><br></span>"

        ' --------- CUST-07 ---------
        ElseIf tCode = "CUST-07" Then
            isDraftMail = True
            mTo = "[recipient_to@example.com]"
            mCc = "[recipient_cc@example.com]"
            mSub = "[Company_Code] ([Location_3]) Delivery Note  on ETD " & dd_mmm_yyyy & " Time 13.50 PM Rev.00"
            mBody = "<span style='font-family:Calibri; font-size:10.5pt;'><b>Dear [Contact_Name_5]</b><br><b>CC, [Internal_Team] and BOI team</b><br>" & _
                    sp7 & "Please see Delivery Note (DN) for [Company_Name] ([Company_Code] [Location_3]) on <b>" & dd_mmm_yyyy & "</b> (milk run) ETD 13.50 PM. Rev.00 as attached<br><br>[[TABLE_HERE]]<br><br></span>"

        ' --------- CUST-14 ---------
        ElseIf tCode = "CUST-14" Then
            isDraftMail = True
            mTo = "[recipient_to@example.com]"
            mCc = "[recipient_cc@example.com]"
            mSub = "Delivery note on " & dd_mmm_yyyy & " [Company_Name] ([Location_4]) Time 10.00 AM. Rev.00"
            mBody = "<span style='font-family:Calibri; font-size:10.5pt;'><b>Dear [Contact_Name_5]</b><br><b>CC, [Internal_Team]</b><br>" & _
                    sp7 & "Please see Delivery Note (DN) for [Company_Name] ([Location_4]) on <b>" & dd_mmm_yyyy & "</b> ETD 10.00 AM. Rev.00 as attached file.<br><br>[[TABLE_HERE]]<br><br></span>"
        End If

        ' =========================================================
        ' [Req 1] EXECUTE EMAIL GENERATION
        ' Uses RangeToHTML helper to embed the formatted table
        ' =========================================================
        If isDraftMail Then
            Dim OutApp As Object, OutMail As Object
            Dim tableHTML As String
            Dim lastCol As String

            ' Determine table range based on whether Invoice column was deleted
            If isInvEmpty Then
                lastCol = "J"
            Else
                lastCol = "K"
            End If

            ' [Req 1] Convert the formatted Excel range to an HTML table string
            tableHTML = RangeToHTML(ws.Range("A1:" & lastCol & lastRow))

            ' [Req 1] Replace the [[TABLE_HERE]] placeholder with the real HTML table
            mBody = Replace(mBody, "[[TABLE_HERE]]", tableHTML)

            ' [Req 1] Create Outlook instance (reuse existing if open)
            On Error Resume Next
            Set OutApp = GetObject(, "Outlook.Application")
            If OutApp Is Nothing Then Set OutApp = CreateObject("Outlook.Application")
            On Error GoTo ErrorHandler ' [Req 2] Resume main error handler

            If OutApp Is Nothing Then
                MsgBox "Could not launch Outlook. Email draft was skipped.", vbExclamation, "Outlook Not Available"
            Else
                ' [Req 1] Create and display the draft email
                Set OutMail = OutApp.CreateItem(0) ' 0 = olMailItem

                With OutMail
                    .To = mTo
                    .CC = mCc
                    .Subject = mSub
                    .HTMLBody = mBody   ' HTML body with embedded table
                    .Display                  ' Open draft window (do NOT .Send)
                End With

                Set OutMail = Nothing
                Set OutApp = Nothing
            End If

            Application.CutCopyMode = False
        End If

        MsgBox "Formatting, saving, and email drafting completed!", vbInformation, "Success"

    Else
        MsgBox "Formatting completed! (Save path not found)", vbExclamation, "Done (No Save Path)"
    End If

    ' =========================================================
    ' [Req 2] CLEANUP: Always runs on success OR error
    ' =========================================================
Cleanup:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.CutCopyMode = False
    Exit Sub

ErrorHandler:
    ' [Req 2] Show a descriptive error message
    MsgBox "An error occurred:" & vbCrLf & vbCrLf & _
           "Error #" & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           "The script has been stopped. Please check your data and try again.", _
           vbCritical, "Unexpected Error"
    Resume Cleanup ' [Req 2] Always restore Excel state even after error

End Sub


' =============================================================================
' [Req 1] HELPER FUNCTION: RangeToHTML
'
' Converts a given Excel Range into an HTML table string.
' This approach uses a temporary file via PublishObjects, which preserves
' Excel's own cell formatting (colors, borders, fonts) in the HTML output.
'
' Usage: tableHTML = RangeToHTML(ws.Range("A1:K20"))
' =============================================================================
Function RangeToHTML(rng As Range) As String

    Dim fso As Object
    Dim ts As Object
    Dim tempFilePath As String
    Dim htmlContent As String
    Dim startMarker As String
    Dim endMarker As String
    Dim startPos As Long
    Dim endPos As Long

    ' Create a unique temporary HTML file path
    tempFilePath = Environ("TEMP") & "\RangeToHTML_" & Format(Now, "yyyymmddhhmmss") & ".htm"

    ' Publish the range to a temporary HTML file using Excel's built-in publisher
    ' This preserves cell colors, borders, and font formatting
    With ThisWorkbook.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         FileName:=tempFilePath, _
         Sheet:=rng.Parent.Name, _
         Source:=rng.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    ' Read the generated HTML file
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(tempFilePath).OpenAsTextStream(1, -2) ' 1=ForReading, -2=TristateUseDefault
    htmlContent = ts.ReadAll
    ts.Close

    ' Extract only the table portion from the full HTML document.
    ' Excel wraps the table in a specific structure; we find the <table> tag.
    startMarker = "<table"
    endMarker = "</table>"

    startPos = InStr(1, htmlContent, startMarker, vbTextCompare)
    endPos = InStr(1, htmlContent, endMarker, vbTextCompare)

    If startPos > 0 And endPos > 0 Then
        RangeToHTML = Mid(htmlContent, startPos, (endPos - startPos) + Len(endMarker))
    Else
        ' Fallback: return the full HTML body if table markers not found
        RangeToHTML = htmlContent
    End If

    ' Clean up the temporary file
    On Error Resume Next
    fso.DeleteFile tempFilePath
    On Error GoTo 0

    Set ts = Nothing
    Set fso = Nothing

End Function


' =============================================================================
' HELPER SUBROUTINE: FormatErrorCell
' Formats a cell to display validation errors (red background, white bold text)
' =============================================================================
Sub FormatErrorCell(rng As Range)
    With rng
        .Interior.Color = RGB(255, 0, 0)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .EntireColumn.AutoFit
    End With
End Sub
