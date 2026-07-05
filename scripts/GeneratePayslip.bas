Sub GeneratePayslip()
    Dim payslipWs As Worksheet
    Dim payEntriesWs As Worksheet
    Dim targetWs As Worksheet
    Dim workerId As String
    Dim monthName As String
    Dim netPay As Double
    Dim netPayRaw As Variant
    Dim mainCategory As String
    Dim subCategory As String
    Dim paymentMode As String
    Dim role As String
    Dim entryRow As Long
    Dim payRowRefCell As Range
    Dim targetRow As Long
    Dim r As Long
    Dim exportPath As String
    Dim fileName As String

    Set payslipWs = ThisWorkbook.Sheets("Payslip")
    Set payEntriesWs = ThisWorkbook.Sheets("PayEntries")

    workerId = payslipWs.Range("B6").Value
    monthName = payslipWs.Range("B4").Value
    role = payslipWs.Range("B7").Value
    netPayRaw = payslipWs.Range("B16").Value
    paymentMode = payslipWs.Range("B17").Value

    If workerId = "" Or monthName = "" Or netPayRaw = "" Then
        MsgBox "No pay entry found for this Worker/Month. Fill in PayEntries first.", vbExclamation
        Exit Sub
    End If

    netPay = CDbl(netPayRaw)

    ' Find the matching PayEntries row (Worker ID + Month) to read/write Payments Row Ref
    entryRow = 0
    For r = 3 To 502
        If payEntriesWs.Cells(r, 1).Value = workerId And payEntriesWs.Cells(r, 3).Value = monthName Then
            entryRow = r
            Exit For
        End If
    Next r

    If entryRow = 0 Then
        MsgBox "No matching PayEntries row found. Fill in PayEntries first.", vbExclamation
        Exit Sub
    End If

    mainCategory = "Labour & Contractors"
    If role = "Labourer" Then
        subCategory = "Labour Wages"
    Else
        subCategory = "Contractor Payments"
    End If

    Set targetWs = ThisWorkbook.Sheets(monthName)
    Set payRowRefCell = payEntriesWs.Cells(entryRow, 9) ' Payments Row Ref column

    If payRowRefCell.Value <> "" Then
        ' Regenerating: update the existing Payments row in place
        Dim refParts() As String
        refParts = Split(payRowRefCell.Value, "K")
        targetRow = CLng(refParts(1))
    Else
        ' First generation: find the first blank row in the Payments table (rows 3-100)
        targetRow = 0
        For r = 3 To 100
            If targetWs.Cells(r, 11).Value = "" Then ' column K = Amount
                targetRow = r
                Exit For
            End If
        Next r
        If targetRow = 0 Then
            MsgBox "Payments table for " & monthName & " is full (rows 3-100 all used).", vbCritical
            Exit Sub
        End If
    End If

    targetWs.Cells(targetRow, 11).Value = netPay              ' Amount (K)
    targetWs.Cells(targetRow, 12).Value = mainCategory         ' Main Category (L)
    targetWs.Cells(targetRow, 13).Value = subCategory          ' Sub-Category (M)
    targetWs.Cells(targetRow, 14).Value = paymentMode          ' Payment Mode (N)
    targetWs.Cells(targetRow, 15).Value = "Payslip: " & payslipWs.Range("B3").Value & " - " & monthName ' Comments (O)

    payRowRefCell.Value = monthName & "!K" & targetRow

    fileName = "Payslip_" & workerId & "_" & monthName & ".pdf"
    exportPath = ThisWorkbook.Path & Application.PathSeparator & fileName
    payslipWs.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=exportPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

    MsgBox "Payslip generated and payment recorded in " & monthName & "!K" & targetRow & vbCrLf & _
           "PDF: " & exportPath, vbInformation
End Sub
