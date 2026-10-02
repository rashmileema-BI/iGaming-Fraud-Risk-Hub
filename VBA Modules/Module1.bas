Attribute VB_Name = "Module1"
Option Explicit

' ================= CONFIG =================
Private Const PAY_TABLE As String = "Payments"
Private Const C_TXN As String = "txn_id"
Private Const C_PLAYER As String = "player_id"
Private Const C_DATE As String = "txn_date"
Private Const C_TYPE As String = "txn_type"
Private Const C_AMT As String = "amount"
Private Const C_STATUS As String = "status"      ' set to "" if you have no status column
Private Const OK_STATUS As String = "|completed|success|successful|approved|"

Private Const P_TXN As String = "txn_id"
Private Const P_AMT As String = "amount"
Private Const P_FEE As String = "fee"

Private Const DEP_FEE_RATE As Double = 0.015
Private Const WD_FEE_FLAT As Double = 0.5
Private Const AMT_TOL As Double = 0.005
Private Const FEE_TOL As Double = 0.02

Private Const LARGE_TXN As Double = 1000         ' report threshold
Private Const VELOCITY_MIN As Long = 5           ' txns per player per day
' ==========================================

' ---------- helpers ----------
Private Function GetTable() As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If LCase(lo.Name) = LCase(PAY_TABLE) Then
                Set GetTable = lo
                Exit Function
            End If
        Next lo
    Next ws
    MsgBox "Table '" & PAY_TABLE & "' not found. Load the Payments query to a worksheet table first.", vbExclamation
End Function

Private Function Col(hdr As Variant, nm As String) As Long
    Dim i As Long
    If nm = "" Then Exit Function
    For i = 1 To UBound(hdr, 2)
        If LCase(Trim(CStr(hdr(1, i)))) = LCase(nm) Then
            Col = i
            Exit Function
        End If
    Next i
End Function

Private Function FreshSheet(nm As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = nm
    Else
        ws.Cells.Clear
    End If
    Set FreshSheet = ws
End Function

Private Function StatusOk(pv As Variant, i As Long, cStat As Long) As Boolean
    If cStat = 0 Then
        StatusOk = True
    Else
        StatusOk = InStr(1, OK_STATUS, "|" & LCase(Trim(CStr(pv(i, cStat)))) & "|", vbTextCompare) > 0
    End If
End Function

Private Function ExpectedFee(txnType As String, amt As Double) As Double
    If txnType = "deposit" Then
        ExpectedFee = Round(amt * DEP_FEE_RATE, 2)
    Else
        ExpectedFee = WD_FEE_FLAT
    End If
End Function

Private Sub AddRow(ByRef out As Variant, ByRef k As Long, cat As String, txn As String, _
                   ledAmt As Variant, pspAmt As Variant, pspFee As Variant, expFee As Variant, note As String)
    k = k + 1
    out(k, 1) = cat
    out(k, 2) = txn
    out(k, 3) = ledAmt
    out(k, 4) = pspAmt
    out(k, 5) = pspFee
    out(k, 6) = expFee
    out(k, 7) = note
End Sub

Private Sub StyleHeader(rng As Range)
    With rng
        .Font.Bold = True
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = vbWhite
    End With
End Sub

' ================= 1. PSP RECONCILIATION =================
Public Sub ReconcilePSP()
    Dim lo As ListObject: Set lo = GetTable()
    If lo Is Nothing Then Exit Sub

    Dim f As Variant
    f = Application.GetOpenFilename("CSV files (*.csv),*.csv", , "Select PSP settlement file")
    If VarType(f) = vbBoolean Then Exit Sub

    Application.ScreenUpdating = False

    ' import the settlement file into PSP_Raw
    Dim wsRaw As Worksheet: Set wsRaw = FreshSheet("PSP_Raw")
    Dim wb As Workbook: Set wb = Workbooks.Open(CStr(f), ReadOnly:=True)
    wb.Sheets(1).UsedRange.Copy wsRaw.Range("A1")
    wb.Close SaveChanges:=False

    Dim nC As Long, nR As Long
    nC = wsRaw.Cells(1, wsRaw.Columns.Count).End(xlToLeft).Column
    nR = wsRaw.Cells(wsRaw.Rows.Count, 1).End(xlUp).Row
    If nR < 2 Then MsgBox "Settlement file is empty.", vbExclamation: Exit Sub

    Dim ph As Variant, pv As Variant, sh As Variant, sv As Variant
    ph = lo.HeaderRowRange.Value
    pv = lo.DataBodyRange.Value
    sh = wsRaw.Range(wsRaw.Cells(1, 1), wsRaw.Cells(1, nC)).Value
    sv = wsRaw.Range(wsRaw.Cells(2, 1), wsRaw.Cells(nR, nC)).Value

    Dim cTxn As Long, cType As Long, cAmt As Long, cStat As Long
    Dim sTxn As Long, sAmt As Long, sFee As Long
    cTxn = Col(ph, C_TXN): cType = Col(ph, C_TYPE): cAmt = Col(ph, C_AMT): cStat = Col(ph, C_STATUS)
    sTxn = Col(sh, P_TXN): sAmt = Col(sh, P_AMT): sFee = Col(sh, P_FEE)
    If cTxn * cType * cAmt = 0 Then MsgBox "A required column is missing in the Payments table (check the CONFIG names).", vbExclamation: Exit Sub
    If sTxn * sAmt * sFee = 0 Then MsgBox "A required column is missing in the settlement file (txn_id, amount, fee).", vbExclamation: Exit Sub

    ' ledger lookup
    Dim led As Object, seen As Object, cnt As Object
    Set led = CreateObject("Scripting.Dictionary")
    Set seen = CreateObject("Scripting.Dictionary")
    Set cnt = CreateObject("Scripting.Dictionary")
    Dim i As Long, j As Long, key As String
    For i = 1 To UBound(pv, 1)
        If StatusOk(pv, i, cStat) Then
            key = Trim(CStr(pv(i, cTxn)))
            If key <> "" Then led(key) = i
        End If
    Next i

    Dim out As Variant, k As Long
    ReDim out(1 To UBound(pv, 1) + UBound(sv, 1) + 1, 1 To 7)

    Dim lAmt As Double, pAmt As Double, pFee As Double, eFee As Double
    Dim clean As Long, bad As Boolean

    For j = 1 To UBound(sv, 1)
        key = Trim(CStr(sv(j, sTxn)))
        If key <> "" Then
            If seen.Exists(key) Then
                AddRow out, k, "Duplicate at PSP", key, "", sv(j, sAmt), sv(j, sFee), "", "Same txn_id settled more than once"
            ElseIf Not led.Exists(key) Then
                AddRow out, k, "Missing in ledger", key, "", sv(j, sAmt), sv(j, sFee), "", "PSP settled a txn we have no record of"
            Else
                seen(key) = True
                i = led(key)
                lAmt = CDbl(pv(i, cAmt))
                pAmt = CDbl(sv(j, sAmt))
                pFee = CDbl(sv(j, sFee))
                eFee = ExpectedFee(LCase(Trim(CStr(pv(i, cType)))), lAmt)
                bad = False
                If Abs(pAmt - lAmt) > AMT_TOL Then
                    AddRow out, k, "Amount mismatch", key, lAmt, pAmt, pFee, eFee, "Difference: " & Format(pAmt - lAmt, "0.00")
                    bad = True
                End If
                If Abs(pFee - eFee) > FEE_TOL Then
                    AddRow out, k, "Fee mismatch", key, lAmt, pAmt, pFee, eFee, "Fee off by " & Format(pFee - eFee, "0.00")
                    bad = True
                End If
                If Not bad Then clean = clean + 1
            End If
        End If
    Next j

    Dim v As Variant
    For Each v In led.Keys
        If Not seen.Exists(CStr(v)) Then
            i = led(v)
            AddRow out, k, "Missing at PSP", CStr(v), pv(i, cAmt), "", "", "", "In ledger but not in the settlement file"
        End If
    Next v

    ' exceptions sheet
    Dim wsE As Worksheet: Set wsE = FreshSheet("Recon_Exceptions")
    wsE.Range("A1:G1").Value = Array("category", "txn_id", "ledger_amount", "psp_amount", "psp_fee", "expected_fee", "note")
    StyleHeader wsE.Range("A1:G1")
    If k > 0 Then
        wsE.Range("A2").Resize(UBound(out, 1), 7).Value = out
        wsE.Range("A1").Resize(k + 1, 7).AutoFilter
        For i = 1 To k
            cnt(out(i, 1)) = cnt(out(i, 1)) + 1
        Next i
    End If
    wsE.Columns("A:G").AutoFit

    ' summary sheet
    Dim wsS As Worksheet: Set wsS = FreshSheet("Recon_Summary")
    wsS.Range("A1").Value = "PSP Reconciliation Summary"
    wsS.Range("A1").Font.Bold = True: wsS.Range("A1").Font.Size = 14
    wsS.Range("A2").Value = "Run": wsS.Range("B2").Value = Format(Now, "yyyy-mm-dd hh:nn")
    wsS.Range("A3").Value = "File": wsS.Range("B3").Value = CStr(f)
    wsS.Range("A5:B5").Value = Array("Metric", "Count")
    StyleHeader wsS.Range("A5:B5")
    wsS.Range("A6").Value = "Ledger transactions checked": wsS.Range("B6").Value = led.Count
    wsS.Range("A7").Value = "PSP rows": wsS.Range("B7").Value = UBound(sv, 1)
    wsS.Range("A8").Value = "Clean matches": wsS.Range("B8").Value = clean
    Dim cats As Variant, r As Long
    cats = Array("Missing at PSP", "Missing in ledger", "Amount mismatch", "Fee mismatch", "Duplicate at PSP")
    For r = 0 To 4
        wsS.Cells(9 + r, 1).Value = cats(r)
        If cnt.Exists(cats(r)) Then wsS.Cells(9 + r, 2).Value = cnt(cats(r)) Else wsS.Cells(9 + r, 2).Value = 0
    Next r
    wsS.Columns("A:B").AutoFit
    wsS.Activate

    Application.ScreenUpdating = True
    MsgBox "Reconciliation done. " & k & " exceptions found. See Recon_Summary and Recon_Exceptions.", vbInformation
End Sub

' ================= 2. ONE-CLICK DAILY REPORT =================
Public Sub DailyReport()
    Dim lo As ListObject: Set lo = GetTable()
    If lo Is Nothing Then Exit Sub

    ' refresh queries first
    ThisWorkbook.RefreshAll
    On Error Resume Next
    Application.CalculateUntilAsyncQueriesDone
    On Error GoTo 0

    Dim ph As Variant, pv As Variant
    ph = lo.HeaderRowRange.Value
    pv = lo.DataBodyRange.Value

    Dim cTxn As Long, cPl As Long, cDt As Long, cTy As Long, cAmt As Long, cStat As Long
    cTxn = Col(ph, C_TXN): cPl = Col(ph, C_PLAYER): cDt = Col(ph, C_DATE)
    cTy = Col(ph, C_TYPE): cAmt = Col(ph, C_AMT): cStat = Col(ph, C_STATUS)
    If cTxn * cPl * cDt * cTy * cAmt = 0 Then MsgBox "A required column is missing (check the CONFIG names).", vbExclamation: Exit Sub

    ' latest date in the data as default
    Dim i As Long, mx As Double
    For i = 1 To UBound(pv, 1)
        If IsDate(pv(i, cDt)) Then If CDbl(Int(CDate(pv(i, cDt)))) > mx Then mx = CDbl(Int(CDate(pv(i, cDt))))
    Next i

    Dim ans As String
    ans = InputBox("Report date (yyyy-mm-dd):", "Daily report", Format(CDate(mx), "yyyy-mm-dd"))
    If ans = "" Then Exit Sub
    If Not IsDate(ans) Then MsgBox "Not a valid date.", vbExclamation: Exit Sub
    Dim tgt As Double: tgt = CDbl(Int(CDate(ans)))

    Application.ScreenUpdating = False

    Dim depN As Long, wdN As Long, depS As Double, wdS As Double
    Dim pl As Object: Set pl = CreateObject("Scripting.Dictionary")
    Dim big As Variant, nb As Long
    ReDim big(1 To UBound(pv, 1), 1 To 4)
    Dim ty As String, a As Double, p As String

    For i = 1 To UBound(pv, 1)
        If IsDate(pv(i, cDt)) Then
            If CDbl(Int(CDate(pv(i, cDt)))) = tgt And StatusOk(pv, i, cStat) Then
                ty = LCase(Trim(CStr(pv(i, cTy))))
                a = CDbl(pv(i, cAmt))
                p = CStr(pv(i, cPl))
                If ty = "deposit" Then depN = depN + 1: depS = depS + a
                If ty = "withdrawal" Then wdN = wdN + 1: wdS = wdS + a
                pl(p) = pl(p) + 1
                If a >= LARGE_TXN Then
                    nb = nb + 1
                    big(nb, 1) = pv(i, cTxn): big(nb, 2) = p: big(nb, 3) = ty: big(nb, 4) = a
                End If
            End If
        End If
    Next i

    Dim ws As Worksheet: Set ws = FreshSheet("Daily_Report")
    ws.Range("A1").Value = "FraudHub Daily Report": ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 16
    ws.Range("A2").Value = "Date": ws.Range("B2").Value = Format(CDate(tgt), "yyyy-mm-dd")
    ws.Range("A3").Value = "Generated": ws.Range("B3").Value = Format(Now, "yyyy-mm-dd hh:nn")

    ws.Range("A5:C5").Value = Array("Type", "Count", "Amount")
    StyleHeader ws.Range("A5:C5")
    ws.Range("A6:C6").Value = Array("Deposits", depN, depS)
    ws.Range("A7:C7").Value = Array("Withdrawals", wdN, wdS)
    ws.Range("A8:C8").Value = Array("Net (deposits - withdrawals)", depN + wdN, depS - wdS)
    ws.Range("A8:C8").Font.Bold = True
    ws.Range("C6:C8").NumberFormat = "#,##0.00"

    Dim r As Long: r = 10
    ws.Cells(r, 1).Value = "Large transactions (>= " & Format(LARGE_TXN, "#,##0") & ")"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)).Value = Array("txn_id", "player_id", "type", "amount")
    StyleHeader ws.Range(ws.Cells(r, 1), ws.Cells(r, 4))
    If nb > 0 Then
        ws.Cells(r + 1, 1).Resize(UBound(big, 1), 4).Value = big
        ws.Range(ws.Cells(r + 1, 4), ws.Cells(r + nb, 4)).NumberFormat = "#,##0.00"
        r = r + nb + 2
    Else
        ws.Cells(r + 1, 1).Value = "None"
        r = r + 3
    End If

    ws.Cells(r, 1).Value = "High-velocity players (>= " & VELOCITY_MIN & " txns in the day)"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Value = Array("player_id", "txn_count")
    StyleHeader ws.Range(ws.Cells(r, 1), ws.Cells(r, 2))
    Dim v As Variant, found As Boolean
    For Each v In pl.Keys
        If pl(v) >= VELOCITY_MIN Then
            r = r + 1
            ws.Cells(r, 1).Value = v
            ws.Cells(r, 2).Value = pl(v)
            found = True
        End If
    Next v
    If Not found Then ws.Cells(r + 1, 1).Value = "None"

    ws.Columns("A:D").AutoFit
    ws.Activate
    Application.ScreenUpdating = True
    MsgBox "Daily report ready for " & Format(CDate(tgt), "yyyy-mm-dd") & ".", vbInformation
End Sub

