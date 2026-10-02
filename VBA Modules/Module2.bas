Attribute VB_Name = "Module2"
Option Explicit

' ================= CONFIG =================
Private Const PAY_TABLE As String = "Payments"
Private Const C_PLAYER As String = "player_id"
Private Const C_DATE As String = "txn_time"
Private Const C_AMT As String = "amount"
Private Const C_STATUS As String = "status"
Private Const C_CARD As String = "card_hash"
Private Const OK_STATUS As String = "|success|"
Private Const FAIL_STATUS As String = "failed"

Private Const SHARED_PTS As Long = 40
Private Const VEL_PTS As Long = 15
Private Const VEL_CAP As Long = 30
Private Const VELOCITY_MIN As Long = 5
Private Const FAIL_MIN As Long = 4
Private Const FAIL_PTS As Long = 20
Private Const LARGE_AMT As Double = 300
Private Const LARGE_PTS As Long = 5
Private Const LARGE_CAP As Long = 15
Private Const HIGH_SCORE As Long = 50
Private Const MED_SCORE As Long = 30

Private Const REFRESH_FIRST As Boolean = True
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
        If ws.AutoFilterMode Then ws.AutoFilterMode = False
        ws.Cells.Clear
    End If
    Set FreshSheet = ws
End Function

Private Sub StyleHeader(rng As Range)
    With rng
        .Font.Bold = True
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = vbWhite
    End With
End Sub

' ================= RISK QUEUE =================
Public Sub RiskQueue()
    Dim lo As ListObject: Set lo = GetTable()
    If lo Is Nothing Then Exit Sub

    If REFRESH_FIRST Then
        ThisWorkbook.RefreshAll
        On Error Resume Next
        Application.CalculateUntilAsyncQueriesDone
        On Error GoTo 0
    End If

    Dim ph As Variant, pv As Variant
    ph = lo.HeaderRowRange.Value
    pv = lo.DataBodyRange.Value

    Dim cPl As Long, cDt As Long, cAmt As Long, cSt As Long, cCard As Long
    cPl = Col(ph, C_PLAYER): cDt = Col(ph, C_DATE): cAmt = Col(ph, C_AMT)
    cSt = Col(ph, C_STATUS): cCard = Col(ph, C_CARD)
    If cPl * cDt * cAmt * cSt * cCard = 0 Then
        MsgBox "A required column is missing in the Payments table (check the CONFIG names).", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim players As Object, cardP As Object, cardN As Object, cardAmt As Object
    Dim vel As Object, fails As Object, large As Object
    Set players = CreateObject("Scripting.Dictionary")
    Set cardP = CreateObject("Scripting.Dictionary")
    Set cardN = CreateObject("Scripting.Dictionary")
    Set cardAmt = CreateObject("Scripting.Dictionary")
    Set vel = CreateObject("Scripting.Dictionary")
    Set fails = CreateObject("Scripting.Dictionary")
    Set large = CreateObject("Scripting.Dictionary")

    Dim i As Long, j As Long
    Dim pl As String, st As String, cd As String, k As String, a As Double

    ' ---- pass 1: collect signals ----
    For i = 1 To UBound(pv, 1)
        pl = Trim(CStr(pv(i, cPl)))
        If pl <> "" Then
            players(pl) = True
            st = LCase(Trim(CStr(pv(i, cSt))))
            If st = FAIL_STATUS Then
                fails(pl) = fails(pl) + 1
            ElseIf InStr(1, OK_STATUS, "|" & st & "|", vbTextCompare) > 0 Then
                a = CDbl(pv(i, cAmt))
                If a >= LARGE_AMT Then large(pl) = large(pl) + 1
                If IsDate(pv(i, cDt)) Then
                    k = CStr(CLng(Int(CDate(pv(i, cDt))))) & "|" & pl
                    vel(k) = vel(k) + 1
                End If
                cd = Trim(CStr(pv(i, cCard)))
                If cd <> "" Then
                    If Not cardP.Exists(cd) Then Set cardP(cd) = CreateObject("Scripting.Dictionary")
                    cardP(cd)(pl) = True
                    cardN(cd) = cardN(cd) + 1
                    cardAmt(cd) = cardAmt(cd) + a
                End If
            End If
        End If
    Next i

    ' ---- high-velocity days per player ----
    Dim velDays As Object: Set velDays = CreateObject("Scripting.Dictionary")
    Dim v As Variant, nVelDays As Long
    For Each v In vel.Keys
        If vel(v) >= VELOCITY_MIN Then
            pl = Mid(CStr(v), InStr(CStr(v), "|") + 1)
            velDays(pl) = velDays(pl) + 1
            nVelDays = nVelDays + 1
        End If
    Next v

    ' ---- shared cards sheet ----
    Dim sharedCard As Object, linked As Object
    Set sharedCard = CreateObject("Scripting.Dictionary")
    Set linked = CreateObject("Scripting.Dictionary")
    Dim wsC As Worksheet: Set wsC = FreshSheet("Shared_Cards")
    wsC.Range("A1:E1").Value = Array("card_hash", "players", "player_ids", "successful_txns", "total_amount")
    StyleHeader wsC.Range("A1:E1")
    Dim nShared As Long, rC As Long, pk As Variant, lst As String
    rC = 1
    For Each v In cardP.Keys
        If cardP(v).Count >= 2 Then
            nShared = nShared + 1
            rC = rC + 1
            lst = Join(cardP(v).Keys, ", ")
            wsC.Cells(rC, 1).Value = CStr(v)
            wsC.Cells(rC, 2).Value = cardP(v).Count
            wsC.Cells(rC, 3).Value = lst
            wsC.Cells(rC, 4).Value = cardN(v)
            wsC.Cells(rC, 5).Value = cardAmt(v)
            For Each pk In cardP(v).Keys
                If Not sharedCard.Exists(CStr(pk)) Then
                    sharedCard(CStr(pk)) = CStr(v)
                    linked(CStr(pk)) = lst
                End If
            Next pk
        End If
    Next v
    wsC.Range("E2:E" & rC).NumberFormat = "#,##0.00"
    wsC.Columns("A:E").AutoFit

    ' ---- scoring ----
    Dim tmp() As Variant
    ReDim tmp(1 To players.Count, 1 To 9)
    Dim n As Long, s As Long, reasons As String
    Dim nvd As Long, nf As Long, nl As Long, nFail As Long, nHigh As Long, nMed As Long
    Dim lvl As String

    For Each v In players.Keys
        pl = CStr(v)
        s = 0
        reasons = ""

        If sharedCard.Exists(pl) Then
            s = s + SHARED_PTS
            reasons = reasons & "Shared card; "
        End If

        nvd = 0
        If velDays.Exists(pl) Then nvd = velDays(pl)
        If nvd > 0 Then
            s = s + IIf(nvd * VEL_PTS > VEL_CAP, VEL_CAP, nvd * VEL_PTS)
            reasons = reasons & "High velocity x" & nvd & " day(s); "
        End If

        nf = 0
        If fails.Exists(pl) Then nf = fails(pl)
        If nf >= FAIL_MIN Then
            s = s + FAIL_PTS
            nFail = nFail + 1
            reasons = reasons & nf & " failed payments; "
        End If

        nl = 0
        If large.Exists(pl) Then nl = large(pl)
        If nl > 0 Then
            s = s + IIf(nl * LARGE_PTS > LARGE_CAP, LARGE_CAP, nl * LARGE_PTS)
            reasons = reasons & nl & " large txn(s); "
        End If

        If s >= HIGH_SCORE Then
            lvl = "High": nHigh = nHigh + 1
        ElseIf s >= MED_SCORE Then
            lvl = "Medium": nMed = nMed + 1
        Else
            lvl = ""
        End If

        If lvl <> "" Then
            If Len(reasons) > 2 Then reasons = Left(reasons, Len(reasons) - 2)
            n = n + 1
            tmp(n, 1) = pl
            tmp(n, 2) = s
            tmp(n, 3) = lvl
            If sharedCard.Exists(pl) Then
                tmp(n, 4) = sharedCard(pl)
                tmp(n, 5) = linked(pl)
            End If
            tmp(n, 6) = nvd
            tmp(n, 7) = nf
            tmp(n, 8) = nl
            tmp(n, 9) = reasons
        End If
    Next v

    ' ---- queue sheet ----
    Dim wsQ As Worksheet: Set wsQ = FreshSheet("Risk_Queue")
    wsQ.Range("A1:I1").Value = Array("player_id", "risk_score", "risk_level", "shared_card", _
        "players_on_card", "high_vel_days", "failed_txns", "large_txns", "reasons")
    StyleHeader wsQ.Range("A1:I1")

    If n > 0 Then
        Dim fin() As Variant
        ReDim fin(1 To n, 1 To 9)
        For i = 1 To n
            For j = 1 To 9
                fin(i, j) = tmp(i, j)
            Next j
        Next i
        wsQ.Range("A2").Resize(n, 9).Value = fin

        With wsQ.Sort
            .SortFields.Clear
            .SortFields.Add key:=wsQ.Range("B2").Resize(n, 1), SortOn:=xlSortOnValues, Order:=xlDescending
            .SortFields.Add key:=wsQ.Range("A2").Resize(n, 1), SortOn:=xlSortOnValues, Order:=xlAscending
            .SetRange wsQ.Range("A1").Resize(n + 1, 9)
            .Header = xlYes
            .Apply
        End With

        For i = 2 To n + 1
            If wsQ.Cells(i, 3).Value = "High" Then
                wsQ.Range(wsQ.Cells(i, 1), wsQ.Cells(i, 9)).Interior.Color = RGB(255, 199, 206)
            Else
                wsQ.Range(wsQ.Cells(i, 1), wsQ.Cells(i, 9)).Interior.Color = RGB(255, 235, 156)
            End If
        Next i
        wsQ.Range("A1").Resize(n + 1, 9).AutoFilter
    End If
    wsQ.Columns("A:I").AutoFit
    If wsQ.Columns("E").ColumnWidth > 40 Then wsQ.Columns("E").ColumnWidth = 40

    ' ---- summary sheet ----
    Dim wsS As Worksheet: Set wsS = FreshSheet("Risk_Summary")
    wsS.Range("A1").Value = "Risk Summary"
    wsS.Range("A1").Font.Bold = True: wsS.Range("A1").Font.Size = 14
    wsS.Range("A2").Value = "Run": wsS.Range("B2").Value = Format(Now, "yyyy-mm-dd hh:nn")
    wsS.Range("A4:B4").Value = Array("Metric", "Value")
    StyleHeader wsS.Range("A4:B4")
    wsS.Range("A5").Value = "Shared cards (2+ players)": wsS.Range("B5").Value = nShared
    wsS.Range("A6").Value = "Players on shared cards": wsS.Range("B6").Value = sharedCard.Count
    wsS.Range("A7").Value = "High-velocity player-days": wsS.Range("B7").Value = nVelDays
    wsS.Range("A8").Value = "Players with " & FAIL_MIN & "+ failed payments": wsS.Range("B8").Value = nFail
    wsS.Range("A9").Value = "High risk (>= " & HIGH_SCORE & ")": wsS.Range("B9").Value = nHigh
    wsS.Range("A10").Value = "Medium risk (" & MED_SCORE & "-" & (HIGH_SCORE - 1) & ")": wsS.Range("B10").Value = nMed
    wsS.Range("A11").Value = "Queue size": wsS.Range("B11").Value = n
    wsS.Columns("A:B").AutoFit

    wsQ.Activate
    Application.ScreenUpdating = True
    MsgBox "Risk queue built: " & n & " players (" & nHigh & " High, " & nMed & " Medium)." & vbCrLf & _
           "See Risk_Queue, Shared_Cards and Risk_Summary.", vbInformation
End Sub

' ================= CONTROL SHEET WITH BUTTONS =================
Public Sub SetupControlSheet()
    Dim ws As Worksheet: Set ws = FreshSheet("Control")
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i

    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 70

    ws.Range("B1").Value = "FraudHub Control Panel"
    ws.Range("B1").Font.Bold = True: ws.Range("B1").Font.Size = 18
    ws.Range("B2").Value = "Refreshes the daily payment files, then runs each check with one click."
    ws.Range("B2").Font.Italic = True

    Dim names As Variant, macros As Variant, descs As Variant
    names = Array("1. PSP Reconciliation", "2. Daily Report", "3. Risk Queue")
    macros = Array("ReconcilePSP", "DailyReport", "RiskQueue")
    descs = Array("Match the ledger against a PSP settlement file. Finds missing, duplicate, amount and fee breaks.", _
                  "Deposits, withdrawals, large transactions and high-velocity players for one day.", _
                  "Scores every player (shared cards, velocity, failures, large txns) and lists who to review first.")

    Dim r As Long, shp As Shape
    For i = 0 To 2
        r = 4 + i * 2
        ws.Rows(r).RowHeight = 38
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Cells(r, 2).Left + 2, ws.Cells(r, 2).Top + 3, _
                                     ws.Cells(r, 2).Width - 4, ws.Rows(r).Height - 6)
        shp.OnAction = macros(i)
        shp.Fill.ForeColor.RGB = RGB(31, 56, 100)
        shp.Line.Visible = msoFalse
        shp.TextFrame2.TextRange.Text = names(i)
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        shp.TextFrame2.TextRange.Font.Bold = msoTrue
        shp.TextFrame2.TextRange.Font.Size = 12
        ws.Cells(r, 3).Value = descs(i)
        ws.Cells(r, 3).VerticalAlignment = xlCenter
        ws.Cells(r, 3).WrapText = True
    Next i

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Range("A1").Select
End Sub

