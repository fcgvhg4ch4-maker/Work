Attribute VB_Name = "modQIS"
'==============================================================================
' QIS Fund Monitoring & Weekly Sales Intelligence Workbook - VBA module
'
' 導入: VBE（Alt+F11）> ファイル > ファイルのインポート で本ファイルを取り込み、
'       ブックを .xlsm で保存する。00_Control にボタンを置く場合は各マクロを割り当てる。
'
' 方針:
'   - メールは「下書き保存」のみ。.Send は一切呼ばない（自動送信しない）。
'   - ポジション履歴は値で追記し、同一 Snapshot ID（基準日|Index Name|ファイル名）は追記しない。
'   - 名前定義（EvalDate, EM_Rep, EM_Body 等）とシート名に依存する。列構成を変える場合は定数を修正。
'   - Bloomberg Add-in のマクロ名は環境により異なるため、QIS_RefreshBloomberg の候補リストを要確認。
'==============================================================================
Option Explicit

Private Const SH_RAW As String = "05_Position_Raw"
Private Const SH_HIST As String = "06_Position_History"
Private Const SH_EMAIL As String = "16_Email_Generator"
Private Const SH_LOG As String = "17_Email_History"
Private Const SH_CLIENT As String = "03_Client_Master"

Private Const RAW_FIRST_ROW As Long = 6
Private Const RAW_NCOLS As Long = 19      ' 17 列 + Source File + Imported At
Private Const HIST_FIRST_ROW As Long = 6
Private Const HIST_CAP As Long = 1500     ' build_workbook.py の CAP_PH と一致させる
Private Const LOG_FIRST_ROW As Long = 6
Private Const LOG_CAP As Long = 500
Private Const REP_FIRST_ROW As Long = 61
Private Const REP_CAP As Long = 20

'------------------------------------------------------------------------------
' 1) 一括更新：Power Query → Bloomberg → 再計算
'------------------------------------------------------------------------------
Public Sub QIS_RefreshAll()
    Dim t As Double: t = Timer
    Application.StatusBar = "Power Query を更新中..."
    On Error GoTo EH
    ThisWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone
    If UCase$(CStr(Range("S_DataMode").Value)) = "BLOOMBERG" Then
        QIS_RefreshBloomberg
    End If
    Application.StatusBar = "再計算中..."
    Application.CalculateFull
    Application.StatusBar = False
    MsgBox "更新完了（" & Format(Timer - t, "0.0") & " 秒）" & vbCrLf & _
           "データ品質: " & Range("DQ_Overall").Value & vbCrLf & _
           "次の手順: QIS_ArchivePositions でポジションを履歴へ保存", vbInformation, "QIS"
    Exit Sub
EH:
    Application.StatusBar = False
    MsgBox "更新でエラー: " & Err.Description, vbExclamation, "QIS"
End Sub

'------------------------------------------------------------------------------
' 2) Bloomberg 再取得（Add-in のマクロ名は環境で要確認）
'------------------------------------------------------------------------------
Public Sub QIS_RefreshBloomberg()
    Dim names As Variant, i As Long, ok As Boolean
    names = Array("RefreshEntireWorkbook", "BloombergUI.RefreshEntireWorkbook", "RefreshAllStaticData")
    For i = LBound(names) To UBound(names)
        On Error Resume Next
        Err.Clear
        Application.Run names(i)
        If Err.Number = 0 Then ok = True: Exit For
        On Error GoTo 0
    Next i
    On Error GoTo 0
    If Not ok Then
        MsgBox "Bloomberg Add-in の更新マクロを呼び出せませんでした。Bloomberg リボンの Refresh を手動で実行してください。", vbExclamation, "QIS"
        Exit Sub
    End If
    ' BDH/BDP は非同期。結果がそろうまで待つ（最大 60 秒）
    Dim t0 As Double: t0 = Timer
    Do While Timer - t0 < 60
        DoEvents
        If Application.CalculationState = xlDone Then Exit Do
    Loop
    Application.Calculate
End Sub

'------------------------------------------------------------------------------
' 3) 05_Position_Raw → 06_Position_History へ値で追記（重複防止）
'------------------------------------------------------------------------------
Public Sub QIS_ArchivePositions()
    Dim wsR As Worksheet, wsH As Worksheet
    Set wsR = ThisWorkbook.Worksheets(SH_RAW)
    Set wsH = ThisWorkbook.Worksheets(SH_HIST)

    Dim lastRaw As Long: lastRaw = LastDataRow(wsR, 1, RAW_FIRST_ROW)
    If lastRaw < RAW_FIRST_ROW Then MsgBox "05_Position_Raw にデータがありません。", vbExclamation, "QIS": Exit Sub

    ' 既存の Snapshot ID と（基準日|Index Name）→ ファイル名
    Dim existing As Object: Set existing = CreateObject("Scripting.Dictionary")
    Dim dateIdx As Object: Set dateIdx = CreateObject("Scripting.Dictionary")
    Dim lastHist As Long: lastHist = LastDataRow(wsH, 4, HIST_FIRST_ROW)
    Dim r As Long, k As String, k2 As String
    For r = HIST_FIRST_ROW To lastHist
        k = CStr(wsH.Cells(r, 1).Value)
        If Len(k) > 0 Then existing(k) = True
        k2 = Format(wsH.Cells(r, 4).Value, "yyyymmdd") & "|" & CStr(wsH.Cells(r, 5).Value)
        If Not dateIdx.Exists(k2) Then dateIdx(k2) = CStr(wsH.Cells(r, 3).Value)
    Next r

    Dim added As Long, skipped As Long, conflicts As String, nowStamp As Date
    nowStamp = Now
    Dim dest As Long: dest = IIf(lastHist < HIST_FIRST_ROW, HIST_FIRST_ROW, lastHist + 1)
    Dim c As Long, snap As String, src As String
    For r = RAW_FIRST_ROW To lastRaw
        If Len(CStr(wsR.Cells(r, 1).Value)) = 0 Then GoTo NextRow
        src = CStr(wsR.Cells(r, 18).Value)
        snap = Format(wsR.Cells(r, 1).Value, "yyyymmdd") & "|" & CStr(wsR.Cells(r, 2).Value) & "|" & src
        k2 = Format(wsR.Cells(r, 1).Value, "yyyymmdd") & "|" & CStr(wsR.Cells(r, 2).Value)
        If existing.Exists(snap) Then
            skipped = skipped + 1
            GoTo NextRow
        End If
        If dateIdx.Exists(k2) Then
            If dateIdx(k2) <> src Then
                If InStr(conflicts, k2) = 0 Then conflicts = conflicts & vbCrLf & k2 & "（既存: " & dateIdx(k2) & " / 今回: " & src & "）"
                skipped = skipped + 1
                GoTo NextRow
            End If
        End If
        If dest > HIST_FIRST_ROW + HIST_CAP - 1 Then
            MsgBox "06_Position_History の容量（" & HIST_CAP & " 行）を超えます。容量拡張または古い履歴の退避が必要です。", vbCritical, "QIS"
            Exit For
        End If
        wsH.Cells(dest, 1).Value = snap
        wsH.Cells(dest, 2).Value = nowStamp
        wsH.Cells(dest, 3).Value = src
        For c = 1 To 17
            wsH.Cells(dest, 3 + c).Value = wsR.Cells(r, c).Value
        Next c
        dest = dest + 1
        added = added + 1
NextRow:
    Next r
    ' 同一スナップショット内の 2 行目以降も追記できるよう、既存判定はループ開始前の状態のみで行っている
    Application.Calculate
    Dim msg As String
    msg = "追記: " & added & " 行 / スキップ（既存または衝突）: " & skipped & " 行"
    If Len(conflicts) > 0 Then msg = msg & vbCrLf & vbCrLf & "同一基準日・同一指数で別ファイルが既に保存されているため追記しませんでした（要確認）:" & conflicts
    MsgBox msg, IIf(Len(conflicts) > 0, vbExclamation, vbInformation), "QIS"
End Sub

'------------------------------------------------------------------------------
' 4) 現在表示中のメールを 17_Email_History に記録
'------------------------------------------------------------------------------
Public Sub QIS_LogEmail(Optional ByVal draftCreated As Boolean = False)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_LOG)
    Dim r As Long: r = LastDataRow(ws, 1, LOG_FIRST_ROW) + 1
    If r < LOG_FIRST_ROW Then r = LOG_FIRST_ROW
    If r > LOG_FIRST_ROW + LOG_CAP - 1 Then MsgBox "17_Email_History が満杯です。", vbCritical, "QIS": Exit Sub
    ws.Cells(r, 1).Value = Format(Now, "yyyymmddhhnnss") & "-" & CStr(Range("EM_Rep").Value)
    ws.Cells(r, 2).Value = Now
    ws.Cells(r, 3).Value = Range("EvalDate").Value
    ws.Cells(r, 4).Value = Range("EM_Rep").Value
    ws.Cells(r, 5).Value = Range("EM_To").Value
    ws.Cells(r, 6).Value = Range("EM_Subject").Value
    ws.Cells(r, 7).Value = Range("EM_Body").Value
    ws.Cells(r, 8).Value = Range("DQ_Overall").Value
    ws.Cells(r, 9).Value = Range("EM_Status").Value
    ws.Cells(r, 10).Value = Environ$("USERNAME")
    ws.Cells(r, 11).Value = Range("S_WorkbookVersion").Value
    ws.Cells(r, 12).Value = IIf(draftCreated, "Y", "N")
End Sub

'------------------------------------------------------------------------------
' 5) 全営業担当者の Outlook 下書きを作成（送信はしない）
'------------------------------------------------------------------------------
Public Sub QIS_CreateOutlookDrafts()
    If CStr(Range("DQ_Overall").Value) = "要確認" Then
        If MsgBox("データ品質チェックに未解決の重大項目があります。件名に【要確認】を付けて下書きを作成しますか？", _
                  vbYesNo + vbExclamation, "QIS") = vbNo Then Exit Sub
    End If
    Dim ol As Object
    On Error Resume Next
    Set ol = GetObject(, "Outlook.Application")
    If ol Is Nothing Then Set ol = CreateObject("Outlook.Application")
    On Error GoTo 0
    If ol Is Nothing Then MsgBox "Outlook を起動できません。本文は 16_Email_Generator からコピーしてください。", vbExclamation, "QIS": Exit Sub

    Dim wsC As Worksheet: Set wsC = ThisWorkbook.Worksheets(SH_CLIENT)
    Dim original As String: original = CStr(Range("EM_Rep").Value)
    Dim r As Long, nm As String, n As Long, mail As Object, subj As String
    For r = REP_FIRST_ROW To REP_FIRST_ROW + REP_CAP - 1
        nm = CStr(wsC.Cells(r, 2).Value)
        If Len(nm) > 0 And UCase$(CStr(wsC.Cells(r, 5).Value)) = "Y" Then
            Range("EM_Rep").Value = nm
            Application.Calculate
            If CStr(Range("EM_Status").Value) <> "対象なし" Then
                subj = CStr(Range("EM_Subject").Value)
                If Left$(CStr(Range("EM_Status").Value), 3) = "要確認" Then subj = "【要確認】" & subj
                Set mail = ol.CreateItem(0)   ' olMailItem
                mail.To = CStr(Range("EM_To").Value)
                mail.Subject = subj
                mail.Body = Replace(CStr(Range("EM_Body").Value), vbLf, vbCrLf)
                mail.Save                       ' 下書き保存のみ（.Send は呼ばない）
                QIS_LogEmail True
                n = n + 1
            End If
        End If
    Next r
    Range("EM_Rep").Value = original
    Application.Calculate
    MsgBox n & " 件の下書きを作成しました（Outlook の下書きフォルダ）。内容を確認のうえ手動で送信してください。", vbInformation, "QIS"
End Sub

'------------------------------------------------------------------------------
' 6) 表示中のメール本文をクリップボードへ
'------------------------------------------------------------------------------
Public Sub QIS_CopyEmailBody()
    Dim txt As String: txt = Replace(CStr(Range("EM_Body").Value), vbLf, vbCrLf)
    Dim dobj As Object
    Set dobj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  ' MSForms.DataObject
    dobj.SetText "件名: " & Range("EM_Subject").Value & vbCrLf & vbCrLf & txt
    dobj.PutInClipboard
    MsgBox "件名と本文をクリップボードにコピーしました。", vbInformation, "QIS"
End Sub

'------------------------------------------------------------------------------
Private Function LastDataRow(ws As Worksheet, col As Long, firstRow As Long) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    Do While r >= firstRow
        If Len(CStr(ws.Cells(r, col).Value)) > 0 Then Exit Do
        r = r - 1
    Loop
    LastDataRow = r
End Function
