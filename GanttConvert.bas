Attribute VB_Name = "GanttConvert"
Option Explicit

' ============================================================
' ガントチャート取込データ変換VBA v5
' ============================================================
' [変更点 v5]
'   - 祝日・長期休暇をVBAハードコードから「休日」シート管理に変更
'   - 休日シート: A列に日付を列挙するだけ（ヘッダー行は自動スキップ）
'   - 土日は自動除外（休日シートへの記載不要）

' モジュールレベル: 起動時に一度だけ読み込む休日配列
Private holidayList() As Date
Private holidayCount As Long

' ---- 休日シートから休日を読み込む ----
Private Sub LoadHolidays()
    holidayCount = 0
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("休日")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "「休日」シートが見つかりません。" & vbCrLf & _
               "ガントCV.xlsm に「休日」シートを作成し、A列に休日日付を入力してください。", vbExclamation
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ReDim holidayList(1 To lastRow)

    Dim r As Long
    For r = 1 To lastRow
        Dim cell As Range
        Set cell = ws.Cells(r, 1)
        If IsDate(cell.Value) Then
            holidayCount = holidayCount + 1
            holidayList(holidayCount) = CDate(cell.Value)
        End If
    Next r
End Sub

' ---- 祝日・長期休暇チェック（土日含む） ----
Private Function IsHoliday(d As Date) As Boolean
    ' 土日は無条件で休日
    If Weekday(d, vbMonday) >= 6 Then IsHoliday = True: Exit Function
    ' 休日シートの日付と照合
    Dim i As Long
    For i = 1 To holidayCount
        If Year(holidayList(i)) = Year(d) And _
           Month(holidayList(i)) = Month(d) And _
           Day(holidayList(i)) = Day(d) Then
            IsHoliday = True: Exit Function
        End If
    Next i
    IsHoliday = False
End Function

' ---- N営業日前後の日付を返す (days<0で過去) ----
Private Function ShiftBizDays(startDate As Date, days As Integer) As Date
    Dim d As Date
    d = startDate
    Dim stp As Integer
    stp = IIf(days < 0, -1, 1)
    Dim remaining As Integer
    remaining = Abs(days)
    Do While remaining > 0
        d = d + stp
        If Not IsHoliday(d) Then remaining = remaining - 1
    Loop
    ShiftBizDays = d
End Function

' ---- 品番マッピング ----
Private Function GetPartMapping() As Variant
    GetPartMapping = Array( _
        Array("23J7011109",  "GD955", "アジャスター", "23J7013231"), _
        Array("23J7013231",  "GD955", "アジャスター", "23J7013231"), _
        Array("23J7011109",  "GD955", "ドローバ",    "23J7012614"), _
        Array("23J7012614",  "GD955", "ドローバ",    "23J7012614"), _
        Array("23J7011109",  "GD955", "サークル",    "23J7011355"), _
        Array("23J7011355",  "GD955", "サークル",    "23J7011355"), _
        Array("23J7015103NK","GD955", "ブレード",    ""), _
        Array("23J7015104NK","GD955", "ブレード",    ""), _
        Array("23J7015804NK","GD955", "ブレード",    ""), _
        Array("X251933200",  "GD825", "アジャスター", "2357043252"), _
        Array("2357043252",  "GD825", "アジャスター", "2357043252"), _
        Array("2077001170NK","特殊機", "大）マテハン", "") _
    )
End Function

' ---- 工番コード ----
Private Function GetKoban(partType As String, refDate As Date) As String
    Dim yr As Integer
    If Month(refDate) >= 4 Then
        yr = Year(refDate) Mod 100
    Else
        yr = (Year(refDate) - 1) Mod 100
    End If
    Select Case partType
        Case "アジャスター": GetKoban = yr & "RCIR"
        Case "ドローバ":    GetKoban = yr & "RDOR"
        Case "ブレード":    GetKoban = yr & "RGDB"
        Case "サークル":    GetKoban = yr & "RCIR"
        Case "大）マテハン": GetKoban = yr & "R"
        Case Else:          GetKoban = yr & "R???"
    End Select
End Function

' ---- LMMDD形式 → Date ----
Private Function CodeToDate(code As String) As Date
    code = Trim(code)
    CodeToDate = 0
    If Len(code) <> 5 Or Not IsNumeric(code) Then Exit Function
    Dim yr As Integer, mm As Integer, dd As Integer
    yr = 2020 + CInt(Left(code, 1))
    mm = CInt(Mid(code, 2, 2))
    dd = CInt(Mid(code, 4, 2))
    If yr < 2024 Or yr > 2035 Or mm < 1 Or mm > 12 Or dd < 1 Or dd > 31 Then Exit Function
    On Error Resume Next
    CodeToDate = DateSerial(yr, mm, dd)
    On Error GoTo 0
End Function

Private Function DateCodeToLong(code As String) As Long
    code = Trim(code)
    If Len(code) = 5 And IsNumeric(code) Then DateCodeToLong = CLng(code) Else DateCodeToLong = 0
End Function

' ---- メイン処理 ----
Public Sub ConvertToGantt()
    ' 休日読み込み（最初に実行）
    LoadHolidays
    If holidayCount = 0 Then
        If MsgBox("休日シートに有効な日付がありません。" & vbCrLf & _
                  "土日のみ除外して続行しますか？", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    End If

    Dim srcFolder As String
    srcFolder = ThisWorkbook.Path & "\元データ\"
    If Dir(srcFolder, vbDirectory) = "" Then
        MsgBox "元データフォルダが見つかりません:" & vbCrLf & srcFolder, vbExclamation
        Exit Sub
    End If

    Dim mapping As Variant
    mapping = GetPartMapping()

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim fName As String
    fName = Dir(srcFolder & "*.xlsm")
    Dim fileCount As Integer
    fileCount = 0
    Do While fName <> ""
        If Left(fName, 2) <> "~$" Then
            ProcessFile srcFolder & fName, mapping, dict
            fileCount = fileCount + 1
        End If
        fName = Dir()
    Loop

    If fileCount = 0 Then
        MsgBox "元データフォルダに .xlsm ファイルが見つかりません。", vbExclamation: Exit Sub
    End If
    If dict.Count = 0 Then
        MsgBox "対象品番のデータが見つかりませんでした。" & vbCrLf & _
               "GetPartMapping の品番パターンを確認してください。", vbExclamation: Exit Sub
    End If

    ' ---- 出力ファイル作成 ----
    Dim outWb As Workbook
    Set outWb = Workbooks.Add
    Dim ws As Worksheet
    Set ws = outWb.Sheets(1)
    ws.Name = "取込データ"

    With ws
        .Cells(1, 1) = "ガントチャート取込用データ（手修正不要）"
        .Cells(2, 1) = "作成日時: " & Format(Now, "YYYY/MM/DD HH:MM")
        .Cells(3, 1) = "読込ファイル数: " & fileCount & " 件"
        .Cells(5, 1) = "仕様名"
        .Cells(5, 2) = "品番"
        .Cells(5, 3) = "工番"
        .Cells(5, 5) = "コマツ納期"
        .Cells(5, 8) = "搬出日"
        .Cells(5, 9) = "搬入日"
        .Range("A5:I5").Interior.Color = RGB(180, 200, 220)
        .Range("A5:I5").Font.Bold = True
    End With

    ws.Columns("E:E").NumberFormat = "@"
    ws.Columns("H:H").NumberFormat = "@"
    ws.Columns("I:I").NumberFormat = "@"
    Dim specOrder As Variant
    specOrder = Array("GD955", "GD825", "K56", "特殊機")
    Dim outRow As Integer
    outRow = 7

    Dim sp As Variant
    For Each sp In specOrder
        Dim lastSpec As String
        lastSpec = ""
        Dim k As Variant
        For Each k In dict.Keys
            Dim v As Variant
            v = dict(k)
            If v(0) = CStr(sp) Then
                Dim kDate As Date
                kDate = CodeToDate(CStr(v(3)))
                If kDate = 0 Then GoTo SkipRow

                Dim hikDate As Date
                hikDate = ShiftBizDays(kDate, -6)
                Dim hanshutsuDate As Date
                hanshutsuDate = ShiftBizDays(hikDate, -5)

                Dim specVal As String
                If v(0) <> lastSpec Then
                    specVal = v(0)
                    lastSpec = v(0)
                Else
                    specVal = ""
                End If

                ws.Cells(outRow, 1) = specVal
                ws.Cells(outRow, 2) = v(2)
                ws.Cells(outRow, 3) = v(5)
                ws.Cells(outRow, 5) = Format(kDate, "yyyy-mm-dd")
                ws.Cells(outRow, 8) = Format(hanshutsuDate, "yyyy-mm-dd")
                ws.Cells(outRow, 9) = Format(hikDate, "yyyy-mm-dd")
                outRow = outRow + 1
SkipRow:
            End If
        Next k
    Next sp

    ws.Columns("A:I").AutoFit
    If outRow > 6 Then
        With ws.Range("A5:I" & (outRow - 1))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = RGB(180, 180, 180)
        End With
    End If

    Dim savePath As String
    savePath = ThisWorkbook.Path & "\取込データ_" & Format(Now, "YYYYMMDD") & ".xlsx"
    Application.DisplayAlerts = False
    outWb.SaveAs savePath, xlOpenXMLWorkbook
    Application.DisplayAlerts = True

    MsgBox "完了！" & vbCrLf & "出力レコード: " & (outRow - 7) & " 件" & vbCrLf & savePath, vbInformation
    outWb.Activate
    Set dict = Nothing
End Sub

' ---- ファイル読込サブルーチン ----
Private Sub ProcessFile(filePath As String, mapping As Variant, dict As Object)
    Dim wb As Workbook
    On Error GoTo ErrHandler
    Set wb = Workbooks.Open(filePath, False, True)

    Dim ws As Worksheet
    Dim s As Worksheet
    For Each s In wb.Sheets
        If s.Name = "Sheet1" Then Set ws = s: Exit For
    Next s
    If ws Is Nothing Then wb.Close False: Exit Sub

    Dim lastRow As Long
    lastRow = ws.UsedRange.Rows.Count

    Dim r As Long
    For r = 2 To lastRow
        Dim partNo As String
        partNo = Trim(ws.Cells(r, 5).Text)
        If partNo = "" Or partNo = "品番" Then GoTo NextRow

        Dim noki As String
        noki = Trim(ws.Cells(r, 8).Text)
        If DateCodeToLong(noki) = 0 Then GoTo NextRow

        Dim kobanBase As String
        Dim spIdx As Integer
        spIdx = InStr(partNo, "  ")
        If spIdx > 0 Then kobanBase = Left(partNo, spIdx - 1) Else kobanBase = partNo

        Dim mi As Integer
        For mi = 0 To UBound(mapping)
            If InStr(partNo, CStr(mapping(mi)(0))) > 0 Then
                Dim specName As String
                Dim partType As String
                Dim outPartNo As String
                specName  = CStr(mapping(mi)(1))
                partType  = CStr(mapping(mi)(2))
                outPartNo = CStr(mapping(mi)(3))

                Dim dictKey As String
                dictKey = specName & "|" & partType & "|" & kobanBase & "|" & noki

                Dim bCol As String
                If outPartNo <> "" Then
                    bCol = partType & outPartNo
                Else
                    bCol = partType & Replace(partNo, " ", "")
                End If

                If Not dict.Exists(dictKey) Then
                    Dim kd As Date
                    kd = CodeToDate(noki)
                    Dim koban As String
                    koban = GetKoban(partType, kd)
                    dict(dictKey) = Array(specName, partType, bCol, noki, outPartNo, koban)
                End If
            End If
        Next mi
NextRow:
    Next r

    wb.Close False
    Exit Sub
ErrHandler:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
End Sub

' ---- 休日シートのセットアップ（初回実行用） ----
Public Sub SetupHolidaySheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("休日")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "休日"
    End If

    With ws
        .Cells(1, 1) = "休日一覧"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Interior.Color = RGB(255, 230, 153)
        .Cells(2, 1) = "※土日は自動除外のため入力不要"
        .Cells(2, 1).Font.Color = RGB(128, 128, 128)
        .Cells(2, 1).Font.Italic = True
        .Cells(3, 1) = "※日付形式で入力（例: 2026/01/01）"
        .Cells(3, 1).Font.Color = RGB(128, 128, 128)
        .Cells(3, 1).Font.Italic = True

        ' サンプル: 2026年祝日
        Dim sample As Variant
        sample = Array( _
            "2026/01/01", "2026/01/02", "2026/01/03", _
            "2026/01/12", "2026/02/11", "2026/02/23", _
            "2026/03/20", "2026/04/29", _
            "2026/05/03", "2026/05/04", "2026/05/05", "2026/05/06", _
            "2026/07/20", "2026/08/11", _
            "2026/08/13", "2026/08/14", "2026/08/15", _
            "2026/09/21", "2026/09/22", "2026/09/23", _
            "2026/10/12", "2026/11/03", "2026/11/23", _
            "2026/12/29", "2026/12/30", "2026/12/31" _
        )
        Dim i As Integer
        For i = 0 To UBound(sample)
            .Cells(4 + i, 1) = CDate(sample(i))
            .Cells(4 + i, 1).NumberFormat = "yyyy/mm/dd"
        Next i
        .Columns("A").AutoFit
    End With

    MsgBox "「休日」シートを作成しました。" & vbCrLf & _
           "2026年の祝日サンプルを入力済みです。" & vbCrLf & _
           "会社カレンダーに合わせて追加・修正してください。", vbInformation
End Sub
