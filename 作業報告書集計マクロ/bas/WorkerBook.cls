VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkerBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private m_WorkerName As String
Private m_BookPath As String

Private WkBook As Workbook
Private SheetDic As Object

' --- WorkerNameのプロパティ設定 ---
Public Property Let WorkerName(Value As String)
    m_WorkerName = Value
End Property

Public Property Get WorkerName() As String
    WorkerName = m_WorkerName
End Property

' --- BookPathのプロパティ設定 ---
Public Property Let BookPath(Value As String)
    m_BookPath = Value
End Property

Public Property Get BookPath() As String
    BookPath = m_BookPath
End Property

' --- 初期化処理: ワークシートオブジェクトを生成する ---
Public Sub Initialize()
    Dim TargetSheet As Worksheet
    Dim TargetWorkerSheet As WorkerSheet

    Set WkBook = Workbooks.Open(m_BookPath)
    Set SheetDic = CreateObject("Scripting.Dictionary")

    For Each TargetSheet In WkBook.Worksheets
        Set TargetWorkerSheet = New WorkerSheet
        Set TargetWorkerSheet.WorkerSheet = TargetSheet
        SheetDic.Add TargetSheet.Name, TargetWorkerSheet
    Next TargetSheet
End Sub

' --- 終了処理 ---
Public Sub Finalize()
    If Not WkBook Is Nothing Then
        WkBook.Close SaveChanges:=False
        Set WkBook = Nothing
        Set SheetDic = Nothing
    End If
End Sub


' --- 原本から作業報告シートをコピーしてシート名を設定する ---
Public Sub CopySheetFrom(TargetBook As Workbook, SheetName As String)
    Dim TargetSheet As Worksheet
    Dim CopiedSheet As Worksheet

    If SheetDic.Exists(SheetName) Then
        Call MsgBox(SheetName & " シートが既に存在します。ブック名=" & WkBook.Name)
        Exit Sub
    End If

    Set TargetSheet = TargetBook.Sheets("原本")
    TargetSheet.Copy After:=WkBook.Sheets(WkBook.Sheets.Count)
    Set CopiedSheet = WkBook.Sheets(WkBook.Sheets.Count)
    CopiedSheet.Name = SheetName

    Call WkBook.Save
End Sub

' --- 原本から祝日リストをコピーする ---
Public Sub UpdateHolidayListFrom(TargetBook As Workbook, SheetName As String)
    Dim TargetSheet As Worksheet
    Dim CopiedSheet As Worksheet
    Dim lastRow As Long
    Dim data As Variant

    If Not SheetDic.Exists(SheetName) Then
        Call MsgBox(SheetName & " シートが存在しません。ブック名=" & WkBook.Name)
        Exit Sub
    End If

    Set TargetSheet = TargetBook.Sheets(SheetName)
    Set CopiedSheet = WkBook.Sheets(SheetName)

    lastRow = TargetSheet.Cells(TargetSheet.Rows.Count, "A").End(xlUp).Row
    data = TargetSheet.Range("A1:B" & lastRow).Value

    Call CopiedSheet.Range("A1:B").ClearContents
    CopiedSheet.Range("A1").Resize(UBound(data, 1), UBound(data, 2)).Value = data

    Call WkBook.Save
End Sub

' --- 指定日の管理者設定に作業報告〆を追記 ---
Public Sub SettingWorkingReportEnd(TargetDate As Date)
    Dim SheetName As String
    Dim TargetWorkerSheet As WorkerSheet

    SheetName = Format(TargetDate, "yyyy-mm")
    If SheetDic.Exists(SheetName) Then
        Set TargetWorkerSheet = SheetDic(SheetName)
        Call TargetWorkerSheet.SettingWorkingReportEnd(TargetDate)
        Call WkBook.Save
    Else
        Call MsgBox(SheetName & " シートが存在しません。ブック名=" & WkBook.Name)
    End If
End Sub


' --- 指定した月の予定時間の編集を不可にする (& セルを塗りつぶす) ---
Public Sub LockYoteiTime(SheetName As String)
    Dim TargetWorkerSheet As WorkerSheet

    If SheetDic.Exists(SheetName) Then
        Set TargetWorkerSheet = SheetDic(SheetName)
        Call TargetWorkerSheet.LockYoteiTime
        Call WkBook.Save
    Else
        Call MsgBox(SheetName & " シートが存在しません。ブック名=" & WkBook.Name)
    End If
End Sub

' --- 指定期間の稼働時間を取得 ---
Public Function GetWorkingHoursBetween(StartDate As Date, EndDate As Date) As Double
    Dim TotalHours As Double
    Dim CurrentDate As Date
    Dim SheetName As String
    Dim TargetWorkerSheet As WorkerSheet

    TotalHours = 0

    For CurrentDate = StartDate To EndDate
        SheetName = Format(CurrentDate, "yyyy-mm")
        If SheetDic.Exists(SheetName) Then
            Set TargetWorkerSheet = SheetDic(SheetName)
            TotalHours = TotalHours + TargetWorkerSheet.GetWorkingHours(CurrentDate)
        End If
    Next CurrentDate

    GetWorkingHoursBetween = TotalHours
End Function

' --- 指定した月の実働時間の差分を取得 ---
Public Function GetDiffWorkingHours(SheetName As String) As String
    Dim Result As String
    Dim TargetWorkerSheet As WorkerSheet

    If SheetDic.Exists(SheetName) Then
        Set TargetWorkerSheet = SheetDic(SheetName)
        Result = m_WorkerName & " : " & Format(TargetWorkerSheet.GetDiffWorkingHours, "0.0") & " H"
        GetDiffWorkingHours = Result
    Else
        Call MsgBox(SheetName & " シートが存在しません。ブック名=" & WkBook.Name)
    End If
End Function











' --- 週報のデータを作成 ---
Public Function CreateSyuhoData(StartDate As Date, EndDate As Date) As Collection
    Dim Result As New Collection
    Dim CurrentStart As Date
    Dim CurrentEnd As Date
    Dim CurrentSyuhoData As SyuhoData

    ' 月マタギ対応
    CurrentStart = StartDate
    Do While CurrentStart <= EndDate
        CurrentEnd = Application.WorksheetFunction.Min(EndDate, _
                      DateSerial(Year(CurrentStart), Month(CurrentStart) + 1, 0))

        Set CurrentSyuhoData = CreateMonthSyuhoData(CurrentStart, CurrentEnd)
        Call Result.Add(CurrentSyuhoData)

        CurrentStart = CurrentEnd + 1
    Loop

    Set CreateSyuhoData = Result
End Function

' --- 月の稼働 (年休、予定時間、稼働時間) を取得 ---
Public Function GetTukiKado(SheetName As String) As TukiKadoData
    Dim Result As TukiKadoData
    Dim TargetWorkerSheet As WorkerSheet

    If SheetDic.Exists(SheetName) Then
        Set TargetWorkerSheet = SheetDic(SheetName)
        Set Result = TargetWorkerSheet.GetTukiKado()
        Result.Name = m_WorkerName
        Set GetTukiKado = Result
    Else
        Call MsgBox(SheetName & " シートが存在しません。ブック名=" & WkBook.Name)
    End If
End Function

' --- 選択期間 (月マタギなしの期間) の週報データを作成 ---
Private Function CreateMonthSyuhoData(StartDate As Date, EndDate As Date) As SyuhoData
    Dim CurrentDate As Date
    Dim SheetName As String
    Dim TargetWorkerSheet As WorkerSheet
    Dim Result As New SyuhoData
    Dim TargetSheetDict As Object
    Dim Key As Variant

    Set TargetSheetDict = CreateObject("Scripting.Dictionary")
    Result.Name = m_WorkerName
    Result.Kikan = StartDate & " ～ " & EndDate

    ' 日毎の処理
    For CurrentDate = StartDate To EndDate
        SheetName = Format(CurrentDate, "yyyy-mm")
        If SheetDic.Exists(SheetName) Then
            Set TargetWorkerSheet = SheetDic(SheetName)
            TargetSheetDict(SheetName) = TargetWorkerSheet

            ' 稼働時間
            Result.SyukadoJikan = Result.SyukadoJikan + TargetWorkerSheet.GetYoteiTime(CurrentDate)
            ' 出勤日数
            Result.SyukkinNissu = Result.SyukkinNissu + TargetWorkerSheet.GetSyukkinNissu(CurrentDate)
        End If
    Next CurrentDate

    ' 月毎の処理
    For Each Key In TargetSheetDict.Keys
        Set TargetWorkerSheet = SheetDic(Key)

        ' 年休予定
        Result.NenkyuYotei = Result.NenkyuYotei & TargetWorkerSheet.GetNenkyuYotei
    Next Key

    ' 最後の日
    SheetName = Format(EndDate, "yyyy-mm")
    If SheetDic.Exists(SheetName) Then
        Set TargetWorkerSheet = SheetDic(SheetName)
        With TargetWorkerSheet
            ' 実働時間合計
            Result.JitudoTimeTotal = .GetJitudoTimeTotal(EndDate)
            ' 当月稼働見込
            Result.TougetuKadoMikomi = .GetTougetuKadoMikomi()
            ' 請求時間
            Result.SeikyuJikan = .GetSeikyuTotal()
            ' 先月調整
            Result.ZengetuTyosei = .GetSengetuTyoseiTotal()
            ' 翌月調整
            Result.YokgetuTyosei = .GetYokgetuTyoseiTotal()
            ' 残業時間
            Result.ZangyoJikan = .GetZangyoTotal()

            ' 進捗状況
            Result.SintyokuJokyo = .GetSintyokuJokyo(EndDate)
            ' 負荷状況
            Result.HukaJokyo = .GetHukaJokyo(EndDate)
            ' 作業内容
            Result.SagyoNaiyo = .GetSagyoNaiyo(EndDate)
            ' 課題
            Result.Kadai = .GetKadai(EndDate)
            ' 今週の一言
            Result.Hitokoto = .GetHitokoto(EndDate)
        End With
    End If

    Set CreateMonthSyuhoData = Result
End Function
