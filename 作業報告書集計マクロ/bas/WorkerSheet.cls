VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkerSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' --- 列挙型: 行インデックス ---
Private Enum RowIndex
    YEAR_ROW = 2          ' 年
    MONTH_ROW             ' 月
    DAYS_ROW              ' 日 (1～31)
    DISPLAY_DAY_ROW       ' 日 (表示)
    WEEK                  ' 曜日
    YOTEI_TIME            ' 予定時間
    JITUDO_TIME           ' 実働時間
    SEIKYU_GAI_TIME       ' 請求時間（外）
    SENGETU_THOSEI_TIME   ' 先月調整時間
    SEIKYU_HOUR           ' 請求時間
    SEIKYU_SECOND         ' 請求分
    YOKUGETU_THOSEI_TIME  ' 翌月調整時間
    ZANGYO_TIME           ' 残業時間
    BIKO                  ' 備考
    ADMIN_SETTING         ' 管理者設定
    SYU_TITLE = 21        ' 週のタイトル
End Enum

' --- 列挙型: 列インデックス ---
Private Enum ColumnIndex
    StartColumn = 2                ' 列開始
    SYU_SINTYOKU_JOKYO = 2         ' 進捗状況
    SYU_FUKA_JOKYO = 4             ' 負荷状況
    SYU_SAGYO_NAIYO = 6            ' 作業内容
    SYU_KADAI = 14                 ' 課題
    SYU_HITOKOTO = 20              ' 今週の一言
    EndColumn = 32                 ' 列終了
    TOTAL_COLUMN                   ' 合計列
End Enum

' --- ワークシート変数 ---
Private m_WorkerSheet As Worksheet

' --- ワークシートの設定プロパティ ---
Public Property Set WorkerSheet(Value As Worksheet)
    Set m_WorkerSheet = Value
End Property




' --- 予定時間の取得 ---
Public Function GetYoteiTime(TargetDate As Date) As Double
    GetYoteiTime = GetTargetValue(TargetDate, YOTEI_TIME)
End Function

' --- 実働時間の取得 ---
Public Function GetJitudoTime(TargetDate As Date) As Double
    GetJitudoTime = GetTargetValue(TargetDate, JITUDO_TIME)
End Function


' --- 月の実働時間合計の取得 ---
Public Function GetJitudoTimeTotal(EndDate As Date) As Double
    GetJitudoTimeTotal = GetHoursBetween(GetDate(ColumnIndex.StartColumn), EndDate, JITUDO_TIME)
End Function


' --- 出勤日数の取得 (8時間で判定) ---
Public Function GetSyukkinNissu(TargetDate As Date) As Double
    Dim Jitudo As Double
    Jitudo = GetTargetValue(TargetDate, JITUDO_TIME)

    If Jitudo > 0 And Jitudo < 8 Then
        GetSyukkinNissu = 0.5
    ElseIf Jitudo >= 8 Then
        GetSyukkinNissu = 1
    Else
        GetSyukkinNissu = 0
    End If
End Function




' --- 指定日の進捗状況を取得 ---
Public Function GetSintyokuJokyo(TargetDate As Date) As String
    Dim SyuRowIndex As Integer
    SyuRowIndex = RowIndex.SYU_TITLE + GetSyuNo(TargetDate)
    GetSintyokuJokyo = m_WorkerSheet.Cells(SyuRowIndex, ColumnIndex.SYU_SINTYOKU_JOKYO).Value
End Function




' --- 負荷状況の取得 ---
Public Function GetFukaJokyo(TargetDate As Date) As String
    Dim SyuRowIndex As Integer
    SyuRowIndex = RowIndex.SYU_TITLE + GetSyuNo(TargetDate)
    GetFukaJokyo = m_WorkerSheet.Cells(SyuRowIndex, ColumnIndex.SYU_FUKA_JOKYO).Value
End Function

' --- 作業内容の取得 ---
Public Function GetSagyoNaiyo(TargetDate As Date) As String
    Dim SyuRowIndex As Integer
    SyuRowIndex = RowIndex.SYU_TITLE + GetSyuNo(TargetDate)
    GetSagyoNaiyo = m_WorkerSheet.Cells(SyuRowIndex, ColumnIndex.SYU_SAGYO_NAIYO).Value
End Function


' --- 課題の取得 ---
Public Function GetKadai(TargetDate As Date) As String
    Dim SyuRowIndex As Integer
    SyuRowIndex = RowIndex.SYU_TITLE + GetSyuNo(TargetDate)
    GetKadai = m_WorkerSheet.Cells(SyuRowIndex, ColumnIndex.SYU_KADAI).Value

    ' 空欄の場合は "特になし" を返す
    If GetKadai = "" Then
        GetKadai = "特になし"
    End If
End Function

' --- 今週の一言の取得 ---
Public Function GetHitokoto(TargetDate As Date) As String
    Dim SyuRowIndex As Integer
    SyuRowIndex = RowIndex.SYU_TITLE + GetSyuNo(TargetDate)
    GetHitokoto = m_WorkerSheet.Cells(SyuRowIndex, ColumnIndex.SYU_HITOKOTO).Value

    ' 空欄の場合は "特になし" を返す
    If GetHitokoto = "" Then
        GetHitokoto = "特になし"
    End If
End Function


' --- 月の先月調整時間の合計 (0.0h) を取得 ---
Public Function GetSengetuTyoseiTotal() As Double
    GetSengetuTyoseiTotal = m_WorkerSheet.Cells(RowIndex.SENGETU_THOSEI_TIME, ColumnIndex.TOTAL_COLUMN + 1).Value
    If Not IsNumeric(GetSengetuTyoseiTotal) Then
        GetSengetuTyoseiTotal = 0
    End If
End Function

' --- 月の請求時間 (0.0h) を取得 ---
Public Function GetSeikyuTotal() As Double
    GetSeikyuTotal = m_WorkerSheet.Cells(RowIndex.SEIKYU_HOUR, ColumnIndex.TOTAL_COLUMN).Value + _
                     m_WorkerSheet.Cells(RowIndex.SEIKYU_SECOND, ColumnIndex.TOTAL_COLUMN) / 60
End Function

' --- 月の年休予定を取得 (MM/DD(曜日): 年休内容) ---
Public Function GetNenkyuYotei() As String
    Dim i As Integer
    Dim CurrentDate As Date
    Dim WeekValue As Range
    Dim CurrentValue As Variant

    For i = ColumnIndex.StartColumn To ColumnIndex.EndColumn
        CurrentDate = GetDate(i)
        Set WeekValue = m_WorkerSheet.Cells(RowIndex.WEEK, i)

        If CurrentDate <> WeekValue Then & "・" & _
            GetNenkyuYotei = GetNenkyuYotei & _
                             Format(CurrentDate, "MM/DD(aaa)") & "：" & _
                             WeekValue.Text & vbLf
        End If
    Next i
End Function




' --- 月の翌月調整時間の合計を取得 ---
Public Function GetYokugetuTyoseiTotal() As Double
    Dim i As Integer
    Dim CurrentDate As Date
    Dim JitudoValue As Variant
    Dim Result As Double

    Result = 0
    For i = ColumnIndex.StartColumn To ColumnIndex.EndColumn
        CurrentDate = GetDate(i)
        If m_WorkerSheet.Cells(RowIndex.JITUDO_TIME, i).Value <> "" Then
            Result = Result + GetTargetValue(CurrentDate, YOKUGETU_THOSEI_TIME)
        End If
    Next i
    GetYokugetuTyoseiTotal = Result
End Function

' --- 月の残業時間を取得 ---
Public Function GetZangyoTotal() As Double
    GetZangyoTotal = m_WorkerSheet.Cells(RowIndex.ZANGYO_TIME, ColumnIndex.TOTAL_COLUMN).Value
End Function

' --- 月の当月稼働見込みを取得 ---
Public Function GetTougetuKadoMikomi() As Double
    Dim i As Integer
    Dim CurrentDate As Date
    Dim Result As Double

    Result = 0
    For i = ColumnIndex.StartColumn To ColumnIndex.EndColumn
        CurrentDate = GetDate(i)
        If m_WorkerSheet.Cells(RowIndex.JITUDO_TIME, i).Value <> "" Then
            Result = Result + GetTargetValue(CurrentDate, JITUDO_TIME)
        ElseIf m_WorkerSheet.Cells(RowIndex.YOTEI_TIME, i).Value <> "" Then
            Result = Result + GetTargetValue(CurrentDate, YOTEI_TIME)
        End If
    Next i
    GetTougetuKadoMikomi = Result
End Function

' --- 指定日の管理者設定に作業報告を追記 ---
Public Sub SettingWorkingReportEnd(TargetDate As Date)
    Dim i As Integer
    Dim CurrentDate As Date

    For i = ColumnIndex.StartColumn To ColumnIndex.EndColumn
        If IsDate(m_WorkerSheet.Cells(RowIndex.DISPLAY_DAY_ROW, i).Value) Then
            CurrentDate = m_WorkerSheet.Cells(RowIndex.DISPLAY_DAY_ROW, i).Value
            If CurrentDate = TargetDate Then
                m_WorkerSheet.Cells(RowIndex.ADMIN_SETTING, i).Value = "作業報告済"
            End If
        End If
    Next i
End Sub

' --- 予定時間の編集を不可にする (セルを塗りつぶす) ---
Public Sub LockYoteiTime()
    Dim TargetRange As Range

    With m_WorkerSheet
        Set TargetRange = .Range(.Cells(RowIndex.YOTEI_TIME, ColumnIndex.StartColumn), _
                                 .Cells(RowIndex.YOTEI_TIME, ColumnIndex.EndColumn))
        .Unprotect
        .Cells.Locked = False
        
        TargetRange.Interior.Color = RGB(192, 192, 192)
        TargetRange.Locked = True
        
        .Protect
    End With
End Sub

' --- 指定日の稼働時間を取得 ---
Public Function GetWorkingHours(TargetDate As Date) As Double
    Dim TotalHours As Double
    Dim CurrentDate As Date
    Dim i As Integer

    TotalHours = 0

    For i = ColumnIndex.StartColumn To ColumnIndex.EndColumn
        If IsDate(m_WorkerSheet.Cells(RowIndex.DISPLAY_DAY_ROW, i).Value) Then
            CurrentDate = m_WorkerSheet.Cells(RowIndex.DISPLAY_DAY_ROW, i).Value
            If CurrentDate = TargetDate Then
                TotalHours = TotalHours + m_WorkerSheet.Cells(RowIndex.JITUDO_TIME, i).Value
            End If
        End If
    Next i

    GetWorkingHours = TotalHours
End Function

' --- 指定期間の稼働時間を取得 ---
Public Function GetWorkingHoursBetween(StartDate As Date, EndDate As Date) As Double
    GetWorkingHoursBetween = GetHoursBetween(StartDate, EndDate, RowIndex.JITUDO_TIME)
End Function

' --- 月の実働時間の差分を取得 ---
Public Function GetDiffWorkingHours() As Double
    Dim TotalHours As Double
    Dim CurrentDate As Date
    Dim i As Integer
    Dim TargetValue As String

    TotalHours = 0

    For i = ColumnIndex.StartColumn To ColumnIndex.EndColumn
        TargetValue = m_WorkerSheet.Cells(RowIndex.JITUDO_TIME, i).Value
        If TargetValue <> "" Then
            TotalHours = TotalHours + m_WorkerSheet.Cells(RowIndex.YOKUGETU_THOSEI_TIME, i).Value
        End If
    Next i

    GetDiffWorkingHours = TotalHours
End Function


' --- 月の稼働 (年休、予定時間、稼働時間) を取得 ---
Public Function GetTukiKado() As TukiKadoData
    Dim Result As New TukiKadoData
    Dim CurrentDate As Date
    Dim CurrentValue As Variant
    Dim i As Integer
    Dim Index As Integer
    Dim Nenkyu() As Variant

    ReDim Nenkyu(1 To 31)

    With m_WorkerSheet
        Index = 1
        For i = ColumnIndex.StartColumn To ColumnIndex.EndColumn
            CurrentDate = GetDate(i)
            CurrentValue = GetTargetValue(CurrentDate, WEEK)
            If Not IsNumeric(CurrentValue) Then
                Nenkyu(Index) = CurrentValue
            End If
            Index = Index + 1
        Next i

        Result.NenkyuYotei = Nenkyu
        Result.YoteiTime = .Range(.Cells(RowIndex.YOTEI_TIME, ColumnIndex.StartColumn), _
                                  .Cells(RowIndex.YOTEI_TIME, ColumnIndex.TOTAL_COLUMN)).Value
        Result.JitudoTime = .Range(.Cells(RowIndex.JITUDO_TIME, ColumnIndex.StartColumn), _
                                   .Cells(RowIndex.JITUDO_TIME, ColumnIndex.TOTAL_COLUMN)).Value
    End With

    Set GetTukiKado = Result
End Function

' --- 指定列の年月日のDateを取得 ---
Private Function GetDate(ColumnValue As Integer) As Date
    GetDate = DateSerial( _
        CInt(m_WorkerSheet.Cells(RowIndex.YEAR_ROW, ColumnValue).Value), _
        CInt(m_WorkerSheet.Cells(RowIndex.MONTH_ROW, ColumnValue).Value), _
        CInt(m_WorkerSheet.Cells(RowIndex.DAYS_ROW, ColumnValue).Value))
End Function

' --- 指定日の値を取得 ---
Private Function GetTargetValue(TargetDate As Date, TargetRow As RowIndex) As Variant
    Dim CurrentDate As Date
    Dim i As Integer

    For i = ColumnIndex.StartColumn To ColumnIndex.EndColumn
        CurrentDate = GetDate(i)
        If CurrentDate = TargetDate Then
            GetTargetValue = m_WorkerSheet.Cells(TargetRow, i).Value
            Exit Function
        End If
    Next i
End Function


' --- 指定日の値を取得 ---
Private Function GetTargetValue(TargetDate As Date, TargetRow As RowIndex) As Variant
    Dim CurrentDate As Date
    Dim i As Integer

    For i = ColumnIndex.StartColumn To ColumnIndex.EndColumn
        CurrentDate = GetDate(i)
        If CurrentDate = TargetDate Then
            GetTargetValue = m_WorkerSheet.Cells(TargetRow, i).Value
            Exit Function
        End If
    Next i
End Function

' --- 指定日の週の何週目かを取得 ---
Private Function GetSyuNo(TargetDate As Date) As Integer
    GetSyuNo = Int((Day(TargetDate) - 1) / 7) + 1
End Function

' --- 指定した行の期間内の合計時間を取得する ---
Private Function GetHoursBetween(StartDate As Date, EndDate As Date, TargetRowIndex As RowIndex) As Double
    Dim TotalHours As Double
    Dim CurrentDate As Date
    Dim i As Integer

    TotalHours = 0

    For i = ColumnIndex.StartColumn To ColumnIndex.EndColumn
        If IsDate(m_WorkerSheet.Cells(RowIndex.DISPLAY_DAY_ROW, i).Value) Then
            CurrentDate = m_WorkerSheet.Cells(RowIndex.DISPLAY_DAY_ROW, i).Value
            If CurrentDate >= StartDate And CurrentDate <= EndDate Then
                TotalHours = TotalHours + m_WorkerSheet.Cells(TargetRowIndex, i).Value
            End If
        End If
    Next i

    GetHoursBetween = TotalHours
End Function


