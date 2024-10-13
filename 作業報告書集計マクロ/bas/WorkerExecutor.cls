VERSION 1.0 CLASS
BEGIN
  MultiUse = -1 'True
END
Attribute VB_Name = "WorkerExecutor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 指定年月シートの作成
Public Sub 指定年月シートの作成()
    Dim Col As Collection
    Dim TargetBook As WorkerBook
    Dim GenponBook As Workbook

    Set GenponBook = Workbooks.Open(Sheet3.Get指定年月シートの作成_対象ブック())

    Set Col = Sheet1.GetTargetWorkerList()
    For Each TargetBook In Col
        Call TargetBook.Initialize
        Call TargetBook.CopySheetFrom(GenponBook, Sheet3.Get指定年月シートの作成_対象シート())
        Call TargetBook.Finalize
    Next TargetBook

    Call GenponBook.Close(False)
End Sub

' 作業報告〆日の設定
Public Sub 作業報告〆日の設定()
    Dim Col As Collection
    Dim TargetBook As WorkerBook

    Set Col = Sheet1.GetTargetWorkerList()
    For Each TargetBook In Col
        Call TargetBook.Initialize
        Call TargetBook.SettingWorkingReportEnd(Sheet3.Get作業報告〆日の設定_作業報告〆日())
        Call TargetBook.Finalize
    Next TargetBook
End Sub

' 作業報告〆
Public Sub 作業報告〆()
    Dim Col As Collection
    Dim TargetBook As WorkerBook

    Set Col = Sheet1.GetTargetWorkerList()
    For Each TargetBook In Col
        Call TargetBook.Initialize
        Call TargetBook.LockYoteiTime(Sheet3.Get作業報告〆_対象シート())
        Call TargetBook.Finalize
    Next TargetBook
End Sub




Public Function 実働時間の差分出力() As String
    Dim Result As String
    Dim Col As Collection
    Dim TargetBook As WorkerBook

    Result = ""
    Set Col = Sheet1.GetTargetWorkerList()
    For Each TargetBook In Col
        Call TargetBook.Initialize
        Result = Result & TargetBook.GetDiffWorkingHours(Sheet3.Get実働時間の差分出力_対象シート()) & vbCrLf
        Call TargetBook.Finalize
    Next TargetBook

    実働時間の差分出力 = Result
End Function

Public Sub 祝日リストの更新()
    Dim Col As Collection
    Dim TargetBook As WorkerBook
    Dim GenponBook As Workbook

    Set GenponBook = Workbooks.Open(Sheet3.Get祝日リストの更新_対象ブック())

    Set Col = Sheet1.GetTargetWorkerList()
    For Each TargetBook In Col
        Call TargetBook.Initialize
        Call TargetBook.UpdateHolydayListFrom(GenponBook, Sheet3.Get祝日リストの更新_対象シート())
        Call TargetBook.Finalize
    Next TargetBook

    Call GenponBook.Close(False)
End Sub




Public Sub 週報の集計()
    Dim Col As Collection
    Dim TargetBook As WorkerBook
    Dim Result As Collection
    Dim CurrentShuhoData As SyuhoData
    Dim SyuhoSheet As Worksheet
    Dim i As Long
    Dim Index As Long

    Const StartRowIndex As Long = 4
    Const StartColumnIndex As Long = 1
    Const EndColumnIndex As Long = 17

    Set SyuhoSheet = ThisWorkbook.Sheets("週報")
    With SyuhoSheet
        Set Col = Sheet1.GetTargetWorkerList()
        i = StartRowIndex

        ' 値と罫線を初期化
        .Range(.Cells(i, StartColumnIndex), .Cells(300, EndColumnIndex)).ClearContents
        .Range(.Cells(i, StartColumnIndex), .Cells(300, EndColumnIndex)).Borders.LineStyle = xlNone

        Index = 1
        For Each TargetBook In Col
            Call TargetBook.Initialize
            Set Result = TargetBook.CreateSyuhoData(Sheet3.Get週報の集計_開始日(), Sheet3.Get週報の集計_終了日())
            Call TargetBook.Finalize

            For Each CurrentShuhoData In Result
                Call AroundBorderLine(SyuhoSheet, i, EndColumnIndex)
                .Cells(i, StartColumnIndex) = Index
                .Range(.Cells(i, StartColumnIndex + 1), .Cells(i, EndColumnIndex)) = CurrentShuhoData.ToSyuhoRowData

                i = i + 1
                Index = Index + 1
            Next CurrentShuhoData
        Next TargetBook

        SyuhoSheet.Activate
    End With
End Sub


Public Sub 月稼働の集計()
    Dim Col As Collection
    Dim TargetBook As WorkerBook
    Dim Result As TukikadoData
    Dim TukikadoSheet As Worksheet
    Dim i As Long

    Const StartRowIndex As Long = 7
    Const StartColumnIndex As Long = 1
    Const EndColumnIndex As Long = 34

    Dim Parts() As String
    Set TukikadoSheet = ThisWorkbook.Sheets("月稼働")

    Parts = Split(Sheet3.Get月稼働の集計_対象シート(), "・")
    .Range("C2") = Parts(0)
    .Range("C3") = Parts(1)

    Set Col = Sheet1.GetTargetWorkerList()
    i = StartRowIndex

    ' 値と罫線を初期化
    .Range(.Cells(i, StartColumnIndex), .Cells(300, EndColumnIndex)).ClearContents
    .Range(.Cells(i, StartColumnIndex), .Cells(300, EndColumnIndex)).Borders.LineStyle = xlNone

    For Each TargetBook In Col
        Call TargetBook.Initialize
        Set Result = TargetBook.GetTukiKado(Sheet3.Get月稼働の集計_対象シート())
        Call TargetBook.Finalize

        ' データの出力
        .Cells(i, StartColumnIndex) = Result.Name
        .Cells(i, StartColumnIndex + 1) = "年休"
        .Range(.Cells(i, StartColumnIndex + 2), .Cells(i, EndColumnIndex)) = Result.NenkyuYotei
        i = i + 1

        .Cells(i, StartColumnIndex + 1) = "予定時間"
        .Range(.Cells(i, StartColumnIndex + 2), .Cells(i, EndColumnIndex)) = Result.YoteiTime
        i = i + 1

        .Cells(i, StartColumnIndex + 1) = "実働時間"
        .Range(.Cells(i, StartColumnIndex + 2), .Cells(i, EndColumnIndex)) = Result.JitudoTime
        i = i + 1
    Next TargetBook

    .Activate
End Sub


Public Sub 月稼働の集計()
    Dim Col As Collection
    Dim TargetBook As WorkerBook
    Dim Result As TukiKadoData
    Dim TukiKadoSheet As Worksheet
    Dim i As Long

    Const StartRowIndex As Long = 7
    Const StartColumnIndex As Long = 1
    Const EndColumnIndex As Long = 34

    Dim Parts() As String

    ' 月稼働シートの取得
    Set TukiKadoSheet = ThisWorkbook.Sheets("月稼働")
    With TukiKadoSheet
        ' シート名を分割して表示
        Parts = Split(Sheet3.Get月稼働の集計_対象シート(), "-")
        .Range("C2") = Parts(0)
        .Range("C3") = Parts(1)

        ' ワーカーリストの取得
        Set Col = Sheet1.GetTargetWorkerList()
        i = StartRowIndex

        ' 初期化: 値と罫線のクリア
        .Range(.Cells(i, StartColumnIndex), .Cells(300, EndColumnIndex)).ClearContents
        .Range(.Cells(i, StartColumnIndex), .Cells(300, EndColumnIndex)).Borders.LineStyle = xlNone

        ' 各ワーカーのデータ処理
        For Each TargetBook In Col
            Call TargetBook.Initialize
            Set Result = TargetBook.GetTukiKado(Sheet3.Get月稼働の集計_対象シート())
            Call TargetBook.Finalize

            ' 罫線の設定 (上部)
            Call TopBorderLine(TukiKadoSheet, i)

            ' 名前の設定
            .Cells(i, StartColumnIndex) = Result.Name

            ' 年休付与
            .Cells(i, StartColumnIndex + 1) = "年休付与"
            .Range(.Cells(i, StartColumnIndex + 2), .Cells(i, EndColumnIndex - 1)) = Result.NenkyuYotei
            i = i + 1
            
            ' MiddiumBorderLineの呼び出し
            Call MiddiumBorderLine(TukiKadoSheet, i)

            ' 予定時間
            .Cells(i, StartColumnIndex + 1) = "予定時間"
            .Range(.Cells(i, StartColumnIndex + 2), .Cells(i, EndColumnIndex)) = Result.YoteiTime

            ' 実働時間
            i = i + 1
            .Cells(i, StartColumnIndex + 1) = "実働時間"
            .Range(.Cells(i, StartColumnIndex + 2), .Cells(i, EndColumnIndex)) = Result.JitudoTime

            ' BottomBorderLineの呼び出し
            Call BottomBorderLine(TukiKadoSheet, i)

            ' 次のワーカーの処理
            i = i + 1
        Next TargetBook

        ' シートをアクティブ化
        .Activate
    End With
End Sub



' 週報用の境界線を設定
Private Sub AroundBorderLine(TargetSheet As Worksheet, Index As Long, EndColumnIndex As Long)
    Dim TargetRange As Range
    
    ' 対象範囲を設定
    Set TargetRange = TargetSheet.Range(TargetSheet.Cells(Index, 1), TargetSheet.Cells(Index, EndColumnIndex))
    
    ' 対角線の線を非表示に設定
    TargetRange.Borders(xlDiagonalDown).LineStyle = xlNone
    TargetRange.Borders(xlDiagonalUp).LineStyle = xlNone

    ' 左側の境界線を設定
    With TargetRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    ' 上側の境界線を設定
    With TargetRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    ' 下側の境界線を設定
    With TargetRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    ' 右側の境界線を設定
    With TargetRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    ' 垂直線の内部の境界線を設定
    With TargetRange.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    ' 水平線の内部の境界線を設定
    With TargetRange.Borders(xlInsideHorizontal)
        .LineStyle = xlNone
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub


' 月稼働用の上部の境界線を設定
Private Sub TopBorderLine(TargetSheet As Worksheet, Index As Long)
    Dim TargetRange As Range

    ' 対象範囲を設定
    Set TargetRange = TargetSheet.Range(TargetSheet.Cells(Index, 1), TargetSheet.Cells(Index, 34))

    ' 対角線の線を非表示に設定
    TargetRange.Borders(xlDiagonalDown).LineStyle = xlNone
    TargetRange.Borders(xlDiagonalUp).LineStyle = xlNone

    ' 左側の境界線を設定
    With TargetRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With

    ' 上側の境界線を設定
    With TargetRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With

    ' 下側の境界線を設定
    With TargetRange.Borders(xlEdgeBottom)
        .LineStyle = xlNone
    End With

    ' 右側の境界線を設定
    With TargetRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With

    ' 垂直方向の内部の境界線を設定
    With TargetRange.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With

    ' 水平方向の内部の境界線を設定
    With TargetRange.Borders(xlInsideHorizontal)
        .LineStyle = xlNone
    End With
End Sub

' 月稼働用の中央の境界線を設定
Private Sub MiddiumBorderLine(TargetSheet As Worksheet, Index As Long)
    Dim TargetRange As Range

    ' 対象範囲を設定
    Set TargetRange = TargetSheet.Range(TargetSheet.Cells(Index, 1), TargetSheet.Cells(Index, 34))

    ' 対角線の線を非表示に設定
    TargetRange.Borders(xlDiagonalDown).LineStyle = xlNone
    TargetRange.Borders(xlDiagonalUp).LineStyle = xlNone

    ' 左側の境界線を設定
    With TargetRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With

    ' 上側の境界線を設定
    With TargetRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With

    ' 下側の境界線を設定
    With TargetRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With

    ' 右側の境界線を設定
    With TargetRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With

    ' 垂直方向の内部の境界線を設定
    With TargetRange.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With

    ' 水平方向の内部の境界線を設定
    With TargetRange.Borders(xlInsideHorizontal)
        .LineStyle = xlNone
    End With
End Sub


' 月稼働用の下部の境界線を設定
Private Sub BottomBorderLine(TargetSheet As Worksheet, Index As Long)
    Dim TargetRange As Range

    ' 対象範囲を設定
    Set TargetRange = TargetSheet.Range(TargetSheet.Cells(Index, 1), TargetSheet.Cells(Index, 34))

    ' 対角線の線を非表示に設定
    TargetRange.Borders(xlDiagonalDown).LineStyle = xlNone
    TargetRange.Borders(xlDiagonalUp).LineStyle = xlNone

    ' 左側の境界線を設定
    With TargetRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With

    ' 上側の境界線を設定
    With TargetRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    ' 下側の境界線を設定
    With TargetRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With

    ' 右側の境界線を設定
    With TargetRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With

    ' 垂直方向の内部の境界線を設定
    With TargetRange.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    ' 水平方向の内部の境界線を非表示
    With TargetRange.Borders(xlInsideHorizontal)
        .LineStyle = xlNone
    End With
End Sub
