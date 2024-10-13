VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Enum RowIndex
    指定年月シートの作成 = 3
    作業報告〆日の設定
    作業報告〆
    実働時間の差分出力
    祝日リストの更新
    週報の集計
    月稼働の集計
    NTTD稼働集計
End Enum


Private Enum ColumnIndex
    対象シート = 5
    設定1 = 7
    設定2 = 9
    設定3 = 11
End Enum

Public Function Get指定年月シートの作成_対象シート() As String
    Get指定年月シートの作成_対象シート = Sheet3.Cells(RowIndex.指定年月シートの作成, ColumnIndex.対象シート).Value
End Function

Public Function Get指定年月シートの作成_対象ブック() As String
    Get指定年月シートの作成_対象ブック = Sheet3.Cells(RowIndex.指定年月シートの作成, ColumnIndex.設定1).Value
End Function

Public Function Get作業報告〆日の設定_作業報告〆日() As Date
    Get作業報告〆日の設定_作業報告〆日 = Sheet3.Cells(RowIndex.作業報告〆日の設定, ColumnIndex.設定1).Value
End Function

Public Function Get作業報告〆_対象シート() As String
    Get作業報告〆_対象シート = Sheet3.Cells(RowIndex.作業報告〆, ColumnIndex.対象シート).Value
End Function

Public Function Get実働時間の差分出力_対象シート() As String
    Get実働時間の差分出力_対象シート = Sheet3.Cells(RowIndex.実働時間の差分出力, ColumnIndex.対象シート).Value
End Function

Public Function Get祝日リストの更新_対象シート() As String
    Get祝日リストの更新_対象シート = Sheet3.Cells(RowIndex.祝日リストの更新, ColumnIndex.対象シート).Value
End Function

Public Function Get祝日リストの更新_対象ブック() As String
    Get祝日リストの更新_対象ブック = Sheet3.Cells(RowIndex.祝日リストの更新, ColumnIndex.設定1).Value
End Function

Public Function Get週報の集計_開始日() As Date
    Get週報の集計_開始日 = Sheet3.Cells(RowIndex.週報の集計, ColumnIndex.設定1).Value
End Function

Public Function Get週報の集計_終了日() As Date
    Get週報の集計_終了日 = Sheet3.Cells(RowIndex.週報の集計, ColumnIndex.設定2).Value
End Function

Public Function Get月稼働の集計_対象シート() As String
    Get月稼働の集計_対象シート = Sheet3.Cells(RowIndex.月稼働の集計, ColumnIndex.対象シート).Value
End Function
