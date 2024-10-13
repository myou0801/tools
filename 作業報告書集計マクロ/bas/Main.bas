Attribute VB_Name = "Main"
Option Explicit

Public Sub 指定年月シートの作成()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanUp

    ' 実行開始
    Dim Exe As New WorkerExecutor
    Call Exe.指定年月シートの作成

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    If Err.Number <> 0 Then
        Call MsgBox("エラーが発生: " & Err.Number & ", " & Err.Source & ", " & Err.Description)
    End If
    Call MsgBox("完了しました")
End Sub

Public Sub 作業報告〆日の設定()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanUp

    ' 実行開始
    Dim Exe As New WorkerExecutor
    Call Exe.作業報告〆日の設定

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    If Err.Number <> 0 Then
        Call MsgBox("エラーが発生: " & Err.Number & ", " & Err.Source & ", " & Err.Description)
    End If
    Call MsgBox("完了しました")
End Sub

Public Sub 作業報告〆()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanUp

    ' 実行開始
    Dim Exe As New WorkerExecutor
    Call Exe.作業報告〆

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    If Err.Number <> 0 Then
        Call MsgBox("エラーが発生: " & Err.Number & ", " & Err.Source & ", " & Err.Description)
    End If
    Call MsgBox("完了しました")
End Sub

Public Sub 実働時間の差分出力()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanUp

    ' 実行開始
    Dim Result As String
    Dim Exe As New WorkerExecutor
    Result = Exe.実働時間の差分出力

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    If Err.Number <> 0 Then
        Call MsgBox("エラーが発生: " & Err.Number & ", " & Err.Source & ", " & Err.Description)
    End If

    TextForm.TextBox.Text = Result
    TextForm.Show
End Sub

Public Sub 祝日リストの更新()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanUp

    ' 実行開始
    Dim Exe As New WorkerExecutor
    Exe.祝日リストの更新

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    If Err.Number <> 0 Then
        Call MsgBox("エラーが発生: " & Err.Number & ", " & Err.Source & ", " & Err.Description)
    End If
    Call MsgBox("完了しました")
End Sub

Public Sub 週報の集計()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanUp

    ' 実行開始
    Dim Exe As New WorkerExecutor
    Exe.週報の集計

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    If Err.Number <> 0 Then
        Call MsgBox("エラーが発生: " & Err.Number & ", " & Err.Source & ", " & Err.Description)
    End If
    Call MsgBox("完了しました")
End Sub

Public Sub 月稼働の集計()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanUp

    ' 実行開始
    Dim Exe As New WorkerExecutor
    Exe.月稼働の集計

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    If Err.Number <> 0 Then
        Call MsgBox("エラーが発生: " & Err.Number & ", " & Err.Source & ", " & Err.Description)
    End If
    Call MsgBox("完了しました")
End Sub
