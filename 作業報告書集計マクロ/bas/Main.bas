Attribute VB_Name = "Main"
Option Explicit

Public Sub �w��N���V�[�g�̍쐬()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanUp

    ' ���s�J�n
    Dim Exe As New WorkerExecutor
    Call Exe.�w��N���V�[�g�̍쐬

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    If Err.Number <> 0 Then
        Call MsgBox("�G���[������: " & Err.Number & ", " & Err.Source & ", " & Err.Description)
    End If
    Call MsgBox("�������܂���")
End Sub

Public Sub ��ƕ񍐁Y���̐ݒ�()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanUp

    ' ���s�J�n
    Dim Exe As New WorkerExecutor
    Call Exe.��ƕ񍐁Y���̐ݒ�

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    If Err.Number <> 0 Then
        Call MsgBox("�G���[������: " & Err.Number & ", " & Err.Source & ", " & Err.Description)
    End If
    Call MsgBox("�������܂���")
End Sub

Public Sub ��ƕ񍐁Y()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanUp

    ' ���s�J�n
    Dim Exe As New WorkerExecutor
    Call Exe.��ƕ񍐁Y

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    If Err.Number <> 0 Then
        Call MsgBox("�G���[������: " & Err.Number & ", " & Err.Source & ", " & Err.Description)
    End If
    Call MsgBox("�������܂���")
End Sub

Public Sub �������Ԃ̍����o��()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanUp

    ' ���s�J�n
    Dim Result As String
    Dim Exe As New WorkerExecutor
    Result = Exe.�������Ԃ̍����o��

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    If Err.Number <> 0 Then
        Call MsgBox("�G���[������: " & Err.Number & ", " & Err.Source & ", " & Err.Description)
    End If

    TextForm.TextBox.Text = Result
    TextForm.Show
End Sub

Public Sub �j�����X�g�̍X�V()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanUp

    ' ���s�J�n
    Dim Exe As New WorkerExecutor
    Exe.�j�����X�g�̍X�V

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    If Err.Number <> 0 Then
        Call MsgBox("�G���[������: " & Err.Number & ", " & Err.Source & ", " & Err.Description)
    End If
    Call MsgBox("�������܂���")
End Sub

Public Sub �T��̏W�v()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanUp

    ' ���s�J�n
    Dim Exe As New WorkerExecutor
    Exe.�T��̏W�v

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    If Err.Number <> 0 Then
        Call MsgBox("�G���[������: " & Err.Number & ", " & Err.Source & ", " & Err.Description)
    End If
    Call MsgBox("�������܂���")
End Sub

Public Sub ���ғ��̏W�v()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanUp

    ' ���s�J�n
    Dim Exe As New WorkerExecutor
    Exe.���ғ��̏W�v

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    If Err.Number <> 0 Then
        Call MsgBox("�G���[������: " & Err.Number & ", " & Err.Source & ", " & Err.Description)
    End If
    Call MsgBox("�������܂���")
End Sub
