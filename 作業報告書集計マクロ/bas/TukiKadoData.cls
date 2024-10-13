VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TukiKadoData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' プライベート変数の宣言
Private m_Name As String            ' 名前
Private m_NenkyuYotei As Variant    ' 年休予定
Private m_YoteiTime As Variant      ' 予定時間
Private m_JitudoTime As Variant     ' 実働時間

' 名前プロパティ
Public Property Let Name(Value As String)
    m_Name = Value
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

' 年休予定プロパティ
Public Property Let NenkyuYotei(Value As Variant)
    m_NenkyuYotei = Value
End Property

Public Property Get NenkyuYotei() As Variant
    NenkyuYotei = m_NenkyuYotei
End Property

' 予定時間プロパティ
Public Property Let YoteiTime(Value As Variant)
    m_YoteiTime = Value
End Property

Public Property Get YoteiTime() As Variant
    YoteiTime = m_YoteiTime
End Property

' 実働時間プロパティ
Public Property Let JitudoTime(Value As Variant)
    m_JitudoTime = Value
End Property

Public Property Get JitudoTime() As Variant
    JitudoTime = m_JitudoTime
End Property
