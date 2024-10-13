VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SyuhoData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' プライベート変数の宣言
Private m_Name As String               ' 名前
Private m_Kikan As String              ' 期間
Private m_SyukadoJikan As Double       ' 週稼働時間
Private m_SyukkinNissu As Double       ' 出勤日数
Private m_JitukadoTotal As Double      ' 実働時間合計
Private m_TougetukadoMikomi As Double  ' 当月稼働見込み
Private m_SeikyuJikan As Double        ' 請求時間
Private m_ZengetuTyosei As Double      ' 前月調整
Private m_YokugetuTyosei As Double     ' 翌月調整
Private m_ZangyoJikan As Double        ' 残業時間
Private m_SintyokuJokyo As String      ' 進捗状況
Private m_FukaJokyo As String          ' 負荷状況
Private m_Kadai As String              ' 課題
Private m_Hitokoto As String           ' 一言
Private m_NenkyuYotei As String        ' 年休予定
Private m_SagyoNaiyou As String        ' 作業内容

' プロパティの定義

' 名前プロパティ
Public Property Let Name(Value As String)
    m_Name = Value
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

' 期間プロパティ
Public Property Let Kikan(Value As String)
    m_Kikan = Value
End Property

Public Property Get Kikan() As String
    Kikan = m_Kikan
End Property

' 週稼働時間プロパティ
Public Property Let SyukadoJikan(Value As Double)
    m_SyukadoJikan = Value
End Property

Public Property Get SyukadoJikan() As Double
    SyukadoJikan = m_SyukadoJikan
End Property

' 出勤日数プロパティ
Public Property Let SyukkinNissu(Value As Double)
    m_SyukkinNissu = Value
End Property

Public Property Get SyukkinNissu() As Double
    SyukkinNissu = m_SyukkinNissu
End Property

' 実働時間合計プロパティ
Public Property Let JitukadoTotal(Value As Double)
    m_JitukadoTotal = Value
End Property

Public Property Get JitukadoTotal() As Double
    JitukadoTotal = m_JitukadoTotal
End Property

' 当月稼働見込みプロパティ
Public Property Let TougetukadoMikomi(Value As Double)
    m_TougetukadoMikomi = Value
End Property

Public Property Get TougetukadoMikomi() As Double
    TougetukadoMikomi = m_TougetukadoMikomi
End Property

' 請求時間プロパティ
Public Property Let SeikyuJikan(Value As Double)
    m_SeikyuJikan = Value
End Property

Public Property Get SeikyuJikan() As Double
    SeikyuJikan = m_SeikyuJikan
End Property

' 前月調整プロパティ
Public Property Let ZengetuTyosei(Value As Double)
    m_ZengetuTyosei = Value
End Property

Public Property Get ZengetuTyosei() As Double
    ZengetuTyosei = m_ZengetuTyosei
End Property

' 翌月調整プロパティ
Public Property Let YokugetuTyosei(Value As Double)
    m_YokugetuTyosei = Value
End Property

Public Property Get YokugetuTyosei() As Double
    YokugetuTyosei = m_YokugetuTyosei
End Property

' 残業時間プロパティ
Public Property Let ZangyoJikan(Value As Double)
    m_ZangyoJikan = Value
End Property

Public Property Get ZangyoJikan() As Double
    ZangyoJikan = m_ZangyoJikan
End Property

' 進捗状況プロパティ
Public Property Let SintyokuJokyo(Value As String)
    m_SintyokuJokyo = Value
End Property

Public Property Get SintyokuJokyo() As String
    SintyokuJokyo = m_SintyokuJokyo
End Property

' 負荷状況プロパティ
Public Property Let FukaJokyo(Value As String)
    m_FukaJokyo = Value
End Property

Public Property Get FukaJokyo() As String
    FukaJokyo = m_FukaJokyo
End Property

' 課題プロパティ
Public Property Let Kadai(Value As String)
    m_Kadai = Value
End Property

Public Property Get Kadai() As String
    Kadai = m_Kadai
End Property

' 一言プロパティ
Public Property Let Hitokoto(Value As String)
    m_Hitokoto = Value
End Property

Public Property Get Hitokoto() As String
    Hitokoto = m_Hitokoto
End Property

' 年休予定プロパティ
Public Property Let NenkyuYotei(Value As String)
    m_NenkyuYotei = Value
End Property

Public Property Get NenkyuYotei() As String
    NenkyuYotei = m_NenkyuYotei
End Property

' 作業内容プロパティ
Public Property Let SagyoNaiyou(Value As String)
    m_SagyoNaiyou = Value
End Property

Public Property Get SagyoNaiyou() As String
    SagyoNaiyou = m_SagyoNaiyou
End Property

' データを配列で取得するメソッド
Public Function ToSyuhoRowData() As Variant
    Dim Result(1 To 16) As Variant
    Result(1) = m_Name
    Result(2) = m_Kikan
    Result(3) = m_SyukadoJikan
    Result(4) = m_SyukkinNissu
    Result(5) = m_JitukadoTotal
    Result(6) = m_TougetukadoMikomi
    Result(7) = m_SeikyuJikan
    Result(8) = m_ZengetuTyosei
    Result(9) = m_YokugetuTyosei
    Result(10) = m_ZangyoJikan
    Result(11) = m_SintyokuJokyo
    Result(12) = m_FukaJokyo
    Result(13) = m_Kadai
    Result(14) = m_Hitokoto
    Result(15) = m_NenkyuYotei
    Result(16) = m_SagyoNaiyou
    ToSyuhoRowData = Result
End Function
