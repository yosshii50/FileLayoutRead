Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

'フィールド
Public Class FLR_Field

    'フィールド名
    Private _FieldName As String = ""
    Public Property FieldName() As String
        Get
            Return _FieldName
        End Get
        Set(ByVal value As String)
            _FieldName = value
        End Set
    End Property

    '桁数
    Private _FieldLength As Integer = 0
    Public Property FieldLength() As Integer
        Get
            Return _FieldLength
        End Get
        Set(ByVal value As Integer)
            _FieldLength = value
        End Set
    End Property

    '属性
    Public Enum FieldType_Enum
        None
        StrX
        Num9
    End Enum
    Private _FieldType As FieldType_Enum = FieldType_Enum.None
    Public Property FieldType() As FieldType_Enum
        Get
            Return _FieldType
        End Get
        Set(ByVal value As FieldType_Enum)
            _FieldType = value
        End Set
    End Property

    'サンプルパターン
    Private _SamplePattern As String = ""
    Public Property SamplePattern() As String
        Get
            Return _SamplePattern
        End Get
        Set(ByVal value As String)
            _SamplePattern = value
        End Set
    End Property

End Class
