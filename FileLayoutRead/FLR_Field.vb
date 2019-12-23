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
    Public Function GetSampleData() As String

        Static WrkIdx As Integer = 0

        Dim WrkStrs() As String = Split(_SamplePattern, ",")

        If WrkIdx >= WrkStrs.Count Then
            WrkIdx = 0
        End If

        Dim WrkStr As String = WrkStrs(WrkIdx)
        Select Case WrkStr
            Case "YYYYMMDD"
                WrkStr = "20191231"
            Case "HHNN"
                WrkStr = "2359"
            Case "SP"
                WrkStr = Space(FieldLength)
            Case ""
                WrkStr = Space(FieldLength)
            Case "CRLF"
                WrkStr = vbCrLf
            Case Else
                If Strings.Left(WrkStr, 1) = "[" And Strings.Right(WrkStr, 1) = "]" Then
                    WrkStr = Strings.Mid(WrkStr, 2, Len(WrkStr) - 2)
                End If
                If Strings.Left(WrkStr, 1) = """" And Strings.Right(WrkStr, 1) = """" Then
                    WrkStr = Strings.Mid(WrkStr, 2, Len(WrkStr) - 2)
                End If
        End Select

        WrkIdx = WrkIdx + 1

        Return WrkStr

    End Function

End Class
