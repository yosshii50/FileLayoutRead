Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

'フィールドタイプ
Public Class FLR_FieldType

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
        BinX
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

    'DB用フィールド名
    Private _DBFieldName As String = ""
    Public Property DBFieldName() As String
        Get
            Return _DBFieldName
        End Get
        Set(ByVal value As String)

            If value = "-" Then
                _DBFieldName = ""
            Else
                _DBFieldName = value
            End If

        End Set
    End Property

    'サンプルデータ取得
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

                WrkStr = Strings.Left(WrkStr & Space(FieldLength), FieldLength)

        End Select

        WrkIdx = WrkIdx + 1

        Return WrkStr

    End Function

    'サンプルパターンと一致しているか確認
    Public Function CheckSamplePattern(ByRef WrkRecordData As Byte()) As Boolean

        Dim WrkStrs() As String = Split(_SamplePattern, ",")

        For Each WrkStr As String In WrkStrs

            If Strings.Left(WrkStr, 1) = "[" And Strings.Right(WrkStr, 1) = "]" Then
                WrkStr = Strings.Mid(WrkStr, 2, Len(WrkStr) - 2)

                Dim CnvStr As String
                CnvStr = System.Text.Encoding.GetEncoding(932).GetString(WrkRecordData)

                If WrkStr = CnvStr Then
                    Return True
                End If

            End If

        Next

        Return False
    End Function


End Class
