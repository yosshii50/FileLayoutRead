Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

'レコード

'構成
'　レコードタイプ
'　　+-フィールド
'　　+-フィールド

Public Class FLR_RecordType

    'レコードタイプ名
    Private _RecordTypeName As String = ""
    Public Property RecordTypeName() As String
        Get
            Return _RecordTypeName
        End Get
        Set(ByVal value As String)
            _RecordTypeName = value.Trim
        End Set
    End Property

    'フィールドの追加
    Public Sub AddField()
        Call AddField("", 0, FLR_Field.FieldType_Enum.None, "")
    End Sub
    Public Sub AddField(ByVal FieldName As String, ByVal FieldLength As Integer, ByVal FieldType As FLR_Field.FieldType_Enum, ByVal SamplePattern As String)

        Dim GetFLR_Field As New FLR_Field

        With GetFLR_Field
            .FieldName = FieldName
            .FieldLength = FieldLength
            .FieldType = FieldType
            .SamplePattern = SamplePattern
        End With

        Call AddField(GetFLR_Field)

    End Sub
    Public Sub AddField(ByRef GetFLR_Field As FLR_Field)

        If _Fields Is Nothing Then
            ReDim _Fields(0)
        Else
            ReDim Preserve _Fields(_Fields.Count)
        End If

        _Fields(_Fields.Count - 1) = GetFLR_Field

    End Sub

    'フィールド
    Private _Fields() As FLR_Field
    Public Property Fields() As FLR_Field()
        Get
            Return _Fields
        End Get
        Set(ByVal value As FLR_Field())
            _Fields = value
        End Set
    End Property

End Class
