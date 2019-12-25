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

    'DBテーブル名
    Private _DB_TableName As String = ""
    Public Property DB_TableName() As String
        Get
            Return _DB_TableName
        End Get
        Set(ByVal value As String)
            _DB_TableName = value.Trim
        End Set
    End Property

    'テーブル作成時の追加文字列
    Private _DB_TableCreateAddStr As String
    Public Property DB_TableCreateAddStr() As String
        Get
            Return _DB_TableCreateAddStr
        End Get
        Set(ByVal value As String)
            _DB_TableCreateAddStr = value
        End Set
    End Property
    Public Sub DB_TableCreateAddStr_Add(ByVal AddStr As String)
        If AddStr <> "" Then
            _DB_TableCreateAddStr = _DB_TableCreateAddStr & AddStr
        End If
    End Sub

    'フィールドの追加
    Public Sub AddField(ByVal FieldName As String, ByVal FieldLength As Integer, ByVal FieldType As FLR_FieldType.FieldType_Enum, ByVal SamplePattern As String)

        Dim GetFLR_Field As New FLR_FieldType

        With GetFLR_Field
            .FieldName = FieldName
            .FieldLength = FieldLength
            .FieldType = FieldType
            .SamplePattern = SamplePattern
        End With

        Call AddField(GetFLR_Field)

    End Sub
    Public Sub AddField(ByRef GetFLR_Field As FLR_FieldType)

        If _Fields Is Nothing Then
            ReDim _Fields(0)
        Else
            ReDim Preserve _Fields(_Fields.Count)
        End If

        _Fields(_Fields.Count - 1) = GetFLR_Field

    End Sub

    'フィールド
    Private _Fields() As FLR_FieldType
    Public Property Fields() As FLR_FieldType()
        Get
            Return _Fields
        End Get
        Set(ByVal value As FLR_FieldType())
            _Fields = value
        End Set
    End Property

    '識別パターンと一致しているか確認
    Public Function CheckPattern(ByRef WrkRecordData As Byte()) As Boolean

        For Each WrkField As FLR_FieldType In _Fields

            Dim WrkStPos As Integer

            '作業用配列に転送
            Dim WrkRD() As Byte
            ReDim WrkRD(WrkField.FieldLength - 1)
            For WrkIdx As Integer = 0 To WrkField.FieldLength - 1
                WrkRD(WrkIdx) = WrkRecordData(WrkIdx + WrkStPos)
            Next

            'サンプルパターンと一致しているか確認
            If WrkField.CheckSamplePattern(WrkRD) = True Then
                Return True
            End If

            WrkStPos = WrkStPos + WrkField.FieldLength
        Next

        Return False
    End Function






End Class
