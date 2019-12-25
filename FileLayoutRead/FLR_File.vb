Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

'ファイル

'構成
'　ファイル
'　　+-レコードタイプ
'　　  　+-フィールド

Public Class FLR_File

    Public Data As New FLR_File_Data(Me)
    Public Parameter As New FLR_File_Parameter(Me)
    Public SampleData As New FLR_File_SampleData(Me)
    Public DB As New FLR_DB(Me)

    'レコード長 0:変動
    Private _RecordSize As Integer = 0
    Public Property RecordSize() As Integer
        Get
            Return _RecordSize
        End Get
        Set(ByVal value As Integer)
            _RecordSize = value
        End Set
    End Property

    '改行コード [RecordSize]=0の場合のみ使用
    Private _NewLineCode As String = ""
    Public Property NewLineCode() As String
        Get
            Return _NewLineCode
        End Get
        Set(ByVal value As String)
            _NewLineCode = value
        End Set
    End Property

    'ファイル名
    Private _FileName As String = ""
    Public Property FileName() As String
        Get
            Return _FileName
        End Get
        Set(ByVal value As String)
            _FileName = value
        End Set
    End Property

    'DB接続文字列
    Private _DBConnect As String = ""
    Public Property DBConnect() As String
        Get
            Return _DBConnect
        End Get
        Set(ByVal value As String)
            _DBConnect = value
        End Set
    End Property

    'レコード種類
    Dim _RecordType() As FLR_RecordType

    'レコード種類の追加
    Public Sub RecordTypeAdd()
        Call RecordTypeAdd("")
    End Sub
    Public Sub RecordTypeAdd(ByVal RecordTypeName As String)

        If _RecordType Is Nothing Then
            ReDim _RecordType(0)
        Else
            ReDim Preserve _RecordType(_RecordType.Count)
        End If

        _RecordType(_RecordType.Count - 1) = New FLR_RecordType

        With RecordType(_RecordType.Count - 1)
            .RecordTypeName = RecordTypeName
        End With

    End Sub

    'レコード種類の件数
    Public ReadOnly Property RecordTypeCount() As Integer
        Get
            If _RecordType Is Nothing Then
                Return 0
            Else
                Return _RecordType.Count
            End If
        End Get
    End Property

    'レコード
    Public ReadOnly Property RecordTypes() As FLR_RecordType()
        Get
            Return _RecordType
        End Get
    End Property

    Public ReadOnly Property RecordType(ByVal Index As Integer) As FLR_RecordType
        Get
            Return _RecordType(Index)
        End Get
    End Property

    Public ReadOnly Property RecordTypeMax() As Integer
        Get
            Return _RecordType.Count - 1
        End Get
    End Property

End Class
