Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

'ファイル

'構成
'　ファイル
'　　+-レコード
'　　| 　+-フィールド
'　　+-レコード
'　　  　+-フィールド
'　　  　+-フィールド

Public Class FLR_File

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

    'レコード構成
    Private Structure RecordConstitution_Str
        Dim Record As FLR_Record
        Dim ParentRecord As FLR_Record
        Dim ChildRecord() As FLR_Record
    End Structure
    Private RecordConstitution() As RecordConstitution_Str

    'レコード種類の追加
    Public Sub AddRecordType()
        Call AddRecordType("")
    End Sub
    Public Sub AddRecordType(ByVal RecordName As String)

        If RecordConstitution Is Nothing Then
            ReDim RecordConstitution(0)
        Else
            ReDim Preserve RecordConstitution(RecordConstitution.Count)
        End If

        With RecordConstitution(RecordConstitution.Count - 1)
            .Record = New FLR_Record
            .Record.RecordName = RecordName
        End With

    End Sub

    'レコード種類の件数
    Public ReadOnly Property RecordTypeCount() As Integer
        Get
            If RecordConstitution Is Nothing Then
                Return 0
            Else
                Return RecordConstitution.Count
            End If
        End Get
    End Property

    '最大のレコード種類
    Public ReadOnly Property LastRecord() As FLR_Record
        Get
            If RecordConstitution Is Nothing Then
                Return Nothing
            Else
                Return RecordConstitution(RecordConstitution.Count - 1).Record
            End If
        End Get
    End Property

    'レコード
    Public ReadOnly Property Records(ByVal Index As Integer) As FLR_Record
        Get
            Return RecordConstitution(Index).Record
        End Get
    End Property

End Class
