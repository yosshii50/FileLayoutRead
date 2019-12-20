Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

'レコード

'構成
'　レコード
'　　+-フィールド
'　　+-フィールド

Public Class FLR_Record

    'レコード名
    Private _RecordName As String = ""
    Public Property RecordName() As String
        Get
            Return _RecordName
        End Get
        Set(ByVal value As String)
            _RecordName = value
        End Set
    End Property

    'フィールドの追加
    Public Sub AddField(ByVal RecordName As String)

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
