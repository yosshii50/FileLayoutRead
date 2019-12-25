Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

Public Class FLR_DBTable

    'テーブル名
    Private _TableName As String
    Public Property TableName() As String
        Get
            Return _TableName
        End Get
        Set(ByVal value As String)
            _TableName = value
        End Set
    End Property

    'テーブル作成時の追加文字列
    Private _TableCreateAddStr As String
    Public Property TableCreateAddStr() As String
        Get
            Return _TableCreateAddStr
        End Get
        Set(ByVal value As String)
            _TableCreateAddStr = value
        End Set
    End Property
    Public Sub TableCreateAddStrAdd(ByVal AddStr As String)
        If AddStr <> "" Then
            _TableCreateAddStr = _TableCreateAddStr & AddStr
        End If
    End Sub

    'フィールド
    Private _DBFields() As FLR_DBField
    Public ReadOnly Property DBFields() As FLR_DBField()
        Get
            Return _DBFields
        End Get
    End Property

    'フィールドの追加
    Public Function DBFieldAdd(ByVal WrkDBFieldName As String, ByRef WrkFLR_FieldType As FLR_FieldType) As FLR_DBField

        Dim WrkDBField As FLR_DBField = Nothing

        If WrkDBFieldName = "" Then
            Return Nothing
        End If

        If Not _DBFields Is Nothing Then
            '既に同じ名前があるか確認
            For Each LoopDBField As FLR_DBField In _DBFields
                If LoopDBField.DBFieldName = WrkDBFieldName Then
                    WrkDBField = LoopDBField
                    Exit For
                End If
            Next
        End If

        If WrkDBField Is Nothing Then
            '新規に追加

            WrkDBField = New FLR_DBField(WrkFLR_FieldType)

            With WrkDBField
                .DBFieldName = WrkDBFieldName
            End With

            If _DBFields Is Nothing Then
                ReDim _DBFields(0)
            Else
                ReDim Preserve _DBFields(_DBFields.Count)
            End If

            _DBFields(_DBFields.Count - 1) = WrkDBField

        End If

        Return WrkDBField
    End Function

End Class
