Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

Public Class FLR_DB

    Public Sub New(ByRef WrkFLR_File As FLR_File)
        BaseFLR_File = WrkFLR_File
    End Sub
    Private BaseFLR_File As FLR_File

    'テーブル
    Private _DBTables() As FLR_DBTable
    Public ReadOnly Property DBTables() As FLR_DBTable()
        Get
            Return _DBTables
        End Get
    End Property

    'テーブルの追加
    Public Function DBTableAdd(ByVal TableName As String) As FLR_DBTable

        Dim WrkDBTable As FLR_DBTable = Nothing

        If TableName = "" Then
            Return Nothing
        End If

        If Not _DBTables Is Nothing Then
            '既に同じ名前があるか確認
            For Each LoopDBTable As FLR_DBTable In _DBTables
                If LoopDBTable.TableName = TableName Then
                    WrkDBTable = LoopDBTable
                    Exit For
                End If
            Next
        End If

        If WrkDBTable Is Nothing Then
            '新規に追加

            WrkDBTable = New FLR_DBTable

            With WrkDBTable
                .TableName = TableName
            End With

            If _DBTables Is Nothing Then
                ReDim _DBTables(0)
            Else
                ReDim Preserve _DBTables(_DBTables.Count)
            End If

            _DBTables(_DBTables.Count - 1) = WrkDBTable

        End If

        Return WrkDBTable
    End Function

    'テーブル作成時の追加文字列
    Private _TableCreateAddStr As String = ""
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
            _TableCreateAddStr = _TableCreateAddStr & AddStr & vbCrLf
        End If
    End Sub

End Class
