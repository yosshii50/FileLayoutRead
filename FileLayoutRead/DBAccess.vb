Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

Public Module DBAccess

    Private DBFactory As System.Data.Common.DbProviderFactory
    Private DBConnection As System.Data.Common.DbConnection

    '変換処理実行
    Public Function Start(ByVal WrkFLR_File As FLR_File) As Boolean

        'DBに接続
        If DB_Connect(WrkFLR_File.DBConnect) = False Then
            Return False
        End If

        'トランザクション実行
        Call StartTransaction(WrkFLR_File)

        'DB切断
        If DB_DisConnect() = False Then
            Return False
        End If

        Return True
    End Function

    'トランザクション実行
    Private Function StartTransaction(ByVal WrkFLR_File As FLR_File) As Boolean

        Dim WrkTransaction As System.Data.Common.DbTransaction = DBConnection.BeginTransaction

        '変換処理実行
        If StartCnv(WrkFLR_File, WrkTransaction) = True Then
            WrkTransaction.Commit()
        Else
            WrkTransaction.Rollback()
        End If

        WrkTransaction.Dispose()

        Return True
    End Function

    '変換処理実行
    Private Function StartCnv(ByVal WrkFLR_File As FLR_File, ByRef WrkTransaction As System.Data.Common.DbTransaction) As Boolean

        Dim cmd1 As System.Data.Common.DbCommand = DBFactory.CreateCommand()
        cmd1.Connection = DBConnection
        cmd1.Transaction = WrkTransaction
        cmd1.CommandType = CommandType.Text

        '必要テーブル生成
        For Each WrkDBTable As FLR_DBTable In WrkFLR_File.DB.DBTables
            If Not WrkDBTable.DBFields Is Nothing Then
                Dim WrkStr As String

                '一度テーブル削除
                WrkStr = ""
                WrkStr = WrkStr & "DROP TABLE " & WrkDBTable.TableName & vbCrLf
                cmd1.CommandText = WrkStr
                Try
                    cmd1.ExecuteNonQuery()
                Catch ex As Exception
                    If InStr(ex.ToString, "ORA-00942") > 0 Then 'OracleClientを組み込まないため、文字列で判断
                        'まだテーブルが存在しない場合
                        Exit Try '処理を続行
                    End If
                    MsgBox(ex.ToString)
                    Return False
                End Try

                '一度テーブル作成
                WrkStr = ""
                WrkStr = WrkStr & "CREATE TABLE " & WrkDBTable.TableName & vbCrLf
                For WrkIdx As Integer = 0 To WrkDBTable.DBFields.Count - 1
                    If WrkIdx = 0 Then
                        WrkStr = WrkStr & "("
                    Else
                        WrkStr = WrkStr & ","
                    End If
                    WrkStr = WrkStr & " " & WrkDBTable.DBFields(WrkIdx).DBFieldName
                    WrkStr = WrkStr & " " & WrkDBTable.DBFields(WrkIdx).Parameter
                    WrkStr = WrkStr & " " & WrkDBTable.DBFields(WrkIdx).Comment
                    WrkStr = WrkStr & vbCrLf
                Next
                WrkStr = WrkStr & ")" & vbCrLf
                If WrkDBTable.TableCreateAddStr = "" Then
                    If WrkFLR_File.DB.TableCreateAddStr <> "" Then
                        WrkStr = WrkStr & WrkFLR_File.DB.TableCreateAddStr
                    End If
                Else
                    WrkStr = WrkStr & WrkDBTable.TableCreateAddStr
                End If

                cmd1.CommandText = WrkStr
                cmd1.ExecuteNonQuery()
            End If
        Next

        For Each WrkRecordData As FLR_RecordData In WrkFLR_File.Data.RecordDatas
            If WrkRecordData.RecordType.DB_TableName <> "" Then

                'Insert文用SQL生成
                Dim WrkStr As String = WrkRecordData.GetInsertTableListStr()

                Clipboard.SetText(WrkStr) 'デバッグ用

                cmd1.CommandText = WrkStr
                cmd1.ExecuteNonQuery()

            End If
        Next

        Return True
    End Function

    Private Function OracleConnectStr2ConnectionString(ByVal OraConStr As String) As String

        '"USER ID=user;PASSWORD=pass;DATA SOURCE=//server:1521/orcl/"

        Dim StrIdx As Integer
        Dim WrkSVStr As String '接続先
        Dim WrkUSPSStr As String 'ユーザーとパスワード
        Dim WrkUSStr As String 'ユーザー
        Dim WrkPSStr As String 'パスワード

        '接続先の分離
        StrIdx = InStr(OraConStr, "@")
        If StrIdx > 0 Then
            '[@]が存在する場合

            WrkUSPSStr = OraConStr.Substring(0, StrIdx - 1)
            WrkSVStr = OraConStr.Substring(StrIdx)

        Else
            '[@]が存在しない場合
            '接続先は無しのまま、ユーザーとパスワードを設定
            WrkUSPSStr = OraConStr
            WrkSVStr = ""
        End If

        'ユーザーとパスワードの分離
        StrIdx = InStr(WrkUSPSStr, "/")
        If StrIdx > 0 Then
            '[/]が存在する場合

            WrkPSStr = WrkUSPSStr.Substring(StrIdx)
            WrkUSStr = WrkUSPSStr.Substring(0, StrIdx - 1)

        Else
            '[/]が存在しない場合
            'ユーザーとパスワードを同じにする

            WrkPSStr = WrkUSPSStr
            WrkUSStr = WrkUSPSStr

        End If

        '帰り値セット
        If WrkSVStr = "" Then
            Return "USER ID=" & WrkUSStr & ";PASSWORD=" & WrkPSStr & ""
        Else
            Return "USER ID=" & WrkUSStr & ";PASSWORD=" & WrkPSStr & ";DATA SOURCE=" & WrkSVStr
        End If

    End Function

    Private Function DB_Connect(ByVal ConnectStr As String) As Boolean

        DBFactory = System.Data.Common.DbProviderFactories.GetFactory("Oracle.DataAccess.Client")
        DBConnection = DBFactory.CreateConnection()
        DBConnection.ConnectionString = OracleConnectStr2ConnectionString(ConnectStr)
        DBConnection.Open()

        Return True
    End Function

    Private Function DB_DisConnect() As Boolean

        DBConnection.Close()
        DBConnection.Dispose()

        Return True
    End Function

End Module
