Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

'設定情報からファイルレイアウト展開
Public Class FLR_File_Parameter

    Public Sub New(ByRef WrkFLR_File As FLR_File)
        BaseFLR_File = WrkFLR_File
    End Sub
    Private BaseFLR_File As FLR_File

    'ファイルを指定して生成実行
    Public Function LoadFile(ByVal ReadFileName As String) As Boolean

        BaseFLR_File.ConfigFileName = StrStEdDel(ReadFileName)

        '設定情報からファイルレイアウト展開
        Dim WrkStr As String
        Try
            WrkStr = System.IO.File.ReadAllText(BaseFLR_File.ConfigFileName, System.Text.Encoding.GetEncoding("Shift_JIS"))
        Catch ex As System.IO.FileNotFoundException
            MsgBox("定義ファイルが見つかりません。")
            Return False
        Catch ex As System.IO.DirectoryNotFoundException
            MsgBox("定義ファイルパスが見つかりません。")
            Return False
        Catch ex As System.NotSupportedException
            MsgBox("定義ファイルパス形式に対応していません。")
            Return False
        Catch ex As Exception
            MsgBox(ex.ToString)
            MsgBox("定義ファイルの読み込みに失敗しました。")
            Return False
        End Try

        If Load(WrkStr) = False Then
            Return False
        End If

        Return True
    End Function

    '生成実行
    Public Function Load(ByVal BaseString As String) As Boolean

        For Each WrkStr As String In Split(BaseString, vbCrLf)

            Dim RetStr As String

            '[RECSIZE]取得
            RetStr = GetSettingInformation(WrkStr, "RECSIZE")
            If Not RetStr Is Nothing Then
                BaseFLR_File.RecordSize = CInt(RetStr)
            End If

            '[FILENAME]取得
            RetStr = GetSettingInformation(WrkStr, "FILENAME")
            If Not RetStr Is Nothing Then
                BaseFLR_File.FileName = StrStEdDel(RetStr)
            End If

            '[SAMPLEPATTERN]取得
            RetStr = GetSettingInformation(WrkStr, "SAMPLEPATTERN")
            If Not RetStr Is Nothing Then
                BaseFLR_File.SampleData.DataPattern = RetStr
            End If

            '[DBCONNECT]取得
            RetStr = GetSettingInformation(WrkStr, "DBCONNECT")
            If Not RetStr Is Nothing Then
                BaseFLR_File.DBConnect = RetStr
            End If

            '[RECTYPE]取得
            RetStr = GetSettingInformation(WrkStr, "RECTYPE")
            If Not RetStr Is Nothing Then
                BaseFLR_File.RecordTypeAdd(RetStr)
            End If

            If BaseFLR_File.RecordTypeCount = 0 Then

                '[TABLEADDSTR]取得
                RetStr = GetSettingInformation(WrkStr, "TABLEADDSTR")
                If Not RetStr Is Nothing Then
                    BaseFLR_File.DB.TableCreateAddStrAdd(RetStr)
                End If

            Else
                '[RECTYPE]が存在する場合

                '[TABLENAME]取得
                RetStr = GetSettingInformation(WrkStr, "TABLENAME")
                If Not RetStr Is Nothing Then
                    BaseFLR_File.RecordTypes(BaseFLR_File.RecordTypeMax).DB_TableName = RetStr
                End If

                '[TABLEADDSTR]取得
                RetStr = GetSettingInformation(WrkStr, "TABLEADDSTR")
                If Not RetStr Is Nothing Then
                    BaseFLR_File.RecordTypes(BaseFLR_File.RecordTypeMax).DB_TableCreateAddStr_Add(RetStr)
                End If

                'フィールド情報取得
                Dim GetFLR_Field As FLR_FieldType = GetFieldInformation(WrkStr)
                If Not GetFLR_Field Is Nothing Then

                    'フィールドの追加
                    BaseFLR_File.RecordTypes(BaseFLR_File.RecordTypeMax).AddField(GetFLR_Field)

                End If

            End If

        Next

        'DBTableの生成
        Call CreateDBTable(BaseFLR_File)

        Return True
    End Function

    'DBTableの生成
    Private Sub CreateDBTable(ByRef WrkFLR_File As FLR_File)

        If WrkFLR_File.RecordTypes Is Nothing Then
            Exit Sub
        End If

        For Each WrkRecordType As FLR_RecordType In WrkFLR_File.RecordTypes

            Dim WrkDBTable As FLR_DBTable

            WrkDBTable = WrkFLR_File.DB.DBTableAdd(WrkRecordType.DB_TableName)

            If WrkRecordType.DB_TableCreateAddStr <> "" Then
                WrkDBTable.TableCreateAddStrAdd(WrkRecordType.DB_TableCreateAddStr)
            End If

            If Not WrkDBTable Is Nothing Then

                For Each WrkFieldType As FLR_FieldType In WrkRecordType.Fields

                    If WrkFieldType.DBFieldName <> "" Then
                        WrkDBTable.DBFieldAdd(WrkFieldType.DBFieldName, WrkFieldType)
                    End If

                Next

            End If

        Next

    End Sub

    'フィールド情報取得
    Private Function GetFieldInformation(ByVal WrkStr As String) As FLR_FieldType

        Dim GetFLR_Field As New FLR_FieldType

        '区切り文字をTABに共通化
        Dim DisassStr As String = GetDeleteComment(WrkStr)
        DisassStr = DisassStr.Replace("　", vbTab)
        DisassStr = DisassStr.Replace(" ", vbTab)

        If GetDeleteComment(WrkStr).Trim = "" Then
            Return Nothing
        End If

        '[TAB]で分解
        Dim ParNo As Integer = 0 'パラメータ番号
        For Each SglStr As String In Split(DisassStr, vbTab)

            If SglStr <> "" Then
                '空白で無い場合のみ
                ParNo = ParNo + 1

                Select Case ParNo

                    Case 1 '１つ目の項目 / フィールドの桁数
                        If IsNumeric(SglStr) = True Then
                            '最初の項目が数字であれば
                            GetFLR_Field.FieldLength = CInt(SglStr)
                        Else
                            Return Nothing
                        End If

                    Case 2 '２つ目の項目 / フィールドの属性
                        Select Case SglStr
                            Case "X"
                                GetFLR_Field.FieldType = FLR_FieldType.FieldType_Enum.StrX
                            Case "9"
                                GetFLR_Field.FieldType = FLR_FieldType.FieldType_Enum.Num9
                            Case "B"
                                GetFLR_Field.FieldType = FLR_FieldType.FieldType_Enum.BinX
                            Case Else
                                Return Nothing
                        End Select

                    Case 3 '３つ目の項目 / DB用フィールド名
                        GetFLR_Field.DBFieldName = SglStr

                    Case 4 '４つ目の項目 / サンプルパターン
                        GetFLR_Field.SamplePattern = SglStr

                End Select

            End If

        Next

        'フィールド名
        GetFLR_Field.FieldName = GetComment(WrkStr)

        Return GetFLR_Field
    End Function

    'コメント部分削除
    Private Function GetDeleteComment(ByVal WrkStr As String) As String

        If InStrComment(WrkStr) <> 0 Then
            WrkStr = Strings.Left(WrkStr, InStrComment(WrkStr) - 1)
        End If

        Return WrkStr
    End Function

    'コメント部分取得
    Private Function GetComment(ByVal WrkStr As String) As String

        If InStrComment(WrkStr) = 0 Then
            WrkStr = ""
        Else
            WrkStr = Strings.Right(WrkStr, Len(WrkStr) - InStrComment(WrkStr) - 1)
            WrkStr = WrkStr.TrimStart
        End If

        Return WrkStr
    End Function

    'コメント判定
    Private Function InStrComment(ByVal WrkStr As String) As Integer

        If WrkStr = "" Then
            Return 0
        End If

        If Strings.Left(WrkStr, 2) = "//" Then
            Return 1
        End If

        Dim WrkPos As Integer

        WrkPos = InStr(WrkStr, " //")
        If WrkPos <> 0 Then
            Return WrkPos + 1
        End If

        WrkPos = InStr(WrkStr, vbTab & "//")
        If WrkPos <> 0 Then
            Return WrkPos + 1
        End If

        WrkPos = InStr(WrkStr, "　//")
        If WrkPos <> 0 Then
            Return WrkPos + 1
        End If

        Return 0
    End Function

    'パラメータ取得
    Private Function GetSettingInformation(ByVal WrkStr As String, ByVal WrkSettingName As String) As String

        WrkStr = GetDeleteComment(WrkStr)

        If InStr(WrkStr.ToUpper, WrkSettingName) <> 0 Then

            If InStr(WrkStr, "=") <> 0 Then

                WrkStr = GetDeleteComment(WrkStr) 'コメント部分削除
                WrkStr = Strings.Right(WrkStr, Len(WrkStr) - InStr(WrkStr, "=")) 'コメント部分削除

                Return WrkStr

            End If

        End If

        Return Nothing
    End Function


End Class
