Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

'設定情報からファイルレイアウト展開
Public Class FLR_File_Parameter

    Public Sub New(ByRef WrkFLR_File As FLR_File)
        BaseFLR_File = WrkFLR_File
    End Sub
    Private BaseFLR_File As FLR_File

    '生成実行
    Public Sub Read(ByVal BaseString As String)

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
                BaseFLR_File.FileName = RetStr
            End If

            '[SAMPLEFILE]取得
            RetStr = GetSettingInformation(WrkStr, "SAMPLEFILE")
            If Not RetStr Is Nothing Then
                BaseFLR_File.SampleData.FileName = RetStr
            End If

            '[SAMPLEPATTERN]取得
            RetStr = GetSettingInformation(WrkStr, "SAMPLEPATTERN")
            If Not RetStr Is Nothing Then
                BaseFLR_File.SampleData.DataPattern = RetStr
            End If

            '[RECTYPE]取得
            RetStr = GetSettingInformation(WrkStr, "RECTYPE")
            If Not RetStr Is Nothing Then
                BaseFLR_File.RecordTypeAdd(RetStr)
            End If

            If BaseFLR_File.RecordTypeCount <> 0 Then
                '[RECTYPE]が存在する場合

                'フィールド情報取得
                Dim GetFLR_Field As FLR_Field = GetFieldInformation(WrkStr)
                If Not GetFLR_Field Is Nothing Then

                    'フィールドの追加
                    BaseFLR_File.RecordTypes(BaseFLR_File.RecordTypeMax).AddField(GetFLR_Field)

                End If

            End If

        Next

    End Sub

    'フィールド情報取得
    Private Function GetFieldInformation(ByVal WrkStr As String) As FLR_Field

        Dim GetFLR_Field As New FLR_Field

        '区切り文字をTABに共通化
        Dim DisassStr As String = GetDeleteComment(WrkStr)
        DisassStr = DisassStr.Replace("　", vbTab)
        DisassStr = DisassStr.Replace(" ", vbTab)

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
                                GetFLR_Field.FieldType = FLR_Field.FieldType_Enum.StrX
                            Case "9"
                                GetFLR_Field.FieldType = FLR_Field.FieldType_Enum.Num9
                            Case Else
                                Return Nothing
                        End Select

                    Case 3 '３つ目の項目 / サンプルパターン
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

        If InStr(WrkStr, "//") <> 0 Then
            WrkStr = Strings.Left(WrkStr, InStr(WrkStr, "//") - 1)
        End If

        Return WrkStr
    End Function

    'コメント部分取得
    Private Function GetComment(ByVal WrkStr As String) As String

        If InStr(WrkStr, "//") = 0 Then
            WrkStr = ""
        Else
            WrkStr = Strings.Right(WrkStr, Len(WrkStr) - InStr(WrkStr, "//") - 1)
            WrkStr = WrkStr.TrimStart
        End If

        Return WrkStr
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
