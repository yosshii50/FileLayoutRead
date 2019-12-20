Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

'設定情報からファイルレイアウト展開
Public Module CreateFLRFile

    '生成実行
    Public Sub Generate(ByRef WrkFLR_File As FLR_File, ByVal BaseString As String)

        For Each WrkStr As String In Split(BaseString, vbCrLf)

            Dim RetStr As String

            '[RECSIZE]取得
            RetStr = GetSettingInformation(WrkStr, "RECSIZE")
            If Not RetStr Is Nothing Then
                WrkFLR_File.RecordSize = CInt(RetStr)
            End If

            '[FILENAME]取得
            RetStr = GetSettingInformation(WrkStr, "FILENAME")
            If Not RetStr Is Nothing Then
                WrkFLR_File.FileName = RetStr
            End If

            '[RECTYPE]取得
            RetStr = GetSettingInformation(WrkStr, "RECTYPE")
            If Not RetStr Is Nothing Then
                WrkFLR_File.AddRecordType(RetStr)
            End If

            If WrkFLR_File.RecordTypeCount <> 0 Then
                'レコードタイプが存在する場合のみ

                'フィールド情報取得
                RetStr = GetFieldInformation(WrkStr)
                If Not RetStr Is Nothing Then



                    WrkFLR_File.LastRecord

                End If




            End If

        Next

    End Sub

    'コメント部分削除
    Private Function GetDeleteComment(ByVal WrkStr As String) As String

        If InStr(WrkStr, "//") <> 0 Then
            WrkStr = Strings.Left(WrkStr, InStr(WrkStr, "//") - 1)
        End If

        Return WrkStr
    End Function

    'パラメータ取得
    Private Function GetSettingInformation(ByVal WrkStr As String, ByVal WrkSettingName As String) As String

        If InStr(WrkStr.ToUpper, WrkSettingName) <> 0 Then

            If InStr(WrkStr, "=") <> 0 Then

                WrkStr = GetDeleteComment(WrkStr) 'コメント部分削除
                WrkStr = Strings.Right(WrkStr, Len(WrkStr) - InStr(WrkStr, "=")) 'コメント部分削除

                Return WrkStr

            End If

        End If

        Return Nothing
    End Function


End Module
