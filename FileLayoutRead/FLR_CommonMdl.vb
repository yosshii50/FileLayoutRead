Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

Public Module FLR_CommonMdl

    '文字列前後の指定文字を削除
    Public Function StrStEdDel(ByVal FileName As String) As String
        Return StrStEdDel(FileName, """", """")
    End Function
    Public Function StrStEdDel(ByVal WrkStr As String, ByVal StStr As String, ByVal EdStr As String) As String
        If Strings.Left(WrkStr, 1) = StStr And Strings.Right(WrkStr, 1) = EdStr Then
            WrkStr = Strings.Mid(WrkStr, 2, Len(WrkStr) - 2)
        End If
        Return WrkStr
    End Function

End Module
