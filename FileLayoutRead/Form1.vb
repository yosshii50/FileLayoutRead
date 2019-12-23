Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

Public Class Form1

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Dim MyFLR_File As New FLR_File

        Dim WrkStr As String = System.IO.File.ReadAllText(TextBox1.Text, System.Text.Encoding.GetEncoding("Shift_JIS"))

        '設定情報からファイルレイアウト展開
        Call MyFLR_File.Parameter.Read(WrkStr)

        'サンプルデータ生成
        Call MyFLR_File.SampleData.Create(TextBox2.Text)

        MsgBox("済")

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim MyFLR_File As New FLR_File

        Dim WrkStr As String = System.IO.File.ReadAllText(TextBox1.Text, System.Text.Encoding.GetEncoding("Shift_JIS"))

        '設定情報からファイルレイアウト展開
        Call MyFLR_File.Parameter.Read(WrkStr)

        '読み込み実行
        Call MyFLR_File.Data.Load(TextBox2.Text)

        MsgBox("済")

    End Sub
End Class
