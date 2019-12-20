Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim WrkStr As String = ""

        For Each WrKLineData As String In System.IO.File.ReadLines("C:\Install\FileLayoutRead_Data\DataSet.txt", System.Text.Encoding.GetEncoding("Shift_JIS"))

            WrkStr = WrkStr & WrKLineData & vbCrLf

        Next

        TextBox1.Text = WrkStr

    End Sub

    Private MyFLR_File As FLR_File

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        MyFLR_File = New FLR_File

        '設定情報からファイルレイアウト展開
        Call CreateFLRFile.Generate(MyFLR_File, TextBox1.Text)

    End Sub
End Class
