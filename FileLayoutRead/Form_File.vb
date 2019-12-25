Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

Public Class Form_File

    Private MyFLR_File As FLR_File

    'データ読み込み実行
    Public Function View(ByRef WrkFLR_File As FLR_File) As Boolean

        MyFLR_File = WrkFLR_File

        Me.Text = WrkFLR_File.FileName

        '画面にリスト表示
        If ViewList(MyFLR_File.Data.RecordDatas) = False Then
            Return False
        End If

        Call Me.Show()

        Return True
    End Function

    '画面にリスト表示
    Private Function ViewList(ByRef RecordDatas() As FLR_RecordData) As Boolean

        '画面にリスト表示
        ListBox1.Items.Clear()
        For Each RecordData As FLR_RecordData In RecordDatas

            Dim WrkList As String = "[" & RecordData.RecordType.RecordTypeName & "]"

            For Each FieldData As FLR_FieldData In RecordData.FieldDatas
                WrkList = WrkList & FieldData.GetString & " "
            Next

            ListBox1.Items.Add(WrkList)
        Next

        Return True
    End Function

    'リストボックスダブルクリック
    Private Sub ListBox1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.DoubleClick

        Dim WrkRecordView As New Form_Record

        Call WrkRecordView.View(ListBox1.SelectedIndex.ToString, MyFLR_File.Data.RecordDatas(ListBox1.SelectedIndex))

    End Sub

End Class