﻿Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

Public Class Form1

    Dim MyFLR_File As FLR_File

    Public Sub FromStart(ByVal GetConfigFileName As String, ByVal GetDataFileName As String)

        Call Me.Show()

        TxtConfig.Text = GetConfigFileName
        TxtData.Text = GetDataFileName

    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        '実行パス取得
        Dim WrkExeName As String
        WrkExeName = System.Reflection.Assembly.GetExecutingAssembly().Location
        WrkExeName = System.IO.Path.GetFileNameWithoutExtension(WrkExeName)
        TxtConfig.Text = WrkExeName & ".conf.txt"

    End Sub

    '定義ファイル読込
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If LoadConfig(MyFLR_File, TxtConfig.Text) = True Then
            TextBox3.Text = TextBox3.Text & "定義ファイル読込 済" & vbCrLf

            '結果表示
            Dim WrkStr As String = ""
            For Each WrkRecordType As FLR_RecordType In MyFLR_File.RecordTypes
                WrkStr = WrkStr & WrkRecordType.RecordTypeName & " "
                WrkStr = WrkStr & WrkRecordType.Fields.Count & "項目" & vbCrLf
            Next

            For Each WrkDBTable As FLR_DBTable In MyFLR_File.DB.DBTables
                WrkStr = WrkStr & WrkDBTable.TableName & " "
                If WrkDBTable.DBFields Is Nothing Then
                    WrkStr = WrkStr & "項目なし" & vbCrLf
                Else
                    WrkStr = WrkStr & WrkDBTable.DBFields.Count & "項目" & vbCrLf
                End If
            Next

            TextBox3.Text = TextBox3.Text & WrkStr

            'データファイル取得
            If MyFLR_File.FileName <> "" Then
                TxtData.Text = MyFLR_File.FileName
            End If

        End If
    End Sub
    Private Function LoadConfig(ByRef WrkFLR_File As FLR_File, ByVal ConfigFileName As String) As Boolean

        WrkFLR_File = New FLR_File

        '定義ファイル読込
        If WrkFLR_File.Parameter.LoadFile(ConfigFileName) = False Then
            Return False
        End If

        Return True
    End Function

    'サンプルデータ生成
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If CreateSample(MyFLR_File, TxtConfig.Text, TxtData.Text) = True Then
            TextBox3.Text = TextBox3.Text & "サンプルデータ生成 済" & vbCrLf
        End If
    End Sub
    Private Function CreateSample(ByRef WrkFLR_File As FLR_File, ByVal ConfigFileName As String, ByVal SampleFileName As String) As Boolean

        If WrkFLR_File Is Nothing Then
            MsgBox("定義ファイルが読み込まれていません。")
            Return False
        End If

        'サンプルデータ生成
        If WrkFLR_File.SampleData.Create(SampleFileName) = False Then
            Return False
        End If

        Return True
    End Function

    'データ読込
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If DataLoad(MyFLR_File, TxtConfig.Text, TxtData.Text, True) = True Then
            TextBox3.Text = TextBox3.Text & "データ読込 済" & vbCrLf
        End If
    End Sub
    Private Function DataLoad(ByRef WrkFLR_File As FLR_File, ByVal ConfigFileName As String, ByVal DataFileName As String, ByVal IsDebuMode As Boolean) As Boolean

        If WrkFLR_File Is Nothing Then
            MsgBox("定義ファイルが読み込まれていません。")
            Return False
        End If

        'データ読込実行
        If WrkFLR_File.Data.Load(DataFileName) = False Then
            Return False
        End If

        Return True
    End Function

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Call DataView(MyFLR_File)
    End Sub
    Private Function DataView(ByRef WrkFLR_File As FLR_File) As Boolean

        If WrkFLR_File Is Nothing Then
            MsgBox("定義ファイルが読み込まれていません。")
            Return False
        End If

        If WrkFLR_File.Data.RecordDatas Is Nothing Then
            MsgBox("データが読み込まれていません。")
            Return False
        End If

        'データ読込結果表示
        Dim WrkDataView As New Form_File
        If WrkDataView.View(WrkFLR_File) = False Then
            Return False
        End If

        Return True
    End Function

    'DBデータ作成
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If DBCreate(MyFLR_File) = True Then
            TextBox3.Text = TextBox3.Text & "DBデータ作成 済" & vbCrLf
        End If
    End Sub
    Private Function DBCreate(ByRef WrkFLR_File As FLR_File) As Boolean

        If WrkFLR_File Is Nothing Then
            MsgBox("定義ファイルが読み込まれていません。")
            Return False
        End If

        If WrkFLR_File.Data.RecordDatas Is Nothing Then
            MsgBox("データが読み込まれていません。")
            Return False
        End If

        '変換処理実行
        If DBAccess.Start(WrkFLR_File) = False Then
            Return False
        End If

        Return True
    End Function

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click

        'OpenFileDialogクラスのインスタンスを作成
        Dim WrkOpenFile As New OpenFileDialog()

        With WrkOpenFile

            .FileName = System.IO.Path.GetFileName(TxtData.Text)
            If TxtData.Text <> "" Then
                .InitialDirectory = System.IO.Path.GetDirectoryName(TxtData.Text)
            End If
            .Filter = "すべてのファイル(*.*)|*.*"
            .FilterIndex = 1
            '.Title = "開くファイルを選択してください"
            ''ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
            'ofd.RestoreDirectory = True

            'ダイアログを表示する
            If .ShowDialog() = DialogResult.OK Then
                'OKボタンがクリックされたとき、選択されたファイル名を表示する
                TxtData.Text = .FileName
            End If

        End With

    End Sub
End Class
