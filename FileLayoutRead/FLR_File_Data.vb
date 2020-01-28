Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

'構成
'　ファイルデータ
'　　+-レコードデータ
'　　| 　+-レコードタイプ
'　　| 　  　+-フィールド
'　　| 　  　+-フィールド
'　　+-レコードデータ
'　　  　+-レコードタイプ
'　　  　  　+-フィールド
'　　  　  　+-フィールド

Public Class FLR_File_Data

    Public Sub New(ByRef WrkFLR_File As FLR_File)
        BaseFLR_File = WrkFLR_File
    End Sub
    Private BaseFLR_File As FLR_File

    'レコードデータ
    Private _RecordDatas() As FLR_RecordData 'Idx0から使用
    Public ReadOnly Property RecordDatas() As FLR_RecordData()
        Get
            Return _RecordDatas
        End Get
    End Property
    Public ReadOnly Property RecordCount() As Integer
        Get
            If _RecordDatas Is Nothing Then
                Return 0
            Else
                Return _RecordDatas.Count
            End If

        End Get
    End Property

    '読み込み実行
    Public Sub Load()
        Call Load(BaseFLR_File.FileName)
    End Sub
    Public Function Load(ByVal FileName As String) As Boolean

        If FileName <> "" Then
            BaseFLR_File.FileName = FileName
        End If

        'データファイルオープン
        Dim WrkFileStream As System.IO.FileStream
        Try
            WrkFileStream = New System.IO.FileStream(BaseFLR_File.FileName, System.IO.FileMode.Open, System.IO.FileAccess.Read)
        Catch ex As System.IO.FileNotFoundException
            MsgBox("データファイルが見つかりません。")
            Return False
        Catch ex As System.IO.DirectoryNotFoundException
            MsgBox("ファイルパスが見つかりません。")
            Return False
        Catch ex As System.NotSupportedException
            MsgBox("ファイルパス形式に対応していません。")
            Return False
        Catch ex As Exception
            MsgBox(ex.ToString)
            MsgBox("データファイルの読み込みに失敗しました。")
            Return False
        End Try

        'データ読み込み領域
        Dim WrkBuf(BaseFLR_File.RecordSize - 1) As Byte

        '残りのデータサイズ
        Dim WrkRemainSize As Integer = CInt(WrkFileStream.Length)

        Erase _RecordDatas

        While WrkRemainSize > 0

            Dim WrkReadSize As Integer

            'ファイルからデータ読み込み
            WrkReadSize = WrkFileStream.Read(WrkBuf, 0, Math.Min(BaseFLR_File.RecordSize, WrkRemainSize))

            For Each WrkRecordType As FLR_RecordType In BaseFLR_File.RecordTypes

                '識別パターンと一致しているか確認
                If WrkRecordType.CheckPattern(WrkBuf) = True Then

                    If _RecordDatas Is Nothing Then
                        ReDim _RecordDatas(0)
                    Else
                        ReDim Preserve _RecordDatas(_RecordDatas.Count)
                    End If
                    _RecordDatas(_RecordDatas.Count - 1) = New FLR_RecordData

                    'レコードデータの追加
                    Call _RecordDatas(_RecordDatas.Count - 1).RecordDataAdd(WrkBuf, WrkRecordType, Nothing)

                    Exit For
                End If

            Next

            WrkRemainSize -= WrkReadSize
        End While

        WrkFileStream.Close()

        Return True
    End Function
End Class
