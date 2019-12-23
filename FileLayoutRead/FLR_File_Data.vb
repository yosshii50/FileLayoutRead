Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

'構成
'　ファイル
'　　+-レコードデータ
'　　| 　+-フィールド
'　　+-レコードデータ
'　　  　+-フィールド
'　　  　+-フィールド

Public Class FLR_File_Data

    Public Sub New(ByRef WrkFLR_File As FLR_File)
        BaseFLR_File = WrkFLR_File
    End Sub
    Private BaseFLR_File As FLR_File

    'レコード構成
    Private Structure RecordConstitution_Str
        Dim RecordData() As Byte
        Dim RecordType As FLR_RecordType
        Dim ParentRecordData As FLR_RecordType
        Dim ChildRecordData() As FLR_RecordType
    End Structure
    Private RecordConstitution() As RecordConstitution_Str

    '読み込み実行
    Public Sub Load()
        Call Load(BaseFLR_File.FileName)
    End Sub
    Public Sub Load(ByVal FileName As String)

        If FileName = "" Then
            FileName = BaseFLR_File.FileName
        End If

        'データファイルオープン
        Dim WrkFileStream As System.IO.FileStream
        WrkFileStream = New System.IO.FileStream(FileName, System.IO.FileMode.Open, System.IO.FileAccess.Read)

        'データ読み込み領域
        Dim WrkBuf(BaseFLR_File.RecordSize - 1) As Byte

        '残りのデータサイズ
        Dim WrkRemainSize As Integer = CInt(WrkFileStream.Length)

        While WrkRemainSize > 0

            Dim WrkReadSize As Integer

            'ファイルからデータ読み込み
            WrkReadSize = WrkFileStream.Read(WrkBuf, 0, Math.Min(BaseFLR_File.RecordSize, WrkRemainSize))

            For Each WrkRecordType As FLR_RecordType In BaseFLR_File.RecordTypes
                'BaseFLR_File.Records




            Next





            WrkRemainSize -= WrkReadSize
        End While

        WrkFileStream.Close()

    End Sub
End Class
