Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

Public Class FLR_File_SampleData

    Public Sub New(ByRef WrkFLR_File As FLR_File)
        BaseFLR_File = WrkFLR_File
    End Sub
    Private BaseFLR_File As FLR_File

    '生成するサンプルファイルのファイル名
    Private _FileName As String = ""
    Public Property FileName() As String
        Get
            Return _FileName
        End Get
        Set(ByVal value As String)
            _FileName = value
        End Set
    End Property

    '生成するサンプルデータのパターン
    Private _DataPattern As String = ""
    Public Property DataPattern() As String
        Get
            Return _DataPattern
        End Get
        Set(ByVal value As String)
            _DataPattern = value
        End Set
    End Property

    '生成実行
    Public Function Create() As Boolean
        Return Create("", "")
    End Function
    Public Function Create(ByVal FileName As String) As Boolean
        Return Create(FileName, "")
    End Function
    Public Function Create(ByVal FileName As String, ByVal SamplePattern As String) As Boolean

        If FileName = "" Then
            FileName = Me.FileName
        End If
        If SamplePattern = "" Then
            SamplePattern = Me.DataPattern
        End If

        Dim sw As New System.IO.StreamWriter(FileName, False, System.Text.Encoding.GetEncoding("shift_jis"))

        With BaseFLR_File
            For Each WrkPtn As String In Split(SamplePattern, ",")

                For WrkRIdx As Integer = 0 To .RecordTypeMax

                    If .RecordTypes(WrkRIdx).RecordTypeName = WrkPtn Then

                        For Each WrkField As FLR_FieldType In .RecordTypes(WrkRIdx).Fields


                            sw.Write(WrkField.GetSampleData)


                        Next

                        Exit For
                    End If

                Next

            Next
        End With

        sw.Close()

        Return True
    End Function

End Class
