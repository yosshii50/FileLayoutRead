Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

'フィールドデータ
Public Class FLR_FieldData

    Private _RecordData As Byte() 'レコードデータ

    'フィールドタイプ
    Private _FieldType As FLR_FieldType
    Public ReadOnly Property FieldType() As FLR_FieldType
        Get
            Return _FieldType
        End Get
    End Property

    '開始位置
    Private _StartPos As Integer
    Public ReadOnly Property StartPos() As Integer
        Get
            Return _StartPos
        End Get
    End Property

    'フィールド文字列データ
    Public ReadOnly Property GetString() As String
        Get

            '作業用配列に転送
            Dim WrkRD() As Byte
            ReDim WrkRD(_FieldType.FieldLength - 1)
            For WrkIdx As Integer = 0 To _FieldType.FieldLength - 1
                WrkRD(WrkIdx) = _RecordData(WrkIdx + _StartPos)
            Next

            If _FieldType.FieldType = FLR_FieldType.FieldType_Enum.BinX Then
                Return "x" & Hex2Str(WrkRD)
            Else
                Return System.Text.Encoding.GetEncoding(932).GetString(WrkRD)
            End If

        End Get
    End Property
    Private Function Hex2Str(ByVal WrkRD() As Byte) As String

        Dim WrkStr As String = ""

        For WrkIdx As Integer = 0 To WrkRD.Count - 1
            WrkStr = WrkStr & Strings.Right("0" & Hex(WrkRD(WrkIdx)), 2)
        Next

        Return WrkStr
    End Function

    'フィールドデータの追加
    Public Sub FieldDataAdd(ByRef WrkRecordData As Byte(), ByRef WrkFieldType As FLR_FieldType, ByVal StartPos As Integer)

        'フィールドタイプ
        _FieldType = WrkFieldType

        'レコードデータ
        _RecordData = WrkRecordData

        '開始位置
        _StartPos = StartPos

    End Sub

End Class
