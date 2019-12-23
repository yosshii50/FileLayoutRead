Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

Public Class FLR_RecordData

    'レコードデータ
    Private _RecordData As Byte()
    Public ReadOnly Property RecordData() As Byte()
        Get
            Return _RecordData
        End Get
    End Property

    'レコードデータの追加
    Public Sub RecordDataAdd(ByRef WrkRecordData As Byte(), ByRef WrkRecordType As FLR_RecordType, ByRef WrkParentRecordData As FLR_RecordData)

        'レコードデータセット
        ReDim _RecordData(WrkRecordData.Count - 1)
        For WrkIdx As Integer = 0 To WrkRecordData.Count - 1
            _RecordData(WrkIdx) = WrkRecordData(WrkIdx)
        Next

        'レコードタイプセット
        _RecordType = WrkRecordType

        If Not WrkParentRecordData Is Nothing Then
            '親レコードが設定されている場合

            '親レコードセット
            Me._ParentRecordData = WrkParentRecordData

            '親レコードの子レコードセット
            Call WrkParentRecordData.ChildRecordDataAdd(Me)

        End If

    End Sub

    'レコードタイプ
    Private _RecordType As FLR_RecordType
    Public ReadOnly Property RecordType() As FLR_RecordType
        Get
            Return _RecordType
        End Get
    End Property

    '親レコード
    Private _ParentRecordData As FLR_RecordData
    Public ReadOnly Property ParentRecordData() As FLR_RecordData
        Get
            Return _ParentRecordData
        End Get
    End Property

    '子レコード
    Private _ChildRecordData() As FLR_RecordData
    Public ReadOnly Property ChildRecordData() As FLR_RecordData()
        Get
            Return _ChildRecordData
        End Get
    End Property
    Protected Sub ChildRecordDataAdd(ByRef WrkChildRecordData As FLR_RecordData)

        If _ChildRecordData Is Nothing Then
            ReDim _ChildRecordData(0)
        Else
            ReDim Preserve _ChildRecordData(_ChildRecordData.Count)
        End If

        _ChildRecordData(_ChildRecordData.Count - 1) = WrkChildRecordData

    End Sub

End Class
