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

    'フィールド
    Private _FieldDatas() As FLR_FieldData
    Public Property FieldDatas() As FLR_FieldData()
        Get
            Return _FieldDatas
        End Get
        Set(ByVal value As FLR_FieldData())
            _FieldDatas = value
        End Set
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

        Dim StartPos As Integer = 0
        For Each WrkFieldType As FLR_FieldType In _RecordType.Fields

            If _FieldDatas Is Nothing Then
                ReDim _FieldDatas(0)
            Else
                ReDim Preserve _FieldDatas(_FieldDatas.Count)
            End If

            _FieldDatas(_FieldDatas.Count - 1) = New FLR_FieldData

            'フィールドデータ
            Call _FieldDatas(_FieldDatas.Count - 1).FieldDataAdd(_RecordData, WrkFieldType, StartPos)

            StartPos = StartPos + WrkFieldType.FieldLength
        Next

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

    'Insert文用SQL生成
    Public Function GetInsertTableListStr() As String

        Dim IsFirst As Boolean
        Dim WrkStr As String = ""

        WrkStr = WrkStr & "INSERT INTO " & RecordType.DB_TableName & vbCrLf
        IsFirst = True
        For Each WrkFieldData As FLR_FieldData In FieldDatas
            If WrkFieldData.FieldType.DBFieldName <> "" Then

                If IsFirst = True Then
                    WrkStr = WrkStr & "("
                Else
                    WrkStr = WrkStr & ","
                End If
                IsFirst = False

                WrkStr = WrkStr & WrkFieldData.FieldType.DBFieldName & vbCrLf

            End If
        Next
        WrkStr = WrkStr & ")" & vbCrLf
        WrkStr = WrkStr & "VALUES" & vbCrLf
        IsFirst = True
        For Each WrkFieldData As FLR_FieldData In FieldDatas
            If WrkFieldData.FieldType.DBFieldName <> "" Then

                If IsFirst = True Then
                    WrkStr = WrkStr & "("
                Else
                    WrkStr = WrkStr & ","
                End If
                IsFirst = False

                If WrkFieldData.FieldType.FieldType = FLR_FieldType.FieldType_Enum.Num9 Then
                    WrkStr = WrkStr & WrkFieldData.GetString() & vbCrLf
                Else
                    WrkStr = WrkStr & "'" & WrkFieldData.GetString() & "'" & vbCrLf
                End If

            End If
        Next
        WrkStr = WrkStr & ")" & vbCrLf

        Return WrkStr
    End Function

End Class
