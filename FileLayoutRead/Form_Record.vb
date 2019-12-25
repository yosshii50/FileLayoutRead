Public Class Form_Record

    Public Sub View(ByVal WrkTitle As String, ByRef WrkRecordData As FLR_RecordData)

        Dim WrkStr As String = ""

        For Each FieldData As FLR_FieldData In WrkRecordData.FieldDatas

            WrkStr = WrkStr & FieldData.StartPos & vbTab
            WrkStr = WrkStr & FieldData.FieldType.FieldLength & vbTab
            WrkStr = WrkStr & "[" & FieldData.GetString & "]" & vbTab

            If FieldData.FieldType.DBFieldName = "" Then
                WrkStr = WrkStr & vbTab
            Else
                WrkStr = WrkStr & "(" & FieldData.FieldType.DBFieldName & ")" & vbTab
            End If

            WrkStr = WrkStr & FieldData.FieldType.FieldName
            WrkStr = WrkStr & vbCrLf

        Next

        WrkStr = WrkStr & "------------------------" & vbCrLf
        If WrkRecordData.RecordType.DB_TableName <> "" Then
            WrkStr = WrkStr & "TABLENAME=[" & WrkRecordData.RecordType.DB_TableName & "]" & vbCrLf
        End If

        TextBox1.Text = WrkStr
        TextBox1.Select(0, 0)

        Me.Text = WrkTitle
        Call Me.Show()

    End Sub

End Class