Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

Public Class FLR_DBField

    Public Sub New(ByRef WrkFLR_FieldType As FLR_FieldType)
        BaseFLR_FieldType = WrkFLR_FieldType
    End Sub
    Private BaseFLR_FieldType As FLR_FieldType

    'フィールド名
    Private _DBFieldName As String
    Public Property DBFieldName() As String
        Get
            Return _DBFieldName
        End Get
        Set(ByVal value As String)
            _DBFieldName = value
        End Set
    End Property

    'パラメータ
    Public ReadOnly Property Parameter() As String
        Get
            Dim WrkStr As String = ""

            Select Case BaseFLR_FieldType.FieldType
                Case FLR_FieldType.FieldType_Enum.Num9
                    WrkStr = WrkStr & "NUMBER   "
                    WrkStr = WrkStr & "(" & BaseFLR_FieldType.FieldLength & ")"
                    WrkStr = WrkStr & "DEFAULT  0  NOT NULL"
                Case Else
                    WrkStr = WrkStr & "VARCHAR2 "
                    WrkStr = WrkStr & "(" & BaseFLR_FieldType.FieldLength & ")"
            End Select
            Return WrkStr
        End Get
    End Property

    'コメント
    Public ReadOnly Property Comment() As String
        Get
            Dim WrkStr As String = ""

            WrkStr = WrkStr & "-- " & BaseFLR_FieldType.FieldName

            Return WrkStr
        End Get
    End Property

End Class
