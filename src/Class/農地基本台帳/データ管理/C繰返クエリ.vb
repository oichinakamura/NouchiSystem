
Public Class C繰返クエリ
    Inherits HimTools2012.clsAccessor

    Public 繰返名称 As String
    Public Params As New Dictionary(Of String, Object)

    Public Sub New(ByVal s繰返名称 As String)
        繰返名称 = s繰返名称
    End Sub


    Public Overrides Sub Execute()
        Me.Message = 繰返名称 & "の実行"
        Select Case 繰返名称
            Case "10a賃借料金額換算"
                Dim pView As DataView = New DataView(Params.Item("データテーブル"), "[処理区分]='換算成功'", "", DataViewRowState.CurrentRows)
                Me.Value = 0
                Me.Maximum = pView.Count
                For Each pRow As DataRowView In pView
                    Me.Value += 1
                    If Not Val(pRow.Item("10a賃借料").ToString) = pRow.Item("計算結果10a賃借料") Then
                        SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [10a賃借料]=" & pRow.Item("計算結果10a賃借料") & " WHERE [ID]=" & pRow.Item("ID"))
                        pRow.Item("処理区分") = "換算済み"
                        Me.Message = 繰返名称 & "の実行 (" & Me.Value & "/" & Me.Maximum & ")"
                    End If
                    If Me._Cancel Then
                        Exit For
                    End If
                Next
        End Select
    End Sub
End Class
