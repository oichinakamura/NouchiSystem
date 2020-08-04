

Public Class CTabPage申請10a当金額変換
    Inherits HimTools2012.controls.CTabPageWithToolStrip
    Private mvarGrid As HimTools2012.controls.DataGridViewWithDataView

    Public Sub New()
        MyBase.New(True, True, "10a当金額変換", "10a当金額変換")
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT D_申請.ID, D_申請.受付年月日, D_申請.名称, D_申請.小作料, D_申請.小作料単位, D_申請.法令, D_申請.[10a金額],'' AS [処理区分] FROM D_申請 WHERE (((D_申請.小作料) Is Not Null) AND ((D_申請.法令) In (31,61,62))) ORDER BY D_申請.受付年月日;")
        mvarGrid = New HimTools2012.controls.DataGridViewWithDataView

        For Each pRow As DataRow In pTBL.Rows
            Dim s小作料 As String = pRow.Item("小作料").ToString
            If Not IsDBNull(pRow.Item("小作料")) Then
                If IsNumeric(s小作料) Then
                Else
                    pRow.Item("処理区分") = "換算失敗"
                End If
            End If

            Select Case pRow.Item("小作料単位").ToString.ToLower
                Case "円/10a"
                    If IsDBNull(pRow.Item("10a金額")) OrElse pRow.Item("10a金額") = 0 Then
                        pRow.Item("10a金額") = Val(Replace(s小作料, ",", ""))
                        SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D_申請] SET [10a金額]=" & pRow.Item("10a金額") & " WHERE [ID]=" & pRow.Item("ID"))
                        pRow.Item("処理区分") = "換算成功"
                    ElseIf Not pRow.Item("10a金額") = Val(Replace(s小作料, ",", "")) Then
                        pRow.Item("10a金額") = Val(Replace(s小作料, ",", ""))
                        SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D_申請] SET [10a金額]=" & pRow.Item("10a金額") & " WHERE [ID]=" & pRow.Item("ID"))
                        pRow.Item("処理区分") = "換算成功"
                    ElseIf pRow.Item("10a金額") = Val(Replace(s小作料, ",", "")) Then
                        pRow.Item("処理区分") = "換算済み"
                    End If
                Case "円"
                    pRow.Item("処理区分") = "換算失敗"
                Case "俵"
                    pRow.Item("処理区分") = "換算失敗"
                Case ""
                    pRow.Item("処理区分") = "換算失敗"
                Case Else
                    pRow.Item("処理区分") = "換算失敗"
            End Select
        Next

        Me.ControlPanel.Add(mvarGrid)
        mvarGrid.SetDataView(pTBL, "", "")
    End Sub

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.CloseOK
        End Get
    End Property

End Class
