Imports HimTools2012.CommonFunc

Public Class classpage転用非農地整合性確認
    Inherits HimTools2012.controls.CTabPageWithToolStrip
    Private mvarGrid As New HimTools2012.controls.DataGridViewWithDataView
    Private WithEvents mvarLoadBtn As New ToolStripButton
    Private pViewTable As DataTable


    Public Sub New()
        MyBase.New(True, True, "転用非農地整合性確認", "転用非農地整合性確認")

        mvarGrid.AllowUserToAddRows = False
        Me.ControlPanel.Add(mvarGrid)
        mvarLoadBtn.Text = "バックアップからの復元"
        Me.ToolStrip.Items.Add(mvarLoadBtn)

        pViewTable = New DataTable
        pViewTable.Columns.Add("ID", GetType(Decimal))
        pViewTable.Columns.Add("許可日", GetType(String))
        pViewTable.Columns.Add("修復済み", GetType(Boolean))

        DoComp()
    End Sub

    Private Sub DoComp()
        pViewTable.Rows.Clear()
        Dim pDic As New Dictionary(Of Decimal, String)
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT D_申請.法令, D_申請.農地リスト, D_申請.状態, D_申請.許可年月日 FROM D_申請 WHERE (((D_申請.法令) In (40,50,51,52,53,600)) AND ((D_申請.農地リスト) Is Not Null) AND ((D_申請.状態)=2));")

        For Each pRow As DataRow In pTBL.Rows
            Dim Ar() As String = Split(pRow.Item("農地リスト").ToString, ";")

            For Each St As String In Ar
                If St.Length Then
                    Dim nID As Decimal = GetKeyCode(St)
                    If Not nID = 0 AndAlso Not pDic.ContainsKey(nID) Then
                        pDic.Add(nID, pRow.Item("許可年月日").ToString)
                    End If
                End If
            Next
        Next

        Dim sID As New System.Text.StringBuilder
        Dim p転用済みTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [ID] FROM [D_転用農地] WHERE [ID]=0")
        p転用済みTBL.PrimaryKey = New DataColumn() {p転用済みTBL.Columns("ID")}
        For Each xID As Decimal In pDic.Keys
            sID.Append(IIf(sID.Length > 0, ",", "") & xID)
            If sID.Length > 512 Then
                Dim pTBL申請地A As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [ID] FROM [D_転用農地] WHERE [ID] IN (" & sID.ToString & ")")
                Dim pTBL申請地B As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [ID] FROM [D:農地Info] WHERE [ID] IN (" & sID.ToString & ")")
                p転用済みTBL.Merge(pTBL申請地A, False, MissingSchemaAction.AddWithKey)
                p転用済みTBL.Merge(pTBL申請地B, False, MissingSchemaAction.AddWithKey)
                sID.Clear()
            End If
        Next
        If sID.Length > 0 Then
            Dim pTBL申請地A As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [ID] FROM [D_転用農地] WHERE [ID] IN (" & sID.ToString & ")")
            Dim pTBL申請地B As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [ID] FROM [D:農地Info] WHERE [ID] IN (" & sID.ToString & ")")
            p転用済みTBL.Merge(pTBL申請地A, False, MissingSchemaAction.AddWithKey)
            p転用済みTBL.Merge(pTBL申請地B, False, MissingSchemaAction.AddWithKey)
            sID.Clear()
        End If

        For Each nID As Decimal In pDic.Keys
            Dim pRow As DataRow = p転用済みTBL.Rows.Find(nID)
            If pRow Is Nothing Then
                pRow = pViewTable.NewRow
                pRow.Item("ID") = nID
                pRow.Item("許可日") = pDic.Item(nID)
                pRow.Item("修復済み") = False
                pViewTable.Rows.Add(pRow)
            End If
        Next
        mvarGrid.SetDataView(pViewTable, "", "")
    End Sub


    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.CloseOK
        End Get
    End Property


    Private Sub mvarLoadBtn_Click(sender As Object, e As System.EventArgs) Handles mvarLoadBtn.Click
        With New OpenFileDialog
            .Filter = "*.MDB|*.MDB"
            If .ShowDialog = DialogResult.OK Then
                Dim pCn As New HimTools2012.Data.CLocalDataEngine("")
                pCn.LocalPath = .FileName

                Dim sSQL As New System.Text.StringBuilder
                For Each pRow As DataRow In New DataView(pViewTable, "[修復済み]=False", "", DataViewRowState.CurrentRows).ToTable.Rows
                    Dim p農地 As DataTable = pCn.GetTableBySqlSelect_Local("SELECT * FROM [D:農地Info] WHERE [ID]=" & pRow.Item("ID"))
                    If p農地 IsNot Nothing AndAlso p農地.Rows.Count > 0 Then

                        Dim pXRow As DataRow = p農地.Rows(0)
                        For Each pCol As DataColumn In p農地.Columns
                            If pCol.ColumnName = "ID" Then
                            Else
                                Select Case pCol.DataType.Name
                                    Case "Int32", "Int16", "Double", "Decimal", "Single" : sSQL.Append(IIf(sSQL.Length > 0, ", ", "") & "[" & pCol.ColumnName & "]=" & pXRow.Item(pCol.ColumnName))

                                    Case "Boolean" : sSQL.Append(IIf(sSQL.Length > 0, ", ", "") & "[" & pCol.ColumnName & "]=" & IIf(pRow.Item(pCol.ColumnName) = True, "True", "False"))
                                    Case "String" : sSQL.Append(IIf(sSQL.Length > 0, ", ", "") & "[" & pCol.ColumnName & "]=" & IIf(IsDBNull(pRow.Item(pCol.ColumnName)), "Null", """" & pRow.Item(pCol.ColumnName) & """"))
                                    Case "DateTime"
                                        If IsDBNull(pRow.Item(pCol.ColumnName)) Then
                                            sSQL.Append(IIf(sSQL.Length > 0, ", ", "") & "[" & pCol.ColumnName & "]=Null")
                                        Else
                                            With CDate(pRow.Item(pCol.ColumnName))
                                                sSQL.Append(IIf(sSQL.Length > 0, ", ", "") & "[" & pCol.ColumnName & "]=#" & .Month & "/" & .Day & "/" & .Year & "#")
                                            End With
                                        End If
                                    Case Else
                                        CasePrint(pCol.DataType.Name)
                                End Select
                            End If
                        Next
                        If sSQL.Length > 0 Then
                            SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D_転用農地]([ID]) VALUES(" & pRow.Item("ID") & ")")
                            SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D_転用農地] SET " & sSQL.ToString & " WHERE [ID]=" & pRow.Item("ID"))
                            pRow.Item("修復済み") = True
                        End If
                    End If
                Next
            End If
        End With
    End Sub
End Class