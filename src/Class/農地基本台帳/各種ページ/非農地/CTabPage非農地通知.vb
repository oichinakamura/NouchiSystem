

Public Class CTabPage非農地通知
    Inherits HimTools2012.controls.CTabPageWithToolStrip
    Implements HimTools2012.controls.XMLLayoutContainer


    Private WithEvents mvar全印刷 As ToolStripButton
    Private WithEvents mvar全解除 As ToolStripButton
    Private WithEvents mvar発行番号 As New ToolStripTextBox
    Private WithEvents mvar印刷開始 As New ToolStripButton
    Private WithEvents mvarExcel出力 As New ToolStripButton
    Private WithEvents mvar送付先出力 As New ToolStripButton("送付先一覧出力")
    Private mvarTxt所有者 As New ToolStripTextBox
    Private WithEvents mvar検索 As ToolStripButton

    Private WithEvents mvar決定総会年月日 As New HimTools2012.controls.ToolStripDateTimePicker
    Private WithEvents mvar発行年月日 As New HimTools2012.controls.ToolStripDateTimePicker

    Private pTBL As New DataTable
    Private WithEvents mvarGrid送付先 As New HimTools2012.controls.DataGridViewWithDataView
    Private pTBL送付先 As New DataTable

    Protected Property mvarXMLLayout As HimTools2012.controls.XMLLayout
    Protected WithEvents mvarGrid As HimTools2012.controls.DataGridViewWithDataView

    Private mvarTabControl As TabControl

    Public Sub New()
        MyBase.New(True, True, "非農地通知", "非農地通知")

        Me.Body.SuspendLayout()
        mvarXMLLayout = New HimTools2012.controls.XMLLayout(SysAD.DB(sLRDB), ObjectMan, Me)
        With mvarXMLLayout
            .StartLayout(SysAD.SystemInfo.画面設定, "非農地通知")
            mvarTabControl = .Controls("MainTab")
            pTBL = SetTBL(SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT D_非農地通知判定.* FROM (D_非農地通知判定 LEFT JOIN [D:個人Info] ON D_非農地通知判定.送付先ID = [D:個人Info].ID) LEFT JOIN V_行政区 ON [D:個人Info].行政区ID = V_行政区.ID WHERE (((D_非農地通知判定.調査時現況地目) In ('山林','原野','のり面'))) ORDER BY V_行政区.ID, D_非農地通知判定.送付先ID;"))
            'pTBL = SetTBL(SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D_非農地通知判定]"))
            pTBL.Columns.Add("印刷", GetType(Boolean))

            pTBL送付先 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT V_行政区.ID AS 行政区ID, V_行政区.行政区, D_非農地通知判定.送付先ID, D_非農地通知判定.送付先氏名, D_非農地通知判定.送付先住所, D_非農地通知判定.送付先郵便番号, D_非農地通知判定.所有者住民区分名 AS 送付先住民区分 FROM (D_非農地通知判定 INNER JOIN [D:個人Info] ON D_非農地通知判定.送付先ID = [D:個人Info].ID) LEFT JOIN V_行政区 ON [D:個人Info].行政区ID = V_行政区.ID WHERE (((D_非農地通知判定.調査時現況地目) In ('山林','原野','のり面'))) GROUP BY V_行政区.ID, V_行政区.行政区, D_非農地通知判定.送付先ID, D_非農地通知判定.送付先氏名, D_非農地通知判定.送付先住所, D_非農地通知判定.送付先郵便番号, D_非農地通知判定.所有者住民区分名 ORDER BY V_行政区.ID, D_非農地通知判定.送付先ID;")
            'pTBL送付先 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT V_行政区.ID AS 行政区ID, V_行政区.行政区, D_非農地通知判定.送付先ID, D_非農地通知判定.送付先氏名, D_非農地通知判定.送付先住所, D_非農地通知判定.送付先郵便番号 FROM (D_非農地通知判定 INNER JOIN [D:個人Info] ON D_非農地通知判定.送付先ID = [D:個人Info].ID) LEFT JOIN V_行政区 ON [D:個人Info].行政区ID = V_行政区.ID GROUP BY V_行政区.ID, V_行政区.行政区, D_非農地通知判定.送付先ID, D_非農地通知判定.送付先氏名, D_非農地通知判定.送付先住所, D_非農地通知判定.送付先郵便番号 ORDER BY V_行政区.ID, D_非農地通知判定.送付先ID;")

            Dim p印刷Col As New DataGridViewCheckBoxColumn
            With p印刷Col
                .HeaderText = "印刷"
                .DataPropertyName = "印刷"
                .Name = "印刷"
            End With

            mvarGrid = .Controls("G出力")
            With mvarGrid
                .Columns.Add(p印刷Col)
                .VirtualMode = True
                .AutoGenerateColumns = True
                .AllowUserToAddRows = False
                .SetDataView(pTBL, "", "")
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                .Create件数表示Ctrl(Me.ToolStrip)
                .Createエクセル出力Ctrl(Me.ToolStrip)
            End With

            mvarGrid送付先 = .Controls("G送付先")
            With mvarGrid送付先
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                .VirtualMode = True
                .AutoGenerateColumns = True
                .AllowUserToAddRows = False
                .SetDataView(pTBL送付先, "", "")
            End With

            mvar全印刷 = New ToolStripButton("全選択")
            mvar全解除 = New ToolStripButton("全解除")
            mvar印刷開始 = New ToolStripButton("印刷開始")
            mvarExcel出力 = New ToolStripButton("印刷開始(Excel)")
            mvar検索 = New ToolStripButton("検索")
            mvarTxt所有者.Width = 200

            Me.ToolStrip.Items.AddRange({New ToolStripSeparator, mvar全印刷, mvar全解除, New ToolStripSeparator,
                                         New ToolStripLabel("決定総会年月日"), mvar決定総会年月日, New ToolStripSeparator,
                                         New ToolStripLabel("発行年月日"), mvar発行年月日, New ToolStripSeparator,
                                         New ToolStripLabel("発行番号"), mvar発行番号, mvar印刷開始, New ToolStripSeparator, mvarExcel出力, New ToolStripSeparator, mvar送付先出力,
                                         New ToolStripSeparator, New ToolStripLabel("所有者検索"), mvarTxt所有者, mvar検索})
            mvar決定総会年月日.Value = Now.Date
            mvar発行年月日.Value = Now.Date
        End With
        Me.Body.ResumeLayout()
    End Sub

    Private Function SetTBL(ByRef pTBL As DataTable) As DataTable
        Dim pResultTBL As DataTable = New DataTable
        With pResultTBL
            .Columns.Add("備考", GetType(String))
            .Columns.Add("所有者ID", GetType(Integer))
            .Columns.Add("所有者氏名", GetType(String))
            .Columns.Add("所有者住所", GetType(String))
            .Columns.Add("所有者住民区分名", GetType(String))
            .Columns.Add("NID", GetType(Integer))
            .Columns.Add("大字", GetType(String))
            .Columns.Add("小字", GetType(String))
            .Columns.Add("調査時地番", GetType(String))
            .Columns.Add("調査時登記地目", GetType(String))
            .Columns.Add("調査時台帳地目", GetType(String))
            .Columns.Add("調査時現況地目", GetType(String))
            .Columns.Add("調査時面積", GetType(Decimal))
            .Columns.Add("送付先ID", GetType(Integer))
            .Columns.Add("送付先氏名", GetType(String))
            .Columns.Add("送付先住所", GetType(String))
            .Columns.Add("送付先郵便番号", GetType(String))
            .Columns.Add("発行番号", GetType(String))
            .Columns.Add("発行年月日", GetType(String))
            .Columns.Add("通知番号", GetType(Integer))
            .Columns.Add("所有者郵便番号", GetType(String))
            .Columns.Add("所有者住民区分", GetType(Integer))
            .Columns.Add("大字CD", GetType(Integer))
            .Columns.Add("ID", GetType(Integer))
        End With

        For Each pRow As DataRow In pTBL.Rows
            Dim AddRow As DataRow = pResultTBL.NewRow
            AddRow.Item("ID") = pRow.Item("ID")
            AddRow.Item("NID") = pRow.Item("NID")
            AddRow.Item("発行番号") = pRow.Item("発行番号")
            AddRow.Item("発行年月日") = pRow.Item("発行年月日")
            AddRow.Item("通知番号") = pRow.Item("通知番号")
            AddRow.Item("大字CD") = pRow.Item("大字CD")
            AddRow.Item("大字") = pRow.Item("大字")
            AddRow.Item("小字") = pRow.Item("小字")
            AddRow.Item("調査時地番") = pRow.Item("調査時地番")
            AddRow.Item("調査時登記地目") = pRow.Item("調査時登記地目")
            AddRow.Item("調査時台帳地目") = pRow.Item("調査時台帳地目")
            AddRow.Item("調査時現況地目") = pRow.Item("調査時現況地目")
            AddRow.Item("調査時面積") = pRow.Item("調査時面積")
            AddRow.Item("所有者ID") = pRow.Item("所有者ID")
            AddRow.Item("所有者氏名") = pRow.Item("所有者氏名")
            AddRow.Item("所有者住所") = pRow.Item("所有者住所")
            AddRow.Item("所有者郵便番号") = pRow.Item("所有者郵便番号")
            AddRow.Item("所有者住民区分") = pRow.Item("所有者住民区分")
            AddRow.Item("所有者住民区分名") = pRow.Item("所有者住民区分名")
            AddRow.Item("送付先ID") = pRow.Item("送付先ID")
            AddRow.Item("送付先氏名") = pRow.Item("送付先氏名")
            AddRow.Item("送付先住所") = pRow.Item("送付先住所")
            AddRow.Item("送付先郵便番号") = pRow.Item("送付先郵便番号")
            AddRow.Item("備考") = pRow.Item("備考")

            pResultTBL.Rows.Add(AddRow)
        Next

        Return pResultTBL
    End Function

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.NoMessage
        End Get
    End Property

    Private Sub mvar印刷開始_Click(sender As Object, e As System.EventArgs) Handles mvar印刷開始.Click
        Print("印刷")
    End Sub

    Private Sub mvarExcel出力_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarExcel出力.Click
        Dim sFolder As String = SysAD.OutputFolder & String.Format("\非農地通知書{0}_{1}", Now.Year, Now.Month)

        If Not IO.Directory.Exists(sFolder) Then
            IO.Directory.CreateDirectory(sFolder)
        End If

        Print("Excel", sFolder)
    End Sub

    Private Sub Print(ByVal sOutPutType As String, Optional ByVal sFolder As String = "")
        If mvar発行番号.Text.Length = 0 Then
            MsgBox("発行番号を入力してください。")
        Else
            Dim p発送先TBL As New DataTable("発送先")

            p発送先TBL.Columns.Add("ID", GetType(Decimal))
            p発送先TBL.Columns.Add("郵便番号", GetType(String))
            p発送先TBL.Columns.Add("氏名", GetType(String))
            p発送先TBL.Columns.Add("住所", GetType(String))
            p発送先TBL.Columns.Add("通知番号", GetType(Int32))
            p発送先TBL.PrimaryKey = {p発送先TBL.Columns("ID")}

            Dim pMax通知番号 As Integer = 0
            For Each pRow As DataRow In pTBL.Rows
                If Not IsDBNull(pRow.Item("通知番号")) AndAlso pRow.Item("通知番号") > pMax通知番号 Then
                    pMax通知番号 = pRow.Item("通知番号")
                End If
            Next

            For Each pRow As DataRowView In New DataView(pTBL, "[印刷]=True", "", DataViewRowState.CurrentRows)
                If Not IsDBNull(pRow.Item("送付先ID")) AndAlso p発送先TBL.Rows.Find(pRow.Item("送付先ID")) Is Nothing Then
                    Dim pNewRow As DataRow = p発送先TBL.NewRow
                    pNewRow.Item("ID") = pRow.Item("送付先ID")
                    pNewRow.Item("郵便番号") = pRow.Item("送付先郵便番号").ToString
                    pNewRow.Item("氏名") = pRow.Item("送付先氏名").ToString
                    pNewRow.Item("住所") = pRow.Item("送付先住所").ToString
                    If Val(pRow.Item("通知番号").ToString) = 0 Then
                        pNewRow.Item("通知番号") = pMax通知番号 + 1
                        pMax通知番号 += 1
                    Else
                        pNewRow.Item("通知番号") = Val(pRow.Item("通知番号").ToString)
                    End If

                    p発送先TBL.Rows.Add(pNewRow)
                End If
            Next

            Dim sDefaultXML01 As String = HimTools2012.TextAdapter.LoadTextFile(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\非農地通知書の送付について.xml")
            Dim sDefaultXML02 As String = HimTools2012.TextAdapter.LoadTextFile(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\非農地通知書.xml")
            Dim sSavePath01 As String = SysAD.OutputFolder & "\非農地通知書の送付について.xml"
            Dim sSavePath02 As String = SysAD.OutputFolder & "\非農地通知書.xml"
            Dim sPath As String = ""

            Using pExcel As New HimTools2012.Excel.Automation.ExcelAutomation

                For Each pRow As DataRow In p発送先TBL.Rows
                    Dim sXML01 As String = sDefaultXML01
                    Dim sXML02 As String = sDefaultXML02

                    sXML01 = Replace(sXML01, "{郵便番号}", pRow.Item("郵便番号").ToString)
                    sXML01 = Replace(sXML01, "{様分}", "")
                    sXML01 = Replace(sXML01, "{発行番号}", mvar発行番号.Text)
                    sXML02 = Replace(sXML02, "{発行番号}", mvar発行番号.Text)
                    sXML01 = Replace(sXML01, "{発行年月日}", 和暦Format(mvar発行年月日.Value))
                    sXML02 = Replace(sXML02, "{発行年月日}", 和暦Format(mvar発行年月日.Value))
                    sXML01 = Replace(sXML01, "{氏名}", pRow.Item("氏名").ToString)
                    sXML02 = Replace(sXML02, "{氏名}", pRow.Item("氏名").ToString)

                    sXML01 = Replace(sXML01, "{住所}", pRow.Item("住所").ToString)
                    sXML01 = Replace(sXML01, "{通知番号}", HimTools2012.StringF.Right("00000" & pRow.Item("通知番号"), 5))
                    sXML02 = Replace(sXML02, "{通知番号}", HimTools2012.StringF.Right("00000" & pRow.Item("通知番号"), 5))

                    'sysad.db(slrdb).executesql("Update [] SET ")

                    sXML02 = Replace(sXML02, "{決定総会年月日}", 和暦Format(mvar決定総会年月日.Value))

                    Dim pView As New DataView(pTBL, "[送付先ID]=" & pRow.Item("ID"), "[所有者ID],[大字],[調査時地番]", DataViewRowState.CurrentRows)
                    Dim n As Integer = 1
                    Dim nArea As Decimal = 0

                    For Each pRV As DataRowView In pView
                        Dim sNo As String = HimTools2012.StringF.Right("00" & n, 2)
                        sXML02 = Replace(sXML02, "{大字" & sNo & "}", pRV.Item("大字").ToString)
                        sXML02 = Replace(sXML02, "{小字" & sNo & "}", pRV.Item("小字").ToString)
                        sXML02 = Replace(sXML02, "{地番" & sNo & "}", pRV.Item("調査時地番").ToString)
                        sXML02 = Replace(sXML02, "{地目" & sNo & "}", pRV.Item("調査時登記地目").ToString)
                        sXML02 = Replace(sXML02, "{面積" & sNo & "}", CDec(pRV.Item("調査時面積")).ToString("F2"))
                        nArea += Val(pRV.Item("調査時面積").ToString)

                        pRV.Item("発行番号") = mvar発行番号.Text
                        pRV.Item("発行年月日") = 和暦Format(mvar発行年月日.Value)
                        pRV.Item("通知番号") = Val(pRow.Item("通知番号").ToString)
                        pRV.Item("備考") = "出力済み"
                        n += 1
                        sXML02 = Replace(sXML02, "{所有者" & sNo & "}", pRV.Item("所有者氏名").ToString)

                        SysAD.DB(sLRDB).ExecuteSQL(String.Format("UPDATE D_非農地通知判定 SET D_非農地通知判定.発行番号 = {1}, D_非農地通知判定.発行年月日 = #{2}#, D_非農地通知判定.通知番号 = {3}, D_非農地通知判定.備考 = '出力済み' WHERE (((D_非農地通知判定.ID)={0}));", Val(pRV.Item("ID").ToString), Val(pRV.Item("発行番号").ToString), mvar発行年月日.Value, Val(pRV.Item("通知番号").ToString)))

                    Next
                    sXML02 = Replace(sXML02, "{件数}", pView.Count)
                    sXML02 = Replace(sXML02, "{面積計}", nArea.ToString("F2"))

                    For i As Integer = n To 44
                        Dim sNo As String = HimTools2012.StringF.Right("00" & i, 2)
                        sXML02 = Replace(sXML02, "{大字" & sNo & "}", "")
                        sXML02 = Replace(sXML02, "{小字" & sNo & "}", "")
                        sXML02 = Replace(sXML02, "{地番" & sNo & "}", "")
                        sXML02 = Replace(sXML02, "{地目" & sNo & "}", "")
                        sXML02 = Replace(sXML02, "{面積" & sNo & "}", "")
                        sXML02 = Replace(sXML02, "{所有者" & sNo & "}", "")
                    Next

                    If n < 14 Then
                        Dim sDelStr As String = Mid(sXML02, InStr(sXML02, " <Worksheet ss:Name=""Page02"">"))
                        sDelStr = HimTools2012.StringF.Left(sDelStr, InStr(sDelStr, " </Worksheet>") + Len(" </Worksheet>") + 1)
                        sXML02 = Replace(sXML02, sDelStr, "")
                    End If


                    Select Case sOutPutType
                        Case "Excel"
                            sPath = sFolder & String.Format("\非農地通知書の送付について({0}).xml", pRow.Item("氏名").ToString)
                            HimTools2012.TextAdapter.SaveTextFile(sPath, sXML01)

                            sPath = sFolder & String.Format("\非農地通知書({0}).xml", pRow.Item("氏名").ToString)
                            HimTools2012.TextAdapter.SaveTextFile(sPath, sXML02)

                        Case "印刷"
                            HimTools2012.TextAdapter.SaveTextFile(sSavePath02, sXML02)
                            pExcel.PrintBook(sSavePath02)
                            HimTools2012.TextAdapter.SaveTextFile(sSavePath01, sXML01)
                            pExcel.PrintBook(sSavePath01)
                    End Select

                Next

                Select Case sOutPutType
                    Case "Excel" : SysAD.ShowFolder(System.IO.Directory.GetParent(sPath).ToString)
                End Select

                mvarbSameOwner = False
                For Each pRow As DataRowView In mvarGrid.DataView
                    If Not IsDBNull(pRow.Item("印刷")) AndAlso pRow.Item("印刷") = True Then
                        pRow.Item("印刷") = False
                    End If
                Next
                mvarbSameOwner = True

                mvarGrid.Update()
                'pTBL = SetTBL(SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D_非農地通知判定]"))
                'pTBL = SetTBL(SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT D_非農地通知判定.* FROM (D_非農地通知判定 LEFT JOIN [D:個人Info] ON D_非農地通知判定.送付先ID = [D:個人Info].ID) LEFT JOIN V_行政区 ON [D:個人Info].行政区ID = V_行政区.ID WHERE (((D_非農地通知判定.調査時現況地目) In ('山林','原野','のり面'))) ORDER BY V_行政区.ID, D_非農地通知判定.送付先ID;"))
                'pTBL.Columns.Add("印刷", GetType(Boolean))


            End Using
        End If
    End Sub

    Private Sub mvar全印刷_Click(sender As Object, e As System.EventArgs) Handles mvar全印刷.Click
        mvarbSameOwner = False
        For Each pRow As DataRowView In mvarGrid.DataView
            pRow.Item("印刷") = True
        Next
        mvarbSameOwner = True
    End Sub

    Private Sub mvar全解除_Click(sender As Object, e As System.EventArgs) Handles mvar全解除.Click
        mvarbSameOwner = False
        For Each pRow As DataRowView In mvarGrid.DataView
            pRow.Item("印刷") = False
        Next
        mvarbSameOwner = True
    End Sub

    Private mvarbSameOwner As Boolean = True
    Private Sub mvarGrid_CellValueChanged(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles mvarGrid.CellValueChanged

        If mvarbSameOwner AndAlso mvarGrid.Columns(e.ColumnIndex).DataPropertyName = "印刷" Then

            Dim pB As Boolean = mvarGrid.Rows(e.RowIndex).Cells("印刷").Value
            If pB = True AndAlso Not IsDBNull(CType(mvarGrid.Rows(e.RowIndex).DataBoundItem, DataRowView).Item("送付先ID")) Then
                Dim pRows() As DataRow = pTBL.Select("[送付先ID]=" & CType(mvarGrid.Rows(e.RowIndex).DataBoundItem, DataRowView).Item("送付先ID"))
                For Each pRow As DataRow In pRows
                    pRow.Item("印刷") = True
                Next
            End If
        End If
    End Sub

    Private Sub mvar送付先出力_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar送付先出力.Click
        mvarGrid送付先.ToExcel()
    End Sub

    Private Sub mvarGrid_CellContextMenuStripNeeded(sender As Object, e As System.Windows.Forms.DataGridViewCellContextMenuStripNeededEventArgs) Handles mvarGrid.CellContextMenuStripNeeded
        If e.ColumnIndex = -1 Then
            If mvarRowMenu Is Nothing Then
                mvarRowMenu = New ContextMenuStrip
                AddHandler mvarRowMenu.Items.Add("「印刷する」に設定").Click, AddressOf Set印刷
                AddHandler mvarRowMenu.Items.Add("「未提出」に設定").Click, AddressOf SetUndispatched
                AddHandler mvarRowMenu.Items.Add("「発送済み」に設定").Click, AddressOf SetSent
            End If
            e.ContextMenuStrip = mvarRowMenu
        End If

    End Sub

    Private mvarRowMenu As ContextMenuStrip
    Private Sub mvarGrid_RowContextMenuStripNeeded(sender As Object, e As System.Windows.Forms.DataGridViewRowContextMenuStripNeededEventArgs) Handles mvarGrid.RowContextMenuStripNeeded
    End Sub
    Private Sub Set印刷()
        If mvarGrid.SelectedRows IsNot Nothing AndAlso mvarGrid.SelectedRows.Count > 0 Then
            Dim s送付先ID As New List(Of String)
            For Each pRow As DataGridViewRow In mvarGrid.SelectedRows
                If pRow.Cells("送付先ID").Value.ToString.Length > 0 AndAlso Not s送付先ID.Contains(CStr(pRow.Cells("送付先ID").Value.ToString)) Then
                    s送付先ID.Add(CStr(pRow.Cells("送付先ID").Value.ToString))
                End If
            Next
            If s送付先ID.Count > 0 Then
                For Each pRowV As DataRowView In New DataView(pTBL, "[送付先ID] IN (" & Join(s送付先ID.ToArray, ",") & ")", "", DataViewRowState.CurrentRows)
                    pRowV.Item("印刷") = True
                Next
            End If

            mvarGrid.ClearSelection()
        End If
    End Sub

    Private Sub SetUndispatched()
        If mvarGrid.SelectedRows IsNot Nothing AndAlso mvarGrid.SelectedRows.Count > 0 Then
            Dim s送付先ID As New List(Of String)
            For Each pRow As DataGridViewRow In mvarGrid.SelectedRows
                If pRow.Cells("送付先ID").Value.ToString.Length > 0 AndAlso Not s送付先ID.Contains(CStr(pRow.Cells("送付先ID").Value.ToString)) Then
                    s送付先ID.Add(CStr(pRow.Cells("送付先ID").Value.ToString))
                End If
            Next
            If s送付先ID.Count > 0 Then
                For Each pRowV As DataRowView In New DataView(pTBL, "[送付先ID] IN (" & Join(s送付先ID.ToArray, ",") & ")", "", DataViewRowState.CurrentRows)
                    SysAD.DB(sLRDB).ExecuteSQL(String.Format("UPDATE D_非農地通知判定 SET D_非農地通知判定.備考 = '未発送' WHERE (((D_非農地通知判定.ID)={0}));", pRowV.Item("ID")))
                    pRowV.Item("備考") = "未発送"
                Next
            End If

            mvarGrid.ClearSelection()
        End If
    End Sub

    Private Sub SetSent()
        If mvarGrid.SelectedRows IsNot Nothing AndAlso mvarGrid.SelectedRows.Count > 0 Then
            Dim s送付先ID As New List(Of String)
            For Each pRow As DataGridViewRow In mvarGrid.SelectedRows
                If pRow.Cells("送付先ID").Value.ToString.Length > 0 AndAlso Not s送付先ID.Contains(CStr(pRow.Cells("送付先ID").Value.ToString)) Then
                    s送付先ID.Add(CStr(pRow.Cells("送付先ID").Value.ToString))
                End If
            Next
            If s送付先ID.Count > 0 Then
                For Each pRowV As DataRowView In New DataView(pTBL, "[送付先ID] IN (" & Join(s送付先ID.ToArray, ",") & ")", "", DataViewRowState.CurrentRows)
                    SysAD.DB(sLRDB).ExecuteSQL(String.Format("UPDATE D_非農地通知判定 SET D_非農地通知判定.備考 = '発送済み' WHERE (((D_非農地通知判定.ID)={0}));", pRowV.Item("ID")))
                    pRowV.Item("備考") = "発送済み"
                Next
            End If

            mvarGrid.ClearSelection()
        End If
    End Sub

    Private Sub mvar検索_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar検索.Click
        Dim RowCount As Integer = 0
        Dim pGrid As HimTools2012.controls.DataGridViewWithDataView

        If mvarTabControl.SelectedTab.Name = "T出力" Then
            pGrid = mvarGrid
        Else
            pGrid = mvarGrid送付先
        End If

        For Each pRow As DataRowView In pGrid.DataView
            If pRow.Item("送付先氏名") = mvarTxt所有者.Text Then
                pGrid.FirstDisplayedScrollingRowIndex = RowCount
                pGrid.Rows(RowCount).Selected = True

                Exit For
            End If

            RowCount += 1
        Next

        
    End Sub
End Class
