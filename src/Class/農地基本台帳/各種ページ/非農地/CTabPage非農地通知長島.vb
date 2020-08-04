Imports System.ComponentModel
Imports HimTools2012.Excel.XMLSS2003

Public Class CTabPage非農地通知長島
    Inherits HimTools2012.controls.CTabPageWithToolStrip
    Implements HimTools2012.controls.XMLLayoutContainer


    Private WithEvents mvar全印刷 As ToolStripButton
    Private WithEvents mvar全解除 As ToolStripButton
    Private WithEvents mvar発行番号 As New ToolStripTextBox
    Public WithEvents mvar印刷開始 As ToolStripSplitButton
    Public WithEvents mvar印刷開始Excel As ToolStripMenuItem
    Private WithEvents mvar送付先出力 As New ToolStripButton("送付先一覧出力")
    Private WithEvents mvar総会資料 As New ToolStripButton("総会資料作成")

    Private WithEvents mvar決定総会年月日 As New HimTools2012.controls.ToolStripDateTimePicker
    Private WithEvents mvar発行年月日 As New HimTools2012.controls.ToolStripDateTimePicker


    Private mvarTabPage出力 As New TabPage("非農地通知出力")
    Private mvarTabPage送付先 As New TabPage("送付先一覧")
    Private WithEvents mvarGrid送付先 As New HimTools2012.controls.DataGridViewWithDataView

    Private pTBL As DataTable
    Private pTBL送付先 As New DataTable

    Protected Property mvarXMLLayout As HimTools2012.controls.XMLLayout
    Protected WithEvents mvarGrid As HimTools2012.controls.DataGridViewWithDataView

    Public Sub New()
        MyBase.New(True, True, "非農地通知", "非農地通知")

        Me.Body.SuspendLayout()
        mvarXMLLayout = New HimTools2012.controls.XMLLayout(SysAD.DB(sLRDB), ObjectMan, Me)
        With mvarXMLLayout
            .StartLayout(SysAD.SystemInfo.画面設定, "非農地通知")
            Dim mvarTabControl As TabControl = .Controls("MainTab")

            pTBL = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT D_非農地通知判定.ID, D_非農地通知判定.NID, D_非農地通知判定.一筆コード, D_非農地通知判定.発行番号, D_非農地通知判定.通知番号, D_非農地通知判定.大字, D_非農地通知判定.小字, D_非農地通知判定.調査時地番, D_非農地通知判定.調査時登記地目, D_非農地通知判定.調査時現況地目, D_非農地通知判定.調査時面積, IIf(IsNull([D:農地Info].[農振法区分]),IIf([D:農地Info].[農業振興地域]=0,'農振地域',IIf([D:農地Info].[農業振興地域]=2,'農振地域外','農用地区域')),IIf([D:農地Info].[農振法区分]=1,'農用地区域',IIf([D:農地Info].[農振法区分]=2,'農振地域',IIf([D:農地Info].[農振法区分]=3,'農振地域外','その他')))) AS 農振法, D_非農地通知判定.所有者ID, D_非農地通知判定.所有者氏名, D_非農地通知判定.所有者住所, D_非農地通知判定.所有者郵便番号, D_非農地通知判定.所有者住民区分, V_住民区分.名称 AS 住民区分, D_非農地通知判定.納税義務者ID, D_非農地通知判定.納税義務者氏名, D_非農地通知判定.納税義務者住所, D_非農地通知判定.納税義務者郵便番号, D_非農地通知判定.納税義務者住民区分, D_非農地通知判定.送付先ID, D_非農地通知判定.送付先氏名, D_非農地通知判定.送付先住所, D_非農地通知判定.送付先郵便番号, D_非農地通知判定.発行年月日 " & _
                                                       "FROM (((D_非農地通知判定 LEFT JOIN V_大字 ON D_非農地通知判定.大字 = V_大字.大字) LEFT JOIN [D:個人Info] ON D_非農地通知判定.送付先ID = [D:個人Info].ID) LEFT JOIN V_住民区分 ON D_非農地通知判定.所有者住民区分 = V_住民区分.ID) LEFT JOIN [D:農地Info] ON D_非農地通知判定.NID = [D:農地Info].ID " & _
                                                       "ORDER BY [D:個人Info].行政区ID, D_非農地通知判定.送付先ID, V_大字.ID, IIf(InStr([調査時地番],'-')>0,Left([調査時地番],InStr([調査時地番],'-')-1),[調査時地番]), IIf(InStr([調査時地番],'-')>0,Mid([調査時地番],InStr([調査時地番],'-')+1),'');")
            pTBL.Columns.Add("印刷", GetType(Boolean))
            pTBL.TableName = "出力用テーブル"

            pTBL送付先 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT V_行政区.ID AS 行政区ID, V_行政区.行政区, D_非農地通知判定.送付先ID, D_非農地通知判定.送付先氏名, D_非農地通知判定.送付先住所, D_非農地通知判定.送付先郵便番号, V_住民区分.名称 AS 住民区分 " & _
                                                             "FROM ((D_非農地通知判定 INNER JOIN [D:個人Info] ON D_非農地通知判定.送付先ID = [D:個人Info].ID) LEFT JOIN V_行政区 ON [D:個人Info].行政区ID = V_行政区.ID) LEFT JOIN V_住民区分 ON [D:個人Info].住民区分 = V_住民区分.ID " & _
                                                             "GROUP BY V_行政区.ID, V_行政区.行政区, D_非農地通知判定.送付先ID, D_非農地通知判定.送付先氏名, D_非農地通知判定.送付先住所, D_非農地通知判定.送付先郵便番号, V_住民区分.名称 " & _
                                                             "ORDER BY V_行政区.ID, D_非農地通知判定.送付先ID;")
            pTBL送付先.Columns.Add("印刷", GetType(Boolean))
            pTBL送付先.TableName = "送付先テーブル"

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

            Dim p出力Col As New DataGridViewCheckBoxColumn
            With p出力Col
                .HeaderText = "印刷"
                .DataPropertyName = "印刷"
                .Name = "印刷"
            End With

            mvarGrid送付先 = .Controls("G送付先")
            With mvarGrid送付先
                .Columns.Add(p出力Col)
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                .VirtualMode = True
                .AutoGenerateColumns = True
                .AllowUserToAddRows = False
                .SetDataView(pTBL送付先, "", "")
            End With

            mvar全印刷 = New ToolStripButton("全選択")
            mvar全解除 = New ToolStripButton("全解除")
            mvar印刷開始 = New ToolStripSplitButton("印刷開始")
            mvar印刷開始Excel = New ToolStripMenuItem("印刷開始(Excel)")

            Me.ToolStrip.Items.AddRange({New ToolStripSeparator, mvar全印刷, mvar全解除, New ToolStripSeparator,
                                         New ToolStripLabel("決定総会年月日"), mvar決定総会年月日, New ToolStripSeparator,
                                         New ToolStripLabel("発行年月日"), mvar発行年月日, New ToolStripSeparator,
                                         New ToolStripLabel("発行番号"), mvar発行番号, mvar印刷開始, New ToolStripSeparator,
                                         mvar送付先出力, New ToolStripSeparator, mvar総会資料})
            mvar印刷開始.DropDownItems.Add(mvar印刷開始Excel)
            mvar決定総会年月日.Value = Now.Date
            mvar発行年月日.Value = Now.Date
        End With

        If System.IO.File.Exists(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\非農地通知書出力履歴(送付先).xml") Then
            If MsgBox("前回の出力履歴を表示しますか？", vbOKCancel) = vbOK Then
                Dim reader As IO.StreamReader = New IO.StreamReader(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\非農地通知書出力履歴(送付先).xml", System.Text.Encoding.GetEncoding("Shift_Jis"))
                XMLCheck(reader, "送付先", pTBL送付先)
                reader = New IO.StreamReader(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\非農地通知書出力履歴(出力用).xml", System.Text.Encoding.GetEncoding("Shift_Jis"))
                XMLCheck(reader, "出力用", pTBL)
            Else
                mvarGrid送付先.SetDataView(pTBL送付先, "", "")
                mvarGrid.SetDataView(pTBL, "", "")
            End If
        Else
            mvarGrid送付先.SetDataView(pTBL送付先, "", "")
            mvarGrid.SetDataView(pTBL, "", "")
        End If

        Me.Body.ResumeLayout()
    End Sub

    Private Sub mvar印刷開始_ButtonClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar印刷開始.ButtonClick
        PrintStart(False)
    End Sub
    Private Sub mvar印刷開始Excel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar印刷開始Excel.Click
        PrintStart(True)
    End Sub
    Private Sub PrintStart(ByVal pExcel出力 As Boolean)
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
                    pNewRow.Item("通知番号") = pMax通知番号 + 1
                    pMax通知番号 += 1

                    p発送先TBL.Rows.Add(pNewRow)
                End If
            Next

            Dim sDefaultXML01 As String = HimTools2012.TextAdapter.LoadTextFile(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\非農地通知書の送付について.xml")
            Dim sDefaultXML02 As String = HimTools2012.TextAdapter.LoadTextFile(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\非農地通知書.xml")
            Dim sSavePath01 As String = SysAD.OutputFolder & "\非農地通知書の送付について.xml"
            Dim sSavePath02 As String = SysAD.OutputFolder & "\非農地通知書.xml"

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

                    Dim pView As New DataView(pTBL, "[印刷]=True AND [送付先ID]=" & pRow.Item("ID"), "[所有者ID],[大字],[調査時地番]", DataViewRowState.CurrentRows)
                    Dim n As Integer = 1
                    Dim nArea As Decimal = 0

                    For Each pRV As DataRowView In pView
                        Dim sNo As String = HimTools2012.StringF.Right("00" & n, 2)
                        sXML02 = Replace(sXML02, "{大字" & sNo & "}", pRV.Item("大字"))
                        sXML02 = Replace(sXML02, "{小字" & sNo & "}", pRV.Item("小字"))
                        sXML02 = Replace(sXML02, "{地番" & sNo & "}", pRV.Item("調査時地番"))
                        sXML02 = Replace(sXML02, "{地目" & sNo & "}", pRV.Item("調査時登記地目"))
                        sXML02 = Replace(sXML02, "{現況地目" & sNo & "}", pRV.Item("調査時現況地目"))
                        sXML02 = Replace(sXML02, "{面積" & sNo & "}", CDec(pRV.Item("調査時面積")).ToString("F2"))
                        nArea += pRV.Item("調査時面積")

                        pRV.Item("発行番号") = mvar発行番号.Text
                        pRV.Item("通知番号") = pRow.Item("通知番号")
                        pRV.Item("発行年月日") = 和暦Format(mvar発行年月日.Value)
                        n += 1
                        sXML02 = Replace(sXML02, "{所有者" & sNo & "}", pRV.Item("所有者氏名").ToString)
                    Next
                    sXML02 = Replace(sXML02, "{件数}", pView.Count)
                    sXML02 = Replace(sXML02, "{面積計}", nArea.ToString("F2"))

                    For i As Integer = n To 44
                        Dim sNo As String = HimTools2012.StringF.Right("00" & i, 2)
                        sXML02 = Replace(sXML02, "{大字" & sNo & "}", "")
                        sXML02 = Replace(sXML02, "{小字" & sNo & "}", "")
                        sXML02 = Replace(sXML02, "{地番" & sNo & "}", "")
                        sXML02 = Replace(sXML02, "{地目" & sNo & "}", "")
                        sXML02 = Replace(sXML02, "{現況地目" & sNo & "}", "")
                        sXML02 = Replace(sXML02, "{面積" & sNo & "}", "")
                        sXML02 = Replace(sXML02, "{所有者" & sNo & "}", "")
                    Next

                    If n < 14 Then
                        Dim sDelStr As String = Mid(sXML02, InStr(sXML02, " <Worksheet ss:Name=""Page02"">"))
                        sDelStr = HimTools2012.StringF.Left(sDelStr, InStr(sDelStr, " </Worksheet>") + Len(" </Worksheet>") + 1)
                        sXML02 = Replace(sXML02, sDelStr, "")
                    End If

                    If pExcel出力 Then
                        sSavePath02 = SysAD.OutputFolder & String.Format("\非農地通知書({0}).xml", pRow.Item("氏名").ToString)
                        HimTools2012.TextAdapter.SaveTextFile(sSavePath02, sXML02)
                        sSavePath01 = SysAD.OutputFolder & String.Format("\非農地通知書の送付について({0}).xml", pRow.Item("氏名").ToString)
                        HimTools2012.TextAdapter.SaveTextFile(sSavePath01, sXML01)
                    Else
                        HimTools2012.TextAdapter.SaveTextFile(sSavePath02, sXML02)
                        pExcel.PrintBook(sSavePath02)
                        HimTools2012.TextAdapter.SaveTextFile(sSavePath01, sXML01)
                        pExcel.PrintBook(sSavePath01)
                    End If
                Next

                'DataTableのデータをXMLに書き込む
                Dim SW As System.IO.StreamWriter = New System.IO.StreamWriter(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\非農地通知書出力履歴(送付先).xml", False, System.Text.Encoding.GetEncoding("Shift_Jis"))
                pTBL送付先.WriteXml(SW)
                SW = New System.IO.StreamWriter(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\非農地通知書出力履歴(出力用).xml", False, System.Text.Encoding.GetEncoding("Shift_Jis"))
                pTBL.WriteXml(SW)

                If pExcel出力 Then
                    System.Diagnostics.Process.Start(System.IO.Directory.GetParent(sSavePath02).ToString)
                End If
            End Using

            MsgBox("出力が終了しました。")
        End If
    End Sub

    Private Sub XMLCheck(ByRef reader As IO.StreamReader, ByVal sKey As String, ByRef pTBL As DataTable)
        Try
            pTBL.Clear()
            pTBL.ReadXml(reader)
            Select Case sKey
                Case "送付先" : mvarGrid送付先.SetDataView(pTBL, "", "")
                Case "出力用" : mvarGrid.SetDataView(pTBL, "", "")
            End Select


        Catch ex As Exception
            MsgBox("履歴XMLのデータがない、あるいは破損しているためデータベースから読み込みます。")
            reader.Close()
            If System.IO.File.Exists(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & String.Format("\非農地通知書出力履歴({0})BK.xml", sKey)) Then
                My.Computer.FileSystem.DeleteFile(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & String.Format("\非農地通知書出力履歴({0})BK.xml", sKey))
            End If
            My.Computer.FileSystem.RenameFile(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & String.Format("\非農地通知書出力履歴({0}).xml", sKey), String.Format("非農地通知書出力履歴({0})BK.xml", sKey))
            'SetpTBL(pTBL)
            mvarGrid.SetDataView(pTBL, "", "")
        End Try
    End Sub
    Private Sub SetpTBL(ByRef pTBL As DataTable)
        For Each pRow As DataRow In pTBL.Rows
            Dim pNRow As DataRow = pTBL.Rows.Find(pRow.Item("ID"))
            If pNRow Is Nothing Then
                pNRow = pTBL.NewRow

                pNRow.Item("行政区ID") = Val(pRow.Item("行政区ID").ToString)
                pNRow.Item("行政区") = pRow.Item("行政区").ToString
                pNRow.Item("所有者ID") = Val(pRow.Item("ID").ToString)
                pNRow.Item("所有者氏名") = pRow.Item("氏名").ToString
                pNRow.Item("所有者フリガナ") = pRow.Item("フリガナ").ToString
                pNRow.Item("所有者住所") = pRow.Item("住所").ToString
                pNRow.Item("所有者郵便番号") = pRow.Item("郵便番号").ToString
                pNRow.Item("住民区分") = pRow.Item("住民区分名").ToString
                pNRow.Item("集計フラグ") = 1

                pTBL.Rows.Add(pNRow)
            End If
        Next
    End Sub

    Private Sub mvar全印刷_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar全印刷.Click
        mvarbSameOwner = False
        For Each pRow As DataRow In pTBL.Rows
            pRow.Item("印刷") = True
        Next
        mvarbSameOwner = True
    End Sub

    Private Sub mvar全解除_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar全解除.Click
        mvarbSameOwner = False
        For Each pRow As DataRow In pTBL.Rows
            pRow.Item("印刷") = False
        Next
        mvarbSameOwner = True
    End Sub

    Private mvarbSameOwner As Boolean = True
    Private Sub mvarGrid_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles mvarGrid.CellValueChanged

        If mvarbSameOwner AndAlso mvarGrid.Columns(e.ColumnIndex).DataPropertyName = "印刷" Then

            Dim pB As Boolean = mvarGrid.Rows(e.RowIndex).Cells("印刷").Value
            If pB = True AndAlso Not IsDBNull(CType(mvarGrid.Rows(e.RowIndex).DataBoundItem, DataRowView).Item("送付先ID")) Then
                Dim pRows() As DataRow = pTBL.Select("[送付先ID]=" & CType(mvarGrid.Rows(e.RowIndex).DataBoundItem, DataRowView).Item("送付先ID"))
                For Each pRow As DataRow In pRows
                    pRow.Item("印刷") = True
                Next
            Else
                'Dim pRows() As DataRow = pTBL.Select("[送付先ID]=" & CType(mvarGrid.Rows(e.RowIndex).DataBoundItem, DataRowView).Item("送付先ID"))
                'For Each pRow As DataRow In pRows
                '    pRow.Item("印刷") = False
                'Next
            End If
        End If
    End Sub

    Private Sub mvarGrid送付先_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles mvarGrid送付先.CellValueChanged

        If mvarbSameOwner AndAlso mvarGrid送付先.Columns(e.ColumnIndex).DataPropertyName = "印刷" Then

            Dim pB As Boolean = mvarGrid送付先.Rows(e.RowIndex).Cells("印刷").Value
            If pB = True AndAlso Not IsDBNull(CType(mvarGrid送付先.Rows(e.RowIndex).DataBoundItem, DataRowView).Item("送付先ID")) Then
                Dim pRows() As DataRow = pTBL.Select("[送付先ID]=" & CType(mvarGrid送付先.Rows(e.RowIndex).DataBoundItem, DataRowView).Item("送付先ID"))
                For Each pRow As DataRow In pRows
                    pRow.Item("印刷") = True
                Next
            Else
                Dim pRows() As DataRow = pTBL.Select("[送付先ID]=" & CType(mvarGrid送付先.Rows(e.RowIndex).DataBoundItem, DataRowView).Item("送付先ID"))
                For Each pRow As DataRow In pRows
                    pRow.Item("印刷") = False
                Next
            End If
        End If
    End Sub

    Private Sub mvar送付先出力_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar送付先出力.Click
        mvarGrid送付先.ToExcel()
    End Sub

    Private Sub mvarGrid_CellContextMenuStripNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellContextMenuStripNeededEventArgs) Handles mvarGrid.CellContextMenuStripNeeded
        If e.ColumnIndex = -1 Then
            If mvarRowMenu Is Nothing Then
                mvarRowMenu = New ContextMenuStrip
                AddHandler mvarRowMenu.Items.Add("「印刷する」に設定").Click, AddressOf Set印刷

            End If
            e.ContextMenuStrip = mvarRowMenu
        End If

    End Sub

    Private mvarRowMenu As ContextMenuStrip
    Private Sub mvarGrid_RowContextMenuStripNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowContextMenuStripNeededEventArgs) Handles mvarGrid.RowContextMenuStripNeeded
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

    Private Sub mvar総会資料_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar総会資料.Click
        If pTBL送付先 IsNot Nothing AndAlso pTBL送付先.Rows.Count > 0 Then
            Dim sFile As String = My.Application.Info.DirectoryPath & "\" & SysAD.市町村.市町村名 & "\非農地通知書について.xml"
            Dim Int議案番号 As Integer = InputBox("議案番号入力", "議案番号を入力してください", 0)
            Dim RowCount As Integer = 1

            If SysAD.IsClickOnceDeployed Then
                sFile = SysAD.ClickOnceSetupPath & "\" & SysAD.市町村.市町村名 & "\非農地通知書について.xml"
            End If

            If IO.File.Exists(sFile) Then
                Dim sXMLOg As String = HimTools2012.TextAdapter.LoadTextFile(sFile)
                Dim sXML As String = sXMLOg

                Dim pXMLSS As New CXMLSS2003(sXML)
                Dim pSheet01 As XMLSSWorkSheet = pXMLSS.WorkBook.WorkSheets.Items("明細")
                Dim pLoopRow01 As New XMLLoopRows(pSheet01, "{No}")
                Dim nLoop As Integer = 1
                Dim pXRow01 As XMLSSRow = Nothing

                pSheet01.ValueReplace("{議案番号}", Int議案番号)

                For Each pRow As DataRowView In New DataView(pTBL送付先, "[印刷]=True", "", DataViewRowState.CurrentRows)
                    If nLoop = 1 Then
                        pXRow01 = pSheet01.FindRowInstrText("{No}")(0)
                    Else
                        For Each pXRow As XMLSSRow In pLoopRow01
                            pXRow01 = pXRow.CopyRow
                            pSheet01.Table.Rows.InsertRow(pLoopRow01.InsetRow, pXRow01)
                            pLoopRow01.InsetRow += 1
                        Next
                    End If

                    pSheet01.ValueReplace("{No}", nLoop)
                    pSheet01.ValueReplace("{申請者Ａ氏名}", pRow.Item("送付先氏名").ToString)
                    pSheet01.ValueReplace("{申請者Ａ住所}", pRow.Item("送付先住所").ToString)
                    pSheet01.ValueReplace("{申請者X氏名}", pRow.Item("送付先氏名").ToString)
                    pSheet01.ValueReplace("{申請者X住所}", pRow.Item("送付先住所").ToString)

                    Dim pView As New DataView(pTBL, "[印刷]=True AND [送付先ID] IN (" & pRow.Item("送付先ID") & ")", "", DataViewRowState.CurrentRows)
                    Dim s土地所在 As String = ""
                    Dim s登記地目 As String = ""
                    Dim s現況地目 As String = ""
                    Dim s面積 As String = ""
                    Dim Dec面積計 As Decimal = 0
                    For n As Integer = 1 To pView.Count
                        Dim pRowV As DataRowView = pView(n - 1)
                        Select Case s土地所在
                            Case ""
                                s土地所在 = Replace地番(pRowV.Item("大字").ToString & pRowV.Item("小字").ToString & pRowV.Item("調査時地番").ToString)
                                s登記地目 = pRowV.Item("調査時登記地目").ToString
                                s現況地目 = pRowV.Item("調査時現況地目").ToString
                                s面積 = Format(pRowV.Item("調査時面積"), "#,###")
                                Dec面積計 += pRowV.Item("調査時面積")
                            Case Else
                                s土地所在 = s土地所在 & vbCrLf & Replace地番(pRowV.Item("大字").ToString & pRowV.Item("小字").ToString & pRowV.Item("調査時地番").ToString)
                                s登記地目 = s登記地目 & vbCrLf & pRowV.Item("調査時登記地目").ToString
                                s現況地目 = s現況地目 & vbCrLf & pRowV.Item("調査時現況地目").ToString
                                s面積 = s面積 & vbCrLf & Format(pRowV.Item("調査時面積"), "#,###")
                                Dec面積計 += pRowV.Item("調査時面積")
                        End Select
                    Next

                    pSheet01.ValueReplace("{土地の所在}", s土地所在)
                    pSheet01.ValueReplace("{筆数計}", pView.Count)
                    pSheet01.ValueReplace("{登記地目}", s登記地目)
                    pSheet01.ValueReplace("{現況地目}", s現況地目)
                    pSheet01.ValueReplace("{面積}", s面積)
                    pSheet01.ValueReplace("{面積計}", Format(Dec面積計, "#,###.00"))
                    pSheet01.ValueReplace("{備考}", "")

                    RowCount += 1
                    nLoop += 1
                Next

                Dim sOutPutFile As String = SysAD.OutputFolder & String.Format("\{0}_非農地通知書について.xml", 和暦Format(Now))
                HimTools2012.TextAdapter.SaveTextFile(sOutPutFile, pXMLSS.OutPut(False))
                SysAD.ShowFolder(sOutPutFile)

                'DataTableのデータをXMLに書き込む
                Dim SW As System.IO.StreamWriter = New System.IO.StreamWriter(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\非農地通知書出力履歴(送付先).xml", False, System.Text.Encoding.GetEncoding("Shift_Jis"))
                pTBL送付先.WriteXml(SW)
                SW = New System.IO.StreamWriter(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\非農地通知書出力履歴(出力用).xml", False, System.Text.Encoding.GetEncoding("Shift_Jis"))
                pTBL.WriteXml(SW)

                MsgBox("終了しました。")
            Else
                MsgBox("指定様式がありません。「非農地通知書について.xml」を作成してDeploymentPlaceに保存してください。")
            End If
        End If
    End Sub

    Private Function Replace地番(ByVal 土地所在 As String) As String
        Dim 土地所在B As String = ""
        Dim CountB As Integer = 0

        If Not IsDBNull(土地所在) AndAlso 土地所在 <> "" Then
            If InStr(土地所在, "-") > 0 Then
                For n As Integer = 1 To Len(土地所在)
                    Select Case Mid(土地所在, n, 1)
                        Case "-"
                            If CountB = 0 Then
                                土地所在B = 土地所在B & Replace(Mid(土地所在, n, 1), "-", "番")
                                CountB += 1
                            Else : 土地所在B = 土地所在B & Replace(Mid(土地所在, n, 1), "-", "の")
                            End If
                        Case Else : 土地所在B = 土地所在B & Mid(土地所在, n, 1)
                    End Select
                Next
            Else : 土地所在B = 土地所在 & "番"
            End If
        End If

        Return 土地所在B
    End Function


End Class
