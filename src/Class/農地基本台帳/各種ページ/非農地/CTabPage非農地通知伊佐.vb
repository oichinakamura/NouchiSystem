Imports System.ComponentModel
Imports HimTools2012.Excel.XMLSS2003

Public Class CTabPage非農地通知伊佐
    Inherits HimTools2012.controls.CTabPageWithToolStrip
    Implements HimTools2012.controls.XMLLayoutContainer


    Private WithEvents mvar全印刷 As ToolStripButton
    Private WithEvents mvar全解除 As ToolStripButton
    Private WithEvents mvar発行番号 As New ToolStripTextBox
    Public WithEvents mvar印刷開始 As ToolStripSplitButton
    Public WithEvents mvar印刷開始Excel As ToolStripMenuItem
    Private WithEvents mvar送付先出力 As New ToolStripButton("送付先一覧出力")
    Private WithEvents mvar宛名 As New ToolStripButton("宛名作成")
    Private WithEvents mvar対象農地更新 As ToolStripButton

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

            pTBL = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT D_非農地通知判定.ID, D_非農地通知判定.NID, D_非農地通知判定.一筆コード, D_非農地通知判定.発行番号, D_非農地通知判定.通知番号, D_非農地通知判定.大字, D_非農地通知判定.小字, D_非農地通知判定.調査時地番, D_非農地通知判定.調査時登記地目, D_非農地通知判定.調査時現況地目, D_非農地通知判定.調査時面積, IIf(IsNull([D:農地Info].[農振法区分]),IIf([D:農地Info].[農業振興地域]=0,'農振地域',IIf([D:農地Info].[農業振興地域]=2,'農振地域外','農用地区域')),IIf([D:農地Info].[農振法区分]=1,'農用地区域',IIf([D:農地Info].[農振法区分]=2,'農振地域',IIf([D:農地Info].[農振法区分]=3,'農振地域外','その他')))) AS 農振法, D_非農地通知判定.所有者ID, D_非農地通知判定.所有者氏名, D_非農地通知判定.所有者住所, D_非農地通知判定.所有者郵便番号, D_非農地通知判定.所有者住民区分, V_住民区分.名称 AS 住民区分, D_非農地通知判定.納税義務者ID, D_非農地通知判定.納税義務者氏名, D_非農地通知判定.納税義務者住所, D_非農地通知判定.納税義務者郵便番号, D_非農地通知判定.納税義務者住民区分, D_非農地通知判定.送付先ID, D_非農地通知判定.送付先氏名, D_非農地通知判定.送付先住所, D_非農地通知判定.送付先郵便番号, D_非農地通知判定.発行年月日 " &
                                                       "FROM (((D_非農地通知判定 LEFT JOIN V_大字 ON D_非農地通知判定.大字 = V_大字.大字) LEFT JOIN [D:個人Info] ON D_非農地通知判定.送付先ID = [D:個人Info].ID) LEFT JOIN V_住民区分 ON D_非農地通知判定.所有者住民区分 = V_住民区分.ID) LEFT JOIN [D:農地Info] ON D_非農地通知判定.NID = [D:農地Info].ID " &
                                                       "ORDER BY [D:個人Info].行政区ID, D_非農地通知判定.送付先ID, V_大字.ID, IIf(InStr([調査時地番],'-')>0,Left([調査時地番],InStr([調査時地番],'-')-1),[調査時地番]), IIf(InStr([調査時地番],'-')>0,Mid([調査時地番],InStr([調査時地番],'-')+1),'');")
            pTBL.Columns.Add("印刷", GetType(Boolean))
            pTBL.TableName = "出力用テーブル"

            pTBL送付先 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT V_行政区.ID AS 行政区ID, V_行政区.行政区, D_非農地通知判定.送付先ID, D_非農地通知判定.送付先氏名, D_非農地通知判定.送付先住所, D_非農地通知判定.送付先郵便番号, V_住民区分.名称 AS 住民区分, D_非農地通知判定.所有者氏名, D_非農地通知判定.通知番号 " &
                                                             "FROM ((D_非農地通知判定 INNER JOIN [D:個人Info] ON D_非農地通知判定.送付先ID = [D:個人Info].ID) LEFT JOIN V_行政区 ON [D:個人Info].行政区ID = V_行政区.ID) LEFT JOIN V_住民区分 ON [D:個人Info].住民区分 = V_住民区分.ID " &
                                                             "GROUP BY V_行政区.ID, V_行政区.行政区, D_非農地通知判定.送付先ID, D_非農地通知判定.送付先氏名, D_非農地通知判定.送付先住所, D_非農地通知判定.送付先郵便番号, V_住民区分.名称, D_非農地通知判定.所有者氏名, D_非農地通知判定.通知番号 " &
                                                             "ORDER BY V_行政区.ID, D_非農地通知判定.送付先ID;")
            pTBL送付先.Columns.Add("印刷", GetType(Boolean))
            pTBL送付先.PrimaryKey = {pTBL送付先.Columns("送付先ID")}
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
            mvar対象農地更新 = New ToolStripButton("対象農地の更新")

            Me.ToolStrip.Items.AddRange({New ToolStripSeparator, mvar全印刷, mvar全解除, New ToolStripSeparator,
                                         New ToolStripLabel("決定総会年月日"), mvar決定総会年月日, New ToolStripSeparator,
                                         New ToolStripLabel("発行年月日"), mvar発行年月日, New ToolStripSeparator,
                                         New ToolStripLabel("発行番号"), mvar発行番号, mvar印刷開始, New ToolStripSeparator,
                                         mvar送付先出力, New ToolStripSeparator, mvar宛名, New ToolStripSeparator, mvar対象農地更新})
            mvar印刷開始.DropDownItems.Add(mvar印刷開始Excel)
            mvar決定総会年月日.Value = Now.Date
            mvar発行年月日.Value = Now.Date
        End With

        Dim nNo As Integer = 1
        For Each pRow As DataRow In pTBL送付先.Rows
            pRow.Item("通知番号") = nNo

            nNo += 1
        Next

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
            p発送先TBL.Columns.Add("所有者氏名", GetType(String))
            p発送先TBL.Columns.Add("住所", GetType(String))
            p発送先TBL.Columns.Add("通知番号", GetType(Int32))
            p発送先TBL.PrimaryKey = {p発送先TBL.Columns("ID")}

            'Dim pMax通知番号 As Integer = 0
            'For Each pRow As DataRow In pTBL.Rows
            '    If Not IsDBNull(pRow.Item("通知番号")) AndAlso pRow.Item("通知番号") > pMax通知番号 Then
            '        pMax通知番号 = pRow.Item("通知番号")
            '    End If
            'Next

            For Each pRow As DataRowView In New DataView(pTBL, "[印刷]=True", "", DataViewRowState.CurrentRows)
                If Not IsDBNull(pRow.Item("送付先ID")) AndAlso p発送先TBL.Rows.Find(pRow.Item("送付先ID")) Is Nothing Then
                    Dim pNewRow As DataRow = p発送先TBL.NewRow
                    pNewRow.Item("ID") = pRow.Item("送付先ID")
                    pNewRow.Item("郵便番号") = pRow.Item("送付先郵便番号").ToString
                    pNewRow.Item("氏名") = pRow.Item("送付先氏名").ToString
                    pNewRow.Item("住所") = pRow.Item("送付先住所").ToString
                    pNewRow.Item("所有者氏名") = pRow.Item("所有者氏名").ToString

                    Dim fRow As DataRow = pTBL送付先.Rows.Find(pRow.Item("送付先ID"))
                    pNewRow.Item("通知番号") = fRow.Item("通知番号").ToString
                    'pMax通知番号 += 1

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
                    sXML01 = Replace(sXML01, "{送付先氏名}", pRow.Item("氏名").ToString)
                    sXML01 = Replace(sXML01, "{所有者名}", pRow.Item("所有者氏名").ToString)

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
                        SysAD.DB(sLRDB).ExecuteSQL(String.Format("UPDATE D_非農地通知判定 SET D_非農地通知判定.通知番号 = {1}, D_非農地通知判定.発行年月日 = Now() WHERE (((D_非農地通知判定.ID)={0}));", Val(pRV.Item("ID").ToString), Val(pRow.Item("通知番号"))))
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

    Private Sub mvar宛名_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar宛名.Click
        If IO.File.Exists(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\宛名シール.xml") Then
            Dim sT As String = HimTools2012.TextAdapter.LoadTextFile(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\宛名シール.xml")
            Dim sT2 As String = sT.Substring(sT.IndexOf(" <Worksheet ss:Name=""Page2"">"))
            Dim sT3 As String = sT2.Substring(0, sT2.LastIndexOf("</Worksheet>") + Len("</Worksheet>") + 2)
            sT = sT.Replace(sT3, "{X}")
            'Dim pExcel As Object = CreateObject("Excel.Application")
            Dim sPage(200) As String

            Dim sWhere As String = ""
            Dim nCount As Integer = 0
            Dim sOutPut As String = sT
            Dim sPath As String = SysAD.OutputFolder & "\宛名シール.xml"

            Dim MaxPage As Integer = (pTBL送付先.Rows.Count - 12) \ 12 + 1 - (((pTBL送付先.Rows.Count - 12) Mod 12) > 0)
            Dim nMax As Integer = 12

            For nPage As Integer = 2 To 200
                sPage(nPage) = sT3.Replace("{PageNo}", nPage).Replace("Page2", "Page" & nPage)

                If nPage >= 3 Then
                    For nL As Integer = 13 To 24
                        sPage(nPage) = KStrLineRep(sPage(nPage), nL, nL + 12 * (nPage - 2))
                    Next
                End If
            Next

            Dim SG As String = ""
            For nSG As Integer = 2 To MaxPage
                SG = SG & sPage(nSG)
            Next
            nMax = (MaxPage - 1) * 12 + 12 + 15 '?
            sOutPut = sOutPut.Replace("{X}", SG)

            Using pExcel As New HimTools2012.Excel.Automation.ExcelAutomation
                For Each pRow As DataRow In pTBL送付先.Rows
                    sOutPut = sOutPut.Replace("{郵便番号" & (nCount + 1) & "}", pRow.Item("送付先郵便番号").ToString)
                    sOutPut = sOutPut.Replace("{住所" & (nCount + 1) & "}", pRow.Item("送付先住所").ToString)
                    sOutPut = sOutPut.Replace("{送付先氏名" & (nCount + 1) & "}", pRow.Item("送付先氏名").ToString)
                    sOutPut = sOutPut.Replace("{所有者氏名" & (nCount + 1) & "}", pRow.Item("所有者氏名").ToString)
                    sOutPut = sOutPut.Replace("{発行番号" & (nCount + 1) & "}", pRow.Item("所有者氏名").ToString)

                    nCount += 1
                Next

                For X As Integer = 1 To nMax
                    sOutPut = sOutPut.Replace("{郵便番号" & X & "}", "")
                    sOutPut = sOutPut.Replace("{住所" & X & "}", "")
                    sOutPut = sOutPut.Replace("{送付先氏名" & X & "}", "")
                    sOutPut = sOutPut.Replace("{所有者氏名" & X & "}", "")
                    sOutPut = sOutPut.Replace("{発行番号" & X & "}", "")
                Next

                sOutPut = sOutPut.Replace("{X}", "")

                HimTools2012.TextAdapter.SaveTextFile(sPath, sOutPut)

                'Dim pBook As Object = pExcel.Workbooks.Open(sPath)
                pExcel.ShowPreview(sPath)
            End Using


        Else
            MsgBox("指定されたフォルダにＸＭＬファイルがありません")

        End If
    End Sub

    Public Function KStrLineRep(ByVal sPage As String, ByVal n As Integer, ByVal n2 As Integer) As String
        sPage = sPage.Replace("{郵便番号" & n & "}", "{郵便番号" & n2 & "}")
        sPage = sPage.Replace("{送付先氏名" & n & "}", "{送付先氏名" & n2 & "}")
        sPage = sPage.Replace("{所有者氏名" & n & "}", "{所有者氏名" & n2 & "}")
        sPage = sPage.Replace("{住所" & n & "}", "{住所" & n2 & "}")
        sPage = sPage.Replace("{通知番号" & n & "}", "{通知番号" & n2 & "}")

        Return sPage
    End Function

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

    Private Sub mvar対象農地更新_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar対象農地更新.Click
        If MsgBox("非農地の更新を行ってもよろしいですか？", vbOKCancel) = vbOK Then
            SysAD.DB(sLRDB).ExecuteSQL("DELETE D_非農地通知判定.ID FROM D_非農地通知判定;")
            SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO D_非農地通知判定 ( ID, NID, 一筆コード, 大字, 小字, 調査時地番, 調査時登記地目, 調査時現況地目, 調査時面積, 所有者ID, 所有者氏名, 所有者住所, 所有者郵便番号, 所有者住民区分, 送付先ID, 送付先氏名, 送付先住所, 送付先郵便番号 ) VALUES ( [D:農地Info].ID, [D:農地Info].ID, [D:農地Info].一筆コード, V_大字.名称, V_小字.名称, [D:農地Info].地番, V_地目.名称, V_現況地目.名称, [D:農地Info].登記簿面積, [D:農地Info].所有者ID, [D:個人Info].氏名, [D:個人Info].住所, [D:個人Info].郵便番号, [D:個人Info].住民区分, [D:農地Info].所有者ID, [D:個人Info].氏名, [D:個人Info].住所, [D:個人Info].郵便番号 ) FROM (((([D:農地Info] LEFT JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) LEFT JOIN V_小字 ON [D:農地Info].小字ID = V_小字.ID) LEFT JOIN V_地目 ON [D:農地Info].登記簿地目 = V_地目.ID) LEFT JOIN V_現況地目 ON [D:農地Info].現況地目 = V_現況地目.ID) LEFT JOIN [D:個人Info] ON [D:農地Info].所有者ID = [D:個人Info].ID WHERE ((([D:農地Info].利用状況調査荒廃)=2));")
        End If
    End Sub
End Class
