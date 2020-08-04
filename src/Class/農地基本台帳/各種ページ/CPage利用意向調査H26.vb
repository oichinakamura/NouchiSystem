'/20160229霧島
'農用地・農用地外…  1 ・ 2( ｱ ･ ｲ ) ・  3 ・ 4(      　　　　　　　　　　　　　　　  　　　　　　　　　　　　　　　　   　　　　 　　   　　      )
'農振外…  　　 2( ｱ ･ ｲ ) ・  3 ・ 4(      　　　　　　　　　　　　　　　  　　　　　　　　　　　　　　　　   　　　　 　　   　　      )


Public Class CPage利用意向調査H26
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Public WithEvents mvarTabCtrl As HimTools2012.controls.TabControlBase

    Public Sub New()
        MyBase.New(True, False, "利用意向調査出力", "利用意向調査出力")
        mvarTabCtrl = New HimTools2012.controls.TabControlBase()
        Me.ControlPanel.Add(mvarTabCtrl)

        mvarTab市内住民 = New CTabPage利用意向調査Page("市内住民", "", "(([D:個人Info].住民区分)=0)")
        mvarTabその他住民 = New CTabPage利用意向調査Page("その他住民", "", "(([D:個人Info].住民区分) Is Null Or ([D:個人Info].住民区分)<>0)")

        Me.SuspendLayout()
        mvarTabCtrl.TabPages.Add(mvarTab市内住民)
        mvarTabCtrl.AddPage(mvarTabその他住民)
        mvarTabCtrl.SelectedTab = mvarTab市内住民
        mvarTabCtrl_SelectedIndexChanged(Nothing, Nothing)
        Me.ResumeLayout()

    End Sub

    Private mvarTab市内住民 As CTabPage利用意向調査Page
    Private mvarTabその他住民 As CTabPage利用意向調査Page

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.NoMessage
        End Get
    End Property

    Private Sub mvarTabCtrl_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles mvarTabCtrl.SelectedIndexChanged
        With CType(mvarTabCtrl.SelectedTab, CTabPage利用意向調査Page)
            If .Init Then
                .LoadData()
            End If
        End With
    End Sub
End Class

Public Class CTabPage利用意向調査Page
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Private mvarTabControl As New TabControl
    Private mvarTPage印刷 As New TabPage
    Private mvarTPage明細 As New TabPage
    Private mvarTSContainer As New ToolStripContainer
    Private mvarTStrip As New ToolStrip
    Private mvarDTSContainer As New ToolStripContainer
    Private mvarDTStrip As New ToolStrip
    Private mvarDGrid As New HimTools2012.controls.DataGridViewWithDataView

    Private mvarTreeTabControl As New TabControl
    Private mvarTPage行政区 As New TabPage
    Private mvarTPage住民区分 As New TabPage
    Private WithEvents mvarTree As TreeView
    Private WithEvents mvarTree区分 As TreeView
    Private mvarSP As SplitContainer
    Private mvarG As HimTools2012.controls.DataGridViewWithDataView
    Private WithEvents mvarSelectAll As ToolStripButton
    Private WithEvents mvarSelectClear As ToolStripButton
    Private WithEvents mvarSelectRev As ToolStripButton
    Public WithEvents mvarPrint登記面積 As ToolStripSplitButton
    Public WithEvents mvarExcel登記面積 As ToolStripMenuItem
    Public WithEvents mvarPrint実面積 As ToolStripSplitButton
    Public WithEvents mvarExcel実面積 As ToolStripMenuItem
    Public WithEvents mvarSave送付先 As ToolStripButton
    Private DSET As New DataSet
    Private p農家 As New DataTable("TBL農地")
    Private p合計 As New DataTable("TBL合計")
    Private p明細 As New DataTable("TBL明細")
    Private sKey As String = ""
    Private sWhere As String = ""
    Public Init As Boolean = True

    Public Sub New(ByVal pKey As String, ByVal sTitle As String, ByVal sLoadWhere As String)
        MyBase.New(False, False, pKey, IIf(sTitle = "", pKey, sTitle))

        mvarSP = New SplitContainer
        mvarSP.Dock = DockStyle.Fill
        sKey = pKey
        sWhere = sLoadWhere

        Me.ControlPanel.Add(mvarTabControl)
        mvarTabControl.Dock = DockStyle.Fill
        mvarTabControl.Controls.Add(mvarTPage印刷)
        mvarTabControl.Controls.Add(mvarTPage明細)

        mvarTPage印刷.Controls.Add(mvarTSContainer)
        mvarTSContainer.Dock = DockStyle.Fill
        mvarTSContainer.TopToolStripPanel.Controls.Add(mvarTStrip)
        mvarTSContainer.ContentPanel.Controls.Add(mvarSP)

        mvarTPage明細.Controls.Add(mvarDTSContainer)
        mvarDTSContainer.Dock = DockStyle.Fill
        mvarDTSContainer.TopToolStripPanel.Controls.Add(mvarDTStrip)
        mvarDTSContainer.ContentPanel.Controls.Add(mvarDGrid)

        mvarTPage印刷.Text = "利用意向出力"
        mvarTPage明細.Text = "各筆明細"

        mvarSP.Panel1.Controls.Add(mvarTreeTabControl)
        mvarTreeTabControl.Dock = DockStyle.Fill
        mvarTreeTabControl.Controls.Add(mvarTPage行政区)
        mvarTreeTabControl.Controls.Add(mvarTPage住民区分)

        mvarTree = New TreeView
        mvarTPage行政区.Controls.Add(mvarTree)
        mvarTree.CheckBoxes = True
        mvarTree.Dock = DockStyle.Fill
        mvarTPage行政区.Text = "行政区"

        mvarTree区分 = New TreeView
        mvarTPage住民区分.Controls.Add(mvarTree区分)
        mvarTree区分.CheckBoxes = True
        mvarTree区分.Dock = DockStyle.Fill
        mvarTPage住民区分.Text = "住民区分"

        mvarG = New HimTools2012.controls.DataGridViewWithDataView
        mvarG.Dock = DockStyle.Fill
        mvarSP.Panel2.Controls.Add(mvarG)
        mvarG.AllowUserToAddRows = False

        mvarSelectAll = New ToolStripButton("全選択")
        mvarSelectClear = New ToolStripButton("全解除")
        mvarSelectRev = New ToolStripButton("選択反転")

        mvarPrint登記面積 = New ToolStripSplitButton("印刷開始(登記簿面積)")
        mvarExcel登記面積 = New ToolStripMenuItem("Excel出力")
        mvarPrint実面積 = New ToolStripSplitButton("印刷開始(実面積)")
        mvarExcel実面積 = New ToolStripMenuItem("Excel出力")

        Select Case sKey
            Case "その他住民"
                mvarSave送付先 = New ToolStripButton("入力内容を保存")
                mvarTStrip.Items.AddRange({mvarSelectAll, mvarSelectClear, mvarSelectRev, New ToolStripSeparator, mvarPrint登記面積, mvarPrint実面積, New ToolStripSeparator, mvarSave送付先})
            Case Else
                mvarTStrip.Items.AddRange({mvarSelectAll, mvarSelectClear, mvarSelectRev, New ToolStripSeparator, mvarPrint登記面積, mvarPrint実面積})
        End Select

        mvarPrint登記面積.DropDownItems.Add(mvarExcel登記面積)
        mvarPrint実面積.DropDownItems.Add(mvarExcel実面積)

        With CType(mvarTStrip.Items.Add("エクセルへ出力"), ToolStripButton)
            AddHandler .Click, AddressOf ToExcel
        End With

        mvarTStrip.Items.Add(New ToolStripSeparator)

        With CType(mvarDTStrip.Items.Add("エクセルへ出力"), ToolStripButton)
            AddHandler .Click, AddressOf ToDExcel
        End With


        mvarDGrid.AllowUserToAddRows = False
    End Sub

    Public Sub LoadData()
        Init = False
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:個人Info].ID, [D:個人Info].行政区ID, V_行政区.行政区, [D:個人Info].氏名, [D:個人Info].[フリガナ], [D:個人Info].住所, [D:個人Info].郵便番号, V_農地.ID AS 筆ID, V_農地.土地所在, V_地目.名称 AS 登記地目名, V_農地.登記簿面積, IIf([実面積]=0 Or IsNull([実面積]),[登記簿面積],[実面積]) AS 現況面積, V_住民区分.名称 AS 住民区分名, V_農地.農振法区分 " & _
                                                        "FROM (V_農地 LEFT JOIN V_地目 ON V_農地.登記簿地目 = V_地目.ID) LEFT JOIN (([D:個人Info] LEFT JOIN V_行政区 ON [D:個人Info].行政区ID = V_行政区.ID) LEFT JOIN V_住民区分 ON [D:個人Info].住民区分 = V_住民区分.ID) ON V_農地.所有者ID = [D:個人Info].ID " & _
                                                        "WHERE ({0} AND ((V_農地.利用状況調査荒廃)=1) And ((V_農地.利用意向調査日) Is Null)) " & _
                                                        "ORDER BY [D:個人Info].行政区ID, [D:個人Info].[フリガナ];", sWhere)
        DSET.Tables.Add(pTBL)

        With p農家
            .Columns.Add("印刷", GetType(Boolean))
            .Columns.Add("発送番号", GetType(Integer))
            .Columns.Add("行政区ID", GetType(Integer))
            .Columns.Add("行政区", GetType(String))
            .Columns.Add("所有者ID", GetType(Decimal))
            .Columns.Add("所有者氏名", GetType(String))
            .Columns.Add("所有者フリガナ", GetType(String))
            .Columns.Add("所有者住所", GetType(String))
            .Columns.Add("所有者郵便番号", GetType(String))
            .Columns.Add("住民区分", GetType(String))
            Select Case sKey
                Case "その他住民"
                    .Columns.Add("送付先氏名", GetType(String))
                    .Columns.Add("送付先住所", GetType(String))
                    .Columns.Add("送付先郵便番号", GetType(String))
            End Select
            DSET.Tables.Add(p農家)
            DSET.Relations.Add("所有地", .Columns("所有者ID"), pTBL.Columns("ID"), False)
            .Columns.Add("筆数", GetType(Integer), "Count(Child(所有地).筆ID)")
            .Columns.Add("面積", GetType(Integer), "Sum(Child(所有地).登記簿面積)")
            .Columns.Add("実面積", GetType(Integer), "Sum(Child(所有地).現況面積)")
            .Columns.Add("集計フラグ", GetType(Integer))
            .PrimaryKey = New DataColumn() {.Columns("所有者ID")}
        End With


        Setp農家TBL(pTBL, sKey)

        With p合計
            .Columns.Add("ID", GetType(Integer))
            .Columns.Add("印刷", GetType(Boolean))
            DSET.Tables.Add(p合計)
            DSET.Relations.Add("合計", .Columns("ID"), p農家.Columns("集計フラグ"), False)
            .Columns.Add("農家数", GetType(Integer), "Count(Child(合計).所有者ID)")
            .Columns.Add("筆数計", GetType(Integer), "Sum(Child(合計).筆数)")
            .Columns.Add("面積計", GetType(Integer), "Sum(Child(合計).面積)")
            .Columns.Add("実面積計", GetType(Integer), "Sum(Child(合計).実面積)")
            .Rows.Add("1")
        End With

        With p明細
            .Columns.Add("行政区ID", GetType(Integer))
            .Columns.Add("行政区", GetType(String))
            .Columns.Add("農家ID", GetType(Integer))
            .Columns.Add("氏名", GetType(String))
            .Columns.Add("フリガナ", GetType(String))
            .Columns.Add("住所", GetType(String))
            .Columns.Add("郵便番号", GetType(String))
            .Columns.Add("筆ID", GetType(Integer))
            .Columns.Add("土地所在", GetType(String))
            .Columns.Add("登記地目", GetType(String))
            .Columns.Add("登記簿面積", GetType(Integer))
            .Columns.Add("実面積", GetType(Integer))
            .Columns.Add("農振法区分", GetType(String))
        End With

        For Each pRow As DataRow In pTBL.Rows
            Dim pMRow As DataRow = p明細.NewRow

            pMRow.Item("行政区ID") = Val(pRow.Item("行政区ID").ToString)
            pMRow.Item("行政区") = pRow.Item("行政区").ToString
            pMRow.Item("農家ID") = Val(pRow.Item("ID").ToString)
            pMRow.Item("氏名") = pRow.Item("氏名").ToString
            pMRow.Item("フリガナ") = pRow.Item("フリガナ").ToString
            pMRow.Item("住所") = pRow.Item("住所").ToString
            pMRow.Item("郵便番号") = pRow.Item("郵便番号").ToString
            pMRow.Item("筆ID") = Val(pRow.Item("筆ID").ToString)
            pMRow.Item("土地所在") = pRow.Item("土地所在").ToString
            pMRow.Item("登記地目") = pRow.Item("登記地目名").ToString
            pMRow.Item("登記簿面積") = Val(pRow.Item("登記簿面積").ToString)
            pMRow.Item("実面積") = Val(pRow.Item("現況面積").ToString)
            pMRow.Item("農振法区分") = IIf(Val(pRow.Item("農振法区分").ToString) = 1, "農振農用地", IIf(Val(pRow.Item("農振法区分").ToString) = 2, "農振地域", IIf(Val(pRow.Item("農振法区分").ToString) = 3, "農振地域外", "-")))

            p明細.Rows.Add(pMRow)
        Next

        If System.IO.File.Exists(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & String.Format("\利用意向調査出力履歴({0}).xml", sKey)) Then
            If MsgBox("前回の出力履歴を表示しますか？", vbOKCancel) = vbOK Then
                Dim reader As IO.StreamReader = New IO.StreamReader(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & String.Format("\利用意向調査出力履歴({0}).xml", sKey), System.Text.Encoding.GetEncoding("Shift_Jis"))
                XMLCheck(reader, sKey, pTBL)
            Else
                mvarG.SetDataView(p農家, "", "[行政区ID],[所有者フリガナ]")
            End If
        Else
            mvarG.SetDataView(p農家, "", "[行政区ID],[所有者フリガナ]")
        End If


        mvarG.Columns("集計フラグ").Visible = False

        mvarTStrip.Items.Add("農家合計=" & p合計.Rows(0).Item("農家数"))
        mvarTStrip.Items.Add("筆数合計=" & p合計.Rows(0).Item("筆数計"))
        mvarTStrip.Items.Add("面積合計=" & p合計.Rows(0).Item("面積計"))
        mvarDGrid.SetDataView(p明細, "", "[行政区ID],[フリガナ]")
    End Sub


    Private Sub XMLCheck(ByRef reader As IO.StreamReader, ByVal sKey As String, ByRef pTBL As DataTable)
        Try
            p農家.Clear()
            p農家.ReadXml(reader)
            mvarG.SetDataView(p農家, "", "[行政区ID],[所有者フリガナ]")
            reader.Close()
        Catch ex As Exception
            MsgBox("履歴XMLのデータがない、あるいは破損しているためデータベースから読み込みます。")
            reader.Close()
            If System.IO.File.Exists(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & String.Format("\利用意向調査出力履歴({0})BK.xml", sKey)) Then
                My.Computer.FileSystem.DeleteFile(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & String.Format("\利用意向調査出力履歴({0})BK.xml", sKey))
            End If
            My.Computer.FileSystem.RenameFile(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & String.Format("\利用意向調査出力履歴({0}).xml", sKey), String.Format("利用意向調査出力履歴({0})BK.xml", sKey))
            Setp農家TBL(pTBL, sKey)
            mvarG.SetDataView(p農家, "", "[行政区ID],[所有者フリガナ]")
        End Try
    End Sub

    Private pID As Integer = 0
    Private Sub Setp農家TBL(ByRef pTBL As DataTable, ByVal sKey As String)
        Try
            For Each pRow As DataRow In pTBL.Rows
                If Val(pRow.Item("ID").ToString) <> 0 Then
                    pID = Val(pRow.Item("ID").ToString)

                    Dim pNRow As DataRow = p農家.Rows.Find(pRow.Item("ID"))
                    If pNRow Is Nothing Then
                        pNRow = p農家.NewRow
                        pNRow.Item("行政区ID") = Val(pRow.Item("行政区ID").ToString)
                        pNRow.Item("行政区") = pRow.Item("行政区").ToString
                        pNRow.Item("所有者ID") = Val(pRow.Item("ID").ToString)
                        pNRow.Item("所有者氏名") = pRow.Item("氏名").ToString
                        pNRow.Item("所有者フリガナ") = pRow.Item("フリガナ").ToString
                        pNRow.Item("所有者住所") = pRow.Item("住所").ToString
                        pNRow.Item("所有者郵便番号") = pRow.Item("郵便番号").ToString
                        pNRow.Item("住民区分") = pRow.Item("住民区分名").ToString
                        pNRow.Item("集計フラグ") = 1

                        p農家.Rows.Add(pNRow)
                    End If
                    If mvarTree.Nodes.Find(pNRow.Item("行政区ID"), True).Length = 0 Then
                        With mvarTree.Nodes.Add(pNRow.Item("行政区").ToString)
                            .Name = pNRow.Item("行政区ID")
                        End With
                    End If

                    If mvarTree区分.Nodes.Find(pNRow.Item("住民区分"), True).Length = 0 Then
                        With mvarTree区分.Nodes.Add(pNRow.Item("住民区分").ToString)
                            .Name = pNRow.Item("住民区分")
                        End With
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message & pID & "のIDが不正です。")
        End Try

    End Sub

    Private Sub ToExcel()
        mvarG.ToExcel()
    End Sub
    Private Sub ToDExcel()
        mvarDGrid.ToExcel()
    End Sub

    Private Sub mvarSelectAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarSelectAll.Click
        For Each pNode As TreeNode In mvarTree.Nodes
            pNode.Checked = True
        Next

        For Each pNode区分 As TreeNode In mvarTree区分.Nodes
            pNode区分.Checked = True
        Next
    End Sub
    Private Sub mvarSelectClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarSelectClear.Click
        For Each pNode As TreeNode In mvarTree.Nodes
            pNode.Checked = False
        Next

        For Each pNode区分 As TreeNode In mvarTree区分.Nodes
            pNode区分.Checked = False
        Next
    End Sub
    Private Sub mvarTree_AfterCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles mvarTree.AfterCheck
        Dim nID As Integer = Val(e.Node.Name)
        Dim bCheck As Boolean = e.Node.Checked
        Dim pView As New DataView(p農家, "[行政区ID]=" & nID, "", DataViewRowState.CurrentRows)
        For Each pRowV As DataRowView In pView
            pRowV.Item("印刷") = bCheck
        Next
    End Sub
    Private Sub mvarTree区分_AfterCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles mvarTree区分.AfterCheck
        Dim n名称 As String = e.Node.Name
        Dim bCheck As Boolean = e.Node.Checked
        Dim pView As New DataView(p農家, "[住民区分]='" & n名称 & "'", "", DataViewRowState.CurrentRows)
        For Each pRowV As DataRowView In pView
            pRowV.Item("印刷") = bCheck
        Next
    End Sub
    Private Sub mvarSelectRev_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarSelectRev.Click
        For Each pNode As TreeNode In mvarTree.Nodes
            pNode.Checked = Not pNode.Checked
        Next

        For Each pNode区分 As TreeNode In mvarTree区分.Nodes
            pNode区分.Checked = Not pNode区分.Checked
        Next
    End Sub


    Private Sub mvarPrint登記面積_ButtonClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarPrint登記面積.ButtonClick
        mvarPrintStart("登記面積")
    End Sub
    Private Sub mvarExcel登記面積_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarExcel登記面積.Click
        mvarPrintStart("登記面積", True)
    End Sub

    Private Sub mvarPrint実面積_ButtonClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarPrint実面積.ButtonClick
        mvarPrintStart("実面積")
    End Sub
    Private Sub mvarExcel実面積_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarExcel実面積.Click
        mvarPrintStart("実面積", True)
    End Sub

    Private Sub mvarPrintStart(ByVal 面積分類 As String, Optional ByVal pExcel出力 As Boolean = False)
        If IO.File.Exists(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\農地における利用の意向について.xml") Then
            p農家.AcceptChanges()
            Dim pView As New DataView(p農家, "[印刷]=True", "", DataViewRowState.CurrentRows)
            Dim sPath As String = SysAD.OutputFolder & "\農地における利用の意向について.xml"
            Dim MaxValue As Integer = IIf(IsDBNull(p農家.Compute("Max(発送番号)", "")), 0, Val(p農家.Compute("Max(発送番号)", "").ToString))

            Dim bDate As Boolean = False
            If MsgBox("利用意向調査日に本日の日付を自動入力しますか？", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                bDate = True
            End If

            Using pExcel As New HimTools2012.Excel.Automation.ExcelAutomation
                For Each pRow As DataRowView In pView
                    Dim sT As String = HimTools2012.TextAdapter.LoadTextFile(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\農地における利用の意向について.xml")

                    Select Case Me.Name
                        Case "その他住民"
                            If pRow.Item("所有者氏名").ToString <> pRow.Item("送付先氏名").ToString And pRow.Item("送付先氏名").ToString <> "" Then
                                If IO.File.Exists(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\農地における利用の意向について(送付先).xml") Then
                                    sT = HimTools2012.TextAdapter.LoadTextFile(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\農地における利用の意向について(送付先).xml")
                                End If
                            End If
                    End Select

                    Dim sOutPut As String = sT

                    sOutPut = Replace(sOutPut, "{自治会番号}", pRow.Item("行政区ID").ToString)
                    sOutPut = Replace(sOutPut, "{自治会名}", pRow.Item("行政区").ToString)
                    sOutPut = Replace(sOutPut, "{郵便番号}", pRow.Item("所有者郵便番号").ToString)
                    sOutPut = Replace(sOutPut, "{氏名}", pRow.Item("所有者氏名").ToString)
                    sOutPut = Replace(sOutPut, "{住所}", pRow.Item("所有者住所").ToString)

                    Select Case Me.Name
                        Case "その他住民"
                            sOutPut = Replace(sOutPut, "{送付先郵便番号}", IIf(pRow.Item("送付先郵便番号").ToString = "", pRow.Item("所有者郵便番号").ToString, pRow.Item("送付先郵便番号").ToString))
                            sOutPut = Replace(sOutPut, "{送付先氏名}", IIf(pRow.Item("送付先氏名").ToString = "", pRow.Item("所有者氏名").ToString, pRow.Item("送付先氏名").ToString))
                            sOutPut = Replace(sOutPut, "{送付先}", IIf(pRow.Item("送付先住所").ToString = "", pRow.Item("所有者住所").ToString, pRow.Item("送付先住所").ToString))
                        Case Else
                            sOutPut = Replace(sOutPut, "{送付先郵便番号}", pRow.Item("所有者郵便番号").ToString)
                            sOutPut = Replace(sOutPut, "{送付先氏名}", pRow.Item("所有者氏名").ToString)
                            sOutPut = Replace(sOutPut, "{送付先}", pRow.Item("所有者住所").ToString)
                    End Select

                    sOutPut = Replace(sOutPut, "{発行年月日}", 和暦Format(Now))
                    sOutPut = Replace(sOutPut, "{会長名}", SysAD.DB(sLRDB).DBProperty("会長名").ToString)

                    Dim pChildView As DataView = pRow.CreateChildView("所有地")
                    Dim list = New List(Of Decimal)
                    For n As Integer = 0 To 11
                        If pChildView.Count > n Then
                            Dim pRowV As DataRowView = pChildView(n)
                            sOutPut = Replace(sOutPut, "{" & String.Format("土地の所在{0:D2}", (n + 1)) & "}", pRowV.Item("土地所在").ToString)
                            sOutPut = Replace(sOutPut, "{" & String.Format("地目{0:D2}", (n + 1)) & "}", pRowV.Item("登記地目名").ToString)
                            sOutPut = Replace(sOutPut, "{" & String.Format("農振法区分{0:D2}", (n + 1)) & "}", IIf(Val(pRowV.Item("農振法区分").ToString) = 1, "農振農用地", IIf(Val(pRowV.Item("農振法区分").ToString) = 2, "農振地域", IIf(Val(pRowV.Item("農振法区分").ToString) = 3, "農振地域外", "-"))))
                            sOutPut = Replace(sOutPut, "{" & String.Format("解答欄{0:D2}", (n + 1)) & "}", "  1 ・ 2( ｱ ･ ｲ ) ・  3 ・  4 ・  5(      　　　　　　　　　　　　　  　　　　　 　　　　　　　　   　　　　 　　   　　      )")

                            Select Case 面積分類
                                Case "登記面積" : sOutPut = Replace(sOutPut, "{" & String.Format("面積{0:D2}", (n + 1)) & "}", DecConv(Val(pRowV.Item("登記簿面積").ToString)))
                                Case "実面積" : sOutPut = Replace(sOutPut, "{" & String.Format("面積{0:D2}", (n + 1)) & "}", DecConv(Val(pRowV.Item("現況面積").ToString)))
                            End Select

                            sOutPut = Replace(sOutPut, "{" & String.Format("所有者{0:D2}", (n + 1)) & "}", pRowV.Item("氏名").ToString)

                            list.Add(Val(pRowV.Item("筆ID").ToString))
                        Else
                            sOutPut = Replace(sOutPut, "{" & String.Format("土地の所在{0:D2}", (n + 1)) & "}", "")
                            sOutPut = Replace(sOutPut, "{" & String.Format("地目{0:D2}", (n + 1)) & "}", "")
                            sOutPut = Replace(sOutPut, "{" & String.Format("農振法区分{0:D2}", (n + 1)) & "}", "")
                            sOutPut = Replace(sOutPut, "{" & String.Format("面積{0:D2}", (n + 1)) & "}", "")
                            sOutPut = Replace(sOutPut, "{" & String.Format("解答欄{0:D2}", (n + 1)) & "}", "")
                            sOutPut = Replace(sOutPut, "{" & String.Format("所有者{0:D2}", (n + 1)) & "}", "")
                        End If
                    Next

                    If bDate = True Then
                        Dim IDList As String = String.Join(",", list)
                        SysAD.DB(sLRDB).ExecuteSQL(String.Format("UPDATE [D:農地Info] SET [D:農地Info].利用意向調査日 = #{1}# WHERE ((([D:農地Info].ID) IN({0})));", IDList, Now.ToShortDateString))
                    End If

                    If IsDBNull(pRow.Item("発送番号")) Then
                        pRow.Item("発送番号") = MaxValue + 1
                        MaxValue += 1
                    Else
                    End If
                    sOutPut = Replace(sOutPut, "{発送番号}", Val(pRow.Item("発送番号").ToString))

                    Select Case pExcel出力
                        Case True
                            'pRow.Item("印刷") = False
                            Dim St送付先 As String = ""
                            Select Case Me.Name
                                Case "その他住民" : St送付先 = IIf(pRow.Item("送付先氏名").ToString = "", pRow.Item("所有者氏名").ToString, "【送】" & pRow.Item("送付先氏名").ToString & "【所】" & pRow.Item("所有者氏名").ToString)
                                Case Else : St送付先 = pRow.Item("所有者氏名").ToString
                            End Select
                            sPath = SysAD.OutputFolder & String.Format("\{0}農地における利用の意向について({1}).xml", 和暦Format(Now, "yyyyMMdd"), St送付先)
                            HimTools2012.TextAdapter.SaveTextFile(sPath, sOutPut)

                        Case False
                            HimTools2012.TextAdapter.SaveTextFile(sPath, sOutPut)
                            pExcel.PrintBook(sPath)
                    End Select
                Next
            End Using

            MsgBox("出力が終了しました。")

            Select Case pExcel出力
                Case True : SysAD.ShowFolder(System.IO.Directory.GetParent(sPath).ToString)
                Case False
            End Select

            'DataTableのデータをXMLに書き込む
            Dim SW As System.IO.StreamWriter = New System.IO.StreamWriter(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & String.Format("\利用意向調査出力履歴({0}).xml", Me.Name), False, System.Text.Encoding.GetEncoding("Shift_Jis"))
            Try

                p農家.WriteXml(SW)
            Finally
                If Not (SW Is Nothing) Then
                    SW.Close()
                End If
            End Try
        Else
            MsgBox("指定されたフォルダにＸＭＬファイルがありません")
        End If
    End Sub

    Private Sub mvarSave送付先_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarSave送付先.Click
        'DataTableのデータをXMLに書き込む
        Dim SW As System.IO.StreamWriter = New System.IO.StreamWriter(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & String.Format("\利用意向調査出力履歴({0}).xml", Me.Name), False, System.Text.Encoding.GetEncoding("Shift_Jis"))
        Try

            p農家.WriteXml(SW)
        Finally
            If Not (SW Is Nothing) Then
                SW.Close()
            End If
        End Try

        MsgBox("保存しました。")
    End Sub

    Private Function DecConv(ByVal pDec As Decimal) As String
        If Fix(pDec) = pDec Then
            Return Val(pDec).ToString("#,##0")
        Else
            Return Val(pDec).ToString("#,##0.##") '小数点第2位まで表示
        End If

        Return pDec
    End Function
End Class

Public Class CPage利用意向調査H26南種子
    Inherits HimTools2012.controls.CTabPageWithToolStrip
    Private WithEvents mvarTree As TreeView
    Private mvarSP As SplitContainer
    Private mvarG As HimTools2012.controls.DataGridViewWithDataView
    Private WithEvents mvarSelectAll As ToolStripButton
    Private WithEvents mvarSelectClear As ToolStripButton
    Private WithEvents mvarSelectRev As ToolStripButton
    Private WithEvents mvarPrintStart As ToolStripButton
    Private DSET As New DataSet
    Private p農家 As New DataTable
    Private p合計 As New DataTable


    Public Sub New()
        MyBase.New(True, True, "利用意向調査出力", "利用意向調査出力")

        mvarSP = New SplitContainer
        mvarSP.Dock = DockStyle.Fill

        Me.ControlPanel.Add(mvarSP)
        mvarTree = New TreeView
        mvarTree.CheckBoxes = True
        mvarTree.Dock = DockStyle.Fill
        mvarSP.Panel1.Controls.Add(mvarTree)
        mvarG = New HimTools2012.controls.DataGridViewWithDataView
        mvarG.Dock = DockStyle.Fill
        mvarSP.Panel2.Controls.Add(mvarG)
        mvarG.AllowUserToAddRows = False

        mvarSelectAll = New ToolStripButton("全選択")
        mvarSelectClear = New ToolStripButton("全解除")
        mvarSelectRev = New ToolStripButton("選択反転")

        mvarPrintStart = New ToolStripButton("印刷開始")

        Me.ToolStrip.Items.AddRange({mvarSelectAll, mvarSelectClear, mvarSelectRev, mvarPrintStart})

        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:個人Info].ID,[D:個人Info].行政区ID, V_行政区.行政区, [D:個人Info].氏名, [D:個人Info].[フリガナ], [D:個人Info].住所, [D:個人Info].郵便番号,[D:個人Info].住民区分, V_農地.ID AS 筆ID, V_農地.土地所在, V_地目.名称 AS 登記地目名, V_農地.登記簿面積 FROM (V_農地 INNER JOIN (([D:世帯Info] INNER JOIN [D:個人Info] ON [D:世帯Info].世帯主ID = [D:個人Info].ID) INNER JOIN V_行政区 ON [D:個人Info].行政区ID = V_行政区.ID) ON V_農地.所有世帯ID = [D:世帯Info].ID) INNER JOIN V_地目 ON V_農地.登記簿地目 = V_地目.ID WHERE ((V_農地.地番) Is Not Null) ORDER BY [D:個人Info].行政区ID, [D:個人Info].[フリガナ], [V_農地].大字ID, [V_農地].ソート文字")
        Dim pView As New DataView(pTBL, "", "", DataViewRowState.CurrentRows)



        DSET.Tables.Add(pTBL)
        With p農家
            .Columns.Add("印刷", GetType(Boolean))
            .Columns.Add("ID", GetType(Decimal))
            .Columns.Add("氏名", GetType(String))
            .Columns.Add("フリガナ", GetType(String))
            .Columns.Add("住所", GetType(String))
            .Columns.Add("郵便番号", GetType(String))
            .Columns.Add("行政区", GetType(String))
            .Columns.Add("行政区ID", GetType(Integer))
            DSET.Tables.Add(p農家)
            DSET.Relations.Add("所有地", .Columns("ID"), pTBL.Columns("ID"), False)
            .Columns.Add("筆数", GetType(Integer), "Count(Child(所有地).筆ID)")
            .Columns.Add("面積", GetType(Integer), "Sum(Child(所有地).登記簿面積)")
            .Columns.Add("集計フラグ", GetType(Integer))
            .PrimaryKey = New DataColumn() {.Columns("ID")}
        End With

        With p合計
            .Columns.Add("ID", GetType(Integer))
            .Columns.Add("印刷", GetType(Boolean))
            DSET.Tables.Add(p合計)
            DSET.Relations.Add("合計", .Columns("ID"), p農家.Columns("集計フラグ"), False)
            .Columns.Add("筆数計", GetType(Integer), "Sum(Child(合計).筆数)")
            .Columns.Add("面積計", GetType(Integer), "Sum(Child(合計).面積)")
            .Rows.Add("1")
        End With

        For Each pRow As DataRow In pTBL.Rows
            Dim pNRow As DataRow = p農家.Rows.Find(pRow.Item("ID"))
            If pNRow Is Nothing Then
                pNRow = p農家.NewRow
                pNRow.Item("ID") = pRow.Item("ID")
                pNRow.Item("氏名") = pRow.Item("氏名")
                pNRow.Item("フリガナ") = pRow.Item("フリガナ")

                pNRow.Item("住所") = pRow.Item("住所")
                pNRow.Item("郵便番号") = pRow.Item("郵便番号")
                Select Case Val(pRow.Item("住民区分").ToString)
                    Case 0
                        pNRow.Item("行政区") = pRow.Item("行政区")
                        pNRow.Item("行政区ID") = pRow.Item("行政区ID")
                    Case 1, 3
                        pNRow.Item("行政区") = "町外住民"
                        pNRow.Item("行政区ID") = 9999001
                    Case 2
                        pNRow.Item("行政区") = "死亡者"
                        pNRow.Item("行政区ID") = 9999002
                    Case Else
                        pNRow.Item("行政区") = "その他"
                        pNRow.Item("行政区ID") = 9999003
                End Select
                pNRow.Item("集計フラグ") = 1

                p農家.Rows.Add(pNRow)
            End If
            If mvarTree.Nodes.Find(pNRow.Item("行政区ID"), True).Length = 0 Then
                With mvarTree.Nodes.Add(pNRow.Item("行政区"))
                    .name = pNRow.Item("行政区ID")
                End With
            End If
        Next

        mvarG.SetDataView(p農家, "", "[行政区ID],[フリガナ]")
        mvarG.Columns("集計フラグ").Visible = False

        Me.ToolStrip.Items.Add("筆数合計=" & p合計.Rows(0).Item("筆数計"))
        Me.ToolStrip.Items.Add("面積合計=" & p合計.Rows(0).Item("面積計"))

        With CType(Me.ToolStrip.Items.Add("エクセルへ出力"), ToolStripButton)
            AddHandler .Click, AddressOf ToExcel
        End With
    End Sub

    Private Sub ToExcel()
        mvarG.ToExcel()
    End Sub

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.NoMessage
        End Get
    End Property



    Private Sub mvarSelectAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarSelectAll.Click
        For Each pNode As TreeNode In mvarTree.Nodes
            pNode.Checked = True
        Next
    End Sub

    Private Sub mvarTree_AfterCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles mvarTree.AfterCheck
        Dim nID As Integer = Val(e.Node.Name)
        Dim bCheck As Boolean = e.Node.Checked
        Dim pView As New DataView(p農家, "[行政区ID]=" & nID, "", DataViewRowState.CurrentRows)
        For Each pRowV As DataRowView In pView
            pRowV.Item("印刷") = bCheck
        Next
    End Sub

    Private Sub mvarSelectClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarSelectClear.Click
        For Each pNode As TreeNode In mvarTree.Nodes
            pNode.Checked = False
        Next
    End Sub

    Private Sub mvarSelectRev_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarSelectRev.Click
        For Each pNode As TreeNode In mvarTree.Nodes
            pNode.Checked = Not pNode.Checked
        Next
    End Sub

    Private Sub mvarPrintStart_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarPrintStart.Click
        If IO.File.Exists(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\農地における利用の意向について.xml") Then
            Dim sT As String = HimTools2012.TextAdapter.LoadTextFile(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\農地における利用の意向について.xml")
            Dim pView As New DataView(p農家, "[印刷]=True", "", DataViewRowState.CurrentRows)
            Dim sPath As String = SysAD.OutputFolder & "\農地における利用の意向について.xml"
            Using pExcel As New HimTools2012.Excel.Automation.ExcelAutomation

                For Each pRow As DataRowView In pView
                    Dim sOutPut As String = sT

                    Dim pChildView As DataView = pRow.CreateChildView("所有地")

                    Select Case pChildView.Count
                        Case 0 To 3 : sOutPut = Replace(sOutPut, fncDelPage(sOutPut, " <Worksheet ss:Name=""Page002"">", " </Worksheet>" & vbCrLf), "")
                        Case 4 To 20 : sOutPut = Replace(sOutPut, fncDelPage(sOutPut, " <Worksheet ss:Name=""Page003"">", " </Worksheet>" & vbCrLf), "")
                        Case 21 To 37 : sOutPut = Replace(sOutPut, fncDelPage(sOutPut, " <Worksheet ss:Name=""Page004"">", " </Worksheet>" & vbCrLf), "")
                        Case 38 To 54 : sOutPut = Replace(sOutPut, fncDelPage(sOutPut, " <Worksheet ss:Name=""Page005"">", " </Worksheet>" & vbCrLf), "")
                        Case 55 To 71 : sOutPut = Replace(sOutPut, fncDelPage(sOutPut, " <Worksheet ss:Name=""Page006"">", " </Worksheet>" & vbCrLf), "")
                        Case 72 To 88 : sOutPut = Replace(sOutPut, fncDelPage(sOutPut, " <Worksheet ss:Name=""Page007"">", " </Worksheet>" & vbCrLf), "")
                        Case 98 To 105 : sOutPut = Replace(sOutPut, fncDelPage(sOutPut, " <Worksheet ss:Name=""Page008"">", " </Worksheet>" & vbCrLf), "")
                        Case 106 To 122
                        Case Else
                            Stop
                    End Select


                    sOutPut = Replace(sOutPut, "{郵便番号}", pRow.Item("郵便番号").ToString)
                    sOutPut = Replace(sOutPut, "{氏名}", pRow.Item("氏名").ToString)
                    sOutPut = Replace(sOutPut, "{所有者名}", pRow.Item("氏名").ToString)
                    sOutPut = Replace(sOutPut, "{住所}", pRow.Item("住所").ToString)
                    sOutPut = Replace(sOutPut, "{発行年月日}", 和暦Format(Now))
                    sOutPut = Replace(sOutPut, "{自治会名}", pRow.Item("行政区").ToString)
                    sOutPut = Replace(sOutPut, "{自治会番号}", pRow.Item("行政区ID").ToString)
                    sOutPut = Replace(sOutPut, "{会長名}", SysAD.DB(sLRDB).DBProperty("会長名").ToString)


                    For n As Integer = 0 To 180
                        If pChildView.Count > n Then
                            Dim pRowV As DataRowView = pChildView(n)
                            sOutPut = Replace(sOutPut, "{" & String.Format("土地所在{0:D2}", (n + 1)) & "}", pRowV.Item("土地所在").ToString)
                            sOutPut = Replace(sOutPut, "{" & String.Format("地目{0:D2}", (n + 1)) & "}", pRowV.Item("登記地目名").ToString)
                            sOutPut = Replace(sOutPut, "{" & String.Format("面積{0:D2}", (n + 1)) & "}", pRowV.Item("登記簿面積").ToString)
                        Else
                            sOutPut = Replace(sOutPut, "{" & String.Format("土地所在{0:D2}", (n + 1)) & "}", "")
                            sOutPut = Replace(sOutPut, "{" & String.Format("地目{0:D2}", (n + 1)) & "}", "")
                            sOutPut = Replace(sOutPut, "{" & String.Format("面積{0:D2}", (n + 1)) & "}", "")
                        End If
                    Next

                    HimTools2012.TextAdapter.SaveTextFile(sPath, sOutPut)
#If DEBUG Then
                    pExcel.PrintBook(sPath)
#Else
                                        pExcel.PrintBook(sPath)
#End If

                Next
            End Using
        Else
            MsgBox("指定されたフォルダにＸＭＬファイルがありません")

        End If

    End Sub
    Private Function fncDelPage(ByVal sText, ByVal sStart, ByVal sEnd) As String
        Dim sRet As String = HimTools2012.StringF.Mid(sText, InStr(sText, sStart))
        sRet = HimTools2012.StringF.Left(sRet, InStrRev(sRet, sEnd) - 1) & sEnd
        Return sRet
    End Function
End Class
