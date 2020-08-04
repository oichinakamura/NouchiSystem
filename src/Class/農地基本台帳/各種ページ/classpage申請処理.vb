
Imports System.ComponentModel
Public Class classpage申請処理
    Inherits HimTools2012.controls.CTabPageWithToolStrip


    Private mvarVSplitter As HimTools2012.controls.SplitContainerEx
    Private mvarSplitterMain As HimTools2012.controls.SplitContainerEx
    Private mvarSplitterSecond As HimTools2012.controls.SplitContainerEx

    Private mvarNavi As HimTools2012.controls.TabControlBase
    Private WithEvents mvarGrid As DataGridView
    Private mvarList As ViewEx

    Private mvarCenterTabPage As HimTools2012.controls.TabControlBase

    Private WithEvents mvarTreeView As TreeView
    Public DataProperty As PropertyGrid
    Public Label申請 As ToolStripLabel
    Public labelList As ToolStripLabel
    Public labelGrid As ToolStripLabel


    Public Sub New()
        MyBase.New(True)
        Me.Name = "申請処理"
        Me.Text = "申請処理"

        '    mvarVSplitter = New SplitContainerEX("申請縦スプリッタ")
        '    mvarSplitterMain = New SplitContainerEX("申請処理スプリッタ")
        '    mvarSplitterSecond = New SplitContainerEX("申請処理スプリッタ2")
        '    mvarCenterTabPage = New CTabCtrlExt
        '    mvarCenterTabPage.Dock = DockStyle.Fill
        '    mvarNavi = New CTabCtrlExt
        '    mvarNavi.Dock = DockStyle.Fill

        '    mvarGrid = New DataGridView
        '    mvarGrid.Dock = DockStyle.Fill

        '    mvarList = New ViewEx()

        '    Label申請 = New ToolStripLabel
        '    Label申請.Text = "↓表示する項目を選択してください。"
        '    Me.ToolStrip.Items.Add(Label申請)

        '    labelList = New ToolStripLabel
        '    labelList.Text = ""
        '    labelGrid = New ToolStripLabel
        '    labelGrid.Text = ""

        '    DataProperty = New PropertyGrid
        '    DataProperty.Dock = DockStyle.Fill

        '    mvarCenterTabPage.AddPage("リスト形式", "リスト形式", mvarList, False).ToolStrip.Items.Add(labelList)
        '    mvarCenterTabPage.AddPage("表形式", "表形式", mvarGrid, False).ToolStrip.Items.Add(labelGrid)

        '    mvarSplitterMain.Panel1.Controls.Add(mvarNavi)
        '    mvarSplitterMain.Panel2.Controls.Add(mvarSplitterSecond)
        '    mvarSplitterSecond.Panel1.Controls.Add(mvarCenterTabPage)
        '    mvarSplitterSecond.Panel2.Controls.Add(DataProperty)
        '    mvarVSplitter.Panel1.Controls.Add(mvarSplitterMain)

        '    mvarVSplitter.Dock = DockStyle.Fill
        '    mvarVSplitter.Orientation = Orientation.Horizontal
        '    Me.ControlPanel.Add(mvarVSplitter)

        '    mvarTreeView = New TreeView
        '    mvarNavi.AddPage("総会", "総会", mvarTreeView)

        '    With New CommonTools.CDataBase(SysAD.DatabaseProperty.Path)
        '        .Open()
        '        If Not .TableExists("D_総会資料") Then
        '            .Execute("CREATE TABLE D_総会資料([ID] LONG CONSTRAINT pkey PRIMARY KEY,[年度] LONG,[月] LONG,[受付開始日] DATE,[受付終了日] DATE,[総会実施日] DATE)")
        '        End If

        '        Dim pArrayLst As ArrayList = .GetFieldList("D_申請")
        '        If Not pArrayLst.Contains("転用目的種類") Then
        '            .Execute("ALTER TABLE D_申請 ADD [転用目的種類] LONG;")
        '        End If

        '        Dim mvarTBL総会資料 As DataTable = .GetTable("SELECT * FROM [D_総会資料] ORDER BY [ID]")

        '        For Each pRow As DataRow In mvarTBL総会資料.Rows
        '            Dim pNode As TreeNode
        '            Dim pTNode() As TreeNode = mvarTreeView.Nodes.Find("年度." & pRow.Item("年度"), True)

        '            If pTNode Is Nothing OrElse pTNode.Length = 0 Then
        '                pNode = mvarTreeView.Nodes.Add("年度." & pRow.Item("年度"), pRow.Item("年度") & "年(平成" & pRow.Item("年度") - 1988 & ")")
        '            Else
        '                pNode = pTNode(0)
        '            End If

        '            Dim pTNodeC() As TreeNode = pNode.Nodes.Find("月." & pRow.Item("月"), True)
        '            If pTNodeC.Length = 0 Then
        '                pNode.Nodes.Add("月." & pRow.Item("ID"), pRow.Item("月") & "月").Tag = pRow
        '            Else

        '            End If
        '        Next

        '        .Close()
        '    End With
        'End Sub

        'Private Sub mvarTreeView_NodeMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles mvarTreeView.NodeMouseClick

        '    With New CommonTools.CDataBase(SysAD.DatabaseProperty.Path)
        '        .Open()

        '        Select Case GetKeyHead(e.Node.Name)
        '            Case "年度"
        '                Dim pTable As DataTable = .GetTable(
        '                        String.Format("SELECT * FROM [D_総会資料] WHERE [年度]={0} ORDER BY [ID]", GetKeyCode(e.Node.Name)))
        '                labelList.Text = "総会情報" & GetKeyCode(e.Node.Name) & "年"
        '                'Dim pB As New BindingSource
        '                mvarList.Items.Clear()
        '                For Each pRowX As DataRow In pTable.Rows
        '                    'Dim p申請月 As New CObj申請月(pRowX, False)
        '                    'pB.Add(p申請月)

        '                    mvarList.Items.Add(New ViewItem総会資料(pRowX))
        '                Next
        '                mvarList.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent)

        '                mvarCenterTabPage.SelectedTab = mvarCenterTabPage.TabPages("リスト形式")
        '                mvarGrid.DataSource = Nothing
        '            Case "月"
        '                If e.Node.Tag IsNot Nothing Then
        '                    Dim pRow As DataRow = e.Node.Tag

        '                    labelGrid.Text = String.Format("総会情報 {2}年/{3}月  [{0:yyyy/MM/dd} ～ {1:yyyy/MM/dd}]", pRow.Item("受付開始日"), pRow.Item("受付終了日"), pRow.Item("年度"), pRow.Item("月"))
        '                    Dim pTable As DataTable = .GetTable(
        '                            String.Format("SELECT * FROM [D_申請] WHERE [受付年月日]>=#{0}# AND [受付年月日]<=#{1}# ORDER BY 法令,状態,受付年月日",
        '                            CType(pRow.Item("受付開始日"), DateTime).ToString("MM/dd/yyyy"),
        '                            CType(pRow.Item("受付終了日"), DateTime).ToString("MM/dd/yyyy")))

        '                    Dim pB As New BindingSource

        '                    For Each pRowX As DataRow In pTable.Rows
        '                        Dim p申請 As New CObj申請(pRowX, False)
        '                        pB.Add(p申請)
        '                    Next

        '                    mvarCenterTabPage.SelectedTab = mvarCenterTabPage.TabPages("表形式")
        '                    mvarGrid.DataSource = pB
        '                End If
        '        End Select

        '        .Close()
        '    End With

    End Sub


    Private Class ViewEx
        Inherits ListView
        Private imgList As ImageList

        Public Sub New()
            MyBase.New()

            Me.Columns.Clear()
            Me.Columns.Add("ID")
            Me.Columns.Add("月").Name = "月"
            Me.Columns.Add("受付開始日").Name = "受付開始日"
            Me.Columns.Add("受付終了日").Name = "受付終了日"
            Me.Columns.Add("総会実施日").Name = "総会実施日"

            imgList = New ImageList
            imgList.ImageSize = New Size(32, 32)
            imgList.ColorDepth = ColorDepth.Depth24Bit
            imgList.Images.Add("Folder", My.Resources.Resource1.Folder_Back)

            Me.SmallImageList = imgList
            Me.View = Windows.Forms.View.Details
        End Sub

        Private Sub ViewEx_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDoubleClick
            Dim pHitTest As ListViewHitTestInfo = Me.HitTest(e.Location)
            If e.Button = Windows.Forms.MouseButtons.Left AndAlso pHitTest.Item IsNot Nothing Then
                Dim pItem As CListViewItemEx = pHitTest.Item
                pItem.DblClick()
            End If
        End Sub

        Private Sub ViewEx_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown
            Dim pHitTest As ListViewHitTestInfo = Me.HitTest(e.Location)
            If e.Button = Windows.Forms.MouseButtons.Right AndAlso pHitTest.Item IsNot Nothing Then
                Dim pItem As CListViewItemEx = pHitTest.Item
                pItem.PopMenu(Me.PointToScreen(e.Location))
            End If
        End Sub
    End Class

    Private Class ViewItem総会資料
        Inherits CListViewItemEx

        Public Sub New(ByVal pRow As DataRow)
            MyBase.New(pRow)
            Me.ImageKey = "Folder"
            Me.Text = pRow.Item("ID")
            Me.SubItems.Add(Row.Item("月"))
            Me.SubItems.Add(Row.Item("受付開始日"))
            Me.SubItems.Add(Row.Item("受付終了日"))
            Me.SubItems.Add(Row.Item("総会実施日"))
        End Sub

        Public Overrides Sub DblClick()
            'Dim pB As New BindingSource
            'With New CommonTools.CDataBase(SysAD.DatabaseProperty.Path)
            '    .Open()

            '    Dim pTable As DataTable = .GetTable(
            '            String.Format("SELECT * FROM [D_申請] WHERE [受付年月日]>=#{0}# AND [受付年月日]<=#{1}# ORDER BY 法令,状態,受付年月日",
            '            CType(Me.Row.Item("受付開始日"), DateTime).ToString("MM/dd/yyyy"),
            '            CType(Me.Row.Item("受付終了日"), DateTime).ToString("MM/dd/yyyy")))


            '    For Each pRowX As DataRow In pTable.Rows
            '        Dim p申請 As New CObj申請(pRowX, False)
            '        pB.Add(p申請)
            '    Next
            '    pB.Sort = "法令,状態,受付年月日"
            '    .Close()
            '    SysAD.page申請処理.labelGrid.Text = String.Format("総会情報 {2}年/{3}月  [{0:yyyy/MM/dd} ～ {1:yyyy/MM/dd}]", Row.Item("受付開始日"), Row.Item("受付終了日"), Row.Item("年度"), Row.Item("月"))
            'End With
            'With SysAD.page申請処理
            '    .mvarCenterTabPage.SelectedTab = .mvarCenterTabPage.TabPages("表形式")
            '    .mvarGrid.DataSource = pB
            'End With
        End Sub

        Public Overrides Sub PopMenu(ByVal pPoint As Point)
            Dim pContext As New ContextMenuStrip

            pContext.Show(pPoint)
        End Sub

    End Class

    Private Sub mvarGrid_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles mvarGrid.DataError
        'Nop
    End Sub

    Private Sub mvarGrid_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles mvarGrid.MouseClick
        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                Dim hti As DataGridView.HitTestInfo = mvarGrid.HitTest(e.X, e.Y)
                If hti.RowIndex > -1 Then
                    Dim pB As BindingSource = mvarGrid.DataSource
                    Dim p申請 As CObj申請 = pB.Item(hti.RowIndex)
                    Select Case p申請.法令
                        'Case enum法令.農地法3条所有権 : DataProperty.SelectedObject = New CObj申請農地法3条所有権移転(p申請.Row, False)
                        'Case enum法令.農地法4条 : DataProperty.SelectedObject = New CObj申請農地法4条(p申請.Row, False)
                        'Case enum法令.農地法5条所有権 : DataProperty.SelectedObject = New CObj申請農地法5条(p申請.Row, False)
                        'Case enum法令.農地法5条貸借 : DataProperty.SelectedObject = New CObj申請農地法5条(p申請.Row, False)
                        'Case enum法令.基盤強化法所有権 : DataProperty.SelectedObject = New CObj申請農地法3条所有権移転(p申請.Row, False)
                        'Case enum法令.利用権設定 : DataProperty.SelectedObject = New CObj申請基盤強化利用権設定(p申請.Row, False)
                        'Case enum法令.あっせん申出出し手 : DataProperty.SelectedObject = New CObj申請あっせん申出出し手(p申請.Row, False)
                        'Case enum法令.あっせん申出受け手 : DataProperty.SelectedObject = New CObj申請あっせん申出受け手(p申請.Row, False)
                        Case Else
                            Stop
                            DataProperty.SelectedObject = p申請
                    End Select

                End If
            Case Windows.Forms.MouseButtons.Right
                Dim hti As DataGridView.HitTestInfo = mvarGrid.HitTest(e.X, e.Y)
                If hti.RowIndex > -1 Then
                    Dim pB As BindingSource = mvarGrid.DataSource
                    Dim p申請 As CObj申請 = pB.Item(hti.RowIndex)
                    Dim b As Boolean = False

                    Select Case p申請.法令
                        'Case enum法令.農地法3条所有権 : CType(New CObj申請農地法3条所有権移転(p申請.Row, False).PopMenu(), ContextMenuStrip).Show(mvarGrid.PointToScreen(e.Location))
                        'Case enum法令.農地法4条 : CType(New CObj申請農地法4条(p申請.Row, False).PopMenu(), ContextMenuStrip).Show(mvarGrid.PointToScreen(e.Location))
                        'Case enum法令.農地法5条所有権 : CType(New CObj申請農地法5条(p申請.Row, False).PopMenu(), ContextMenuStrip).Show(mvarGrid.PointToScreen(e.Location))
                        'Case enum法令.農地法5条貸借 : CType(New CObj申請農地法5条(p申請.Row, False).PopMenu(), ContextMenuStrip).Show(mvarGrid.PointToScreen(e.Location))
                        'Case enum法令.基盤強化法所有権 : CType(New CObj申請基盤強化法所有権(p申請.Row, False).PopMenu(), ContextMenuStrip).Show(mvarGrid.PointToScreen(e.Location))
                        'Case enum法令.利用権設定 : CType(New CObj申請基盤強化利用権設定(p申請.Row, False).PopMenu(), ContextMenuStrip).Show(mvarGrid.PointToScreen(e.Location))
                        'Case enum法令.あっせん申出出し手 : CType(New CObj申請あっせん申出出し手(p申請.Row, False).PopMenu(), ContextMenuStrip).Show(mvarGrid.PointToScreen(e.Location))
                        'Case enum法令.あっせん申出受け手 : CType(New CObj申請あっせん申出受け手(p申請.Row, False).PopMenu(), ContextMenuStrip).Show(mvarGrid.PointToScreen(e.Location))
                        Case Else

                            p申請.GetContextMenu(Nothing).Show(mvarGrid.PointToScreen(e.Location))
                    End Select
                End If
        End Select
    End Sub

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.NoMessage
        End Get
    End Property
End Class

Public Class CListViewEx2
    Inherits ListView

    Public Sub New()
        MyBase.New()

        Me.Dock = DockStyle.Fill
    End Sub
End Class
Public MustInherit Class CListViewItemEx
    Inherits ListViewItem
    Public Row As DataRow

    Public Sub New(ByVal pRow As DataRow)
        MyBase.New()
        Row = pRow
    End Sub
    Public MustOverride Sub PopMenu(ByVal pPoint As Point)
    Public MustOverride Sub DblClick()

End Class
