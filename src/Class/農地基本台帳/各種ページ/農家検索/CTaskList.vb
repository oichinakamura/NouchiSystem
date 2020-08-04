

Public Class CTaskList
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Private mvarSplitter As HimTools2012.controls.SplitContainerEx
    Private mvarSplitterV As SplitContainer
    Private mvarFlowPanel As HimTools2012.controls.FlowLayoutPanelEX
    Private WithEvents mvarCal As HimTools2012.controls.CalenderCtrl
    Private WithEvents mvarDTP As DateTimePicker
    Private WithEvents mvarTime1 As DateTimePicker

    Private WithEvents mvarDTP2 As DateTimePicker
    Private WithEvents mvarTime2 As DateTimePicker


    Private mnuGContext As HimTools2012.controls.ContextMenuEx
    Private WithEvents mvarGrid As HimTools2012.controls.DataGridViewWithDataView
    Private mvarTBL As DataTable
    Private mvarXMLFile As String = ""
    Private mvarXMLUpdate As DateTime
    Private mvarTab As HimTools2012.controls.TabControlBase
    Private watcher As System.IO.FileSystemWatcher = Nothing

    Public Sub New(ByRef pNode As Xml.XmlNode, ByRef pLayout As HimTools2012.controls.XMLLayout)
        MyBase.New(pNode, pLayout)

        Text = "タスク一覧"

        mvarXMLFile = SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\Task.xml"

        mvarTBL = New DataTable
        mvarTBL.TableName = "共有タスク"
        mvarTBL.Columns.Add("GID", GetType(Guid))
        mvarTBL.Columns.Add("日時", GetType(Date))
        mvarTBL.Columns.Add("種類", GetType(Integer))
        mvarTBL.Columns.Add("内容", GetType(String))
        mvarTBL.Columns.Add("削除", GetType(Boolean))
        mvarTBL.PrimaryKey = New DataColumn() {mvarTBL.Columns("GID")}
        Me.CanMove = False

        mvarSplitter = New HimTools2012.controls.SplitContainerEX("タスク管理SP")
        mvarSplitter.Dock = DockStyle.Fill

        mvarSplitterV = New SplitContainer
        mvarSplitterV.BackColor = Color.DarkGray
        mvarSplitterV.Dock = DockStyle.Fill
        mvarSplitterV.Orientation = Orientation.Horizontal
        mvarSplitterV.SplitterWidth = 16
        mvarSplitterV.Panel1.BackColor = Color.White

        Dim pPanel As New Panel
        pPanel.Dock = DockStyle.Fill
        pPanel.AutoScroll = True

        mvarCal = New HimTools2012.controls.CalenderCtrl
        mvarCal.Dock = DockStyle.Fill
        mvarCal.BackColor = Color.WhiteSmoke
        mvarCal.Location = New Point(0, 0)
        mvarCal.Margin = New Padding(0, 0, 0, 0)
        mvarCal.Padding = New Padding(0, 0, 0, 0)
        '        mvarCal.MaxSelectionCount = 365 * 100

        If IO.File.Exists(mvarXMLFile) Then
            Try
                Dim pTBL As New DataTable
                pTBL.ReadXml(mvarXMLFile)
                mvarTBL.Merge(pTBL, True, MissingSchemaAction.AddWithKey)
            Catch ex As Exception
                Stop
            End Try
            Dim ar日付 As New List(Of Date)
            For Each pRow As DataRowView In New DataView(mvarTBL, "[削除]=False", "", DataViewRowState.CurrentRows)
                If Not ar日付.Contains(pRow.Item("日時")) Then
                    ar日付.Add(CDate(pRow.Item("日時")).Date)
                End If
            Next
            'mvarCal.BoldedDates = ar日付.ToArray
            mvarXMLUpdate = Now
        Else
            SaveTBL()

        End If
        StartW()

        pPanel.Controls.Add(mvarCal)
        pPanel.BackColor = Color.DarkGreen
        mvarSplitterV.Panel1.Controls.Add(pPanel)

        mvarDTP = New DateTimePicker
        mvarDTP.Format = DateTimePickerFormat.Custom
        mvarDTP.Font = New Font(Me.Font.FontFamily, 12)
        mvarDTP.CustomFormat = "yyyy年 M月 d日"

        mvarTime1 = New DateTimePicker
        mvarTime1.Format = DateTimePickerFormat.Time
        mvarTime1.Font = New Font(Me.Font.FontFamily, 12)
        mvarTime1.ShowUpDown = True

        mvarDTP2 = New DateTimePicker
        mvarDTP2.Format = DateTimePickerFormat.Custom
        mvarDTP2.Font = New Font(Me.Font.FontFamily, 12)
        mvarDTP2.CustomFormat = "yyyy年 M月 d日"

        mvarTime2 = New DateTimePicker
        mvarTime2.Format = DateTimePickerFormat.Time
        mvarTime2.Font = New Font(Me.Font.FontFamily, 12)
        mvarTime2.ShowUpDown = True

        mvarFlowPanel = New HimTools2012.controls.FlowLayoutPanelEx
        mvarFlowPanel.BackColor = Color.White
        mvarSplitterV.Panel2.Controls.Add(mvarFlowPanel)
        mvarFlowPanel.Controls.Add(mvarDTP)
        mvarFlowPanel.Controls.Add(mvarTime1)
        mvarFlowPanel.Controls.Add(mvarDTP2)
        mvarFlowPanel.Controls.Add(mvarTime2)
        mvarFlowPanel.SetFlowBreak(mvarTime1, True)
        mvarFlowPanel.SetFlowBreak(mvarTime2, True)
        mvarFlowPanel.AddButton("現地調査日を追加", True, AddressOf DoClick)
        mvarFlowPanel.AddButton("総会日を追加", True, AddressOf DoClick)


        mvarGrid = New HimTools2012.controls.DataGridViewWithDataView
        mvarGrid.Dock = DockStyle.Fill
        mvarGrid.AllowUserToAddRows = False
        mvarGrid.AllowUserToDeleteRows = False
        mvarGrid.SetDataView(mvarTBL, "[削除]=False AND [日時]=" & HimTools2012.StringF.Toリテラル日付(mvarDTP.Value), "[日時]")
        mvarGrid.AutoGenerateColumns = False

        Dim pCol0 As New DataGridViewTextBoxColumn
        pCol0.Name = "GID"
        pCol0.DataPropertyName = "GID"
        pCol0.Visible = False

        Dim pCol1 As New DataGridViewDateTimePickerColumn '【保留】HimTools2012.controls.
        pCol1.Name = "日時"
        pCol1.DataPropertyName = "日時"
        mvarGrid.Columns.Add(pCol1)

        Dim pCol2 As New DataGridViewComboBoxColumn
        Dim pKind As New DataTable
        pCol2.Name = "種類"
        pCol2.DataPropertyName = "種類"
        pCol2.DisplayMember = "Display"
        pCol2.ValueMember = "Value"
        pKind.Columns.Add("Display", GetType(String))
        pKind.Columns.Add("Value", GetType(Integer))
        pKind.Rows.Add("定例総会", 41)
        pKind.Rows.Add("現地調査日", 51)
        pKind.Rows.Add("その他", 91)

        pCol2.DataSource = pKind
        mvarGrid.Columns.Add(pCol2)

        Dim pCol3 As New DataGridViewTextBoxColumn
        pCol3.Name = "内容"
        pCol3.DataPropertyName = "内容"
        mvarGrid.Columns.Add(pCol3)
        pCol3.MinimumWidth = 500
        Dim pCol4 As New DataGridViewCheckBoxColumn
        pCol4.Name = "削除"
        pCol4.DataPropertyName = "削除"
        pCol4.Visible = False
        mvarGrid.Columns.Add(pCol4)
        mvarGrid.Columns.Add(pCol0)

        mvarCal.Value = Now.Date

        mvarTab = New HimTools2012.controls.TabControlBase()
        mvarTab.AddNewPage(mvarGrid, "予定", "予定", False)

        mnuGContext = New HimTools2012.controls.ContextMenuEx(Nothing)
        AddHandler CType(mnuGContext.Items.Add("予定を削除"), ToolStripMenuItem).Click, AddressOf DoGClick
        mvarGrid.ContextMenuStrip = mnuGContext

        Me.ControlPanel.Add(mvarSplitter)
        mvarSplitter.Panel1.Controls.Add(mvarSplitterV)
        mvarSplitter.Panel2.Controls.Add(mvarTab)
        mvarDTP.Value = Now.Date
    End Sub



    Public Sub StartW()
        If watcher IsNot Nothing Then
            Return
        End If

        watcher = New System.IO.FileSystemWatcher
        watcher.Path = SysAD.CustomReportFolder(SysAD.市町村.市町村名)
        watcher.NotifyFilter = System.IO.NotifyFilters.LastAccess Or System.IO.NotifyFilters.LastWrite Or System.IO.NotifyFilters.FileName Or System.IO.NotifyFilters.DirectoryName
        watcher.Filter = "Task.xml"

        'watcher.SynchronizingObject = Me

        'イベントハンドラの追加
        AddHandler watcher.Changed, AddressOf watcher_Changed
        AddHandler watcher.Created, AddressOf watcher_Changed

        '監視を開始する
        watcher.EnableRaisingEvents = True

    End Sub
    Private Sub watcher_Changed(ByVal source As System.Object, ByVal e As System.IO.FileSystemEventArgs)
        Select Case e.ChangeType
            Case System.IO.WatcherChangeTypes.Changed, System.IO.WatcherChangeTypes.Created
                Dim pInfo As New IO.FileInfo(e.FullPath)
                If pInfo.LastWriteTime > mvarXMLUpdate Then
                    Try
                        Dim pTBL As New DataTable
                        pTBL.ReadXml(mvarXMLFile)
                        mvarTBL.Merge(pTBL, True, MissingSchemaAction.AddWithKey)
                    Catch ex As Exception
                        Stop
                    End Try
                    Dim ar日付 As New List(Of Date)
                    For Each pRow As DataRowView In New DataView(mvarTBL, "[削除]=False", "", DataViewRowState.CurrentRows)
                        If Not ar日付.Contains(pRow.Item("日時")) Then
                            ar日付.Add(CDate(pRow.Item("日時")).Date)
                        End If
                    Next
                    'mvarCal.BoldedDates = ar日付.ToArray
                    mvarXMLUpdate = Now
                End If
        End Select
    End Sub


    Public Sub DoGClick(s As Object, e As EventArgs)
        Select Case s.text
            Case "予定を削除"
                mvarGrid.Rows(mvarGrid.CurrentCell.RowIndex).Cells("削除").Value = True
                SaveTBL()
                mvarGrid.SetDataView(mvarTBL, "[削除]=False AND [日時]=" & HimTools2012.StringF.Toリテラル日付(mvarDTP.Value), "[日時]")
        End Select

    End Sub

    Public Sub DoClick(s As Object, e As EventArgs)
        Select Case s.text
            Case "総会日を追加"
                If MsgBox("総会日を[" & mvarValue.Date.ToString("yyyy/MM/dd") & "]に追加しますか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    Dim pRow As DataRow = mvarTBL.NewRow
                    pRow.Item("GID") = System.Guid.NewGuid
                    pRow.Item("日時") = mvarValue
                    pRow.Item("内容") = "総会日[" & mvarValue.Date.ToString("yyyy/MM/dd") & "]"
                    pRow.Item("種類") = 41
                    pRow.Item("削除") = False

                    mvarTBL.Rows.Add(pRow)
                    SaveTBL()
                End If
            Case "現地調査日を追加"
                If MsgBox("現地調査日を[" & mvarValue.Date.ToString("yyyy/MM/dd") & "]に追加しますか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    Dim pRow As DataRow = mvarTBL.NewRow
                    pRow.Item("GID") = System.Guid.NewGuid
                    pRow.Item("日時") = mvarValue
                    pRow.Item("内容") = "現地調査日[" & mvarValue.Date.ToString("yyyy/MM/dd") & "]"
                    pRow.Item("種類") = 51
                    pRow.Item("削除") = False
                    mvarTBL.Rows.Add(pRow)
                    SaveTBL()
                End If
            Case "その他予定を追加"
                If MsgBox("その他予定を[" & mvarValue.Date.ToString("yyyy/MM/dd") & "]に追加しますか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    Dim pRow As DataRow = mvarTBL.NewRow
                    pRow.Item("GID") = System.Guid.NewGuid
                    pRow.Item("日時") = mvarValue
                    pRow.Item("内容") = "その他予定[" & mvarValue.Date.ToString("yyyy/MM/dd") & "]:ここにコメントを入力"
                    pRow.Item("種類") = 91
                    pRow.Item("削除") = False
                    mvarTBL.Rows.Add(pRow)
                    SaveTBL()
                End If
            Case "指定範囲に受けた申請"
                Dim pGrid As HimTools2012.controls.DataGridViewWithDataView

                If mvarTab.TabPages.ContainsKey("受付申請") Then
                    pGrid = CType(mvarTab.TabPages("受付申請"), HimTools2012.controls.CTabPageWithToolStrip).ControlPanel(0)
                Else
                    pGrid = New HimTools2012.controls.DataGridViewWithDataView
                    mvarTab.AddNewPage(pGrid, "受付申請", "受付申請", True, True)
                End If
                pGrid.AllowUserToAddRows = False

                Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D_申請] WHERE [受付年月日]>={0} AND [受付年月日]<={1}", HimTools2012.StringF.Toリテラル日付(mvarDTP.Value), HimTools2012.StringF.Toリテラル日付(mvarDTP2.Value))
                App農地基本台帳.TBL申請.MergePlus(pTBL)
                pGrid.SetDataView(App農地基本台帳.TBL申請.Body, String.Format("[受付年月日]>={0} AND [受付年月日]<={1}", HimTools2012.StringF.Toリテラル日付(mvarDTP.Value), HimTools2012.StringF.Toリテラル日付(mvarDTP2.Value)), "[受付年月日]")
            Case "農地法４条申請許可"
                Dim pList As C申請リスト
                Dim sKey As String = "農地法4条許可済み"
                Dim sName As String = sKey
                Dim sWhere As String = String.Format("[法令]=40 AND [許可年月日]>={0} AND [許可年月日]<={1}", HimTools2012.StringF.Toリテラル日付(mvarDTP.Value), HimTools2012.StringF.Toリテラル日付(mvarDTP2.Value))

                If Not SysAD.page農家世帯.TabPageContainKey(sKey) Then
                    pList = New C申請リスト(SysAD.page農家世帯, sKey, sName)
                    pList.Name = sKey
                    SysAD.page農家世帯.中央Tab.AddPage(pList)
                    pList.ImageKey = pList.IconKey
                Else
                    pList = SysAD.page農家世帯.GetItem(sKey)
                End If


                pList.検索開始(sWhere, sWhere, "[許可年月日],[受付番号],[受付補助記号]")
            Case "農地法５条申請許可"
                Dim pList As C申請リスト
                Dim sKey As String = "農地法5条許可済み"
                Dim sName As String = sKey
                Dim sWhere As String = String.Format("[法令] IN (50,51,52) AND [許可年月日]>={0} AND [許可年月日]<={1}", HimTools2012.StringF.Toリテラル日付(mvarDTP.Value), HimTools2012.StringF.Toリテラル日付(mvarDTP2.Value))

                If Not SysAD.page農家世帯.TabPageContainKey(sKey) Then
                    pList = New C申請リスト(SysAD.page農家世帯, sKey, sName)
                    pList.Name = sKey
                    SysAD.page農家世帯.中央Tab.AddPage(pList)
                    pList.ImageKey = pList.IconKey
                Else
                    pList = SysAD.page農家世帯.GetItem(sKey)
                End If

                pList.検索開始(sWhere, sWhere, "[許可年月日],[受付番号],[受付補助記号]")
            Case Else

        End Select
    End Sub

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.NoMessage
        End Get
    End Property
    Private mvarMenu As ToolStripMenuItem

    Private mvarValue As DateTime
    Private mvarValue2 As DateTime

    Private Sub mvarCal_ClickDayItem(s As Object, pDate As Date, pButton As System.Windows.Forms.MouseButtons, ByVal pLocation As System.Drawing.Point) Handles mvarCal.ClickDayItem
        If pButton = Windows.Forms.MouseButtons.Right Then
            Dim mnuContext As New HimTools2012.controls.ContextMenuEx(Nothing)
            AddHandler mnuContext.AddMenu(和暦Format(pDate), , , , , , Color.White, Color.Navy).Click, AddressOf SetDay
            AddHandler CType(mnuContext.Items.Add("現地調査日を追加"), ToolStripMenuItem).Click, AddressOf DoClick
            AddHandler CType(mnuContext.Items.Add("総会日を追加"), ToolStripMenuItem).Click, AddressOf DoClick
            AddHandler CType(mnuContext.Items.Add("その他予定を追加"), ToolStripMenuItem).Click, AddressOf DoClick
            mnuContext.Items.Add(New ToolStripSeparator)

            AddHandler CType(mnuContext.Items.Add("指定範囲に受けた申請"), ToolStripMenuItem).Click, AddressOf DoClick
            With CType(mnuContext.Items.Add("指定範囲に許可された申請"), ToolStripMenuItem)
                AddHandler CType(.DropDownItems.Add("農地法４条申請許可"), ToolStripMenuItem).Click, AddressOf DoClick
                AddHandler CType(.DropDownItems.Add("農地法５条申請許可"), ToolStripMenuItem).Click, AddressOf DoClick
            End With
            mnuContext.Show(mvarCal, pLocation)
        End If
    End Sub
    Private Sub SetDay(s As Object, e As EventArgs)

        mvarCal.Value = CDate(s.text)
    End Sub

    Private Sub mvarCal_SelectionChanged(s As Object, e As System.EventArgs) Handles mvarCal.SelectionChanged
        mvarValue = mvarCal.SelectStart
        mvarDTP.Value = mvarValue
        mvarTime1.Value = mvarValue

        mvarValue2 = mvarCal.SelectEnd
        mvarDTP2.Value = mvarValue2
        mvarTime2.Value = mvarValue2

        mvarGrid.RowFilter = "[削除]=False AND [日時]>=" & HimTools2012.StringF.Toリテラル日付(mvarCal.SelectStart) & " AND [日時]<=" & HimTools2012.StringF.Toリテラル日付(mvarCal.SelectEnd)
    End Sub

    Private Sub mvarDTP_ValueChanged(sender As Object, e As System.EventArgs) Handles mvarDTP.ValueChanged
        mvarValue = mvarDTP.Value
        mvarCal.SelectStart = mvarValue
        mvarTime1.Value = mvarValue
    End Sub

    Private Sub mvarTime1_ValueChanged(sender As Object, e As System.EventArgs) Handles mvarTime1.ValueChanged
        mvarValue = mvarTime1.Value
        mvarCal.SelectStart = mvarValue
        mvarDTP.Value = mvarValue
    End Sub

    Private Sub mvarDTP2_ValueChanged(sender As Object, e As System.EventArgs) Handles mvarDTP2.ValueChanged
        mvarValue2 = mvarDTP2.Value
        mvarCal.SelectEnd = mvarValue2
        mvarTime2.Value = mvarValue2
    End Sub
    Private Sub mvarTime2_ValueChanged(sender As Object, e As System.EventArgs) Handles mvarTime2.ValueChanged
        mvarValue2 = mvarTime2.Value
        mvarCal.SelectEnd = mvarValue2
        mvarDTP2.Value = mvarValue2
    End Sub

    Private Sub mvarGrid_CellValidating(sender As Object, e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles mvarGrid.CellValidating
        Select Case mvarGrid.Columns(e.ColumnIndex).DataPropertyName
            Case "日時"
                If Not IsDate(mvarGrid.Item(e.ColumnIndex, e.RowIndex).Value) Then
                    e.Cancel = True
                End If
            Case "種類", "内容"
                If mvarGrid.Item(e.ColumnIndex, e.RowIndex).Value Is Nothing OrElse IsDBNull(mvarGrid.Item(e.ColumnIndex, e.RowIndex).Value) Then

                    e.Cancel = True
                End If
            Case "削除"
            Case "GID"
            Case Else
                If Not SysAD.IsClickOnceDeployed Then
                    Stop
                End If
        End Select
    End Sub

    Private Sub mvarGrid_CellValueChanged(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles mvarGrid.CellValueChanged
        SaveTBL()
    End Sub

    Private Sub mvarGrid_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles mvarGrid.KeyDown
        If e.KeyCode = Keys.Delete AndAlso MsgBox("削除しますか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            mvarGrid.Rows(mvarGrid.CurrentCell.RowIndex).Cells("削除").Value = True

            SaveTBL()
            mvarGrid.SetDataView(mvarTBL, "[削除]=False AND [日時]=" & HimTools2012.StringF.Toリテラル日付(mvarDTP.Value), "[日時]")
        End If
    End Sub

    Private Sub SaveTBL()
        SyncLock Me
            If watcher IsNot Nothing Then
                watcher.EnableRaisingEvents = False
            End If
            Dim ar日付 As New List(Of Date)
            For Each pRow As DataRowView In New DataView(mvarTBL, "[削除]=False", "", DataViewRowState.CurrentRows)
                If Not ar日付.Contains(pRow.Item("日時")) Then
                    ar日付.Add(CDate(pRow.Item("日時")).Date)
                End If
            Next
            'mvarCal.BoldedDates = ar日付.ToArray

            If Not IO.Directory.Exists(SysAD.CustomReportFolder(SysAD.市町村.市町村名)) Then
                IO.Directory.CreateDirectory(SysAD.CustomReportFolder(SysAD.市町村.市町村名))
            End If

            mvarTBL.WriteXml(mvarXMLFile, XmlWriteMode.WriteSchema)
            mvarXMLUpdate = Now

            StartW()
        End SyncLock
    End Sub


End Class

Public Class MonthView
    Inherits Windows.Forms.MonthCalendar

    Public Sub New()
        MyBase.New()

    End Sub



    'Protected Overrides Sub WndProc(ByRef m As Message)
    '    MyBase.WndProc(m)
    '    Select Case m.Msg
    '        Case 15
    '            Using gr As Graphics = Me.CreateGraphics
    '                Using sb As SolidBrush = New SolidBrush(Color.LightPink)
    '                    Dim CR As Rectangle = Me.Bounds

    '                    gr.DrawEllipse(Pens.Red, New Rectangle(15, 50, 15, Me.Font.Size))
    '                    '                gr.FillRectangle(sb, MyBase.ClientRectangle)
    '                End Using
    '            End Using

    '    End Select
    'End Sub


    'Private Sub MonthView_ClientSizeChanged(sender As Object, e As System.EventArgs) Handles Me.ClientSizeChanged
    '    Using gr As Graphics = Me.CreateGraphics
    '        Using sb As SolidBrush = New SolidBrush(Color.LightPink)
    '            gr.DrawEllipse(Pens.Red, New Rectangle(MyBase.ClientRectangle.X, MyBase.ClientRectangle.Y, Me.PreferredSize.Width, Me.PreferredSize.Height))
    '            '                gr.FillRectangle(sb, MyBase.ClientRectangle)
    '        End Using
    '    End Using

    'End Sub

    'Private Sub MonthView_RegionChanged(sender As Object, e As System.EventArgs) Handles Me.RegionChanged
    '    Using gr As Graphics = Me.CreateGraphics
    '        Using sb As SolidBrush = New SolidBrush(Color.LightPink)
    '            gr.FillRectangle(sb, MyBase.ClientRectangle)
    '        End Using
    '    End Using
    'End Sub
End Class
