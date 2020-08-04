
Imports System.ComponentModel

Public MustInherit Class 個人リスト共通
    Inherits CNList農地台帳

    Public Sub New(ByVal sKey As String, ByVal sText As String, ByVal bCloseable As Boolean)
        MyBase.New(sText, sKey, bCloseable)
    End Sub

    Public Sub New(ByRef pNode As Xml.XmlNode, ByRef pLayout As HimTools2012.controls.XMLLayout)
        MyBase.New(pNode, pLayout)
    End Sub

    Public Overrides Property IconKey As String
        Get
            Return "List"
        End Get
        Set(value As String)

        End Set
    End Property

End Class

Public Class C個人リスト
    Inherits 個人リスト共通

    Private WithEvents mvar基本台帳印刷 As ToolStripButton
    Private WithEvents mvar住民追加 As ToolStripButton
    Private WithEvents mvarToMap経営 As ToolStripButton

    Public Sub New()
        MyBase.New("個人リスト", "個人リスト", False)

        App農地基本台帳.ListColumnDesign.SetGridColumns(GView, "個人リストColumns")

        mvar基本台帳印刷 = New ToolStripButton("基本台帳印刷")
        Me.ToolStrip.Items.Add(mvar基本台帳印刷)

        mvar住民追加 = New ToolStripButton("住民・法人の追加")
        Me.ToolStrip.Items.Add(mvar住民追加)

        mvar検索Page = New CPage検索("個人検索条件", "個人検索", New C個人検索条件, False, Me)

        mvarToMap経営 = New ToolStripButton("経営農地を地図で呼ぶ")
        mvarToMap経営.Enabled = False
        Me.ToolStrip.Items.Add(mvarToMap経営)

        AddHandler mvarGrid.CellPainting, AddressOf mvarGrid_CellPainting
        mvarlabel = New ToolStripLabel("")
        Me.ToolStrip.Items.Add(mvarlabel)
    End Sub

    Public Sub New(ByRef pNode As Xml.XmlNode, ByRef pLayout As HimTools2012.controls.XMLLayout)
        MyBase.New(pNode, pLayout)

        App農地基本台帳.ListColumnDesign.SetGridColumns(GView, "個人リストColumns")

        mvar基本台帳印刷 = New ToolStripButton("基本台帳印刷")
        Me.ToolStrip.Items.Add(mvar基本台帳印刷)

        mvar住民追加 = New ToolStripButton("住民・法人の追加")
        Me.ToolStrip.Items.Add(mvar住民追加)

        mvar検索Page = New CPage検索("個人検索条件", "個人検索", New C個人検索条件, False, Me)
        For Each pSideNode As Xml.XmlNode In pNode.ChildNodes
            Select Case pSideNode.Name
                Case "SearchBlcokPanel"
                    If mvar検索Page IsNot Nothing AndAlso pLayout.Controls.ContainsKey(pSideNode.Attributes("ContainerBlockPanel").Value) Then
                        CType(pLayout.Controls.Item(pSideNode.Attributes("ContainerBlockPanel").Value), HimTools2012.controls.BlockPanelControl).BlockPanels.Add(mvar検索Page, False)
                    End If
                Case Else
                    Stop
            End Select
        Next

        mvarToMap経営 = New ToolStripButton("経営農地を地図で呼ぶ")
        mvarToMap経営.Enabled = False
        Me.ToolStrip.Items.Add(mvarToMap経営)

        AddHandler mvarGrid.CellPainting, AddressOf mvarGrid_CellPainting
        mvarlabel = New ToolStripLabel("")
        Me.ToolStrip.Items.Add(mvarlabel)
    End Sub

    Private Sub mvarGrid_CellPainting(sender As Object, e As System.Windows.Forms.DataGridViewCellPaintingEventArgs)
        If e.RowIndex > 0 Then
            With CType(mvarGrid.Rows(e.RowIndex).DataBoundItem, DataRowView)
                If InStr(.Item("住民区分名").ToString, "死亡") Then
                    e.CellStyle.BackColor = Color.LightGray
                End If
            End With

        End If
    End Sub

    Private Sub mvar基本台帳印刷_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar基本台帳印刷.Click
        If GView.SelectedCells IsNot Nothing AndAlso GView.SelectedCells.Count > 0 AndAlso GView.SelectedCells(0).RowIndex > -1 Then
            mod農地基本台帳.基本台帳印刷("個人." & GView.Item("ID", GView.SelectedCells(0).RowIndex).Value, ExcelViewMode.Preview, 印刷Mode.フル印刷)
        End If
    End Sub

    Public Overrides Sub 検索開始(ByVal sWhere As String, ByVal sViewWhere As String, Optional sOrderBy As String = "", Optional sColumnStyle As String = "")
        SyncLock Me

            If sWhere IsNot Nothing AndAlso sWhere.Length > 0 Then
                Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:個人Info] WHERE " & sWhere)
                App農地基本台帳.TBL個人.MergePlus(pTBL)

                If GView.DataView Is Nothing Then
                    GView.SetDataView(App農地基本台帳.TBL個人.Body, sViewWhere, "[フリガナ]")
                Else
                    GView.DataView.RowFilter = sViewWhere
                End If
                mvarlabel.Text = GView.Rows.Count & "件表示"

            ElseIf GView.DataView Is Nothing Then
                GView.SetDataView(App農地基本台帳.TBL個人.Body, sViewWhere, "[フリガナ]")
            Else
                GView.DataView.RowFilter = sViewWhere
            End If

        End SyncLock

        Me.Active()
        GView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCellsExceptHeader)
        GView.SetColumnSortMode(DataGridViewColumnSortMode.NotSortable)
    End Sub

    Private Sub mvar住民追加_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar住民追加.Click
        App農地基本台帳.CreateFarmer()
    End Sub

    Private Sub mvarToMap経営_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarToMap経営.Click
        If SysAD.MapConnection.HasMap Then
            Dim sA As New System.Text.StringBuilder
            Dim sB As New System.Text.StringBuilder

            sB.AppendLine("Clear:0")
            sB.AppendLine("LogicMode:2")
            sB.AppendLine("PaintMode:1")

            For Each pRow As DataGridViewRow In GView.SelectedRows
                sA.Append(IIf(sA.Length > 0, ",", "") & pRow.Cells("ID").Value.ToString)
            Next
            Dim pTable As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(String.Format("SELECT [ID] FROM [D:農地Info] WHERE ([自小作別]=0 AND [所有者ID] IN ({0})) Or ([自小作別]<>0 AND [借受人ID] IN ({0}))", sA.ToString))


            For Each pRow As DataRow In pTable.Rows
                sB.AppendLine("LotIDP:" & pRow.Item("ID").ToString & ",1")
            Next

            SysAD.MapConnection.SelectMap(sB.ToString)
        Else
            mvarToMap経営.Enabled = False
        End If

    End Sub
End Class


Public Class C住民記録リスト
    Inherits 個人リスト共通

    Public Sub New(ByVal sName As String, ByVal sText As String)
        MyBase.New(sName, sText, True)
        App農地基本台帳.ListColumnDesign.SetGridColumns(GView, "住民記録リストColumns")
    End Sub

    Public Overrides Sub 検索開始(sWhere As String, sViewWhere As String, Optional sOrderBy As String = "", Optional sColumnStyle As String = "")
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT 'Human' AS [アイコン],'住基.' & [ID] AS [Key],* FROM [M_住民情報] WHERE " & sWhere)

        GView.SetDataView(pTBL, sViewWhere, "")
        GView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCellsExceptHeader)

    End Sub
End Class

Public Class C削除個人リスト
    Inherits 個人リスト共通

    Public Sub New(ByVal sName As String, ByVal sText As String)
        MyBase.New(sName, sText, True)
        App農地基本台帳.ListColumnDesign.SetGridColumns(GView, "削除個人リストColumns")
    End Sub

    Public Overrides Sub 検索開始(sWhere As String, sViewWhere As String, Optional sOrderBy As String = "", Optional sColumnStyle As String = "")
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT 'Human' AS [アイコン],'削除個人.' & [ID] AS [Key],* FROM [D_削除個人] WHERE " & sWhere)
        App農地基本台帳.TBL削除個人.MergePlus(pTBL)

        GView.SetDataView(App農地基本台帳.TBL削除個人.Body, sViewWhere, "")
        GView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCellsExceptHeader)
    End Sub
End Class
