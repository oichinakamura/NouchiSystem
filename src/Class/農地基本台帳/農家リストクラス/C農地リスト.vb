
Public MustInherit Class 農地共通リスト
    Inherits CNList農地台帳

    Public Sub New(ByVal sText As String, ByVal sName As String, Optional bCloseable As Boolean = False)
        MyBase.New(sText, sName, bCloseable)

    End Sub
    Public Sub New(ByRef pPNode As Xml.XmlNode, ByRef pLayout As HimTools2012.controls.XMLLayout)
        MyBase.New(pPNode, pLayout)
    End Sub

    Public Sub Sub検索開始(ByRef BaseTBL As HimTools2012.Data.DataTableWith, ByVal sTableName As String, ByVal s検索文字列 As String, ByVal sView検索文字列 As String, Optional sOrderBy As String = "")
        SyncLock Me
            ColumnsEx.ClearFilter()
            If s検索文字列.Length > 0 Then
                Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [{0}] WHERE {1}", sTableName, s検索文字列)
                BaseTBL.MergePlus(pTBL)
                GView.SetDataView(BaseTBL.Body, Replace(sView検索文字列, "%", "*"), "")
            ElseIf GView.DataView Is Nothing AndAlso s検索文字列.Length > 0 Then
                GView.SetDataView(BaseTBL.Body, Replace(sView検索文字列, "%", "*"), "")
            ElseIf s検索文字列.Length > 0 Then
                GView.DataView.RowFilter = Replace(sView検索文字列, "%", "*")
            End If
        End SyncLock
        Me.Active()
        GView.SetColumnSortMode(DataGridViewColumnSortMode.NotSortable)

        GView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCellsExceptHeader)
        GView.Refresh()

    End Sub
    Public Sub Sub検索開始(ByRef BaseTBL As HimTools2012.Data.DataTableEx, ByVal sTableName As String, ByVal s検索文字列 As String, ByVal sView検索文字列 As String, Optional sOrderBy As String = "")
        SyncLock Me
            ColumnsEx.ClearFilter()
            If s検索文字列.Length > 0 Then
                Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [{0}] WHERE {1}", sTableName, s検索文字列)
                BaseTBL.MergePlus(pTBL)
                GView.SetDataView(BaseTBL, Replace(sView検索文字列, "%", "*"), "")
            ElseIf GView.DataView Is Nothing AndAlso s検索文字列.Length > 0 Then
                GView.SetDataView(BaseTBL, Replace(sView検索文字列, "%", "*"), "")
            ElseIf s検索文字列.Length > 0 Then
                GView.DataView.RowFilter = Replace(sView検索文字列, "%", "*")
            End If
        End SyncLock
        Me.Active()
        GView.SetColumnSortMode(DataGridViewColumnSortMode.NotSortable)

        GView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader)
        GView.Refresh()

    End Sub

    Public Overrides Property IconKey As String
        Get
            Return "List"
        End Get
        Set(value As String)

        End Set
    End Property
End Class


Public Class C農地リスト
    Inherits 農地共通リスト

    Private WithEvents mvar基本台帳印刷 As ToolStripButton
    Private WithEvents mvarToMap As ToolStripButton
    Private WithEvents mvar換地テーブルより As New ToolStripButton("換地テーブルより")

    Public Sub New()
        MyBase.New("農地リスト", "農地リスト")
        App農地基本台帳.ListColumnDesign.SetGridColumns(GView, "農地リストColumns")

        mvarToMap = New ToolStripButton("地図を呼ぶ", My.Resources.Resource1.Map.ToBitmap, AddressOf mvarToMap_Click)
        mvarToMap.Enabled = False
        Me.ToolStrip.Items.Add(mvarToMap)

        Dim Btn複数地番 = New ToolStripButton("複数地番で検索")
        AddHandler Btn複数地番.Click, AddressOf 複数地番検索
        Me.ToolStrip.Items.Add(Btn複数地番)

        mvar検索Page = New CPage検索("農地検索条件", "農地検索", New C農地検索条件, False, Me)

    End Sub

    Public Sub New(ByRef pNode As Xml.XmlNode, ByRef pLayout As HimTools2012.controls.XMLLayout)
        MyBase.New(pNode, pLayout)
        App農地基本台帳.ListColumnDesign.SetGridColumns(GView, "農地リストColumns")

        mvarToMap = New ToolStripButton("地図を呼ぶ", My.Resources.Resource1.Map.ToBitmap, AddressOf mvarToMap_Click)
        mvarToMap.Enabled = False
        Me.ToolStrip.Items.Add(mvarToMap)

        Dim Btn複数地番 = New ToolStripButton("複数地番で検索")
        AddHandler Btn複数地番.Click, AddressOf 複数地番検索
        Me.ToolStrip.Items.Add(Btn複数地番)

        mvar検索Page = New CPage検索("農地検索条件", "農地検索", New C農地検索条件, False, Me)
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
    End Sub


    Public Overrides Sub 検索開始(ByVal sWhere As String, ByVal sViewWhere As String, Optional sOrderBy As String = "", Optional sColumnStyle As String = "")
        Sub検索開始(App農地基本台帳.TBL農地, "D:農地Info", sWhere, sViewWhere)
    End Sub

    Public Sub 複数地番検索()
        Using pDlg As New dlg複数地番検索
            With pDlg
                If .ShowDialog() = DialogResult.OK Then
                    検索開始(.sResult, .sResult)
                End If
            End With
        End Using
    End Sub

    Private Sub mvarToMap_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarToMap.Click
        If SysAD.MapConnection.HasMap Then
            Dim sB As New System.Text.StringBuilder

            sB.AppendLine("Clear:0")
            sB.AppendLine("LogicMode:2")
            sB.AppendLine("PaintMode:1")

            For Each pRow As DataGridViewRow In GView.SelectedRows
                sB.AppendLine("LotIDP:" & pRow.Cells("ID").Value.ToString & ",1")
            Next

            SysAD.MapConnection.SelectMap(sB.ToString)
        Else
            mvarToMap.Enabled = False
        End If
    End Sub

    Private Sub mvar換地テーブルより_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar換地テーブルより.Click
        If SysAD.page農家世帯.TabPageContainKey("換地処理") Then
            MsgBox("既に換地処理が実行されています。処理が完了するか中断してください", MsgBoxStyle.Critical)
        Else
            Dim pTBL換地 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM 換地データ INNER JOIN [D:農地Info] ON 換地データ.旧ID = [D:農地Info].ID WHERE [新ID] Is Null ORDER BY [No]")
            Dim GID As Integer = -1

            Dim sGroup As New List(Of String)
            Dim pNewID As Decimal = 0
            Dim s新地番 As String = ""
            Dim n大字 As Integer = 0

            For Each pRow As DataRow In pTBL換地.Rows
                If GID = -1 Then
                    GID = pRow.Item("No")
                    sGroup.Clear()
                    sGroup.Add(pRow.Item("旧ID"))
                    Dim pNRow As DataRow = App農地基本台帳.TBL農地.FindRowByID(pRow.Item("旧ID"))
                    n大字 = pNRow.Item("大字ID")
                    s新地番 = pRow.Item("新地番")
                ElseIf GID = pRow.Item("No") Then
                    If Not s新地番 = pRow.Item("新地番") Then
                        MsgBox("換地後地番が異なります", MsgBoxStyle.Critical)
                        Exit Sub
                    End If
                    sGroup.Add(pRow.Item("旧ID"))
                Else
                    If sGroup.Count > 0 Then
                        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [M_固定情報] WHERE [大字ID]={0} AND [地番]='{1}'", n大字, s新地番)
                        If pTBL.Rows.Count > 0 Then
                            Dim pPage = New CTabPage換地処理(sGroup.ToArray, pTBL.Rows(0).Item("nID"))
                        Else
                            Dim pPage = New CTabPage換地処理(sGroup.ToArray, 0)
                        End If

                    End If
                    GID = pRow.Item("No")
                    sGroup.Clear()
                    sGroup.Add(pRow.Item("旧ID"))
                    Dim pNRow As DataRow = App農地基本台帳.TBL農地.FindRowByID(pRow.Item("旧ID"))
                    n大字 = pNRow.Item("大字ID")
                    s新地番 = pRow.Item("新地番")
                End If
            Next

            If sGroup.Count > 0 Then
                Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [M_固定情報] WHERE [大字ID]={0} AND [地番]='{1}'", n大字, s新地番)
                If pTBL.Rows.Count > 0 Then
                    Dim pPage = New CTabPage換地処理(sGroup.ToArray, pTBL.Rows(0).Item("nID"))
                Else
                    Dim pPage = New CTabPage換地処理(sGroup.ToArray, 0)
                End If

            End If
        End If
    End Sub
End Class

Public Class C土地台帳リスト
    Inherits 農地共通リスト

    Public Sub New(ByVal sKey As String, ByVal sText As String)
        MyBase.New(sText, sKey, True)
        App農地基本台帳.ListColumnDesign.SetGridColumns(GView, "固定資産土地リスト")
    End Sub
    Public Overrides Sub 検索開始(ByVal sWhere As String, ByVal sViewWhere As String, Optional sOrderBy As String = "", Optional sColumnStyle As String = "")
        Sub検索開始(App農地基本台帳.TBL固定情報, "M_固定情報", sWhere, sViewWhere)
    End Sub
End Class

Public Class C転用農地リスト
    Inherits 農地共通リスト

    Public Sub New(ByVal sKey As String, ByVal sText As String)
        MyBase.New(sText, sKey, True)
        App農地基本台帳.ListColumnDesign.SetGridColumns(GView, "転用農地Columns")
    End Sub

    Public Overrides Sub 検索開始(ByVal sWhere As String, ByVal sViewWhere As String, Optional sOrderBy As String = "", Optional sColumnStyle As String = "")
        Sub検索開始(App農地基本台帳.TBL転用農地, "D_転用農地", sWhere, sViewWhere)
    End Sub
End Class

Public Class C削除農地リスト
    Inherits 農地共通リスト

    Public Sub New(ByVal sKey As String, ByVal sText As String)
        MyBase.New(sText, sKey, True)
        App農地基本台帳.ListColumnDesign.SetGridColumns(GView, "削除農地Columns")
    End Sub

    Public Overrides Sub 検索開始(ByVal sWhere As String, ByVal sViewWhere As String, Optional sOrderBy As String = "", Optional sColumnStyle As String = "")
        Sub検索開始(App農地基本台帳.TBL削除農地, "D_削除農地", sWhere, sViewWhere)
    End Sub
End Class



