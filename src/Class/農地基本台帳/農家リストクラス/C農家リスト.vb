
Imports System.Reflection
Imports System.Globalization

Public Class C農家リスト
    Inherits CNList農地台帳

    Private WithEvents mvar基本台帳印刷 As ToolStripButton
    Private WithEvents mvar検索条件Combo As ToolStripComboBox


    Public Overrides Property IconKey As String
        Get
            Return "List"
        End Get
        Set(value As String)

        End Set
    End Property

    Public Sub New()
        MyBase.New("農家リスト", "農家リスト")

        App農地基本台帳.ListColumnDesign.SetGridColumns(GView, "農家リストColumns")

        mvar基本台帳印刷 = New ToolStripButton("基本台帳印刷")
        Me.ToolStrip.Items.Add(mvar基本台帳印刷)

        mvar検索条件Combo = New ToolStripComboBox
        mvar検索条件Combo.Alignment = ToolStripItemAlignment.Right

        Me.ToolStrip.Items.Add(mvar検索条件Combo)
        mvar検索条件Combo.AutoSize = False
        mvar検索条件Combo.Width = mvar検索条件Combo.Width * 2

        mvar検索Page = New CPage検索("農家検索条件", "農家検索", New C農家検索条件, False)

    End Sub
    Public Sub New(ByRef pNode As Xml.XmlNode, ByRef pLayout As HimTools2012.controls.XMLLayout)
        MyBase.New(pNode, pLayout)
        App農地基本台帳.ListColumnDesign.SetGridColumns(GView, "農家リストColumns")

        mvar基本台帳印刷 = New ToolStripButton("基本台帳印刷")
        Me.ToolStrip.Items.Add(mvar基本台帳印刷)

        mvar検索条件Combo = New ToolStripComboBox
        mvar検索条件Combo.Alignment = ToolStripItemAlignment.Right

        Me.ToolStrip.Items.Add(mvar検索条件Combo)
        mvar検索条件Combo.AutoSize = False
        mvar検索条件Combo.Width = mvar検索条件Combo.Width * 2

        mvar検索Page = New CPage検索("農家検索条件", "農家検索", New C農家検索条件, False)
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
    Private Sub mvar基本台帳_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar基本台帳印刷.Click
        If GView.SelectedCells IsNot Nothing AndAlso GView.SelectedCells.Count > 0 AndAlso GView.SelectedCells(0).RowIndex > -1 Then
            mod農地基本台帳.基本台帳印刷(GView.Item("Key", GView.SelectedCells(0).RowIndex).Value, ExcelViewMode.Preview, 印刷Mode.フル印刷)
        End If
    End Sub

    Private Sub mvar検索Page_検索(ByVal s検索文字列 As String, ByVal sView検索文字列 As String) Handles mvar検索Page.検索
        mvar検索条件Combo.Items.Insert(0, sView検索文字列)
        mvar検索条件Combo.Text = sView検索文字列
        Me.検索開始(s検索文字列, sView検索文字列)
    End Sub

    Public Overrides Sub 検索開始(ByVal sWhere As String, ByVal sViewWhere As String, Optional sOrderBy As String = "", Optional sColumnStyle As String = "")
        If sWhere.Length > 0 Then
            Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(
            "SELECT [D:世帯Info].* FROM [D:個人Info] INNER JOIN [D:世帯Info] ON [D:個人Info].ID = [D:世帯Info].世帯主ID WHERE " & sWhere)

            App農地基本台帳.TBL世帯.MergePlus(pTBL)
        Else

        End If

        If GView.DataView Is Nothing Then
            Dim b As Boolean = False
            Do Until b
                Try
                    GView.SetDataView(App農地基本台帳.TBL世帯.Body, sViewWhere, "[ID]")

                    b = True
                Catch ex As Exception

                End Try
            Loop
        Else
            GView.DataView.RowFilter = sViewWhere
        End If


        Me.Active()
        GView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCellsExceptHeader)
        GView.SetColumnSortMode(DataGridViewColumnSortMode.NotSortable)


    End Sub

    Public Sub あっせん希望()
        検索開始("[D:世帯Info].あっせん希望=True", "[あっせん希望]=True")
        Me.Active()
    End Sub


End Class

