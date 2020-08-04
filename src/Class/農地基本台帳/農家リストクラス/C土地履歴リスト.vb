
Imports System.ComponentModel
Imports System.Drawing.Design
Imports System.Windows.Forms.Design

Public Class C土地履歴リスト
    Inherits CNList農地台帳

    Private WithEvents mvar基本台帳印刷 As ToolStripButton
    Private WithEvents mvar住民追加 As ToolStripButton
    Private WithEvents mvarToMap経営 As ToolStripButton

    Public Sub New()
        MyBase.New("土地履歴リスト", "土地履歴リスト", False)

        Dim imageColumn As New DataGridViewImageColumn()
        With imageColumn
            .Image = Nothing
            .ImageLayout = DataGridViewImageCellLayout.Zoom
            .Name = "アイコンImage"
            .HeaderText = ""
            .DefaultCellStyle.NullValue = Nothing
            .Width = 50
        End With

        GView.Columns.Add(imageColumn)

        AddColumn("ID", "ID")
        AddColumn("LID", "LID")
        AddColumn("農地土地所在", "土地所在")
        AddColumn("異動日", "異動日")
        AddColumn("内容", "内容")

        'AddColumn("更新日", "更新日")
        AddColumn("入力日", "更新日")

        Dim pKeyColumn As New DataGridViewTextBoxColumn
        pKeyColumn.DataPropertyName = "Key"
        pKeyColumn.Name = "Key"
        pKeyColumn.Visible = False
        GView.Columns.Add(pKeyColumn)

        Dim pIconColumn As New DataGridViewTextBoxColumn
        pIconColumn.DataPropertyName = "アイコン"
        pIconColumn.Name = "アイコン"
        pIconColumn.Visible = False
        GView.Columns.Add(pIconColumn)

    End Sub
    Public Sub New(ByRef pNode As Xml.XmlNode, ByRef pLayout As HimTools2012.controls.XMLLayout)
        MyBase.New(pNode, pLayout)

        Dim imageColumn As New DataGridViewImageColumn()
        With imageColumn
            .Image = Nothing
            .ImageLayout = DataGridViewImageCellLayout.Zoom
            .Name = "アイコンImage"
            .HeaderText = ""
            .DefaultCellStyle.NullValue = Nothing
            .Width = 50
        End With

        GView.Columns.Add(imageColumn)

        AddColumn("ID", "ID")
        AddColumn("LID", "LID")
        AddColumn("農地土地所在", "土地所在")
        AddColumn("異動日", "異動日")
        AddColumn("内容", "内容")
        AddColumn("結果", "結果")

        'AddColumn("更新日", "更新日")
        AddColumn("入力日", "更新日")

        Dim pKeyColumn As New DataGridViewTextBoxColumn
        pKeyColumn.DataPropertyName = "Key"
        pKeyColumn.Name = "Key"
        pKeyColumn.Visible = False
        GView.Columns.Add(pKeyColumn)

        Dim pIconColumn As New DataGridViewTextBoxColumn
        pIconColumn.DataPropertyName = "アイコン"
        pIconColumn.Name = "アイコン"
        pIconColumn.Visible = False
        GView.Columns.Add(pIconColumn)
    End Sub
    Private Sub AddColumn(sFiled As String, sHeaderText As String)
        Dim pColumn As DataGridViewColumn

        Select Case App農地基本台帳.TBL土地履歴.Columns(sFiled).DataType.ToString
            Case "System.Int32", "System.Decimal"
                pColumn = New DataGridViewTextBoxColumn
            Case "System.DateTime"
                pColumn = New DataGridViewTextBoxColumn
            Case "System.String"
                pColumn = New DataGridViewTextBoxColumn
            Case Else
                pColumn = New DataGridViewTextBoxColumn
                Stop
        End Select
        pColumn.DataPropertyName = sFiled
        pColumn.Name = sFiled
        pColumn.HeaderText = sHeaderText
        pColumn.Visible = True
        GView.Columns.Add(pColumn)
    End Sub

    Public Overrides Sub 検索開始(sWhere As String, sViewWhere As String, Optional sOrderBy As String = "", Optional sColumnStyle As String = "")
        SyncLock Me
            If sWhere IsNot Nothing AndAlso sWhere.Length > 0 Then
                Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D_土地履歴] WHERE " & sWhere)
                App農地基本台帳.TBL土地履歴.MergePlus(pTBL)
                If GView.DataView Is Nothing Then
                    GView.SetDataView(App農地基本台帳.TBL土地履歴.Body, sWhere, "異動日")
                Else
                    GView.SetDataView(App農地基本台帳.TBL土地履歴.Body, sWhere, "異動日")
                End If


            ElseIf GView.DataView Is Nothing Then
                GView.SetDataView(App農地基本台帳.TBL土地履歴.Body, sViewWhere, "異動日")
            Else
                GView.DataView.RowFilter = sViewWhere
            End If
            GView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCellsExceptHeader)

        End SyncLock
        GView.SetColumnSortMode(DataGridViewColumnSortMode.NotSortable)

        Me.Active()
    End Sub

    Public Overrides Property IconKey As String
        Get
            Return "List"
        End Get
        Set(value As String)

        End Set
    End Property
End Class
