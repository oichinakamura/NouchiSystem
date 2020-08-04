
Public Class CPage重複農地検索
    Inherits HimTools2012.TabPages.CTabPageWithDataGridView
    Public Sub New()
        MyBase.New(True, "重複農地検索.0", "重複農地検索", ObjectMan)

        App農地基本台帳.ListColumnDesign.SetGridColumns(mvarGrid, "農地リストColumns")
        mvarGrid.Columns.Remove("更新")
        Me.ControlPanel.Add(mvarGrid)
        If App農地基本台帳.TBL農地.Body.Columns.Contains("重複A") Then
            App農地基本台帳.TBL農地.Body.Columns.Remove(App農地基本台帳.TBL農地.Body.Columns("重複A"))
        End If
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:農地Info].*,1 AS [重複A] FROM [D:農地Info] WHERE ((([D:農地Info].[大字ID]) In (SELECT [大字ID] FROM [D:農地Info] As Tmp GROUP BY [大字ID],[地番] HAVING Count(*)>1  And [地番] = [D:農地Info].[地番]) And ([D:農地Info].[大字ID])>0)) ORDER BY [D:農地Info].[大字ID],[地番];")

        App農地基本台帳.TBL農地.MergePlus(pTBL)
        mvarGrid.Create件数表示Ctrl(Me.ToolStrip)
        mvarGrid.SetDataView(App農地基本台帳.TBL農地.Body, "[重複A]=1", "[大字],[地番]", HimTools2012.controls.DataGridViewWithDataView.AutoGenerateColumnsMode.AutoGenerateEnable)

    End Sub
    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.CloseOK
        End Get
    End Property
End Class
