
Public Class CTabPage各種グリッド設定
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Private mvarSP As SplitContainer
    Private mvarSP2 As SplitContainer
    Private WithEvents mvarLst As HimTools2012.controls.DataGridViewWithDataView

    Private WithEvents mvarLstTBL As ListView
    Private mvarGrid As HimTools2012.controls.DataGridViewWithDataView
    Private mvarDSet As CListColumnDesign

    Public Sub New()
        MyBase.New(True, True, "各種グリッド設定", "各種グリッド設定")

        mvarSP = New SplitContainer
        mvarSP.Dock = DockStyle.Fill

        mvarSP2 = New SplitContainer
        mvarSP2.Dock = DockStyle.Fill
        mvarSP2.Orientation = Orientation.Horizontal

        mvarLst = New HimTools2012.controls.DataGridViewWithDataView
        mvarLst.MultiSelect = False
        mvarGrid = New HimTools2012.controls.DataGridViewWithDataView
        Dim mvarTab As New HimTools2012.controls.TabControlBase()

        mvarLstTBL = New ListView
        mvarLstTBL.Dock = DockStyle.Fill

        mvarLstTBL.BackColor = mvarGrid.BackgroundColor

        For Each pTBL As DataTable In App農地基本台帳.DSet.Tables
            If Not pTBL.TableName.EndsWith("Table") Then
                mvarLstTBL.Items.Add(pTBL.TableName)
            End If
        Next

        Dim strXML As String = My.Resources.Resource1.DataListViewカラム設定
        mvarDSet = New CListColumnDesign


        If Not mvarDSet.Tables.Contains("GridTable管理") Then
            Dim pTBL As New DataTable("GridTable管理")
            pTBL.Columns.Add("名称")
            pTBL.Columns.Add("BaseTable")
            pTBL.PrimaryKey = {pTBL.Columns("名称")}
            mvarDSet.Tables.Add(pTBL)
        End If

        For Each pTBL As DataTable In mvarDSet.Tables
            If Not pTBL.TableName = "GridTable管理" AndAlso mvarDSet.Tables("GridTable管理").Rows.Find(pTBL.TableName) Is Nothing Then
                mvarDSet.Tables("GridTable管理").Rows.Add(pTBL.TableName)
            End If
        Next

        Dim pDelList As New List(Of DataRow)
        For Each pRow As DataRow In mvarDSet.Tables("GridTable管理").Rows
            If Not mvarDSet.Tables.Contains(pRow.Item("名称")) Then
                pDelList.Add(pRow)
            End If
        Next
        For Each pDelRow As DataRow In pDelList
            mvarDSet.Tables("GridTable管理").Rows.Remove(pDelRow)
        Next

        mvarLst.SetDataView(mvarDSet.Tables("GridTable管理"), "", "")
        mvarTab.AddNewPage(mvarGrid, "Grid詳細", "Grid詳細", False, True, True)

        mvarSP2.Panel1.Controls.Add(mvarLst)
        mvarSP2.Panel2.Controls.Add(mvarLstTBL)

        mvarSP.Panel1.Controls.Add(mvarSP2)
        mvarSP.Panel2.Controls.Add(mvarTab)
        Me.ControlPanel.Add(mvarSP)

        AddHandler Me.ToolStrip.Items.Add("保存").Click, Sub(s, e) mvarDSet.Save("農地基本台帳")
    End Sub

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.NoMessage
        End Get
    End Property

    Private Sub mvarLst_SelectionChanged(sender As Object, e As System.EventArgs) Handles mvarLst.SelectionChanged
        If mvarLst.SelectedRows IsNot Nothing AndAlso mvarLst.SelectedRows.Count > 0 Then
            Dim sName As String = mvarLst.SelectedRows(0).Cells("名称").Value
            If mvarDSet.Tables.Contains(sName) Then
                mvarGrid.SetDataView(mvarDSet.Tables(sName), "", "")

            End If

        End If
    End Sub

    Private Sub mvarLstTBL_MouseMove(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles mvarLstTBL.MouseMove
        If e.Button = Windows.Forms.MouseButtons.Left Then
            Dim pItem As ListViewHitTestInfo = mvarLstTBL.HitTest(e.Location)

            If pItem.Item IsNot Nothing Then
                mvarLstTBL.DoDragDrop(pItem.Item, DragDropEffects.Copy)
            End If
        End If
    End Sub
End Class
