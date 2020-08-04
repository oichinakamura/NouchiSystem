
Public Class CTabPageClassGenerator
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Private WithEvents mvarText As New RichTextBox
    Private WithEvents mvarError As New RichTextBox
    Private pTBL As DataTable
    Private mvarPropertyGrid As New PropertyGrid

    Public Sub New()
        MyBase.New(True, True, "クラス作成", "クラス作成")

        Dim sP As New SplitContainer
        sP.Dock = DockStyle.Fill
        sP.Orientation = Orientation.Horizontal

        Dim sP2 As New SplitContainer
        sP2.Dock = DockStyle.Fill
        sP2.Orientation = Orientation.Vertical

        mvarPropertyGrid.Dock = DockStyle.Fill

        mvarText.Dock = DockStyle.Fill
        sP2.Panel1.Controls.Add(mvarText)
        sP2.Panel2.Controls.Add(mvarPropertyGrid)

        sP.Panel1.Controls.Add(sP2)

        mvarError.Dock = DockStyle.Fill
        sP.Panel2.Controls.Add(mvarError)

        Me.ControlPanel.Add(sP)

        Dim pCls As New CCreateClass

        pTBL = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM D:農地Info WHERE [ID]>0 AND [ID]<100")

        Dim pRow As DataRow = pTBL.Rows(0)
        mvarText.Text = pCls.GetCode(pRow)
        AddHandler ToolStrip.Items.Add("実行").Click, AddressOf 実行
    End Sub

    Public Sub 実行()
        Dim pRow As DataRow = pTBL.Rows(0)
        Dim pCls As New CCreateClass
        Dim pObj As Object = pCls.GetObject("CDataRow", pRow, mvarError)
        If pObj IsNot Nothing Then
            mvarPropertyGrid.SelectedObject = pObj

        End If
    End Sub


    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.CloseOK
        End Get
    End Property


End Class
