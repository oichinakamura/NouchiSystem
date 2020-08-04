

Public Class CH26各種集計
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Private mvarTabControl As New TabControl
    Private mvarTab利用状況 As New TabPage
    Private mvarTab利用意向 As New TabPage

    Private mvarG利用状況 As New HimTools2012.controls.DataGridViewWithDataView
    Private mvarG利用状況市外 As New HimTools2012.controls.DataGridViewWithDataView
    Private mvarG農地法32の1 As New HimTools2012.controls.DataGridViewWithDataView
    Private mvarG農地法32の4 As New HimTools2012.controls.DataGridViewWithDataView
    Private mvarG農地法33の1 As New HimTools2012.controls.DataGridViewWithDataView
    Private mvarG根拠未入力 As New HimTools2012.controls.DataGridViewWithDataView

    Public Sub New()
        MyBase.New(True, False, "H26各種集計", "H26各種集計")

        Dim mvarSP利用状況 As New SplitContainer
        With mvarSP利用状況
            .Orientation = Orientation.Horizontal
            .Dock = DockStyle.Fill
            .Panel1.Controls.Add(SetPanel("市内住民のみ対象", mvarG利用状況))
            .Panel2.Controls.Add(SetPanel("市外住民を含む", mvarG利用状況市外))

            mvarTab利用状況.Controls.Add(mvarSP利用状況)
        End With

        Dim mvarSP利用意向 As New SplitContainer
        With mvarSP利用意向
            .Orientation = Orientation.Horizontal
            .Dock = DockStyle.Fill

            Dim mvarSP1 As New SplitContainer
            Dim mvarSP2 As New SplitContainer

            mvarSP1.Orientation = Orientation.Horizontal
            mvarSP1.Dock = DockStyle.Fill

            mvarSP2.Orientation = Orientation.Horizontal
            mvarSP2.Dock = DockStyle.Fill

            .Panel1.Controls.Add(mvarSP1)
            .Panel2.Controls.Add(mvarSP2)
            mvarSP1.Panel1.Controls.Add(SetPanel("根拠条項:農地法32の1", mvarG農地法32の1))
            mvarSP1.Panel2.Controls.Add(SetPanel("根拠条項:農地法32の4", mvarG農地法32の4))
            mvarSP2.Panel1.Controls.Add(SetPanel("根拠条項:農地法33の1", mvarG農地法33の1))
            mvarSP2.Panel2.Controls.Add(SetPanel("根拠条項:未入力", mvarG根拠未入力))

            mvarTab利用意向.Controls.Add(mvarSP利用意向)
        End With


        mvarTabControl.Controls.Add(mvarTab利用状況)
        mvarTabControl.Controls.Add(mvarTab利用意向)
        mvarTabControl.Dock = DockStyle.Fill
        Me.ControlPanel.Add(mvarTabControl)


        Dim TBL利用意向調査 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:農地Info].利用意向意向内容区分, '' AS 利用意向内容, [D:農地Info].利用意向根拠条項, '' AS 利用意向, Int(Sum([D:農地Info].登記簿面積)) AS 登記簿面積の合計, Int(Sum([D:農地Info].実面積)) AS 実面積の合計, Int(Sum([D:農地Info].田面積)) AS 田面積の合計, Int(Sum([D:農地Info].畑面積)) AS 畑面積の合計, Int(Sum([D:農地Info].樹園地)) AS 樹園地の合計, Count([D:農地Info].ID) AS 筆数 FROM [D:農地Info] GROUP BY [D:農地Info].利用意向意向内容区分, '', [D:農地Info].利用意向根拠条項, '' HAVING ((([D:農地Info].利用意向意向内容区分)>0));")

        For Each pRow As DataRow In TBL利用意向調査.Rows
            Select Case Val(pRow.Item("利用意向意向内容区分").ToString)
                Case 1 : pRow.Item("利用意向内容") = "自ら耕作"
                Case 2 : pRow.Item("利用意向内容") = "機構事業"
                Case 3 : pRow.Item("利用意向内容") = "所有者代理事業"
                Case 4 : pRow.Item("利用意向内容") = "権利設定または転移"
                Case 5 : pRow.Item("利用意向内容") = "その他"
                Case Else
                    pRow.Item("利用意向内容") = "自ら耕作"
            End Select
            Select Case Val(pRow.Item("利用意向根拠条項").ToString)
                Case 1 : pRow.Item("利用意向") = "農地法第32条第１項"
                Case 2 : pRow.Item("利用意向") = "農地法第32条第４項"
                Case 3 : pRow.Item("利用意向") = "農地法第33条第１項"
                Case Else : pRow.Item("利用意向") = "-"
            End Select
        Next

        Dim TBL利用状況調査 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT IIf([利用状況調査荒廃]=0,'-',IIf([利用状況調査荒廃]=1,'A分類','B分類')) AS 利用状況調査, Sum([D:農地Info].登記簿面積) AS 登記簿面積の合計, Sum([D:農地Info].実面積) AS 実面積の合計, Sum([D:農地Info].田面積) AS 田面積の合計, Sum([D:農地Info].畑面積) AS 畑面積の合計, Sum([D:農地Info].樹園地) AS 樹園地の合計, Count([D:農地Info].ID) AS 筆数 FROM ([D:農地Info] INNER JOIN [D:世帯Info] ON [D:農地Info].所有世帯ID = [D:世帯Info].ID) INNER JOIN [D:個人Info] ON [D:世帯Info].世帯主ID = [D:個人Info].ID WHERE ((([D:農地Info].利用状況調査荒廃)>0) AND (([D:個人Info].住民区分)=0)) GROUP BY IIf([利用状況調査荒廃]=0,'-',IIf([利用状況調査荒廃]=1,'A分類','B分類'));")
        Dim TBL利用状況調査市外 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:農地Info].利用状況調査荒廃, Sum([D:農地Info].登記簿面積) AS 登記簿面積の合計, Sum([D:農地Info].実面積) AS 実面積の合計, Sum([D:農地Info].田面積) AS 田面積の合計, Sum([D:農地Info].畑面積) AS 畑面積の合計, Sum([D:農地Info].樹園地) AS 樹園地の合計, Count([D:農地Info].ID) AS 筆数, V_住民区分.名称 AS 所有者住民区分 FROM ([D:農地Info] INNER JOIN [D:個人Info] AS [D:個人Info_1] ON [D:農地Info].所有者ID = [D:個人Info_1].ID) INNER JOIN V_住民区分 ON [D:個人Info_1].住民区分 = V_住民区分.ID GROUP BY [D:農地Info].利用状況調査荒廃, V_住民区分.名称 HAVING ((([D:農地Info].利用状況調査荒廃)>0));")

        mvarG利用状況.SetDataView(TBL利用状況調査, "", "")
        mvarG利用状況市外.SetDataView(TBL利用状況調査市外, "", "")

        mvarG農地法32の1.SetDataView(TBL利用意向調査, "[利用意向] = '農地法第32条第１項'", "")
        mvarG農地法32の4.SetDataView(TBL利用意向調査, "[利用意向] = '農地法第32条第４項'", "")
        mvarG農地法33の1.SetDataView(TBL利用意向調査, "[利用意向] = '農地法第33条第１項'", "")
        mvarG根拠未入力.SetDataView(TBL利用意向調査, "[利用意向] = '-'", "")

        With mvarG利用状況
            .AllowUserToAddRows = False
        End With
        With mvarTab利用状況
            .Text = "利用状況調査"
        End With
        With mvarTab利用意向
            .Text = "利用意向調査"
        End With

    End Sub

    Private Function SetPanel(sTitle As String, pGrid As HimTools2012.controls.DataGridViewWithDataView) As ToolStripContainer
        Dim pTs As New HimTools2012.controls.ToolStripContainerEX(pGrid, True)
        pTs.ToolBar.Items.Add(New ToolStripLabel(sTitle))
        pGrid.AllowUserToAddRows = False
        AddHandler pTs.ToolBar.Items.Add("エクセルへ").Click, AddressOf pGrid.ToExcel
        Return pTs
    End Function

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.NoMessage
        End Get
    End Property

 
End Class

