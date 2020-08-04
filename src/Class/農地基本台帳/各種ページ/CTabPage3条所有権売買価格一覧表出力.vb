'20160411霧島

Public Class CTabPage3条所有権売買価格一覧表出力
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Public mvarStart As ToolStripDateTimePickerWithlabel
    Public mvarEnd As ToolStripDateTimePickerWithlabel
    Public WithEvents mvar検索開始 As ToolStripButton
    Public WithEvents mvarExcel出力 As ToolStripButton

    Private mvarGrid As HimTools2012.controls.DataGridViewWithDataView
    Private TBL農地 As DataTable
    Private TBL転用農地 As DataTable
    Private TBL削除農地 As DataTable

    Public Sub New()
        MyBase.New(True, True, "所有権売買価格一覧表", "所有権売買価格一覧表")

        mvarStart = New ToolStripDateTimePickerWithlabel("対象許可年月日")
        mvarEnd = New ToolStripDateTimePickerWithlabel("～")
        mvarStart.Value = Now.Date
        mvarEnd.Value = Now.Date

        mvar検索開始 = New ToolStripButton("検索開始")
        mvarExcel出力 = New ToolStripButton("Excel出力")

        CreatemvarGrid(mvarGrid)

        Me.ControlPanel.Add(mvarGrid)
        Me.ToolStrip.Items.AddRange({mvarStart, mvarEnd, mvar検索開始, New ToolStripSeparator, mvarExcel出力})

        TBL農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:農地Info].ID, [D:農地Info].大字ID, [D:農地Info].小字ID, [D:農地Info].地番, [D:農地Info].登記簿地目, [D:農地Info].登記簿面積, [D:農地Info].実面積 FROM [D:農地Info]")
        'App農地基本台帳.TBL農地.MergePlus(TBL農地)
        TBL農地.PrimaryKey = New DataColumn() {TBL農地.Columns("ID")}
        TBL転用農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D_転用農地].ID, [D_転用農地].大字ID, [D_転用農地].小字ID, [D_転用農地].地番, [D_転用農地].登記簿地目, [D_転用農地].登記簿面積, [D_転用農地].実面積 FROM [D_転用農地]")
        'App農地基本台帳.TBL転用農地.MergePlus(TBL転用農地)
        TBL転用農地.PrimaryKey = New DataColumn() {TBL転用農地.Columns("ID")}
        TBL削除農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D_削除農地].ID, [D_削除農地].大字ID, [D_削除農地].小字ID, [D_削除農地].地番, [D_削除農地].登記簿地目, [D_削除農地].登記簿面積, [D_削除農地].実面積 FROM [D_削除農地]")
        'App農地基本台帳.TBL削除農地.MergePlus(TBL削除農地)
        TBL削除農地.PrimaryKey = New DataColumn() {TBL削除農地.Columns("ID")}
    End Sub

    Private DSet As HimTools2012.Data.DataSetEx
    Private mvarTable As DataTable
    Private mvarMasterTable As DataTable
    Private mvarMaster大字 As DataTable
    Private mvarMaster小字 As DataTable
    Private mvarMaster地目 As DataTable
    Private mvarOutPutTable As DataTable

    Private Total登記面積 As Decimal = 0
    Private Total実面積 As Decimal = 0
    Private Total対象筆数 As Integer = 0

    Private Sub mvar検索開始_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar検索開始.Click
        DSet = New HimTools2012.Data.DataSetEx
        If mvarStart.Value < mvarEnd.Value Then
            Dim sWhere As String = String.Format("[法令] IN (30,50) AND [状態]=2 AND [許可年月日]>=#{0}/{1}/{2}# AND [許可年月日]<=#{3}/{4}/{5}# AND [農地リスト] IS NOT NULL",
                                                mvarStart.Value.Month, mvarStart.Value.Day, mvarStart.Value.Year,
                                                mvarEnd.Value.Month, mvarEnd.Value.Day, mvarEnd.Value.Year)

            mvarTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D_申請] WHERE " & sWhere & " ORDER BY [法令],[許可番号]")
            mvarMasterTable = App農地基本台帳.GetMasterView("所有権移転の種類").ToTable
            mvarMaster大字 = App農地基本台帳.GetMasterView("大字").ToTable
            mvarMaster大字.PrimaryKey = New DataColumn() {mvarMaster大字.Columns("ID")}
            mvarMaster小字 = App農地基本台帳.GetMasterView("小字").ToTable
            mvarMaster小字.PrimaryKey = New DataColumn() {mvarMaster小字.Columns("ID")}
            mvarMaster地目 = App農地基本台帳.GetMasterView("地目").ToTable
            mvarMaster地目.PrimaryKey = New DataColumn() {mvarMaster地目.Columns("ID")}
            DSet.Tables.AddRange({mvarTable, mvarMasterTable})

            If mvarTable.Columns.Contains("所有権移転") Then
            Else
                DSet.Relations.Add(New DataRelation("所有権移転", mvarMasterTable.Columns("ID"), mvarTable.Columns("所有権移転の種類"), False))
                mvarTable.Columns.Add(New DataColumn("所有権移転内容", GetType(String), "Parent(所有権移転).名称"))
            End If

            CreatemvarOutPutTable(mvarOutPutTable)
            For Each pRow As DataRow In mvarTable.Rows
                Dim Ar As Object = Split(pRow.Item("農地リスト").ToString, ";")

                'Dim Dic筆情報 As New Dictionary(Of String, String)
                Total登記面積 = 0
                Total実面積 = 0
                Total対象筆数 = 0
                For n As Integer = 0 To UBound(Ar)
                    Dim Ar農地 As Object = Split(Ar(n), ".")
                    Dim pFindRow As DataRow = TBL農地.Rows.Find(Ar農地(1))
                    If Not pFindRow Is Nothing Then : CountArea(pFindRow)
                    Else
                        pFindRow = TBL転用農地.Rows.Find(Ar農地(1))
                        If Not pFindRow Is Nothing Then : CountArea(pFindRow)
                        Else
                            pFindRow = TBL削除農地.Rows.Find(Ar農地(1))
                            If Not pFindRow Is Nothing Then : CountArea(pFindRow)
                            End If
                        End If
                    End If
                Next

                For n As Integer = 0 To UBound(Ar)
                    Dim Ar農地 As Object = Split(Ar(n), ".")
                    Dim pFindRow As DataRow = TBL農地.Rows.Find(Ar農地(1))
                    If Not pFindRow Is Nothing Then
                        AddRowOutPutTable(pRow, pFindRow)
                    Else
                        pFindRow = TBL転用農地.Rows.Find(Ar農地(1))
                        If Not pFindRow Is Nothing Then : AddRowOutPutTable(pRow, pFindRow)
                        Else
                            pFindRow = TBL削除農地.Rows.Find(Ar農地(1))
                            If Not pFindRow Is Nothing Then : AddRowOutPutTable(pRow, pFindRow)
                            End If
                        End If
                    End If
                Next
            Next

            mvarGrid.SetDataView(mvarOutPutTable, "[所有権移転] = '売買' And [小作料] > 0", "[法令], [許可年月日], [許可番号]")
        End If
    End Sub
    Private Sub CountArea(ByRef pFindRow As DataRow)
        Total登記面積 += pFindRow.Item("登記簿面積")
        Total実面積 += pFindRow.Item("実面積")
        Total対象筆数 += 1
    End Sub
    Private Sub AddRowOutPutTable(ByRef pRow As DataRow, ByRef pFindRow As DataRow)
        Dim AddRow As DataRow = mvarOutPutTable.NewRow
        AddRow.Item("法令") = pRow.Item("法令")
        AddRow.Item("許可年月日") = pRow.Item("許可年月日")
        AddRow.Item("許可番号") = pRow.Item("許可番号")
        AddRow.Item("名称") = pRow.Item("名称")
        AddRow.Item("所有権移転") = pRow.Item("所有権移転内容")
        AddRow.Item("小作料") = pRow.Item("小作料")
        AddRow.Item("小作料単位") = pRow.Item("小作料単位")
        AddRow.Item("登記面積") = Total登記面積
        AddRow.Item("現況面積") = Total実面積
        AddRow.Item("対象筆数") = Total対象筆数

        Dim p大字 As DataRow = mvarMaster大字.Rows.Find(pFindRow.Item("大字ID"))
        AddRow.Item("筆大字") = IIf(p大字 Is Nothing, "", p大字.Item("名称"))
        Dim p小字 As DataRow = mvarMaster小字.Rows.Find(pFindRow.Item("小字ID"))
        AddRow.Item("筆小字") = IIf(p小字 Is Nothing, "", p小字.Item("名称"))
        AddRow.Item("筆地番") = pFindRow.Item("地番")
        Dim p地目 As DataRow = mvarMaster地目.Rows.Find(pFindRow.Item("登記簿地目"))
        AddRow.Item("筆登記地目") = IIf(p地目 Is Nothing, "", p地目.Item("名称"))
        AddRow.Item("筆登記面積") = pFindRow.Item("登記簿面積")


        mvarOutPutTable.Rows.Add(AddRow)
    End Sub

    Private Sub CreatemvarGrid(ByRef pDGrid As HimTools2012.controls.DataGridViewWithDataView)
        pDGrid = New HimTools2012.controls.DataGridViewWithDataView
        AddCol(pDGrid, "法令", "法令", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "許可年月日", "許可年月日", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "許可番号", "許可番号", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "名称", "名称", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "所有権移転", "所有権移転", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "小作料", "小作料", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "小作料単位", "小作料単位", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "登記面積", "登記面積", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "現況面積", "現況面積", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "対象筆数", "対象筆数", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, " ", " ", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "筆大字", "筆大字", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "筆小字", "筆小字", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "筆地番", "筆地番", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "筆登記地目", "筆登記地目", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "筆登記面積", "筆登記面積", New DataGridViewTextBoxColumn)

        pDGrid.AutoGenerateColumns = False
        pDGrid.AllowUserToAddRows = False
        pDGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
    End Sub
    Private Sub AddCol(ByRef pDGrid As HimTools2012.controls.DataGridViewWithDataView, ByVal sHeader As String, ByVal sData As String, ByVal pCol As DataGridViewColumn)
        pCol.HeaderText = sHeader
        pCol.DataPropertyName = sData
        pDGrid.Columns.Add(pCol)
    End Sub

    Private Sub CreatemvarOutPutTable(ByRef pTable As DataTable)
        pTable = New DataTable
        pTable.Columns.Add(New DataColumn("法令", GetType(Integer)))
        pTable.Columns.Add(New DataColumn("許可年月日", GetType(Date)))
        pTable.Columns.Add(New DataColumn("許可番号", GetType(Integer)))
        pTable.Columns.Add(New DataColumn("名称", GetType(String)))
        pTable.Columns.Add(New DataColumn("所有権移転", GetType(String)))
        pTable.Columns.Add(New DataColumn("小作料", GetType(Decimal)))
        pTable.Columns.Add(New DataColumn("小作料単位", GetType(String)))
        pTable.Columns.Add(New DataColumn("登記面積", GetType(Decimal)))
        pTable.Columns.Add(New DataColumn("現況面積", GetType(Decimal)))
        pTable.Columns.Add(New DataColumn("対象筆数", GetType(Integer)))
        pTable.Columns.Add(New DataColumn(" ", GetType(Integer)))
        pTable.Columns.Add(New DataColumn("筆大字", GetType(String)))
        pTable.Columns.Add(New DataColumn("筆小字", GetType(String)))
        pTable.Columns.Add(New DataColumn("筆地番", GetType(String)))
        pTable.Columns.Add(New DataColumn("筆登記地目", GetType(String)))
        pTable.Columns.Add(New DataColumn("筆登記面積", GetType(Decimal)))
    End Sub

    Private Sub mvarExcel出力_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarExcel出力.Click
        mvarGrid.ToExcel()
    End Sub
End Class
