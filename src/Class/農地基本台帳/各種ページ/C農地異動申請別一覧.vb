Imports HimTools2012.CommonFunc

Public Class C農地異動申請別一覧
    Inherits HimTools2012.TabPages.NListSK

    Public mvarStart As ToolStripDateTimePickerWithlabel
    Public mvarEnd As ToolStripDateTimePickerWithlabel
    Public WithEvents mvar検索開始 As ToolStripButton

    'Private TBL個人 As DataTable
    'Private TBL農地 As DataTable
    'Private TBL転用農地 As DataTable
    'Private TBL削除農地 As DataTable

    Public Sub New()
        MyBase.New(True, "農地異動申請別一覧", "農地異動申請別一覧", ObjectMan, SysAD.ImageList48, True)
        mvarStart = New ToolStripDateTimePickerWithlabel("対象許可年月日")
        mvarEnd = New ToolStripDateTimePickerWithlabel("～")
        mvarStart.Value = Now.Date
        mvarEnd.Value = Now.Date

        mvar検索開始 = New ToolStripButton("検索開始")

        CreatemvarGrid()

        Me.ControlPanel.Add(mvarGrid)
        Me.ToolStrip.Items.AddRange({mvarStart, mvarEnd, mvar検索開始, New ToolStripSeparator})

        'TBL個人 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:個人Info].ID, [D:個人Info].性別 FROM [D:個人Info];")
        'App農地基本台帳.TBL個人.MergePlus(TBL個人)
        'TBL個人.PrimaryKey = New DataColumn() {TBL個人.Columns("ID")}
        'TBL農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:農地Info].ID, [D:農地Info].大字ID, [D:農地Info].小字ID, [D:農地Info].地番, [D:農地Info].登記簿地目, [D:農地Info].登記簿面積, [D:農地Info].現況地目, [D:農地Info].実面積 FROM [D:農地Info]")
        'App農地基本台帳.TBL農地.MergePlus(TBL農地)
        'TBL農地.PrimaryKey = New DataColumn() {TBL農地.Columns("ID")}
        'TBL転用農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D_転用農地].ID, [D_転用農地].大字ID, [D_転用農地].小字ID, [D_転用農地].地番, [D_転用農地].登記簿地目, [D_転用農地].登記簿面積, [D_転用農地].現況地目, [D_転用農地].実面積 FROM [D_転用農地]")
        'App農地基本台帳.TBL転用農地.MergePlus(TBL転用農地)
        'TBL転用農地.PrimaryKey = New DataColumn() {TBL転用農地.Columns("ID")}
        'TBL削除農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D_削除農地].ID, [D_削除農地].大字ID, [D_削除農地].小字ID, [D_削除農地].地番, [D_削除農地].登記簿地目, [D_削除農地].登記簿面積, [D_削除農地].現況地目, [D_削除農地].実面積 FROM [D_削除農地]")
        'App農地基本台帳.TBL削除農地.MergePlus(TBL削除農地)
        'TBL削除農地.PrimaryKey = New DataColumn() {TBL削除農地.Columns("ID")}
    End Sub

    Private Sub CreatemvarGrid()
        AddCol(mvarGrid, "法令", "", New DataGridViewTextBoxColumn)
        AddCol(mvarGrid, "許可年月日", "", New DataGridViewTextBoxColumn)
        AddCol(mvarGrid, "許可番号", "", New DataGridViewTextBoxColumn)
        AddCol(mvarGrid, "名称", "", New DataGridViewTextBoxColumn)
        AddCol(mvarGrid, "所有権移転", "", New DataGridViewTextBoxColumn)
        AddCol(mvarGrid, "権利の種類", "", New DataGridViewTextBoxColumn)
        AddCol(mvarGrid, "個人法人の別_出し手", "", New DataGridViewTextBoxColumn)
        AddCol(mvarGrid, "小作料", "対価/小作料", New DataGridViewTextBoxColumn)
        AddCol(mvarGrid, "小作料単位", "単位", New DataGridViewTextBoxColumn)
        AddCol(mvarGrid, "対象筆数", "同一申請時の筆数", New DataGridViewTextBoxColumn)
        AddCol(mvarGrid, " ", "", New DataGridViewTextBoxColumn)
        AddCol(mvarGrid, "筆大字", "", New DataGridViewTextBoxColumn)
        AddCol(mvarGrid, "筆小字", "", New DataGridViewTextBoxColumn)
        AddCol(mvarGrid, "筆地番", "", New DataGridViewTextBoxColumn)
        AddCol(mvarGrid, "筆登記地目", "", New DataGridViewTextBoxColumn)
        AddCol(mvarGrid, "筆登記面積", "", New DataGridViewTextBoxColumn)
        AddCol(mvarGrid, "筆現況地目", "", New DataGridViewTextBoxColumn)
        AddCol(mvarGrid, "筆現況面積", "", New DataGridViewTextBoxColumn)
        AddCol(mvarGrid, "所有者ID", "", New DataGridViewTextBoxColumn)

        Dim pKeyColumn As New DataGridViewTextBoxColumn
        pKeyColumn.DataPropertyName = "Key"
        pKeyColumn.Name = "Key"
        pKeyColumn.Visible = False
        mvarGrid.Columns.Add(pKeyColumn)

        Dim pIconColumn As New DataGridViewTextBoxColumn
        pIconColumn.DataPropertyName = "アイコン"
        pIconColumn.Name = "アイコン"
        pIconColumn.Visible = False
        mvarGrid.Columns.Add(pIconColumn)
    End Sub

    Private Sub AddCol(ByRef pDGrid As HimTools2012.controls.DataGridViewWithDataView, ByVal sName As String, ByVal sHeaderName As String, ByVal pCol As DataGridViewColumn)
        pCol.HeaderText = IIf(sHeaderName = "", sName, sHeaderName)
        pCol.DataPropertyName = sName
        pCol.Name = sName
        pDGrid.Columns.Add(pCol)
    End Sub

    Private DSet As HimTools2012.Data.DataSetEx
    Private mvarTable As DataTable
    Private mvarTBL所移 As DataTable
    Private mvarTBL形態 As DataTable
    Private mvarMaster大字 As DataTable
    Private mvarMaster小字 As DataTable
    Private mvarMaster地目 As DataTable
    Private mvarMaster現況地目 As DataTable
    Private mvarOutPutTable As DataTable
    Private mvarTotalTable As DataTable

    Private Total対象筆数 As Integer = 0

    Private Sub mvar検索開始_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar検索開始.Click
        DSet = New HimTools2012.Data.DataSetEx
        If mvarStart.Value < mvarEnd.Value Then
            Dim sWhere As String = String.Format("[法令] IN (30,31,60,61,62) AND [状態]=2 AND [許可年月日]>=#{0}/{1}/{2}# AND [許可年月日]<=#{3}/{4}/{5}# AND [許可番号] IS NOT NULL AND [農地リスト] IS NOT NULL",
                                                mvarStart.Value.Month, mvarStart.Value.Day, mvarStart.Value.Year,
                                                mvarEnd.Value.Month, mvarEnd.Value.Day, mvarEnd.Value.Year)

            mvarTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D_申請] WHERE " & sWhere & " ORDER BY [法令],[許可番号]")

            mvarTBL所移 = App農地基本台帳.GetMasterView("所有権移転の種類").ToTable
            mvarTBL所移.TableName = "所有権移転マスタ"
            mvarTBL形態 = App農地基本台帳.GetMasterView("小作形態").ToTable
            mvarTBL形態.TableName = "小作形態マスタ"
            mvarMaster大字 = App農地基本台帳.GetMasterView("大字").ToTable
            mvarMaster大字.PrimaryKey = New DataColumn() {mvarMaster大字.Columns("ID")}
            mvarMaster小字 = App農地基本台帳.GetMasterView("小字").ToTable
            mvarMaster小字.PrimaryKey = New DataColumn() {mvarMaster小字.Columns("ID")}
            mvarMaster地目 = App農地基本台帳.GetMasterView("地目").ToTable
            mvarMaster地目.PrimaryKey = New DataColumn() {mvarMaster地目.Columns("ID")}
            mvarMaster現況地目 = App農地基本台帳.GetMasterView("課税地目").ToTable
            mvarMaster現況地目.PrimaryKey = New DataColumn() {mvarMaster現況地目.Columns("ID")}
            DSet.Tables.AddRange({mvarTable, mvarTBL所移, mvarTBL形態})

            If mvarTable.Columns.Contains("所有権移転") Then
            Else
                DSet.Relations.Add(New DataRelation("所有権移転", mvarTBL所移.Columns("ID"), mvarTable.Columns("所有権移転の種類"), False))
                mvarTable.Columns.Add(New DataColumn("所有権移転", GetType(String), "Parent(所有権移転).名称"))
            End If

            If mvarTable.Columns.Contains("権利の種類") Then
            Else
                DSet.Relations.Add(New DataRelation("権利の種類", mvarTBL形態.Columns("ID"), mvarTable.Columns("権利種類"), False))
                mvarTable.Columns.Add(New DataColumn("権利の種類", GetType(String), "Parent(権利の種類).名称"))
            End If

            CreatemvarOutPutTable(mvarOutPutTable)
            CreatemvarOutPutTable(mvarTotalTable)

            Dim nList As New List(Of Long)
            Dim sSource As New System.Text.StringBuilder

            For Each pRow As DataRow In mvarTable.Rows
                Dim Ar() As String = Split(pRow.Item("農地リスト").ToString, ";")
                For Each sKey As String In Ar
                    Dim nID As Long = GetKeyCode(sKey)
                    If sKey.Length > 0 AndAlso Not nList.Contains(nID) Then
                        nList.Add(nID)
                        sSource.Append(nID.ToString() & ",")
                        If sSource.Length > 1024 Then
                            Dim sList As String = HimTools2012.StringF.Left(sSource.ToString(), sSource.Length - 1)
                            Dim pTBL1 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(
                                "SELECT * FROM [D:農地Info] WHERE [ID] IN ({0})",
                                 sList
                            )
                            App農地基本台帳.TBL農地.MergePlus(pTBL1)
                            Dim pTBL2 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(
                                "SELECT * FROM [D_転用農地] WHERE [ID] IN ({0})",
                                 sList
                            )
                            App農地基本台帳.TBL転用農地.MergePlus(pTBL2)
                            Dim pTBL3 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(
                                "SELECT * FROM [D_削除農地] WHERE [ID] IN ({0})",
                                 sList
                            )
                            App農地基本台帳.TBL削除農地.MergePlus(pTBL3)

                            sSource.Clear()
                        End If
                    End If
                Next
            Next
            If sSource.Length > 0 Then
                Dim sList As String = HimTools2012.StringF.Left(sSource.ToString(), sSource.Length - 1)
                Dim pTBL1 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(
                    "SELECT * FROM [D:農地Info] WHERE [ID] IN ({0})",
                    sList
                )
                App農地基本台帳.TBL農地.MergePlus(pTBL1)
                Dim pTBL2 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(
                    "SELECT * FROM [D_転用農地] WHERE [ID] IN ({0})",
                    sList
                )
                App農地基本台帳.TBL転用農地.MergePlus(pTBL2)
                Dim pTBL3 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(
                    "SELECT * FROM [D_削除農地] WHERE [ID] IN ({0})",
                    sList
                )
                App農地基本台帳.TBL削除農地.MergePlus(pTBL3)
            End If

            For Each pRow As DataRow In mvarTable.Rows
                Dim Ar() As String = Split(pRow.Item("農地リスト").ToString, ";")

                Total対象筆数 = 0

                For Each sKey As String In Ar
                    If sKey.Length > 0 Then
                        Dim 農地ID As Long = GetKeyCode(sKey)
                        Dim pFindRow As DataRow = App農地基本台帳.TBL農地.Rows.Find(農地ID)

                        If Not pFindRow Is Nothing Then : Total対象筆数 += 1
                        Else
                            pFindRow = App農地基本台帳.TBL転用農地.Rows.Find(農地ID)
                            If Not pFindRow Is Nothing Then : Total対象筆数 += 1
                            Else
                                pFindRow = App農地基本台帳.TBL削除農地.Rows.Find(農地ID)
                                If Not pFindRow Is Nothing Then : Total対象筆数 += 1
                                End If
                            End If
                        End If
                    End If
                Next

                For Each sKey As String In Ar
                    If sKey.Length > 0 Then
                        Dim 農地ID As Long = GetKeyCode(sKey)

                        Dim pFindRow As DataRow = App農地基本台帳.TBL農地.FindRowByID(農地ID)
                        If Not pFindRow Is Nothing Then
                            AddRowOutPutTable(pRow, pFindRow, "農地." & 農地ID)
                        Else
                            pFindRow = App農地基本台帳.TBL転用農地.Rows.Find(農地ID)
                            If Not pFindRow Is Nothing Then : AddRowOutPutTable(pRow, pFindRow, "転用農地." & 農地ID)
                            Else
                                pFindRow = App農地基本台帳.TBL削除農地.Rows.Find(農地ID)
                                If Not pFindRow Is Nothing Then : AddRowOutPutTable(pRow, pFindRow, "削除農地." & 農地ID)
                                End If
                            End If
                        End If
                    End If
                Next
            Next

            mvarOutPutTable.Merge(mvarTotalTable)

            mvarGrid.SetDataView(mvarOutPutTable, "", "[法令], [許可年月日], [許可番号]")
        End If
    End Sub

    Private Sub AddRowOutPutTable(ByRef pRow As DataRow, ByRef pFindRow As DataRow, ByVal pKey As String)
        Dim AddRow As DataRow = mvarOutPutTable.NewRow
        AddRow.Item("Key") = pKey
        AddRow.Item("法令") = pRow.Item("法令")
        AddRow.Item("許可年月日") = pRow.Item("許可年月日")
        AddRow.Item("許可番号") = pRow.Item("許可番号")
        AddRow.Item("名称") = pRow.Item("名称")
        AddRow.Item("所有権移転") = pRow.Item("所有権移転")
        AddRow.Item("権利の種類") = pRow.Item("権利の種類")

        Dim pFind農家 As DataRow = App農地基本台帳.TBL個人.FindRowByID(pRow.Item("申請者A"))
        If Not pFind農家 Is Nothing Then
            If Val(pFind農家.Item("性別").ToString) = 3 Then
                AddRow.Item("個人法人の別_出し手") = "法人"
            Else
                AddRow.Item("個人法人の別_出し手") = "個人"
            End If
        Else
            AddRow.Item("個人法人の別_出し手") = "個人"
        End If
        AddRow.Item("小作料") = pRow.Item("小作料")
        AddRow.Item("小作料単位") = pRow.Item("小作料単位")
        AddRow.Item("対象筆数") = Total対象筆数

        Dim p大字 As DataRow = mvarMaster大字.Rows.Find(pFindRow.Item("大字ID"))
        AddRow.Item("筆大字") = FindValue(p大字)
        Dim p小字 As DataRow = mvarMaster小字.Rows.Find(pFindRow.Item("小字ID"))
        AddRow.Item("筆小字") = FindValue(p小字)
        AddRow.Item("筆地番") = pFindRow.Item("地番")
        Dim p地目 As DataRow = mvarMaster地目.Rows.Find(pFindRow.Item("登記簿地目"))
        AddRow.Item("筆登記地目") = FindValue(p地目)
        AddRow.Item("筆登記面積") = pFindRow.Item("登記簿面積")
        Dim p現況地目 As DataRow = mvarMaster地目.Rows.Find(pFindRow.Item("現況地目"))
        AddRow.Item("筆現況地目") = FindValue(p現況地目)
        AddRow.Item("筆現況面積") = pFindRow.Item("実面積")

        AddRow.Item("所有者ID") = pRow.Item("申請者A")

        AddRow.Item("アイコン") = IIf(InStr(AddRow.Item("筆現況地目").ToString, "田") > 0, "田", "畑")

        mvarOutPutTable.Rows.Add(AddRow)
    End Sub

    Private Function FindValue(ByVal pRow As DataRow)
        Dim pValue As Object = Nothing
        If pRow Is Nothing Then
            pValue = ""
        Else
            pValue = pRow.Item("名称")
        End If

        Return pValue
    End Function

    Private Sub CreatemvarOutPutTable(ByRef pTable As DataTable)
        pTable = New DataTable
        With pTable
            .Columns.Add(New DataColumn("Key", GetType(String)))
            .Columns.Add(New DataColumn("アイコン", GetType(String)))
            .Columns.Add(New DataColumn("法令", GetType(Integer)))
            .Columns.Add(New DataColumn("許可年月日", GetType(Date)))
            .Columns.Add(New DataColumn("許可番号", GetType(Integer)))
            .Columns.Add(New DataColumn("名称", GetType(String)))
            .Columns.Add(New DataColumn("所有権移転", GetType(String)))
            .Columns.Add(New DataColumn("権利の種類", GetType(String)))
            .Columns.Add(New DataColumn("個人法人の別_出し手", GetType(String)))
            .Columns.Add(New DataColumn("小作料", GetType(String)))
            .Columns.Add(New DataColumn("小作料単位", GetType(String)))
            .Columns.Add(New DataColumn("対象筆数", GetType(Integer)))
            .Columns.Add(New DataColumn(" ", GetType(Integer)))
            .Columns.Add(New DataColumn("筆大字", GetType(String)))
            .Columns.Add(New DataColumn("筆小字", GetType(String)))
            .Columns.Add(New DataColumn("筆地番", GetType(String)))
            .Columns.Add(New DataColumn("筆登記地目", GetType(String)))
            .Columns.Add(New DataColumn("筆登記面積", GetType(Decimal)))
            .Columns.Add(New DataColumn("筆現況地目", GetType(String)))
            .Columns.Add(New DataColumn("筆現況面積", GetType(Decimal)))
            .Columns.Add(New DataColumn("所有者ID", GetType(Decimal)))
        End With
    End Sub

    Private Sub SetTotalTable(ByRef pTBL As DataTable, ByRef pRefTBL As DataTable)
        Dim s法令 As Integer = 0
        Dim s許可年月日 As DateTime = Nothing
        Dim s許可番号 As Integer = 0

        Dim s名称 As String = ""
        Dim s所有権移転 As String = ""
        Dim s権利の種類 As String = ""
        Dim s個人法人の別 As String = ""
        Dim s小作料 As String = ""
        Dim s小作料単位 As String = ""
        Dim s対象筆数 As Integer = 0

        Dim s登記面積 As Decimal = 0
        Dim s現況面積 As Decimal = 0

        For Each pRow As DataRow In pRefTBL.Rows
            If s法令 = 0 Or pRow.Item("法令") = s法令 AndAlso pRow.Item("許可年月日") = s許可年月日 AndAlso pRow.Item("許可番号") = s許可番号 Then
                s法令 = pRow.Item("法令")
                s許可年月日 = pRow.Item("許可年月日")
                s許可番号 = pRow.Item("許可番号")
                s名称 = pRow.Item("名称")
                s所有権移転 = pRow.Item("所有権移転").ToString
                s権利の種類 = pRow.Item("権利の種類").ToString
                s個人法人の別 = pRow.Item("個人法人の別_出し手")
                s小作料 = pRow.Item("小作料").ToString
                s小作料単位 = pRow.Item("小作料単位").ToString
                s対象筆数 += 1

                s登記面積 += pRow.Item("筆登記面積")
                s現況面積 += pRow.Item("筆現況面積")
            Else
                Dim pTRow As DataRow = pTBL.NewRow
                With pTRow
                    .Item("法令") = s法令
                    .Item("許可年月日") = s許可年月日
                    .Item("許可番号") = s許可番号
                    .Item("筆ID") = 999999999
                    .Item("名称") = s名称
                    .Item("所有権移転") = s所有権移転
                    .Item("権利の種類") = s権利の種類
                    .Item("個人法人の別_出し手") = s個人法人の別
                    .Item("小作料") = s小作料
                    .Item("小作料単位") = s小作料単位
                    .Item("対象筆数") = s対象筆数
                    .Item("筆登記面積") = s登記面積
                    .Item("筆現況面積") = s現況面積
                    pTBL.Rows.Add(pTRow)

                    s法令 = pRow.Item("法令")
                    s許可年月日 = pRow.Item("許可年月日")
                    s許可番号 = pRow.Item("許可番号")
                    s対象筆数 = 0
                    s登記面積 = pRow.Item("筆登記面積")
                    s現況面積 = pRow.Item("筆現況面積")
                End With
            End If
        Next

        Dim pLRow As DataRow = pTBL.NewRow
        With pLRow
            .Item("法令") = s法令
            .Item("許可年月日") = s許可年月日
            .Item("許可番号") = s許可番号
            .Item("筆ID") = 999999999
            .Item("名称") = s名称
            .Item("所有権移転") = s所有権移転
            .Item("権利の種類") = s権利の種類
            .Item("個人法人の別_出し手") = s個人法人の別
            .Item("小作料") = s小作料
            .Item("小作料単位") = s小作料単位
            .Item("対象筆数") = s対象筆数
            .Item("筆登記面積") = s登記面積
            .Item("筆現況面積") = s現況面積
            pTBL.Rows.Add(pLRow)
        End With
    End Sub

    Public Overrides ReadOnly Property 検索Page As HimTools2012.TabPages.CPage検索SK
        Get
            Return Nothing
        End Get
    End Property

    Public Overrides Sub 検索開始(ByVal sWhere As String, ByVal sViewWhere As String, Optional ByVal sOrderBy As String = "", Optional ByVal sColumnStyle As String = "")

    End Sub
End Class
