'20160411霧島
Imports HimTools2012
Imports HimTools2012.Excel.XMLSS2003

''' <summary>
'''　目次: \\172.29.11.1\システム配置\農地台帳\伊佐市
''' </summary>
''' <remarks></remarks>
Public Class C伊佐市
    Inherits C市町村別
    Public Sub New()
        MyBase.New("伊佐市")
    End Sub

    Public Overrides Function Get選挙世帯一覧() As System.Data.DataTable
        Return Nothing
    End Function
    Public Overrides ReadOnly Property 旧農振都市計画使用 As Boolean
        Get
            Return True
        End Get
    End Property
    Public Overrides Sub InitMenu(ByVal pMain As HimTools2012.SystemWindows.CMainPageSK)
        With pMain
            .ListView.Clear()
            .ListView.ItemAdd("農家検索", "農家検索", ImageKey.閲覧検索, "閲覧・検索", AddressOf 農家一覧)

            .ListView.ItemAdd("非農地通知", "非農地通知", ImageKey.閲覧検索, "閲覧・検索", AddressOf ClickMenu)
            .ListView.ItemAdd("非農地通知済み管理", "非農地通知済み管理", ImageKey.閲覧検索, "閲覧・検索", AddressOf ClickMenu)
            .ListView.ItemAdd("固定資産比較", "固定資産比較", "他システム連携", "操作", AddressOf 固定資産比較)

            .ListView.ItemAdd("CSVto農地", "CSVto農地", ImageKey.作業, "操作", AddressOf CSVto農地)

            .ListView.ItemAdd("現地調査表作成", "現地調査表作成", ImageKey.作業, "操作", AddressOf 現地調査)
            .ListView.ItemAdd("申請内容出力", "申請内容出力", ImageKey.作業, "操作", AddressOf 申請内容出力)
            .ListView.ItemAdd("総会資料作成", "総会資料作成", ImageKey.作業, "操作", AddressOf ClickMenu)

            .ListView.ItemAdd("受付公布簿", "伊佐市用受付公布簿", "印刷", "印刷", Sub(s, e) SysAD.MainForm.MainTabCtrl.ExistPage("受付公布簿", True, GetType(CTabPage受付公布簿)))
            .ListView.ItemAdd("自然解約の実行", "自然解約の実行", ImageKey.作業, "操作", AddressOf ClickMenu)
            .ListView.ItemAdd("所有権売買価格一覧表", "所有権売買価格一覧表", "印刷", "印刷", Sub(s, e) SysAD.MainForm.MainTabCtrl.ExistPage("所有権売買価格一覧表", True, GetType(CTabPage3条所有権売買価格一覧表出力)))

            .ListView.ItemAdd("重複農地検索.0", "重複農地検索", ImageKey.閲覧検索, "閲覧・検索", AddressOf ClickMenu)
            .ListView.ItemAdd("農地台帳一括印刷", "農地台帳一括印刷", "印刷", "印刷", AddressOf ClickMenu)
            .ListView.ItemAdd("農家一覧", "農家一覧", "印刷", "印刷", AddressOf ClickMenu)
            .ListView.ItemAdd("工事進歩状況一覧表", "工事進歩状況一覧表", "印刷", "印刷", AddressOf 工事進歩状況一覧表)

            .ListView.ItemAdd("利用権終期台帳", "利用権終期台帳", "集計一覧", "印刷", AddressOf sub利用権終期管理)

            .ListView.ItemAdd("固定読み込み", "固定読み込み", "他システム連携", "他システム連携", AddressOf ClickMenu)
            .ListView.ItemAdd("農業委員毎利用権実績", "農業委員毎利用権実績", "印刷", "印刷", AddressOf SUB農業委員毎利用権実績)

            .ListView.ItemAdd("フェーズ２利用状況調査・意向調査情報", "フェーズ２利用状況調査・意向調査情報", ImageKey.他システム連携, "他システム連携", AddressOf ClickMenu)
            MyBase.InitMenu(pMain)
        End With
    End Sub
    Public Sub SUB農業委員毎利用権実績()
        If Not SysAD.MainForm.MainTabCtrl.ExistPage("農業委員毎利用権実績", True, GetType(CTabPage農業委員毎利用権実績)) Then

        End If
    End Sub


    Private Sub 工事進歩状況一覧表()
        Dim sFile As String = SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\工事進歩状況一覧表.xml"

        If IO.File.Exists(sFile) Then
            With New dlgInputBWDate(SysAD.GetXMLProperty("工事進捗状況一覧", "一覧を作成する転用申請が許可された期間を入力してください。", Now.Date))
                If .ShowDialog = DialogResult.OK Then
                    Dim StartDate As DateTime = New DateTime(.StartDate.Year, .StartDate.Month, .StartDate.Day)
                    Dim EndDate As DateTime = New DateTime(.EndDate.Year, .EndDate.Month, .EndDate.Day)

                    Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(String.Format("SELECT D_申請.ID, D_申請.法令, D_申請.状態, D_申請.名称, D_申請.許可年月日, D_申請.許可番号, D_申請.農地リスト, D_申請.申請者A, D_申請.氏名A, D_申請.住所A, D_申請.申請者B, D_申請.氏名B, D_申請.住所B, D_申請.申請理由A, D_申請.所有権移転の種類, D_申請.完了報告年月日 FROM D_申請 WHERE (((D_申請.法令) In (40,50,51)) AND ((D_申請.状態)=2) AND ((D_申請.許可年月日)>=#{0}# And (D_申請.許可年月日)<=#{1}#) AND ((D_申請.完了報告年月日) Is Null Or (D_申請.完了報告年月日)<#1/1/2000#)) ORDER BY D_申請.許可年月日, D_申請.許可番号;", StartDate, EndDate))
                    Dim sXML As String = HimTools2012.TextAdapter.LoadTextFile(sFile)
                    Dim XMLSS As New CXMLSS2003(sXML)

                    With XMLSS.WorkBook.WorkSheets.Items("一覧表")
                        Dim LoopRows As New XMLLoopRows(._object)
                        Dim nLoop As Integer = -1

                        For Each pRow As DataRow In pTBL.Rows
                            If nLoop = -1 Then

                            Else
                                For Each pXRow As XMLSSRow In LoopRows
                                    Dim pCopyRow = pXRow.CopyRow


                                    .Table.Rows.InsertRow(LoopRows.InsetRow, pCopyRow)
                                    LoopRows.InsetRow += 1
                                Next
                            End If

                            nLoop += 1
                            Dim 申請者ID As Decimal = 0

                            Select Case Val(pRow("法令").ToString)
                                Case enum法令.農地法4条
                                    .ValueReplace("{転用事業者氏名}", pRow("氏名A").ToString)
                                    .ValueReplace("{転用事業者住所}", pRow("住所A").ToString)
                                    申請者ID = Val(pRow.Item("申請者A").ToString)
                                Case enum法令.農地法5条所有権, enum法令.農地法5条貸借
                                    .ValueReplace("{転用事業者氏名}", pRow("氏名B").ToString)
                                    .ValueReplace("{転用事業者住所}", pRow("住所B").ToString)
                                    申請者ID = Val(pRow.Item("申請者B").ToString)
                                Case enum法令.事業計画変更
                                    .ValueReplace("{転用事業者氏名}", pRow("氏名C").ToString)
                                    .ValueReplace("{転用事業者住所}", pRow("住所C").ToString)
                                    申請者ID = Val(pRow.Item("申請者C").ToString)
                            End Select

                            .ValueReplace("{No}", "" & (nLoop + 1).ToString)
                            .ValueReplace("{許可年月日}", Replace(CDate(pRow.Item("許可年月日")).ToString("yyyy.M.d"), "/", "."))
                            .ValueReplace("{許可番号}", pRow.Item("許可番号").ToString)

                            .ValueReplace("{転用目的}", pRow("申請理由A").ToString)

                            Dim pPRow = App農地基本台帳.TBL個人.FindRowByID(CommonFunc.GetKeyCode(申請者ID))

                            If pPRow IsNot Nothing Then
                                .ValueReplace("{転用事業者電話番号}", pPRow.Item("電話番号").ToString)
                                .ValueReplace("{転用事業者集落名}", pPRow.Item("行政区名").ToString)
                            End If


                            Dim 転用許可地 As String = ""
                            Dim 転用農地Count As Integer = -1
                            Dim 転用面積_田 As Decimal = 0
                            Dim 転用面積_畑 As Decimal = 0

                            Dim St農地 As String = pRow.Item("農地リスト").ToString
                            For Each s農地Key As String In Split(St農地, ";")
                                If s農地Key.Length > 0 Then
                                    Dim pNRow As DataRow = Nothing
                                    Select Case CommonFunc.GetKeyHead(s農地Key)
                                        Case "農地"
                                            pNRow = App農地基本台帳.TBL農地.FindRowByID(CommonFunc.GetKeyCode(s農地Key))
                                        Case "転用農地"
                                            pNRow = App農地基本台帳.TBL転用農地.FindRowByID(CommonFunc.GetKeyCode(s農地Key))
                                    End Select

                                    If pNRow IsNot Nothing Then
                                        If 転用許可地.Length = 0 Then
                                            転用許可地 = pNRow.Item("土地所在")
                                        End If
                                        If Val(pNRow.Item("田面積").ToString) > 0 OrElse pNRow.Item("登記簿地目名") = "田" Then
                                            転用面積_田 += pNRow.Item("実面積")
                                        ElseIf Val(pNRow.Item("畑面積").ToString) > 0 OrElse pNRow.Item("登記簿地目名") = "畑" Then
                                            転用面積_畑 += pNRow.Item("実面積")
                                        End If
                                    End If

                                    転用農地Count += 1
                                End If
                                If 転用農地Count > 0 Then
                                    転用許可地 &= String.Format(" 外{0}筆", 転用農地Count)
                                End If
                            Next
                            .ValueReplace("{土地の所在}", 転用許可地)
                            .ValueReplace("{田面積計}", String.Format("{0,8}", 転用面積_田))
                            .ValueReplace("{畑面積計}", String.Format("{0,8}", 転用面積_畑))
                        Next

                        Dim SavePath As String = SysAD.OutputFolder & String.Format("\工事進捗状況報告書.xml")

                        HimTools2012.TextAdapter.SaveTextFile(SavePath, XMLSS.OutPut(True))
                        Using pExcel As New HimTools2012.Excel.Automation.ExcelAutomation
                            pExcel.ShowPreview(SavePath)
                        End Using
                    End With
                End If
            End With
        End If
    End Sub


    Private Sub 申請内容出力()
        If Not SysAD.MainForm.MainTabCtrl.ExistPage("申請内容", True, GetType(CTabPage申請内容)) Then

        End If
    End Sub

    Public Overrides Sub InitLocalData()
        Dim sPath As String = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\農地台帳LocaloData"
        If Not IO.Directory.Exists(sPath) Then
            'MsgBox("=>" & sPath & "を作成します。")
            Try
                IO.Directory.CreateDirectory(sPath)
            Catch ex As Exception
                MsgBox("失敗しました:" & ex.Message)
            End Try
        End If
        SysAD.SystemInfo.LocalDataPath = sPath
        SysAD.SystemInfo.XMLDataPath = sPath

#If DEBUG Then
        sub農地期間満了の終了()
#End If

        Dim FilePath = SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\AutoSQL.txt"


        If IO.File.Exists(FilePath) Then
            MsgBox("データベースの加工/修復を行います。")
            Dim sText As String = HimTools2012.TextAdapter.LoadTextFile(FilePath)
            Dim sLines As String() = Split(sText, vbCrLf)

            With SysAD.DB(sLRDB)
                For Each sLine As String In sLines
                    .ExecuteSQL(sLine)
                Next
            End With
            System.IO.File.Delete(FilePath)
            MsgBox("終了しました。")
        End If

    End Sub
    Public Overrides ReadOnly Property 市町村別現況地目CD(ByVal nType As C市町村別.地目Type) As Integer()
        Get
            Select Case nType
                Case 地目Type.田地目 : Return {10, 11}
                Case 地目Type.畑地目 : Return {20, 21}
                Case 地目Type.農地地目 : Return {10, 11, 20, 21}
                Case 地目Type.その他地目
                    Return MyBase.Make市町村別現況地目コード(nType)
                Case Else
                    Return MyBase.市町村別現況地目CD(nType)
            End Select

            Return MyBase.市町村別現況地目CD(nType)
        End Get
    End Property

End Class

Public Class CTabPage申請内容
    Inherits HimTools2012.TabPages.CTabPageWithDataGridView

    Private mvarTable As DataTable
    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.NoMessage
        End Get
    End Property

    Private mvar検索開始 As New ToolStripDateTimePickerWithlabel("検索範囲")
    Private mvar検索終了 As New ToolStripDateTimePickerWithlabel("～")
    Private WithEvents mvarBtn検索開始 As New ToolStripButton("検索開始")
    Private mvarDSet As HimTools2012.Data.DataSetEx

    Public Sub New()
        MyBase.New(True, "申請内容", "申請内容", ObjectMan, False)
        mvarGrid.ReadOnly = True
        mvar検索開始.Value = DateSerial(Now.Year, Now.Month, 1).AddMonths(IIf(Now.Day > 15, 0, -1))
        mvar検索終了.Value = DateSerial(Now.Year, mvar検索開始.Value.Month, HimTools2012.DateFunctions.GetMaxDay(mvar検索開始.Value.Year, mvar検索開始.Value.Month, 31))

        AddColumn("調査年月日", "調査日")
        AddColumn("調査員1名称", "調査員１名称")
        AddColumn("調査員2名称", "調査員１名称")
        AddColumn("調査員3名称", "調査員１名称")
        AddColumn("法令名称", "法令名称")
        AddColumn("権利名称", "権利名称")
        AddColumn("設定名称", "設定名称")
        AddColumn("農地区分名称", "農地区分名称")
        AddColumn("氏名", "氏名")
        AddColumn("行政書士名", "受任者")
        AddColumn("申請地", "申請地")
        ',,,,

        AddColumn("目的", "目的")
        ',集合時間場所

        AddColumn("農業委員1", "調査員１コード")
        AddColumn("農業委員2", "調査員２コード")
        AddColumn("農業委員3", "調査員３コード")
        ',支所コード
        AddColumn("法令", "法令コード")
        AddColumn("権利コード", "権利コード")
        AddColumn("所有権移転の種類", "設定コード")
        AddColumn("農地区分", "農地区分コード")
        ',農地番号,土地番号,市町村コード,区コード,町＋字丁目コード,小字コード,記号１,記号１名称,本番,記号２,記号２名称,枝番,記号３,記号３名称,孫番,曾孫番,現況分割


        AddColumn("申請世帯A", "譲渡人農家コード")
        AddColumn("申請者A", "譲渡人住民コード")
        AddColumn("氏名A", "譲渡人氏名")
        AddColumn("年齢A", "譲渡人年齢")
        AddColumn("住所A", "譲渡人住所")
        AddColumn("申請者住民区分A", "譲渡人住民区分コード")
        AddColumn("譲渡人住民区分名称", "譲渡人住民区分名称")
        AddColumn("譲渡人認定農業者コード", "譲渡人認定農業者コード")
        AddColumn("譲渡人認定農業者名称", "譲渡人認定農業者名称")
        AddColumn("経営面積A", "譲渡人経営面積")

        AddColumn("申請世帯B", "譲受人農家コード")
        AddColumn("申請者B", "譲受人住民コード")
        AddColumn("氏名B", "譲受人氏名")
        AddColumn("年齢B", "譲受人年齢")
        AddColumn("住所B", "譲受人住所")
        AddColumn("申請者住民区分B", "譲受人住民区分コード")
        AddColumn("譲受人住民区分名称", "譲受人住民区分名称")
        AddColumn("譲受人認定農業者コード", "譲受人認定農業者コード")
        AddColumn("譲受人認定農業者名称", "譲受人認定農業者名称")
        AddColumn("経営面積B", "譲受人経営面積")

        ',,筆数_田,筆数_畑,筆数_計,面積_田,面積_畑,面積_計,受人の経営形態,工事内容

        AddColumn("土地改良区意見", "土地改良区意見")

        'mvarGrid.Createエクセル出力Ctrl(Me.ToolStrip)
        Me.ToolStrip.Items.AddRange(
            {
                mvar検索開始, mvar検索終了,
                New ToolStripSeparator,
                mvarBtn検索開始
            }
        )
    End Sub

    Private Sub mvarBtn検索開始_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarBtn検索開始.Click
        Dim sWhere As String
        With mvar検索開始.Value
            sWhere = String.Format("[受付年月日]>=#{0}/{1}/{2}#", .Month, .Day, .Year)
        End With
        With mvar検索終了.Value
            sWhere &= String.Format(" AND [受付年月日]<=#{0}/{1}/{2}#", .Month, .Day, .Year)
        End With
        mvarDSet = New HimTools2012.Data.DataSetEx

        mvarTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT IIf([法令]=30 Or [法令]=31,'',[申請理由A]) AS 目的, * FROM D_申請 WHERE (((D_申請.[法令]) In (30,31,40,50,51,52,302,303,602)) AND ((D_申請.[状態])=0) AND " & sWhere & ") ORDER BY IIf([法令]=302 Or [法令]=303,35,[法令]), D_申請.[受付番号];")
        mvarDSet.Tables.Add(mvarTable)

        Dim p農業委員 As DataTable = GetMaster("農業委員")

        mvarDSet.Relations.Add(New DataRelation("農業委員001", p農業委員.Columns("ID"), mvarTable.Columns("農業委員1"), False))
        mvarDSet.Relations.Add(New DataRelation("農業委員002", p農業委員.Columns("ID"), mvarTable.Columns("農業委員2"), False))
        mvarDSet.Relations.Add(New DataRelation("農業委員003", p農業委員.Columns("ID"), mvarTable.Columns("農業委員3"), False))

        Dim p法令 As DataTable = GetMaster("法令")
        mvarDSet.Relations.Add(New DataRelation("R法令", p法令.Columns("ID"), mvarTable.Columns("法令"), False))

        Dim p所有権移転の種類 As DataTable = GetMaster("所有権移転の種類")
        mvarDSet.Relations.Add(New DataRelation("R所有権移転の種類", p所有権移転の種類.Columns("ID"), mvarTable.Columns("所有権移転の種類"), False))

        Dim p農地種別 As DataTable = GetMaster("農地区分")
        mvarDSet.Relations.Add(New DataRelation("R農地区分", p農地種別.Columns("ID"), mvarTable.Columns("農地区分"), False))

        mvarTable.Columns.Add("調査員1名称", GetType(String), "Parent(農業委員001).名称")
        mvarTable.Columns.Add("調査員2名称", GetType(String), "Parent(農業委員002).名称")
        mvarTable.Columns.Add("調査員3名称", GetType(String), "Parent(農業委員003).名称")

        mvarTable.Columns.Add("法令名称", GetType(String), "Parent(R法令).名称")
        mvarTable.Columns.Add("権利コード", GetType(Integer), "IsNull([所有権移転の種類],0)*10")
        mvarTable.Columns.Add("権利名称", GetType(String))
        mvarTable.Columns.Add("設定名称", GetType(String), "Parent(R所有権移転の種類).名称")
        mvarTable.Columns.Add("農地区分名称", GetType(String), "Parent(R農地区分).名称")

        mvarTable.Columns.Add("氏名", GetType(String), "IsNull([氏名B],[氏名A])")
        mvarTable.Columns.Add("行政書士名", GetType(String), "[行政書士]")
        mvarTable.Columns.Add("申請地", GetType(String))

        mvarTable.Columns.Add("申請者住民区分A", GetType(Integer))
        mvarTable.Columns.Add("申請者住民区分B", GetType(Integer))
        mvarTable.Columns.Add("譲渡人認定農業者コード", GetType(Integer))
        mvarTable.Columns.Add("譲受人認定農業者コード", GetType(Integer))

        Dim p住民区分 As DataTable = GetMaster("住民区分")
        mvarDSet.Relations.Add(New DataRelation("R住民区分A", p住民区分.Columns("ID"), mvarTable.Columns("申請者住民区分A"), False))
        mvarDSet.Relations.Add(New DataRelation("R住民区分B", p住民区分.Columns("ID"), mvarTable.Columns("申請者住民区分B"), False))

        mvarTable.Columns.Add("譲渡人住民区分名称", GetType(String), "Parent(R住民区分A).名称")
        mvarTable.Columns.Add("譲受人住民区分名称", GetType(String), "Parent(R住民区分B).名称")
        mvarTable.Columns.Add("譲渡人認定農業者名称", GetType(String))
        mvarTable.Columns.Add("譲受人認定農業者名称", GetType(String))
        '
        Dim p申請書_土地改良区の意見書について = GetMaster("申請書_土地改良区の意見書について")
        mvarDSet.Relations.Add(New DataRelation("R申請書_土地改良区の意見書について", p申請書_土地改良区の意見書について.Columns("ID"), mvarTable.Columns("土地改良区の意見書の有不用"), False))
        mvarTable.Columns.Add("土地改良区意見", GetType(String), "Parent(R申請書_土地改良区の意見書について).名称")

        For Each pRow As DataRow In mvarTable.Rows
            Dim nCount As Integer = -1
            Dim s申請地 As String = ""
            Dim SList As String() = Split(Trim(pRow.Item("農地リスト").ToString), ";")
            Array.Sort(SList)

            If SList.Length > 0 Then
                For Each s As String In SList
                    Dim ID As String = Val(CommonFunc.GetKeyCode(s))
                    Dim p農地 As CObj農地 = ObjectMan.GetObject("農地." & ID)
                    Dim p転用農地 As CObj転用農地 = ObjectMan.GetObject("転用農地." & ID)
                    nCount += 1

                    If p農地 IsNot Nothing AndAlso s申請地 = "" Then
                        s申請地 = p農地.土地所在
                    ElseIf p転用農地 IsNot Nothing AndAlso s申請地 = "" Then
                        s申請地 = p転用農地.土地所在
                    End If
                Next

                pRow.Item("申請地") = s申請地 & IIf(nCount > 0, " 外 " & nCount & " 筆", "")


            End If

            Select Case pRow.Item("法令")
                Case enum法令.農用地計画変更
                    pRow.Item("権利名称") = IIf(Val(pRow.Item("区分").ToString) = 1, "農振除外", IIf(Val(pRow.Item("区分").ToString) = 2, "用途区分変更", IIf(Val(pRow.Item("区分").ToString) = 3, "農振編入", "")))
                Case Else
                    pRow.Item("権利名称") = IIf(pRow.Item("権利コード") > 1, "所有権", IIf(pRow.Item("権利コード") = 1, "所有権", ""))
            End Select


            If Not IsDBNull(pRow.Item("申請者A")) AndAlso pRow.Item("申請者A") <> 0 Then
                Dim pRowK As DataRow = App農地基本台帳.TBL個人.FindRowByID(pRow.Item("申請者A"))
                If pRowK IsNot Nothing Then
                    pRow.Item("申請者住民区分A") = Val(pRowK.Item("住民区分").ToString)
                    pRow.Item("譲渡人認定農業者コード") = Val(pRowK.Item("農業改善計画認定").ToString)
                    Select Case pRow.Item("譲渡人認定農業者コード")
                        Case 1 : pRow.Item("譲渡人認定農業者名称") = "認定農業"
                        Case 2 : pRow.Item("譲渡人認定農業者名称") = ""
                        Case 3 : pRow.Item("譲渡人認定農業者名称") = ""
                        Case 4 : pRow.Item("譲渡人認定農業者名称") = "認定農業"
                    End Select
                End If
            End If
            If Not IsDBNull(pRow.Item("申請者B")) AndAlso pRow.Item("申請者B") <> 0 Then
                Dim pRowK As DataRow = App農地基本台帳.TBL個人.FindRowByID(pRow.Item("申請者B"))
                If pRowK IsNot Nothing Then
                    pRow.Item("申請者住民区分B") = Val(pRowK.Item("住民区分").ToString)
                    pRow.Item("譲受人認定農業者コード") = Val(pRowK.Item("農業改善計画認定").ToString)
                    Select Case pRow.Item("譲受人認定農業者コード")
                        Case 1 : pRow.Item("譲受人認定農業者名称") = "認定農業"
                        Case 2 : pRow.Item("譲受人認定農業者名称") = ""
                        Case 3 : pRow.Item("譲受人認定農業者名称") = ""
                        Case 4 : pRow.Item("譲受人認定農業者名称") = "認定農業"
                    End Select
                End If
            End If
        Next

        mvarGrid.AutoGenerateColumns = False
        mvarGrid.SetDataView(mvarTable, "", "")

    End Sub

    Private Sub AddColumn(ByVal sDataName As String, ByVal sHeaderName As String)
        Dim pColumn As DataGridViewColumn = New DataGridViewTextBoxColumn
        pColumn.DataPropertyName = sDataName
        pColumn.HeaderText = sHeaderName
        mvarGrid.Columns.Add(pColumn)
    End Sub

    Private Function GetMaster(ByRef sClass As String) As DataTable
        Dim pM As DataTable = New DataView(App農地基本台帳.DataMaster.Body, "[Class]='" & sClass & "'", "", DataViewRowState.CurrentRows).ToTable
        pM.TableName = sClass
        mvarDSet.Tables.Add(pM)
        Return pM
    End Function
End Class