Imports System.IO
Imports System.Text
Imports HimTools2012.controls

'日付の時間等の取り扱い
'数値型のその他
'取り込み

Public Class CTabPageF2利用状況意向調査
    Inherits CTabPageWithToolStrip

    Public mvarTarget As ToolStripComboBox
    Public mvarTabControl As TabControl
    Public mvarTabPage出力 As TabPage
    Public mvarTabPage読込 As TabPage
    Public WithEvents mvar出力準備 As ToolStripButton
    Public WithEvents mvar読込準備 As ToolStripButton
    Public WithEvents mvar出力開始 As ToolStripButton
    Public WithEvents mvar読込開始 As ToolStripButton
    Private mvarGrid出力 As DataGridViewWithDataView
    Private mvarGrid読込 As DataGridViewWithDataView

    Private TBL農地 As DataTable
    Private mvarGridTBL As DataTable
    Private TBL市町村コード As New DataTable

    Public Sub New()
        MyBase.New(True, True, "フェーズ２利用状況調査・意向調査情報", "フェーズ２利用状況調査・意向調査情報")

        mvarTarget = New ToolStripComboBox
        With mvarTarget
            .Items.Add("利用状況調査・意向調査情報")
            .SelectedIndex = 0
            .AutoSize = False
            .Width = 240
        End With

        mvar出力準備 = New ToolStripButton("出力準備")
        mvar読込準備 = New ToolStripButton("読込準備")
        mvar出力開始 = New ToolStripButton("出力開始")
        mvar読込開始 = New ToolStripButton("読込開始")

        CreateGrid(mvarGrid出力)
        CreateGrid(mvarGrid読込)

        mvarTabControl = New TabControl
        With mvarTabControl
            .Dock = DockStyle.Fill
        End With
        mvarTabPage出力 = New TabPage
        With mvarTabPage出力
            .Text = "出力内容"
            .Controls.Add(mvarGrid出力)
        End With
        mvarTabPage読込 = New TabPage
        With mvarTabPage読込
            .Text = "読込内容"
            .Controls.Add(mvarGrid読込)
        End With
        mvarTabControl.TabPages.AddRange(New TabPage() {mvarTabPage出力, mvarTabPage読込})

        Me.ControlPanel.Add(mvarTabControl)
        Me.ToolStrip.Items.AddRange({New ToolStripLabel("出力対象"), mvarTarget, mvar出力準備, mvar出力開始, New ToolStripSeparator, mvar読込準備, mvar読込開始})

        TBL農地 = New DataTable

        CreateTable(mvarGridTBL)

        都道府県ID = Val(SysAD.DB(sLRDB).DBProperty("都道府県ID").ToString)
        市町村CD = Val(SysAD.DB(sLRDB).DBProperty("市町村ID").ToString)
        市町村名 = SysAD.DB(sLRDB).DBProperty("市町村名")
        Set市町村コード()
    End Sub

    Private Sub CreateGrid(ByRef pDGrid As DataGridViewWithDataView)
        pDGrid = New DataGridViewWithDataView
        AddCol(pDGrid, "連携用所在地キー", "連携用所在地キー", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "市町村コード", "市町村コード", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "市町村名", "市町村名", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "大字コード", "大字コード", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "大字名", "大字名", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "小字コード", "小字コード", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "小字名", "小字名", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "本番区分", "本番区分", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "本番", "本番", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "枝番区分", "枝番区分", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "枝番", "枝番", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "孫番区分", "孫番区分", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "孫番", "孫番", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "曾孫番区分", "曾孫番区分", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "曾孫番", "曾孫番", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "玄孫番区分", "玄孫番区分", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "玄孫番", "玄孫番", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "区分", "区分", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "利用状況調査年月日", "利用状況調査年月日", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "状況調査結果", "状況調査結果", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "荒廃農地調査分類", "荒廃農地調査分類", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "調査利用状況", "調査利用状況", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "調査委員名", "調査委員名", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "耕作放棄地通し番号", "耕作放棄地通し番号", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "一時転用", "一時転用", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "無断転用", "無断転用", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "違反転用", "違反転用", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "調査不可判断年月日", "調査不可判断年月日", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "調査不可理由", "調査不可理由", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "調査不可理由その他内訳", "調査不可理由その他内訳", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "利用意向調査年月日", "利用意向調査年月日", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "意向根拠条項", "意向根拠条項", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "意向表明年月日", "意向表明年月日", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "意向調査結果", "意向調査結果", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "意向調査結果その他内訳", "意向調査結果その他内訳", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "措置の実施状況", "措置の実施状況", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "権利関係調査", "権利関係調査", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "権利関係調査結果年月日", "権利関係調査結果年月日", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "権利関係調査結果", "権利関係調査結果", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "権利関係調査結果その他内訳", "権利関係調査結果その他内訳", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "公示年月日(第32条第3項)", "公示年月日", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "通知年月日(第41条第1項)", "通知年月日", New DataGridViewTextBoxColumn)

        pDGrid.AutoGenerateColumns = False
        pDGrid.AllowUserToAddRows = False
        pDGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
    End Sub

    Private Sub CreateTable(ByRef pTBL As DataTable)
        pTBL = New DataTable
        pTBL.Columns.Add("連携用所在地キー")
        pTBL.Columns.Add("市町村コード")
        pTBL.Columns.Add("市町村名")
        pTBL.Columns.Add("大字コード")
        pTBL.Columns.Add("大字名")
        pTBL.Columns.Add("小字コード")
        pTBL.Columns.Add("小字名")
        pTBL.Columns.Add("本番区分")
        pTBL.Columns.Add("本番")
        pTBL.Columns.Add("枝番区分")
        pTBL.Columns.Add("枝番")
        pTBL.Columns.Add("孫番区分")
        pTBL.Columns.Add("孫番")
        pTBL.Columns.Add("曾孫番区分")
        pTBL.Columns.Add("曾孫番")
        pTBL.Columns.Add("玄孫番区分")
        pTBL.Columns.Add("玄孫番")
        pTBL.Columns.Add("区分")
        pTBL.Columns.Add("利用状況調査年月日")
        pTBL.Columns.Add("状況調査結果")
        pTBL.Columns.Add("荒廃農地調査分類")
        pTBL.Columns.Add("調査利用状況")
        pTBL.Columns.Add("調査委員名")
        pTBL.Columns.Add("耕作放棄地通し番号")
        pTBL.Columns.Add("一時転用")
        pTBL.Columns.Add("無断転用")
        pTBL.Columns.Add("違反転用")
        pTBL.Columns.Add("調査不可判断年月日")
        pTBL.Columns.Add("調査不可理由")
        pTBL.Columns.Add("調査不可理由その他内訳")
        pTBL.Columns.Add("利用意向調査年月日")
        pTBL.Columns.Add("意向根拠条項")
        pTBL.Columns.Add("意向表明年月日")
        pTBL.Columns.Add("意向調査結果")
        pTBL.Columns.Add("意向調査結果その他内訳")
        pTBL.Columns.Add("措置の実施状況")
        pTBL.Columns.Add("権利関係調査")
        pTBL.Columns.Add("権利関係調査結果年月日")
        pTBL.Columns.Add("権利関係調査結果")
        pTBL.Columns.Add("権利関係調査結果その他内訳")
        pTBL.Columns.Add("公示年月日")
        pTBL.Columns.Add("通知年月日")
    End Sub

    Private Sub AddCol(ByRef pDGrid As DataGridViewWithDataView, ByVal sHeader As String, ByVal sData As String, ByVal pCol As DataGridViewColumn)
        pCol.HeaderText = sHeader
        pCol.DataPropertyName = sData
        pDGrid.Columns.Add(pCol)
    End Sub

    Private pView As DataView
    'Private Sub mvar検索開始_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar検索開始.Click
    'If mvarStart.Value < mvarEnd.Value Then
    '    Dim sWhere As String = String.Format("[利用意向調査日]>=#{0}/{1}/{2}# AND [利用意向調査日]<=#{3}/{4}/{5}#",
    '                                        mvarStart.Value.Month, mvarStart.Value.Day, mvarStart.Value.Year, mvarEnd.Value.Month, mvarEnd.Value.Day, mvarEnd.Value.Year)

    '    pView = New DataView(TBL農地, sWhere, "[所有者ID]", DataViewRowState.CurrentRows)

    '    Dim RowCount As Integer = 1
    '    Dim SaveUserNo As Integer = 0
    '    Dim UserNoCount As Integer = 0
    '    For Each pRow As DataRowView In pView
    '        pRow.Item("農地整理番号") = RowCount
    '        pRow.Item("意向調査年度") = Replace(Replace(和暦Format(mvarStart.Text, "gggyy"), "平成", ""), "令和", "")
    '        If SaveUserNo = 0 OrElse SaveUserNo <> pRow.Item("所有者ID") Then
    '            UserNoCount += 1
    '            SaveUserNo = pRow.Item("所有者ID")
    '        End If
    '        pRow.Item("所有者等番号") = UserNoCount
    '        pRow.Item("権利種類") = 1

    '        Select Case pRow.Item("名称")
    '            Case "田", "畑"
    '            Case Else : pRow.Item("名称") = ""
    '        End Select

    '        Select Case Val(pRow.Item("農振法区分").ToString)
    '            Case enum農振法区分.農用地区域 : pRow.Item("農用地区分") = 1
    '            Case enum農振法区分.農振地域 : pRow.Item("農用地区分") = 2
    '            Case Else
    '                Select Case Val(pRow.Item("農業振興地域").ToString)
    '                    Case enum農業振興地域.農用地内 : pRow.Item("農用地区分") = 1
    '                    Case enum農業振興地域.農用地外 : pRow.Item("農用地区分") = 2
    '                    Case Else : pRow.Item("農用地区分") = ""
    '                End Select
    '        End Select

    '        Select Case Val(pRow.Item("利用状況調査農地法").ToString)
    '            Case 1 To 2 : pRow.Item("現地調査状況") = pRow.Item("利用状況調査農地法")
    '            Case Else : pRow.Item("現地調査状況") = ""
    '        End Select

    '        Select Case Val(pRow.Item("利用意向意向内容区分").ToString)
    '            Case 1 : pRow.Item("利用意向意向内容区分") = 4
    '            Case 2 : pRow.Item("利用意向意向内容区分") = 1
    '            Case 3 : pRow.Item("利用意向意向内容区分") = 2
    '            Case 4 : pRow.Item("利用意向意向内容区分") = 3
    '            Case 5 : pRow.Item("利用意向意向内容区分") = 5
    '            Case Else : pRow.Item("利用意向意向内容区分") = 6
    '        End Select

    '        RowCount += 1
    '    Next

    '    mvarGrid.SetDataView(pView.ToTable, sWhere, "[所有者ID]")
    'End If
    'End Sub

    'Private Sub mvar帳票出力_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar帳票出力.Click
    '    If IO.File.Exists(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\利用意向調査実施農地状況一覧.xml") Then
    '        Dim sXML As String = HimTools2012.TextAdapter.LoadTextFile(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\利用意向調査実施農地状況一覧.xml")
    '        Dim pXMLSS As New HimTools2012.Excel.XMLSS2003.CXMLSS2003(sXML)
    '        Dim pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet = pXMLSS.WorkBook.WorkSheets.Items("農地状況一覧")
    '        Dim pLoopRow As New HimTools2012.Excel.XMLSS2003.XMLLoopRows(pSheet, "{No}")
    '        Dim nLoop As Integer = -1

    '        For Each pRow As DataRowView In pView
    '            Dim pXRow As HimTools2012.Excel.XMLSS2003.XMLSSRow = Nothing

    '            If nLoop = -1 Then
    '                pXRow = pSheet.FindRowInstrText("{No}")(0)
    '            Else

    '                For Each pCXRow As HimTools2012.Excel.XMLSS2003.XMLSSRow In pLoopRow
    '                    pXRow = pCXRow.CopyRow

    '                    pSheet.Table.Rows.InsertRow(pLoopRow.InsetRow, pXRow)
    '                    pLoopRow.InsetRow += 1
    '                Next
    '            End If
    '            nLoop += 1

    '            With pXRow
    '                .ValueReplace("{No}", pRow.Item("農地整理番号").ToString)
    '                .ValueReplace("{年度}", pRow.Item("意向調査年度").ToString)
    '                .ValueReplace("{市町村CD}", pRow.Item("市町村CD").ToString)
    '                .ValueReplace("{市町村名}", pRow.Item("市町村名").ToString)
    '                .ValueReplace("{所有者No}", pRow.Item("所有者等番号").ToString)
    '                .ValueReplace("{氏名}", pRow.Item("氏名").ToString)
    '                .ValueReplace("{権利種類}", pRow.Item("権利種類").ToString)
    '                .ValueReplace("{大字}", pRow.Item("大字").ToString)
    '                .ValueReplace("{小字}", pRow.Item("小字").ToString)
    '                .ValueReplace("{地番}", pRow.Item("地番").ToString)
    '                .ValueReplace("{地目}", pRow.Item("名称").ToString)
    '                .ValueReplace("{面積}", DecConv(CDec(pRow.Item("実面積").ToString)))
    '                .ValueReplace("{区分}", pRow.Item("農用地区分").ToString)
    '                .ValueReplace("{調査状況}", pRow.Item("現地調査状況").ToString)
    '                .ValueReplace("{再生可否}", "○")    '/***初期値***/
    '                .ValueReplace("{借受希望}", "×")
    '                .ValueReplace("{適正賃料}", "○")
    '                .ValueReplace("{集積適地}", IIf(Val(pRow.Item("農用地区分").ToString) = 1, "○", IIf(Val(pRow.Item("農用地区分").ToString) = 2, "×", "")))
    '                .ValueReplace("{適否結果○×}", "―")
    '                .ValueReplace("{適否結果×○}", "―")
    '                .ValueReplace("{活用可能性}", "×")
    '                If Not IsDBNull(pRow.Item("利用意向意思表明日")) Then
    '                    .ValueReplace("{回答日}", 和暦Format(pRow.Item("利用意向意思表明日"), "gyy.MM.dd"))
    '                Else
    '                    .ValueReplace("{回答日}", "")
    '                End If
    '                .ValueReplace("{回答内容}", pRow.Item("利用意向意向内容区分").ToString)
    '            End With
    '        Next

    '        Dim sFile As String = My.Computer.FileSystem.SpecialDirectories.Desktop & "\xx利用意向調査実施農地状況一覧.xml"
    '        Try
    '            HimTools2012.TextAdapter.SaveTextFile(sFile, pXMLSS.OutPut(False))
    '            MsgBox("エクセルをデスクトップに作成しました。", MsgBoxStyle.Information)
    '            SysAD.ShowFolder(sFile)
    '        Catch ex As Exception

    '        End Try

    '    End If
    'End Sub
    Private Function DecConv(ByVal pDec As Decimal) As String
        If Fix(pDec) = pDec Then
            Return Val(pDec).ToString("#,##0")
        Else
            Return Val(pDec).ToString("#,##0.##") '小数点第2位まで表示
        End If

        Return pDec
    End Function


    Private Sub mvar出力準備_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar出力準備.Click
        If TBL農地.Rows.Count = 0 Then
            Dim DSet As DataSet = New DataSet
            TBL農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [V_農地]")
            DSet.Tables.Add(TBL農地)

            Dim TBL大字 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [V_大字]")
            DSet.Tables.Add(TBL大字)
            Dim TBL小字 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [V_小字]")
            DSet.Tables.Add(TBL小字)

            DSet.Relations.Add("大字情報", TBL大字.Columns("ID"), TBL農地.Columns("大字ID"), False)
            TBL農地.Columns.Add("大字名", GetType(String), "Parent(大字情報).名称")

            DSet.Relations.Add("小字情報", TBL小字.Columns("ID"), TBL農地.Columns("小字ID"), False)
            TBL農地.Columns.Add("小字名", GetType(String), "Parent(小字情報).名称")
        End If

        For Each pRow As DataRow In TBL農地.Rows
            Dim addRow As DataRow = mvarGridTBL.NewRow

            Dim 市町村CD As String = 市町村コード(pRow.Item("土地所在").ToString)
            Dim Conv市町村CD As String = IIf(Len(pRow.Item("市町村ID").ToString) = 6, pRow.Item("市町村ID").ToString, IIf(市町村CD = "", Find市町村コード(pRow.Item("土地所在").ToString), 市町村CD))

            addRow.Item("連携用所在地キー") = Fnc連携Key(pRow, Conv市町村CD)
            addRow.Item("市町村コード") = Conv市町村CD
            addRow.Item("市町村名") = 市町村名
            addRow.Item("大字コード") = Val(pRow.Item("大字ID").ToString)
            addRow.Item("大字名") = pRow.Item("大字名").ToString
            addRow.Item("小字コード") = Val(pRow.Item("小字ID").ToString)
            addRow.Item("小字名") = pRow.Item("小字名").ToString
            addRow.Item("本番区分") = s本番区分
            addRow.Item("本番") = s本番
            addRow.Item("枝番区分") = s枝番区分
            addRow.Item("枝番") = s枝番
            addRow.Item("孫番区分") = s孫番区分
            addRow.Item("孫番") = s孫番
            addRow.Item("曾孫番区分") = ""
            addRow.Item("曾孫番") = ""
            addRow.Item("玄孫番区分") = ""
            addRow.Item("玄孫番") = ""
            addRow.Item("区分") = Val(pRow.Item("一部現況").ToString)
            addRow.Item("利用状況調査年月日") = pRow.Item("利用状況調査日")
            addRow.Item("状況調査結果") = Val(pRow.Item("利用状況調査農地法").ToString)
            addRow.Item("荒廃農地調査分類") = Val(pRow.Item("利用状況調査荒廃").ToString)
            addRow.Item("調査利用状況") = pRow.Item("利用状況").ToString
            addRow.Item("調査委員名") = ""
            addRow.Item("耕作放棄地通し番号") = Val(pRow.Item("利用状況耕作放棄地通し番号").ToString)
            addRow.Item("一時転用") = IIf(Val(pRow.Item("利用状況調査転用").ToString) = 1, pRow.Item("利用状況一時転用区分"), 0)
            addRow.Item("無断転用") = IIf(Val(pRow.Item("利用状況調査転用").ToString) = 2, 2, 1)
            addRow.Item("違反転用") = IIf(Val(pRow.Item("利用状況調査転用").ToString) = 3, 2, 1)
            addRow.Item("調査不可判断年月日") = pRow.Item("利用状況調査不可判断日")
            addRow.Item("調査不可理由") = Val(pRow.Item("利用状況調査不可判断理由").ToString)
            addRow.Item("調査不可理由その他内訳") = pRow.Item("利用状況調査不可判断その他理由").ToString
            addRow.Item("利用意向調査年月日") = pRow.Item("利用意向調査日")
            addRow.Item("意向根拠条項") = Val(pRow.Item("利用意向根拠条項").ToString)
            addRow.Item("意向表明年月日") = pRow.Item("利用意向意思表明日")
            addRow.Item("意向調査結果") = Val(pRow.Item("利用意向意向内容区分").ToString)
            addRow.Item("意向調査結果その他内訳") = pRow.Item("利用意向調査結果その他理由").ToString
            addRow.Item("措置の実施状況") = pRow.Item("利用意向措置実施状況").ToString
            addRow.Item("権利関係調査") = Val(pRow.Item("利用意向権利関係調査区分").ToString)
            addRow.Item("権利関係調査結果年月日") = pRow.Item("利用意向調査不可年月日")
            addRow.Item("権利関係調査結果") = Val(pRow.Item("利用意向調査不可結果").ToString)
            addRow.Item("権利関係調査結果その他内訳") = pRow.Item("利用意向権利関係調査記録")
            addRow.Item("公示年月日") = pRow.Item("利用意向公示年月日")
            addRow.Item("通知年月日") = pRow.Item("利用意向通知年月日")


            mvarGridTBL.Rows.Add(addRow)
        Next

        mvarGrid出力.DataSource = New DataView(mvarGridTBL)


    End Sub

    Private Sub mvar出力開始_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar出力開始.Click
        Dim savePath As String = ""
        With New SaveFileDialog
            .FileName = "利用状況調査・意向調査情報.csv"
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            .Filter = "CSVファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*"

            If .ShowDialog = DialogResult.OK Then
                savePath = .FileName
            End If
        End With

        SaveToCSV(mvarGridTBL, savePath, True, ",", "", "")
    End Sub

    Private Sub SaveToCSV(ByVal dt As DataTable, ByVal fileName As String, ByVal hasHeader As Boolean, ByVal separator As String, ByVal quote As String, ByVal replace As String)
        Dim rows As Integer = dt.Rows.Count
        Dim cols As Integer = dt.Columns.Count
        Dim text As String

        Dim writer As StreamWriter = New StreamWriter(fileName, False, Encoding.GetEncoding("shift_jis"))
        'カラム名を保存するか
        If hasHeader Then
            For i As Integer = 0 To cols - 1 Step 1
                If quote <> "" Then
                    text = dt.Columns(i).ColumnName.Replace(quote, replace)
                Else
                    text = dt.Columns(i).ColumnName
                End If
                If i <> cols - 1 Then
                    writer.Write(quote + text + quote + separator)
                Else
                    writer.WriteLine(quote + text + quote)
                End If
            Next
        End If
        'データの保存処理
        For i As Integer = 0 To rows - 1 Step 1
            For j As Integer = 0 To cols - 1 Step 1
                If quote <> "" Then
                    text = dt.Rows(i)(j).ToString().Replace(quote, replace)
                Else
                    text = dt.Rows(i)(j).ToString()
                End If
                If j <> cols - 1 Then
                    writer.Write(quote + text + quote + separator)
                Else
                    writer.WriteLine(quote + text + quote)
                End If
            Next j
        Next i

        writer.Close()
    End Sub

    Dim s本番区分 As String = "" : Dim s本番 As String = ""
    Dim s枝番区分 As String = "" : Dim s枝番 As String = ""
    Dim s孫番区分 As String = "" : Dim s孫番 As String = ""
    Private Function Fnc連携Key(ByVal pRow As DataRow, ByVal CityCode As String) As String
        Dim pAddress As String = Replace(pRow.Item("地番").ToString, "の", "")

        pAddress = StrConv(pAddress, vbNarrow)

        If InStr(pAddress, "-") > 0 Then        '地番が"-"を含むかどうか
            s本番 = Val(HimTools2012.StringF.Left(pAddress, InStr(pAddress, "-") - 1))
            Dim s分岐1 As String = Mid(pAddress, InStr(pAddress, "-") + 1)

            If InStr(s分岐1, "-") > 0 Then        '枝番以降が"-"を含むかどうか
                Dim s分岐2 As String
                Dim s分岐3 As String
                If Char.IsNumber(s分岐1, 0) Then
                    s枝番 = Val(HimTools2012.StringF.Left(s分岐1, InStr(s分岐1, "-") - 1))
                    s分岐2 = Mid(s分岐1, InStr(s分岐1, "-") + 1)

                    If InStr(s分岐2, "-") > 0 Then
                        If Char.IsNumber(s分岐2, 0) Then
                            s孫番 = Val(HimTools2012.StringF.Left(s分岐2, InStr(s分岐2, "-") - 1))
                            s孫番区分 = StrConv(Mid(s分岐2, InStr(s分岐2, "-") + 1), VbStrConv.Wide)
                            '終了
                        Else
                            s枝番区分 = StrConv(HimTools2012.StringF.Left(s分岐2, InStr(s分岐2, "-") - 1), VbStrConv.Wide)
                            s分岐3 = Mid(s分岐2, InStr(s分岐2, "-") + 1)

                            If InStr(s分岐3, "-") > 0 Then
                                s孫番 = Val(HimTools2012.StringF.Left(s分岐3, InStr(s分岐3, "-") - 1))
                                s孫番区分 = StrConv(Mid(s分岐3, InStr(s分岐3, "-") + 1), VbStrConv.Wide)
                                '終了
                            Else
                                If Char.IsNumber(s分岐3, 0) Then : s孫番 = Val(s分岐3)
                                Else : s孫番区分 = StrConv(s分岐3, VbStrConv.Wide)
                                End If
                                '終了
                            End If
                        End If
                    Else
                        If Char.IsNumber(s分岐2, 0) Then : s孫番 = Val(s分岐2)
                        Else : s枝番区分 = StrConv(s分岐2, VbStrConv.Wide)
                        End If
                        '終了
                    End If
                Else
                    s本番区分 = StrConv(HimTools2012.StringF.Left(s分岐1, InStr(s分岐1, "-") - 1), VbStrConv.Wide)
                    s分岐2 = Mid(s分岐1, InStr(s分岐1, "-") + 1)

                    If InStr(s分岐2, "-") > 0 Then
                        If Char.IsNumber(s分岐2, 0) Then
                            s枝番 = Val(HimTools2012.StringF.Left(s分岐2, InStr(s分岐2, "-") - 1))
                            s分岐3 = Mid(s分岐2, InStr(s分岐2, "-") + 1)

                            If InStr(s分岐3, "-") > 0 Then
                                If Char.IsNumber(s分岐3, 0) Then
                                    s孫番 = Val(HimTools2012.StringF.Left(s分岐3, InStr(s分岐3, "-") - 1))
                                    s孫番区分 = StrConv(Mid(s分岐3, InStr(s分岐3, "-") + 1), VbStrConv.Wide)
                                    '終了
                                Else
                                    s枝番区分 = StrConv(HimTools2012.StringF.Left(s分岐3, InStr(s分岐3, "-") - 1), VbStrConv.Wide)
                                    Dim s分岐4 As String = Mid(s分岐3, InStr(s分岐3, "-") + 1)

                                    If InStr(s分岐4, "-") > 0 Then
                                        s孫番 = Val(HimTools2012.StringF.Left(s分岐4, InStr(s分岐4, "-") - 1))
                                        s孫番区分 = StrConv(Mid(s分岐4, InStr(s分岐4, "-") + 1), VbStrConv.Wide)
                                    Else
                                        s孫番区分 = StrConv(Mid(s分岐4, InStr(s分岐4, "-") + 1), VbStrConv.Wide)
                                    End If
                                    '終了
                                End If
                            Else
                                If Char.IsNumber(s分岐3, 0) Then : s孫番 = Val(s分岐3)
                                Else : s枝番区分 = StrConv(s分岐3, VbStrConv.Wide)
                                End If
                                '終了
                            End If
                        Else
                            s枝番区分 = StrConv(HimTools2012.StringF.Left(s分岐2, InStr(s分岐2, "-") - 1), VbStrConv.Wide)
                            s分岐3 = Mid(s分岐2, InStr(s分岐2, "-") + 1)

                            If InStr(s分岐3, "-") > 0 Then
                                s孫番 = Val(HimTools2012.StringF.Left(s分岐3, InStr(s分岐3, "-") - 1))
                                s孫番区分 = StrConv(Mid(s分岐3, InStr(s分岐3, "-") + 1), VbStrConv.Wide)
                                '終了
                            Else
                                If Char.IsNumber(s分岐3, 0) Then : s孫番 = Val(s分岐3)
                                Else : s孫番区分 = StrConv(s分岐3, VbStrConv.Wide)
                                End If
                                '終了
                            End If
                        End If
                    Else
                        If Char.IsNumber(s分岐2, 0) Then : s枝番 = Val(s分岐2)
                        Else : s枝番区分 = StrConv(s分岐2, VbStrConv.Wide)
                        End If
                        '終了
                    End If
                End If
            Else
                If Char.IsNumber(s分岐1, 0) Then : s枝番 = Val(s分岐1)
                Else : s本番区分 = StrConv(s分岐1, VbStrConv.Wide)
                End If
                '終了
            End If
        Else
            s本番 = Val(pAddress)
            '終了
        End If

        Return String.Format("{0}/{1}-{2}:{3}-{4}-{5}-{6}-{7}/{8}",
                             CityCode,
                             Format(Val(pRow.Item("大字ID").ToString), "00000000"),
                             Format(Val(pRow.Item("小字ID").ToString), "0000000"),
                             s本番区分 & s本番,
                             s枝番区分 & s枝番,
                             s孫番区分 & s孫番,
                             "",
                             "",
                             StrConv(Val(pRow.Item("一部現況").ToString), vbWide))
    End Function

    Public Sub Set市町村コード()
        If System.IO.File.Exists(SysAD.CustomReportFolder("共通様式") & "\code_list.csv") Then
            Dim cReader As New System.IO.StreamReader(SysAD.CustomReportFolder("共通様式") & "\code_list.csv", System.Text.Encoding.Default)
            Dim LoopCount As Integer = 0

            While (cReader.Peek() >= 0)
                Dim stBuffer As String = cReader.ReadLine() ' ファイルを 1 行ずつ読み込む
                Dim cAr As Object = Split(stBuffer, ",")

                If LoopCount = 0 Then
                    With TBL市町村コード
                        .Columns.Add(cAr(0))
                        .Columns.Add(cAr(1))
                        .Columns.Add(cAr(2))
                        .Columns.Add(cAr(3))
                        .Columns.Add(cAr(4))
                    End With

                    LoopCount += 1
                Else
                    Dim pRow As DataRow = TBL市町村コード.NewRow
                    pRow.Item("団体コード") = cAr(0)
                    pRow.Item("都道府県名（漢字）") = cAr(1)
                    pRow.Item("市区町村名（漢字）") = cAr(2)
                    pRow.Item("都道府県名（カナ）") = cAr(3)
                    pRow.Item("市区町村名（カナ）") = cAr(4)

                    TBL市町村コード.Rows.Add(pRow)
                End If
            End While
        End If
    End Sub

    Private cityCode As String = ""
    Private otherJusyo As String = ""
    Private kenCity As String = ""
    Public ReadOnly Property 市町村コード(ByVal sValue As String)
        Get
            Dim CodePath As String = ""
            If System.IO.File.Exists(SysAD.CustomReportFolder("共通様式") & "\code_list.csv") Then
                CodePath = SysAD.CustomReportFolder("共通様式") & "\code_list.csv"
            End If

            Dim cityCodeModel As CitiesCode.Interface.ICityCodeModel = New CitiesCode.Factory.CityCodeFactory().CreateCityCodeModel(CodePath)

            Dim jusyoModel As CitiesCode.Interface.IJusyoModel = cityCodeModel.GetCityCode(sValue)  ' 文字列より市町村コード取得
            If jusyoModel.MatchState = CitiesCode.Types.MatchType.Match Then

                ' Match以外はnull
                cityCode = jusyoModel.CityCode
                otherJusyo = jusyoModel.OtherJusyoText  ' その他の住所
                If Len(jusyoModel.OtherJusyoText) > 0 Then
                    kenCity = jusyoModel.JusyoText.Replace(jusyoModel.OtherJusyoText, "")  ' その他の住所以外
                Else
                    kenCity = ""
                End If

                Return cityCode
            End If

            Return ""
        End Get
    End Property

    Public 都道府県ID As String = ""
    Public 市町村CD As String = ""
    Public 市町村名 As String = ""
    Public Function Find市町村コード(ByVal sValue As String)
        Dim Find大分類 As String = ""
        Dim Find小分類 As String = ""

        For Each pRow As DataRow In TBL市町村コード.Rows
            If InStr(sValue, pRow.Item("都道府県名（漢字）")) > 0 Then
                Find大分類 = pRow.Item("都道府県名（漢字）")
                Exit For
            End If
        Next

        For Each pRow As DataRow In TBL市町村コード.Rows
            If InStr(sValue, pRow.Item("市区町村名（漢字）")) > 0 Then
                Find小分類 = pRow.Item("市区町村名（漢字）")
                Exit For
            End If
        Next

        Dim pTBL As DataTable
        If Len(Find大分類) > 0 AndAlso Len(Find小分類) > 0 Then
            pTBL = New DataView(TBL市町村コード, String.Format("[都道府県名（漢字）] = '{0}' And [市区町村名（漢字）] = '{1}'", Find大分類, Find小分類), "", DataViewRowState.CurrentRows).ToTable

            If pTBL.Rows.Count = 1 Then
                Dim FindDataRow() As DataRow = pTBL.Select(String.Format("[都道府県名（漢字）] = '{0}' And [市区町村名（漢字）] = '{1}'", Find大分類, Find小分類), "")
                Return FindDataRow(0).Item("団体コード")
            Else
                Return 都道府県ID & 市町村CD
            End If
        ElseIf Len(Find小分類) > 0 Then
            pTBL = New DataView(TBL市町村コード, String.Format("[市区町村名（漢字）] = '{0}'", Find小分類), "", DataViewRowState.CurrentRows).ToTable

            If pTBL.Rows.Count = 1 Then
                Dim FindDataRow() As DataRow = pTBL.Select(String.Format("[市区町村名（漢字）] = '{0}'", Find小分類), "")
                Return FindDataRow(0).Item("団体コード")
            Else
                Return 都道府県ID & 市町村CD
            End If
        Else
            Return 都道府県ID & 市町村CD
        End If

    End Function
End Class
