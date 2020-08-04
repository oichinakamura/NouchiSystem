Imports HimTools2012.controls

Public Class CTabPage利用意向調査実施農地状況一覧
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Public mvarStart As ToolStripDateTimePickerWithlabel
    Public mvarEnd As ToolStripDateTimePickerWithlabel
    Public WithEvents mvar検索開始 As ToolStripButton
    Public WithEvents mvar帳票出力 As ToolStripButton
    Public WithEvents mvarExcel出力 As ToolStripButton
    Private mvarGrid As HimTools2012.controls.DataGridViewWithDataView

    Private TBL農地 As DataTable
    Private TBL転用農地 As DataTable
    Private TBL削除農地 As DataTable

    Public Sub New()
        MyBase.New(True, True, "利用意向調査実施農地状況一覧", "利用意向調査実施農地状況一覧")

        mvarStart = New ToolStripDateTimePickerWithlabel("対象利用意向調査年月日")
        mvarEnd = New ToolStripDateTimePickerWithlabel("～")
        mvarStart.Value = Now.Date
        mvarEnd.Value = Now.Date

        mvar検索開始 = New ToolStripButton("検索開始")
        mvar帳票出力 = New ToolStripButton("帳票出力")
        mvarExcel出力 = New ToolStripButton("Excel出力")

        CreatemvarGrid(mvarGrid)

        Me.ControlPanel.Add(mvarGrid)
        Me.ToolStrip.Items.AddRange({mvarStart, mvarEnd, mvar検索開始, New ToolStripSeparator, mvar帳票出力, New ToolStripSeparator, mvarExcel出力})

        '/*****農地Info,転用農地.削除農地より検索*****/
        '/↓↓↓↓↓↓↓↓↓A分類農地のみの出力↓↓↓↓↓↓↓↓↓↓/
        'TBL農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:農地Info].所有者ID, [D:個人Info].氏名, V_大字.大字, V_小字.小字, [D:農地Info].地番, V_現況地目.名称, [D:農地Info].実面積, [D:農地Info].農業振興地域, [D:農地Info].農振法区分, [D:農地Info].利用状況調査農地法, [D:農地Info].利用意向調査日, [D:農地Info].利用意向意思表明日, [D:農地Info].利用意向意向内容区分 " & _
        '                                              "FROM ((([D:農地Info] LEFT JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) LEFT JOIN V_小字 ON [D:農地Info].小字ID = V_小字.ID) LEFT JOIN V_現況地目 ON [D:農地Info].現況地目 = V_現況地目.ID) LEFT JOIN [D:個人Info] ON [D:農地Info].所有者ID = [D:個人Info].ID " & _
        '                                              "WHERE ((([D:農地Info].利用意向調査日) Is Not Null) AND (([D:農地Info].利用状況調査荒廃)=1)) " & _
        '                                              "ORDER BY [D:農地Info].所有者ID, [D:農地Info].大字ID;")
        'TBL転用農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D_転用農地].所有者ID, [D:個人Info].氏名, V_大字.大字, V_小字.小字, [D_転用農地].地番, V_現況地目.名称, [D_転用農地].実面積, [D_転用農地].農業振興地域, [D_転用農地].農振法区分, [D_転用農地].利用状況調査農地法, [D_転用農地].利用意向調査日, [D_転用農地].利用意向意思表明日, [D_転用農地].利用意向意向内容区分 " & _
        '                                              "FROM ((([D_転用農地] LEFT JOIN V_大字 ON [D_転用農地].大字ID = V_大字.ID) LEFT JOIN V_小字 ON [D_転用農地].小字ID = V_小字.ID) LEFT JOIN V_現況地目 ON [D_転用農地].現況地目 = V_現況地目.ID) LEFT JOIN [D:個人Info] ON [D_転用農地].所有者ID = [D:個人Info].ID " & _
        '                                              "WHERE ((([D_転用農地].利用意向調査日) Is Not Null) AND (([D_転用農地].利用状況調査荒廃)=1)) " & _
        '                                              "ORDER BY [D_転用農地].所有者ID, [D_転用農地].大字ID;")
        'TBL農地.Merge(TBL転用農地)
        'TBL削除農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D_削除農地].所有者ID, [D:個人Info].氏名, V_大字.大字, V_小字.小字, [D_削除農地].地番, V_現況地目.名称, [D_削除農地].実面積, [D_削除農地].農業振興地域, [D_削除農地].農振法区分, [D_削除農地].利用状況調査農地法, [D_削除農地].利用意向調査日, [D_削除農地].利用意向意思表明日, [D_削除農地].利用意向意向内容区分 " & _
        '                                              "FROM ((([D_削除農地] LEFT JOIN V_大字 ON [D_削除農地].大字ID = V_大字.ID) LEFT JOIN V_小字 ON [D_削除農地].小字ID = V_小字.ID) LEFT JOIN V_現況地目 ON [D_削除農地].現況地目 = V_現況地目.ID) LEFT JOIN [D:個人Info] ON [D_削除農地].所有者ID = [D:個人Info].ID " & _
        '                                              "WHERE ((([D_削除農地].利用意向調査日) Is Not Null) AND (([D_削除農地].利用状況調査荒廃)=1)) " & _
        '                                              "ORDER BY [D_削除農地].所有者ID, [D_削除農地].大字ID;")
        'TBL農地.Merge(TBL削除農地)

        '/↓↓↓↓↓↓↓↓↓全農地が対象↓↓↓↓↓↓↓↓↓↓/
        TBL農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:農地Info].所有者ID, [D:個人Info].氏名, V_大字.大字, V_小字.小字, [D:農地Info].地番, V_現況地目.名称, [D:農地Info].実面積, [D:農地Info].農業振興地域, [D:農地Info].農振法区分, [D:農地Info].利用状況調査農地法, [D:農地Info].利用意向調査日, [D:農地Info].利用意向意思表明日, [D:農地Info].利用意向意向内容区分 " &
                                                      "FROM ((([D:農地Info] LEFT JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) LEFT JOIN V_小字 ON [D:農地Info].小字ID = V_小字.ID) LEFT JOIN V_現況地目 ON [D:農地Info].現況地目 = V_現況地目.ID) LEFT JOIN [D:個人Info] ON [D:農地Info].所有者ID = [D:個人Info].ID " &
                                                      "WHERE ((([D:農地Info].利用意向調査日) Is Not Null)) " &
                                                      "ORDER BY [D:農地Info].所有者ID, [D:農地Info].大字ID;")
        TBL転用農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D_転用農地].所有者ID, [D:個人Info].氏名, V_大字.大字, V_小字.小字, [D_転用農地].地番, V_現況地目.名称, [D_転用農地].実面積, [D_転用農地].農業振興地域, [D_転用農地].農振法区分, [D_転用農地].利用状況調査農地法, [D_転用農地].利用意向調査日, [D_転用農地].利用意向意思表明日, [D_転用農地].利用意向意向内容区分 " &
                                                      "FROM ((([D_転用農地] LEFT JOIN V_大字 ON [D_転用農地].大字ID = V_大字.ID) LEFT JOIN V_小字 ON [D_転用農地].小字ID = V_小字.ID) LEFT JOIN V_現況地目 ON [D_転用農地].現況地目 = V_現況地目.ID) LEFT JOIN [D:個人Info] ON [D_転用農地].所有者ID = [D:個人Info].ID " &
                                                      "WHERE ((([D_転用農地].利用意向調査日) Is Not Null)) " &
                                                      "ORDER BY [D_転用農地].所有者ID, [D_転用農地].大字ID;")
        TBL農地.Merge(TBL転用農地)
        TBL削除農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D_削除農地].所有者ID, [D:個人Info].氏名, V_大字.大字, V_小字.小字, [D_削除農地].地番, V_現況地目.名称, [D_削除農地].実面積, [D_削除農地].農業振興地域, [D_削除農地].農振法区分, [D_削除農地].利用状況調査農地法, [D_削除農地].利用意向調査日, [D_削除農地].利用意向意思表明日, [D_削除農地].利用意向意向内容区分 " &
                                                      "FROM ((([D_削除農地] LEFT JOIN V_大字 ON [D_削除農地].大字ID = V_大字.ID) LEFT JOIN V_小字 ON [D_削除農地].小字ID = V_小字.ID) LEFT JOIN V_現況地目 ON [D_削除農地].現況地目 = V_現況地目.ID) LEFT JOIN [D:個人Info] ON [D_削除農地].所有者ID = [D:個人Info].ID " &
                                                      "WHERE ((([D_削除農地].利用意向調査日) Is Not Null)) " &
                                                      "ORDER BY [D_削除農地].所有者ID, [D_削除農地].大字ID;")
        TBL農地.Merge(TBL削除農地)

        '/*****市町村コード、市町村名の取得*****/
        Dim Int市町村CD As Integer = Val(SysAD.DB(sLRDB).DBProperty("市町村ID").ToString)
        Dim St市町村名 As String = SysAD.DB(sLRDB).DBProperty("市町村名")

        TBL農地.Columns.Add("農地整理番号", GetType(Integer))
        TBL農地.Columns.Add("意向調査年度", GetType(Integer))
        TBL農地.Columns.Add("市町村CD", GetType(Integer), Int市町村CD)
        TBL農地.Columns.Add("市町村名", GetType(String), "'" & St市町村名 & "'")
        TBL農地.Columns.Add("所有者等番号", GetType(Integer))
        TBL農地.Columns.Add("権利種類", GetType(Integer))
        TBL農地.Columns.Add("農用地区分", GetType(String))
        TBL農地.Columns.Add("現地調査状況", GetType(String))
    End Sub

    Private Sub CreatemvarGrid(ByRef pDGrid As HimTools2012.controls.DataGridViewWithDataView)
        pDGrid = New HimTools2012.controls.DataGridViewWithDataView
        AddCol(pDGrid, "農地整理番号", "農地整理番号", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "意向調査年度", "意向調査年度", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "市町村CD", "市町村CD", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "市町村名", "市町村名", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "所有者等番号", "所有者等番号", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "氏名", "氏名", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "権利種類", "権利種類", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "大字", "大字", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "字", "小字", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "地番", "地番", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "現況地目", "名称", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "面積", "実面積", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "農用地区分", "農用地区分", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "現地調査状況", "現地調査状況", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "回答日", "利用意向意思表明日", New DataGridViewTextBoxColumn)
        AddCol(pDGrid, "回答内容", "利用意向意向内容区分", New DataGridViewTextBoxColumn)

        pDGrid.AutoGenerateColumns = False
        pDGrid.AllowUserToAddRows = False
        pDGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
    End Sub

    Private Sub AddCol(ByRef pDGrid As HimTools2012.controls.DataGridViewWithDataView, ByVal sHeader As String, ByVal sData As String, ByVal pCol As DataGridViewColumn)
        pCol.HeaderText = sHeader
        pCol.DataPropertyName = sData
        pDGrid.Columns.Add(pCol)
    End Sub

    Private pView As DataView
    Private Sub mvar検索開始_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar検索開始.Click
        If mvarStart.Value < mvarEnd.Value Then
            Dim sWhere As String = String.Format("[利用意向調査日]>=#{0}/{1}/{2}# AND [利用意向調査日]<=#{3}/{4}/{5}#",
                                                mvarStart.Value.Month, mvarStart.Value.Day, mvarStart.Value.Year, mvarEnd.Value.Month, mvarEnd.Value.Day, mvarEnd.Value.Year)

            pView = New DataView(TBL農地, sWhere, "[所有者ID]", DataViewRowState.CurrentRows)

            Dim RowCount As Integer = 1
            Dim SaveUserNo As Integer = 0
            Dim UserNoCount As Integer = 0
            For Each pRow As DataRowView In pView
                pRow.Item("農地整理番号") = RowCount
                pRow.Item("意向調査年度") = Replace(Replace(和暦Format(mvarStart.Text, "gggyy"), "平成", ""), "令和", "")
                If SaveUserNo = 0 OrElse SaveUserNo <> pRow.Item("所有者ID") Then
                    UserNoCount += 1
                    SaveUserNo = pRow.Item("所有者ID")
                End If
                pRow.Item("所有者等番号") = UserNoCount
                pRow.Item("権利種類") = 1

                Select Case pRow.Item("名称")
                    Case "田", "畑"
                    Case Else : pRow.Item("名称") = ""
                End Select

                Select Case Val(pRow.Item("農振法区分").ToString)
                    Case enum農振法区分.農用地区域 : pRow.Item("農用地区分") = 1
                    Case enum農振法区分.農振地域 : pRow.Item("農用地区分") = 2
                    Case Else
                        Select Case Val(pRow.Item("農業振興地域").ToString)
                            Case enum農業振興地域.農用地内 : pRow.Item("農用地区分") = 1
                            Case enum農業振興地域.農用地外 : pRow.Item("農用地区分") = 2
                            Case Else : pRow.Item("農用地区分") = ""
                        End Select
                End Select

                Select Case Val(pRow.Item("利用状況調査農地法").ToString)
                    Case 1 To 2 : pRow.Item("現地調査状況") = pRow.Item("利用状況調査農地法")
                    Case Else : pRow.Item("現地調査状況") = ""
                End Select

                Select Case Val(pRow.Item("利用意向意向内容区分").ToString)
                    Case 1 : pRow.Item("利用意向意向内容区分") = 4
                    Case 2 : pRow.Item("利用意向意向内容区分") = 1
                    Case 3 : pRow.Item("利用意向意向内容区分") = 2
                    Case 4 : pRow.Item("利用意向意向内容区分") = 3
                    Case 5 : pRow.Item("利用意向意向内容区分") = 5
                    Case Else : pRow.Item("利用意向意向内容区分") = 6
                End Select

                RowCount += 1
            Next

            mvarGrid.SetDataView(pView.ToTable, sWhere, "[所有者ID]")
        End If
    End Sub

    Private Sub mvar帳票出力_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar帳票出力.Click
        If IO.File.Exists(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\利用意向調査実施農地状況一覧.xml") Then
            Dim sXML As String = HimTools2012.TextAdapter.LoadTextFile(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\利用意向調査実施農地状況一覧.xml")
            Dim pXMLSS As New HimTools2012.Excel.XMLSS2003.CXMLSS2003(sXML)
            Dim pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet = pXMLSS.WorkBook.WorkSheets.Items("農地状況一覧")
            Dim pLoopRow As New HimTools2012.Excel.XMLSS2003.XMLLoopRows(pSheet, "{No}")
            Dim nLoop As Integer = -1

            For Each pRow As DataRowView In pView
                Dim pXRow As HimTools2012.Excel.XMLSS2003.XMLSSRow = Nothing

                If nLoop = -1 Then
                    pXRow = pSheet.FindRowInstrText("{No}")(0)
                Else

                    For Each pCXRow As HimTools2012.Excel.XMLSS2003.XMLSSRow In pLoopRow
                        pXRow = pCXRow.CopyRow

                        pSheet.Table.Rows.InsertRow(pLoopRow.InsetRow, pXRow)
                        pLoopRow.InsetRow += 1
                    Next
                End If
                nLoop += 1

                With pXRow
                    .ValueReplace("{No}", pRow.Item("農地整理番号").ToString)
                    .ValueReplace("{年度}", pRow.Item("意向調査年度").ToString)
                    .ValueReplace("{市町村CD}", pRow.Item("市町村CD").ToString)
                    .ValueReplace("{市町村名}", pRow.Item("市町村名").ToString)
                    .ValueReplace("{所有者No}", pRow.Item("所有者等番号").ToString)
                    .ValueReplace("{氏名}", pRow.Item("氏名").ToString)
                    .ValueReplace("{権利種類}", pRow.Item("権利種類").ToString)
                    .ValueReplace("{大字}", pRow.Item("大字").ToString)
                    .ValueReplace("{小字}", pRow.Item("小字").ToString)
                    .ValueReplace("{地番}", pRow.Item("地番").ToString)
                    .ValueReplace("{地目}", pRow.Item("名称").ToString)
                    .ValueReplace("{面積}", DecConv(CDec(pRow.Item("実面積").ToString)))
                    .ValueReplace("{区分}", pRow.Item("農用地区分").ToString)
                    .ValueReplace("{調査状況}", pRow.Item("現地調査状況").ToString)
                    .ValueReplace("{再生可否}", "○")    '/***初期値***/
                    .ValueReplace("{借受希望}", "×")
                    .ValueReplace("{適正賃料}", "○")
                    .ValueReplace("{集積適地}", IIf(Val(pRow.Item("農用地区分").ToString) = 1, "○", IIf(Val(pRow.Item("農用地区分").ToString) = 2, "×", "")))
                    .ValueReplace("{適否結果○×}", "―")
                    .ValueReplace("{適否結果×○}", "―")
                    .ValueReplace("{活用可能性}", "×")
                    If Not IsDBNull(pRow.Item("利用意向意思表明日")) Then
                        .ValueReplace("{回答日}", 和暦Format(pRow.Item("利用意向意思表明日"), "gyy.MM.dd"))
                    Else
                        .ValueReplace("{回答日}", "")
                    End If
                    .ValueReplace("{回答内容}", pRow.Item("利用意向意向内容区分").ToString)
                End With
            Next

            Dim sFile As String = My.Computer.FileSystem.SpecialDirectories.Desktop & "\xx利用意向調査実施農地状況一覧.xml"
            Try
                HimTools2012.TextAdapter.SaveTextFile(sFile, pXMLSS.OutPut(False))
                MsgBox("エクセルをデスクトップに作成しました。", MsgBoxStyle.Information)
                SysAD.ShowFolder(sFile)
            Catch ex As Exception

            End Try

        End If
    End Sub
    Private Function DecConv(ByVal pDec As Decimal) As String
        If Fix(pDec) = pDec Then
            Return Val(pDec).ToString("#,##0")
        Else
            Return Val(pDec).ToString("#,##0.##") '小数点第2位まで表示
        End If

        Return pDec
    End Function

    Private Sub mvarExcel出力_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarExcel出力.Click
        mvarGrid.ToExcel()
    End Sub
End Class
