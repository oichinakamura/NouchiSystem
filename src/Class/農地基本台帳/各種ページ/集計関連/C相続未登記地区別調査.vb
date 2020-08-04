'南種子町役場…データが壊れてるかも？（最適化の必要あり）
'   ※IDがNullのレコードがある気がする…

Public Class C相続未登記地区別調査
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Public mvarTabCtrl As HimTools2012.controls.TabControlBase
    Private WithEvents mvarBtnSearch As ToolStripButton
    Private mvarTSC明細 As ToolStripContainer
    Private mvarTS明細 As ToolStrip
    Private mvarGrid明細 As New HimTools2012.controls.DataGridViewWithDataView
    Private WithEvents mvarExcel明細 As ToolStripButton
    Private mvarTSC集計 As ToolStripContainer
    Private mvarTS集計 As ToolStrip
    Private mvarGrid集計 As New HimTools2012.controls.DataGridViewWithDataView
    Private WithEvents mvarExcel集計 As ToolStripButton
    Private mvarTSProg As New ToolStripProgressBar
    Private mvarTSLabel As New ToolStripLabel

    Private s市町村 As String = ""
    Public Sub New()
        MyBase.New(True, True, "相続未登記地区別調査", "相続未登記地区別調査")

        mvarTSProg.Visible = False
        mvarTSLabel.Visible = False

        mvarBtnSearch = New ToolStripButton("データ読み込み")
        Me.ToolStrip.Items.AddRange({mvarBtnSearch, mvarTSProg, mvarTSLabel})

        mvarTabCtrl = New HimTools2012.controls.TabControlBase
        mvarTabCtrl.Dock = DockStyle.Fill

        Me.Controls.Add(mvarTabCtrl)

        Dim pPage明細 As New TabPage("明細")
        Dim pPage集計 As New TabPage("集計")

        mvarTabCtrl.TabPages.Add(pPage明細)
        mvarTabCtrl.TabPages.Add(pPage集計)

        mvarTSC明細 = New ToolStripContainer
        pPage明細.Controls.Add(mvarTSC明細)
        mvarTSC明細.Dock = DockStyle.Fill
        mvarTS明細 = New ToolStrip
        mvarTSC明細.TopToolStripPanel.Controls.Add(mvarTS明細)
        mvarExcel明細 = New ToolStripButton("Excel出力")
        mvarTS明細.Items.Add(mvarExcel明細)
        mvarTSC明細.ContentPanel.Controls.Add(mvarGrid明細)
        mvarGrid明細.Dock = DockStyle.Fill
        mvarGrid明細.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        mvarTSC集計 = New ToolStripContainer
        pPage集計.Controls.Add(mvarTSC集計)
        mvarTSC集計.Dock = DockStyle.Fill
        mvarTS集計 = New ToolStrip
        mvarTSC集計.TopToolStripPanel.Controls.Add(mvarTS集計)
        mvarExcel集計 = New ToolStripButton("Excel出力")
        mvarTS集計.Items.Add(mvarExcel集計)
        mvarTSC集計.ContentPanel.Controls.Add(mvarGrid集計)
        mvarGrid集計.Dock = DockStyle.Fill
        mvarGrid集計.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

        s市町村 = SysAD.DB(sLRDB).DBProperty("市町村名")
    End Sub

    Private pTBL農地 As DataTable
    Private pTBL対象行政区 As DataTable
    Private pTBL住記情報 As DataTable
    Private pTBL住民区分 As DataTable
    Private Sub mvarBtnSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarBtnSearch.Click
        mvarTSProg.Visible = True
        mvarTSLabel.Visible = True
        mvarBtnSearch.Visible = False

        '田・畑・樹・採いずれかに面積が入っているものを農地として扱う
        pTBL農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT V_農地.ID, V_大字.名称 AS 大字名, V_農地.土地所在, V_農地.登記簿面積, V_農地.実面積, V_農地.所有者ID, V_農地.利用状況調査農地法, V_農地.利用状況調査荒廃 FROM V_農地 LEFT JOIN V_大字 ON V_農地.大字ID = V_大字.ID WHERE (((V_農地.ID) Is Not Null) AND ((V_大字.名称) Is Not Null) AND ((V_農地.所有者ID) Is Not Null) AND ((V_農地.田面積)>0) AND ((V_農地.大字ID)<>-1) AND ((V_農地.農地状況)<=19)) OR (((V_農地.ID) Is Not Null) AND ((V_大字.名称) Is Not Null) AND ((V_農地.所有者ID) Is Not Null) AND ((V_農地.大字ID)<>-1) AND ((V_農地.農地状況)<=19) AND ((V_農地.畑面積)>0)) OR (((V_農地.ID) Is Not Null) AND ((V_大字.名称) Is Not Null) AND ((V_農地.所有者ID) Is Not Null) AND ((V_農地.大字ID)<>-1) AND ((V_農地.農地状況)<=19) AND ((V_農地.樹園地)>0)) OR (((V_農地.ID) Is Not Null) AND ((V_大字.名称) Is Not Null) AND ((V_農地.所有者ID) Is Not Null) AND ((V_農地.大字ID)<>-1) AND ((V_農地.農地状況)<=19) AND ((V_農地.採草放牧面積)>0)) ORDER BY V_農地.大字ID, V_農地.所有者ID;")
        'pTBL住記情報 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT M_住民情報.ID, CDbl([ID]) AS ReID, M_住民情報.世帯No, M_住民情報.氏名, M_住民情報.住所, M_住民情報.住民区分, M_住民情報.生年月日 FROM(M_住民情報) WHERE (((M_住民情報.ID) Is Not Null) AND ((CDbl([ID])) Is Not Null) AND ((M_住民情報.世帯No) Is Not Null));")
        pTBL住記情報 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT M_住民情報.ID, M_住民情報.世帯No, M_住民情報.氏名, M_住民情報.住所, M_住民情報.住民区分, M_住民情報.生年月日, M_住民情報.行政区 AS 行政区ID, V_行政区.名称 AS 行政区 FROM M_住民情報 LEFT JOIN V_行政区 ON M_住民情報.行政区 = V_行政区.ID WHERE (((M_住民情報.ID) Is Not Null));")
        pTBL住記情報.PrimaryKey = {pTBL住記情報.Columns("ID")}
        'pTBL住記情報.PrimaryKey = {pTBL住記情報.Columns("ReID")}
        pTBL住民区分 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [V_住民区分]")

        Dim DSet As New DataSet
        DSet.Tables.AddRange({pTBL農地, pTBL住記情報, pTBL住民区分})

        DSet.Relations.Add(New DataRelation("住民区分情報", pTBL住民区分.Columns("ID"), pTBL住記情報.Columns("住民区分"), False))
        pTBL住記情報.Columns.Add(New DataColumn("住民区分名", GetType(String), "Parent(住民区分情報).名称"))

        CreateTBL対象行政区()
        CreateTBL集計()

        mvarTSProg.Value = 0
        mvarTSProg.Maximum = pTBL農地.Rows.Count
        For Each pRow As DataRow In pTBL農地.Rows
            If Not IsDBNull(pRow.Item("所有者ID")) Then
                Set対象行政区テーブル(pRow)
            Else
                'Stop
            End If

            mvarTSProg.Increment(1)
            mvarTSLabel.Text = "データ読み込み中...(" & mvarTSProg.Value & "/" & mvarTSProg.Maximum & ")"
            My.Application.DoEvents()
        Next

        mvarGrid明細.SetDataView(pTBL対象行政区, "", "[行政区ID]")
        mvarGrid集計.SetDataView(pTBL集計, "", "")

        mvarTSProg.Visible = False
        mvarTSLabel.Visible = False
        mvarBtnSearch.Visible = True
    End Sub

    Private Sub CreateTBL対象行政区()
        pTBL対象行政区 = New DataTable
        With pTBL対象行政区
            .Columns.Add("行政区ID", GetType(Decimal))
            .Columns.Add("行政区", GetType(String))
            .Columns.Add("筆数計", GetType(Integer))
            .Columns.Add("面積計", GetType(Decimal))
            .Columns.Add("うち遊休農地筆数計", GetType(Integer))
            .Columns.Add("うち遊休農地面積計", GetType(Decimal))
            .Columns.Add("所有者ID", GetType(Decimal))
            .Columns.Add("借受者ID", GetType(Decimal))

            .Columns.Add("要因区分", GetType(String))
            .PrimaryKey = { .Columns("ID")}
        End With
    End Sub

    Private Sub Set対象行政区テーブル(ByRef pRow As DataRow)
        Dim pFRow住記情報 As DataRow = pTBL住記情報.Rows.Find(pRow.Item("所有者ID"))

        If pFRow住記情報 IsNot Nothing Then
            Dim pAddRow As DataRow = pTBL対象行政区.NewRow

            With pAddRow
                .Item("行政区ID") = pFRow住記情報.Item("行政区ID")
                .Item("行政区") = pFRow住記情報.Item("行政区")
                .Item("筆数計") = 1
                .Item("面積計") = Val(pRow.Item("登記簿面積").ToString)
                .Item("うち遊休農地筆数計") = IIf(Check遊休(pRow), 1, 0)
                .Item("うち遊休農地面積計") = IIf(Check遊休(pRow), Val(pRow.Item("登記簿面積").ToString), 0)
                .Item("要因区分") = Check住民区分(pFRow住記情報)
            End With

            pTBL対象行政区.Rows.Add(pAddRow)

            Allocation要因(pAddRow, pRow)
        ElseIf pFRow住記情報 Is Nothing Then
            Dim pAddRow As DataRow = pTBL対象行政区.NewRow

            With pAddRow
                .Item("行政区ID") = 9999999999
                .Item("行政区") = "行政区なし"
                .Item("筆数計") = 1
                .Item("面積計") = Val(pRow.Item("登記簿面積").ToString)
                .Item("うち遊休農地筆数計") = IIf(Check遊休(pRow), 1, 0)
                .Item("うち遊休農地面積計") = IIf(Check遊休(pRow), Val(pRow.Item("登記簿面積").ToString), 0)
                .Item("要因区分") = "②所有者が市町村外在住等"
            End With

            pTBL対象行政区.Rows.Add(pAddRow)

            Allocation要因(pAddRow, pRow)
        Else
            Stop
        End If
    End Sub

    Private Function Check住民区分(ByVal pRow As DataRow)
        '氏名に「○○　外○名」と含む場合の処理
        If Strings.Right(pRow.Item("氏名").ToString, 1) = "名" Then
            If IsNumeric(Strings.Left(Strings.Right(pRow.Item("氏名").ToString, 2).ToString, 1)) Then
                Return "⑧共有者状況不明"
            End If
        End If

        Select Case pRow.Item("住民区分名").ToString
            Case "死亡", "死亡者", "死亡所有者", "死亡所有", "死亡(世帯主)"
                Return "①相続未登記"
            Case "市外住民", "市外居住者", "町外住民", "町外者", "転出", "転出者", "転出確定者", "転出確認", "転居出者", "除票（住）", "その他除票者", "戸籍転籍除籍者", "外国人転出確定者", "市外外国人", "住登外", "住登外（住）", "住登外（法人）", "住登外（共有者）", "住登外個人", "住登外法人", "住民喪失後住登外"
                Return "②所有者が市町村外在住等"
            Case "未登録住民"
                Return "②所有者が市町村外在住等"
            Case "共有名義", "共有者", "共有"
                Return "⑧共有者状況不明"
            Case "不明住民"
                Return "⑨不明"
            Case ""
                If IsDate(pRow.Item("生年月日").ToString) Then
                    If Year(pRow.Item("生年月日").ToString) <= 1900 Then
                        Return "①相続未登記"
                    Else
                        Return "⑨不明"
                    End If
                Else
                    Return "⑨不明"
                End If
            Case Else
                'Debug.Print(pRow.Item("住民区分名").ToString)
                Return ""
        End Select
    End Function

    Private Sub Allocation要因(ByRef pRowKey As DataRow, ByRef pRow As DataRow)
        Select Case pRowKey.Item("要因区分").ToString
            Case "①相続未登記" : Set集計テーブル(pRow, 1)
            Case "②所有者が市町村外在住等" : Set集計テーブル(pRow, 2)
            Case "⑧共有者状況不明" : Set集計テーブル(pRow, 8)
            Case "⑨不明" : Set集計テーブル(pRow, 9)
            Case Else : Set集計テーブル(pRow, 0)
        End Select
    End Sub

    Private pTBL集計 As DataTable
    Private Sub CreateTBL集計()
        pTBL集計 = New DataTable
        With pTBL集計
            .Columns.Add("市町村名", GetType(String))
            .Columns.Add("台帳全体筆数", GetType(Integer))
            .Columns.Add("台帳全体面積", GetType(Decimal))
            .Columns.Add("相続未登記筆数", GetType(Integer))
            .Columns.Add("相続未登記面積", GetType(Decimal))
            .Columns.Add("相続未登記遊休筆数", GetType(Integer))
            .Columns.Add("相続未登記遊休面積", GetType(Decimal))
            .Columns.Add("市町村外筆数", GetType(Integer))
            .Columns.Add("市町村外面積", GetType(Decimal))
            .Columns.Add("市町村外遊休筆数", GetType(Integer))
            .Columns.Add("市町村外遊休面積", GetType(Decimal))
            .Columns.Add("不明筆数", GetType(Integer))
            .Columns.Add("不明面積", GetType(Decimal))
            .Columns.Add("共有者状況不明筆数", GetType(Integer))
            .Columns.Add("共有者状況不明面積", GetType(Decimal))
            .PrimaryKey = { .Columns("市町村名")}
        End With
    End Sub

    Private Sub Set集計テーブル(ByRef pRow As DataRow, ByVal p要因 As Integer)
        Dim pFindRow As DataRow = pTBL集計.Rows.Find(s市町村)

        If pFindRow IsNot Nothing Then
            With pFindRow
                .Item("台帳全体筆数") += 1
                .Item("台帳全体面積") += Val(pRow.Item("登記簿面積").ToString)
                .Item("相続未登記筆数") += IIf(p要因 = 1, 1, 0)
                .Item("相続未登記面積") += IIf(p要因 = 1, Val(pRow.Item("登記簿面積").ToString), 0)
                .Item("相続未登記遊休筆数") += IIf(Check遊休(pRow), IIf(p要因 = 1, 1, 0), 0)
                .Item("相続未登記遊休面積") += IIf(Check遊休(pRow), IIf(p要因 = 1, Val(pRow.Item("登記簿面積").ToString), 0), 0)
                .Item("市町村外筆数") += IIf(p要因 = 2, 1, 0)
                .Item("市町村外面積") += IIf(p要因 = 2, Val(pRow.Item("登記簿面積").ToString), 0)
                .Item("市町村外遊休筆数") += IIf(Check遊休(pRow), IIf(p要因 = 2, 1, 0), 0)
                .Item("市町村外遊休面積") += IIf(Check遊休(pRow), IIf(p要因 = 2, Val(pRow.Item("登記簿面積").ToString), 0), 0)
                .Item("不明筆数") += IIf(p要因 = 9, 1, 0)
                .Item("不明面積") += IIf(p要因 = 9, Val(pRow.Item("登記簿面積").ToString), 0)
                .Item("共有者状況不明筆数") += IIf(p要因 = 8, 1, 0)
                .Item("共有者状況不明面積") += IIf(p要因 = 8, Val(pRow.Item("登記簿面積").ToString), 0)
            End With
        Else
            Dim pAddRow As DataRow = pTBL集計.NewRow

            With pAddRow
                .Item("市町村名") = s市町村
                .Item("台帳全体筆数") = 1
                .Item("台帳全体面積") = Val(pRow.Item("登記簿面積").ToString)
                .Item("相続未登記筆数") = IIf(p要因 = 1, 1, 0)
                .Item("相続未登記面積") = IIf(p要因 = 1, Val(pRow.Item("登記簿面積").ToString), 0)
                .Item("相続未登記遊休筆数") = IIf(Check遊休(pRow), IIf(p要因 = 1, 1, 0), 0)
                .Item("相続未登記遊休面積") = IIf(Check遊休(pRow), IIf(p要因 = 1, Val(pRow.Item("登記簿面積").ToString), 0), 0)
                .Item("市町村外筆数") = IIf(p要因 = 2, 1, 0)
                .Item("市町村外面積") = IIf(p要因 = 2, Val(pRow.Item("登記簿面積").ToString), 0)
                .Item("市町村外遊休筆数") = IIf(Check遊休(pRow), IIf(p要因 = 2, 1, 0), 0)
                .Item("市町村外遊休面積") = IIf(Check遊休(pRow), IIf(p要因 = 2, Val(pRow.Item("登記簿面積").ToString), 0), 0)
                .Item("不明筆数") = IIf(p要因 = 9, 1, 0)
                .Item("不明面積") = IIf(p要因 = 9, Val(pRow.Item("登記簿面積").ToString), 0)
                .Item("共有者状況不明筆数") = IIf(p要因 = 8, 1, 0)
                .Item("共有者状況不明面積") = IIf(p要因 = 8, Val(pRow.Item("登記簿面積").ToString), 0)
            End With

            pTBL集計.Rows.Add(pAddRow)
        End If
    End Sub

    ''' <summary>
    ''' 遊休農地にはB分類は含まない
    ''' </summary>
    ''' <param name="pRow"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function Check遊休(ByVal pRow As DataRow)
        Dim sResult As Boolean = False

        Select Case Val(pRow.Item("利用状況調査農地法").ToString)
            Case 1, 2 '農地法第32条第1項1号、農地法第32条第1項2号
                sResult = True
        End Select

        Select Case Val(pRow.Item("利用状況調査荒廃").ToString)
            Case 1 'A分類のみ
                sResult = True
        End Select

        Return sResult
    End Function

    Private Sub mvarExcel明細_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarExcel明細.Click
        mvarGrid明細.ToExcel()
    End Sub

    Private Sub mvarExcel集計_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarExcel集計.Click
        mvarGrid集計.ToExcel()
    End Sub
End Class
