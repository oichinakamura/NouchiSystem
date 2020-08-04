Imports ClosedXML.Excel
Imports HimTools2012

Public Class CTabPage標準フォーマット出力
    Inherits C標準フォーマット

    Private mvarTabCtrl As controls.TabControlBase
    Private DTP開始日 As ToolStripDateTimePickerWithlabel
    Private DTP終了日 As ToolStripDateTimePickerWithlabel
    Private WithEvents mvar絞込条件 As ToolStripComboBox
    Private WithEvents mvar検索開始 As ToolStripButton
    Private mvarTSProg As New ToolStripProgressBar
    Private mvarTSLabel As New ToolStripLabel
    Private WithEvents mvar出力開始 As ToolStripButton

    Private mvarGrid申請 As controls.DataGridViewWithDataView

    Public Sub New()
        mvarTabCtrl = New controls.TabControlBase
        mvarTabCtrl.Dock = DockStyle.Fill

        ControlPanel.Add(mvarTabCtrl)

        With New 議案書作成パラメータ
            DTP開始日 = New ToolStripDateTimePickerWithlabel("検索範囲")
            DTP開始日.Value = .開始年月日

            DTP終了日 = New ToolStripDateTimePickerWithlabel("～")
            DTP終了日.Value = .終了年月日
        End With

        mvar絞込条件 = New ToolStripComboBox
        mvar絞込条件.Items.AddRange(New String() {"全対象", "農地法3条", "農地法4条", "農地法5条", "基盤強化法", "解約", "非農地判断"})
        mvar絞込条件.Text = "全対象"

        mvar検索開始 = New ToolStripButton
        mvar検索開始.Text = "データ読込"

        mvarTSProg = New ToolStripProgressBar
        mvarTSProg.Visible = False

        mvarTSLabel = New ToolStripLabel
        mvarTSLabel.Visible = False

        mvar出力開始 = New ToolStripButton
        mvar出力開始.Text = "標準フォーマット出力"

        Me.ToolStrip.Items.AddRange({DTP開始日, DTP終了日, mvar絞込条件, mvar検索開始, mvarTSProg, mvarTSLabel, New ToolStripSeparator, mvar出力開始})

        mvarGrid申請 = New controls.DataGridViewWithDataView
        CreateGrid申請()

        Dim pPage01 As New controls.CTabPageWithToolStrip(False, True, "申請", "申請一覧")
        pPage01.ControlPanel.Add(mvarGrid申請)
        mvarGrid申請.Createエクセル出力Ctrl(pPage01.ToolStrip)
        mvarTabCtrl.AddPage(pPage01)

        '// 申請以外のデータを読み込む
        DSet = New DataSet
        TBL農家 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:個人Info].ID, [D:個人Info].氏名, [D:個人Info].住所, [D:個人Info].職業 FROM [D:個人Info];")
        TBL農家.PrimaryKey = New DataColumn() {TBL農家.Columns("ID")}
        TBL農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:農地Info].ID, [D:農地Info].大字ID, V_大字.名称 AS 大字名, [D:農地Info].小字ID, V_小字.名称 AS 小字名, [D:農地Info].地番, [D:農地Info].一部現況, [D:農地Info].現況地目, [D:農地Info].所有者ID FROM ([D:農地Info] LEFT JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) LEFT JOIN V_小字 ON [D:農地Info].小字ID = V_小字.ID;")
        TBL転用農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D_転用農地].ID, [D_転用農地].大字ID, V_大字.名称 AS 大字名, [D_転用農地].小字ID, V_小字.名称 AS 小字名, [D_転用農地].地番, [D_転用農地].一部現況, [D_転用農地].現況地目, [D_転用農地].所有者ID FROM ([D_転用農地] LEFT JOIN V_大字 ON [D_転用農地].大字ID = V_大字.ID) LEFT JOIN V_小字 ON [D_転用農地].小字ID = V_小字.ID;")

        TBL農地.Merge(TBL転用農地)
        TBL農地.PrimaryKey = New DataColumn() {TBL農地.Columns("ID")}

        Dim TBL現況地目マスタ As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [V_現況地目];")
        DSet.Tables.AddRange(New DataTable() {TBL農家, TBL農地, TBL現況地目マスタ})

        DSet.Relations.Add("現況地目", TBL現況地目マスタ.Columns("ID"), TBL農地.Columns("現況地目"), False)
        TBL農地.Columns.Add("現況地目名", GetType(String), "Parent(現況地目).名称")

        ColumnCheck(TBL農地, "本番区分", GetType(String))
        ColumnCheck(TBL農地, "本番", GetType(Integer))
        ColumnCheck(TBL農地, "枝番区分", GetType(String))
        ColumnCheck(TBL農地, "枝番", GetType(Integer))
        ColumnCheck(TBL農地, "孫番区分", GetType(String))
        ColumnCheck(TBL農地, "孫番", GetType(Integer))

        For Each pRow As DataRow In TBL農地.Rows
            Conv地番(pRow)
        Next
    End Sub

    Private Sub CreateGrid申請()
        With mvarGrid申請
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .AutoGenerateColumns = False
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            .AddColumnText("ID", "ID", "ID", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("法令", "法令", "法令", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("名称", "名称", "名称", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("許可番号", "受付/許可番号", "許可番号", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("許可年月日", "受付/許可年月日", "許可年月日", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("借受人住所", "借受人住所", "借受人住所", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("借受人氏名", "借受人氏名", "借受人氏名", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("転用目的", "転用目的", "転用目的", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("農地区分", "農地区分", "農地区分", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("小作料", "小作料", "小作料", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("小作料単位", "小作料単位", "小作料単位", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("筆数計", "筆数計", "筆数計", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("面積計", "面積計", "面積計", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
        End With
    End Sub

    Private Sub mvar検索開始_Click(sender As Object, e As EventArgs) Handles mvar検索開始.Click
        Try
            'mvarTSProg.Visible = True
            'mvarTSLabel.Visible = True
            'mvar検索開始.Visible = False

            Dim s法令 As String = ""
            Select Case mvar絞込条件.Text
                Case "全対象" : s法令 = "30,31,311,40,50,51,52,60,61,62,602,600,180,200,210"
                Case "農地法3条" : s法令 = "30,31,311"
                Case "農地法4条" : s法令 = "40"
                Case "農地法5条" : s法令 = "50,51,52"
                Case "基盤強化法" : s法令 = "60,61,62"
                Case "非農地判断" : s法令 = "602,600"
                Case "解約" : s法令 = "180,200,210"
            End Select

            TBL申請 = SysAD.DB(sLRDB).GetTableBySqlSelect(String.Format("SELECT * FROM D_申請 WHERE [法令] IN({0}) AND [状態]=2 AND [許可年月日] >= #{1}# AND [許可年月日] <= #{2}# ORDER BY [受付年月日], [受付番号];", s法令, DTP開始日.Value.ToShortDateString, DTP終了日.Value.ToShortDateString))
            mvarGrid申請.SetDataView(TBL申請, "", "")

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private AppNo As Integer = 1
    Private rowNo As Integer = 4
    Private AppNoATo中 As Integer = 1
    Private rowNoATo中 As Integer = 4
    Private AppNo中ToB As Integer = 1
    Private rowNo中ToB As Integer = 4
    Private Sub mvar出力開始_Click(sender As Object, e As EventArgs) Handles mvar出力開始.Click
        Dim sFilePath As String = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "RPA標準フォーマット.xlsx")
        Dim sDirectory As String = ""

        If System.IO.File.Exists(sFilePath) AndAlso Not TBL申請 Is Nothing Then
            Using book As New XLWorkbook(sFilePath)
                Select Case mvar絞込条件.Text
                    Case "全対象"
                        Dim sheet = book.Worksheet("3条")
                        農地法第3条(sheet)
                        sheet = book.Worksheet("4条")
                        農地法第4条(sheet)
                        sheet = book.Worksheet("5条")
                        農地法第5条(sheet)
                    Case "農地法3条"
                        Dim sheet = book.Worksheet("3条")
                        農地法第3条(sheet)
                    Case "農地法4条"
                        Dim sheet = book.Worksheet("4条")
                        農地法第4条(sheet)
                    Case "農地法5条"
                        Dim sheet = book.Worksheet("5条")
                        農地法第5条(sheet)
                    Case "基盤強化法"
                        基盤強化法第19条(book)
                    Case "非農地判断"
                        Dim sheet = book.Worksheet("非農地判断")
                        非農地判断(sheet)
                    Case "解約"
                        Dim sheet = book.Worksheet("18条解約（許可）")
                        農地法第18条(sheet)
                End Select

                Dim savePath As String = 名前を付けて保存("RPA標準フォーマット", OutputType.xlsx, Now.ToString("yyyyMMdd"))
                If savePath <> "" Then
                    book.SaveAs(savePath)
                    sDirectory = System.IO.Path.GetDirectoryName(savePath)

                    MsgBox("終了しました。")
                    Process.Start(sDirectory)
                End If
            End Using
        End If
    End Sub

    Private Sub 農地法第3条(ByRef targetSheet As IXLWorksheet)
        Dim sData(,) As Object

        Select Case mvar絞込条件.Text
            Case "全対象"
                Dim pView As DataView = New DataView(TBL申請, "[法令] In (30,31,311)", "", DataViewRowState.CurrentRows)
                ReDim sData(pView.Count, 61)   '配列の次元を変更
            Case Else
                ReDim sData(TBL申請.Rows.Count, 61)
        End Select

        AppNo = 1
        rowNo = 4

        For Each pRow As DataRow In TBL申請.Rows
            Select Case Val(pRow.Item("法令").ToString)
                Case 30, 31, 311
                    Dim s農地List As String() = Split(pRow.Item("農地リスト").ToString(), ";")
                    For Each s農地ID As String In s農地List
                        sData(0, 0) = AppNo
                        '// 申請種別
                        sData(0, 1) = 1
                        Select Case Val(pRow.Item("法令").ToString)
                            Case 30, 311 : sData(0, 2) = 1
                            Case 31 : sData(0, 2) = 2
                        End Select
                        sData(0, 3) = ""
                        sData(0, 4) = Val(pRow.Item("受付番号").ToString)
                        '// 受け人選択
                        sData(0, 5) = Get農家情報(Val(pRow.Item("申請者B").ToString), PersonInfo.Name)
                        sData(0, 6) = Get農家情報(Val(pRow.Item("申請者B").ToString), PersonInfo.Address)
                        sData(0, 7) = Get農家情報(Val(pRow.Item("申請者B").ToString), PersonInfo.Job)
                        sData(0, 8) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Name)
                        sData(0, 9) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Address)
                        '// 対象地選択
                        sData(0, 10) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Name)
                        sData(0, 11) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Address)
                        sData(0, 12) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字ID)
                        sData(0, 13) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字名)
                        sData(0, 14) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字ID)
                        sData(0, 15) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字名)
                        sData(0, 16) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番区分)
                        sData(0, 17) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番)
                        sData(0, 18) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番区分)
                        sData(0, 19) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番)
                        sData(0, 20) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番区分)
                        sData(0, 21) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番)
                        sData(0, 22) = ""
                        sData(0, 23) = ""
                        sData(0, 24) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.一部現況)
                        sData(0, 25) = Cnv農地ID(CommonFunc.GetKeyCode(s農地ID))
                        '// 受け人情報
                        sData(0, 26) = pRow.Item("申請理由B").ToString
                        sData(0, 27) = ""
                        sData(0, 28) = ""
                        sData(0, 29) = ""
                        sData(0, 30) = ""
                        sData(0, 31) = ""
                        sData(0, 32) = ""
                        '// 渡し人情報
                        sData(0, 33) = pRow.Item("申請理由A").ToString
                        sData(0, 34) = ""
                        sData(0, 35) = ""
                        sData(0, 36) = ""
                        sData(0, 37) = ""
                        sData(0, 38) = ""
                        sData(0, 39) = ""
                        '// 申請情報
                        sData(0, 40) = IIf(Val(pRow.Item("法令")) = 311, 3, IIf(Val(pRow.Item("小作料").ToString) > 0, 1, 2))
                        sData(0, 41) = ""
                        sData(0, 42) = Conv権利種類(pRow)
                        sData(0, 43) = Val(pRow.Item("10a金額").ToString)
                        sData(0, 44) = Conv小作料(pRow, FarmRent.小作料)
                        sData(0, 45) = Conv小作料(pRow, FarmRent.物納)
                        sData(0, 46) = ""
                        sData(0, 47) = ""
                        sData(0, 48) = ""
                        sData(0, 49) = ""
                        sData(0, 50) = ""
                        sData(0, 51) = ""
                        sData(0, 52) = ""
                        sData(0, 53) = CnvDate(pRow.Item("受付年月日"))
                        sData(0, 54) = CnvDate(pRow.Item("始期"))
                        sData(0, 55) = CnvDate(pRow.Item("終期"))
                        sData(0, 56) = Conv期間(pRow, Time.年)
                        sData(0, 57) = Conv期間(pRow, Time.月)
                        sData(0, 58) = pRow.Item("備考").ToString
                        '// 申請情報
                        sData(0, 59) = CnvDate(pRow.Item("総会日"))
                        sData(0, 60) = CnvDate(pRow.Item("許可年月日"))
                        sData(0, 61) = Val(pRow.Item("許可番号").ToString)

                        If AppNo > 5 Then
                            Dim Range As IXLRange = targetSheet.Range($"{rowNo - 1}:{rowNo - 1}")
                            Dim cell As IXLCell = targetSheet.Cell($"A{rowNo}")
                            Range.CopyTo(cell)
                            targetSheet.Rows(1, rowNo).Height = 40
                            targetSheet.Cell(String.Format("A{0}", rowNo)).InsertData(New Object() {sData})
                        Else
                            targetSheet.Cell(String.Format("A{0}", rowNo)).InsertData(New Object() {sData})
                        End If

                        AppNo += 1
                        rowNo += 1
                    Next
                Case Else
            End Select
        Next
    End Sub

    Private Sub 農地法第4条(ByRef targetSheet As IXLWorksheet)
        Dim sData(,) As Object

        Select Case mvar絞込条件.Text
            Case "全対象"
                Dim pView As DataView = New DataView(TBL申請, "[法令] In (40)", "", DataViewRowState.CurrentRows)
                ReDim sData(pView.Count, 40)
            Case Else
                ReDim sData(TBL申請.Rows.Count, 40)
        End Select

        AppNo = 1
        rowNo = 4

        For Each pRow As DataRow In TBL申請.Rows
            Select Case Val(pRow.Item("法令").ToString)
                Case 40
                    Dim s農地List As String() = Split(pRow.Item("農地リスト").ToString(), ";")

                    For Each s農地ID As String In s農地List
                        sData(0, 0) = AppNo
                        '// 申請種別
                        sData(0, 1) = 42
                        sData(0, 2) = 4
                        sData(0, 3) = ""
                        sData(0, 4) = Val(pRow.Item("受付番号").ToString)
                        '// 申請代理人選択
                        sData(0, 5) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Name)
                        sData(0, 6) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Address)
                        '// 対象地選択
                        sData(0, 7) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Name)
                        sData(0, 8) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Address)
                        sData(0, 9) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字ID)
                        sData(0, 10) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字名)
                        sData(0, 11) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字ID)
                        sData(0, 12) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字名)
                        sData(0, 13) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番区分)
                        sData(0, 14) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番)
                        sData(0, 15) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番区分)
                        sData(0, 16) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番)
                        sData(0, 17) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番区分)
                        sData(0, 18) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番)
                        sData(0, 19) = ""
                        sData(0, 20) = ""
                        sData(0, 21) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.一部現況)
                        sData(0, 22) = Cnv農地ID(CommonFunc.GetKeyCode(s農地ID))
                        '// 申請代理人情報
                        sData(0, 23) = ""
                        sData(0, 24) = ""
                        '// 申請情報
                        sData(0, 25) = 1
                        sData(0, 26) = ""
                        sData(0, 27) = Val(pRow.Item("調査転用用途").ToString)
                        sData(0, 28) = pRow.Item("申請理由A").ToString
                        sData(0, 29) = ""
                        sData(0, 30) = ""
                        sData(0, 31) = CnvDate(pRow.Item("受付年月日"))
                        sData(0, 32) = CnvDate(pRow.Item("始期"))
                        sData(0, 33) = CnvDate(pRow.Item("終期"))
                        sData(0, 34) = Conv期間(pRow, Time.年)
                        sData(0, 35) = Conv期間(pRow, Time.月)
                        sData(0, 36) = pRow.Item("備考").ToString
                        '// 議案内容
                        sData(0, 37) = CnvDate(pRow.Item("総会日"))
                        sData(0, 38) = CnvDate(pRow.Item("進達年月日"))
                        sData(0, 39) = CnvDate(pRow.Item("許可年月日"))
                        sData(0, 40) = CnvDate(pRow.Item("許可番号"))

                        If AppNo > 5 Then
                            Dim Range As IXLRange = targetSheet.Range($"{rowNo - 1}:{rowNo - 1}")
                            Dim cell As IXLCell = targetSheet.Cell($"A{rowNo}")
                            Range.CopyTo(cell)
                            targetSheet.Rows(1, rowNo).Height = 40
                            targetSheet.Cell(String.Format("A{0}", rowNo)).InsertData(New Object() {sData})
                        Else
                            targetSheet.Cell(String.Format("A{0}", rowNo)).InsertData(New Object() {sData})
                        End If

                        AppNo += 1
                        rowNo += 1
                    Next
                Case Else
            End Select
        Next
    End Sub

    Private Sub 農地法第5条(ByRef targetSheet As IXLWorksheet)
        Dim sData(,) As Object

        Select Case mvar絞込条件.Text
            Case "全対象"
                Dim pView As DataView = New DataView(TBL申請, "[法令] In (50,51,52)", "", DataViewRowState.CurrentRows)
                ReDim sData(pView.Count, 52)
            Case Else
                ReDim sData(TBL申請.Rows.Count, 52)
        End Select

        AppNo = 1
        rowNo = 4

        For Each pRow As DataRow In TBL申請.Rows
            Select Case Val(pRow.Item("法令").ToString)
                Case 50, 51, 52
                    Dim s農地List As String() = Split(pRow.Item("農地リスト").ToString(), ";")

                    For Each s農地ID As String In s農地List
                        sData(0, 0) = AppNo
                        '// 申請種別
                        sData(0, 1) = 62
                        Select Case Val(pRow.Item("法令").ToString)
                            Case 50 : sData(0, 2) = 1
                            Case 51, 52 : sData(0, 2) = 2
                        End Select
                        sData(0, 3) = ""
                        sData(0, 4) = Val(pRow.Item("受付番号").ToString)
                        '// 受け人選択
                        sData(0, 5) = Get農家情報(Val(pRow.Item("申請者B").ToString), PersonInfo.Name)
                        sData(0, 6) = Get農家情報(Val(pRow.Item("申請者B").ToString), PersonInfo.Address)
                        sData(0, 7) = Get農家情報(Val(pRow.Item("申請者B").ToString), PersonInfo.Job)
                        '// 申請代理人選択
                        sData(0, 8) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Name)
                        sData(0, 9) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Address)
                        '// 対象地選択
                        sData(0, 10) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Name)
                        sData(0, 11) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Address)
                        sData(0, 12) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字ID)
                        sData(0, 13) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字名)
                        sData(0, 14) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字ID)
                        sData(0, 15) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字名)
                        sData(0, 16) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番区分)
                        sData(0, 17) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番)
                        sData(0, 18) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番区分)
                        sData(0, 19) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番)
                        sData(0, 20) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番区分)
                        sData(0, 21) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番)
                        sData(0, 22) = ""
                        sData(0, 23) = ""
                        sData(0, 24) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.一部現況)
                        sData(0, 25) = Cnv農地ID(CommonFunc.GetKeyCode(s農地ID))
                        '// 受け人情報
                        sData(0, 26) = ""
                        sData(0, 27) = ""
                        '// 渡し人情報
                        sData(0, 28) = ""
                        sData(0, 29) = ""
                        '// 申請情報
                        sData(0, 30) = IIf(Val(pRow.Item("法令").ToString) = 52, 2, 1)
                        sData(0, 31) = ""
                        sData(0, 32) = Val(pRow.Item("調査転用用途").ToString)
                        sData(0, 33) = pRow.Item("申請理由A").ToString
                        sData(0, 34) = ""
                        sData(0, 35) = ""
                        sData(0, 36) = IIf(Val(pRow.Item("小作料").ToString) > 0, 1, 2)
                        sData(0, 37) = ""
                        sData(0, 38) = Conv権利種類(pRow)
                        sData(0, 39) = Val(pRow.Item("10a金額").ToString)
                        sData(0, 40) = Conv小作料(pRow, FarmRent.小作料)
                        sData(0, 41) = Conv小作料(pRow, FarmRent.物納)
                        sData(0, 42) = ""
                        sData(0, 43) = CnvDate(pRow.Item("受付年月日"))
                        sData(0, 44) = CnvDate(pRow.Item("始期"))
                        sData(0, 45) = CnvDate(pRow.Item("終期"))
                        sData(0, 46) = Conv期間(pRow, Time.年)
                        sData(0, 47) = Conv期間(pRow, Time.月)
                        sData(0, 48) = pRow.Item("備考").ToString
                        '// 議案内容
                        sData(0, 49) = CnvDate(pRow.Item("総会日"))
                        sData(0, 50) = CnvDate(pRow.Item("進達年月日"))
                        sData(0, 51) = CnvDate(pRow.Item("許可年月日"))
                        sData(0, 52) = CnvDate(pRow.Item("許可番号"))

                        If AppNo > 5 Then
                            Dim Range As IXLRange = targetSheet.Range($"{rowNo - 1}:{rowNo - 1}")
                            Dim cell As IXLCell = targetSheet.Cell($"A{rowNo}")
                            Range.CopyTo(cell)
                            targetSheet.Rows(1, rowNo).Height = 40
                            targetSheet.Cell(String.Format("A{0}", rowNo)).InsertData(New Object() {sData})
                        Else
                            targetSheet.Cell(String.Format("A{0}", rowNo)).InsertData(New Object() {sData})
                        End If

                        AppNo += 1
                        rowNo += 1
                    Next
                Case Else
            End Select
        Next
    End Sub

    Private Sub 基盤強化法第19条(ByRef book As XLWorkbook)
        AppNo = 1
        rowNo = 4
        AppNoATo中 = 1
        rowNoATo中 = 4
        AppNo中ToB = 1
        rowNo中ToB = 4

        Dim targetSheet As IXLWorksheet = Nothing
        For Each pRow As DataRow In TBL申請.Rows
            If InStr(pRow.Item("名称").ToString, "利用権設定_配分計画に基づく中間管理機構から貸人") > 0 Then
                targetSheet = book.Worksheet("未確定【機構転貸】機構法18条")
                基盤強化法第19条中ToB処理(targetSheet, pRow, FCVia.中ToB)
            ElseIf InStr(pRow.Item("名称").ToString, "利用権設定_中間管理機構") > 0 Then
                Dim n中間管理機構ID As Integer = Val(SysAD.DB(sLRDB).DBProperty("中間管理機構ID"))
                If Val(pRow.Item("申請者B").ToString) = n中間管理機構ID Then
                    targetSheet = book.Worksheet("【機構転貸】基盤法19条")
                    基盤強化法第19条ATo中処理(targetSheet, pRow, FCVia.ATo中)
                Else
                    targetSheet = book.Worksheet("【機構転貸】基盤法19条")
                    基盤強化法第19条ATo中処理(targetSheet, pRow, FCVia.ATo中)
                    targetSheet = book.Worksheet("未確定【機構転貸】機構法18条")
                    基盤強化法第19条中ToB処理(targetSheet, pRow, FCVia.中ToB)
                End If
            Else
                targetSheet = book.Worksheet("19条")
                基盤強化法第19条処理(targetSheet, pRow, FCVia.対象外)
            End If
        Next
    End Sub
    Private Sub 基盤強化法第19条処理(ByRef targetSheet As IXLWorksheet, ByVal pRow As DataRow, ByVal FCType As FCVia)
        Dim sData(,) As Object
        ReDim sData(TBL申請.Rows.Count, 62)

        Dim s農地List As String() = Split(pRow.Item("農地リスト").ToString(), ";")

        For Each s農地ID As String In s農地List
            sData(0, 0) = AppNo
            '// 申請種別
            sData(0, 1) = 6
            Select Case Val(pRow.Item("法令").ToString)
                Case 60 : sData(0, 2) = 1
                Case 61 : sData(0, 2) = 2
                Case 62 : sData(0, 2) = 3
            End Select
            sData(0, 3) = ""
            sData(0, 4) = Val(pRow.Item("受付番号").ToString)
            '// 受け人選択
            sData(0, 5) = Get農家情報(Val(pRow.Item("申請者B").ToString), PersonInfo.Name)
            sData(0, 6) = Get農家情報(Val(pRow.Item("申請者B").ToString), PersonInfo.Address)
            sData(0, 7) = Get農家情報(Val(pRow.Item("申請者B").ToString), PersonInfo.Job)
            '// 申請代理人
            sData(0, 8) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Name)
            sData(0, 9) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Address)
            '// 対象地選択
            sData(0, 10) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Name)
            sData(0, 11) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Address)
            sData(0, 12) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字ID)
            sData(0, 13) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字名)
            sData(0, 14) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字ID)
            sData(0, 15) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字名)
            sData(0, 16) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番区分)
            sData(0, 17) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番)
            sData(0, 18) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番区分)
            sData(0, 19) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番)
            sData(0, 20) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番区分)
            sData(0, 21) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番)
            sData(0, 22) = ""
            sData(0, 23) = ""
            sData(0, 24) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.一部現況)
            sData(0, 25) = Cnv農地ID(CommonFunc.GetKeyCode(s農地ID))
            '// 受け人情報
            sData(0, 26) = ""
            sData(0, 27) = ""
            sData(0, 28) = ""
            sData(0, 29) = ""
            sData(0, 30) = ""
            sData(0, 31) = ""
            '// 渡し人情報
            sData(0, 32) = ""
            sData(0, 33) = ""
            '// 申請情報
            sData(0, 34) = IIf(Val(pRow.Item("小作料").ToString) > 0, 1, 2)
            sData(0, 35) = ""
            sData(0, 36) = Conv権利種類(pRow)
            sData(0, 37) = Val(pRow.Item("10a金額").ToString)
            sData(0, 38) = Conv小作料(pRow, FarmRent.小作料)
            sData(0, 39) = Conv小作料(pRow, FarmRent.物納)
            sData(0, 40) = ""
            sData(0, 41) = ""
            sData(0, 42) = ""
            sData(0, 43) = ""
            sData(0, 44) = ""
            sData(0, 45) = ""
            sData(0, 46) = ""
            sData(0, 47) = ""
            sData(0, 48) = CnvDate(pRow.Item("公告年月日"))
            Select Case Val(Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.現況地目).ToString)
                Case 1 : sData(0, 49) = 1
                Case 2 : sData(0, 49) = 2
                Case Else : sData(0, 49) = 9
            End Select
            sData(0, 50) = pRow.Item("利用権内容").ToString
            sData(0, 51) = IIf(CnvBool(pRow.Item("再設定")) = True, 2, 1)
            sData(0, 52) = ""
            sData(0, 53) = ""
            sData(0, 54) = CnvDate(pRow.Item("受付年月日"))
            sData(0, 55) = CnvDate(pRow.Item("始期"))
            sData(0, 56) = CnvDate(pRow.Item("終期"))
            sData(0, 57) = Conv期間(pRow, Time.年)
            sData(0, 58) = Conv期間(pRow, Time.月)
            sData(0, 59) = pRow.Item("備考").ToString
            '// 申請情報
            sData(0, 60) = CnvDate(pRow.Item("総会日"))
            sData(0, 61) = CnvDate(pRow.Item("許可年月日"))
            sData(0, 62) = Val(pRow.Item("許可番号").ToString)

            If AppNo > 5 Then
                Dim Range As IXLRange = targetSheet.Range($"{rowNo - 1}:{rowNo - 1}")
                Dim cell As IXLCell = targetSheet.Cell($"A{rowNo}")
                'Range.CopyTo(cell)
                targetSheet.Rows(1, rowNo).Height = 40
                targetSheet.Cell(String.Format("A{0}", rowNo)).InsertData(New Object() {sData})
            Else
                targetSheet.Cell(String.Format("A{0}", rowNo)).InsertData(New Object() {sData})
            End If

            AppNo += 1
            rowNo += 1
        Next
    End Sub

    Private Sub 基盤強化法第19条ATo中処理(ByRef targetSheet As IXLWorksheet, ByVal pRow As DataRow, ByVal FCType As FCVia)
        Dim sData(,) As Object
        ReDim sData(TBL申請.Rows.Count, 56)

        Dim n中間管理機構ID As Integer = Val(SysAD.DB(sLRDB).DBProperty("中間管理機構ID"))
        Dim s農地List As String() = Split(pRow.Item("農地リスト").ToString(), ";")
        For Each s農地ID As String In s農地List
            sData(0, 0) = AppNoATo中
            '// 申請種別
            sData(0, 1) = 9
            sData(0, 2) = 2
            sData(0, 3) = ""
            sData(0, 4) = Val(pRow.Item("受付番号").ToString)
            '// 経由転貸人選択
            sData(0, 5) = Get農家情報(n中間管理機構ID, PersonInfo.Name)
            sData(0, 6) = Get農家情報(n中間管理機構ID, PersonInfo.Address)
            sData(0, 7) = Get農家情報(n中間管理機構ID, PersonInfo.Job)
            '// 申請代理人
            sData(0, 8) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Name)
            sData(0, 9) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Address)
            '// 対象地選択
            sData(0, 10) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Name)
            sData(0, 11) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Address)
            sData(0, 12) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字ID)
            sData(0, 13) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字名)
            sData(0, 14) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字ID)
            sData(0, 15) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字名)
            sData(0, 16) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番区分)
            sData(0, 17) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番)
            sData(0, 18) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番区分)
            sData(0, 19) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番)
            sData(0, 20) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番区分)
            sData(0, 21) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番)
            sData(0, 22) = ""
            sData(0, 23) = ""
            sData(0, 24) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.一部現況)
            sData(0, 25) = Cnv農地ID(CommonFunc.GetKeyCode(s農地ID))
            '// 受け人情報
            sData(0, 26) = ""
            sData(0, 27) = ""
            '// 渡し人情報
            sData(0, 28) = pRow.Item("申請理由A").ToString
            sData(0, 29) = ""
            sData(0, 30) = ""
            '// 申請情報
            sData(0, 31) = Conv権利種類(pRow)
            sData(0, 32) = Val(pRow.Item("10a金額").ToString)
            sData(0, 33) = Conv小作料(pRow, FarmRent.小作料)
            sData(0, 34) = Conv小作料(pRow, FarmRent.物納)
            sData(0, 35) = ""
            sData(0, 36) = ""
            sData(0, 37) = ""
            sData(0, 38) = ""
            sData(0, 39) = ""
            sData(0, 40) = ""
            sData(0, 41) = ""
            sData(0, 42) = CnvDate(pRow.Item("公告年月日"))
            Select Case Val(Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.現況地目).ToString)
                Case 1 : sData(0, 43) = 1
                Case 2 : sData(0, 43) = 2
                Case Else : sData(0, 43) = 9
            End Select
            sData(0, 44) = pRow.Item("利用権内容").ToString
            sData(0, 45) = IIf(CnvBool(pRow.Item("再設定")) = True, 2, 1)
            sData(0, 46) = ""
            sData(0, 47) = ""
            sData(0, 48) = CnvDate(pRow.Item("受付年月日"))
            sData(0, 49) = CnvDate(pRow.Item("始期"))
            sData(0, 50) = CnvDate(pRow.Item("終期"))
            sData(0, 51) = Conv期間(pRow, Time.年)
            sData(0, 52) = Conv期間(pRow, Time.月)
            sData(0, 53) = pRow.Item("備考").ToString
            '// 申請情報
            sData(0, 54) = CnvDate(pRow.Item("総会日"))
            sData(0, 55) = CnvDate(pRow.Item("許可年月日"))
            sData(0, 56) = Val(pRow.Item("許可番号").ToString)

            If AppNoATo中 > 5 Then
                Dim Range As IXLRange = targetSheet.Range($"{rowNoATo中 - 1}:{rowNoATo中 - 1}")
                Dim cell As IXLCell = targetSheet.Cell($"A{rowNoATo中}")
                'Range.CopyTo(cell)
                targetSheet.Rows(1, rowNoATo中).Height = 40
                targetSheet.Cell(String.Format("A{0}", rowNoATo中)).InsertData(New Object() {sData})
            Else
                targetSheet.Cell(String.Format("A{0}", rowNoATo中)).InsertData(New Object() {sData})
            End If

            AppNoATo中 += 1
            rowNoATo中 += 1
        Next
    End Sub

    Private Sub 基盤強化法第19条中ToB処理(ByRef targetSheet As IXLWorksheet, ByVal pRow As DataRow, ByVal FCType As FCVia)
        Dim sData(,) As Object
        ReDim sData(TBL申請.Rows.Count, 64)

        Dim n中間管理機構ID As Integer = Val(SysAD.DB(sLRDB).DBProperty("中間管理機構ID"))
        Dim s農地List As String() = Split(pRow.Item("農地リスト").ToString(), ";")
        For Each s農地ID As String In s農地List
            sData(0, 0) = AppNo中ToB
            '// 申請種別
            sData(0, 1) = 12
            Select Case Val(pRow.Item("法令").ToString)
                Case 61 : sData(0, 2) = 2
                Case 62 : sData(0, 2) = 3
            End Select
            sData(0, 3) = Val(pRow.Item("受付番号").ToString)
            '// 受け人選択
            sData(0, 4) = Get農家情報(Val(pRow.Item("申請者B").ToString), PersonInfo.Name)
            sData(0, 5) = Get農家情報(Val(pRow.Item("申請者B").ToString), PersonInfo.Address)
            sData(0, 6) = Get農家情報(Val(pRow.Item("申請者B").ToString), PersonInfo.Job)
            '// 経由転貸人選択
            sData(0, 7) = Get農家情報(n中間管理機構ID, PersonInfo.Name)
            sData(0, 8) = Get農家情報(n中間管理機構ID, PersonInfo.Address)
            sData(0, 9) = Get農家情報(n中間管理機構ID, PersonInfo.Job)
            '// 申請代理人
            sData(0, 10) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Name)
            sData(0, 11) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Address)
            '// 対象地選択
            sData(0, 12) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Name)
            sData(0, 13) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Address)
            sData(0, 14) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字ID)
            sData(0, 15) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字名)
            sData(0, 16) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字ID)
            sData(0, 17) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字名)
            sData(0, 18) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番区分)
            sData(0, 19) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番)
            sData(0, 20) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番区分)
            sData(0, 21) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番)
            sData(0, 22) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番区分)
            sData(0, 23) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番)
            sData(0, 24) = ""
            sData(0, 25) = ""
            sData(0, 26) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.一部現況)
            sData(0, 27) = Cnv農地ID(CommonFunc.GetKeyCode(s農地ID))
            '// 受け人情報
            sData(0, 28) = ""
            sData(0, 29) = ""
            '// 渡し人情報
            sData(0, 30) = pRow.Item("申請理由A").ToString
            sData(0, 31) = ""
            sData(0, 32) = ""
            '// 申請情報
            sData(0, 33) = Conv権利種類(pRow)
            sData(0, 34) = Val(pRow.Item("10a金額").ToString)
            sData(0, 35) = Conv小作料(pRow, FarmRent.小作料)
            sData(0, 36) = Conv小作料(pRow, FarmRent.物納)
            sData(0, 37) = ""
            sData(0, 38) = ""
            sData(0, 39) = Conv権利種類(pRow)
            sData(0, 40) = Val(pRow.Item("10a金額").ToString)
            sData(0, 41) = Conv小作料(pRow, FarmRent.小作料)
            sData(0, 42) = Conv小作料(pRow, FarmRent.物納)
            sData(0, 43) = ""
            sData(0, 44) = ""
            sData(0, 45) = ""
            sData(0, 46) = ""
            sData(0, 47) = ""
            sData(0, 48) = ""
            sData(0, 49) = ""
            sData(0, 50) = CnvDate(pRow.Item("公告年月日"))
            Select Case Val(Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.現況地目).ToString)
                Case 1 : sData(0, 51) = 1
                Case 2 : sData(0, 51) = 2
                Case Else : sData(0, 51) = 9
            End Select
            sData(0, 52) = pRow.Item("利用権内容").ToString
            sData(0, 53) = IIf(CnvBool(pRow.Item("再設定")) = True, 2, 1)
            sData(0, 54) = ""
            sData(0, 55) = ""
            sData(0, 56) = CnvDate(pRow.Item("受付年月日"))
            sData(0, 57) = CnvDate(pRow.Item("始期"))
            sData(0, 58) = CnvDate(pRow.Item("終期"))
            sData(0, 59) = Conv期間(pRow, Time.年)
            sData(0, 60) = Conv期間(pRow, Time.月)
            sData(0, 61) = pRow.Item("備考").ToString
            '// 申請情報
            sData(0, 62) = CnvDate(pRow.Item("総会日"))
            sData(0, 63) = CnvDate(pRow.Item("許可年月日"))
            sData(0, 64) = Val(pRow.Item("許可番号").ToString)

            If AppNo中ToB > 5 Then
                Dim Range As IXLRange = targetSheet.Range($"{rowNo中ToB - 1}:{rowNo中ToB - 1}")
                Dim cell As IXLCell = targetSheet.Cell($"A{rowNo中ToB}")
                'Range.CopyTo(cell)
                targetSheet.Rows(1, rowNo中ToB).Height = 40
                targetSheet.Cell(String.Format("A{0}", rowNo中ToB)).InsertData(New Object() {sData})
            Else
                targetSheet.Cell(String.Format("A{0}", rowNo中ToB)).InsertData(New Object() {sData})
            End If

            AppNo中ToB += 1
            rowNo中ToB += 1
        Next
    End Sub

    Private Sub 農地法第18条(ByRef targetSheet As IXLWorksheet)
        Dim sData(,) As Object
        ReDim sData(TBL申請.Rows.Count, 52)

        AppNo = 1
        rowNo = 4

        For Each pRow As DataRow In TBL申請.Rows
            Dim s農地List As String() = Split(pRow.Item("農地リスト").ToString(), ";")

            For Each s農地ID As String In s農地List
                sData(0, 0) = AppNo
                '// 申請種別
                sData(0, 1) = 22
                sData(0, 2) = 5
                sData(0, 3) = ""
                sData(0, 4) = Val(pRow.Item("受付番号").ToString)
                '// 申請代理人選択
                sData(0, 5) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Name)
                sData(0, 6) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Address)
                '// 対象地選択
                sData(0, 7) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Name)
                sData(0, 8) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Address)
                sData(0, 9) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字ID)
                sData(0, 10) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字名)
                sData(0, 11) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字ID)
                sData(0, 12) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字名)
                sData(0, 13) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番区分)
                sData(0, 14) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番)
                sData(0, 15) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番区分)
                sData(0, 16) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番)
                sData(0, 17) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番区分)
                sData(0, 18) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番)
                sData(0, 19) = ""
                sData(0, 20) = ""
                sData(0, 21) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.一部現況)
                sData(0, 22) = Cnv農地ID(CommonFunc.GetKeyCode(s農地ID))
                '// 受け人情報
                sData(0, 23) = ""
                sData(0, 24) = ""
                '// 渡し人情報
                sData(0, 25) = ""
                sData(0, 26) = ""
                sData(0, 27) = Conv解約形態(pRow)
                sData(0, 28) = ""
                sData(0, 29) = ""
                sData(0, 30) = ""
                sData(0, 31) = ""
                sData(0, 32) = CnvDate(pRow.Item("受付年月日"))
                sData(0, 33) = pRow.Item("備考").ToString
                '// 議案内容
                sData(0, 34) = CnvDate(pRow.Item("総会日"))
                sData(0, 35) = CnvDate(pRow.Item("許可年月日"))
                sData(0, 36) = CnvDate(pRow.Item("許可番号"))

                If AppNo > 5 Then
                    Dim Range As IXLRange = targetSheet.Range($"{rowNo - 1}:{rowNo - 1}")
                    Dim cell As IXLCell = targetSheet.Cell($"A{rowNo}")
                    Range.CopyTo(cell)
                    targetSheet.Rows(1, rowNo).Height = 40
                    targetSheet.Cell(String.Format("A{0}", rowNo)).InsertData(New Object() {sData})
                Else
                    targetSheet.Cell(String.Format("A{0}", rowNo)).InsertData(New Object() {sData})
                End If

                AppNo += 1
                rowNo += 1
            Next
        Next
    End Sub

    Private Sub 非農地判断(ByRef targetSheet As IXLWorksheet)
        Dim sData(,) As Object
        ReDim sData(TBL申請.Rows.Count, 52)

        AppNo = 1
        rowNo = 4

        For Each pRow As DataRow In TBL申請.Rows
            Dim s農地List As String() = Split(pRow.Item("農地リスト").ToString(), ";")

            For Each s農地ID As String In s農地List
                sData(0, 0) = AppNo
                '// 申請種別
                sData(0, 1) = 42
                sData(0, 2) = 4
                sData(0, 3) = ""
                sData(0, 4) = Val(pRow.Item("受付番号").ToString)
                '// 申請代理人選択
                sData(0, 5) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Name)
                sData(0, 6) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Address)
                '// 対象地選択
                sData(0, 7) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Name)
                sData(0, 8) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Address)
                sData(0, 9) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字ID)
                sData(0, 10) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字名)
                sData(0, 11) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字ID)
                sData(0, 12) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字名)
                sData(0, 13) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番区分)
                sData(0, 14) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番)
                sData(0, 15) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番区分)
                sData(0, 16) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番)
                sData(0, 17) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番区分)
                sData(0, 18) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番)
                sData(0, 19) = ""
                sData(0, 20) = ""
                sData(0, 21) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.一部現況)
                sData(0, 22) = Cnv農地ID(CommonFunc.GetKeyCode(s農地ID))
                '// 申請者情報
                sData(0, 23) = ""
                sData(0, 24) = ""
                sData(0, 25) = 1
                sData(0, 26) = IIf(Val(pRow.Item("調査転用用途").ToString) = 0, 51, Val(pRow.Item("調査転用用途").ToString))
                sData(0, 27) = Conv判定地目(pRow)
                sData(0, 28) = CnvDate(pRow.Item("受付年月日"))
                sData(0, 29) = ""
                sData(0, 30) = ""
                sData(0, 31) = ""
                sData(0, 32) = ""
                sData(0, 33) = pRow.Item("備考").ToString
                '// 議案内容
                sData(0, 34) = CnvDate(pRow.Item("総会日"))
                sData(0, 35) = CnvDate(pRow.Item("許可年月日"))
                sData(0, 36) = CnvDate(pRow.Item("許可番号"))


                If AppNo > 5 Then
                    Dim Range As IXLRange = targetSheet.Range($"{rowNo - 1}:{rowNo - 1}")
                    Dim cell As IXLCell = targetSheet.Cell($"A{rowNo}")
                    Range.CopyTo(cell)
                    targetSheet.Rows(1, rowNo).Height = 40
                    targetSheet.Cell(String.Format("A{0}", rowNo)).InsertData(New Object() {sData})
                Else
                    targetSheet.Cell(String.Format("A{0}", rowNo)).InsertData(New Object() {sData})
                End If

                AppNo += 1
                rowNo += 1
            Next
        Next
    End Sub

    Private Sub CreateTBL申請(ByRef pTBL As DataTable)
        With pTBL
            .Columns.Add("ID", GetType(Decimal))
            .Columns.Add("法令", GetType(String))
            .Columns.Add("名称", GetType(String))
            .Columns.Add("許可番号", GetType(Integer))
            .Columns.Add("許可年月日", GetType(Date))
            .Columns.Add("借受人住所", GetType(String))
            .Columns.Add("借受人氏名", GetType(String))
            .Columns.Add("転用目的", GetType(String))
            .Columns.Add("農地区分", GetType(String))
            .Columns.Add("小作料", GetType(Decimal))
            .Columns.Add("小作料単位", GetType(String))
            .Columns.Add("筆数計", GetType(Integer))
            .Columns.Add("面積計", GetType(Decimal))
        End With




    End Sub



End Class
