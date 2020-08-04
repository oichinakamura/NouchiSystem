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

        Me.ToolStrip.Items.AddRange({DTP開始日, DTP終了日, mvar絞込条件, mvar検索開始, New ToolStripSeparator, mvar出力開始, mvarTSProg, mvarTSLabel})

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
            .AddColumnText("許可年月日", "許可年月日", "許可年月日", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("許可番号", "許可番号", "許可番号", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("申請者氏名", "申請者氏名", "申請者氏名", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("申請者住所", "申請者住所", "申請者住所", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("受け人氏名", "受け人氏名", "受け人氏名", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("受け人住所", "受け人住所", "受け人住所", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("小作料", "小作料", "小作料", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("小作料単位", "小作料単位", "小作料単位", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("筆数計", "筆数計", "筆数計", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
        End With
    End Sub

    Private Sub mvar検索開始_Click(sender As Object, e As EventArgs) Handles mvar検索開始.Click
        Try
            Dim s法令 As String = ""
            Select Case mvar絞込条件.Text
                Case "全対象" : s法令 = "30,31,311,40,50,51,52,60,61,62,180,200,210"
                Case "農地法3条" : s法令 = "30,31,311"
                Case "農地法4条" : s法令 = "40"
                Case "農地法5条" : s法令 = "50,51,52"
                Case "基盤強化法" : s法令 = "60,61,62"
                Case "非農地判断" : s法令 = "40" '602,600
                Case "解約" : s法令 = "180,200,210"
            End Select

            TBL申請 = SysAD.DB(sLRDB).GetTableBySqlSelect(String.Format("SELECT * FROM D_申請 WHERE [法令] IN({0}) AND [状態]=2 AND [許可年月日] >= #{1}# AND [許可年月日] <= #{2}# ORDER BY [法令], [許可年月日], [許可番号];", s法令, DTP開始日.Value.ToShortDateString, DTP終了日.Value.ToShortDateString))
            ColumnCheck(TBL申請, "申請者氏名", GetType(String))
            ColumnCheck(TBL申請, "申請者住所", GetType(String))
            ColumnCheck(TBL申請, "受け人氏名", GetType(String))
            ColumnCheck(TBL申請, "受け人住所", GetType(String))
            ColumnCheck(TBL申請, "筆数計", GetType(String))

            For Each pRow As DataRow In TBL申請.Rows
                pRow.Item("申請者氏名") = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Name)
                pRow.Item("申請者住所") = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Address)
                pRow.Item("受け人氏名") = Get農家情報(Val(pRow.Item("申請者B").ToString), PersonInfo.Name)
                pRow.Item("受け人住所") = Get農家情報(Val(pRow.Item("申請者B").ToString), PersonInfo.Address)
                Dim ar As Object = Split(pRow.Item("農地リスト").ToString, ";")
                pRow.Item("筆数計") = UBound(ar) + 1

            Next

            mvarGrid申請.SetDataView(TBL申請, "", "")

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private excelApp As Object = Nothing
    Private excelBooks As Object = Nothing
    Private excelBook As Object = Nothing
    Private excelSheets As Object = Nothing
    Private excelSheet As Object = Nothing

    Private AppNo As Integer = 1
    Private rowNo As Integer = 4
    Private AppNoATo中 As Integer = 1
    Private rowNoATo中 As Integer = 4
    Private AppNo中ToB As Integer = 1
    Private rowNo中ToB As Integer = 4
    Private Sub mvar出力開始_Click(sender As Object, e As EventArgs) Handles mvar出力開始.Click
        mvarTSProg.Visible = True
        mvarTSLabel.Visible = True
        mvar出力開始.Visible = False

        Dim sFilePath As String = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "RPA標準フォーマット.xlsx")
        Dim sDirectory As String = ""

        If System.IO.File.Exists(sFilePath) AndAlso Not TBL申請 Is Nothing Then
            excelApp = CreateObject("Excel.Application")
            excelBooks = excelApp.Workbooks
            excelBook = excelBooks.Open(sFilePath)
            excelSheets = excelBook.Sheets

            Select Case mvar絞込条件.Text
                Case "全対象"
                    excelSheet = excelSheets("3条")
                    農地法第3条(excelSheet)
                    excelSheet = excelSheets("4条")
                    農地法第4条(excelSheet)
                    excelSheet = excelSheets("5条")
                    農地法第5条(excelSheet)
                    基盤強化法第19条()
                    excelSheet = excelSheets("非農地判断")
                    非農地判断(excelSheet)
                    excelSheet = excelSheets("18条解約（許可）")
                    農地法第18条(excelSheet)
                Case "農地法3条"
                    excelSheet = excelSheets("3条")
                    農地法第3条(excelSheet)
                Case "農地法4条"
                    excelSheet = excelSheets("4条")
                    農地法第4条(excelSheet)
                Case "農地法5条"
                    excelSheet = excelSheets("5条")
                    農地法第5条(excelSheet)
                Case "基盤強化法"
                    基盤強化法第19条()
                Case "非農地判断"
                    excelSheet = excelSheets("非農地判断")
                    非農地判断(excelSheet)
                Case "解約"
                    excelSheet = excelSheets("18条解約（許可）")
                    農地法第18条(excelSheet)
            End Select

            Dim savePath As String = 名前を付けて保存("RPA標準フォーマット", OutputType.xlsx, Now.ToString("yyyyMMdd"))
            If savePath <> "" Then
                excelApp.DisplayAlerts = False
                excelBook.SaveAs(savePath)
                excelApp.DisplayAlerts = True

                excelBook.Close()
                ReleaseObject(excelSheets)
                ReleaseObject(excelBooks)
                ReleaseObject(excelBook)
                ReleaseObject(excelApp)

                sDirectory = System.IO.Path.GetDirectoryName(savePath)
                MsgBox("終了しました。")
                Process.Start(sDirectory)
            End If
        End If

        mvarTSProg.Visible = False
        mvarTSLabel.Visible = False
        mvar出力開始.Visible = True
    End Sub

    Private Sub 農地法第3条(ByRef targetSheet As Object)
        Dim sData(,) As Object

        Select Case mvar絞込条件.Text
            Case "全対象"
                Dim pView As DataView = New DataView(TBL申請, "[法令] In (30,31,311)", "", DataViewRowState.CurrentRows)
                ReDim sData(pView.Count * 4, 61)   '配列の次元を変更

                mvarTSProg.Value = 0
                mvarTSProg.Maximum = pView.Count
            Case Else
                ReDim sData(TBL申請.Rows.Count * 4, 61)

                mvarTSProg.Value = 0
                mvarTSProg.Maximum = TBL申請.Rows.Count
        End Select

        AppNo = 1
        rowNo = 4

        For Each pRow As DataRow In TBL申請.Rows
            Select Case Val(pRow.Item("法令").ToString)
                Case 30, 31, 311
                    Dim s農地List As String() = Split(pRow.Item("農地リスト").ToString(), ";")
                    For Each s農地ID As String In s農地List
                        Dim x As Integer = AppNo - 1
                        sData(x, 0) = AppNo
                        '// 申請種別
                        sData(x, 1) = 1
                        Select Case Val(pRow.Item("法令").ToString)
                            Case 30, 311 : sData(x, 2) = 1
                            Case 31 : sData(x, 2) = 2
                        End Select
                        sData(x, 3) = ""
                        sData(x, 4) = Val(pRow.Item("受付番号").ToString)
                        '// 受け人選択
                        sData(x, 5) = Get農家情報(Val(pRow.Item("申請者B").ToString), PersonInfo.Name)
                        sData(x, 6) = Get農家情報(Val(pRow.Item("申請者B").ToString), PersonInfo.Address)
                        sData(x, 7) = Get農家情報(Val(pRow.Item("申請者B").ToString), PersonInfo.Job)
                        sData(x, 8) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Name)
                        sData(x, 9) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Address)
                        '// 対象地選択
                        sData(x, 10) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Name)
                        sData(x, 11) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Address)
                        sData(x, 12) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字ID)
                        sData(x, 13) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字名)
                        sData(x, 14) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字ID)
                        sData(x, 15) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字名)
                        sData(x, 16) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番区分)
                        sData(x, 17) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番)
                        sData(x, 18) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番区分)
                        sData(x, 19) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番)
                        sData(x, 20) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番区分)
                        sData(x, 21) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番)
                        sData(x, 22) = ""
                        sData(x, 23) = ""
                        sData(x, 24) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.一部現況)
                        sData(x, 25) = Cnv農地ID(CommonFunc.GetKeyCode(s農地ID))
                        '// 受け人情報
                        sData(x, 26) = pRow.Item("申請理由B").ToString
                        sData(x, 27) = ""
                        sData(x, 28) = ""
                        sData(x, 29) = ""
                        sData(x, 30) = ""
                        sData(x, 31) = ""
                        sData(x, 32) = ""
                        '// 渡し人情報
                        sData(x, 33) = pRow.Item("申請理由A").ToString
                        sData(x, 34) = ""
                        sData(x, 35) = ""
                        sData(x, 36) = ""
                        sData(x, 37) = ""
                        sData(x, 38) = ""
                        sData(x, 39) = ""
                        '// 申請情報
                        sData(x, 40) = IIf(Val(pRow.Item("法令")) = 311, 3, IIf(Val(pRow.Item("小作料").ToString) > 0, 1, 2))
                        sData(x, 41) = ""
                        sData(x, 42) = Conv権利種類(pRow)
                        sData(x, 43) = Val(pRow.Item("10a金額").ToString)
                        sData(x, 44) = Conv小作料(pRow, FarmRent.小作料)
                        sData(x, 45) = Conv小作料(pRow, FarmRent.物納)
                        sData(x, 46) = ""
                        sData(x, 47) = ""
                        sData(x, 48) = ""
                        sData(x, 49) = ""
                        sData(x, 50) = ""
                        sData(x, 51) = ""
                        sData(x, 52) = ""
                        sData(x, 53) = CnvDate(pRow.Item("受付年月日"))
                        sData(x, 54) = CnvDate(pRow.Item("始期"))
                        sData(x, 55) = CnvDate(pRow.Item("終期"))
                        sData(x, 56) = Conv期間(pRow, Time.年)
                        sData(x, 57) = Conv期間(pRow, Time.月)
                        sData(x, 58) = pRow.Item("備考").ToString
                        '// 申請情報
                        sData(x, 59) = CnvDate(pRow.Item("総会日"))
                        sData(x, 60) = CnvDate(pRow.Item("許可年月日"))
                        sData(x, 61) = Val(pRow.Item("許可番号").ToString)

                        AppNo += 1
                        rowNo += 1
                    Next
                Case Else
            End Select

            mvarTSProg.Increment(1)
            mvarTSLabel.Text = "農地法第3条データ出力中...(" & mvarTSProg.Value & "/" & mvarTSProg.Maximum & ")"
            My.Application.DoEvents()
        Next

        With targetSheet
            .Range(String.Format("A4:BJ{0}", rowNo - 1)) = sData
            .Range("A4:BJ4").Copy
            .Range(String.Format("A4:BJ{0}", rowNo - 1)).PasteSpecial(Paste:=-4122, Operation:=-4142, SkipBlanks:=False, Transpose:=False) '"xlPasteFormats" "xlNone"
            .Range(String.Format("4:{0}", rowNo - 1)).RowHeight = 36
        End With
    End Sub

    Private Sub 農地法第4条(ByRef targetSheet As Object)
        Dim sData(,) As Object

        Select Case mvar絞込条件.Text
            Case "全対象"
                Dim pView As DataView = New DataView(TBL申請, "[法令] In (40)", "", DataViewRowState.CurrentRows)
                ReDim sData(pView.Count * 4, 40)

                mvarTSProg.Value = 0
                mvarTSProg.Maximum = pView.Count
            Case Else
                ReDim sData(TBL申請.Rows.Count * 4, 40)

                mvarTSProg.Value = 0
                mvarTSProg.Maximum = TBL申請.Rows.Count
        End Select

        AppNo = 1
        rowNo = 4

        For Each pRow As DataRow In TBL申請.Rows
            Select Case Val(pRow.Item("法令").ToString)
                Case 40
                    Dim s農地List As String() = Split(pRow.Item("農地リスト").ToString(), ";")

                    For Each s農地ID As String In s農地List
                        Dim x As Integer = AppNo - 1
                        sData(x, 0) = AppNo
                        '// 申請種別
                        sData(x, 1) = 42
                        sData(x, 2) = 4
                        sData(x, 3) = ""
                        sData(x, 4) = Val(pRow.Item("受付番号").ToString)
                        '// 申請代理人選択
                        sData(x, 5) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Name)
                        sData(x, 6) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Address)
                        '// 対象地選択
                        sData(x, 7) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Name)
                        sData(x, 8) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Address)
                        sData(x, 9) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字ID)
                        sData(x, 10) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字名)
                        sData(x, 11) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字ID)
                        sData(x, 12) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字名)
                        sData(x, 13) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番区分)
                        sData(x, 14) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番)
                        sData(x, 15) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番区分)
                        sData(x, 16) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番)
                        sData(x, 17) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番区分)
                        sData(x, 18) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番)
                        sData(x, 19) = ""
                        sData(x, 20) = ""
                        sData(x, 21) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.一部現況)
                        sData(x, 22) = Cnv農地ID(CommonFunc.GetKeyCode(s農地ID))
                        '// 申請代理人情報
                        sData(x, 23) = ""
                        sData(x, 24) = ""
                        '// 申請情報
                        sData(x, 25) = 1
                        sData(x, 26) = ""
                        sData(x, 27) = Val(pRow.Item("調査転用用途").ToString)
                        sData(x, 28) = pRow.Item("申請理由A").ToString
                        sData(x, 29) = ""
                        sData(x, 30) = ""
                        sData(x, 31) = CnvDate(pRow.Item("受付年月日"))
                        sData(x, 32) = CnvDate(pRow.Item("始期"))
                        sData(x, 33) = CnvDate(pRow.Item("終期"))
                        sData(x, 34) = Conv期間(pRow, Time.年)
                        sData(x, 35) = Conv期間(pRow, Time.月)
                        sData(x, 36) = pRow.Item("備考").ToString
                        '// 議案内容
                        sData(x, 37) = CnvDate(pRow.Item("総会日"))
                        sData(x, 38) = CnvDate(pRow.Item("進達年月日"))
                        sData(x, 39) = CnvDate(pRow.Item("許可年月日"))
                        sData(x, 40) = CnvDate(pRow.Item("許可番号"))

                        AppNo += 1
                        rowNo += 1
                    Next
                Case Else
            End Select

            mvarTSProg.Increment(1)
            mvarTSLabel.Text = "農地法第4条データ出力中...(" & mvarTSProg.Value & "/" & mvarTSProg.Maximum & ")"
            My.Application.DoEvents()
        Next

        With targetSheet
            .Range(String.Format("A4:AO{0}", rowNo - 1)) = sData
            .Range("A4:AO4").Copy
            .Range(String.Format("A4:AO{0}", rowNo - 1)).PasteSpecial(Paste:=-4122, Operation:=-4142, SkipBlanks:=False, Transpose:=False) '"xlPasteFormats" "xlNone"
            .Range(String.Format("4:{0}", rowNo - 1)).RowHeight = 36
        End With
    End Sub

    Private Sub 農地法第5条(ByRef targetSheet As Object)
        Dim sData(,) As Object

        Select Case mvar絞込条件.Text
            Case "全対象"
                Dim pView As DataView = New DataView(TBL申請, "[法令] In (50,51,52)", "", DataViewRowState.CurrentRows)
                ReDim sData(pView.Count * 4, 52)

                mvarTSProg.Value = 0
                mvarTSProg.Maximum = pView.Count
            Case Else
                ReDim sData(TBL申請.Rows.Count * 4, 52)

                mvarTSProg.Value = 0
                mvarTSProg.Maximum = TBL申請.Rows.Count
        End Select

        AppNo = 1
        rowNo = 4

        For Each pRow As DataRow In TBL申請.Rows
            Select Case Val(pRow.Item("法令").ToString)
                Case 50, 51, 52
                    Dim s農地List As String() = Split(pRow.Item("農地リスト").ToString(), ";")

                    For Each s農地ID As String In s農地List
                        Dim x As Integer = AppNo - 1
                        sData(x, 0) = AppNo
                        '// 申請種別
                        sData(x, 1) = 62
                        Select Case Val(pRow.Item("法令").ToString)
                            Case 50 : sData(x, 2) = 1
                            Case 51, 52 : sData(x, 2) = 2
                        End Select
                        sData(x, 3) = ""
                        sData(x, 4) = Val(pRow.Item("受付番号").ToString)
                        '// 受け人選択
                        sData(x, 5) = Get農家情報(Val(pRow.Item("申請者B").ToString), PersonInfo.Name)
                        sData(x, 6) = Get農家情報(Val(pRow.Item("申請者B").ToString), PersonInfo.Address)
                        sData(x, 7) = Get農家情報(Val(pRow.Item("申請者B").ToString), PersonInfo.Job)
                        '// 申請代理人選択
                        sData(x, 8) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Name)
                        sData(x, 9) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Address)
                        '// 対象地選択
                        sData(x, 10) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Name)
                        sData(x, 11) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Address)
                        sData(x, 12) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字ID)
                        sData(x, 13) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字名)
                        sData(x, 14) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字ID)
                        sData(x, 15) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字名)
                        sData(x, 16) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番区分)
                        sData(x, 17) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番)
                        sData(x, 18) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番区分)
                        sData(x, 19) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番)
                        sData(x, 20) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番区分)
                        sData(x, 21) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番)
                        sData(x, 22) = ""
                        sData(x, 23) = ""
                        sData(x, 24) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.一部現況)
                        sData(x, 25) = Cnv農地ID(CommonFunc.GetKeyCode(s農地ID))
                        '// 受け人情報
                        sData(x, 26) = ""
                        sData(x, 27) = ""
                        '// 渡し人情報
                        sData(x, 28) = ""
                        sData(x, 29) = ""
                        '// 申請情報
                        sData(x, 30) = IIf(Val(pRow.Item("法令").ToString) = 52, 2, 1)
                        sData(x, 31) = ""
                        sData(x, 32) = Val(pRow.Item("調査転用用途").ToString)
                        sData(x, 33) = pRow.Item("申請理由A").ToString
                        sData(x, 34) = ""
                        sData(x, 35) = ""
                        sData(x, 36) = IIf(Val(pRow.Item("小作料").ToString) > 0, 1, 2)
                        sData(x, 37) = ""
                        sData(x, 38) = Conv権利種類(pRow)
                        sData(x, 39) = Val(pRow.Item("10a金額").ToString)
                        sData(x, 40) = Conv小作料(pRow, FarmRent.小作料)
                        sData(x, 41) = Conv小作料(pRow, FarmRent.物納)
                        sData(x, 42) = ""
                        sData(x, 43) = CnvDate(pRow.Item("受付年月日"))
                        sData(x, 44) = CnvDate(pRow.Item("始期"))
                        sData(x, 45) = CnvDate(pRow.Item("終期"))
                        sData(x, 46) = Conv期間(pRow, Time.年)
                        sData(x, 47) = Conv期間(pRow, Time.月)
                        sData(x, 48) = pRow.Item("備考").ToString
                        '// 議案内容
                        sData(x, 49) = CnvDate(pRow.Item("総会日"))
                        sData(x, 50) = CnvDate(pRow.Item("進達年月日"))
                        sData(x, 51) = CnvDate(pRow.Item("許可年月日"))
                        sData(x, 52) = CnvDate(pRow.Item("許可番号"))

                        AppNo += 1
                        rowNo += 1
                    Next
                Case Else
            End Select

            mvarTSProg.Increment(1)
            mvarTSLabel.Text = "農地法第5条データ出力中...(" & mvarTSProg.Value & "/" & mvarTSProg.Maximum & ")"
            My.Application.DoEvents()
        Next

        With targetSheet
            .Range(String.Format("A4:BA{0}", rowNo - 1)) = sData
            .Range("A4:BA4").Copy
            .Range(String.Format("A4:BA{0}", rowNo - 1)).PasteSpecial(Paste:=-4122, Operation:=-4142, SkipBlanks:=False, Transpose:=False) '"xlPasteFormats" "xlNone"
            .Range(String.Format("4:{0}", rowNo - 1)).RowHeight = 36
        End With
    End Sub

    Private Sub 基盤強化法第19条()
        Dim sData(,) As Object
        Dim sDataAto中(,) As Object
        Dim sData中toB(,) As Object

        Select Case mvar絞込条件.Text
            Case "全対象"
                Dim pView As DataView = New DataView(TBL申請, "[法令] In (60,61,62)", "", DataViewRowState.CurrentRows)
                ReDim sData(pView.Count * 4, 62)
                ReDim sDataAto中(pView.Count * 4, 56)
                ReDim sData中toB(pView.Count * 4, 64)

                mvarTSProg.Value = 0
                mvarTSProg.Maximum = pView.Count
            Case Else
                ReDim sData(TBL申請.Rows.Count * 4, 62)
                ReDim sDataAto中(TBL申請.Rows.Count * 4, 56)
                ReDim sData中toB(TBL申請.Rows.Count * 4, 64)

                mvarTSProg.Value = 0
                mvarTSProg.Maximum = TBL申請.Rows.Count
        End Select

        AppNo = 1
        rowNo = 4
        AppNoATo中 = 1
        rowNoATo中 = 4
        AppNo中ToB = 1
        rowNo中ToB = 4

        Dim targetSheet As Object = Nothing
        For Each pRow As DataRow In TBL申請.Rows
            Select Case Val(pRow.Item("法令").ToString)
                Case 60, 61, 62
                    If InStr(pRow.Item("名称").ToString, "利用権設定_配分計画に基づく中間管理機構から貸人") > 0 Then
                        targetSheet = excelSheets("【機構転貸】機構法18条")
                        基盤強化法第19条中ToB処理(targetSheet, sData中toB, pRow)
                    ElseIf InStr(pRow.Item("名称").ToString, "利用権設定_中間管理機構") > 0 Then
                        Dim n中間管理機構ID As Integer = Val(SysAD.DB(sLRDB).DBProperty("中間管理機構ID"))
                        If Val(pRow.Item("申請者B").ToString) = n中間管理機構ID Then
                            targetSheet = excelSheets("【機構転貸】基盤法19条")
                            基盤強化法第19条ATo中処理(targetSheet, sDataAto中, pRow)
                        Else
                            targetSheet = excelSheets("【機構転貸】基盤法19条")
                            基盤強化法第19条ATo中処理(targetSheet, sDataAto中, pRow)
                            targetSheet = excelSheets("【機構転貸】機構法18条")
                            基盤強化法第19条中ToB処理(targetSheet, sData中toB, pRow)
                        End If
                    Else
                        targetSheet = excelSheets("19条")
                        基盤強化法第19条処理(targetSheet, sData, pRow)
                    End If
                Case Else
            End Select

            mvarTSProg.Increment(1)
            mvarTSLabel.Text = "基盤強化法(機構法含む)データ出力中...(" & mvarTSProg.Value & "/" & mvarTSProg.Maximum & ")"
            My.Application.DoEvents()
        Next

        '// 基盤強化法(AtoB)
        targetSheet = excelSheets("19条")
        With targetSheet
            .Range(String.Format("A4:BK{0}", rowNo - 1)) = sData
            .Range("A4:BK4").Copy
            .Range(String.Format("A4:BK{0}", rowNo - 1)).PasteSpecial(Paste:=-4122, Operation:=-4142, SkipBlanks:=False, Transpose:=False) '"xlPasteFormats" "xlNone"
            .Range(String.Format("4:{0}", rowNo - 1)).RowHeight = 36
        End With
        '// 基盤強化法(Ato中)
        targetSheet = excelSheets("【機構転貸】基盤法19条")
        With targetSheet
            .Range(String.Format("A4:BE{0}", rowNo - 1)) = sDataAto中
            .Range("A4:BE4").Copy
            .Range(String.Format("A4:BE{0}", rowNo - 1)).PasteSpecial(Paste:=-4122, Operation:=-4142, SkipBlanks:=False, Transpose:=False) '"xlPasteFormats" "xlNone"
            .Range(String.Format("4:{0}", rowNo - 1)).RowHeight = 36
        End With
        '// 機構法(中toB)
        targetSheet = excelSheets("【機構転貸】機構法18条")
        With targetSheet
            .Range(String.Format("A4:BM{0}", rowNo - 1)) = sData中toB
            .Range("A4:BM4").Copy
            .Range(String.Format("A4:BM{0}", rowNo - 1)).PasteSpecial(Paste:=-4122, Operation:=-4142, SkipBlanks:=False, Transpose:=False) '"xlPasteFormats" "xlNone"
            .Range(String.Format("4:{0}", rowNo - 1)).RowHeight = 36
        End With
    End Sub
    Private Sub 基盤強化法第19条処理(ByRef targetSheet As Object, ByRef sData As Object, ByVal pRow As DataRow)
        Dim s農地List As String() = Split(pRow.Item("農地リスト").ToString(), ";")

        For Each s農地ID As String In s農地List
            Dim x As Integer = AppNo - 1
            sData(x, 0) = AppNo
            '// 申請種別
            sData(x, 1) = 6
            Select Case Val(pRow.Item("法令").ToString)
                Case 60 : sData(x, 2) = 1
                Case 61 : sData(x, 2) = 2
                Case 62 : sData(x, 2) = 3
            End Select
            sData(x, 3) = ""
            sData(x, 4) = Val(pRow.Item("受付番号").ToString)
            '// 受け人選択
            sData(x, 5) = Get農家情報(Val(pRow.Item("申請者B").ToString), PersonInfo.Name)
            sData(x, 6) = Get農家情報(Val(pRow.Item("申請者B").ToString), PersonInfo.Address)
            sData(x, 7) = Get農家情報(Val(pRow.Item("申請者B").ToString), PersonInfo.Job)
            '// 申請代理人
            sData(x, 8) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Name)
            sData(x, 9) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Address)
            '// 対象地選択
            sData(x, 10) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Name)
            sData(x, 11) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Address)
            sData(x, 12) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字ID)
            sData(x, 13) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字名)
            sData(x, 14) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字ID)
            sData(x, 15) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字名)
            sData(x, 16) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番区分)
            sData(x, 17) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番)
            sData(x, 18) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番区分)
            sData(x, 19) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番)
            sData(x, 20) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番区分)
            sData(x, 21) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番)
            sData(x, 22) = ""
            sData(x, 23) = ""
            sData(x, 24) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.一部現況)
            sData(x, 25) = Cnv農地ID(CommonFunc.GetKeyCode(s農地ID))
            '// 受け人情報
            sData(x, 26) = ""
            sData(x, 27) = ""
            sData(x, 28) = ""
            sData(x, 29) = ""
            sData(x, 30) = ""
            sData(x, 31) = ""
            '// 渡し人情報
            sData(x, 32) = ""
            sData(x, 33) = ""
            '// 申請情報
            sData(x, 34) = IIf(Val(pRow.Item("小作料").ToString) > 0, 1, 2)
            sData(x, 35) = ""
            sData(x, 36) = Conv権利種類(pRow)
            sData(x, 37) = Val(pRow.Item("10a金額").ToString)
            sData(x, 38) = Conv小作料(pRow, FarmRent.小作料)
            sData(x, 39) = Conv小作料(pRow, FarmRent.物納)
            sData(x, 40) = ""
            sData(x, 41) = ""
            sData(x, 42) = ""
            sData(x, 43) = ""
            sData(x, 44) = ""
            sData(x, 45) = ""
            sData(x, 46) = ""
            sData(x, 47) = ""
            sData(x, 48) = CnvDate(pRow.Item("公告年月日"))
            Select Case Val(Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.現況地目).ToString)
                Case 1 : sData(x, 49) = 1
                Case 2 : sData(x, 49) = 2
                Case Else : sData(x, 49) = 9
            End Select
            sData(x, 50) = pRow.Item("利用権内容").ToString
            sData(x, 51) = IIf(CnvBool(pRow.Item("再設定")) = True, 2, 1)
            sData(x, 52) = ""
            sData(x, 53) = ""
            sData(x, 54) = CnvDate(pRow.Item("受付年月日"))
            sData(x, 55) = CnvDate(pRow.Item("始期"))
            sData(x, 56) = CnvDate(pRow.Item("終期"))
            sData(x, 57) = Conv期間(pRow, Time.年)
            sData(x, 58) = Conv期間(pRow, Time.月)
            sData(x, 59) = pRow.Item("備考").ToString
            '// 申請情報
            sData(x, 60) = CnvDate(pRow.Item("総会日"))
            sData(x, 61) = CnvDate(pRow.Item("許可年月日"))
            sData(x, 62) = Val(pRow.Item("許可番号").ToString)

            AppNo += 1
            rowNo += 1
        Next
    End Sub

    Private Sub 基盤強化法第19条ATo中処理(ByRef targetSheet As Object, ByRef sData As Object, ByVal pRow As DataRow)
        Dim n中間管理機構ID As Integer = Val(SysAD.DB(sLRDB).DBProperty("中間管理機構ID"))
        Dim s農地List As String() = Split(pRow.Item("農地リスト").ToString(), ";")
        For Each s農地ID As String In s農地List
            Dim x As Integer = AppNoATo中 - 1
            sData(x, 0) = AppNoATo中
            '// 申請種別
            sData(x, 1) = 9
            sData(x, 2) = 2
            sData(x, 3) = ""
            sData(x, 4) = Val(pRow.Item("受付番号").ToString)
            '// 経由転貸人選択
            sData(x, 5) = Get農家情報(n中間管理機構ID, PersonInfo.Name)
            sData(x, 6) = Get農家情報(n中間管理機構ID, PersonInfo.Address)
            sData(x, 7) = Get農家情報(n中間管理機構ID, PersonInfo.Job)
            '// 申請代理人
            sData(x, 8) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Name)
            sData(x, 9) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Address)
            '// 対象地選択
            sData(x, 10) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Name)
            sData(x, 11) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Address)
            sData(x, 12) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字ID)
            sData(x, 13) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字名)
            sData(x, 14) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字ID)
            sData(x, 15) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字名)
            sData(x, 16) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番区分)
            sData(x, 17) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番)
            sData(x, 18) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番区分)
            sData(x, 19) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番)
            sData(x, 20) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番区分)
            sData(x, 21) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番)
            sData(x, 22) = ""
            sData(x, 23) = ""
            sData(x, 24) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.一部現況)
            sData(x, 25) = Cnv農地ID(CommonFunc.GetKeyCode(s農地ID))
            '// 受け人情報
            sData(x, 26) = ""
            sData(x, 27) = ""
            '// 渡し人情報
            sData(x, 28) = pRow.Item("申請理由A").ToString
            sData(x, 29) = ""
            sData(x, 30) = ""
            '// 申請情報
            sData(x, 31) = Conv権利種類(pRow)
            sData(x, 32) = Val(pRow.Item("10a金額").ToString)
            sData(x, 33) = Conv小作料(pRow, FarmRent.小作料)
            sData(x, 34) = Conv小作料(pRow, FarmRent.物納)
            sData(x, 35) = ""
            sData(x, 36) = ""
            sData(x, 37) = ""
            sData(x, 38) = ""
            sData(x, 39) = ""
            sData(x, 40) = ""
            sData(x, 41) = ""
            sData(x, 42) = CnvDate(pRow.Item("公告年月日"))
            Select Case Val(Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.現況地目).ToString)
                Case 1 : sData(x, 43) = 1
                Case 2 : sData(x, 43) = 2
                Case Else : sData(x, 43) = 9
            End Select
            sData(x, 44) = pRow.Item("利用権内容").ToString
            sData(x, 45) = IIf(CnvBool(pRow.Item("再設定")) = True, 2, 1)
            sData(x, 46) = ""
            sData(x, 47) = ""
            sData(x, 48) = CnvDate(pRow.Item("受付年月日"))
            sData(x, 49) = CnvDate(pRow.Item("始期"))
            sData(x, 50) = CnvDate(pRow.Item("終期"))
            sData(x, 51) = Conv期間(pRow, Time.年)
            sData(x, 52) = Conv期間(pRow, Time.月)
            sData(x, 53) = pRow.Item("備考").ToString
            '// 申請情報
            sData(x, 54) = CnvDate(pRow.Item("総会日"))
            sData(x, 55) = CnvDate(pRow.Item("許可年月日"))
            sData(x, 56) = Val(pRow.Item("許可番号").ToString)

            AppNoATo中 += 1
            rowNoATo中 += 1
        Next
    End Sub

    Private Sub 基盤強化法第19条中ToB処理(ByRef targetSheet As Object, ByRef sData As Object, ByVal pRow As DataRow)
        Dim n中間管理機構ID As Integer = Val(SysAD.DB(sLRDB).DBProperty("中間管理機構ID"))
        Dim s農地List As String() = Split(pRow.Item("農地リスト").ToString(), ";")
        For Each s農地ID As String In s農地List
            Dim x As Integer = AppNo中ToB - 1
            sData(x, 0) = AppNo中ToB
            '// 申請種別
            sData(x, 1) = 12
            Select Case Val(pRow.Item("法令").ToString)
                Case 61 : sData(x, 2) = 2
                Case 62 : sData(x, 2) = 3
            End Select
            sData(x, 3) = Val(pRow.Item("受付番号").ToString)
            '// 受け人選択
            sData(x, 4) = Get農家情報(Val(pRow.Item("申請者B").ToString), PersonInfo.Name)
            sData(x, 5) = Get農家情報(Val(pRow.Item("申請者B").ToString), PersonInfo.Address)
            sData(x, 6) = Get農家情報(Val(pRow.Item("申請者B").ToString), PersonInfo.Job)
            '// 経由転貸人選択
            sData(x, 7) = Get農家情報(n中間管理機構ID, PersonInfo.Name)
            sData(x, 8) = Get農家情報(n中間管理機構ID, PersonInfo.Address)
            sData(x, 9) = Get農家情報(n中間管理機構ID, PersonInfo.Job)
            '// 申請代理人
            sData(x, 10) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Name)
            sData(x, 11) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Address)
            '// 対象地選択
            sData(x, 12) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Name)
            sData(x, 13) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Address)
            sData(x, 14) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字ID)
            sData(x, 15) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字名)
            sData(x, 16) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字ID)
            sData(x, 17) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字名)
            sData(x, 18) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番区分)
            sData(x, 19) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番)
            sData(x, 20) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番区分)
            sData(x, 21) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番)
            sData(x, 22) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番区分)
            sData(x, 23) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番)
            sData(x, 24) = ""
            sData(x, 25) = ""
            sData(x, 26) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.一部現況)
            sData(x, 27) = Cnv農地ID(CommonFunc.GetKeyCode(s農地ID))
            '// 受け人情報
            sData(x, 28) = ""
            sData(x, 29) = ""
            '// 渡し人情報
            sData(x, 30) = pRow.Item("申請理由A").ToString
            sData(x, 31) = ""
            sData(x, 32) = ""
            '// 申請情報
            sData(x, 33) = Conv権利種類(pRow)
            sData(x, 34) = Val(pRow.Item("10a金額").ToString)
            sData(x, 35) = Conv小作料(pRow, FarmRent.小作料)
            sData(x, 36) = Conv小作料(pRow, FarmRent.物納)
            sData(x, 37) = ""
            sData(x, 38) = ""
            sData(x, 39) = Conv権利種類(pRow)
            sData(x, 40) = Val(pRow.Item("10a金額").ToString)
            sData(x, 41) = Conv小作料(pRow, FarmRent.小作料)
            sData(x, 42) = Conv小作料(pRow, FarmRent.物納)
            sData(x, 43) = ""
            sData(x, 44) = ""
            sData(x, 45) = ""
            sData(x, 46) = ""
            sData(x, 47) = ""
            sData(x, 48) = ""
            sData(x, 49) = ""
            sData(x, 50) = CnvDate(pRow.Item("公告年月日"))
            Select Case Val(Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.現況地目).ToString)
                Case 1 : sData(x, 51) = 1
                Case 2 : sData(x, 51) = 2
                Case Else : sData(x, 51) = 9
            End Select
            sData(x, 52) = pRow.Item("利用権内容").ToString
            sData(x, 53) = IIf(CnvBool(pRow.Item("再設定")) = True, 2, 1)
            sData(x, 54) = ""
            sData(x, 55) = ""
            sData(x, 56) = CnvDate(pRow.Item("受付年月日"))
            sData(x, 57) = CnvDate(pRow.Item("始期"))
            sData(x, 58) = CnvDate(pRow.Item("終期"))
            sData(x, 59) = Conv期間(pRow, Time.年)
            sData(x, 60) = Conv期間(pRow, Time.月)
            sData(x, 61) = pRow.Item("備考").ToString
            '// 申請情報
            sData(x, 62) = CnvDate(pRow.Item("総会日"))
            sData(x, 63) = CnvDate(pRow.Item("許可年月日"))
            sData(x, 64) = Val(pRow.Item("許可番号").ToString)

            AppNo中ToB += 1
            rowNo中ToB += 1
        Next
    End Sub

    Private Sub 農地法第18条(ByRef targetSheet As Object)
        Dim sData(,) As Object

        Select Case mvar絞込条件.Text
            Case "全対象"
                Dim pView As DataView = New DataView(TBL申請, "[法令] In (180,200,210)", "", DataViewRowState.CurrentRows)
                ReDim sData(pView.Count * 4, 36)

                mvarTSProg.Value = 0
                mvarTSProg.Maximum = pView.Count
            Case Else
                ReDim sData(TBL申請.Rows.Count * 4, 36)

                mvarTSProg.Value = 0
                mvarTSProg.Maximum = TBL申請.Rows.Count
        End Select

        AppNo = 1
        rowNo = 4

        For Each pRow As DataRow In TBL申請.Rows
            Select Case Val(pRow.Item("法令").ToString)
                Case 180, 200, 210
                    Dim s農地List As String() = Split(pRow.Item("農地リスト").ToString(), ";")

                    For Each s農地ID As String In s農地List
                        Dim x As Integer = AppNo - 1
                        sData(x, 0) = AppNo
                        '// 申請種別
                        sData(x, 1) = 22
                        sData(x, 2) = 5
                        sData(x, 3) = ""
                        sData(x, 4) = Val(pRow.Item("受付番号").ToString)
                        '// 申請代理人選択
                        sData(x, 5) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Name)
                        sData(x, 6) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Address)
                        '// 対象地選択
                        sData(x, 7) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Name)
                        sData(x, 8) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Address)
                        sData(x, 9) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字ID)
                        sData(x, 10) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字名)
                        sData(x, 11) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字ID)
                        sData(x, 12) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字名)
                        sData(x, 13) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番区分)
                        sData(x, 14) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番)
                        sData(x, 15) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番区分)
                        sData(x, 16) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番)
                        sData(x, 17) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番区分)
                        sData(x, 18) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番)
                        sData(x, 19) = ""
                        sData(x, 20) = ""
                        sData(x, 21) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.一部現況)
                        sData(x, 22) = Cnv農地ID(CommonFunc.GetKeyCode(s農地ID))
                        '// 受け人情報
                        sData(x, 23) = ""
                        sData(x, 24) = ""
                        '// 渡し人情報
                        sData(x, 25) = ""
                        sData(x, 26) = ""
                        sData(x, 27) = Conv解約形態(pRow)
                        sData(x, 28) = ""
                        sData(x, 29) = ""
                        sData(x, 30) = ""
                        sData(x, 31) = ""
                        sData(x, 32) = CnvDate(pRow.Item("受付年月日"))
                        sData(x, 33) = pRow.Item("備考").ToString
                        '// 議案内容
                        sData(x, 34) = CnvDate(pRow.Item("総会日"))
                        sData(x, 35) = CnvDate(pRow.Item("許可年月日"))
                        sData(x, 36) = CnvDate(pRow.Item("許可番号"))

                        AppNo += 1
                        rowNo += 1
                    Next
                Case Else
            End Select

            mvarTSProg.Increment(1)
            mvarTSLabel.Text = "農地法第18条データ出力中...(" & mvarTSProg.Value & "/" & mvarTSProg.Maximum & ")"
            My.Application.DoEvents()
        Next

        With targetSheet
            .Range(String.Format("A4:AK{0}", rowNo - 1)) = sData
            .Range("A4:AK4").Copy
            .Range(String.Format("A4:AK{0}", rowNo - 1)).PasteSpecial(Paste:=-4122, Operation:=-4142, SkipBlanks:=False, Transpose:=False) '"xlPasteFormats" "xlNone"
            .Range(String.Format("4:{0}", rowNo - 1)).RowHeight = 36
        End With
    End Sub

    Private Sub 非農地判断(ByRef targetSheet As Object)
        Dim sData(,) As Object

        Select Case mvar絞込条件.Text
            Case "全対象"
                Dim pView As DataView = New DataView(TBL申請, "[法令] In (40)", "", DataViewRowState.CurrentRows)
                ReDim sData(pView.Count * 4, 36)

                mvarTSProg.Value = 0
                mvarTSProg.Maximum = pView.Count
            Case Else
                ReDim sData(TBL申請.Rows.Count * 4, 36)

                mvarTSProg.Value = 0
                mvarTSProg.Maximum = TBL申請.Rows.Count
        End Select

        AppNo = 1
        rowNo = 4

        For Each pRow As DataRow In TBL申請.Rows
            Select Case Val(pRow.Item("法令").ToString)
                Case 40
                    Dim s農地List As String() = Split(pRow.Item("農地リスト").ToString(), ";")

                    For Each s農地ID As String In s農地List
                        Dim x As Integer = AppNo - 1
                        sData(x, 0) = AppNo
                        '// 申請種別
                        sData(x, 1) = 42
                        sData(x, 2) = 4
                        sData(x, 3) = ""
                        sData(x, 4) = Val(pRow.Item("受付番号").ToString)
                        '// 申請代理人選択
                        sData(x, 5) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Name)
                        sData(x, 6) = Get農家情報(Val(pRow.Item("代理人A").ToString), PersonInfo.Address)
                        '// 対象地選択
                        sData(x, 7) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Name)
                        sData(x, 8) = Get農家情報(Val(pRow.Item("申請者A").ToString), PersonInfo.Address)
                        sData(x, 9) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字ID)
                        sData(x, 10) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.大字名)
                        sData(x, 11) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字ID)
                        sData(x, 12) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.小字名)
                        sData(x, 13) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番区分)
                        sData(x, 14) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.本番)
                        sData(x, 15) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番区分)
                        sData(x, 16) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.枝番)
                        sData(x, 17) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番区分)
                        sData(x, 18) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.孫番)
                        sData(x, 19) = ""
                        sData(x, 20) = ""
                        sData(x, 21) = Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.一部現況)
                        sData(x, 22) = Cnv農地ID(CommonFunc.GetKeyCode(s農地ID))
                        '// 申請者情報
                        sData(x, 23) = ""
                        sData(x, 24) = ""
                        sData(x, 25) = 1
                        sData(x, 26) = IIf(Val(pRow.Item("調査転用用途").ToString) = 0, 51, Val(pRow.Item("調査転用用途").ToString))
                        sData(x, 27) = Conv判定地目(Get農地情報(CommonFunc.GetKeyCode(s農地ID), LandInfo.現況地目名))
                        sData(x, 28) = CnvDate(pRow.Item("受付年月日"))
                        sData(x, 29) = ""
                        sData(x, 30) = ""
                        sData(x, 31) = ""
                        sData(x, 32) = ""
                        sData(x, 33) = pRow.Item("備考").ToString
                        '// 議案内容
                        sData(x, 34) = CnvDate(pRow.Item("総会日"))
                        sData(x, 35) = CnvDate(pRow.Item("許可年月日"))
                        sData(x, 36) = CnvDate(pRow.Item("許可番号"))

                        AppNo += 1
                        rowNo += 1
                    Next
                Case Else
            End Select

            mvarTSProg.Increment(1)
            mvarTSLabel.Text = "非農地判断データ出力中...(" & mvarTSProg.Value & "/" & mvarTSProg.Maximum & ")"
            My.Application.DoEvents()
        Next

        With targetSheet
            .Range(String.Format("A4:AK{0}", rowNo - 1)) = sData
            .Range("A4:AK4").Copy
            .Range(String.Format("A4:AK{0}", rowNo - 1)).PasteSpecial(Paste:=-4122, Operation:=-4142, SkipBlanks:=False, Transpose:=False) '"xlPasteFormats" "xlNone"
            .Range(String.Format("4:{0}", rowNo - 1)).RowHeight = 36
        End With
    End Sub
End Class
