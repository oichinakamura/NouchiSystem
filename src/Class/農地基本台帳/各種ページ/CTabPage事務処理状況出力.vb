
Imports HimTools2012
Public Class CTabPage事務処理状況出力
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Private mvarTabCtrl As HimTools2012.controls.TabControlBase
    Private mvar許可範囲開始 As New ToolStripDateTimePickerWithlabel("検索範囲")
    Private mvar許可範囲終了 As New ToolStripDateTimePickerWithlabel("～")
    Private mvar受付日検索 As ToolStripDropDownButton
    Private WithEvents mvar受付中読込 As ToolStripMenuItem
    Private WithEvents mvar許可済読込 As ToolStripMenuItem
    Private WithEvents mvar許可読込 As ToolStripButton
    Private WithEvents mvar絞込 As ToolStripComboBox
    Private mvarTSProg As ToolStripProgressBar
    Private mvarTSLabel As ToolStripLabel
    Private WithEvents mvar事務処理状況出力 As ToolStripButton
    Private mvarGrid申請 As HimTools2012.controls.DataGridViewWithDataView
    Private mvarGrid明細 As HimTools2012.controls.DataGridViewWithDataView

    Public TBL申請 As DataTable
    Public TBL農地 As DataTable
    Public TBL転用農地 As DataTable
    Public TBL削除農地 As DataTable
    Public TBL土地履歴 As DataTable

    Public Sub New()
        MyBase.New(True, True, "事務処理状況出力", "事務処理状況出力")

        mvarTabCtrl = New HimTools2012.controls.TabControlBase
        mvarTabCtrl.Dock = DockStyle.Fill

        ControlPanel.Add(mvarTabCtrl)

        With New 議案書作成パラメータ
            mvar許可範囲開始.Value = .開始年月日
            mvar許可範囲終了.Value = .終了年月日
        End With

        mvar受付日検索 = New ToolStripDropDownButton
        mvar受付日検索.Text = "データ読込(受付日)"

        mvar受付中読込 = New ToolStripMenuItem
        mvar受付中読込.Text = "受付中"
        mvar許可済読込 = New ToolStripMenuItem
        mvar許可済読込.Text = "許可済"
        mvar受付日検索.DropDownItems.AddRange({mvar受付中読込, mvar許可済読込})

        mvar許可読込 = New ToolStripButton
        mvar許可読込.Text = "データ読込(許可日)"

        mvar絞込 = New ToolStripComboBox
        mvar絞込.Items.AddRange(New String() {"全対象", "農地法3条", "農地法4条", "農地法5条", "基盤強化法", "非農地証明", "解約"})
        mvar絞込.Text = "全対象"

        mvarTSProg = New ToolStripProgressBar
        mvarTSProg.Visible = False

        mvarTSLabel = New ToolStripLabel
        mvarTSLabel.Visible = False

        mvar事務処理状況出力 = New ToolStripButton
        mvar事務処理状況出力.Text = "事務処理状況出力"

        Me.ToolStrip.Items.AddRange({mvar許可範囲開始, mvar許可範囲終了, New ToolStripSeparator, mvar絞込, New ToolStripSeparator, mvar受付日検索, mvar許可読込, mvarTSProg, mvarTSLabel, New ToolStripSeparator, mvar事務処理状況出力})

        mvarGrid申請 = New HimTools2012.controls.DataGridViewWithDataView
        CreateGrid申請()

        mvarGrid明細 = New HimTools2012.controls.DataGridViewWithDataView
        CreateGrid明細()

        Dim pPage01 As New HimTools2012.controls.CTabPageWithToolStrip(False, True, "申請", "申請一覧")
        pPage01.ControlPanel.Add(mvarGrid申請)
        mvarGrid申請.Createエクセル出力Ctrl(pPage01.ToolStrip)
        mvarTabCtrl.AddPage(pPage01)

        Dim pPage02 As New HimTools2012.controls.CTabPageWithToolStrip(False, True, "明細", "申請明細")
        pPage02.ControlPanel.Add(mvarGrid明細)
        mvarGrid明細.Createエクセル出力Ctrl(pPage02.ToolStrip)
        mvarTabCtrl.AddPage(pPage02)
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
            .AddColumnText("状態", "状態", "状態", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("許可番号", "受付/許可番号", "許可番号", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("許可年月日", "受付/許可年月日", "許可年月日", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("借受人住所", "借受人住所", "借受人住所", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("借受人氏名", "借受人氏名", "借受人氏名", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("転用目的", "転用目的", "転用目的", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("農地区分", "農地区分", "農地区分", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("小作料", "小作料", "小作料", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("小作料単位", "小作料単位", "小作料単位", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("工事開始年月", "", "", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("工事終了年月", "", "", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("完了報告年月日", "完了報告日", "完了報告年月日", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("筆数計", "筆数計", "筆数計", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("面積計", "面積計", "面積計", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("代表地番", "代表地番", "代表地番", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
        End With
    End Sub

    Private Sub CreateGrid明細()
        With mvarGrid明細
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .AutoGenerateColumns = False
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            .AddColumnText("ID", "ID", "ID", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("法令", "法令", "法令", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("名称", "名称", "名称", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("状態", "状態", "状態", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("権利の種類", "権利の種類", "権利の種類", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("期間", "期間", "期間", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("始期", "始期年月日", "始期", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("終期", "終期年月日", "終期", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("小作料", "小作料", "小作料", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("小作料単位", "小作料単位", "小作料単位", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("許可番号", "受付/許可番号", "許可番号", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("許可年月日", "受付/許可年月日", "許可年月日", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("公告年月日", "公告年月日", "公告年月日", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("申請者", "申請者", "申請者", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("代理人", "代理人", "代理人", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("代理人住所", "代理人住所", "代理人住所", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("転用目的", "転用目的", "転用目的", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("農地区分", "農地区分", "農地区分", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("不許可例外", "不許可の例外", "不許可例外", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("始末書", "始末書の有無", "始末書", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("(借)認定状況", "(借)認定状況", "(借)認定状況", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("農業委員1", "農業委員1", "農業委員1", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("農業委員2", "農業委員2", "農業委員2", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("農業委員3", "農業委員3", "農業委員3", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("再設定", "再設定の有無", "再設定", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)

            .AddColumnText(" ", " ", " ", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("筆ID", "筆ID", "筆ID", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("大字", "大字", "大字", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("小字", "小字", "小字", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("地番", "地番", "地番", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("地目", "地目", "地目", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("面積", "面積", "面積", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("田面積", "田面積", "田面積", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("畑面積", "畑面積", "畑面積", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("樹園地", "樹園地", "樹園地", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("採草放牧面積", "採草放牧面積", "採草放牧面積", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)

        End With
    End Sub

    Private Sub Set利用権設定等事務(ByRef sXML As String)
        Dim pView As DataView = New DataView(TBL申請, "[法令] = 30 OR [法令] = 31 OR [法令] = 60 OR [法令] = 61 OR [法令] = 62", "", DataViewRowState.CurrentRows)
        Dim RowCount As Integer = 0
        Dim Int3条件数 As Integer = 0 : Dim Int利用権件数 As Integer = 0
        Dim Int3条田筆数 As Integer = 0 : Dim Int3条畑筆数 As Integer = 0
        Dim Int3条田所有権件数 As Decimal = 0 : Dim Int3条畑所有権件数 As Decimal = 0
        Dim Int3条田賃貸借件数 As Decimal = 0 : Dim Int3条畑賃貸借件数 As Decimal = 0
        Dim Int3条田使用貸借件数 As Decimal = 0 : Dim Int3条畑使用貸借件数 As Decimal = 0

        Dim Int利用権田筆数 As Integer = 0 : Dim Int利用権畑筆数 As Integer = 0
        Dim Int利用権田所有権件数 As Decimal = 0 : Dim Int利用権畑所有権件数 As Decimal = 0
        Dim Int利用権田賃貸借件数 As Decimal = 0 : Dim Int利用権畑賃貸借件数 As Decimal = 0
        Dim Int利用権田使用貸借件数 As Decimal = 0 : Dim Int利用権畑使用貸借件数 As Decimal = 0
        Dim Int利用権田新規設定件数 As Decimal = 0 : Dim Int利用権畑新規設定件数 As Decimal = 0
        Dim Int利用権田再設定件数 As Decimal = 0 : Dim Int利用権畑再設定件数 As Decimal = 0

        For Each pRow As DataRowView In pView
            'Me.Value += 1 : RowCount += 1
            'Message = "利用権設定等事務データ処理中(" & RowCount & "/" & TBL申請.Rows.Count & ")..."

            Select Case Val(pRow.Item("法令").ToString)
                Case "30", "31" : Int3条件数 += 1
                Case "60", "61", "62" : Int利用権件数 += 1
            End Select
            Dim Ar筆リスト As Object = Split(pRow.Item("農地リスト").ToString, ";")
            For n As Integer = 0 To UBound(Ar筆リスト)
                Dim Ar筆情報 As Object = Split(Ar筆リスト(n), ".")
                Dim pRowFind As DataRow = Nothing

                pRowFind = TBL農地.Rows.Find(Ar筆情報(1))
                If pRowFind Is Nothing Then
                    pRowFind = TBL転用農地.Rows.Find(Ar筆情報(1))
                    If pRowFind Is Nothing Then
                        pRowFind = TBL削除農地.Rows.Find(Ar筆情報(1))
                    End If
                End If

                If Not pRowFind Is Nothing Then
                    Select Case Val(pRow.Item("法令").ToString)
                        Case "30", "31"
                            If Val(pRowFind.Item("田面積").ToString) > 0 Then : Int3条田筆数 += 1
                            ElseIf Val(pRowFind.Item("畑面積").ToString) > 0 Then : Int3条畑筆数 += 1
                            End If

                            Select Case Val(pRow.Item("法令").ToString)
                                Case "30"
                                    If Val(pRowFind.Item("田面積").ToString) > 0 Then : Int3条田所有権件数 += Val(pRowFind.Item("田面積").ToString)
                                    ElseIf Val(pRowFind.Item("畑面積").ToString) > 0 Then : Int3条畑所有権件数 += Val(pRowFind.Item("畑面積").ToString)
                                    End If
                                Case "31"
                                    If Val(pRowFind.Item("田面積").ToString) > 0 Then
                                        Select Case pRow.Item("権利種類")
                                            Case 1 : Int3条田賃貸借件数 += pRowFind.Item("田面積")
                                            Case 2 : Int3条田使用貸借件数 += pRowFind.Item("田面積")
                                        End Select
                                    ElseIf Val(pRowFind.Item("畑面積").ToString) > 0 Then
                                        Select Case pRow.Item("権利種類")
                                            Case 1 : Int3条畑賃貸借件数 += pRowFind.Item("畑面積")
                                            Case 2 : Int3条畑使用貸借件数 += pRowFind.Item("畑面積")
                                        End Select
                                    End If
                            End Select
                        Case "60", "61", "62"
                            If Val(pRowFind.Item("田面積").ToString) > 0 Then
                                Int利用権田筆数 += 1

                                Select Case pRow.Item("再設定")
                                    Case True : Int利用権田再設定件数 += Val(pRowFind.Item("田面積").ToString)
                                    Case False : Int利用権田新規設定件数 += Val(pRowFind.Item("田面積").ToString)
                                End Select
                            ElseIf Val(pRowFind.Item("畑面積").ToString) > 0 Then
                                Int利用権畑筆数 += 1

                                Select Case pRow.Item("再設定")
                                    Case True : Int利用権畑再設定件数 += Val(pRowFind.Item("畑面積").ToString)
                                    Case False : Int利用権畑新規設定件数 += Val(pRowFind.Item("畑面積").ToString)
                                End Select
                            End If

                            Select Case Val(pRow.Item("法令").ToString)
                                Case "60", "62"
                                    If Val(pRowFind.Item("田面積").ToString) > 0 Then : Int利用権田所有権件数 += Val(pRowFind.Item("田面積").ToString)
                                    ElseIf Val(pRowFind.Item("畑面積").ToString) > 0 Then : Int利用権畑所有権件数 += Val(pRowFind.Item("畑面積").ToString)
                                    End If
                                Case "61"
                                    If Val(pRowFind.Item("田面積").ToString) > 0 Then
                                        Select Case pRow.Item("権利種類")
                                            Case 1 : Int利用権田賃貸借件数 += pRowFind.Item("田面積")
                                            Case 2 : Int利用権田使用貸借件数 += pRowFind.Item("田面積")
                                        End Select
                                    ElseIf Val(pRowFind.Item("畑面積").ToString) > 0 Then
                                        Select Case pRow.Item("権利種類")
                                            Case 1 : Int利用権畑賃貸借件数 += pRowFind.Item("畑面積")
                                            Case 2 : Int利用権畑使用貸借件数 += pRowFind.Item("畑面積")
                                        End Select
                                    End If
                            End Select
                    End Select
                End If
            Next
        Next

        sXML = Replace(sXML, "{3条件数}", Int3条件数)
        sXML = Replace(sXML, "{3条田筆数}", Int3条田筆数)
        sXML = Replace(sXML, "{3条畑筆数}", Int3条畑筆数)
        sXML = Replace(sXML, "{3条田所有権件数}", Int3条田所有権件数)
        sXML = Replace(sXML, "{3条畑所有権件数}", Int3条畑所有権件数)
        sXML = Replace(sXML, "{3条田賃貸借件数}", Int3条田賃貸借件数)
        sXML = Replace(sXML, "{3条畑賃貸借件数}", Int3条畑賃貸借件数)
        sXML = Replace(sXML, "{3条田使用賃貸件数}", Int3条田使用貸借件数)
        sXML = Replace(sXML, "{3条畑使用賃貸件数}", Int3条畑使用貸借件数)

        sXML = Replace(sXML, "{3条田面積}", Int3条田所有権件数 + Int3条田賃貸借件数 + Int3条田使用貸借件数)
        sXML = Replace(sXML, "{3条畑面積}", Int3条畑所有権件数 + Int3条畑賃貸借件数 + Int3条畑使用貸借件数)

        sXML = Replace(sXML, "{利用権件数}", Int利用権件数)
        sXML = Replace(sXML, "{利用権田筆数}", Int利用権田筆数)
        sXML = Replace(sXML, "{利用権畑筆数}", Int利用権畑筆数)
        sXML = Replace(sXML, "{利用権田所有権件数}", Int利用権田所有権件数)
        sXML = Replace(sXML, "{利用権畑所有権件数}", Int利用権畑所有権件数)
        sXML = Replace(sXML, "{利用権田賃貸借件数}", Int利用権田賃貸借件数)
        sXML = Replace(sXML, "{利用権畑賃貸借件数}", Int利用権畑賃貸借件数)
        sXML = Replace(sXML, "{利用権田使用賃貸件数}", Int利用権田使用貸借件数)
        sXML = Replace(sXML, "{利用権畑使用賃貸件数}", Int利用権畑使用貸借件数)

        sXML = Replace(sXML, "{利用権田新規設定件数}", Int利用権田新規設定件数)
        sXML = Replace(sXML, "{利用権畑新規設定件数}", Int利用権畑新規設定件数)
        sXML = Replace(sXML, "{利用権田再設定件数}", Int利用権田再設定件数)
        sXML = Replace(sXML, "{利用権畑再設定件数}", Int利用権畑再設定件数)

        sXML = Replace(sXML, "{利用権田面積}", Int利用権田所有権件数 + Int利用権田賃貸借件数 + Int利用権田使用貸借件数 + Int利用権田新規設定件数 + Int利用権田再設定件数)
        sXML = Replace(sXML, "{利用権畑面積}", Int利用権畑所有権件数 + Int利用権畑賃貸借件数 + Int利用権畑使用貸借件数 + Int利用権畑新規設定件数 + Int利用権畑再設定件数)
    End Sub

    Private Sub Set4条5条関係事務(ByRef sXML As String)
        Dim pView As DataView = New DataView(TBL申請, "[法令] = 40 OR [法令] = 50 OR [法令] = 51 OR [法令] = 52", "", DataViewRowState.CurrentRows)
        Dim RowCount As Integer = 0
        Dim Int4条件数 As Integer = 0 : Dim Int5条件数 As Integer = 0
        Dim Int4条田筆数 As Integer = 0 : Dim Int4条畑筆数 As Integer = 0
        Dim Int4条田面積 As Decimal = 0 : Dim Int4条畑面積 As Decimal = 0

        Dim Int5条田筆数 As Integer = 0 : Dim Int5条畑筆数 As Integer = 0
        Dim Int5条田面積 As Decimal = 0 : Dim Int5条畑面積 As Decimal = 0

        Dim pTBL4条5条 As New DataTable
        With pTBL4条5条
            .Columns.Add("名称", GetType(String))
            .Columns.Add("田筆数", GetType(Integer))
            .Columns.Add("畑筆数", GetType(Integer))
            .Columns.Add("田面積", GetType(Decimal))
            .Columns.Add("畑面積", GetType(Decimal))
            .PrimaryKey = New DataColumn() { .Columns("名称")}
        End With

        For Each pRow As DataRowView In pView
            'Me.Value += 1 : RowCount += 1
            'Message = "4条5条関係事務データ処理中(" & RowCount & "/" & TBL申請.Rows.Count & ")..."

            Select Case Val(pRow.Item("法令").ToString)
                Case "40" : Int4条件数 += 1
                Case "50", "51", "52" : Int5条件数 += 1
            End Select
            Dim Ar筆リスト As Object = Split(pRow.Item("農地リスト").ToString, ";")
            For n As Integer = 0 To UBound(Ar筆リスト)
                Dim Ar筆情報 As Object = Split(Ar筆リスト(n), ".")
                Dim pRowFind As DataRow = Nothing

                pRowFind = TBL農地.Rows.Find(Ar筆情報(1))
                If pRowFind Is Nothing Then
                    pRowFind = TBL転用農地.Rows.Find(Ar筆情報(1))
                    If pRowFind Is Nothing Then
                        pRowFind = TBL削除農地.Rows.Find(Ar筆情報(1))
                    End If
                End If

                If Not pRowFind Is Nothing Then
                    Select Case Val(pRow.Item("法令").ToString)
                        Case "40"
                            If Val(pRowFind.Item("田面積").ToString) > 0 Then
                                Int4条田筆数 += 1
                                Int4条田面積 += Val(pRowFind.Item("田面積").ToString)
                            ElseIf Val(pRowFind.Item("畑面積").ToString) > 0 Then
                                Int4条畑筆数 += 1
                                Int4条畑面積 += Val(pRowFind.Item("畑面積").ToString)
                            End If

                            Set申請内訳(pTBL4条5条, pRowFind, "4条", "宅地")
                            Set申請内訳(pTBL4条5条, pRowFind, "4条", "道路")
                            Set申請内訳(pTBL4条5条, pRowFind, "4条", "雑種地")
                            Set申請内訳(pTBL4条5条, pRowFind, "4条", "保安林")
                            Set申請内訳(pTBL4条5条, pRowFind, "4条", "牧場")
                        Case "50", "51", "52"
                            If Val(pRowFind.Item("田面積").ToString) > 0 Then
                                Int5条田筆数 += 1
                                Int5条田面積 += Val(pRowFind.Item("田面積").ToString)
                            ElseIf Val(pRowFind.Item("畑面積").ToString) > 0 Then
                                Int5条畑筆数 += 1
                                Int5条畑面積 += Val(pRowFind.Item("畑面積").ToString)
                            End If

                            Set申請内訳(pTBL4条5条, pRowFind, "5条", "宅地")
                            Set申請内訳(pTBL4条5条, pRowFind, "5条", "道路")
                            Set申請内訳(pTBL4条5条, pRowFind, "5条", "雑種地")
                            Set申請内訳(pTBL4条5条, pRowFind, "5条", "保安林")
                            Set申請内訳(pTBL4条5条, pRowFind, "5条", "牧場")
                    End Select
                End If
            Next
        Next

        sXML = Replace(sXML, "{4条件数}", Int4条件数)
        sXML = Replace(sXML, "{4条田筆数}", Int4条田筆数)
        sXML = Replace(sXML, "{4条畑筆数}", Int4条畑筆数)
        sXML = Replace(sXML, "{4条田面積}", Int4条田面積)
        sXML = Replace(sXML, "{4条畑面積}", Int4条畑面積)

        sXML = Replace(sXML, "{4条宅地田筆数}", Find申請面積(pTBL4条5条, "4条宅地", "田筆数"))
        sXML = Replace(sXML, "{4条宅地田面積}", Find申請面積(pTBL4条5条, "4条宅地", "田面積"))
        sXML = Replace(sXML, "{4条宅地畑筆数}", Find申請面積(pTBL4条5条, "4条宅地", "畑筆数"))
        sXML = Replace(sXML, "{4条宅地畑面積}", Find申請面積(pTBL4条5条, "4条宅地", "畑面積"))
        sXML = Replace(sXML, "{4条道路田筆数}", Find申請面積(pTBL4条5条, "4条道路", "田筆数"))
        sXML = Replace(sXML, "{4条道路田面積}", Find申請面積(pTBL4条5条, "4条道路", "田面積"))
        sXML = Replace(sXML, "{4条道路畑筆数}", Find申請面積(pTBL4条5条, "4条道路", "畑筆数"))
        sXML = Replace(sXML, "{4条道路畑面積}", Find申請面積(pTBL4条5条, "4条道路", "畑面積"))
        sXML = Replace(sXML, "{4条雑種地田筆数}", Find申請面積(pTBL4条5条, "4条雑種地", "田筆数"))
        sXML = Replace(sXML, "{4条雑種地田面積}", Find申請面積(pTBL4条5条, "4条雑種地", "田面積"))
        sXML = Replace(sXML, "{4条雑種地畑筆数}", Find申請面積(pTBL4条5条, "4条雑種地", "畑筆数"))
        sXML = Replace(sXML, "{4条雑種地畑面積}", Find申請面積(pTBL4条5条, "4条雑種地", "畑面積"))
        sXML = Replace(sXML, "{4条保安林田筆数}", Find申請面積(pTBL4条5条, "4条保安林", "田筆数"))
        sXML = Replace(sXML, "{4条保安林田面積}", Find申請面積(pTBL4条5条, "4条保安林", "田面積"))
        sXML = Replace(sXML, "{4条保安林畑筆数}", Find申請面積(pTBL4条5条, "4条保安林", "畑筆数"))
        sXML = Replace(sXML, "{4条保安林畑面積}", Find申請面積(pTBL4条5条, "4条保安林", "畑面積"))
        sXML = Replace(sXML, "{4条牧場田筆数}", Find申請面積(pTBL4条5条, "4条牧場", "田筆数"))
        sXML = Replace(sXML, "{4条牧場田面積}", Find申請面積(pTBL4条5条, "4条牧場", "田面積"))
        sXML = Replace(sXML, "{4条牧場畑筆数}", Find申請面積(pTBL4条5条, "4条牧場", "畑筆数"))
        sXML = Replace(sXML, "{4条牧場畑面積}", Find申請面積(pTBL4条5条, "4条牧場", "畑面積"))

        sXML = Replace(sXML, "{5条件数}", Int5条件数)
        sXML = Replace(sXML, "{5条田筆数}", Int5条田筆数)
        sXML = Replace(sXML, "{5条畑筆数}", Int5条畑筆数)
        sXML = Replace(sXML, "{5条田面積}", Int5条田面積)
        sXML = Replace(sXML, "{5条畑面積}", Int5条畑面積)

        sXML = Replace(sXML, "{5条宅地田筆数}", Find申請面積(pTBL4条5条, "5条宅地", "田筆数"))
        sXML = Replace(sXML, "{5条宅地田面積}", Find申請面積(pTBL4条5条, "5条宅地", "田面積"))
        sXML = Replace(sXML, "{5条宅地畑筆数}", Find申請面積(pTBL4条5条, "5条宅地", "畑筆数"))
        sXML = Replace(sXML, "{5条宅地畑面積}", Find申請面積(pTBL4条5条, "5条宅地", "畑面積"))
        sXML = Replace(sXML, "{5条道路田筆数}", Find申請面積(pTBL4条5条, "5条道路", "田筆数"))
        sXML = Replace(sXML, "{5条道路田面積}", Find申請面積(pTBL4条5条, "5条道路", "田面積"))
        sXML = Replace(sXML, "{5条道路畑筆数}", Find申請面積(pTBL4条5条, "5条道路", "畑筆数"))
        sXML = Replace(sXML, "{5条道路畑面積}", Find申請面積(pTBL4条5条, "5条道路", "畑面積"))
        sXML = Replace(sXML, "{5条雑種地田筆数}", Find申請面積(pTBL4条5条, "5条雑種地", "田筆数"))
        sXML = Replace(sXML, "{5条雑種地田面積}", Find申請面積(pTBL4条5条, "5条雑種地", "田面積"))
        sXML = Replace(sXML, "{5条雑種地畑筆数}", Find申請面積(pTBL4条5条, "5条雑種地", "畑筆数"))
        sXML = Replace(sXML, "{5条雑種地畑面積}", Find申請面積(pTBL4条5条, "5条雑種地", "畑面積"))
        sXML = Replace(sXML, "{5条保安林田筆数}", Find申請面積(pTBL4条5条, "5条保安林", "田筆数"))
        sXML = Replace(sXML, "{5条保安林田面積}", Find申請面積(pTBL4条5条, "5条保安林", "田面積"))
        sXML = Replace(sXML, "{5条保安林畑筆数}", Find申請面積(pTBL4条5条, "5条保安林", "畑筆数"))
        sXML = Replace(sXML, "{5条保安林畑面積}", Find申請面積(pTBL4条5条, "5条保安林", "畑面積"))
        sXML = Replace(sXML, "{5条牧場田筆数}", Find申請面積(pTBL4条5条, "5条牧場", "田筆数"))
        sXML = Replace(sXML, "{5条牧場田面積}", Find申請面積(pTBL4条5条, "5条牧場", "田面積"))
        sXML = Replace(sXML, "{5条牧場畑筆数}", Find申請面積(pTBL4条5条, "5条牧場", "畑筆数"))
        sXML = Replace(sXML, "{5条牧場畑面積}", Find申請面積(pTBL4条5条, "5条牧場", "畑面積"))
    End Sub

    Private Sub Set非農地証明願(ByRef sXML As String)
        Dim pView As DataView = New DataView(TBL申請, "[法令] = 600 OR [法令] = 602", "", DataViewRowState.CurrentRows)
        Dim RowCount As Integer = 0
        Dim Int非農地件数 As Integer = 0
        Dim Int非農地田筆数 As Integer = 0 : Dim Int非農地畑筆数 As Integer = 0
        Dim Int非農地田面積 As Decimal = 0 : Dim Int非農地畑面積 As Decimal = 0

        Dim pTBL非農地 As New DataTable
        With pTBL非農地
            .Columns.Add("名称", GetType(String))
            .Columns.Add("田筆数", GetType(Integer))
            .Columns.Add("畑筆数", GetType(Integer))
            .Columns.Add("田面積", GetType(Decimal))
            .Columns.Add("畑面積", GetType(Decimal))
            .PrimaryKey = New DataColumn() { .Columns("名称")}
        End With

        For Each pRow As DataRowView In pView
            'Me.Value += 1 : RowCount += 1 : Int非農地件数 += 1
            'Message = "4条5条関係事務データ処理中(" & RowCount & "/" & TBL申請.Rows.Count & ")..."

            Dim Ar筆リスト As Object = Split(pRow.Item("農地リスト").ToString, ";")
            For n As Integer = 0 To UBound(Ar筆リスト)
                Dim Ar筆情報 As Object = Split(Ar筆リスト(n), ".")
                Dim pRowFind As DataRow = Nothing

                pRowFind = TBL農地.Rows.Find(Ar筆情報(1))
                If pRowFind Is Nothing Then
                    pRowFind = TBL転用農地.Rows.Find(Ar筆情報(1))
                    If pRowFind Is Nothing Then
                        pRowFind = TBL削除農地.Rows.Find(Ar筆情報(1))
                    End If
                End If

                If Not pRowFind Is Nothing Then
                    If Val(pRowFind.Item("田面積").ToString) > 0 Then
                        Int非農地田筆数 += 1
                        Int非農地田面積 += Val(pRowFind.Item("田面積").ToString)
                    ElseIf Val(pRowFind.Item("畑面積").ToString) > 0 Then
                        Int非農地畑筆数 += 1
                        Int非農地畑面積 += Val(pRowFind.Item("畑面積").ToString)
                    End If

                    Set申請内訳(pTBL非農地, pRowFind, "非農地", "宅地")
                    Set申請内訳(pTBL非農地, pRowFind, "非農地", "山林")
                    Set申請内訳(pTBL非農地, pRowFind, "非農地", "雑種地")
                    Set申請内訳(pTBL非農地, pRowFind, "非農地", "原野")
                    Set申請内訳(pTBL非農地, pRowFind, "非農地", "道路")
                End If
            Next
        Next

        sXML = Replace(sXML, "{非農地件数}", Int非農地件数)
        sXML = Replace(sXML, "{非農地田筆数}", Int非農地田筆数)
        sXML = Replace(sXML, "{非農地畑筆数}", Int非農地畑筆数)
        sXML = Replace(sXML, "{非農地田面積}", Int非農地田面積)
        sXML = Replace(sXML, "{非農地畑面積}", Int非農地畑面積)

        sXML = Replace(sXML, "{非農地宅地田筆数}", Find申請面積(pTBL非農地, "非農地宅地", "田筆数"))
        sXML = Replace(sXML, "{非農地宅地田面積}", Find申請面積(pTBL非農地, "非農地宅地", "田面積"))
        sXML = Replace(sXML, "{非農地宅地畑筆数}", Find申請面積(pTBL非農地, "非農地宅地", "畑筆数"))
        sXML = Replace(sXML, "{非農地宅地畑面積}", Find申請面積(pTBL非農地, "非農地宅地", "畑面積"))
        sXML = Replace(sXML, "{非農地山林田筆数}", Find申請面積(pTBL非農地, "非農地山林", "田筆数"))
        sXML = Replace(sXML, "{非農地山林田面積}", Find申請面積(pTBL非農地, "非農地山林", "田面積"))
        sXML = Replace(sXML, "{非農地山林畑筆数}", Find申請面積(pTBL非農地, "非農地山林", "畑筆数"))
        sXML = Replace(sXML, "{非農地山林畑面積}", Find申請面積(pTBL非農地, "非農地山林", "畑面積"))
        sXML = Replace(sXML, "{非農地雑種地田筆数}", Find申請面積(pTBL非農地, "非農地雑種地", "田筆数"))
        sXML = Replace(sXML, "{非農地雑種地田面積}", Find申請面積(pTBL非農地, "非農地雑種地", "田面積"))
        sXML = Replace(sXML, "{非農地雑種地畑筆数}", Find申請面積(pTBL非農地, "非農地雑種地", "畑筆数"))
        sXML = Replace(sXML, "{非農地雑種地畑面積}", Find申請面積(pTBL非農地, "非農地雑種地", "畑面積"))
        sXML = Replace(sXML, "{非農地原野田筆数}", Find申請面積(pTBL非農地, "非農地原野", "田筆数"))
        sXML = Replace(sXML, "{非農地原野田面積}", Find申請面積(pTBL非農地, "非農地原野", "田面積"))
        sXML = Replace(sXML, "{非農地原野畑筆数}", Find申請面積(pTBL非農地, "非農地原野", "畑筆数"))
        sXML = Replace(sXML, "{非農地原野畑面積}", Find申請面積(pTBL非農地, "非農地原野", "畑面積"))
        sXML = Replace(sXML, "{非農地道路田筆数}", Find申請面積(pTBL非農地, "非農地道路", "田筆数"))
        sXML = Replace(sXML, "{非農地道路田面積}", Find申請面積(pTBL非農地, "非農地道路", "田面積"))
        sXML = Replace(sXML, "{非農地道路畑筆数}", Find申請面積(pTBL非農地, "非農地道路", "畑筆数"))
        sXML = Replace(sXML, "{非農地道路畑面積}", Find申請面積(pTBL非農地, "非農地道路", "畑面積"))
    End Sub

    Private Sub Set申請内訳(ByRef pTBL As DataTable, ByRef pRowFind As DataRow, ByVal pKey As String, ByVal p地目 As String)
        If InStr(pRowFind.Item("現況地目名").ToString, p地目) > 0 Then
            If Val(pRowFind.Item("田面積").ToString) > 0 Then
                Dim pRow As DataRow = pTBL.Rows.Find(String.Format("{0}{1}", pKey, p地目))
                If pRow Is Nothing Then
                    Dim pAddRow As DataRow = pTBL.NewRow
                    pAddRow.Item("名称") = String.Format("{0}{1}", pKey, p地目)
                    pAddRow.Item("田筆数") = 1
                    pAddRow.Item("田面積") = Val(pRowFind.Item("田面積").ToString)
                    pTBL.Rows.Add(pAddRow)
                Else
                    pRow.Item("田筆数") = Val(pRow.Item("田筆数").ToString) + 1
                    pRow.Item("田面積") = Val(pRow.Item("田面積").ToString) + Val(pRowFind.Item("田面積").ToString)
                End If
            ElseIf Val(pRowFind.Item("畑面積").ToString) > 0 Then
                Dim pRow As DataRow = pTBL.Rows.Find(String.Format("{0}{1}", pKey, p地目))
                If pRow Is Nothing Then
                    Dim pAddRow As DataRow = pTBL.NewRow
                    pAddRow.Item("名称") = String.Format("{0}{1}", pKey, p地目)
                    pAddRow.Item("畑筆数") = 1
                    pAddRow.Item("畑面積") = Val(pRowFind.Item("畑面積").ToString)
                    pTBL.Rows.Add(pAddRow)
                Else
                    pRow.Item("畑筆数") = Val(pRow.Item("畑筆数").ToString) + 1
                    pRow.Item("畑面積") = Val(pRow.Item("畑面積").ToString) + Val(pRowFind.Item("畑面積").ToString)
                End If
            End If
        End If
    End Sub

    Private Function Find申請面積(ByRef pTBL As DataTable, ByVal pKey As String, ByVal pValue As String) As Decimal
        Dim pResult As Decimal = 0
        Dim pRow As DataRow = pTBL.Rows.Find(pKey)
        If Not pRow Is Nothing Then
            pResult = pRow.Item(pValue)
        End If

        Return pResult
    End Function

    Private Sub mvar事務処理状況出力_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar事務処理状況出力.Click
        Dim sFile As String = ""
        If Not SysAD.IsClickOnceDeployed Then
            sFile = My.Application.Info.DirectoryPath & "\" & SysAD.市町村.市町村名 & "\農業委員会関係事務処理状況.xml"
        Else
            sFile = SysAD.ClickOnceSetupPath & "\" & SysAD.市町村.市町村名 & "\農業委員会関係事務処理状況.xml"
        End If

        If IO.File.Exists(sFile) Then
            Dim sXML As String = HimTools2012.TextAdapter.LoadTextFile(sFile)

            Set利用権設定等事務(sXML)
            Set4条5条関係事務(sXML)
            Set非農地証明願(sXML)

            Dim sOutPutFile As String = SysAD.OutputFolder & String.Format("\農業委員会関係事務処理状況({0}～{1}).xml", 和暦Format(mvar許可範囲開始.Value), 和暦Format(mvar許可範囲終了.Value))
            HimTools2012.TextAdapter.SaveTextFile(sOutPutFile, sXML)
            SysAD.ShowFolder(System.IO.Directory.GetParent(sOutPutFile).ToString)
        Else
            MsgBox("指定様式がありません。")
        End If
    End Sub

    Private Sub mvar受付中読込_Click(sender As Object, e As EventArgs) Handles mvar受付中読込.Click
        検索開始(Enum状態.受付中)
    End Sub
    Private Sub mvar許可済読込_Click(sender As Object, e As EventArgs) Handles mvar許可済読込.Click
        検索開始(Enum状態.許可済)
    End Sub

    Private Sub mvar許可読込_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar許可読込.Click
        検索開始(Enum状態.許可)
    End Sub

    Private Sub 検索開始(申請状態 As Enum状態)
        Dim sError As String = "000"
        Dim sError2 As String = ""
        Try
            mvarTSProg.Visible = True
            mvarTSLabel.Visible = True
            mvar受付日検索.Visible = False
            mvar許可読込.Visible = False

            sError = "001"
            TBL農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:農地Info].ID, V_大字.大字, V_小字.小字, [D:農地Info].地番, [D:農地Info].登記簿地目, [D:農地Info].現況地目, V_現況地目.名称 AS 現況地目名, [D:農地Info].農委地目ID, [D:農地Info].登記簿面積, [D:農地Info].実面積, [D:農地Info].田面積, [D:農地Info].畑面積, [D:農地Info].樹園地, [D:農地Info].採草放牧面積 " &
                                                            "FROM (([D:農地Info] LEFT JOIN V_現況地目 ON [D:農地Info].現況地目 = V_現況地目.ID) LEFT JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) LEFT JOIN V_小字 ON [D:農地Info].小字ID = V_小字.ID;")
            TBL農地.PrimaryKey = {TBL農地.Columns("ID")}

            sError = "002"
            TBL転用農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D_転用農地].ID, V_大字.大字, V_小字.小字, [D_転用農地].地番, [D_転用農地].登記簿地目, [D_転用農地].現況地目, V_現況地目.名称 AS 現況地目名, [D_転用農地].農委地目ID, [D_転用農地].登記簿面積, [D_転用農地].実面積, [D_転用農地].田面積, [D_転用農地].畑面積, [D_転用農地].樹園地, [D_転用農地].採草放牧面積 " &
                                                            "FROM (([D_転用農地] LEFT JOIN V_現況地目 ON [D_転用農地].現況地目 = V_現況地目.ID) LEFT JOIN V_大字 ON [D_転用農地].大字ID = V_大字.ID) LEFT JOIN V_小字 ON [D_転用農地].小字ID = V_小字.ID;")
            TBL転用農地.PrimaryKey = {TBL転用農地.Columns("ID")}

            sError = "003"
            TBL削除農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D_削除農地].ID, V_大字.大字, V_小字.小字, [D_削除農地].地番, [D_削除農地].登記簿地目, [D_削除農地].現況地目, V_現況地目.名称 AS 現況地目名, [D_削除農地].農委地目ID, [D_削除農地].登記簿面積, [D_削除農地].実面積, [D_削除農地].田面積, [D_削除農地].畑面積, [D_削除農地].樹園地, [D_削除農地].採草放牧面積 " &
                                                            "FROM (([D_削除農地] LEFT JOIN V_現況地目 ON [D_削除農地].現況地目 = V_現況地目.ID) LEFT JOIN V_大字 ON [D_削除農地].大字ID = V_大字.ID) LEFT JOIN V_小字 ON [D_削除農地].小字ID = V_小字.ID;")
            TBL削除農地.PrimaryKey = {TBL削除農地.Columns("ID")}

            '/*【20161025】解約分の追加*/
            sError = "005"
            Dim s検索日 As String = ""
            Dim n状態 As Integer = 0
            Select Case 申請状態
                Case Enum状態.受付
                    s検索日 = "受付年月日"
                    n状態 = 0
                Case Enum状態.許可
                    s検索日 = "許可年月日"
                    n状態 = 2
                Case Enum状態.受付中
                    s検索日 = "受付年月日"
                    n状態 = 0
                Case Enum状態.許可済
                    s検索日 = "受付年月日"
                    n状態 = 2
            End Select

            sError = "006"
            Dim sWhere As String = "30,31,311,40,50,51,52,60,61,62,602,600,180,200,210"
            Select Case mvar絞込.Text
                Case "全対象" : sWhere = "30,31,311,40,50,51,52,60,61,62,602,600,180,200,210"
                Case "農地法3条" : sWhere = "30,31,311"
                Case "農地法4条" : sWhere = "40"
                Case "農地法5条" : sWhere = "50,51,52"
                Case "基盤強化法" : sWhere = "60,61,62"
                Case "非農地証明" : sWhere = "602,600"
                Case "解約" : sWhere = "180,200,210"
            End Select


            TBL申請 = SysAD.DB(sLRDB).GetTableBySqlSelect(String.Format("SELECT D_申請.ID, D_申請.状態, D_申請.法令, D_申請.農地リスト, D_申請.名称, D_申請.受付番号, D_申請.許可番号, D_申請.受付年月日, D_申請.許可年月日, D_申請.公告年月日, D_申請.権利種類, D_申請.再設定, D_申請.申請理由A, D_申請.区分, D_申請.申請者A, D_申請.申請者B, D_申請.農地区分, D_申請.不許可例外, D_申請.始末書, [D:個人Info].氏名 AS 申請者B氏名, [D:個人Info].住所 AS 申請者B住所, [D:個人Info].農業改善計画認定 AS 申請者B認定状況, D_申請.申請者C, D_申請.経由法人ID, D_申請.代理人A, V_小作形態.名称 AS 権利の種類, D_申請.始期, D_申請.終期, D_申請.期間, D_申請.農業委員1, D_申請.農業委員2, D_申請.農業委員3, D_申請.完了報告年月日, D_申請.小作料, D_申請.小作料単位, D_申請.工事開始年1, D_申請.工事開始月1, D_申請.工事終了年1, D_申請.工事終了月1 FROM (D_申請 LEFT JOIN [D:個人Info] ON D_申請.申請者B = [D:個人Info].ID) LEFT JOIN V_小作形態 ON D_申請.権利種類 = V_小作形態.ID WHERE (((D_申請.状態)={3}) AND ((D_申請.法令) In ({4})) AND ((D_申請.{2})>=#{0}# And (D_申請.{2})<=#{1}#)) ORDER BY D_申請.法令, D_申請.{2};", mvar許可範囲開始.Value.ToShortDateString, mvar許可範囲終了.Value.ToShortDateString, s検索日, n状態, sWhere))                'App農地基本台帳.TBL申請.MergePlus(TBL申請)
            sError = "007"
            TBL土地履歴 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT D_土地履歴.申請ID, D_土地履歴.LID, V_大字.大字 AS 申請時大字, V_小字.小字 AS 申請時小字, D_土地履歴.申請時地番, V_現況地目.名称 AS 申請時地目, D_土地履歴.申請時実面積, D_土地履歴.申請時田面積, D_土地履歴.申請時畑面積, D_土地履歴.申請時樹園地, D_土地履歴.申請時採草放牧面積 FROM ((D_土地履歴 LEFT JOIN V_大字 ON D_土地履歴.申請時大字ID = V_大字.ID) LEFT JOIN V_小字 ON D_土地履歴.申請時小字ID = V_小字.ID) LEFT JOIN V_現況地目 ON D_土地履歴.申請時登記簿地目 = V_現況地目.ID WHERE(((D_土地履歴.申請ID) > 0)) GROUP BY D_土地履歴.申請ID, D_土地履歴.LID, V_大字.大字, V_小字.小字, D_土地履歴.申請時地番, V_現況地目.名称, D_土地履歴.申請時実面積, D_土地履歴.申請時田面積, D_土地履歴.申請時畑面積, D_土地履歴.申請時樹園地, D_土地履歴.申請時採草放牧面積;")

            sError = "008"
            Dim TBL農家 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM D:個人Info")
            TBL農家.PrimaryKey = {TBL農家.Columns("ID")}

            sError = "009"
            Dim TBL認定状況 As DataTable = App農地基本台帳.GetMasterView("農業改善計画認定項目").ToTable
            TBL認定状況.PrimaryKey = {TBL認定状況.Columns("ID")}

            sError = "010"
            Dim TBL農業委員 As DataTable = App農地基本台帳.GetMasterView("農業委員").ToTable
            TBL農業委員.PrimaryKey = {TBL農業委員.Columns("ID")}

            sError = "011"
            Dim TBL農地区分 As DataTable = App農地基本台帳.GetMasterView("農地区分").ToTable
            TBL農地区分.PrimaryKey = {TBL農地区分.Columns("ID")}

            sError = "012"
            Dim mvarTBL申請 As New DataTable
            CreateTBL申請(mvarTBL申請)

            sError = "013"
            Dim mvarTBL明細 As New DataTable
            CreateTBL明細(mvarTBL明細)

            '//受付中or許可済を受付にもどす
            If (申請状態 = Enum状態.受付中 OrElse 申請状態 = Enum状態.許可済) Then
                申請状態 = Enum状態.受付
            End If

            mvarTSProg.Value = 0
            mvarTSProg.Maximum = TBL申請.Rows.Count
            If TBL申請.Rows.Count > 0 Then
                For Each pRow As DataRow In TBL申請.Rows
                    sError = "014"
                    sError = pRow.Item("ID") & ":" & pRow.Item("名称")
                    Dim Ar筆リスト As Object = Split(pRow.Item("農地リスト").ToString, ";")
                    Dim Int筆数計 As Integer = 0
                    Dim Dec面積計 As Decimal = 0

                    Dim s代表地番 As String = ""
                    Dim nCount As Integer = 1
                    For n As Integer = 0 To UBound(Ar筆リスト)
                        Dim Ar筆情報 As Object = Split(Ar筆リスト(n), ".")
                        Dim pRowFind As DataRow = Nothing
                        Dim pRow履歴 As DataRow = Nothing

                        If UBound(Ar筆情報) > 0 Then
                            If IsNumeric(Ar筆情報(1)) Then
                                pRowFind = TBL農地.Rows.Find(CDec(Ar筆情報(1)))
                                If pRowFind Is Nothing Then
                                    pRowFind = TBL転用農地.Rows.Find(CDec(Ar筆情報(1)))
                                    If pRowFind Is Nothing Then
                                        pRowFind = TBL削除農地.Rows.Find(CDec(Ar筆情報(1)))
                                    End If
                                End If
                            End If
                        Else
                            'Stop
                        End If

                        sError = "015"
                        If Not pRowFind Is Nothing Then
                            Dim pTBL As DataTable = New DataView(TBL土地履歴, "[LID] = " & Val(pRowFind.Item("ID").ToString) & " AND [申請ID] = " & Val(pRow.Item("ID").ToString), "", DataViewRowState.CurrentRows).ToTable
                            pTBL.PrimaryKey = {pTBL.Columns("申請ID")}
                            pRow履歴 = pTBL.Rows.Find(pRow.Item("ID"))

                            Int筆数計 += 1
                            Dec面積計 += Get土地履歴("申請時実面積", "実面積", pRow履歴, pRowFind)
                        End If

                        sError = "016"
                        Dim Add明細Row As DataRow = mvarTBL明細.NewRow
                        With Add明細Row
                            '/*【20161025】レイアウトの変更（空欄なしへ）*/
                            '.Item("ID") = IIf(nCount = 1, Val(pRow.Item("ID").ToString), DBNull.Value)
                            '.Item("名称") = IIf(nCount = 1, pRow.Item("名称"), "")
                            '.Item("状態") = IIf(nCount = 1, "許可済み", "")
                            '.Item("許可番号") = IIf(nCount = 1, Val(pRow.Item("許可番号").ToString), DBNull.Value)
                            '.Item("許可年月日") = IIf(nCount = 1, CnvDate(pRow.Item("許可年月日")), DBNull.Value)
                            '.Item("転用目的") = IIf(nCount = 1, IIf(Not IsDBNull(pRow.Item("区分")), IIf(Val(pRow.Item("区分").ToString) = 1, "農振除外", IIf(Val(pRow.Item("区分").ToString) = 2, "用途区分変更", IIf(Val(pRow.Item("区分").ToString) = 3, "農振編入", ""))), pRow.Item("申請理由A").ToString), "")

                            .Item("ID") = Val(pRow.Item("ID").ToString)
                            .Item("法令") = Val(pRow.Item("法令").ToString)
                            .Item("名称") = pRow.Item("名称").ToString
                            .Item("状態") = IIf(申請状態 = Enum状態.受付, "受付中", "許可済み")
                            .Item("権利の種類") = pRow.Item("権利の種類").ToString

                            sError = "017"
                            Dim n期間年 As Integer = 999
                            If IsDBNull(pRow.Item("期間")) OrElse Val(pRow.Item("期間").ToString) = 0 Then
                                If Not IsDBNull(pRow.Item("始期")) AndAlso Not IsDBNull(pRow.Item("終期")) Then
                                    n期間年 = DateDiff(DateInterval.Year, pRow.Item("始期"), CDate(pRow.Item("終期")).AddDays(1))
                                End If
                            Else
                                n期間年 = Val(pRow.Item("期間").ToString)
                            End If

                            Dim n期間月 As Integer = 0
                            If Not IsDBNull(pRow.Item("始期")) AndAlso Not IsDBNull(pRow.Item("終期")) Then
                                n期間月 = DateDiff(DateInterval.Month, pRow.Item("始期"), CDate(pRow.Item("終期")).AddDays(1)) Mod 12
                            End If

                            If n期間年 = 999 Then
                                If Not IsDBNull(pRow.Item("始期")) AndAlso IsDBNull(pRow.Item("終期")) Then
                                    .Item("期間") = "永年"
                                Else
                                    .Item("期間") = ""
                                End If
                            Else
                                If IsDBNull(pRow.Item("期間")) OrElse Val(pRow.Item("期間").ToString) = 0 Then
                                    .Item("期間") = n期間年 & "年" & IIf(n期間月 > 0, n期間月 & "月", "") & "間"
                                Else
                                    .Item("期間") = n期間年 & "年間"
                                End If
                            End If

                            sError = "018"
                            .Item("始期") = CnvDate(pRow.Item("始期"))
                            .Item("終期") = CnvDate(pRow.Item("終期"))
                            .Item("小作料") = Val(pRow.Item("小作料").ToString)
                            .Item("小作料単位") = pRow.Item("小作料単位").ToString()

                            .Item("許可番号") = IIf(申請状態 = Enum状態.受付, Val(pRow.Item("受付番号").ToString), Val(pRow.Item("許可番号").ToString))
                            .Item("許可年月日") = IIf(申請状態 = Enum状態.受付, CnvDate(pRow.Item("受付年月日")), CnvDate(pRow.Item("許可年月日")))
                            .Item("公告年月日") = CnvDate(pRow.Item("公告年月日"))

                            Dim pRow申請者 As DataRow = TBL農家.Rows.Find(Val(pRow.Item("申請者A").ToString))
                            If pRow申請者 IsNot Nothing Then
                                .Item("申請者") = pRow申請者.Item("氏名").ToString
                            End If

                            Dim pRow代理人 As DataRow = TBL農家.Rows.Find(Val(pRow.Item("代理人A").ToString))
                            If pRow代理人 IsNot Nothing Then
                                .Item("代理人") = pRow代理人.Item("氏名").ToString
                                .Item("代理人住所") = pRow代理人.Item("住所").ToString
                            End If

                            sError = "019"
                            .Item("転用目的") = IIf(Not IsDBNull(pRow.Item("区分")), IIf(Val(pRow.Item("区分").ToString) = 1, "農振除外", IIf(Val(pRow.Item("区分").ToString) = 2, "用途区分変更", IIf(Val(pRow.Item("区分").ToString) = 3, "農振編入", ""))), pRow.Item("申請理由A").ToString)

                            Dim pRow農地区分 As DataRow = TBL農地区分.Rows.Find(Val(pRow.Item("農地区分").ToString))
                            .Item("農地区分") = pRow農地区分.Item("名称").ToString

                            .Item("不許可例外") = pRow.Item("不許可例外").ToString

                            If IsDBNull(pRow.Item("始末書")) Then
                                .Item("始末書") = "□"
                            Else
                                .Item("始末書") = IIf(pRow.Item("始末書") = True, "☑", "□")
                            End If

                            sError = "020"
                            Dim pFindRow As DataRow = TBL認定状況.Rows.Find(Val(pRow.Item("申請者B認定状況").ToString))
                            .Item("(借)認定状況") = pFindRow.Item("名称").ToString

                            Dim pRow農委 As DataRow = TBL農業委員.Rows.Find(Val(pRow.Item("農業委員1").ToString))
                            .Item("農業委員1") = pRow農委.Item("名称").ToString
                            pRow農委 = TBL農業委員.Rows.Find(Val(pRow.Item("農業委員2").ToString))
                            .Item("農業委員2") = pRow農委.Item("名称").ToString
                            pRow農委 = TBL農業委員.Rows.Find(Val(pRow.Item("農業委員3").ToString))
                            .Item("農業委員3") = pRow農委.Item("名称").ToString
                            If IsDBNull(pRow.Item("再設定")) AndAlso Val(pRow.Item("法令").ToString) = 61 Then
                                .Item("再設定") = "新"
                            ElseIf Val(pRow.Item("法令").ToString) = 61 Then
                                .Item("再設定") = IIf(pRow.Item("再設定") = True, "再", "新")
                            Else
                                .Item("再設定") = ""
                            End If

                            If Not pRowFind Is Nothing Then
                                sError = "021"
                                .Item("筆ID") = Val(pRowFind.Item("ID").ToString)
                                .Item("大字") = Get土地履歴("申請時大字", "大字", pRow履歴, pRowFind)
                                .Item("小字") = Get土地履歴("申請時小字", "小字", pRow履歴, pRowFind)
                                .Item("地番") = Get土地履歴("申請時地番", "地番", pRow履歴, pRowFind)
                                .Item("地目") = Get土地履歴("申請時地目", "現況地目名", pRow履歴, pRowFind)
                                .Item("面積") = Get土地履歴("申請時実面積", "実面積", pRow履歴, pRowFind)

                                .Item("田面積") = Get土地履歴("申請時田面積", "田面積", pRow履歴, pRowFind)
                                .Item("畑面積") = Get土地履歴("申請時畑面積", "畑面積", pRow履歴, pRowFind)
                                .Item("樹園地") = Get土地履歴("申請時樹園地", "樹園地", pRow履歴, pRowFind)
                                .Item("採草放牧面積") = Get土地履歴("申請時採草放牧面積", "採草放牧面積", pRow履歴, pRowFind)

                                If s代表地番 = "" Then
                                    s代表地番 = .Item("大字") & .Item("小字") & .Item("地番")
                                End If
                            Else
                                sError = "022"
                                If UBound(Ar筆情報) > 0 Then
                                    If IsNumeric(Ar筆情報(1)) Then
                                        .Item("筆ID") = CDec(Ar筆情報(1))
                                    Else
                                        .Item("筆ID") = 9999999999
                                    End If
                                Else
                                    .Item("筆ID") = 9999999999
                                End If

                                .Item("大字") = "筆情報なし"
                            End If

                            mvarTBL明細.Rows.Add(Add明細Row)
                        End With

                        nCount += 1

                        sError = "023"
                        mvarTSProg.Increment(1)
                        mvarTSLabel.Text = "データ読み込み中...(" & mvarTSProg.Value & "/" & mvarTSProg.Maximum & ")"
                        My.Application.DoEvents()
                    Next

                    sError = "024"
                    Dim Add申請Row As DataRow = mvarTBL申請.NewRow
                    With Add申請Row
                        .Item("ID") = Val(pRow.Item("ID").ToString)
                        .Item("法令") = Val(pRow.Item("法令").ToString)
                        .Item("名称") = pRow.Item("名称").ToString
                        .Item("状態") = IIf(申請状態 = Enum状態.受付, "受付中", "許可済み")
                        .Item("許可番号") = IIf(申請状態 = Enum状態.受付, Val(pRow.Item("受付番号").ToString), Val(pRow.Item("許可番号").ToString))
                        .Item("許可年月日") = IIf(申請状態 = Enum状態.受付, CnvDate(pRow.Item("受付年月日")), CnvDate(pRow.Item("許可年月日")))
                        .Item("筆数計") = Int筆数計
                        .Item("面積計") = Dec面積計
                        .Item("代表地番") = s代表地番
                        '.Item("農地区分") = IIf(Not IsDBNull(pRow.Item("区分")), IIf(Val(pRow.Item("区分").ToString) = 1, "農振除外", IIf(Val(pRow.Item("区分").ToString) = 2, "用途区分変更", IIf(Val(pRow.Item("区分").ToString) = 3, "農振編入", ""))), pRow.Item("申請理由A").ToString)
                        Dim pRow農地区分 As DataRow = TBL農地区分.Rows.Find(Val(pRow.Item("農地区分").ToString))
                        .Item("農地区分") = pRow農地区分.Item("名称").ToString
                        .Item("小作料") = Val(pRow.Item("小作料").ToString)
                        .Item("小作料単位") = pRow.Item("小作料単位").ToString
                        If Val(pRow.Item("工事開始年1").ToString) > 0 Then
                            .Item("工事開始年月") = String.Format("{0}年{1}月", pRow.Item("工事開始年1").ToString, pRow.Item("工事開始月1").ToString)
                        End If
                        If Val(pRow.Item("工事終了年1").ToString) > 0 Then
                            .Item("工事終了年月") = String.Format("{0}年{1}月", pRow.Item("工事終了年1").ToString, pRow.Item("工事終了月1").ToString)
                        End If
                        .Item("完了報告年月日") = CnvDate(pRow.Item("完了報告年月日"))
                        .Item("借受人氏名") = pRow.Item("申請者B氏名").ToString
                        .Item("借受人住所") = pRow.Item("申請者B住所").ToString

                        mvarTBL申請.Rows.Add(Add申請Row)
                    End With
                Next
            Else
                MsgBox("検索範囲に該当する申請がありません。")
            End If

            sError = "025"
            mvarGrid申請.SetDataView(mvarTBL申請, "", "")
            mvarGrid明細.SetDataView(mvarTBL明細, "", "")

            mvarTSProg.Visible = False
            mvarTSLabel.Visible = False
            mvar受付日検索.Visible = True
            mvar許可読込.Visible = True
        Catch ex As Exception
            MsgBox(sError & ":" & sError2 & ":" & ex.Message)
        End Try

    End Sub

    Private Function Get土地履歴(ByVal ValTrue As String, ByVal ValFalse As String, ByRef pRow履歴 As DataRow, ByRef pRowFind As DataRow)
        If pRow履歴 IsNot Nothing Then
            If Len(pRow履歴.Item(ValTrue).ToString) > 0 Or Val(pRow履歴.Item(ValTrue).ToString) > 0 Then
                Return pRow履歴.Item(ValTrue)
            Else
                Return pRowFind.Item(ValFalse)
            End If
        Else
            Return pRowFind.Item(ValFalse)
        End If
    End Function

    Private Function Get代表土地履歴(ByVal ValTrue As String, ByVal ValFalse As String, ByRef pRow履歴 As DataRow, ByRef pRowFind As DataRow)
        If pRow履歴 IsNot Nothing Then
            If Len(pRow履歴.Item(ValTrue).ToString) > 0 Or Val(pRow履歴.Item(ValTrue).ToString) > 0 Then
                Return pRow履歴.Item("申請時大字").ToString & pRow履歴.Item("申請時小字").ToString & pRow履歴.Item("申請時地番").ToString
            Else
                Return pRowFind.Item("大字").ToString & pRowFind.Item("小字").ToString & pRowFind.Item("地番").ToString
            End If
        Else
            Return pRowFind.Item("大字").ToString & pRowFind.Item("小字").ToString & pRowFind.Item("地番").ToString
        End If
    End Function

    Private Sub CreateTBL申請(ByRef pTBL As DataTable)
        With pTBL
            .Columns.Add("ID", GetType(Decimal))
            .Columns.Add("法令", GetType(String))
            .Columns.Add("名称", GetType(String))
            .Columns.Add("状態", GetType(String))
            .Columns.Add("許可番号", GetType(Integer))
            .Columns.Add("許可年月日", GetType(Date))
            .Columns.Add("筆数計", GetType(Integer))
            .Columns.Add("面積計", GetType(Decimal))
            .Columns.Add("代表地番", GetType(String))
            .Columns.Add("農地区分", GetType(String))
            .Columns.Add("借受人氏名", GetType(String))
            .Columns.Add("借受人住所", GetType(String))
            .Columns.Add("小作料", GetType(Decimal))
            .Columns.Add("小作料単位", GetType(String))
            .Columns.Add("工事開始年月", GetType(String))
            .Columns.Add("工事終了年月", GetType(String))
            .Columns.Add("完了報告年月日", GetType(Date))
        End With
    End Sub

    Private Sub CreateTBL明細(ByRef pTBL As DataTable)
        With pTBL
            .Columns.Add("ID", GetType(Decimal))
            .Columns.Add("法令", GetType(String))
            .Columns.Add("名称", GetType(String))
            .Columns.Add("状態", GetType(String))
            .Columns.Add("権利の種類", GetType(String))
            .Columns.Add("期間", GetType(String))
            .Columns.Add("始期", GetType(Date))
            .Columns.Add("終期", GetType(Date))
            .Columns.Add("小作料", GetType(Decimal))
            .Columns.Add("小作料単位", GetType(String))
            .Columns.Add("許可番号", GetType(Integer))
            .Columns.Add("許可年月日", GetType(Date))
            .Columns.Add("公告年月日", GetType(Date))
            .Columns.Add("申請者", GetType(String))
            .Columns.Add("代理人", GetType(String))
            .Columns.Add("代理人住所", GetType(String))
            .Columns.Add("転用目的", GetType(String))
            .Columns.Add("農地区分", GetType(String))
            .Columns.Add("不許可例外", GetType(String))
            .Columns.Add("始末書", GetType(String))
            .Columns.Add("(借)認定状況", GetType(String))
            .Columns.Add("農業委員1", GetType(String))
            .Columns.Add("農業委員2", GetType(String))
            .Columns.Add("農業委員3", GetType(String))
            .Columns.Add("再設定", GetType(String))
            .Columns.Add(" ", GetType(String))
            .Columns.Add("筆ID", GetType(Decimal))
            .Columns.Add("大字", GetType(String))
            .Columns.Add("小字", GetType(String))
            .Columns.Add("地番", GetType(String))
            .Columns.Add("地目", GetType(String))
            .Columns.Add("面積", GetType(Decimal))
            .Columns.Add("田面積", GetType(Decimal))
            .Columns.Add("畑面積", GetType(Decimal))
            .Columns.Add("樹園地", GetType(Decimal))
            .Columns.Add("採草放牧面積", GetType(Decimal))
        End With
    End Sub

    Private Function CnvDate(ByVal pValue As Object)
        If pValue IsNot Nothing AndAlso IsDate(pValue) Then
            Return Format(pValue, "yyyy/MM/dd")
        Else
            Return DBNull.Value
        End If
    End Function

    Private Enum Enum状態
        受付 = 0
        許可 = 2
        受付中 = 3
        許可済 = 4
    End Enum
End Class
