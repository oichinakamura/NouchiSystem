
Imports System.ComponentModel
Imports System.Threading
Imports HimTools2012
Imports HimTools2012.controls


Public Class C利用権終期管理
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Private mvarSPC As controls.SplitContainerEx
    Private mvarTabCtrl As controls.TabControlBase

    Private WithEvents mvar条件 As CPropertyGridPlus
    Private WithEvents mvar終期通知発行 As ToolStripButton
    Private WithEvents mvarExcel終期通知発行 As ToolStripButton
    Private mvar通知発行日 As controls.ToolStripDateTimePicker
    Private mvar受付締切日 As controls.ToolStripDateTimePicker
    Private mvar締切判定 As controls.ToolStripCheckBoxEx
    Private WithEvents mvarBtn開始 As Windows.Forms.Button

    Private WithEvents mvarXLPath As HimTools2012.controls.ToolStripSpringTextBox

    Private mvar条件Object As 条件Object
    Private mvarGridView As controls.DataGridViewWithDataView
    Private WithEvents mvar農家View As controls.DataGridViewWithDataView
    Private mvar農家TBL As DataTable
    Private WithEvents mvar全選択 As ToolStripButton
    Private WithEvents mvar全解除 As ToolStripButton

    Public Sub New()
        MyBase.New(True, True, "利用権終期管理", "利用権終期管理")
        mvar条件Object = New 条件Object

        mvarSPC = New controls.SplitContainerEx("利用権終期管理", "80%")
        mvarSPC.Dock = DockStyle.Fill
        ControlPanel.Add(mvarSPC)

        mvar条件 = New CPropertyGridPlus
        With mvar条件
            .Dock = DockStyle.Fill
            .SelectedObject = mvar条件Object
        End With

        mvarBtn開始 = New Windows.Forms.Button
        With mvarBtn開始
            .Text = "検索開始"
            .Dock = DockStyle.Top
        End With

        mvar終期通知発行 = New ToolStripButton
        With mvar終期通知発行
            .Text = "終期通知書発行"
        End With

        mvarExcel終期通知発行 = New ToolStripButton
        With mvarExcel終期通知発行
            .Text = "終期通知書発行(Excel)"
        End With

        mvar通知発行日 = New ToolStripDateTimePicker
        mvar受付締切日 = New ToolStripDateTimePicker
        mvar締切判定 = New ToolStripCheckBoxEx
        With mvar締切判定
            .Checked = True
        End With

        Me.ToolStrip.Items.AddRange(New ToolStripItem() {
                                    New ToolStripLabel("通知書発行日"),
                                    mvar通知発行日,
                                    New ToolStripLabel("受け付け締め切り締日"),
                                    mvar受付締切日,
                                    mvar締切判定,
                                    New ToolStripSeparator,
                                    mvar終期通知発行,
                                    mvarExcel終期通知発行
                                    })

        mvarSPC.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {mvarBtn開始, mvar条件})

        mvar農家View = New controls.DataGridViewWithDataView
        With mvar農家View
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
        End With

        mvar農家TBL = New DataTable
        With mvar農家TBL.Columns
            .Add(New DataColumn("印刷", GetType(Boolean)))
            .Add(New DataColumn("公告日", GetType(String)))
            .Add(New DataColumn("開始日", GetType(String)))
            .Add(New DataColumn("終了日", GetType(String)))
            .Add(New DataColumn("利用権設定者ID", GetType(Decimal)))
            .Add(New DataColumn("利用権設定者世帯ID", GetType(Decimal)))
            .Add(New DataColumn("利用権設定者氏名", GetType(String)))
            .Add(New DataColumn("利用権設定者名前", GetType(String)))
            .Add(New DataColumn("利用権設定者郵便番号", GetType(String)))
            .Add(New DataColumn("利用権設定者住所", GetType(String)))
            .Add(New DataColumn("利用権設定者集落", GetType(String)))
            .Add(New DataColumn("利用権受け者ID", GetType(Decimal)))
            .Add(New DataColumn("利用権受け者世帯ID", GetType(Decimal)))
            .Add(New DataColumn("利用権受け者氏名", GetType(String)))
            .Add(New DataColumn("利用権受け者郵便番号", GetType(String)))
            .Add(New DataColumn("利用権受け者住所", GetType(String)))
            .Add(New DataColumn("利用権受け者集落", GetType(String)))
            .Add(New DataColumn("利用権受け者個人法人の別", GetType(String)))
            .Add(New DataColumn("利用権管理者ID", GetType(Decimal)))
            .Add(New DataColumn("利用権管理者氏名", GetType(String)))
            .Add(New DataColumn("農地ID", GetType(String)))

            mvar農家TBL.PrimaryKey = New DataColumn() { .Item("公告日"), .Item("終了日"), .Item("利用権設定者ID"), .Item("利用権受け者ID")}
        End With
        mvar農家View.SetDataView(mvar農家TBL, "", "")

        mvarGridView = New controls.DataGridViewWithDataView
        With mvarGridView
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .AutoGenerateColumns = False
        End With

        With mvarGridView
            .AddColumnText("ID", "ID", "ID", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("大字", "大字", "大字", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("小字", "小字", "小字", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("所在", "所在", "所在", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("地番", "地番", "地番", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)

            .AddColumnText("所有者氏名", "利用権設定するもの氏名", "所有者氏名", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("所有者住所", "利用権設定するもの住所", "所有者住所", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)

            .AddColumnText("借受人氏名", "利用権設定をうけるもの氏名", "借受人氏名", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("借受人住所", "利用権設定をうけるもの住所", "借受人住所", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("借受人個人法人の別", "利用権設定をうけるものの個人法人の別", "借受人個人法人の別", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)

            .AddColumnText("管理者氏名", "農地所有者氏名", "管理者氏名", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("管理者住所", "農地所有者住所", "管理者住所", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)

            .AddColumnText("小作形態種別", "契約の内容", "小作形態種別", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)

            .AddColumnText("登記簿地目名", "登記簿地目", "登記簿地目名", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("現況地目名", "現況地目", "現況地目名", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)

            .AddColumnText("登記簿面積", "登記簿面積", "登記簿面積", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("実面積", "実面積", "実面積", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)

            .AddColumnText("小作料", "小作料", "小作料", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("小作料単位", "小作料単位", "小作料単位", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)

            .AddColumnDateTime("小作開始年月日", "開始日", "小作開始年月日", enumReadOnly.bReadOnly, , DataGridViewContentAlignment.MiddleLeft)
            .AddColumnDateTime("小作終了年月日", "終了日", "小作終了年月日", enumReadOnly.bReadOnly, , DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("経由農業生産法人名", "法人経由", "経由農業生産法人名", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("農振法区分名", "農振法区分名", "農振法区分名", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("備考", "備考", "備考", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleLeft)
        End With

        mvarTabCtrl = New controls.TabControlBase
        mvarTabCtrl.Dock = DockStyle.Fill

        mvarGridView.CodeMasterTable = App農地基本台帳.DataMaster.Body
        mvarSPC.Panel2.Controls.Add(mvarTabCtrl)

        'With mvarTabCtrl.AddNewPage(mvarGridView, "農地", "対象農地", False, False, True)
        '    mvarGridView.Createエクセル出力Ctrl(.ToolStrip)
        'End With
        Dim pPage01 As New controls.CTabPageWithToolStrip(False, True, "農地", "対象農地")
        pPage01.ControlPanel.Add(mvarGridView)
        mvarGridView.Createエクセル出力Ctrl(pPage01.ToolStrip)
        mvarTabCtrl.AddPage(pPage01)


        Dim pPage02 As New controls.CTabPageWithToolStrip(False, True, "終期通知対象", "終期通知対象")
        pPage02.ControlPanel.Add(mvar農家View)
        mvar農家View.Createエクセル出力Ctrl(pPage02.ToolStrip)
        mvar全選択 = New ToolStripButton("全選択")
        pPage02.ToolStrip.Items.Add(mvar全選択)
        mvar全解除 = New ToolStripButton("全解除")
        pPage02.ToolStrip.Items.Add(mvar全解除)
        mvarTabCtrl.AddPage(pPage02)

        Me.ToolStrip.ItemAdd("書式ファイル", New ToolStripButton("書式ファイル"), AddressOf FL)
        mvarXLPath = Me.ToolStrip.ItemAdd("書式ファイルパス", New HimTools2012.controls.ToolStripSpringTextBox())
        mvarXLPath.ReadOnly = True
        mvarXLPath.AutoSize = True
        mvarXLPath.Text = SysAD.ClickOnceSetupPath & "\" & SysAD.市町村.市町村名 & "\利用権終期通知書.xml"
    End Sub

    '20170522 DataGridViewがチェックされたことを知る
    'Private Sub DataGridView1_CurrentCellDirtyStateChanged(ByVal sender As Object, ByVal e As EventArgs) Handles mvar農家View.CurrentCellDirtyStateChanged
    '    If mvar農家View.CurrentCellAddress.X = 0 AndAlso mvar農家View.IsCurrentCellDirty Then
    '        'コミットする
    '        mvar農家View.CommitEdit(DataGridViewDataErrorContexts.Commit)
    '    End If
    'End Sub

    'Private Sub DataGridView1_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles mvar農家View.CellEndEdit
    '    With Me.mvar農家View
    '        If .Columns(e.ColumnIndex).GetType Is GetType(DataGridViewCheckBoxColumn) Then
    '            Dim checkboxcell As DataGridViewCell = .Item(e.ColumnIndex, e.RowIndex)
    '            Select Case Convert.ToString(.Item(e.ColumnIndex, e.RowIndex).Value)
    '                Case "", False
    '                    checkboxcell.Value = True
    '                Case True
    '                    checkboxcell.Value = False
    '            End Select
    '        End If
    '    End With
    'End Sub

    Private Class 条件Object
        Public Enum Enum終期検索開始時期
            なし = 0
            今から01月先 = 1
            今から02月先 = 2
            今から03月先 = 3
            今から04月先 = 4
            今から05月先 = 5
            今から06月先 = 6
            今から07月先 = 7
            今から08月先 = 8
            今から09月先 = 9
            今から10月先 = 10
            今から11月先 = 11
            今から12月先 = 12
            今から02年先 = 24
        End Enum

        Public Enum Enum終期検索期間
            なし = 0
            指定月から1ヶ月 = 1
            指定月から2ヶ月 = 2
            指定月から3ヶ月 = 3
            指定月から6ヶ月 = 6
            指定月から12ヶ月 = 12
        End Enum

        Public Enum EnumYesNo
            いいえ = 0
            はい = 1
            農地法のみ = 2
        End Enum

        Private mvar終期選択範囲 As New InputSupport.C日付範囲入力
        Private mvar終期検索開始時期 As Integer = 0
        Private mvar終期検索期間 As Integer = 0
        Private mvar農地法検索 As Integer = 0
        Private mvarTable As DataTable

        Public Sub New()
            mvarTable = New DataTable("利用権終期管理検索条件")
            mvarTable.Columns.Add("終期検索開始時期", GetType(Enum終期検索開始時期))
            mvarTable.Columns.Add("終期検索期間", GetType(Enum終期検索期間))


            Dim sFile As String = SysAD.ClickOnceSetupPath & "\終期検索条件.XML"
            If Not SysAD.IsClickOnceDeployed Then
                sFile = My.Application.Info.DirectoryPath & "\終期検索条件.XML"
            End If

            mvar終期選択範囲.範囲開始 = Now.Date
            mvar終期選択範囲.範囲終了 = Now.Date

            If IO.File.Exists(sFile) Then
                mvarTable.ReadXml(sFile)
                If mvarTable.Rows.Count > 0 Then
                    Me.終期検索開始時期 = mvarTable.Rows(0).Item("終期検索開始時期")
                    Me.終期検索期間 = mvarTable.Rows(0).Item("終期検索期間")
                End If
            End If

        End Sub

        <Category("01_終期選択条件")> <Description("検索する終期の期間を指示します。")>
        Public Property 終期検索期間 As Enum終期検索期間
            Get
                Return mvar終期検索期間
            End Get
            Set(ByVal value As Enum終期検索期間)
                mvar終期検索期間 = value

                If mvar終期検索期間 > 0 Then
                    Dim n As Integer = CType(value, Integer)
                    終期選択範囲.範囲終了 = DateAdd(DateInterval.Month, n, 終期選択範囲.範囲開始)

                End If

            End Set
        End Property

        <Category("01_終期選択条件")> <Description("検索する終期の絞り込みを行います。")>
        Public Property 終期選択範囲 As InputSupport.C日付範囲入力
            Get
                Return mvar終期選択範囲
            End Get
            Set(ByVal value As InputSupport.C日付範囲入力)
                mvar終期選択範囲 = value
            End Set
        End Property

        <Category("01_終期選択条件")> <Description("検索する終期の期間を指示します。")>
        Public Property 終期検索開始時期 As Enum終期検索開始時期
            Get
                Return mvar終期検索開始時期
            End Get
            Set(ByVal value As Enum終期検索開始時期)
                mvar終期検索開始時期 = value

                If value > 0 Then
                    Dim n As Integer = CType(value, Integer)
                    Dim def As TimeSpan = 終期選択範囲.範囲期間
                    終期選択範囲.範囲開始 = DateAdd(DateInterval.Month, n, Now.Date)

                    終期選択範囲.範囲終了 = 終期選択範囲.範囲開始 + def
                End If
            End Set
        End Property

        <Category("02_検索対象条件")> <Description("絞込み条件を変更します。")>
        Public Property 農地法を含む As EnumYesNo
            Get
                Return mvar農地法検索
            End Get
            Set(ByVal value As EnumYesNo)
                mvar農地法検索 = value
            End Set
        End Property

        Public Overrides Function ToString() As String
            Dim St As String = mvar終期選択範囲.ToStringSQL("小作終了年月日")

            Select Case 農地法を含む
                Case EnumYesNo.いいえ : St = St & " AND [小作地適用法]=2"
                Case EnumYesNo.はい
                Case EnumYesNo.農地法のみ : St = St & " AND [小作地適用法]=1"
            End Select
            Return St
        End Function

        Public Sub Save条件()
            mvarTable.Rows.Clear()
            mvarTable.Rows.Add({Me.終期検索開始時期, Me.終期検索期間})
            mvarTable.WriteXml(SysAD.ClickOnceSetupPath & "終期検索条件.XML", System.Data.XmlWriteMode.WriteSchema, True)
        End Sub

    End Class

    Private Sub mvar条件_PropertyValueChanged(ByVal s As Object, ByVal e As System.Windows.Forms.PropertyValueChangedEventArgs) Handles mvar条件.PropertyValueChanged
        Select Case e.ChangedItem.Label
            Case "終期検索開始時期"
                '        Dim n As Integer = CType(e.ChangedItem.Value, Integer)
                '        Dim def As TimeSpan = mvar条件Object.終期選択範囲.範囲期間
                '        mvar条件Object.終期選択範囲.範囲開始 = DateAdd(DateInterval.Month, n, Now.Date)
                '        mvar条件Object.終期選択範囲.範囲終了 = mvar条件Object.終期選択範囲.範囲開始 + def
            Case "終期検索期間"
                '        Dim n As Integer = CType(e.ChangedItem.Value, Integer)
                '        mvar条件Object.終期選択範囲.範囲終了 = DateAdd(DateInterval.Month, n, mvar条件Object.終期選択範囲.範囲開始)
            Case Else
        End Select
        mvar条件Object.Save条件()
    End Sub

    Private mvar農地 As DataTable

    Private Sub mvarBtn開始_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarBtn開始.Click
        With SysAD.DB(sLRDB)
            .ExecuteSQL("UPDATE [D:農地Info] SET [D:農地Info].小作開始年月日 = IIf(InStr([小作開始年月日],' ')>0,Left([小作開始年月日],InStr([小作開始年月日],' ')-1),[小作開始年月日]), [D:農地Info].小作終了年月日 = IIf(InStr([小作終了年月日],' ')>0,Left([小作終了年月日],InStr([小作終了年月日],' ')-1),[小作終了年月日]);")

            Dim sWhere As String = mvar条件Object.ToString
            If Len(sWhere) Then
                sWhere = "[自小作別]>0 AND (" & sWhere & ")"
                mvar農地 = .GetTableBySqlSelect("SELECT * FROM [D:農地Info] Where " & sWhere, "")
                App農地基本台帳.TBL農地.MergePlus(mvar農地)
                If Not App農地基本台帳.TBL農地.Columns.Contains("農振法区分名") Then
                    App農地基本台帳.TBL農地.Columns.Add("農振法区分名", GetType(String), "IIF(農振法区分=1,'農用地',IIF(農振法区分=2,'農振地',IIF(農振法区分=3,'農振外','-')))")
                End If
                mvarGridView.SetDataView(App農地基本台帳.TBL農地.Body, sWhere, "")

                mvar農家TBL.Rows.Clear()
                For Each pRow As DataRowView In mvarGridView.DataView
                    Dim s開始日 As String = pRow.Item("小作開始年月日")
                    Dim s終了日 As String = pRow.Item("小作終了年月日")
                    Dim nID As Decimal = 0
                    If IsDBNull(pRow.Item("管理者ID")) OrElse Val(pRow.Item("管理者ID").ToString) = 0 Then
                        nID = pRow.Item("所有者ID")
                    Else
                        nID = Val(pRow.Item("管理者ID").ToString)
                    End If

                    Dim pFind農家 As DataRow = mvar農家TBL.Rows.Find({s開始日, s終了日, nID, pRow.Item("借受人ID")})
                    If pFind農家 Is Nothing Then
                        Dim addRow = mvar農家TBL.NewRow()
                        addRow.Item("公告日") = s開始日
                        addRow.Item("開始日") = s開始日
                        addRow.Item("終了日") = s終了日

                        addRow.Item("農地ID") = pRow.Item("ID")

                        SetOwnerInfo(addRow, pRow)

                        addRow.Item("利用権受け者ID") = pRow.Item("借受人ID")
                        addRow.Item("利用権受け者世帯ID") = pRow.Item("借受世帯ID")
                        'pFind農家.Item("利用権受け者氏名") = pRow.Item("借受人氏名")
                        Dim p借受者 As DataRow = App農地基本台帳.TBL個人.FindRowByID(pRow.Item("借受人ID"))
                        If p借受者 IsNot Nothing Then
                            If addRow.Item("利用権受け者氏名").ToString.Length = 0 Then
                                addRow.Item("利用権受け者氏名") = p借受者.Item("氏名")
                            End If

                            If p借受者.Item("送付先郵便番号").ToString.Length > 0 Then
                                addRow.Item("利用権受け者郵便番号") = p借受者.Item("送付先郵便番号")
                            Else
                                addRow.Item("利用権受け者郵便番号") = p借受者.Item("郵便番号")
                            End If
                            If Trim(p借受者.Item("送付先住所").ToString).Length > 0 Then
                                addRow.Item("利用権受け者住所") = p借受者.Item("送付先住所")
                            Else
                                addRow.Item("利用権受け者住所") = p借受者.Item("住所")
                            End If
                            If p借受者.Item("行政区名").ToString.Length > 0 Then
                                addRow.Item("利用権受け者集落") = p借受者.Item("行政区名")
                            Else
                                addRow.Item("利用権受け者集落") = "                      "
                            End If
                            If p借受者.Item("性別").ToString.Length = 3 Then
                                addRow.Item("利用権受け者個人法人の別") = "法人"
                            Else
                                addRow.Item("利用権受け者個人法人の別") = "個人"
                            End If
                        End If

                        Dim pFind管理者 As DataRow = mvar農家TBL.Rows.Find({s開始日, s終了日, Val(pRow.Item("管理者ID").ToString), pRow.Item("管理者ID")})
                        If pFind管理者 Is Nothing Then
                            addRow.Item("利用権管理者ID") = pRow.Item("管理者ID")
                            addRow.Item("利用権管理者氏名") = pRow.Item("管理者氏名").ToString
                        End If

                        Try
                            mvar農家TBL.Rows.Add(addRow)
                        Catch ex As Exception

                        End Try

                    Else
                        pFind農家.Item("農地ID") = pFind農家.Item("農地ID") & "," & pRow.Item("ID")
                    End If
                Next
            End If
        End With
    End Sub

    Private Sub SetOwnerInfo(ByRef addRow As DataRow, ByVal pRow As DataRowView)
        With addRow
            Dim bManager As Boolean = False
            If IsDBNull(pRow.Item("管理者ID")) OrElse Val(pRow.Item("管理者ID").ToString) = 0 Then
                .Item("利用権設定者ID") = pRow.Item("所有者ID")
                .Item("利用権設定者世帯ID") = pRow.Item("所有世帯ID")
            Else
                .Item("利用権設定者ID") = Val(pRow.Item("管理者ID").ToString)
                .Item("利用権設定者世帯ID") = Val(pRow.Item("管理世帯ID").ToString)
                bManager = True
            End If
            .Item("利用権管理者ID") = Val(pRow.Item("管理者ID").ToString)

            Dim p利用権設定者 As DataRow = App農地基本台帳.TBL個人.FindRowByID(.Item("利用権設定者ID"))
            If p利用権設定者 IsNot Nothing Then
                If bManager = True Then
                    Dim p所有者 As DataRow = App農地基本台帳.TBL個人.FindRowByID(pRow.Item("所有者ID"))
                    If p所有者 IsNot Nothing Then
                        .Item("利用権設定者氏名") = String.Format("{0}({1})", p利用権設定者.Item("氏名").ToString, p所有者.Item("氏名").ToString)
                        .Item("利用権設定者名前") = p所有者.Item("氏名").ToString
                    Else
                        .Item("利用権設定者氏名") = p利用権設定者.Item("氏名").ToString
                        .Item("利用権設定者名前") = p所有者.Item("氏名").ToString
                    End If
                Else
                    Dim p所有者 As DataRow = App農地基本台帳.TBL個人.FindRowByID(pRow.Item("所有者ID"))
                    If p所有者 IsNot Nothing Then
                        .Item("利用権設定者名前") = p所有者.Item("氏名").ToString
                    End If
                    .Item("利用権設定者氏名") = p利用権設定者.Item("氏名").ToString
                End If

                If p利用権設定者.Item("送付先郵便番号").ToString.Length > 0 Then
                    .Item("利用権設定者郵便番号") = p利用権設定者.Item("送付先郵便番号")
                Else
                    .Item("利用権設定者郵便番号") = p利用権設定者.Item("郵便番号")
                End If

                If Trim(p利用権設定者.Item("送付先住所").ToString).Length > 0 Then
                    .Item("利用権設定者住所") = p利用権設定者.Item("送付先住所")
                Else
                    .Item("利用権設定者住所") = p利用権設定者.Item("住所")
                End If

                If p利用権設定者.Item("行政区名").ToString.Length > 0 Then
                    .Item("利用権設定者集落") = p利用権設定者.Item("行政区名")
                Else
                    .Item("利用権設定者集落") = "                      "
                End If
            End If
        End With
    End Sub

    Private Sub FL()
        With New OpenFileDialog
            .Filter = "*.xml|*.xml"
            If .ShowDialog = DialogResult.OK Then
                mvarXLPath.Text = .FileName
            End If
        End With
    End Sub

    Public Overrides ReadOnly Property PageCloseMode As controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.NoMessage
        End Get
    End Property

    Private Sub mvar終期通知発行_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar終期通知発行.Click
        終期通知書発行(False)
        農用地利用権設定申出書(True)
        MsgBox("終了しました。")
    End Sub

    Private Sub mvarExcel終期通知発行_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarExcel終期通知発行.Click
        終期通知書発行(True)
        農用地利用権設定申出書(True)
        MsgBox("終了しました。")
    End Sub

    Private Sub 終期通知書発行(ByVal pExcel出力 As Boolean)
        If mvar農家TBL IsNot Nothing AndAlso mvar農家TBL.Rows.Count > 0 Then
            Dim sFile As String = ""

            sFile = mvarXLPath.Text

            If IO.File.Exists(sFile) Then
                Dim sXMLOg As String = TextAdapter.LoadTextFile(sFile)
                Dim s会長名 As String = SysAD.DB(sLRDB).DBProperty("会長名")
                For Each pRow As DataRowView In New DataView(mvar農家TBL, "[印刷]=True", "", DataViewRowState.CurrentRows)
                    Dim sXML As String = sXMLOg
                    Dim list As New List(Of DateTime)

                    sXML = Replace(sXML, "{市町村}", SysAD.市町村.市町村名)
                    sXML = Replace(sXML, "{会長名}", s会長名)
                    sXML = Replace(sXML, "{公告日年月日}", 和暦Format(pRow.Item("公告日").ToString))
                    sXML = Replace(sXML, "{開始日年月日}", 和暦Format(pRow.Item("開始日").ToString))
                    sXML = Replace(sXML, "{終了日年月日}", 和暦Format(pRow.Item("終了日").ToString))
                    sXML = Replace(sXML, "{申請者Ａ名前}", pRow.Item("利用権設定者名前").ToString)
                    sXML = Replace(sXML, "{申請者Ａ氏名}", pRow.Item("利用権設定者氏名").ToString)

                    If InStr(pRow.Item("利用権設定者氏名").ToString, "(") Then
                        sXML = Replace(sXML, "{申請者Ａ送付先氏名}", Mid(pRow.Item("利用権設定者氏名").ToString, 1, InStr(pRow.Item("利用権設定者氏名").ToString, "(") - 1))
                    Else
                        sXML = Replace(sXML, "{申請者Ａ送付先氏名}", pRow.Item("利用権設定者氏名").ToString)
                    End If

                    sXML = Replace(sXML, "{申請者Ａ住所}", pRow.Item("利用権設定者住所").ToString)
                    sXML = Replace(sXML, "{申請者Ａ郵便番号}", pRow.Item("利用権設定者郵便番号").ToString)
                    If pRow.Item("利用権設定者集落").ToString <> "-" Then
                        sXML = Replace(sXML, "{申請者Ａ集落名}", pRow.Item("利用権設定者集落").ToString)
                    Else
                        sXML = Replace(sXML, "{申請者Ａ集落名}", "             ")
                    End If

                    sXML = Replace(sXML, "{申請者Ｂ氏名}", pRow.Item("利用権受け者氏名").ToString)
                    sXML = Replace(sXML, "{申請者Ｂ名前}", pRow.Item("利用権受け者氏名").ToString)
                    sXML = Replace(sXML, "{申請者Ｂ住所}", pRow.Item("利用権受け者住所").ToString)
                    sXML = Replace(sXML, "{申請者Ｂ郵便番号}", pRow.Item("利用権受け者郵便番号").ToString)
                    If pRow.Item("利用権受け者集落").ToString <> "-" Then
                        sXML = Replace(sXML, "{申請者Ｂ集落名}", pRow.Item("利用権受け者集落").ToString)
                    Else
                        sXML = Replace(sXML, "{申請者Ｂ集落名}", "             ")
                    End If

                    Dim s管理者内訳 As String = ""
                    Dim pView As New DataView(App農地基本台帳.TBL農地.Body, "[ID] IN (" & pRow.Item("農地ID") & ")", "[大字],[本番],[地番]", DataViewRowState.CurrentRows)
                    For n As Integer = 1 To 40
                        If n > pView.Count Then
                            sXML = Replace(sXML, "{所在地" & StringF.Right("00" & n, 2) & "}", "")
                            sXML = Replace(sXML, "{地目" & StringF.Right("00" & n, 2) & "}", "")
                            sXML = Replace(sXML, "{面積" & StringF.Right("00" & n, 2) & "}", "")
                            sXML = Replace(sXML, "{利用権等の種類" & StringF.Right("00" & n, 2) & "}", "")
                            sXML = Replace(sXML, "{小作料" & StringF.Right("00" & n, 2) & "}", "")
                            sXML = Replace(sXML, "{所有者" & StringF.Right("00" & n, 2) & "}", "")
                            sXML = Replace(sXML, "{備考" & StringF.Right("00" & n, 2) & "}", "")
                        Else
                            Dim pRowV As DataRowView = pView(n - 1)
                            If IsDBNull(pRowV.Item("大字ID")) OrElse Val(pRowV.Item("大字ID").ToString) <= 0 Then
                                sXML = Replace(sXML, "{所在地" & StringF.Right("00" & n, 2) & "}", pRowV.Item("所在").ToString & pRowV.Item("地番").ToString)
                            Else
                                sXML = Replace(sXML, "{所在地" & StringF.Right("00" & n, 2) & "}", pRowV.Item("土地所在"))
                            End If
                            sXML = Replace(sXML, "{地目" & StringF.Right("00" & n, 2) & "}", pRowV.Item("登記簿地目名").ToString)
                            sXML = Replace(sXML, "{面積" & StringF.Right("00" & n, 2) & "}", DecConv(pRowV.Item("登記簿面積")))
                            sXML = Replace(sXML, "{利用権等の種類" & StringF.Right("00" & n, 2) & "}", pRowV.Item("小作形態種別").ToString)

                            If Not IsDBNull(pRowV.Item("小作料")) AndAlso Not IsDBNull(pRowV.Item("小作料単位")) Then
                                sXML = Replace(sXML, "{小作料" & StringF.Right("00" & n, 2) & "}", "10aあたり " & DecConv(Val(pRowV.Item("小作料").ToString)) & pRowV.Item("小作料単位").ToString)
                            Else
                                sXML = Replace(sXML, "{小作料" & StringF.Right("00" & n, 2) & "}", "")
                            End If
                            sXML = Replace(sXML, "{所有者" & StringF.Right("00" & n, 2) & "}", pRowV.Item("所有者氏名").ToString)
                            sXML = Replace(sXML, "{備考" & StringF.Right("00" & n, 2) & "}", "")

                            If Not IsDBNull(pRowV.Item("管理者ID")) AndAlso Val(pRowV.Item("管理者ID").ToString) <> 0 Then
                                If Not IsDBNull(pRowV.Item("農地所有内訳")) Then
                                    Select Case Val(pRowV.Item("農地所有内訳").ToString)
                                        Case 0, 2 : s管理者内訳 = "代理人"
                                        Case 1 : s管理者内訳 = "管理人"
                                    End Select
                                End If
                            End If

                            If Not IsDBNull(pRowV.Item("小作終了年月日")) Then
                                list.Add(pRowV.Item("小作終了年月日"))
                            End If
                        End If
                    Next

                    If s管理者内訳.Length <> 0 Then
                        sXML = Replace(sXML, "{申請者Ａ管理者}", "（" & s管理者内訳 & "）" & pRow.Item("利用権管理者氏名").ToString)
                        sXML = Replace(sXML, "{申請者Ａ貸付管理者}", s管理者内訳 & " ： " & pRow.Item("利用権管理者氏名").ToString)
                    Else
                        sXML = Replace(sXML, "{申請者Ａ管理者}", "")
                        sXML = Replace(sXML, "{申請者Ａ貸付管理者}", "")
                    End If

                    Dim s出Page1 As String = GetPageXML(sXML, "出Page1")
                    Dim s出Page2 As String = GetPageXML(sXML, "出Page2")

                    Dim s受Page1 As String = GetPageXML(sXML, "受Page1")
                    Dim s受Page2 As String = GetPageXML(sXML, "受Page2")

                    If pView.Count < 16 Then
                        s出Page2 = ""
                        s受Page2 = ""
                    End If

                    sXML = Replace(sXML, "XX出Page1XX", s出Page1)
                    sXML = Replace(sXML, "XX出Page2XX", s出Page2)
                    sXML = Replace(sXML, "XX受Page1XX", s受Page1)
                    sXML = Replace(sXML, "XX受Page2XX", s受Page2)

                    sXML = Replace(sXML, "{通知年月日}", 和暦Format(mvar通知発行日.Value))
                    If mvar締切判定.Checked = True Then
                        sXML = Replace(sXML, "{締日年月日}", 和暦Format(mvar受付締切日.Value))
                    Else
                        list.Sort()
                        Dim oldDate As DateTime = New DateTime(list(0).Year, list(0).Month, 1)
                        Dim newDate = oldDate.AddDays(-1.0)
                        sXML = Replace(sXML, "{締日年月日}", 和暦Format(newDate))
                    End If

                    Select Case pExcel出力
                        Case True
                            Dim sOutPutFile As String = SysAD.OutputFolder & String.Format("\利用権終期通知書({0}→{1} {2}～{3}まで).xml", pRow.Item("利用権設定者氏名").ToString, pRow.Item("利用権受け者氏名").ToString, 和暦Format(pRow.Item("開始日").ToString), 和暦Format(pRow.Item("終了日").ToString))
                            TextAdapter.SaveTextFile(sOutPutFile, sXML)
                            System.Diagnostics.Process.Start(SysAD.OutputFolder)
                        Case False
                            Using pExcel As New Excel.Automation.ExcelAutomation
                                Dim sOutPutFile As String = SysAD.OutputFolder & "\利用権終期通知書.xml"
                                TextAdapter.SaveTextFile(sOutPutFile, sXML)
                                pExcel.PrintBook(sOutPutFile)
                            End Using
                    End Select


                Next
            Else
                MsgBox("指定様式がありません。「利用権終期通知書.xml」を作成してDeploymentPlaceに保存してください。")
            End If
        End If
    End Sub

    Private Sub 農用地利用権設定申出書(ByVal pExcel出力 As Boolean)
        If mvar農家TBL IsNot Nothing AndAlso mvar農家TBL.Rows.Count > 0 Then
            Dim sFile As String = SysAD.ClickOnceSetupPath & "\" & SysAD.市町村.市町村名 & "\農用地利用権設定申出書.xml"

            If IO.File.Exists(sFile) Then
                Dim sXMLOg As String = TextAdapter.LoadTextFile(sFile)

                For Each pRow As DataRowView In New DataView(mvar農家TBL, "[印刷]=True", "", DataViewRowState.CurrentRows)
                    Dim sXML As String = sXMLOg

                    Dim pViewB As New DataView(App農地基本台帳.TBL個人.Body, "[ID] IN (" & pRow.Item("利用権受け者ID") & ")", "", DataViewRowState.CurrentRows)
                    Dim pRowB As DataRowView = pViewB(0)
                    sXML = Replace(sXML, "{申請者Ｂ郵便番号}", pRow.Item("利用権受け者郵便番号").ToString)
                    sXML = Replace(sXML, "{申請者Ｂ住所}", pRow.Item("利用権受け者住所").ToString)
                    sXML = Replace(sXML, "{申請者Ｂ年齢}", GetAge(pRowB.Item("生年月日")))
                    Select Case Val(pRowB.Item("性別").ToString)
                        Case 0 : sXML = Replace(sXML, "{申請者Ｂ性別}", "男")
                        Case 1 : sXML = Replace(sXML, "{申請者Ｂ性別}", "女")
                        Case Else : sXML = Replace(sXML, "{申請者Ｂ性別}", "")
                    End Select
                    sXML = Replace(sXML, "{申請者Ｂ電話番号}", IIf(Val(pRowB.Item("電話番号").ToString) = 0, "", pRowB.Item("電話番号").ToString))

                    Dim pViewA As New DataView(App農地基本台帳.TBL個人.Body, "[ID] IN (" & pRow.Item("利用権設定者ID") & ")", "", DataViewRowState.CurrentRows)
                    Dim pRowA As DataRowView = pViewA(0)
                    sXML = Replace(sXML, "{申請者Ａ郵便番号}", pRow.Item("利用権設定者郵便番号").ToString)
                    sXML = Replace(sXML, "{申請者Ａ住所}", pRow.Item("利用権設定者住所").ToString)
                    sXML = Replace(sXML, "{申請者Ａ年齢}", GetAge(pRowA.Item("生年月日")))
                    Select Case Val(pRowA.Item("性別").ToString)
                        Case 0 : sXML = Replace(sXML, "{申請者Ａ性別}", "男")
                        Case 1 : sXML = Replace(sXML, "{申請者Ａ性別}", "女")
                        Case Else : sXML = Replace(sXML, "{申請者Ａ性別}", "")
                    End Select
                    sXML = Replace(sXML, "{申請者Ａ電話番号}", IIf(Val(pRowA.Item("電話番号").ToString) = 0, "", pRowA.Item("電話番号").ToString))

                    Dim pView As New DataView(App農地基本台帳.TBL農地.Body, "[ID] IN (" & pRow.Item("農地ID") & ")", "[大字],[本番],[地番]", DataViewRowState.CurrentRows)
                    For n As Integer = 1 To 10
                        If n > pView.Count Then
                            sXML = Replace(sXML, "{大字" & StringF.Right("00" & n, 2) & "}", "")
                            sXML = Replace(sXML, "{小字" & StringF.Right("00" & n, 2) & "}", "")
                            sXML = Replace(sXML, "{地番" & StringF.Right("00" & n, 2) & "}", "")
                            sXML = Replace(sXML, "{地目" & StringF.Right("00" & n, 2) & "}", "")
                            sXML = Replace(sXML, "{面積" & StringF.Right("00" & n, 2) & "}", "")
                        Else
                            Dim pRowV As DataRowView = pView(n - 1)
                            If IsDBNull(pRowV.Item("大字ID")) OrElse Val(pRowV.Item("大字ID").ToString) <= 0 Then
                                sXML = Replace(sXML, "{大字" & StringF.Right("00" & n, 2) & "}", pRowV.Item("所在").ToString)
                            Else
                                sXML = Replace(sXML, "{大字" & StringF.Right("00" & n, 2) & "}", pRowV.Item("大字").ToString)
                            End If
                            sXML = Replace(sXML, "{小字" & StringF.Right("00" & n, 2) & "}", pRowV.Item("小字").ToString)
                            sXML = Replace(sXML, "{地番" & StringF.Right("00" & n, 2) & "}", pRowV.Item("地番").ToString)
                            sXML = Replace(sXML, "{地目" & StringF.Right("00" & n, 2) & "}", pRowV.Item("現況地目名").ToString)
                            sXML = Replace(sXML, "{面積" & StringF.Right("00" & n, 2) & "}", DecConv(pRowV.Item("実面積")))
                        End If
                    Next

                    Dim pView家族 As DataView = New DataView(App農地基本台帳.TBL個人.Body, "[住民区分] IN (0,110) AND [世帯ID]<>0 AND [世帯ID]=" & Val(pRow.Item("利用権受け者世帯ID").ToString), "続柄1,続柄2,続柄2", DataViewRowState.CurrentRows)
                    For n As Integer = 1 To 6
                        If pView家族.Count >= n Then
                            Dim pRow家族 As DataRowView = pView家族(n - 1)
                            sXML = Replace(sXML, "{申請者Ｂ世帯員氏名0" & n & "}", pRow家族.Item("氏名").ToString)
                            sXML = Replace(sXML, "{申請者Ｂ世帯員年齢0" & n & "}", GetAge(pRow家族.Item("生年月日")))
                            Select Case Val(pRow家族.Item("性別").ToString)
                                Case 0 : sXML = Replace(sXML, "{申請者Ｂ世帯員性別0" & n & "}", "男")
                                Case 1 : sXML = Replace(sXML, "{申請者Ｂ世帯員性別0" & n & "}", "女")
                                Case Else : sXML = Replace(sXML, "{申請者Ｂ世帯員性別0" & n & "}", "")
                            End Select
                            sXML = Replace(sXML, "{申請者Ｂ世帯員続柄0" & n & "}", CObj個人.Get続柄(pRow家族.Row))
                            sXML = Replace(sXML, "{申請者Ｂ世帯員職業0" & n & "}", pRow家族.Item("職業").ToString)
                            If IsDBNull(pRow家族.Item("農業従事日数")) OrElse Val(pRow家族.Item("農業従事日数").ToString) = 0 Then
                                sXML = Replace(sXML, "{申請者Ｂ世帯員従事日数0" & n & "}", "")
                            Else
                                sXML = Replace(sXML, "{申請者Ｂ世帯員従事日数0" & n & "}", pRow家族.Item("農業従事日数").ToString)
                            End If
                        Else
                            sXML = Replace(sXML, "{申請者Ｂ世帯員氏名0" & n & "}", "")
                            sXML = Replace(sXML, "{申請者Ｂ世帯員年齢0" & n & "}", "")
                            sXML = Replace(sXML, "{申請者Ｂ世帯員性別0" & n & "}", "")
                            sXML = Replace(sXML, "{申請者Ｂ世帯員続柄0" & n & "}", "")
                            sXML = Replace(sXML, "{申請者Ｂ世帯員職業0" & n & "}", "")
                            sXML = Replace(sXML, "{申請者Ｂ世帯員従事日数0" & n & "}", "")
                        End If
                    Next

                    借受人経営面積(sXML, pRow)
                    所有者経営面積(sXML, pRow)

                    Select Case pExcel出力
                        Case True
                            Dim sOutPutFile As String = SysAD.OutputFolder & String.Format("\農用地利用権設定申出書({0}→{1} {2}まで).xml", pRow.Item("利用権設定者氏名").ToString, pRow.Item("利用権受け者氏名").ToString, 和暦Format(pRow.Item("終了日").ToString))
                            TextAdapter.SaveTextFile(sOutPutFile, sXML)
                            System.Diagnostics.Process.Start(SysAD.OutputFolder)
                        Case False
                            Using pExcel As New Excel.Automation.ExcelAutomation
                                Dim sOutPutFile As String = SysAD.OutputFolder & "\農用地利用権設定申出書.xml"
                                TextAdapter.SaveTextFile(sOutPutFile, sXML)
                                pExcel.PrintBook(sOutPutFile)
                            End Using
                    End Select
                Next
            Else
                'MsgBox("指定様式がありません。「農用地権利設定申出書.xml」を作成してDeploymentPlaceに保存してください。")
            End If
        End If
    End Sub

    Private Sub 借受人経営面積(ByRef sXML As String, ByRef pRow As DataRowView)
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(String.Format("SELECT * FROM [D:農地Info] WHERE [所有者ID]={0} Or [借受人ID]={0}", Val(pRow.Item("利用権受け者ID").ToString)))
        App農地基本台帳.TBL農地.MergePlus(pTBL)
        'Dim pTBL所有農地 As DataView = New DataView(App農地基本台帳.TBL農地.Body, "([所有世帯ID]=" & Val(pRow.Item("利用権受け者世帯ID").ToString) & ") AND [自小作別]<1", "", DataViewRowState.CurrentRows)
        Dim pTBL所有農地 As DataView = New DataView(App農地基本台帳.TBL農地.Body, "([所有者ID]=" & Val(pRow.Item("利用権受け者ID").ToString) & " AND [自小作別]<1) OR ([所有者ID]=" & Val(pRow.Item("利用権受け者ID").ToString) & " AND [借受人ID]=" & Val(pRow.Item("利用権受け者ID").ToString) & " AND [自小作別]>0)", "", DataViewRowState.CurrentRows)
        'Dim pTBL借受農地 As DataView = New DataView(App農地基本台帳.TBL農地.Body, "([借受世帯ID]=" & Val(pRow.Item("利用権受け者世帯ID").ToString) & ") AND [自小作別]>0", "", DataViewRowState.CurrentRows)
        Dim pTBL借受農地 As DataView = New DataView(App農地基本台帳.TBL農地.Body, "([借受人ID]=" & Val(pRow.Item("利用権受け者ID").ToString) & ") AND ([所有者ID]<>" & Val(pRow.Item("利用権受け者ID").ToString) & ") AND [自小作別]>0", "", DataViewRowState.CurrentRows)
        Dim pTBL貸付農地 As DataView = New DataView(App農地基本台帳.TBL農地.Body, "([所有者ID]=" & Val(pRow.Item("利用権受け者ID").ToString) & ") AND ([借受人ID]<>" & pRow.Item("利用権受け者ID") & ") AND ([自小作別]>0)", "", DataViewRowState.CurrentRows)

        Dim 自田面積 As Decimal = 0
        Dim 自畑面積 As Decimal = 0
        Dim 自樹面積 As Decimal = 0
        For Each pOwnweRow As DataRowView In pTBL所有農地
            If Val(pOwnweRow.Item("田面積").ToString) > 0 Then : 自田面積 += Val(pOwnweRow.Item("田面積").ToString)
            ElseIf Val(pOwnweRow.Item("畑面積").ToString) > 0 Then : 自畑面積 += Val(pOwnweRow.Item("畑面積").ToString)
            ElseIf Val(pOwnweRow.Item("樹園地").ToString) > 0 Then : 自樹面積 += Val(pOwnweRow.Item("樹園地").ToString)
            End If
        Next
        sXML = Replace(sXML, "{申請者Ｂ自作地田}", 自田面積)
        sXML = Replace(sXML, "{申請者Ｂ自作地畑}", 自畑面積)
        sXML = Replace(sXML, "{申請者Ｂ自作地樹}", 自樹面積)

        Dim 小田面積 As Decimal = 0
        Dim 小畑面積 As Decimal = 0
        Dim 小樹面積 As Decimal = 0
        For Each pTenancyRow As DataRowView In pTBL借受農地
            If Val(pTenancyRow.Item("田面積").ToString) > 0 Then : 小田面積 += Val(pTenancyRow.Item("田面積").ToString)
            ElseIf Val(pTenancyRow.Item("畑面積").ToString) > 0 Then : 小畑面積 += Val(pTenancyRow.Item("畑面積").ToString)
            ElseIf Val(pTenancyRow.Item("樹園地").ToString) > 0 Then : 小樹面積 += Val(pTenancyRow.Item("樹園地").ToString)
            End If
        Next
        sXML = Replace(sXML, "{申請者Ｂ小作地田}", 小田面積)
        sXML = Replace(sXML, "{申請者Ｂ小作地畑}", 小畑面積)
        sXML = Replace(sXML, "{申請者Ｂ小作地樹}", 小樹面積)

        Dim 貸田面積 As Decimal = 0
        Dim 貸畑面積 As Decimal = 0
        Dim 貸樹面積 As Decimal = 0
        For Each pLoanRow As DataRowView In pTBL貸付農地
            If Val(pLoanRow.Item("田面積").ToString) > 0 Then : 貸田面積 += Val(pLoanRow.Item("田面積").ToString)
            ElseIf Val(pLoanRow.Item("畑面積").ToString) > 0 Then : 貸畑面積 += Val(pLoanRow.Item("畑面積").ToString)
            ElseIf Val(pLoanRow.Item("樹園地").ToString) > 0 Then : 貸樹面積 += Val(pLoanRow.Item("樹園地").ToString)
            End If
        Next
        sXML = Replace(sXML, "{申請者Ｂ貸付地田}", 貸田面積)
        sXML = Replace(sXML, "{申請者Ｂ貸付地畑}", 貸畑面積)
        sXML = Replace(sXML, "{申請者Ｂ貸付地樹}", 貸樹面積)
    End Sub

    Private Sub 所有者経営面積(ByRef sXML As String, ByRef pRow As DataRowView)
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(String.Format("SELECT * FROM [D:農地Info] WHERE [所有者ID]={0} Or [借受人ID]={0}", Val(pRow.Item("利用権設定者ID").ToString)))
        App農地基本台帳.TBL農地.MergePlus(pTBL)

        Dim pTBL所有農地 As DataView
        Dim pTBL借受農地 As DataView
        Dim pTBL貸付農地 As DataView
        If SysAD.市町村.市町村名 = "三股町" Then 'テスト
            pTBL所有農地 = New DataView(App農地基本台帳.TBL農地.Body, "([所有者ID]=" & Val(pRow.Item("利用権設定者ID").ToString) & " AND [自小作別]<1) OR ([所有者ID]=" & Val(pRow.Item("利用権設定者ID").ToString) & " AND [借受人ID]=" & Val(pRow.Item("利用権設定者ID").ToString) & " AND [自小作別]>0)", "", DataViewRowState.CurrentRows)
            pTBL借受農地 = New DataView(App農地基本台帳.TBL農地.Body, "([借受人ID]=" & Val(pRow.Item("利用権設定者ID").ToString) & ") AND ([所有者ID]<>" & Val(pRow.Item("利用権設定者ID").ToString) & ") AND [自小作別]>0", "", DataViewRowState.CurrentRows)
            pTBL貸付農地 = New DataView(App農地基本台帳.TBL農地.Body, "([所有者ID]=" & Val(pRow.Item("利用権設定者ID").ToString) & ") AND ([借受人ID]<>" & Val(pRow.Item("利用権設定者ID").ToString) & ") AND ([自小作別]>0)", "", DataViewRowState.CurrentRows)
        Else
            pTBL所有農地 = New DataView(App農地基本台帳.TBL農地.Body, "([所有世帯ID]=" & Val(pRow.Item("利用権設定者世帯ID").ToString) & ") AND [自小作別]<1", "", DataViewRowState.CurrentRows)
            pTBL借受農地 = New DataView(App農地基本台帳.TBL農地.Body, "([借受世帯ID]=" & Val(pRow.Item("利用権設定者世帯ID").ToString) & ") AND [自小作別]>0", "", DataViewRowState.CurrentRows)
            pTBL貸付農地 = New DataView(App農地基本台帳.TBL農地.Body, "([所有者ID]=" & Val(pRow.Item("利用権設定者ID").ToString) & ") AND ([借受人ID]<>" & Val(pRow.Item("利用権設定者ID").ToString) & ") AND ([自小作別]>0)", "", DataViewRowState.CurrentRows)
        End If

        Dim 自田面積 As Decimal = 0
        Dim 自畑面積 As Decimal = 0
        Dim 自樹面積 As Decimal = 0
        For Each pOwnweRow As DataRowView In pTBL所有農地
            If Val(pOwnweRow.Item("田面積").ToString) > 0 Then : 自田面積 += Val(pOwnweRow.Item("田面積").ToString)
            ElseIf Val(pOwnweRow.Item("畑面積").ToString) > 0 Then : 自畑面積 += Val(pOwnweRow.Item("畑面積").ToString)
            ElseIf Val(pOwnweRow.Item("樹園地").ToString) > 0 Then : 自樹面積 += Val(pOwnweRow.Item("樹園地").ToString)
            End If
        Next
        sXML = Replace(sXML, "{申請者Ａ自作地田}", 自田面積)
        sXML = Replace(sXML, "{申請者Ａ自作地畑}", 自畑面積)
        sXML = Replace(sXML, "{申請者Ａ自作地樹}", 自樹面積)

        Dim 小田面積 As Decimal = 0
        Dim 小畑面積 As Decimal = 0
        Dim 小樹面積 As Decimal = 0
        For Each pTenancyRow As DataRowView In pTBL借受農地
            If Val(pTenancyRow.Item("田面積").ToString) > 0 Then : 小田面積 += Val(pTenancyRow.Item("田面積").ToString)
            ElseIf Val(pTenancyRow.Item("畑面積").ToString) > 0 Then : 小畑面積 += Val(pTenancyRow.Item("畑面積").ToString)
            ElseIf Val(pTenancyRow.Item("樹園地").ToString) > 0 Then : 小樹面積 += Val(pTenancyRow.Item("樹園地").ToString)
            End If
        Next
        sXML = Replace(sXML, "{申請者Ａ小作地田}", 小田面積)
        sXML = Replace(sXML, "{申請者Ａ小作地畑}", 小畑面積)
        sXML = Replace(sXML, "{申請者Ａ小作地樹}", 小樹面積)

        Dim 貸田面積 As Decimal = 0
        Dim 貸畑面積 As Decimal = 0
        Dim 貸樹面積 As Decimal = 0
        For Each pLoanRow As DataRowView In pTBL貸付農地
            If Val(pLoanRow.Item("田面積").ToString) > 0 Then : 貸田面積 += Val(pLoanRow.Item("田面積").ToString)
            ElseIf Val(pLoanRow.Item("畑面積").ToString) > 0 Then : 貸畑面積 += Val(pLoanRow.Item("畑面積").ToString)
            ElseIf Val(pLoanRow.Item("樹園地").ToString) > 0 Then : 貸樹面積 += Val(pLoanRow.Item("樹園地").ToString)
            End If
        Next
        sXML = Replace(sXML, "{申請者Ａ貸付地田}", 貸田面積)
        sXML = Replace(sXML, "{申請者Ａ貸付地畑}", 貸畑面積)
        sXML = Replace(sXML, "{申請者Ａ貸付地樹}", 貸樹面積)
    End Sub

    Private Function GetAge(ByVal pValue As Object) As Integer
        If IsDBNull(pValue) Then
            Return 0
            'ElseIf pValue = 0 Then
            '    Return 0
        Else
            Dim nNow As Integer = CInt(Format(Now, "yyyyMMdd"))
            Dim nDay As Integer = CInt(Format(pValue, "yyyyMMdd"))

            Return CInt(Fix((nNow - nDay) / 10000))
        End If
    End Function

    Private Function GetPageXML(ByRef sXML As String, ByVal sKey As String) As String
        Dim sPage01 As String = StringF.Mid(sXML, InStr(sXML, " <Worksheet ss:Name=""" & sKey & """>"))
        sPage01 = StringF.Left(sPage01, InStr(sPage01, " </Worksheet>") + Len(" </Worksheet>") + 1)
        sXML = Replace(sXML, sPage01, "XX" & sKey & "XX")
        Return sPage01
    End Function

    Private Sub mvar全選択_Click(sender As Object, e As System.EventArgs) Handles mvar全選択.Click
        For Each pRow As DataRow In mvar農家TBL.Rows
            If IsDBNull(pRow.Item("印刷")) OrElse Not pRow.Item("印刷") Then
                pRow.Item("印刷") = True
            End If
        Next
    End Sub

    Private Sub mvar全解除_Click(sender As Object, e As System.EventArgs) Handles mvar全解除.Click
        For Each pRow As DataRow In mvar農家TBL.Rows
            If IsDBNull(pRow.Item("印刷")) OrElse pRow.Item("印刷") Then
                pRow.Item("印刷") = False
            End If
        Next
    End Sub

    Private Function DecConv(ByVal pDec As Decimal) As String
        If Fix(pDec) = pDec Then
            Return Val(pDec).ToString("#,##0")
        Else
            Return Val(pDec).ToString("#,##0.##") '小数点第2位まで表示
        End If

        Return pDec
    End Function

End Class

Public Class C利用権終期通知書
    Private mvarKey As String
    Private dtStart As Date
    Private dtEnd As Date
    Private dtPrintDay As Date
    Private dtXX As Date
    Private mvarData As String
    Private mvarTPH As Single
    Private mvarPn(0 To 2) As String     '   Private mvarPn(1 To 2) As String
    Private mvarQn(0 To 3) As String     '   Private mvarQn(1 To 3) As String
    Private mvarRn(0 To 2) As String     '   Private mvarRn(1 To 2) As String
    Private b封筒 As Boolean
    Private s連絡先 As String
    Private mvar〆日 As Integer
    Private dt提出 As Date


    Public Function DataInit(Optional ByVal sParam As String = "") As Boolean
        '    Dim Ar As Object
        '    Dim Rs As RecordsetEx
        '    Dim pNode As Object = Nothing  'MSComctlLib.Node
        '    Dim pDic As CDataDictionary = Nothing
        '    Dim sSQL As String

        '    Ar = Split(sParam, ";")
        '    If IsDate(Ar(0)) Then dtPrintDay = CDate(Ar(0))
        '    If Ar(2) Then
        '        mvar〆日 = Val(Ar(2))
        '    End If

        '    If mvar〆日 = 0 Then mvar〆日 = 1
        '    dtStart = DateSerial(Year(dtPrintDay), Month(dtPrintDay) + SysAD.DB(sLRDB).DBProperty("利用権終期検索開始"), 1)
        '    'dtEnd = DateSerial(Year(dtStart), Month(dtStart) + SysAD.DB(sLRDB).DBProperty("利用権終期検索終了"), 1) - 1
        '    dtXX = DateSerial(Year(dtPrintDay), Month(dtPrintDay) + 2, mvar〆日)

        '    'dt提出 = 0
        '    If UBound(Ar) > 2 Then
        '        dt提出 = CDate(Ar(3))
        '    End If

        '    mvarParent.UsedTreeView = True

        '    mvarData = SysAD.Func.s市町村名
        '    Select Case SysAD.Func.s市町村名
        '        Case "大崎町"
        '            mvarTPH = 8
        '            '1 ～ 4
        '            mvarPn(1) = "耕作者"
        '            mvarPn(2) = "所有者"
        '            mvarRn(1) = "継続して貸し付けたい場合は"
        '            mvarRn(2) = "引き続き耕作する希望があれば"
        '            mvarQn(1) = "を設定した"
        '            mvarQn(2) = "の設定を受けた"
        '            mvarQn(3) = "(契約されない場合は、利用権は自動的に消滅します。)"
        '        Case Else
        '            If SysAD.Func.s市町村名 = "枕崎市" Then
        '                mvarTPH = 8
        '            Else
        '                mvarTPH = 10
        '            End If

        '            '1 ～ 4
        '            mvarPn(1) = "耕作者"
        '            mvarPn(2) = "所有者"
        '            mvarRn(1) = "貸付けたい"
        '            mvarRn(2) = "耕作したい"
        '            mvarQn(1) = "を設定した"
        '            mvarQn(2) = "の設定を受けた"
        '            mvarQn(3) = ""
        '    End Select

        '    Select Case SysAD.Func.s市町村名
        '        Case "姶良市"
        '            b封筒 = True
        '            mvarData = mvarData & ";;"
        '            mvarData = mvarData & ";" & SysAD.Func.s市町村名 & "農業委員会;"
        '        Case "宗像市"
        '            b封筒 = True
        '            mvarData = mvarData & ";;"
        '            mvarData = mvarData & ";" & SysAD.Func.s市町村名 & "長;(農業振興課)"
        '        Case Else
        '            b封筒 = False
        '            mvarData = mvarData & ";;"
        '            mvarData = mvarData & ";" & SysAD.Func.s市町村名 & "農業委員会会長;" & SysAD.DB(sLRDB).DBProperty("会長名")
        '    End Select


        '    '5 ～
        '    If SysAD.DB(sLRDB).DBProperty("利用権終期文字列") = "" Then
        '        mvarData = mvarData & ";" & "%D1%付け公告の農用地利用集積計画によって利用権を設定した下記の土地は、%D2%を持って貸借契約の期間が終了しますので通知します。\n　また、%Pn%の%An%氏にも契約が終了する旨、連絡いたしましたので\n申し添えます｡なお、契約期間が終了する土地を継続して%Rn%場合、農業委員等を通じて%D3%までに農業委員会へ申し出て下さい｡"
        '    Else
        '        mvarData = mvarData & ";" & SysAD.DB(sLRDB).DBProperty("利用権終期文字列")
        '    End If


        '    mvarData = mvarData & ";" & Ar(1)

        '    Rs = SysAD.DB(sLRDB).GetRecordsetEx("SELECT V_農地.土地所在, V_農地.所有者ID, V_農地.管理人ID, [D:個人Info].氏名 AS 貸し手, [D:個人Info].住所 AS 貸し手住所, V_農地.借受人ID, V_農地.小作形態, [D:個人Info_1].氏名 AS 借り手, [D:個人Info_1].住所 AS 借り手住所, V_農地.小作開始年月日, V_農地.小作終了年月日, V_地目.名称 AS 地目名, ([田面積]+[畑面積]) AS 面積, [D:個人Info].郵便番号 AS 郵便番号A, V_行政区.名称 AS 集落名A, [D:個人Info_1].郵便番号 AS 郵便番号B, V_行政区_1.名称 AS 集落名B FROM ((((V_農地 LEFT JOIN [D:個人Info] ON V_農地.管理人ID = [D:個人Info].ID) LEFT JOIN [D:個人Info] AS [D:個人Info_1] ON V_農地.借受人ID = [D:個人Info_1].ID) LEFT JOIN V_地目 ON V_農地.登記簿地目 = V_地目.ID) LEFT JOIN V_行政区 ON [D:個人Info].行政区ID = V_行政区.ID) LEFT JOIN V_行政区 AS V_行政区_1 ON [D:個人Info_1].行政区ID = V_行政区_1.ID " & _
        '                                     "WHERE ([小作地適用法]=2) AND ([V_農地].[小作終了年月日]>=#" & Format(dtStart, "yyyy/mm/dd") & "# And (V_農地.小作終了年月日)<=#" & Format(dtEnd, "yyyy/mm/dd") & "#) AND ((V_農地.自小作別)>0) ORDER BY V_農地.小作終了年月日;")
        '    Do Until Rs.EOF
        '        For Each pNode In mvarParent.Nodes
        '            If pNode.Key = "n" & Rs.Value("管理人ID") & "X" & Rs.Value("借受人ID") Then
        '                Exit For
        '            End If
        '        Next

        '        If pNode Is Nothing Then
        '            pNode = mvarParent.TreeNodeAdd("n" & Rs.Value("管理人ID") & "X" & Rs.Value("借受人ID"), "(" & Format(Rs.Value("小作終了年月日"), "GGEE年MM月DD日") & ")" & Rs.Value("貸し手") & "→" & Rs.Value("借り手"), "Unit")
        '            pNode.Tag = Rs.Value("小作開始年月日") & "," & Rs.Value("小作終了年月日") & "," & Rs.Value("貸し手") & "," & Rs.Value("借り手")
        '            pNode.Tag = pNode.Tag & "," & Fnc.GetShortAddress(Rs.NullCast("貸し手住所", ""), "鹿児島県", SysAD.DB(sLRDB).DBProperty("郡部"), , , True)
        '            pNode.Tag = pNode.Tag & "," & Fnc.GetShortAddress(Rs.NullCast("借り手住所", ""), "鹿児島県", SysAD.DB(sLRDB).DBProperty("郡部"), , , True)
        '            pNode.Tag = pNode.Tag & "," & Rs.NullCast("郵便番号A", "") & "," & Rs.NullCast("集落名A", "") & "," & Rs.NullCast("郵便番号B", "") & "," & Rs.NullCast("集落名B", "")
        '            pNode.Tag = pNode.Tag & ";" & Rs.Value("土地所在") & "," & Rs.Value("地目名") & "," & Rs.Value("面積") & "," & Rs.Value("小作形態")
        '        Else
        '            pNode.Tag = pNode.Tag & ";" & Rs.Value("土地所在") & "," & Rs.Value("地目名") & "," & Rs.Value("面積") & "," & Rs.Value("小作形態")
        '        End If

        '        Rs.MoveNext()
        '    Loop
        '    Rs.CloseRs()

        '    MsgBox("全 " & mvarParent.Nodes.Count & " 件(" & dtStart & "～" & dtEnd & ")")
        '    mvarParent.Page = 1
        '    mvarParent.MaxPage = 1

        '    sSQL = "SELECT '世帯員.' & [D:個人Info_1].[ID] AS [Key], [D:個人Info_1].氏名, [D:個人Info_1].フリガナ, [D:個人Info_1].郵便番号 AS 郵便番号, [D:個人Info_1].住所 " & _
        '    "FROM V_農地 LEFT JOIN [D:個人Info] AS [D:個人Info_1] ON V_農地.管理人ID = [D:個人Info_1].ID " & _
        '    "WHERE (((V_農地.小作終了年月日)>=#" & Format(dtStart, "yyyy/mm/dd") & "# And (V_農地.小作終了年月日)<=#" & Format(dtEnd, "yyyy/mm/dd") & "#) AND ((V_農地.自小作別)>0) AND ((V_農地.小作地適用法)=2)) " & _
        '    "GROUP BY '世帯員.' & [D:個人Info_1].[ID], [D:個人Info_1].氏名, [D:個人Info_1].フリガナ, [D:個人Info_1].郵便番号, [D:個人Info_1].住所; " & _
        '    "UNION SELECT '世帯員.' & [D:個人Info_1].[ID] AS [Key], [D:個人Info_1].氏名, [D:個人Info_1].フリガナ, [D:個人Info_1].郵便番号 AS 郵便番号, [D:個人Info_1].住所 " & _
        '    "FROM V_農地 LEFT JOIN [D:個人Info] AS [D:個人Info_1] ON V_農地.借受人ID = [D:個人Info_1].ID " & _
        '    "WHERE (((V_農地.小作終了年月日)>=#" & Format(dtStart, "yyyy/mm/dd") & "# And (V_農地.小作終了年月日)<=#" & Format(dtEnd, "yyyy/mm/dd") & "#) AND ((V_農地.自小作別)>0) AND ((V_農地.小作地適用法)=2)) " & _
        '    "GROUP BY '世帯員.' & [D:個人Info_1].[ID], [D:個人Info_1].氏名, [D:個人Info_1].フリガナ, [D:個人Info_1].郵便番号, [D:個人Info_1].住所;"

        '    mvarPDW.SQLListview.SQLListviewCHead(sSQL, "氏名;氏名;郵便番号;郵便番号;住所;住所", "利用権終了通知通知対象者リスト【宛名ラベル】")
        '    Return Nothing
        Return Nothing
    End Function

    Private Sub PrintSub(ByVal sKey As String, ByVal nPage As Integer)
        'Dim sData As String
        'Dim Ar As Object = Nothing
        'Dim Cm As Object
        'Dim Ax As Object = Nothing

        'Dim n As Single
        'Dim i As Integer = 0
        'Dim nPit As Single, HH As Single, Rows As Integer, Ay As Object       'DX1個目
        'Dim P As Object = Nothing
        'Dim st変更 As Object = Nothing
        'Dim Ln As String
        'Dim LX(20) As String
        'Dim pDate As Date
        'Dim dX(5) As Single, hX(5) As Single
        'Dim nL As Integer

        'If nPage = 0 Then Exit Sub
        'sData = mvarParent.Nodes(sKey).Tag
        'n = 0

        'HH = 23     'ヘッダーの高さ
        'Rows = 15   '何セル分
        'nPit = 8    '表の中の繰り返しの巾
        'Ay = Fnc.GetArray(70, 15, 23, 28, 30)

        'dX(0) = 23 : hX(0) = dX(0) + Ay(0) / 2
        'For n = 0 To UBound(Ay)
        '    dX(n + 1) = dX(n) + Ay(n)
        '    If n < UBound(Ay) Then hX(n + 1) = dX(n + 1) + Ay(n + 1) / 2
        'Next
        '    '----------------------------------
        'With mvarParent
        '    .ClearScreen()
        '    .Offset(0, 0)


        '    If SysAD.Func.s市町村名 = "姶良市" Then
        '        nL = 22
        '    Else
        '        nL = 0
        '    End If

        'If nPage <= 2 Then
        '    '        '    ----------------------------表作成
        '    '.LineBoxDX(dX(0), 141, HH, Rows, nPit, Ay)   '地目→相手方

        '    '.DLine(9 + IIf(SysAD.Func.s市町村名 = "姶良市", nL, 0), 43, 157 + nL, 43, vbBlack)     '殿の下線
        '    '.DLine(45, 72, 155, 72, vbBlack)    'タイトルの下線
        '    '-----------------------------------
        '    If Len(mvarData) Then
        '        Ax = Split(mvarData, ";")
        '    Else
        '        Exit Sub
        '    End If
        '    '            '------------------------------共通部
        '    '.TextOut(Format(dtPrintDay, "GGG  E年  M月   D日"), 175 + IIf(SysAD.Func.s市町村名 = "姶良市", -10, 0), 21, 11, 1)
        '    'If Val(SysAD.DB(sLRDB).DBProperty("終期通知番号指定")) = 1 Then .TextOut("農委第  " & Ax(6) & " 号", 175, 15, 11, 1)
        '    'If Len(Ax(1)) Then
        '    '    .TextOut(Ax(1), 115, 45, 11, 0)
        '    '    .TextOut(Ax(2), 165, 45, 11, 0)
        'End If

        'If Len(Ax(4)) Then
        '    '                .TextOut(Ax(3), 95, 51, 11, 0)
        '    '                .TextOut(Ax(4), 145, 51, 11, 0)
        'Else
        '    '                .TextOut(Ax(3), 145, 51, 11, 0)
        'End If

        '            .TextOut("利用権" & mvarQn(nPage) & "農地の契約期間終了通知書", 100, 67, 13, 2)

        '            .TextOut("記", (dX(0) + dX(5)) / 2, 130, 11, 2)

        '            .TextOut("所　　　在　　　地", hX(0), 150, 11, 2)
        '            If SysAD.Func.s市町村名 = "三股町" Then .TextOut("(三股町)", hX(0), 155, 11, 2)
        '            .TextOut("地 目", hX(1), 150, 11, 2)
        '            .TextOut("備  考", hX(4), 150, 11, 2)
        '            .DrawText(" 面  積\n\n(  ㎡  )", dX(2) + 3, 146, Ay(2), 4, 11)
        '            .DrawText("利用権等\n\nの 種 類", dX(3) + 6, 146, Ay(3), 4, 11)
        '            '------------------------------個人別
        '            Ar = Split(sData, ";")
        '            Cm = Split(Ar(0), ",")
        '            '.TextOut "殿", 54 + nL, 38, 11, 2
        '            '------------------------------市町村別

        '            If SysAD.Func.s市町村名 <> "三股町" Then
        '                .TextOut(Format(Cm(0), "公告日　GGG E年 M月 D日"), dX(0), 132, 11, 0)
        '            Else
        '                .TextOut(Format(Cm(0), "開始日　GGG E年 M月 D日"), dX(0), 132, 11, 0)
        '            End If
        '            .TextOut(Format(Cm(1), "終了日　GGG E年 M月 D日"), dX(0), 136, 11, 0)

        'If UBound(Ax) > 4 Then
        '    Ln = Ax(5)
        '    Ln = Replace(Ln, "%D1%", " " & Format(Cm(0), "GGG E年 M月 D日") & " ")
        '    Ln = Replace(Ln, "%D2%", " " & Format(Cm(1), "GGG E年 M月 D日") & " ")

        '    If SysAD.Func.s市町村名 = "南種子町" Then
        '        pDate = DateAdd("m", -1, CDate(Cm(1)))
        '        dtXX = DateSerial(Year(pDate), Month(pDate), mvar〆日)
        '        'If dt提出 = 0 Then
        '        '    Ln = Replace(Ln, "%D3%", " " & Format(dtXX, "GGG E年 M月 D日") & " ")
        '        'Else
        '        '    Ln = Replace(Ln, "%D3%", " " & Format(dt提出, "GGG E年 M月 D日") & " ")
        '    End If
        'Else
        '    If dt提出 = 0 Then
        '        '    Ln = Replace(Ln, "%D3%", " " & Format(dtXX, "GGG E年 M月 D日") & " ")
        '    Else
        '        '    Ln = Replace(Ln, "%D3%", " " & Format(dt提出, "GGG E年 M月 D日") & " ")
        '    End If
        'End If
        ''                Ln = Replace(Ln, "%Pn%", mvarPn(nPage))
        ''                Ln = Replace(Ln, "%Rn%", mvarRn(nPage))
        ''                Ln = Replace(Ln, "%Qn%", mvarQn(nPage))
        ''                Ln = Ln & mvarQn(3)

        'If nPage = 1 And SysAD.Func.s市町村名 = "三股町" Then
        '    '                    .TextOut("貸し手農家に対する通知書", 175, 15, 11, 1, 0)
        '                    .DLine(124, 19, 176, 19)
        '                    .TextOut(Cm(4), 9, 31, 11)   '窓空き封筒で出すという話になれば表示させる
        '                    .TextOut(Cm(2), 25, 38, 11, 2)

        '                    .TextOut("借受人:" & Cm(3), 125, 136, 11, 2)
        '                    Ln = Replace(Ln, "%An%", " " & Cm(3) & " ")
        '                    If Len(Cm(6)) = 7 Then Cm(6) = Left$(Cm(6), 3) & "-" & Mid$(Cm(6), 4)

        'ElseIf nPage = 1 And (SysAD.Func.s市町村名 <> "姶良市" And SysAD.Func.s市町村名 <> "大崎町") Then

        '                    .TextOut(Cm(4), 9, 31, 11)
        '                    .TextOut(Cm(2) & "       様", 9, 38, 11)
        '                    .TextOut("借受人:" & Cm(3), 190, 136, 11, 1)
        '                    Ln = Replace(Ln, "%An%", " " & Cm(3) & " ")
        '                    If Len(Cm(6)) = 7 Then Cm(6) = Left$(Cm(6), 3) & "-" & Mid$(Cm(6), 4)
        '                    If b封筒 Then .TextOut(Cm(6), 9, 24, 11) '郵便番号
        ''                    If b封筒 Then .TextOut(Cm(7), 9, 46, 11) '集落名
        'ElseIf nPage = 1 And (SysAD.Func.s市町村名 = "大崎町") Then
        ''                    .TextOut(Cm(5), 9 + nL, 31, 11)
        ''                    .TextOut(Cm(3) & " 殿", 9 + nL, 38, 11)
        ''                    .TextOut("貸付人:" & Cm(2), 125, 136, 11, 2)
        ''                    Ln = Replace(Ln, "%An%", " " & Cm(2) & " ")
        ''                    If Len(Cm(6)) = 7 Then Cm(6) = Left$(Cm(6), 3) & "-" & Mid$(Cm(6), 4)
        ''                    If b封筒 Then .TextOut("〒" & Cm(6), 9 + nL, 24, 11) '郵便番号
        ''                    If b封筒 And Len(Cm(7)) > 0 Then .TextOut("(" & Trim(Cm(7)) & ")", 9 + nL, 46, 11) '集落名
        'ElseIf nPage = 1 And (SysAD.Func.s市町村名 = "姶良市") Then
        ''                    .TextOut(Cm(4), 9 + nL, 31, 11)
        '                    .TextOut(Cm(2), 25 + nL, 38, 11, 2)
        '                    .TextOut("借受人:" & Cm(3), 125, 136, 11, 2)
        '                    Ln = Replace(Ln, "%An%", " " & Cm(3) & " ")
        '                    If Len(Cm(6)) = 7 Then Cm(6) = Left$(Cm(6), 3) & "-" & Mid$(Cm(6), 4)
        '                    If b封筒 Then .TextOut("〒" & Cm(6), 9 + nL, 24, 11) '郵便番号
        '                    If b封筒 And Len(Cm(7)) > 0 Then .TextOut("(" & Trim(Cm(7)) & ")", 9 + nL, 46, 11) '集落名
        '                Else
        'If SysAD.Func.s市町村名 <> "姶良市" Then .TextOut(Cm(5), 9, 31, 11)
        '                    .TextOut(Cm(3), 25 + nL, 38, 11, 2)
        '                    .TextOut("貸付人:" & Cm(2), 125, 136, 11, 2)
        '                    Ln = Replace(Ln, "%An%", " " & Cm(2) & " ")

        '                    If Len(Cm(8)) = 7 Then Cm(8) = Left$(Cm(8), 3) & "-" & Mid$(Cm(8), 4)
        '                    If SysAD.Func.s市町村名 = "姶良市" Then
        '                        Cm(5) = Replace(Cm(5), SysAD.Func.s市町村名, "")
        '                        If InStr(Cm(5), "姶良町") > 0 Or InStr(Cm(5), "加治木町") > 0 Or InStr(Cm(5), "蒲生町") > 0 Then
        '                            .TextOut(SysAD.Func.s市町村名 & Cm(5), 9 + nL, 31, 11)
        '                        Else
        '                            .TextOut(Cm(5), 9 + nL, 31, 11)
        'End If

        '                        If b封筒 Then .TextOut("〒" & Cm(8), 9 + nL, 24, 11) '郵便番号
        '                        If b封筒 And Len(Cm(9)) > 0 Then .TextOut("(" & Trim(Cm(9)) & ")", 9 + nL, 46, 11) '集落名
        '                    Else
        '                        If b封筒 Then .TextOut(Cm(8), 9, 24, 11) '郵便番号
        '                        If b封筒 Then .TextOut(Cm(9), 9, 46, 11) '集落名
        '                    End If

        '                    If SysAD.Func.s市町村名 = "三股町" Then
        '                        .TextOut("借り手農家に対する通知書", 175, 15, 11, 1, 0)
        '                        .DLine(124, 19, 176, 19)
        '                    End If

        '                End If
        '            End If

        '            '利用権終期文字列
        '            If InStr(Ln, "＄＄＄") Then
        '                Cm = Split(Ln, "＄＄＄")
        '                If mvarParent.Page = 1 Then
        '                    .DrawText(Cm(0), 29, 82, 160, mvarTPH, 12)
        '                    .TextOut("【貸し手】", 6, 16, 11)
        '                Else
        '                    .DrawText(Cm(1), 29, 82, 160, mvarTPH, 12)
        '                    .TextOut("【借り手】", 6, 16, 11)
        '                End If
        '            Else
        '                .DrawText(Ln, 29, 82, 162, mvarTPH, 12)
        '            End If
        '            '                Select Case mvarData & mvarParent.Page
        '            Select Case SysAD.Func.s市町村名 & mvarParent.Page
        '                Case "三股町1", "三股町2"
        '                    If mvarParent.Page = 1 Then
        '                        st変更 = "利用権を設定した,耕作者,連絡いたしましたので申し添えます。   　賃借契約が終了した後、貴殿が耕作される場合は借手にも連絡しておいて下さい。また耕作されず、,貸,　貸借契約が終了した後、貴殿が耕作される場合は借手にも連絡しておいてください。また耕作されず、継続して貸される場合には必ず、利用権の設定をしてください。"
        '                        P = Split(st変更, ",")
        '                    Else
        '                        st変更 = "利用権の設定を受けた,所有者,連絡いたしましたが、さらに,耕作"
        '                        P = Split(st変更, ",")
        '                    End If
        '                    '                        .hTextOut "　平成　　　年　　　月　　　日付公告の農用地利用集積計画によって" & P(0) & "下記の農地は、平成　　年　　月　　日をもって賃借契約の期間が終了しますので通知致します。", 29, 82, 37, 8
        '                    '                        .hTextOut "　なお、" & P(1) & "にも契約が終了する旨、" & P(2) & "継続して" & P(3) & "される場合は必ず、利用権の設定をして下さい｡", 29, 106, 38, 8
        '                    '                        .hTextOut "※この通知書の写しは最寄の農業委員（農地集積促進員）にも配布してありますので利用権の設定をされる場合はご相談下さい。", 29, 138, 37, 8
        '            End Select
        '            '-----------------------------------
        '            '------------------------------当込み
        '            For n = 1 To UBound(Ar)
        '                '                    If n > 10 Then Exit For
        '                If n > 15 Then Exit For
        '                Cm = Split(Ar(n), ",")
        '                .TextOut(Cm(0), dX(0) + 1, 160 + n * 8, 11, , 2)
        '                .TextOut(Cm(1), hX(1), 160 + n * 8, 11, 2, 2)
        '                .TextOut(FormatNumber(Cm(2), 0), dX(3) - 1, 160 + n * 8, 11, 1, 2)
        '                .TextOut(Choose(Val(Cm(3)) + 1, "", "賃貸借権", "使用貸借権", "その他", ""), hX(3), 160 + n * 8, 11, 2, 2)
        '            Next
        '            '-----------------------------------
        '        Else
        '            Ar = Split(sData, ";")
        '            .TextOut("所　　　在　　　地", hX(0), 35, 11, 2)
        '            If SysAD.Func.s市町村名 = "三股町" Then .TextOut("(三股町)", hX(0), 40, 11, 2)
        '            .TextOut("地 目", hX(1), 35, 11, 2)
        '            .TextOut("備  考", hX(4), 35, 11, 2)
        '            .DrawText(" 面  積\n\n(  ㎡  )", dX(2) + 3, 31, Ay(2), 4, 11)
        '            .DrawText("利用権等\n\nの 種 類", dX(3) + 6, 31, Ay(3), 4, 11)

        '            .LineBoxDX(dX(0), 26, HH, 28, nPit, Ay)   '地目→相手方
        '            If nPage = 3 Then
        '                i = 11
        '            Else
        '                i = 11 + (nPage - 3) * 27
        '            End If

        '            n = 1
        '            For n = i To n + 27
        '                If n > UBound(Ar) Then Exit For
        '                Cm = Split(Ar(n), ",")
        '                .TextOut(Cm(0), dX(0) + 1, 53 + (n - i) * 8, 11, , 2)
        '                .TextOut(Cm(1), hX(1), 53 + (n - i) * 8, 11, 2, 2)
        '                .TextOut(FormatNumber(Cm(2), 0), dX(3) - 1, 53 + (n - i) * 8, 11, 1, 2)
        '                .TextOut(Choose(Val(Cm(3)) + 1, "", "賃貸借権", "使用貸借権", "その他", ""), hX(3), 53 + (n - i) * 8, 11, 2, 2)
        '            Next

        '        End If
        '    End With


    End Sub




End Class

Public Class CPrint利用権終期台帳
    Public Function DataInit(Optional ByVal sParam As String = "") As Boolean
        '        St = "自治会集落名行政区公民館公民会"
        '        s名称変更 = MID(St, Val(SysAD.DB(sLRDB).DBProperty("選挙申請名簿用")) * 3 + 1, 3)

        '        Ar = Split(sParam, ";")
        '        mvarStartDT = Format(Ar(0), "GGGEE年MM月DD日")
        '        mvarEndDT = Format(Ar(1), "GGGEE年MM月DD日")

        '        Rs = SysAD.DB(sLRDB).GetRecordsetEx("SELECT  & _
        '        "FROM ((((((([D:農地info] INNER JOIN [D:個人Info] ON [D:農地info].所有者ID = [D:個人Info].ID) INNER JOIN [D:個人Info] AS [D:個人Info_1] ON [D:農地info].借受人ID = [D:個人Info_1].ID) LEFT JOIN [V_大字] ON [D:農地info].大字ID = [V_大字].ID) LEFT JOIN [V_小字] ON [D:農地info].小字ID = [V_小字].ID) INNER JOIN M_BASICALL ON [D:農地info].小作形態 = M_BASICALL.ID) INNER JOIN V_地目 ON [D:農地info].現況地目 = V_地目.ID) LEFT JOIN V_行政区 AS V_行政区_1 ON [D:個人Info_1].行政区ID = V_行政区_1.ID) LEFT JOIN V_行政区 ON [D:個人Info].行政区ID = V_行政区.ID " & _
        '        "WHERE ((([D:農地info].小作終了年月日)>=CDate('" & mvarStartDT & "') And ([D:農地info].小作終了年月日)<=CDate('" & mvarEndDT & "')) AND (([D:農地info].自小作別)>0) AND ((M_BASICALL.Class)='利用権種類')) " & _
        Return Nothing
    End Function
    Private Sub PrintSub(ByVal nPage As Integer)
        '            .TextOut("利用権終期農地一覧表", 120, Y, 16, 0)
        '            .TextOut("終期の範囲：" & mvarStartDT & "～" & mvarEndDT, 1, Y, 11, 0)

    End Sub




End Class
