'/20160316霧島

Imports System.ComponentModel


Public Class C権利移動借賃MainPage
    Inherits HimTools2012.SystemWindows.CMainPageSK

    Public Sub New()
        MyBase.New(True, False, "権利移動借賃データ出力一覧", "権利移動借賃データ出力一覧")

        mvarListView.Groups.Add("操作", "操作>>").HeaderAlignment = HorizontalAlignment.Left
        mvarListView.Groups.Add("他システム連携", "他システム連携").HeaderAlignment = HorizontalAlignment.Left
        mvarListView.Groups.Add("設定", "設定").HeaderAlignment = HorizontalAlignment.Left

        With Me
            .ListView.ItemAdd("権利移動借賃データ出力(提出用)", "権利移動借賃データ出力(提出用)", "他システム連携", "他システム連携", AddressOf ClickMenu)
            .ListView.ItemAdd("権利移動借賃データ出力(確認用)", "権利移動借賃データ出力(確認用)", "他システム連携", "他システム連携", AddressOf ClickMenu)
            .ListView.ItemAdd("戻る", "戻る", "作業", "操作", AddressOf ClickMenu)
        End With
    End Sub

    Public Sub ClickMenu(ByVal s As Object, ByVal e As EventArgs)
        Select Case CType(s, ListViewItem).Text
            Case "権利移動借賃データ出力(提出用)"
                StartOutPut(EnumOutPut.提出用)
            Case "権利移動借賃データ出力(確認用)"
                StartOutPut(EnumOutPut.確認用)
            Case "戻る"
                If SysAD.MainForm.MainTabCtrl.ExistPage("Main") Then
                    SysAD.MainForm.MainTabCtrl.TabPages.Remove(Me)
                    Me.Dispose()
                End If
        End Select
    End Sub

    Private Sub StartOutPut(ByVal pOutPutType As Integer)
        If MsgBox("農地権利移動・借賃等調査データの作成を開始しますか？", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Dim St調査年度 As String = InputBox("調査年度を入力してください。" & vbCrLf & "※2018年の場合「30」と入力してください。")
            If Len(St調査年度) > 0 Then
                Dim pデータ出力 As New C権利移動借賃データ出力(Val(St調査年度), pOutPutType)
                'pデータ出力.Execute()
                My.Application.DoEvents()

                If MessageBox.Show("終了しました。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly) = DialogResult.OK Then
                End If
            Else
                MsgBox("調査年度の入力をお願いします。")
            End If
        End If
    End Sub
End Class


Public Class C権利移動借賃データ出力
    Inherits HimTools2012.clsAccessor

    Private p調査年度 As Integer = 0
    Private p和暦年度 As Integer = 0
    Private pOutPutType As Integer = 0
    Private CountTBL As DataTable

    Public 都道府県ID As Integer = 0
    Public 支庁郡ID As Integer = 0
    Public 市町村ID As Integer = 0
    Public 中間管理機構ID As Decimal = 0
    Public 中間管理機構名 As String = ""

    Public TBL申請 As DataTable
    Public TBL個人 As DataTable
    Public TBL農地 As DataTable
    Public TBL転用農地 As DataTable
    Public TBL削除農地 As DataTable
    Public TBL申請農地 As DataTable
    Public TBL合計面積 As DataTable

    Public Sub New(ByVal p年度 As Integer, ByVal pOutPut As EnumOutPut)
        p調査年度 = p年度
        p和暦年度 = p年度
        pOutPutType = pOutPut
        Me.Start(True, True)
    End Sub

    Public Overrides Sub Execute()
        '【20161028】南種子町役場　農業委員会　河野様　出力内容に中間管理機構の情報を含んでいるかどうか？
        '
        'Try
        Message = "申請情報読み込み中..."

        Select Case Len(p調査年度.ToString)
            Case 1
                p調査年度 += 2018
            Case 2
                If p調査年度 > 26 Then
                    p調査年度 += 1988
                Else
                    p調査年度 += 2018
                End If
            Case Else
        End Select

        Dim sWhere As String = String.Format("([法令] In (30,31,311,40,50,51,52,60,61,62,180,200,210,602) AND [状態]=2 AND [許可年月日]>=#1/1/{0}# And [許可年月日]<=#12/31/{0}#) OR ([法令] In (31,60,61,62) AND [状態]=2 AND [終期]>=#1/1/{0}# And [終期]<=#12/31/{0}#)", p調査年度)
        TBL申請 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT IIf([始期]<[許可年月日],[許可年月日],[始期]) AS 始期年月日, * FROM D_申請 WHERE {0}", sWhere)
        Dim pView As New DataView(TBL申請, "", "[許可年月日]", DataViewRowState.CurrentRows)
        Me.Value = 0
        '10a賃借料
        Message = "個人情報読み込み中..."
        TBL個人 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:個人Info].ID, [D:個人Info].氏名, [D:個人Info].性別, V_住民区分.名称 AS 住民区分名, [D:個人Info].農業改善計画認定 FROM [D:個人Info] LEFT JOIN V_住民区分 ON [D:個人Info].住民区分 = V_住民区分.ID;")
        TBL個人.PrimaryKey = {TBL個人.Columns("ID")}

        Message = "農地情報読み込み中..."
        TBL農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:農地Info].ID, [D:農地Info].大字ID, V_大字.名称 AS 大字, [D:農地Info].小字ID, V_小字.名称 AS 小字, [D:農地Info].地番, [D:農地Info].登記簿地目, V_地目.名称 AS 登記簿地目名, [D:農地Info].登記簿面積, [D:農地Info].実面積, [D:農地Info].調査土地利用区域地目, [D:農地Info].農振法区分, [D:農地Info].農業振興地域, [D:農地Info].都市計画法, [D:農地Info].都市計画法区分, [D:農地Info].田面積, [D:農地Info].畑面積, [D:農地Info].自小作別, [D:農地Info].借受人ID, [D:農地Info].経由農業生産法人ID, [D:農地Info].小作料, [D:農地Info].小作料単位, [D:農地Info].[10a賃借料] FROM (([D:農地Info] LEFT JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) LEFT JOIN V_小字 ON [D:農地Info].小字ID = V_小字.ID) LEFT JOIN V_地目 ON [D:農地Info].登記簿地目 = V_地目.ID;")
        TBL農地.PrimaryKey = {TBL農地.Columns("ID")}

        TBL転用農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D_転用農地].ID, [D_転用農地].大字ID, V_大字.名称 AS 大字, [D_転用農地].小字ID, V_小字.名称 AS 小字, [D_転用農地].地番, [D_転用農地].登記簿地目, V_地目.名称 AS 登記簿地目名, [D_転用農地].登記簿面積, [D_転用農地].実面積, [D_転用農地].調査土地利用区域地目, [D_転用農地].農振法区分, [D_転用農地].農業振興地域, [D_転用農地].都市計画法, [D_転用農地].都市計画法区分, [D_転用農地].田面積, [D_転用農地].畑面積, [D_転用農地].自小作別, [D_転用農地].借受人ID, [D_転用農地].経由農業生産法人ID, [D_転用農地].小作料, [D_転用農地].小作料単位, [D_転用農地].[10a賃借料] FROM (([D_転用農地] LEFT JOIN V_大字 ON [D_転用農地].大字ID = V_大字.ID) LEFT JOIN V_小字 ON [D_転用農地].小字ID = V_小字.ID) LEFT JOIN V_地目 ON [D_転用農地].登記簿地目 = V_地目.ID;")
        TBL転用農地.PrimaryKey = {TBL転用農地.Columns("ID")}

        TBL削除農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D_削除農地].ID, [D_削除農地].大字ID, V_大字.名称 AS 大字, [D_削除農地].小字ID, V_小字.名称 AS 小字, [D_削除農地].地番, [D_削除農地].登記簿地目, V_地目.名称 AS 登記簿地目名, [D_削除農地].登記簿面積, [D_削除農地].実面積, [D_削除農地].調査土地利用区域地目, [D_削除農地].農振法区分, [D_削除農地].農業振興地域, [D_削除農地].都市計画法, [D_削除農地].都市計画法区分, [D_削除農地].田面積, [D_削除農地].畑面積, [D_削除農地].自小作別, [D_削除農地].借受人ID, [D_削除農地].経由農業生産法人ID, [D_削除農地].小作料, [D_削除農地].小作料単位, [D_削除農地].[10a賃借料] FROM (([D_削除農地] LEFT JOIN V_大字 ON [D_削除農地].大字ID = V_大字.ID) LEFT JOIN V_小字 ON [D_削除農地].小字ID = V_小字.ID) LEFT JOIN V_地目 ON [D_削除農地].登記簿地目 = V_地目.ID;")
        TBL削除農地.PrimaryKey = {TBL削除農地.Columns("ID")}

        CreateCountTable()

        都道府県ID = Val(SysAD.DB(sLRDB).DBProperty("都道府県ID").ToString)
        支庁郡ID = Val(SysAD.DB(sLRDB).DBProperty("支庁郡ID").ToString)
        市町村ID = IIf(Len(Val(SysAD.DB(sLRDB).DBProperty("市町村ID").ToString)) > 3, Left(Val(SysAD.DB(sLRDB).DBProperty("市町村ID").ToString), 3), Val(SysAD.DB(sLRDB).DBProperty("市町村ID").ToString))
        中間管理機構ID = Val(SysAD.DB(sLRDB).DBProperty("中間管理機構ID"))
        Dim pFindRow As DataRow = TBL個人.Rows.Find(中間管理機構ID)
        中間管理機構名 = pFindRow.Item("氏名").ToString

        Set権利設定移転()
        Split中間管理機構 = Enum中間管理機構.設定無し
        Set貸借終了()
        Set農地転用()
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try
    End Sub

    Private MSplit中間管理機構 As Enum中間管理機構 = Enum中間管理機構.設定無し
    Private Split中間管理機構 As Enum中間管理機構 = Enum中間管理機構.設定無し

    Private Sub Set権利設定移転()
        Try
            Dim sCSV As New StringBEx("")
            Dim h権利の設定移転 As String() = {"調査年", "都道府県", "振興局･郡等", "市区町村", "適用法令", "整理番号", "許可・受理・協議・公告年月日", "譲受人(借人）", "譲渡人(貸人）", "権利の種類", "農地法第３条２項５号（下限面積）不許可の例外該当の有無", "農地法第３条２項１号、２号、４号不許可の例外該当の有無", "貸借期間：始期", "貸借期間：終期", "貸借期間", "権利の設定・移転を受ける者（譲受人、借人）個人・法人の別", "権利の設定・移転を受ける者（譲受人、借人）法人の形態別", "権利の設定・移転を受ける者（譲受人、借人）経営改善計画の認定の有無", "権利の設定・移転をする者（譲渡人、貸人）の個人・法人の別", "筆通し番号", "大字", "小字", "地番", "土地利用計画の区域区分・地目", "面積(㎡)", "賃借料情報区分", "借賃等(百円/10ａ)", "地域項目１", "地域項目２"}
            Dim pLineHeader As New StringBEx("様式番号")
            Dim pView As New DataView(TBL申請, String.Format("[法令] In (30,31,311,60,61,62) AND [許可年月日]>=#1/1/{0}# And [許可年月日]<=#12/31/{0}# AND [農地リスト] IS NOT NULL", p調査年度), "[法令],[許可年月日]", DataViewRowState.CurrentRows)   '仮
            Dim RowCount As Integer = 0

            For n As Integer = 0 To UBound(h権利の設定移転) '仮
                pLineHeader.mvarBody.Append("," & h権利の設定移転(n))
            Next
            sCSV.Body.AppendLine(pLineHeader.Body.ToString)

            Dim n整理番号 As New Dictionary(Of Integer, Integer)
            Dim s筆通し番号 As Integer = 1

            If MsgBox("中間管理機構を介する場合、申請は２種類作成していますか？（①Aさん→中間管理機構 ②中間管理機構→Bさん）", vbYesNo) = vbNo Then
                MSplit中間管理機構 = Enum中間管理機構.分割処理
                Split中間管理機構 = Enum中間管理機構.分割処理
            End If

            Me.Maximum = pView.Count
            For Each pRow As DataRowView In pView '仮
                Me.Value += 1
                RowCount += 1
                Message = "権利の設定・移転データファイル出力中(" & RowCount & "/" & pView.Count & ")..."

                Dim pLineRow As New StringBEx("1")  '初期値(様式番号)
                Dim NotAppend As Boolean = False
                With pLineRow.mvarBody
                    If MSplit中間管理機構 = Enum中間管理機構.分割処理 AndAlso Val(pRow.Item("申請者A").ToString) <> 中間管理機構ID AndAlso Val(pRow.Item("申請者B").ToString) <> 中間管理機構ID AndAlso Val(pRow.Item("経由法人ID").ToString) = 中間管理機構ID Then
                        '/***Aさんから中間管理機構へ***/
                        Split中間管理機構 = Enum中間管理機構.Aから中間管理機構
                        申請情報参照(pLineRow, pRow, n整理番号, False, NotAppend, Enum様式番号.権利設定移転, False, True)

                        Dim n As Integer = 0
                        Get筆情報(pRow)
                        For Each pRowView As DataRowView In p農地View
                            If n = 0 Then
                                If NotAppend Then
                                Else
                                    筆情報参照(pLineRow, pRow, pRowView, s筆通し番号, Enum様式番号.権利設定移転)
                                    sCSV.Body.AppendLine(pLineRow.Body.ToString)
                                    n += 1
                                End If
                            Else
                                s筆通し番号 += 1
                                pLineRow = New StringBEx("1")
                                If NotAppend Then
                                Else
                                    申請情報参照(pLineRow, pRow, n整理番号, True, NotAppend, Enum様式番号.権利設定移転, False, True)
                                    筆情報参照(pLineRow, pRow, pRowView, s筆通し番号, Enum様式番号.権利設定移転)
                                    sCSV.Body.AppendLine(pLineRow.Body.ToString)
                                End If
                            End If
                        Next

                        s筆通し番号 = 1

                        '/***中間管理機構からBさんへ***/
                        pLineRow = New StringBEx("1")
                        NotAppend = False

                        Split中間管理機構 = Enum中間管理機構.中間管理機構からB
                        申請情報参照(pLineRow, pRow, n整理番号, False, NotAppend, Enum様式番号.権利設定移転)

                        n = 0
                        Get筆情報(pRow)
                        For Each pRowView As DataRowView In p農地View
                            If n = 0 Then
                                If NotAppend Then
                                Else
                                    筆情報参照(pLineRow, pRow, pRowView, s筆通し番号, Enum様式番号.権利設定移転)
                                    sCSV.Body.AppendLine(pLineRow.Body.ToString)
                                    n += 1
                                End If
                            Else
                                s筆通し番号 += 1
                                pLineRow = New StringBEx("1")
                                If NotAppend Then
                                Else
                                    申請情報参照(pLineRow, pRow, n整理番号, True, NotAppend, Enum様式番号.権利設定移転)
                                    筆情報参照(pLineRow, pRow, pRowView, s筆通し番号, Enum様式番号.権利設定移転)
                                    sCSV.Body.AppendLine(pLineRow.Body.ToString)
                                End If
                            End If
                        Next

                        Split中間管理機構 = Enum中間管理機構.分割処理
                    Else
                        If Val(pRow.Item("申請者A").ToString) <> 中間管理機構ID AndAlso Val(pRow.Item("申請者B").ToString) = 中間管理機構ID AndAlso Val(pRow.Item("経由法人ID").ToString) = 中間管理機構ID Then
                            Split中間管理機構 = Enum中間管理機構.Aから中間管理機構
                        ElseIf Val(pRow.Item("申請者A").ToString) = 中間管理機構ID AndAlso Val(pRow.Item("申請者B").ToString) <> 中間管理機構ID AndAlso Val(pRow.Item("経由法人ID").ToString) = 中間管理機構ID Then
                            Split中間管理機構 = Enum中間管理機構.中間管理機構からB
                        Else
                            Split中間管理機構 = Enum中間管理機構.設定無し
                        End If
                        申請情報参照(pLineRow, pRow, n整理番号, False, NotAppend, Enum様式番号.権利設定移転)

                        Dim n As Integer = 0
                        Get筆情報(pRow)
                        For Each pRowView As DataRowView In p農地View
                            If n = 0 Then
                                If NotAppend Then
                                Else
                                    筆情報参照(pLineRow, pRow, pRowView, s筆通し番号, Enum様式番号.権利設定移転)
                                    sCSV.Body.AppendLine(pLineRow.Body.ToString)
                                    n += 1
                                End If
                            Else
                                s筆通し番号 += 1
                                pLineRow = New StringBEx("1")
                                If NotAppend Then
                                Else
                                    申請情報参照(pLineRow, pRow, n整理番号, True, NotAppend, Enum様式番号.権利設定移転)
                                    筆情報参照(pLineRow, pRow, pRowView, s筆通し番号, Enum様式番号.権利設定移転)
                                    sCSV.Body.AppendLine(pLineRow.Body.ToString)
                                End If
                            End If
                        Next
                    End If
                End With

                s筆通し番号 = 1

                If _Cancel Then
                    Throw New Exception("Cancel")
                    Exit Sub
                End If
            Next

            名前を付けて保存(sCSV, "権利の設定・移転データファイル", True)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Set貸借終了()
        Try
            Dim sCSV As New StringBEx("")
            Dim h貸借の終了 As String() = {"調査年", "都道府県", "振興局･郡等", "市区町村", "適用法令", "整理番号", "許可・受理・協議・公告年月日", "譲受人(借人）", "譲渡人(貸人）", "権利の種類", "返還する者（借人）個人・法人の別", "返還する者（借人）法人の形態別", "返還を受ける者（貸人）の個人・法人の別", "許可・通知・取消しの根拠条項", "基盤強化法による利用権の終了後の農地の状況", "機構法による貸借の終了後の農地の状況", "筆通し番号", "大字", "小字", "地番", "土地利用計画の区域区分・地目", "面積(㎡)", "地域項目１", "地域項目２"}
            Dim pLineHeader As New StringBEx("様式番号")
            Dim pView As New DataView(TBL申請, String.Format("([法令] In (180,200,210) AND [農地リスト] IS NOT NULL)", p調査年度), "[法令],[許可年月日]", DataViewRowState.CurrentRows)
            If pView.Count > 0 Then
                If MsgBox("議案書のない申請情報を強制的に作成しますか？", vbOKCancel) = MsgBoxResult.Ok Then
                    pView = New DataView(TBL申請, String.Format("([法令] In (180,200,210) AND [農地リスト] IS NOT NULL) OR ([法令] In (31,60,61,62) AND [状態]=2 AND [終期]>=#1/1/{0}# And [終期]<=#12/31/{0}# AND [農地リスト] IS NOT NULL)", p調査年度), "[法令],[許可年月日]", DataViewRowState.CurrentRows)
                End If
            Else
                pView = New DataView(TBL申請, String.Format("([法令] In (180,200,210) AND [農地リスト] IS NOT NULL) OR ([法令] In (31,60,61,62) AND [状態]=2 AND [終期]>=#1/1/{0}# And [終期]<=#12/31/{0}# AND [農地リスト] IS NOT NULL)", p調査年度), "[法令],[許可年月日]", DataViewRowState.CurrentRows)
            End If

            For n As Integer = 0 To UBound(h貸借の終了)
                pLineHeader.mvarBody.Append("," & h貸借の終了(n))
            Next
            sCSV.Body.AppendLine(pLineHeader.Body.ToString)

            Dim RowCount As Integer = 0
            Dim n整理番号 As New Dictionary(Of Integer, Integer)
            Dim s筆通し番号 As Integer = 1
            Me.Value = 0
            Me.Maximum = pView.Count
            For Each pRow As DataRowView In pView
                Me.Value += 1
                RowCount += 1
                Message = "貸借の終了データファイル出力中(" & RowCount & "/" & pView.Count & ")..."

                Dim pLineRow As New StringBEx("2")  '初期値(様式番号)
                Dim NotAppend As Boolean = False
                With pLineRow.mvarBody
                    申請情報参照(pLineRow, pRow, n整理番号, False, NotAppend, Enum様式番号.貸借終了, True)

                    Dim n As Integer = 0
                    Get筆情報(pRow)
                    For Each pRowView As DataRowView In p農地View
                        If n = 0 Then
                            If NotAppend Then
                            Else
                                筆情報参照(pLineRow, pRow, pRowView, s筆通し番号, Enum様式番号.貸借終了)
                                sCSV.Body.AppendLine(pLineRow.Body.ToString)
                                n += 1
                            End If
                        Else
                            s筆通し番号 += 1
                            pLineRow = New StringBEx("2")
                            If NotAppend Then
                            Else
                                申請情報参照(pLineRow, pRow, n整理番号, True, NotAppend, Enum様式番号.貸借終了, True)
                                筆情報参照(pLineRow, pRow, pRowView, s筆通し番号, Enum様式番号.貸借終了)
                                sCSV.Body.AppendLine(pLineRow.Body.ToString)
                            End If
                        End If
                    Next

                    If NotAppend = True AndAlso n整理番号(n根拠条例) > 0 Then
                        n整理番号(n根拠条例) -= 1
                    End If
                End With

                s筆通し番号 = 1

                If _Cancel Then
                    Throw New Exception("Cancel")
                    Exit Sub
                End If
            Next

            名前を付けて保存(sCSV, "貸借の終了データファイル")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    Private Sub Set農地転用()
        Try
            Dim sCSV As New StringBEx("")
            Dim h農地等の転用 As String() = {"調査年", "都道府県", "振興局･郡等", "市区町村", "適用法令", "整理番号", "許可・受理・協議・公告年月日", "譲受人(借人）", "譲渡人(貸人）", "権利の種類", "許可・届出・協議・公告と許可除外条項", "土地利用計画の区域区分（細区分）", "転用に伴う農用地区域除外", "転用主体", "用途", "一時転用の該当の有無", "農地の区分", "優良農地の許可判断の根拠", "補助番号", "土地利用計画の区域区分・地目", "面積(㎡)", "地域項目１", "地域項目２"}
            Dim pLineHeader As New StringBEx("様式番号")
            Dim pView As New DataView(TBL申請, String.Format("[法令] In (40,42,50,51,52,602) AND [許可年月日]>=#1/1/{0}# And [許可年月日]<=#12/31/{0}# AND [農地リスト] IS NOT NULL", p調査年度), "[法令],[許可年月日]", DataViewRowState.CurrentRows)
            Dim RowCount As Integer = 0

            For n As Integer = 0 To UBound(h農地等の転用)
                pLineHeader.mvarBody.Append("," & h農地等の転用(n))
            Next
            sCSV.Body.AppendLine(pLineHeader.Body.ToString)

            Dim n整理番号 As New Dictionary(Of Integer, Integer)
            Dim s筆通し番号 As Integer = 1
            Me.Value = 0
            Me.Maximum = pView.Count
            For Each pRow As DataRowView In pView
                'If pRow.Item("ID") = 2915 Then
                '    Stop
                'End If
                'If n整理番号.ContainsKey(28) = 18 Then
                '    Stop
                'End If

                Me.Value += 1
                RowCount += 1
                Message = "農地等の転用データファイル出力中(" & RowCount & "/" & pView.Count & ")..."

                Dim pLineRow As New StringBEx("3")  '初期値(様式番号)
                Dim NotAppend As Boolean = False
                With pLineRow.mvarBody
                    申請情報参照(pLineRow, pRow, n整理番号, False, NotAppend, Enum様式番号.農地転用)

                    Dim n As Integer = 0
                    Get合計面積(pRow)
                    For Each pRowView As DataRowView In p転用View
                        If n = 0 Then
                            If NotAppend Then
                            Else
                                筆情報参照(pLineRow, pRow, pRowView, s筆通し番号, Enum様式番号.農地転用)
                                sCSV.Body.AppendLine(pLineRow.Body.ToString)
                                n += 1
                            End If
                        Else
                            s筆通し番号 += 1
                            pLineRow = New StringBEx("3")
                            If NotAppend Then
                            Else
                                申請情報参照(pLineRow, pRow, n整理番号, True, NotAppend, Enum様式番号.農地転用)
                                筆情報参照(pLineRow, pRow, pRowView, s筆通し番号, Enum様式番号.農地転用)
                                sCSV.Body.AppendLine(pLineRow.Body.ToString)
                            End If
                        End If
                    Next

                    If NotAppend = True AndAlso n整理番号(n根拠条例) > 0 Then
                        n整理番号(n根拠条例) -= 1
                    End If
                End With

                s筆通し番号 = 1

                If _Cancel Then
                    Throw New Exception("Cancel")
                    Exit Sub
                End If
            Next

#If DEBUG Then
            名前を付けて保存(sCSV, "農地等の転用データファイル", , True)
#Else
        名前を付けて保存(sCSV, "農地等の転用データファイル", , True)
#End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private n根拠条例 As Integer = 0
    Private n農地区分 As Integer = 0
    Private Flag権利移動借賃 As Enum権利移動借賃
    Private Flag貸借終了 As Enum貸借終了
    Private Sub 申請情報参照(ByRef pLineRow As StringBEx, ByRef pRow As DataRowView, ByRef n整理番号 As Dictionary(Of Integer, Integer), ByRef SameRequest As Boolean, ByRef NotAppend As Boolean, ByVal p様式番号 As Enum様式番号, Optional ByVal pException As Boolean = False, Optional ByVal bForce As Boolean = False)
        Flag権利移動借賃 = Enum権利移動借賃.設定無
        Flag貸借終了 = Enum貸借終了.設定無
        NotAppend = False

        With pLineRow.mvarBody
            .Append("," & p調査年度) '調査年度　20190731変更
            .Append("," & 都道府県ID)  '都道府県
            .Append("," & 支庁郡ID)    '振興局・郡等
            .Append("," & 市町村ID)    '市区町村

            If (IsDBNull(pRow.Item("調査適用法令")) Or Val(pRow.Item("調査適用法令").ToString) = 0) Or bForce = True Then    '適用法令
                Select Case pRow.Item("法令")
                    Case enum法令.農地法3条所有権, enum法令.農地法3条耕作権
                        If pException = True Then
                            .Append("," & 11) : n根拠条例 = 11
                            Flag貸借終了 = Enum貸借終了.根拠条項
                        Else
                            .Append("," & 1) : n根拠条例 = 1
                        End If
                    Case enum法令.農地法3条の3第1項 : .Append("," & 3) : n根拠条例 = 3
                    Case enum法令.基盤強化法所有権, enum法令.利用権設定, enum法令.利用権移転
                        If pException = True Then
                            .Append("," & 14) : n根拠条例 = 14
                            Flag貸借終了 = Enum貸借終了.基盤農地状況
                        Else
                            If Split中間管理機構 = Enum中間管理機構.中間管理機構からB Then
                                .Append("," & 7) : n根拠条例 = 7
                            Else
                                .Append("," & 6) : n根拠条例 = 6
                            End If
                        End If
                    Case enum法令.農地法4条, enum法令.農地法4条一時転用 : .Append("," & 21) : n根拠条例 = 21
                    Case enum法令.農地法5条貸借, enum法令.農地法5条所有権, enum法令.農地法5条一時転用 : .Append("," & 28) : n根拠条例 = 28
                    Case enum法令.非農地証明願 : .Append("," & 35) : n根拠条例 = 35
                    Case enum法令.農地法18条解約
                        'If Val(pRow.Item("調査権利の種類").ToString) = 24 Then
                        '    .Append("," & 14) : n根拠条例 = 14
                        'Else
                        '    .Append("," & 12)
                        '    n根拠条例 = 12
                        '    If Val(pRow.Item("調査権利の種類").ToString) = 22 Then
                        '        NotAppend = True
                        '    End If
                        'End If
                        .Append("," & 14) : n根拠条例 = 14
                        Flag貸借終了 = Enum貸借終了.根拠条項
                    Case enum法令.農地法20条解約, enum法令.合意解約
                        .Append("," & 14) : n根拠条例 = 14
                        Flag貸借終了 = Enum貸借終了.基盤農地状況
                End Select
            Else
                Select Case Val(pRow.Item("調査適用法令").ToString)
                    Case 12
                        If Val(pRow.Item("調査権利の種類").ToString) = 24 Then
                            .Append("," & 14) : n根拠条例 = 14
                            Flag貸借終了 = Enum貸借終了.基盤農地状況
                        Else
                            .Append("," & Val(pRow.Item("調査適用法令").ToString))
                            n根拠条例 = Val(pRow.Item("調査適用法令").ToString)

                            If Val(pRow.Item("調査権利の種類").ToString) = 22 Then
                                NotAppend = True
                            End If
                            Flag貸借終了 = Enum貸借終了.根拠条項
                        End If
                    Case Else
                        .Append("," & Val(pRow.Item("調査適用法令").ToString))
                        n根拠条例 = Val(pRow.Item("調査適用法令").ToString)

                        Select Case Val(pRow.Item("調査適用法令").ToString)
                            Case 10, 11, 13 : Flag貸借終了 = Enum貸借終了.根拠条項
                            Case 14 : Flag貸借終了 = Enum貸借終了.基盤農地状況
                            Case 15 : Flag貸借終了 = Enum貸借終了.機構農地状況
                        End Select
                End Select
            End If
            If Not n整理番号.ContainsKey(n根拠条例) Then
                n整理番号.Add(n根拠条例, 1)
            Else
                If SameRequest = True Then
                Else
                    n整理番号(n根拠条例) += 1
                End If
            End If

            .Append("," & n整理番号(n根拠条例))    '整理番号   '仮（農地が複数ある場合の処理を追加する）

            Select Case pRow.Item("法令")    '許可・受理・協議・公告年月日(現在は許可日で)
                Case enum法令.農地法3条所有権, enum法令.農地法3条耕作権, enum法令.基盤強化法所有権, enum法令.利用権設定, enum法令.利用権移転
                    If pException = True Then
                        .Append("," & Cnv日付変換(pRow.Item("終期")))
                    Else
                        If Split中間管理機構 = Enum中間管理機構.Aから中間管理機構 Then
                            If IsDBNull(pRow.Item("公告年月日")) OrElse Not IsDate(pRow.Item("公告年月日")) Then
                                .Append("," & Cnv日付変換(pRow.Item("許可年月日")))
                            Else
                                .Append("," & Cnv日付変換(pRow.Item("公告年月日")))
                            End If
                        ElseIf Split中間管理機構 = Enum中間管理機構.中間管理機構からB Then
                            If IsDBNull(pRow.Item("機構配分計画知事公告日")) OrElse Not IsDate(pRow.Item("機構配分計画知事公告日")) Then
                                .Append("," & Cnv日付変換(pRow.Item("許可年月日")))
                            Else
                                .Append("," & Cnv日付変換(pRow.Item("機構配分計画知事公告日")))
                            End If
                        Else
                            .Append("," & Cnv日付変換(pRow.Item("許可年月日")))
                        End If
                    End If
                Case Else
                    .Append("," & Cnv日付変換(pRow.Item("許可年月日")))
            End Select

            Select Case n根拠条例
                Case 20 To 26
                    .Append("," & pRow.Item("氏名A").ToString)    '譲受人(借人)
                    .Append(",")    '譲渡人(貸人) 
                Case Else
                    Select Case Split中間管理機構
                        Case Enum中間管理機構.Aから中間管理機構
                            .Append("," & 中間管理機構名)
                            .Append("," & pRow.Item("氏名A").ToString)
                        Case Enum中間管理機構.中間管理機構からB
                            .Append("," & pRow.Item("氏名B").ToString)
                            .Append("," & 中間管理機構名)
                        Case Else
                            .Append("," & pRow.Item("氏名B").ToString)
                            .Append("," & pRow.Item("氏名A").ToString)
                    End Select
            End Select

            If IsDBNull(pRow.Item("調査権利の種類")) Then    '権利の種類
                Select Case pRow.Item("法令")
                    Case enum法令.農地法3条所有権
                        If pException = True Then
                            .Append("," & 21)
                        Else
                            Select Case Val(pRow.Item("所有権移転の種類").ToString)
                                Case 1 : .Append("," & 1)
                                Case Else : .Append("," & 2)
                            End Select
                        End If
                    Case enum法令.基盤強化法所有権
                        If pException = True Then
                            Select Case Val(pRow.Item("権利種類").ToString)
                                Case 1 : .Append("," & 23)
                                Case Else : .Append("," & 24)
                            End Select
                        Else
                            If Val(pRow.Item("小作料").ToString) > 0 Then : .Append("," & 1)
                            Else : .Append("," & 2)
                            End If
                        End If
                    Case enum法令.農地法3条耕作権
                        If pException = True Then
                            .Append("," & 21)
                        Else
                            Select Case Val(pRow.Item("権利種類").ToString)
                                Case 1
                                    .Append("," & 4)
                                    Flag権利移動借賃 = Enum権利移動借賃.貸借期間
                                Case Else
                                    .Append("," & 7)
                                    Flag権利移動借賃 = Enum権利移動借賃.期間
                            End Select
                        End If
                    Case enum法令.利用権設定, enum法令.利用権移転
                        If pException = True Then
                            Select Case Val(pRow.Item("権利種類").ToString)
                                Case 1 : .Append("," & 23)
                                Case Else : .Append("," & 24)
                            End Select
                        Else
                            Select Case Val(pRow.Item("権利種類").ToString)
                                Case 1
                                    .Append("," & 4)
                                    Flag権利移動借賃 = Enum権利移動借賃.貸借期間
                                Case Else
                                    .Append("," & 7)
                                    Flag権利移動借賃 = Enum権利移動借賃.期間
                            End Select
                        End If
                    Case enum法令.農地法3条の3第1項 : .Append("," & 2)
                    Case enum法令.農地法4条, enum法令.農地法4条一時転用 : .Append("," & 35)
                    Case enum法令.農地法5条貸借, enum法令.農地法5条一時転用
                        Select Case Val(pRow.Item("権利種類").ToString)
                            Case 1 : .Append("," & 33)
                            Case Else : .Append("," & 34)
                        End Select
                    Case enum法令.農地法5条所有権
                        Select Case Val(pRow.Item("所有権移転の種類").ToString)
                            Case 1 : .Append("," & 31)
                            Case Else : .Append("," & 32)
                        End Select
                    Case enum法令.非農地証明願 : .Append("," & 35)
                    Case enum法令.農地法18条解約, enum法令.農地法20条解約, enum法令.合意解約
                        Select Case n根拠条例
                            Case 10, 11, 12 : .Append("," & 21)
                            Case 13, 14 : .Append("," & 23)
                            Case 15 : .Append("," & 26)
                        End Select
                End Select
            Else
                Select Case pRow.Item("法令")
                    Case enum法令.農地法18条解約, enum法令.農地法20条解約, enum法令.合意解約
                        Select Case n根拠条例
                            Case 10, 11, 12
                                Select Case Val(pRow.Item("調査権利の種類").ToString)
                                    Case 21, 22, 23, 26 : .Append("," & Val(pRow.Item("調査権利の種類").ToString))
                                    Case Else : .Append("," & 21)
                                End Select
                            Case 13, 14
                                Select Case Val(pRow.Item("調査権利の種類").ToString)
                                    Case 23, 24, 25 : .Append("," & Val(pRow.Item("調査権利の種類").ToString))
                                    Case Else : .Append("," & 23)
                                End Select
                            Case 15
                                Select Case Val(pRow.Item("調査権利の種類").ToString)
                                    Case 26, 27 : .Append("," & Val(pRow.Item("調査権利の種類").ToString))
                                    Case Else : .Append("," & 26)
                                End Select
                        End Select

                        If Val(pRow.Item("調査権利の種類").ToString) = 24 Then
                            Flag貸借終了 = Enum貸借終了.基盤農地状況
                        Else
                            Select Case pRow.Item("調査権利の種類")
                                Case 4 To 6, 10 : Flag権利移動借賃 = Enum権利移動借賃.貸借期間
                                Case 7 To 9, 11 To 12 : Flag権利移動借賃 = Enum権利移動借賃.期間
                            End Select
                        End If
                    Case Else
                        .Append("," & Val(pRow.Item("調査権利の種類").ToString))

                        Select Case pRow.Item("調査権利の種類")
                            Case 4 To 6, 10 : Flag権利移動借賃 = Enum権利移動借賃.貸借期間
                            Case 7 To 9, 11 To 12 : Flag権利移動借賃 = Enum権利移動借賃.期間
                        End Select
                End Select
            End If

            Select Case p様式番号
                Case Enum様式番号.権利設定移転 : Append権利設定移転(pLineRow, pRow, bForce)
                Case Enum様式番号.貸借終了 : Append貸借終了(pLineRow, pRow)
                Case Enum様式番号.農地転用 : Append農地転用(pLineRow, pRow, SameRequest, p様式番号)
            End Select
        End With
    End Sub
    Private Sub Append権利設定移転(ByRef pLineRow As StringBEx, ByRef pRow As DataRowView, Optional ByVal bForce As Boolean = False)
        Try
            With pLineRow.mvarBody
                Select Case pRow.Item("法令")
                    Case enum法令.農地法3条の3第1項, enum法令.基盤強化法所有権, enum法令.利用権設定, enum法令.利用権移転
                        .Append(",")    '農地法第３条２項５号(下限面積)不許可の例外該当の有無
                        .Append(",")    '農地法第３条２項１号、２号、４号不許可の例外該当の有無
                    Case Else
                        .Append("," & IIf(Val(pRow.Item("調査農法3条2項5号").ToString) > 0, Val(pRow.Item("調査農法3条2項5号").ToString), 1))
                        .Append("," & IIf(Val(pRow.Item("調査農法3条2項124号").ToString) > 0, Val(pRow.Item("調査農法3条2項124号").ToString), 1))
                End Select

                If Flag権利移動借賃 = Enum権利移動借賃.期間 Or Flag権利移動借賃 = Enum権利移動借賃.貸借期間 Then
                    .Append("," & Cnv日付変換(pRow.Item("始期年月日")))    '貸借期間：始期
                    .Append("," & Cnv日付変換(pRow.Item("終期")))    '貸借期間：終期
                    .Append("," & Cnv期間変換(pRow.Item("期間"), pRow))    '貸借期間
                Else
                    .Append(",")
                    .Append(",")
                    .Append(",")
                End If

                Dim 譲受人ID As Decimal = IIf(Split中間管理機構 = Enum中間管理機構.Aから中間管理機構, 中間管理機構ID, Val(pRow.Item("申請者B").ToString))
                If Val(pRow.Item("調査個人法人の別A").ToString) > 0 And bForce = False Then
                    .Append("," & Val(pRow.Item("調査個人法人の別A").ToString))    '6-権利の設定・移転を受ける者（譲受人、借人）個人・法人の別
                    If Val(pRow.Item("調査法人の形態別").ToString) > 0 Then
                        .Append("," & Val(pRow.Item("調査法人の形態別").ToString))
                    Else
                        If Val(pRow.Item("調査個人法人の別A").ToString) = 1 Then
                            .Append(",")
                        Else
                            Dim pFind申請者 As DataRow = TBL個人.Rows.Find(譲受人ID)
                            If Not pFind申請者 Is Nothing AndAlso InStr(pFind申請者.Item("住民区分名").ToString, "法人") > 0 Then
                                If Val(pFind申請者.Item("ID").ToString) = 中間管理機構ID Then
                                    .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 8))    '7-権利の設定・移転を受ける者（譲受人、借人）法人の形態別
                                Else
                                    If InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "(株)") > 0 Or InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "株式会社") > 0 Then
                                        .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 1))
                                    ElseIf InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "(有)") > 0 Or InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "有限会社") > 0 Then
                                        .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 2))
                                    ElseIf InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "農事組合法人") > 0 Then
                                        .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 4))
                                    Else
                                        Select Case Val(pRow.Item("調査個人法人の別A").ToString)
                                            Case 2
                                                .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 1))
                                            Case 4
                                                .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 8))
                                            Case 5
                                                .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 10))
                                        End Select
                                    End If
                                End If
                            ElseIf Not pFind申請者 Is Nothing AndAlso Val(pFind申請者.Item("性別").ToString) = 3 Then
                                If Val(pFind申請者.Item("ID").ToString) = 中間管理機構ID Then
                                    .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 8))    '7-権利の設定・移転を受ける者（譲受人、借人）法人の形態別
                                Else
                                    If InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "(株)") > 0 Or InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "株式会社") > 0 Then
                                        .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 1))
                                    ElseIf InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "(有)") > 0 Or InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "有限会社") > 0 Then
                                        .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 2))
                                    ElseIf InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "農事組合法人") > 0 Then
                                        .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 4))
                                    Else
                                        Select Case Val(pRow.Item("調査個人法人の別A").ToString)
                                            Case 2
                                                .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 1))
                                            Case 4
                                                .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 8))
                                            Case 5
                                                .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 10))
                                        End Select
                                    End If
                                End If
                            Else
                                .Append("," & 10)
                            End If
                        End If

                    End If
                Else
                    Dim pFind申請者 As DataRow = TBL個人.Rows.Find(譲受人ID)
                    If Not pFind申請者 Is Nothing AndAlso InStr(pFind申請者.Item("住民区分名").ToString, "法人") > 0 Then
                        If Val(pFind申請者.Item("ID").ToString) = 中間管理機構ID Then
                            .Append("," & 3)
                            .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 8))
                        Else
                            .Append("," & 2)
                            If InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "(株)") > 0 Or InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "株式会社") > 0 Then
                                .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 1))
                            ElseIf InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "(有)") > 0 Or InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "有限会社") > 0 Then
                                .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 2))
                            ElseIf InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "農事組合法人") > 0 Then
                                .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 4))
                            Else
                                .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 1))
                            End If
                        End If
                    ElseIf Not pFind申請者 Is Nothing AndAlso Val(pFind申請者.Item("性別").ToString) = 3 Then
                        If Val(pFind申請者.Item("ID").ToString) = 中間管理機構ID Then
                            .Append("," & 3)
                            .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 8))
                        Else
                            .Append("," & 2)
                            If InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "(株)") > 0 Or InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "株式会社") > 0 Then
                                .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 1))
                            ElseIf InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "(有)") > 0 Or InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "有限会社") > 0 Then
                                .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 2))
                            ElseIf InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "農事組合法人") > 0 Then
                                .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 4))
                            Else
                                .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 1))
                            End If
                        End If
                    ElseIf Not pFind申請者 Is Nothing AndAlso Val(pFind申請者.Item("ID").ToString) = 中間管理機構ID Then    '20190405追加
                        Select Case Split中間管理機構
                            Case Enum中間管理機構.Aから中間管理機構
                                .Append("," & 3)
                                .Append("," & 8)
                            Case Else
                                .Append("," & 1)
                                .Append(",")
                        End Select
                    Else
                        .Append("," & 1)
                        .Append(",")
                    End If
                End If

                If Val(pRow.Item("調査経営改善計画の有無").ToString) > 0 Then
                    .Append("," & Val(pRow.Item("調査経営改善計画の有無").ToString))    '8-権利の設定・移転を受ける者（譲受人、借人）経営改善計画の認定の有無
                Else
                    Dim pFind申請者 As DataRow = TBL個人.Rows.Find(譲受人ID)
                    If pFind申請者 Is Nothing Then
                        .Append("," & 1)
                    Else
                        Select Case Val(pFind申請者.Item("農業改善計画認定").ToString)
                            Case 1 : .Append("," & 1)
                            Case Else : .Append("," & 2)
                        End Select
                    End If
                End If

                Dim 譲渡人ID As Decimal = IIf(Split中間管理機構 = Enum中間管理機構.中間管理機構からB, 中間管理機構ID, Val(pRow.Item("申請者A").ToString))
                If Val(pRow.Item("調査個人法人の別B").ToString) > 0 And bForce = False Then
                    .Append("," & Val(pRow.Item("調査個人法人の別B").ToString))    '9-権利の設定・移転をする者（譲渡人、貸人）の個人・法人の別
                Else
                    Dim pFind申請者 As DataRow = TBL個人.Rows.Find(譲渡人ID)
                    If Not pFind申請者 Is Nothing AndAlso InStr(pFind申請者.Item("住民区分名").ToString, "法人") > 0 Then
                        If Val(pFind申請者.Item("ID").ToString) = 中間管理機構ID Then : .Append("," & 3)
                        Else : .Append("," & 2)
                        End If
                    ElseIf Not pFind申請者 Is Nothing AndAlso Val(pFind申請者.Item("ID").ToString) = 中間管理機構ID Then
                        Select Case Split中間管理機構
                            Case Enum中間管理機構.中間管理機構からB
                                .Append("," & 3)
                            Case Else
                                .Append("," & 1)
                        End Select
                    Else : .Append("," & 1)
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Append貸借終了(ByRef pLineRow As StringBEx, ByRef pRow As DataRowView)
        With pLineRow.mvarBody
            If Val(pRow.Item("調査個人法人の別A").ToString) > 0 Then
                .Append("," & Val(pRow.Item("調査個人法人の別A").ToString))    '21-返還する者（借人）個人・法人の別
                If Val(pRow.Item("調査個人法人の別A").ToString) = 1 Then
                    .Append(",")
                Else
                    If Val(pRow.Item("調査法人の形態別").ToString) > 0 Then
                        .Append("," & Val(pRow.Item("調査法人の形態別").ToString))
                    Else
                        Dim pFind申請者 As DataRow = TBL個人.Rows.Find(Val(pRow.Item("申請者B").ToString))
                        If Not pFind申請者 Is Nothing AndAlso InStr(pFind申請者.Item("住民区分名").ToString, "法人") > 0 Then
                            If Val(pFind申請者.Item("ID").ToString) = 中間管理機構ID Then
                                .Append("," & 3)
                                .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 8))    '22-返還する者（借人）法人の形態別
                            Else
                                .Append("," & 2)
                                If InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "(株)") > 0 Or InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "株式会社") > 0 Then
                                    .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 1))
                                ElseIf InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "(有)") > 0 Or InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "有限会社") > 0 Then
                                    .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 2))
                                ElseIf InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "農事組合法人") > 0 Then
                                    .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 4))
                                Else
                                    .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 1))
                                End If
                            End If
                        Else
                            .Append(",")
                        End If
                    End If
                End If
            Else
                Dim pFind申請者 As DataRow = TBL個人.Rows.Find(Val(pRow.Item("申請者B").ToString))
                If Not pFind申請者 Is Nothing AndAlso InStr(pFind申請者.Item("住民区分名").ToString, "法人") > 0 Then
                    If Val(pFind申請者.Item("ID").ToString) = 中間管理機構ID Then
                        .Append("," & 3)
                        .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 8))    '22-返還する者（借人）法人の形態別
                    Else
                        .Append("," & 2)
                        If InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "(株)") > 0 Or InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "株式会社") > 0 Then
                            .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 1))
                        ElseIf InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "(有)") > 0 Or InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "有限会社") > 0 Then
                            .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 2))
                        ElseIf InStr(StrConv(pRow.Item("氏名B").ToString, vbNarrow), "農事組合法人") > 0 Then
                            .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 4))
                        Else
                            .Append("," & IIf(Val(pRow.Item("調査法人の形態別").ToString) > 0, Val(pRow.Item("調査法人の形態別").ToString), 1))
                        End If
                    End If
                Else
                    .Append("," & 1)
                    .Append(",")
                End If
            End If

            If Val(pRow.Item("調査個人法人の別B").ToString) > 0 Then
                .Append("," & Val(pRow.Item("調査個人法人の別B").ToString))    '23-返還を受ける者（貸人）の個人・法人の別
            Else
                Dim pFind申請者 As DataRow = TBL個人.Rows.Find(Val(pRow.Item("申請者A").ToString))
                If Not pFind申請者 Is Nothing AndAlso InStr(pFind申請者.Item("住民区分名").ToString, "法人") > 0 Then
                    If Val(pFind申請者.Item("ID").ToString) = 中間管理機構ID Then : .Append("," & 3)
                    Else : .Append("," & 2)
                    End If
                Else : .Append("," & 1)
                End If
            End If

            Select Case Flag貸借終了
                Case Enum貸借終了.根拠条項    '24-許可・通知・取消しの根拠条項
                    If Val(pRow.Item("調査許可等の根拠条項").ToString) > 0 Then
                        .Append("," & Val(pRow.Item("調査許可等の根拠条項").ToString))
                    Else
                        Select Case n根拠条例
                            Case 10 : .Append("," & 1)
                            Case 11 : .Append("," & 15)
                            Case 12 : .Append("," & 22)
                            Case 13 : .Append("," & 3)
                        End Select
                    End If
                    .Append(",")
                    .Append(",")
                Case Enum貸借終了.基盤農地状況    '25-基盤強化法による利用権の終了後の農地の状況
                    .Append(",")
                    .Append("," & IIf(Val(pRow.Item("調査基盤法満了農地状況").ToString) > 0, Val(pRow.Item("調査基盤法満了農地状況").ToString), Get小作情報(pRow)))
                    .Append(",")
                Case Enum貸借終了.機構農地状況    '26-機構法による貸借の終了後の農地の状況
                    .Append(",")
                    .Append(",")
                    .Append("," & IIf(Val(pRow.Item("調査中間管理事業法満了農地状況").ToString) > 0, Val(pRow.Item("調査中間管理事業法満了農地状況").ToString), Get小作情報(pRow)))
            End Select
        End With
    End Sub
    Private Sub Append農地転用(ByRef pLineRow As StringBEx, ByRef pRow As DataRowView, ByRef SameRequest As Boolean, ByVal p様式番号 As Enum様式番号)
        With pLineRow.mvarBody
            If Not SameRequest Then
                .Append("," & IIf(Val(pRow.Item("調査許可等の除外条項").ToString) > 0, Val(pRow.Item("調査許可等の除外条項").ToString), IIf(n根拠条例 = 35, 11, 1)))    '許可・届出・協議・公告と許可除外条項

                If p様式番号 = 3 Then
                    If n根拠条例 = 35 Then : .Append("," & IIf(Val(pRow.Item("調査土地利用計画区域区分").ToString) > 0, Val(pRow.Item("調査土地利用計画区域区分").ToString), 1))    '土地利用計画の区域区分（細区分）
                    Else : .Append("," & IIf(Val(pRow.Item("調査土地利用計画区域区分").ToString) > 0, Val(pRow.Item("調査土地利用計画区域区分").ToString), 2))
                    End If
                Else : .Append("," & IIf(Val(pRow.Item("調査土地利用計画区域区分").ToString) > 0, Val(pRow.Item("調査土地利用計画区域区分").ToString), 1))
                End If
                .Append("," & IIf(Val(pRow.Item("調査転用に伴う農用地区域除外").ToString), Val(pRow.Item("調査転用に伴う農用地区域除外").ToString), 2))    '転用に伴う農用地区域除外
                .Append("," & IIf(Val(pRow.Item("調査転用主体").ToString) > 0, Val(pRow.Item("調査転用主体").ToString), 9))    '転用主体
                .Append("," & IIf(Val(pRow.Item("調査転用用途").ToString) > 0, IIf(Val(pRow.Item("調査転用用途").ToString) > 70, 45, Val(pRow.Item("調査転用用途").ToString)), 45))    '用途
                .Append("," & IIf(Val(pRow.Item("調査一時転用該当有無").ToString) > 0, Val(pRow.Item("調査一時転用該当有無").ToString), IIf(pRow.Item("法令") = enum法令.農地法5条一時転用, 1, IIf(pRow.Item("法令") = enum法令.農地法4条一時転用, 1, 2))))   '一時転用の該当の有無
                Select Case n根拠条例   '農地の区分
                    Case 23, 30, 34, 35 : .Append(",")
                    Case Else
                        If Val(pRow.Item("調査転用農地区分").ToString) > 0 Then
                            .Append("," & Val(pRow.Item("調査転用農地区分").ToString))
                            n農地区分 = Val(pRow.Item("調査転用農地区分").ToString)
                        Else
                            If Val(pRow.Item("農振区分").ToString) = 2 Then
                                .Append("," & 1)
                                n農地区分 = 1
                            Else
                                .Append("," & 51)
                                n農地区分 = 51
                            End If
                        End If
                End Select
                Select Case n農地区分    '優良農地の許可判断の根拠
                    Case 1 To 22
                        .Append("," & IIf(Val(pRow.Item("調査優良農地許可判断根拠").ToString) > 0, Val(pRow.Item("調査優良農地許可判断根拠").ToString), 11))
                    Case Else : .Append(",")
                End Select
            Else
                .Append(",")
                .Append(",")
                .Append(",")
                .Append(",")
                .Append(",")
                .Append(",")
                .Append(",")
                .Append(",")
            End If
        End With
    End Sub

    Private Sub 筆情報参照(ByRef pLineRow As StringBEx, ByRef pRow As DataRowView, ByRef pRowView As DataRowView, ByVal Int筆通し番号 As Integer, ByVal p様式番号 As Enum様式番号)
        With pLineRow.mvarBody
            Select Case p様式番号
                Case Enum様式番号.農地転用
                    Set通し番号(pLineRow, pRowView, Int筆通し番号)
                Case Else : .Append("," & Int筆通し番号)    '筆通し番号
            End Select

            Select Case p様式番号
                Case Enum様式番号.権利設定移転, Enum様式番号.貸借終了
                    .Append("," & pRowView.Item("大字").ToString)    '大字
                    .Append("," & pRowView.Item("小字").ToString)    '小字
                    Select Case pOutPutType
                        Case EnumOutPut.提出用
                            .Append("," & pRowView.Item("地番").ToString)    '地番
                        Case EnumOutPut.確認用
                            .Append(",=T(""" & pRowView.Item("地番").ToString & """)")    '地番
                    End Select
                    .Append("," & Cnv土地利用区域地目(pRowView))    '土地利用計画の区域区分・地目
                    .Append("," & pRowView.Item("面積"))    '面積(㎡)

                    'If pRowView.Item("地番").ToString = "451-1" Then
                    '    Stop
                    'End If

                    Select Case p様式番号
                        Case Enum様式番号.権利設定移転
                            If Flag権利移動借賃 = Enum権利移動借賃.貸借 Or Flag権利移動借賃 = Enum権利移動借賃.貸借期間 Then
                                If SysAD.市町村.市町村名 = "宗像市" Then
                                    Select Case Cnv土地利用区域地目(pRowView)
                                        Case 1 : .Append("," & 1)    '賃借料情報区分
                                        Case 2 : .Append("," & 2)
                                        Case 11 : .Append("," & 3)
                                        Case 12 : .Append("," & 4)
                                        Case Else : .Append("," & 1)
                                    End Select
                                Else
                                    .Append("," & 1)    '賃借料情報区分
                                End If
                                If InStr(pRowView.Item("小作料単位").ToString, "円") > 0 Then
                                    .Append("," & Val(pRowView.Item("小作料").ToString))    '借賃等(百円/10ａ)
                                ElseIf SysAD.市町村.市町村名 = "宗像市" AndAlso InStr(pRowView.Item("小作料単位").ToString, "kg") > 0 Then
                                    .Append("," & Val(pRowView.Item("小作料").ToString))    '借賃等(百円/10ａ)
                                Else
                                    .Append("," & 10)    '借賃等(百円/10ａ)
                                End If
                            Else
                                .Append(",")
                                .Append(",")
                            End If
                        Case Enum様式番号.貸借終了
                    End Select
                Case Enum様式番号.農地転用
                    If Val(pRow.Item("調査土地利用計画区域区分").ToString) = 1 Then
                        Select Case Cnv土地利用区域地目(pRowView)
                            Case 1, 6, 11 : .Append("," & 6)
                            Case 2, 7, 12 : .Append("," & 7)
                        End Select
                    Else
                        .Append("," & Cnv土地利用区域地目(pRowView))
                    End If

                    .Append("," & pRowView.Item("面積"))
            End Select

            .Append(",")    '地域項目１
            .Append(",")    '地域項目２
        End With
    End Sub

    Public p農地View As DataView
    Public p転用View As DataView

    Private Sub Get筆情報(ByRef pRow As DataRowView)
        Dim Ar筆リスト As Object = Split(pRow.Item("農地リスト").ToString, ";")
        Dim pRowFind As DataRow = Nothing

        TBL申請農地.Rows.Clear()

        For n As Integer = 0 To UBound(Ar筆リスト)
            Dim Ar筆情報 As Object = Split(Ar筆リスト(n), ".")
            If InStr(Ar筆情報(0), "転用農地") > 0 Then
                pRowFind = TBL転用農地.Rows.Find(Ar筆情報(1))
                If pRowFind Is Nothing Then
                    pRowFind = TBL農地.Rows.Find(Ar筆情報(1))
                    If pRowFind Is Nothing Then
                        pRowFind = TBL削除農地.Rows.Find(Ar筆情報(1))
                    End If
                End If
            ElseIf InStr(Ar筆情報(0), "農地") > 0 Then
                pRowFind = TBL農地.Rows.Find(Ar筆情報(1))
                If pRowFind Is Nothing Then
                    pRowFind = TBL転用農地.Rows.Find(Ar筆情報(1))
                    If pRowFind Is Nothing Then
                        pRowFind = TBL削除農地.Rows.Find(Ar筆情報(1))
                    End If
                End If
            End If

            If Not pRowFind Is Nothing Then
                Dim pAddRow As DataRow = TBL申請農地.NewRow
                pAddRow.Item("ID") = pRowFind.Item("ID")
                pAddRow.Item("大字") = pRowFind.Item("大字").ToString
                pAddRow.Item("小字") = pRowFind.Item("小字").ToString
                pAddRow.Item("地番") = pRowFind.Item("地番").ToString
                'If pRowFind.Item("地番") = "451-1" Then
                '    Stop
                'End If
                pAddRow.Item("面積") = Math.Floor(IIf(Val(pRowFind.Item("実面積").ToString) = 0, Val(pRowFind.Item("登記簿面積").ToString), Val(pRowFind.Item("実面積").ToString)) + 0.5)

                pAddRow.Item("田面積") = Val(pRowFind.Item("田面積").ToString)
                pAddRow.Item("畑面積") = Val(pRowFind.Item("畑面積").ToString)
                pAddRow.Item("調査土地利用区域地目") = Val(pRowFind.Item("調査土地利用区域地目").ToString)
                pAddRow.Item("農振法区分") = Val(pRowFind.Item("農振法区分").ToString)
                pAddRow.Item("農業振興地域") = Val(pRowFind.Item("農業振興地域").ToString)
                pAddRow.Item("都市計画法") = Val(pRowFind.Item("都市計画法").ToString)
                pAddRow.Item("都市計画法区分") = Val(pRowFind.Item("都市計画法区分").ToString)
                pAddRow.Item("自小作別") = Val(pRowFind.Item("自小作別").ToString)
                pAddRow.Item("小作料単位") = pRowFind.Item("小作料単位").ToString
                If Val(pRowFind.Item("10a賃借料").ToString) > 0 Then
                    pAddRow.Item("小作料") = Math.Floor((pRowFind.Item("10a賃借料") / 100) + 0.5)
                Else
                    If InStr(pRowFind.Item("小作料単位").ToString, "円") > 0 Then
                        pAddRow.Item("小作料") = Math.Floor((Val(pRowFind.Item("小作料").ToString) / 100) + 0.5)
                    ElseIf SysAD.市町村.市町村名 = "宗像市" AndAlso InStr(pRowFind.Item("小作料単位").ToString, "kg") > 0 Then
                        pAddRow.Item("小作料") = Math.Floor((Val(pRowFind.Item("小作料").ToString * 200) / 100) + 0.5)
                    Else
                        pAddRow.Item("小作料") = 0
                    End If

                    'Select Case pRowFind.Item("小作料単位").ToString
                    '    Case "円/10a", "円/反"
                    '        pAddRow.Item("小作料") = Math.Floor((Val(pRowFind.Item("小作料").ToString) / 100) + 0.5)
                    '    Case Else
                    '        pAddRow.Item("小作料") = 0
                    'End Select
                End If

                TBL申請農地.Rows.Add(pAddRow)
            End If
        Next

        p農地View = New DataView(TBL申請農地, "", "[面積] DESC", DataViewRowState.CurrentRows)
    End Sub
    Private Function Get小作情報(ByRef pRow As DataRowView) As Integer
        Dim Ar筆リスト As Object = Split(pRow.Item("農地リスト").ToString, ";")
        Dim pRowFind As DataRow = Nothing

        For n As Integer = 0 To UBound(Ar筆リスト)
            Dim Ar筆情報 As Object = Split(Ar筆リスト(n), ".")
            If InStr(Ar筆情報(0), "転用農地") > 0 Then
                pRowFind = TBL転用農地.Rows.Find(Ar筆情報(1))
                If pRowFind Is Nothing Then
                    pRowFind = TBL農地.Rows.Find(Ar筆情報(1))
                    If pRowFind Is Nothing Then
                        pRowFind = TBL削除農地.Rows.Find(Ar筆情報(1))
                    End If
                End If
            ElseIf InStr(Ar筆情報(0), "農地") > 0 Then
                pRowFind = TBL農地.Rows.Find(Ar筆情報(1))
                If pRowFind Is Nothing Then
                    pRowFind = TBL転用農地.Rows.Find(Ar筆情報(1))
                    If pRowFind Is Nothing Then
                        pRowFind = TBL削除農地.Rows.Find(Ar筆情報(1))
                    End If
                End If
            End If

            If Not pRowFind Is Nothing Then
                If Val(pRowFind.Item("自小作別").ToString) = 0 Then
                    Return 13
                Else
                    If Val(pRowFind.Item("借受人ID").ToString) = 中間管理機構ID Then
                        Return 12
                    ElseIf Val(pRowFind.Item("経由農業生産法人ID").ToString) = 中間管理機構ID Then
                        Return 11
                    ElseIf Val(pRow.Item("申請者B").ToString) = Val(pRowFind.Item("借受人ID").ToString) Then
                        Return 1
                    Else
                        Return 2
                    End If
                End If
            Else
                Return 21
            End If
        Next

        Return 21
    End Function
    Private Sub Get合計面積(ByRef pRow As DataRowView)
        Dim Ar筆リスト As Object = Split(pRow.Item("農地リスト").ToString, ";")
        Dim pRowFind As DataRow = Nothing
        Dim pRowArea As DataRow = Nothing

        TBL合計面積.Rows.Clear()

        For n As Integer = 0 To UBound(Ar筆リスト)
            Dim Ar筆情報 As Object = Split(Ar筆リスト(n), ".")
            If InStr(Ar筆情報(0), "転用農地") > 0 Then
                pRowFind = TBL転用農地.Rows.Find(Ar筆情報(1))
                If pRowFind Is Nothing Then
                    pRowFind = TBL農地.Rows.Find(Ar筆情報(1))
                    If pRowFind Is Nothing Then
                        pRowFind = TBL削除農地.Rows.Find(Ar筆情報(1))
                    End If
                End If
            ElseIf InStr(Ar筆情報(0), "農地") > 0 Then
                pRowFind = TBL農地.Rows.Find(Ar筆情報(1))
                If pRowFind Is Nothing Then
                    pRowFind = TBL転用農地.Rows.Find(Ar筆情報(1))
                    If pRowFind Is Nothing Then
                        pRowFind = TBL削除農地.Rows.Find(Ar筆情報(1))
                    End If
                End If
            End If

            If Not pRowFind Is Nothing Then
                Dim pAddRow As DataRow = TBL合計面積.NewRow
                pRowArea = TBL合計面積.Rows.Find(IIf(Val(pRowFind.Item("調査土地利用区域地目").ToString) > 0, Val(pRowFind.Item("調査土地利用区域地目").ToString), Cnv土地利用区域転用地目(pRowFind)))

                If Not pRowArea Is Nothing Then
                    pRowArea.Item("面積") = Val(pRowArea.Item("面積").ToString) + Math.Floor(IIf(Val(pRowFind.Item("実面積").ToString) = 0, Val(pRowFind.Item("登記簿面積").ToString), Val(pRowFind.Item("実面積").ToString)) + 0.5)
                Else
                    pAddRow.Item("調査土地利用区域地目") = IIf(Val(pRowFind.Item("調査土地利用区域地目").ToString) > 0, Val(pRowFind.Item("調査土地利用区域地目").ToString), Cnv土地利用区域転用地目(pRowFind))
                    pAddRow.Item("面積") = Math.Floor(IIf(Val(pRowFind.Item("実面積").ToString) = 0, Val(pRowFind.Item("登記簿面積").ToString), Val(pRowFind.Item("実面積").ToString)) + 0.5)
                    TBL合計面積.Rows.Add(pAddRow)
                End If
            End If
        Next

        p転用View = New DataView(TBL合計面積, "", "[面積] DESC", DataViewRowState.CurrentRows)
    End Sub

    Private sPath As String = ""
    Private Sub 名前を付けて保存(ByVal sCSV As StringBEx, ByVal SaveFileName As String, Optional ByVal OpenDialog As Boolean = False, Optional ByVal OpenFolder As Boolean = False)
        '/***名前を付けて保存***/
        If OpenDialog = True Then
            With New SaveFileDialog
                .FileName = String.Format("{0}.csv", SaveFileName)
                .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
                .Filter = "CSVファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*"

                If .ShowDialog = DialogResult.OK Then
                    sPath = .FileName
                End If
            End With
        End If

        Dim ArSavePath As Object = Split(sPath, "\")
        For n As Integer = 0 To UBound(ArSavePath)
            If n = 0 Then : sPath = ArSavePath(0)
            ElseIf n = UBound(ArSavePath) Then : sPath = sPath & "\" & String.Format("{0}.csv", SaveFileName)
            Else : sPath = sPath & "\" & ArSavePath(n)
            End If
        Next

        Dim CSVText As System.IO.StreamWriter
        Select Case pOutPutType
            Case EnumOutPut.提出用
                CSVText = New System.IO.StreamWriter(sPath, False, System.Text.Encoding.GetEncoding(932))
            Case Else
                CSVText = New System.IO.StreamWriter(sPath, False, System.Text.Encoding.UTF8)
        End Select

        CSVText.Write(sCSV.Body.ToString)
        CSVText.Dispose()

        If OpenFolder = True Then
            MsgBox("CSVの出力が完了しました。")
            SysAD.ShowFolder(System.IO.Directory.GetParent(sPath).ToString)
        End If
    End Sub

    Private Sub CreateCountTable()
        CountTBL = New DataTable
        CountTBL.Columns.Add("ID", GetType(Integer))
        CountTBL.Columns.Add("CountValue", GetType(Integer))

        CountTBL.PrimaryKey = New DataColumn() {CountTBL.Columns("ID")}

        TBL申請農地 = New DataTable
        With TBL申請農地
            .Columns.Add("ID", GetType(Integer))
            .Columns.Add("大字", GetType(String))
            .Columns.Add("小字", GetType(String))
            .Columns.Add("地番", GetType(String))
            .Columns.Add("面積", GetType(Decimal))
            .Columns.Add("田面積", GetType(Decimal))
            .Columns.Add("畑面積", GetType(Decimal))
            .Columns.Add("調査土地利用区域地目", GetType(Integer))
            .Columns.Add("農振法区分", GetType(Integer))
            .Columns.Add("農業振興地域", GetType(Integer))
            .Columns.Add("都市計画法", GetType(Integer))
            .Columns.Add("都市計画法区分", GetType(Integer))
            .Columns.Add("自小作別", GetType(Integer))
            .Columns.Add("小作料", GetType(Integer))
            .Columns.Add("小作料単位", GetType(String))
        End With

        TBL合計面積 = New DataTable
        With TBL合計面積
            .Columns.Add("調査土地利用区域地目", GetType(Integer))
            .Columns.Add("面積", GetType(Decimal))

            .PrimaryKey = New DataColumn() { .Columns("調査土地利用区域地目")}
        End With
    End Sub

    Private Function Cnv日付変換(ByVal pDate As Object) As String
        Dim StDate As String = ""
        If IsDate(pDate) Then
            StDate = Format(pDate, "yyyy.M.d")

            '//20190726 農地の権利移動借賃等調査回収により「yy.M.d」⇒「yyyy.M.d」へ
            'StDate = 和暦Format(pDate)
            'StDate = Replace(StDate, "平成", "")
            'StDate = Replace(StDate, "令和", "") '//追加
            'StDate = Replace(StDate, "年", ".")
            'StDate = Replace(StDate, "月", ".")
            'StDate = Replace(StDate, "日", "")
        End If

        Return StDate
    End Function
    Private Function Cnv期間変換(ByVal pPeriod As Object, ByRef pRow As DataRowView) As Integer
        If Not IsDBNull(pPeriod) Then
            Select Case Val(pPeriod)
                Case Is < 1 : Return 1 '// 1年未満
                Case 1 To 2 : Return 2 '// 1年以上～3年未満
                Case 3 To 5 : Return 3 '// 3年以上～6年未満
                Case 6 To 9 : Return 4 '// 6年以上～10年未満
                Case 10 To 19 : Return 5 '// 10年以上～20年未満
                Case 20 To 29 : Return 6 '// 20年以上～30年未満
                Case 30 To 39 : Return 7 '// 30年以上～40年未満
                Case 40 To 49 : Return 8 '// 40年以上～50年未満
                Case 50 : Return 9
                Case Else : Return 0
            End Select
        Else
            If Not IsDBNull(pRow.Item("始期年月日")) AndAlso Not IsDBNull(pRow.Item("終期")) Then
                Dim p期間 As Integer = DateDiff(DateInterval.Year, pRow.Item("始期年月日"), pRow.Item("終期"))
                Return p期間

            Else : Return 0
            End If
        End If
    End Function

    Private Sub Set通し番号(ByRef pLineRow As StringBEx, ByRef pRowView As DataRowView, ByVal Int筆通し番号 As Integer)
        With pLineRow.mvarBody
            If Not IsDBNull(pRowView.Item("調査土地利用区域地目")) Then
                Create通し番号(pLineRow, Val(pRowView.Item("調査土地利用区域地目").ToString), Int筆通し番号)
            Else
                If Val(pRowView.Item("農振法区分").ToString) = 1 Or Val(pRowView.Item("農業振興地域").ToString) = 1 Then
                    If Val(pRowView.Item("田面積").ToString) > 0 Then : Create通し番号(pLineRow, 1, Int筆通し番号)
                    ElseIf Val(pRowView.Item("畑面積").ToString) > 0 Then : Create通し番号(pLineRow, 2, Int筆通し番号)
                    Else : Create通し番号(pLineRow, 3, Int筆通し番号)
                    End If
                ElseIf Val(pRowView.Item("都市計画法区分").ToString) = 1 Or Val(pRowView.Item("都市計画法").ToString) = 4 Then
                    If Val(pRowView.Item("田面積").ToString) > 0 Then : Create通し番号(pLineRow, 6, Int筆通し番号)
                    ElseIf Val(pRowView.Item("畑面積").ToString) > 0 Then : Create通し番号(pLineRow, 7, Int筆通し番号)
                    Else : Create通し番号(pLineRow, 8, Int筆通し番号)
                    End If
                Else
                    If Val(pRowView.Item("田面積").ToString) > 0 Then : Create通し番号(pLineRow, 11, Int筆通し番号)
                    ElseIf Val(pRowView.Item("畑面積").ToString) > 0 Then : Create通し番号(pLineRow, 12, Int筆通し番号)
                    Else : Create通し番号(pLineRow, 13, Int筆通し番号)
                    End If
                End If
            End If
        End With
    End Sub
    Private Sub Create通し番号(ByRef pLineRow As StringBEx, ByVal pKey As Integer, ByVal Int筆通し番号 As Integer)
        Dim pAddRow As DataRow = CountTBL.Rows.Find(pKey)
        If pAddRow Is Nothing Then
            pAddRow = CountTBL.NewRow()

            pAddRow.Item("ID") = pKey
            pAddRow.Item("CountValue") = Int筆通し番号

            CountTBL.Rows.Add(pAddRow)
        Else
            If Int筆通し番号 = 1 Then : pAddRow.Item("CountValue") = 1
            Else : pAddRow.Item("CountValue") += 1
            End If
        End If

        pLineRow.mvarBody.Append("," & pAddRow.Item("CountValue"))    '補助番号


    End Sub

    '農地の区分が1（農用地区域内～）の場合地目！
    Private Function Cnv土地利用区域地目(ByRef pRowFind As DataRowView) As Integer
        If Not IsDBNull(pRowFind.Item("調査土地利用区域地目")) And Val(pRowFind.Item("調査土地利用区域地目").ToString) Then
            Return Val(pRowFind.Item("調査土地利用区域地目").ToString)
        Else
            If Val(pRowFind.Item("農振法区分").ToString) = 1 Or Val(pRowFind.Item("農業振興地域").ToString) = 1 Then
                If Val(pRowFind.Item("田面積").ToString) > 0 Then : Return 1
                ElseIf Val(pRowFind.Item("畑面積").ToString) > 0 Then : Return 2
                Else : Return 3
                End If
            ElseIf Val(pRowFind.Item("都市計画法区分").ToString) = 1 Or Val(pRowFind.Item("都市計画法").ToString) = 4 Then
                If Val(pRowFind.Item("田面積").ToString) > 0 Then : Return 6
                ElseIf Val(pRowFind.Item("畑面積").ToString) > 0 Then : Return 7
                Else : Return 8
                End If
            Else
                If Val(pRowFind.Item("田面積").ToString) > 0 Then : Return 11
                ElseIf Val(pRowFind.Item("畑面積").ToString) > 0 Then : Return 12
                Else : Return 13
                End If
            End If
        End If
    End Function
    Private Function Cnv土地利用区域転用地目(ByRef pRowFind As DataRow) As Integer
        If Not IsDBNull(pRowFind.Item("調査土地利用区域地目")) And Val(pRowFind.Item("調査土地利用区域地目").ToString) Then
            Return Val(pRowFind.Item("調査土地利用区域地目").ToString)
        Else
            If n根拠条例 = 35 Then
                If Left(pRowFind.Item("登記簿地目名").ToString, 1) = "田" Then : Return 11
                ElseIf Left(pRowFind.Item("登記簿地目名").ToString, 1) = "畑" Then : Return 12
                Else : Return 13
                End If
            ElseIf Val(pRowFind.Item("農振法区分").ToString) = 1 Or Val(pRowFind.Item("農業振興地域").ToString) = 1 Then
                Select Case n根拠条例
                    Case 23, 30, 34
                        If Left(pRowFind.Item("登記簿地目名").ToString, 1) = "田" Then : Return 1
                        ElseIf Left(pRowFind.Item("登記簿地目名").ToString, 1) = "畑" Then : Return 2
                        Else : Return 3
                        End If
                    Case Else
                        If n農地区分 = 51 Then
                            If Left(pRowFind.Item("登記簿地目名").ToString, 1) = "田" Then : Return 11
                            ElseIf Left(pRowFind.Item("登記簿地目名").ToString, 1) = "畑" Then : Return 12
                            Else : Return 13
                            End If
                        Else
                            Select Case n農地区分
                                Case 1, 2, 3
                                    If Left(pRowFind.Item("登記簿地目名").ToString, 1) = "田" Then : Return 1
                                    ElseIf Left(pRowFind.Item("登記簿地目名").ToString, 1) = "畑" Then : Return 2
                                    Else : Return 3
                                    End If
                                Case Else
                                    If Left(pRowFind.Item("登記簿地目名").ToString, 1) = "田" Then : Return 11
                                    ElseIf Left(pRowFind.Item("登記簿地目名").ToString, 1) = "畑" Then : Return 12
                                    Else : Return 13
                                    End If
                            End Select

                        End If
                End Select
            ElseIf Val(pRowFind.Item("都市計画法区分").ToString) = 1 Or Val(pRowFind.Item("都市計画法").ToString) = 4 Then
                Select Case n根拠条例
                    Case 23, 30, 34
                        If Left(pRowFind.Item("登記簿地目名").ToString, 1) = "田" Then : Return 6
                        ElseIf Left(pRowFind.Item("登記簿地目名").ToString, 1) = "畑" Then : Return 7
                        Else : Return 8
                        End If
                    Case Else
                        If n農地区分 = 51 Then
                            If Left(pRowFind.Item("登記簿地目名").ToString, 1) = "田" Then : Return 11
                            ElseIf Left(pRowFind.Item("登記簿地目名").ToString, 1) = "畑" Then : Return 12
                            Else : Return 13
                            End If
                        Else
                            If n根拠条例 = 34 Then
                                If Left(pRowFind.Item("登記簿地目名").ToString, 1) = "田" Then : Return 6
                                ElseIf Left(pRowFind.Item("登記簿地目名").ToString, 1) = "畑" Then : Return 7
                                Else : Return 8
                                End If
                            Else
                                Select Case n農地区分
                                    Case 1, 2, 3
                                        If Left(pRowFind.Item("登記簿地目名").ToString, 1) = "田" Then : Return 1
                                        ElseIf Left(pRowFind.Item("登記簿地目名").ToString, 1) = "畑" Then : Return 2
                                        Else : Return 3
                                        End If
                                    Case Else
                                        If Left(pRowFind.Item("登記簿地目名").ToString, 1) = "田" Then : Return 11
                                        ElseIf Left(pRowFind.Item("登記簿地目名").ToString, 1) = "畑" Then : Return 12
                                        Else : Return 13
                                        End If
                                End Select
                            End If
                        End If
                End Select
            Else
                Select Case n農地区分
                    Case 1, 2, 3
                        If Left(pRowFind.Item("登記簿地目名").ToString, 1) = "田" Then : Return 1
                        ElseIf Left(pRowFind.Item("登記簿地目名").ToString, 1) = "畑" Then : Return 2
                        Else : Return 3
                        End If
                    Case Else
                        If Left(pRowFind.Item("登記簿地目名").ToString, 1) = "田" Then : Return 11
                        ElseIf Left(pRowFind.Item("登記簿地目名").ToString, 1) = "畑" Then : Return 12
                        Else : Return 13
                        End If
                End Select
            End If
        End If
    End Function

    Enum Enum様式番号
        権利設定移転 = 1
        貸借終了 = 2
        農地転用 = 3
    End Enum

    Enum Enum権利移動借賃
        設定無 = 0
        貸借 = 1
        期間 = 2
        貸借期間 = 3
    End Enum

    Enum Enum貸借終了
        設定無 = 0
        根拠条項 = 1
        基盤農地状況 = 2
        機構農地状況 = 3
    End Enum

    Enum Enum中間管理機構
        設定無し = 0
        分割処理 = 1
        Aから中間管理機構 = 2
        中間管理機構からB = 3
    End Enum
End Class

Public Enum EnumOutPut
    提出用 = 0
    確認用 = 1
End Enum