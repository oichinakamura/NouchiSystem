Namespace CSV農地一覧
    Public Class COutPutCSV農地一覧
        Inherits COutPutCSV

        Private mvarGrid As New HimTools2012.controls.DataGridViewWithDataView
        Private WithEvents mvarTSBtn As New ToolStripButton("CSV出力")
        Private WithEvents mvarTSBtn調査済 As New ToolStripButton("CSV出力(農地のみ)")

        Private p出力View As DataView
        Private p公開データ構築 As 公開用データ初期化

        Public Sub New()
            MyBase.New(True, True, "CSV出力", "CSV出力")
            Me.ControlPanel.Add(mvarGrid)
            Me.ToolStrip.Items.AddRange({mvarTSBtn, mvarTSBtn調査済})

            p公開データ構築 = New 公開用データ初期化()
            Try
                With p公開データ構築
                    .Dialog.StartProc(True, True)
                    If .Dialog._objException Is Nothing = False Then
                        If .Dialog._objException.Message = "Cancel" Then
                            MsgBox("処理を中止しました。　", , "処理中止")
                        End If
                    End If
                End With

                p出力View = New DataView(p公開データ構築.出力TBL, "", "[市町村コード], [大字ID], [小字ID], [本番区分], [本番], [枝番区分], [枝番], [孫番区分], [孫番], [一部現況]", DataViewRowState.CurrentRows)
                mvarGrid.DataSource = p出力View
                mvarGrid.AllowUserToAddRows = False
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Sub

        Private Sub mvarTSBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarTSBtn.Click
            Try
                Dim objAcc As New CSV出力(p出力View, p公開データ構築.TBL公開用個人, p公開データ構築.都道府県ID, p公開データ構築.市町村CD)
                With objAcc
                    .Dialog.StartProc(True, True)

                    If .Dialog._objException Is Nothing = False Then
                        If .Dialog._objException.Message = "Cancel" Then
                            MsgBox("処理を中止しました。　", , "処理中止")
                        Else
                            'Throw objDlg._objException
                        End If
                    End If
                End With
            Catch ex As Exception

            End Try
        End Sub

        Private Sub mvarTSBtn調査済_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarTSBtn調査済.Click
            Try
                p出力View = New DataView(p公開データ構築.出力TBL, "[利用状況調査日] Is Not Null", "[市町村コード], [大字ID], [小字ID], [本番区分], [本番], [枝番区分], [枝番], [孫番区分], [孫番], [一部現況]", DataViewRowState.CurrentRows)
                Dim objAcc As New CSV出力(p出力View, p公開データ構築.TBL公開用個人, p公開データ構築.都道府県ID, p公開データ構築.市町村CD)
                With objAcc
                    .Dialog.StartProc(True, True)

                    If .Dialog._objException Is Nothing = False Then
                        If .Dialog._objException.Message = "Cancel" Then
                            MsgBox("処理を中止しました。　", , "処理中止")
                        Else
                            'Throw objDlg._objException
                        End If
                    End If
                End With
            Catch ex As Exception

            End Try
        End Sub
    End Class

    Public Class 公開用データ初期化
        Inherits HimTools2012.clsAccessor
        Public DSET As New DataSet
        Public 出力TBL As DataTable
        Public TBL公開用個人 As DataTable

        Public 都道府県ID As Integer = 0
        Public 市町村CD As Integer = 0
        Public 市町村名 As String = ""

        Public Sub New()
        End Sub

        Public Overrides Sub Execute()
            My.Application.DoEvents()
            Message = "データ初期化中...数分かかります"
            Me.BarStyle = ProgressBarStyle.Marquee
            My.Application.DoEvents()

            '/*****市町村コード、市町村名の取得*****/
            都道府県ID = Val(SysAD.DB(sLRDB).DBProperty("都道府県ID").ToString)
            市町村CD = Val(SysAD.DB(sLRDB).DBProperty("市町村ID").ToString)
            市町村名 = SysAD.DB(sLRDB).DBProperty("市町村名")


            Dim pDelTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT D_公開用個人.[PID], D_公開用個人.[AutoID] FROM D_公開用個人 WHERE (((D_公開用個人.[PID]) In (SELECT [PID] FROM [D_公開用個人] As Tmp GROUP BY [PID] HAVING Count(*)>1 ))) ORDER BY D_公開用個人.[PID];")
            Do
                Dim nID As New List(Of String)
                For Each pRow As DataRow In pDelTBL.Rows
                    nID.Add(CStr(pRow.Item("AutoID")))
                Next
                If nID.Count > 0 Then
                    SysAD.DB(sLRDB).ExecuteSQL("DELETE FROM [D_公開用個人] WHERE [AutoID] IN (" & Join(nID.ToArray, ",") & ")")
                End If
                pDelTBL = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT D_公開用個人.[PID], D_公開用個人.[AutoID] FROM D_公開用個人 WHERE (((D_公開用個人.[PID]) In (SELECT [PID] FROM [D_公開用個人] As Tmp GROUP BY [PID] HAVING Count(*)>1 ))) ORDER BY D_公開用個人.[PID];")
            Loop Until pDelTBL.Rows.Count = 0

            出力TBL = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:農地Info].ID AS 筆ID, [D:農地Info].所有者ID, [D:個人Info].氏名 AS 所有者氏名, [D:個人Info].フリガナ AS 所有者フリガナ, [D:個人Info].生年月日 AS 所有者生年月日, [D:農地Info].大字ID, V_大字.大字, [D:農地Info].小字ID, V_小字.小字, [D:農地Info].一部現況, [D:農地Info].地番, [D:農地Info].現況地目, V_現況地目.名称 AS 現況地目名, [D:農地Info].登記簿地目, V_地目.名称 AS 登記簿地目名, [D:農地Info].登記簿面積, [D:農地Info].実面積, [D:農地Info].農振法区分, [D:農地Info].都市計画法区分, [D:農地Info].生産緑地法, [D:農地Info].所有者農地意向, [D:農地Info].自小作別, [D:農地Info].小作形態," & _
                                            " [D:農地Info].小作開始年月日, [D:農地Info].小作終了年月日, [D:農地Info].利用状況調査日, [D:農地Info].利用状況調査農地法, [D:農地Info].利用意向調査日, [D:農地Info].利用意向根拠条項, [D:農地Info].利用意向意思表明日, [D:農地Info].利用意向意向内容区分, [D:農地Info].利用意向権利関係調査区分, [D:農地Info].利用意向公示年月日, [D:農地Info].利用意向通知年月日, [D:農地Info].農地法35の1通知日, [D:農地Info].農地法35の2通知日, [D:農地Info].農地法35の3通知日, [D:農地Info].勧告年月日, [D:農地Info].勧告内容, [D:農地Info].中間管理勧告日, [D:農地Info].再生利用困難農地," & _
                                            " [D:農地Info].農地法40裁定公告日, [D:農地Info].農地法43裁定公告日, [D:農地Info].農地法44の1裁定公告日, [D:農地Info].農地法44の3裁定公告日, [D:農地Info].中間管理権取得日, [D:農地Info].権利設定内容, [D:農地Info].利用配分計画設定期間年, [D:農地Info].利用配分計画設定期間月, [D:農地Info].利用配分計画始期日, [D:農地Info].利用配分計画終期日, [D:農地Info].管理者ID, [D:個人Info_1].氏名 AS 管理者氏名, [D:農地Info].借受人ID, [D:個人Info_2].氏名 AS 借受人氏名, [D:個人Info_2].フリガナ AS 借受人フリガナ, [D:個人Info_2].生年月日 AS 借受人生年月日 " & _
                                            " FROM (((((([D:農地Info] LEFT JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) LEFT JOIN V_小字 ON [D:農地Info].小字ID = V_小字.ID) LEFT JOIN V_地目 ON [D:農地Info].登記簿地目 = V_地目.ID) LEFT JOIN V_現況地目 ON [D:農地Info].現況地目 = V_現況地目.ID) LEFT JOIN [D:個人Info] ON [D:農地Info].所有者ID = [D:個人Info].ID) LEFT JOIN [D:個人Info] AS [D:個人Info_1] ON [D:農地Info].管理者ID = [D:個人Info_1].ID) LEFT JOIN [D:個人Info] AS [D:個人Info_2] ON [D:農地Info].借受人ID = [D:個人Info_2].ID" & _
                                            " WHERE ((([D:農地Info].大字ID)>0) AND (([D:農地Info].都市計画法区分) Is Null Or ([D:農地Info].都市計画法区分)<>1));")
            With 出力TBL
                .Columns.Add(New DataColumn("本番区分", GetType(String)))
                .Columns.Add(New DataColumn("本番", GetType(Integer)))
                .Columns.Add(New DataColumn("枝番区分", GetType(String)))
                .Columns.Add(New DataColumn("枝番", GetType(Integer)))
                .Columns.Add(New DataColumn("孫番区分", GetType(String)))
                .Columns.Add(New DataColumn("孫番", GetType(Integer)))
                .Columns.Add(New DataColumn("耕作者整理番号", GetType(String)))
                .Columns.Add(New DataColumn("小作設定期間年", GetType(Integer)))
                .Columns.Add(New DataColumn("小作設定期間月", GetType(Integer)))
                .Columns.Add(New DataColumn("市町村コード", GetType(Integer), Val(都道府県ID) & HimTools2012.StringF.Left(市町村CD & "0000", 4)))
                .Columns.Add(New DataColumn("市町村名", GetType(String), "'" & SysAD.市町村.市町村名 & "'"))
            End With

            TBL公開用個人 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM D_公開用個人")
            TBL公開用個人.PrimaryKey = New DataColumn() {TBL公開用個人.Columns("PID")}

            Dim TBL登記簿地目 As New 登記簿地目変換()
            TBL登記簿地目.Init()
            Dim TBL現況地目 As New 現況地目変換()
            TBL現況地目.Init()

            出力TBL.TableName = "対象農地明細"
            TBL登記簿地目.TableName = "登記簿地目変換TBL"
            TBL現況地目.TableName = "現況地目変換TBL"
            DSET.Tables.Add(出力TBL)
            DSET.Tables.Add(TBL登記簿地目)
            DSET.Tables.Add(TBL現況地目)

            DSET.Relations.Add("登記", TBL登記簿地目.Columns("登記名称"), 出力TBL.Columns("登記簿地目名"), False)
            出力TBL.Columns.Add("変換登記簿地目", GetType(Integer), "Parent(登記).登記ID")
            DSET.Relations.Add("現況", TBL現況地目.Columns("現況名称"), 出力TBL.Columns("現況地目名"), False)
            出力TBL.Columns.Add("変換現況地目", GetType(Integer), "Parent(現況).現況ID")

            For Each pRow As DataRow In 出力TBL.Rows
                Sub地番変換(pRow)
                Sub小作形態(pRow)
                Sub日付関連(pRow)
            Next
        End Sub

        Private Sub Sub地番変換(ByRef pRow As DataRow)
            Dim s分岐1 As String = ""
            Dim s分岐2 As String = ""
            Dim s分岐3 As String = ""
            Dim s分岐4 As String = ""

            pRow.Item("地番") = StrConv(pRow.Item("地番").ToString, vbNarrow)

            If InStr(pRow.Item("地番").ToString, "-") > 0 Then        '地番が"-"を含むかどうか
                pRow.Item("本番") = Val(HimTools2012.StringF.Left(pRow.Item("地番").ToString, InStr(pRow.Item("地番").ToString, "-") - 1))
                s分岐1 = Mid(pRow.Item("地番").ToString, InStr(pRow.Item("地番").ToString, "-") + 1)

                If InStr(s分岐1, "-") > 0 Then        '枝番以降が"-"を含むかどうか
                    If Char.IsNumber(s分岐1, 0) Then
                        pRow.Item("枝番") = Val(HimTools2012.StringF.Left(s分岐1, InStr(s分岐1, "-") - 1))
                        s分岐2 = Mid(s分岐1, InStr(s分岐1, "-") + 1)

                        If InStr(s分岐2, "-") > 0 Then
                            If Char.IsNumber(s分岐2, 0) Then
                                pRow.Item("孫番") = Val(HimTools2012.StringF.Left(s分岐2, InStr(s分岐2, "-") - 1))
                                pRow.Item("孫番区分") = StrConv(Mid(s分岐2, InStr(s分岐2, "-") + 1), VbStrConv.Wide)
                                '終了
                            Else
                                pRow.Item("枝番区分") = StrConv(HimTools2012.StringF.Left(s分岐2, InStr(s分岐2, "-") - 1), VbStrConv.Wide)
                                s分岐3 = Mid(s分岐2, InStr(s分岐2, "-") + 1)

                                If InStr(s分岐3, "-") > 0 Then
                                    pRow.Item("孫番") = Val(HimTools2012.StringF.Left(s分岐3, InStr(s分岐3, "-") - 1))
                                    pRow.Item("孫番区分") = StrConv(Mid(s分岐3, InStr(s分岐3, "-") + 1), VbStrConv.Wide)
                                    '終了
                                Else
                                    If Char.IsNumber(s分岐3, 0) Then
                                        pRow.Item("孫番") = Val(s分岐3)
                                    Else
                                        pRow.Item("孫番区分") = StrConv(s分岐3, VbStrConv.Wide)
                                    End If
                                    '終了
                                End If
                            End If
                        Else
                            If Char.IsNumber(s分岐2, 0) Then
                                pRow.Item("孫番") = Val(s分岐2)
                            Else
                                pRow.Item("枝番区分") = StrConv(s分岐2, VbStrConv.Wide)
                            End If
                            '終了
                        End If
                    Else
                        pRow.Item("本番区分") = StrConv(HimTools2012.StringF.Left(s分岐1, InStr(s分岐1, "-") - 1), VbStrConv.Wide)
                        s分岐2 = Mid(s分岐1, InStr(s分岐1, "-") + 1)

                        If InStr(s分岐2, "-") > 0 Then
                            If Char.IsNumber(s分岐2, 0) Then
                                pRow.Item("枝番") = Val(HimTools2012.StringF.Left(s分岐2, InStr(s分岐2, "-") - 1))
                                s分岐3 = Mid(s分岐2, InStr(s分岐2, "-") + 1)

                                If InStr(s分岐3, "-") > 0 Then
                                    If Char.IsNumber(s分岐3, 0) Then
                                        pRow.Item("孫番") = Val(HimTools2012.StringF.Left(s分岐3, InStr(s分岐3, "-") - 1))
                                        pRow.Item("孫番区分") = StrConv(Mid(s分岐3, InStr(s分岐3, "-") + 1), VbStrConv.Wide)
                                        '終了
                                    Else
                                        pRow.Item("枝番区分") = StrConv(HimTools2012.StringF.Left(s分岐3, InStr(s分岐3, "-") - 1), VbStrConv.Wide)
                                        s分岐4 = Mid(s分岐3, InStr(s分岐3, "-") + 1)

                                        If InStr(s分岐4, "-") > 0 Then
                                            pRow.Item("孫番") = Val(HimTools2012.StringF.Left(s分岐4, InStr(s分岐4, "-") - 1))
                                            pRow.Item("孫番区分") = StrConv(Mid(s分岐4, InStr(s分岐4, "-") + 1), VbStrConv.Wide)
                                        Else
                                            pRow.Item("孫番区分") = StrConv(Mid(s分岐4, InStr(s分岐4, "-") + 1), VbStrConv.Wide)
                                        End If
                                        '終了
                                    End If
                                Else
                                    If Char.IsNumber(s分岐3, 0) Then
                                        pRow.Item("孫番") = Val(s分岐3)
                                    Else
                                        pRow.Item("枝番区分") = StrConv(s分岐3, VbStrConv.Wide)
                                    End If
                                    '終了
                                End If
                            Else
                                pRow.Item("枝番区分") = StrConv(HimTools2012.StringF.Left(s分岐2, InStr(s分岐2, "-") - 1), VbStrConv.Wide)
                                s分岐3 = Mid(s分岐2, InStr(s分岐2, "-") + 1)

                                If InStr(s分岐3, "-") > 0 Then
                                    pRow.Item("孫番") = Val(HimTools2012.StringF.Left(s分岐3, InStr(s分岐3, "-") - 1))
                                    pRow.Item("孫番区分") = StrConv(Mid(s分岐3, InStr(s分岐3, "-") + 1), VbStrConv.Wide)
                                    '終了
                                Else
                                    If Char.IsNumber(s分岐3, 0) Then
                                        pRow.Item("孫番") = Val(s分岐3)
                                    Else
                                        pRow.Item("孫番区分") = StrConv(s分岐3, VbStrConv.Wide)
                                    End If
                                    '終了
                                End If
                            End If
                        Else
                            If Char.IsNumber(s分岐2, 0) Then
                                pRow.Item("枝番") = Val(s分岐2)
                            Else
                                pRow.Item("枝番区分") = StrConv(s分岐2, VbStrConv.Wide)
                            End If
                            '終了
                        End If
                    End If
                Else
                    If Char.IsNumber(s分岐1, 0) Then
                        pRow.Item("枝番") = Val(s分岐1)
                    Else
                        pRow.Item("本番区分") = StrConv(s分岐1, VbStrConv.Wide)
                    End If
                    '終了
                End If
            Else
                pRow.Item("本番") = Val(pRow.Item("地番").ToString)
                '終了
            End If
        End Sub

        Private Sub Sub小作形態(ByRef pRow As DataRow)
            Select Case pRow.Item("小作形態").ToString
                Case "1" : pRow.Item("小作形態") = 5
                Case "2" : pRow.Item("小作形態") = 4
                Case "3", "7", "8", "9" : pRow.Item("小作形態") = 6
                Case "4" : pRow.Item("小作形態") = 1
                Case "5" : pRow.Item("小作形態") = 2
                Case "6" : pRow.Item("小作形態") = 3
                Case Else
            End Select
        End Sub

        Private Sub Sub日付関連(ByRef pRow As DataRow)
            Dim 設定年 As Integer = 0
            Dim 設定月 As Integer = 0
            'Debug.Print(pRow.Item("小作開始年月日").ToString)
            If Val(pRow.Item("自小作別").ToString) = 0 Then
                pRow.Item("借受人ID") = DBNull.Value '20160915 新規追加
                pRow.Item("借受人氏名") = DBNull.Value '20160915 新規追加
                pRow.Item("借受人フリガナ") = DBNull.Value '20161110 新規追加
                pRow.Item("借受人生年月日") = DBNull.Value '20161110 新規追加
                pRow.Item("小作開始年月日") = DBNull.Value
                pRow.Item("小作終了年月日") = DBNull.Value
                pRow.Item("小作設定期間年") = DBNull.Value
                pRow.Item("小作設定期間月") = DBNull.Value
                pRow.Item("利用配分計画始期日") = DBNull.Value
                pRow.Item("利用配分計画終期日") = DBNull.Value
                pRow.Item("利用配分計画設定期間年") = DBNull.Value
                pRow.Item("利用配分計画設定期間月") = DBNull.Value
            Else
                pRow.Item("借受人ID") = Val(pRow.Item("借受人ID").ToString) '20160915 新規追加
                pRow.Item("借受人氏名") = pRow.Item("借受人氏名").ToString '20160915 新規追加
                pRow.Item("借受人フリガナ") = Val(pRow.Item("借受人フリガナ").ToString) '20161110 新規追加
                If Not IsDBNull(pRow.Item("借受人生年月日")) AndAlso IsDate(pRow.Item("借受人生年月日")) Then
                    pRow.Item("借受人生年月日") = pRow.Item("借受人生年月日").ToString '20161110 新規追加
                Else
                    pRow.Item("借受人生年月日") = DBNull.Value
                End If

                If Not IsDBNull(pRow.Item("小作開始年月日")) And Not IsDBNull(pRow.Item("小作終了年月日")) Then
                    If Year(pRow.Item("小作開始年月日")) >= 1945 And Year(pRow.Item("小作終了年月日")) >= 1945 Then
                        設定年 = Format(CDate(pRow.Item("小作終了年月日")), "yyyyMMdd") - Format(CDate(pRow.Item("小作開始年月日")), "yyyyMMdd")
                        設定年 = Math.Floor(設定年 / 100)
                        If HimTools2012.StringF.Right(設定年, 2) >= 1 And HimTools2012.StringF.Right(設定年, 2) <= 12 Then
                            設定月 = HimTools2012.StringF.Right(設定年, 2)
                        ElseIf HimTools2012.StringF.Right(設定年, 2) > 0 Then
                            設定月 = 12 - (100 - HimTools2012.StringF.Right(設定年, 2))
                        Else
                            設定月 = 0
                        End If
                        設定年 = Math.Floor(設定年 / 100)
                        pRow.Item("小作設定期間年") = 設定年
                        pRow.Item("小作設定期間月") = 設定月
                    Else
                        pRow.Item("小作開始年月日") = DBNull.Value
                        pRow.Item("小作終了年月日") = DBNull.Value
                        pRow.Item("小作設定期間年") = DBNull.Value
                        pRow.Item("小作設定期間月") = DBNull.Value
                    End If
                Else
                    pRow.Item("小作開始年月日") = DBNull.Value
                    pRow.Item("小作終了年月日") = DBNull.Value
                    pRow.Item("小作設定期間年") = DBNull.Value
                    pRow.Item("小作設定期間月") = DBNull.Value
                End If
                If Not IsDBNull(pRow.Item("利用配分計画始期日")) And Not IsDBNull(pRow.Item("利用配分計画終期日")) Then
                    If Year(pRow.Item("利用配分計画始期日")) >= 1945 And Year(pRow.Item("利用配分計画終期日")) >= 1945 Then
                        設定年 = Format(CDate(pRow.Item("利用配分計画終期日")), "yyyyMMdd") - Format(CDate(pRow.Item("利用配分計画始期日")), "yyyyMMdd")
                        設定年 = Math.Floor(設定年 / 100)
                        If HimTools2012.StringF.Right(設定年, 2) >= 1 And HimTools2012.StringF.Right(設定年, 2) <= 12 Then
                            設定月 = HimTools2012.StringF.Right(設定年, 2)
                        ElseIf HimTools2012.StringF.Right(設定年, 2) > 0 Then
                            設定月 = 12 - (100 - HimTools2012.StringF.Right(設定年, 2))
                        Else
                            設定月 = 0
                        End If
                        設定年 = Math.Floor(設定年 / 100)
                        pRow.Item("利用配分計画設定期間年") = 設定年
                        pRow.Item("利用配分計画設定期間月") = 設定月
                    Else
                        pRow.Item("利用配分計画始期日") = DBNull.Value
                        pRow.Item("利用配分計画終期日") = DBNull.Value
                        pRow.Item("利用配分計画設定期間年") = DBNull.Value
                        pRow.Item("利用配分計画設定期間月") = DBNull.Value
                    End If
                Else
                    pRow.Item("利用配分計画始期日") = DBNull.Value
                    pRow.Item("利用配分計画終期日") = DBNull.Value
                    pRow.Item("利用配分計画設定期間年") = DBNull.Value
                    pRow.Item("利用配分計画設定期間月") = DBNull.Value
                End If
            End If
        End Sub

    End Class

    Public Class CSV出力
        Inherits HimTools2012.clsAccessor

        Private mvar都道府県ID As Integer = 0
        Private mvar市町村CD As Integer = 0
        Private p出力View As DataView
        Private mvarTBL公開用個人 As DataTable

        Public Sub New(ByRef pView As DataView, ByRef pTBL公開用個人 As DataTable, ByRef n都道府県ID As Integer, ByRef n市町村CD As Integer)
            p出力View = pView
            mvar都道府県ID = n都道府県ID
            mvar市町村CD = n市町村CD
            mvarTBL公開用個人 = pTBL公開用個人
        End Sub

        Private Function GetHeader()
            Dim sHeader As String() = {"筆ID", "市町村名", "大字ID", "大字", "小字ID", "小字", "本番区分", "本番", "枝番区分", "枝番", "孫番区分", "孫番", "一部現況", _
                                       "登記簿地目", "現況地目", "登記簿面積", "実面積", "農振法区分", "都市計画法区分", "所有者ID", "所有者氏名", "所有者フリガナ", "所有者生年月日", "借受人ID", "借受人氏名", "借受人フリガナ", "借受人生年月日", _
                                       "生産緑地法", "所有者農地意向", "耕作者整理番号", "小作形態", "小作設定期間年", "小作設定期間月", "小作開始年月日", "小作終了年月日", _
                                       "中間管理権取得日", "権利設定内容", "利用配分計画設定期間年", "利用配分計画設定期間月", "利用配分計画始期日", "利用配分計画終期日", _
                                       "利用状況調査日", "利用状況調査農地法", _
                                       "利用意向調査日", "利用意向根拠条項", "利用意向意思表明日", "利用意向意向内容区分", "利用意向権利関係調査区分", "利用意向公示年月日", "利用意向通知年月日", _
                                       "農地法35の1通知日", "農地法35の2通知日", "農地法35の3通知日", _
                                       "勧告年月日", "勧告内容", "中間管理勧告日", "再生利用困難農地", _
                                       "農地法40裁定公告日", "農地法43裁定公告日", "農地法44の1裁定公告日", "農地法44の3裁定公告日"}

            Return sHeader
        End Function

        Public Overrides Sub Execute()
            Dim sCSV As New StringBEx("")
            Dim s列名取得 As String() = GetHeader() '20160916 新規追加

            Message = "CSV出力中.."
            Me.Maximum = p出力View.Count

            Me.Value = 0
            Dim s市町村CD As String = HimTools2012.StringF.Right("000" & mvar都道府県ID, 2) & HimTools2012.StringF.Left(mvar市町村CD & "0000", 4)

            '/******20160916 新規追加******/
            Dim pLineHeader As New StringBEx("市町村コード")
            For n As Integer = 0 To UBound(s列名取得)
                pLineHeader.SetNumber(s列名取得(n), 変換区分.全角)
            Next
            sCSV.Body.AppendLine(pLineHeader.Body.ToString)
            '/*****************************/

            For Each pRow As DataRowView In p出力View
                Try
                    Dim pLine As New StringBEx("""" & s市町村CD & """")
                    Me.Value += 1
                    Message = "CSV出力中(" & Me.Value & "/" & Me.Maximum & ").."
                    '基本的事項－基本的事項
                    pLine.SetNumber(pRow.Item("筆ID").ToString, 変換区分.半角) '20160915 新規追加
                    pLine.SetNumber(pRow.Item("市町村名").ToString, 変換区分.全角)
                    pLine.SetNumber(pRow.Item("大字ID").ToString, 変換区分.半角)
                    pLine.SetNumber(pRow.Item("大字").ToString, 変換区分.全角)
                    pLine.SetNumber(pRow.Item("小字ID").ToString, 変換区分.字コード)
                    pLine.SetNumber(pRow.Item("小字").ToString, 変換区分.外字)
                    pLine.SetNumber(pRow.Item("本番区分").ToString, 変換区分.全角)
                    pLine.SetNumber(pRow.Item("本番").ToString, 変換区分.半角)
                    pLine.SetNumber(pRow.Item("枝番区分").ToString, 変換区分.全角)
                    pLine.SetNumber(pRow.Item("枝番").ToString, 変換区分.半角)
                    pLine.SetNumber(pRow.Item("孫番区分").ToString, 変換区分.全角)
                    pLine.SetNumber(pRow.Item("孫番").ToString, 変換区分.半角)
                    pLine.SetNumber(IIf(Val(pRow.Item("一部現況").ToString) > 0, Val(pRow.Item("一部現況").ToString), ""), 変換区分.全角)
                    pLine.SetNumber(pRow.Item("変換登記簿地目").ToString, 変換区分.登記簿地目)
                    pLine.SetNumber(pRow.Item("変換現況地目").ToString, 変換区分.現況地目)
                    pLine.SetNumber(Val(pRow.Item("登記簿面積").ToString), 変換区分.面積)
                    pLine.SetNumber(Val(pRow.Item("実面積").ToString), 変換区分.面積)
                    pLine.SetNumber(Val(pRow.Item("農振法区分").ToString), 変換区分.半角)
                    If SysAD.市町村.市町村名 = "日置市" Then : pLine.SetNumber("", 変換区分.半角)
                    Else : pLine.SetNumber(Val(pRow.Item("都市計画法区分").ToString), 変換区分.半角)
                    End If

                    pLine.SetNumber(pRow.Item("所有者ID").ToString, 変換区分.半角) '20160915 新規追加
                    pLine.SetNumber(pRow.Item("所有者氏名").ToString, 変換区分.全角) '20160915 新規追加
                    pLine.SetNumber(pRow.Item("所有者フリガナ").ToString, 変換区分.半角) '20161110 新規追加
                    pLine.SetNumber(pRow.Item("所有者生年月日").ToString, 変換区分.日付) '20161110 新規追加

                    '/******20160915 新規追加******/
                    If Val(pRow.Item("自小作別").ToString) > 0 AndAlso IsDate(pRow.Item("小作終了年月日")) AndAlso pRow.Item("小作終了年月日") > Now Then
                        pLine.SetNumber(pRow.Item("借受人ID").ToString, 変換区分.半角)
                        pLine.SetNumber(pRow.Item("借受人氏名").ToString, 変換区分.全角)
                        pLine.SetNumber(pRow.Item("借受人フリガナ").ToString, 変換区分.半角) '20161110 新規追加
                        pLine.SetNumber(pRow.Item("借受人生年月日").ToString, 変換区分.日付) '20161110 新規追加
                    Else
                        pLine.SetNumber("", 変換区分.半角)
                        pLine.SetNumber("", 変換区分.全角)
                        pLine.SetNumber("", 変換区分.半角) '20161110 新規追加
                        pLine.SetNumber("", 変換区分.日付) '20161110 新規追加
                    End If
                    '/*****************************/

                    pLine.SetNumber(IIf(pRow.Item("生産緑地法").ToString = True, 2, 1), 変換区分.半角)
                    pLine.SetNumber(IIf(Val(pRow.Item("所有者農地意向").ToString) > 5, 5, Val(pRow.Item("所有者農地意向").ToString)), 変換区分.半角)
                    pLine.SetNumber(Fnc耕作者整理番号(pRow), 変換区分.半角)
                    '農地等の借地権等の設定状況－農地等の借地権等の設定状況
                    If Val(pRow.Item("自小作別").ToString) > 0 AndAlso IsDate(pRow.Item("小作終了年月日")) AndAlso pRow.Item("小作終了年月日") > Now Then
                        pLine.SetNumber(Val(pRow.Item("小作形態").ToString), 変換区分.半角)
                        pLine.SetNumber(Val(pRow.Item("小作設定期間年").ToString), 変換区分.半角)
                        pLine.SetNumber(Val(pRow.Item("小作設定期間月").ToString), 変換区分.半角)
                        pLine.SetNumber(IIf(pRow.Item("小作開始年月日").ToString = "", "", pRow.Item("小作開始年月日")), 変換区分.日付)
                        pLine.SetNumber(IIf(pRow.Item("小作終了年月日").ToString = "", "", pRow.Item("小作終了年月日")), 変換区分.日付)
                    Else
                        pLine.SetNumber(0, 変換区分.半角)
                        pLine.SetNumber("", 変換区分.半角)
                        pLine.SetNumber("", 変換区分.半角)
                        pLine.SetNumber("", 変換区分.日付)
                        pLine.SetNumber("", 変換区分.日付)
                    End If
                    '農地中間管理権と利用配分計画等－機構が農地中間管理権を取得した年月日
                    pLine.SetNumber(IIf(pRow.Item("中間管理権取得日").ToString = "", "", pRow.Item("中間管理権取得日")), 変換区分.日付)
                    '農地中間管理権と利用配分計画等－利用配分計画
                    pLine.SetNumber(Val(pRow.Item("権利設定内容").ToString), 変換区分.半角)
                    pLine.SetNumber(Val(pRow.Item("利用配分計画設定期間年").ToString), 変換区分.半角)
                    pLine.SetNumber(Val(pRow.Item("利用配分計画設定期間月").ToString), 変換区分.半角)
                    pLine.SetNumber(IIf(pRow.Item("利用配分計画始期日").ToString = "", "", pRow.Item("利用配分計画始期日")), 変換区分.日付)
                    pLine.SetNumber(IIf(pRow.Item("利用配分計画終期日").ToString = "", "", pRow.Item("利用配分計画終期日")), 変換区分.日付)
                    '利用状況調査－利用状況調査
                    pLine.SetNumber(IIf(pRow.Item("利用状況調査日").ToString = "", "", pRow.Item("利用状況調査日")), 変換区分.日付)
                    pLine.SetNumber(IIf(Val(pRow.Item("利用状況調査農地法").ToString) > 0, Val(pRow.Item("利用状況調査農地法").ToString), 3), 変換区分.半角)
                    '利用意向調査－利用意向調査
                    pLine.SetNumber(IIf(pRow.Item("利用意向調査日").ToString = "", "", pRow.Item("利用意向調査日")), 変換区分.日付)
                    pLine.SetNumber(Val(pRow.Item("利用意向根拠条項").ToString), 変換区分.半角)
                    '利用意向調査－調査結果
                    pLine.SetNumber(IIf(pRow.Item("利用意向意思表明日").ToString = "", "", pRow.Item("利用意向意思表明日")), 変換区分.日付)
                    pLine.SetNumber(Val(pRow.Item("利用意向意向内容区分").ToString), 変換区分.半角)
                    '利用意向調査－所有者が確知できない農地
                    pLine.SetNumber(IIf(Val(pRow.Item("利用意向権利関係調査区分").ToString) > 2, 0, Val(pRow.Item("利用意向権利関係調査区分").ToString)), 変換区分.半角)
                    pLine.SetNumber(IIf(pRow.Item("利用意向公示年月日").ToString = "", "", pRow.Item("利用意向公示年月日")), 変換区分.日付)
                    pLine.SetNumber(IIf(pRow.Item("利用意向通知年月日").ToString = "", "", pRow.Item("利用意向通知年月日")), 変換区分.日付)
                    '農地中間管理機構等との協議等－農地中間管理機構との協議
                    pLine.SetNumber(IIf(pRow.Item("農地法35の1通知日").ToString = "", "", pRow.Item("農地法35の1通知日")), 変換区分.日付)
                    pLine.SetNumber(IIf(pRow.Item("農地法35の2通知日").ToString = "", "", pRow.Item("農地法35の2通知日")), 変換区分.日付)
                    pLine.SetNumber(IIf(pRow.Item("農地法35の3通知日").ToString = "", "", pRow.Item("農地法35の3通知日")), 変換区分.日付)
                    '農地中間管理機構等との協議等－農地所有者への勧告
                    pLine.SetNumber(IIf(pRow.Item("勧告年月日").ToString = "", "", pRow.Item("勧告年月日")), 変換区分.日付)
                    pLine.SetNumber(Val(pRow.Item("勧告内容").ToString), 変換区分.半角)
                    pLine.SetNumber(IIf(pRow.Item("中間管理勧告日").ToString = "", "", pRow.Item("中間管理勧告日")), 変換区分.日付)
                    '農地中間管理機構等との協議等－再生利用困難な農地
                    pLine.SetNumber(Val(pRow.Item("再生利用困難農地").ToString), 変換区分.半角)
                    '裁定－裁定公告の状況
                    pLine.SetNumber(IIf(pRow.Item("農地法40裁定公告日").ToString = "", "", pRow.Item("農地法40裁定公告日")), 変換区分.日付)
                    pLine.SetNumber(IIf(pRow.Item("農地法43裁定公告日").ToString = "", "", pRow.Item("農地法43裁定公告日")), 変換区分.日付)
                    '措置命令－措置命令の内容
                    pLine.SetNumber(IIf(pRow.Item("農地法44の1裁定公告日").ToString = "", "", pRow.Item("農地法44の1裁定公告日")), 変換区分.日付)
                    pLine.SetNumber(IIf(pRow.Item("農地法44の3裁定公告日").ToString = "", "", pRow.Item("農地法44の3裁定公告日")), 変換区分.日付)

                    sCSV.Body.AppendLine(pLine.Body.ToString)
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try

                If _Cancel Then
                    Throw New Exception("Cancel")
                    Exit Sub
                End If
            Next

            '↓保存場所の確認だけ設定
            'Shift-JISで保存します。
            Dim dtNow As DateTime = DateTime.Now
            Dim dtToday As DateTime = dtNow.Date
            Dim StToday As String = dtToday.ToString("yyyyMMdd")
            Dim sPath As String = ""

            '/***名前を付けて保存***/
            With New SaveFileDialog
                .FileName = mvar都道府県ID & mvar市町村CD & "_" & StToday & "_公表用農地情報.csv"
                .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
                .Filter = "CSVファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*"

                If .ShowDialog = DialogResult.OK Then
                    sPath = .FileName
                End If
            End With

            Dim CSVText As New System.IO.StreamWriter(sPath, False, System.Text.Encoding.GetEncoding(932))
            CSVText.Write(sCSV.Body.ToString)
            CSVText.Dispose()

            MsgBox("CSVの出力が完了しました。")
            SysAD.ShowFolder(System.IO.Directory.GetParent(sPath).ToString)
        End Sub

        Private Function Fnc耕作者整理番号(ByRef pRow As DataRowView) As Decimal
            Dim 耕作者ID As Decimal = 0
            Dim 耕作者名 As String = ""

            If Val(pRow.Item("自小作別").ToString) > 0 Then
                耕作者ID = Val(pRow.Item("借受人ID").ToString)
                耕作者名 = pRow.Item("借受人氏名").ToString
            Else
                If Val(pRow.Item("管理者ID").ToString) <> 0 Then
                    耕作者ID = Val(pRow.Item("管理者ID").ToString)
                    耕作者名 = pRow.Item("管理者氏名").ToString
                Else
                    耕作者ID = Val(pRow.Item("所有者ID").ToString)
                    耕作者名 = pRow.Item("所有者氏名").ToString
                End If
            End If

            Dim p耕作者Row As DataRow = mvarTBL公開用個人.Rows.Find(耕作者ID)
            If p耕作者Row Is Nothing Then
                SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D_公開用個人]([PID],[氏名]) VALUES({0},'{1}')", 耕作者ID, 耕作者名)

                Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM D_公開用個人 WHERE [PID]=" & 耕作者ID)
                mvarTBL公開用個人.Merge(pTBL)
                p耕作者Row = mvarTBL公開用個人.Rows.Find(耕作者ID)

                pRow.Item("耕作者整理番号") = Val(p耕作者Row.Item("AutoID").ToString)
            Else
                pRow.Item("耕作者整理番号") = Val(p耕作者Row.Item("AutoID").ToString)
            End If

            Return pRow.Item("耕作者整理番号")
        End Function
    End Class
End Namespace



