Imports HimTools2012
Imports HimTools2012.controls

Public Class CTabPage農業委員毎利用権実績
    Inherits CTabPageWithToolStrip

    Private mvarGrid As DataGridViewWithDataView
    Private mvarGrid2 As DataGridViewWithDataView
    Private mvarTabCtrl As TabControlBase
    Private mvar開始 As ToolStripDateTimePicker
    Private mvar終了 As ToolStripDateTimePicker
    Private WithEvents mvar検索 As ToolStripButtonEX

    Public Sub New()
        MyBase.New(True, True, "農業委員毎利用権実績", "農業委員毎利用権実績", CloseMode.NoMessage)

        Dim mvarLabel As New ToolStripLabel("検索範囲")
        mvar開始 = New ToolStripDateTimePicker()
        Dim mvarLabel2 As New ToolStripLabel("～")
        mvar終了 = New ToolStripDateTimePicker()
        mvar検索 = New ToolStripButtonEX()
        mvar検索.Text = "許可年月日検索"
        Me.ToolStrip.Items.AddRange({mvarLabel, mvar開始, mvarLabel2, mvar終了, mvar検索})

        Try

            mvarGrid = New DataGridViewWithDataView
            With mvarGrid
                .AddColumnDateTime("許可年月日", "許可年月日", "許可年月日", enumReadOnly.bReadOnly, "yyyy/MM/dd")
                .AddColumnText("許可番号", "許可番号", "許可番号", enumReadOnly.bReadOnly)
                '.AddColumnText("ID", "処理番号", "ID", enumReadOnly.bReadOnly)
                .AddColumnText("設定内容", "設定内容", "設定内容", enumReadOnly.bReadOnly)
                .AddColumnText("委員名1", "委員名1", "農業委員01氏名", enumReadOnly.bReadOnly)
                .AddColumnText("委員名2", "委員名2", "農業委員02氏名", enumReadOnly.bReadOnly)
                .AddColumnText("委員名3", "委員名3", "農業委員03氏名", enumReadOnly.bReadOnly)
                .AddColumnDateTime("始期", "始期", "始期", enumReadOnly.bReadOnly)
                .AddColumnDateTime("終期", "終期", "終期", enumReadOnly.bReadOnly)
                .AddColumnText("再設定", "再設定", "再設定", enumReadOnly.bReadOnly)

                '20170725農地情報の追加
                .AddColumnText("大字", "", "", enumReadOnly.bReadOnly)
                .AddColumnText("小字", "", "", enumReadOnly.bReadOnly)
                .AddColumnText("地番", "", "", enumReadOnly.bReadOnly)
                .AddColumnText("登記地目", "", "", enumReadOnly.bReadOnly)
                .AddColumnText("現況地目", "", "", enumReadOnly.bReadOnly)
                .AddColumnText("登記面積", "", "", enumReadOnly.bReadOnly)
                .AddColumnText("現況面積", "", "", enumReadOnly.bReadOnly)

                .AddColumnText("所有者名", "", "", enumReadOnly.bReadOnly)
                '.AddColumnText("農地リスト", "", "", enumReadOnly.bReadOnly)

                mvarGrid.AutoGenerateColumns = False
            End With

            mvarGrid2 = New DataGridViewWithDataView
            mvarTabCtrl = New TabControlBase

            Me.ControlPanel.Add(mvarTabCtrl)

            Dim pTabPage As New CTabPageWithToolStrip(False, True, "明細", "明細")
            pTabPage.Controls.Add(mvarGrid)
            mvarGrid.Createエクセル出力Ctrl(pTabPage.ToolStrip)
            mvarTabCtrl.AddPage(pTabPage)

            Dim pTabPage2 As New CTabPageWithToolStrip(False, True, "集計", "集計")
            pTabPage2.Controls.Add(mvarGrid2)
            mvarGrid2.Createエクセル出力Ctrl(pTabPage2.ToolStrip)
            mvarTabCtrl.AddPage(pTabPage2)

        Catch ex As Exception
            Stop
        End Try
    End Sub

    Private Sub mvar検索_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar検索.Click
        If mvar開始.Value < mvar終了.Value Then
            Dim sWhere As String = String.Format("[法令] IN (61,62) AND [状態]=2 AND [許可年月日]>=#{0}/{1}/{2}# AND [許可年月日]<=#{3}/{4}/{5}#",
                            mvar開始.Value.Month, mvar開始.Value.Day, mvar開始.Value.Year,
                            mvar終了.Value.Month, mvar終了.Value.Day, mvar終了.Value.Year
                        )
            Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D_申請] WHERE " & sWhere)
            Dim pTBL農地 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:農地Info].ID, V_大字.大字, V_小字.小字, [D:農地Info].地番, V_地目.名称 AS 登記地目名, [D:農地Info].登記簿面積, V_現況地目.名称 AS 現況地目名, [D:農地Info].実面積, [D:農地Info].所有者ID, [D:個人Info].氏名 AS 所有者名 FROM (((([D:農地Info] LEFT JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) LEFT JOIN V_小字 ON [D:農地Info].小字ID = V_小字.ID) LEFT JOIN V_地目 ON [D:農地Info].登記簿地目 = V_地目.ID) LEFT JOIN V_現況地目 ON [D:農地Info].現況地目 = V_現況地目.ID) LEFT JOIN [D:個人Info] ON [D:農地Info].所有者ID = [D:個人Info].ID;")
            pTBL農地.PrimaryKey = {pTBL農地.Columns("ID")}
            Dim pTBL転用農地 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D_転用農地].ID, V_大字.大字, V_小字.小字, [D_転用農地].地番, V_地目.名称 AS 登記地目名, [D_転用農地].登記簿面積, V_現況地目.名称 AS 現況地目名, [D_転用農地].実面積, [D_転用農地].所有者ID, [D:個人Info].氏名 AS 所有者名 FROM (((([D_転用農地] LEFT JOIN V_大字 ON [D_転用農地].大字ID = V_大字.ID) LEFT JOIN V_小字 ON [D_転用農地].小字ID = V_小字.ID) LEFT JOIN V_地目 ON [D_転用農地].登記簿地目 = V_地目.ID) LEFT JOIN V_現況地目 ON [D_転用農地].現況地目 = V_現況地目.ID) LEFT JOIN [D:個人Info] ON [D_転用農地].所有者ID = [D:個人Info].ID;")
            pTBL転用農地.PrimaryKey = {pTBL転用農地.Columns("ID")}
            Dim pTBLResult As New DataTable : CreatepTBLResult(pTBLResult)
            App農地基本台帳.TBL申請.MergePlus(pTBL)

            For Each pRow As DataRow In App農地基本台帳.TBL申請.Body.Rows
                Dim Ar筆リスト As Object = Split(pRow.Item("農地リスト").ToString, ";")
                For n As Integer = 0 To UBound(Ar筆リスト)
                    Dim Ar筆情報 As Object = Split(Ar筆リスト(n), ".")
                    Dim pRowFind As DataRow = Nothing

                    If InStr(Ar筆情報(0), "転用農地") > 0 Then
                        pRowFind = pTBL転用農地.Rows.Find(Ar筆情報(1))
                        If pRowFind Is Nothing Then
                            pRowFind = pTBL農地.Rows.Find(Ar筆情報(1))
                        End If
                    ElseIf InStr(Ar筆情報(0), "農地") > 0 Then
                        pRowFind = pTBL農地.Rows.Find(Ar筆情報(1))
                        If pRowFind Is Nothing Then
                            pRowFind = pTBL転用農地.Rows.Find(Ar筆情報(1))
                        End If
                    End If

                    If Not pRowFind Is Nothing Then
                        Dim pAddRow As DataRow = pTBLResult.NewRow
                        With pAddRow
                            .Item("法令") = pRow.Item("法令")
                            .Item("状態") = pRow.Item("状態")
                            .Item("許可年月日") = pRow.Item("許可年月日")
                            .Item("許可番号") = pRow.Item("許可番号")
                            .Item("処理番号") = pRow.Item("ID")
                            .Item("設定内容") = pRow.Item("名称")
                            .Item("農業委員1") = pRow.Item("農業委員1")
                            .Item("農業委員01氏名") = pRow.Item("農業委員01氏名")
                            .Item("農業委員02氏名") = pRow.Item("農業委員02氏名")
                            .Item("農業委員03氏名") = pRow.Item("農業委員03氏名")
                            .Item("始期") = pRow.Item("始期")
                            .Item("終期") = pRow.Item("終期")
                            .Item("再設定") = pRow.Item("再設定表示")
                            .Item("農地ID") = pRowFind.Item("ID")
                            .Item("大字") = pRowFind.Item("大字")
                            .Item("小字") = pRowFind.Item("小字")
                            .Item("地番") = pRowFind.Item("地番")
                            .Item("登記地目") = pRowFind.Item("登記地目名")
                            .Item("現況地目") = pRowFind.Item("現況地目名")
                            .Item("登記面積") = pRowFind.Item("登記簿面積")
                            .Item("現況面積") = pRowFind.Item("実面積")
                            .Item("所有者名") = pRowFind.Item("所有者名")
                            .Item("農地リスト") = pRow.Item("農地リスト")
                        End With

                        pTBLResult.Rows.Add(pAddRow)
                    End If
                Next
            Next
            mvarGrid.SetDataView(pTBLResult, "([農業委員01氏名] Is Not Null Or [農業委員02氏名] Is Not Null Or [農業委員03氏名] Is Not Null) AND " & sWhere, "[許可年月日],[許可番号]")

            Dim pView As DataView = mvarGrid.DataView
            Dim pTBL2 As New DataTable

            pTBL2.Columns.Add("管理番号", GetType(Long))
            pTBL2.Columns.Add("委員名", GetType(String))
            pTBL2.Columns.Add("設定件数", GetType(Integer))
            pTBL2.Columns.Add("設定筆数", GetType(Integer))
            pTBL2.Columns.Add("再設定件数", GetType(Integer))
            pTBL2.Columns.Add("再設定筆数", GetType(Integer))


            pTBL2.PrimaryKey = New DataColumn() {pTBL2.Columns("管理番号")}

            For Each mvarV As DataRowView In pView
                If Not IsDBNull(mvarV.Item("農業委員1")) Then
                    Dim pRow As DataRow = pTBL2.Rows.Find(mvarV.Item("農業委員1"))
                    If pRow Is Nothing Then
                        pRow = pTBL2.NewRow
                        pRow.Item("管理番号") = mvarV.Item("農業委員1")
                        pRow.Item("委員名") = mvarV.Item("農業委員01氏名")
                        pRow.Item("設定件数") = Not IIf(mvarV.Item("再設定") = "新規", False, True)
                        pRow.Item("再設定件数") = IIf(mvarV.Item("再設定") = "新規", False, True)
                        If IIf(mvarV.Item("再設定") = "新規", False, True) Then
                            pRow.Item("設定筆数") = 0
                            pRow.Item("再設定筆数") = Split(mvarV.Item("農地リスト").ToString, ";").Length
                        Else
                            pRow.Item("設定筆数") = Split(mvarV.Item("農地リスト").ToString, ";").Length
                            pRow.Item("再設定筆数") = 0
                        End If

                        pTBL2.Rows.Add(pRow)
                    Else
                        If IIf(mvarV.Item("再設定") = "新規", False, True) Then
                            pRow.Item("再設定件数") += 1
                            pRow.Item("再設定筆数") += Split(mvarV.Item("農地リスト").ToString, ";").Length
                        Else
                            pRow.Item("設定件数") += 1
                            pRow.Item("設定筆数") += Split(mvarV.Item("農地リスト").ToString, ";").Length
                        End If
                    End If
                End If
            Next

            mvarGrid2.SetDataView(pTBL2, "", "管理番号")
        Else
            MsgBox("期間を設定してください", MsgBoxStyle.Critical)
        End If
    End Sub

    Private Sub CreatepTBLResult(ByRef pTBL As DataTable)
        With pTBL
            .Columns.Add("法令", GetType(Integer))
            .Columns.Add("状態", GetType(Integer))
            .Columns.Add("許可年月日", GetType(Date))
            .Columns.Add("許可番号", GetType(Integer))
            .Columns.Add("処理番号", GetType(Integer))
            .Columns.Add("設定内容", GetType(String))
            .Columns.Add("農業委員1", GetType(Integer))
            .Columns.Add("農業委員01氏名", GetType(String))
            .Columns.Add("農業委員02氏名", GetType(String))
            .Columns.Add("農業委員03氏名", GetType(String))
            .Columns.Add("始期", GetType(Date))
            .Columns.Add("終期", GetType(Date))
            .Columns.Add("再設定", GetType(String))

            .Columns.Add("", GetType(String))
            .Columns.Add("農地ID", GetType(Decimal))
            .Columns.Add("大字", GetType(String))
            .Columns.Add("小字", GetType(String))
            .Columns.Add("地番", GetType(String))
            .Columns.Add("登記地目", GetType(String))
            .Columns.Add("現況地目", GetType(String))
            .Columns.Add("登記面積", GetType(Decimal))
            .Columns.Add("現況面積", GetType(Decimal))
            .Columns.Add("所有者名", GetType(String))
            .Columns.Add("農地リスト", GetType(String))
        End With
    End Sub
End Class
