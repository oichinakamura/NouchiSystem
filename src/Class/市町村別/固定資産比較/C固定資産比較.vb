

Public Class C固定資産比較
    Inherits HimTools2012.controls.CTabPageWithToolStrip

#Region "コントロール"

    Private WithEvents mvarBtn固定Load As New ToolStripButton("固定資産テーブル読込")
    Private WithEvents mvarBtn農地Load As New ToolStripButton("農地テーブル読込")
    Private WithEvents mvarBtn比較フラグの初期化 As New ToolStripButton("比較フラグの初期化")
    Private WithEvents mvarBtn照合開始 As New ToolStripButton("照合開始")
    Private WithEvents mvarBtn農地外固定処理 As New ToolStripButton("農地外固定処理")
    Private WithEvents mvarBtn行の比較 As New ToolStripButton("行の比較")

    Private mvar固定Prog As New ToolStripProgressBar
    Private mvar固定ProgTxt As New ToolStripLabel
    Private WithEvents mvar固定Stop As New ToolStripButton("停止")
    Private mvar農地Prog As New ToolStripProgressBar
    Private mvar農地ProgTxt As New ToolStripLabel
    Private WithEvents mvar農地Stop As New ToolStripButton("停止")


    Private mvarSplitter As New SplitContainer
    Private WithEvents mvarGrid固定 As HimTools2012.controls.DataGridViewWithDataView
    Private WithEvents mvarGrid農地 As HimTools2012.controls.DataGridViewWithDataView


    Private mvar固定 As DataTable
    Private mvar農地 As DataTable
    Private mvar転用 As DataTable

    Private mvarlbl固定筆数 As New ToolStripLabel("")
    Private mvarlbl農地筆数 As New ToolStripLabel("")
    Private mvar固定なし As Integer = -12
    Private mvar固定分筆不明 As Integer = -13
    Private mvar削除 As Integer = -38
    Private mvar固定分筆可能 As Integer = -39
    Private mvar地番不整合 As Integer = -37
    Private mvar一部現況 As Integer = -36
#End Region

    Public Sub New(n固定なし As Integer)
        MyBase.New(True, False, "固定資産比較", "固定資産比較")
        mvar固定なし = n固定なし
        mvarSplitter.Dock = DockStyle.Fill
        mvarSplitter.Orientation = Orientation.Horizontal
        Me.ControlPanel.Add(mvarSplitter)

        mvarGrid固定 = New HimTools2012.controls.DataGridViewWithDataView()
        mvarGrid固定.ReadOnly = True
        mvarGrid農地 = New HimTools2012.controls.DataGridViewWithDataView()
        Dim TB01 As New HimTools2012.controls.ToolStripContainerEX(mvarGrid固定, True)
        Dim TB02 As New HimTools2012.controls.ToolStripContainerEX(mvarGrid農地, True)

        mvarSplitter.Panel1.Controls.Add(TB01)
        mvarSplitter.Panel2.Controls.Add(TB02)

        TB01.ToolBar.Items.Add(mvarBtn固定Load)
        TB01.ToolBar.Items.Add(mvarBtn比較フラグの初期化)
        mvar固定Stop.Enabled = False
        TB01.ToolBar.Items.AddRange({mvarBtn行の比較, mvarlbl固定筆数, New ToolStripSeparator, mvarBtn農地外固定処理, mvar固定Prog, mvar固定ProgTxt, mvar固定Stop})

        mvar農地Prog.Enabled = False
        Dim pExcelBtn As New ToolStripButton("エクセル出力")
        AddHandler pExcelBtn.Click, AddressOf mvarGrid農地.ToExcel
        TB02.ToolBar.Items.AddRange({mvarBtn農地Load, mvarlbl農地筆数, New ToolStripSeparator, pExcelBtn, New ToolStripSeparator, mvarBtn照合開始, mvar農地Prog, mvar農地ProgTxt, mvar農地Stop})
    End Sub

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.CloseOK
        End Get
    End Property

    Private Sub mvarBtn固定Load_Click(sender As Object, e As System.EventArgs) Handles mvarBtn固定Load.Click
        SysAD.DB(sLRDB).ExecuteSQL("UPDATE [M_固定情報] SET [一部現況]=0 WHERE [一部現況] Is Null")

        mvar固定 = SysAD.DB(sLRDB).GetTableBySqlSelectWithDialog("SELECT [nID],[農地ID],[大字ID],[小字ID],[地番],0 AS [本番],0 AS [枝番],[一部現況],[登記面積],[現況面積],[登記地目],[現況地目],[所有者ID],[異動事由],[異動事由TXT],[異動年月日]  FROM [M_固定情報] WHERE [農地ID]=0")
        mvar固定.TableName = "固定資産"

        Dim sSQL As New SQLStr
        For Each pRow As DataRow In mvar固定.Rows
            Dim sName As String = pRow.Item("地番")
            If sName.Length > 0 AndAlso sName <> "-" Then

                Do Until IsNumeric(sName.Substring(0, 1))
                    sName = sName.Substring(1)
                Loop
                'If Split(sName, "-").Length = 5 Then
                '    Dim n一部 As String = HimTools2012.StringF.Mid(sName, InStrRev(sName, "-") + 1)
                '    sName = HimTools2012.StringF.Left(sName, sName.Length - n一部.Length - 1)

                '    Do Until sName.EndsWith("-") = False
                '        sName = HimTools2012.StringF.Left(sName, Len(sName) - 1)
                '    Loop

                '    sSQL.AddUpdate("M_固定情報", "地番", sName, "nID", pRow.Item("nID"))
                '    sSQL.AddUpdate("M_固定情報", "一部現況", Val(n一部), "nID", pRow.Item("nID"))
                '    pRow.Item("地番") = sName
                '    pRow.Item("一部現況") = Val(n一部)
                'End If
                Dim sNameX() As String = Split(sName, "-")
                pRow.Item("本番") = Val(sNameX(0))
                If sNameX.Length > 1 Then
                    pRow.Item("枝番") = Val(sNameX(1))
                End If

                If sSQL.Length > 1024 Then
                    SysAD.DB(sLRDB).ExecuteSQL(sSQL.ToString)
                    sSQL.Clear()
                End If
            End If

        Next
        If sSQL.Length > 0 Then
            SysAD.DB(sLRDB).ExecuteSQL(sSQL.ToString)
            sSQL.Clear()
        End If
        mvarGrid固定.SetDataView(mvar固定, "[農地ID]=0", "[大字ID],[本番],[枝番]")
        Set固定残数()
    End Sub

    Private Sub mvarBtn農地Load_Click(sender As Object, e As System.EventArgs) Handles mvarBtn農地Load.Click
        SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [一筆コード] = -1 WHERE [大字ID]<0")
        SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [一筆コード] = -1 WHERE [所在] Is Not Null")
        SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [一部現況] = 0 WHERE [一部現況] Is Null")

        mvar農地 = SysAD.DB(sLRDB).GetTableBySqlSelectWithDialog("SELECT [ID],[一筆コード],[大字ID],[小字ID],[地番],0 AS [本番],0 AS [枝番],[一部現況],[登記簿面積],[実面積],[登記簿地目],[現況地目],[所有者ID],[先行異動],[先行異動日],[更新日],[自小作別] FROM [D:農地Info] WHERE [大字ID]>0 AND [一筆コード]=0")
        For Each pRow As DataRow In mvar農地.Rows
            Dim s地番 As String = pRow.Item("地番").ToString
            If InStr(s地番, "(") Then
                s地番 = HimTools2012.StringF.Left(s地番, InStr(s地番, "(") - 1)
            End If
            'pRow.Item("地番") = s地番
            Dim sName() As String = Split(s地番, "-")

            pRow.Item("本番") = Val(sName(0))
            If sName.Length > 1 Then
                pRow.Item("枝番") = Val(sName(1))
            End If
        Next


        mvarGrid農地.SetDataView(mvar農地, "[一筆コード]=0", "[大字ID],[本番],[枝番]")
        mvarlbl農地筆数.Text = "残り農地筆数" & mvar農地.Rows.Count

    End Sub

    Private Sub mvarBtn比較フラグの初期化_Click(sender As Object, e As System.EventArgs) Handles mvarBtn比較フラグの初期化.Click
        If MsgBox("全ての比較フラグを初期化しますか？" & vbCrLf & "比較済みのデータも最初からの処理になります。", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            If MsgBox("本当にいいですか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                SysAD.DB(sLRDB).ExecuteSQL("UPDATE [M_固定情報] SET [農地ID]=0")
                SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [一筆コード]=0")
                For Each pRow As DataRow In mvar固定.Rows
                    pRow.Item("農地ID") = 0
                Next
                For Each pRow As DataRow In mvar農地.Rows
                    pRow.Item("一筆コード") = 0
                Next
            End If
        End If
    End Sub

    Private Sub mvarBtn照合開始_Click(sender As Object, e As System.EventArgs) Handles mvarBtn照合開始.Click
        Dim sSQL As New SQLStr
        If mvar固定 IsNot Nothing AndAlso
           mvar農地 IsNot Nothing AndAlso
           MsgBox("開始します。", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then

            Dim pView As New DataView(mvar農地, "[一筆コード]=0", "[大字ID],[本番],[枝番]", DataViewRowState.CurrentRows)
            mvar農地Prog.Value = 0
            mvar農地Prog.Maximum = pView.Count
            Set農地Prog()
            sSQL.Clear()

            For Each pRow As DataRowView In pView
                If Not IsDBNull(pRow.Item("一筆コード")) AndAlso pRow.Item("一筆コード") > 0 Then

                Else
                    Dim fxRow() As DataRow = mvar固定.Select(String.Format("[大字ID]={0} AND [地番]='{1}' AND [農地ID]=0", pRow.Item("大字ID"), pRow.Item("地番")))
                    If fxRow IsNot Nothing AndAlso fxRow.Length > 0 Then
                        If fxRow.Length = 1 Then
                            CheckSQL(sSQL, 0)

                            My.Application.DoEvents()
                            If mvar農地Stop.Enabled = False Then
                                Exit For
                            End If

                            Dim fRow As DataRow = fxRow(0)

                            If fRow.Item("一部現況") = pRow.Item("一部現況") AndAlso
                                    CDec(fRow.Item("登記面積")) = CDec(Val(pRow.Item("登記簿面積").ToString)) AndAlso
                                    fRow.Item("登記地目") = pRow.Item("登記簿地目") AndAlso
                                    fRow.Item("現況地目") = pRow.Item("現況地目") AndAlso
                                    CDec(fRow.Item("現況面積")) = CDec(Val(pRow.Item("実面積").ToString)) AndAlso
                                    fRow.Item("所有者ID") = pRow.Item("所有者ID") Then
                                sSQL.AddUpdate("D:農地Info", pRow.Row, "一筆コード", fRow.Item("nID"), "ID", pRow.Item("ID"))
                                sSQL.AddUpdate("M_固定情報", fRow, "農地ID", pRow.Item("ID"), "nID", fRow.Item("nID"))
                            ElseIf fRow.Item("一部現況") = pRow.Item("一部現況") AndAlso
                                                                CDec(fRow.Item("登記面積")) = CDec(Val(pRow.Item("登記簿面積").ToString)) AndAlso
                                                                fRow.Item("登記地目") = pRow.Item("登記簿地目") AndAlso
                                                                fRow.Item("現況地目") = pRow.Item("現況地目") AndAlso
                                                                CDec(fRow.Item("現況面積")) = CDec(pRow.Item("実面積")) AndAlso
                                                                fRow.Item("所有者ID") <> pRow.Item("所有者ID") AndAlso Not pRow.Item("先行異動") Then
                                sSQL.AddUpdate("D:農地Info", pRow.Row, "所有者ID", fRow.Item("所有者ID"), "ID", pRow.Item("ID"))

                                sSQL.AddUpdate("D:農地Info", pRow.Row, "一筆コード", fRow.Item("nID"), "ID", pRow.Item("ID"))
                                sSQL.AddUpdate("M_固定情報", fRow, "農地ID", pRow.Item("ID"), "nID", fRow.Item("nID"))
                            ElseIf fRow.Item("一部現況") = pRow.Item("一部現況") AndAlso
                                                                fRow.Item("所有者ID") = pRow.Item("所有者ID") Then
                                sSQL.AddUpdate("D:農地Info", pRow.Row, "登記簿地目", fRow.Item("登記地目"), "ID", pRow.Item("ID"))
                                sSQL.AddUpdate("D:農地Info", pRow.Row, "登記簿面積", fRow.Item("登記面積"), "ID", pRow.Item("ID"))
                                sSQL.AddUpdate("D:農地Info", pRow.Row, "実面積", fRow.Item("現況面積"), "ID", pRow.Item("ID"))
                                sSQL.AddUpdate("D:農地Info", pRow.Row, "現況地目", fRow.Item("現況地目"), "ID", pRow.Item("ID"))
                            Else
                                If Not fRow.Item("一部現況") = pRow.Item("一部現況") Then
                                    sSQL.AddUpdate("D:農地Info", pRow.Row, "一部現況", fRow.Item("一部現況"), "ID", pRow.Item("ID"))
                                End If
                                If Not CDec(fRow.Item("登記面積")) = CDec(pRow.Item("登記簿面積")) Then
                                    sSQL.AddUpdate("D:農地Info", pRow.Row, "登記簿面積", fRow.Item("登記面積"), "ID", pRow.Item("ID"))
                                End If
                                If Not CDec(fRow.Item("登記地目")) = CDec(pRow.Item("登記簿地目")) Then
                                    sSQL.AddUpdate("D:農地Info", pRow.Row, "登記簿地目", fRow.Item("登記地目"), "ID", pRow.Item("ID"))
                                End If
                                If Not CDec(fRow.Item("現況面積")) = CDec(pRow.Item("実面積")) Then
                                    sSQL.AddUpdate("D:農地Info", pRow.Row, "実面積", fRow.Item("現況面積"), "ID", pRow.Item("ID"))
                                End If
                                If Not CDec(fRow.Item("現況地目")) = CDec(pRow.Item("現況地目")) Then
                                    sSQL.AddUpdate("D:農地Info", pRow.Row, "現況地目", fRow.Item("現況地目"), "ID", pRow.Item("ID"))
                                End If
                            End If
                        Else
                            If fxRow.Count = 2 And pRow.Item("一部現況") = 0 AndAlso fxRow(0).Item("一部現況") = 1 AndAlso fxRow(1).Item("一部現況") = 2 Then
                                sSQL.AddUpdate("D:農地Info", pRow.Row, "一筆コード", mvar一部現況, "ID", pRow.Item("ID"))
                            ElseIf fxRow.Count = 2 And pRow.Item("一部現況") = 0 AndAlso fxRow(0).Item("一部現況") = 2 AndAlso fxRow(1).Item("一部現況") = 1 Then
                                sSQL.AddUpdate("D:農地Info", pRow.Row, "一筆コード", mvar一部現況, "ID", pRow.Item("ID"))
                            Else
                                For Each fRow As DataRow In fxRow
                                    If fRow.Item("一部現況") = pRow.Item("一部現況") Then
                                        If CDec(fRow.Item("登記面積")) = CDec(pRow.Item("登記簿面積")) Then
                                            If fRow.Item("登記地目") = pRow.Item("登記簿地目") Then
                                                If fRow.Item("現況地目") = pRow.Item("現況地目") Then
                                                    If CDec(fRow.Item("現況面積")) = CDec(pRow.Item("実面積")) Then
                                                        If fRow.Item("所有者ID") = pRow.Item("所有者ID") Then
                                                            sSQL.AddUpdate("D:農地Info", pRow.Row, "一筆コード", fRow.Item("nID"), "ID", pRow.Item("ID"))
                                                            sSQL.AddUpdate("M_固定情報", fRow, "農地ID", pRow.Item("ID"), "nID", fRow.Item("nID"))
                                                        Else
                                                            If Not IsDBNull(pRow.Item("先行異動")) AndAlso pRow.Item("先行異動") = True Then
                                                                sSQL.AddUpdate("D:農地Info", pRow.Row, "一筆コード", fRow.Item("nID"), "ID", pRow.Item("ID"))
                                                                sSQL.AddUpdate("M_固定情報", fRow, "農地ID", pRow.Item("ID"), "nID", fRow.Item("nID"))
                                                            ElseIf IsDBNull(pRow.Item("更新日")) OrElse pRow.Item("更新日") < #1/1/2014# Then
                                                                sSQL.AddUpdate("D:農地Info", pRow.Row, "所有者ID", fRow.Item("所有者ID"), "ID", pRow.Item("ID"))
                                                            End If
                                                        End If
                                                    ElseIf IsDBNull(pRow.Item("更新日")) OrElse pRow.Item("更新日") < #1/1/2014# Then
                                                        sSQL.AddUpdate("D:農地Info", pRow.Row, "実面積", fRow.Item("現況面積"), "ID", pRow.Item("ID"))
                                                    End If
                                                Else
                                                    If CDec(fRow.Item("現況面積")) = CDec(pRow.Item("実面積")) Then
                                                        If IsDBNull(pRow.Item("更新日")) OrElse pRow.Item("更新日") < #1/1/2014# Then
                                                            sSQL.AddUpdate("D:農地Info", pRow.Row, "現況地目", fRow.Item("現況地目"), "ID", pRow.Item("ID"))
                                                        End If
                                                    Else

                                                    End If
                                                End If
                                            Else
                                                'If fRow.Item("現況地目") = pRow.Item("現況地目") Then
                                                '    If CDec(fRow.Item("現況面積")) = CDec(pRow.Item("実面積")) Then
                                                '        If fRow.Item("所有者ID") = pRow.Item("所有者ID") Then
                                                '            sSQL.AddUpdate("D:農地Info", "登記簿地目", fRow.Item("登記地目"), "ID", pRow.Item("ID"))
                                                '            pRow.Item("登記簿地目") = fRow.Item("登記地目")
                                                '        End If
                                                '    End If
                                                'End If
                                            End If
                                        Else

                                        End If
                                    Else
                                        If CDec(fRow.Item("登記面積")) = CDec(pRow.Item("登記簿面積")) Then
                                            If fRow.Item("登記地目") = pRow.Item("登記簿地目") Then
                                                If fRow.Item("現況地目") = pRow.Item("現況地目") Then
                                                    If CDec(fRow.Item("現況面積")) = CDec(pRow.Item("実面積")) Then
                                                        If fRow.Item("所有者ID") = pRow.Item("所有者ID") Then

                                                            sSQL.AddUpdate("D:農地Info", pRow.Row, "一部現況", fRow.Item("一部現況"), "ID", pRow.Item("ID"))
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Next
                            End If

                        End If
                    Else
                        ''中津のみ
                        'Dim pF As New DataView(mvar固定, "[大字ID]=" & pRow.Item("大字ID") & " AND [本番]=" & pRow.Item("本番"), "", DataViewRowState.CurrentRows)
                        'If pF.Count = 0 Then
                        '    sSQL.AddUpdate("D:農地Info", "一筆コード", mvar固定なし, "ID", pRow.Item("ID"))
                        '    pRow.Item("一筆コード") = mvar固定なし

                        'ElseIf pRow.Item("枝番") = 1 AndAlso pF.Count = 1 AndAlso pF(0).Item("枝番") = 2 Then
                        '    sSQL.AddUpdate("D:農地Info", "地番", pF(0).Item("地番"), "ID", pRow.Item("ID"))
                        '    sSQL.AddUpdate("D:農地Info", "一筆コード", pF(0).Item("nID"), "ID", pRow.Item("ID"))
                        '    sSQL.AddUpdate("M_固定情報", "農地ID", pRow.Item("ID"), "nID", pF(0).Item("nID"))

                        '    pRow.Item("一筆コード") = pF(0).Item("nID")
                        '    pF(0).Item("農地ID") = pRow.Item("ID")
                        'ElseIf pRow.Item("枝番") = 1 AndAlso pF.Count = 1 AndAlso pF(0).Item("枝番") = 3 Then
                        '    sSQL.AddUpdate("D:農地Info", "地番", pF(0).Item("地番"), "ID", pRow.Item("ID"))
                        '    sSQL.AddUpdate("D:農地Info", "一筆コード", pF(0).Item("nID"), "ID", pRow.Item("ID"))
                        '    sSQL.AddUpdate("M_固定情報", "農地ID", pRow.Item("ID"), "nID", pF(0).Item("nID"))

                        '    pRow.Item("一筆コード") = pF(0).Item("nID")
                        '    pF(0).Item("農地ID") = pRow.Item("ID")
                        'ElseIf pRow.Item("枝番") = 1 AndAlso pF.Count = 1 AndAlso pF(0).Item("枝番") = 0 Then
                        '    sSQL.AddUpdate("D:農地Info", "一筆コード", mvar地番不整合, "ID", pRow.Item("ID"))
                        '    pRow.Item("一筆コード") = mvar地番不整合
                        '    pF(0).Item("農地ID") = pRow.Item("ID")
                        'ElseIf pRow.Item("枝番") = 0 AndAlso pF.Count > 1 Then
                        '    Dim pX As New DataView(mvar農地, "[大字ID]=" & pRow.Item("大字ID") & " AND [本番]=" & pRow.Item("本番"), "", DataViewRowState.CurrentRows)

                        '    If pX.Count = 1 Then
                        '        If pX(0).Item("枝番") = 0 Then
                        '            sSQL.AddUpdate("D:農地Info", "一筆コード", mvar固定分筆可能, "ID", pRow.Item("ID"))
                        '            pRow.Item("一筆コード") = mvar固定分筆可能
                        '        End If
                        '    Else
                        '        For Each ppX As DataRowView In pX
                        '            ppX.Item("一筆コード") = mvar固定分筆不明
                        '            sSQL.AddUpdate("D:農地Info", "一筆コード", mvar固定分筆不明, "ID", pRow.Item("ID"))
                        '        Next
                        '    End If
                        'ElseIf pF.Count > 1 Then
                        '    Dim nngF As Boolean = True
                        '    For Each ppf As DataRowView In pF

                        '        If ppf.Item("登記地目") < 30 OrElse ppf.Item("現況地目") < 30 Then
                        '            nngF = False
                        '        End If
                        '    Next
                        '    If nngF Then
                        '        pRow.Item("一筆コード") = mvar削除
                        '        sSQL.AddUpdate("D:農地Info", "一筆コード", mvar削除, "ID", pRow.Item("ID"))
                        '    End If
                        'Else

                        'End If

                    End If

                End If
                CheckSQL(sSQL, 1024)
                mvar農地Prog.Value += 1
            Next
            CheckSQL(sSQL, 0)
            '住記異動
            SysAD.DB(sLRDB).ExecuteSQL("UPDATE M_住民情報 INNER JOIN [D:個人Info] ON M_住民情報.ID = [D:個人Info].ID SET [D:個人Info].世帯ID = [世帯No] WHERE (((M_住民情報.世帯No)<>0) AND (([D:個人Info].世帯ID)<>[世帯No])) OR (((M_住民情報.世帯No)<>0) AND (([D:個人Info].世帯ID) Is Null));")


            SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D:個人Info] ( ID, 世帯ID, 氏名, フリガナ, 住所, 住民区分, 生年月日, 性別, 続柄1, 続柄2, 続柄3, 行政区ID, 郵便番号, 電話番号 ) SELECT M_住民情報.ID, M_住民情報.世帯No, M_住民情報.氏名, M_住民情報.[フリガナ], M_住民情報.住所, M_住民情報.住民区分, M_住民情報.生年月日, M_住民情報.性別, M_住民情報.続柄, M_住民情報.続柄2, M_住民情報.続柄3, M_住民情報.行政区, M_住民情報.郵便番号, M_住民情報.電話番号 FROM ([D:農地Info] INNER JOIN M_住民情報 ON [D:農地Info].所有者ID = M_住民情報.ID) LEFT JOIN [D:個人Info] ON [D:農地Info].所有者ID = [D:個人Info].ID GROUP BY M_住民情報.ID, M_住民情報.世帯No, M_住民情報.氏名, M_住民情報.[フリガナ], M_住民情報.住所, M_住民情報.住民区分, M_住民情報.生年月日, M_住民情報.性別, M_住民情報.続柄, M_住民情報.続柄2, M_住民情報.続柄3, M_住民情報.行政区, M_住民情報.郵便番号, M_住民情報.電話番号, [D:個人Info].ID HAVING ((([D:個人Info].ID) Is Null));")
            SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D:世帯Info] ( 世帯主ID, ID ) SELECT [D:個人Info].ID, [D:個人Info].世帯ID FROM [D:個人Info] LEFT JOIN [D:世帯Info] ON [D:個人Info].世帯ID = [D:世帯Info].ID GROUP BY [D:個人Info].ID, [D:個人Info].世帯ID, [D:世帯Info].ID HAVING ((([D:世帯Info].ID) Is Null));")
            SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] INNER JOIN [D:個人Info] ON [D:農地Info].所有者ID = [D:個人Info].ID SET [D:農地Info].所有世帯ID = [D:個人Info].[世帯ID] WHERE ((([D:農地Info].所有世帯ID)<>[D:個人Info].[世帯ID]) AND (([D:個人Info].世帯ID)<>0));")
            SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:個人Info] INNER JOIN [D:世帯Info] ON [D:個人Info].世帯ID = [D:世帯Info].ID SET [D:世帯Info].世帯主ID = [D:個人Info].[ID] WHERE ((([D:世帯Info].世帯主ID) Is Null) AND (([D:個人Info].続柄1)=2) AND (([D:個人Info].住民区分)=0)) OR ((([D:世帯Info].世帯主ID)<>[D:個人Info].[ID]) AND (([D:個人Info].続柄1)=2) AND (([D:個人Info].住民区分)=0));")



            mvar農地ProgTxt.Text = ""
            mvar農地Prog.Value = 0
            MsgBox("終了しました")
        End If
    End Sub
    Private Sub CheckSQL(sSQL As SQLStr, nLim As Integer)
        If sSQL.Length > nLim Then
            Set農地Prog()
            Set固定残数()
            My.Application.DoEvents()

            SysAD.DB(sLRDB).ExecuteSQL(sSQL.ToString)
            sSQL.Clear()
        End If
    End Sub

    Public Class SQLStr
        Public Property Body As New System.Text.StringBuilder
        Public Sub New()

        End Sub
        Public Sub Clear()
            Body.Clear()
        End Sub

        Public Sub AddUpdate(ByVal sTable As String, ByRef pRow As DataRow, ByVal sField As String, ByVal oValue As Object, ByVal sKeyName As String, ByVal nID As Integer)
            pRow.Item(sField) = oValue

            Select Case pRow.Table.Columns(sField).DataType.ToString
                Case "System.Int32", "System.Single", "System.Int16", "System.Decimal", "System.Double"
                    Body.AppendLine(String.Format("UPDATE [{0}] SET [{1}]={2} WHERE [{4}]={3}", sTable, sField, oValue.ToString, nID, sKeyName))
                Case "String"
                    Body.AppendLine(String.Format("UPDATE [{0}] SET [{1}]='{2}' WHERE [{4}]={3}", sTable, sField, oValue.ToString, nID, sKeyName))
                Case Else
                    Stop
            End Select
        End Sub

        Public Function Length() As Integer
            Return Body.Length
        End Function

        Public Overrides Function ToString() As String
            Return Body.ToString
        End Function
    End Class

    Private mvarID As Long = 0
    Private sName As String = ""
    Private oValue As Object = Nothing

    Private Sub mvarGrid農地_CellBeginEdit(sender As Object, e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles mvarGrid農地.CellBeginEdit
        mvarID = mvarGrid農地.Rows(e.RowIndex).Cells("ID").Value
        sName = mvarGrid農地.Columns(e.ColumnIndex).DataPropertyName
        oValue = mvarGrid農地.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
    End Sub

    Private Sub mvarGrid農地_UserDeletingRow(sender As Object, e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles mvarGrid農地.UserDeletingRow
        mvarID = e.Row.Cells("ID").Value
        If MsgBox("削除しますか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then

        Else
            e.Cancel = True
        End If
    End Sub

    Private Sub mvarGrid農地_UserDeletedRow(sender As Object, e As System.Windows.Forms.DataGridViewRowEventArgs) Handles mvarGrid農地.UserDeletedRow
        Dim sSQL As String = ""
        If App農地基本台帳.TBL削除農地.FindRowByID(mvarID) IsNot Nothing Then
            農地Exec("DELETE FROM [D_削除農地] WHERE [ID]=" & mvarID)
        End If
        農地Exec("INSERT INTO [D_削除農地] SELECT * FROM [D:農地Info] WHERE [ID]=" & mvarID)
        農地Exec("DELETE FROM [D:農地Info] WHERE [ID]=" & mvarID)
    End Sub

    Private Sub 農地Exec(ByVal sSQL As String)
        HimTools2012.TextAdapter.AppendTextFile(My.Computer.FileSystem.SpecialDirectories.Desktop & "\農地基本台帳関連\固定比較操作.TXT", sSQL & vbCrLf, "SJIS")
        SysAD.DB(sLRDB).ExecuteSQL(sSQL)
    End Sub

    Private Sub mvarGrid農地_CellEndEdit(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles mvarGrid農地.CellEndEdit
        If sName.Length > 0 AndAlso mvarGrid農地.Columns(e.ColumnIndex).DataPropertyName = sName Then
            Dim pNewData As Object = mvarGrid農地.Rows(e.RowIndex).Cells(e.ColumnIndex).Value

            Dim sField As String = mvarGrid農地.Columns(e.ColumnIndex).DataPropertyName
            Select Case mvar農地.Columns(sField).DataType.ToString
                Case "System.String"
                    農地Exec("UPDATE [D:農地Info] SET [" & sName & "]='" & pNewData.ToString & "' WHERE [ID]=" & mvarID)
                Case Else
                    農地Exec("UPDATE [D:農地Info] SET [" & sName & "]=" & pNewData.ToString & " WHERE [ID]=" & mvarID)
            End Select
        End If

    End Sub

    Private mvar地番 As String = ""
    Private mvar大字 As Integer = 0

    Private Sub mvarGrid農地_RowHeaderMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles mvarGrid農地.RowHeaderMouseClick
        If e.Button = Windows.Forms.MouseButtons.Right Then
            Dim pMenu As New ContextMenuStrip
            mvarID = mvarGrid農地.Rows(e.RowIndex).Cells("ID").Value
            mvar大字 = mvarGrid農地.Rows(e.RowIndex).Cells("大字ID").Value
            mvar地番 = mvarGrid農地.Rows(e.RowIndex).Cells("地番").Value
            pMenu.Items.Add("分筆", Nothing, AddressOf 分筆)
            pMenu.Items.Add("不明地に設定", Nothing, AddressOf 不明地に設定)
            pMenu.Show(mvarGrid農地, e.Location)

        End If
    End Sub

    Private Sub 不明地に設定()
        農地Exec("UPDATE [D:農地Info] SET [一筆コード]=-144 WHERE [ID]=" & mvarID)
        Dim pView As New DataView(mvar農地, "[ID]=" & mvarID, "", DataViewRowState.CurrentRows)
        If pView.Count = 1 Then
            pView.Item(0).Item("一筆コード") = -144
        End If
    End Sub

    Private Sub 分筆()
        Dim sName As String = InputBox("分筆後の地番を入力してください。", "分筆処理", mvar地番)
        If sName.Length > 0 AndAlso Not sName = mvar地番 Then
            With New dlgSelectDataGridView
                Dim sPut As String = HimTools2012.StringF.Left(sName, InStr(sName & "-", "-") - 1)
                Dim pTBL As DataTable = New DataView(mvar固定, "[大字ID]=" & mvar大字 & " AND [本番]='" & sPut & "'", "", DataViewRowState.CurrentRows).ToTable

                .Grid.ReadOnly = False
                .AddViewColumn("選択", GetType(Boolean), pTBL)
                .AddViewColumn("元地の条件を引き継ぐ土地", GetType(Boolean), pTBL)
                .AddViewColumn("地番", GetType(String))
                .Grid.AutoGenerateColumns = False

                For Each pRow As DataRow In pTBL.Rows
                    If pRow.Item("地番") = sName Then
                        pRow.Item("元地の条件を引き継ぐ土地") = True
                    End If
                Next

                .Grid.DataSource = New DataView(mvar固定, "[大字ID]=" & mvar大字 & " AND [本番]='" & sPut & "'", "", DataViewRowState.CurrentRows).ToTable
                .Grid.MultiSelect = True
                If .ShowDialog() = DialogResult.OK Then
                    Dim nCount As Integer = 0
                    For Each pRow As DataRow In pTBL.Rows
                        nCount -= (pRow.Item("元地の条件を引き継ぐ土地") = True)
                    Next
                    If nCount > 1 Then
                        MsgBox("複数の筆に引き継ぐ土地が設定されています。", MsgBoxStyle.Critical)
                        Exit Sub
                    End If

                    For Each pRow As DataRow In pTBL.Rows
                        If pRow.Item("元地の条件を引き継ぐ土地") = True Then
                            農地Exec("UPDATE [D:農地Info] SET [地番]='" & pRow.Item("地番").ToString & "' WHERE [ID]=" & mvarID)
                        ElseIf Not IsDBNull(pRow.Item("選択")) AndAlso pRow.Item("選択") Then

                            'Dim sField As New System.Text.StringBuilder("[ID]")
                            'Dim pTBLX As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE [ID]=" & mvarID)

                            'For Each pCol As DataColumn In pTBLX.Columns
                            '    If sField.Length > 500 Then
                            '        Exit For
                            '    Else
                            '        sField.Append(",[" & pCol.ColumnName & "]")
                            '    End If
                            'Next
                            '農地Exec("INSERT INTO [D:農地Info] SELECT " & sField.ToString & " FROM [D:農地Info] WHERE [ID]=" & mvarID)


                        End If
                    Next

                End If
            End With

        Else
            MsgBox("中断しました")
        End If
    End Sub

    Private Sub mvarBtn農地外固定処理_Click(sender As Object, e As System.EventArgs) Handles mvarBtn農地外固定処理.Click
        Dim sSQL As New SQLStr
        Dim p農地地目 As New List(Of Integer)

        p農地地目.AddRange(CType(SysAD.市町村, C市町村別).市町村別登記地目CD(C市町村別.地目Type.農地地目))
        If mvar固定 IsNot Nothing AndAlso
              mvar農地 IsNot Nothing AndAlso
              MsgBox("開始します。", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Dim pView As New DataView(mvar固定, "[農地ID]=0", "[大字ID],[本番],[枝番]", DataViewRowState.CurrentRows)
            mvar固定Prog.Value = 0
            mvar固定Prog.Maximum = pView.Count
            mvar固定ProgTxt.Text = "進行中(" & mvar固定Prog.Value & "/" & mvar固定Prog.Maximum & ")"
            mvar固定Stop.Enabled = True
            My.Application.DoEvents()

            For Each pRow As DataRowView In pView
                If Not p農地地目.Contains(pRow.Item("登記地目")) AndAlso Not p農地地目.Contains(pRow.Item("現況地目")) Then
                    Dim sName2() As String = Split(pRow.Item("地番"), "-")

                    Dim pView2 As New DataView(mvar農地, "[一筆コード]=0 AND [大字ID]=" & pRow.Item("大字ID") & " AND ([地番]='" & sName2(0) & "' Or [地番] Like '" & sName2(0) & "-*')", "", DataViewRowState.CurrentRows)
                    If pView2.Count = 0 Then
                        sSQL.AddUpdate("M_固定情報", pRow.Row, "農地ID", -11, "nID", pRow.Item("nID"))
                    End If
                ElseIf Not p農地地目.Contains(pRow.Item("現況地目")) Then
                    Dim sName2() As String = Split(pRow.Item("地番"), "-")

                    Dim pView2 As New DataView(mvar農地, "[一筆コード]=0 AND [大字ID]=" & pRow.Item("大字ID") & " AND ([地番]='" & sName2(0) & "' Or [地番] Like '" & sName2(0) & "-*')", "", DataViewRowState.CurrentRows)
                    If pView2.Count = 0 Then
                        sSQL.AddUpdate("M_固定情報", pRow.Row, "農地ID", -12, "nID", pRow.Item("nID"))
                    End If
                Else
                    Dim sName2() As String = Split(pRow.Item("地番"), "-")

                    Dim pView2 As New DataView(mvar農地, "[一筆コード]=0 AND [大字ID]=" & pRow.Item("大字ID") & " AND ([地番]='" & sName2(0) & "' Or [地番] Like '" & sName2(0) & "-*')", "", DataViewRowState.CurrentRows)
                    If pView2.Count = 0 Then
                        sSQL.AddUpdate("M_固定情報", pRow.Row, "農地ID", -13, "nID", pRow.Item("nID"))
                    End If
                End If
                If sSQL.Length > 1024 Then
                    SysAD.DB(sLRDB).ExecuteSQL(sSQL.ToString)
                    sSQL.Clear()
                    Set固定Prog()
                    Set固定残数()
                    My.Application.DoEvents()
                End If
                My.Application.DoEvents()
                If mvar固定Stop.Enabled = False Then
                    Exit For
                End If
                mvar固定Prog.Value += 1
            Next
            If sSQL.Length > 0 Then
                SysAD.DB(sLRDB).ExecuteSQL(sSQL.ToString)
                sSQL.Clear()
                mvar固定Prog.Value = 0
                mvar固定ProgTxt.Text = ""
                Set固定残数()
                My.Application.DoEvents()
            End If

        End If
        mvar固定Stop.Enabled = False

        MsgBox("終了しました")
    End Sub

    Private Sub mvarBtn行の比較_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarBtn行の比較.Click
        If mvarGrid固定.SelectedRows IsNot Nothing AndAlso mvarGrid固定.SelectedRows.Count > 0 AndAlso mvarGrid農地.SelectedRows IsNot Nothing AndAlso mvarGrid農地.SelectedRows.Count > 0 Then
            Dim bCheck As Boolean = True
            bCheck = bCheck And CheckField("大字ID", "大字ID")
            bCheck = bCheck And CheckField("地番", "地番")
            bCheck = bCheck And CheckField("一部現況", "一部現況")
            bCheck = bCheck And CheckField("登記面積", "登記簿面積")
            bCheck = bCheck And CheckField("登記地目", "登記簿地目")
            bCheck = bCheck And CheckField("現況面積", "実面積")
            bCheck = bCheck And CheckField("現況地目", "現況地目")
            Dim bCheck所有者 As Boolean = CheckField("所有者ID", "所有者ID")


         

            If bCheck AndAlso Not bCheck所有者 AndAlso mvarGrid農地.SelectedRows(0).Cells("先行異動").Value Then

                Select Case MsgBox("所有者が異なるだけですが、農地台帳に先行異動が設定されています。" & vbCrLf & "農地台帳を優先する＝Yes、固定資産を優先する=No、処理しない=Cancel", MsgBoxStyle.YesNoCancel)
                    Case MsgBoxResult.Yes
                        mvarID = mvarGrid農地.SelectedRows(0).Cells("ID").Value
                        Dim n固定ID As Long = mvarGrid固定.SelectedRows(0).Cells("nID").Value
                        農地Exec("UPDATE [D:農地Info] SET [一筆コード]=" & n固定ID & " WHERE [ID]=" & mvarID)
                        農地Exec("UPDATE [M_固定情報] SET [農地ID]=" & mvarID & " WHERE [nID]=" & n固定ID)

                        mvarGrid農地.SelectedRows(0).Cells("一筆コード").Value = n固定ID
                        mvarGrid固定.SelectedRows(0).Cells("農地ID").Value = mvarID
                    Case MsgBoxResult.No
                    Case MsgBoxResult.Cancel
                End Select
            End If
            mvarGrid固定.ClearSelection()
            mvarGrid農地.ClearSelection()
        End If
    End Sub

    Private Function CheckField(ByVal s固定 As String, ByVal s農地 As String) As Boolean

        If mvarGrid固定.SelectedRows(0).Cells(s固定).Value <> mvarGrid農地.SelectedRows(0).Cells(s農地).Value Then
            mvarGrid固定.SelectedRows(0).Cells(s固定).Style.BackColor = Color.Pink
            mvarGrid農地.SelectedRows(0).Cells(s農地).Style.BackColor = Color.Pink
            Return False
        Else
            mvarGrid固定.SelectedRows(0).Cells(s固定).Style.BackColor = mvarGrid固定.Columns(s固定).DefaultCellStyle.BackColor
            mvarGrid農地.SelectedRows(0).Cells(s農地).Style.BackColor = mvarGrid農地.Columns(s農地).DefaultCellStyle.BackColor
            Return True
        End If

    End Function

    Private Sub mvarGrid固定_DataBindingComplete(sender As Object, e As System.Windows.Forms.DataGridViewBindingCompleteEventArgs) Handles mvarGrid固定.DataBindingComplete
        Set固定残数()
    End Sub

    Private Sub mvarGrid農地_DataBindingComplete(sender As Object, e As System.Windows.Forms.DataGridViewBindingCompleteEventArgs) Handles mvarGrid農地.DataBindingComplete
        mvarlbl農地筆数.Text = "残り農地筆数" & mvar農地.Rows.Count
    End Sub

    Private Sub Set固定残数()
        mvarlbl固定筆数.Text = "残り固定筆数" & mvarGrid固定.Rows.Count
    End Sub
    Private Sub Set農地残数()
        mvarlbl農地筆数.Text = "残り農地筆数" & mvarGrid農地.Rows.Count
    End Sub
    Private Sub Set固定Prog()

        mvar固定ProgTxt.Text = "進行中(" & mvar固定Prog.Value & "/" & mvar固定Prog.Maximum & ")"
        Set固定残数()
    End Sub
    Private Sub Set農地Prog()

        mvar農地ProgTxt.Text = "進行中(" & mvar農地Prog.Value & "/" & mvar農地Prog.Maximum & ")"
        Set農地残数()
    End Sub

    Private Sub mvar固定Stop_Click(sender As Object, e As System.EventArgs) Handles mvar固定Stop.Click
        mvar固定Stop.Enabled = False
    End Sub
    Private Sub mvar農地Stop_Click(sender As Object, e As System.EventArgs) Handles mvar農地Stop.Click
        mvar農地Stop.Enabled = False
    End Sub

End Class

