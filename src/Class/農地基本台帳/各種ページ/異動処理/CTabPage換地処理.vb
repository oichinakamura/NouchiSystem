
Imports HimTools2012
Imports HimTools2012.CommonFunc

Public Class CTabPage換地処理
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Private mvarSP As SplitContainer
    Private mvarType As String
    Private WithEvents mvarGridView01 As HimTools2012.controls.DataGridViewWithDataView
    Private WithEvents mvarGridView02 As HimTools2012.controls.DataGridViewWithDataView

    Private mvarInnerTabC01 As HimTools2012.controls.ToolStripContainerEX
    Private mvarInnerTabC02 As HimTools2012.controls.ToolStripContainerEX

    Private sLst As New List(Of String)
    Private WithEvents mvar追加 As New ToolStripButton("追加")
    Private WithEvents mvar換地開始 As New ToolStripButton
    Private mvar換地後TBL As DataTable
    Private mvar共通大字 As Decimal = 0
    Private mvar小字Table As DataTable
    Private mvarExtNewID As Decimal = 0

    Public Sub New(ByVal nID() As String, ByVal NewID As Long, Optional ByVal sType As String = "換地")
        MyBase.New(True, True, sType & "処理", sType & "処理")
        mvarType = sType
        mvar換地開始.Text = mvarType & "開始"
        mvarGridView01 = New HimTools2012.controls.DataGridViewWithDataView
        Me.ImageKey = "List"
        mvarExtNewID = NewID
        With mvarGridView01
            .AddColumnText("ID", "ID", "ID", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("大字", "大字", "大字", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("小字", "小字", "小字", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("地番", "地番", "地番", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("登記簿面積", "登記簿面積", "登記簿面積", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("実面積", "実面積", "実面積", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("登記簿地目名", "登記簿地目名", "登記簿地目名", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("現況地目名", "現況地目名", "現況地目名", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("所有者氏名", "所有者氏名", "所有者氏名", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("自小作", "自小作", "自小作", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("借受人氏名", "借受人氏名", "借受人氏名", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("貸借始期", "貸借始期", "貸借始期", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("貸借終期", "小作終了年月日", "小作終了年月日", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("小作料", "小作料", "小作料", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("小作料単位", "小作料単位", "小作料単位", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            .AllowUserToDeleteRows = False
            .AutoGenerateColumns = False
        End With
        mvar換地後TBL = New DataTable
        With mvar換地後TBL
            .Columns.Add("大字ID", GetType(Integer))
            .Columns.Add("小字ID", GetType(Integer))
            .Columns.Add("地番", GetType(String))
            .Columns.Add("面積", GetType(Integer))
            .Columns.Add("地目", GetType(Integer))
            .Columns.Add("条件元地番", GetType(Decimal))
            .Columns.Add("換地済み", GetType(Boolean))
            .Columns.Add("固定ID", GetType(Decimal))
        End With

        mvarGridView01.AllowDrop = True
        sLst.Clear()

        If nID.Length > 0 Then
            For Each nSt As String In nID
                If GetKeyCode(nSt) <> 0 Then
                    Dim n As Decimal = GetKeyCode(nSt)
                    Dim pRow As DataRow = App農地基本台帳.TBL農地.FindRowByID(n)
                    sLst.Add(n)
                End If
            Next
            Dim St As String = Join(sLst.ToArray, ",")

            mvarGridView01.SetDataView(App農地基本台帳.TBL農地.Body, "[ID] In (" & St & ")", "")
            mvar共通大字 = mvarGridView01.DataView(0).Item("大字ID")
        End If

        mvarGridView02 = New HimTools2012.controls.DataGridViewWithDataView
        With mvarGridView02
            Dim 大字Table As DataTable = App農地基本台帳.DataMaster.GetClassTable("大字")
            .AddColumnCombo("大字ID", "大字ID", "大字ID", 大字Table, GetType(Integer), "ID", "名称", enumReadOnly.bCanEdit, DataGridViewContentAlignment.MiddleLeft)
            mvar小字Table = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:農地Info].小字ID AS ID, V_小字.名称 FROM [D:農地Info] INNER JOIN V_小字 ON [D:農地Info].小字ID = V_小字.ID WHERE ((([D:農地Info].大字ID)=" & mvar共通大字 & ")) GROUP BY [D:農地Info].小字ID, V_小字.名称;")
            .AddColumnCombo("小字ID", "小字ID", "小字ID", mvar小字Table, GetType(Integer), "ID", "名称", enumReadOnly.bCanEdit, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("地番", "地番", "地番", enumReadOnly.bCanEdit, DataGridViewContentAlignment.MiddleRight)
            Dim 地目Table As DataTable = App農地基本台帳.DataMaster.GetClassTable("地目")
            .AddColumnCombo("地目", "地目", "地目", 地目Table, GetType(Integer), "ID", "名称", enumReadOnly.bCanEdit, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnText("登記簿面積", "面積", "面積", enumReadOnly.bCanEdit, DataGridViewContentAlignment.MiddleRight)

            .AddColumnCombo("条件元地番", "条件元地番", "条件元地番", mvarGridView01.DataView, GetType(Integer), "ID", "土地所在", enumReadOnly.bCanEdit, DataGridViewContentAlignment.MiddleLeft)
            .AddColumnCheck("換地済み", "換地済み", "換地済み", enumReadOnly.bReadOnly, Color.Beige)
        End With
        mvarGridView02.SetDataView(mvar換地後TBL, "", "")

        mvarInnerTabC01 = New HimTools2012.controls.ToolStripContainerEX(mvarGridView01, True, True)
        mvarInnerTabC01.ToolBar.Items.Add(New ToolStripLabel("従前地"))
        mvarInnerTabC02 = New HimTools2012.controls.ToolStripContainerEX(mvarGridView02, True, True)
        mvarInnerTabC02.ToolBar.Items.Add(New ToolStripLabel(mvarType & "後"))
        mvarInnerTabC02.ToolBar.Items.Add(New ToolStripSeparator)

        mvarInnerTabC02.ToolBar.Items.Add(mvar追加)
        mvarInnerTabC02.ToolBar.Items.Add(mvar換地開始)
        '
        mvarSP = New SplitContainer
        mvarSP.Panel1.Controls.Add(mvarInnerTabC01)
        mvarSP.Panel2.Controls.Add(mvarInnerTabC02)
        mvarSP.Dock = DockStyle.Fill
        Me.Controls.Add(mvarSP)

        If NewID <> 0 Then
            Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [M_固定情報] WHERE [nID]=" & NewID)
            If pTBL.Rows.Count > 0 Then
                Dim pNRow As DataRow = mvar換地後TBL.NewRow
                pNRow.Item("固定ID") = NewID
                pNRow.Item("大字ID") = pTBL.Rows(0).Item("大字ID")
                pNRow.Item("小字ID") = pTBL.Rows(0).Item("小字ID")
                pNRow.Item("地番") = pTBL.Rows(0).Item("地番")
                Select Case SysAD.市町村.市町村名
                    Case "姶良市"
                        Select Case pTBL.Rows(0).Item("登記地目")
                            Case 10 : pNRow.Item("地目") = 1
                            Case 20 : pNRow.Item("地目") = 2
                            Case 30 : pNRow.Item("地目") = 3
                            Case 50 : pNRow.Item("地目") = 5
                            Case 1 : pNRow.Item("地目") = 1


                            Case Else
                                Stop
                        End Select
                    Case Else
                        pNRow.Item("地目") = pTBL.Rows(0).Item("登記地目")
                End Select
                pNRow.Item("面積") = pTBL.Rows(0).Item("登記面積")
                mvar換地後TBL.Rows.Add(pNRow)
                pNRow.Item("条件元地番") = mvarGridView01.DataView.Item(0).Item("ID")
            End If

        End If
        If mvarExtNewID <> 0 Then
            mvar換地開始_Click(mvar換地開始, Nothing)
        End If
    End Sub

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.NoMessage
        End Get
    End Property



    Private Sub mvarGridView02_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles mvarGridView02.CellValueChanged
        Select Case mvarGridView02.Columns(e.ColumnIndex).DataPropertyName
            Case "大字ID"
                mvar共通大字 = mvarGridView02.Item(e.ColumnIndex, e.RowIndex).Value
                mvar小字Table = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:農地Info].小字ID AS ID, V_小字.名称 FROM [D:農地Info] INNER JOIN V_小字 ON [D:農地Info].小字ID = V_小字.ID WHERE ((([D:農地Info].大字ID)=" & mvar共通大字 & ")) GROUP BY [D:農地Info].小字ID, V_小字.名称;")
                With CType(mvarGridView02.Columns("小字ID"), DataGridViewComboBoxColumn)
                    .DataSource = mvar小字Table
                End With
            Case "地番"
                Dim n大字ID As Integer = mvarGridView02.Item("大字ID", e.RowIndex).Value
                Dim s地番 As String = mvarGridView02.Item(e.ColumnIndex, e.RowIndex).Value
                If Not IsDBNull(n大字ID) AndAlso Not IsDBNull(s地番) Then
                    Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [M_固定情報] WHERE [大字ID]={0} AND [地番]='{1}'", n大字ID, s地番)
                    If pTBL.Rows.Count > 0 Then
                        mvarGridView02.Item("小字ID", e.RowIndex).Value = pTBL.Rows(0).Item("小字ID")
                        'Select Case SysAD.市町村.市町村名
                        '    Case "姶良市"
                        '        Select Case pTBL.Rows(0).Item("登記地目")
                        '            Case 10 : mvarGridView02.Item("地目", e.RowIndex).Value = 1
                        '            Case 20 : mvarGridView02.Item("地目", e.RowIndex).Value = 2
                        '            Case 30 : mvarGridView02.Item("地目", e.RowIndex).Value = 3
                        '            Case 40 : mvarGridView02.Item("地目", e.RowIndex).Value = 4

                        '            Case 1 : mvarGridView02.Item("地目", e.RowIndex).Value = 1

                        '            Case Else
                        '                Stop
                        '        End Select
                        '    Case Else
                        mvarGridView02.Item("地目", e.RowIndex).Value = pTBL.Rows(0).Item("登記地目")
                        'End Select
                        mvarGridView02.Item("登記簿面積", e.RowIndex).Value = pTBL.Rows(0).Item("登記面積")
                        mvarGridView02.Item("固定ID", e.RowIndex).Value = pTBL.Rows(0).Item("nID")
                    Else
                        MsgBox("指定された地番が固定資産に見つかりません。条件を変更してください。")
                        mvarGridView02.Item("固定ID", e.RowIndex).Value = 0
                    End If
                End If
            Case Else

        End Select
    End Sub

    Private Sub mvarGridView01_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles mvarGridView01.DragDrop
        Select Case e.Data.GetFormats()(0)
            Case "System.String"
                Dim sKeys() As String = Split(e.Data.GetData("System.String").ToString, ";")
                For Each sKey As String In sKeys
                    Select Case GetKeyHead(sKey)
                        Case "農地"
                            e.Effect = DragDropEffects.All

                            Dim nID As Decimal = GetKeyCode(sKey).ToString
                            If Not sLst.Contains(GetKeyCode(sKey).ToString) Then
                                sLst.Add(nID)
                            End If
                            mvarGridView01.RowFilter = "[ID] In (" & Join(sLst.ToArray, ",") & ")"
                    End Select
                Next
        End Select
    End Sub

    Private Sub mvarGridView01_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles mvarGridView01.DragEnter
        Select Case e.Data.GetFormats()(0)
            Case "System.String"
                Select Case GetKeyHead(e.Data.GetData("System.String"))
                    Case "農地"
                        e.Effect = DragDropEffects.All
                End Select
        End Select
    End Sub

    Private Sub mvar追加_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar追加.Click
        Dim pRow As DataRow = mvar換地後TBL.NewRow
        pRow.Item("大字ID") = mvar共通大字
        pRow.Item("面積") = 0
        pRow.Item("条件元地番") = mvarGridView01.Item("ID", 0).Value
        pRow.Item("固定ID") = 0
        mvar換地後TBL.Rows.Add(pRow)
    End Sub

    Private Sub mvar換地開始_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar換地開始.Click
        Dim sLst元所在 As New List(Of String)
        For Each pRow As DataRowView In mvarGridView01.DataView
            sLst元所在.Add(pRow.Item("土地所在").ToString)
        Next
        Dim s元地番 As String = Join(sLst元所在.ToArray, ",")

        Dim sLst後所在 As New List(Of String)
        Dim nLstUpdateID As New List(Of Decimal)

        If sLst.Count > 0 AndAlso mvar換地後TBL.Rows.Count > 0 Then
            For Each pRow As DataRow In mvar換地後TBL.Rows
                If Not IsDBNull(pRow.Item("地番")) Then
                    pRow.Item("地番") = Trim(Replace(pRow.Item("地番"), "　", ""))
                End If


                If IsDBNull(pRow.Item("条件元地番")) Then
                    MsgBox("条件元地番が設定されていません", MsgBoxStyle.Critical)
                    Exit Sub
                ElseIf IsDBNull(pRow.Item("大字ID")) Then
                    Exit Sub
                ElseIf IsDBNull(pRow.Item("地番")) Then
                    MsgBox("地番が設定されていません", MsgBoxStyle.Critical)
                    Exit Sub
                ElseIf IsDBNull(pRow.Item("面積")) Then
                    MsgBox("面積が設定されていません", MsgBoxStyle.Critical)
                    Exit Sub
                ElseIf IsDBNull(pRow.Item("地目")) Then
                    MsgBox("地目が設定されていません", MsgBoxStyle.Critical)
                    Exit Sub
                End If
            Next

            For Each pRow As DataRow In mvar換地後TBL.Rows
                If Not IsDBNull(pRow.Item("条件元地番")) Then
                    Dim UpdateFlg As Boolean = False
                    Dim nMinID = App農地基本台帳.TBL農地.MinID() - 1
                    If App農地基本台帳.TBL農地.FindRowByID(pRow.Item("固定ID")) IsNot Nothing Then
                        nMinID = pRow.Item("固定ID")
                        nLstUpdateID.Add(nMinID)
                        UpdateFlg = True
                    End If

                    Dim n元ID As Decimal = pRow.Item("条件元地番")
                    Dim Rs1 As DataRow = App農地基本台帳.TBL農地.FindRowByID(n元ID)

                    Dim Rs2 As DataRow = App農地基本台帳.TBL農地.NewRow
                    For Each pCol As DataColumn In Rs1.Table.Columns
                        If pCol.ColumnName = "ID" Then
                            Rs2.Item(pCol.ColumnName) = nMinID
                        Else
                            Rs2.Item(pCol.ColumnName) = Rs1.Item(pCol.ColumnName)
                        End If
                    Next
                    Rs2.Item("大字ID") = pRow.Item("大字ID")
                    Rs2.Item("小字ID") = pRow.Item("小字ID")
                    Rs2.Item("地番") = pRow.Item("地番")
                    Rs2.Item("登記簿面積") = pRow.Item("面積")
                    Rs2.Item("登記簿地目") = pRow.Item("地目")
                    Rs2.Item("実面積") = pRow.Item("面積")
                    Rs2.Item("現況地目") = pRow.Item("地目")
                    Rs2.Item("田面積") = IIf(Rs2.Item("登記簿地目名").ToString.EndsWith("田"), pRow.Item("面積"), 0)
                    Rs2.Item("畑面積") = IIf(Rs2.Item("登記簿地目名").ToString.EndsWith("畑"), pRow.Item("面積"), 0)
                    Rs2.Item("樹園地") = 0
                    Rs2.Item("自小作別") = 0
                    'Rs2.Item("所有者ID") = pRow.Item("所有者ID")

                    Dim pCopyRecord As New RecordSQL(Rs2)

                    Dim pTBLExt As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE [ID]=0")

                    Dim sSQL2 As String = pCopyRecord.InsertSQL("D:農地Info", pTBLExt)

                    Dim sRet As String = SysAD.DB(sLRDB).ExecuteSQL(sSQL2)


                    If sRet = "OK" Or sRet = "" Then
                        App農地基本台帳.TBL農地.Rows.Add(Rs2)
                        Dim s内容 As String = ""
                        Select Case mvarType
                            Case "換地" : s内容 = "[" & s元地番 & "]より換地後として作成"
                            Case "合筆" : s内容 = "[" & s元地番 & "]より合筆として作成"
                        End Select
                        Make農地履歴(nMinID, Now, Now, 土地異動事由.換地処理追加, enum法令.換地処理, s内容)
                        sLst後所在.Add(Rs2.Item("土地所在"))

                        For Each pRowV As DataRowView In mvarGridView01.DataView
                            SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D_土地系図]([自ID],[元ID],[元土地所在]) VALUES({0},{1},'{2}')", nMinID, pRowV.Item("ID"), pRowV.Item("土地所在"))
                        Next

                        pRow.Item("換地済み") = True
                    ElseIf UpdateFlg = True Then
                        pCopyRecord.NewValue("大字ID") = pRow.Item("大字ID")
                        pCopyRecord.NewValue("小字ID") = pRow.Item("小字ID")
                        pCopyRecord.NewValue("地番") = pRow.Item("地番")
                        pCopyRecord.NewValue("登記簿面積") = pRow.Item("面積")
                        pCopyRecord.NewValue("登記簿地目") = pRow.Item("地目")
                        pCopyRecord.NewValue("実面積") = pRow.Item("面積")
                        pCopyRecord.NewValue("現況地目") = pRow.Item("地目")
                        pCopyRecord.NewValue("田面積") = IIF(Rs2.Item("登記簿地目名").ToString.EndsWith("田"), pRow.Item("面積"), 0)
                        pCopyRecord.NewValue("畑面積") = IIF(Rs2.Item("登記簿地目名").ToString.EndsWith("畑"), pRow.Item("面積"), 0)
                        pCopyRecord.NewValue("樹園地") = 0
                        pCopyRecord.NewValue("自小作別") = 0

                        Dim sSQL3 As String = pCopyRecord.UpdateSQL("D:農地Info")
                        Dim sRet2 As String = SysAD.DB(sLRDB).ExecuteSQL(sSQL3)

                        If sRet2 = "OK" Or sRet2 = "" Then
                            Dim s内容 As String = ""
                            Select Case mvarType
                                Case "換地" : s内容 = "[" & s元地番 & "]より換地後として作成"
                                Case "合筆" : s内容 = "[" & s元地番 & "]より合筆として作成"
                            End Select
                            Make農地履歴(nMinID, Now, Now, 土地異動事由.換地処理追加, enum法令.換地処理, s内容)
                            sLst後所在.Add(Rs2.Item("土地所在"))

                            For Each pRowV As DataRowView In mvarGridView01.DataView
                                SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D_土地系図]([自ID],[元ID],[元土地所在]) VALUES({0},{1},'{2}')", nMinID, pRowV.Item("ID"), pRowV.Item("土地所在"))
                            Next

                            pRow.Item("換地済み") = True
                        End If
                    Else
                        pRow.Item("換地済み") = False
                    End If
                Else
                End If
            Next
            Dim s後地番 As String = Join(sLst後所在.ToArray, ",")

            If 1 Then
                For Each pRow As DataRowView In mvarGridView01.DataView
                    If SysAD.page農家世帯.DataViewCollection.ContainsKey("農地." & pRow.Item("ID")) Then
                        SysAD.page農家世帯.DataViewCollection.Item("農地." & pRow.Item("ID")).ClosePage()
                    End If

                    If Not nLstUpdateID.Contains(Val(pRow.Item("ID").ToString)) Then
                        農地削除(New DataView(App農地基本台帳.TBL農地.Body, "[ID]=" & pRow.Item("ID"), "", DataViewRowState.CurrentRows).ToTable, 土地異動事由.換地処理削除, C農地削除.enum転送先.削除農地, "[" & s後地番 & "]へ換地処理")
                    End If
                Next
            End If

            If mvarExtNewID = 0 Then
                MsgBox("処理が終了しました") : Me.DoClose()
            Else

            End If

        Else
            MsgBox(String.Format("{0}前、{0}後が正しく選択されていません", mvarType), MsgBoxStyle.Critical)
        End If
    End Sub
    Public Overrides Function QueryClose() As Boolean
        Return True
    End Function
    Private Sub mvarGridView01_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles mvarGridView01.UserDeletingRow
        MsgBox("選択した筆の情報は削除できません。一度閉じて選択しなおしてください。", MsgBoxStyle.Critical)
        'Dim nRow As Integer = e.Row.Index
        'Dim nID As Decimal = e.Row.Cells("ID").Value
        'Try
        '    sLst.Remove(nID)
        '    'mvar農地Col.DataSource = Nothing
        '    'mvarGridView01.ClearView()
        '    'mvarGridView01.Rows.Clear()
        '    mvarGridView01.SetDataView(App農地基本台帳.TBL農地.body, "[ID] In (" & Join(sLst.ToArray, ",") & ")", "")
        '    'mvar農地Col.DataSource = mvarGridView01.DataView.ToTable
        'Catch ex As Exception

        'End Try
    End Sub
End Class
