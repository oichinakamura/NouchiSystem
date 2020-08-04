
Imports System.ComponentModel
Imports System.CodeDom.Compiler
Imports System.Reflection
Imports HimTools2012

Public Class classPage非農地通知
    Inherits HimTools2012.controls.CTabPageWithToolStrip
    Private mvarDSet As New HimTools2012.Data.DataSetEx
    Private mvarTable As DataTable
    Private mvar住民区分 As DataTable

    Private WithEvents mvarGridView As DataGridLocal
    Private mvarTreeView As TreeViewLocal

    Public WithEvents mvarDBPath As HimTools2012.controls.ToolStripSpringTextBox

    Public WithEvents mvar発行番号 As ToolStripTextBoxWithLabel
    Public WithEvents mvar通知番号開始 As ToolStripTextBoxWithLabel
    Public WithEvents mvar決定総会日 As ToolStripDateTimePickerWithlabel
    Public WithEvents mvar通知日 As ToolStripDateTimePickerWithlabel
    Private mvarLDB As HimTools2012.Data.CLocalDataEngine
    Public Tbl宛名 As DataTable

    Public ReadOnly Property TBL非農地 As DataTable
        Get
            Return mvarTable
        End Get
    End Property


    Private Function GetLDB() As HimTools2012.Data.CLocalDataEngine
        If mvarLDB Is Nothing Then
            mvarLDB = New HimTools2012.Data.CLocalDataEngine("非農地通知")
            mvarLDB.LocalPath = mvarDBPath.Text
        End If

        Return mvarLDB
    End Function


    Public Sub New()
        MyBase.New(True, , , "非農地通知")


        Dim mvarSP As SplitContainer


        mvarSP = New SplitContainer
        mvarSP.Dock = DockStyle.Fill

        mvarTreeView = New TreeViewLocal(Me)
        mvarGridView = New DataGridLocal(Me)


        mvar発行番号 = New ToolStripTextBoxWithLabel("発行番号", "農地基本台帳;非農地通知;非農地通知書発行番号")
        mvar発行番号.WarningMessage = "必ず発行番号を入力してください。"

        mvar通知番号開始 = New ToolStripTextBoxWithLabel("通知番号開始", "農地基本台帳;非農地通知;非農地通知書通知番号開始")
        mvar通知番号開始.WarningMessage = "必ず通知開始番号を入力してください。"

        mvar決定総会日 = New ToolStripDateTimePickerWithlabel("決定総会日", "農地基本台帳;非農地通知;非農地通知書決定総会日")
        mvar通知日 = New ToolStripDateTimePickerWithlabel("通知日", "農地基本台帳;非農地通知;非農地通知書通知日")

        Me.ToolStrip.Items.AddRange({mvar発行番号, mvar通知番号開始, mvar決定総会日, mvar通知日})

        mvarTreeView.mvar通知書Path = GetSetting("農地基本台帳", "非農地通知", "非農地通知書Path", "")
        mvarTreeView.mvar通知発送文書Path = GetSetting("農地基本台帳", "非農地通知", "非農地通知発送文書Path", "")

        Me.ToolStrip.ItemAdd("データベース", New ToolStripButton("データベース"), AddressOf PC)
        mvarDBPath = Me.ToolStrip.ItemAdd("データパス", New HimTools2012.controls.ToolStripSpringTextBox())

        mvarDBPath.ReadOnly = True
        mvarDBPath.AutoSize = True
        mvarDBPath.Text = GetSetting(My.Application.Info.AssemblyName, "File", "DBPath", "'農政管理.MDB'を選択してください。")


        mvarSP.Panel1.Controls.Add(mvarTreeView.TreeCont)
        mvarSP.Panel2.Controls.Add(mvarGridView.GridCont)
        Me.ControlPanel.Add(mvarSP)
    End Sub

    Public Sub Init()
        If IO.File.Exists(mvarDBPath.Text) Then
            LoadDB()
        End If
    End Sub

    Private Sub PC()
        With New OpenFileDialog
            .Filter = "*.MDB|*.MDB"
            If .ShowDialog = DialogResult.OK Then
                mvarDBPath.Text = .FileName
                SaveSetting(My.Application.Info.AssemblyName, "File", "DBPath", .FileName)
                LoadDB()
            End If
        End With
    End Sub

    '20170515 農政情報.mdb 以外のデータベースが登録されている場合画面を表示できない
    '間違ったデータベースが登録されている場合、データベース参照画面を表示するよう修正
    Private Sub LoadDB()
        Try
            With GetLDB()

                mvarGridView.bLoading = True
                .ExecuteSQL_local("UPDATE D_非農地通知管理 SET D_非農地通知管理.発送番号 = Left$([発送番号],4) & '0' & Mid$([発送番号],5) WHERE (((Len([発送番号]))=7));")
                .ExecuteSQL_local("UPDATE D_非農地通知管理 SET D_非農地通知管理.発送番号 = Left$([発送番号],4) & '00' & Mid$([発送番号],5) WHERE (((Len([発送番号]))=6));")
                .ExecuteSQL_local("UPDATE D_非農地通知管理 SET D_非農地通知管理.発送番号 = Left$([発送番号],4) & '000' & Mid$([発送番号],5) WHERE (((Len([発送番号]))=5));")

                mvarTable = .GetTableBySqlSelect_Local("SELECT CODE,LID,NID,大字ID,大字,小字,地番,発送モード,本番,枝番,地目,面積,決定日,通知日,発送番号,発送先ID,発送先氏名,発送先郵便番号,発送先住所,登記名義人ID,登記名義人氏名,登記名義人郵便番号,登記名義人住所,登記名義人区分,現所有者ID,現所有者氏名,現所有者郵便番号,現所有者住所,現所有者区分,納税管理者ID,納税管理者氏名,納税管理者郵便番号,納税管理者住所,納税管理者区分,解消分類,解消年月日,備考,耕作作物,調査現況,調査備考,樹齢,エクセルにある FROM [D_非農地通知管理]")
                mvarTable.TableName = "D_非農地通知管理"
                mvarTable.Columns.Add(New DataColumn("削除", GetType(Boolean)))
                mvarDSet.Tables.Add(mvarTable)
                mvarTable.PrimaryKey = New DataColumn() {mvarTable.Columns("CODE")}

                mvarGridView.Columns.Clear()

                mvar住民区分 = New DataView(App農地基本台帳.DataMaster.Body, "[Class]='住民区分'", "", DataViewRowState.CurrentRows).ToTable
                mvar住民区分.TableName = "住民区分"
                mvarDSet.Tables.Add(mvar住民区分)

                mvarDSet.Relations.Add(New DataRelation("登記名義人住民区分", mvar住民区分.Columns("ID"), mvarTable.Columns("登記名義人区分"), False))
                mvarTable.Columns.Add("登記名義人区分名", GetType(String), "Parent(登記名義人住民区分).名称")
                mvarDSet.Relations.Add(New DataRelation("現所有者住民区分", mvar住民区分.Columns("ID"), mvarTable.Columns("現所有者区分"), False))
                mvarTable.Columns.Add("現所有者区分名", GetType(String), "Parent(現所有者住民区分).名称")
                mvarDSet.Relations.Add(New DataRelation("納税管理者住民区分", mvar住民区分.Columns("ID"), mvarTable.Columns("納税管理者区分"), False))
                mvarTable.Columns.Add("納税管理者区分名", GetType(String), "Parent(納税管理者住民区分).名称")

                mvarGridView.SetDataView(mvarTable, "[削除] Is Null", "大字ID,本番,枝番")
                mvarGridView.bFirst = True
                mvarGridView.bLoading = False
                mvarGridView.Columns("登記名義人区分名").DisplayIndex = mvarGridView.Columns("登記名義人区分").DisplayIndex + 1
                mvarGridView.Columns("現所有者区分名").DisplayIndex = mvarGridView.Columns("現所有者区分").DisplayIndex + 1
                mvarGridView.Columns("納税管理者区分名").DisplayIndex = mvarGridView.Columns("納税管理者区分").DisplayIndex + 1

            End With
            mvarGridView.ColumnVisible()
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & "データベース「農政管理.MDB」を選択しなおしてください。")
            PC()
        End Try
    End Sub



    Private Class ExtRow
        Private mvarS1 As String = ""
        Public Sub New(ByVal s1 As String)
            mvarS1 = s1
        End Sub
        Public ReadOnly Property ABC() As String
            Get
                Return mvarS1
            End Get
        End Property
        Public Property DEF() As String
            Get
                Return mvarS1
            End Get
            Set(ByVal value As String)
                mvarS1 = value
            End Set
        End Property
        Public Property DateNow() As DateTime
            Get
                Return Now
            End Get
            Set(ByVal value As DateTime)
                mvarS1 = value.ToString
            End Set
        End Property
    End Class

    Private Class UpdateList
        Inherits List(Of String)

        Public Sub AddList(ByVal ParamArray sFields() As String)
            For Each sField As String In sFields
                If Not Me.Contains(sField) Then
                    Me.Add(sField)
                End If
            Next
        End Sub

        Public Function UpdateValue(ByVal pRow As DataRow, ByVal sField As String, ByVal pValue As Object) As Boolean
            AddList(sField)
            pRow.Item(sField) = pValue

            Return True
        End Function

        Public Sub SaveRow(ByVal pDB As HimTools2012.Data.CLocalDataEngine, ByVal pRow As DataRow)
            If Me.Count > 0 Then
                Dim sUpdate As String = ""

                With pDB.GetOleConnection()
                    Dim command As OleDb.OleDbCommand = .CreateCommand()

                    For Each St As String In Me
                        sUpdate &= IIf(sUpdate.Length > 0, ", ", "") & "[" & St & "]=@" & St

                        Dim pColumn As DataColumn = pRow.Table.Columns(St)
                        Dim DT As System.Data.OleDb.OleDbParameter

                        Select Case pColumn.DataType.FullName
                            Case "System.Int32"
                                DT = New System.Data.OleDb.OleDbParameter("@" & St, OleDb.OleDbType.Integer)
                                DT.Value = pRow.Item(St)
                            Case "System.String"
                                DT = New System.Data.OleDb.OleDbParameter("@" & St, OleDb.OleDbType.VarChar, pRow.Item(St).ToString.Length)
                                DT.Value = pRow.Item(St)
                            Case "System.DateTime"
                                DT = New System.Data.OleDb.OleDbParameter("@" & St, OleDb.OleDbType.Date)
                                DT.Value = pRow.Item(St)
                            Case Else
                                Stop
                                DT = New System.Data.OleDb.OleDbParameter("@" & St, SqlDbType.NVarChar, St.Length)
                                DT.Value = St
                        End Select

                        command.Parameters.Add(DT)
                    Next
                    Dim sSQL As String = "UPDATE [D_非農地通知管理] SET " & sUpdate & " WHERE [CODE]='" & pRow.Item(sPrimaryKey) & "'"

                    command.CommandText = sSQL
                    command.ExecuteNonQuery()
                    .Close()
                End With
            End If
            Me.Clear()
            pRow.AcceptChanges()
        End Sub
        Dim sPrimaryKey As String = ""
        Public Sub New(ByVal sKey As String)
            sPrimaryKey = sKey
        End Sub
    End Class

    Private Class DataGridLocal
        Inherits HimTools2012.controls.DataGridViewWithDataView
        Public GridCont As ToolStripContainer

        Public DefaultMenu As ContextMenuStrip

        Public 非農地ContextMenu As New 非農地RowMenu
        Public 登記名義ContextMenu As ContextMenuStrip
        Public 納税義務者ContextMenu As ContextMenuStrip
        Public 非農地MultiContextMenu As ContextMenuStrip


        Public ToolStrip As New ToolStrip
        Public EditedRowIndex As Integer = -1
        Public ChangeList As New UpdateList("CODE")
        Private WithEvents mvarCHk発送済み非表示 As HimTools2012.controls.ToolStripCheckBoxEx
        Private WithEvents mvar発送先空欄を埋める As ToolStripButton

        Private WithEvents mvarID採番 As ToolStripButton
        Private mvarParent As classPage非農地通知
        Private txt検索 As ToolStripTextBoxWithLabel
        Private WithEvents txt検索Btn As ToolStripButton
        Private mvarViewSelect As ToolStripDropDownButton

        Public Sub New(ByVal pParent As classPage非農地通知)
            MyBase.New()
            mvarParent = pParent

            Me.AllowUserToAddRows = False
            ToolStrip.GripStyle = False

            登記名義ContextMenu = New ContextMenuStrip
            AddHandler 登記名義ContextMenu.Items.Add("この筆のみ登記名義人を発送先に設定").Click, AddressOf Change発送モード
            AddHandler 登記名義ContextMenu.Items.Add("同一登記名義人の筆を発送先に設定").Click, AddressOf Change発送モード

            納税義務者ContextMenu = New ContextMenuStrip
            AddHandler 納税義務者ContextMenu.Items.Add("この筆のみ納税義務者を発送先に設定").Click, AddressOf Change発送モード
            AddHandler 納税義務者ContextMenu.Items.Add("同一納税義務者の筆を発送先に設定").Click, AddressOf Change発送モード

            AddHandler 非農地ContextMenu.通知済み解除.Click, AddressOf 発送済み解除
            AddHandler 非農地ContextMenu.発送済み解除.Click, AddressOf 発送済み解除
            'AddHandler 非農地ContextMenu.誤入力解除.Click, AddressOf sub誤入
            AddHandler 非農地ContextMenu.非農地リスト削除.Click, AddressOf sub非農地レコードの削除

            非農地MultiContextMenu = New ContextMenuStrip
            AddHandler 非農地MultiContextMenu.Items.Add("選択された非農地の通知済みを解除する").Click, AddressOf 発送済み解除
            AddHandler 非農地MultiContextMenu.Items.Add("決定日を設定する").Click, AddressOf sub決定日
            AddHandler 非農地MultiContextMenu.Items.Add("非農地リストから削除する").Click, AddressOf sub非農地レコードの削除

            mvarCHk発送済み非表示 = New HimTools2012.controls.ToolStripCheckBoxEx("発送済みを非表示")
            Me.ToolStrip.Items.Add(mvarCHk発送済み非表示)

            mvar発送先空欄を埋める = New ToolStripButton("発送先空欄を埋める")
            Me.ToolStrip.Items.Add(mvar発送先空欄を埋める)

            Dock = DockStyle.Fill
            GridCont = New ToolStripContainer
            GridCont.Dock = DockStyle.Fill
            GridCont.TopToolStripPanel.Controls.Add(Me.ToolStrip)
            GridCont.ContentPanel.Controls.Add(Me)


            mvarViewSelect = New ToolStripDropDownButton
            mvarViewSelect.Text = "表示選択"
            Me.ToolStrip.Items.Add(mvarViewSelect)

            AddHandler mvarViewSelect.DropDownItems.Add("CODE").Click, AddressOf 表示選択MenuClick
            AddHandler mvarViewSelect.DropDownItems.Add("LID").Click, AddressOf 表示選択MenuClick
            AddHandler mvarViewSelect.DropDownItems.Add("NID").Click, AddressOf 表示選択MenuClick
            AddHandler mvarViewSelect.DropDownItems.Add("大字ID").Click, AddressOf 表示選択MenuClick
            AddHandler mvarViewSelect.DropDownItems.Add("本番").Click, AddressOf 表示選択MenuClick
            AddHandler mvarViewSelect.DropDownItems.Add("枝番").Click, AddressOf 表示選択MenuClick
            AddHandler mvarViewSelect.DropDownItems.Add("現所有者ID").Click, AddressOf 表示選択MenuClick
            AddHandler mvarViewSelect.DropDownItems.Add("現所有者氏名").Click, AddressOf 表示選択MenuClick
            AddHandler mvarViewSelect.DropDownItems.Add("現所有者郵便番号").Click, AddressOf 表示選択MenuClick
            AddHandler mvarViewSelect.DropDownItems.Add("現所有者住所").Click, AddressOf 表示選択MenuClick
            AddHandler mvarViewSelect.DropDownItems.Add("現所有者区分").Click, AddressOf 表示選択MenuClick
            AddHandler mvarViewSelect.DropDownItems.Add("現所有者区分名").Click, AddressOf 表示選択MenuClick
            '
            For Each pItem As ToolStripMenuItem In mvarViewSelect.DropDownItems
                Select Case pItem.Text
                    Case "LID" : pItem.Checked = True
                    Case "大字ID" : pItem.Checked = True
                End Select
            Next

            txt検索 = New ToolStripTextBoxWithLabel("文字検索", "非農地グリッド内検索")
            txt検索.AutoSize = False
            txt検索.Width = 400
            Me.ToolStrip.Items.Add(txt検索)
            txt検索Btn = New ToolStripButton("検索開始")
            Me.ToolStrip.Items.Add(txt検索Btn)

            DefaultMenu = New ContextMenuStrip
            Dim pMenuCopy As ToolStripMenuItem = DefaultMenu.Items.Add("コピー")
            pMenuCopy.ShortcutKeys = Keys.Control Or Keys.C
            AddHandler pMenuCopy.Click, AddressOf GCopy

            Dim pMenuPaste As ToolStripMenuItem = DefaultMenu.Items.Add("ペースト")
            pMenuPaste.ShortcutKeys = Keys.Control Or Keys.V
            AddHandler pMenuPaste.Click, AddressOf GPaste

            Dim pMenuDelete As ToolStripMenuItem = DefaultMenu.Items.Add("削除")
            pMenuDelete.ShortcutKeys = Keys.Delete
            AddHandler pMenuDelete.Click, AddressOf GDelete

            Dim pMenuNextF As ToolStripMenuItem = DefaultMenu.Items.Add("次を検索")
            pMenuNextF.ShortcutKeys = Keys.F3
            AddHandler pMenuNextF.Click, AddressOf txt検索Btn_Click

            mvarID採番 = New ToolStripButton("空白発送者IDを採番")
            Me.ToolStrip.Items.Add(mvarID採番)


            'Me.ContextMenuStrip = DefaultMenu
            AddHandler Me.ToolStrip.Items.Add("発送済み確定").Click, AddressOf TranseFar発送済み確定
            Me.Createエクセル出力Ctrl(Me.ToolStrip)
        End Sub

        Private mvar年度 As Integer = 1 '28

        Private Sub TranseFar発送済み確定()
            With New HimTools2012.PropertyGridDialog(New CInput年度設定(), "発送済み確定年度入力")
                CType(.ResultProperty, CInput年度設定).年度 = mvar年度 + 2018 '1988
                Dim 有効確定 As Integer = 0
                Dim 誤入力 As Integer = 0
                Dim 失敗 As Integer = 0
                Dim 総数 As Integer = 0
                If .ShowDialog = DialogResult.OK Then
                    Dim n年度 As String = Val(HimTools2012.StringF.Mid(CType(.ResultProperty, CInput年度設定).年度.ToString, 2))
                    If n年度 = 0 Then Exit Sub Else mvar年度 = n年度
                    If Me.SelectedRows IsNot Nothing AndAlso Me.SelectedRows.Count > 0 AndAlso MsgBox("選択行(誤入力・発送保留を除く)を確定しますか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then

                        If MsgBox("重要な項目なので再確認します。確定してリストを移しますか", MsgBoxStyle.YesNo) Then
                            With mvarParent.GetLDB()
                                Dim p2 As DataTable = .GetTableBySqlSelect_Local("SELECT * FROM [D_非農地情報] WHERE CODE='000'")
                                総数 = Me.SelectedRows.Count
                                For Each pRowG As DataGridViewRow In Me.SelectedRows
                                    Select Case DoTrance(mvarParent.GetLDB(), pRowG, p2, mvar年度 + 2018) '1988
                                        Case -1 : 失敗 += 1
                                        Case 0 : 誤入力 += 1
                                        Case 1 : 有効確定 += 1
                                    End Select

                                    My.Application.DoEvents()
                                Next
                                MsgBox(String.Format("終了しました。総数{0}、成功{1}、誤入力・発送保留により無効{2}、送信失敗{3}", 総数, 有効確定, 誤入力, 失敗))
                                Me.ClearSelection()
                                Try
                                    Me.CurrentCell = Me.Item(Me.CurrentCell.ColumnIndex, Me.CurrentCell.RowIndex + 1)
                                Catch ex As Exception

                                End Try

                                Me.Refresh()

                            End With
                        End If
                    ElseIf (Me.SelectedRows Is Nothing OrElse Me.SelectedRows.Count = 0) AndAlso MsgBox("選択行がないので全件(誤入力・発送保留を除く)を確定しますか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        If MsgBox("重要な項目なので再確認します。確定してリストを移しますか", MsgBoxStyle.YesNo) Then
                            With mvarParent.GetLDB()
                                Dim p2 As DataTable = .GetTableBySqlSelect_Local("SELECT * FROM [D_非農地情報] WHERE CODE='000'")

                                総数 = Me.Rows.Count
                                For Each pRowG As DataGridViewRow In Me.Rows
                                    Select Case DoTrance(mvarParent.GetLDB(), pRowG, p2, mvar年度 + 2018) '1988
                                        Case -1 : 失敗 += 1
                                        Case 0 : 誤入力 += 1
                                        Case 1 : 有効確定 += 1
                                    End Select

                                Next
                                MsgBox(String.Format("終了しました。総数{0}、成功{1}、誤入力・発送保留により無効{2}、送信失敗{3}", 総数, 有効確定, 誤入力, 失敗))
                                Me.ClearSelection()

                                Try
                                    Me.CurrentCell = Me.Item(Me.CurrentCell.ColumnIndex, Me.CurrentCell.RowIndex + 1)
                                Catch ex As Exception

                                End Try

                                Me.Refresh()
                            End With
                        End If
                    End If
                End If
            End With
        End Sub

        Private Function DoTrance(ByVal pDB As HimTools2012.Data.CLocalDataEngine, ByVal pRowG As DataGridViewRow, ByVal p2 As DataTable, n年度 As Integer) As Integer
            Try
                With pDB.GetOleConnection

                    If Not pRowG.Cells("発送モード").Value = 5 AndAlso Not pRowG.Cells("発送モード").Value = 4 Then

                        Dim command As OleDb.OleDbCommand = .CreateCommand()
                        Dim sSQLDel As String = "DELETE FROM [D_非農地情報] WHERE [CODE]='" & pRowG.Cells("CODE").Value.ToString & "'"
                        command.CommandText = sSQLDel
                        command.ExecuteNonQuery()

                        Dim sFields As New List(Of String)
                        Dim sValues As New ArrayList


                        For Each pCol As DataColumn In DataView.Table.Columns
                            If p2.Columns.Contains(pCol.ColumnName) Then
                                sFields.Add(pCol.ColumnName)

                                If IsDBNull(pRowG.Cells(pCol.ColumnName).Value) Then
                                    sValues.Add("Null")
                                Else
                                    Select Case pCol.DataType.FullName
                                        Case "System.DateTime"
                                            Dim pDT As DateTime = pRowG.Cells(pCol.ColumnName).Value
                                            sValues.Add(IIf(Not IsDate(pRowG.Cells(pCol.ColumnName).Value), "Null", "#" & pDT.Month & "/" & pDT.Day & "/" & pDT.Year & "#"))
                                        Case "System.Int32", "System.Decimal"
                                            sValues.Add(pRowG.Cells(pCol.ColumnName).Value)
                                        Case "System.String"
                                            sValues.Add(IIf(pRowG.Cells(pCol.ColumnName).Value = "", "Null", "'" & pRowG.Cells(pCol.ColumnName).Value & "'"))
                                        Case Else
                                            Stop
                                    End Select
                                End If
                            End If
                        Next
                        sFields.Add("年度")
                        sValues.Add(n年度)

                        Dim sSQL As String = "INSERT INTO [D_非農地情報](" & Join(sFields.ToArray, ",") & ") VALUES(" & Join(sValues.ToArray, ",") & ")"
                        command.CommandText = sSQL

                        command.ExecuteNonQuery()
                        Dim sSQLDelA As String = "DELETE FROM [D_非農地通知管理] WHERE [CODE]='" & pRowG.Cells("CODE").Value.ToString & "'"
                        command.CommandText = sSQLDelA
                        command.ExecuteNonQuery()
                        pRowG.Cells("削除").Value = True
                    Else
                        Return 0
                    End If
                    .Close()
                End With
                pRowG.Selected = False
                Return 1
            Catch ex As Exception
                Return -1
            End Try
        End Function



        Public Sub sub非農地レコードの削除()
            If Me.SelectedRows IsNot Nothing AndAlso Me.SelectedRows.Count > 0 AndAlso MsgBox("選択筆を非農地リストから削除しますか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                Dim nCODEs As New List(Of String)
                For Each pRowG As DataGridViewRow In Me.SelectedRows
                    Dim pRow As DataRow = CType(pRowG.DataBoundItem, DataRowView).Row
                    nCODEs.Add("""" & pRow.Item("CODE") & """")
                Next
                If nCODEs.Count > 0 Then
                    For Each nCode As String In nCODEs
                        Dim pDelRow As DataRow = mvarParent.TBL非農地.Rows.Find(Replace(nCode, """", ""))
                        mvarParent.TBL非農地.Rows.Remove(pDelRow)
                    Next
                    With mvarParent.GetLDB()
                        .ExecuteSQL_local("DELETE FROM [D_非農地通知管理] WHERE [CODE] IN (" & Join(nCODEs.ToArray, ",") & ")")
                        .ExecuteSQL_local("UPDATE [D_荒廃農地] SET [区分]=7 WHERE [CODE] IN (" & Join(nCODEs.ToArray, ",") & ")")
                    End With
                End If
            End If
        End Sub

        Private Sub GCopy()
            If Me.CurrentCell IsNot Nothing Then
                Clipboard.SetText(Me.CurrentCell.Value.ToString)
                Dim pRow As DataRow = mvarParent.TBL非農地.Rows.Find(Me.Rows(Me.CurrentCell.RowIndex).Cells("CODE").Value)
            End If
        End Sub
        Private Sub GPaste()
            If Me.CurrentCell IsNot Nothing AndAlso Clipboard.ContainsText Then
                Me.CurrentCell.Value = Clipboard.GetText()
                Dim pRow As DataRow = mvarParent.TBL非農地.Rows.Find(Me.Rows(Me.CurrentCell.RowIndex).Cells("CODE").Value)
                ChangeList.UpdateValue(pRow, "発送モード", 3)
                ChangeList.UpdateValue(pRow, Me.Columns(Me.CurrentCell.ColumnIndex).DataPropertyName, Me.CurrentCell.Value)
                ChangeList.SaveRow(mvarParent.GetLDB(), pRow)
            End If
        End Sub

        Private Sub GDelete()
            If Me.CurrentCell IsNot Nothing Then
                Me.CurrentCell.Value = DBNull.Value
                Dim pRow As DataRow = mvarParent.TBL非農地.Rows.Find(Me.Rows(Me.CurrentCell.RowIndex).Cells("CODE").Value)
                Select Case Me.Columns(Me.CurrentCell.ColumnIndex).DataPropertyName
                    Case "決定日"
                        ChangeList.UpdateValue(pRow, "発送モード", 3)
                    Case Else
                End Select
                ChangeList.UpdateValue(pRow, Me.Columns(Me.CurrentCell.ColumnIndex).DataPropertyName, Me.CurrentCell.Value)
                ChangeList.SaveRow(mvarParent.GetLDB(), pRow)
            End If
        End Sub

        Private Sub 発送済み解除(ByVal s As Object, ByVal e As EventArgs)
            Select Case s.text
                Case "通知済みを解除する"
                    Dim pRow As DataRow = mvarParent.TBL非農地.Rows.Find(Me.Rows(Me.CurrentCell.RowIndex).Cells("CODE").Value)
                    ChangeList.Clear()
                    ChangeList.UpdateValue(pRow, "発送番号", DBNull.Value)
                    ChangeList.UpdateValue(pRow, "通知日", DBNull.Value)
                    ChangeList.SaveRow(mvarParent.GetLDB(), pRow)
                Case "同発送先非農地を解除する"
                    With mvarParent.GetLDB()
                        Dim RX As Integer = Me.CurrentCell.RowIndex
                        Dim ID As Integer = Me.Rows(RX).Cells("登記名義人ID").Value

                        Dim Lst As New List(Of String)
                        For Each pDRow As DataGridViewRow In Me.Rows
                            If Not IsDBNull(pDRow.Cells("登記名義人ID").Value) AndAlso pDRow.Cells("登記名義人ID").Value = ID Then
                                Lst.Add(pDRow.Cells("CODE").Value)
                            End If
                        Next
                        For Each sCODE As String In Lst
                            Dim pRow As DataRow = mvarParent.TBL非農地.Rows.Find(sCODE)
                            ChangeList.Clear()
                            ChangeList.UpdateValue(pRow, "発送番号", DBNull.Value)
                            ChangeList.UpdateValue(pRow, "通知日", DBNull.Value)
                            ChangeList.SaveRow(mvarParent.GetLDB(), pRow)
                        Next
                    End With
                Case "選択された非農地の通知済みを解除する"
                    With mvarParent.GetLDB()
                        Dim Lst As New List(Of String)
                        For Each pDRow As DataGridViewRow In Me.SelectedRows
                            Lst.Add(pDRow.Cells("CODE").Value)
                        Next

                        For Each sCODE As String In Lst
                            Dim pRow As DataRow = mvarParent.TBL非農地.Rows.Find(sCODE)
                            ChangeList.Clear()
                            ChangeList.UpdateValue(pRow, "発送番号", DBNull.Value)
                            ChangeList.UpdateValue(pRow, "通知日", DBNull.Value)
                            ChangeList.SaveRow(mvarParent.GetLDB(), pRow)
                        Next
                    End With
            End Select

        End Sub

        Private Sub DataGridLocal_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles Me.CellBeginEdit
            If EditedRowIndex > -1 AndAlso EditedRowIndex <> e.RowIndex AndAlso ChangeList.Count > 0 Then
                ChangeList.SaveRow(mvarParent.GetLDB(), mvarParent.TBL非農地.Rows.Find(Me.Rows(EditedRowIndex).Cells("CODE").Value))
                EditedRowIndex = -1
            End If
        End Sub

        Private Sub 表示選択MenuClick(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim pItem As ToolStripMenuItem = CType(sender, ToolStripMenuItem)

            If Me.Columns.Contains(pItem.Text) Then
                pItem.Checked = Not pItem.Checked
                Me.ColumnVisible()
            End If
        End Sub

        Private Sub mvarGridView_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles Me.DataError

        End Sub


        Public Sub ColumnVisible()
            For Each pItem As ToolStripMenuItem In mvarViewSelect.DropDownItems
                Dim sName As String = pItem.Text
                If Me.Columns.Contains(sName) Then

                    Me.Columns(sName).Visible = pItem.Checked
                End If
            Next
            If Me.Columns.Contains("地番") Then
                Me.Columns("地番").Frozen = True
            End If
        End Sub

        Private Sub mvarGridView_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles Me.RowPostPaint
            If e.RowIndex Then

            End If
            If Me.Rows(e.RowIndex).Cells("発送モード").Value = -1 Then
                Me.Rows(e.RowIndex).ErrorText = "発送先が決定できません。"
            ElseIf Not Me.Rows(e.RowIndex).Cells("発送モード").Value AndAlso Me.Rows(e.RowIndex).Cells("発送先氏名").Value.ToString = "" Then
                Me.Rows(e.RowIndex).ErrorText = "発送先氏名がありません。"
            ElseIf Not Me.Rows(e.RowIndex).Cells("発送モード").Value AndAlso Me.Rows(e.RowIndex).Cells("発送先住所").Value.ToString = "" Then
                Me.Rows(e.RowIndex).ErrorText = "発送先住所がありません。"
            ElseIf Me.Rows(e.RowIndex).Cells("発送モード").Value = 4 OrElse Me.Rows(e.RowIndex).Cells("発送モード").Value = 5 Then
                If Me.Rows(e.RowIndex).DefaultCellStyle.Font IsNot Nothing AndAlso Me.Rows(e.RowIndex).DefaultCellStyle.Font.Strikeout = True Then

                Else
                    Dim cellStyle As New DataGridViewCellStyle()
                    cellStyle.Font = New Font(Me.Font.FontFamily, Me.Font.Size, FontStyle.Strikeout)
                    cellStyle.ForeColor = Color.DarkGray
                    Me.Rows(e.RowIndex).DefaultCellStyle = cellStyle
                End If
            ElseIf Me.Rows(e.RowIndex).DefaultCellStyle.Font IsNot Nothing AndAlso Me.Rows(e.RowIndex).DefaultCellStyle.Font.Strikeout = True Then
                Dim cellStyle As New DataGridViewCellStyle()
                cellStyle.Font = New Font(Me.Font.FontFamily, Me.Font.Size)
                Me.Rows(e.RowIndex).DefaultCellStyle = cellStyle
            Else
                Me.Rows(e.RowIndex).ErrorText = ""
            End If
        End Sub

        Private Sub mvarGridView_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
            If Me.bFirst AndAlso mvarParent.TBL非農地 IsNot Nothing Then
                Me.bFirst = False
                ColumnVisible()
                If Me.Columns.Contains("発送モード") Then
                    Dim table発送モード As New DataTable("発送モード")
                    table発送モード.Columns.Add("Display", GetType(String))
                    table発送モード.Columns.Add("Value", GetType(Integer))
                    table発送モード.Rows.Add("登記名義人", 0)
                    table発送モード.Rows.Add("現所有者", 1)
                    table発送モード.Rows.Add("納税管理人", 2)
                    table発送モード.Rows.Add("手入力", 3)
                    table発送モード.Rows.Add("発送保留", 4)
                    table発送モード.Rows.Add("誤入力", 5)
                    table発送モード.Rows.Add("発送先不明", -1)

                    Dim column As New DataGridViewComboBoxColumn()
                    column.DataPropertyName = "発送モード"
                    column.DataSource = table発送モード
                    column.ValueMember = "Value"
                    column.DisplayMember = "Display"

                    Me.Columns.Insert(Me.Columns("発送モード").Index, column)
                    Me.Columns.Remove("発送モード")
                    column.Name = "発送モード"
                    column.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
                End If


            End If
        End Sub

        Private Sub Change発送モード(ByVal s As Object, ByVal e As EventArgs)
            Select Case s.text
                Case "この筆のみ登記名義人を発送先に設定"
                    Dim pRow As DataRow = mvarParent.TBL非農地.Rows.Find(Me.Rows(Me.CurrentCell.RowIndex).Cells("CODE").Value)
                    ChangeList.Clear()
                    ChangeList.UpdateValue(pRow, "発送モード", 3)
                    ChangeList.UpdateValue(pRow, "発送先ID", pRow.Item(("登記名義人ID")))
                    ChangeList.UpdateValue(pRow, "発送先氏名", pRow.Item(("登記名義人氏名")))
                    ChangeList.UpdateValue(pRow, "発送先住所", pRow.Item(("登記名義人住所")))
                    ChangeList.UpdateValue(pRow, "発送先郵便番号", pRow.Item(("登記名義人郵便番号")))
                    ChangeList.SaveRow(mvarParent.GetLDB(), pRow)
                Case "この筆のみ納税義務者を発送先に設定"
                    Dim pRow As DataRow = mvarParent.TBL非農地.Rows.Find(Me.Rows(Me.CurrentCell.RowIndex).Cells("CODE").Value)
                    ChangeList.Clear()
                    ChangeList.UpdateValue(pRow, "発送モード", 3)
                    ChangeList.UpdateValue(pRow, "発送先ID", pRow.Item(("納税義務者ID")))
                    ChangeList.UpdateValue(pRow, "発送先氏名", pRow.Item(("納税義務者氏名")))
                    ChangeList.UpdateValue(pRow, "発送先住所", pRow.Item(("納税義務者住所")))
                    ChangeList.UpdateValue(pRow, "発送先郵便番号", pRow.Item(("納税義務者郵便番号")))
                    ChangeList.SaveRow(mvarParent.GetLDB(), pRow)
                Case "同一登記名義人の筆を発送先に設定"
                    Dim RX As Integer = Me.CurrentCell.RowIndex
                    Dim ID As Integer = Me.Rows(RX).Cells("登記名義人ID").Value
                    Dim Lst As New List(Of String)
                    Dim pRowX As DataRow = mvarParent.TBL非農地.Rows.Find(Me.Rows(Me.CurrentCell.RowIndex).Cells("CODE").Value)

                    For Each pDRow As DataGridViewRow In Me.Rows
                        If Not IsDBNull(pDRow.Cells("登記名義人ID").Value) AndAlso pDRow.Cells("登記名義人ID").Value = ID Then
                            Lst.Add(pDRow.Cells("CODE").Value)
                        End If
                    Next

                    With mvarParent.GetLDB()
                        For Each sCODE As String In Lst
                            Dim pRow As DataRow = mvarParent.TBL非農地.Rows.Find(sCODE)

                            ChangeList.Clear()
                            ChangeList.UpdateValue(pRow, "発送モード", 3)
                            ChangeList.UpdateValue(pRow, "発送先ID", pRowX.Item(("登記名義人ID")))
                            ChangeList.UpdateValue(pRow, "発送先氏名", pRowX.Item(("登記名義人氏名")))
                            ChangeList.UpdateValue(pRow, "発送先住所", pRowX.Item(("登記名義人住所")))
                            ChangeList.UpdateValue(pRow, "発送先郵便番号", pRowX.Item(("登記名義人郵便番号")))
                            ChangeList.SaveRow(mvarParent.GetLDB(), pRow)
                        Next
                    End With

                Case "同一納税義務者の筆を発送先に設定"
                    Dim RX As Integer = Me.CurrentCell.RowIndex
                    Dim ID As Integer = Me.Rows(RX).Cells("納税管理者ID").Value
                    Dim Lst As New List(Of String)
                    Dim pRowX As DataRow = mvarParent.TBL非農地.Rows.Find(Me.Rows(Me.CurrentCell.RowIndex).Cells("CODE").Value)

                    For Each pDRow As DataGridViewRow In Me.Rows
                        If Not IsDBNull(pDRow.Cells("納税管理者ID").Value) AndAlso pDRow.Cells("納税管理者ID").Value = ID Then
                            Lst.Add(pDRow.Cells("CODE").Value)
                        End If
                    Next

                    With mvarParent.GetLDB()
                        For Each sCODE As String In Lst
                            Dim pRow As DataRow = mvarParent.TBL非農地.Rows.Find(sCODE)

                            ChangeList.Clear()
                            ChangeList.UpdateValue(pRow, "発送モード", 3)
                            ChangeList.UpdateValue(pRow, "発送先ID", pRowX.Item(("納税管理者ID")))
                            ChangeList.UpdateValue(pRow, "発送先氏名", pRowX.Item(("納税管理者氏名")))
                            ChangeList.UpdateValue(pRow, "発送先住所", pRowX.Item(("納税管理者住所")))
                            ChangeList.UpdateValue(pRow, "発送先郵便番号", pRowX.Item(("納税管理者郵便番号")))
                            ChangeList.SaveRow(mvarParent.GetLDB(), pRow)
                        Next
                    End With
            End Select
        End Sub

        Private Sub mvarGridView_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles Me.CellEndEdit
            Dim pDRow As DataGridViewRow = Me.Rows(e.RowIndex)

            Dim pRow As DataRow = mvarParent.TBL非農地.Rows.Find(Me.Rows(e.RowIndex).Cells("CODE").Value)
            If pDRow.DataBoundItem Is Nothing Then
                pRow = mvarParent.TBL非農地.Rows.Find(Me.Rows(e.RowIndex).Cells("CODE").Value)
            ElseIf TypeOf pDRow.DataBoundItem Is DataRowView Then
                pRow = CType(pDRow.DataBoundItem, DataRowView).Row
            ElseIf TypeOf pDRow.DataBoundItem Is DataRow Then
                pRow = CType(pDRow.DataBoundItem, DataRowView).Row
            End If


            EditedRowIndex = e.RowIndex

            Select Case Me.Columns(e.ColumnIndex).Name
                Case "発送モード"
                    Select Case Me.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                        Case 0
                            ChangeList.UpdateValue(pRow, "発送モード", 0)
                            ChangeList.UpdateValue(pRow, "発送先ID", pRow.Item("登記名義人ID"))
                            ChangeList.UpdateValue(pRow, "発送先氏名", pRow.Item("登記名義人氏名"))
                            ChangeList.UpdateValue(pRow, "発送先住所", pRow.Item("登記名義人住所"))
                            ChangeList.UpdateValue(pRow, "発送先郵便番号", pRow.Item("登記名義人郵便番号"))
                        Case 1
                            ChangeList.UpdateValue(pRow, "発送モード", 1)
                            ChangeList.UpdateValue(pRow, "発送先ID", pRow.Item("現所有者ID"))
                            ChangeList.UpdateValue(pRow, "発送先氏名", pRow.Item("現所有者氏名"))
                            ChangeList.UpdateValue(pRow, "発送先住所", pRow.Item("現所有者住所"))
                            ChangeList.UpdateValue(pRow, "発送先郵便番号", pRow.Item("現所有者郵便番号"))
                        Case 2
                            ChangeList.UpdateValue(pRow, "発送モード", 2)
                            ChangeList.UpdateValue(pRow, "発送先ID", pRow.Item("納税管理者ID"))
                            ChangeList.UpdateValue(pRow, "発送先氏名", pRow.Item("納税管理者氏名"))
                            ChangeList.UpdateValue(pRow, "発送先住所", pRow.Item("納税管理者住所"))
                            ChangeList.UpdateValue(pRow, "発送先郵便番号", pRow.Item("納税管理者郵便番号"))

                        Case 3, 4, 5
                            ChangeList.UpdateValue(pRow, "発送モード", Me.Rows(e.RowIndex).Cells(e.ColumnIndex).Value)
                        Case -1
                            ChangeList.UpdateValue(pRow, "発送モード", Me.Rows(e.RowIndex).Cells("発送モード").Value)
                            ChangeList.UpdateValue(pRow, "発送先ID", DBNull.Value)
                            ChangeList.UpdateValue(pRow, "発送先氏名", DBNull.Value)
                            ChangeList.UpdateValue(pRow, "発送先住所", DBNull.Value)
                            ChangeList.UpdateValue(pRow, "発送先郵便番号", DBNull.Value)
                    End Select
                    '                    If Not ChangeList.Contains("発送モード") Then ChangeList.Add("発送モード")

                Case "発送先ID"
                    ChangeList.UpdateValue(pRow, "発送モード", 3)
                    ChangeList.UpdateValue(pRow, "発送先ID", pRow.Item("発送先ID"))
                Case "発送先氏名"
                    ChangeList.UpdateValue(pRow, "発送モード", 3)
                    ChangeList.UpdateValue(pRow, "発送先氏名", pRow.Item("発送先氏名"))
                Case "発送先住所"
                    ChangeList.UpdateValue(pRow, "発送モード", 3)
                    ChangeList.UpdateValue(pRow, "発送先住所", pRow.Item("発送先住所"))
                Case "発送先郵便番号"
                    ChangeList.UpdateValue(pRow, "発送モード", 3)
                    ChangeList.UpdateValue(pRow, "発送先郵便番号", pRow.Item("発送先郵便番号"))
            End Select
            ChangeList.SaveRow(mvarParent.GetLDB(), pRow)
        End Sub

        Private Sub mvarCHk発送済み非表示_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarCHk発送済み非表示.Click
            If mvarCHk発送済み非表示.Checked Then
                DataView.RowFilter = "[削除] Is Null AND [通知日] Is Null AND [発送番号] Is Null"
            Else
                DataView.RowFilter = "[削除] Is Null"
            End If
        End Sub

        Public bFirst As Boolean = True
        Public bLoading As Boolean = False

        Private Sub DataGridLocal_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SelectionChanged
            If EditedRowIndex > -1 AndAlso ChangeList.Count > 0 Then
                ChangeList.SaveRow(mvarParent.GetLDB(), mvarParent.TBL非農地.Rows.Find(Me.Rows(EditedRowIndex).Cells("CODE").Value))
                EditedRowIndex = -1
            End If
        End Sub

        Private Sub DataGridLocal_DataSourceChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.DataSourceChanged
            Dim Ar() As String = New String() {"決定日", "通知日", "発送番号", "発送モード", "発送先ID", "発送先氏名", "発送先住所", "発送先郵便番号", "解消分類", "解消年月日", "備考", "耕作作物", "調査現況", "調査備考", "樹齢"}
            For Each c As DataGridViewColumn In Me.Columns
                c.SortMode = DataGridViewColumnSortMode.NotSortable
                c.ReadOnly = Not Ar.Contains(c.DataPropertyName)
            Next
        End Sub


        Private pStartRow As Integer = 0
        Private pStartColumn As Integer = 0

        Private Sub sub決定日(ByVal s As Object, ByVal e As EventArgs)
            If Me.SelectedRows.Count > 0 Then
                Dim St As String = InputBox("決定日を入力してください", "非農地決定日設定", Now.Date.ToString.Replace(" 0:00:00", ""))
                If IsDate(St) Then
                    Dim Dt As DateTime = CDate(St)
                    With mvarParent.GetLDB()
                        For Each pRoX As DataGridViewRow In Me.SelectedRows
                            Dim sCODE As String = pRoX.Cells("CODE").Value
                            Dim pRow As DataRow = mvarParent.TBL非農地.Rows.Find(sCODE)
                            ChangeList.Clear()
                            pRow.Item("決定日") = Dt
                            ChangeList.UpdateValue(pRow, "決定日", Dt)

                            ChangeList.SaveRow(mvarParent.GetLDB(), pRow)
                        Next

                    End With

                End If
            End If
        End Sub


        Private Sub sub誤入()
            If Me.SelectedRows.Count = 0 Then

            ElseIf Me.SelectedRows.Count = 1 Then
                '    Dim pRow As DataRow = tbl荒廃農地.Rows.Find(Me.SelectedRows(0).Cells("LID").Value)
                '    If Not pRow Is Nothing Then
                '        Dim pInput誤入力訂正 As New CInput誤入力訂正
                '        Dim pDlg As New dlgInputMulti(pInput誤入力訂正, "誤入力訂正", "パラメータを入力してください")

                '        If pDlg.ShowDialog = DialogResult.OK Then
                '            Stop
                '        End If
                '    Else
                '        MsgBox("対象になる荒廃農地がありません。もとに戻せません。", MsgBoxStyle.Critical)
                '    End If
            End If
        End Sub

        Private Sub txt検索Btn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt検索Btn.Click
            If Len(txt検索.Text) Then

                If Me.CurrentCell IsNot Nothing Then
                    If pStartRow = Me.CurrentCell.RowIndex AndAlso pStartColumn = Me.CurrentCell.ColumnIndex Then
                        pStartColumn = (Me.CurrentCell.ColumnIndex + 1) Mod Me.Columns.Count
                    Else
                        pStartRow = Me.CurrentCell.RowIndex
                        pStartColumn = Me.CurrentCell.ColumnIndex
                    End If
                End If
                Dim pFRow As Integer = pStartRow

                If pStartColumn > 0 Then
                    For n As Integer = pStartColumn To Me.Columns.Count - 1
                        If Me.Item(n, pFRow).Visible AndAlso
                            Not IsDBNull(Me.Item(n, pFRow).Value) AndAlso Me.Item(n, pFRow).Value Like txt検索.Text Then
                            Me.CurrentCell = Me.Item(n, pFRow)
                            pStartRow = pFRow
                            pStartColumn = n
                            Exit Sub
                        End If
                    Next
                End If
                For l As Integer = 0 To Me.Rows.Count - 2
                    pFRow = (pFRow + 1) Mod Me.Rows.Count
                    For n As Integer = 0 To Me.Columns.Count - 1
                        If Me.Item(n, pFRow).Visible AndAlso
                            Not IsDBNull(Me.Item(n, pFRow).Value) AndAlso Me.Item(n, pFRow).Value Like txt検索.Text Then
                            Me.CurrentCell = Me.Item(n, pFRow)
                            pStartRow = pFRow
                            pStartColumn = n
                            Exit Sub
                        End If
                    Next
                Next

                MsgBox("検索した文字列は見つかりませんでした。", MsgBoxStyle.Information)
            End If
        End Sub

        Private Sub mvarID採番_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarID採番.Click
            Dim MinNo As Integer = -1000
            For Each pRow As DataRow In mvarParent.TBL非農地.Rows
                If Not IsDBNull(pRow.Item("発送先ID")) AndAlso pRow.Item("発送先ID") < MinNo Then
                    MinNo = pRow.Item("発送先ID") - 1
                End If
            Next

            Dim pView採番 As New DataView(mvarParent.TBL非農地, "[発送モード]=3 AND [発送先ID] IS NULL AND [発送先氏名] IS NOT NULL AND [発送先住所] Is Not Null AND [発送先郵便番号] Is Not Null", "", DataViewRowState.CurrentRows)

            For Each pViewRow As DataRowView In pView採番
                If IsDBNull(pViewRow.Item("発送先ID")) AndAlso pViewRow.Item("発送先氏名").ToString.Length > 0 AndAlso pViewRow.Item("発送先住所").ToString.Length > 0 AndAlso pViewRow.Item("発送先郵便番号").ToString.Length > 0 Then

                    Dim pView既存 As New DataView(mvarParent.TBL非農地,
                            String.Format("[発送モード]=3 AND [発送先ID] IS Not NULL AND [発送先氏名]='{0}' AND [発送先住所]='{1}' AND [発送先郵便番号]='{2}'", pViewRow.Item("発送先氏名"), pViewRow.Item("発送先住所"), pViewRow.Item("発送先郵便番号")),
                            "", DataViewRowState.CurrentRows)

                    If pView既存.Count > 0 Then
                        Dim XNO As Integer = pView既存(0).Item("発送先ID")
                        Dim bOK As Boolean = True
                        For Each pRow既存 As DataRowView In pView既存
                            bOK = bOK And (pRow既存.Item("発送先ID") = XNO)
                        Next
                        If bOK Then

                            Dim pView同選択 As New DataView(mvarParent.TBL非農地,
                                    String.Format("[発送モード]=3 AND [発送先ID] IS NULL AND [発送先氏名]='{0}' AND [発送先住所]='{1}' AND [発送先郵便番号]='{2}'", pViewRow.Item("発送先氏名"), pViewRow.Item("発送先住所"), pViewRow.Item("発送先郵便番号")),
                                    "", DataViewRowState.CurrentRows)
                            If pView同選択.Count > 0 Then

                                For Each pSetRow As DataRowView In pView同選択
                                    pSetRow.Item("発送先ID") = XNO
                                Next
                            End If
                        End If
                    End If


                    Dim pView選択 As New DataView(mvarParent.TBL非農地,
                            String.Format("[発送モード]=3 AND [発送先ID] IS NULL AND [発送先氏名]='{0}' AND [発送先住所]='{1}' AND [発送先郵便番号]='{2}'", pViewRow.Item("発送先氏名"), pViewRow.Item("発送先住所"), pViewRow.Item("発送先郵便番号")),
                            "", DataViewRowState.CurrentRows)
                    If pView選択.Count > 0 Then

                        For Each pSetRow As DataRowView In pView選択
                            pSetRow.Item("発送先ID") = MinNo
                        Next
                        MinNo -= 1
                    End If
                End If
            Next

        End Sub

        Private Sub DataGridLocal_CellErrorTextNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellErrorTextNeededEventArgs) Handles Me.CellErrorTextNeeded
            If e.RowIndex > -1 AndAlso e.ColumnIndex > -1 Then
                If InStr(Me.Item(e.ColumnIndex, e.RowIndex).Value.ToString, "XXX-") Then
                    e.ErrorText = "仮番号"
                End If
            End If
        End Sub

        Private Sub mvar発送先空欄を埋める_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar発送先空欄を埋める.Click
            With mvarParent.GetLDB()

                For Each pRowV As DataRowView In CType(Me.DataSource, DataView)
                    If IsDBNull(pRowV.Item("発送先ID")) Then
                        Select Case pRowV.Item("登記名義人区分")
                            Case 0, 5, 31, 33, 36, 37
                                ChangeList.UpdateValue(pRowV.Row, "発送モード", 0)
                                ChangeList.UpdateValue(pRowV.Row, "発送先ID", pRowV.Item("登記名義人ID"))
                                ChangeList.UpdateValue(pRowV.Row, "発送先氏名", pRowV.Item("登記名義人氏名"))
                                ChangeList.UpdateValue(pRowV.Row, "発送先住所", pRowV.Item("登記名義人住所"))
                                ChangeList.UpdateValue(pRowV.Row, "発送先郵便番号", pRowV.Item("登記名義人郵便番号"))
                            Case 6, 32, 34, 35

                                Select Case pRowV.Item("現所有者区分")
                                    Case 0, 5, 31, 33, 36, 37
                                        ChangeList.UpdateValue(pRowV.Row, "発送モード", 1)
                                        ChangeList.UpdateValue(pRowV.Row, "発送先ID", pRowV.Item("現所有者ID"))
                                        ChangeList.UpdateValue(pRowV.Row, "発送先氏名", pRowV.Item("現所有者氏名"))
                                        ChangeList.UpdateValue(pRowV.Row, "発送先住所", pRowV.Item("現所有者住所"))
                                        ChangeList.UpdateValue(pRowV.Row, "発送先郵便番号", pRowV.Item("現所有者郵便番号"))
                                    Case 6, 32, 34, 35
                                        If IsDBNull(pRowV.Item("納税管理者区分")) Then pRowV.Item("納税管理者区分") = -1
                                        Select Case pRowV.Item("納税管理者区分")
                                            Case 0, 5, 31, 33, 36, 37
                                                ChangeList.UpdateValue(pRowV.Row, "発送モード", 2)
                                                ChangeList.UpdateValue(pRowV.Row, "発送先ID", pRowV.Item("納税管理者ID"))
                                                ChangeList.UpdateValue(pRowV.Row, "発送先氏名", pRowV.Item("納税管理者氏名"))
                                                ChangeList.UpdateValue(pRowV.Row, "発送先住所", pRowV.Item("納税管理者住所"))
                                                ChangeList.UpdateValue(pRowV.Row, "発送先郵便番号", pRowV.Item("納税管理者郵便番号"))
                                            Case Else
                                                ChangeList.UpdateValue(pRowV.Row, "発送モード", 3)
                                                ChangeList.UpdateValue(pRowV.Row, "発送先ID", pRowV.Item(("登記名義人ID")))
                                                ChangeList.UpdateValue(pRowV.Row, "発送先氏名", pRowV.Item(("登記名義人氏名")))
                                                ChangeList.UpdateValue(pRowV.Row, "発送先住所", pRowV.Item(("登記名義人住所")))
                                                ChangeList.UpdateValue(pRowV.Row, "発送先郵便番号", pRowV.Item(("登記名義人郵便番号")))
                                        End Select
                                    Case Else
                                        Stop
                                End Select
                            Case Else
                                Stop
                        End Select
                    End If
                    ChangeList.SaveRow(mvarParent.GetLDB(), pRowV.Row)
                Next

            End With
            MsgBox("保存しました。")
        End Sub



        Private Sub DataGridLocal_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles Me.ColumnHeaderMouseClick
            If DataView.Sort = "[" & Me.Columns(e.ColumnIndex).DataPropertyName & "] ASC" Then
                DataView.Sort = "[" & Me.Columns(e.ColumnIndex).DataPropertyName & "] DESC"
            ElseIf DataView.Sort = "[" & Me.Columns(e.ColumnIndex).DataPropertyName & "] DESC" Then
                DataView.Sort = "[" & Me.Columns(e.ColumnIndex).DataPropertyName & "] ASC"
            ElseIf DataView.Sort = Me.Columns(e.ColumnIndex).DataPropertyName Then
                DataView.Sort = "[" & Me.Columns(e.ColumnIndex).DataPropertyName & "] ASC"
            Else
                DataView.Sort = Me.Columns(e.ColumnIndex).DataPropertyName
            End If
        End Sub

        Private Sub DataGridLocal_CellContextMenuStripNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellContextMenuStripNeededEventArgs) Handles Me.CellContextMenuStripNeeded
            If e.ColumnIndex = -1 Then
                If e.RowIndex > -1 AndAlso Me.SelectedRows.Count = 1 Then
                    非農地ContextMenu.SetEnable(
                        Not IsDBNull(Me.Rows(e.RowIndex).Cells("通知日").Value)
                    )
                    e.ContextMenuStrip = 非農地ContextMenu
                ElseIf Me.SelectedRows.Count > 1 Then
                    e.ContextMenuStrip = 非農地MultiContextMenu
                End If

            ElseIf e.RowIndex = -1 AndAlso e.ColumnIndex > -1 Then

            Else
                Select Case Me.Columns(e.ColumnIndex).Name
                    Case "登記名義人ID", "登記名義人氏名", "登記名義人住所", "登記名義人郵便番号"
                        e.ContextMenuStrip = 登記名義ContextMenu
                    Case "納税管理者ID", "納税管理者氏名", "納税管理者住所", "納税管理者郵便番号"
                        e.ContextMenuStrip = 納税義務者ContextMenu
                    Case Else
                        If e.RowIndex > -1 AndAlso Not IsDBNull(Me.Rows(e.RowIndex).Cells("通知日").Value) AndAlso Me.SelectedRows.Count = 1 Then
                            e.ContextMenuStrip = 非農地ContextMenu
                        ElseIf Me.SelectedRows.Count > 1 Then
                            e.ContextMenuStrip = 非農地MultiContextMenu
                        ElseIf Me.SelectedCells.Count = 1 Then
                            e.ContextMenuStrip = DefaultMenu
                        End If
                End Select
            End If
        End Sub

        Public Class 非農地RowMenu
            Inherits ContextMenuStrip

            Public 通知済み解除 As ToolStripMenuItem
            Public 発送済み解除 As ToolStripMenuItem
            Public 誤入力解除 As ToolStripMenuItem
            Public 非農地リスト削除 As ToolStripMenuItem

            Public Sub New()
                通知済み解除 = Items.Add("通知済みを解除する")
                発送済み解除 = Items.Add("同発送先非農地を解除する")
                '誤入力解除 = Items.Add("誤入力をもどす")
                Me.Items.Add(New ToolStripSeparator)
                非農地リスト削除 = Items.Add("非農地リストから削除する")
            End Sub
            Public Sub SetEnable(b通知済み解除 As Boolean)
                通知済み解除.Enabled = b通知済み解除
                発送済み解除.Enabled = b通知済み解除

            End Sub
        End Class

    End Class

    Private Class TreeViewLocal
        Inherits TreeView
        Public TreeTools As ToolStrip
        Private WithEvents BtnAllClear As ToolStripButton
        Private WithEvents BtnAllSelect As ToolStripButton
        Private WithEvents mvar宛名 As ToolStripDropDownButton
        Private mvarParent As classPage非農地通知
        Private WithEvents mvar印刷 As ToolStripButton

        Public mvar通知書Path As String = ""
        Public mvar通知発送文書Path As String = ""

        Private mvarPrintID As New List(Of Integer)

        Public TreeCont As ToolStripContainer

        Public Sub New(ByVal pParent As classPage非農地通知)
            mvarParent = pParent

            Me.TreeTools = New ToolStrip
            Me.TreeTools.GripStyle = False
            Me.TreeTools.AutoSize = False
            Me.TreeTools.Stretch = True

            Me.TreeTools.ImageScalingSize = New System.Drawing.Size(32, 32)


            mvar宛名 = New ToolStripDropDownButton
            mvar宛名.Text = "宛名作成"
            mvar宛名.Image = My.Resources.Resource1._112_DownArrowShort_Blue.ToBitmap
            AddHandler CType(mvar宛名.DropDownItems.Add("全ての農家で作成"), ToolStripDropDownItem).Click, AddressOf mvar宛名_Click
            AddHandler CType(mvar宛名.DropDownItems.Add("選択筆のみ作成"), ToolStripDropDownItem).Click, AddressOf mvar宛名_Click
            mvar宛名.AutoSize = True

            Me.TreeTools.Items.Add(mvar宛名)

            mvar印刷 = New ToolStripButton
            mvar印刷.Text = "印刷開始"
            mvar印刷.Image = SysAD.ImageList48.Images("printer")

            Me.TreeTools.Items.Add(mvar印刷)
            Dim mvar印刷Sub = New ToolStripDropDownButton
            mvar印刷Sub.Text = ""
            Me.TreeTools.Items.Add(mvar印刷Sub)
            AddHandler mvar印刷Sub.DropDownItems.Add("通知書ファイルを指定する").Click, AddressOf Set通知書Path
            AddHandler mvar印刷Sub.DropDownItems.Add("通知発送文書ファイルを指定する").Click, AddressOf Set通知発送文書Path



            Me.TreeTools.Items.Add(New ToolStripSeparator)

            BtnAllSelect = New ToolStripButton("全て選択")
            BtnAllSelect.Image = My.Resources.Resource1.SuccessComplete
            BtnAllSelect.ImageTransparentColor = Color.Magenta
            TreeTools.Items.Add(BtnAllSelect)

            BtnAllClear = New ToolStripButton("選択を解除")
            BtnAllClear.Image = My.Resources.Resource1.UnCheck
            BtnAllClear.ImageTransparentColor = Color.Magenta
            TreeTools.Items.Add(BtnAllClear)


            Me.CheckBoxes = True
            Me.Dock = DockStyle.Fill
            Me.TreeCont = New ToolStripContainer
            Me.TreeCont.Dock = DockStyle.Fill
            Me.TreeCont.TopToolStripPanel.Controls.Add(Me.TreeTools)
            Me.TreeCont.ContentPanel.Controls.Add(Me)

        End Sub

        Private Sub Set通知書Path(ByVal s As Object, ByVal e As EventArgs)
            With New OpenFileDialog
                .Filter = "非農地通知書.xml|非農地通知書.xml"
                If IO.File.Exists(mvar通知書Path) Then
                    .FileName = mvar通知書Path
                End If

                If .ShowDialog = DialogResult.OK Then
                    mvar通知書Path = .FileName
                    SaveSetting("農地基本台帳", "非農地通知", "非農地通知書Path", .FileName)
                End If

            End With
        End Sub

        Private Sub Set通知発送文書Path(ByVal s As Object, ByVal e As EventArgs)
            With New OpenFileDialog
                .Filter = "非農地通知発送文書.xml|非農地通知発送文書.xml"
                If IO.File.Exists(mvar通知発送文書Path) Then
                    .FileName = mvar通知発送文書Path
                End If

                If .ShowDialog = DialogResult.OK Then
                    mvar通知発送文書Path = .FileName
                    SaveSetting("農地基本台帳", "非農地通知", "非農地通知発送文書Path", .FileName)
                End If

            End With
        End Sub

        Private Sub mvar宛名_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar宛名.Click
            Dim sMode As String = ""
            If TypeOf sender Is ToolStripDropDownItem Then
                Select Case CType(sender, ToolStripDropDownItem).Text
                    Case "宛名作成"
                        Return
                    Case "選択筆のみ作成" : sMode = "選択筆のみ作成"
                    Case "全ての農家で作成" : sMode = "全ての農家で作成"
                    Case Else
                        Stop
                End Select

            End If
            If mvarParent.TBL非農地 IsNot Nothing AndAlso mvarParent.TBL非農地.Rows.Count > 0 Then
                mvarParent.Tbl宛名 = New DataTable
                With mvarParent.Tbl宛名
                    .Columns.Clear()
                    .Columns.Add(New DataColumn("ID", GetType(Integer)))
                    .Columns.Add(New DataColumn("氏名", GetType(String)))
                    .Columns.Add(New DataColumn("住所", GetType(String)))
                    .Columns.Add(New DataColumn("郵便番号", GetType(String)))
                    .Columns.Add(New DataColumn("通知文書情報", GetType(String)))

                    .PrimaryKey = New DataColumn() { .Columns("ID")}

                    Nodes.Clear()
                    Nodes.Add("姶良", "姶良")
                    Nodes.Add("加治木", "加治木")
                    Nodes.Add("蒲生", "蒲生")
                    Nodes.Add("市外", "市外")

                    Select Case sMode
                        Case "選択筆のみ作成"
                            MakeFromSelectData()
                        Case "全ての農家で作成"
                            MakeFromAllLandData()
                    End Select


                End With

            End If
        End Sub

        Private Sub MakeFromAllLandData()
            With mvarParent.Tbl宛名
                Dim pViewSort As New DataView(mvarParent.TBL非農地, "", "[発送先住所],[発送先氏名]", DataViewRowState.CurrentRows)
                For Each pRow As DataRowView In pViewSort

                    If pRow.Item("発送モード") < 4 AndAlso Not IsDBNull(pRow.Item("発送先ID")) AndAlso
                        Not IsDBNull(pRow.Item("発送先氏名")) AndAlso
                        Not IsDBNull(pRow.Item("発送先住所")) AndAlso
                        (IsDBNull(pRow.Item("通知日")) OrElse pRow.Item("通知日").ToString.Length = 0) AndAlso
                        (IsDBNull(pRow.Item("発送番号")) OrElse pRow.Item("発送番号").ToString.Length = 0) AndAlso
                            .Rows.Find(pRow.Item("発送先ID")) Is Nothing Then

                        Dim pNewRow As DataRow = .NewRow
                        pNewRow.Item("ID") = pRow.Item("発送先ID")
                        pNewRow.Item("氏名") = pRow.Item("発送先氏名")
                        pNewRow.Item("住所") = pRow.Item("発送先住所")
                        pNewRow.Item("郵便番号") = pRow.Item("発送先郵便番号")

                        .Rows.Add(pNewRow)
                        Dim sText As String = pRow.Item("発送先氏名").ToString & "(" & pRow.Item("発送先住所").ToString & ")"
                        Dim pNode As New TreeNode

                        If InStr(pRow.Item("発送先住所"), "加治木") Then
                            pNode = Nodes("加治木").Nodes.Add(sText)
                        ElseIf InStr(pRow.Item("発送先住所"), "蒲生") Then
                            pNode = Nodes("蒲生").Nodes.Add(sText)
                        ElseIf InStr(pRow.Item("発送先住所"), "姶良") Then
                            pNode = Nodes("姶良").Nodes.Add(sText)
                        Else
                            pNode = Nodes("市外").Nodes.Add(sText)
                        End If

                        pNode.Tag = pNewRow

                    End If
                Next
            End With
        End Sub

        Private Sub MakeFromSelectData()
            With mvarParent.Tbl宛名
                For Each pGRow As DataGridViewRow In mvarParent.mvarGridView.SelectedRows
                    Dim pRow As DataRow = mvarParent.TBL非農地.Rows.Find(pGRow.Cells("CODE").Value)


                    If pRow.Item("発送モード") < 4 AndAlso Not IsDBNull(pRow.Item("発送先ID")) AndAlso
                        Not IsDBNull(pRow.Item("発送先氏名")) AndAlso
                        Not IsDBNull(pRow.Item("発送先住所")) AndAlso
                                          .Rows.Find(pRow.Item("発送先ID")) Is Nothing Then

                        Dim pNewRow As DataRow = .NewRow
                        pNewRow.Item("ID") = pRow.Item("発送先ID")
                        pNewRow.Item("氏名") = pRow.Item("発送先氏名")
                        pNewRow.Item("住所") = pRow.Item("発送先住所")
                        pNewRow.Item("郵便番号") = pRow.Item("発送先郵便番号")

                        .Rows.Add(pNewRow)
                        Dim sText As String = pRow.Item("発送先氏名").ToString & "(" & pRow.Item("発送先住所").ToString & ")"
                        Dim pNode As New TreeNode

                        If InStr(pRow.Item("発送先住所"), "加治木") Then
                            pNode = Nodes("加治木").Nodes.Add(sText)
                        ElseIf InStr(pRow.Item("発送先住所"), "蒲生") Then
                            pNode = Nodes("蒲生").Nodes.Add(sText)
                        ElseIf InStr(pRow.Item("発送先住所"), "姶良") Then
                            pNode = Nodes("姶良").Nodes.Add(sText)
                        Else
                            pNode = Nodes("市外").Nodes.Add(sText)
                        End If

                        pNode.Tag = pNewRow

                    End If
                Next
            End With
        End Sub

        Private Sub TreeViewLocal_AfterCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles Me.AfterCheck
            For Each pNode As TreeNode In e.Node.Nodes
                pNode.Checked = e.Node.Checked
            Next
        End Sub

        Private Sub TreeViewLocal_ParentChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ParentChanged
            If TypeName(Me.Parent) = "ToolStripContentPanel" Then

            Else
                Stop
            End If
        End Sub

        Private Sub BtnAllClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnAllClear.Click
            SetCheck(Me.Nodes, False)
        End Sub
        Private Sub BtnAllSelect_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnAllSelect.Click
            SetCheck(Me.Nodes, True)
        End Sub
        Private Sub SetCheck(ByVal pNodes As TreeNodeCollection, ByVal b As Boolean)
            For Each pNode As TreeNode In pNodes
                If pNode.Nodes.Count > 0 Then
                    SetCheck(pNode.Nodes, b)
                End If
                pNode.Checked = b
            Next
        End Sub


        Private b仮印刷 As Boolean = False
        Private n元番号 As Integer = 0
        Private Sub 通知書印刷(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar印刷.Click
            Dim sFileName As String = ""
            Dim sFileName2 As String = ""
            If IO.File.Exists(mvar通知書Path) Then
                sFileName = mvar通知書Path
            Else
                sFileName = My.Application.Info.DirectoryPath & "\非農地通知書.xml"
            End If

            If IO.File.Exists(mvar通知発送文書Path) Then
                sFileName2 = mvar通知発送文書Path
            Else
                sFileName2 = My.Application.Info.DirectoryPath & "\非農地通知発送文書.xml"
            End If
            Dim sWriteFile As String = SysAD.OutputFolder & "\非農地通知書.xml"
            Dim sWriteFile2 As String = SysAD.OutputFolder & "\非農地通知発送文書.xml"

            b仮印刷 = False
            Select Case MsgBox("本印刷を行いますか", MsgBoxStyle.YesNoCancel)
                Case MsgBoxResult.Yes
                    b仮印刷 = False
                Case MsgBoxResult.No
                    b仮印刷 = True
                Case MsgBoxResult.Cancel
                    Exit Sub
            End Select
            mvarPrintID.Clear()

            If Not IO.File.Exists(sFileName) AndAlso Not IO.File.Exists(sFileName2) Then
                MsgBox("ファイルが見つかりません", vbCritical)
            ElseIf Not IsDate(mvarParent.mvar決定総会日.Text) Then
                MsgBox("総会日を入力してください", vbCritical)
            ElseIf Not IsDate(mvarParent.mvar通知日.Text) Then
                MsgBox("通知日を入力してください", vbCritical)
            ElseIf Val(mvarParent.mvar通知番号開始.Text) = 0 Then
                MsgBox("通知番号の開始Noを入力してください", vbCritical)
            ElseIf Val(mvarParent.mvar発行番号.Text) = 0 Then
                MsgBox("発行番号を入力してください", vbCritical)
            Else

                Dim pExcel As New HimTools2012.Excel.Automation.ExcelAutomation
                Dim sXML As String = HimTools2012.TextAdapter.LoadTextFile(sFileName)

                If Val(mvarParent.mvar通知番号開始.Text) = 0 Then
                    mvarParent.mvar通知番号開始.Text = "1"
                End If
                n元番号 = Val(mvarParent.mvar通知番号開始.Text)

                With mvarParent.GetLDB()
                    PrintSub(mvarParent.GetLDB(), Nodes, sXML, sWriteFile, pExcel)

                    If mvarPrintID.Count > 0 Then
                        Dim sXML2 As String = HimTools2012.TextAdapter.LoadTextFile(sFileName2)

                        For Each nID As Integer In mvarPrintID

                            With mvarParent.Tbl宛名.Rows.Find(nID)
                                Dim s As String = .Item("通知文書情報")
                                If s.Length > 0 Then
                                    For Each St As String In Split(s, ";")
                                        Dim Ar() As String = Split(St, "＄")

                                        Dim sXml3 As String = Replace(sXML2, "{発行番号}", mvarParent.mvar発行番号.Text)
                                        Dim BBan As String
                                        If .Item("郵便番号").ToString.Length = 7 Then
                                            BBan = .Item("郵便番号").ToString.Substring(0, 3) & "-" & .Item("郵便番号").ToString.Substring(3, 4)
                                        Else
                                            If InStr(.Item("郵便番号").ToString, "-") > 0 Then
                                                BBan = .Item("郵便番号").ToString
                                            Else
                                                BBan = "   -    "
                                            End If
                                        End If

                                        sXml3 = Replace(sXml3, "{郵便番号}", BBan)
                                        sXml3 = Replace(sXml3, "{発行年月日}", 和暦Format(mvarParent.mvar通知日.Text))
                                        sXml3 = Replace(sXml3, "{住所}", .Item("住所"))
                                        sXml3 = Replace(sXml3, "{氏名}", .Item("氏名") & " 様")

                                        If Ar(0) = .Item("氏名") Then
                                            sXml3 = Replace(sXml3, "{様分}", "")
                                        Else
                                            sXml3 = Replace(sXml3, "{様分}", "( " & Ar(0) & " 様分 )")
                                        End If

                                        sXml3 = Replace(sXml3, "{通知番号}", Ar(1))

                                        If Not b仮印刷 Then
                                            HimTools2012.TextAdapter.SaveTextFile(sWriteFile2, sXml3)
                                            pExcel.PrintBook(sWriteFile2)
                                            'pExcel.PrintFile(sWriteFile2, 1, 1)
                                            IO.File.Delete(sWriteFile2)
                                        End If
                                    Next

                                End If
                            End With
                        Next
                    End If
                End With

                If b仮印刷 Then
                    mvarParent.mvar通知番号開始.Text = n元番号
                End If

                pExcel.Close()
            End If

        End Sub

        Private Sub PrintSub(ByVal pDB As HimTools2012.Data.CLocalDataEngine, ByVal pNodeP As TreeNodeCollection, ByVal sXMLA As String, ByVal sWriteFile As String, ByVal pExcel As HimTools2012.Excel.Automation.ExcelAutomation)
            For Each pNode As TreeNode In pNodeP

                If pNode.Nodes.Count > 0 Then
                    PrintSub(pDB, pNode.Nodes, sXMLA, sWriteFile, pExcel)
                End If
                '
                If pNode.Checked AndAlso pNode.Tag IsNot Nothing Then
                    Dim pRow As DataRow = CType(pNode.Tag, DataRow)

                    mvarPrintID.Add(pRow.Item("ID"))

                    Dim pView名義 As New DataView(mvarParent.TBL非農地, "[発送先ID]=" & pRow.Item("ID"), "登記名義人ID", DataViewRowState.CurrentRows)

                    Dim query = From mvar登記名義 As DataRowView In pView名義 Group By 名義ID = mvar登記名義.Item("登記名義人ID"), s氏名 = mvar登記名義.Item("登記名義人氏名") Into g = Group
                                Select New With
                            {
                                .ID = 名義ID,
                                .氏名 = s氏名
                            }


                    Dim sSemiC As String = ""
                    pRow.Item("通知文書情報") = ""
                    For Each p登記名義 In query
                        Dim sXML As String = sXMLA
                        sXML = Replace(sXML, "{氏名}", p登記名義.氏名 & " 　様")
                        Dim pView As New DataView(mvarParent.TBL非農地, "[発送モード]>=0 AND [発送モード]<4 AND [登記名義人ID]=" & p登記名義.ID, "大字ID,本番,枝番", DataViewRowState.CurrentRows)

                        If pView.Count > 0 Then

                            sXML = Replace(sXML, "{発行番号}", mvarParent.mvar発行番号.Text)
                            sXML = Replace(sXML, "{発行年月日}", 和暦Format(mvarParent.mvar通知日.Text))
                            sXML = Replace(sXML, "{決定総会年月日}", 和暦Format(mvarParent.mvar決定総会日.Text))
                            sXML = Replace(sXML, "{通知番号}", Val(mvarParent.mvar通知番号開始.Text))

                            Dim nCount As Integer = 0
                            Dim nArea As Decimal = 0

                            pRow.Item("通知文書情報") &= sSemiC & p登記名義.氏名 & "＄" & Val(mvarParent.mvar通知番号開始.Text)
                            sSemiC = ";"
                            For i = 0 To 50
                                If pView.Count > i Then
                                    nCount += 1
                                    Dim ChangeList As New UpdateList("CODE")
                                    Dim pLRow As DataRowView = pView.Item(i)

                                    ChangeList.UpdateValue(pLRow.Row, "通知日", mvarParent.mvar通知日.Value)

                                    If b仮印刷 AndAlso pRow.Item("通知文書情報").ToString.Length > 0 Then
                                        pLRow.Row.Item("発送番号") = "XXX-" & Strings.Right("0000" & mvarParent.mvar通知番号開始.Text, 4)
                                    ElseIf pRow.Item("通知文書情報").ToString.Length > 0 Then
                                        ChangeList.UpdateValue(pLRow.Row, "発送番号", mvarParent.mvar発行番号.Text & "-" & Strings.Right("0000" & mvarParent.mvar通知番号開始.Text, 4))
                                        ChangeList.SaveRow(pDB, pLRow.Row)
                                    End If

                                    sXML = Replace(sXML, "{" & String.Format("大字{0:D2}", (i + 1)) & "}", pLRow.Item("大字"))
                                    sXML = Replace(sXML, "{" & String.Format("小字{0:D2}", (i + 1)) & "}", pLRow.Item("小字"))
                                    sXML = Replace(sXML, "{" & String.Format("地番{0:D2}", (i + 1)) & "}", pLRow.Item("地番"))
                                    sXML = Replace(sXML, "{" & String.Format("地目{0:D2}", (i + 1)) & "}", pLRow.Item("地目"))
                                    sXML = Replace(sXML, "{" & String.Format("面積{0:D2}", (i + 1)) & "}", String.Format("{0:#,##0}", pLRow.Item("面積")).Replace(".00", ""))
                                    nArea += pLRow.Item("面積")
                                Else
                                    sXML = Replace(sXML, "{" & String.Format("大字{0:D2}", (i + 1)) & "}", "")
                                    sXML = Replace(sXML, "{" & String.Format("小字{0:D2}", (i + 1)) & "}", "")
                                    sXML = Replace(sXML, "{" & String.Format("地番{0:D2}", (i + 1)) & "}", "")
                                    sXML = Replace(sXML, "{" & String.Format("地目{0:D2}", (i + 1)) & "}", "")
                                    sXML = Replace(sXML, "{" & String.Format("面積{0:D2}", (i + 1)) & "}", "")
                                End If
                            Next

                            sXML = Replace(sXML, "{件数}", nCount)
                            sXML = Replace(sXML, "{面積計}", String.Format("{0:#,##0}", nArea).Replace(".00", ""))

                            mvarParent.mvar通知番号開始.Text = Val(mvarParent.mvar通知番号開始.Text) + 1
                            If Not b仮印刷 Then
                                HimTools2012.TextAdapter.SaveTextFile(sWriteFile, sXML)
                                'pExcel.PrintFile(sWriteFile, 1, 1)
                                pExcel.PrintBook(sWriteFile)

                                mvarParent.mvar通知番号開始.SaveReg()
                                IO.File.Delete(sWriteFile)
                            End If
                        End If
                    Next
                End If

            Next

        End Sub

    End Class

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            If MsgBox("閉じてもよろしいですか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                mvarGridView.DataSource = Nothing
                Return HimTools2012.controls.CloseMode.CloseOK
            Else
                Return HimTools2012.controls.CloseMode.CancelClose
            End If
        End Get
    End Property
End Class

Public Class ToolStripDateTimePickerWithlabel
    Inherits HimTools2012.controls.ToolStripDateTimePicker
    Private WithEvents mvarLabel As ToolStripLabel
    Private mvarTRegKey As String = ""

    Public Sub New(ByVal sCaption As String, Optional ByVal sRegKey As String = "")
        mvarLabel = New ToolStripLabel
        mvarLabel.Text = sCaption
        If Split(sRegKey, ";").Length = 3 Then
            mvarTRegKey = sRegKey
            Dim Ar() As String = Split(sRegKey, ";")
            Dim St As String = GetSetting(Ar(0), Ar(1), Ar(2), "")
            If IsDate(St) Then
                Me.Value = CDate(St)
            Else
                Me.Value = Me.DateTimePicker.MinDate
            End If
        End If
    End Sub

    Private Sub mvarLabel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarLabel.Click
        Me.Focus()
    End Sub
    Private Sub SOwnerChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.OwnerChanged
        If Me.Owner IsNot Nothing Then
            Me.Owner.Items.Add(New ToolStripSeparator)
            Me.Owner.Items.Add(mvarLabel)
        End If
    End Sub

    Private Sub SValidated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Validated
        SaveReg()
    End Sub
    Public Sub SaveReg()
        If mvarTRegKey.Length > 0 Then
            Dim Ar() As String = Split(mvarTRegKey, ";")
            If Me.Value <> Me.DateTimePicker.MinDate Then
                SaveSetting(Ar(0), Ar(1), Ar(2), Me.Value.ToString)
            Else
                SaveSetting(Ar(0), Ar(1), Ar(2), "")
            End If
        End If
    End Sub
End Class

Public Class ToolStripTextBoxWithLabel
    Inherits ToolStripTextBox

    Private WithEvents mvarLabel As ToolStripLabel
    Private mvarTRegKey As String = ""
    Public WarningMessage As String = ""

    Public Sub New(ByVal sCaption As String, Optional ByVal sRegKey As String = "")
        mvarLabel = New ToolStripLabel
        mvarLabel.Text = sCaption

        If Split(sRegKey, ";").Length = 3 Then
            mvarTRegKey = sRegKey
            Dim Ar() As String = Split(sRegKey, ";")
            Me.Text = GetSetting(Ar(0), Ar(1), Ar(2), "")
        End If
    End Sub

    Private Sub ToolStripTextBoxWithLabel_OwnerChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.OwnerChanged
        Me.Owner.Items.Add(New ToolStripSeparator)
        Me.Owner.Items.Add(mvarLabel)
    End Sub

    Private Sub mvarLabel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarLabel.Click
        Me.Focus()
    End Sub

    Private Sub ToolStripTextBoxWithLabel_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.TextChanged
        If Me.Text.Length = 0 AndAlso WarningMessage.Length > 0 Then
            mvarLabel.Image = My.Resources.Resource1.Warning
            mvarLabel.ToolTipText = WarningMessage
        Else
            mvarLabel.Image = Nothing
            mvarLabel.ToolTipText = ""
        End If
    End Sub

    Private Sub ToolStripTextBoxWithLabel_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Validated
        SaveReg()
    End Sub

    Public Sub SaveReg()
        If mvarTRegKey.Length > 0 Then
            Dim Ar() As String = Split(mvarTRegKey, ";")
            SaveSetting(Ar(0), Ar(1), Ar(2), Me.Text)
        End If
    End Sub
End Class


Public Class CInput誤入力訂正
    Inherits InputObjectParam

    Public Enum enum耕作作物
        なし = 0
        水稲 = 1
        飼料用米 = 2
        加工用米 = 3
        自己保全管理 = 4
        野菜 = 50
        有機野菜 = 51
        大根 = 52
        ネギ = 53
        レタス = 54
        キャベツ = 55
        ハクサイ = 56
        たまねぎ = 57
        なす = 58
        にんじん = 59
        トマト = 60
        ピーマン = 61
        ほうれんそう = 62
        にがうり = 63
        アスパラ = 64
        ミョウガ = 65
        ウコン = 66
        ｷｭｳﾘ = 67
        しょうが = 68
        ﾆﾝﾆｸ = 69
        ｵｸﾗ = 70
        たかな = 71
        麦 = 100
        ソバ = 101
        ゴマ = 102
        とうもろこし = 103
        雑穀 = 104
        甘藷 = 150
        かぼちゃ = 151
        さといも = 152
        こんにゃく芋 = 153
        大豆 = 154
        小豆 = 155
        ナタマメ = 156
        その他豆類 = 157
        ｼﾞｬｶﾞｲﾓ = 158
        山芋 = 159
        その他芋類 = 160
        苺_ﾊｳｽ含 = 200
        スイカ = 201
        メロン = 202
        なたね = 250
        キク = 251
        コスモス = 252
        花卉 = 253
        花木 = 254
        飼料 = 300
        ソルガム = 301
        イタリアン = 302
        芝 = 350
        たばこ = 351
        採草地 = 352
        樹園地_茶 = 400
        樹園地_ﾐｶﾝ = 401
        樹園地_柿 = 402
        樹園地_梨 = 403
        樹園地_栗 = 404
        樹園地_梅 = 405
        樹園地_桃 = 406
        樹園地_ﾌﾞﾄﾞｳ = 407
        樹園地_ﾂｹﾞ = 408
        樹園地_ｷｨｳｨ = 409
        樹園地_ﾀﾗﾉﾒ = 410
        樹園地_桑 = 411
        樹園地_緑竹 = 412
        樹園地_ｸﾇｷﾞ = 413
        樹園地_造園林 = 414
        ビワ = 415
        樹園地 = 449
        緑 = 450
        緑_自己保全 = 451
        黄 = 460
        黄_自己保全 = 461
        黄_町場_赤状態 = 462
        黄_ほ場整備区域_赤状態 = 463
        黄_農用地_赤状態 = 464
        黄_ｸﾇｷﾞ幼木等 = 465
        黄_経営移譲_赤状態 = 466
        黄_法人_赤状態 = 467
        黄_開発許可地_触らない_赤状態 = 468
        農道通路 = 500
        農業用倉庫等 = 501
        堆肥置場 = 502
        牛舎 = 503
        豚舎 = 504
        鶏舎 = 505
        牧場 = 506
        牧場_ヤギ = 507
        ダチョウ飼育場 = 508
        法面_残地 = 509
        公共地 = 600
        転用許可地 = 650
        公衆道路 = 651
        用悪水路 = 652
        池 = 653
        電気通信施設 = 654
        宅地 = 655
        駐車場 = 656
        雑種地 = 657
        資材置場 = 658
        ゴルフ場 = 659
        墓地 = 660
        公民館 = 661
        学校敷地 = 662
        防火水そう = 663
        ｸﾚｰ射撃場 = 664
        境内地 = 665
        赤 = 700
        赤_山林 = 701
        赤_原野 = 702
        赤_杉 = 703
        赤_ﾋﾉｷ = 704
        赤_竹林 = 705
        赤_ｸﾇｷﾞ = 706
        赤_耕作不能地 = 707
        赤_雑木 = 708
        緑を解消 = 800
        黄を解消 = 801
        赤を解消 = 802
        農政調査済 = 810
        林班図確認 = 820
        林班図該当なし = 821
        再確認 = 830
        再再確認 = 831
        地目農地外 = 832
        非経営耕地 = 833
        地目_田 = 1001
        地目_畑 = 1002
        耕作放棄地_緑 = 1003
        耕作放棄地_黄 = 1004
        耕作放棄地_赤 = 1005
    End Enum
    Public Enum enum区分
        耕作 = -2
        保全管理 = -1
        なし = 0
        Ａ分類 = 1
        Ｂ分類 = 3
        非農地判断済み農地 = 4
        転用 = 5
        不作地 = 6
    End Enum

    <Category("修正内容")> Public Property 耕作作物 As enum耕作作物
    <Category("修正内容")> Public Property 調査状況 As String = ""
    <Category("修正内容")> Public Property 調査備考 As String = ""
    <Category("修正内容")> Public Property 樹齢 As String = ""
    <Category("修正内容")> Public Property 区分 As enum区分

    Public Overrides Function AddRecord() As Long
        Return 0
    End Function


    Public Overrides Function CheckValues() As Boolean

        Return True
    End Function

    Public Overrides Sub SetUpdateRow(ByRef pUpdateRow As HimTools2012.Data.UpdateRow)

    End Sub
End Class

