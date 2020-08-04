
Imports System.ComponentModel
Imports HimTools2012.Data

Public Class CommonMASTER管理
    Inherits HimTools2012.controls.CTabPageWithToolStrip
    Private mvarSP As SplitContainer
    Private WithEvents mvarG As HimTools2012.controls.DataGridViewWithDataView
    Private WithEvents mvarT As TreeView
    Private WithEvents 追加 As ToolStripButton
    Private WithEvents 削除 As ToolStripButton
    Private WithEvents 保存 As ToolStripButton
    Private TBLXML As DataTable

    Public Sub New()
        MyBase.New(True, True, "CommonMASTER管理", "コモンマスタ管理")

        Dim pTBL As New DataTableEx()
        pTBL.LoadText(My.Resources.Resource1.M_BASICALL)

        TBLXML.PrimaryKey = {TBLXML.Columns("ID"), TBLXML.Columns("Class")}

        mvarT = New TreeView
        mvarT.CheckBoxes = True
        mvarT.Dock = DockStyle.Fill
        mvarG = New HimTools2012.controls.DataGridViewWithDataView

        mvarSP = New SplitContainer
        mvarSP.Panel1.Controls.Add(mvarT)
        mvarSP.Panel2.Controls.Add(mvarG)
        mvarSP.Dock = DockStyle.Fill
        Me.ControlPanel.Add(mvarSP)

        For Each pRow As DataRow In TBLXML.Rows
            If mvarT.Nodes.Find(pRow.Item("Class"), True).Length = 0 Then
                mvarT.Nodes.Add(pRow.Item("Class"), pRow.Item("Class"))
            End If
        Next

        mvarG.SetDataView(TBLXML, "", "[Class],[ID]")
        追加 = New ToolStripButton("追加")
        追加.Enabled = False

        削除 = New ToolStripButton("削除")
        削除.Enabled = False

        保存 = New ToolStripButton("保存")
        保存.Enabled = False

        Me.ToolStrip.Items.AddRange({追加, 削除, 保存})

        Dim pMenu As New ContextMenuStrip
        With CType(pMenu.Items.Add("追加", Nothing, AddressOf 追加_Click), ToolStripMenuItem)
            .ShowShortcutKeys = True
            .ShortcutKeys = Keys.P Or Keys.Control
        End With
        With CType(pMenu.Items.Add("挿入", Nothing, AddressOf 挿入_Click), ToolStripMenuItem)
            .ShowShortcutKeys = True
            .ShortcutKeys = Keys.Insert Or Keys.Control
        End With
        With CType(pMenu.Items.Add("貼り付け", Nothing, AddressOf 貼り付け_Click), ToolStripMenuItem)
            .ShowShortcutKeys = True
            .ShortcutKeys = Keys.V Or Keys.Control
        End With
        With CType(pMenu.Items.Add("保存", Nothing, AddressOf 保存_Click), ToolStripMenuItem)
            .ShowShortcutKeys = True
            .ShortcutKeys = Keys.S Or Keys.Control
        End With

        mvarG.ContextMenuStrip = pMenu
        mvarG.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
    End Sub

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            If 保存.Enabled Then
                Select Case MsgBox("未保存の項目があります。保存しますか", MsgBoxStyle.YesNoCancel)
                    Case MsgBoxResult.Yes
                        保存_Click(Nothing, Nothing)
                        Return HimTools2012.controls.CloseMode.CloseOK
                    Case MsgBoxResult.No
                        Return HimTools2012.controls.CloseMode.CloseOK
                    Case Else
                        Return HimTools2012.controls.CloseMode.CancelClose
                End Select
            Else
                Return HimTools2012.controls.CloseMode.CloseOK
            End If
        End Get
    End Property

    Private Sub mvarT_AfterCheck(sender As Object, e As System.Windows.Forms.TreeViewEventArgs) Handles mvarT.AfterCheck
        Dim sList As New List(Of String)
        For Each pNode As TreeNode In mvarT.Nodes
            If pNode.Checked Then
                sList.Add("'" & pNode.Name & "'")
            End If
        Next
        If sList.Count > 0 Then
            mvarG.RowFilter = "[Class] IN (" & Join(sList.ToArray, ",") & ")"
        Else
            mvarG.RowFilter = ""
        End If
    End Sub

    Private Sub mvarT_AfterSelect(sender As Object, e As System.Windows.Forms.TreeViewEventArgs) Handles mvarT.AfterSelect
        追加.Enabled = (mvarT.SelectedNode IsNot Nothing)
    End Sub

    Private Sub 追加_Click(sender As Object, e As System.EventArgs) Handles 追加.Click
        If mvarT.SelectedNode IsNot Nothing Then
            With New HimTools2012.PropertyGridDialog(New Input追加(TBLXML, mvarT.SelectedNode.Name), "追加")
                If .ShowDialog = DialogResult.OK Then
                    Dim sClass As String = CType(.ResultProperty, Input追加).クラス
                    If sClass.Length > 0 Then
                        Dim pNewRow As DataRow = TBLXML.NewRow
                        Dim pU As Input追加 = CType(.ResultProperty, Input追加)
                        pNewRow.Item("ID") = pU.ID
                        pNewRow.Item("Class") = sClass
                        pNewRow.Item("名称") = pU.名称
                        Dim pT() As TreeNode = mvarT.Nodes.Find(sClass, True)

                        If pT.Length = 0 Then
                            mvarT.Nodes.Add(sClass, sClass)
                        End If

                        TBLXML.Rows.Add(pNewRow)
                        保存.Enabled = True
                    End If
                End If
            End With
        Else

        End If
    End Sub
    Private Sub 挿入_Click(sender As Object, e As System.EventArgs)
        If mvarT.SelectedNode IsNot Nothing Then
            If mvarG.CurrentRow IsNot Nothing Then
                Dim nID As Integer = mvarG.CurrentRow.Cells("ID").Value - 1
                Dim sClass As String = mvarG.CurrentRow.Cells("Class").Value
                Dim pV As New DataView(TBLXML, String.Format("[class]='{0}' AND [ID]={1}", sClass, nID), "", DataViewRowState.CurrentRows)

                If pV.Count = 0 Then
                    Dim pRow As DataRow = TBLXML.NewRow
                    pRow.Item("ID") = nID
                    pRow.Item("Class") = sClass
                    TBLXML.Rows.Add(pRow)
                End If

            End If
        End If
    End Sub

    Private Sub 貼り付け_Click(sender As Object, e As System.EventArgs)
        Dim pasteText As String = Clipboard.GetText()
        If tb Is Nothing Then
            Dim insertRowIndex As Integer = mvarG.CurrentCell.RowIndex
            Dim insertColIndex As Integer = mvarG.CurrentCell.ColumnIndex

            If String.IsNullOrEmpty(pasteText) Then
                Return
            End If
            pasteText = pasteText.Replace(vbCrLf, vbLf)
            pasteText = pasteText.Replace(vbCr, vbLf)
            pasteText = pasteText.TrimEnd(New Char() {vbLf})
            Dim lines As String() = pasteText.Split(vbLf)

            Dim isHeader As Boolean = True
            For Each line As String In lines
                '列ヘッダーならば飛ばす
                If isHeader Then
                    isHeader = False
                Else
                    'タブで分割
                    Dim vals As String() = line.Split(ControlChars.Tab)
                    '列数が合っているか調べる
                    If vals.Length - 1 <> mvarG.ColumnCount Then
                        '    Throw New ApplicationException("列数が違います。")
                    End If
                    Dim row As DataGridViewRow = mvarG.Rows(insertRowIndex)

                    'ヘッダーを設定
                    For i As Integer = 0 To vals.Length - 2
                        row.Cells(i + insertColIndex).Value = vals((i + 1))
                    Next

                    '次の行へ
                    insertRowIndex += 1
                End If
            Next
        Else
            tb.Paste()
        End If

    End Sub

    Private Sub 複写_Click(sender As Object, e As System.EventArgs)
        If mvarG.SelectedRows IsNot Nothing AndAlso mvarG.SelectedRows.Count = 1 Then

        ElseIf mvarG.SelectedCells IsNot Nothing AndAlso mvarG.SelectedCells.Count = 1 Then

        End If
    End Sub
    Private Sub mvarT_NodeMouseClick(sender As Object, e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles mvarT.NodeMouseClick
        e.Node.Checked = Not e.Node.Checked
    End Sub

    Private Class Input追加
        Inherits HimTools2012.InputSupport.CInputSupport
        Private mvarT As DataTable
        <Category("01基本情報")>
        Public Property ID As Integer = 0
        Private mvarClass As String
        <Category("01基本情報")>
        Public Property クラス As String
            Get
                Return mvarClass
            End Get
            Set(value As String)
                mvarClass = value
            End Set
        End Property
        <Category("01基本情報")>
        Public Property 名称 As String

        Public Sub New(pTable As DataTable, sClass As String)
            MyBase.New(pTable)
            mvarT = pTable
            mvarClass = sClass
            Dim pV As New DataView(mvarT, String.Format("[class]='{0}'", Me.クラス), "", DataViewRowState.CurrentRows)
            For Each pRowV As DataRowView In pV
                If pRowV.Item("ID") > Me.ID Then
                    Me.ID = pRowV.Item("ID") + 1
                End If
            Next
        End Sub

        Public Overrides Function DataCompleate() As Boolean
            Dim pV As New DataView(mvarT, String.Format("[class]='{0}' AND [ID]={1}", Me.クラス, Me.ID), "", DataViewRowState.CurrentRows)
            If pV.Count = 0 Then
                Return True
            Else
                Return False
            End If
        End Function

    End Class
    Private Sub 保存_Click(sender As Object, e As System.EventArgs) Handles 保存.Click
        Dim MyPath As String = My.Application.Info.DirectoryPath

        If InStr(MyPath, "\農政パック\農地基本台帳\bin") Then
            Try
                Dim sDataPath As String = HimTools2012.StringF.Left(MyPath, InStr(MyPath.ToLower, "\bin"))
                TBLXML.WriteXml(sDataPath & "Resources\M_BASICALL.xml", System.Data.XmlWriteMode.WriteSchema, False)
                MsgBox("保存しました。(保存内容は再起動・コンパイル後に有効になります)", MsgBoxStyle.Information)
                保存.Enabled = False
            Catch ex As Exception
                MsgBox("保存に失敗しました。", MsgBoxStyle.Exclamation)
            End Try
        End If
    End Sub

    Private Sub mvarG_CellValueChanged(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles mvarG.CellValueChanged
        保存.Enabled = True
    End Sub

 
    Private Sub mvarG_SelectionChanged(sender As Object, e As System.EventArgs) Handles mvarG.SelectionChanged
        削除.Enabled = (mvarG.SelectedRows IsNot Nothing AndAlso mvarG.SelectedRows.Count > 0)
    End Sub

    Dim WithEvents dataGridViewTextBox As TextBox
    Dim tb As DataGridViewTextBoxEditingControl

    Private Sub mvarG_CellEndEdit(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles mvarG.CellEndEdit
        If tb IsNot Nothing Then
            Try
                RemoveHandler tb.KeyDown, AddressOf dataGridViewTextBox_KeyDown
            Catch ex As Exception

            End Try
            tb = Nothing
        End If
    End Sub

    Private Sub mvarG_EditingControlShowing(ByVal sender As Object, ByVal e As DataGridViewEditingControlShowingEventArgs) Handles mvarG.EditingControlShowing
        If TypeOf e.Control Is DataGridViewTextBoxEditingControl Then
            tb = CType(e.Control, DataGridViewTextBoxEditingControl)

            Try
                RemoveHandler tb.KeyDown, AddressOf dataGridViewTextBox_KeyDown
            Catch ex As Exception

            End Try
            AddHandler tb.KeyDown, AddressOf dataGridViewTextBox_KeyDown

        End If
    End Sub

    Private Sub dataGridViewTextBox_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs)
        With CType(sender, DataGridViewTextBoxEditingControl)
            If (e.KeyCode And Keys.X) = Keys.X AndAlso e.Control = True AndAlso .SelectedText.Length > 0 Then
                .Cut()
            ElseIf (e.KeyCode And Keys.C) = Keys.C AndAlso e.Control = True AndAlso .SelectedText.Length > 0 Then
                .Copy()
            ElseIf (e.KeyCode And Keys.V) = Keys.V AndAlso e.Control = True Then
                .Paste()
            End If
        End With
    End Sub

End Class


