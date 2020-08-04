

Public Class CTabPage重複農地
    Inherits HimTools2012.TabPages.CTabPageWithDataGridView

    Public Sub New()
        MyBase.New(True, "重複農地", "重複農地", ObjectMan)

        If App農地基本台帳.TBL農地.Columns.Contains("重複X") Then
            App農地基本台帳.TBL農地.Columns.Remove(App農地基本台帳.TBL農地.Columns("重複X"))
        End If
        App農地基本台帳.TBL農地.Columns.Add("重複X", GetType(Integer))

        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:農地Info].*,1 as [重複X] FROM [D:農地Info] WHERE ((([D:農地Info].大字ID)>0) AND (([D:農地Info].[大字ID]) In (SELECT [大字ID] FROM [D:農地Info] As Tmp GROUP BY [大字ID],[小字ID],[地番],[一部現況] HAVING Count(*)>1  And [小字ID] = [D:農地Info].[小字ID] And [地番] = [D:農地Info].[地番] And [一部現況] = [D:農地Info].[一部現況]))) ORDER BY [D:農地Info].[大字ID];")
        App農地基本台帳.TBL農地.MergePlus(pTBL)

        mvarGrid.Columns(mvarGrid.Columns.Add("ID", "ID")).DataPropertyName = "ID"
        mvarGrid.Columns(mvarGrid.Columns.Add("大字", "大字")).DataPropertyName = "大字"
        mvarGrid.Columns(mvarGrid.Columns.Add("小字", "小字")).DataPropertyName = "小字"
        mvarGrid.Columns(mvarGrid.Columns.Add("所在", "所在")).DataPropertyName = "所在"
        mvarGrid.Columns(mvarGrid.Columns.Add("地番", "地番")).DataPropertyName = "地番"
        mvarGrid.Columns(mvarGrid.Columns.Add("一部現況", "一部現況")).DataPropertyName = "一部現況"

        mvarGrid.Columns(mvarGrid.Columns.Add("登記簿地目名", "登記簿地目名")).DataPropertyName = "登記簿地目名"
        mvarGrid.Columns(mvarGrid.Columns.Add("現況地目名", "現況地目名")).DataPropertyName = "現況地目名"
        mvarGrid.Columns(mvarGrid.Columns.Add("登記簿面積", "登記簿面積")).DataPropertyName = "登記簿面積"
        mvarGrid.Columns(mvarGrid.Columns.Add("実面積", "実面積")).DataPropertyName = "実面積"
        mvarGrid.Columns(mvarGrid.Columns.Add("所有者氏名", "所有者氏名")).DataPropertyName = "所有者氏名"
        mvarGrid.Columns(mvarGrid.Columns.Add("自小作", "自小作")).DataPropertyName = "自小作"
        mvarGrid.Columns(mvarGrid.Columns.Add("借受人氏名", "借受人氏名")).DataPropertyName = "借受人氏名"
        mvarGrid.AutoGenerateColumns = False

        For Each pCol As DataGridViewColumn In mvarGrid.Columns
            Select Case pCol.DataPropertyName
                Case "地番"
                Case "一部現況"
                Case Else
                    pCol.ReadOnly = True
            End Select
        Next



        mvarGrid.SetDataView(App農地基本台帳.TBL農地.Body, "[重複X]=1", "[大字],[地番]")
        For Each pRowV As DataRowView In mvarGrid.DataView

        Next

        AddHandler mvarGrid.CellEndEdit, AddressOf mvarGrid_CellEndEdit
        AddHandler mvarGrid.RowHeaderMouseClick, AddressOf mvarGrid_RowHeaderMouseClick

        Me.ControlPanel.Add(mvarGrid)
    End Sub

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.CloseOK
        End Get
    End Property


    Private mvarMenuRow As DataGridViewRow

    Private Sub mvarGrid_CellEndEdit(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim sSQL As String = ""
        Select Case mvarGrid.Columns(e.ColumnIndex).DataPropertyName
            Case "地番" : sSQL = DataRowHelper.GetUpdateSQL(App農地基本台帳.TBL農地.FindRowByID(mvarGrid.Item("ID", e.RowIndex).Value), "ID", GetType(Int32), "地番")
            Case "一部現況" : sSQL = DataRowHelper.GetUpdateSQL(App農地基本台帳.TBL農地.FindRowByID(mvarGrid.Item("ID", e.RowIndex).Value), "ID", GetType(Int32), "一部現況")
        End Select
        If sSQL.Length > 0 Then
            SysAD.DB(sLRDB).ExecuteSQL(sSQL)
        End If
    End Sub

    Private Sub mvarGrid_RowHeaderMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs)
        mvarMenuRow = Nothing
        If e.RowIndex > -1 AndAlso mvarGrid.SelectedRows IsNot Nothing AndAlso e.Button = Windows.Forms.MouseButtons.Right Then
            Dim pMenu As New ContextMenuStrip

            Select Case mvarGrid.SelectedRows.Count
                Case 0
                Case 1
                    mvarMenuRow = mvarGrid.Rows(e.RowIndex)
                    AddHandler pMenu.Items.Add("削除").Click, AddressOf Del
                Case Else
                    'mvarMenuRow = mvarGrid.Rows(e.RowIndex)
                    'AddHandler pMenu.Items.Add("同定処理").Click, AddressOf sub同定処理
            End Select


            pMenu.Show(Control.MousePosition)
        End If
    End Sub
    Private Sub Del()
        If mvarMenuRow IsNot Nothing AndAlso MsgBox("農地を削除しますか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            農地削除(New DataView(App農地基本台帳.TBL農地.Body, "[ID]=" & mvarMenuRow.Cells("ID").Value, "", DataViewRowState.CurrentRows).ToTable, 261, C農地削除.enum転送先.転用農地, "旧台帳より非農地確定")
        End If
    End Sub
    Private Sub sub同定処理()
        Dim pList As New List(Of DataGridViewRow)
        For Each pRow As DataGridViewRow In mvarGrid.SelectedRows
            If pList.Count = 0 Then
                pList.Add(pRow)
            ElseIf pList(0).Cells("大字ID").Value = pRow.Cells("大字ID").Value AndAlso pList(0).Cells("地番").Value = pRow.Cells("地番").Value Then

            End If
        Next
    End Sub
End Class

Public Class DataRowHelper
    Private mvarRow As DataRow = Nothing
    Private mvarKeyName As String = ""
    Private mvarKeyType As System.Type = Nothing

    Public Sub New(ByRef pRow As DataRow, ByVal sKeyName As String, ByVal KeyType As System.Type)
        mvarRow = pRow
        If mvarRow.Table.TableName.ToString = "" Then
            MsgBox("テーブルネームを設定してください", MsgBoxStyle.Exclamation)
            Stop
        End If
        mvarKeyName = sKeyName
        mvarKeyType = KeyType
    End Sub

    Public Shared Function GetUpdateSQL(ByRef pRow As DataRow, ByVal sKeyName As String, ByVal KeyType As System.Type, ByVal sField As String)
        Dim mvarRow As DataRow = pRow
        Dim mvarKeyName As String = sKeyName
        Dim mvarKeyType As System.Type = KeyType

        If mvarRow.Table.TableName.ToString = "" Then
            MsgBox("テーブルネームを設定してください", MsgBoxStyle.Exclamation)
            Return ""
        End If


        Dim sSQL As String = ""
        If IsDBNull(pRow.Item(sField)) Then
            Return "UPDATE [" & mvarRow.Table.TableName & "] SET [" & sField & "]=Null WHERE [" & sKeyName & "]=" & pRow.Item(sKeyName)
        Else
            Select Case pRow.Table.Columns(sField).DataType.ToString
                Case "System.String"
                    Return "UPDATE [" & mvarRow.Table.TableName & "] SET [" & sField & "]='" & pRow.Item(sField).ToString & "' WHERE [" & sKeyName & "]=" & pRow.Item(sKeyName)
                Case "System.Boolean"
                    Return "UPDATE [" & mvarRow.Table.TableName & "] SET [" & sField & "]=" & CBool(pRow.Item(sField)).ToString & " WHERE [" & sKeyName & "]=" & pRow.Item(sKeyName)
                Case "System.Int32"
                    Return "UPDATE [" & mvarRow.Table.TableName & "] SET [" & sField & "]=" & pRow.Item(sField).ToString & " WHERE [" & sKeyName & "]=" & pRow.Item(sKeyName)
                Case Else
                    Stop
            End Select
        End If

        Return sSQL
    End Function

End Class


