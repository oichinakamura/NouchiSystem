Imports System.Windows.Forms

Public Class dlg複数地番検索
    Public sResult As String = ""
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Me.ComboBox1.SelectedItem IsNot Nothing AndAlso mvarData.Rows.Count > 0 Then
            sResult = "[大字ID]=" & CType(Me.ComboBox1.SelectedItem, DataRowView).Item("ID")

            sResult &= " AND [地番] IN ("
            Dim SC As String = ""
            For Each pRow As DataRow In mvarData.Rows
                If pRow.Item("地番").ToString.Length > 0 Then
                    sResult &= SC & "'" & pRow.Item("地番").ToString & "'"
                    SC = ","
                End If
            Next
            sResult &= ")"
            SysAD.SetXMLProperty("複数地番検索", "選択大字", CType(Me.ComboBox1.SelectedItem, DataRowView).Item("ID"))
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        End If
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub
    Private mvarData As DataTable
    Private Sub dlg複数地番検索_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.ComboBox1.ValueMember = "ID"
        Me.ComboBox1.DisplayMember = "名称"
        Me.ComboBox1.DataSource = App農地基本台帳.TBL大字
        Me.ComboBox1.SelectedValue = Val(SysAD.GetXMLProperty("複数地番検索", "選択大字", "0"))
        mvarData = New DataTable("検索地番")

        mvarData.Columns.Add("地番")
        mvarData.PrimaryKey = {mvarData.Columns("地番")}
        DataGridView1.DataSource = mvarData

    End Sub

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError

    End Sub

    Private Sub DataGridView1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DataGridView1.KeyDown
        If e.Control = True AndAlso (e.KeyCode And Keys.V) = Keys.V Then
            Dim iData As IDataObject = System.Windows.Forms.Clipboard.GetDataObject
            Dim myStr As String = CType(iData.GetData(DataFormats.UnicodeText), String)
            With DataGridView1
                Dim sLines() As String = Split(myStr, vbCrLf)
                For Each sLine As String In sLines
                    If sLine.Length > 0 Then
                        Dim sColumns() As String = Split(sLine, vbTab)
                        For Each sColumn As String In sColumns
                            If Trim(sColumn).Length > 0 Then
                                If mvarData.Rows.Find(Trim(sColumn)) Is Nothing Then
                                    mvarData.Rows.Add(Trim(sColumn))
                                End If
                            End If
                        Next
                    End If
                Next
            End With
        End If
    End Sub
End Class
