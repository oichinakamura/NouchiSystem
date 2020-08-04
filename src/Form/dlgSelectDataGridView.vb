Imports System.Windows.Forms

Public Class dlgSelectDataGridView
    Public ResultRow As DataRow() = Nothing
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If DataGridView1.SelectedRows.Count = 1 Then
            If TypeOf DataGridView1.DataSource Is DataTable Then
                Dim pList As New List(Of DataRow)

                For Each pRow As DataGridViewRow In DataGridView1.SelectedRows
                    pList.Add(CType(DataGridView1.DataSource, DataTable).Rows(pRow.Index))
                Next

                ResultRow = pList.ToArray
            Else

            End If
        End If
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Public Sub AddViewColumn(ByVal sField As String, ByVal pType As System.Type, Optional ByRef pTBL As DataTable = Nothing)
        If pTBL IsNot Nothing Then
            pTBL.Columns.Add(sField, pType)
        End If

        Select Case pType.Name
            Case "Boolean"
                Dim pItem As New DataGridViewCheckBoxColumn()
                pItem.HeaderText = sField
                pItem.DataPropertyName = sField
                Grid.Columns.Add(pItem)
            Case "String"
                Dim pItem As New DataGridViewTextBoxColumn()
                pItem.HeaderText = sField
                pItem.DataPropertyName = sField
                Grid.Columns.Add(pItem)
            Case Else
                Stop
        End Select

    End Sub

    Public ReadOnly Property Grid As DataGridView
        Get
            Return DataGridView1
        End Get
    End Property

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        If DataGridView1.SelectedRows.Count = 1 Then
            OK_Button.Enabled = True
        Else
            OK_Button.Enabled = False
        End If
    End Sub

    Public Sub New()

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        OK_Button.Enabled = False

    End Sub

    Private Sub ToExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToExcel.Click
        ToExcelCmd(DataGridView1.DataSource, My.Computer.FileSystem.SpecialDirectories.Desktop & "\ToExcel.xml")
    End Sub

    Private Sub ToExcelCmd(ByVal pTable As System.Data.DataTable, ByVal sFileName As String)
        With New HimTools2012.ExcelSpreadSheet.XMLSpreadSheet
            With .WorkBook()
                With .Styles
                    With .Add("Default", "Normal", "Default")
                    End With
                    With .Add("s24", "")
                        .Alignment.Horizontal = "Center"
                        .Alignment.Vertical = "Center"
                        .NumberFormat = .CreateNumberFormat("Short Date")
                    End With
                    With .Add("s61", "")
                        .Alignment.Horizontal = "Left"
                        .Alignment.Vertical = "Center"
                    End With

                    With .Add("s62", "")
                        .Alignment.Horizontal = "Center"
                        .Alignment.Vertical = "Center"
                    End With
                    With .Add("s63", "")
                        .Alignment.Horizontal = "Center"
                        .Alignment.Vertical = "Center"
                        .NumberFormat = .CreateNumberFormat("0_ ;&quot;△&quot;0_ ")
                    End With
                End With

                Dim pSheet As New HimTools2012.ExcelSpreadSheet.ExcelWorkSheet(pTable.TableName)
                With pSheet
                    With .Table
                        Dim nCount As Integer = 0

                        With .AddRow
                            For Each pCol As DataColumn In pTable.Columns
                                .AddCell("s62", "String", pCol.ColumnName)
                            Next
                        End With

                        For Each pRow As DataRow In pTable.Rows
                            nCount += 1
                            With .AddRow
                                For Each pCol As DataColumn In pTable.Columns
                                    Select Case pCol.DataType.FullName
                                        Case "System.Boolean"
                                            .AddCell("s61", "String", pRow.Item(pCol.ColumnName).ToString)
                                        Case "System.Type"
                                            .AddCell("s61", "String", pRow.Item(pCol.ColumnName).ToString)
                                        Case "System.Int16", "System.Int32"
                                            .AddCell("s62", "Number", pRow.Item(pCol.ColumnName))
                                        Case "System.String"
                                            .AddCell("s61", "String", pRow.Item(pCol.ColumnName).ToString)
                                        Case Else
                                            .AddCell("s61", "String", pRow.Item(pCol.ColumnName).ToString)
                                    End Select
                                Next
                            End With
                        Next
                    End With

                End With
                .WorkSheets.Add(pTable.TableName, pSheet)

            End With
            HimTools2012.TextAdapter.SaveTextFile(sFileName, .ToString)
        End With

    End Sub

End Class
