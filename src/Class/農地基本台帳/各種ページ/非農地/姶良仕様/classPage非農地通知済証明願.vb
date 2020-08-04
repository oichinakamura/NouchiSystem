

Public Class classPage非農地通知済証明願
    Inherits HimTools2012.controls.CTabPageWithToolStrip
    Private mvarTS As HimTools2012.controls.ToolStripContainerEX
    Private mvarTS2 As HimTools2012.controls.ToolStripContainerEX

    Private WithEvents mvarSP As New SplitContainer

    Private WithEvents mvarDBPath As HimTools2012.controls.ToolStripSpringTextBox
    Private WithEvents mvarXLPath As HimTools2012.controls.ToolStripSpringTextBox
    Private WithEvents mvarGridView As New DataGrid非農地証明
    'Public Overrides Function ViewMenuDropDownOpening(ByRef pViewMenu As ToolStripMenuItem) As System.Windows.Forms.ToolStripMenuItem
    '    Return Nothing
    'End Function
    Public Sub New()
        MyBase.new(True, , "非農地通知済証明願", "非農地通知済証明願")

        mvarSP.Dock = DockStyle.Fill

        mvarGridView.Dock = DockStyle.Fill
        mvarGridView.AllowUserToAddRows = False


        mvarGridView.G2.Dock = DockStyle.Fill
        mvarGridView.G2.AllowUserToAddRows = False

        Dim TB As New ToolStrip
        TB.Stretch = True
        TB.GripStyle = ToolStripGripStyle.Hidden
        AddHandler TB.Items.Add("印刷").Click, AddressOf 印刷

        mvarTS2 = New HimTools2012.controls.ToolStripContainerEX(mvarGridView.G2)
        mvarTS2.TopToolStripPanel.Controls.Add(TB)
        mvarSP.Panel1.Controls.Add(mvarGridView)
        mvarSP.Panel2.Controls.Add(mvarTS2)

        mvarTS = New HimTools2012.controls.ToolStripContainerEX(mvarSP)
        Me.Controls.Add(mvarTS)

        Me.ToolStrip.ItemAdd("データベース", New ToolStripButton("データベース"), AddressOf PC)
        mvarDBPath = Me.ToolStrip.ItemAdd("データパス", New HimTools2012.controls.ToolStripSpringTextBox())
        mvarDBPath.ReadOnly = True
        mvarDBPath.AutoSize = True

        Me.ToolStrip.ItemAdd("書式ファイル", New ToolStripButton("書式ファイル"), AddressOf FL)
        mvarXLPath = Me.ToolStrip.ItemAdd("書式ファイルパス", New HimTools2012.controls.ToolStripSpringTextBox())
        mvarXLPath.ReadOnly = True
        mvarXLPath.AutoSize = True


        mvarDBPath.Text = GetSetting(My.Application.Info.AssemblyName, "File", "DBPath", "'農政管理.MDB'を選択してください。")
        If IO.File.Exists(mvarDBPath.Text) Then
            LoadDB()
        End If
        mvarXLPath.Text = GetSetting(My.Application.Info.AssemblyName, "File", "XMLPath", "'非農地通知済証明書.xml'を選択してください。")

    End Sub

    Private Sub 印刷()
        If IO.File.Exists(mvarXLPath.Text) Then
            If mvarGridView.G2.Rows.Count > 0 Then
                Dim DInput As New Input印刷
                Dim pDlg As New HimTools2012.PropertyGridDialog(DInput, "非農地通知証明願い")

                If pDlg.ShowDialog = DialogResult.OK AndAlso DInput.発行年月日.Year > 2000 AndAlso DInput.発行番号.Length > 0 Then
                    Dim sXML As String = HimTools2012.TextAdapter.LoadTextFile(mvarXLPath.Text)
                    Dim sOutPutFile As String = My.Computer.FileSystem.SpecialDirectories.Temp & "\" & Strings.Mid(mvarXLPath.Text, InStrRev(mvarXLPath.Text, "\") + 1)

                    sXML = Replace(sXML, "{発行番号}", DInput.発行番号)
                    sXML = Replace(sXML, "{発行年月日}", 和暦Format(DInput.発行年月日))
                    Dim n As Integer = 1
                    Dim 面積計 As Decimal = 0

                    For Each pRow As DataRowView In mvarGridView.G2.DataSource
                        sXML = Replace(sXML, "{氏名}", pRow.Item("登記名義人氏名").ToString)
                        sXML = Replace(sXML, "{大字" & n & "}", pRow.Item("大字").ToString)
                        sXML = Replace(sXML, "{小字" & n & "}", pRow.Item("小字").ToString)
                        sXML = Replace(sXML, "{地番" & n & "}", pRow.Item("地番").ToString)
                        sXML = Replace(sXML, "{地目" & n & "}", pRow.Item("地目").ToString)
                        sXML = Replace(sXML, "{面積" & n & "}", Val(pRow.Item("面積").ToString).ToString("#,##0"))
                        面積計 += Val(pRow.Item("面積").ToString)
                        If Not IsDBNull(pRow.Item("決定日")) AndAlso pRow.Item("決定日") > #1/1/2000# Then
                            sXML = Replace(sXML, "{決定日" & n & "}", 和暦Format(pRow.Item("決定日"), "gyy.MM.dd"))
                        Else
                            sXML = Replace(sXML, "{決定日" & n & "}", "")
                        End If
                        If Not IsDBNull(pRow.Item("通知日")) AndAlso pRow.Item("通知日") > #1/1/2000# Then
                            sXML = Replace(sXML, "{通知日" & n & "}", 和暦Format(pRow.Item("通知日"), "gyy.MM.dd"))
                        Else
                            sXML = Replace(sXML, "{通知日" & n & "}", "")
                        End If
                        If Not IsDBNull(pRow.Item("発送番号")) AndAlso pRow.Item("発送番号").ToString.Length >= 7 Then
                            sXML = Replace(sXML, "{通知番号" & n & "}", pRow.Item("発送番号").ToString.Substring(0, 3))
                        Else
                            sXML = Replace(sXML, "{通知番号" & n & "}", "")
                        End If

                        n += 1
                    Next
                    For i = n To 13
                        sXML = Replace(sXML, "{大字" & i & "}", "")
                        sXML = Replace(sXML, "{小字" & i & "}", "")
                        sXML = Replace(sXML, "{地番" & i & "}", "")
                        sXML = Replace(sXML, "{地目" & i & "}", "")
                        sXML = Replace(sXML, "{面積" & i & "}", "")
                        sXML = Replace(sXML, "{決定日" & i & "}", "")
                        sXML = Replace(sXML, "{通知日" & i & "}", "")
                        sXML = Replace(sXML, "{通知番号" & i & "}", "")
                    Next
                    sXML = Replace(sXML, "{件数}", n - 1)
                    sXML = Replace(sXML, "{面積計}", 面積計.ToString("#,##0"))
                    sXML = Replace(sXML, "姶農委第号", "")
                    HimTools2012.TextAdapter.SaveTextFile(sOutPutFile, sXML)
                    SaveAndOpen(sOutPutFile)
                End If
            End If
        Else
            MsgBox("正しく「非農地通知済証明願.XML」ファイルが設定されていません。", MsgBoxStyle.Critical)
        End If
    End Sub

    Public Sub SaveAndOpen(ByVal sFile As String)
        Dim pExcel As Object = CreateObject("Excel.Application")
        Try
            Dim oBook As Object = Nothing

            Try
                oBook = pExcel.WorkBooks.open(sFile)
                pExcel.visible = True

                Try
                    oBook.sheets.PrintPreview(True)
                Catch ex As Exception

                End Try
            Catch ex As Exception
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook)
            End Try
            pExcel.Quit()
        Catch ex As Exception

        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pExcel)
        End Try
    End Sub

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get

            Return HimTools2012.controls.CloseMode.NoMessage
        End Get
    End Property

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

    Private Sub FL()
        With New OpenFileDialog
            .Filter = "*.xml|*.xml"
            If .ShowDialog = DialogResult.OK Then
                mvarXLPath.Text = .FileName
                SaveSetting(My.Application.Info.AssemblyName, "File", "XMLPath", .FileName)
            End If
        End With
    End Sub

    Private Sub LoadDB()
        Try
            With New HimTools2012.Data.CLocalDataEngine("非農地証明")
                .LocalPath = mvarDBPath.Text
                mvarGridView.bLoading = True

                Dim pTable As DataTable
                pTable = .GetTableBySqlSelect_Local("SELECT 0 AS 印刷順, CODE, 決定日, 通知日, 大字, 小字, 地番,地目,面積,登記名義人ID, 登記名義人氏名, 登記名義人住所, 発送番号,本番,枝番 FROM [D_非農地情報]")
                pTable.PrimaryKey = New DataColumn() {pTable.Columns("CODE")}

                Dim pColumn As New DataGridViewTextBoxColumn
                pColumn.Name = "印刷順"
                pColumn.HeaderText = "印刷順"
                pColumn.DataPropertyName = "印刷順"

                mvarGridView.Columns.Add(pColumn)
                mvarGridView.SetDataView(pTable, "", "大字,本番,枝番")
                mvarGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
                mvarGridView.Columns("本番").Visible = False
                mvarGridView.Columns("枝番").Visible = False
                mvarGridView.ReadOnly = True
                mvarGridView.ClearSelection()

                mvarGridView.G2.DataSource = New DataView(mvarGridView.DataTable, "[印刷順]>0", "印刷順", DataViewRowState.CurrentRows)
                mvarGridView.G2.Columns("CODE").Visible = False
                mvarGridView.G2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End With
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & "データベース「農政管理.MDB」を選択しなおしてください。")
            PC()
        End Try
    End Sub

    Private mvarV As Boolean = False
    Private Sub mvarSP_Paint(sender As Object, e As System.Windows.Forms.PaintEventArgs) Handles mvarSP.Paint
        If Not mvarV Then
            mvarSP.SplitterDistance = mvarSP.ClientRectangle.Width / 2
            mvarGridView.bLoading = False
            mvarV = True
        End If
    End Sub
End Class

Public Class Input印刷
    Inherits HimTools2012.InputSupport.CInputSupport
    Public Sub New()
        MyBase.New(Nothing)
        発行年月日 = Now.Date
    End Sub

    Public Property 発行番号 As String = ""
    Public Property 発行年月日 As DateTime
End Class

Public Class DataGrid非農地証明
    Inherits HimTools2012.controls.DataGridViewWithDataView

    Private mvarSelCount As Integer = 0
    Private mvarHeader As HimTools2012.controls.ColumnHeaderEX
    Public WithEvents G2 As New DataGridView

    Public Sub New()
        MyBase.New()
        mvarSelCount = 0
        mvarHeader = New HimTools2012.controls.ColumnHeaderEX(Me)
        Me.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
    End Sub


    Public bLoading As Boolean = False

    Private Sub DataGrid非農地証明_SelectionChanged(sender As Object, e As System.EventArgs) Handles Me.SelectionChanged
        If Not bLoading Then

            G2.SuspendLayout()
            If SelectedRows.Count = 0 Then
                Dim pView As New DataView(Me.DataTable, "[印刷順]>0", "", DataViewRowState.CurrentRows)
                For Each pRow As DataRowView In pView
                    pRow.Item("印刷順") = 0
                Next
                mvarSelCount = 0
            ElseIf SelectedRows.Count = 1 Then
                For Each pRow As DataGridViewRow In Me.Rows
                    If Not pRow.Selected Then
                        pRow.Cells("印刷順").Value = 0
                    End If
                Next
                mvarSelCount = 1
                Me.DataTable.Rows.Find(SelectedRows(0).Cells("CODE").Value).Item("印刷順") = mvarSelCount
            Else
                For Each pRow As DataGridViewRow In Me.Rows
                    If Not pRow.Selected Then
                        pRow.Cells("印刷順").Value = 0
                    End If
                Next
                Dim pView As New DataView(Me.DataTable, "[印刷順] >0", "印刷順", DataViewRowState.CurrentRows)
                Dim n As Integer = 1
                For Each pRowV As DataRowView In pView
                    pRowV.Item("印刷順") = n
                    n += 1
                Next

                mvarSelCount = pView.Count
                Dim nMax As Integer = SelectedRows.Count
                For Each pRow As DataGridViewRow In SelectedRows
                    If pRow.Cells("印刷順").Value = 0 Then
                        mvarSelCount += 1
                        Me.DataTable.Rows.Find(pRow.Cells("CODE").Value).Item("印刷順") = nMax
                        nMax -= 1
                    End If
                Next
            End If
            G2.ResumeLayout()

            G2.Sort(G2.Columns("印刷順"), System.ComponentModel.ListSortDirection.Ascending)
        End If

    End Sub

    Private Sub G2_CellErrorTextNeeded(sender As Object, e As System.Windows.Forms.DataGridViewCellErrorTextNeededEventArgs) Handles G2.CellErrorTextNeeded
        If G2.Columns(e.ColumnIndex).DataPropertyName = "登記名義人氏名" Then
            Dim query = From order In CType(G2.DataSource, DataView).ToTable.AsEnumerable()
                          Group order By s項目 = order.Field(Of String)("登記名義人氏名") Into g = Group
                          Select New With
                          {
                              .選択枝 = s項目
                          }

            If query.Count > 1 Then
                e.ErrorText = "登記名義が複数あります"
            End If
        End If
    End Sub

    Private Sub G2_DataError(sender As Object, e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles G2.DataError

    End Sub
End Class
