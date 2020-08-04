
Public Class CListColumnDesign
    Inherits HimTools2012.Data.DataSetEx

    Public Sub New()
        MyBase.New("DataListViewカラム設定")

        Dim strXML As String = My.Resources.Resource1.DataListViewカラム設定
        Me.ReadXml(New IO.StringReader(strXML))

    End Sub

    Public Sub CreateGridViewDesign(ByRef pGrid As HimTools2012.controls.DataGridViewWithDataView, ByVal sKey As String)
        If Me.Tables.Contains(sKey) Then
            Me.Tables.Remove(sKey)
        End If

        If Not pGrid.Columns.Contains("Key") Then
            Dim pKeyColumn As New DataGridViewTextBoxColumn
            pKeyColumn.DataPropertyName = "Key"
            pKeyColumn.Name = "Key"
            pKeyColumn.Visible = False
            pGrid.Columns.Add(pKeyColumn)
        End If

        If Not pGrid.Columns.Contains("アイコン") Then
            Dim pIconColumn As New DataGridViewTextBoxColumn
            pIconColumn.DataPropertyName = "アイコン"
            pIconColumn.Name = "アイコン"
            pIconColumn.Visible = False
            pGrid.Columns.Add(pIconColumn)
        End If


        Dim pTBLX As New DataTable(sKey)
        pTBLX.Columns.Add("Index", GetType(Integer))
        pTBLX.Columns.Add("ColumnName", GetType(String))
        pTBLX.Columns.Add("DataPropertyName", GetType(String))
        pTBLX.Columns.Add("HeaderText", GetType(String))
        pTBLX.Columns.Add("ReadOnly", GetType(Boolean))
        pTBLX.Columns.Add("DisplayIndex", GetType(Integer))
        pTBLX.Columns.Add("TextAlignment", GetType(Integer))
        pTBLX.Columns.Add("DataType", GetType(String))
        pTBLX.Columns.Add("M_BASICClass", GetType(String))
        pTBLX.Columns.Add("Visible", GetType(Boolean))

        For Each pCol As DataGridViewColumn In pGrid.Columns
            Dim pAddRow As DataRow = pTBLX.NewRow
            With pCol
                pAddRow.Item("Index") = pCol.Index
                pAddRow.Item("ColumnName") = pCol.Name
                pAddRow.Item("DataPropertyName") = pCol.DataPropertyName
                pAddRow.Item("HeaderText") = pCol.HeaderText
                pAddRow.Item("DisplayIndex") = pCol.DisplayIndex
                pAddRow.Item("TextAlignment") = pCol.DefaultCellStyle.Alignment
                If pCol.ValueType IsNot Nothing Then
                    pAddRow.Item("DataType") = pCol.ValueType.FullName
                End If
                pAddRow.Item("Visible") = pCol.Visible
            End With

            pTBLX.Rows.Add(pAddRow)
        Next

        Me.Tables.Add(pTBLX)
    End Sub

    Public Sub Save(ByVal sProject名 As String)
        Dim pPath As IO.DirectoryInfo = New IO.DirectoryInfo(My.Application.Info.DirectoryPath)

        If Not SysAD.IsClickOnceDeployed Then
            Do Until pPath.Name = sProject名
                pPath = pPath.Parent
            Loop

            For Each pTarget As IO.DirectoryInfo In pPath.GetDirectories
                If pTarget.Name = "Resources" Then
                    Try
                        Me.WriteXml(pTarget.FullName & "\DataListViewカラム設定.xml", XmlWriteMode.WriteSchema)
                        Exit For
                    Catch ex As Exception

                    End Try
                End If
            Next

        End If
    End Sub

    Public Sub SetGridColumns(ByRef pGrid As HimTools2012.controls.DataGridViewWithDataView, ByVal sKey As String)
        With pGrid
            .SuspendLayout()
            If Me.Tables.Contains(sKey) Then
                .AutoGenerateColumns = False
                .Columns.Clear()

                Dim pDesignMaster As DataTable = Me.Tables(sKey)

                If Not pDesignMaster.Columns.Contains("M_BASICClass") Then
                    pDesignMaster.Columns.Add("M_BASICClass", GetType(String))
                End If

                For Each pRow As DataRowView In New DataView(pDesignMaster, "", "Index", DataViewRowState.CurrentRows)
                    Select Case pRow.Item("ColumnName")
                        Case "更新" : pGrid.AddButtonColumn("更新", "更新", "更新")
                        Case Else
                            If Not IsDBNull(pRow.Item("M_BASICClass")) Then
                                Dim pCreateColumn As New DataGridViewComboBoxColumn

                                pCreateColumn.DataPropertyName = pRow.Item("DataPropertyName")
                                pCreateColumn.DataSource = New DataView(App農地基本台帳.DataMaster.Body, "[Class]='" & pRow.Item("M_BASICClass") & "'", "[ID]", DataViewRowState.CurrentRows)
                                pCreateColumn.ValueMember = "ID"
                                pCreateColumn.DisplayMember = "名称"
                                pCreateColumn.Name = pRow.Item("ColumnName")
                                pCreateColumn.HeaderText = pRow.Item("HeaderText")
                                pCreateColumn.Visible = pRow.Item("Visible")
                                pCreateColumn.DisplayStyleForCurrentCellOnly = True
                                .Columns.Add(pCreateColumn)
                            Else
                                Select Case pRow.Item("DataType").ToString
                                    Case "System.Boolean"
                                        Dim pChkCol As New DataGridViewCheckBoxColumn
                                        pChkCol.Name = pRow.Item("ColumnName")
                                        pChkCol.DataPropertyName = pRow.Item("DataPropertyName")
                                        pChkCol.HeaderText = pRow.Item("HeaderText")
                                        pChkCol.Visible = pRow.Item("Visible")
                                        .Columns.Add(pChkCol)
                                    Case "System.DateTime"
                                        Dim pDTPCol As New DataGridViewDateTimePickerColumn '【保留】HimTools2012.controls.
                                        pDTPCol.Format = "yyyy/MM/dd" '【保留】"gyy/MM/dd"⇒"yyyy/MM/dd"
                                        pDTPCol.Name = pRow.Item("ColumnName")
                                        pDTPCol.DataPropertyName = pRow.Item("DataPropertyName")
                                        pDTPCol.HeaderText = pRow.Item("HeaderText")
                                        pDTPCol.Visible = pRow.Item("Visible")
                                        .Columns.Add(pDTPCol)
                                    Case Else
                                        With .Columns(.Columns.Add(pRow.Item("ColumnName"), pRow.Item("HeaderText")))
                                            .DataPropertyName = pRow.Item("DataPropertyName")
                                            .DefaultCellStyle.Alignment = pRow.Item("TextAlignment")
                                            .Visible = pRow.Item("Visible")
                                        End With
                                End Select
                            End If
                    End Select
                Next
                For Each pRow As DataRowView In New DataView(pDesignMaster, "", "DisplayIndex", DataViewRowState.CurrentRows)
                    With CType(.Columns(pRow.Item("ColumnName")), DataGridViewColumn)
                        .DisplayIndex = pRow.Item("DisplayIndex")
                        '    pAddRow.Item("DataType") = pCol.ValueType.FullName
                    End With
                Next
            Else
                Stop
            End If

            If Not .Columns.Contains("Key") Then
                Dim pKeyColumn As New DataGridViewTextBoxColumn
                pKeyColumn.DataPropertyName = "Key"
                pKeyColumn.Name = "Key"
                pKeyColumn.Visible = False
                .Columns.Add(pKeyColumn)
            End If
            If Not .Columns.Contains("アイコン") Then
                Dim pIconColumn As New DataGridViewTextBoxColumn
                pIconColumn.DataPropertyName = "アイコン"
                pIconColumn.Name = "アイコン"
                pIconColumn.Visible = False
                .Columns.Add(pIconColumn)
            End If
            .ResumeLayout()
        End With
    End Sub

End Class
