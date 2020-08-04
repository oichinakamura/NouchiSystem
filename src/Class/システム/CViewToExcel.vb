
Imports System.Windows.Forms

Public Class CViewToExcel
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Private mvarSQL As String
    Private mvarTable As DataTable
    Private mvarDir As String = ""
    Private mvarFileName As String = ""
    Private WithEvents mvarGridView As DataGridLocal
    Private WithEvents mvarExcelList As ToolStripButton


    Public Sub New(ByVal Name As String, ByVal sTitle As String, ByVal sSQL As String, sOutPutDir As String, sFileName As String)
        MyBase.New(True)
        Me.Name = Name
        Me.Text = sTitle
        mvarSQL = sSQL
        mvarDir = sOutPutDir
        mvarFileName = sFileName
        mvarGridView = New DataGridLocal
        mvarTable = SysAD.DB("LRDB").GetTableBySqlSelect(sSQL)
        mvarGridView.DataSource = mvarTable
        Me.ControlPanel.Add(mvarGridView)

        mvarExcelList = New ToolStripButton("エクセル出力")
        Me.ToolStrip.Items.Add(mvarExcelList)

    End Sub

    Public Sub New(ByVal Name As String, ByVal sTitle As String, ByVal pTable As DataTable, sOutPutDir As String, sFileName As String)
        MyBase.New(True)
        Me.Name = Name
        Me.Text = sTitle
        mvarTable = pTable
        mvarDir = sOutPutDir
        mvarFileName = sFileName

        mvarGridView = New DataGridLocal
        mvarGridView.DataSource = mvarTable
        Me.ControlPanel.Add(mvarGridView)

        mvarExcelList = New ToolStripButton("エクセル出力")
        Me.ToolStrip.Items.Add(mvarExcelList)

    End Sub



    Private Class DataGridLocal
        Inherits DataGridView

        Public Sub New()
            MyBase.New()
            Me.Dock = DockStyle.Fill
            Me.AllowUserToAddRows = False

        End Sub

        Private Sub DataGridLocal_DataSourceChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.DataSourceChanged
            For Each pCol As DataGridViewColumn In Me.Columns
                Select Case pCol.ValueType.FullName

                    Case "System.Double"
                        pCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    Case Else
                        Stop
                End Select
            Next

        End Sub
    End Class

    Private Sub mvarExcelList_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarExcelList.Click
        SaveGridViewToExcel(mvarGridView, mvarDir & If(mvarDir.EndsWith("\") OrElse mvarFileName.StartsWith("\"), "", "\") & mvarFileName)
    End Sub
    Public Sub SaveGridViewToExcel(ByVal pGridView As DataGridView, ByVal sFileName As String)
        Dim Data(,) As Object
        ReDim Data(pGridView.Rows.Count, pGridView.Columns.Count - 1)

        Dim excelApp As Object = CreateObject("Excel.Application")

        Dim bk As Object = excelApp.Workbooks.Add()
        Dim sheet As Object = bk.Worksheets(1)

        Dim range As Object = Nothing
        Dim range1 As Object = Nothing
        Dim range2 As Object = Nothing

        Dim startX As Integer = 1
        Dim startY As Integer = 1

        For nCol As Integer = 0 To pGridView.Columns.Count - 1
            Data(0, nCol) = pGridView.Columns(nCol).HeaderText
        Next

        For nCol As Integer = 0 To pGridView.Columns.Count - 1
            For nRow As Integer = 0 To pGridView.Rows.Count - 1
                Data(nRow + 1, nCol) = pGridView.Item(nCol, nRow).FormattedValue
            Next
        Next

        '始点
        range1 = sheet.Cells(startY, startX)
        '終点
        range2 = sheet.Cells(startY + UBound(Data), startX + UBound(Data, 2))
        'セル範囲
        range = sheet.Range(range1, range2)
        '貼り付け
        range.Value = Data

        'bk.SaveAs(sFileName, -4143, "", "", False, False)


        With range.Borders(7)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With range.Borders(8)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With range.Borders(9)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With range.Borders(10)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With range.Borders(11)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With range.Borders(12)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With

        For nCol As Integer = 0 To pGridView.Columns.Count - 1
            With sheet.Range(sheet.Cells(1, nCol + 1), sheet.Cells(pGridView.Rows.Count, nCol + 1))
                If pGridView.Columns(nCol).Visible Then
                    sheet.Columns(nCol + 1).EntireColumn.AutoFit()
                    Select Case mvarTable.Columns(pGridView.Columns(nCol).DataPropertyName).DataType.FullName
                        Case "System.DateTime"
                            .NumberFormatLocal = "[$-411]ge.m.d;@"
                        Case "System.Decimal"
                        Case "System.Int32"
                        Case "System.String"
                            .NumberFormatLocal = "@"
                        Case Else
                    End Select
                Else
                    sheet.Columns(nCol + 1).Hidden = True
                End If
            End With
        Next

        'Excelを表示する

        excelApp.Visible = True
        'sheet.Activate()
        'sheet.range("A1").Select()

        'excelApp.ActiveWindow.FreezePanes = True

        System.Runtime.InteropServices.Marshal.ReleaseComObject(range)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(range2)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(range1)


        System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(bk)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)
    End Sub

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.NoMessage
        End Get
    End Property
End Class
