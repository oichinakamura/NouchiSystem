

Public Class C農地賃借料データ
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Private mvarInsideTabCtrl As HimTools2012.controls.TabControlBase
    Private WithEvents mvarMakeExcel As ToolStripButton

    Public Sub New()
        MyBase.New(True, True, "農地賃借料データ(過去3年分)", "農地賃借料データ(過去3年分)")
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT CInt([大字ID]/100) AS 旧町ID, [D:農地Info].大字ID, V_大字.大字, [D:農地Info].地番, [D:農地Info].小作開始年月日, [D:農地Info].小作終了年月日, IIf([田面積]>0,[小作料],0) AS 田小作料, IIf([畑面積]>0,[小作料],0) AS 畑小作料, IIf([樹園地]>0,[小作料],0) AS 茶畑小作料, '' AS 備考 FROM [D:農地Info] INNER JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID WHERE ((([D:農地Info].自小作別)>0) AND (([D:農地Info].小作開始年月日)>#12/31/2011#) AND (([D:農地Info].小作地適用法)=2) AND (([D:農地Info].小作形態)=1) AND (([D:農地Info].小作料単位) Like '%円%')) ORDER BY CInt([大字ID]/100), [D:農地Info].大字ID, Val([地番]);")
        mvarInsideTabCtrl = New HimTools2012.controls.TabControlBase

        Me.ControlPanel.Add(mvarInsideTabCtrl)


        For Each p大字 As DataRow In pTBL.Rows
            If Not mvarInsideTabCtrl.TabPages.ContainsKey("大字." & p大字.Item("大字ID")) Then
                Dim pGrid As New HimTools2012.controls.DataGridViewWithDataView

                pGrid.AllowUserToAddRows = False
                pGrid.SetDataView(pTBL, "[大字ID]=" & p大字.Item("大字ID"), "")
                mvarInsideTabCtrl.AddNewPage(pGrid, "大字." & p大字.Item("大字ID"), p大字.Item("大字"), False)
            End If
        Next

        mvarMakeExcel = New ToolStripButton
        mvarMakeExcel.Text = "エクセル作成"
        Me.ToolStrip.Items.Add(mvarMakeExcel)

    End Sub

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.CloseOK
        End Get
    End Property

    Private Sub mvarMakeExcel_Click(sender As Object, e As System.EventArgs) Handles mvarMakeExcel.Click
        Dim excelApp As Object = CreateObject("Excel.Application")
        Dim bk As Object = excelApp.Workbooks.Add()

        Dim pPageNo As Integer = 0

        For Each pPage As TabPage In mvarInsideTabCtrl.TabPages
            pPageNo += 1
            If bk.Worksheets.count >= pPageNo Then
            Else
                bk.Worksheets.add()
            End If
        Next

        pPageNo = 0
        For Each pPage As TabPage In mvarInsideTabCtrl.TabPages
            pPageNo += 1

            Dim pGrid As HimTools2012.controls.DataGridViewWithDataView = CType(pPage, HimTools2012.controls.CTabPageWithToolStrip).ControlPanel(0)

            Dim sheet As Object
            sheet = bk.Worksheets(pPageNo)
            sheet.name = pPage.Text

            With pGrid
                Dim Data(,) As Object
                ReDim Data(.Rows.Count, .Columns.Count - 1)

                Dim range As Object = Nothing
                Dim range1 As Object = Nothing
                Dim range2 As Object = Nothing

                Dim startX As Integer = 1
                Dim startY As Integer = 1

                For nCol As Integer = 0 To .Columns.Count - 1
                    If Not .Columns(nCol).DataPropertyName = "" Then
                        Data(0, nCol) = .Columns(nCol).HeaderText
                    End If
                Next

                For nCol As Integer = 0 To .Columns.Count - 1
                    If Not .Columns(nCol).DataPropertyName = "" Then
                        For nRow As Integer = 0 To .Rows.Count - 1
                            If .Columns(nCol).ValueType IsNot Nothing Then
                                Select Case .Columns(nCol).ValueType.FullName
                                    Case "System.Int16", "System.Int32"
                                        Data(nRow + 1, nCol) = .Item(nCol, nRow).FormattedValue
                                    Case "System.Decimal"
                                        Data(nRow + 1, nCol) = .Item(nCol, nRow).Value
                                    Case "System.String", "System.DateTime"
                                        Data(nRow + 1, nCol) = .Item(nCol, nRow).FormattedValue
                                    Case Else
                                        Data(nRow + 1, nCol) = .Item(nCol, nRow).FormattedValue
                                End Select
                            End If
                        Next
                    End If
                Next

                '始点
                range1 = sheet.Cells(startY, startX)
                '終点
                range2 = sheet.Cells(startY + UBound(Data), startX + UBound(Data, 2))
                'セル範囲
                range = sheet.Range(range1, range2)
                '貼り付け
                range.Value = Data

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

                For nCol As Integer = 0 To .Columns.Count - 1
                    If pGrid.Columns(nCol).Visible AndAlso Not pGrid.Columns(nCol).DataPropertyName = "" Then
                        sheet.Columns(nCol + 1).EntireColumn.AutoFit()
                        If pGrid.Columns(nCol) IsNot Nothing AndAlso pGrid.DataView.Table.Columns(pGrid.Columns(nCol).DataPropertyName) IsNot Nothing Then
                            Dim range3 As Object = Nothing
                            range3 = sheet.Range(sheet.Cells(1, nCol + 1), sheet.Cells(.Rows.Count, nCol + 1))
                            Select Case pGrid.DataView.Table.Columns(pGrid.Columns(nCol).DataPropertyName).DataType.FullName
                                Case "System.DateTime"
                                    range3.NumberFormatLocal = "[$-411]ge.m.d;@"
                                Case "System.Decimal"
                                Case "System.Int32"
                                Case "System.String"
                                    range3.NumberFormatLocal = "@"
                                Case Else
                            End Select

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(range3)
                        End If
                    Else
                        sheet.Columns(nCol + 1).Hidden = True
                    End If
                Next

                'Excelを表示する




                System.Runtime.InteropServices.Marshal.ReleaseComObject(range)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range2)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range1)

                System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet)
            End With
        Next
        excelApp.Visible = True

        System.Runtime.InteropServices.Marshal.ReleaseComObject(bk)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)

    End Sub
End Class
