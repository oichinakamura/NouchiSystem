Imports System.Reflection

Public Class CDataToExcel
    Inherits HimTools2012.clsAccessor

    Private mvarPath As String = ""
    Private mvarStartDate As DateTime
    Private mvarEndDate As DateTime

    Private mvar農家 As DataTable
    Private mvar農地 As DataTable
    Private mvar転用 As DataTable

    Public excelApp As Object = Nothing
    Public excelBooks As Object = Nothing
    Public excelBook As Object = Nothing
    Public excelSheets As Object = Nothing
    Public excelSheet As Object = Nothing
    Public excelCell As Object = Nothing
    Public Sub New(ByVal sFileName As String)
        With New dlgInputBWDate(SysAD.GetXMLProperty("集計", "", Now.Date))
            If .ShowDialog = DialogResult.OK Then
                mvarPath = System.IO.Path.Combine(SysAD.CustomReportFolder(SysAD.市町村.市町村名), sFileName)
                mvarStartDate = String.Format("#{0}/{1}/{2}#", .StartDate.Month, .StartDate.Day, .StartDate.Year)
                mvarEndDate = String.Format("#{0}/{1}/{2}#", .EndDate.Month, .EndDate.Day, .EndDate.Year)

                Me.Start(True, True)
            End If
        End With
    End Sub

    Public Overrides Sub Execute()
        Try
            If IO.File.Exists(mvarPath) Then
                excelApp = CreateObject("Excel.Application")
                excelBooks = excelApp.Workbooks
                excelBook = excelBooks.Open(mvarPath)
                excelSheets = excelBook.Sheets
                excelSheet = excelSheets(1)
                excelCell = excelSheet.Cells

                OutPutToExcel()

                SaveAndOpen($"利用権等実績({DateTime.Now.ToLongDateString()}).xls")
            Else
                MsgBox("様式が見つかりません。")
            End If
        Catch ex As IO.FileNotFoundException
            MsgBox(ex.Message)
            _releaseObject(excelSheets)
            _releaseObject(excelBooks)
            _releaseObject(excelBook)
            _releaseObject(excelApp)
        End Try
    End Sub

    Private Sub OutPutToExcel()
        mvar農家 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:個人Info]")
        mvar農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:農地Info].* AS D, [D:個人Info].農業改善計画認定 FROM [D:農地Info] LEFT JOIN [D:個人Info] ON [D:農地Info].所有者ID = [D:個人Info].ID;")
        mvar転用 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D_転用農地].* AS D, [D:個人Info].農業改善計画認定 FROM [D_転用農地] LEFT JOIN [D:個人Info] ON [D_転用農地].所有者ID = [D:個人Info].ID;")

        利用権設定()
        農地法３条()
        '一部変更()
        農地法転用()
        非農地買受適格証明()
        貸借終了()
        農地中間管理機構()
    End Sub

    Private Sub 利用権設定()
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(String.Format("SELECT D_申請.ID, D_申請.法令, D_申請.許可年月日, D_申請.農地リスト, D_申請.申請者A, D_申請.申請者B, D_申請.権利種類, D_申請.再設定 FROM D_申請 WHERE (((D_申請.法令) In ({0},{1})) AND ((D_申請.許可年月日)>=#{2}# And (D_申請.許可年月日)<=#{3}#) AND ((D_申請.状態)=2)) ORDER BY D_申請.許可年月日;", 61, 60, mvarStartDate, mvarEndDate))
        Me.Maximum = pTBL.Rows.Count
        Me.Value = 0

        For Each pRow As DataRow In pTBL.Rows
            Me.Value += 1
            Message = "利用権設定情報出力中(" & Me.Value & "/" & pTBL.Rows.Count & ")..."

            Select Case Month(pRow.Item("許可年月日"))
                Case 4 : excelCell(5, 4) = CnvDecimal(Nothing, 5, 4)
                Case 5 : excelCell(12, 4) = CnvDecimal(Nothing, 12, 4)
                Case 6 : excelCell(19, 4) = CnvDecimal(Nothing, 19, 4)
                Case 7 : excelCell(26, 4) = CnvDecimal(Nothing, 26, 4)
                Case 8 : excelCell(33, 4) = CnvDecimal(Nothing, 33, 4)
                Case 9 : excelCell(40, 4) = CnvDecimal(Nothing, 40, 4)
                Case 10 : excelCell(47, 4) = CnvDecimal(Nothing, 47, 4)
                Case 11 : excelCell(54, 4) = CnvDecimal(Nothing, 54, 4)
                Case 12 : excelCell(61, 4) = CnvDecimal(Nothing, 61, 4)
                Case 1 : excelCell(68, 4) = CnvDecimal(Nothing, 68, 4)
                Case 2 : excelCell(75, 4) = CnvDecimal(Nothing, 75, 4)
                Case 3 : excelCell(82, 4) = CnvDecimal(Nothing, 82, 4)
            End Select

            Dim Ars As Object = Split(pRow.Item("農地リスト"), ";")
            For Each Ar As String In Ars
                Dim nID As Integer = GetID(Ar)
                Dim row農地 As DataRow = Find農地Info(nID)
                If row農地 IsNot Nothing Then
                    If Not IsDBNull(pRow.Item("許可年月日")) Then
                        Select Case Month(pRow.Item("許可年月日"))
                            Case 4 : SelectMonth利用権設定(pRow, row農地, 6, 9)
                            Case 5 : SelectMonth利用権設定(pRow, row農地, 13, 9)
                            Case 6 : SelectMonth利用権設定(pRow, row農地, 20, 9)
                            Case 7 : SelectMonth利用権設定(pRow, row農地, 27, 9)
                            Case 8 : SelectMonth利用権設定(pRow, row農地, 34, 9)
                            Case 9 : SelectMonth利用権設定(pRow, row農地, 41, 9)
                            Case 10 : SelectMonth利用権設定(pRow, row農地, 48, 9)
                            Case 11 : SelectMonth利用権設定(pRow, row農地, 55, 9)
                            Case 12 : SelectMonth利用権設定(pRow, row農地, 62, 9)
                            Case 1 : SelectMonth利用権設定(pRow, row農地, 69, 9)
                            Case 2 : SelectMonth利用権設定(pRow, row農地, 76, 9)
                            Case 3 : SelectMonth利用権設定(pRow, row農地, 83, 9)
                        End Select
                    End If
                End If
            Next
        Next

        _releaseObject(excelCell)
        _releaseObject(excelSheet)
    End Sub
    Private Sub SelectMonth利用権設定(ByVal pRow As DataRow, ByVal row農地 As DataRow, ByVal nRow As Integer, ByVal nCol As Integer)
        Select Case Val(pRow.Item("法令").ToString)
            Case enum法令.利用権設定
                Select Case Val(pRow.Item("権利種類").ToString)
                    Case 1
                        SetCell利用権設定(row農地, pRow, nRow, nCol)
                    Case 2
                        SetCell利用権設定(row農地, pRow, nRow, nCol + 3)
                End Select

                Select Case CnvBool(pRow.Item("再設定"))
                    Case True
                        SetCell利用権設定(row農地, pRow, nRow, nCol + 9)
                    Case False
                        SetCell利用権設定(row農地, pRow, nRow, nCol + 6)
                End Select
            Case enum法令.基盤強化法所有権
                SetCell利用権設定(row農地, pRow, nRow, nCol + 12)
        End Select
    End Sub
    Private Sub SetCell利用権設定(ByVal pRow As DataRow, ByVal pRow申請 As DataRow, ByVal nRow As Integer, ByVal nCol As Integer)
        Dim pRow個人 As DataRow = Find個人Info(Val(pRow申請.Item("申請者A").ToString))
        If Not pRow個人 Is Nothing Then
            Select Case Val(pRow.Item("現況地目").ToString)
                Case 10
                    Select Case Val(pRow個人.Item("農業改善計画認定").ToString)
                        Case 1
                            excelCell(nRow, nCol) = CnvDecimal(Nothing, nRow, nCol)
                            excelCell(nRow, nCol + 1) = CnvDecimal(Nothing, nRow, nCol + 1)
                            excelCell(nRow, nCol + 2) = CnvDecimal(pRow.Item("実面積"), nRow, nCol + 2)
                        Case 2
                            excelCell(nRow + 1, nCol) = CnvDecimal(Nothing, nRow + 1, nCol)
                            excelCell(nRow + 1, nCol + 1) = CnvDecimal(Nothing, nRow + 1, nCol + 1)
                            excelCell(nRow + 1, nCol + 2) = CnvDecimal(pRow.Item("実面積"), nRow + 1, nCol + 2)
                    End Select
                Case 20
                    Select Case Val(pRow個人.Item("農業改善計画認定").ToString)
                        Case 1
                            excelCell(nRow + 3, nCol) = CnvDecimal(Nothing, nRow + 3, nCol)
                            excelCell(nRow + 3, nCol + 1) = CnvDecimal(Nothing, nRow + 3, nCol + 1)
                            excelCell(nRow + 3, nCol + 2) = CnvDecimal(pRow.Item("実面積"), nRow + 3, nCol + 2)
                        Case 2
                            excelCell(nRow + 4, nCol) = CnvDecimal(Nothing, nRow + 4, nCol)
                            excelCell(nRow + 4, nCol + 1) = CnvDecimal(Nothing, nRow + 4, nCol + 1)
                            excelCell(nRow + 4, nCol + 2) = CnvDecimal(pRow.Item("実面積"), nRow + 4, nCol + 2)
                    End Select
            End Select
        End If
    End Sub

    Private Sub 農地法３条()
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(String.Format("SELECT D_申請.ID, D_申請.法令, D_申請.許可年月日, D_申請.農地リスト, D_申請.権利種類, D_申請.再設定 FROM D_申請 WHERE (((D_申請.法令) In ({0})) AND ((D_申請.許可年月日)>=#{1}# And (D_申請.許可年月日)<=#{2}#) AND ((D_申請.状態)=2));", "31, 30, 801, 802, 803, 804", mvarStartDate, mvarEndDate))
        Me.Maximum = pTBL.Rows.Count
        Me.Value = 0

        excelSheet = excelSheets(2)
        excelCell = excelSheet.Cells

        For Each pRow As DataRow In pTBL.Rows
            Me.Value += 1
            Message = "農地法３条情報出力中(" & Me.Value & "/" & pTBL.Rows.Count & ")..."

            Select Case Month(pRow.Item("許可年月日"))
                Case 4 : 農地法３条月別件数(pRow, 5, 8)
                Case 5 : 農地法３条月別件数(pRow, 12, 8)
                Case 6 : 農地法３条月別件数(pRow, 19, 8)
                Case 7 : 農地法３条月別件数(pRow, 26, 8)
                Case 8 : 農地法３条月別件数(pRow, 33, 8)
                Case 9 : 農地法３条月別件数(pRow, 40, 8)
                Case 10 : 農地法３条月別件数(pRow, 47, 8)
                Case 11 : 農地法３条月別件数(pRow, 54, 8)
                Case 12 : 農地法３条月別件数(pRow, 61, 8)
                Case 1 : 農地法３条月別件数(pRow, 68, 8)
                Case 2 : 農地法３条月別件数(pRow, 75, 8)
                Case 3 : 農地法３条月別件数(pRow, 82, 8)
            End Select

            Dim Ars As Object = Split(pRow.Item("農地リスト"), ";")
            For Each Ar As String In Ars
                Dim nID As Integer = GetID(Ar)
                Dim row農地 As DataRow = Find農地Info(nID)
                If row農地 IsNot Nothing Then
                    If Not IsDBNull(pRow.Item("許可年月日")) Then
                        Select Case Month(pRow.Item("許可年月日"))
                            Case 4 : SelectMonth農地法３条(pRow, row農地, 5, 9)
                            Case 5 : SelectMonth農地法３条(pRow, row農地, 12, 9)
                            Case 6 : SelectMonth農地法３条(pRow, row農地, 19, 9)
                            Case 7 : SelectMonth農地法３条(pRow, row農地, 26, 9)
                            Case 8 : SelectMonth農地法３条(pRow, row農地, 33, 9)
                            Case 9 : SelectMonth農地法３条(pRow, row農地, 40, 9)
                            Case 10 : SelectMonth農地法３条(pRow, row農地, 47, 9)
                            Case 11 : SelectMonth農地法３条(pRow, row農地, 54, 9)
                            Case 12 : SelectMonth農地法３条(pRow, row農地, 61, 9)
                            Case 1 : SelectMonth農地法３条(pRow, row農地, 68, 9)
                            Case 2 : SelectMonth農地法３条(pRow, row農地, 75, 9)
                            Case 3 : SelectMonth農地法３条(pRow, row農地, 82, 9)
                        End Select
                    End If
                End If
            Next
        Next

        _releaseObject(excelCell)
        _releaseObject(excelSheet)
    End Sub
    Private Sub 農地法３条月別件数(ByVal pRow As DataRow, ByVal nRow As Integer, ByVal nCol As Integer)
        Select Case Val(pRow.Item("法令").ToString)
            Case enum法令.農地法3条所有権 : excelCell(nRow, nCol) = CnvDecimal(Nothing, nRow, nCol)
            Case enum法令.農地法3条耕作権
                Select Case Val(pRow.Item("権利種類").ToString)
                    Case 1 : excelCell(nRow, nCol + 6) = CnvDecimal(Nothing, nRow, nCol + 6)
                    Case 2 : excelCell(nRow, nCol + 9) = CnvDecimal(Nothing, nRow, nCol + 9)
                End Select
            Case enum法令.買受適格耕公, enum法令.買受適格耕競, enum法令.買受適格転公, enum法令.買受適格転競 : excelCell(nRow, nCol + 3) = CnvDecimal(Nothing, nRow, nCol + 3)
        End Select
    End Sub
    Private Sub SelectMonth農地法３条(ByVal pRow As DataRow, ByVal row農地 As DataRow, ByVal nRow As Integer, ByVal nCol As Integer)
        Select Case Val(pRow.Item("法令").ToString)
            Case enum法令.農地法3条所有権 : SetCell農地法３条(row農地, nRow, nCol)
            Case enum法令.農地法3条耕作権
                Select Case Val(pRow.Item("権利種類").ToString)
                    Case 1 : SetCell農地法３条(row農地, nRow, nCol + 6)
                    Case 2 : SetCell農地法３条(row農地, nRow, nCol + 9)
                End Select
            Case enum法令.買受適格耕公, enum法令.買受適格耕競, enum法令.買受適格転公, enum法令.買受適格転競 : SetCell農地法３条(row農地, nRow, nCol + 3)
        End Select
    End Sub
    Private Sub SetCell農地法３条(ByVal pRow As DataRow, ByVal nRow As Integer, ByVal nCol As Integer)
        Select Case Val(pRow.Item("現況地目").ToString)
            Case 10
                excelCell(nRow, nCol) = CnvDecimal(Nothing, nRow, nCol)
                excelCell(nRow, nCol + 1) = CnvDecimal(pRow.Item("実面積"), nRow, nCol + 1)
            Case 20
                excelCell(nRow + 1, nCol) = CnvDecimal(Nothing, nRow + 1, nCol)
                excelCell(nRow + 1, nCol + 1) = CnvDecimal(pRow.Item("実面積"), nRow + 1, nCol + 1)
        End Select
    End Sub

    Private Sub 一部変更()
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(String.Format("SELECT D_申請.ID, D_申請.法令, D_申請.許可年月日, D_申請.農地リスト, D_申請.権利種類, D_申請.再設定 FROM D_申請 WHERE (((D_申請.法令) In ({0})) AND ((D_申請.許可年月日)>=#{1}# And (D_申請.許可年月日)<=#{2}#) AND ((D_申請.状態)=2));", "302", mvarStartDate, mvarEndDate))
        Me.Maximum = pTBL.Rows.Count
        Me.Value = 0

        excelSheet = excelSheets(3)
        excelCell = excelSheet.Cells

        For Each pRow As DataRow In pTBL.Rows
            Me.Value += 1
            Message = "一部変更出力中(" & Me.Value & "/" & pTBL.Rows.Count & ")..."

            Select Case Month(pRow.Item("許可年月日"))
                Case 4 : 一部変更月別件数(pRow, 4, 2)
                Case 5 : 一部変更月別件数(pRow, 5, 2)
                Case 6 : 一部変更月別件数(pRow, 6, 2)
                Case 7 : 一部変更月別件数(pRow, 7, 2)
                Case 8 : 一部変更月別件数(pRow, 8, 2)
                Case 9 : 一部変更月別件数(pRow, 9, 2)
                Case 10 : 一部変更月別件数(pRow, 10, 2)
                Case 11 : 一部変更月別件数(pRow, 11, 2)
                Case 12 : 一部変更月別件数(pRow, 12, 2)
                Case 1 : 一部変更月別件数(pRow, 13, 2)
                Case 2 : 一部変更月別件数(pRow, 14, 2)
                Case 3 : 一部変更月別件数(pRow, 15, 2)
            End Select

            Dim Ars As Object = Split(pRow.Item("農地リスト"), ";")
            For Each Ar As String In Ars
                Dim nID As Integer = GetID(Ar)
                Dim row農地 As DataRow = Find農地Info(nID)
                If row農地 IsNot Nothing Then
                    If Not IsDBNull(pRow.Item("許可年月日")) Then
                        Select Case Month(pRow.Item("許可年月日"))
                            Case 4 : SetCell一部変更(row農地, 4, 3)
                            Case 5 : SetCell一部変更(row農地, 5, 3)
                            Case 6 : SetCell一部変更(row農地, 6, 3)
                            Case 7 : SetCell一部変更(row農地, 7, 3)
                            Case 8 : SetCell一部変更(row農地, 8, 3)
                            Case 9 : SetCell一部変更(row農地, 9, 3)
                            Case 10 : SetCell一部変更(row農地, 10, 3)
                            Case 11 : SetCell一部変更(row農地, 11, 3)
                            Case 12 : SetCell一部変更(row農地, 12, 3)
                            Case 1 : SetCell一部変更(row農地, 13, 3)
                            Case 2 : SetCell一部変更(row農地, 14, 3)
                            Case 3 : SetCell一部変更(row農地, 15, 3)
                        End Select
                    End If
                End If
            Next
        Next

        _releaseObject(excelCell)
        _releaseObject(excelSheet)
    End Sub
    Private Sub 一部変更月別件数(ByVal pRow As DataRow, ByVal nRow As Integer, ByVal nCol As Integer)
        excelCell(nRow, nCol) = CnvDecimal(Nothing, nRow, nCol)
    End Sub
    Private Sub SetCell一部変更(ByVal pRow As DataRow, ByVal nRow As Integer, ByVal nCol As Integer)
        Select Case Val(pRow.Item("現況地目").ToString)
            Case 10
                excelCell(nRow, nCol) = CnvDecimal(Nothing, nRow, nCol)
                excelCell(nRow, nCol + 1) = CnvDecimal(pRow.Item("実面積"), nRow, nCol + 1)
            Case 20
                excelCell(nRow, nCol + 2) = CnvDecimal(Nothing, nRow, nCol + 2)
                excelCell(nRow, nCol + 3) = CnvDecimal(pRow.Item("実面積"), nRow, nCol + 3)
            Case -301, -302, -303
                excelCell(nRow, nCol + 4) = CnvDecimal(Nothing, nRow, nCol + 4)
                excelCell(nRow, nCol + 5) = CnvDecimal(pRow.Item("実面積"), nRow, nCol + 5)
            Case Else
                excelCell(nRow, nCol + 6) = CnvDecimal(Nothing, nRow, nCol + 6)
                excelCell(nRow, nCol + 7) = CnvDecimal(pRow.Item("実面積"), nRow, nCol + 7)
        End Select
    End Sub

    Private Sub 農地法転用()
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(String.Format("SELECT D_申請.ID, D_申請.法令, D_申請.許可年月日, D_申請.農地リスト, D_申請.申請理由A, D_申請.権利種類, D_申請.再設定 FROM D_申請 WHERE (((D_申請.法令) In ({0})) AND ((D_申請.許可年月日)>=#{1}# And (D_申請.許可年月日)<=#{2}#) AND ((D_申請.状態)=2));", "40,42,50,51,52", mvarStartDate, mvarEndDate))
        Me.Maximum = pTBL.Rows.Count
        Me.Value = 0

        excelSheet = excelSheets(4)
        excelCell = excelSheet.Cells

        For Each pRow As DataRow In pTBL.Rows
            Me.Value += 1
            Message = "農地法転用出力中(" & Me.Value & "/" & pTBL.Rows.Count & ")..."

            Select Case Month(pRow.Item("許可年月日"))
                Case 4 : 農地法転用月別件数(pRow, 5, 2)
                Case 5 : 農地法転用月別件数(pRow, 6, 2)
                Case 6 : 農地法転用月別件数(pRow, 7, 2)
                Case 7 : 農地法転用月別件数(pRow, 8, 2)
                Case 8 : 農地法転用月別件数(pRow, 9, 2)
                Case 9 : 農地法転用月別件数(pRow, 10, 2)
                Case 10 : 農地法転用月別件数(pRow, 11, 2)
                Case 11 : 農地法転用月別件数(pRow, 12, 2)
                Case 12 : 農地法転用月別件数(pRow, 13, 2)
                Case 1 : 農地法転用月別件数(pRow, 14, 2)
                Case 2 : 農地法転用月別件数(pRow, 15, 2)
                Case 3 : 農地法転用月別件数(pRow, 16, 2)
            End Select

            Dim Ars As Object = Split(pRow.Item("農地リスト"), ";")
            For Each Ar As String In Ars
                Dim nID As Integer = GetID(Ar)
                Dim row農地 As DataRow = Find農地Info(nID)
                If row農地 IsNot Nothing Then
                    If Not IsDBNull(pRow.Item("許可年月日")) Then
                        Select Case Month(pRow.Item("許可年月日"))
                            Case 4 : SelectMonth農地法転用(pRow, row農地, 5, 3)
                            Case 5 : SelectMonth農地法転用(pRow, row農地, 6, 3)
                            Case 6 : SelectMonth農地法転用(pRow, row農地, 7, 3)
                            Case 7 : SelectMonth農地法転用(pRow, row農地, 8, 3)
                            Case 8 : SelectMonth農地法転用(pRow, row農地, 9, 3)
                            Case 9 : SelectMonth農地法転用(pRow, row農地, 10, 3)
                            Case 10 : SelectMonth農地法転用(pRow, row農地, 11, 3)
                            Case 11 : SelectMonth農地法転用(pRow, row農地, 12, 3)
                            Case 12 : SelectMonth農地法転用(pRow, row農地, 13, 3)
                            Case 1 : SelectMonth農地法転用(pRow, row農地, 14, 3)
                            Case 2 : SelectMonth農地法転用(pRow, row農地, 15, 3)
                            Case 3 : SelectMonth農地法転用(pRow, row農地, 16, 3)
                        End Select
                    End If
                End If
            Next
        Next

        _releaseObject(excelCell)
        _releaseObject(excelSheet)
    End Sub
    Private Sub 農地法転用月別件数(ByVal pRow As DataRow, ByVal nRow As Integer, ByVal nCol As Integer)
        Select Case Val(pRow.Item("法令").ToString)
            Case 40, 42
                excelCell(nRow, nCol) = CnvDecimal(Nothing, nRow, nCol)
            Case 50, 51, 52
                excelCell(nRow + 17, nCol) = CnvDecimal(Nothing, nRow + 17, nCol)
        End Select
    End Sub
    Private Sub SelectMonth農地法転用(ByVal pRow As DataRow, ByVal row農地 As DataRow, ByVal nRow As Integer, ByVal nCol As Integer)
        Select Case Val(pRow.Item("法令").ToString)
            Case 40, 42
                If InStr(pRow.Item("申請理由A").ToString, "宅") > 0 Then
                    SetCell農地法転用(row農地, nRow, nCol)
                ElseIf InStr(pRow.Item("申請理由A").ToString, "畜舎") > 0 Then
                    SetCell農地法転用(row農地, nRow, nCol + 4)
                ElseIf InStr(pRow.Item("申請理由A").ToString, "道路") > 0 Then
                    SetCell農地法転用(row農地, nRow, nCol + 8)
                Else
                    SetCell農地法転用(row農地, nRow, nCol + 12)
                End If
            Case 50, 51, 52
                If InStr(pRow.Item("申請理由A").ToString, "農家住宅") > 0 Then
                    SetCell農地法転用(row農地, nRow + 17, nCol + 4)
                ElseIf InStr(pRow.Item("申請理由A").ToString, "宅") > 0 Then
                    SetCell農地法転用(row農地, nRow + 17, nCol)
                ElseIf InStr(pRow.Item("申請理由A").ToString, "畜舎") > 0 Then
                    SetCell農地法転用(row農地, nRow + 17, nCol + 8)
                Else
                    SetCell農地法転用(row農地, nRow + 17, nCol + 12)
                End If
        End Select
    End Sub
    Private Sub SetCell農地法転用(ByVal pRow As DataRow, ByVal nRow As Integer, ByVal nCol As Integer)
        Select Case Val(pRow.Item("現況地目").ToString)
            Case 10
                excelCell(nRow, nCol) = CnvDecimal(Nothing, nRow, nCol)
                excelCell(nRow, nCol + 1) = CnvDecimal(pRow.Item("実面積"), nRow, nCol + 1)
            Case 20
                excelCell(nRow, nCol + 2) = CnvDecimal(Nothing, nRow, nCol + 2)
                excelCell(nRow, nCol + 3) = CnvDecimal(pRow.Item("実面積"), nRow, nCol + 3)
        End Select
    End Sub

    Private Sub 非農地買受適格証明()
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(String.Format("SELECT D_申請.ID, D_申請.法令, D_申請.許可年月日, D_申請.農地リスト, D_申請.権利種類, D_申請.再設定 FROM D_申請 WHERE (((D_申請.法令) In ({0})) AND ((D_申請.許可年月日)>=#{1}# And (D_申請.許可年月日)<=#{2}#) AND ((D_申請.状態)=2));", "600,602,801,802,803,804", mvarStartDate, mvarEndDate))
        Me.Maximum = pTBL.Rows.Count
        Me.Value = 0

        excelSheet = excelSheets(6)
        excelCell = excelSheet.Cells

        For Each pRow As DataRow In pTBL.Rows
            Me.Value += 1
            Message = "非農地/買受適格証明出力中(" & Me.Value & "/" & pTBL.Rows.Count & ")..."

            Select Case Month(pRow.Item("許可年月日"))
                Case 4 : 非農地買受適格証明月別件数(pRow, 4, 3)
                Case 5 : 非農地買受適格証明月別件数(pRow, 5, 3)
                Case 6 : 非農地買受適格証明月別件数(pRow, 6, 3)
                Case 7 : 非農地買受適格証明月別件数(pRow, 7, 3)
                Case 8 : 非農地買受適格証明月別件数(pRow, 8, 3)
                Case 9 : 非農地買受適格証明月別件数(pRow, 9, 3)
                Case 10 : 非農地買受適格証明月別件数(pRow, 10, 3)
                Case 11 : 非農地買受適格証明月別件数(pRow, 11, 3)
                Case 12 : 非農地買受適格証明月別件数(pRow, 12, 3)
                Case 1 : 非農地買受適格証明月別件数(pRow, 13, 3)
                Case 2 : 非農地買受適格証明月別件数(pRow, 14, 3)
                Case 3 : 非農地買受適格証明月別件数(pRow, 15, 3)
            End Select

            Dim Ars As Object = Split(pRow.Item("農地リスト"), ";")
            For Each Ar As String In Ars
                Dim nID As Integer = GetID(Ar)
                Dim row農地 As DataRow = Find農地Info(nID)
                If row農地 IsNot Nothing Then
                    If Not IsDBNull(pRow.Item("許可年月日")) Then
                        Select Case Month(pRow.Item("許可年月日"))
                            Case 4 : SelectMonth非農地買受適格証明(pRow, row農地, 4, 6)
                            Case 5 : SelectMonth非農地買受適格証明(pRow, row農地, 5, 6)
                            Case 6 : SelectMonth非農地買受適格証明(pRow, row農地, 6, 6)
                            Case 7 : SelectMonth非農地買受適格証明(pRow, row農地, 7, 6)
                            Case 8 : SelectMonth非農地買受適格証明(pRow, row農地, 8, 6)
                            Case 9 : SelectMonth非農地買受適格証明(pRow, row農地, 9, 6)
                            Case 10 : SelectMonth非農地買受適格証明(pRow, row農地, 10, 6)
                            Case 11 : SelectMonth非農地買受適格証明(pRow, row農地, 11, 6)
                            Case 12 : SelectMonth非農地買受適格証明(pRow, row農地, 12, 6)
                            Case 1 : SelectMonth非農地買受適格証明(pRow, row農地, 13, 6)
                            Case 2 : SelectMonth非農地買受適格証明(pRow, row農地, 14, 6)
                            Case 3 : SelectMonth非農地買受適格証明(pRow, row農地, 15, 6)
                        End Select
                    End If
                End If
            Next
        Next

        _releaseObject(excelCell)
        _releaseObject(excelSheet)
    End Sub
    Private Sub 非農地買受適格証明月別件数(ByVal pRow As DataRow, ByVal nRow As Integer, ByVal nCol As Integer)
        Select Case Val(pRow.Item("法令").ToString)
            Case 600, 602
                excelCell(nRow, nCol) = CnvDecimal(Nothing, nRow, nCol)
            Case 801, 802, 803, 804
                excelCell(nRow, nCol + 28) = CnvDecimal(Nothing, nRow, nCol + 28)
        End Select
    End Sub
    Private Sub SelectMonth非農地買受適格証明(ByVal pRow As DataRow, ByVal row農地 As DataRow, ByVal nRow As Integer, ByVal nCol As Integer)
        Select Case Val(pRow.Item("法令").ToString)
            Case 600, 602
                SetCell非農地買受適格証明(row農地, nRow, nCol)
            Case 801, 802, 803, 804
                SetCell非農地買受適格証明(row農地, nRow, nCol + 28)
        End Select
    End Sub
    Private Sub SetCell非農地買受適格証明(ByVal pRow As DataRow, ByVal nRow As Integer, ByVal nCol As Integer)
        Select Case Val(pRow.Item("現況地目").ToString)
            Case 10
                excelCell(nRow, nCol) = CnvDecimal(Nothing, nRow, nCol)
                excelCell(nRow, nCol + 1) = CnvDecimal(pRow.Item("実面積"), nRow, nCol + 1)
            Case 20
                excelCell(nRow, nCol + 2) = CnvDecimal(Nothing, nRow, nCol + 2)
                excelCell(nRow, nCol + 3) = CnvDecimal(pRow.Item("実面積"), nRow, nCol + 3)
        End Select
    End Sub

    Private nDay As String = ""
    Private Sub 貸借終了()
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(String.Format("SELECT D_申請.ID, D_申請.法令, D_申請.許可年月日, D_申請.農地リスト, D_申請.権利種類, D_申請.再設定 FROM D_申請 WHERE (((D_申請.法令) In ({0})) AND ((D_申請.許可年月日)>=#{1}# And (D_申請.許可年月日)<=#{2}#) AND ((D_申請.状態)=2)) ORDER BY D_申請.許可年月日;", "31,61", mvarStartDate, mvarEndDate))
        Me.Maximum = pTBL.Rows.Count
        Me.Value = 0

        excelSheet = excelSheets(7)
        excelCell = excelSheet.Cells

        Dim nAdd As Integer = 0
        For Each pRow As DataRow In pTBL.Rows
            Me.Value += 1
            Message = "貸借終了出力中(" & Me.Value & "/" & pTBL.Rows.Count & ")..."

            Select Case Month(pRow.Item("許可年月日"))
                Case 4 : 貸借終了月別件数(pRow, 4, 3, nAdd)
                Case 5 : 貸借終了月別件数(pRow, 6, 3, nAdd)
                Case 6 : 貸借終了月別件数(pRow, 8, 3, nAdd)
                Case 7 : 貸借終了月別件数(pRow, 10, 3, nAdd)
                Case 8 : 貸借終了月別件数(pRow, 12, 3, nAdd)
                Case 9 : 貸借終了月別件数(pRow, 14, 3, nAdd)
                Case 10 : 貸借終了月別件数(pRow, 16, 3, nAdd)
                Case 11 : 貸借終了月別件数(pRow, 18, 3, nAdd)
                Case 12 : 貸借終了月別件数(pRow, 20, 3, nAdd)
                Case 1 : 貸借終了月別件数(pRow, 22, 3, nAdd)
                Case 2 : 貸借終了月別件数(pRow, 24, 3, nAdd)
                Case 3 : 貸借終了月別件数(pRow, 26, 3, nAdd)
            End Select
            nDay = DateAndTime.Month(pRow.Item("許可年月日")) & "/" & DateAndTime.Day(pRow.Item("許可年月日"))

            Dim Ars As Object = Split(pRow.Item("農地リスト"), ";")
            For Each Ar As String In Ars
                Dim nID As Integer = GetID(Ar)
                Dim row農地 As DataRow = Find農地Info(nID)
                If row農地 IsNot Nothing Then
                    If Not IsDBNull(pRow.Item("許可年月日")) Then
                        Select Case Month(pRow.Item("許可年月日"))
                            Case 4 : SetCell貸借終了(row農地, 4, 7, nAdd)
                            Case 5 : SetCell貸借終了(row農地, 6, 7, nAdd)
                            Case 6 : SetCell貸借終了(row農地, 8, 7, nAdd)
                            Case 7 : SetCell貸借終了(row農地, 10, 7, nAdd)
                            Case 8 : SetCell貸借終了(row農地, 12, 7, nAdd)
                            Case 9 : SetCell貸借終了(row農地, 14, 7, nAdd)
                            Case 10 : SetCell貸借終了(row農地, 16, 7, nAdd)
                            Case 11 : SetCell貸借終了(row農地, 18, 7, nAdd)
                            Case 12 : SetCell貸借終了(row農地, 20, 7, nAdd)
                            Case 1 : SetCell貸借終了(row農地, 22, 7, nAdd)
                            Case 2 : SetCell貸借終了(row農地, 24, 7, nAdd)
                            Case 3 : SetCell貸借終了(row農地, 26, 7, nAdd)
                        End Select
                    End If
                End If
            Next
        Next

        _releaseObject(excelCell)
        _releaseObject(excelSheet)
    End Sub
    Private Sub 貸借終了月別件数(ByVal pRow As DataRow, ByVal nRow As Integer, ByVal nCol As Integer, ByRef nAdd As Integer)
        Dim sDay As String = DateAndTime.Month(pRow.Item("許可年月日")) & "/" & DateAndTime.Day(pRow.Item("許可年月日"))

        If nDay <> "" AndAlso sDay <> "" AndAlso (Split(nDay, "/")(0) = 3 AndAlso Split(sDay, "/")(0) = 4) Then
            nAdd = 0
        End If

        If nDay <> "" AndAlso sDay <> nDay AndAlso Split(nDay, "/")(0) = Split(sDay, "/")(0) Then
            Dim R As Object = excelSheet.Range(nRow + nAdd + 1 & ":" & nRow + nAdd + 1)
            R.Copy()
            R.Insert()
            _releaseObject(R)

            nAdd += 1
        End If
        nRow += nAdd

        excelCell(nRow, nCol) = sDay
        excelCell(nRow, nCol + 1) = CnvDecimal(Nothing, nRow, nCol + 1)
    End Sub
    Private Sub SetCell貸借終了(ByVal pRow As DataRow, ByVal nRow As Integer, ByVal nCol As Integer, ByVal nAdd As Integer)
        nRow += nAdd
        Select Case Val(pRow.Item("現況地目").ToString)
            Case 10
                excelCell(nRow, nCol) = CnvDecimal(Nothing, nRow, nCol)
                excelCell(nRow, nCol + 3) = CnvDecimal(pRow.Item("実面積"), nRow, nCol + 3)
            Case 20
                excelCell(nRow, nCol + 10) = CnvDecimal(Nothing, nRow, nCol + 10)
                excelCell(nRow, nCol + 13) = CnvDecimal(pRow.Item("実面積"), nRow, nCol + 13)
        End Select
    End Sub

    Private Sub 農地中間管理機構()
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(String.Format("SELECT D_申請.ID, D_申請.法令, D_申請.許可年月日, D_申請.農地リスト, D_申請.権利種類, D_申請.再設定 FROM D_申請 WHERE (((D_申請.法令) In ({0})) AND ((D_申請.許可年月日)>=#{1}# And (D_申請.許可年月日)<=#{2}#) AND ((D_申請.状態)=2) AND ((D_申請.経由法人ID)={3}));", "61", mvarStartDate, mvarEndDate, Val(SysAD.DB(sLRDB).DBProperty("中間管理機構ID"))))
        Me.Maximum = pTBL.Rows.Count
        Me.Value = 0

        excelSheet = excelSheets(9)
        excelCell = excelSheet.Cells

        For Each pRow As DataRow In pTBL.Rows
            Me.Value += 1
            Message = "農地中間管理機構出力中(" & Me.Value & "/" & pTBL.Rows.Count & ")..."

            Select Case Month(pRow.Item("許可年月日"))
                Case 4 : 農地中間管理機構月別件数(pRow, 4, 2)
                Case 5 : 農地中間管理機構月別件数(pRow, 5, 2)
                Case 6 : 農地中間管理機構月別件数(pRow, 6, 2)
                Case 7 : 農地中間管理機構月別件数(pRow, 7, 2)
                Case 8 : 農地中間管理機構月別件数(pRow, 8, 2)
                Case 9 : 農地中間管理機構月別件数(pRow, 9, 2)
                Case 10 : 農地中間管理機構月別件数(pRow, 10, 2)
                Case 11 : 農地中間管理機構月別件数(pRow, 11, 2)
                Case 12 : 農地中間管理機構月別件数(pRow, 12, 2)
                Case 1 : 農地中間管理機構月別件数(pRow, 13, 2)
                Case 2 : 農地中間管理機構月別件数(pRow, 14, 2)
                Case 3 : 農地中間管理機構月別件数(pRow, 15, 2)
            End Select

            Dim Ars As Object = Split(pRow.Item("農地リスト"), ";")
            For Each Ar As String In Ars
                Dim nID As Integer = GetID(Ar)
                Dim row農地 As DataRow = Find農地Info(nID)
                If row農地 IsNot Nothing Then
                    If Not IsDBNull(pRow.Item("許可年月日")) Then
                        Select Case Month(pRow.Item("許可年月日"))
                            Case 4 : SetCell農地中間管理機構(row農地, 4, 3)
                            Case 5 : SetCell農地中間管理機構(row農地, 5, 3)
                            Case 6 : SetCell農地中間管理機構(row農地, 6, 3)
                            Case 7 : SetCell農地中間管理機構(row農地, 7, 3)
                            Case 8 : SetCell農地中間管理機構(row農地, 8, 3)
                            Case 9 : SetCell農地中間管理機構(row農地, 9, 3)
                            Case 10 : SetCell農地中間管理機構(row農地, 10, 3)
                            Case 11 : SetCell農地中間管理機構(row農地, 11, 3)
                            Case 12 : SetCell農地中間管理機構(row農地, 12, 3)
                            Case 1 : SetCell農地中間管理機構(row農地, 13, 3)
                            Case 2 : SetCell農地中間管理機構(row農地, 14, 3)
                            Case 3 : SetCell農地中間管理機構(row農地, 15, 3)
                        End Select
                    End If
                End If
            Next
        Next

        _releaseObject(excelCell)
        _releaseObject(excelSheet)
    End Sub
    Private Sub 農地中間管理機構月別件数(ByVal pRow As DataRow, ByVal nRow As Integer, ByVal nCol As Integer)
        excelCell(nRow, nCol) = CnvDecimal(Nothing, nRow, nCol)
    End Sub
    Private Sub SetCell農地中間管理機構(ByVal pRow As DataRow, ByVal nRow As Integer, ByVal nCol As Integer)
        Select Case Val(pRow.Item("現況地目").ToString)
            Case 10
                excelCell(nRow, nCol) = CnvDecimal(Nothing, nRow, nCol)
                excelCell(nRow, nCol + 1) = CnvDecimal(pRow.Item("実面積"), nRow, nCol + 1)
            Case 20
                excelCell(nRow, nCol + 2) = CnvDecimal(Nothing, nRow, nCol + 2)
                excelCell(nRow, nCol + 3) = CnvDecimal(pRow.Item("実面積"), nRow, nCol + 3)
            Case -301, -302, -303
                excelCell(nRow, nCol + 4) = CnvDecimal(Nothing, nRow, nCol + 4)
                excelCell(nRow, nCol + 5) = CnvDecimal(pRow.Item("実面積"), nRow, nCol + 5)
            Case Else
                excelCell(nRow, nCol + 6) = CnvDecimal(Nothing, nRow, nCol + 6)
                excelCell(nRow, nCol + 7) = CnvDecimal(pRow.Item("実面積"), nRow, nCol + 7)
        End Select
    End Sub

    Private Function GetID(ByVal sID As String) As Integer
        Return Val(Replace(Replace(sID, "転用農地.", ""), "農地.", ""))
    End Function

    Private Function Find個人Info(ByVal nID As Integer) As DataRow
        Dim p個人Row As DataRow() = mvar農家.Select("[ID]=" & nID)
        If Not IsDBNull(p個人Row) Then
            If p個人Row.Count > 0 Then
                Return p個人Row(0)
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If
    End Function

    Private Function Find農地Info(ByVal nID As Integer) As DataRow
        Dim p農地Row As DataRow() = mvar農地.Select("[ID]=" & nID)
        If p農地Row.Count > 0 Then
            Return p農地Row(0)
        End If

        Dim p転用Row As DataRow() = mvar転用.Select("[ID]=" & nID)
        If p転用Row.Count > 0 Then
            Return p転用Row(0)
        Else
            Return Nothing
        End If
    End Function

    Private Function CnvDecimal(ByVal pVal As Object, ByVal nRow As Integer, ByVal nCol As Integer) As Decimal
        If pVal IsNot Nothing Then
            If excelCell(nRow, nCol).Value IsNot Nothing Then
                Return Val(excelCell(nRow, nCol).Value) + Val(pVal.ToString)
            Else
                Return Val(pVal.ToString)
            End If
        Else
            If excelCell(nRow, nCol).Value IsNot Nothing Then
                Return Val(excelCell(nRow, nCol).Value) + 1
            Else
                Return 1
            End If
        End If
    End Function

    Private Function CnvBool(ByVal pVal As Object) As Boolean
        If Not IsDBNull(pVal) Then
            Return pVal
        Else
            Return False
        End If
    End Function

    Public Sub SaveAndOpen(ByVal sFileName As String)
        Dim FillFullPath As String = System.IO.Path.Combine(System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), sFileName)
        excelApp.DisplayAlerts = False
        excelBook.SaveAs(FillFullPath)
        excelApp.DisplayAlerts = True

        excelBook.Close()
        _releaseObject(excelSheets)
        _releaseObject(excelBooks)
        _releaseObject(excelBook)
        _releaseObject(excelApp)
        System.Diagnostics.Process.Start(FillFullPath)
    End Sub

    Public Sub _releaseObject(ByRef obj As Object)
        Try
            Runtime.InteropServices.Marshal.FinalReleaseComObject(obj)
            obj = Nothing
        Catch ex As IO.FileNotFoundException
            obj = Nothing
            Console.WriteLine("オブジェクトを解放できない" + ex.ToString())
        Finally
            GC.Collect()
        End Try
    End Sub


End Class
