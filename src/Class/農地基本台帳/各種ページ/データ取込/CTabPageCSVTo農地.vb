Imports HimTools2012.CommonFunc

Public Class CTabPageCSVTo農地
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Private mvarSP As SynchronizeTwinGridView
    Private WithEvents mvarGridView01 As HimTools2012.controls.DataGridViewWithDataView
    Private WithEvents mvarGridView02 As HimTools2012.controls.DataGridViewWithDataView
    Private WithEvents mvarTBLData As HimTools2012.Data.CSV2Table
    Private mvarMenu As MenuStrip

    Public Sub New()
        MyBase.New(True, True, "CSVTo農地", "CSVTo農地")

        mvarGridView01 = New HimTools2012.controls.DataGridViewWithDataView
        mvarGridView02 = New HimTools2012.controls.DataGridViewWithDataView

        mvarSP = New SynchronizeTwinGridView(mvarGridView01, mvarGridView02, 1)
        mvarSP.Dock = DockStyle.Fill
        mvarSP.Orientation = Orientation.Horizontal

        mvarSP.SplitterDistance = 200

        mvarMenu = New MenuStrip
        With CType(mvarMenu.Items.Add("ファイル(&F)"), ToolStripMenuItem)
            AddHandler .DropDownItems.Add("開く(&O)").Click, AddressOf LoadCSV
        End With
        Me.TopToolStripPanel.Controls.Add(mvarMenu)
        AddHandler Me.ToolStrip.Items.Add("検査").Click, AddressOf 検査
        AddHandler Me.ToolStrip.Items.Add("開始").Click, AddressOf 開始

        Me.ControlPanel.Add(mvarSP)
    End Sub

    Public Sub LoadCSV()
        With New OpenFileDialog
            .Filter = "CSVファイル(*.csv)|*.csv|エクセルファイル(*.xls;*.xlsx)|*.xls;*.xlsx"

            mvarTBLData = Nothing
            If .ShowDialog() = DialogResult.OK Then
                Dim sExt As String = File拡張子(.FileName)
                Select Case sExt.ToLower
                    Case "csv"
                        mvarTBLData = New HimTools2012.Data.CSV2Table(.FileName, System.Text.Encoding.GetEncoding("shift_jis"))
                    Case "xlsx", "xls"
                        If MsgBox("CSV形式の一時ファイルを作成します(複数シートがある場合、最初のシートのみ変換します。)", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                            .FileName = ExcelFileConv(.FileName, HimTools2012.Excel.ExcelCnvType.ToCSV)
                            If IO.File.Exists(.FileName) Then
                                mvarTBLData = New HimTools2012.Data.CSV2Table(.FileName, System.Text.Encoding.GetEncoding("shift_jis"))
                            Else
                                MsgBox("CSVに変換できませんでした。")
                            End If
                        End If
                    Case Else

                End Select

                If mvarTBLData IsNot Nothing Then
                    If mvarTBLData.Columns.Contains("ＩＤ") Then
                        mvarTBLData.Columns("ＩＤ").ColumnName = "ID"
                    End If
                    If Not mvarTBLData.Columns.Contains("ID") Then
                        mvarTBLData.Columns.Add(New DataColumn("ID", GetType(Decimal)))
                    End If
                    If Not mvarTBLData.Columns.Contains("大字ID") Then
                        mvarTBLData.Columns.Add(New DataColumn("大字ID", GetType(Decimal)))
                    End If
                    If Not mvarTBLData.Columns.Contains("地番") Then
                        mvarTBLData.Columns.Add(New DataColumn("地番", GetType(String)))
                    End If
                    If Not mvarTBLData.Columns.Contains("ID合致") Then
                        mvarTBLData.Columns.Add(New DataColumn("ID合致", GetType(Boolean)))
                    End If
                    If Not mvarTBLData.Columns.Contains("大字ID合致") Then
                        mvarTBLData.Columns.Add(New DataColumn("大字ID合致", GetType(Boolean)))
                    End If
                    If Not mvarTBLData.Columns.Contains("地番合致") Then
                        mvarTBLData.Columns.Add(New DataColumn("地番合致", GetType(Boolean)))
                    End If


                    mvarGridView02.SetDataView(mvarTBLData, "", "")

                    mvarGridView01.Columns.Clear()
                    For Each pCol As DataColumn In mvarTBLData.Columns
                        Dim pCombo As New DataGridViewComboBoxColumn
                        pCombo.HeaderText = pCol.ColumnName
                        pCombo.Name = pCol.ColumnName
                        pCombo.Items.Add("ID")
                        pCombo.Items.Add("ID合致")
                        pCombo.Items.Add("大字")
                        pCombo.Items.Add("大字ID")
                        pCombo.Items.Add("大字ID合致")
                        pCombo.Items.Add("地番")
                        pCombo.Items.Add("地番合致")

                        pCombo.Items.Add("-")

                        'pCombo.Items.Add("管理者ID")
                        pCombo.Items.Add("環境保全型農業直接支払交付金")
                        pCombo.Items.Add("多目的機能支払_農地維持支払")
                        pCombo.Items.Add("多目的機能支払_資源向上支払")
                        pCombo.Items.Add("中山間地域等直接支払")

                        mvarGridView01.Columns.Add(pCombo)
                    Next

                    mvarGridView01.Rows.Add()
                    Dim pNew As DataGridViewRow = mvarGridView01.Rows(0)
                    For Each pCol As DataColumn In mvarTBLData.Columns
                        Select Case pCol.ColumnName
                            Case "ID", "ＩＤ" : pNew.Cells(pCol.ColumnName).Value = "ID"
                                pCol.ColumnName = "ID"
                            Case "大字" : pNew.Cells(pCol.ColumnName).Value = "大字"
                            Case "大字ID" : pNew.Cells(pCol.ColumnName).Value = "大字ID"
                            Case "大字ID合致" : pNew.Cells(pCol.ColumnName).Value = "大字ID合致"
                            Case "ID合致" : pNew.Cells(pCol.ColumnName).Value = "ID合致"
                            Case "地番" : pNew.Cells(pCol.ColumnName).Value = "地番"
                            Case "地番合致" : pNew.Cells(pCol.ColumnName).Value = "地番合致"
                        End Select
                    Next
                End If
            End If
        End With
    End Sub

    Public Sub 検査()
        Dim pID As New ArrayList
        Dim p大字IDList As New List(Of Decimal)

        Dim s大字ID列 As String = ""
        Dim s地番列 As String = ""
        Dim p農地TBL As DataTable = Nothing

        For Each pRow As DataGridViewRow In mvarGridView02.Rows
            Dim nID As Decimal = 0
            Dim s地番 As String = ""

            Dim n大字ID As Decimal = 0
            For Each pCol As DataGridViewColumn In mvarGridView01.Columns
                Dim pCell As DataGridViewCell = mvarGridView01.Rows(0).Cells(pCol.Name)
                If pCell.Value IsNot Nothing Then
                    Select Case pCell.Value.ToString
                        Case "ID"
                            If Not IsDBNull(pRow.Cells("ID").Value) Then



                                If Val(pRow.Cells("ID").Value.ToString) <> 0 Then
                                    If Not pID.Contains(pRow.Cells("ID").Value.ToString) Then
                                        pID.Add(pRow.Cells("ID").Value.ToString)
                                    Else
                                        Stop
                                    End If
                                    nID = pRow.Cells("ID").Value
                                Else
                                    Stop
                                End If
                            End If
                            If pID.Count > 20 Then
                                Dim n As Integer = pCheckTBL(pID).Rows.Count
                            End If

                        Case "大字"
                            Dim pORow As DataRow() = App農地基本台帳.TBL大字.Select("[名称]='" & pRow.Cells("大字").Value.ToString & "'")
                            If pORow IsNot Nothing AndAlso pORow.Length > 0 Then
                                pRow.Cells("大字ID").Value = pORow(0).Item("ID")
                                n大字ID = pORow(0).Item("ID")
                            Else
                                pRow.Cells("大字").ErrorText = "一致する大字が見つかりません"
                            End If
                        Case "大字ID"
                            s大字ID列 = pCol.Name
                        Case "地番"
                            s地番列 = pCol.Name
                            If Not IsDBNull(pRow.Cells(s地番列).Value) Then
                                s地番 = pRow.Cells(s地番列).Value
                            Else
                                pRow.Cells(s地番列).ErrorText = "入力が不正です"
                            End If
                        Case "管理者ID"
                            If pRow.Cells(pCol.Name).Value.ToString.Length = 0 Then
                            ElseIf Not IsNumeric(pRow.Cells(pCol.Name).Value.ToString) Then
                                pRow.Cells(s地番列).ErrorText = "入力は必ず個人コード(ID)でお願いします"
                            End If
                        Case "多目的機能支払_農地維持支払"
                            Select Case pRow.Cells(pCol.Name).Value.ToString
                                Case "0", "1", "○"
                                Case Else
                                    pRow.Cells(s地番列).ErrorText = "入力が不正です"
                            End Select

                    End Select
                End If
            Next



            If nID = 0 AndAlso n大字ID <> 0 AndAlso s地番.Length > 0 Then
                If Not p大字IDList.Contains(n大字ID) Then
                    If p農地TBL Is Nothing Then
                        p農地TBL = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE [大字ID]=" & n大字ID)
                    Else
                        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE [大字ID]=" & n大字ID)
                        p農地TBL.Merge(pTBL)
                    End If

                    p大字IDList.Add(n大字ID)
                End If

                Dim pView As New DataView(p農地TBL, "[地番] = '" & s地番 & "'", "", DataViewRowState.CurrentRows)
                Select Case pView.Count
                    Case 0
                        pRow.ErrorText = "合致する大字・地番がありません"
                    Case 1
                        pRow.Cells("ID").Value = pView.Item(0).Item("ID")

                        Dim n As Integer = pCheckTBL(pView.ToTable).Rows.Count
                    Case 2
                        pRow.ErrorText = "合致する大字・地番が複数行あります。"
                End Select
            End If
        Next

        If pID.Count > 0 Then
            Dim n As Integer = pCheckTBL(pID).Rows.Count
        End If

        If pCheckTBL IsNot Nothing AndAlso pCheckTBL.Rows.Count > 0 Then
            For Each pRow As DataGridViewRow In mvarGridView02.Rows
                If IsDBNull(pRow.Cells("ID").Value) Then

                ElseIf pCheckTBL IsNot Nothing Then
                    Dim pXRow As DataRow = pCheckTBL.Rows.Find(Val(pRow.Cells("ID").Value.ToString))
                    If pXRow IsNot Nothing Then
                        pRow.Cells("ID合致").Value = True

                        If s大字ID列.Length > 0 Then
                            If pRow.Cells(s大字ID列).Value.ToString = pXRow.Item("大字ID").ToString Then
                                pRow.Cells("大字ID合致").Value = True
                            End If
                        End If
                        If s地番列.Length > 0 Then
                            If pRow.Cells(s地番列).Value.ToString = pXRow.Item("地番").ToString Then
                                pRow.Cells("地番合致").Value = True
                            End If
                        End If

                    Else
                        If pRow.Cells("ID").Value.ToString.Length > 0 Then

                            Debug.Print(pRow.Cells("ID").Value.ToString)

                        End If

                        pRow.ErrorText = "IDが合致しません"
                    End If
                End If
            Next
        End If


    End Sub

    Public Sub 開始()
        If Not mvarGridView02.Rows.Count > 0 AndAlso Not mvarGridView01.Rows.Count > 0 Then
            MsgBox("データが正しく読み込まれていません。", MsgBoxStyle.Critical)
        Else
            Dim sID As String = ""
            Dim sID合致 As String = ""
            Dim s変換列 As New Dictionary(Of String, String)

            For Each pCol As DataGridViewColumn In mvarGridView01.Columns
                If mvarGridView01.Rows(0).Cells(pCol.Name).Value IsNot Nothing Then
                    Select Case mvarGridView01.Rows(0).Cells(pCol.Name).Value.ToString
                        Case "ID"
                            sID = pCol.Name
                        Case "ID合致"
                            sID合致 = pCol.Name
                        Case "-"
                        Case "大字"
                        Case "大字ID"
                        Case "大字ID合致"
                        Case "地番"
                        Case "地番合致"
                        Case ""
                        Case Else
                            s変換列.Add(mvarGridView01.Rows(0).Cells(pCol.Name).Value.ToString, pCol.Name)
                    End Select
                End If
            Next
            If s変換列.Count = 0 Then
                MsgBox("取込先の項目が設定されていません", MsgBoxStyle.Critical)
            ElseIf sID.Length * sID合致.Length > 0 AndAlso s変換列.Count > 0 Then
                Dim nCount As Integer = 0
                For Each pRow As DataRow In mvarTBLData.Rows
                    nCount += CheckBool(pRow.Item("ID合致")) * CheckBool(pRow.Item("大字ID合致")) * CheckBool(pRow.Item("地番合致"))
                Next

                If Not nCount > 0 Then
                    MsgBox("検査結果で合致するデータがありません。データが不正であるか正しく検査されていません。", MsgBoxStyle.Critical)
                ElseIf mvarTBLData.Rows.Count <> nCount AndAlso MsgBox("取込み先農地で合致しないデータがあります。実行しますか", MsgBoxStyle.YesNo) = MsgBoxResult.No Then

                Else
                    Dim sSQL As New System.Text.StringBuilder

                    For Each pRow As DataRow In mvarTBLData.Rows
                        If CheckBool(pRow.Item("ID合致")) * CheckBool(pRow.Item("大字ID合致")) * CheckBool(pRow.Item("地番合致")) Then
                            Dim sSQLF As New List(Of String)

                            For Each sType As String In s変換列.Keys
                                Dim s列 As String = s変換列.Item(sType)
                                Select Case sType
                                    Case "管理者ID"
                                        sSQLF.Add("[管理者ID]=" & ConvNumber(pRow.Item(s列)))
                                    Case "環境保全型農業直接支払交付金"
                                        sSQLF.Add("[環境保全交付金]=" & ConvNumber(pRow.Item(s列)))
                                    Case "多目的機能支払_農地維持支払"
                                        sSQLF.Add("[農地維持交付金]=" & ConvNumber(pRow.Item(s列)))
                                    Case "多目的機能支払_資源向上支払"
                                        sSQLF.Add("[資源向上交付金]=" & ConvNumber(pRow.Item(s列)))
                                    Case "中山間地域等直接支払"
                                        sSQLF.Add("[中山間直接支払]=" & ConvNumber(pRow.Item(s列)))
                                    Case Else
                                        Stop
                                End Select
                            Next

                            sSQL.AppendLine("UPDATE [D:農地Info] SET " & Join(sSQLF.ToArray, ", ") & " WHERE [ID]=" & pRow.Item("ID"))
                        End If
                        If sSQL.Length > 1024 Then
                            SysAD.DB(sLRDB).ExecuteSQL(sSQL.ToString)
                            sSQL.Clear()
                        End If
                    Next
                    If sSQL.Length > 0 Then
                        SysAD.DB(sLRDB).ExecuteSQL(sSQL.ToString)
                        sSQL.Clear()
                    End If

                End If
            End If
        End If
    End Sub


    Private Function CheckBool(pData As Object) As Integer
        If pData Is Nothing Then
            Return 0
        ElseIf IsDBNull(pData) Then
            Return 0
        ElseIf pData = True Then
            Return 1
        Else
            Return 0
        End If
    End Function

    Private Function ConvBool(pdata As Object) As String
        If TypeName(pdata).ToString Then
            Select Case CStr(pdata)
                Case "True", "TRUE" : Return "True"
                Case "False", "FALSE" : Return "False"
                Case "1", "○", "☑" : Return "True"
                Case "0", "□" : Return "False"
                Case "-1" : Return "True"
                Case "", "-" : Return "False"
                Case Else
            End Select
        End If

        Return "Null"
    End Function

    Private Function ConvNumber(pdata As Object) As String
        If TypeName(pdata) = "String" Then
            If IsNumeric(pdata) Then
                Return Val(pdata)
            ElseIf pdata = "○" Then
                Return 1
            End If
        End If

        Return "Null"
    End Function

    Private mvarCheckTBL As DataTable
    Private ReadOnly Property pCheckTBL(Optional pID As ArrayList = Nothing) As DataTable
        Get
            If pID IsNot Nothing Then
                If mvarCheckTBL Is Nothing Then
                    mvarCheckTBL = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE [ID] IN(" & Join(pID.ToArray, ",") & ")")
                    mvarCheckTBL.PrimaryKey = {mvarCheckTBL.Columns("ID")}
                Else
                    mvarCheckTBL.Merge(SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE [ID] IN(" & Join(pID.ToArray, ",") & ")"))
                End If
                pID.Clear()
            Else

            End If

            Return mvarCheckTBL
        End Get
    End Property
    Private ReadOnly Property pCheckTBL(pTBL As DataTable) As DataTable
        Get
            If mvarCheckTBL Is Nothing Then
                mvarCheckTBL = pTBL
                mvarCheckTBL.PrimaryKey = {mvarCheckTBL.Columns("ID")}
            Else
                mvarCheckTBL.Merge(pTBL)
            End If
            Return mvarCheckTBL
        End Get
    End Property

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.CloseOK
        End Get
    End Property



End Class

