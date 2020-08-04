Imports HimTools2012.CommonFunc

Public Class C申請リスト
    Inherits CNList農地台帳

    Private WithEvents mvar総会資料 As ToolStripButton
    Private WithEvents mvar同一許可番号設定 As ToolStripButton
    Private WithEvents mvar許可連番設定 As ToolStripButton

    Private WithEvents mvar一括許可 As ToolStripButton
    Private WithEvents mvar受付交付簿 As ToolStripButton
    Private WithEvents mvar並べ替え As ToolStripDropDownButton
    Private WithEvents mvar受付番号順 As ToolStripMenuItem
    Private WithEvents mvar許可番号順 As ToolStripMenuItem

    Private WithEvents mvar申請データCSV As ToolStripButton
    Private WithEvents mvar宛名CSV As ToolStripButton

    Public Sub New(ByVal pParent As classPage農家世帯, ByVal sKey As String, ByVal sText As String)
        MyBase.New(sText, sKey, True)
        SetGridColumn(GetType(CObj申請))

        mvar総会資料 = New ToolStripButton("総会資料作成")
        Me.ToolStrip.Items.Add(mvar総会資料)
        Me.ToolStrip.Items.Add(New ToolStripSeparator)

        mvar並べ替え = New ToolStripDropDownButton("並び順")
        mvar受付番号順 = New ToolStripMenuItem("受付番号順")
        mvar許可番号順 = New ToolStripMenuItem("許可番号順")
        mvar並べ替え.DropDownItems.Add(mvar受付番号順)
        mvar並べ替え.DropDownItems.Add(mvar許可番号順)
        Me.ToolStrip.Items.Add(mvar並べ替え)

        mvar許可連番設定 = New ToolStripButton("許可連番設定")
        Me.ToolStrip.Items.Add(mvar許可連番設定)

        mvar同一許可番号設定 = New ToolStripButton("同一許可番号設定")
        Me.ToolStrip.Items.Add(mvar同一許可番号設定)

        Me.ToolStrip.Items.Add(New ToolStripSeparator)

        mvar一括許可 = New ToolStripButton("許可/承認する")
        Me.ToolStrip.Items.Add(mvar一括許可)

        mvar受付交付簿 = New ToolStripButton("受付交付簿")
        Me.ToolStrip.Items.Add(mvar受付交付簿)

        mvar申請データCSV = New ToolStripButton("申請データCSV")
        Me.ToolStrip.Items.Add(mvar申請データCSV)

        mvar宛名CSV = New ToolStripButton("宛名CSV")
        Me.ToolStrip.Items.Add(mvar宛名CSV)



    End Sub

    Private mvarWhere As String = ""

    Public Overrides Sub 検索開始(ByVal sWhere As String, ByVal sViewWhere As String, Optional sOrderBy As String = "", Optional sColumnStyle As String = "")
        If sColumnStyle = "" Then
            sColumnStyle = "申請リストColumns"
        End If

        App農地基本台帳.ListColumnDesign.SetGridColumns(GView, sColumnStyle)

        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D_申請] WHERE " & sWhere)
        App農地基本台帳.TBL申請.MergePlus(pTBL)

        mvarWhere = sViewWhere
        GView.SetDataView(App農地基本台帳.TBL申請.Body, sViewWhere, sOrderBy)


        If GView.DataView.Count > 0 Then
            GView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCellsExceptHeader)
        End If
    End Sub

    Private Sub mvar総会資料_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar総会資料.Click
        総会資料作成()
    End Sub
    Private Sub mvar同一許可番号設定_Click(sender As Object, e As System.EventArgs) Handles mvar同一許可番号設定.Click
        If GView.SelectedRows IsNot Nothing AndAlso MsgBox("同じ許可番号を設定しますか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Dim nNo As String = InputBox("許可番号を入力してください", "同一許可番号設定", "")
            If Val(nNo) > 0 AndAlso GView.SelectedRows IsNot Nothing Then
                For Each pRow As DataGridViewRow In GView.SelectedRows
                    If pRow.Cells("Key").Value.ToString.Length > 0 Then
                        Dim sKey As String = pRow.Cells("Key").Value
                        pRow.Cells("許可番号").Value = Val(nNo)
                        SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D_申請] SET [許可番号]=" & nNo & " WHERE [ID]=" & GetKeyCode(sKey))
                    End If
                Next
            End If
        End If
    End Sub
    Public Overrides Property IconKey As String
        Get
            Return "List"
        End Get
        Set(value As String)

        End Set
    End Property

    Private Sub mvar許可連番設定_Click(sender As Object, e As System.EventArgs) Handles mvar許可連番設定.Click
        Dim param1 As New InputParam("開始許可番号", "開始許可番号", 1, GetType(Decimal))
        Dim param2 As New InputParam("許可年月日", "許可年月日", Now().Date, GetType(DateTime))
        Dim dlg As New dlgInputForm("許可番号設定", "設定する許可番号と許可日を入力してください。", param1, param2)

        If GView.SelectedRows IsNot Nothing AndAlso GView.SelectedRows.Count > 0 AndAlso dlg.ShowDialog() = DialogResult.OK Then
            If param1.Value >= 1 AndAlso IsDate(param2.ToString()) Then
                Dim sNo As String = param1.Value

                If Val(sNo) > 0 AndAlso GView.SelectedRows IsNot Nothing Then
                    Dim nList As New List(Of Integer)

                    For Each pRow As DataGridViewRow In GView.SelectedRows
                        nList.Add(pRow.Cells("ID").Value)
                    Next

                    Dim nNo As Integer = Val(sNo)
                    For Each pVRow As DataRowView In App農地基本台帳.TBL申請.ToDataView("", "受付番号")
                        If nList.Contains(pVRow.Item("ID")) Then
                            pVRow.Item("許可番号") = nNo
                            Dim Dt As DateTime = CDate(param2.ToString())
                            pVRow.Item("許可年月日") = Dt
                            Dim sDT As String = String.Format("#{0}/{1}/{2}#", Dt.Month, Dt.Day, Dt.Year)
                            SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D_申請] SET [許可番号]={0},[許可年月日]={1} WHERE [ID]={2}", nNo, sDT, pVRow.Item("ID"))
                            nNo += 1
                        End If
                    Next
                End If
            End If
        Else
            If GView.SelectedRows.Count = 0 Then
                MsgBox("設定する筆を選択してください。")
            End If
        End If

        Exit Sub

    End Sub


    Private Sub mvar一括許可_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar一括許可.Click
        If GView.SelectedRows IsNot Nothing AndAlso GView.SelectedRows.Count > 0 AndAlso
            MsgBox(String.Format("選択した申請[{0}]を一括して許可しますか", GView.SelectedRows.Count), MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Dim sDate As String = InputBox("許可日を入力してください", "許可処理", Now.Date.ToString.Replace(" 0:00:00", ""))
            Dim s発行日 As String = InputBox("許可書/通知書の発行日を入力してください", "許可/承認", Now.Date.ToString.Replace(" 0:00:00", ""))

            If IsDate(sDate) AndAlso IsDate(s発行日) Then
                Try
                    Dim pDate As Date = CDate(sDate)
                    Dim sFolder As String = SysAD.OutputFolder & String.Format("\許可書{0}_{1}", pDate.Year, pDate.Month)
                    If Not IO.Directory.Exists(sFolder) Then
                        IO.Directory.CreateDirectory(sFolder)
                    End If

                    Dim pList As New List(Of Integer)

                    For Each pRow As DataGridViewRow In GView.SelectedRows
                        pList.Add(pRow.Cells("ID").Value)
                    Next
                    pList.Reverse()

                    For Each nID As Integer In pList
                        Dim pRow As New HimTools2012.Data.DataRowPlus(App農地基本台帳.TBL申請.Rows.Find(nID))
                        Dim skey = pRow.Item("Key")
                        Select Case pRow.Item("法令")
                            Case enum法令.農地法3条所有権, enum法令.基盤強化法所有権 : sub異動所有権移転(CDate(sDate), CDate(s発行日), pRow, sFolder, False)
                            Case enum法令.農地法3条耕作権 : fnc設置利用権(skey, pRow.Item("法令"), sFolder, False, CDate(sDate), CDate(s発行日))
                            Case enum法令.農地法4条, enum法令.農地法4条一時転用 : sub農地転用(pRow, 10040, sFolder, False, CDate(sDate))
                            Case enum法令.農地法5条所有権, enum法令.農地法5条貸借, enum法令.農地法5条一時転用 : sub農地転用(pRow, 10050, sFolder, False, CDate(sDate))
                            Case enum法令.利用権設定, enum法令.利用権移転 : fnc設置利用権(skey, 10000 + pRow.Item("法令"), sFolder, False, CDate(sDate), CDate(s発行日))
                            Case enum法令.農地法18条解約, enum法令.合意解約, enum法令.中間管理機構へ農地の返還
                                If SysAD.市町村.市町村名 = "伊佐市" Then
                                    Dim tFolder As String = SysAD.OutputFolder & String.Format("\通知書{0}_{1}", pDate.Year, pDate.Month)
                                    If Not IO.Directory.Exists(tFolder) Then
                                        IO.Directory.CreateDirectory(tFolder)
                                    End If

                                    fnc通知書発行(skey, 10000 + pRow.Item("法令"), tFolder, False, CDate(Now), Nothing)
                                    sFolder = SysAD.OutputFolder & String.Format("\通知書{0}_{1}", pDate.Year, pDate.Month)
                                End If

                                Dim nRow As DataRow = App農地基本台帳.TBL申請.Rows.Find(nID)
                                Dim p申請 As New CObj申請(nRow, False)
                                Select Case pRow.Item("法令")
                                    Case enum法令.農地法18条解約 : p申請.RentEnd(skey, CDate(sDate), "18条解約")
                                    Case enum法令.合意解約 : p申請.RentEnd(skey, CDate(sDate), "20条解約")
                                    Case enum法令.中間管理機構へ農地の返還 : p申請.RentEnd(skey, CDate(sDate), "農地返還")
                                End Select
                            Case Else
                                MsgBox("未対応の処理を実行しました")
                                Exit Sub
                        End Select

                    Next
                    SysAD.ShowFolder(sFolder)

                    MsgBox("処理しました")

                Catch ex As Exception
                    Stop
                End Try
            End If

        End If
    End Sub
    Public Function InputDateTime(ByVal sPrompt As String, ByVal sTitle As String, ByVal pDefault As Date) As Object
        Dim sDate As String = InputBox(sPrompt, sTitle, pDefault.ToString)
        If IsDate(sDate) Then
            Return CDate(sDate)
        Else
            Return Nothing
        End If
    End Function
    Private Sub mvar受付交付簿_Click(sender As Object, e As System.EventArgs) Handles mvar受付交付簿.Click
        MsgBox("対象データがありません", MsgBoxStyle.Critical)

    End Sub

    Private Sub mvar申請データCSV_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar申請データCSV.Click
        Dim pSet As New HimTools2012.Data.DataSetEx
        Dim pTBL申請 As DataTable = GView.DataView.ToTable
        Dim pTbl農地Link As New DataTable("D_農地Link")

        For Each pCol As DataColumn In pTBL申請.Columns
            pTbl農地Link.Columns.Add(New DataColumn(pCol.ColumnName, pCol.DataType, pCol.Expression))
        Next

        pTbl農地Link.Columns.Add(New DataColumn("農地ID", App農地基本台帳.TBL農地.Columns("ID").DataType))
        pTbl農地Link.Columns.Add(New DataColumn("申請ID", GetType(Integer)))

        pTbl農地Link.PrimaryKey = New DataColumn() {pTbl農地Link.Columns("農地ID"), pTbl農地Link.Columns("申請ID")}
        pSet.Tables.Add(pTBL申請)
        pSet.Tables.Add(pTbl農地Link)
        pSet.Relations.Add(New DataRelation("申請情報", pTBL申請.Columns("ID"), pTbl農地Link.Columns("申請ID"), False))

        Dim pList農地 As New List(Of String)
        Dim pList転用 As New List(Of String)
        Dim pLoadList農地 As New List(Of String)
        Dim pLoadList転用 As New List(Of String)

        pTBL申請.TableName = "D_申請"

        Dim pPage As New CNList農地台帳("申請リスト", "申請リスト", True)

        For Each pRow As DataRow In pTBL申請.Rows
            If Not IsDBNull(pRow.Item("農地リスト")) Then
                Dim St As String = pRow.Item("農地リスト")
                Dim Ar() As String = Split(St, ";")

                For Each Key As String In Ar

                    If Key.Length > 0 Then
                        If Key.StartsWith("農地.") Then
                            Dim pNewRow As DataRow = pTbl農地Link.NewRow
                            For Each pCol As DataColumn In pTBL申請.Columns
                                pNewRow.Item(pCol.ColumnName) = pRow.Item(pCol.ColumnName)
                            Next

                            Dim nID As Integer = Val(Strings.Mid(Key, 4))

                            pNewRow.Item("農地ID") = nID
                            pNewRow.Item("申請ID") = pRow.Item("ID")

                            pTbl農地Link.Rows.Add(pNewRow)
                            Dim pRow農地 As DataRow = App農地基本台帳.TBL農地.Rows.Find(nID)

                            If pRow農地 Is Nothing AndAlso Not pLoadList農地.Contains(nID) Then
                                pLoadList農地.Add(nID.ToString)
                            End If
                            pList農地.Add(nID.ToString)
                        ElseIf Key.StartsWith("転用農地.") Then
                            Dim pNewRow As DataRow = pTbl農地Link.NewRow
                            For Each pCol As DataColumn In pTBL申請.Columns
                                pNewRow.Item(pCol.ColumnName) = pRow.Item(pCol.ColumnName)
                            Next

                            Dim nID As Integer = Val(Strings.Mid(Key, 6))

                            pNewRow.Item("農地ID") = nID
                            pNewRow.Item("申請ID") = pRow.Item("ID")

                            pTbl農地Link.Rows.Add(pNewRow)
                            Dim pRow農地 As DataRow = App農地基本台帳.TBL転用農地.Rows.Find(nID)

                            If pRow農地 Is Nothing AndAlso Not pLoadList転用.Contains(nID) Then
                                pLoadList転用.Add(nID.ToString)
                            End If
                            pList転用.Add(nID.ToString)
                        Else
#If DEBUG Then
                            Stop
#End If
                        End If
                    End If
                Next
            End If
        Next


        If pLoadList農地.Count > 0 Then
            App農地基本台帳.TBL農地.MergePlus(SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE [ID] IN (" & Join(pLoadList農地.ToArray, ",") & ")"))
        End If
        If pLoadList転用.Count > 0 Then
            App農地基本台帳.TBL転用農地.MergePlus(SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D_転用農地] WHERE [ID] IN (" & Join(pLoadList転用.ToArray, ",") & ")"))
        End If

        Dim p農地TMP As DataTable = Nothing
        If pList農地.Count > 0 Then
            Dim sList As String = Join(pList農地.ToArray, ",")
            p農地TMP = New DataView(App農地基本台帳.TBL農地.Body, "ID IN(" & sList & ")", "", DataViewRowState.CurrentRows).ToTable
            pSet.Tables.Add(p農地TMP)
        End If
        If pList転用.Count > 0 Then
            Dim sList As String = Join(pList転用.ToArray, ",")
            If p農地TMP Is Nothing Then
                p農地TMP = New DataView(App農地基本台帳.TBL転用農地.Body, "ID IN(" & sList & ")", "", DataViewRowState.CurrentRows).ToTable
                pSet.Tables.Add(p農地TMP)
            Else

                p農地TMP.Merge(New DataView(App農地基本台帳.TBL転用農地.Body, "ID IN(" & sList & ")", "", DataViewRowState.CurrentRows).ToTable)
            End If
        End If
        If p農地TMP Is Nothing Then
            MsgBox("対象農地が見つかりません")
        Else
            pSet.Relations.Add(New DataRelation("転用申請農地", p農地TMP.Columns("ID"), pTbl農地Link.Columns("農地ID"), False))
            'pTbl農地Link.Columns.Add("大字ID", GetType(String), "Parent(転用申請農地).大字ID")
            pTbl農地Link.Columns.Add("大字", GetType(String), "Parent(転用申請農地).大字")
            'pTbl農地Link.Columns.Add("小字ID", GetType(String), "Parent(転用申請農地).小字ID")
            pTbl農地Link.Columns.Add("小字", GetType(String), "Parent(転用申請農地).小字")
            pTbl農地Link.Columns.Add("所在", GetType(String), "Parent(転用申請農地).所在")
            pTbl農地Link.Columns.Add("地番", GetType(String), "Parent(転用申請農地).地番")
            pTbl農地Link.Columns.Add("登記簿地目名", GetType(String), "Parent(転用申請農地).登記簿地目名")
            pTbl農地Link.Columns.Add("現況地目名", GetType(String), "Parent(転用申請農地).現況地目名")
            pTbl農地Link.Columns.Add("登記簿面積", GetType(String), "Parent(転用申請農地).登記簿面積")
            pTbl農地Link.Columns.Add("実面積", GetType(String), "Parent(転用申請農地).実面積")

            pTbl農地Link.Columns.Add("所有者名", GetType(String), "Parent(転用申請農地).所有者氏名")
            pTbl農地Link.Columns.Add("所有者住所", GetType(String), "Parent(転用申請農地).所有者住所")
            pPage.GView.AutoGenerateColumns = True
            pPage.GView.Create件数表示Ctrl(pPage.ToolStrip)
            pPage.GView.SetDataView(pTbl農地Link, "", "")
            pPage.GView.Columns("Key").Visible = False
            pPage.GView.Columns("アイコン").Visible = False

            SysAD.page農家世帯.中央Tab.AddPage(pPage)
        End If
    End Sub

    Private Sub mvar受付番号順_Click(sender As Object, e As System.EventArgs) Handles mvar受付番号順.Click
        If mvarWhere.Length > 0 Then
            検索開始(mvarWhere, mvarWhere, "受付補助記号,受付番号,受付年月日")
        End If
    End Sub

    Private Sub mvar許可番号順_Click(sender As Object, e As System.EventArgs) Handles mvar許可番号順.Click
        If mvarWhere.Length > 0 Then
            検索開始(mvarWhere, mvarWhere, "許可番号,受付番号,許可年月日")
        End If
    End Sub

    Private Sub mvar宛名CSV_Click(sender As Object, e As System.EventArgs) Handles mvar宛名CSV.Click
        Dim pSet As New HimTools2012.Data.DataSetEx
        Dim pTBL申請 As DataTable = GView.DataView.ToTable
        Dim pTbl農家リスト As New DataTable("農家リスト")

        pTbl農家リスト.Columns.Add(New DataColumn("受付補助記号", GetType(String)))
        pTbl農家リスト.Columns.Add(New DataColumn("受付番号", GetType(Integer)))
        pTbl農家リスト.Columns.Add(New DataColumn("個人ID", GetType(Long)))
        pTbl農家リスト.Columns.Add(New DataColumn("氏名", GetType(String)))
        pTbl農家リスト.Columns.Add(New DataColumn("郵便番号", GetType(String)))
        pTbl農家リスト.Columns.Add(New DataColumn("住所", GetType(String)))
        pTbl農家リスト.PrimaryKey = {pTbl農家リスト.Columns("受付番号"), pTbl農家リスト.Columns("個人ID")}

        Dim bMess As Boolean = False
        Dim _list As New List(Of Integer)
        For Each pRow As DataRow In pTBL申請.Rows
            Dim pVal As Integer = Year(pRow.Item("許可年月日")).ToString("0000") & Month(pRow.Item("許可年月日")).ToString("00")
            If _list.Count = 0 Then
                _list.Add(pVal)
            Else
                If _list.Contains(pVal) Then
                    _list.Add(pVal)
                Else
                    bMess = True
                    _list.Add(pVal)
                End If
            End If
        Next

        Dim pView As DataView
        If bMess = True Then
            Dim Start日付 As String = ""
            Dim End日付 As String = ""
            With New dlgInputBWDate(Now.Date)
                If .ShowDialog = DialogResult.OK Then
                    Start日付 = String.Format("#{0}/{1}/{2}#", .StartDate.Month, .StartDate.Day, .StartDate.Year)
                    End日付 = String.Format("#{0}/{1}/{2}#", .EndDate.Month, .EndDate.Day, .EndDate.Year)
                End If
            End With

            If Start日付.Length > 0 AndAlso End日付.Length > 0 Then
                pView = New DataView(pTBL申請, String.Format("[許可年月日] >= {0} AND [許可年月日] <= {1}", Start日付, End日付), "", DataViewRowState.CurrentRows)
            Else
                pView = New DataView(pTBL申請, "", "", DataViewRowState.CurrentRows)
            End If
        Else
            pView = New DataView(pTBL申請, "", "", DataViewRowState.CurrentRows)
        End If

        Dim nList As New List(Of Long)
        For Each pRow As DataRowView In pView
            If Not IsDBNull(pRow.Item("申請者A")) AndAlso Not nList.Contains(pRow.Item("申請者A")) Then
                nList.Add(pRow.Item("申請者A"))
            End If
            If Not IsDBNull(pRow.Item("申請者B")) AndAlso Not nList.Contains(pRow.Item("申請者B")) Then
                nList.Add(pRow.Item("申請者B"))
            End If
        Next

        If nList.Count > 0 Then
            Dim sB As New System.Text.StringBuilder
            For Each n As Long In nList
                sB.Append("," & n)
            Next
            Dim pTBL個人 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:個人Info] WHERE [ID] IN (" & sB.ToString.Substring(1) & ");")
            App農地基本台帳.TBL個人.MergePlus(pTBL個人)
        End If


        For Each pRow As DataRowView In pView
            Try
                Select Case CType(pRow.Item("法令"), enum法令)
                    Case enum法令.農地法3条所有権, enum法令.利用権設定, enum法令.農地法3条耕作権, enum法令.農地法5条所有権, enum法令.農地法5条一時転用, enum法令.農地法5条貸借

                        If pTbl農家リスト.Rows.Find({pRow.Item("受付番号"), pRow.Item("申請者A")}) Is Nothing Then
                            Dim pRNRowA As DataRow = pTbl農家リスト.NewRow
                            Dim p個人RowA As DataRow = App農地基本台帳.TBL個人.FindRowByID(pRow.Item("申請者A"))
                            pRNRowA.Item("受付補助記号") = pRow.Item("受付補助記号").ToString
                            pRNRowA.Item("受付番号") = pRow.Item("受付番号")

                            pRNRowA.Item("個人ID") = pRow.Item("申請者A")
                            pRNRowA.Item("氏名") = pRow.Item("氏名A")
                            pRNRowA.Item("郵便番号") = p個人RowA.Item("郵便番号")
                            pRNRowA.Item("住所") = pRow.Item("住所A")

                            pTbl農家リスト.Rows.Add(pRNRowA)
                        End If

                        If pTbl農家リスト.Rows.Find({pRow.Item("受付番号"), pRow.Item("申請者B")}) Is Nothing Then
                            Dim pRNRowB As DataRow = pTbl農家リスト.NewRow
                            Dim p個人RowB As DataRow = App農地基本台帳.TBL個人.FindRowByID(pRow.Item("申請者B"))

                            pRNRowB.Item("受付補助記号") = pRow.Item("受付補助記号").ToString
                            pRNRowB.Item("受付番号") = pRow.Item("受付番号")

                            pRNRowB.Item("個人ID") = pRow.Item("申請者B")
                            pRNRowB.Item("氏名") = pRow.Item("氏名B")
                            pRNRowB.Item("郵便番号") = p個人RowB.Item("郵便番号")
                            pRNRowB.Item("住所") = pRow.Item("住所B")

                            pTbl農家リスト.Rows.Add(pRNRowB)
                        End If

                    Case enum法令.農地法4条, enum法令.農地法4条一時転用
                        Dim pRNRow As DataRow = pTbl農家リスト.NewRow
                        Dim p個人RowA As DataRow = App農地基本台帳.TBL個人.FindRowByID(pRow.Item("申請者A"))

                        pRNRow.Item("受付補助記号") = pRow.Item("受付補助記号").ToString
                        pRNRow.Item("受付番号") = pRow.Item("受付番号")

                        pRNRow.Item("個人ID") = pRow.Item("申請者A")
                        pRNRow.Item("氏名") = pRow.Item("氏名A")
                        pRNRow.Item("郵便番号") = p個人RowA.Item("郵便番号")
                        pRNRow.Item("住所") = pRow.Item("住所A")

                        pTbl農家リスト.Rows.Add(pRNRow)

                    Case Else
                        Stop
                End Select
            Catch ex As Exception
                If Not SysAD.IsClickOnceDeployed Then
                    Stop
                End If
            End Try
        Next
        Dim pPage As New CNList農地台帳("宛名リスト", "宛名リスト", True)
        pPage.GView.AutoGenerateColumns = True
        pPage.GView.SetDataView(pTbl農家リスト, "", "[受付補助記号],[受付番号]")

        SysAD.page農家世帯.中央Tab.AddPage(pPage)
    End Sub
End Class

Public Class C一括許可入力支援
    Inherits HimTools2012.InputSupport.CInputSupport

    Public Sub New(ByRef p申請 As CObj申請)
        MyBase.New(App農地基本台帳.DataMaster.Body)

        Dim p許可日 As Object = Now
        If p許可日 IsNot Nothing AndAlso p許可日 > #1990/01/01# Then
            _許可年月日 = p許可日
        Else
            _許可年月日 = Now.Date
        End If

        If p申請.許可番号 > 0 Then
            Me.開始許可番号 = p申請.許可番号
        Else
            Me.開始許可番号 = _開始許可番号
        End If

    End Sub

    Private _開始許可番号 As Integer
    Public Property 開始許可番号 As Integer
        Get
            Return _開始許可番号
        End Get
        Set(value As Integer)
            _開始許可番号 = value
        End Set
    End Property

    Private _許可年月日 As DateTime
    Public Property 許可年月日 As DateTime
        Get
            Return _許可年月日
        End Get
        Set(value As DateTime)
            _許可年月日 = value
        End Set
    End Property
End Class
