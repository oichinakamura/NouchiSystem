'20160806霧島
Imports System.ComponentModel
Imports HimTools2012.Excel.XMLSS2003
Public MustInherit Class CPrint諮問資料
    Inherits CPrint諮問意見書作成共通

    Public Sub New()

    End Sub
End Class

Class CPrint諮問総括表
    Inherits CPrint諮問意見書単票作成

    Public Overrides Sub SetDataRow(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRow As System.Data.DataRowView, Optional ByVal pObj As Object = Nothing)

    End Sub

    Public Overloads Overrides Sub Set単票Data(ByRef pDataCreater As C諮問意見書Data作成, ByVal pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet)
        pSheet.ValueReplace("{市町村名}", SysAD.市町村.市町村名)
        Dim sError As String = ""
        Try

            If pDataCreater.TabCtrl.TabPages.ContainsKey("n.40") Then
                Dim pTB4条 As 諮問意見書Page = pDataCreater.TabCtrl.TabPages("n.40")
                Dim pTable As DataTable = CType(pTB4条.List.DataSource, DataView).ToTable
                Dim pView As New DataView(pTable, "[選択]=True", "", DataViewRowState.CurrentRows)
                Dim pTBL4条諮問総括 As New DataTable("4条諮問総括")
                pTBL4条諮問総括.Columns.Add(New DataColumn("ID", GetType(Integer)))
                pTBL4条諮問総括.Columns.Add(New DataColumn("名称", GetType(String)))
                Dim pCt As New DataColumn("Count", GetType(Integer))
                pCt.DefaultValue = 0
                pTBL4条諮問総括.Columns.Add(pCt)
                pTBL4条諮問総括.PrimaryKey = New DataColumn() {pTBL4条諮問総括.Columns("ID")}
                sError = "004-001"
                If pView.Count > 0 Then
                    pDataCreater.Maximum = pView.Count
                    pDataCreater.Value = 0
                    For Each pRow As DataRowView In pView
                        If pRow.Item("選択") Then
                            nLoop += 1
                            pDataCreater.Message = "諮問総括表 処理中 (" & nLoop + 1 & "/" & pView.Count & ")"
                            pDataCreater.Value = nLoop
                            Dim n As Integer = Val(pRow.Item("申請後農地分類").ToString)

                            Dim pRowT As DataRow = pTBL4条諮問総括.Rows.Find(n)
                            If pRowT Is Nothing Then

                                pRowT = pTBL4条諮問総括.NewRow
                                pRowT.Item("ID") = n

                                pRowT.Item("Count") = 0
                                pTBL4条諮問総括.Rows.Add(pRowT)
                            End If
                            pRowT.Item("Count") += 1
                        End If
                    Next
                    For n = 1 To 10
                        Dim pRowT As DataRow = pTBL4条諮問総括.Rows.Find(n)
                        If pRowT Is Nothing Then
                            pSheet.ValueReplace(Replace("{4条件数00}", "00", Strings.Right("00" & n, 2)), "0")
                        Else
                            pSheet.ValueReplace(Replace("{4条件数00}", "00", Strings.Right("00" & n, 2)), pRowT.Item("Count"))
                        End If
                        '
                    Next
                End If
            End If
            If pDataCreater.TabCtrl.TabPages.ContainsKey("n.50") Then
                Dim pTB5条 As 諮問意見書Page = pDataCreater.TabCtrl.TabPages("n.50")
                Dim pTable As DataTable = CType(pTB5条.List.DataSource, DataView).ToTable
                Dim pView As New DataView(pTable, "[選択]=True", "", DataViewRowState.CurrentRows)
                Dim pTBL5条諮問総括 As New DataTable("5条諮問総括")
                pTBL5条諮問総括.Columns.Add(New DataColumn("ID", GetType(Integer)))
                pTBL5条諮問総括.Columns.Add(New DataColumn("名称", GetType(String)))
                Dim pCt As New DataColumn("Count", GetType(Integer))
                pCt.DefaultValue = 0
                pTBL5条諮問総括.Columns.Add(pCt)
                pTBL5条諮問総括.PrimaryKey = New DataColumn() {pTBL5条諮問総括.Columns("ID")}
                sError = "005-001"

                nLoop = -1
                If pView.Count > 0 Then
                    pDataCreater.Maximum = pView.Count
                    pDataCreater.Value = 0
                    sError = "005-002"
                    For Each pRow As DataRowView In pView
                        sError = "005-003"

                        If pRow.Item("選択") Then
                            sError = "005-004"
                            nLoop += 1
                            pDataCreater.Message = "諮問総括表 処理中 (" & nLoop + 1 & "/" & pView.Count & ")"
                            pDataCreater.Value = nLoop
                            Dim n As Integer = Val(pRow.Item("申請後農地分類").ToString)

                            Dim pRowT As DataRow = pTBL5条諮問総括.Rows.Find(n)
                            If pRowT Is Nothing Then

                                pRowT = pTBL5条諮問総括.NewRow
                                pRowT.Item("ID") = n

                                pRowT.Item("Count") = 0
                                pTBL5条諮問総括.Rows.Add(pRowT)
                            End If
                            pRowT.Item("Count") += 1
                        End If
                    Next
                    For n = 1 To 10
                        Dim pRowT As DataRow = pTBL5条諮問総括.Rows.Find(n)
                        If pRowT Is Nothing Then
                            pSheet.ValueReplace(Replace("{5条件数00}", "00", Strings.Right("00" & n, 2)), "0")
                        Else
                            pSheet.ValueReplace(Replace("{5条件数00}", "00", Strings.Right("00" & n, 2)), pRowT.Item("Count"))
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(sError & ":" & ex.Message)
            Stop
        End Try
    End Sub
End Class

Class CPrint諮問農地区分別一覧表
    Inherits CPrint諮問意見書単票作成

    Public Overrides Sub SetDataRow(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRow As System.Data.DataRowView, Optional ByVal pObj As Object = Nothing)

    End Sub

    Public Overloads Overrides Sub Set単票Data(ByRef pDataCreater As C諮問意見書Data作成, ByVal pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet)
        pSheet.ValueReplace("{市町村名}", SysAD.市町村.市町村名)
        Dim sError As String = ""
        Try
            If pDataCreater.TabCtrl.TabPages.ContainsKey("n.40") Then
                Dim pTB４条 As 諮問意見書Page = pDataCreater.TabCtrl.TabPages("n.40")
                Dim pTable As DataTable = CType(pTB４条.List.DataSource, DataView).ToTable
                Dim pView As New DataView(pTable, "[選択]=True", "", DataViewRowState.CurrentRows)

                Dim pTBL4条諮問総括 As New DataTable("4条諮問総括")
                pTBL4条諮問総括.Columns.Add(New DataColumn("ID", GetType(Integer)))
                pTBL4条諮問総括.Columns.Add(New DataColumn("名称", GetType(String)))
                Dim pCt As New DataColumn("Count", GetType(Integer))
                pCt.DefaultValue = 0
                pTBL4条諮問総括.Columns.Add(pCt)
                Dim pCt2 As New DataColumn("面積", GetType(Decimal))
                pCt2.DefaultValue = 0
                pTBL4条諮問総括.Columns.Add(pCt2)
                pTBL4条諮問総括.PrimaryKey = New DataColumn() {pTBL4条諮問総括.Columns("ID")}
                If pView.Count > 0 Then
                    nLoop = -1
                    pDataCreater.Maximum = pView.Count + 1
                    pDataCreater.Value = 0
                    For Each pRow As DataRowView In pView
                        If pRow.Item("選択") Then
                            nLoop += 1
                            pDataCreater.Message = "諮問総括表 処理中 (" & nLoop + 1 & "/" & pView.Count & ")"
                            pDataCreater.Value = nLoop
                            Dim n As Integer = Val(pRow.Item("農地区分").ToString)
                            Dim sNList As String = pRow.Item("農地リスト").ToString
                            Dim sNID As String = ""

                            Dim Ar As String() = Split(sNList, ";")
                            For Each sKey As String In Ar

                                If sKey.StartsWith("農地.") Then
                                    sNID &= IIf(sNID.Length > 0, ",", "") & sKey.Substring(3)
                                ElseIf sKey.StartsWith("転用農地.") Then
                                    sNID &= IIf(sNID.Length > 0, ",", "") & sKey.Substring(5)
                                Else
                                    Stop
                                End If

                            Next

                            App農地基本台帳.TBL農地.MergePlus(SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE [ID] In (" & sNID & ")"))
                            Dim pView40 As New DataView(App農地基本台帳.TBL農地.Body, "[ID] In (" & sNID & ")", "", DataViewRowState.CurrentRows)
                            Dim nArea As Decimal = 0
                            For Each pRowV As DataRowView In pView40
                                nArea += Val(pRowV.Item("登記簿面積").ToString)
                            Next
                            App農地基本台帳.TBL転用農地.MergePlus(SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D_転用農地] WHERE [ID] In (" & sNID & ")"))
                            Dim pViewT40 As New DataView(App農地基本台帳.TBL転用農地.Body, "[ID] In (" & sNID & ")", "", DataViewRowState.CurrentRows)
                            For Each pRowV As DataRowView In pViewT40
                                nArea += Val(pRowV.Item("登記簿面積").ToString)
                            Next

                            Select Case n
                                Case 1, 4, 5
                                    n = 1
                                Case 2, 3, 6
                                    n = 2
                            End Select

                            Dim pRowT As DataRow = pTBL4条諮問総括.Rows.Find(n)
                            If pRowT Is Nothing Then

                                pRowT = pTBL4条諮問総括.NewRow
                                pRowT.Item("ID") = n
                                pRowT.Item("Count") = 1
                                pRowT.Item("面積") = nArea

                                pTBL4条諮問総括.Rows.Add(pRowT)
                            Else
                                pRowT.Item("Count") += 1
                                pRowT.Item("面積") += nArea
                            End If
                        End If
                    Next
                    For n = 1 To 2
                        Dim pRowT As DataRow = pTBL4条諮問総括.Rows.Find(n)
                        If pRowT Is Nothing Then
                            pSheet.ValueReplace(Replace("{4条件数00}", "00", Strings.Right("00" & n, 2)), 0)
                            pSheet.ValueReplace(Replace("{4条面積00}", "00", Strings.Right("00" & n, 2)), 0)
                        Else
                            pSheet.ValueReplace(Replace("{4条件数00}", "00", Strings.Right("00" & n, 2)), pRowT.Item("Count"))
                            pSheet.ValueReplace(Replace("{4条面積00}", "00", Strings.Right("00" & n, 2)), pRowT.Item("面積"))
                        End If
                    Next
                End If
            End If

            If pDataCreater.TabCtrl.TabPages.ContainsKey("n.50") Then
                Dim pTB5条 As 諮問意見書Page = pDataCreater.TabCtrl.TabPages("n.50")
                Dim pTable As DataTable = CType(pTB5条.List.DataSource, DataView).ToTable
                Dim pView As New DataView(pTable, "[選択]=True", "", DataViewRowState.CurrentRows)

                Dim pTBL5条諮問総括 As New DataTable("5条諮問総括")
                pTBL5条諮問総括.Columns.Add(New DataColumn("ID", GetType(Integer)))
                pTBL5条諮問総括.Columns.Add(New DataColumn("名称", GetType(String)))
                Dim pCt As New DataColumn("Count", GetType(Integer))
                pCt.DefaultValue = 0
                pTBL5条諮問総括.Columns.Add(pCt)
                Dim pCt2 As New DataColumn("面積", GetType(Decimal))
                pCt2.DefaultValue = 0
                pTBL5条諮問総括.Columns.Add(pCt2)
                pTBL5条諮問総括.PrimaryKey = New DataColumn() {pTBL5条諮問総括.Columns("ID")}
                nLoop = -1

                If pView.Count > 0 Then
                    pDataCreater.Maximum = pView.Count + 1
                    pDataCreater.Value = 0
                    For Each pRow As DataRowView In pView
                        If pRow.Item("選択") Then
                            nLoop += 1
                            pDataCreater.Message = "諮問総括表 処理中 (" & nLoop + 1 & "/" & pView.Count & ")"
                            pDataCreater.Value = nLoop

                            Dim n As Integer = Val(pRow.Item("農地区分").ToString)
                            Dim sNList As String = pRow.Item("農地リスト").ToString
                            Dim sNID As String = ""

                            Dim Ar As String() = Split(sNList, ";")
                            For Each sKey As String In Ar

                                If sKey.StartsWith("農地.") Then
                                    sNID &= IIf(sNID.Length > 0, ",", "") & sKey.Substring(3)
                                ElseIf sKey.StartsWith("転用農地.") Then
                                    sNID &= IIf(sNID.Length > 0, ",", "") & sKey.Substring(5)
                                Else
                                    Stop
                                End If

                            Next

                            App農地基本台帳.TBL農地.MergePlus(SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE [ID] In (" & sNID & ")"))
                            Dim pView50 As New DataView(App農地基本台帳.TBL農地.Body, "[ID] In (" & sNID & ")", "", DataViewRowState.CurrentRows)
                            Dim nArea As Decimal = 0
                            For Each pRowV As DataRowView In pView50
                                nArea += Val(pRowV.Item("登記簿面積").ToString)
                            Next
                            App農地基本台帳.TBL転用農地.MergePlus(SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D_転用農地] WHERE [ID] In (" & sNID & ")"))
                            Dim pViewT50 As New DataView(App農地基本台帳.TBL転用農地.Body, "[ID] In (" & sNID & ")", "", DataViewRowState.CurrentRows)
                            For Each pRowV As DataRowView In pViewT50
                                nArea += Val(pRowV.Item("登記簿面積").ToString)
                            Next

                            Select Case n
                                Case 1, 4, 5
                                    n = 1
                                Case 2, 3, 6
                                    n = 2
                            End Select

                            Dim pRowT As DataRow = pTBL5条諮問総括.Rows.Find(n)
                            If pRowT Is Nothing Then

                                pRowT = pTBL5条諮問総括.NewRow
                                pRowT.Item("ID") = n
                                pRowT.Item("Count") = 1
                                pRowT.Item("面積") = nArea

                                pTBL5条諮問総括.Rows.Add(pRowT)
                            Else
                                pRowT.Item("Count") += 1
                                pRowT.Item("面積") += nArea
                            End If
                        End If
                    Next
                    For n = 1 To 2
                        Dim pRowT As DataRow = pTBL5条諮問総括.Rows.Find(n)
                        If pRowT Is Nothing Then
                            pSheet.ValueReplace(Replace("{5条件数00}", "00", Strings.Right("00" & n, 2)), "0")
                            pSheet.ValueReplace(Replace("{5条面積00}", "00", Strings.Right("00" & n, 2)), "0")
                        Else
                            pSheet.ValueReplace(Replace("{5条件数00}", "00", Strings.Right("00" & n, 2)), pRowT.Item("Count"))
                            pSheet.ValueReplace(Replace("{5条面積00}", "00", Strings.Right("00" & n, 2)), pRowT.Item("面積"))
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            Stop
        End Try
    End Sub
End Class

Public MustInherit Class CPrint諮問共通
    Inherits CPrint諮問意見書資料作成
    Protected mvarSheet As XMLSSWorkSheet

    Protected Sub 諮問共通(ByVal pRow As HimTools2012.Data.DataRowPlus)
        If Not IsDBNull(pRow.Item("総会日")) AndAlso IsDate(pRow.Item("総会日")) Then
            mvarSheet.ValueReplace("{総会日}", 和暦Format(pRow.Item("総会日").ToString))
        Else
            mvarSheet.ValueReplace("{総会日}", "")
        End If


        Select Case Val(pRow.Item("農地区分").ToString)
            Case 1
                mvarSheet.ValueReplace("{農地区分01}", "○")
                mvarSheet.ValueReplace("{農地区分}", "第１種農地")
            Case 2, 6
                mvarSheet.ValueReplace("{農地区分02}", "○")
                mvarSheet.ValueReplace("{農地区分}", "第２種農地")
            Case 3
                mvarSheet.ValueReplace("{農地区分03}", "○")
                mvarSheet.ValueReplace("{農地区分}", "第３種農地")
            Case 4, 5
                mvarSheet.ValueReplace("{農地区分04}", "○")
                mvarSheet.ValueReplace("{農地区分}", "農振農用地")
        End Select
        mvarSheet.ValueReplace("{農地区分01}", "")
        mvarSheet.ValueReplace("{農地区分02}", "")
        mvarSheet.ValueReplace("{農地区分03}", "")
        mvarSheet.ValueReplace("{農地区分04}", "")
        mvarSheet.ValueReplace("{農地区分}", "")

        Select Case pRow.Item("申請後農地分類", 0)
            Case 1 : mvarSheet.ValueReplace("{転用目的01}", "○")
            Case 2 : mvarSheet.ValueReplace("{転用目的02}", "○")
            Case 3 : mvarSheet.ValueReplace("{転用目的03}", "○")
            Case Else
                mvarSheet.ValueReplace("{転用目的その他}", pRow.Item("申請理由A").ToString)
        End Select
        mvarSheet.ValueReplace("{転用目的01}", "")
        mvarSheet.ValueReplace("{転用目的02}", "")
        mvarSheet.ValueReplace("{転用目的03}", "")
        mvarSheet.ValueReplace("{転用目的その他}", "")

        If pRow.Item("用途").ToString.Length > 0 Then
            mvarSheet.ValueReplace("{用途}", pRow.Item("用途").ToString)
        Else
            mvarSheet.ValueReplace("{用途}", "")
        End If

        Select Case CType(pRow.Item("法令"), enum法令)
            Case enum法令.農地法5条所有権
                mvarSheet.ValueReplace("{権利内容}", "所有権移転")
            Case enum法令.農地法5条貸借
                Select Case Val(pRow.Item("権利種類").ToString)
                    Case 1 : mvarSheet.ValueReplace("{権利内容}", "賃借権")
                    Case 2 : mvarSheet.ValueReplace("{権利内容}", "使用貸借権")
                    Case Else : mvarSheet.ValueReplace("{権利内容}", "その他")
                End Select
            Case enum法令.農地法5条一時転用
                mvarSheet.ValueReplace("{権利内容}", "期間借地")
        End Select

        Dim wSt1 As String = IIf(pRow.Item("工事開始年1").ToString.Length > 0, pRow.Item("工事開始年1"), "　") & "年" & IIf(pRow.Item("工事開始月1").ToString.Length > 0, pRow.Item("工事開始月1"), "　") & "月"
        Dim wSt2 As String = IIf(pRow.Item("工事終了年1").ToString.Length > 0, pRow.Item("工事終了年1"), "　") & "年" & IIf(pRow.Item("工事終了月1").ToString.Length > 0, pRow.Item("工事終了月1"), "　") & "月"

        mvarSheet.ValueReplace("{工事開始}", wSt1)
        mvarSheet.ValueReplace("{工事終了}", wSt2)

        If pRow.Item("数量", 0) > 0 Then
            mvarSheet.ValueReplace("{棟数}", pRow.Item("数量").ToString)
        Else
            mvarSheet.ValueReplace("{棟数}", "")
        End If
        If pRow.Item("建築面積").ToString.Length > 0 AndAlso pRow.Item("建築面積") > 0 Then
            mvarSheet.ValueReplace("{建築面積}", pRow.Item("建築面積").ToString)
        Else
            mvarSheet.ValueReplace("{建築面積}", "")
        End If

        If Val(pRow.Item("都市計画区分").ToString) > 0 Then
            mvarSheet.ValueReplace("{土地利用規制01}", "○")
        Else
            mvarSheet.ValueReplace("{土地利用規制01}", "")
        End If
        If Val(pRow.Item("農振区分").ToString) > 0 Then
            mvarSheet.ValueReplace("{土地利用規制02}", "○")
        Else
            mvarSheet.ValueReplace("{土地利用規制02}", "")
        End If
        If Val(pRow.Item("土地改良事業の有無").ToString) <> 0 Then
            mvarSheet.ValueReplace("{土地利用規制03}", "○")
        Else
            mvarSheet.ValueReplace("{土地利用規制03}", "")
        End If
        mvarSheet.ValueReplace("{備考}", pRow.Item("備考").ToString)
    End Sub
End Class

Class CPrint諮問資料4条
    Inherits CPrint諮問共通

    Public Overrides Sub SetDataRow(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRowX As System.Data.DataRowView, Optional ByVal pObj As Object = Nothing)
        mvarSheet = pSheet
        '複数申請人A
        Dim sError As String = ""
        Try
            Dim pRow As New HimTools2012.Data.DataRowPlus(pRowX.Row)
            pRow.Body.Item("諮問番号") = SetNO(pSheet, False)
            申請者A(pSheet, pRowX)
            pSheet.ValueReplace("{農地所在地}", SysAD.市町村.市町村名)

            Dim p集計結果 As C筆明細と集計作成 = 複数土地設定(pSheet, pRowX, Nothing)
            pRow.Body.Item("総面積") = p集計結果.総面積


            諮問共通(pRow)

        Catch ex As Exception
            Stop
        End Try
    End Sub
End Class

Class CPrint諮問資料5条
    Inherits CPrint諮問共通

    Public Overloads Overrides Sub SetDataRow(ByRef pSheet As XMLSSWorkSheet, ByRef pRowX As System.Data.DataRowView, Optional ByVal pObj As Object = Nothing)
        mvarSheet = pSheet
        '複数申請人A
        Dim sError As String = ""
        Try
            Dim pRow As New HimTools2012.Data.DataRowPlus(pRowX.Row)
            pRow.Body.Item("諮問番号") = SetNO(pSheet, False)
            申請者A(pSheet, pRowX)
            申請者B(pSheet, pRowX)
            pSheet.ValueReplace("{農地所在地}", SysAD.市町村.市町村名)
            pSheet.ValueReplace("{備考}", pRow.Item("備考").ToString)
            Dim p集計結果 As C筆明細と集計作成 = 複数土地設定(pSheet, pRowX, Nothing)
            pRow.Body.Item("総面積") = p集計結果.総面積

            Select Case pRow.Item("法令")
                Case enum法令.農地法5条所有権
                    pSheet.ValueReplace("{権利01}", "○")
                Case enum法令.農地法5条貸借, enum法令.農地法5条一時転用
                    Select Case pRow.Item("権利種類", 0)
                        Case 1 : pSheet.ValueReplace("{権利02}", "○")
                        Case 2 : pSheet.ValueReplace("{権利03}", "○")
                    End Select
            End Select
            pSheet.ValueReplace("{権利01}", "")
            pSheet.ValueReplace("{権利02}", "")
            pSheet.ValueReplace("{権利03}", "")


            諮問共通(pRow)
        Catch ex As Exception
            Stop
        End Try
    End Sub
End Class

Class CPrint諮問資料個別説明
    Inherits CPrint諮問意見書資料作成

    Private mvarDataCreater As C諮問意見書Data作成
    Public Sub New(ByRef pDataCreater As C諮問意見書Data作成)
        mvarDataCreater = pDataCreater
    End Sub

    Public Overrides Function LoopSub(ByVal s処理名称 As String, ByVal sDesktopFolder As String, ByVal p作成 As CPrint諮問意見書作成共通, ByVal sFile As String, ByVal ParamArray pTabs() As Object) As Boolean
        Dim sCount As String = ""
        Dim sNewFile As String = HimTools2012.FileManager.NewFileName(sDesktopFolder & "\" & sFile, 0)

        Dim sXML As String = HimTools2012.TextAdapter.LoadTextFile(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\" & sFile)
        Dim pXMLSS As New CXMLSS2003(sXML)
        Dim pSheet As XMLSSWorkSheet = pXMLSS.WorkBook.WorkSheets.Items("諮問説明資料")

        With pSheet
            If Not InStr(.Table.InnerXML, "{No}") > 0 Then

                Return False
            End If
        End With

        LoopRows = New XMLLoopRows(pSheet)
        mvarDataCreater.Maximum = 0
        nLoop = -1

        If mvarDataCreater.TabCtrl.TabPages.ContainsKey("n.40") Then
            Dim pTB４条 As 諮問意見書Page = mvarDataCreater.TabCtrl.TabPages("n.40")
            Dim pTable As DataTable = CType(pTB４条.List.DataSource, DataView).ToTable
            Dim pView As New DataView(pTable, "([農地区分]=1 Or [農地区分]=5 Or [総面積]>=3000) AND [選択]=True", "受付補助記号,受付番号", DataViewRowState.CurrentRows)
            mvarDataCreater.Maximum += pView.Count

            For Each pRowV As DataRowView In pView
                If nLoop = -1 Then

                Else
                    For Each pXRow As XMLSSRow In LoopRows
                        Dim pCopyRow = pXRow.CopyRow

                        pSheet.Table.Rows.InsertRow(LoopRows.InsetRow, pCopyRow)
                        LoopRows.InsetRow += 1
                    Next
                End If

                nLoop += 1
                mvarDataCreater.Message = s処理名称 & " 処理中 (" & nLoop + 1 & "/" & pView.Count & ")"
                mvarDataCreater.Value = nLoop

                Me.SetDataRow(pSheet, pRowV, "4")
            Next
        End If
        If mvarDataCreater.TabCtrl.TabPages.ContainsKey("n.50") Then
            Dim pTB５条 As 諮問意見書Page = mvarDataCreater.TabCtrl.TabPages("n.50")
            Dim pTable As DataTable = CType(pTB５条.List.DataSource, DataView).ToTable
            Dim pView As New DataView(pTable, "([農地区分]=1 Or [農地区分]=5 Or [総面積]>=3000) AND [選択]=True", "受付補助記号,受付番号", DataViewRowState.CurrentRows)
            mvarDataCreater.Maximum += pView.Count

            For Each pRowV As DataRowView In pView
                If nLoop = -1 Then

                Else
                    For Each pXRow As XMLSSRow In LoopRows
                        Dim pCopyRow = pXRow.CopyRow

                        pSheet.Table.Rows.InsertRow(LoopRows.InsetRow, pCopyRow)
                        LoopRows.InsetRow += 1
                    Next
                End If

                nLoop += 1
                mvarDataCreater.Message = s処理名称 & " 処理中 (" & nLoop + 1 & "/" & pView.Count & ")"
                mvarDataCreater.Value = nLoop

                Me.SetDataRow(pSheet, pRowV, "5")
            Next
        End If
        If mvarDataCreater.Maximum > 0 Then
            HimTools2012.TextAdapter.SaveTextFile(sNewFile, pXMLSS.OutPut(True))
        End If
        Return True
    End Function

    Public Overrides Sub SetDataRow(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRow As System.Data.DataRowView, Optional ByVal pObj As Object = Nothing)
        pSheet.ValueReplace("{説明番号}", SetNO(pSheet, False))
        pSheet.ValueReplace("{条項01}", pObj.ToString)
        pSheet.ValueReplace("{諮問番号01}", pRow.Item("諮問番号").ToString)
        If pRow.Item("不許可例外").ToString.Length > 0 Then
            pSheet.ValueReplace("{不許可例外}", "(" & pRow.Item("不許可例外").ToString & ")")
        Else
            pSheet.ValueReplace("{不許可例外}", "")

        End If

        pSheet.ValueReplace("{市町村名01}", SysAD.市町村.市町村名)

        Select Case pRow.Item("法令")
            Case enum法令.農地法4条, enum法令.農地法4条一時転用
                pSheet.ValueReplace("{転用事業者名01}", pRow.Item("氏名A").ToString)
            Case enum法令.農地法5条所有権, enum法令.農地法5条貸借, enum法令.農地法5条一時転用
                pSheet.ValueReplace("{転用事業者名01}", pRow.Item("氏名B").ToString)
        End Select

        Select Case Val(pRow.Item("農地区分").ToString)
            Case 1 : pSheet.ValueReplace("{農地区分01}", "第1種農地")
            Case 2 : pSheet.ValueReplace("{農地区分01}", "第2種農地")
            Case 3 : pSheet.ValueReplace("{農地区分01}", "第3種農地")
            Case 5 : pSheet.ValueReplace("{農地区分01}", "農用地区域内農地")
            Case Else
                'Stop
        End Select
        Dim s期間 As New System.Text.StringBuilder
        If Not IsDBNull(pRow.Item("始期")) AndAlso pRow.Item("始期") > #1/1/1900# Then
            s期間.Append(和暦Format(pRow.Item("始期"), "gy.M"))
        End If
        If Not IsDBNull(pRow.Item("永久")) AndAlso pRow.Item("永久") Then
            s期間.Append("～(永久)")
        ElseIf Not IsDBNull(pRow.Item("終期")) AndAlso pRow.Item("終期") > #1/1/1900# Then
            s期間.Append("～" & 和暦Format(pRow.Item("終期"), "gy.M"))
        ElseIf s期間.Length > 0 Then
            s期間.Append("～")
        End If
        pSheet.ValueReplace("{期間}", s期間.ToString)


        pSheet.ValueReplace("{転用面積01}", String.Format("{0:0,0}", pRow.Item("総面積")))

        Dim pVX As New DataView(App農地基本台帳.DataMaster.Body, "[Class]='申請後農地分類' AND [ID]=" & Val(pRow.Item("申請後農地分類").ToString), "", DataViewRowState.CurrentRows)
        If pVX.Count = 1 Then
            If pVX(0).Item("名称").ToString = "その他" Then
                pSheet.ValueReplace("{転用目的01}", pRow.Item("申請理由A").ToString())
            Else
                pSheet.ValueReplace("{転用目的01}", pVX(0).Item("名称"))
            End If
        Else
            pSheet.ValueReplace("{転用目的01}", "")
        End If
        pSheet.ValueReplace("{転用目的A}", pRow.Item("申請理由A").ToString())

        pSheet.ValueReplace("{農業01}", "")

        If Not IsDBNull(pRow.Item("土地改良事業の有無")) AndAlso pRow.Item("土地改良事業の有無") Then
            pSheet.ValueReplace("{土地01}", "○")
        Else
            pSheet.ValueReplace("{土地01}", "")
        End If

        If Not IsDBNull(pRow.Item("農地の広がり")) AndAlso pRow.Item("農地の広がり") = 1 Then
            pSheet.ValueReplace("{集団01}", "○")
        Else
            pSheet.ValueReplace("{集団01}", "")
        End If
    End Sub
End Class

Class CPrint諮問説明資料
    Inherits CPrint諮問意見書資料作成
    Private mvarDataCreater As C諮問意見書Data作成
    Public Sub New(ByRef pDataCreater As C諮問意見書Data作成)
        mvarDataCreater = pDataCreater
    End Sub

    Public Overrides Function LoopSub(ByVal s処理名称 As String, ByVal sDesktopFolder As String, ByVal p作成 As CPrint諮問意見書作成共通, ByVal sFile As String, ByVal ParamArray pTabs() As Object) As Boolean
        Dim sCount As String = ""
        Dim sNewFile As String = HimTools2012.FileManager.NewFileName(sDesktopFolder & "\" & sFile, 0)

        Dim sXML As String = HimTools2012.TextAdapter.LoadTextFile(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\" & sFile)
        Dim pXMLSS As New CXMLSS2003(sXML)
        Dim pSheet As XMLSSWorkSheet = pXMLSS.WorkBook.WorkSheets.Items("諮問説明資料")

        With pSheet
            If Not InStr(.Table.InnerXML, "{No}") > 0 Then

                Return False
            End If
        End With

        LoopRows = New XMLLoopRows(pSheet)
        mvarDataCreater.Maximum = 0
        nLoop = -1

        If mvarDataCreater.TabCtrl.TabPages.ContainsKey("n.40") Then
            Dim pTB４条 As 諮問意見書Page = mvarDataCreater.TabCtrl.TabPages("n.40")
            Dim pTable As DataTable = CType(pTB４条.List.DataSource, DataView).ToTable
            Dim pView As New DataView(pTable, "([意見聴取案件]>0 Or [農地区分]=1 Or [農地区分]=5 Or [総面積]>=3000) AND [選択]=True", "受付補助記号,受付番号", DataViewRowState.CurrentRows)
            mvarDataCreater.Maximum += pView.Count

            For Each pRowV As DataRowView In pView
                If nLoop = -1 Then

                Else
                    For Each pXRow As XMLSSRow In LoopRows
                        Dim pCopyRow = pXRow.CopyRow

                        pSheet.Table.Rows.InsertRow(LoopRows.InsetRow, pCopyRow)
                        LoopRows.InsetRow += 1
                    Next
                End If

                nLoop += 1
                mvarDataCreater.Message = s処理名称 & " 処理中 (" & nLoop + 1 & "/" & pView.Count & ")"
                mvarDataCreater.Value = nLoop

                Me.SetDataRow(pSheet, pRowV, "4")
            Next
        End If
        If mvarDataCreater.TabCtrl.TabPages.ContainsKey("n.50") Then
            Dim pTB５条 As 諮問意見書Page = mvarDataCreater.TabCtrl.TabPages("n.50")
            Dim pTable As DataTable = CType(pTB５条.List.DataSource, DataView).ToTable
            Dim pView As New DataView(pTable, "([意見聴取案件]>0 Or [農地区分]=1 Or [農地区分]=5 Or [総面積]>=3000) AND [選択]=True", "受付補助記号,受付番号", DataViewRowState.CurrentRows)
            mvarDataCreater.Maximum += pView.Count

            For Each pRowV As DataRowView In pView
                If nLoop = -1 Then

                Else
                    For Each pXRow As XMLSSRow In LoopRows
                        Dim pCopyRow = pXRow.CopyRow

                        pSheet.Table.Rows.InsertRow(LoopRows.InsetRow, pCopyRow)
                        LoopRows.InsetRow += 1
                    Next
                End If

                nLoop += 1
                mvarDataCreater.Message = s処理名称 & " 処理中 (" & nLoop + 1 & "/" & pView.Count & ")"
                mvarDataCreater.Value = nLoop

                Me.SetDataRow(pSheet, pRowV, "5")
            Next
        End If
        If mvarDataCreater.Maximum > 0 Then
            HimTools2012.TextAdapter.SaveTextFile(sNewFile, pXMLSS.OutPut(True))
        End If
        Return True
    End Function

    Public Overrides Sub SetDataRow(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRow As System.Data.DataRowView, Optional ByVal pObj As Object = Nothing)
        pSheet.ValueReplace("{説明番号}", SetNO(pSheet, False))
        pSheet.ValueReplace("{条項01}", pObj.ToString)
        pSheet.ValueReplace("{諮問番号01}", nLoop + 1)
        pSheet.ValueReplace("{市町村}", SysAD.市町村.市町村名)
        pSheet.ValueReplace("{農地所在地}", SysAD.市町村.市町村名)


        Select Case pRow.Item("法令")
            Case enum法令.農地法4条, enum法令.農地法4条一時転用
                pSheet.ValueReplace("{条項}", "４条")
                pSheet.ValueReplace("{申請者氏名}", pRow.Item("氏名A").ToString)
                pSheet.ValueReplace("{申請者Ａ氏名}", "")
                pSheet.ValueReplace("{申請者Ｂ氏名}", "")
            Case enum法令.農地法5条所有権, enum法令.農地法5条貸借, enum法令.農地法5条一時転用
                pSheet.ValueReplace("{条項}", "５条")
                pSheet.ValueReplace("{申請者氏名}", "")
                pSheet.ValueReplace("{申請者Ａ氏名}", pRow.Item("氏名A").ToString)
                pSheet.ValueReplace("{申請者Ｂ氏名}", pRow.Item("氏名B").ToString)
        End Select



        Dim s期間 As New System.Text.StringBuilder
        If Not IsDBNull(pRow.Item("始期")) AndAlso pRow.Item("始期") > #1/1/1900# Then
            s期間.Append(和暦Format(pRow.Item("始期"), "gy.M"))
        End If
        If Not IsDBNull(pRow.Item("永久")) AndAlso pRow.Item("永久") Then
            s期間.Append("～(永久)")
        ElseIf Not IsDBNull(pRow.Item("終期")) AndAlso pRow.Item("終期") > #1/1/1900# Then
            s期間.Append("～" & 和暦Format(pRow.Item("終期"), "gy.M"))
        ElseIf s期間.Length > 0 Then
            s期間.Append("～")
        End If
        pSheet.ValueReplace("{期間}", s期間.ToString)

        Dim p集計結果 As C筆明細と集計作成 = 複数土地設定(pSheet, pRow, Nothing)

        Dim s地目 As New System.Text.StringBuilder
        Dim s面積 As New System.Text.StringBuilder

        If p集計結果.田面計 > 0 Then
            s地目.Append("田")
            s面積.Append(String.Format("{0:0,0}㎡", p集計結果.田面計))
        End If
        If p集計結果.畑面計 > 0 Then
            s地目.Append(IIf(s地目.Length > 0, "&#10;&#10;", "") & "畑")
            s面積.Append(IIf(s面積.Length > 0, "&#10;&#10;", "") & String.Format("{0:0,0}㎡", p集計結果.畑面計))
        End If
        If p集計結果.田面計 > 0 AndAlso p集計結果.畑面計 > 0 Then
            s地目.Append("&#10;&#10;計")
            s面積.Append(String.Format("&#10;&#10;{0:0,0}㎡", p集計結果.総面積))
        End If

        pSheet.ValueReplace("{地目資}", s地目.ToString())
        pSheet.ValueReplace("{転用面積01}", s面積.ToString())

        Select Case Val(pRow.Item("申請後農地分類").ToString)
            Case 1 : pSheet.ValueReplace("{転用目的01}", "○")
            Case 2 : pSheet.ValueReplace("{転用目的02}", "○")
            Case 3 : pSheet.ValueReplace("{転用目的03}", "○")
            Case Else
                pSheet.ValueReplace("{転用目的その他}", pRow.Item("申請理由A").ToString)
        End Select
        pSheet.ValueReplace("{転用目的01}", "")
        pSheet.ValueReplace("{転用目的02}", "")
        pSheet.ValueReplace("{転用目的03}", "")
        pSheet.ValueReplace("{転用目的その他}", "")

        Select Case Val(pRow.Item("農地区分").ToString)
            Case 1 : pSheet.ValueReplace("{農地区分01}", "○")
            Case 2, 6 : pSheet.ValueReplace("{農地区分02}", "○")
            Case 3 : pSheet.ValueReplace("{農地区分03}", "○")
            Case 4, 5 : pSheet.ValueReplace("{農地区分04}", "○")
        End Select
        pSheet.ValueReplace("{農地区分01}", "")
        pSheet.ValueReplace("{農地区分02}", "")
        pSheet.ValueReplace("{農地区分03}", "")
        pSheet.ValueReplace("{農地区分04}", "")

        Select Case pRow.Item("法令")
            Case enum法令.農地法5条所有権
                pSheet.ValueReplace("{権利01}", "○")
            Case enum法令.農地法5条貸借, enum法令.農地法5条一時転用
                Select Case pRow.Item("権利種類")
                    Case 1 : pSheet.ValueReplace("{権利02}", "○")
                    Case 2 : pSheet.ValueReplace("{権利03}", "○")
                End Select
        End Select
        pSheet.ValueReplace("{権利01}", "")
        pSheet.ValueReplace("{権利02}", "")
        pSheet.ValueReplace("{権利03}", "")
        pSheet.ValueReplace("{権利04}", "")

        If Val(pRow.Item("数量").ToString) > 0 Then
            pSheet.ValueReplace("{棟数}", pRow.Item("数量").ToString)
        Else
            pSheet.ValueReplace("{棟数}", "")
        End If
        If pRow.Item("建築面積").ToString.Length > 0 AndAlso pRow.Item("建築面積") > 0 Then
            pSheet.ValueReplace("{建築面積}", pRow.Item("建築面積").ToString)
        Else
            pSheet.ValueReplace("{建築面積}", "")
        End If

        If Val(pRow.Item("都市計画区分").ToString) > 0 Then
            pSheet.ValueReplace("{土地利用規制01}", "○")
        Else
            pSheet.ValueReplace("{土地利用規制01}", "")
        End If
        If Val(pRow.Item("農振区分").ToString) > 0 Then
            pSheet.ValueReplace("{土地利用規制02}", "○")
        Else
            pSheet.ValueReplace("{土地利用規制02}", "")
        End If
        If Val(pRow.Item("土地改良事業の有無").ToString) <> 0 Then
            pSheet.ValueReplace("{土地利用規制03}", "○")
        Else
            pSheet.ValueReplace("{土地利用規制03}", "")
        End If



        If Not IsDBNull(pRow.Item("土地改良事業の有無")) AndAlso pRow.Item("土地改良事業の有無") Then
            pSheet.ValueReplace("{土地01}", "○")
        Else
            pSheet.ValueReplace("{土地01}", "")
        End If

        If Not IsDBNull(pRow.Item("農地の広がり")) AndAlso pRow.Item("農地の広がり") = 1 Then
            pSheet.ValueReplace("{集団01}", "○")
        Else
            pSheet.ValueReplace("{集団01}", "")
        End If

        pSheet.ValueReplace("{許可}", "○許可")
        pSheet.ValueReplace("{不許可}", "不許可")

        Select Case Val(pRow.Item("農地区分").ToString)
            Case 1, 4, 5
                If pRow.Item("不許可例外").ToString.Length > 0 Then
                    pSheet.ValueReplace("{不許可例外}", pRow.Item("不許可例外").ToString)
                End If
        End Select
        pSheet.ValueReplace("{不許可例外}", "")
        'If pRow.Item("不許可例外").ToString.Length > 0 Then
        '    pSheet.ValueReplace("{不許可例外}", pRow.Item("不許可例外").ToString)
        'Else
        '    pSheet.ValueReplace("{不許可例外}", "")
        'End If
    End Sub
End Class

Class CPrint諮問農地法農地区分別総括表
    Inherits CPrint諮問意見書単票作成

    Private Sub SetTBL4条5条(ByRef pTBL As DataTable)
        pTBL.Columns.Add(New DataColumn("ID", GetType(Integer)))
        pTBL.Columns.Add(New DataColumn("名称", GetType(String)))
        Dim pCt As New DataColumn("Count", GetType(Integer))
        pCt.DefaultValue = 0
        pTBL.Columns.Add(pCt)
        Dim pCt2 As New DataColumn("面積", GetType(Decimal))
        pCt2.DefaultValue = 0
        pTBL.Columns.Add(pCt2)
        Dim pCt3 As New DataColumn("田面積", GetType(Decimal))
        pCt3.DefaultValue = 0
        pTBL.Columns.Add(pCt3)
        Dim pCt4 As New DataColumn("畑面積", GetType(Decimal))
        pCt4.DefaultValue = 0
        pTBL.Columns.Add(pCt4)
        Dim pCt5 As New DataColumn("採草放牧面積", GetType(Decimal))
        pCt5.DefaultValue = 0
        pTBL.Columns.Add(pCt5)
        Dim pCt6 As New DataColumn("3000超件数", GetType(Integer))
        pCt6.DefaultValue = 0
        pTBL.Columns.Add(pCt6)
        Dim pCt7 As New DataColumn("3000超面積", GetType(Decimal))
        pCt7.DefaultValue = 0
        pTBL.Columns.Add(pCt7)
        Dim pCt8 As New DataColumn("3000以下件数", GetType(Integer))
        pCt8.DefaultValue = 0
        pTBL.Columns.Add(pCt8)
        Dim pCt9 As New DataColumn("3000以下面積", GetType(Decimal))
        pCt9.DefaultValue = 0
        pTBL.Columns.Add(pCt9)

        Dim pCt10 As New DataColumn("3000超件数うち農用地", GetType(Decimal))
        pCt10.DefaultValue = 0
        pTBL.Columns.Add(pCt10)
        Dim pCt11 As New DataColumn("3000超面積うち農用地", GetType(Decimal))
        pCt11.DefaultValue = 0
        pTBL.Columns.Add(pCt11)
        Dim pCt12 As New DataColumn("3000以下件数うち農用地", GetType(Decimal))
        pCt12.DefaultValue = 0
        pTBL.Columns.Add(pCt12)
        Dim pCt13 As New DataColumn("3000以下面積うち農用地", GetType(Decimal))
        pCt13.DefaultValue = 0
        pTBL.Columns.Add(pCt13)

        Dim pCt14 As New DataColumn("3000以下全体件数", GetType(Decimal))
        pCt14.DefaultValue = 0
        pTBL.Columns.Add(pCt14)
        Dim pCt15 As New DataColumn("3000以下全体面積", GetType(Decimal))
        pCt15.DefaultValue = 0
        pTBL.Columns.Add(pCt15)

        pTBL.PrimaryKey = New DataColumn() {pTBL.Columns("ID")}
    End Sub

    Public Overrides Sub SetDataRow(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRow As System.Data.DataRowView, Optional ByVal pObj As Object = Nothing)

    End Sub

    Public Overloads Overrides Sub Set単票Data(ByRef pDataCreater As C諮問意見書Data作成, ByVal pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet)
        pSheet.ValueReplace("{市町村名}", SysAD.市町村.市町村名)
        Try
            If pDataCreater.TabCtrl.TabPages.ContainsKey("n.40") Then
                Dim pTB4条 As 諮問意見書Page = pDataCreater.TabCtrl.TabPages("n.40")
                Dim pTable As DataTable = CType(pTB4条.List.DataSource, DataView).ToTable
                Dim pView As New DataView(pTable, "[選択]=True", "", DataViewRowState.CurrentRows)

                Dim pTBL4条諮問農地区分別総括 As New DataTable("4条諮問農地区分別総括")
                SetTBL4条5条(pTBL4条諮問農地区分別総括)

                Dim pTBL4条諮問目的別総括 As New DataTable("4条諮問目的別総括")
                SetTBL4条5条(pTBL4条諮問目的別総括)

                If pView.Count > 0 Then
                    nLoop = -1
                    pDataCreater.Maximum = pView.Count + 1
                    pDataCreater.Value = 0
                    For Each pRow As DataRowView In pView
                        If pRow.Item("選択") And pRow.Item("複数申請人A") = False Then
                            nLoop += 1
                            pDataCreater.Message = "4条諮問総括表 処理中 (" & nLoop + 1 & "/" & pView.Count & ")"
                            pDataCreater.Value = nLoop
                            Dim n As Integer = Val(pRow.Item("農地区分").ToString)
                            Dim i As Integer = Val(pRow.Item("申請後農地分類").ToString)
                            Dim sNList As String = pRow.Item("農地リスト").ToString
                            Dim sNID As String = ""

                            Dim Ar As String() = Split(sNList, ";")
                            For Each sKey As String In Ar
                                If sKey.StartsWith("農地.") Then : sNID &= IIf(sNID.Length > 0, ",", "") & sKey.Substring(3)
                                ElseIf sKey.StartsWith("転用農地.") Then : sNID &= IIf(sNID.Length > 0, ",", "") & sKey.Substring(5)
                                Else : Stop
                                End If
                            Next

                            App農地基本台帳.TBL農地.MergePlus(SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE [ID] In (" & sNID & ")"))
                            Dim pView40 As New DataView(App農地基本台帳.TBL農地.Body, "[ID] In (" & sNID & ")", "", DataViewRowState.CurrentRows)
                            Dim nArea As Decimal = 0
                            Dim n田面積 As Decimal = 0
                            Dim n畑面積 As Decimal = 0

                            For Each pRowV As DataRowView In pView40
                                nArea += Val(pRowV.Item("実面積").ToString)

                                If Left(pRowV.Item("登記簿地目名").ToString, 1) = "田" Then '転用前地目
                                    n田面積 += Val(pRowV.Item("実面積").ToString)
                                ElseIf Left(pRowV.Item("登記簿地目名").ToString, 1) = "畑" Then
                                    n畑面積 += Val(pRowV.Item("実面積").ToString)
                                End If
                            Next

                            App農地基本台帳.TBL転用農地.MergePlus(SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D_転用農地] WHERE [ID] In (" & sNID & ")"))
                            Dim pViewT40 As New DataView(App農地基本台帳.TBL転用農地.Body, "[ID] In (" & sNID & ")", "", DataViewRowState.CurrentRows)
                            For Each pRowV As DataRowView In pViewT40
                                nArea += Val(pRowV.Item("実面積").ToString)

                                If Left(pRowV.Item("登記簿地目名").ToString, 1) = "田" Then
                                    n田面積 += Val(pRowV.Item("実面積").ToString)
                                ElseIf Left(pRowV.Item("登記簿地目名").ToString, 1) = "畑" Then
                                    n畑面積 += Val(pRowV.Item("実面積").ToString)
                                End If
                            Next

                            Select Case n
                                Case 1, 4, 5 : n = 1
                                Case 2, 3, 6 : n = 2
                            End Select

                            Dim pRowN As DataRow = pTBL4条諮問農地区分別総括.Rows.Find(n)
                            If pRowN Is Nothing Then
                                pRowN = pTBL4条諮問農地区分別総括.NewRow
                                pRowN.Item("ID") = n
                                pRowN.Item("Count") = 1
                                pRowN.Item("面積") = nArea

                                If nArea >= 3000 Then
                                    pRowN.Item("3000超件数") = 1
                                    pRowN.Item("3000超面積") = nArea
                                    If Val(pRow.Item("農地区分").ToString) = 1 Or Val(pRow.Item("農地区分").ToString) = 5 Then
                                        pRowN.Item("3000超件数うち農用地") = 1
                                        pRowN.Item("3000超面積うち農用地") = nArea
                                    End If
                                Else
                                    pRowN.Item("3000以下全体件数") = 1
                                    pRowN.Item("3000以下全体面積") = nArea

                                    If Val(pRow.Item("意見聴取案件").ToString) > 0 Or Val(pRow.Item("農地区分").ToString) = 1 Or Val(pRow.Item("農地区分").ToString) = 5 Then
                                        pRowN.Item("3000以下件数") = 1
                                        pRowN.Item("3000以下面積") = nArea
                                        If Val(pRow.Item("農地区分").ToString) = 1 Or Val(pRow.Item("農地区分").ToString) = 5 Then
                                            pRowN.Item("3000以下件数うち農用地") = 1
                                            pRowN.Item("3000以下面積うち農用地") = nArea
                                        End If
                                    End If
                                End If

                                pTBL4条諮問農地区分別総括.Rows.Add(pRowN)
                            Else
                                pRowN.Item("Count") += 1
                                pRowN.Item("面積") += nArea

                                If nArea >= 3000 Then
                                    pRowN.Item("3000超件数") += 1
                                    pRowN.Item("3000超面積") += nArea
                                    If Val(pRow.Item("農地区分").ToString) = 1 Or Val(pRow.Item("農地区分").ToString) = 5 Then
                                        pRowN.Item("3000超件数うち農用地") += 1
                                        pRowN.Item("3000超面積うち農用地") += nArea
                                    End If
                                Else
                                    pRowN.Item("3000以下全体件数") += 1
                                    pRowN.Item("3000以下全体面積") += nArea

                                    If Val(pRow.Item("意見聴取案件").ToString) > 0 Or Val(pRow.Item("農地区分").ToString) = 1 Or Val(pRow.Item("農地区分").ToString) = 5 Then
                                        pRowN.Item("3000以下件数") += 1
                                        pRowN.Item("3000以下面積") += nArea
                                        If Val(pRow.Item("農地区分").ToString) = 1 Or Val(pRow.Item("農地区分").ToString) = 5 Then
                                            pRowN.Item("3000以下件数うち農用地") += 1
                                            pRowN.Item("3000以下面積うち農用地") += nArea
                                        End If
                                    End If
                                End If
                            End If

                            Select Case i
                                Case 11 : i = 10
                                Case Else
                            End Select

                            Dim pRowI As DataRow = pTBL4条諮問目的別総括.Rows.Find(i)
                            If pRowI Is Nothing Then
                                pRowI = pTBL4条諮問目的別総括.NewRow
                                pRowI.Item("ID") = i
                                pRowI.Item("Count") = 1
                                pRowI.Item("田面積") = n田面積
                                pRowI.Item("畑面積") = n畑面積

                                If nArea >= 3000 Then
                                    pRowI.Item("3000超件数") = 1
                                    pRowI.Item("3000超面積") = nArea
                                    If Val(pRow.Item("農地区分").ToString) = 1 Or Val(pRow.Item("農地区分").ToString) = 5 Then
                                        pRowI.Item("3000超件数うち農用地") = 1
                                        pRowI.Item("3000超面積うち農用地") = nArea
                                    End If
                                Else
                                    pRowI.Item("3000以下全体件数") = 1
                                    pRowI.Item("3000以下全体面積") = nArea

                                    If Val(pRow.Item("意見聴取案件").ToString) > 0 Or Val(pRow.Item("農地区分").ToString) = 1 Or Val(pRow.Item("農地区分").ToString) = 5 Then
                                        pRowI.Item("3000以下件数") = 1
                                        pRowI.Item("3000以下面積") = nArea
                                        If Val(pRow.Item("農地区分").ToString) = 1 Or Val(pRow.Item("農地区分").ToString) = 5 Then
                                            pRowI.Item("3000以下件数うち農用地") = 1
                                            pRowI.Item("3000以下面積うち農用地") = nArea
                                        End If
                                    End If
                                End If
                                pTBL4条諮問目的別総括.Rows.Add(pRowI)
                            Else
                                pRowI.Item("Count") += 1
                                pRowI.Item("田面積") += n田面積
                                pRowI.Item("畑面積") += n畑面積

                                If nArea >= 3000 Then
                                    pRowI.Item("3000超件数") += 1
                                    pRowI.Item("3000超面積") += nArea
                                    If Val(pRow.Item("農地区分").ToString) = 1 Or Val(pRow.Item("農地区分").ToString) = 5 Then
                                        pRowI.Item("3000超件数うち農用地") += 1
                                        pRowI.Item("3000超面積うち農用地") += nArea
                                    End If
                                Else
                                    pRowI.Item("3000以下全体件数") += 1
                                    pRowI.Item("3000以下全体面積") += nArea

                                    If Val(pRow.Item("意見聴取案件").ToString) > 0 Or Val(pRow.Item("農地区分").ToString) = 1 Or Val(pRow.Item("農地区分").ToString) = 5 Then
                                        pRowI.Item("3000以下件数") += 1
                                        pRowI.Item("3000以下面積") += nArea
                                        If Val(pRow.Item("農地区分").ToString) = 1 Or Val(pRow.Item("農地区分").ToString) = 5 Then
                                            pRowI.Item("3000以下件数うち農用地") += 1
                                            pRowI.Item("3000以下面積うち農用地") += nArea
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    For n = 1 To 2
                        Dim pRowN As DataRow = pTBL4条諮問農地区分別総括.Rows.Find(n)
                        If pRowN Is Nothing Then
                            pSheet.ValueReplace(Replace("{4条農地区分別件数00}", "00", Strings.Right("00" & n, 2)), 0)
                            pSheet.ValueReplace(Replace("{4条農地区分別面積00}", "00", Strings.Right("00" & n, 2)), 0)
                            pSheet.ValueReplace(Replace("{4条農地区分別↑件数00}", "00", Strings.Right("00" & n, 2)), 0)
                            pSheet.ValueReplace(Replace("{4条農地区分別↑面積00}", "00", Strings.Right("00" & n, 2)), 0)
                            pSheet.ValueReplace(Replace("{4条農地区分別↓件数00}", "00", Strings.Right("00" & n, 2)), 0)
                            pSheet.ValueReplace(Replace("{4条農地区分別↓面積00}", "00", Strings.Right("00" & n, 2)), 0)
                            pSheet.ValueReplace(Replace("{4条農地区分別↑うち農用地件数00}", "00", Strings.Right("00" & n, 2)), 0)
                            pSheet.ValueReplace(Replace("{4条農地区分別↑うち農用地面積00}", "00", Strings.Right("00" & n, 2)), 0)
                            pSheet.ValueReplace(Replace("{4条農地区分別↓うち農用地件数00}", "00", Strings.Right("00" & n, 2)), 0)
                            pSheet.ValueReplace(Replace("{4条農地区分別↓うち農用地面積00}", "00", Strings.Right("00" & n, 2)), 0)

                        Else
                            pSheet.ValueReplace(Replace("{4条農地区分別件数00}", "00", Strings.Right("00" & n, 2)), pRowN.Item("Count"))
                            pSheet.ValueReplace(Replace("{4条農地区分別面積00}", "00", Strings.Right("00" & n, 2)), pRowN.Item("面積"))
                            pSheet.ValueReplace(Replace("{4条農地区分別↑件数00}", "00", Strings.Right("00" & n, 2)), pRowN.Item("3000超件数"))
                            pSheet.ValueReplace(Replace("{4条農地区分別↑面積00}", "00", Strings.Right("00" & n, 2)), pRowN.Item("3000超面積"))
                            pSheet.ValueReplace(Replace("{4条農地区分別↓件数00}", "00", Strings.Right("00" & n, 2)), pRowN.Item("3000以下件数"))
                            pSheet.ValueReplace(Replace("{4条農地区分別↓面積00}", "00", Strings.Right("00" & n, 2)), pRowN.Item("3000以下面積"))
                            pSheet.ValueReplace(Replace("{4条農地区分別↑うち農用地件数00}", "00", Strings.Right("00" & n, 2)), pRowN.Item("3000超件数うち農用地"))
                            pSheet.ValueReplace(Replace("{4条農地区分別↑うち農用地面積00}", "00", Strings.Right("00" & n, 2)), pRowN.Item("3000超面積うち農用地"))
                            pSheet.ValueReplace(Replace("{4条農地区分別↓うち農用地件数00}", "00", Strings.Right("00" & n, 2)), pRowN.Item("3000以下件数うち農用地"))
                            pSheet.ValueReplace(Replace("{4条農地区分別↓うち農用地面積00}", "00", Strings.Right("00" & n, 2)), pRowN.Item("3000以下面積うち農用地"))
                        End If
                    Next

                    For i = 1 To 12
                        Dim pRowI As DataRow = pTBL4条諮問目的別総括.Rows.Find(i)
                        If pRowI Is Nothing Then
                            pSheet.ValueReplace(Replace("{4条目的別件数00}", "00", Strings.Right("00" & i, 2)), "0")
                            pSheet.ValueReplace(Replace("{4条目的別↑件数00}", "00", Strings.Right("00" & i, 2)), "0")
                            pSheet.ValueReplace(Replace("{4条目的別↓件数00}", "00", Strings.Right("00" & i, 2)), "0")
                            pSheet.ValueReplace(Replace("{4条目的別↑面積00}", "00", Strings.Right("00" & i, 2)), "0")
                            pSheet.ValueReplace(Replace("{4条目的別↓面積00}", "00", Strings.Right("00" & i, 2)), "0")
                            pSheet.ValueReplace(Replace("{4条目的別↓全体件数00}", "00", Strings.Right("00" & i, 2)), "0")
                            pSheet.ValueReplace(Replace("{4条目的別↓全体面積00}", "00", Strings.Right("00" & i, 2)), "0")
                            pSheet.ValueReplace(Replace("{4条目的別↑うち農用地件数00}", "00", Strings.Right("00" & i, 2)), "0")
                            pSheet.ValueReplace(Replace("{4条目的別↓うち農用地件数00}", "00", Strings.Right("00" & i, 2)), "0")
                            pSheet.ValueReplace(Replace("{4条目的別↑うち農用地面積00}", "00", Strings.Right("00" & i, 2)), "0")
                            pSheet.ValueReplace(Replace("{4条目的別↓うち農用地面積00}", "00", Strings.Right("00" & i, 2)), "0")

                            pSheet.ValueReplace(Replace("{4条田面積00}", "00", Strings.Right("00" & i, 2)), "0")
                            pSheet.ValueReplace(Replace("{4条畑面積00}", "00", Strings.Right("00" & i, 2)), "0")
                        Else
                            pSheet.ValueReplace(Replace("{4条目的別件数00}", "00", Strings.Right("00" & i, 2)), pRowI.Item("Count"))
                            pSheet.ValueReplace(Replace("{4条目的別↑件数00}", "00", Strings.Right("00" & i, 2)), pRowI.Item("3000超件数"))
                            pSheet.ValueReplace(Replace("{4条目的別↓件数00}", "00", Strings.Right("00" & i, 2)), pRowI.Item("3000以下件数"))
                            pSheet.ValueReplace(Replace("{4条目的別↑面積00}", "00", Strings.Right("00" & i, 2)), pRowI.Item("3000超面積"))
                            pSheet.ValueReplace(Replace("{4条目的別↓面積00}", "00", Strings.Right("00" & i, 2)), pRowI.Item("3000以下面積"))
                            pSheet.ValueReplace(Replace("{4条目的別↓全体件数00}", "00", Strings.Right("00" & i, 2)), pRowI.Item("3000以下全体件数"))
                            pSheet.ValueReplace(Replace("{4条目的別↓全体面積00}", "00", Strings.Right("00" & i, 2)), pRowI.Item("3000以下全体面積"))
                            pSheet.ValueReplace(Replace("{4条目的別↑うち農用地件数00}", "00", Strings.Right("00" & i, 2)), pRowI.Item("3000超件数うち農用地"))
                            pSheet.ValueReplace(Replace("{4条目的別↓うち農用地件数00}", "00", Strings.Right("00" & i, 2)), pRowI.Item("3000以下件数うち農用地"))
                            pSheet.ValueReplace(Replace("{4条目的別↑うち農用地面積00}", "00", Strings.Right("00" & i, 2)), pRowI.Item("3000超面積うち農用地"))
                            pSheet.ValueReplace(Replace("{4条目的別↓うち農用地面積00}", "00", Strings.Right("00" & i, 2)), pRowI.Item("3000以下面積うち農用地"))

                            pSheet.ValueReplace(Replace("{4条田面積00}", "00", Strings.Right("00" & i, 2)), pRowI.Item("田面積"))
                            pSheet.ValueReplace(Replace("{4条畑面積00}", "00", Strings.Right("00" & i, 2)), pRowI.Item("畑面積"))
                        End If
                    Next
                End If
            End If

            For n = 1 To 2
                pSheet.ValueReplace(Replace("{4条農地区分別件数00}", "00", Strings.Right("00" & n, 2)), 0)
                pSheet.ValueReplace(Replace("{4条農地区分別面積00}", "00", Strings.Right("00" & n, 2)), 0)
                pSheet.ValueReplace(Replace("{4条農地区分別↑件数00}", "00", Strings.Right("00" & n, 2)), 0)
                pSheet.ValueReplace(Replace("{4条農地区分別↑面積00}", "00", Strings.Right("00" & n, 2)), 0)
                pSheet.ValueReplace(Replace("{4条農地区分別↓件数00}", "00", Strings.Right("00" & n, 2)), 0)
                pSheet.ValueReplace(Replace("{4条農地区分別↓面積00}", "00", Strings.Right("00" & n, 2)), 0)
                pSheet.ValueReplace(Replace("{4条農地区分別↑うち農用地件数00}", "00", Strings.Right("00" & n, 2)), 0)
                pSheet.ValueReplace(Replace("{4条農地区分別↑うち農用地面積00}", "00", Strings.Right("00" & n, 2)), 0)
                pSheet.ValueReplace(Replace("{4条農地区分別↓うち農用地件数00}", "00", Strings.Right("00" & n, 2)), 0)
                pSheet.ValueReplace(Replace("{4条農地区分別↓うち農用地面積00}", "00", Strings.Right("00" & n, 2)), 0)
            Next
            For i = 1 To 12
                pSheet.ValueReplace(Replace("{4条目的別件数00}", "00", Strings.Right("00" & i, 2)), "0")
                pSheet.ValueReplace(Replace("{4条目的別↑件数00}", "00", Strings.Right("00" & i, 2)), "0")
                pSheet.ValueReplace(Replace("{4条目的別↓件数00}", "00", Strings.Right("00" & i, 2)), "0")
                pSheet.ValueReplace(Replace("{4条目的別↑面積00}", "00", Strings.Right("00" & i, 2)), "0")
                pSheet.ValueReplace(Replace("{4条目的別↓面積00}", "00", Strings.Right("00" & i, 2)), "0")
                pSheet.ValueReplace(Replace("{4条目的別↓全体件数00}", "00", Strings.Right("00" & i, 2)), "0")
                pSheet.ValueReplace(Replace("{4条目的別↓全体面積00}", "00", Strings.Right("00" & i, 2)), "0")
                pSheet.ValueReplace(Replace("{4条目的別↑うち農用地件数00}", "00", Strings.Right("00" & i, 2)), "0")
                pSheet.ValueReplace(Replace("{4条目的別↓うち農用地件数00}", "00", Strings.Right("00" & i, 2)), "0")
                pSheet.ValueReplace(Replace("{4条目的別↑うち農用地面積00}", "00", Strings.Right("00" & i, 2)), "0")
                pSheet.ValueReplace(Replace("{4条目的別↓うち農用地面積00}", "00", Strings.Right("00" & i, 2)), "0")

                pSheet.ValueReplace(Replace("{4条田面積00}", "00", Strings.Right("00" & i, 2)), "0")
                pSheet.ValueReplace(Replace("{4条畑面積00}", "00", Strings.Right("00" & i, 2)), "0")
            Next

            If pDataCreater.TabCtrl.TabPages.ContainsKey("n.50") Then
                Dim pTB5条 As 諮問意見書Page = pDataCreater.TabCtrl.TabPages("n.50")
                Dim pTable As DataTable = CType(pTB5条.List.DataSource, DataView).ToTable
                Dim pView As New DataView(pTable, "[選択]=True", "", DataViewRowState.CurrentRows)

                Dim pTBL5条諮問農地区分別総括 As New DataTable("5条諮問農地区分別総括")
                SetTBL4条5条(pTBL5条諮問農地区分別総括)

                Dim pTBL5条諮問目的別総括 As New DataTable("5条諮問目的別総括")
                SetTBL4条5条(pTBL5条諮問目的別総括)

                nLoop = -1

                If pView.Count > 0 Then
                    pDataCreater.Maximum = pView.Count + 1
                    pDataCreater.Value = 0
                    Dim nCount As Integer = 1
                    For Each pRow As DataRowView In pView
                        If pRow.Item("選択") And pRow.Item("複数申請人A") = False Then
                            nLoop += 1
                            pDataCreater.Message = "5条諮問総括表 処理中 (" & nLoop + 1 & "/" & pView.Count & ")"
                            pDataCreater.Value = nLoop

                            Dim n As Integer = Val(pRow.Item("農地区分").ToString)
                            Dim i As Integer = Val(pRow.Item("申請後農地分類").ToString)
                            Dim sNList As String = pRow.Item("農地リスト").ToString
                            Dim sNID As String = ""
                            Dim Ar As String() = Split(sNList, ";")

                            For Each sKey As String In Ar
                                If sKey.StartsWith("農地.") Then : sNID &= IIf(sNID.Length > 0, ",", "") & sKey.Substring(3)
                                ElseIf sKey.StartsWith("転用農地.") Then : sNID &= IIf(sNID.Length > 0, ",", "") & sKey.Substring(5)
                                Else : Stop
                                End If
                            Next

                            App農地基本台帳.TBL農地.MergePlus(SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE [ID] In (" & sNID & ")"))
                            Dim pView50 As New DataView(App農地基本台帳.TBL農地.Body, "[ID] In (" & sNID & ")", "", DataViewRowState.CurrentRows)
                            Dim nArea As Decimal = 0
                            Dim n田面積 As Decimal = 0
                            Dim n畑面積 As Decimal = 0
                            Dim n採草放牧面積 As Decimal = 0

                            For Each pRowV As DataRowView In pView50
                                nArea += Val(pRowV.Item("実面積").ToString)

                                If Left(pRowV.Item("登記簿地目名").ToString, 1) = "田" Then
                                    n田面積 += Val(pRowV.Item("実面積").ToString)
                                ElseIf Left(pRowV.Item("登記簿地目名").ToString, 1) = "畑" Then
                                    n畑面積 += Val(pRowV.Item("実面積").ToString)
                                Else
                                    n採草放牧面積 += Val(pRowV.Item("実面積").ToString)
                                End If
                            Next
                            App農地基本台帳.TBL転用農地.MergePlus(SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D_転用農地] WHERE [ID] In (" & sNID & ")"))
                            Dim pViewT50 As New DataView(App農地基本台帳.TBL転用農地.Body, "[ID] In (" & sNID & ")", "", DataViewRowState.CurrentRows)
                            For Each pRowV As DataRowView In pViewT50
                                nArea += Val(pRowV.Item("実面積").ToString)

                                If Left(pRowV.Item("登記簿地目名").ToString, 1) = "田" Then
                                    n田面積 += Val(pRowV.Item("実面積").ToString)
                                ElseIf Left(pRowV.Item("登記簿地目名").ToString, 1) = "畑" Then
                                    n畑面積 += Val(pRowV.Item("実面積").ToString)
                                Else
                                    n採草放牧面積 += Val(pRowV.Item("実面積").ToString)
                                End If
                            Next

                            Select Case n
                                Case 1, 4, 5 : n = 1
                                Case 2, 3, 6 : n = 2
                            End Select

                            Dim pRowN As DataRow = pTBL5条諮問農地区分別総括.Rows.Find(n)
                            If pRowN Is Nothing Then
                                pRowN = pTBL5条諮問農地区分別総括.NewRow
                                pRowN.Item("ID") = n
                                pRowN.Item("Count") = 1
                                pRowN.Item("面積") = nArea

                                If nArea >= 3000 Then
                                    pRowN.Item("3000超件数") = 1
                                    pRowN.Item("3000超面積") = nArea
                                    If Val(pRow.Item("農地区分").ToString) = 1 Or Val(pRow.Item("農地区分").ToString) = 5 Then
                                        pRowN.Item("3000超件数うち農用地") = 1
                                        pRowN.Item("3000超面積うち農用地") = nArea
                                    End If
                                Else
                                    pRowN.Item("3000以下全体件数") = 1
                                    pRowN.Item("3000以下全体面積") = nArea

                                    If Val(pRow.Item("意見聴取案件").ToString) > 0 Or Val(pRow.Item("農地区分").ToString) = 1 Or Val(pRow.Item("農地区分").ToString) = 5 Then
                                        pRowN.Item("3000以下件数") = 1
                                        pRowN.Item("3000以下面積") = nArea
                                        If Val(pRow.Item("農地区分").ToString) = 1 Or Val(pRow.Item("農地区分").ToString) = 5 Then
                                            pRowN.Item("3000以下件数うち農用地") = 1
                                            pRowN.Item("3000以下面積うち農用地") = nArea
                                        End If
                                    End If
                                End If

                                pTBL5条諮問農地区分別総括.Rows.Add(pRowN)
                            Else
                                pRowN.Item("Count") += 1
                                pRowN.Item("面積") += nArea

                                If nArea >= 3000 Then
                                    pRowN.Item("3000超件数") += 1
                                    pRowN.Item("3000超面積") += nArea
                                    If Val(pRow.Item("農地区分").ToString) = 1 Or Val(pRow.Item("農地区分").ToString) = 5 Then
                                        pRowN.Item("3000超件数うち農用地") += 1
                                        pRowN.Item("3000超面積うち農用地") += nArea
                                    End If
                                Else
                                    pRowN.Item("3000以下全体件数") += 1
                                    pRowN.Item("3000以下全体面積") += nArea

                                    If Val(pRow.Item("意見聴取案件").ToString) > 0 Or Val(pRow.Item("農地区分").ToString) = 1 Or Val(pRow.Item("農地区分").ToString) = 5 Then
                                        pRowN.Item("3000以下件数") += 1
                                        pRowN.Item("3000以下面積") += nArea
                                        If Val(pRow.Item("農地区分").ToString) = 1 Or Val(pRow.Item("農地区分").ToString) = 5 Then
                                            pRowN.Item("3000以下件数うち農用地") += 1
                                            pRowN.Item("3000以下面積うち農用地") += nArea
                                        End If
                                    End If
                                End If
                            End If

                            Select Case i
                                Case 11 : i = 10
                                Case Else
                            End Select

                            Dim pRowI As DataRow = pTBL5条諮問目的別総括.Rows.Find(i)
                            If pRowI Is Nothing Then
                                pRowI = pTBL5条諮問目的別総括.NewRow
                                pRowI.Item("ID") = i
                                pRowI.Item("Count") = 1
                                pRowI.Item("田面積") = n田面積
                                pRowI.Item("畑面積") = n畑面積
                                pRowI.Item("採草放牧面積") = n採草放牧面積

                                If nArea >= 3000 Then
                                    pRowI.Item("3000超件数") = 1
                                    pRowI.Item("3000超面積") = nArea
                                    If Val(pRow.Item("農地区分").ToString) = 1 Or Val(pRow.Item("農地区分").ToString) = 5 Then
                                        pRowI.Item("3000超件数うち農用地") = 1
                                        pRowI.Item("3000超面積うち農用地") = nArea
                                    End If
                                Else
                                    pRowI.Item("3000以下全体件数") = 1
                                    pRowI.Item("3000以下全体面積") = nArea

                                    If Val(pRow.Item("意見聴取案件").ToString) > 0 Or Val(pRow.Item("農地区分").ToString) = 1 Or Val(pRow.Item("農地区分").ToString) = 5 Then
                                        pRowI.Item("3000以下件数") = 1
                                        pRowI.Item("3000以下面積") = nArea
                                        If Val(pRow.Item("農地区分").ToString) = 1 Or Val(pRow.Item("農地区分").ToString) = 5 Then
                                            pRowI.Item("3000以下件数うち農用地") = 1
                                            pRowI.Item("3000以下面積うち農用地") = nArea
                                        End If
                                    End If
                                End If

                                pTBL5条諮問目的別総括.Rows.Add(pRowI)
                            Else
                                pRowI.Item("Count") += 1
                                pRowI.Item("田面積") += n田面積
                                pRowI.Item("畑面積") += n畑面積
                                pRowI.Item("採草放牧面積") += n採草放牧面積

                                If nArea >= 3000 Then
                                    pRowI.Item("3000超件数") += 1
                                    pRowI.Item("3000超面積") += nArea
                                    If Val(pRow.Item("農地区分").ToString) = 1 Or Val(pRow.Item("農地区分").ToString) = 5 Then
                                        pRowI.Item("3000超件数うち農用地") += 1
                                        pRowI.Item("3000超面積うち農用地") += nArea
                                    End If
                                Else
                                    pRowI.Item("3000以下全体件数") += 1
                                    pRowI.Item("3000以下全体面積") += nArea

                                    If Val(pRow.Item("意見聴取案件").ToString) > 0 Or Val(pRow.Item("農地区分").ToString) = 1 Or Val(pRow.Item("農地区分").ToString) = 5 Then
                                        pRowI.Item("3000以下件数") += 1
                                        pRowI.Item("3000以下面積") += nArea
                                        If Val(pRow.Item("農地区分").ToString) = 1 Or Val(pRow.Item("農地区分").ToString) = 5 Then
                                            pRowI.Item("3000以下件数うち農用地") += 1
                                            pRowI.Item("3000以下面積うち農用地") += nArea
                                        End If
                                    End If
                                End If
                            End If

                            If Val(pRow.Item("申請後農地分類").ToString) = 11 Then
                                pSheet.ValueReplace(Replace("{許可年月日00}", "00", Strings.Right("00" & nCount, 2)), pRow.Item("許可年月日").ToString)
                                pSheet.ValueReplace(Replace("{市町村名00}", "00", Strings.Right("00" & nCount, 2)), SysAD.市町村.市町村名)
                                pSheet.ValueReplace(Replace("{事業者名00}", "00", Strings.Right("00" & nCount, 2)), pRow.Item("氏名B").ToString)
                                pSheet.ValueReplace(Replace("{条項00}", "00", Strings.Right("00" & nCount, 2)), "５条")
                                pSheet.ValueReplace(Replace("{農地区分00}", "00", Strings.Right("00" & nCount, 2)), IIf(Val(pRow.Item("農地区分").ToString) = 1, "第１種農地", IIf(Val(pRow.Item("農地区分").ToString) = 2, "第２種農地", IIf(Val(pRow.Item("農地区分").ToString) = 3, "第３種農地", IIf(Val(pRow.Item("農地区分").ToString) = 4, "甲種農地", IIf(Val(pRow.Item("農地区分").ToString) = 5, "農用地区域内農地", IIf(Val(pRow.Item("農地区分").ToString) = 6, "第２種農地その他の農地", "")))))))
                                pSheet.ValueReplace(Replace("{種類00}", "00", Strings.Right("00" & nCount, 2)), "太陽光発電")
                                pSheet.ValueReplace(Replace("{転用面積00}", "00", Strings.Right("00" & nCount, 2)), nArea)
                                pSheet.ValueReplace(Replace("{備考00}", "00", Strings.Right("00" & nCount, 2)), pRow.Item("不許可例外").ToString)

                                nCount += 1

                            End If
                        End If
                    Next

                    For n = 1 To 2
                        Dim pRowN As DataRow = pTBL5条諮問農地区分別総括.Rows.Find(n)
                        If pRowN Is Nothing Then
                            pSheet.ValueReplace(Replace("{5条農地区分別件数00}", "00", Strings.Right("00" & n, 2)), "0")
                            pSheet.ValueReplace(Replace("{5条農地区分別面積00}", "00", Strings.Right("00" & n, 2)), "0")
                            pSheet.ValueReplace(Replace("{5条農地区分別↑件数00}", "00", Strings.Right("00" & n, 2)), "0")
                            pSheet.ValueReplace(Replace("{5条農地区分別↑面積00}", "00", Strings.Right("00" & n, 2)), "0")
                            pSheet.ValueReplace(Replace("{5条農地区分別↓件数00}", "00", Strings.Right("00" & n, 2)), "0")
                            pSheet.ValueReplace(Replace("{5条農地区分別↓面積00}", "00", Strings.Right("00" & n, 2)), "0")
                            pSheet.ValueReplace(Replace("{5条農地区分別↑うち農用地件数00}", "00", Strings.Right("00" & n, 2)), "0")
                            pSheet.ValueReplace(Replace("{5条農地区分別↑うち農用地面積00}", "00", Strings.Right("00" & n, 2)), "0")
                            pSheet.ValueReplace(Replace("{5条農地区分別↓うち農用地件数00}", "00", Strings.Right("00" & n, 2)), "0")
                            pSheet.ValueReplace(Replace("{5条農地区分別↓うち農用地面積00}", "00", Strings.Right("00" & n, 2)), "0")
                        Else
                            pSheet.ValueReplace(Replace("{5条農地区分別件数00}", "00", Strings.Right("00" & n, 2)), pRowN.Item("Count"))
                            pSheet.ValueReplace(Replace("{5条農地区分別面積00}", "00", Strings.Right("00" & n, 2)), pRowN.Item("面積"))
                            pSheet.ValueReplace(Replace("{5条農地区分別↑件数00}", "00", Strings.Right("00" & n, 2)), pRowN.Item("3000超件数"))
                            pSheet.ValueReplace(Replace("{5条農地区分別↑面積00}", "00", Strings.Right("00" & n, 2)), pRowN.Item("3000超面積"))
                            pSheet.ValueReplace(Replace("{5条農地区分別↓件数00}", "00", Strings.Right("00" & n, 2)), pRowN.Item("3000以下件数"))
                            pSheet.ValueReplace(Replace("{5条農地区分別↓面積00}", "00", Strings.Right("00" & n, 2)), pRowN.Item("3000以下面積"))
                            pSheet.ValueReplace(Replace("{5条農地区分別↑うち農用地件数00}", "00", Strings.Right("00" & n, 2)), pRowN.Item("3000超件数うち農用地"))
                            pSheet.ValueReplace(Replace("{5条農地区分別↑うち農用地面積00}", "00", Strings.Right("00" & n, 2)), pRowN.Item("3000超面積うち農用地"))
                            pSheet.ValueReplace(Replace("{5条農地区分別↓うち農用地件数00}", "00", Strings.Right("00" & n, 2)), pRowN.Item("3000以下件数うち農用地"))
                            pSheet.ValueReplace(Replace("{5条農地区分別↓うち農用地面積00}", "00", Strings.Right("00" & n, 2)), pRowN.Item("3000以下面積うち農用地"))
                        End If
                    Next

                    For i = 1 To 12
                        Dim pRowI As DataRow = pTBL5条諮問目的別総括.Rows.Find(i)
                        If pRowI Is Nothing Then
                            pSheet.ValueReplace(Replace("{5条目的別件数00}", "00", Strings.Right("00" & i, 2)), "0")
                            pSheet.ValueReplace(Replace("{5条目的別↑件数00}", "00", Strings.Right("00" & i, 2)), "0")
                            pSheet.ValueReplace(Replace("{5条目的別↓件数00}", "00", Strings.Right("00" & i, 2)), "0")
                            pSheet.ValueReplace(Replace("{5条目的別↑面積00}", "00", Strings.Right("00" & i, 2)), "0")
                            pSheet.ValueReplace(Replace("{5条目的別↓面積00}", "00", Strings.Right("00" & i, 2)), "0")
                            pSheet.ValueReplace(Replace("{5条目的別↓全体件数00}", "00", Strings.Right("00" & i, 2)), "0")
                            pSheet.ValueReplace(Replace("{5条目的別↓全体面積00}", "00", Strings.Right("00" & i, 2)), "0")
                            pSheet.ValueReplace(Replace("{5条目的別↑うち農用地件数00}", "00", Strings.Right("00" & i, 2)), "0")
                            pSheet.ValueReplace(Replace("{5条目的別↓うち農用地件数00}", "00", Strings.Right("00" & i, 2)), "0")
                            pSheet.ValueReplace(Replace("{5条目的別↑うち農用地面積00}", "00", Strings.Right("00" & i, 2)), "0")
                            pSheet.ValueReplace(Replace("{5条目的別↓うち農用地面積00}", "00", Strings.Right("00" & i, 2)), "0")

                            pSheet.ValueReplace(Replace("{5条田面積00}", "00", Strings.Right("00" & i, 2)), "0")
                            pSheet.ValueReplace(Replace("{5条畑面積00}", "00", Strings.Right("00" & i, 2)), "0")
                            pSheet.ValueReplace(Replace("{5条採草放牧面積00}", "00", Strings.Right("00" & i, 2)), "0")
                        Else
                            pSheet.ValueReplace(Replace("{5条目的別件数00}", "00", Strings.Right("00" & i, 2)), pRowI.Item("Count"))
                            pSheet.ValueReplace(Replace("{5条目的別↑件数00}", "00", Strings.Right("00" & i, 2)), pRowI.Item("3000超件数"))
                            pSheet.ValueReplace(Replace("{5条目的別↓件数00}", "00", Strings.Right("00" & i, 2)), pRowI.Item("3000以下件数"))
                            pSheet.ValueReplace(Replace("{5条目的別↑面積00}", "00", Strings.Right("00" & i, 2)), pRowI.Item("3000超面積"))
                            pSheet.ValueReplace(Replace("{5条目的別↓面積00}", "00", Strings.Right("00" & i, 2)), pRowI.Item("3000以下面積"))
                            pSheet.ValueReplace(Replace("{5条目的別↓全体件数00}", "00", Strings.Right("00" & i, 2)), pRowI.Item("3000以下全体件数"))
                            pSheet.ValueReplace(Replace("{5条目的別↓全体面積00}", "00", Strings.Right("00" & i, 2)), pRowI.Item("3000以下全体面積"))
                            pSheet.ValueReplace(Replace("{5条目的別↑うち農用地件数00}", "00", Strings.Right("00" & i, 2)), pRowI.Item("3000超件数うち農用地"))
                            pSheet.ValueReplace(Replace("{5条目的別↓うち農用地件数00}", "00", Strings.Right("00" & i, 2)), pRowI.Item("3000以下件数うち農用地"))
                            pSheet.ValueReplace(Replace("{5条目的別↑うち農用地面積00}", "00", Strings.Right("00" & i, 2)), pRowI.Item("3000超面積うち農用地"))
                            pSheet.ValueReplace(Replace("{5条目的別↓うち農用地面積00}", "00", Strings.Right("00" & i, 2)), pRowI.Item("3000以下面積うち農用地"))

                            pSheet.ValueReplace(Replace("{5条田面積00}", "00", Strings.Right("00" & i, 2)), pRowI.Item("田面積"))
                            pSheet.ValueReplace(Replace("{5条畑面積00}", "00", Strings.Right("00" & i, 2)), pRowI.Item("畑面積"))
                            pSheet.ValueReplace(Replace("{5条採草放牧面積00}", "00", Strings.Right("00" & i, 2)), pRowI.Item("採草放牧面積"))
                        End If
                    Next
                End If
            End If

            For n = 1 To 2
                pSheet.ValueReplace(Replace("{5条農地区分別件数00}", "00", Strings.Right("00" & n, 2)), "0")
                pSheet.ValueReplace(Replace("{5条農地区分別面積00}", "00", Strings.Right("00" & n, 2)), "0")
                pSheet.ValueReplace(Replace("{5条農地区分別↑件数00}", "00", Strings.Right("00" & n, 2)), "0")
                pSheet.ValueReplace(Replace("{5条農地区分別↑面積00}", "00", Strings.Right("00" & n, 2)), "0")
                pSheet.ValueReplace(Replace("{5条農地区分別↓件数00}", "00", Strings.Right("00" & n, 2)), "0")
                pSheet.ValueReplace(Replace("{5条農地区分別↓面積00}", "00", Strings.Right("00" & n, 2)), "0")
                pSheet.ValueReplace(Replace("{5条農地区分別↑うち農用地件数00}", "00", Strings.Right("00" & n, 2)), "0")
                pSheet.ValueReplace(Replace("{5条農地区分別↑うち農用地面積00}", "00", Strings.Right("00" & n, 2)), "0")
                pSheet.ValueReplace(Replace("{5条農地区分別↓うち農用地件数00}", "00", Strings.Right("00" & n, 2)), "0")
                pSheet.ValueReplace(Replace("{5条農地区分別↓うち農用地面積00}", "00", Strings.Right("00" & n, 2)), "0")
            Next

            For i = 1 To 12
                pSheet.ValueReplace(Replace("{5条目的別件数00}", "00", Strings.Right("00" & i, 2)), "0")
                pSheet.ValueReplace(Replace("{5条目的別↑件数00}", "00", Strings.Right("00" & i, 2)), "0")
                pSheet.ValueReplace(Replace("{5条目的別↓件数00}", "00", Strings.Right("00" & i, 2)), "0")
                pSheet.ValueReplace(Replace("{5条目的別↑面積00}", "00", Strings.Right("00" & i, 2)), "0")
                pSheet.ValueReplace(Replace("{5条目的別↓面積00}", "00", Strings.Right("00" & i, 2)), "0")
                pSheet.ValueReplace(Replace("{5条目的別↓全体件数00}", "00", Strings.Right("00" & i, 2)), "0")
                pSheet.ValueReplace(Replace("{5条目的別↓全体面積00}", "00", Strings.Right("00" & i, 2)), "0")
                pSheet.ValueReplace(Replace("{5条目的別↑うち農用地件数00}", "00", Strings.Right("00" & i, 2)), "0")
                pSheet.ValueReplace(Replace("{5条目的別↓うち農用地件数00}", "00", Strings.Right("00" & i, 2)), "0")
                pSheet.ValueReplace(Replace("{5条目的別↑うち農用地面積00}", "00", Strings.Right("00" & i, 2)), "0")
                pSheet.ValueReplace(Replace("{5条目的別↓うち農用地面積00}", "00", Strings.Right("00" & i, 2)), "0")

                pSheet.ValueReplace(Replace("{5条田面積00}", "00", Strings.Right("00" & i, 2)), "0")
                pSheet.ValueReplace(Replace("{5条畑面積00}", "00", Strings.Right("00" & i, 2)), "0")
                pSheet.ValueReplace(Replace("{5条採草放牧面積00}", "00", Strings.Right("00" & i, 2)), "0")
            Next

            For r = 1 To 30
                pSheet.ValueReplace(Replace("{許可年月日00}", "00", Strings.Right("00" & r, 2)), "")
                pSheet.ValueReplace(Replace("{市町村名00}", "00", Strings.Right("00" & r, 2)), "")
                pSheet.ValueReplace(Replace("{事業者名00}", "00", Strings.Right("00" & r, 2)), "")
                pSheet.ValueReplace(Replace("{条項00}", "00", Strings.Right("00" & r, 2)), "")
                pSheet.ValueReplace(Replace("{農地区分00}", "00", Strings.Right("00" & r, 2)), "")
                pSheet.ValueReplace(Replace("{種類00}", "00", Strings.Right("00" & r, 2)), "")
                pSheet.ValueReplace(Replace("{転用面積00}", "00", Strings.Right("00" & r, 2)), "")
                pSheet.ValueReplace(Replace("{備考00}", "00", Strings.Right("00" & r, 2)), "")
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class

Public Class C諮問意見書Data作成
    Inherits HimTools2012.clsAccessor
    Private mvarTab As TabControl
    Public Sub New(ByVal pTabC As TabControl)
        MyBase.new()
        mvarTab = pTabC
    End Sub
    Public ReadOnly Property TabCtrl() As TabControl
        Get
            Return mvarTab
        End Get
    End Property

    Public Overrides Sub Execute()
        Dim sFolder As String = SysAD.OutputFolder & String.Format("\総会資料{0}_{1}", Now.Year, Now.Month)
        If IO.Directory.Exists(SysAD.GetXMLProperty("総会資料関連", "出力先フォルダ", "")) Then
            sFolder = SysAD.GetXMLProperty("総会資料関連", "出力先フォルダ", "") & String.Format("\総会資料{0}_{1}", Now.Year, Now.Month)
        End If

        If Not IO.Directory.Exists(sFolder) Then
            IO.Directory.CreateDirectory(sFolder)
        End If

        Dim p諮問資料 As Boolean = False

        For Each p申請 As 諮問意見書Page In mvarTab.TabPages
            If p申請.印刷 Then

                Select Case p申請.Name
                    Case "n.40"
                        sub諮問資料作成("諮問4条一覧表", sFolder, New CPrint諮問資料4条, p申請, "農地法第4条審査表.xml")
                        sub諮問資料作成("諮問4条一覧表", sFolder, New CPrint諮問資料4条, p申請, "農地法第4条意見書.xml")
                        p諮問資料 = p諮問資料 Or True
                    Case "n.50"
                        sub諮問資料作成("諮問5条一覧表", sFolder, New CPrint諮問資料5条, p申請, "農地法第5条審査表.xml")
                        sub諮問資料作成("諮問5条一覧表", sFolder, New CPrint諮問資料5条, p申請, "農地法第5条意見書.xml")
                        p諮問資料 = p諮問資料 Or True
                    Case "n.600", "n.602"
                        'sub総会資料作成("非農地証明", sFolder, New CPrint総会資料作成非農地証明, p申請, "非農地証明書願いについて.xml")
                    Case Else
                        'Stop
                End Select
            End If
        Next

        If p諮問資料 Then
            sub諮問資料作成("農地法目的別諮問総括表", sFolder, New CPrint諮問総括表, Nothing, "00_農地法目的別諮問総括表.xml")
            sub諮問資料作成("農地区分別諮問一覧表", sFolder, New CPrint諮問農地区分別一覧表, Nothing, "01_農地区分別一覧表.xml")
            sub諮問資料作成("諮問資料個別説明", sFolder, New CPrint諮問資料個別説明(Me), Nothing, "04_諮問説明資料.xml")
            sub諮問資料作成("諮問資料個別説明", sFolder, New CPrint諮問説明資料(Me), Nothing, "05_諮問説明資料.xml")
            sub諮問資料作成("諮問資料個別説明", sFolder, New CPrint諮問農地法農地区分別総括表, Nothing, "06_農地区分別諮問総括表.xml")
            sub諮問資料作成("諮問資料個別説明", sFolder, New CPrint諮問農地法農地区分別総括表, Nothing, "農転用許可状況報告書.xml")

        End If

        System.Diagnostics.Process.Start(sFolder)
    End Sub


    Private Function sub諮問資料作成(ByVal s処理名称 As String, ByVal sDesktopFolder As String, ByVal p作成 As CPrint諮問意見書作成共通, ByRef p申請 As 諮問意見書Page, ByVal sFile As String) As Boolean
        If System.IO.File.Exists(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\" & sFile) Then
            If p申請 Is Nothing Then
                If p作成.LoopSub(s処理名称, sDesktopFolder, p作成, sFile) Then
                    Return True
                Else
                    Dim sNewFile As String = HimTools2012.FileManager.NewFileName(sDesktopFolder & "\" & sFile, 0)
                    Dim sXML As String = HimTools2012.TextAdapter.LoadTextFile(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\" & sFile)

                    Dim pXMLSS As New CXMLSS2003(sXML)
                    p作成.SetData(pXMLSS, p申請, s処理名称, Me)

                    HimTools2012.TextAdapter.SaveTextFile(sNewFile, pXMLSS.OutPut(True))
                    Return True
                End If
            ElseIf IO.File.Exists(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\" & sFile) Then
                If p作成.LoopSub(s処理名称, sDesktopFolder, p作成, sFile) Then
                    Return True
                Else

                    If p申請.SelectCount > 0 Then
                        Dim sNewFile As String = HimTools2012.FileManager.NewFileName(sDesktopFolder & "\" & sFile, 0)

                        Dim sXML As String = HimTools2012.TextAdapter.LoadTextFile(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\" & sFile)
                        Dim pXMLSS As New CXMLSS2003(sXML)

                        p作成.SetData(pXMLSS, p申請, s処理名称, Me)

                        HimTools2012.TextAdapter.SaveTextFile(sNewFile, pXMLSS.OutPut(True))


                        Return True
                    Else
                        Return False
                    End If

                End If
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

End Class

Public MustInherit Class CPrint諮問意見書作成共通
    Protected nLoop As Integer = -1
    Public LoopRows As XMLLoopRows

    MustOverride Sub SetDataRow(ByRef pSheet As XMLSSWorkSheet, ByRef pRow As DataRowView, Optional ByVal pObj As Object = Nothing)
    MustOverride Sub SetData(ByRef XMLSS As CXMLSS2003, ByRef pTab As 諮問意見書Page, ByVal s処理名称 As String, ByRef pDataCreater As C諮問意見書Data作成)

    Public Overridable Function LoopSub(ByVal s処理名称 As String, ByVal sDesktopFolder As String, ByVal p作成 As CPrint諮問意見書作成共通, ByVal sFile As String, ByVal ParamArray pTabs() As Object) As Boolean
        Return False
    End Function

    Public Function SetNO(ByRef pSheet As XMLSSWorkSheet) As Integer

        pSheet.ValueReplace("{No}", (nLoop + 1).ToString)

        Return nLoop + 1
    End Function
End Class

Public MustInherit Class CPrint諮問意見書資料作成
    Inherits CPrint諮問意見書作成共通

    Protected mvar総合計 As New C筆明細と集計作成

    Protected 総括Data As New Dictionary(Of String, dt総括表)

    Public Sub New()

    End Sub

    Public Overloads Function SetNO(ByRef pSheet As XMLSSWorkSheet, ByVal b行高調整 As Boolean) As Integer
        If b行高調整 Then
            pSheet.ValueReplace("{No}", "" & (nLoop + 1).ToString)
        Else
            pSheet.ValueReplace("{No}", (nLoop + 1).ToString)
        End If
        Return nLoop + 1
    End Function

    Public Sub Set複数行(ByRef pSheet As XMLSSWorkSheet, ByRef pTab As 諮問意見書Page, ByVal s処理名称 As String, ByRef pDataCreater As C諮問意見書Data作成, ByVal sWhere As String)
        If Not InStr(pSheet.Table.InnerXML, "{No}") > 0 Then
            Exit Sub
        End If
        LoopRows = New XMLLoopRows(pSheet)

        Dim pTable As DataTable = CType(pTab.List.DataSource, DataView).Table

        Dim pView As New DataView(pTable, CType(pTab.List.DataSource, DataView).RowFilter & " AND [選択]=True", "中間管理Flag,受付補助記号,受付番号", DataViewRowState.CurrentRows)

        pDataCreater.Maximum = pView.Count
        nLoop = -1

        For Each pRow As DataRowView In pView
            If pRow.Item("選択") Then

                If nLoop = -1 Then

                Else
                    For Each pXRow As XMLSSRow In LoopRows
                        Dim pCopyRow = pXRow.CopyRow

                        pSheet.Table.Rows.InsertRow(LoopRows.InsetRow, pCopyRow)
                        LoopRows.InsetRow += 1
                    Next
                End If

                nLoop += 1
                pDataCreater.Message = s処理名称 & " 処理中 (" & nLoop + 1 & "/" & pView.Count & ")"
                pDataCreater.Value = nLoop

                Me.SetDataRow(pSheet, pRow)
            End If
        Next
    End Sub

    Protected Function 複数土地設定(ByRef pSheet As XMLSSWorkSheet, ByRef pRow As DataRowView, ByRef p総括Item As dt総括表, Optional ByVal b再設定 As Boolean = False) As C筆明細と集計作成
        Dim sNList As String = pRow.Item("農地リスト").ToString
        Dim s農振区分 As String = ""

        Dim p案件内集計 As New C筆明細と集計作成

        Dim nCount As Integer = 0
        Dim sNID As String = ""

        Dim Ar As String() = Split(sNList, ";")
        For Each sKey As String In Ar

            If sKey.StartsWith("農地.") Then
                sNID &= IIf(sNID.Length > 0, ",", "") & sKey.Substring(3)
            ElseIf sKey.StartsWith("転用農地.") Then
                sNID &= IIf(sNID.Length > 0, ",", "") & sKey.Substring(5)
            Else
                Stop
            End If
        Next

        App農地基本台帳.TBL農地.FindRowBySQL("[ID] In (" & sNID & ")")
        Dim pView As New DataView(App農地基本台帳.TBL農地.Body, "[ID] In (" & sNID & ")", "", DataViewRowState.CurrentRows)

        For Each pRowV As DataRowView In pView
            If InStr("," & sNID & ",", "," & pRowV.Item("ID") & ",") Then
                sNID = Replace("," & sNID & ",", "," & pRowV.Item("ID") & ",", ",")
            End If
            nCount += 1

            s農振区分 = s農振区分 & p案件内集計.R & IIf(Val(pRowV.Item("農振法区分").ToString) = 0, _
                                                    IIf(Val(pRowV.Item("農業振興地域").ToString) = 0, "他", IIf(Val(pRowV.Item("農業振興地域").ToString) = 1, "内", IIf(Val(pRowV.Item("農業振興地域").ToString) = 2, "外", "-"))), _
                                                    IIf(Val(pRowV.Item("農振法区分").ToString) = 1, "内", IIf(Val(pRowV.Item("農振法区分").ToString) = 2, "他", IIf(Val(pRowV.Item("農振法区分").ToString) = 3, "外", "-"))))
            p案件内集計.Set筆情報(pRowV.Row, p総括Item, b再設定)
        Next

        sNID = Replace(sNID, ",,", ",")
        Do Until Not sNID.StartsWith(",") AndAlso Not sNID.EndsWith(",")
            If sNID.StartsWith(",") Then sNID = Strings.Mid(sNID, 2)
            If sNID.EndsWith(",") Then sNID = Strings.Left(sNID, Len(sNID) - 1)
        Loop

        If Len(sNID) Then
            App農地基本台帳.TBL転用農地.FindRowBySQL("[ID] In (" & sNID & ")")
            Dim pViewT As New DataView(App農地基本台帳.TBL転用農地.Body, "[ID] In (" & sNID & ")", "", DataViewRowState.CurrentRows)
            For Each pRowV As DataRowView In pViewT
                If InStr("," & sNID & ",", "," & pRowV.Item("ID") & ",") Then
                    sNID = Replace("," & sNID & ",", "," & pRowV.Item("ID") & ",", ",")
                End If
                nCount += 1

                s農振区分 = s農振区分 & p案件内集計.R & IIf(Val(pRowV.Item("農振法区分").ToString) > 0, _
                                                    IIf(Val(pRowV.Item("農業振興地域").ToString) = 0, "他", IIf(Val(pRowV.Item("農業振興地域").ToString) = 1, "内", IIf(Val(pRowV.Item("農業振興地域").ToString) = 2, "外", "-"))), _
                                                    IIf(Val(pRowV.Item("農振法区分").ToString) = 1, "内", IIf(Val(pRowV.Item("農振法区分").ToString) = 2, "他", IIf(Val(pRowV.Item("農振法区分").ToString) = 3, "外", "-"))))
                p案件内集計.Set筆情報(pRowV.Row, p総括Item, b再設定)
            Next
        End If

        With p案件内集計
            pSheet.ValueReplace("{残筆数}", nCount - 1)
            pSheet.ValueReplace("{調査票土地の所在}", (Split(.明細作成.To土地所在文字列("&#10;"), "&#10;")(0)))
            pSheet.ValueReplace("{農振区分}", s農振区分)

            p案件内集計.Replace案件毎集計(pSheet)
            mvar総合計.Add合計(p案件内集計)

            pSheet.ValueReplace("{面積}", .明細作成.To面積文字列("&#10;"))
            pSheet.ValueReplace("{面積内}", .明細作成.To面積内文字列("&#10;"))
        End With

        Return p案件内集計
    End Function

    Protected Class C筆明細と集計作成
        Public 田数計 As Integer = 0
        Public 畑数計 As Integer = 0
        Public 樹数計 As Integer = 0
        Public 他数計 As Integer = 0

        Public 田面計 As Decimal = 0
        Public 畑面計 As Decimal = 0
        Public 樹面計 As Decimal = 0
        Public 他面計 As Decimal = 0

        Public Is田内 As Boolean = False
        Public Is畑内 As Boolean = False
        Public Is樹内 As Boolean = False
        Public Is他内 As Boolean = False

        Public 田面計内 As Decimal = 0
        Public 畑面計内 As Decimal = 0
        Public 樹面計内 As Decimal = 0
        Public 他面計内 As Decimal = 0



        Public 明細作成 As New C明細作成
        Public R As String = ""

        Public Sub New()

        End Sub

        Public ReadOnly Property 総面積() As Decimal
            Get
                Return 田面計 + 畑面計 + 樹面計 + 他面計
            End Get
        End Property

        Public Sub Set筆情報(ByVal pRow As DataRow, ByVal p総括Item As dt総括表, ByVal b再設定 As Boolean)
            Dim n田面積 As Decimal = 0
            Dim n田面積内 As Decimal = 0
            Dim n畑面積 As Decimal = 0
            Dim n畑面積内 As Decimal = 0
            Dim n樹面積 As Decimal = 0
            Dim n樹面積内 As Decimal = 0
            Dim n他面積 As Decimal = 0
            Dim n他面積内 As Decimal = 0

            If pRow.Item("登記簿地目名").ToString.IndexOf("田") > -1 AndAlso pRow.Item("登記簿地目名").ToString <> "塩田" Then
                田数計 += 1
                n田面積 += Val(pRow.Item("登記簿面積").ToString)
                n田面積内 = IIf(Val(pRow.Item("部分面積").ToString) > 0, Val(pRow.Item("部分面積").ToString), Val(pRow.Item("登記簿面積").ToString))
                Is田内 = Is田内 AndAlso (Val(pRow.Item("部分面積").ToString) > 0)

                田面計 += n田面積
                田面計内 += n田面積内
            ElseIf pRow.Item("登記簿地目名").ToString.IndexOf("畑") > -1 Then
                畑数計 += 1
                n畑面積 += Val(pRow.Item("登記簿面積").ToString)
                n畑面積内 = IIf(Val(pRow.Item("部分面積").ToString) > 0, Val(pRow.Item("部分面積").ToString), Val(pRow.Item("登記簿面積").ToString))
                Is畑内 = Is畑内 AndAlso (Val(pRow.Item("部分面積").ToString) > 0)

                畑面計 += n畑面積
                畑面計内 += n畑面積内
            ElseIf pRow.Item("登記簿地目名").ToString.IndexOf("樹") > -1 Then
                樹数計 += 1
                n樹面積 += Val(pRow.Item("登記簿面積").ToString)
                n樹面積内 = IIf(Val(pRow.Item("部分面積").ToString) > 0, Val(pRow.Item("部分面積").ToString), Val(pRow.Item("登記簿面積").ToString))
                Is樹内 = Is樹内 AndAlso (Val(pRow.Item("部分面積").ToString) > 0)

                樹面計 += n樹面積
                樹面計内 += n樹面積内
            Else
                他数計 += 1
                n他面積 += Val(pRow.Item("登記簿面積").ToString)
                n他面積内 = IIf(Val(pRow.Item("部分面積").ToString) > 0, Val(pRow.Item("部分面積").ToString), Val(pRow.Item("登記簿面積").ToString))
                Is他内 = Is他内 AndAlso (Val(pRow.Item("部分面積").ToString) > 0)

                他面計 += n他面積
                他面計内 += n他面積内
            End If

            明細作成.Plus(pRow, Val(pRow.Item("登記簿面積").ToString), Val(pRow.Item("部分面積").ToString), n田面積, n畑面積)
            R = "&#10;"

            If p総括Item IsNot Nothing Then
                If pRow.Item("登記簿地目名").ToString.IndexOf("田") > -1 Then
                    p総括Item.Set面積(dt総括表.enum地目.田, n田面積内, b再設定)
                ElseIf pRow.Item("登記簿地目名").ToString.IndexOf("畑") > -1 Then
                    p総括Item.Set面積(dt総括表.enum地目.畑, n畑面積内, b再設定)
                ElseIf pRow.Item("登記簿地目名").ToString.IndexOf("樹") > -1 Then
                    p総括Item.Set面積(dt総括表.enum地目.樹, n樹面積内, b再設定)
                Else
                    p総括Item.Set面積(dt総括表.enum地目.他, n他面積内, b再設定)
                End If
            End If

        End Sub

        Public Sub Replace案件毎集計(ByRef pSheet As XMLSSWorkSheet)
            pSheet.ValueReplace("{田筆数計}", 田数計)
            pSheet.ValueReplace("{田面積計}", IIf(田面計内 > 0, Replace(田面計内.ToString("#,##0.00"), ".00", ""), ""))
            pSheet.ValueReplace("{田面積計内}", IIf(田面計 > 0, 田面計.ToString("#,##0"), "") & IIf(Is田内, "(内" & 田面計内.ToString("#,##0") & ")", ""))

            pSheet.ValueReplace("{畑筆数計}", 畑数計)
            pSheet.ValueReplace("{畑面積計}", IIf(畑面計内 > 0, Replace(畑面計内.ToString("#,##0"), ".00", ""), ""))
            pSheet.ValueReplace("{畑面積計内}", IIf(畑面計 > 0, 畑面計.ToString("#,##0"), "") & IIf(Is畑内, "(内" & 畑面計内.ToString("#,##0") & ")", ""))

            pSheet.ValueReplace("{樹筆数計}", 樹数計)
            pSheet.ValueReplace("{樹面積計}", IIf(樹面計内 > 0, Replace(樹面計内.ToString("#,##0"), ".00", ""), ""))
            pSheet.ValueReplace("{樹面積計内}", IIf(樹面計 > 0, 樹面計.ToString("#,##0"), "") & IIf(Is樹内, "(内" & 樹面計内.ToString("#,##0") & ")", ""))

            pSheet.ValueReplace("{他筆数計}", 他数計)
            pSheet.ValueReplace("{他面積計}", IIf(他面計内 > 0, Replace(他面計内.ToString("#,##0"), ".00", ""), ""))
            pSheet.ValueReplace("{他面積計内}", IIf(他面計 > 0, 他面計.ToString("#,##0"), "") & IIf(Is他内, "(内" & 他面計内.ToString("#,##0") & ")", ""))

            Dim 総計 As Decimal = 田面計 + 畑面計 + 樹面計 + 他面計
            Dim 総計内 As Decimal = 田面計内 + 畑面計内 + 樹面計内 + 他面計内
            pSheet.ValueReplace("{面積計}", Replace(総計内.ToString("#,##0.00"), ".00", ""))
            pSheet.ValueReplace("{面積計内}", 総計.ToString("#,##0") & IIf((Is田内 Or Is畑内 Or Is樹内 Or Is他内), "(内" & 総計内.ToString("#,##0") & ")", ""))
        End Sub

        Public Sub Replace明細総合計(ByRef pSheet As XMLSSWorkSheet)
            Dim 総筆数 As Decimal = 田数計 + 畑数計 + 樹数計 + 他数計
            Dim 総面計 As Decimal = 田面計 + 畑面計 + 樹面計 + 他面計
            Dim 総面計内 As Decimal = 田面計内 + 畑面計内 + 樹面計内 + 他面計内

            With pSheet
                .ValueReplace("{田筆数総合計}", 田数計.ToString("#,##0"))
                .ValueReplace("{田面積総合計}", 田面計内.ToString("#,##0"))
                .ValueReplace("{田面積総合計内}", 田面計.ToString("#,##0") & IIf(Is田内, "(内" & 田面計内.ToString("#,##0") & ")", ""))

                .ValueReplace("{畑筆数総合計}", 畑数計.ToString("#,##0"))
                .ValueReplace("{畑面積総合計}", 畑面計内.ToString("#,##0"))
                .ValueReplace("{畑面積総合計内}", 畑面計.ToString("#,##0") & IIf(Is畑内, "(内" & 畑面計内.ToString("#,##0") & ")", ""))

                .ValueReplace("{樹筆数総合計}", 樹数計.ToString("#,##0"))
                .ValueReplace("{樹面積総合計}", 樹面計内.ToString("#,##0"))
                .ValueReplace("{樹面積総合計内}", 樹面計.ToString("#,##0") & IIf(Is樹内, "(内" & 樹面計内.ToString("#,##0") & ")", ""))

                .ValueReplace("{他筆数総合計}", 他数計.ToString("#,##0"))
                .ValueReplace("{他面積総合計}", 他面計内.ToString("#,##0"))
                .ValueReplace("{他面積総合計内}", 他面計.ToString("#,##0") & IIf(Is他内, "(内" & 他面計内.ToString("#,##0") & ")", ""))

                .ValueReplace("{筆数総合計}", 総筆数.ToString("#,##0"))
                .ValueReplace("{面積総合計}", 総面計内.ToString("#,##0"))
                .ValueReplace("{面積総合計内}", 総面計.ToString("#,##0") & IIf(Is田内 Or Is畑内 Or Is樹内 Or Is他内, "(内" & 総面計内.ToString("#,##0") & ")", ""))
            End With
        End Sub

        Public Sub Add合計(ByVal p案件毎 As C筆明細と集計作成)
            田数計 += p案件毎.田数計
            畑数計 += p案件毎.畑数計
            樹数計 += p案件毎.樹数計
            他数計 += p案件毎.他数計

            田面計 += p案件毎.田面計
            畑面計 += p案件毎.畑面計
            樹面計 += p案件毎.樹面計
            他面計 += p案件毎.他面計

            Is田内 = Is田内 Or p案件毎.Is田内
            Is畑内 = Is畑内 Or p案件毎.Is畑内
            Is樹内 = Is樹内 Or p案件毎.Is樹内
            Is他内 = Is他内 Or p案件毎.Is他内

            田面計内 += p案件毎.田面計内
            畑面計内 += p案件毎.畑面計内
            樹面計内 += p案件毎.樹面計内
            他面計内 += p案件毎.他面計内
        End Sub

        Public Class C明細作成
            Inherits List(Of 筆毎面積)

            Public Sub Plus(ByVal pRow As DataRow, ByVal p本地面積 As Decimal, ByVal p部分面積 As Decimal, ByVal p田面積 As Decimal, ByVal p畑面積 As Decimal)
                Dim s土地所在 As String
                If pRow.Item("小字").ToString = "" OrElse pRow.Item("小字").ToString = "-" Then
                    s土地所在 = IIf(pRow.Item("所在").ToString.Length > 0, pRow.Item("所在").ToString, pRow.Item("大字").ToString & pRow.Item("地番").ToString)
                Else
                    s土地所在 = IIf(pRow.Item("所在").ToString.Length > 0, pRow.Item("所在").ToString, pRow.Item("大字").ToString & IIf(pRow.Item("小字").ToString.Length > 0, "字", "") & pRow.Item("小字").ToString) & pRow.Item("地番").ToString
                End If

                Me.Add(New 筆毎面積(s土地所在, p本地面積, p部分面積, p田面積, p畑面積))
            End Sub

            Public Function To面積文字列(Optional ByVal bySp As String = ",") As String
                Dim sB As New System.Text.StringBuilder

                For Each pNum As 筆毎面積 In Me
                    If Fix(pNum.本地面積) = pNum.本地面積 Then
                        sB.Append(IIf(sB.Length > 0, bySp, "") & pNum.本地面積.ToString("#,##0"))
                    Else
                        sB.Append(IIf(sB.Length > 0, bySp, "") & pNum.本地面積.ToString("#,##0.00"))
                    End If
                Next

                Return sB.ToString
            End Function
            Public Function To面積内文字列(Optional ByVal bySp As String = ",") As String
                Dim sB As New System.Text.StringBuilder

                For Each pNum As 筆毎面積 In Me
                    If Fix(pNum.本地面積) = pNum.本地面積 Then
                        sB.Append(IIf(sB.Length > 0, bySp, "") & pNum.本地面積.ToString("#,##0"))
                    Else
                        sB.Append(IIf(sB.Length > 0, bySp, "") & pNum.本地面積.ToString("#,##0.00"))
                    End If
                    Dim pD As Decimal = pNum.部分面積
                    If pD > 0 Then
                        sB.Append("(内")
                        If Fix(pNum.部分面積) = pNum.部分面積 Then
                            sB.Append(pNum.部分面積.ToString("#,##0"))
                        Else
                            sB.Append(pNum.部分面積.ToString("#,##0.00"))
                        End If
                        sB.Append(")")
                    End If
                Next


                Return sB.ToString
            End Function

            Public Function To土地所在文字列(Optional ByVal bySp As String = ",") As String
                Dim sB As New System.Text.StringBuilder

                For Each pNum As 筆毎面積 In Me
                    sB.Append(IIf(sB.Length > 0, bySp, "") & pNum.土地所在)
                Next

                Return sB.ToString
            End Function

            Public Class 筆毎面積
                Public 土地所在 As String = ""
                Public 本地面積 As Decimal
                Public 部分面積 As Decimal
                Public 田面積 As Decimal
                Public 畑面積 As Decimal

                Public Sub New(ByVal s所在 As String, ByVal d本地面積 As Decimal, ByVal d部分面積 As Decimal, ByVal dt田面積 As Decimal, ByVal dt畑面積 As Decimal)
                    土地所在 = s所在
                    本地面積 = d本地面積
                    部分面積 = d部分面積
                    田面積 = dt田面積
                    畑面積 = dt畑面積
                End Sub

            End Class
        End Class

    End Class

    Public Sub 転用共通(ByRef pSheet As XMLSSWorkSheet, ByRef pRow As DataRowView)
        pSheet.ValueReplace("{資金計画}", pRow.Item("資金計画").ToString)
        If pRow.Item("数量").ToString.Length > 0 AndAlso pRow.Item("数量") > 0 Then
            pSheet.ValueReplace("{棟数}", pRow.Item("数量").ToString & "棟")
        Else
            pSheet.ValueReplace("{棟数}", "")
        End If
        If pRow.Item("数量").ToString.Length > 0 AndAlso pRow.Item("数量") > 0 Then
            pSheet.ValueReplace("{棟数n}", pRow.Item("数量").ToString & "棟")
        Else
            pSheet.ValueReplace("{棟数n}", " 0")
        End If

        If pRow.Item("建築面積").ToString.Length > 0 AndAlso pRow.Item("建築面積") > 0 Then
            pSheet.ValueReplace("{建築面積}", pRow.Item("建築面積").ToString & "㎡")
        Else
            pSheet.ValueReplace("{建築面積}", "")
        End If

        If pRow.Item("用途").ToString.Length > 0 Then
            pSheet.ValueReplace("{用途}", pRow.Item("用途").ToString)
        Else
            pSheet.ValueReplace("{用途}", "")
        End If

        If pRow.Item("申請地目安").ToString.Length > 0 Then
            pSheet.ValueReplace("{申請地目安}", pRow.Item("申請地目安").ToString)
        Else
            pSheet.ValueReplace("{申請地目安}", "")
        End If

        Dim St1 As String = IIf(pRow.Item("工事開始年1").ToString.Length > 0, pRow.Item("工事開始年1"), "") & "." & IIf(pRow.Item("工事開始月1").ToString.Length > 0, pRow.Item("工事開始月1"), "")
        Dim St2 As String = IIf(pRow.Item("工事終了年1").ToString.Length > 0, pRow.Item("工事終了年1"), "") & "." & IIf(pRow.Item("工事終了月1").ToString.Length > 0, pRow.Item("工事終了月1"), "")

        Dim wSt1 As String = IIf(pRow.Item("工事開始年1").ToString.Length > 0, pRow.Item("工事開始年1"), "　") & "年" & IIf(pRow.Item("工事開始月1").ToString.Length > 0, pRow.Item("工事開始月1"), "　") & "月"
        Dim wSt2 As String = IIf(pRow.Item("工事終了年1").ToString.Length > 0, pRow.Item("工事終了年1"), "　") & "年" & IIf(pRow.Item("工事終了月1").ToString.Length > 0, pRow.Item("工事終了月1"), "　") & "月"

        If St1.Length > 1 Then St1 = St1 & "～&#10;"
        If St2.Length > 1 Then St1 = St1 & St2
        '
        pSheet.ValueReplace("{工事着工}", St1)
        pSheet.ValueReplace("{工事計画}", St1)
        pSheet.ValueReplace("{工事開始}", wSt1)
        pSheet.ValueReplace("{工事終了}", wSt2)

        pSheet.ValueReplace("{目的}", "")
        If Not IsDBNull(pRow.Item("資金計画")) AndAlso pRow.Item("資金計画").ToString.Length > 0 Then
            pSheet.ValueReplace("{資金}", pRow.Item("資金計画").ToString)
        Else
            pSheet.ValueReplace("{資金}", "")
        End If
        区分設定(pSheet, pRow)
    End Sub

    Public Sub 区分設定(ByRef pSheet As XMLSSWorkSheet, ByRef pRow As DataRowView)
        pSheet.ValueReplace("{農地区分}", pRow.Item("農地区分名称").ToString)
        Select Case Val(pRow.Item("農地区分").ToString)
            Case 1 : pSheet.ValueReplace("{諮問農地区分}", 1)
            Case 2 : pSheet.ValueReplace("{諮問農地区分}", 2)
            Case 3 : pSheet.ValueReplace("{諮問農地区分}", 3)
            Case 4, 5 : pSheet.ValueReplace("{諮問農地区分}", "農振農用地")
            Case Else : pSheet.ValueReplace("{諮問農地区分}", "-")
        End Select

        pSheet.ValueReplace("{農振区分名称}", pRow.Item("農振区分名称").ToString)
        pSheet.ValueReplace("{都計区分名称}", pRow.Item("都計区分名称").ToString)
    End Sub

    Public Sub 受付情報(ByRef pSheet As XMLSSWorkSheet, ByRef pRow As DataRowView)
        pSheet.ValueReplace("{受付番号}", pRow.Item("受付番号"))
        pSheet.ValueReplace("{受付年月日}", 和暦Format(CDate(pRow.Item("受付年月日"))))
    End Sub

    Public Sub 権利内容(ByRef pSheet As XMLSSWorkSheet, ByRef pRow As System.Data.DataRowView)
        Select Case CType(pRow.Item("法令"), enum法令)
            Case enum法令.農地法5条所有権
                pSheet.ValueReplace("{権利内容}", "所有権移転")
            Case enum法令.農地法5条貸借
                Select Case Val(pRow.Item("権利種類").ToString)
                    Case 1 : pSheet.ValueReplace("{権利内容}", "賃借権")
                    Case 2 : pSheet.ValueReplace("{権利内容}", "使用貸借権")
                    Case Else : pSheet.ValueReplace("{権利内容}", "その他")
                End Select
            Case enum法令.農地法5条一時転用
                pSheet.ValueReplace("{権利内容}", "期間借地")
        End Select
    End Sub

    Public Sub 申請者A(ByRef pSheet As XMLSSWorkSheet, ByRef pRow As DataRowView)
        If pRow.Item("代理人名").ToString.Length > 0 Then
            pSheet.ValueReplace("{申請者Ａ氏名}", "[代理人]" & pRow.Item("代理人名").ToString)
            pSheet.ValueReplace("{申請者Ａ住所}", "[代理人]" & pRow.Item("代理人住所").ToString)
        Else
            pSheet.ValueReplace("{申請者Ａ氏名}", pRow.Item("氏名A").ToString)
            pSheet.ValueReplace("{申請者Ａ住所}", pRow.Item("住所A").ToString)
        End If
    End Sub

    Public Sub 申請者B(ByRef pSheet As XMLSSWorkSheet, ByRef pRow As DataRowView)
        pSheet.ValueReplace("{申請者Ｂ氏名}", pRow.Item("氏名B").ToString)
        pSheet.ValueReplace("{申請者Ｂ住所}", pRow.Item("住所B").ToString)
    End Sub

    Public Sub 調査委員(ByRef pSheet As XMLSSWorkSheet, ByRef pRow As DataRowView)
        If Not IsDBNull(pRow.Item("農業委員1")) Then
            Dim pRow1 As DataRow = App農地基本台帳.DataMaster.Rows.Find({pRow.Item("農業委員1"), "農業委員"})
            If pRow1 IsNot Nothing Then
                pSheet.ValueReplace("{調査委員}", IIf(IsDBNull(pRow1.Item("名称")), pRow.Item("調査員A").ToString, pRow1.Item("名称").ToString))
                pSheet.ValueReplace("{調査委員1}", IIf(IsDBNull(pRow1.Item("名称")), pRow.Item("調査員A").ToString, pRow1.Item("名称").ToString))
            End If
        End If
        pSheet.ValueReplace("{調査委員}", pRow.Item("調査員A").ToString)
        pSheet.ValueReplace("{調査委員1}", pRow.Item("調査員A").ToString)

        If Not IsDBNull(pRow.Item("農業委員2")) Then
            Dim pRow2 As DataRow = App農地基本台帳.DataMaster.Rows.Find({pRow.Item("農業委員2"), "農業委員"})
            If pRow2 IsNot Nothing Then
                pSheet.ValueReplace("{調査委員2}", pRow2.Item("名称").ToString)
            End If
        End If
        pSheet.ValueReplace("{調査委員2}", "")

        If Not IsDBNull(pRow.Item("農業委員3")) Then
            Dim pRow3 As DataRow = App農地基本台帳.DataMaster.Rows.Find({pRow.Item("農業委員3"), "農業委員"})
            If pRow3 IsNot Nothing Then
                pSheet.ValueReplace("{調査委員3}", pRow3.Item("名称").ToString)
            End If
        End If
        pSheet.ValueReplace("{調査委員3}", "")

        pSheet.ValueReplace("{担当委員}", pRow.Item("調査員B").ToString)
    End Sub

    Private Function GetCVNumber(ByVal pObj As Object) As String
        If pObj Is Nothing Then
            Return ""
        ElseIf IsDBNull(pObj) Then
            Return ""
        ElseIf CDbl(pObj) > 0 Then
            Return Val(pObj).ToString("#,###")
        Else
            Return ""
        End If
    End Function

    Public Overloads Overrides Sub SetData(ByRef XMLSS As HimTools2012.Excel.XMLSS2003.CXMLSS2003, ByRef pTab As 諮問意見書Page, ByVal s処理名称 As String, ByRef pDataCreater As C諮問意見書Data作成)
        If XMLSS.WorkBook.WorkSheets.Items.ContainsKey("明細") Then
            With XMLSS.WorkBook.WorkSheets.Items("明細")
                Set複数行(._object, pTab, s処理名称, pDataCreater, "")
                mvar総合計.Replace明細総合計(XMLSS.WorkBook.WorkSheets.Items("明細"))

            End With
        Else
            For Each pS As XMLSSWorkSheet In XMLSS.WorkBook.WorkSheets.Items.Values
                Set複数行(pS, pTab, s処理名称, pDataCreater, "")
                mvar総合計.Replace明細総合計(pS)
            Next
        End If
    End Sub

    Public Overrides Sub SetDataRow(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRow As System.Data.DataRowView, Optional ByVal pObj As Object = Nothing)

    End Sub
End Class

Public MustInherit Class CPrint諮問意見書単票作成
    Inherits CPrint諮問意見書作成共通

    Public MustOverride Sub Set単票Data(ByRef pDataCreater As C諮問意見書Data作成, ByVal pSheet As XMLSSWorkSheet)

    Public Sub New()
        MyBase.new()
    End Sub

    Public Overrides Sub SetData(ByRef XMLSS As CXMLSS2003, ByRef pTab As 諮問意見書Page, ByVal s処理名称 As String, ByRef pDataCreater As C諮問意見書Data作成)
        For Each pS As XMLSSWorkSheet In XMLSS.WorkBook.WorkSheets.Items.Values
            Set単票Data(pDataCreater, pS)
        Next
    End Sub

End Class
