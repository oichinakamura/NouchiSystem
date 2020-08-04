Public Class C三股町
    Inherits C市町村別

    ''' <summary>
    ''' \\192.168.1.40\MapFolder\システム配置\三股町\
    ''' </summary>
    ''' <remarks>mmt-svkouzai\Admin_HIM\admin</remarks>
    Public Sub New()
        MyBase.New("三股町")
    End Sub
    Public Overrides ReadOnly Property 旧農振都市計画使用 As Boolean
        Get
            Return True
        End Get
    End Property
    Public Overrides Sub InitMenu(pMain As HimTools2012.SystemWindows.CMainPageSK)
        With pMain
            .ListView.Clear()
            .ListView.ItemAdd("農家検索", "農家検索", "閲覧・検索", "閲覧・検索", AddressOf 農家一覧)
            .ListView.ItemAdd("利用権終期台帳", "利用権終期台帳", "印刷", "印刷", AddressOf sub利用権終期管理)


            .ListView.ItemAdd("固定資産比較", "固定資産比較", "他システム連携", "操作", AddressOf 固定資産比較)
            .ListView.ItemAdd("農地重複", "農地重複", "他システム連携", "操作", AddressOf sub農地重複)
            .ListView.ItemAdd("農地台帳一括印刷", "農地台帳一括印刷", "印刷", "印刷", AddressOf ClickMenu)
            .ListView.ItemAdd("総会資料作成", "総会資料作成", ImageKey.作業, "操作", AddressOf ClickMenu)
            .ListView.ItemAdd("自然解約の実行", "自然解約の実行", ImageKey.作業, "操作", AddressOf ClickMenu)


            .ListView.ItemAdd("公開用CSV正規化", "公開用CSV正規化", ImageKey.作業, "操作", AddressOf 公開用CSV正規化)
            .ListView.ItemAdd("農地台帳一括印刷", "農地台帳一括印刷", "印刷", "印刷", AddressOf ClickMenu)
            MyBase.InitMenu(pMain)
        End With

    End Sub
    Public Overrides Function Get選挙世帯一覧() As System.Data.DataTable
        With SysAD.DB(sLRDB)
            .ExecuteSQL("UPDATE [D:個人Info] INNER JOIN [D:世帯Info] ON [D:個人Info].世帯ID = [D:世帯Info].ID SET [D:世帯Info].世帯主ID = [D:個人Info].[ID] WHERE ((([D:世帯Info].世帯主ID) Is Null Or ([D:世帯Info].世帯主ID)<>[D:個人Info].[ID]) AND (([D:個人Info].続柄1)=2) AND (([D:個人Info].続柄2)=0) AND (([D:個人Info].住民区分)=0));")


            Dim pT1 As DataTable = .GetTableBySqlSelect("SELECT [D:世帯Info].ID,[D:世帯Info].世帯主ID,'-' AS [世帯郵便番号],[D:世帯Info].選挙連番, Sum(IIF([農地状況]<=29,1,0)*([V_農地].[田面積]+[V_農地].[畑面積]+[V_農地].[樹園地])) AS 式2, Sum(([V_農地].[田面積]+[V_農地].[畑面積]+[V_農地].[樹園地])) AS 式1 FROM V_農地 INNER JOIN [D:世帯Info] ON V_農地.耕作世帯ID = [D:世帯Info].ID GROUP BY [D:世帯Info].ID,[D:世帯Info].世帯主ID,選挙連番 HAVING ((([D:世帯Info].ID)>0));")
            pT1.PrimaryKey = New DataColumn() {pT1.Columns("ID")}

            Dim pT2 As DataTable = .GetTableBySqlSelect("SELECT * FROM [D_個人履歴] WHERE [異動事由] IN (20111,20112)")
            pT2.PrimaryKey = New DataColumn() {pT2.Columns("PID"), pT2.Columns("異動事由")}

            Dim pTable As DataTable = .GetTableBySqlSelect("SELECT [世帯ID],[D:個人Info].[ID],[氏名],[フリガナ],[D:個人Info].[行政区ID],[D:個人Info].[行政区ID] & 'g' AS [行政区],[D:個人Info].郵便番号,[D:個人Info].[住所],[生年月日],[続柄1] & '-' & [続柄2] & '-' & [続柄3] AS [続柄],Choose([性別]+1,'男','女') AS 性別s,0 AS 世帯面積,0 AS 耕作面積,[選挙権の有無] AS 前年度の選挙権,0 AS 審査年の年齢 FROM [D:個人Info] LEFT JOIN [D:世帯Info] ON [D:個人Info].世帯ID = [D:世帯Info].ID WHERE [世帯ID]>0 AND [D:個人Info].[ID]>0 AND [住民区分]=0 ORDER BY [D:個人Info].[行政区ID],[世帯ID],[続柄1],[続柄2],[続柄3],[D:個人Info].[ID],[D:個人Info].[生年月日]")
            pTable.Columns.Add(New DataColumn("年齢要件", GetType(Boolean)))

            Dim mvar続柄 As DataTable = .GetTableBySqlSelect("SELECT * FROM [V_続柄]")
            mvar続柄.PrimaryKey = New DataColumn() {mvar続柄.Columns("ID")}

            Dim mvar行政区 As DataTable = .GetTableBySqlSelect("SELECT *,0 As [Used] FROM [V_行政区] ORDER BY [ID]")
            mvar行政区.PrimaryKey = New DataColumn() {mvar行政区.Columns("ID")}


            Dim p世帯 As New DataTable("選挙世帯")
            p世帯.Columns.Add(New DataColumn("ID", GetType(Integer)))
            p世帯.Columns.Add(New DataColumn("印刷Flag", GetType(Boolean)))
            p世帯.Columns.Add(New DataColumn("選挙連番", GetType(Integer)))
            p世帯.Columns.Add(New DataColumn("世帯主ID", GetType(Integer)))
            p世帯.Columns.Add(New DataColumn("世帯主名", GetType(String)))
            p世帯.Columns.Add(New DataColumn("世帯主フリガナ", GetType(String)))
            p世帯.Columns.Add(New DataColumn("世帯住所", GetType(String)))
            p世帯.Columns.Add(New DataColumn("世帯郵便番号", GetType(String)))
            p世帯.Columns.Add(New DataColumn("行政区ID", GetType(Integer)))
            p世帯.Columns.Add(New DataColumn("行政区", GetType(String)))
            p世帯.Columns.Add(New DataColumn("出力エリア", GetType(String)))
            p世帯.Columns.Add(New DataColumn("世帯面積", GetType(Double)))
            p世帯.Columns.Add(New DataColumn("耕作面積", GetType(Double)))
            p世帯.Columns.Add(New DataColumn("面積要件", GetType(Boolean)))
            p世帯.Columns.Add(New DataColumn("強制要件", GetType(Boolean)))
            p世帯.Columns.Add(New DataColumn("部分出力", GetType(Boolean)))
            p世帯.Columns.Add(New DataColumn("男", GetType(Integer)))
            p世帯.Columns.Add(New DataColumn("女", GetType(Integer)))
            p世帯.Columns.Add(New DataColumn("人数", GetType(String)))
            p世帯.Columns.Add(New DataColumn("世帯員", GetType(String)))

            p世帯.PrimaryKey = New DataColumn() {p世帯.Columns("ID")}
            Dim sError As String = ""

            Try
                sError = "001"
                For Each pRow As DataRow In pTable.Rows
                    sError = "002"
                    Dim p強制出力 As DataRow = pT2.Rows.Find(New Object() {pRow.Item("ID"), "20111"})
                    Dim p送付拒否 As DataRow = pT2.Rows.Find(New Object() {pRow.Item("ID"), "20112"})

                    Try
                        If Now.Month < 10 Then
                            pRow.Item("審査年の年齢") = HimTools2012.DateFunctions.年齢(pRow.Item("生年月日"), DateSerial(Now.Year, 4, 1))
                        Else
                            pRow.Item("審査年の年齢") = HimTools2012.DateFunctions.年齢(pRow.Item("生年月日"), DateSerial(Now.Year + 1, 4, 1))
                        End If

                        pRow.Item("年齢要件") = (pRow.Item("審査年の年齢") >= 20)
                    Catch ex As Exception
                        pRow.Item("年齢要件") = False
                    End Try

                    sError = "003"
                    If pRow.Item("年齢要件") = True AndAlso (p強制出力 IsNot Nothing OrElse p送付拒否 Is Nothing) Then
                        sError = "003-1"
                        Dim p世帯Row As DataRow = p世帯.Rows.Find(pRow.Item("世帯ID"))
                        sError = "003-2"

                        If Not IsDBNull(pRow.Item("続柄")) Then
                            sError = "003-3"
                            Dim Ar() As String = pRow.Item("続柄").ToString.Split("-")
                            sError = "003-4"
                            Dim sBL As New System.Text.StringBuilder

                            sError = "003-5"
                            For Each Sx As String In Ar
                                sError = "003-6"
                                If Len(Sx) Then
                                    Dim pZRow As DataRow = mvar続柄.Rows.Find(Integer.Parse(Sx))
                                    sError = "003-7"
                                    If Not Sx = "0" AndAlso pZRow IsNot Nothing Then
                                        sError = "003-8"

                                        If sBL.Length = 0 Then
                                            sError = "003-9"

                                            sBL.Append(pZRow.Item("名称"))
                                        ElseIf pZRow IsNot Nothing Then
                                            sError = "003-10"

                                            sBL.Append("の" & pZRow.Item("名称"))
                                        End If
                                    End If
                                End If
                            Next
                            sError = "003-11"

                            pRow.Item("続柄") = sBL.ToString
                        End If
                        sError = "004"

                        If p世帯Row Is Nothing Then
                            p世帯Row = p世帯.NewRow
                            p世帯Row.Item("ID") = pRow.Item("世帯ID")
                            p世帯Row.Item("男") = 0
                            p世帯Row.Item("女") = 0
                            p世帯Row.Item("人数") = 0
                            p世帯Row.Item("印刷Flag") = False
                            p世帯Row.Item("部分出力") = False
                            p世帯.Rows.Add(p世帯Row)
                            Dim p世帯面積情報 As DataRow = pT1.Rows.Find(pRow.Item("世帯ID"))
                            p世帯Row.Item("強制要件") = False
                            If Not p世帯面積情報 Is Nothing Then
                                p世帯Row.Item("選挙連番") = p世帯面積情報.Item("選挙連番")
                                p世帯Row.Item("世帯主ID") = p世帯面積情報.Item("世帯主ID")
                                p世帯Row.Item("世帯面積") = p世帯面積情報.Item("式1")
                                p世帯Row.Item("耕作面積") = p世帯面積情報.Item("式2")
                                p世帯Row.Item("面積要件") = (p世帯面積情報.Item("式2") >= 1000)
                            Else
                                p世帯Row.Item("世帯主ID") = -1
                                p世帯Row.Item("世帯面積") = 0
                                p世帯Row.Item("耕作面積") = 0
                                p世帯Row.Item("面積要件") = False
                            End If
                        End If

                        sError = "005"
                        p世帯Row.Item("強制要件") = p世帯Row.Item("強制要件") Or (p強制出力 IsNot Nothing)
                        If pRow.Item("続柄") = "世帯主" OrElse (Not IsDBNull(p世帯Row.Item("世帯主ID")) AndAlso p世帯Row.Item("世帯主ID") = pRow.Item("ID")) Then
                            p世帯Row.Item("世帯主名") = pRow.Item("氏名")
                            p世帯Row.Item("世帯主フリガナ") = pRow.Item("フリガナ")
                            p世帯Row.Item("世帯住所") = pRow.Item("住所")
                            p世帯Row.Item("世帯郵便番号") = pRow.Item("郵便番号")

                            If pRow.Item("住所").ToString.StartsWith("姶良市加治木町") Then
                                p世帯Row.Item("出力エリア") = "加治木"
                            ElseIf pRow.Item("住所").ToString.StartsWith("姶良市蒲生町") Then
                                p世帯Row.Item("出力エリア") = "蒲生"
                            Else
                                p世帯Row.Item("出力エリア") = "姶良"
                            End If

                            If Not IsDBNull(pRow.Item("行政区")) Then
                                Dim pZRow As DataRow = mvar行政区.Rows.Find(Val(pRow.Item("行政区")))
                                If pZRow IsNot Nothing Then
                                    p世帯Row.Item("行政区ID") = pRow.Item("行政区ID")
                                    p世帯Row.Item("行政区") = pZRow.Item("名称")

                                End If
                            Else
                                Stop
                            End If
                        End If

                        sError = "006"
                        p世帯Row.Item("男") -= (pRow.Item("性別s") = "男")
                        p世帯Row.Item("女") -= (pRow.Item("性別s") = "女")

                        p世帯Row.Item("人数") = p世帯Row.Item("男") + p世帯Row.Item("女")
                        p世帯Row.Item("部分出力") = p世帯Row.Item("部分出力") Or pRow.Item("前年度の選挙権")
                        p世帯Row.Item("世帯員") = p世帯Row.Item("世帯員").ToString &
                                IIf(Len(p世帯Row.Item("世帯員").ToString) > 0, vbCrLf, "") &
                                pRow.Item("氏名") & ";" &
                                pRow.Item("続柄") & ";" &
                                pRow.Item("生年月日") & ";" & pRow.Item("性別s") & ";" & pRow.Item("前年度の選挙権")
                    ElseIf p送付拒否 Is Nothing Then


                    End If
                    sError = "007"

                Next
            Catch ex As Exception
                MsgBox(sError & ":" & ex.Message)

                Stop
            End Try

            For Each p世帯Row As DataRow In p世帯.Rows
                If IsDBNull(p世帯Row.Item("世帯主名")) Then
                    p世帯Row.Item("世帯主名") = "×"
                    p世帯Row.Item("行政区ID") = "999999"

                End If
            Next

            Dim pView As New DataView(p世帯, "[世帯主名]<>'×' AND ([強制要件] = True Or [面積要件] = True)", "[行政区ID],[世帯主フリガナ]", DataViewRowState.CurrentRows)

            Return pView.ToTable
        End With
    End Function

    Public Overrides Function Get地区情報(ByVal s住所 As String) As String
        Return MyBase.Get地区情報(s住所)
    End Function

    Private mvar利用権終期管理 As C利用権終期管理

    Public Overrides ReadOnly Property Get受付しめ月 As Integer
        Get
            Return -1
        End Get
    End Property
    Public Overrides ReadOnly Property Get総会締日 As Integer
        Get
            Select Case Month(Now) - 1
                Case 1, 3, 5, 7, 8, 10, 12, 0
                    Return 31
                Case 4, 6, 9, 11
                    Return 20
                Case 2
                    Return 28
            End Select

            Return MyBase.Get総会締日
        End Get
    End Property
    Public Overrides ReadOnly Property Get総会次受付日 As Integer
        Get
            Return 1
        End Get
    End Property
    Public Overrides Sub InitLocalData()
        With New dlgLoginForm()

            If Not .ShowDialog() = Windows.Forms.DialogResult.OK Then
                Try
                    End
                Catch ex As Exception

                End Try
            End If
        End With

        Dim pTBLE As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:個人Info] Where [ID]=0")
        If pTBLE.Columns.Contains("作業者・作業PC") Then
            SysAD.DB(sLRDB).ExecuteSQL("ALTER TABLE [D:個人Info] DROP [作業者・作業PC]")
        End If

        Dim pDBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM D_転用農地 WHERE (((D_転用農地.[小作形態])=7) AND ((D_転用農地.小作終了年月日)<Now())) OR (((D_転用農地.[農地状況])=1040) AND ((D_転用農地.小作終了年月日)<Now()));")
        If pDBL.Rows.Count > 0 Then
            If MsgBox("期間終了した一時転用があります。元に戻しますか?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                For Each pRow As DataRow In pDBL.Rows
                    Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE [ID]=" & pRow.Item("ID"))

                    If pTBL.Rows.Count > 0 Then
                        SysAD.DB(sLRDB).ExecuteSQL("DELETE FROM [D:農地Info] WHERE [ID]=" & pRow.Item("ID"))
                    End If

                    SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D:農地Info]([ID]) VALUES({0})", pRow.Item("ID"))
                    Dim sSQL As New System.Text.StringBuilder
                    For Each pCol As DataColumn In pTBL.Columns
                        If pDBL.Columns.Contains(pCol.ColumnName) Then
                            If pCol.ColumnName = "ID" Then
                            ElseIf Not IsDBNull(pRow.Item(pCol.ColumnName)) Then
                                Select Case pCol.DataType.Name
                                    Case "Int16", "Int32", "Single", "Double", "Decimal"
                                        sSQL.Append(IIf(sSQL.Length > 0, ",", "") & String.Format("[{0}]={1}", pCol.ColumnName, pRow.Item(pCol.ColumnName)))
                                    Case "Boolean"
                                        If pRow.Item(pCol.ColumnName) Then
                                            sSQL.Append(IIf(sSQL.Length > 0, ",", "") & String.Format("[{0}]=True", pCol.ColumnName, pRow.Item(pCol.ColumnName)))
                                        Else
                                            sSQL.Append(IIf(sSQL.Length > 0, ",", "") & String.Format("[{0}]=False", pCol.ColumnName, pRow.Item(pCol.ColumnName)))
                                        End If
                                    Case "String"
                                        If pRow.Item(pCol.ColumnName).ToString.Length > 0 Then
                                            sSQL.Append(IIf(sSQL.Length > 0, ",", "") & String.Format("[{0}]='{1}'", pCol.ColumnName, pRow.Item(pCol.ColumnName)))
                                        Else
                                            sSQL.Append(IIf(sSQL.Length > 0, ",", "") & String.Format("[{0}]=Null", pCol.ColumnName))
                                        End If
                                    Case "DateTime"
                                        If CDate(pRow.Item(pCol.ColumnName)) > #1/1/1900# Then
                                            With CDate(pRow.Item(pCol.ColumnName))
                                                sSQL.Append(IIf(sSQL.Length > 0, ",", "") & String.Format("[{0}]=#{2}/{3}/{1}#", pCol.ColumnName, .Year, .Month, .Day))
                                            End With
                                        Else
                                            sSQL.Append(IIf(sSQL.Length > 0, ",", "") & String.Format("[{0}]=Null", pCol.ColumnName))
                                        End If
                                    Case Else
                                        Stop
                                End Select
                                If sSQL.Length > 768 Then
                                    Dim St As String = SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET " & sSQL.ToString & " WHERE [ID]=" & pRow.Item("ID"))
                                    sSQL.Clear()
                                End If
                            End If
                        End If
                    Next
                    If sSQL.Length > 0 Then
                        SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET " & sSQL.ToString & " WHERE [ID]=" & pRow.Item("ID"))
                    End If
                    Dim pEndDT As DateTime = pRow.Item("小作終了年月日")
                    Dim sEndDT As String = "#" & pEndDT.Month & "/" & pEndDT.Day & "/" & pEndDT.Year & "#"
                    SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO D_土地履歴 (LID, 更新日, 異動日, 異動事由, 内容, 入力日) VALUES({0}, Now(),{1}, 441, '一時転用の終了[{2}]', Now())", pRow.Item("ID"), sEndDT, pEndDT.ToShortDateString)
                    SysAD.DB(sLRDB).ExecuteSQL("DELETE FROM [D_転用農地] WHERE [ID]=" & pRow.Item("ID"))
                Next
            End If
        End If

    End Sub

    Public Sub 公開用CSV正規化()
        If MsgBox("公開用CSV正規化を実行しますか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:農地Info].[大字ID], [D:農地Info].[地番], [D:農地Info].[ID], [D:農地Info].一部現況 FROM [D:農地Info] WHERE ((([D:農地Info].大字ID) In (SELECT [大字ID] FROM [D:農地Info] As Tmp GROUP BY [大字ID],[地番] HAVING Count(*)>1  And [地番] = [D:農地Info].[地番]) And ([D:農地Info].大字ID)>0) AND (([D:農地Info].一部現況)=0)) ORDER BY [D:農地Info].大字ID, [D:農地Info].地番, [D:農地Info].一部現況;")
            Dim n大字 As Integer = 0
            Dim s地番 As String = ""
            Dim nList As New List(Of Long)

            For Each pRow As DataRow In pTBL.Rows
                If n大字 = 0 And s地番 = "" Then
                    n大字 = pRow.Item("大字ID")
                    s地番 = pRow.Item("地番")
                    nList.Clear()
                    nList.Add(pRow.Item("ID"))
                ElseIf n大字 <> pRow.Item("大字ID") OrElse s地番 <> pRow.Item("地番") Then
                    Dim n As Integer = 1
                    For Each nID As Long In nList
                        SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [一部現況]={0} WHERE [ID]={1}", n, nID)
                        n += 1
                    Next

                    n大字 = pRow.Item("大字ID")
                    s地番 = pRow.Item("地番")
                    nList.Clear()
                    nList.Add(pRow.Item("ID"))
                Else
                    nList.Add(pRow.Item("ID"))
                End If
            Next

            If nList.Count > 0 Then
                Dim n As Integer = 1
                For Each nID As Long In nList
                    SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [一部現況]={0} WHERE [ID]={1}", n, nID)
                    n += 1
                Next
                nList.Clear()
            End If
        End If
        SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [D:農地Info].利用状況調査農地法 = 3 WHERE ((([D:農地Info].利用状況調査農地法) Is Null));")
        SysAD.DB(sLRDB).ExecuteSQL("'UPDATE [D:農地Info] SET [D:農地Info].利用状況調査荒廃 = 0 WHERE ((([D:農地Info].利用状況調査荒廃) Is Null));")
        MsgBox("終了しました")
    End Sub

End Class
