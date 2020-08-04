

''' <summary>
''' 日置市農業委員会
''' </summary>
''' <remarks>\\10.100.2.16\gis\農家台帳\日置市\</remarks>
Public Class C日置市
    Inherits C市町村別

    Public Overrides Sub InitMenu(pMain As HimTools2012.SystemWindows.CMainPageSK)

        With pMain
            .ListView.Clear()

            .ListView.ItemAdd("農家検索", "農家検索", "閲覧・検索", "閲覧・検索", AddressOf 農家一覧)
            '    'mvarListView.ItemAdd("申請受付・審査・許可・取り下げ・取り消し", "申請受付・審査・許可・取り下げ・取り消し", "閲覧・検索", "閲覧・検索", AddressOf 申請処理)
            '    'mvarListView.ItemAdd("非農地通知", "非農地通知", "閲覧・検索", "閲覧・検索", AddressOf sub非農地通知)

            .ListView.ItemAdd("利用権終期台帳", "利用権終期台帳", "印刷", "印刷", AddressOf sub利用権終期管理)


            If SysAD.SystemInfo.ユーザー.n権利 > 0 Then

                .ListView.ItemAdd("申請・許可名簿", "申請・許可名簿", "印刷", "印刷", AddressOf ClickMenu)

                .ListView.ItemAdd("住記農家照合", "住記―農家照合", "他システム連携", "操作", AddressOf ClickMenu)
                .ListView.ItemAdd("固定資産比較", "固定資産比較", "他システム連携", "操作", AddressOf 固定資産比較)

                .ListView.ItemAdd("農地権利移動・借賃等調査システム出力", "農地権利移動・借賃等調査システム出力", ImageKey.作業, "操作", AddressOf ClickMenu)
                .ListView.ItemAdd("非農地通知", "非農地通知", "閲覧・検索", "閲覧・検索", AddressOf ClickMenu)

                .ListView.ItemAdd("総会資料作成", "総会資料作成", ImageKey.作業, "操作", AddressOf ClickMenu)
                .ListView.ItemAdd("耕作放棄地情報との関連付け", "耕作放棄地情報との関連付け", "メンテナンス", "設定", AddressOf ClickMenu)
                .ListView.ItemAdd("自然解約の実行", "自然解約の実行", ImageKey.作業, "操作", AddressOf ClickMenu)

                '.ListView.ItemAdd("TestXB", "TestXB",ImageKey.作業, "操作", AddressOf XXX)
                '.ListView.ItemAdd("男女設定", "男女設定",ImageKey.作業, "操作", AddressOf ClickMenu)

                .ListView.ItemAdd("転用非農地整合性確認", "転用非農地整合性確認", ImageKey.作業, "操作", AddressOf ClickMenu)
            End If

            .ListView.ItemAdd("農地賃借料データ", "農地賃借料データ", "集計一覧", "印刷", AddressOf 農地賃借料データ3年間)
            .ListView.ItemAdd("農業改善計画認定別経営面積", "農業改善計画認定別経営面積", "印刷", "印刷", AddressOf sub農業改善計画認定別経営面積)
            .ListView.ItemAdd("農用地面積", "農用地面積", ImageKey.作業, "操作", AddressOf sub農用地面積)

            .ListView.ItemAdd("農地台帳一括印刷", "農地台帳一括印刷", "印刷", "印刷", AddressOf ClickMenu)

            .ListView.ItemAdd("フェーズ２利用状況調査・意向調査情報", "フェーズ２利用状況調査・意向調査情報", ImageKey.他システム連携, "他システム連携", AddressOf ClickMenu)
            MyBase.InitMenu(pMain)
        End With

    End Sub
    Public Overrides ReadOnly Property 旧農振都市計画使用 As Boolean
        Get
            Return True
        End Get
    End Property


    Private Sub 農地賃借料データ3年間()
        If Not SysAD.MainForm.MainTabCtrl.ExistPage("農地賃借料データ(過去3年分)") Then
            SysAD.MainForm.MainTabCtrl.AddPage(New C農地賃借料データ)
        End If
    End Sub


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




    Public Sub New()
        MyBase.New("日置市")

    End Sub

    Public Overrides Function Get選挙世帯一覧() As System.Data.DataTable
        With SysAD.DB(sLRDB)

            .ExecuteSQL("UPDATE [D:個人Info] INNER JOIN [D:世帯Info] ON [D:個人Info].世帯ID = [D:世帯Info].ID SET [D:世帯Info].世帯主ID = [D:個人Info].[ID] WHERE ((([D:世帯Info].世帯主ID) Is Null Or ([D:世帯Info].世帯主ID)<>[D:個人Info].[ID]) AND (([D:個人Info].続柄1)=2) AND (([D:個人Info].続柄2)=0) AND (([D:個人Info].住民区分)=0));")

            Dim pT1 As DataTable = .GetTableBySqlSelect("SELECT [D:世帯Info].ID,[D:世帯Info].世帯主ID,'-' AS [世帯郵便番号],[D:世帯Info].選挙連番, Sum(-([農地状況]<=29)*([V_農地].[田面積]+[V_農地].[畑面積]+[V_農地].[樹園地])) AS 式2, Sum(([V_農地].[田面積]+[V_農地].[畑面積]+[V_農地].[樹園地])) AS 式1 FROM V_農地 INNER JOIN [D:世帯Info] ON V_農地.耕作世帯ID = [D:世帯Info].ID GROUP BY [D:世帯Info].ID,[D:世帯Info].世帯主ID,選挙連番 HAVING ((([D:世帯Info].ID)>0));")
            pT1.PrimaryKey = New DataColumn() {pT1.Columns("ID")}

            Dim mvar続柄 As DataTable = .GetTableBySqlSelect("SELECT * FROM [V_続柄]")
            mvar続柄.PrimaryKey = New DataColumn() {mvar続柄.Columns("ID")}

            Dim mvar行政区 As DataTable = .GetTableBySqlSelect("SELECT *,0 As [Used] FROM [V_行政区] ORDER BY [ID]")
            mvar行政区.PrimaryKey = New DataColumn() {mvar行政区.Columns("ID")}

            Dim TB0 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT ID, 名称,' 第' & Right$('0' & nParam,2) & '投票区' AS 投票区,0 AS 世帯数,0 as 男,0 as 女,0 AS 世帯員数計 FROM [M_BASICALL] WHERE [ID]>0 AND [Class]='行政区';")
            TB0.PrimaryKey = New DataColumn() {TB0.Columns("ID")}

            Dim TB1 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT ID,[D:個人Info].行政区ID, [D:個人Info].世帯ID,[氏名],[フリガナ],[D:個人Info].郵便番号, [D:個人Info].性別,[住所],[生年月日],[続柄1] & '-' & [続柄2] & '-' & [続柄3] AS [続柄],Choose([性別]+1,'男','女') AS 性別s FROM [D:個人Info] WHERE [住民区分]=0 AND ((([D:個人Info].選挙権の有無)=True));")

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

            For Each pRow As DataRow In TB1.Rows
                Dim p世帯Row As DataRow = p世帯.Rows.Find(pRow.Item("世帯ID"))

                If Not IsDBNull(pRow.Item("続柄")) Then
                    Dim Ar() As String = pRow.Item("続柄").ToString.Split("-")
                    Dim sBL As New System.Text.StringBuilder

                    For Each Sx As String In Ar
                        Dim pZRow As DataRow = mvar続柄.Rows.Find(Integer.Parse(Sx))
                        If Not Sx = "0" AndAlso pZRow IsNot Nothing Then
                            If sBL.Length = 0 Then
                                sBL.Append(pZRow.Item("名称"))
                            ElseIf pZRow IsNot Nothing Then
                                sBL.Append("の" & pZRow.Item("名称"))
                            End If
                        End If
                    Next
                    If InStr(sBL.ToString, "妻の妻") Then
                        Stop
                    End If
                End If

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
                        p世帯Row.Item("面積要件") = (Val(p世帯面積情報.Item("式1").ToString) >= 1000)
                        p世帯Row.Item("強制要件") = Not p世帯Row.Item("面積要件")
                    Else
                        p世帯Row.Item("世帯主ID") = -1
                        p世帯Row.Item("世帯面積") = 0
                        p世帯Row.Item("耕作面積") = 0
                        p世帯Row.Item("面積要件") = False
                        p世帯Row.Item("強制要件") = True
                    End If

                    p世帯Row.Item("世帯主名") = pRow.Item("氏名")
                    p世帯Row.Item("世帯主フリガナ") = pRow.Item("フリガナ")
                    p世帯Row.Item("世帯住所") = pRow.Item("住所")
                    p世帯Row.Item("世帯郵便番号") = pRow.Item("郵便番号")

                    If pRow.Item("住所").ToString.StartsWith("伊集院町") Then
                        p世帯Row.Item("出力エリア") = "伊集院"
                    ElseIf pRow.Item("住所").ToString.StartsWith("日吉町") Then
                        p世帯Row.Item("出力エリア") = "日吉"
                    ElseIf pRow.Item("住所").ToString.StartsWith("東市来町") Then
                        p世帯Row.Item("出力エリア") = "東市来"
                    ElseIf pRow.Item("住所").ToString.StartsWith("吹上町") Then
                        p世帯Row.Item("出力エリア") = "吹上"
                    End If

                    If Not IsDBNull(pRow.Item("行政区ID")) Then
                        Dim pZRow As DataRow = mvar行政区.Rows.Find(Val(pRow.Item("行政区ID")))
                        If pZRow IsNot Nothing Then
                            p世帯Row.Item("行政区ID") = pRow.Item("行政区ID")
                            p世帯Row.Item("行政区") = pZRow.Item("名称")
                        End If
                    End If

                Else

                End If

                If pRow.Item("続柄") = "世帯主" OrElse (Not IsDBNull(p世帯Row.Item("世帯主ID")) AndAlso p世帯Row.Item("世帯主ID") = pRow.Item("ID")) Then
                    p世帯Row.Item("世帯主名") = pRow.Item("氏名")
                    p世帯Row.Item("世帯主フリガナ") = pRow.Item("フリガナ")
                    p世帯Row.Item("世帯住所") = pRow.Item("住所")
                    p世帯Row.Item("世帯郵便番号") = pRow.Item("郵便番号")

                    If pRow.Item("住所").ToString.StartsWith("伊集院町") Then
                        p世帯Row.Item("出力エリア") = "伊集院"
                    ElseIf pRow.Item("住所").ToString.StartsWith("日吉町") Then
                        p世帯Row.Item("出力エリア") = "日吉"
                    ElseIf pRow.Item("住所").ToString.StartsWith("東市来町") Then
                        p世帯Row.Item("出力エリア") = "東市来"
                    ElseIf pRow.Item("住所").ToString.StartsWith("吹上町") Then
                        p世帯Row.Item("出力エリア") = "吹上"
                    End If

                    If Not IsDBNull(pRow.Item("行政区ID")) Then
                        Dim pZRow As DataRow = mvar行政区.Rows.Find(Val(pRow.Item("行政区ID")))
                        If pZRow IsNot Nothing Then
                            p世帯Row.Item("行政区ID") = pRow.Item("行政区ID")
                            p世帯Row.Item("行政区") = pZRow.Item("名称")
                        End If
                    End If
                End If

                p世帯Row.Item("男") -= (pRow.Item("性別s") = "男")
                p世帯Row.Item("女") -= (pRow.Item("性別s") = "女")

                p世帯Row.Item("人数") = p世帯Row.Item("男") + p世帯Row.Item("女")
                p世帯Row.Item("部分出力") = False
                p世帯Row.Item("世帯員") = p世帯Row.Item("世帯員").ToString &
                                IIf(Len(p世帯Row.Item("世帯員").ToString) > 0, vbCrLf, "") &
                                pRow.Item("氏名") & ";" &
                                pRow.Item("続柄") & ";" &
                                pRow.Item("生年月日") & ";" & pRow.Item("性別s")

            Next

            Dim Lst As New List(Of Integer)
            'For Each p世帯Row As DataRow In p世帯.Rows
            '    If IsDBNull(p世帯Row.Item("世帯主名")) Then
            '        p世帯Row.Item("世帯主名") = "×"
            '        Lst.Add(p世帯Row.Item("ID"))
            '    End If
            'Next

            'For Each nID As Integer In Lst
            '    Dim pRow As DataRow = p世帯.Rows.Find(nID)
            '    p世帯.Rows.Remove(pRow)
            'Next

            Dim pView As New DataView(p世帯, "[強制要件] = True Or [面積要件] = True", "[行政区ID],[世帯主フリガナ]", DataViewRowState.CurrentRows)

            Return pView.ToTable
        End With
    End Function

    Public Overrides Sub InitLocalData()
        With New dlgLoginForm()

            If Not .ShowDialog() = Windows.Forms.DialogResult.OK Then
                Try

                    End
                Catch ex As Exception

                End Try
            Else
                Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [M_BASICALL] WHERE [ID]=5 AND [Class]='農業改善計画認定項目'")
                If pTBL.Rows.Count = 0 Then
                    SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [M_BASICALL]([ID],Class,名称) VALUES(5,'農業改善計画認定項目','認定農業者+農地所有適格法人');")

                    pTBL = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [M_BASICALL] WHERE [ID]=6 AND [Class]='農業改善計画認定項目'")
                    If pTBL.Rows.Count = 0 Then
                        SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [M_BASICALL]([ID],Class,名称) VALUES(6,'農業改善計画認定項目','認定新規就農者');")
                    End If

                    End
                End If

            End If
        End With
    End Sub
    Public Overrides Function Get地区情報(ByVal s住所 As String) As String
        If InStr(s住所, "吹上") > 0 Then
            Return "吹上"
        ElseIf InStr(s住所, "東市来") > 0 Then
            Return "東市来"
        ElseIf InStr(s住所, "日吉") > 0 Then
            Return "日吉"
        Else
            Return "伊集院"
        End If

        Return MyBase.Get地区情報(s住所)
    End Function


End Class
