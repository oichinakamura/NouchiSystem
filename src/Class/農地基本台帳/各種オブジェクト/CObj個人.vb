Imports HimTools2012.CommonFunc


Public Class CObj個人 : Inherits CTargetObjWithView農地台帳

    Public Sub New(ByVal pRow As DataRow, ByVal bAddNew As Boolean)
        MyBase.New(pRow, bAddNew, New HimTools2012.TargetSystem.DataKey("個人", pRow.Item("ID")), "D:個人Info")
        Try
            If IsDBNull(pRow.Item("続柄1")) Then
                pRow.Item("続柄1") = 0
            End If
        Catch ex As Exception

        End Try
    End Sub
    Public Overrides ReadOnly Property DataTableWithUpdateList As HimTools2012.Data.DataTableWithUpdateList
        Get
            Return App農地基本台帳.TBL個人
        End Get
    End Property

#Region "プロパティ"

    Public Property 世帯ID() As System.Int64
        Get
            Return GetDecimalValue("世帯ID")
        End Get
        Set(ByVal value As System.Int64)
            ValueChange("世帯ID", value)
        End Set
    End Property
    Public Property フリガナ() As String
        Get
            Return GetStringValue("フリガナ")
        End Get
        Set(ByVal value As String)
            ValueChange("フリガナ", value)
        End Set
    End Property
    Public Property 氏名() As String
        Get
            Return GetStringValue("氏名")
        End Get
        Set(ByVal value As String)
            ValueChange("氏名", value)
        End Set
    End Property
    Public Property 住所() As String
        Get
            Return GetStringValue("住所")
        End Get
        Set(ByVal value As String)
            ValueChange("住所", value)
        End Set
    End Property

    Public ReadOnly Property 市町村CD As String
        Get
            Dim cityCodeModel As CitiesCode.Interface.ICityCodeModel = New CitiesCode.Factory.CityCodeFactory().CreateCityCodeModel("csvパス")

            Dim cityCode As String = ""
            Dim otherJusyo As String = ""
            Dim kenCity As String = ""

            Dim jusyoModel As CitiesCode.Interface.IJusyoModel = cityCodeModel.GetCityCode("鹿児島県鹿児島市松原町7-6")  ' 文字列より市町村コード取得
            If jusyoModel.MatchState = CitiesCode.Types.MatchType.Match Then

                ' Match以外はnull
                cityCode = jusyoModel.CityCode
                otherJusyo = jusyoModel.OtherJusyoText  ' その他の住所
                kenCity = jusyoModel.JusyoText.Replace(jusyoModel.OtherJusyoText, "")  ' その他の住所以外
                Return cityCode
            End If

            Return ""
        End Get
    End Property


    Public Shared Function Get続柄(ByVal pRow As DataRow) As String
        Dim s続柄 As String = ""

        Dim pRowV1 As String = App農地基本台帳.DataMaster.GetValue("続柄", Val(pRow.Item("続柄1").ToString))
        Dim pRowV2 As String = App農地基本台帳.DataMaster.GetValue("続柄", Val(pRow.Item("続柄2").ToString))
        Dim pRowV3 As String = App農地基本台帳.DataMaster.GetValue("続柄", Val(pRow.Item("続柄3").ToString))
        Return (pRowV1 & "の" & pRowV2 & "の" & pRowV3).ToString.Replace("の-", "")
    End Function


    Public Property 農業改善計画認定() As enum農業改善計画認定
        Get
            Return GetIntegerValue("農業改善計画認定")
        End Get
        Set(ByVal Value As enum農業改善計画認定)
            ValueChange("農業改善計画認定", Value)
        End Set
    End Property


    Public Property 更新日() As DateTime
        Get
            Return GetDateValue("更新日")
        End Get
        Set(ByVal value As DateTime)
            ValueChange("更新日", value)
        End Set
    End Property
#End Region

    Public Overrides Sub ValueChange(ByVal sField As String, ByVal pValue As Object)
        App農地基本台帳.TBL個人.AddUpdateListwithDataViewPage(Me, sField, pValue)
    End Sub

    Public Overrides Function GetContextMenu(Optional ByVal pMenu As HimTools2012.controls.MenuItemEx = Nothing, Optional nDips As Integer = 1, Optional sParam() As String = Nothing) As HimTools2012.controls.MenuPlus
        Dim pMenuPlus As HimTools2012.controls.MenuPlus = CreateMenu(pMenu)

        Dim bEdit As Boolean = (SysAD.SystemInfo.ユーザー.n権利 > 0)
        With pMenuPlus
            If pMenu IsNot Nothing Then .AddMenu("<<" & Me.氏名 & ">>", , , , True, , Color.White, Color.LightSlateGray)
            .AddMenu("開く", , AddressOf ClickMenu)
            .InsertSeparator()

            Dim n世帯 As Decimal = GetDecimalValue("世帯ID")
            If Not n世帯 = 0 Then
                Dim p世帯 As DataRow = App農地基本台帳.TBL世帯.FindRowByID(n世帯)

                If p世帯 Is Nothing Then
                    .AddMenu("世帯の作成", , , bEdit)
                Else
                    .AddSubMenu("世帯", nDips, ObjectMan.GetObject("農家." & Me.GetDecimalValue("世帯ID")), My.Resources.Resource1.農家世帯.ToBitmap)
                End If
            Else
                .AddMenu("世帯の作成", , , bEdit)
            End If

            .AddMenu("世帯番号の修正", , , bEdit)
            .AddMenu("基本台帳印刷")
            .AddMenu("基本台帳印刷(簡易)")
            .AddMenu("基本台帳印刷(編集モード)")

            .AddMenuByText({"耕作証明願", "耕作面積証明書", "-"}, AddressOf ClickMenu, True)
            .AddMenu("総括表")
            .InsertSeparator()
            With .AddMenu("農地一覧")
                .AddMenu("所有農地", , AddressOf ClickMenu)
                .AddMenu("登記名義農地", , AddressOf ClickMenu)
                .InsertSeparator()
                .AddMenu("経営農地", "経営農地の一覧", AddressOf ClickMenu)
                .AddMenu("自作農地", "自作農地の一覧", AddressOf ClickMenu)
                .AddMenu("借受農地", "借受農地の一覧", AddressOf ClickMenu)
                .InsertSeparator()
                .AddMenu("貸付農地", "貸出農地の一覧", AddressOf ClickMenu)
                '.AddMenu("転用農地", AddressOf sub転用農地一覧)
                .InsertSeparator()
                .AddMenu("管理農地", "管理農地の一覧", AddressOf ClickMenu)
            End With

            If SysAD.MapConnection.HasMap Then
                .InsertSeparator()
                With .AddMenu("地図を呼ぶ")
                    .AddMenu("経営農地", , AddressOf sub経営農地地図一覧)
                    .AddMenu("所有農地", , AddressOf sub所有農地地図一覧)
                    .AddMenu("借受農地", , AddressOf sub借受農地地図一覧)
                    .InsertSeparator()
                    .AddMenu("貸付農地", , AddressOf sub貸付農地地図一覧)
                End With
            End If
            With .AddMenu("設定")
                .AddMenu("認定農業者に設定", , AddressOf ClickMenu, , bEdit)
                .AddMenu("担い手農家に設定", , AddressOf ClickMenu, , bEdit)
                .AddMenu("農業生産法人に設定", , AddressOf ClickMenu, , bEdit)
                .AddMenu("認定農業者＋担い手農家に設定", , AddressOf ClickMenu, , bEdit)
            End With
            'St = St & IIf(DVProperty.Rs.Value("選挙権の有無"), "^", "") & n & "選挙権の有無;"
            ' & IIf(DVProperty.Rs.Value("農業改善計画認定") = 1, "認定情報の呼出;", "") & "-;" & n &  & IIf(DVProperty.Rs.Value("合併異動"), "", "~") & n & "<;"
            'St = St & ">所有地の移動;" & n & "農地の所有権移動－３条;" & n & "農地の貸借－３条;" & n & "農地の転用－４条;" & n & "転用目的の所有権移動－５条;" & n & "所有権移転;" & n & "利用権設定;" & n & "利用権移転;" & n & "解約;<;"

            .AddMenu("関連申請一覧", , AddressOf ClickMenu)
            .AddMenu("農地の追加", , AddressOf ClickMenu, , bEdit)
            .AddMenu("あっせん申出", , AddressOf ClickMenu, , bEdit)


        End With

        Return GetCommonMenu(pMenuPlus, pMenu, bEdit)
    End Function

    Private Sub Send農地ListToMap(ByVal pTable As DataTable)
        If pTable.Rows.Count > 0 Then
            Dim sB As New System.Text.StringBuilder

            sB.AppendLine("LogicMode:2")
            sB.AppendLine("PaintMode:1")

            For Each pRow As DataRow In pTable.Rows
                sB.AppendLine("LotIDP:" & pRow.Item("ID").ToString & ",1")
            Next

            SysAD.MapConnection.SelectMap(sB.ToString)
        Else
            MsgBox("該当する農地が見つかりませんでした", MsgBoxStyle.Critical)
        End If
    End Sub

    Public Sub sub経営農地地図一覧()
        Dim pTable As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(String.Format("SELECT [ID] FROM [D:農地Info] WHERE ([自小作別]=0 AND [所有者ID]={0}) Or ([自小作別]<>0 AND [借受人ID]={0})", GetLongIntValue(("ID"))))
        Send農地ListToMap(pTable)
    End Sub
    Public Sub sub所有農地地図一覧()
        Dim pTable As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(String.Format("SELECT [ID] FROM [D:農地Info] WHERE [所有者ID]={0}", GetLongIntValue("ID")))
        Send農地ListToMap(pTable)
    End Sub
    Public Sub sub自作農地地図一覧()
        Dim pTable As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(String.Format("SELECT [ID] FROM [D:農地Info] WHERE [自小作別]=0 AND [所有者ID]={0}", GetLongIntValue("ID")))
        Send農地ListToMap(pTable)
    End Sub
    Public Sub sub借受農地地図一覧()
        Dim pTable As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(String.Format("SELECT [ID] FROM [D:農地Info] WHERE [自小作別]<>0 AND [借受人ID]={0}", GetLongIntValue("ID")))
        Send農地ListToMap(pTable)
    End Sub
    Public Sub sub貸付農地地図一覧()
        Dim pTable As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(String.Format("SELECT [ID] FROM [D:農地Info] WHERE ([自小作別]<>0 AND [所有者ID]={0} AND [借受人ID]<>{0})", GetLongIntValue("ID")))
        Send農地ListToMap(pTable)
    End Sub

    Public Function GetDBString(ByVal pTable As DataTable, ByVal sColumnDelimiter As String, ByVal sRowDelimiter As String, ByVal ParamArray Columns() As String) As String
        Dim pList As New System.Text.StringBuilder
        For Each pRow As DataRow In pTable.Rows
            If pList.Length > 0 Then pList.Append(sRowDelimiter)

            For Each sColumn As String In Columns
                pList.Append(IIF(pList.Length > 0, sColumnDelimiter, "") & pRow.Item(sColumn))
            Next
        Next
        Return pList.ToString
    End Function

    Public Overrides Function ToString() As String
        Return Me.氏名
    End Function


    Public Overrides Function GetProperty(ByVal sParam As String) As Object
        Try
            Select Case sParam
                Case "世帯の有無" : Return (Not Val(Row.Body.Item("世帯ID").ToString) = 0)
                Case "経営面積"
                    If Not IsDBNull(mvarRow.Item("世帯ID")) AndAlso mvarRow.Item("世帯ID") <> 0 Then Return Me.GetProperty("世帯経営面積") Else Return Me.GetProperty("個人経営面積")
                Case "借入面積"
                    If Not Row.IsZero("世帯ID") Then
                        Return GetProperty("世帯借入面積")
                    Else
                        Return GetProperty("個人借入面積")
                    End If
                Case "世帯経営面積"
                    If Not Row.IsZero("世帯ID") Then
                        Dim p農地TBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE [所有世帯ID]={0} Or ([自小作別]>0 AND [借受世帯ID]={0}) Or ([所有者ID]={1})", Me.世帯ID, Me.ID)
                        App農地基本台帳.TBL農地.MergePlus(p農地TBL)

                        App農地基本台帳.TBL農地.FindRowBySQL(String.Format("[所有世帯ID]={0} Or ([自小作別]>0 AND [借受世帯ID]={0}) Or ([所有者ID]={1})", Me.世帯ID, Me.ID))
                        Dim query = From cust In App農地基本台帳.TBL農地.Body Where
                                    (Val(cust.Item("自小作別").ToString) = 0 AndAlso Val(cust.Item("所有世帯ID").ToString) = Me.世帯ID) OrElse
                                    (Val(cust.Item("自小作別").ToString) > 0 AndAlso Val(cust.Item("借受世帯ID").ToString) = Me.世帯ID)

                        Dim pSum = (From q In query Select Val(q.Item("田面積")) + Val(q.Item("畑面積")) + Val(q.Item("樹園地")) + Val(q.Item("採草放牧面積"))).Sum
                        Return pSum
                    Else
                        Return Me.GetProperty("個人経営面積")
                    End If
                Case "世帯借入面積"
                    If Not Row.IsZero("世帯ID") Then
                        Dim pTBL世帯借入面積 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Sum([V_農地].田面積+[V_農地].畑面積) AS 耕作面積 FROM [V_農地] WHERE [V_農地].所有世帯ID<>" & mvarRow.Item("世帯ID") & " AND [V_農地].借受世帯ID=" & mvarRow.Item("世帯ID") & " AND [V_農地].自小作別>0")
                        If pTBL世帯借入面積.Rows.Count > 0 Then Return pTBL世帯借入面積.Rows(0).Item("耕作面積")
                    End If
                    Return 0
                Case "個人経営面積"
                    Dim p農地TBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE [所有者ID]={0} Or ([自小作別]>0 AND [借受人ID]={0})", Me.ID)
                    App農地基本台帳.TBL農地.MergePlus(p農地TBL)

                    App農地基本台帳.TBL農地.FindRowBySQL(String.Format("[所有者ID]={0} Or ([自小作別]>0 AND [借受人ID]={0})", Me.ID))
                    Dim query = From cust In App農地基本台帳.TBL農地.Body Where
                                (Val(cust.Item("自小作別").ToString) = 0 AndAlso Val(cust.Item("所有者ID").ToString) = Me.世帯ID) OrElse
                                (Val(cust.Item("自小作別").ToString) > 0 AndAlso Val(cust.Item("借受人ID").ToString) = Me.世帯ID)

                    Dim pSum = (From q In query Select Val(q.Item("田面積")) + Val(q.Item("畑面積")) + Val(q.Item("樹園地")) + Val(q.Item("採草放牧面積"))).Sum
                    Return pSum
                Case "個人借入面積"
                    Dim pTBL個人借入面積 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Sum([V_農地].田面積+[V_農地].畑面積) AS 耕作面積 FROM [V_農地] WHERE [借受人ID]=" & Me.ID & " AND [V_農地].自小作別>0")
                    If pTBL個人借入面積.Rows.Count > 0 Then
                        Return Val(pTBL個人借入面積.Rows(0).Item("耕作面積").ToString)
                    Else
                        Return 0
                    End If
                Case "集落名" : Return mvarRow.Item("行政区名").ToString()
                Case "年齢" : Return HimTools2012.DateFunctions.年齢(mvarRow.Item("生年月日"), Now)
                Case "世帯員数"
                    If Not Row.IsZero("世帯ID") Then
                        Dim pRows() As DataRow = App農地基本台帳.TBL個人.FindRowBySQL("[世帯ID]=" & Row.Body.Item("世帯ID"))
                        Return pRows.Count
                    Else
                        Return 0
                    End If
                Case Else
                    Return mvarRow.Item(sParam)
            End Select
        Catch ex As Exception

        End Try

        Return ""
    End Function

    Public Overrides Function DoCommand(ByVal sCommand As String, ByVal ParamArray sParams() As String) As Object
        Select Case sCommand
            Case "開く"
                Return Me.OpenDataViewNext(SysAD.page農家世帯.DataViewCollection, ObjectMan)
            Case "所有農地" : SysAD.page農家世帯.農地リスト.検索開始(String.Format("[管理者ID]={0} Or [所有者ID]={0}", mvarRow.Item("ID")), String.Format("[管理者ID]={0} Or [所有者ID]={0}", mvarRow.Item("ID")))
            Case "登記名義農地" : SysAD.page農家世帯.農地リスト.検索開始(String.Format("[登記名義人ID]={0}", mvarRow.Item("ID")), String.Format("[登記名義人ID]={0}", mvarRow.Item("ID")))
            Case "耕作証明願", "耕作証明多筆型", "耕作証明多筆型世帯" : mod農地基本台帳.耕作多筆証明印刷(Me.Key.KeyValue)
            Case "耕作証明書世帯", "耕作面積証明書", "耕作面積証明印刷"
                mod農地基本台帳.耕作面積証明印刷(Me.Key.KeyValue)
            Case "基本台帳印刷" : mod農地基本台帳.基本台帳印刷(Me.Key.KeyValue, ExcelViewMode.Preview, 印刷Mode.フル印刷)
            Case "基本台帳印刷(簡易)" : mod農地基本台帳.基本台帳印刷(Me.Key.KeyValue, ExcelViewMode.Preview, 印刷Mode.簡易印刷)
            Case "基本台帳印刷(編集モード)" : mod農地基本台帳.基本台帳印刷(Me.Key.KeyValue, ExcelViewMode.EditMode, 印刷Mode.フル印刷)
            Case "関連申請一覧" : C各種申請管理.Open申請List("個人関連申請一覧." & Me.Key.ID, "関連申請一覧:" & Me.氏名, "[申請者A]=" & Me.ID & " OR [申請者B]=" & Me.ID)
            Case "データを削除" : 個人削除(Me)
            Case "認定農業者に設定" : ValueChange("農業改善計画認定", 1) : Me.SaveMyself()
            Case "担い手農家に設定" : ValueChange("農業改善計画認定", 2) : Me.SaveMyself()
            Case "農業生産法人に設定" : ValueChange("農業改善計画認定", 3) : Me.SaveMyself()
            Case "認定農業者＋担い手農家に設定" : ValueChange("農業改善計画認定", 4) : Me.SaveMyself()

            Case "世帯を呼ぶ", "世帯呼出" : Open世帯(Me.世帯ID, "指定された世帯は見つかりませんでした。")
            Case "世帯の作成" : mod個人.世帯追加(Me)
            Case "農地の追加" : Sub農地追加(Me, Nothing)
            Case "経営農地の一覧", "経営農地"
                Dim sWhere As String = String.Format("([自小作別]=0 AND [所有者ID]={0}) Or ([自小作別]<>0 AND [借受人ID]={0})", mvarRow.Item("ID"))
                Dim sVWhere As String = String.Format("([自小作別]=0 AND [所有者ID]={0}) Or ([自小作別]<>0 AND [借受人ID]={0})", mvarRow.Item("ID"))
                SysAD.page農家世帯.農地リスト.検索開始(sWhere, sVWhere)
            Case "自作農地", "自作農地の一覧"
                Dim sWhere As String = String.Format("[自小作別]=0 AND [所有者ID]={0}", mvarRow.Item("ID"))
                Dim sVWhere As String = String.Format("[自小作別]=0 AND [所有者ID]={0}", mvarRow.Item("ID"))
                SysAD.page農家世帯.農地リスト.検索開始(sWhere, sVWhere)
            Case "所有農地", "所有農地の一覧", "所有地一覧"
                Dim sWhere As String = String.Format("[所有者ID]={0}", mvarRow.Item("ID"))
                Dim sVWhere As String = String.Format("[所有者ID]={0}", mvarRow.Item("ID"))
                SysAD.page農家世帯.農地リスト.検索開始(sWhere, sVWhere)
            Case "借受農地の一覧", "借受農地"
                Dim sWhere As String = String.Format("[自小作別]<>0 AND [借受人ID]={0}", mvarRow.Item("ID"))
                Dim sVWhere As String = String.Format("[自小作別]<>0 AND [借受人ID]={0}", mvarRow.Item("ID"))
                SysAD.page農家世帯.農地リスト.検索開始(sWhere, sVWhere)
            Case "貸出農地の一覧", "貸付農地"
                Dim sWhere As String = String.Format("[自小作別]<>0 AND [所有者ID]={0} AND [借受人ID]<>{0}", mvarRow.Item("ID"))
                Dim sVWhere As String = String.Format("[自小作別]<>0 AND [所有者ID]={0} AND [借受人ID]<>{0}", mvarRow.Item("ID"))
                SysAD.page農家世帯.農地リスト.検索開始(sWhere, sVWhere)
            Case "管理農地の一覧", "管理農地"
                Dim sWhere As String = String.Format("[管理者ID]={0}", mvarRow.Item("ID"))
                Dim sVWhere As String = String.Format("[管理者ID]={0}", mvarRow.Item("ID"))
                SysAD.page農家世帯.農地リスト.検索開始(sWhere, sVWhere)
            Case "あっせん申出" : C申請データ作成.あっせん申出受(Me)
            Case "総括表" : 総括表()
            Case "農地一覧"

            Case "世帯番号の修正"
                Dim nTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Min([ID]) AS [最小値] FROM [D:世帯Info]")
                Dim nID As Integer = nTBL.Rows(0).Item("最小値") - 1
                Dim sNo As String = InputBox("世帯番号を入力してください", "世帯番号の設定", nID)
                'If Val(sNo) > 0 Then
                '    SetItem("世帯番号", Val(sNo))
                '    SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:個人Info] SET [世帯ID]={0} WHERE [ID]=" & Me.ID, sNo)
                'End If
                'SetItem("世帯番号", Val(sNo))
                App農地基本台帳.TBL個人.AddUpdateListwithDataViewPage(Me, "世帯番号", Val(sNo))
                SaveMyself()
                SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:個人Info] SET [世帯ID]={0} WHERE [ID]=" & Me.ID, sNo)

                'MsgBox("現在この機能は使えません", vbCritical)
            Case Else
                Return MyBase.DoCommand(sCommand, sParams)
        End Select
        Return ""
    End Function

    Public Sub 総括表()
        If Not SysAD.page農家世帯.TabPageContainKey("総括表" & Me.Key.KeyValue) Then
            SysAD.page農家世帯.中央Tab.AddPage(New CTabPage総括表(Me))
        End If
    End Sub

    Public Overrides Sub DropDown(ByVal sSourceList As String, Optional ByVal sOption As String = "")
        Select Case GetKeyHead(sSourceList)
            Case "個人"
                If Split(sSourceList, ";").Length > 1 Then
                    MsgBox("複数人のドラッグは処理できません。", MsgBoxStyle.Critical)
                ElseIf MsgBox("同定処理しますか", vbYesNo) = vbYes Then
                    Dim nID As Long = GetKeyCode(sSourceList)
                    SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [D:農地Info].所有者ID = {0} WHERE ((([D:農地Info].所有者ID)={1}));", Me.Key.ID, nID)
                    SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [D:農地Info].管理者ID = {0} WHERE ((([D:農地Info].管理者ID)={1}));", Me.Key.ID, nID)
                    SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [D:農地Info].借受人ID = {0} WHERE ((([D:農地Info].借受人ID)={1}));", Me.Key.ID, nID)
                    SysAD.DB(sLRDB).ExecuteSQL("UPDATE D_申請 SET D_申請.申請者A = {0} WHERE (((D_申請.申請者A)={1}));", Me.Key.ID, nID)
                    SysAD.DB(sLRDB).ExecuteSQL("UPDATE D_申請 SET D_申請.申請者B = {0} WHERE (((D_申請.申請者B)={1}));", Me.Key.ID, nID)
                    SysAD.DB(sLRDB).ExecuteSQL("UPDATE D_申請 SET D_申請.申請者C = {0} WHERE (((D_申請.申請者C)={1}));", Me.Key.ID, nID)
                    SysAD.DB(sLRDB).ExecuteSQL("UPDATE D_土地履歴 SET D_土地履歴.関係者A = {0} WHERE (((D_土地履歴.関係者A)={1}));", Me.Key.ID, nID)
                    SysAD.DB(sLRDB).ExecuteSQL("UPDATE D_土地履歴 SET D_土地履歴.関係者B = {0} WHERE (((D_土地履歴.関係者B)={1}));", Me.Key.ID, nID)
                    SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] INNER JOIN [D:個人Info] ON [D:農地Info].所有者ID = [D:個人Info].ID SET [D:農地Info].所有世帯ID = [世帯ID] WHERE ((([D:農地Info].所有世帯ID)<>[世帯ID]));")
                    SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] INNER JOIN [D:個人Info] ON [D:農地Info].管理者ID = [D:個人Info].ID SET [D:農地Info].管理世帯ID = [世帯ID] WHERE ((([D:農地Info].管理世帯ID)<>[世帯ID]));")
                    SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] INNER JOIN [D:個人Info] ON [D:農地Info].借受人ID = [D:個人Info].ID SET [D:農地Info].借受世帯ID = [世帯ID] WHERE ((([D:農地Info].借受世帯ID)<>[世帯ID]));")
                End If
            Case "農地" : mod申請データ作成処理.農地申請データ作成処理(sSourceList, Me)
            Case "転用農地"
                mod申請データ作成処理.転用農地申請データ作成処理(sSourceList, Me)
            Case Else
        End Select
    End Sub

    Public Overrides Function CanDropKeyHead(ByVal sKey As String, ByVal sOption As String) As Boolean
        Select Case GetKeyHead(sKey)
            Case "個人" : Return True
            Case "農地" : Return True
            Case "転用農地" : Return True
            Case Else
                If Not SysAD.IsClickOnceDeployed Then
                    Debug.Print(sKey)
                    Stop
                End If
                Return False
        End Select
    End Function

    Public Overrides Function SaveMyself() As Boolean

        Return MyBase.SaveBase("D:個人Info")
    End Function
End Class
