
Imports HimTools2012.CommonFunc
Imports HimTools2012.controls
Imports HimTools2012.controls.DVCtrlCommonBase
#Region "宣言"

Public Enum Enum兼業区分
    不明 = 0
    専業 = 1
    兼業_農業所得多い = 2
    兼業_農業以外所得多い = 3
    農業生産法人 = 4
    非農家 = 5
End Enum
Public Enum Enum希望有無
    不明 = 0
    希望する = 1
    希望しない = 2
End Enum
Public Enum Enum構成員有無
    不明 = 0
    構成員 = 1
    非構成員 = 2
End Enum
Public Enum Enum有無様式1
    不明 = 0
    有 = 1
    無 = 2
End Enum
#End Region

Public Class CObj農家 : Inherits CTargetObjWithView農地台帳
    Public Enum Enumあっせん希望種別
        なし = 0
        農業委員会あっせん = 1
        管理機構あっせん = 2
        農委機構双方 = 3
    End Enum
    Public Sub New(ByVal pRow As DataRow, ByVal bAddNew As Boolean)
        MyBase.New(pRow, bAddNew, New HimTools2012.TargetSystem.DataKey("農家", pRow.Item("ID")), "D:世帯Info")
    End Sub

    Public Overrides Function GetProperty(ByVal sParam As String) As Object
        Select Case sParam
            Case "氏名" : Return GetItem("世帯主氏名", "")
            Case "住所" : Return GetItem("住所", "")
            Case "世帯主KEY" : Return "個人." & GetItem("世帯主ID", 0)
            Case "世帯主名", "世帯主氏名" : Return GetItem("世帯主氏名", "")
            Case "世帯主ID" : Return GetItem("世帯主ID", 0)
            Case "世帯主"
                If Val(Me.Row.Body.Item("世帯主ID").ToString) > 0 Then
                    Dim pRow As DataRow = App農地基本台帳.TBL個人.Rows.Find(Val(Me.Row.Body.Item("世帯主ID").ToString))
                    If pRow IsNot Nothing Then
                        Return ObjectMan.GetObjectDB("個人." & pRow.Item("ID"), pRow, GetType(CObj個人), False)
                    Else
                        Return Nothing
                    End If
                Else
                    Return Nothing
                End If
            Case "世帯主集落"
                If Val(Me.Row.Body.Item("世帯主ID").ToString) > 0 Then
                    Dim pRow As DataRow = App農地基本台帳.TBL個人.Rows.Find(Val(Me.Row.Body.Item("世帯主ID").ToString))
                    Dim p世帯主 As New CObj個人(pRow, False)
                    Return p世帯主.Row.Body.Item("行政区名")
                Else
                    Return ""
                End If
            Case "世帯主職業"
                If Val(Me.Row.Body.Item("世帯主ID").ToString) > 0 Then
                    Dim pRow As DataRow = App農地基本台帳.TBL個人.Rows.Find(Val(Me.Row.Body.Item("世帯主ID").ToString))
                    Dim p世帯主 As New CObj個人(pRow, False)
                    Return p世帯主.Row.Body.Item("職業")
                Else
                    Return ""
                End If
            Case "住民区分"
                Dim pRow As DataRow = App農地基本台帳.TBL個人.FindRowByID(Me.世帯主ID)
                If pRow IsNot Nothing Then
                    Return Val(pRow.Item("住民区分"))
                End If

            Case "経営面積"
                App農地基本台帳.TBL農地.FindRowBySQL(String.Format("[所有世帯ID]={0} Or ([自小作別]<>0 AND [借受世帯ID]={0})", Me.ID))
                Dim query = From cust In App農地基本台帳.TBL農地.Body Where
                            (Val(cust.Item("自小作別").ToString) = 0 AndAlso Val(cust.Item("所有世帯ID").ToString) = Me.ID) OrElse
                            (Val(cust.Item("自小作別").ToString) > 0 AndAlso Val(cust.Item("借受世帯ID").ToString) = Me.ID)

                Dim pSum = (From q In query Select Val(q.Item("田面積")) + Val(q.Item("畑面積"))).Sum
                Return pSum
                'Case "借入面積"
                '    Rs = SysAD.DB(sLRDB).GetRecordsetEx("SELECT Sum([V_農地].田面積+[V_農地].畑面積) AS 耕作面積 FROM [V_農地] WHERE [V_農地].所有世帯ID<>" & DVProperty.ID & " AND [V_農地].借受世帯ID=" & DVProperty.ID & " AND [V_農地].自小作別>0", 0, , Me)
                '    If Not Rs Is Nothing Then CDataviewSK_GetProperty2 = Rs.Value("耕作面積") : SysAD.DB(sLRDB).CloseRs(Rs)
            Case "世帯員数"
                If Me.ID <> 0 Then
                    Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:個人Info] WHERE [D:個人Info].世帯ID=" & Me.ID)
                    App農地基本台帳.TBL個人.MergePlus(pTBL)
                    Return pTBL.Rows.Count
                End If
            Case "年齢"
                Dim pRow As DataRow = App農地基本台帳.TBL個人.FindRowByID(Me.世帯主ID)
                If pRow IsNot Nothing Then
                    If Not IsDBNull(pRow.Item("生年月日")) AndAlso pRow.Item("生年月日") > #1/1/1901# Then
                        Return HimTools2012.DateFunctions.年齢(pRow.Item("生年月日"))
                    End If
                End If
            Case "集落名"
                Dim pRow As DataRow = App農地基本台帳.TBL個人.FindRowByID(Me.世帯主ID)
                If pRow IsNot Nothing Then
                    With New CObj個人(pRow, False)
                        Return .GetStringValue("行政区名")
                    End With
                End If
            Case Else
                Try
                    Return mvarRow.Item(sParam).ToString
                Catch ex As Exception
                    Return ""
                End Try
        End Select

        Return ""
    End Function



#Region "プロパティ"

    Public Property 農家番号() As Integer
        Get
            Return GetIntegerValue("農家番号")
        End Get
        Set(ByVal value As Integer)
            ValueChange("農家番号", value)
        End Set
    End Property
    Public ReadOnly Property 世帯主ID() As Long
        Get
            Return GetLongIntValue("世帯主ID")
        End Get
    End Property

    Public Property 更新日() As Date
        Get
            Return GetDateValue("更新日")
        End Get
        Set(ByVal value As Date)
            ValueChange("更新日", value)
        End Set
    End Property


    Public ReadOnly Property 世帯主() As CObj個人
        Get
            Return ObjectMan.GetObject("個人." & Me.世帯主ID)
        End Get
    End Property
#End Region

    Public Overrides Function GetContextMenu(Optional ByVal pMenu As HimTools2012.controls.MenuItemEx = Nothing, Optional nDips As Integer = 1, Optional sParam() As String = Nothing) As HimTools2012.controls.MenuPlus
        Dim pMenuPlus As HimTools2012.controls.MenuPlus = CreateMenu(pMenu)
        Dim bEdit As Boolean = (SysAD.SystemInfo.ユーザー.n権利 > 0)
        With pMenuPlus
            .AddMenu("開く", , AddressOf ClickMenu)
            .InsertSeparator()
            .AddSubMenu("世帯主", nDips, ObjectMan.GetObject("個人." & Me.GetDecimalValue("世帯主ID")), My.Resources.Resource1.個人.ToBitmap)
            .AddMenuByText({"営農情報の表示", "基本台帳印刷", "基本台帳印刷(簡易)", "基本台帳印刷(編集モード)", "-"}, AddressOf ClickMenu, True)
            .AddMenuByText({"耕作証明願", "耕作面積証明書", "-"}, AddressOf ClickMenu, True)
            .AddMenu("営農計画書", , AddressOf ClickMenu)
            .AddMenu("農業生産法人報告書", , AddressOf ClickMenu)

            .AddMenu("農地等の利用状況報告書", , AddressOf ClickMenu, , False)
            .AddMenu("違反転用事案報告書", , AddressOf ClickMenu, , False)
            .AddMenu("遊休農地勧告書", , AddressOf ClickMenu, , False)
            .AddMenu("非農地証明書", , AddressOf ClickMenu, , False)
            .AddMenu("非農地通知書", , AddressOf ClickMenu, , False)
            .AddMenu("適格要件届出書", , AddressOf ClickMenu, , False)
            .AddMenu("農地転用許可後の工事進捗状況報告書", , AddressOf ClickMenu, , False)
            .InsertSeparator()
            .AddMenu("総括表", , AddressOf ClickMenu)
            .InsertSeparator()
            .AddMenu("家族一覧", , AddressOf ClickMenu)
            .AddMenu("家族の追加", , AddressOf ClickMenu, , bEdit)
            .AddMenu("農家番号の入力", , AddressOf ClickMenu, , bEdit)

            '        St = St & "営農情報;-;" & n & "非農地証明;" & _
            .InsertSeparator()
            With .AddMenu("農地一覧")
                .AddMenu("所有農地", , AddressOf ClickMenu)

                .AddMenu("経営農地", , AddressOf ClickMenu)
                .AddMenu("自作農地", , AddressOf ClickMenu)
                .AddMenu("借受農地", , AddressOf ClickMenu)
                .InsertSeparator()
                .AddMenu("貸付農地", , AddressOf ClickMenu)
                .AddMenu("転用農地", , AddressOf ClickMenu)
                .InsertSeparator()
                .AddMenu("管理農地", , AddressOf ClickMenu)
            End With

            If SysAD.MapConnection.HasMap Then
                .InsertSeparator()
                With .AddMenu("地図を呼ぶ")
                    .AddMenu("経営農地", , AddressOf ClickMenu)
                    .AddMenu("所有農地", , AddressOf ClickMenu)
                    .AddMenu("借受農地", , AddressOf ClickMenu)
                    .InsertSeparator()
                    .AddMenu("貸付農地", , AddressOf ClickMenu)
                End With
            End If

            '        St = St & ">設定;" & n & "家族農地の連動;" & n & "合併世帯の解除;<"
            '        St = St & "-;" & IIf(DVProperty.Rs.Value("農地との関連"), "^", "") & n & "農地との関連"
            .InsertSeparator()

            .AddMenu("農地追加", , AddressOf ClickMenu, , bEdit)
            .AddMenu("関連申請", , AddressOf ClickMenu)
            .AddMenu("あっせん申出", , AddressOf ClickMenu, , bEdit)

            .InsertSeparator()
            Dim getList = Get人農地()

            With .AddMenu("人農地プラン中心経営体内訳")
                .AddMenu("設定無", , AddressOf ClickMenu, getList(0))
                .AddMenu("中心経営体", , AddressOf ClickMenu, getList(1))
                .AddMenu("中心経営体ではない", , AddressOf ClickMenu, getList(2))
                .AddMenu("調査中", , AddressOf ClickMenu, getList(3))
            End With

            SetDVMenu(pMenuPlus, pMenu)
        End With
        Return pMenuPlus
    End Function


    Private Function Get人農地() As List(Of Integer)
        Dim _mvarlist = New List(Of Integer)

        Select Case Me.GetDecimalValue("人農地プラン中心経営体区分")
            Case 0 : _mvarlist.AddRange(New Integer() {1, 0, 0, 0})
            Case 1 : _mvarlist.AddRange(New Integer() {0, 1, 0, 0})
            Case 2 : _mvarlist.AddRange(New Integer() {0, 0, 1, 0})
            Case 3 : _mvarlist.AddRange(New Integer() {0, 0, 0, 1})
        End Select

        Return _mvarlist
    End Function

    Public Overrides Function ToString() As String
        Return mvarRow.Item("世帯主氏名").ToString
    End Function

    Public Overrides Function DoCommand(ByVal sCommand As String, ByVal ParamArray sParams() As String) As Object
        Select Case sCommand
            Case "世帯主を呼ぶ" : Open個人(mvarRow.Item("世帯主ID"), "")
            Case "営農情報の表示" : 営農呼出し(mvarRow.Item("ID"), "指定された所有者が見つかりませんでした。")
            Case "耕作証明願" : mod農地基本台帳.耕作多筆証明印刷(Me.Key.KeyValue)
            Case "耕作面積証明書", "耕作面積証明印刷"
                mod農地基本台帳.耕作面積証明印刷(Me.Key.KeyValue)
            Case "開く" : Me.OpenDataViewNext(SysAD.page農家世帯.DataViewCollection, ObjectMan)
            Case "農地追加" : Sub農地追加(Nothing, Me)
            Case "農地との関連"
                Dim mvarUpdateRow As New HimTools2012.Data.UpdateRow(mvarRow, HimTools2012.Data.UPDateMode.AutoUpdate)
                mvarUpdateRow.SetValue("農地との関連", Not mvarRow.Item("農地との関連"))
                App農地基本台帳.TBL世帯.Update(mvarUpdateRow, mvarAddNew)
            Case "基本台帳印刷" : mod農地基本台帳.基本台帳印刷(Me.Key.KeyValue, ExcelViewMode.Preview, 印刷Mode.フル印刷)
            Case "基本台帳印刷(簡易)" : mod農地基本台帳.基本台帳印刷(Me.Key.KeyValue, ExcelViewMode.Preview, 印刷Mode.簡易印刷)

            Case "基本台帳印刷(編集モード)" : mod農地基本台帳.基本台帳印刷(Me.Key.KeyValue, ExcelViewMode.EditMode, 印刷Mode.フル印刷)

            Case "家族一覧"
                SysAD.page農家世帯.個人リスト.検索開始("[世帯ID]=" & Me.ID, "[世帯ID]=" & Me.ID)
            Case "経営農地", "経営農地の一覧"
                Dim sWhere As String = String.Format("([自小作別]=0 AND [所有世帯ID]={0}) Or ([自小作別]<>0 AND [借受世帯ID]={0})", Me.ID)
                SysAD.page農家世帯.農地リスト.検索開始(sWhere, sWhere)
            Case "所有農地", "所有農地の一覧", "所有地一覧"
                Dim sWhere As String = String.Format("[所有世帯ID]={0}", Me.ID)
                SysAD.page農家世帯.農地リスト.検索開始(sWhere, sWhere)
            Case "自作農地", "自作農地の一覧"
                Dim sWhere As String = String.Format("[自小作別]=0 AND [所有世帯ID]={0}", Me.ID)
                SysAD.page農家世帯.農地リスト.検索開始(sWhere, sWhere)
            Case "借受農地", "借受農地の一覧"
                Dim sWhere As String = String.Format("[自小作別]<>0 AND [借受世帯ID]={0}", Me.ID)
                SysAD.page農家世帯.農地リスト.検索開始(sWhere, sWhere)
            Case "貸出農地の一覧", "貸付農地"
                Dim sWhere As String = String.Format("[自小作別]<>0 AND [所有世帯ID]={0} AND [借受世帯ID]<>{0}", mvarRow.Item("ID"))
                SysAD.page農家世帯.農地リスト.検索開始(sWhere, sWhere)
            Case "管理農地"
                Dim sWhere As String = String.Format("[管理世帯ID]={0}", mvarRow.Item("ID"))
                SysAD.page農家世帯.農地リスト.検索開始(sWhere, sWhere)
            Case "削除"
                '    Select Case DVProperty.ClassStr
                '        Case "農家" : subOBJ削除("農家の削除", "世帯を削除しますか", "DELETE * FROM [D:世帯Info] WHERE [ID]=%DVP.ID%", Me)
                '        Case "営農情報"
                '            If DVProperty.Rs.V("兼業区分") = 4 Then
                '                MsgBox("非農家の営農情報は監理できません", vbCritical)
                '                Exit Function
                '            End If
                '            If MsgBox("営農情報を削除しますか", vbYesNo) = vbYes Then
                '                CDataviewSK_DoCommand2("閉じる")
                '                SystemDB.ExecuteSQL("DELETE * FROM [D_世帯営農] WHERE [ID]=" & DVProperty.ID)
                '                mvarPDW.SQLListview.Refresh()
                '            End If
                '    End Select
                'Case "農家番号の入力"
                '    n = Val(Fnc.InputText("農家番号", "農家番号を入力してください", Val(DVProperty.Rs.V("農家番号")), 1, 3))
                '    If n Then DVProperty.Rs.Update("農家番号", n) : mvarPDW.SQLListview.Refresh(DVProperty.Key)
                'Case "転用農地" : view転用農地一覧(DVProperty.ID)
                'Case "農機具の追加" : SystemDB.ExecuteSQL("INSERT INTO D_営農情報 ( SID, Class ) VALUES(" & DVProperty.ID & ",1);")
                'Case "農機具の一覧" : mvarPDW.SQLListview.SQLListviewCHead("SELECT '営農農機具.' & [ID] AS [KEY],'農機具' AS [名称],'Unit' AS [Icon] FROM [D_営農情報] WHERE [CLASS]=1;", "名称;名称", "農機具の一覧")
                '    '申請関連
                'Case "営農計画書" : mvarPDW.PrintGo(ObjectMan.GetObject("営農計画書.0"), DVProperty.Key)
                'Case "農地の所有権移動－３条", "所有権移転" : 農地法3条所有権移転(Me, "")

                'Case "農地の耕作権移動－３条" : 農地法3条耕作権移転(Me, "")
                'Case "農地の転用－４条" : 農地法4条設定(Me)
                'Case "関連申請"
                '    St = "SELECT  '申請.' & [D_申請].[ID] AS [KEY],D_申請.名称, M_BASICALL.名称 AS 状態, D_申請.受付年月日, D_申請.許可年月日 " & _
                '    "FROM ([D:個人Info] LEFT JOIN D_申請 ON [D:個人Info].ID = D_申請.申請者A) LEFT JOIN M_BASICALL ON D_申請.状態 = M_BASICALL.ID " & _
                '    "WHERE ((([D:個人Info].世帯ID)=" & DVProperty.ID & ") AND ((M_BASICALL.Class)='申請状況')) " & _
                '    "UNION SELECT '申請.' & [D_申請].[ID] AS [KEY], D_申請.名称, M_BASICALL.名称, D_申請.受付年月日, D_申請.許可年月日 " & _
                '    "FROM [D:個人Info] LEFT JOIN (D_申請 LEFT JOIN M_BASICALL ON D_申請.状態 = M_BASICALL.ID) ON [D:個人Info].ID = D_申請.申請者B " & _
                '    "WHERE ((([D:個人Info].世帯ID)=" & DVProperty.ID & ") AND ((M_BASICALL.Class)='申請状況'));"

                '    mvarPDW.SQLListview.SQLListviewCHead(St, "名称;名称;状態;状態;受付年月日;受付年月日;許可年月日;許可年月日", "関連申請")
                'Case ""
            Case "総括表" : 総括表()
                'Case "家族農地の連動"
                '    mvarPDW.WaitMessage = "家族の農地の結び付け処理中・・・"
                '    SystemDB.ExecuteSQL("UPDATE [D:農地Info] INNER JOIN [D:個人Info] ON [D:農地Info].所有者ID = [D:個人Info].ID SET [D:農地Info].所有世帯ID = [世帯ID] WHERE ((([D:農地Info].所有世帯ID)=0 Or ([D:農地Info].所有世帯ID) Is Null) AND (([D:個人Info].世帯ID)>0));")
                '    SystemDB.ExecuteSQL("UPDATE [D:農地Info] INNER JOIN [D:個人Info] ON [D:農地Info].借受人ID = [D:個人Info].ID SET [D:農地Info].借受世帯ID = [世帯ID] WHERE ((([D:個人Info].世帯ID)>0) AND (([D:農地Info].借受世帯ID)=0 Or ([D:農地Info].借受世帯ID) Is Null));")
                '    mvarPDW.WaitMessage = ""
                'Case "合併世帯の解除"
                '    Rs = SystemDB.GetRecordsetEx("SELECT [ID] , [氏名] FROM [D:個人Info] WHERE [合併異動]=TRUE AND [世帯ID]=" & DVProperty.ID, , , Me)
                '    St = Fnc.SelectMultiString(Rs.GetString(":", ";"), "選択してください")
                '    SystemDB.CloseRs(Rs)
                '    If Len(St) Then 合併世帯解除(St)

            Case "営農情報"
                'ObjectMan.GetObject("営農情報." & Me.ID).OpenDataViewNext()
                MsgBox("営農情報がありません", MsgBoxStyle.Critical)

                'Case "世帯情報" : ADApp.DataviewCol.Add(ObjectMan.GetObject("農家." & DVProperty.ID))
            Case "あっせん申出" : C申請データ作成.あっせん申出受(Me)
            Case "農業生産法人報告書"
                MsgBox("フォーマットが指定されていません")
            Case "営農計画書"
                MsgBox("フォーマットが指定されていません")
            Case "非農地証明書"
                MsgBox("フォーマットが指定されていません")
            Case "農地一覧"
            Case "設定無", "中心経営体", "中心経営体ではない", "調査中"
                Select Case sCommand
                    Case "設定無" : ValueChange("人農地プラン中心経営体区分", 0)
                    Case "中心経営体" : ValueChange("人農地プラン中心経営体区分", 1)
                    Case "中心経営体ではない" : ValueChange("人農地プラン中心経営体区分", 2)
                    Case "調査中" : ValueChange("人農地プラン中心経営体区分", 3)
                End Select

                Me.SaveMyself()
            Case Else
                Return MyBase.DoCommand(sCommand, sParams)
        End Select
        Return ""
    End Function

    Public Function 営農呼出し(ByVal nID As Integer, ByVal sMessErroe As String) As Boolean
        Dim pRow営農 As DataRow = App農地基本台帳.TBL世帯営農.Rows.Find(nID)
        If pRow営農 Is Nothing Then
            Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D_世帯営農] WHERE [ID]=" & nID)
            If pTBL.Rows.Count > 0 Then
                App農地基本台帳.TBL世帯営農.MergePlus(pTBL)
            Else
                SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO D_世帯営農(ID) VALUES(" & nID & ")")
                pTBL = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D_世帯営農] WHERE [ID]=" & nID)

                App農地基本台帳.TBL世帯営農.MergePlus(pTBL)
            End If
        End If
        CType(ObjectMan.GetObject("営農情報." & nID), HimTools2012.TargetSystem.CTargetObjWithView).OpenDataViewNext(SysAD.page農家世帯.DataViewCollection)
        Return True
    End Function

    Public Function GetDictionary(ByVal ParamArray pObj() As Object)
        If pObj.Length > 0 Then
            Dim pDic As New Dictionary(Of String, Object)
            For n As Integer = 0 To pObj.Length - 1 Step 2
                If TypeOf pObj(n) Is String And pObj.Length > n + 1 Then
                    pDic.Add(pObj(n).ToString, pObj(n + 1))
                End If
            Next
            Return pDic
        Else
            Return False
        End If
    End Function

    Public Overrides Sub DropDown(ByVal sSourceList As String, Optional ByVal sOption As String = "")
        Select Case GetKeyHead(sSourceList)
            Case "行政区"
                If GetKeyHead(sSourceList) = 0 Then
                ElseIf MsgBox("行政区を変更しますか?", vbYesNoCancel) = vbYes Then
                    mvarRow.Item("行政区ID") = GetKeyCode(sSourceList)
                End If
            Case "個人"
                Dim p個人 As CObj個人 = ObjectMan.GetObject(Strings.Left(sSourceList, InStr(sSourceList & ";", ";") - 1))
                Select Case OptionSelect("世帯を異動する;農地の関連を修復する", "異動する内容を選択してください。", "")
                    Case "世帯を異動する"
                        Sub世帯員異動(sSourceList)
                    Case "農地の関連を修復する"
                        Dim pN As New C農家異動(C農家異動.N処理.n農地の関連補正, GetDictionary("個人", p個人))
                    Case Else
                End Select
            Case "農家" : Sub世帯合併()
            Case "農地" : mod申請データ作成処理.農地申請データ作成処理(sSourceList, Me)
            Case Else

        End Select

    End Sub

#Region "総括表"
    Public Sub 総括表()
        If Not SysAD.page農家世帯.TabPageContainKey("総括表" & Me.Key.KeyValue) Then
            SysAD.page農家世帯.中央Tab.AddPage(New CTabPage総括表(Me))
        End If
    End Sub
#End Region

#Region "世帯員の異動"
    Public Sub Sub世帯員異動(ByVal sList As String)
        Do Until Not sList.EndsWith(";")
            sList = HimTools2012.StringF.Right(sList, sList.Length - 1)
        Loop
        Dim Ar As String() = Split(sList, ";")

        If MsgBox("世帯員 " & UBound(Ar) + 1 & "人を移動します。よろしいですか", vbYesNo) = vbYes Then
            For n = 0 To UBound(Ar)
                If Len(Ar(n)) Then
                    Dim ID As Decimal = GetKeyCode(CStr(Ar(n)))
                    SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:個人Info] SET [世帯ID]={1},[合併世帯ID]=0,[合併世帯]=False WHERE [ID]={0}", ID, Me.Key.ID)
                    SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [D:農地Info].所有世帯ID = " & Me.Key.ID & " WHERE ((([D:農地Info].所有者ID)=" & ID & "));")
                    SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [D:農地Info].管理世帯ID = " & Me.Key.ID & " WHERE ((([D:農地Info].管理者ID)=" & ID & "));")
                    SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [D:農地Info].借受世帯ID = " & Me.Key.ID & " WHERE ((([D:農地Info].借受人ID)=" & ID & ") AND (([D:農地Info].自小作別)>0));")
                End If
            Next
        End If
    End Sub
#End Region
#Region "世帯の合併"
    Private Sub Sub世帯合併()
        'Dim ZID As Long

        'If InStr(sSourceList, ";") Then
        '    MsgBox("複数の世帯のドロップには対応しておりません", vbCritical, "ドロップの失敗")
        'ElseIf DVProperty.ID = FncNet.GetKeyCode(sSourceList) Then
        '    SysAD.MDIForm.Message = "同じ世帯にドロップしました!"
        'ElseIf MsgBox("合併世帯として結び付けますか？", vbYesNo) = vbYes Then
        '    With SystemDB
        '        Rs = .GetRecordsetEx("SELECT ID, 続柄 FROM [V_続柄] WHERE [ID]>0;", , , Me)

        '        ZID = Val(FncNet.OptionSelect("-1:変更しない;" & Rs.GetString(":", ";"), "移動先の世帯主から見た移動元の世帯主との続柄を入力してください", "子"))
        '        .CloseRs(Rs)
        '        If ZID = 0 Then Exit Sub

        '        ID = FncNet.GetKeyCode(sSourceList)
        '        '世帯員の移動
        '        .ExecuteSQL("INSERT INTO D_個人履歴 ( PID, 異動事由, 内容, 異動日時 ) SELECT [D:個人Info].ID, 1000008 AS 式1, '合併世帯による世帯移動[" & ID & "]→[" & DVProperty.ID & "]' AS 式2, Now() AS 式3 FROM [D:個人Info] WHERE ((([D:個人Info].世帯ID)=" & ID & "));")
        '        .ExecuteSQL("UPDATE [D:個人Info] SET [D:個人Info].住記世帯番号 = [D:個人Info].世帯ID, [D:個人Info].住記続柄1 = [D:個人Info].続柄1, [D:個人Info].住記続柄2 = [D:個人Info].[続柄2], [D:個人Info].住記続柄3 = [D:個人Info].[続柄3] WHERE ((([D:個人Info].世帯ID)=" & ID & "));", "世帯員")

        '        If ZID = (-1) Then
        '            .ExecuteSQL("UPDATE [D:個人Info] SET [D:個人Info].世帯ID = " & DVProperty.ID & ",[D:個人Info].合併異動 = True,[D:個人Info].合併異動日=Now() WHERE ((([D:個人Info].世帯ID)=" & ID & "));", "世帯員の移動")
        '        Else
        '            .ExecuteSQL("UPDATE [D:個人Info] SET [D:個人Info].世帯ID = " & DVProperty.ID & ",[D:個人Info].続柄1 = " & ZID & ", [D:個人Info].続柄2 = [続柄1], [D:個人Info].続柄3 = [続柄2], [D:個人Info].続柄4 = [続柄3], [D:個人Info].合併異動 = True,[D:個人Info].合併異動日=Now() WHERE ((([D:個人Info].世帯ID)=" & ID & "));", "世帯員の移動")
        '            .ExecuteSQL("UPDATE [D:個人Info] SET [D:個人Info].続柄2 = 0 WHERE ((([D:個人Info].続柄2)=" & Val(.GetDirectData("S_システムデータ", "DATA", , "世帯主続柄コード")) & "));")
        '        End If

        '        '所有地の移動
        '        .ExecuteSQL("INSERT INTO D_土地履歴 ( LID, 更新日, 異動日, 異動事由, 内容 ) SELECT [D:農地Info].ID, Date() AS 式1, Date() AS 式2, 10008 AS 式3, '合併世帯による管理世帯変更[" & ID & "]→[" & DVProperty.ID & "]' AS 式4 FROM [D:農地Info] WHERE ((([D:農地Info].所有世帯ID)=" & ID & ")) OR ((([D:農地Info].[管理世帯ID])=" & ID & "));", "所有農地の履歴作成")
        '        .ExecuteSQL("UPDATE [D:農地Info] SET [D:農地Info].[管理世帯ID] = " & DVProperty.ID & " WHERE ((([D:農地Info].[管理世帯ID])=" & ID & ")) OR ((([D:農地Info].所有世帯ID)=" & ID & "));", "所有農地の移動")

        '        '小作地の移動
        '        .ExecuteSQL("INSERT INTO D_土地履歴 ( LID, 更新日, 異動日, 異動事由, 内容 ) SELECT [D:農地Info].ID, Date() AS 式1, Date() AS 式2, 10009 AS 式3, '合併世帯による借受世帯変更[" & ID & "]→[" & DVProperty.ID & "]' AS 式4 FROM [D:農地Info] WHERE ((([D:農地Info].借受世帯ID)=" & ID & "));", "小作農地の履歴作成")
        '        .ExecuteSQL("UPDATE [D:農地Info] SET [D:農地Info].借受世帯ID = " & DVProperty.ID & " WHERE ((([D:農地Info].借受世帯ID)=" & ID & "));", "小作農地の移動")

        '        '合併異動
        '        .ExecuteSQL("UPDATE [D:世帯Info] SET [D:世帯Info].農地との関連 = False, [D:世帯Info].合併異動 = True, [D:世帯Info].関連世帯 = " & DVProperty.ID & ",[D:世帯Info].合併異動日=Now() WHERE ((([D:世帯Info].ID)=" & ID & "));", "元世帯の処理")
        '        .ExecuteSQL("UPDATE [D:世帯Info] SET [D:世帯Info].合併異動 = True,[D:世帯Info].合併異動日=Now() WHERE (([D:世帯Info].ID)=" & DVProperty.ID & ");", "先世帯の処理")
        '    End With
        'End If

    End Sub
#End Region

    Public Overrides Function InitDataViewNext(ByRef pDB As HimTools2012.TargetSystem.CDataViewCollection, Optional InterfaceName As String = "") As Boolean
        If Me.DataViewPage Is Nothing Then
            Me.DataViewPage = New DataViewNext農家(Me)
        End If
        Return True
    End Function

    Public Overloads Overrides Function CanDropKeyHead(ByVal sKey As String, ByVal sOption As String) As Boolean
        Select Case GetKeyHead(sKey)
            Case "個人", "農地"
                Return True
            Case Else
                If Not SysAD.IsClickOnceDeployed Then
                    Debug.Print(sKey)
                    Stop
                End If
                Return False
        End Select
    End Function
    Public Overrides ReadOnly Property DataTableWithUpdateList As HimTools2012.Data.DataTableWithUpdateList
        Get
            Return App農地基本台帳.TBL世帯
        End Get
    End Property

    Public Overrides Function SaveMyself() As Boolean
        Return MyBase.SaveBase("D:世帯Info")
    End Function

    Public Overrides Sub ValueChange(ByVal sField As String, ByVal pValue As Object)
        App農地基本台帳.TBL世帯.AddUpdateListwithDataViewPage(Me, sField, pValue)
    End Sub
End Class

Public Class DataViewNext農家営農
    Inherits CDataViewPanel農地台帳


    Public Sub New(pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget, Nothing, SysAD.page農家世帯.DataViewCollection, True, True)
        'App農地基本台帳.TBL営農情報
        SetButtons()
        Dim nID As Integer = pTarget.ID

        Panel.FlowDirection = FlowDirection.LeftToRight

        Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsReadOnly, pTarget, "ID"), "ID")
        Panel.AddCtrl(New DateTimePickerPlus(True, pTarget, "更新日"), "更新日", em改行.改行あり)
        'AddItem","TX世帯ID","VB.TextBox;WithLabel=世帯ID;Alignment=2;BackColor=&HFFFFF0;Locked=True;Height=300;ToTop=0","ID"
        'AddItem","Ck情報公開拒否","VB.CheckBox;Caption=情報公開拒否;Width=1500;Height=330;","情報公開拒否"
        'AddItem","TX住所","VB.Textbox;WithLabel=住所;Width=6000;Locked=True;BackColor=&HFFFFF0;NewLine;","住所"
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("営農情報", Me), "", True), HimTools2012.controls.GroupBoxPlus)

            'AddItem","経営計画Opt","台帳管理EXE.OptionButtonsEX;WithLabel=経営計画;Width=4600;Height=300;DataListString=不明,経営規模拡大,現状維持,経営規模縮小;NewLine;VSkip=50;","経営計画"
            'AddItem","拡大縮小ID","VB.Textbox;WithLabel=拡大縮小方法;Height=300;Width=400;NewLine;VSkip=50;","拡大縮小方法"
            'AddCombo","拡大縮小Combo","MSComctlLib.ImageComboCtl.2;Width=2400;ImageList=MiniIcon;Height=300;","SELECT 'n.' & [ID] AS [KEY],[名称],'Unit' AS [ICON] FROM [M_BASICALL] WHERE [CLASS]='拡大縮小方法' ORDER BY [ID]","拡大縮小ID"
            'AddItem","TX希望面積","VB.Textbox;WithLabel=希望面積;Height=300;Width=1000;","希望面積"
            'AddItem","LB希望面積","VB.Label;Caption=ヘクタール;Width=1000;Height=300;BorderStyle=0;BackStyle=0;HSkip=100;",""
            'AddItem","経営意向Opt","台帳管理EXE.OptionButtonsEX;WithLabel=経営意向;Width=6900;Height=300;DataListString=不明,農業だけでやる,農業中心でやる,兼業中心でやる,農業をやめたい;NewLine;VSkip=50;","経営意向"
            'AddItem","TX希望年数","VB.Textbox;WithLabel=希望年数;Height=300;Width=800;NewLine;VSkip=50;","希望年数"
            'AddItem","LB希望年数","VB.Label;Caption=年後;Width=1000;Height=300;BorderStyle=0;BackStyle=0;HSkip=100;",""
            'BreakLine","Bk営農計画","VSkip=100;"
            'AddItem","経営計画米麦作Opt","台帳管理EXE.OptionButtonsEX;WithLabel=米麦作;Width=3500;Height=300;DataListString=現状維持,拡大,縮小;NewLine;","経営計画米麦作"
            'AddItem","経営計画畜産Opt","台帳管理EXE.OptionButtonsEX;WithLabel=畜産;Width=3500;Height=300;DataListString=現状維持,拡大,縮小;NewLine;","経営計画畜産"
            'AddItem","経営計画果樹Opt","台帳管理EXE.OptionButtonsEX;WithLabel=果樹;Width=3500;Height=300;DataListString=現状維持,拡大,縮小;NewLine;","経営計画果樹"
            'AddItem","経営計画そさいOpt","台帳管理EXE.OptionButtonsEX;WithLabel=そさい;Width=3500;Height=300;DataListString=現状維持,拡大,縮小;NewLine;","経営計画そさい"
            'AddItem","経営計画養蚕Opt","台帳管理EXE.OptionButtonsEX;WithLabel=養蚕;Width=3500;Height=300;DataListString=現状維持,拡大,縮小;NewLine;","経営計画養蚕"
            'BreakLine","Bk農機具","VSkip=100;"
            'AddGroupBtn","農機具・施設","NewLine;","農機具・施設","G農機具"
            'AddItem","農機具ID1","VB.Textbox;Group=G農機具;Height=300;Width=600;BackColor=&H00FFFFC0&;NewLine;","農機具種類1"
            'AddCombo","農機具Combo1","MSComctlLib.ImageComboCtl.2;Group=G農機具;Width=2700;ImageList=MiniIcon;Height=300;BackColor=&H00FFFFC0&;","SELECT 'n.' & [ID] AS [KEY],[名称],'Unit' AS [ICON] FROM [M_BASICALL] WHERE [CLASS]='農機具' ORDER BY [名称]","農機具ID1"
            'AddItem","TX農機具数量1","VB.Textbox;Group=G農機具;width=800;","農機具数量1"
            'AddItem","農機具ID2","VB.Textbox;Group=G農機具;Height=300;Width=600;;BackColor=&H00FFFFC0&;HSkip=100;","農機具種類2"
            'AddCombo","農機具Combo2","MSComctlLib.ImageComboCtl.2;Group=G農機具;Width=2700;ImageList=MiniIcon;Height=300;;BackColor=&H00FFFFC0&","SELECT 'n.' & [ID] AS [KEY],[名称],'Unit' AS [ICON] FROM [M_BASICALL] WHERE [CLASS]='農機具' ORDER BY [名称]","農機具ID2"
            'AddItem","TX農機具数量2","VB.Textbox;Group=G農機具;width=800;","農機具数量2"
            'AddButton","農機具追加BTN","VB.Commandbutton;Caption=農機具・施設追加;Group=G農機具;Width=1800;Height=300;HSkip=300;","農機具追加ボタン",""
            'AddItem","農機具ID3","VB.Textbox;Group=G農機具;Height=300;Width=600;BackColor=&H00FFFFC0&;NewLine;","農機具種類3"
            'AddCombo","農機具Combo3","MSComctlLib.ImageComboCtl.2;Group=G農機具;Width=2700;ImageList=MiniIcon;Height=300;BackColor=&H00FFFFC0&;","SELECT 'n.' & [ID] AS [KEY],[名称],'Unit' AS [ICON] FROM [M_BASICALL] WHERE [CLASS]='農機具' ORDER BY [名称]","農機具ID3"
            'AddItem","TX農機具数量3","VB.Textbox;Group=G農機具;width=800;","農機具数量3"
            'AddItem","農機具ID4","VB.Textbox;Group=G農機具;Height=300;Width=600;BackColor=&H00FFFFC0&;HSkip=100;","農機具種類4"
            'AddCombo","農機具Combo4","MSComctlLib.ImageComboCtl.2;Group=G農機具;Width=2700;ImageList=MiniIcon;Height=300;BackColor=&H00FFFFC0&;","SELECT 'n.' & [ID] AS [KEY],[名称],'Unit' AS [ICON] FROM [M_BASICALL] WHERE [CLASS]='農機具' ORDER BY [名称]","農機具ID4"
            'AddItem","TX農機具数量4","VB.Textbox;Group=G農機具;width=800;","農機具数量4"
            'AddItem","農機具ID5","VB.Textbox;Group=G農機具;Height=300;Width=600;BackColor=&H00FFFFC0&;NewLine;","農機具種類5"
            'AddCombo","農機具Combo5","MSComctlLib.ImageComboCtl.2;Group=G農機具;Width=2700;ImageList=MiniIcon;Height=300;BackColor=&H00FFFFC0&;","SELECT 'n.' & [ID] AS [KEY],[名称],'Unit' AS [ICON] FROM [M_BASICALL] WHERE [CLASS]='農機具' ORDER BY [名称]","農機具ID5"
            'AddItem","TX農機具数量5","VB.Textbox;Group=G農機具;width=800;","農機具数量5"
            'AddItem","農機具ID6","VB.Textbox;Group=G農機具;Height=300;Width=600;BackColor=&H00FFFFC0&;HSkip=100;","農機具種類6"
            'AddCombo","農機具Combo6","MSComctlLib.ImageComboCtl.2;Group=G農機具;Width=2700;ImageList=MiniIcon;Height=300;BackColor=&H00FFFFC0&;","SELECT 'n.' & [ID] AS [KEY],[名称],'Unit' AS [ICON] FROM [M_BASICALL] WHERE [CLASS]='農機具' ORDER BY [名称]","農機具ID6"
            'AddItem","TX農機具数量6","VB.Textbox;Group=G農機具;width=800;","農機具数量6"
            'AddItem","農機具ID7","VB.Textbox;Group=G農機具;Height=300;Width=600;BackColor=&H00FFFFC0&;NewLine;","農機具種類7"
            'AddCombo","農機具Combo7","MSComctlLib.ImageComboCtl.2;Group=G農機具;Width=2700;ImageList=MiniIcon;Height=300;BackColor=&H00FFFFC0&;","SELECT 'n.' & [ID] AS [KEY],[名称],'Unit' AS [ICON] FROM [M_BASICALL] WHERE [CLASS]='農機具' ORDER BY [名称]","農機具ID7"
            'AddItem","TX農機具数量7","VB.Textbox;Group=G農機具;width=800;","農機具数量7"
            'AddItem","農機具ID8","VB.Textbox;Group=G農機具;Height=300;Width=600;BackColor=&H00FFFFC0&;HSkip=100;","農機具種類8"
            'AddCombo","農機具Combo8","MSComctlLib.ImageComboCtl.2;Group=G農機具;Width=2700;ImageList=MiniIcon;Height=300;BackColor=&H00FFFFC0&;","SELECT 'n.' & [ID] AS [KEY],[名称],'Unit' AS [ICON] FROM [M_BASICALL] WHERE [CLASS]='農機具' ORDER BY [名称]","農機具ID8"
            'AddItem","TX農機具数量8","VB.Textbox;Group=G農機具;width=800;","農機具数量8"
            'AddItem","農機具ID9","VB.Textbox;Group=G農機具;Height=300;Width=600;BackColor=&H00FFFFC0&;NewLine;","農機具種類9"
            'AddCombo","農機具Combo9","MSComctlLib.ImageComboCtl.2;Group=G農機具;Width=2700;ImageList=MiniIcon;Height=300;BackColor=&H00FFFFC0&;","SELECT 'n.' & [ID] AS [KEY],[名称],'Unit' AS [ICON] FROM [M_BASICALL] WHERE [CLASS]='農機具' ORDER BY [名称]","農機具ID9"
            'AddItem","TX農機具数量9","VB.Textbox;Group=G農機具;width=800;","農機具数量9"
            'AddItem","農機具ID10","VB.Textbox;Group=G農機具;Height=300;Width=600;BackColor=&H00FFFFC0&;HSkip=100;","農機具種類10"
            'AddCombo","農機具Combo10","MSComctlLib.ImageComboCtl.2;Group=G農機具;Width=2700;ImageList=MiniIcon;Height=300;BackColor=&H00FFFFC0&;","SELECT 'n.' & [ID] AS [KEY],[名称],'Unit' AS [ICON] FROM [M_BASICALL] WHERE [CLASS]='農機具' ORDER BY [名称]","農機具ID10"
            'AddItem","TX農機具数量10","VB.Textbox;Group=G農機具;width=800;","農機具数量10"
            'BreakLine","Bk家畜","VSkip=50;"
            'AddGroupBtn","家畜","NewLine;VSkip=50","家畜","G家畜"
            'AddItem","家畜ID1","VB.Textbox;Group=G家畜;Height=300;Width=600;BackColor=&H00FFFFC0&;HSkip=100;NewLine;","家畜種類1"
            'AddCombo","家畜Combo1","MSComctlLib.ImageComboCtl.2;Group=G家畜;Width=2700;ImageList=MiniIcon;Height=300;BackColor=&H00FFFFC0&;","SELECT 'n.' & [ID] AS [KEY],[名称],'Unit' AS [ICON] FROM [M_BASICALL] WHERE [CLASS]='家畜' ORDER BY [名称]","家畜ID1"
            'AddItem","TX家畜数量1","VB.Textbox;Group=G家畜;width=800;","家畜数量1"
            'AddItem","家畜ID2","VB.Textbox;Group=G家畜;Height=300;Width=600;BackColor=&H00FFFFC0&;HSkip=100;","家畜種類2"
            'AddCombo","家畜Combo2","MSComctlLib.ImageComboCtl.2;Group=G家畜;Width=2700;ImageList=MiniIcon;Height=300;BackColor=&H00FFFFC0&;","SELECT 'n.' & [ID] AS [KEY],[名称],'Unit' AS [ICON] FROM [M_BASICALL] WHERE [CLASS]='家畜' ORDER BY [名称]","家畜ID2"
            'AddItem","TX家畜数量2","VB.Textbox;Group=G家畜;width=800;","家畜数量2"
            'AddButton","家畜追加BTN","VB.Commandbutton;Caption=家畜追加;Group=G家畜;Width=1200;Height=300;HSkip=300;","家畜追加ボタン",""
            'AddItem","家畜ID3","VB.Textbox;Group=G家畜;Height=300;Width=600;BackColor=&H00FFFFC0&;HSkip=100;NewLine;","家畜種類3"
            'AddCombo","家畜Combo3","MSComctlLib.ImageComboCtl.2;Group=G家畜;Width=2700;ImageList=MiniIcon;Height=300;BackColor=&H00FFFFC0&;","SELECT 'n.' & [ID] AS [KEY],[名称],'Unit' AS [ICON] FROM [M_BASICALL] WHERE [CLASS]='家畜' ORDER BY [名称]","家畜ID3"
            'AddItem","TX家畜数量3","VB.Textbox;Group=G家畜;width=800;","家畜数量3"
            'AddItem","家畜ID4","VB.Textbox;Group=G家畜;Height=300;Width=600;BackColor=&H00FFFFC0&;HSkip=100;","家畜種類4"
            'AddCombo","家畜Combo4","MSComctlLib.ImageComboCtl.2;Group=G家畜;Width=2700;ImageList=MiniIcon;Height=300;BackColor=&H00FFFFC0&;","SELECT 'n.' & [ID] AS [KEY],[名称],'Unit' AS [ICON] FROM [M_BASICALL] WHERE [CLASS]='家畜' ORDER BY [名称]","家畜ID4"
            'AddItem","TX家畜数量4","VB.Textbox;Group=G家畜;width=800;","家畜数量4"
            'AddItem","家畜ID5","VB.Textbox;Group=G家畜;Height=300;Width=600;BackColor=&H00FFFFC0&;HSkip=100;NewLine;","家畜種類5"
            'AddCombo","家畜Combo5","MSComctlLib.ImageComboCtl.2;Group=G家畜;Width=2700;ImageList=MiniIcon;Height=300;BackColor=&H00FFFFC0&;","SELECT 'n.' & [ID] AS [KEY],[名称],'Unit' AS [ICON] FROM [M_BASICALL] WHERE [CLASS]='家畜' ORDER BY [名称]","家畜ID5"
            'AddItem","TX家畜数量5","VB.Textbox;Group=G家畜;width=800;","家畜数量5"
            'AddItem","家畜ID6","VB.Textbox;Group=G家畜;Height=300;Width=600;BackColor=&H00FFFFC0&;HSkip=100;","家畜種類6"
            'AddCombo","家畜Combo6","MSComctlLib.ImageComboCtl.2;Group=G家畜;Width=2700;ImageList=MiniIcon;Height=300;BackColor=&H00FFFFC0&;","SELECT 'n.' & [ID] AS [KEY],[名称],'Unit' AS [ICON] FROM [M_BASICALL] WHERE [CLASS]='家畜' ORDER BY [名称]","家畜ID6"
            'AddItem","TX家畜数量6","VB.Textbox;Group=G家畜;width=800;","家畜数量6"
            'AddItem","家畜ID7","VB.Textbox;Group=G家畜;Height=300;Width=600;BackColor=&H00FFFFC0&;HSkip=100;NewLine;","家畜種類7"
            'AddCombo","家畜Combo7","MSComctlLib.ImageComboCtl.2;Group=G家畜;Width=2700;ImageList=MiniIcon;Height=300;BackColor=&H00FFFFC0&;","SELECT 'n.' & [ID] AS [KEY],[名称],'Unit' AS [ICON] FROM [M_BASICALL] WHERE [CLASS]='家畜' ORDER BY [名称]","家畜ID7"
            'AddItem","TX家畜数量7","VB.Textbox;Group=G家畜;width=800;","家畜数量7"
            'AddItem","家畜ID8","VB.Textbox;Group=G家畜;Height=300;Width=600;BackColor=&H00FFFFC0&;HSkip=100;","家畜種類8"
            'AddCombo","家畜Combo8","MSComctlLib.ImageComboCtl.2;Group=G家畜;Width=2700;ImageList=MiniIcon;Height=300;BackColor=&H00FFFFC0&;","SELECT 'n.' & [ID] AS [KEY],[名称],'Unit' AS [ICON] FROM [M_BASICALL] WHERE [CLASS]='家畜' ORDER BY [名称]","家畜ID8"
            'AddItem","TX家畜数量8","VB.Textbox;Group=G家畜;width=800;","家畜数量8"
            'BreakLine","Bk販売収入順位","VSkip=50;"
            'AddGroupBtn","販売収入順位","NewLine;","販売収入順位","G販売収入順位"
            'AddItem","TX販売収入順位米","VB.Textbox;WithLabel=米の順位;Group=G販売収入順位;width=800;NewLine;","販売収入順位米"
            'AddItem","TX販売収入順位畜産","VB.Textbox;WithLabel=畜産の順位;Group=G販売収入順位;width=800;","販売収入順位畜産"
            'AddItem","TX販売収入順位果樹","VB.Textbox;WithLabel=果樹の順位;Group=G販売収入順位;width=800;NewLine;","販売収入順位果樹"
            'AddItem","TX販売収入順位そさい","VB.Textbox;WithLabel=そさいの順位;Group=G販売収入順位;width=800;","販売収入順位そさい"
            'AddItem","TX販売収入順位養蚕","VB.Textbox;WithLabel=養蚕の順位;Group=G販売収入順位;width=800;NewLine;","販売収入順位養蚕"
            'AddItem","TX販売収入名称その他１","VB.Textbox;IMEMode=4;Alignment=0;Group=G販売収入順位;width=1500;","販売収入名称その他１"
            'AddItem","TX販売収入順位その他１","VB.Textbox;Group=G販売収入順位;width=800;","販売収入順位その他１"
            'AddItem","TX販売収入名称その他２","VB.Textbox;IMEMode=4;Alignment=0;Group=G販売収入順位;width=1500;NewLine;","販売収入名称その他２"
            'AddItem","TX販売収入順位その他２","VB.Textbox;Group=G販売収入順位;width=800;","販売収入順位その他２"
            'AddItem","TX販売収入名称その他３","VB.Textbox;IMEMode=4;Alignment=0;Group=G販売収入順位;width=1500;","販売収入名称その他３"
            'AddItem","TX販売収入順位その他３","VB.Textbox;Group=G販売収入順位;width=800;","販売収入順位その他３"
        End With
    End Sub

End Class

Public Class C農家異動
    Inherits HimTools2012.clsAccessor

    Public Enum N処理
        n農地の関連補正 = 1
    End Enum
    Public Result As String = ""
    Private mvarMode As N処理
    Private mvarParam As Dictionary(Of String, Object)

    Public Sub New(nMode As N処理, pDictionary As Dictionary(Of String, Object))
        mvarMode = nMode
        mvarParam = pDictionary
        With Me
            .Dialog.StartProc(True, True)

            If .Dialog._objException Is Nothing = False Then
                If .Dialog._objException.Message = "Cancel" Then
                    MsgBox("処理を中止しました。　", , "処理中止")
                    Exit Sub
                Else
                    Throw .Dialog._objException
                End If
            Else
                MsgBox("終了しました。", MsgBoxStyle.Information)
                Result = ""
            End If
        End With

    End Sub

    Public Overrides Sub Execute()
        Select Case mvarMode
            Case N処理.n農地の関連補正 : Sub農地の関連補正()
        End Select
    End Sub

    Private Sub Sub農地の関連補正()
        Dim p個人 As CObj個人 = mvarParam.Item("個人")
        With Me
            If p個人.世帯ID = 0 Then
                MsgBox("世帯番号が設定されていません", MsgBoxStyle.Critical)
            ElseIf MsgBox("世帯と個人の農地の関係を補正しますか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                Me.Message = "農地情報の取得"
                Dim p関連農地Tbl As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE [所有者ID]={0} Or [管理者ID]={0} Or [借受人ID]={0}", p個人.ID)
                App農地基本台帳.TBL農地.MergePlus(p関連農地Tbl)
                Dim pView As New DataView(App農地基本台帳.TBL農地.Body, String.Format("[所有者ID]={0} Or [管理者ID]={0} Or [借受人ID]={0}", p個人.ID), "", DataViewRowState.CurrentRows)

                .Value = 0
                .Maximum = pView.Count
                For Each pRow As DataRowView In pView
                    Me.Message = "異動処理(" & .Value & "/" & pView.Count & ")"
                    Me.Value += 1
                    Dim sSQL As New System.Text.StringBuilder
                    If Not IsDBNull(pRow.Item("所有者ID")) AndAlso pRow.Item("所有者ID") = p個人.ID Then
                        sSQL.Append(IIF(sSQL.Length > 0, ",", "") & "[所有世帯ID]=" & p個人.世帯ID)
                        pRow.Item("所有世帯ID") = p個人.世帯ID
                    End If
                    If Not IsDBNull(pRow.Item("管理者ID")) AndAlso pRow.Item("管理者ID") = p個人.ID Then
                        sSQL.Append(IIF(sSQL.Length > 0, ",", "") & "[管理世帯ID]=" & p個人.世帯ID)
                        pRow.Item("管理世帯ID") = p個人.世帯ID
                    End If
                    If Not IsDBNull(pRow.Item("借受人ID")) AndAlso pRow.Item("借受人ID") = p個人.ID Then
                        sSQL.Append(IIF(sSQL.Length > 0, ",", "") & "[借受世帯ID]=" & p個人.世帯ID)
                        pRow.Item("借受世帯ID") = p個人.世帯ID
                    End If

                    If sSQL.Length > 0 Then
                        SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET {0} WHERE [ID]={1}", sSQL.ToString, pRow.Item("ID"))
                        Make農地履歴(pRow.Item("ID"), Now, Now, 0, enum法令.職権異動, "世帯間の関連付け設定", p個人.ID)
                    End If
                Next
            End If
        End With
    End Sub
End Class

