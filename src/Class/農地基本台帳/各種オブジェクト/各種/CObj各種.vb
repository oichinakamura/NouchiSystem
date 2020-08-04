'/20160311霧島

Imports HimTools2012
Imports HimTools2012.CommonFunc
Imports HimTools2012.controls
Imports HimTools2012.controls.DVCtrlCommonBase

Public Class CObj各種
    Inherits CTargetObjWithView農地台帳


    Public Sub New(ByVal sKey As String)
        MyBase.New(Nothing, False, New HimTools2012.TargetSystem.DataKey(sKey), "")
    End Sub


    Public Sub New(ByVal pRow As DataRow, ByVal pClass As String, ByVal nCode As Long, ByVal bAddNew As Boolean)
        MyBase.New(pRow, bAddNew, New HimTools2012.TargetSystem.DataKey(pClass, pRow.Item("ID")), "")
    End Sub

    Public MyTreeNode As TreeNode
    Public Sub New(ByVal pTreeNode As TreeNode)
        MyBase.New(Nothing, False, New HimTools2012.TargetSystem.DataKey(pTreeNode.Name), "")
        MyTreeNode = pTreeNode
    End Sub

    Public Overrides Function InitDataViewNext(ByRef pDB As HimTools2012.TargetSystem.CDataViewCollection, Optional ByVal InterfaceName As String = "") As Boolean
        If Me.DataViewPage Is Nothing Then
            Select Case Me.Key.DataClass
                Case Else
                    Stop
            End Select
        End If
        Return True
    End Function



    Public Overridable Function 名称() As String
        Select Case Me.Key.DataClass
            Case "期別農業委員会"
                If Me.Row IsNot Nothing AndAlso Not Me.Row.Exists Then
                    Return App農地基本台帳.DataMaster.Rows.Find({Me.Key.ID, Me.Key.DataClass}).Item("名称").ToString
                End If
            Case ""
            Case Else

        End Select
        Return ""
    End Function

    Public Overrides Function DoCommand(ByVal sCommand As String, ByVal ParamArray sParams() As String) As Object
        Select Case sCommand
            Case "閉じる"
                If Me.DataViewPage IsNot Nothing Then
                    Me.DataViewPage.DoClose()
                End If
            Case "コードの表示" : MsgBox("Key:" & Me.Key.KeyValue, vbOKOnly, "コードの表示")
            Case "開く" : Open(sParams)

            Case Else
                Select Case sCommand & "-" & Me.Key.DataClass
                    Case "転用借地申請一覧-一時転用終期台帳"

                        '    Case "認定農業者-世帯を呼ぶ", "担い手農家-世帯を呼ぶ"
                        '        St = "農家." & SysAD.DB(sLRDB).GetDirectData("D:個人Info", "世帯ID", DVProperty.ID)
                        '        mvarPDW.DataviewCol.Add(ADApp.ObjectMan.GetObject(St))
                        '    Case "認定農業者-個人情報を呼ぶ", "担い手農家-個人情報を呼ぶ"
                        '        mvarPDW.DataviewCol.Add(ADApp.ObjectMan.GetObject("個人." & DVProperty.ID))
                        '    Case "市町村-市町村データの再サンプル"
                        '        St = SysAD.DB(sLRDB).GetDirectData("S_Folder", "名称", , DVProperty.Key)
                        '        If Len(St) Then
                        '            SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:個人Info] SET [D:個人Info].市町村ID = " & DVProperty.ID & " WHERE ((([D:個人Info].住所) Like '" & St & "%'));")
                        '        End If
                        '    Case "行政区-年金受給者一覧"
                        '        Debug.Assert(False)
                        '    Case "行政区-選挙確認リスト" : mvarPDW.PrintGo(New CPrint選挙一覧表, CStr(DVProperty.ID))
                        '    Case "認定農業者-削除"
                        '        If MsgBox("認定農業者を取り消しますか", vbYesNo) = vbYes Then
                        '            CDataviewSK_DoCommand2("閉じる")
                        '            SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:個人Info] SET [D:個人Info].農業改善計画認定 = 0 WHERE ((([D:個人Info].ID)=%DVP.ID%));", , , DVProperty)
                        '            mvarPDW.SQLListview.Refresh()
                        '        End If
                        '    Case "担い手農家-削除"
                        '        If MsgBox("担い手農家を取り消しますか", vbYesNo) = vbYes Then
                        '            CDataviewSK_DoCommand2("閉じる")
                        '            SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:個人Info] SET [D:個人Info].農業改善計画認定 = 0 WHERE ((([D:個人Info].ID)=%DVP.ID%));", , , DVProperty)
                        '            mvarPDW.SQLListview.Refresh()
                        '        End If
                        '    Case "農地大字-遊休農地一覧" : view農地List("WHERE [遊休化]=TRUE AND [大字ID]=" & DVProperty.ID, "遊休農地一覧") '/** 農地0001-0001-0002
                        '    Case "行政区-農家番号の振分け"
                        '        SysAD.DB(sLRDB).Execute("UPDATE [D:世帯Info] INNER JOIN [D:個人Info] ON [D:個人Info].ID = [D:世帯Info].世帯主ID SET [農家番号]=0 WHERE [D:個人Info].行政区ID=" & DVProperty.ID)
                        '        Rs = SysAD.DB(sLRDB).GetRecordsetEx("SELECT [D:世帯Info].[ID],[D:世帯Info].農家番号 FROM [D:個人Info] INNER JOIN [D:世帯Info] ON [D:個人Info].ID = [D:世帯Info].世帯主ID WHERE ((([D:個人Info].行政区ID)=" & DVProperty.ID & ") AND (([D:世帯Info].農地との関連)=True)) AND [D:個人Info].[住民区分]=" & Val(SysAD.DB(sLRDB).DBProperty("記録住民コード")) & " ORDER BY [D:個人Info].フリガナ;", , , Me)
                        '        n = 1
                        '        Do Until Rs.EOF
                        '            SysAD.DB(sLRDB).Execute("UPDATE [D:世帯Info] SET [農家番号]=" & n & " WHERE [ID]=" & Rs.ID)
                        '            Rs.MoveNext()
                        '            n = n + 1
                        '        Loop
                        '        Rs.CloseRs()
                        '        MsgBox("振り直しました")
                        '    Case "担い手農家-メニュー" : CDataviewSK_PopupMenu()


                        '    Case "農地大字-フォルダから削除"
                        '        SysAD.DB(sLRDB).ExecuteSQL("DELETE S_Folder.Key FROM S_Folder WHERE (((S_Folder.Key)='" & DVProperty.Key & "'));")
                        '        mvarPDW.Folder.Redraw(mvarPDW.Folder.ExpandedList)

                        '    Case "行政区-台帳一括印刷"
                        '        MsgBox("メモリが不足しています。一括印刷を行うことができません", vbCritical)
                        '    Case "行政区-世帯経営面積一覧"
                        '        view世帯集計List("[D:個人Info].[行政区ID]=" & DVProperty.ID, "世帯経営面積一覧")
                        '        '                mvarPDW.SQLListview.SQLListviewCHead DVProperty.StrReplace(ADApp.Common.AppData("世帯経営面積一覧SQL")), "世帯主名;氏名;フリガナ;フリガナ;住所;住所;@自作田;自作田;@自作畑;自作畑;@自作計;自作計;@小作田;小作田;@小作畑;小作畑;@小作計;小作計;@経営計;経営計", "世帯経営面積一覧:" & DVProperty.Name

                    Case "世帯員データを一時作成-住基"
                        Dim pRow As DataRow = App農地基本台帳.TBL個人.FindRowByID(Me.Key.ID)
                        If pRow IsNot Nothing Then
                            MsgBox(String.Format("同一の住民番号[{0}:{1}]で既に登録されています。", pRow.Item("ID"), pRow.Item("氏名")), MsgBoxStyle.Information)
                        ElseIf MsgBox("世帯員に一次変換しますか", vbYesNo) = vbYes Then

                            SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D:個人Info] ( ID, 世帯ID, 氏名, フリガナ, 住所, 住民区分, 生年月日, 性別, 続柄1, 続柄2, 続柄3, 行政区ID) " & _
                            "SELECT M_住民情報.ID, IIf(IsNull([M_住民情報].[世帯ID]),IIf(IsNull([M_住民情報].[世帯No]),0,[M_住民情報].[世帯No]),Val([M_住民情報].[世帯ID])), M_住民情報.氏名, M_住民情報.フリガナ, M_住民情報.住所, M_住民情報.住民区分, M_住民情報.生年月日, M_住民情報.性別, M_住民情報.続柄, M_住民情報.続柄2, M_住民情報.続柄3, M_住民情報.行政区 " & _
                            "FROM M_住民情報 WHERE [M_住民情報].ID=" & Me.Key.ID)

                            CType(ObjectMan.GetObject("個人." & Me.Key.ID), HimTools2012.TargetSystem.CTargetObjWithView).OpenDataViewNext(SysAD.page農家世帯.DataViewCollection)
                        End If
                    Case "農地の作成-固定土地"
                        If MsgBox("農地に変換しますか", vbYesNo) = vbYes Then
                            Dim nID As Decimal = App農地基本台帳.TBL農地.MinID - 1
                            Dim nTID As Decimal = App農地基本台帳.TBL転用農地.MinID - 1
                            If nTID < nID Then nID = nTID - 1

                            SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D:農地Info] ( ID, 一筆コード, 大字ID, 小字ID, 地番, 一部現況, 登記簿面積, 登記簿地目, 実面積, 現況地目, 自小作別, 所有者ID, 所有世帯ID, 登記名義人ID, 更新年月日, 田面積, 畑面積 ) " & _
                            "SELECT " & nID & ",M_固定情報.nID, M_固定情報.大字ID, M_固定情報.小字ID, M_固定情報.地番, M_固定情報.一部現況, M_固定情報.登記面積, M_固定情報.登記地目, M_固定情報.現況面積, M_固定情報.現況地目, 0, M_固定情報.所有者ID, [D:個人Info].世帯ID, M_固定情報.登記名義人ID, M_固定情報.異動年月日,IIf(InStr('," & SysAD.DB(sLRDB).DBProperty("田地目") & ",',',' & [現況地目] & ',')=0,0,[現況面積]),IIf(InStr('," & SysAD.DB(sLRDB).DBProperty("畑地目") & ",',',' & [現況地目] & ',')=0,0,[現況面積]) " & _
                            "FROM M_固定情報 LEFT JOIN [D:個人Info] ON M_固定情報.所有者ID = [D:個人Info].ID WHERE [nID]=" & Me.Key.ID)

                            App農地基本台帳.TBL農地.FindRowBySQL("[ID]=" & nID)
                            ObjectMan.GetObject("農地." & nID).DoCommand("開く")
                        End If
                    Case "所有地-公共機関"
                        Dim sWhere As String = String.Format("[所有者ID]={0}", Me.Key.ID)
                        SysAD.page農家世帯.農地リスト.検索開始(sWhere, sWhere)
                    Case "-市内世帯管理" : App農地基本台帳.CreateFarmer()
                    Case "関連農地を農地台帳から除外-公共機関"
                        Dim pID As Decimal = Me.Key.ID
                        Dim p法人 As CObj個人 = ObjectMan.GetObject("個人." & Me.Key.ID)
                        Dim sMess As String
                        If p法人 Is Nothing Then
                            sMess = "公共機関に属する土地の削除"
                        Else
                            sMess = p法人.氏名 & "に属する土地の削除"
                        End If

                        If MsgBoxVB("公共機関に関する土地を農地台帳から除外しますか", , MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                            農地削除(True, SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE [所有者ID]=" & pID), , , sMess)
                        End If

                    Case "農業委員会[n]期の追加-農業委員管理"
                        期別農業委員会の追加()
                    Case "編集-期別農業委員会"
                        期別農業委員会の編集(Me)
                    Case "削除-期別農業委員会"
                        期別農業委員会の削除(Me)
                    Case "今期農業委員会に設定-期別農業委員会"
                        SysAD.DB(sLRDB).DBProperty("今期農業委員会Key") = Me.Key.KeyValue
                        MsgBox("設定を終了しました。")


                    Case "総会資料作成-受付中" : 総会資料作成()
                    Case "総会資料作成-申請管理" : 総会資料作成()
                    Case "削除-農業委員"
                        If MsgBox("削除しますか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                            App農地基本台帳.DataMaster.DeleteData(Me.Key.DataClass, Me.Key.ID)
                        End If
                    Case "住民の追加-市内世帯管理"
                        App農地基本台帳.CreateFarmer()
                    Case "受付簿-所有権移転受付中"
                        '/***********************
                    Case "削除-土地履歴"
                        If MsgBox("履歴を削除しますか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                            Me.DoCommand("閉じる")
                            Dim pRow As DataRow = App農地基本台帳.TBL土地履歴.Rows.Find(Me.Key.ID)

                            If pRow IsNot Nothing Then
                                App農地基本台帳.TBL土地履歴.Rows.Remove(pRow)
                                SysAD.DB(sLRDB).ExecuteSQL("DELETE FROM [D_土地履歴] WHERE [ID]=" & Me.Key.ID)
                            End If
                        End If
                    Case Else
                        CasePrint(sCommand & "-" & Me.Key.DataClass)
                End Select
        End Select

        Return ""

    End Function

    Private Function Open(ByVal ParamArray sParams() As String) As Object
        Select Case Me.Key.DataClass
            Case "市外管理"
            Case "市内管理"
            Case "公共機関管理", "公共機関"
            Case "マスタ管理", "認定・担い手管理", "あっせん希望管理"
            Case "地域農地"
                Select Case Me.Key.ID
                    Case 1 : SysAD.page農家世帯.農地リスト.検索開始(String.Format("Int([大字ID]/100)={0}", Me.Key.ID), String.Format("[大字ID]>={0}00 AND [大字ID]<={0}99", Me.Key.ID))
                    Case 2 : SysAD.page農家世帯.農地リスト.検索開始(String.Format("Int([大字ID]/100)={0}", Me.Key.ID), String.Format("[大字ID]>={0}00 AND [大字ID]<={0}99", Me.Key.ID))
                    Case 3 : SysAD.page農家世帯.農地リスト.検索開始(String.Format("Int([大字ID]/100)={0}", Me.Key.ID), String.Format("[大字ID]>={0}00 AND [大字ID]<={0}99", Me.Key.ID))
                    Case 4 : SysAD.page農家世帯.農地リスト.検索開始(String.Format("Int([大字ID]/100)={0}", Me.Key.ID), String.Format("[大字ID]>={0}00 AND [大字ID]<={0}99", Me.Key.ID))
                End Select
            Case "行政区"
                SysAD.page農家世帯.農家リスト.検索開始("[D:個人Info].行政区ID=" & Me.Key.ID, "[世帯主行政区ID]=" & Me.Key.ID)
                SysAD.page農家世帯.農家リスト.Active()
            Case "買受希望者リスト" : SysAD.page農家世帯.農家リスト.検索開始("[あっせん希望]=True", "[あっせん希望]=True")
            Case "買受希望農地リスト" : SysAD.page農家世帯.農地リスト.検索開始("[あっせん希望]=True", "[あっせん希望]=True")
            Case "売渡希望農地リスト" : SysAD.page農家世帯.農地リスト.検索開始("[売渡希望]=True", "[売渡希望]=True")
            Case "貸付希望農地リスト" : SysAD.page農家世帯.農地リスト.検索開始("[貸付希望]=True", "[貸付希望]=True")
            Case "農業者年金一覧"
                Select Case Me.Key.ID
                    Case 1 : SysAD.page農家世帯.個人リスト.検索開始("[農年加入受給種別]=1", "[農年加入受給種別]=1")
                    Case 2 : SysAD.page農家世帯.個人リスト.検索開始("[農年加入受給種別]=2", "[農年加入受給種別]=2")
                    Case 3 : SysAD.page農家世帯.個人リスト.検索開始("[農年加入受給種別]=3", "[農年加入受給種別]=3")
                    Case 4 : SysAD.page農家世帯.個人リスト.検索開始("[農年加入受給種別]=4", "[農年加入受給種別]=4")
                End Select
            Case "年金対象者一覧"
                'SysAD.page農家世帯.個人リスト.検索開始("[経営移譲の有無]=True", "[経営移譲の有無]=True")
                OpenSQLList("年金対象者一覧", "SELECT '個人.' & [D:個人Info].[ID] AS [Key], [D:個人Info].ID, [D:個人Info].世帯ID, [D:個人Info].[フリガナ], [D:個人Info].氏名, [D:個人Info].生年月日, IIf(Format([生年月日],'mm/dd')>Format(Date(),'mm/dd'),DateDiff('yyyy',[生年月日],Date())-1,DateDiff('yyyy',[生年月日],Date())) AS 年齢, IIf([性別]=0,'男',IIf([性別]=1,'女',IIf([性別]=3,'法人'))) AS 性別2, V_住民区分.名称 AS 住民区分名, V_行政区.名称 AS 行政区名, IIf([農年加入受給種別]=1,'旧制度年金加入者',IIf([農年加入受給種別]=3,'新制度年金加入者',IIf([農年加入受給種別]=2,'旧制度年金受給者',IIf([農年加入受給種別]=4,'新制度年金受給者','')))) AS 農年加入受給種別2, IIf([経営移譲の有無]=True,'あり','なし') AS 経営移譲の有無2, IIf([老齢受給の有無]=True,'あり','なし') AS 老齢受給の有無2 FROM ([D:個人Info] LEFT JOIN V_住民区分 ON [D:個人Info].住民区分 = V_住民区分.ID) LEFT JOIN V_行政区 ON [D:個人Info].行政区ID = V_行政区.ID WHERE (((IIf([農年加入受給種別]=1,'旧制度年金加入者',IIf([農年加入受給種別]=3,'新制度年金加入者',IIf([農年加入受給種別]=2,'旧制度年金受給者',IIf([農年加入受給種別]=4,'新制度年金受給者','')))))<>'')) OR (((IIf([経営移譲の有無]=True,'あり','なし'))='あり')) OR (((IIf([老齢受給の有無]=True,'あり','なし'))='あり'));")
            Case "経営移譲年金"
                SysAD.page農家世帯.個人リスト.検索開始("[経営移譲の有無]=True", "[経営移譲の有無]=True")
            Case "農業者老齢年金"
                SysAD.page農家世帯.個人リスト.検索開始("[老齢受給の有無]=1 or [老齢受給の有無]=-1", "[老齢受給の有無]=1 or [老齢受給の有無]=-1")

            Case "市内世帯管理"

            Case "期別農業委員会"
                OpenTableFilterList("農業委員リスト." & Me.Key.ID, Me.名称, App農地基本台帳.DataMaster.Body, "[ParentKey]='" & Me.Key.KeyValue & "'", "[ID]")
            Case "農業委員管理"
            Case "住基"
            Case "マスタクラス"
                Dim sClass As String = Me.Key.Code
                OpenTableFilterList("マスタクラス" & sClass & ".0", "マスタ管理[" & sClass & "]", App農地基本台帳.DataMaster.Body, "[Class]='" & sClass & "'", "[ID]")
            Case "審査中"
            Case "申請管理"

            Case "利用権終期年毎" : OpenSQLList("利用権終期台帳年毎" & Me.Key.ID & "(" & 和暦Format(New Date(Me.Key.ID, 1, 1), "gggyy年") & ")", "SELECT '農地.' & [V_農地].[ID] AS [Key], '農地' AS アイコン, V_農地.土地所在, V_地目.名称 AS 登記地目, V_農地.登記簿面積, V_農地.小作開始年月日, V_農地.小作終了年月日, [D:個人Info].氏名 AS 借人, [D:個人Info].住所 AS 借人住所, [D:個人Info_1].氏名 AS 貸人, [D:個人Info_1].住所 AS 貸人住所, [D:個人Info_2].氏名 AS 経由農業生産法人,  IIf([小作形態]=1,'賃貸借',IIf([小作形態]=2,'使用貸借','その他')) AS 形態 FROM (((V_農地 INNER JOIN [D:個人Info] ON V_農地.借受人ID = [D:個人Info].ID) INNER JOIN [D:個人Info] AS [D:個人Info_1] ON V_農地.所有者ID = [D:個人Info_1].ID) LEFT JOIN [D:個人Info] AS [D:個人Info_2] ON V_農地.経由農業生産法人ID = [D:個人Info_2].ID) LEFT JOIN V_地目 ON V_農地.登記簿地目 = V_地目.ID WHERE (自小作別>0) AND (((Year([小作終了年月日]))=" & Me.Key.ID & ") AND ((V_農地.小作地適用法)=2));")
            Case "農地法終期年毎" : OpenSQLList("農地法終期台帳年毎" & Me.Key.ID & "(" & 和暦Format(New Date(Me.Key.ID, 1, 1), "gggyy年") & ")", "SELECT '農地.' & [V_農地].[ID] AS [Key], '農地' AS アイコン, V_農地.土地所在, V_地目.名称 AS 登記地目, V_農地.登記簿面積, V_農地.小作開始年月日, V_農地.小作終了年月日, [D:個人Info].氏名 AS 借人, [D:個人Info].住所 AS 借人住所, [D:個人Info_1].氏名 AS 貸人, [D:個人Info_1].住所 AS 貸人住所, [D:個人Info_2].氏名 AS 経由農業生産法人,  IIf([小作形態]=1,'賃貸借',IIf([小作形態]=2,'使用貸借','その他')) AS 形態 FROM (((V_農地 INNER JOIN [D:個人Info] ON V_農地.借受人ID = [D:個人Info].ID) INNER JOIN [D:個人Info] AS [D:個人Info_1] ON V_農地.所有者ID = [D:個人Info_1].ID) LEFT JOIN [D:個人Info] AS [D:個人Info_2] ON V_農地.経由農業生産法人ID = [D:個人Info_2].ID) LEFT JOIN V_地目 ON V_農地.登記簿地目 = V_地目.ID WHERE (自小作別>0) AND (((Year([小作終了年月日]))=" & Me.Key.ID & ") AND ((V_農地.小作地適用法)=1));")
            Case "地目別集計" : OpenSQLList("地目別集計", "SELECT V_大字.大字, V_小字.小字, V_地目.名称 AS 登記地目, V_現況地目.名称 AS 現況地目, Int(Sum([D:農地Info].登記簿面積)) AS 登記簿面積計,Int(Sum([D:農地Info].登記簿面積)*IIf([V_現況地目].[名称]='田',1,0)) AS 登記簿面積計_現況田,Int(Sum([D:農地Info].登記簿面積)*IIf([V_現況地目].[名称]='畑',1,0)) AS 登記簿面積計_現況畑, Int(Sum([D:農地Info].実面積)) AS 実面積の合計, Int(Sum([D:農地Info].実面積)*IIf([V_現況地目].[名称]='田',1,0)) AS 実面積の合計_現況田, Int(Sum([D:農地Info].実面積)*IIf([V_現況地目].[名称]='畑',1,0)) AS 実面積の合計_現況畑, Sum(IIf([V_地目].[名称]='田',1,0)) AS 登記地目田筆数, Sum(IIf([V_現況地目].[名称]='田',1,0)) AS 現況地目田筆数, Sum(IIf([V_地目].[名称]='畑',1,0)) AS 登記地目畑筆数, Sum(IIf([V_現況地目].[名称]='畑',1,0)) AS 現況地目畑筆数 FROM ((([D:農地Info] LEFT JOIN V_小字 ON [D:農地Info].小字ID = V_小字.ID) INNER JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) LEFT JOIN V_地目 ON [D:農地Info].登記簿地目 = V_地目.ID) LEFT JOIN V_現況地目 ON [D:農地Info].現況地目 = V_現況地目.ID GROUP BY V_大字.大字, V_小字.小字, V_地目.名称, V_現況地目.名称;", {"登記簿面積計", "登記簿面積計_現況田", "登記簿面積計_現況畑"})
            Case "農地法終期台帳" : OpenSQLList("農地法終期台帳", "SELECT '農地法終期年毎.' & Year([小作終了年月日]) as [Key], Format([小作終了年月日],'gggee年') AS 貸借終了年,Count([D:農地Info].ID) AS 筆数, Sum([D:農地Info].田面積) AS 田面積の合計, Sum([D:農地Info].畑面積) AS 畑面積の合計,'申請' AS [アイコン] FROM [D:農地Info] WHERE (自小作別>0) AND ((([D:農地Info].小作地適用法)=1) AND (([D:農地Info].小作終了年月日)>#1/1/1900#)) GROUP BY '農地法終期年毎.' & Year([小作終了年月日]),'申請',Format([小作終了年月日],'gggee年') ORDER BY Format([小作終了年月日],'gggee年');")
                Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Year([小作終了年月日]) AS 式1 FROM [D:農地Info] WHERE ((([D:農地Info].自小作別)=True) AND (([D:農地Info].小作地適用法)=1)) GROUP BY Year([小作終了年月日]);")
                Dim pNode As TreeNode = SysAD.page農家世帯.FolderTree.FindTree(Me.Key.KeyValue)
                If pNode IsNot Nothing Then
                    For Each pRow As DataRow In pTBL.Rows
                        Dim pFind As TreeNode() = pNode.Nodes.Find("農地法終期年毎." & pRow.Item("式1"), True)
                        If pFind.Length = 0 Then
                            Dim n年 As Integer = pRow.Item("式1")
                            Dim pCNode As TreeNode = pNode.Nodes.Add(n年 & "年度(" & 和暦Format(New Date(n年, 1, 1), "gggyy年") & ")")
                            pCNode.Name = "農地法終期年毎." & pRow.Item("式1")
                        End If
                    Next
                End If
            Case "中間管理機構を介した貸借農地"
                Dim sWhere As String = String.Format("[自小作別]<>0 AND [経由農業生産法人ID]={0}", Val(SysAD.DB(sLRDB).DBProperty("中間管理機構ID")))
                SysAD.page農家世帯.農地リスト.検索開始(sWhere, sWhere)
            Case "利用権終期台帳" : OpenSQLList("利用権終期台帳", "SELECT '利用権終期年毎.' & Year([小作終了年月日]) as [Key], Format([小作終了年月日],'gggee年') AS 貸借終了年,Count([D:農地Info].ID) AS 筆数, Sum([D:農地Info].田面積) AS 田面積の合計, Sum([D:農地Info].畑面積) AS 畑面積の合計,'申請' AS [アイコン] FROM [D:農地Info] WHERE (自小作別>0) AND ((([D:農地Info].小作地適用法)=2) AND (([D:農地Info].小作終了年月日)>#1/1/1900#)) GROUP BY '利用権終期年毎.' & Year([小作終了年月日]),'申請',Format([小作終了年月日],'gggee年') ORDER BY Format([小作終了年月日],'gggee年');")
                Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Year([小作終了年月日]) AS 式1 FROM [D:農地Info] WHERE ((([D:農地Info].自小作別)=True) AND (([D:農地Info].小作地適用法)=2)) GROUP BY Year([小作終了年月日]);")
                Dim pNode As TreeNode = SysAD.page農家世帯.FolderTree.FindTree(Me.Key.KeyValue)
                If pNode IsNot Nothing Then
                    For Each pRow As DataRow In pTBL.Rows
                        Dim pFind As TreeNode() = pNode.Nodes.Find("利用権終期年毎." & pRow.Item("式1"), True)
                        If pFind.Length = 0 Then
                            Dim n年 As Integer = pRow.Item("式1")
                            Dim pCNode As TreeNode = pNode.Nodes.Add(n年 & "年度(" & 和暦Format(New Date(n年, 1, 1), "gggyy年") & ")")
                            pCNode.Name = "利用権終期年毎." & pRow.Item("式1")
                        End If
                    Next
                End If


            Case "事業計画変更"
            Case "中間管理機構"
                Dim nID As Decimal = Val(SysAD.DB(sLRDB).DBProperty("中間管理機構ID"))
                If Val(SysAD.DB(sLRDB).DBProperty("中間管理機構ID")) = 0 Then
                    Dim sID As String = InputBox("中間管理機構IDを入力してください" & vbCrLf & "(新しい番号を作成する場合は'-1'のまま)", "中間管理機構IDの入力", "-1")
                    If sID = "-1" Then
                        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT MIn([ID])-1 AS [MInID] FROM [D:個人Info]")
                        If pTBL IsNot Nothing AndAlso pTBL.Rows.Count = 1 Then
                            nID = Val(pTBL.Rows(0).Item("MinID").ToString)
                            SysAD.DB(sLRDB).DBProperty("中間管理機構ID") = nID
                            SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D:個人Info] ( ID, 氏名, フリガナ ) VALUES({0},'中間管理機構','ﾁｭｳｶﾝｶﾝﾘｷｺｳ')", nID)
                        End If
                    ElseIf Val(sID) <> 0 Then
                        SysAD.DB(sLRDB).DBProperty("中間管理機構ID") = Val(sID)
                    Else
                        Return Nothing
                    End If
                Else
                    Open個人(nID, "登録された番号で中間管理機構が見つかりません")
                End If

            Case "一時転用終期台帳"
                Dim sKey As String = "転用農地検索.0"
                Dim pList As C転用農地リスト
                If Not SysAD.page農家世帯.TabPageContainKey(sKey) Then
                    pList = New C転用農地リスト(sKey, "転用農地検索")
                    pList.Name = sKey
                    SysAD.page農家世帯.中央Tab.AddPage(pList)
                    pList.ImageKey = pList.IconKey
                Else
                    pList = SysAD.page農家世帯.GetItem(sKey)
                End If

                pList.検索開始("[小作形態]=7 Or [農地状況]=1040", "[小作形態]=7 Or [農地状況]=1040")
            Case "農業改善計画認定"
                Select Case Me.Key.ID
                    Case 1 : SysAD.page農家世帯.個人リスト.検索開始("[農業改善計画認定]=1 Or [農業改善計画認定]=4", "[農業改善計画認定]=1 Or [農業改善計画認定]=4")
                    Case 2 : SysAD.page農家世帯.個人リスト.検索開始("[農業改善計画認定]=2 Or [農業改善計画認定]=4", "[農業改善計画認定]=2 Or [農業改善計画認定]=4")
                    Case 3 : SysAD.page農家世帯.個人リスト.検索開始("[農業改善計画認定]=3", "[農業改善計画認定]=3")
                    Case 4 : SysAD.page農家世帯.個人リスト.検索開始("[農業改善計画認定]=4", "[農業改善計画認定]=4")
                End Select
            Case "家族経営協定管理" : SysAD.page農家世帯.農家リスト.検索開始("[家族経営協定]=True", "[家族経営協定]=True")
            Case "パソコン研修管理" : SysAD.page農家世帯.農家リスト.検索開始("[パソコン研修]=True", "[パソコン研修]=True")
            Case "農業法人研修希望" : SysAD.page農家世帯.農家リスト.検索開始("[農業法人研修希望]=True", "[農業法人研修希望]=True")
            Case "市内農地管理"
                If MsgBox("表示に時間がかかる恐れがあります。よろしいですか？", vbOKCancel) = vbOK Then
                    SysAD.page農家世帯.農地リスト.検索開始("[大字ID]>0 And [大字ID] Is Not Null", "[大字ID]>0 And [大字ID] Is Not Null")
                End If
            Case "市外農地管理"
                If MsgBox("表示に時間がかかる恐れがあります。よろしいですか？", vbOKCancel) = vbOK Then
                    SysAD.page農家世帯.農地リスト.検索開始("[大字ID]<1 Or [大字ID] Is Null", "[大字ID]<1 Or [大字ID] Is Null")
                End If
            Case "不在者管理"
                OpenSQLList("不在地主管理", String.Format("SELECT '個人.' & [D:個人Info].[ID] AS [Key], [D:個人Info].ID, [D:個人Info].世帯ID, [D:個人Info].[フリガナ], [D:個人Info].氏名, [D:個人Info].住所, [D:個人Info].生年月日, M_BASICALL.名称 AS 性別, V_住民区分.名称 AS 住民区分名, V_行政区.行政区, IIf([農年受給の有無]=True,'有','無') AS 農年受給, IIf([老齢受給の有無]=True,'有','無') AS 老齢受給, IIf([経営移譲受給の有無]=True,'有','無') AS 経営移譲受給, [D:個人Info].更新日 FROM ((([D:個人Info] LEFT JOIN V_住民区分 ON [D:個人Info].住民区分 = V_住民区分.ID) LEFT JOIN V_行政区 ON [D:個人Info].行政区ID = V_行政区.ID) INNER JOIN [D:農地Info] ON [D:個人Info].ID = [D:農地Info].所有者ID) LEFT JOIN M_BASICALL ON [D:個人Info].性別 = M_BASICALL.ID WHERE (((M_BASICALL.Class)='性別') AND (([D:個人Info].住民区分)<>{0} And ([D:個人Info].住民区分)<>{1}) AND (([D:農地Info].大字ID)>0)) GROUP BY '個人.' & [D:個人Info].[ID], [D:個人Info].ID, [D:個人Info].世帯ID, [D:個人Info].[フリガナ], [D:個人Info].氏名, [D:個人Info].住所, [D:個人Info].生年月日, M_BASICALL.名称, V_住民区分.名称, V_行政区.行政区, IIf([農年受給の有無]=True,'有','無'), IIf([老齢受給の有無]=True,'有','無'), IIf([経営移譲受給の有無]=True,'有','無'), [D:個人Info].更新日;", Val(SysAD.DB(sLRDB).DBProperty("記録住民コード")), Val(SysAD.DB(sLRDB).DBProperty("死亡住民コード"))))
            Case "特例対象農地" : SysAD.page農家世帯.農地リスト.検索開始("[特例農地区分]=True", "[特例農地区分]=True")
        End Select
        Return Nothing
    End Function

    Protected Function GetBitweenYearStr(ByVal sField As String, ByVal StartYear As Integer) As String
        Return "[" & sField & "]>=#1/1/" & StartYear & "# AND [" & sField & "]<=#12/31/" & StartYear & "#"
    End Function

    Protected Function GetBitweenMonthStr(ByVal sField As String, ByVal StartYear As Integer, ByVal StartMonth As Integer) As String
        If StartMonth < 12 Then
            Return "[" & sField & "]>=#" & StartMonth & "/1/" & StartYear & "# AND [" & sField & "]<#" & (StartMonth + 1) & "/1/" & StartYear & "#"
        Else
            Return "[" & sField & "]>=#12/1/" & StartYear & "# AND [" & sField & "]<=#12/31/" & StartYear & "#"
        End If
    End Function

    Public Overrides Sub DropDown(ByVal sSourceList As String, Optional ByVal sOption As String = "")
        Select Case Me.Key.DataClass & "-" & GetKeyHead(sSourceList)
            '    Case "マスタグループ-大字管理", "マスタグループ-小字管理", "マスタグループ-地目管理", "マスタグループ-町外大字管理", "マスタグループ-住民区分管理", "マスタグループ-性別コード", "マスタグループ-農機具コード", "マスタグループ-自小作区分管理", "マスタグループ-続柄管理"
            '        SysAD.DB(sLRDB).ExecuteSQL("UPDATE S_Folder SET S_Folder.ParentKey = '" & Me.Key & "' WHERE (((S_Folder.Key) = '" & Ar(i) & "'));")
            '        mvarPDW.Folder.Redraw(mvarPDW.Folder.ExpandedList)
            '    Case "大字管理-町外大字"
            '        SysAD.DB(sLRDB).SetDirectDataWithClass("M_BASICALL", "Class", "町外大字", "大字", FncNet.GetKeyCode(CStr(Ar(i))))
            '    Case "町外大字管理-大字"
            '        SysAD.DB(sLRDB).SetDirectDataWithClass("M_BASICALL", "Class", "大字", "町外大字", FncNet.GetKeyCode(CStr(Ar(i))))
            Case "農業委員管理-個人"

            Case "中間管理機構-個人"

            Case "農業改善計画認定-個人"
                Dim p個人 As CObj個人 = ObjectMan.GetObject(sSourceList)
                p個人.ValueChange("農業改善計画認定", Me.Key.ID)
                p個人.更新日 = Now
                p個人.SaveMyself()
            Case "公共機関管理-個人"
                Dim p個人 As CObj個人 = ObjectMan.GetObject(sSourceList)
                Dim St As String = SysAD.DB(sLRDB).DBProperty("公共機関コード", "")
                Dim sList() As String = Split(St, ",")

                If Not sList.Contains(p個人.ID) Then
                    If MsgBox("[" & p個人.氏名 & "]を公共機関に登録しますか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        SysAD.DB(sLRDB).DBProperty("公共機関コード") = SysAD.DB(sLRDB).DBProperty("公共機関コード") & IIf(SysAD.DB(sLRDB).DBProperty("公共機関コード").Length > 0, ",", "") & p個人.ID
                        Dim pNode() As TreeNode = SysAD.page農家世帯.FolderTree.Body.Nodes.Find(Me.Key.KeyValue, True)
                        pNode(0).Nodes.Add(p個人.氏名).Name = "公共機関." & p個人.ID
                    End If
                Else
                    MsgBox("既に存在します", MsgBoxStyle.Critical)
                End If
            Case "買受希望者リスト-農家"
                Dim ar As String() = Split(sSourceList, ";")
                Dim nCount As Integer = HimTools2012.StringF.SplitCount(sSourceList, ";")
                If MsgBox(String.Format("農家{0}件を買受あっせん希望に設定しますか？", nCount), MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    Dim IDList As String = HimTools2012.StringF.IDList(sSourceList)
                    If IDList > 0 Then
                        SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:世帯Info] SET [あっせん希望]=True WHERE [ID] IN ({0})", {IDList})
                    End If
                End If
                '        Case "市町村-世帯員"
                '            SysAD.DB(sLRDB).Execute("UPDATE [D:個人Info] SET [D:個人Info].市町村ID = " & DVProperty.ID & " WHERE [D:個人Info].ID IN (" & GetKeyCodeList(sSourceList, "世帯員") & ");")
                '            mvarPDW.SQLListview.Refresh()

                '        Case "行政区-農業委員"

                '        Case "行政区-農家"
                '            mvarPDW.WaitMessage = "** 移動中 **"

                '            SID = SysAD.DB(sLRDB).GetDirectData("S_システムデータ", "Data", , "市町村ID")
                '            Ar = Split(sSourceList, ";")
                '            For n = LBound(Ar) To UBound(Ar)
                '                SysAD.DB(sLRDB).SetDirectData("D:世帯Info", "市町村ID", SID, FncNet.GetKeyCode(CStr(Ar(n))))
                '                mvarPDW.WaitMessage = "** " & n & "/" & UBound(Ar) & " **"
                '            Next

                '            mvarPDW.WaitMessage = ""
            Case "中間管理機構を介した貸借農地-農地", "中間管理機構-農地"
                If SysAD.SystemInfo.ユーザー.n権利 > 0 Then
                    Dim pKeys As String() = Split(sSourceList, ";")
                    Dim n機構ID As Long = Val(SysAD.DB(sLRDB).DBProperty("中間管理機構ID"))
                    Dim nCount As Integer = 0
                    For Each sKey As String In pKeys

                        Select Case GetKeyHead(sKey)
                            Case "農地"
                                Dim p農地 As CObj農地 = ObjectMan.GetObject(sKey)
                                If p農地 Is Nothing Then
                                    If Not p農地.GetLongIntValue("経由農業生産法人ID").Equals(n機構ID) Then
                                        p農地.ValueChange("経由農業生産法人ID", n機構ID)
                                        nCount += 1
                                        p農地.SaveMyself()
                                    End If
                                End If
                        End Select
                    Next
                    MsgBox(String.Format("管理機構を介した貸借に設定しました。({0} / {1})", nCount, pKeys.Length), vbInformation)
                Else
                    MsgBox("権限がないため処理できませんでした", MsgBoxStyle.Critical)
                End If
            Case "期別農業委員会-個人"
                Dim p個人 As CObj個人 = ObjectMan.GetObject(sSourceList)

                If p個人 IsNot Nothing AndAlso MsgBox("[" & p個人.ToString & "]を" & Me.名称 & "に設定しますか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    Dim Ar() As String = Split(sSourceList, ";")
                    Dim sName As String = p個人.氏名
                    Dim nID As Decimal = App農地基本台帳.DataMaster.GetMaxID("農業委員") + 1
                    App農地基本台帳.DataMaster.AddData("農業委員", nID, sName, 0, p個人.Key.KeyValue, Me.Key.KeyValue)

                End If
            Case Else
                CasePrint(Me.Key.DataClass & "-" & GetKeyHead(sSourceList))
        End Select


    End Sub

    Public Overrides Function GetContextMenu(Optional ByVal pMenu As HimTools2012.controls.MenuItemEx = Nothing, Optional ByVal nDips As Integer = 1, Optional ByVal sParam() As String = Nothing) As HimTools2012.controls.MenuPlus
        Dim pMenuf As New HimTools2012.controls.ContextMenuEx(AddressOf ClickMenu)
        Dim bEdit As Boolean = (SysAD.SystemInfo.ユーザー.n権利 > 0)
        'ID クラスのないオブジェクトで失敗
        Select Case Me.Key.DataClass
            Case "行政区" : pMenuf.AddMenu("開く", , AddressOf ClickMenu)
            Case "４条許可済月別" : pMenuf.AddMenu("公布簿", , AddressOf ClickMenu)
            Case "４条許可済年別" : pMenuf.AddMenu("公布簿", , AddressOf ClickMenu)
            Case "５条許可済年別" : pMenuf.AddMenu("公布簿", , AddressOf ClickMenu)
            Case "４条許可済" : pMenuf.AddMenu("公布簿", , AddressOf ClickMenu)
            Case "５条許可済"
                pMenuf.AddMenu("開く", , AddressOf ClickMenu)
                pMenuf.AddMenu("公布簿", , AddressOf ClickMenu)
            Case "市内世帯管理" : pMenuf.AddMenu("住民の追加", , AddressOf ClickMenu, , bEdit)
            Case "住基" : pMenuf.AddMenu("世帯員データを一時作成", , AddressOf ClickMenu, , bEdit)
            Case "農業委員管理" : pMenuf.AddMenu("農業委員会[n]期の追加", , AddressOf ClickMenu, , bEdit)
            Case "申請管理" : pMenuf.AddMenu("総会資料作成", , AddressOf ClickMenu, , bEdit)
            Case "３条許可済月別"
            Case "18条承認済月別"
            Case "一時転用終期台帳" ':   pMenuf.AddMenu("転用借地申請一覧", AddressOf ClickMenu)
            Case "３条許可済"
            Case "受付中" : pMenuf.AddMenu("総会資料作成", , AddressOf ClickMenu, , bEdit)
            Case "市内農地管理" : pMenuf.AddMenu("固定資産情報より農地追加", , AddressOf ClickMenu)
            Case "市外管理", "市内管理" : pMenuf.AddMenu("開く", , AddressOf ClickMenu)
            Case "農業改善計画認定" : pMenuf.AddMenu("開く", , AddressOf ClickMenu)
            Case "固定土地" : pMenuf.AddMenu("農地の作成", , AddressOf ClickMenu)
            Case "マスタ管理" : pMenuf.AddMenu("開く", , AddressOf ClickMenu)
            Case "利用権終期台帳" : pMenuf.AddMenu("開く", , AddressOf ClickMenu)
            Case "公共機関管理" : pMenuf.AddMenu("開く", , AddressOf ClickMenu)
            Case "中間管理機構" : Return Nothing
            Case "農地法3条1項の届出" : Return Nothing
            Case "３条１項処理済過年度" : Return Nothing
            Case "あっせん申出渡受付中" : Return Nothing
            Case "審査中" : Return Nothing
            Case "中間管理機構を介した貸借農地"
            Case "農地法終期台帳"
            Case "不在者管理"
            Case "公共機関"
                pMenuf.AddMenu("開く", , AddressOf ClickMenu)
                pMenuf.AddMenu("所有地", , AddressOf ClickMenu)
                pMenuf.AddMenu("関連農地を農地台帳から除外", , AddressOf ClickMenu)
            Case "期別農業委員会"
                If SysAD.DB(sLRDB).DBProperty("今期農業委員会Key", "") <> Me.Key.KeyValue Then
                    pMenuf.AddMenu("今期農業委員会に設定", , AddressOf ClickMenu)
                End If
                pMenuf.AddMenu("開く", , AddressOf ClickMenu)
                pMenuf.AddMenu("編集", , AddressOf ClickMenu)
                pMenuf.AddMenu("削除", , AddressOf ClickMenu)
            Case "農業委員" : pMenuf.AddMenu("削除", , AddressOf ClickMenu)
            Case Else
#If DEBUG Then

                CasePrint(Me.Key.DataClass)
                Return Nothing
#End If
        End Select

        'Select Case DVProperty.ClassStr
        '    Case "行政区" : St = "開く;" & n & "世帯の追加;>一覧;世帯経営面積一覧;" & n & "-;年金受給者一覧;貸出農家;<;台帳一括印刷;農家番号の振分け;-;最上位に移動する;最下位置に移動する;"
        '    Case "町内管理" : St = "開く;-;" & n & ";-;最上位に移動する;最下位置に移動する;"
        '    Case "市町村" : St = "開く;農家の追加;-;基本台帳の一括印刷;市町村データの再サンプル;-;コードの表示;最下位置に移動する;" & n & "削除;"
        '    Case "担い手農家" : St = "開く;世帯を呼ぶ;個人情報を呼ぶ;-;" & n & "削除;-;閉じる;"
        '    Case "認定農業者" : St = "開く;世帯を呼ぶ;個人情報を呼ぶ;-;" & n & "削除;-;閉じる;"
        '    Case "年金受給個人" : St = "開く;世帯を呼ぶ;個人情報を呼ぶ;-;" & n & "削除;"

        '    Case "農地追加", "農地追加個人" : St = ""
        '    Case "農地大字" : St = "遊休農地一覧;" & n & "フォルダから削除;"
        '    Case "土地" : St = IIf(DVProperty.Rs.Value("非農地"), "", "~") & n & "農地へ変換;"
        '    Case "固定" : St = n & "農地データを一時作成;"

        '    Case Else
        '        Debug.Assert(CaseAssertPrint(DVProperty.ClassStr, "NK97.CObj各種エラー"))
        'End Select

        Return pMenuf
    End Function


    Private Sub Sub行政区_農家追加()
        'Dim St As String
        'Dim n As Long
        'Dim Rs As NK97.RecordsetEx

        'n = Fnc.DBGetMinID(SystemDB, "D:世帯INFO")
        'St = InputBox("世帯番号を入力してください", "世帯の追加", n)
        'n = Val(St)

        'If n Then
        '    Rs = SysAD.DB(sLRDB).GetRecordsetEx("SELECT * FROM [D:世帯INFO] WHERE [ID]=" & n, , , Me)
        '    If Rs.EOF Then
        '        Rs = SysAD.DB(sLRDB).GetRecordsetEx("SELECT * FROM [D:世帯INFO]", 2, 3, Me)
        '        Rs.UpdateValues(GetArray("ID", "行政区ID", "農地との関連"), GetArray(n, DVProperty.ID, True))
        '        SysAD.DB(sLRDB).CloseRs(Rs)
        '        mvarPDW.DataviewCol.Add(ADApp.ObjectMan.GetObject("農家." & n))
        '    Else
        '        mvarPDW.DataviewCol.Add(ADApp.ObjectMan.GetObject("農家." & n))
        '    End If
        '    Rs.CloseRs()
        'End If
    End Sub



    Public Overrides Function GetProperty(ByVal sParam As String) As Object
        Select Case Me.Key.DataClass

            Case Else
                If Not SysAD.IsClickOnceDeployed Then
                    Debug.Print(Me.Key.DataClass)
                    Stop
                End If
        End Select


        Return ""
        'Private Function CDataviewSK_GetProperty2(ByVal sPropName As String, Optional ByVal sParam As String = "") As Object
        '    Select Case DVProperty.ClassStr & "-" & sPropName
        '        Case "認定農業者-名称" : CDataviewSK_GetProperty2 = "認定農業者:" & SysAD.DB(sLRDB).GetDirectData("D:個人Info", "氏名", DVProperty.ID)
        '        Case "担い手農家-名称" : CDataviewSK_GetProperty2 = "担い手農家:" & SysAD.DB(sLRDB).GetDirectData("D:個人Info", "氏名", DVProperty.ID)
        '        Case "依頼-名称" : CDataviewSK_GetProperty2 = DVProperty.Rs.Value("件名")
        '        Case "年金受給個人-名称" : CDataviewSK_GetProperty2 = DVProperty.Rs.Value("氏名")
        '        Case "行政区-メニュー" : CDataviewSK_GetProperty2 = PopSub()

        '        Case "農地追加-名称" : CDataviewSK_GetProperty2 = "農地追加"
        '        Case "農地追加個人-名称" : CDataviewSK_GetProperty2 = "農地追加個人"
        '        Case Else
        '            Debug.Assert(CaseAssertPrint(DVProperty.ClassStr & "-" & sPropName, "CObj各種_GetProperty2"))
        '    End Select
        'End Function
    End Function


    Public Overloads Overrides Function CanDropKeyHead(ByVal sKey As String, ByVal sOption As String) As Boolean

        Select Case GetKeyHead(sKey) & "-" & Me.Key.DataClass
            Case "個人-農業改善計画認定" : Return True
            Case "農家-買受希望者リスト" : Return True
            Case "個人-公共機関管理" : Return True
            Case "個人-中間管理機構" : Return True
            Case "個人-あっせん希望管理" : Return True
            Case "個人-期別農業委員会" : Return True
            Case "個人-年金加入者" : Return False
            Case "個人-年金受給者" : Return False
            Case "個人-市内管理" : Return False
            Case "個人-市外管理" : Return False
            Case "個人-認定・担い手管理" : Return False
            Case "個人-農家の取り組み管理" : Return False
            Case "個人-農業委員管理" : Return False
            Case "個人-申請管理" : Return False
            Case "個人-公共機関" : Return False
            Case "個人-市内農地管理" : Return False
            Case "個人-取消し" : Return False
            Case "個人-取下げ" : Return False
            Case "個人-審査中" : Return False
            Case "農地-あっせん希望管理" : Return False
            Case "農地-利用権移転承認済" : Return False
            Case "個人-市内世帯管理" : Return False '\CanDropKeyHead
            Case "農地-取下げ"
            Case "農地-申請管理"
            Case "農地-中間管理機構" : Return True
            Case "農地-中間管理機構を介した貸借農地" : Return True
            Case "農地-公共機関管理" : Return False
            Case "農家-行政区" : Return False
            Case "農地-年金加入者" : Return False
            Case "農地-農業委員管理" : Return False
            Case "農地-農家の取り組み管理" : Return False
            Case "農地-年金受給者" : Return False
            Case "利用権設定承認済-利用権設定承認済" : Return False
            Case "個人-農地台帳" : Return False
            Case "個人-集計機能" : Return False
            Case "転用農地-年金加入者" : Return False
            Case "農地-集計機能"
            Case Else
                CasePrint(GetKeyHead(sKey) & "-" & Me.Key.DataClass, "return false")

                Return False
        End Select

        Return False
    End Function
    Public Overrides ReadOnly Property DataTableWithUpdateList As HimTools2012.Data.DataTableWithUpdateList
        Get
            Return Nothing
        End Get
    End Property

    Public Overrides Function SaveMyself() As Boolean
        Return False
    End Function

    Public Overrides Sub ValueChange(ByVal sField As String, ByVal pValue As Object)

    End Sub

End Class

Public Class DataViewNextコード
    Inherits CDataViewPanel農地台帳

    Public Sub New(pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget, Nothing, SysAD.page農家世帯.DataViewCollection, True, True)
        SetButtons()
        Dim nID As Integer = pTarget.ID
        'Aパターン
        '"AddItem","CodeID","VB.TextBox;WithLabel=ID;Alignment=2:BackColor=&HFFFF00;Locked=False;Height=300;MinLeft=550;ToTop=0;","ID"
        '"AddItem","CodeClass","VB.TextBox;Alignment=0;WithLabel=クラス名;Locked=False;Height=300;MinLeft=550;NewLine;","Class"
        '"AddItem","Code名称","VB.TextBox;WithLabel=名称;Alignment=0:BackColor=&HFFFF00;Locked=False;Height=300;Width=2000;MinLeft=550;NewLine;","名称"
        '"AddItem","Code親","VB.TextBox;WithLabel=親コード;Alignment=0:Height=300;Width=2000;MinLeft=550;NewLine;","ParentKey"
        '"AddItem","sParamTX","VB.TextBox;WithLabel=パラメータ;Alignment=0;OLEDropMode=1;Height=300;Width=8000;NewLine;","sParam"
        '"AddItem","nParamTX","VB.TextBox;WithLabel=数値パラメータ;Alignment=1;OLEDropMode=2;Height=300;Width=2000;NewLine;","nParam"


        'Bパターン
        '"DataviewSettingS","ButtonWidth=1000;LabelBorderStyle=1;TextFontSize=11;LabelWidth=1500;TextWidth=1000;TextHeight=330"
        '"AddItem","CodeCODE","VB.TextBox;WithLabel=CODE;Alignment=0:BackColor=&HFFFF00;Locked=False;Height=300;MinLeft=550;ToTop=0;","CODE"
        '"AddItem","CodeClass","VB.TextBox;Alignment=0;WithLabel=クラス名;Locked=False;Height=300;MinLeft=550;NewLine;","Class"
        '"AddItem","Code名称","VB.TextBox;WithLabel=名称;Alignment=0:BackColor=&HFFFF00;Locked=False;Height=300;Width=2000;MinLeft=550;NewLine;","名称"
        '"AddItem","Code親","VB.TextBox;WithLabel=親コード;Alignment=0:Height=300;Width=2000;MinLeft=550;NewLine;","ParentKey"

    End Sub

End Class


Public Class DataViewNext担い手農家
    Inherits CDataViewPanel農地台帳

    Public Sub New(pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget, Nothing, SysAD.page農家世帯.DataViewCollection, True, True)
        SetButtons()
        Dim nID As Integer = pTarget.ID
        With Panel
            Dim nHeight As Integer = .AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsReadOnly, pTarget, "ID", , 80), "ID").Height
            .AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "氏名", , 200), "氏名", em改行.改行あり)
        End With
    End Sub

End Class

Public Class DataViewNext認定農業者
    Inherits CDataViewPanel農地台帳

    Public Sub New(pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget, Nothing, SysAD.page農家世帯.DataViewCollection, True, True)
        SetButtons()
        Dim nID As Integer = pTarget.ID
        With Panel
            Dim nHeight As Integer = .AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsReadOnly, pTarget, "ID", , 80), "ID").Height
            .AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsReadOnly, pTarget, "認定番号", , 200), "認定番号")
            .AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "氏名", , 200), "氏名", em改行.改行あり)
            .AddCtrl(New DateTimePickerPlus(False, mvarTarget, "認定日"), "認定日", em改行.改行あり)
            .AddCtrl(New DateTimePickerPlus(False, mvarTarget, "認定期限日"), "認定期限日", em改行.改行あり)
        End With
    End Sub

End Class

Public Class DataViewNext年金受給個人
    Inherits CDataViewPanel農地台帳

    Public Sub New(pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget, Nothing, SysAD.page農家世帯.DataViewCollection, True, True)
        SetButtons()
        Dim nID As Integer = pTarget.ID

        With Panel
            Dim nHeight As Integer = .AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsReadOnly, pTarget, "ID", , 80), "ID").Height
            .AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "氏名", , 200), "氏名", em改行.改行あり)
        End With

    End Sub

End Class


Public Class DataViewNext農地追加
    Inherits CDataViewPanel農地台帳

    Public Sub New(pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget, Nothing, SysAD.page農家世帯.DataViewCollection, True, True)
        SetButtons()

        Dim nID As Integer = pTarget.ID

        '"AddHeader","農地追加","土地所在,土地所在,ViewText;所在,所在,DataText;地番,地番,DataText;登記簿面積,登記簿面積,DataDBL;田面積,田面積,DataDBL;畑面積,畑面積,DataDBL;登記簿地目,登記簿地目,DataLong;現況地目,現況地目,ComboList,地目"
        '"AddSQL","農地追加","SELECT 'Key.' & [V_農地].[ID] AS [Key], [V_農地].* FROM [V_農地] INNER JOIN [D:農地Info] ON [V_農地].ID = [D:農地Info].ID WHERE [V_農地].[備考]='User追加' AND [V_農地].[所有世帯ID]=%DVP.ID%;"
        '"AddSQL","地目","SELECT 'Key.' & [V_地目].ID AS [KEY],[名称] FROM [V_地目];"
        '"AddGrid","Grid農地追加","台帳管理EXE.NakaGridCtrl;Width=8000;Height=3500;ToTop=0;","%SQL.農地追加%","%HEADER.農地追加%"
        '"AddItem","所有者ID","VB.Textbox;WithLabel=所有者名;Alignment=1;Width=1000;NewLine;",""
        '"AddCombo","所有者Combo","MSComctlLib.ImageComboCtl.2;Text=;ImageList=MiniIcon;Width=3000;","SELECT 'n.' & [ID] AS [KEY],[氏名] AS [名称],'Farmer' AS [ICON] FROM [D:個人Info] WHERE [世帯ID]=%DVP.ID%","所有者ID"
        '"AddButton","農地追加","VB.Commandbutton;Caption=農地追加;Group=G小作;Width=900;Height=300;NewLine","農地追加ボタン",""
        '"AddButton","農地削除","VB.Commandbutton;Caption=農地削除;Group=G小作;Width=900;Height=300;","農地削除ボタン",""

    End Sub

End Class

Public Class DataViewNext農地追加個人
    Inherits CDataViewPanel農地台帳

    Public Sub New(pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget, Nothing, SysAD.page農家世帯.DataViewCollection, True, True)
        SetButtons()

        Dim nID As Integer = pTarget.ID

        '"AddHeader","農地追加","土地所在,土地所在,ViewText;所在,所在,DataText;地番,地番,DataText;登記簿面積,登記簿面積,DataDBL;田面積,田面積,DataDBL;畑面積,畑面積,DataDBL;登記簿地目,登記簿地目,DataLong;現況地目,現況地目,ComboList,地目"
        '"AddSQL","農地追加","SELECT 'Key.' & [V_農地].[ID] AS [Key], [V_農地].* FROM [V_農地] INNER JOIN [D:農地Info] ON [V_農地].ID = [D:農地Info].ID WHERE [V_農地].[備考]='User追加' AND [V_農地].[所有者ID]=%DVP.ID%;"
        '"AddSQL","地目","SELECT 'Key.' & [V_地目].ID AS [KEY],[名称] FROM [V_地目];"
        '"AddGrid","Grid農地追加","台帳管理EXE.NakaGridCtrl;Width=8000;Height=3500;ToTop=0;","%SQL.農地追加%","%HEADER.農地追加%"
        '"AddButton","農地追加","VB.Commandbutton;Caption=農地追加;Group=G小作;Width=900;Height=300;NewLine","農地追加ボタン",""
        '"AddButton","農地削除","VB.Commandbutton;Caption=農地削除;Group=G小作;Width=900;Height=300;","農地削除ボタン",""
    End Sub

End Class


