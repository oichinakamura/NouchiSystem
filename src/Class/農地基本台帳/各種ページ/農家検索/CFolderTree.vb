'20160531霧島

Imports HimTools2012.CommonFunc
Imports HimTools2012.controls

Public Class CFolderTree
    Inherits HimTools2012.TabPageSK.CFolderTreeBase

    Public Sub New(pPNode As Xml.XmlNode, pLayout As HimTools2012.controls.XMLLayout)
        MyBase.New(False, "フォルダ", "フォルダ")
        For Each pAttr As Xml.XmlAttribute In pPNode.Attributes
            Select Case pAttr.Name
                Case "Name" : Me.Name = pAttr.Value
                Case "CustomType"
                Case Else
                    Stop
            End Select
        Next
        For Each pChild As Xml.XmlNode In pPNode.ChildNodes
            Select Case pChild.Name
                Case Else
                    Stop
            End Select
        Next

        InitTree()
    End Sub
    Private 許可済み申請TBL As DataTable
    'SELECT Year([許可年月日]) AS 年, Month([許可年月日]) AS 月, D_申請.法令 FROM D_申請 WHERE (((D_申請.状態)=2)) GROUP BY Year([許可年月日]), Month([許可年月日]), D_申請.法令 HAVING (((Year([許可年月日])) Is Not Null) AND ((Month([許可年月日])) Is Not Null)) ORDER BY Year([許可年月日]), Month([許可年月日]);

    Public Overrides Sub InitTree()
        MyBase.InitTree()
        Body.Nodes.Clear()
        With mvarTreeView.Nodes.Add(CType(SysAD.市町村, C市町村別).市町村種別 & "内")
            .Name = "市内管理.0"
            With .Nodes.Add("世帯")
                .Name = "市内世帯管理.0"
                Dim pView As DataTable = SysAD.MasterView("行政区").ToTable

                For Each pVR As DataRow In pView.Rows
                    If pVR.Item("ID") > 0 Then
                        Dim pItem As TreeNode = .Nodes.Add(pVR.Item("名称"))
                        pItem.Name = "行政区." & pVR.Item("ID")
                        pItem.ImageKey = "CloseFolder"
                    End If
                Next
            End With
            With .Nodes.Add("農地")
                .Name = "市内農地管理.0"
                If SysAD.市町村.市町村名 = "日置市" Then
                    .Nodes.Add("東市来農地").Name = "地域農地.1"
                    .Nodes.Add("伊集院農地").Name = "地域農地.2"
                    .Nodes.Add("日吉農地").Name = "地域農地.3"
                    .Nodes.Add("吹上農地").Name = "地域農地.4"
                End If

            End With
            With .Nodes.Add("公共機関")
                .Name = "公共機関管理.0"
                Dim sKey As String = SysAD.DB(sLRDB).DBProperty("公共機関コード")
                If sKey.Length > 0 Then
                    Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:個人Info] WHERE [ID] IN (" & sKey & ")")

                    For Each pRow As DataRow In pTBL.Rows
                        With .Nodes.Add(pRow.Item("氏名").ToString)
                            .Name = "公共機関." & pRow.Item("ID")
                        End With
                    Next
                End If

                '公共機関コード
            End With
        End With
        With mvarTreeView.Nodes.Add(CType(SysAD.市町村, C市町村別).市町村種別 & "外")
            .Name = "市外管理.0"
            With .Nodes.Add("不在地主管理")
                .Name = "不在者管理.0"

            End With
            With .Nodes.Add("農地")
                .Name = "市外農地管理.0"

            End With
        End With

        許可済み申請TBL = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Year([許可年月日]) AS 年, Month([許可年月日]) AS 月, D_申請.法令 FROM D_申請 WHERE (((D_申請.状態)=2)) GROUP BY Year([許可年月日]), Month([許可年月日]), D_申請.法令 HAVING (((Year([許可年月日])) Is Not Null) AND ((Month([許可年月日])) Is Not Null)) ORDER BY Year([許可年月日]), Month([許可年月日]);")


        With mvarTreeView.Nodes.Add("申請")
            .Name = "申請管理.0"

            With .Nodes.Add("受付中")
                .Name = "受付中.0"
                .Nodes.Add("農地法３条").Name = "３条受付中.0"
                .Nodes.Add("農地法3条1項の届出").Name = "農地法3条1項の届出.0"

                .Nodes.Add("農地法４条").Name = "４条受付中.0"
                .Nodes.Add("農地法５条").Name = "５条受付中.0"
                .Nodes.Add("18条解約受付中").Name = "18条解約受付中.0"

                .Nodes.Add("基盤強化法所有権移転").Name = "所有権移転受付中.0"
                .Nodes.Add("基盤強化法利用権設定").Name = "利用権設定受付中.0"
                .Nodes.Add("基盤強化法利用権移転").Name = "利用権移転受付中.0"
                .Nodes.Add("合意解約受付中").Name = "合意解約受付中.0"

                .Nodes.Add("あっせん申出(渡)").Name = "あっせん申出渡受付中.0"
                .Nodes.Add("あっせん申出(受)").Name = "あっせん申出受受付中.0"

                .Nodes.Add("農地改良届").Name = "農地改良届受付中.0"
                .Nodes.Add("農地利用目的変更").Name = "農地利用目的変更受付中.0"
                .Nodes.Add("農用地利用計画変更").Name = "農用地利用計画変更受付中.0"
                .Nodes.Add("非農地証明願い").Name = "非農地証明願い受付中.0"
                .Nodes.Add("事業計画変更").Name = "事業計画変更受付中.0"
                '
                .Nodes.Add("買受適格-耕作-公売").Name = "買受適格-耕作-公売受付中.0"
                .Nodes.Add("買受適格-耕作-競売").Name = "買受適格-耕作-競売受付中.0"
                .Nodes.Add("買受適格-転用-公売").Name = "買受適格-転用-公売受付中.0"
                .Nodes.Add("買受適格-転用-競売").Name = "買受適格-転用-競売受付中.0"

                .Nodes.Add(New TreeNodeX("現地調査表作成.0", "現地調査表作成", GetType(C各種申請管理)))
            End With
            With .Nodes.Add("審査中")
                .Name = "審査中.0"
                .Nodes.Add("農地法３条").Name = "３条審査中.0"
                .Nodes.Add("農地法４条").Name = "４条審査中.0"
                .Nodes.Add("農地法５条").Name = "５条審査中.0"

                .Nodes.Add("基盤強化法所有権移転").Name = "所有権移転審査中.0"
                .Nodes.Add("基盤強化法利用権設定").Name = "利用権設定審査中.0"
                .Nodes.Add("基盤強化法利用権移転").Name = "利用権移転審査中.0"

                .Nodes.Add("あっせん申出(渡)").Name = "あっせん申出渡審査中.0"
                .Nodes.Add("あっせん申出(受)").Name = "あっせん申出受審査中.0"

                .Nodes.Add("農地利用目的変更").Name = "農地利用目的変更審査中.0"
                .Nodes.Add("農用地利用計画変更").Name = "農用地利用計画変更審査中.0"
                .Nodes.Add("事業計画変更").Name = "事業計画変更審査中.0"
            End With
            With .Nodes.Add("許可・承認・処理済")
                Set過去履歴(.Nodes.Add("農地法３条"), "３条許可済.0", "３条許可済", {enum法令.農地法3条所有権, enum法令.農地法3条耕作権})
                Set過去履歴(.Nodes.Add("農地法３条１項"), "３条１項処理済.0", "３条１項処理済", {enum法令.農地法3条の3第1項})
                Set過去履歴(.Nodes.Add("農地法４条"), "４条許可済.0", "４条許可済", {enum法令.農地法4条, enum法令.農地法4条一時転用})
                Set過去履歴(.Nodes.Add("農地法５条"), "５条許可済.0", "５条許可済", {enum法令.農地法5条所有権, enum法令.農地法5条貸借, enum法令.農地法5条一時転用})
                Set過去履歴(.Nodes.Add("農地法18条"), "18条承認済.0", "18条承認済", {enum法令.農地法18条解約, enum法令.農地法20条解約}, "承認")
                Set過去履歴(.Nodes.Add("基盤強化法所有権移転"), "所有権移転承認済.0", "所有権移転承認済", {enum法令.基盤強化法所有権})
                Set過去履歴(.Nodes.Add("基盤強化法利用権設定"), "利用権設定承認済.0", "利用権設定承認済", {enum法令.利用権設定})
                Set過去履歴(.Nodes.Add("基盤強化法利用権移転"), "利用権移転承認済.0", "利用権移転承認済", {enum法令.利用権移転})
                Set過去履歴(.Nodes.Add("合意解約"), "合意解約承認済.0", "合意解約承認済", {enum法令.合意解約, enum法令.中間管理機構へ農地の返還}, "承認")

                Set過去履歴(.Nodes.Add("あっせん申出(渡)"), "あっせん申出渡承認済.0", "あっせん申出渡承認済", {enum法令.あっせん出手})
                Set過去履歴(.Nodes.Add("あっせん申出(受)"), "あっせん申出受承認済.0", "あっせん申出受承認済", {enum法令.あっせん受手})

                Set過去履歴(.Nodes.Add("農地改良届"), "農地改良届承認済.0", "農地改良届承認済", {enum法令.農地改良届})
                Set過去履歴(.Nodes.Add("農地利用目的変更"), "農地利用目的変更承認済.0", "農地利用目的変更承認済", {enum法令.農地利用目的変更})
                Set過去履歴(.Nodes.Add("農用地利用計画変更"), "農用地利用計画変更承認済.0", "農用地利用計画変更承認済", {enum法令.農用地計画変更})

                Set過去履歴(.Nodes.Add("非農地証明願済"), "非農地証明願済.0", "非農地証明願済", {enum法令.非農地証明願})
                Set過去履歴(.Nodes.Add("事業計画変更"), "事業計画変更承認済.0", "事業計画変更承認済", {enum法令.事業計画変更})

                Set過去履歴(.Nodes.Add("買受適格-耕作-公売"), "買受適格-耕作-公売承認済.0", "買受適格-耕作-公売承認済", {enum法令.非農地証明願})
                Set過去履歴(.Nodes.Add("買受適格-耕作-競売"), "買受適格-耕作-競売承認済.0", "買受適格-耕作-競売承認済", {enum法令.事業計画変更})
                Set過去履歴(.Nodes.Add("買受適格-転用-公売"), "買受適格-転用-公売承認済.0", "買受適格-転用-公売承認済", {enum法令.非農地証明願})
                Set過去履歴(.Nodes.Add("買受適格-転用-競売"), "買受適格-転用-競売承認済.0", "買受適格-転用-競売承認済", {enum法令.事業計画変更})
            End With
            With .Nodes.Add("取下げ")
                .Name = "取下げ.0"
                .Nodes.Add("農地法３条").Name = "３条取下げ.0"
                .Nodes.Add("農地法3条1項の届出").Name = "３条１項取下げ.0"

                .Nodes.Add("農地法４条").Name = "４条取下げ.0"
                .Nodes.Add("農地法５条").Name = "５条取下げ.0"
                .Nodes.Add("18条解約").Name = "18条解約取下げ.0"

                .Nodes.Add("基盤強化法所有権移転").Name = "所有権移転取下げ.0"
                .Nodes.Add("基盤強化法利用権設定").Name = "利用権設定取下げ.0"
                .Nodes.Add("基盤強化法利用権移転").Name = "利用権移転取下げ.0"
                .Nodes.Add("合意解約").Name = "合意解約取下げ.0"

                .Nodes.Add("あっせん申出(渡)").Name = "あっせん申出渡取下げ.0"
                .Nodes.Add("あっせん申出(受)").Name = "あっせん申出受取下げ.0"

                .Nodes.Add("農地改良届").Name = "農地改良届取下げ.0"
                .Nodes.Add("農地利用目的変更").Name = "農地利用目的変更取下げ.0"
                .Nodes.Add("農用地利用計画変更").Name = "農用地利用計画変更取下げ.0"
                .Nodes.Add("非農地証明願い").Name = "非農地証明願い取下げ.0"
                .Nodes.Add("事業計画変更").Name = "事業計画変更取下げ.0"
                '
                .Nodes.Add("買受適格-耕作-公売").Name = "買受適格-耕作-公売取下げ.0"
                .Nodes.Add("買受適格-耕作-競売").Name = "買受適格-耕作-競売取下げ.0"
                .Nodes.Add("買受適格-転用-公売").Name = "買受適格-転用-公売取下げ.0"
                .Nodes.Add("買受適格-転用-競売").Name = "買受適格-転用-競売取下げ.0"
            End With
            With .Nodes.Add("取消し")
                .Name = "取消し.0"
                .Nodes.Add("農地法３条").Name = "３条取消し.0"
                .Nodes.Add("農地法3条1項の届出").Name = "３条１項取消し.0"

                .Nodes.Add("農地法４条").Name = "４条取消し.0"
                .Nodes.Add("農地法５条").Name = "５条取消し.0"
                .Nodes.Add("18条解約").Name = "18条解約取消し.0"

                .Nodes.Add("基盤強化法所有権移転").Name = "所有権移転取消し.0"
                .Nodes.Add("基盤強化法利用権設定").Name = "利用権設定取消し.0"
                .Nodes.Add("基盤強化法利用権移転").Name = "利用権移転取消し.0"
                .Nodes.Add("合意解約").Name = "合意解約取消し.0"

                .Nodes.Add("あっせん申出(渡)").Name = "あっせん申出渡取消し.0"
                .Nodes.Add("あっせん申出(受)").Name = "あっせん申出受取消し.0"

                .Nodes.Add("農地改良届").Name = "農地改良届取消し.0"
                .Nodes.Add("農地利用目的変更").Name = "農地利用目的変更取消し.0"
                .Nodes.Add("農用地利用計画変更").Name = "農用地利用計画変更取消し.0"
                .Nodes.Add("非農地証明願い").Name = "非農地証明願い取消し.0"
                .Nodes.Add("事業計画変更").Name = "事業計画変更取消し.0"
                '
                .Nodes.Add("買受適格-耕作-公売").Name = "買受適格-耕作-公売取消し.0"
                .Nodes.Add("買受適格-耕作-競売").Name = "買受適格-耕作-競売取消し.0"
                .Nodes.Add("買受適格-転用-公売").Name = "買受適格-転用-公売取消し.0"
                .Nodes.Add("買受適格-転用-競売").Name = "買受適格-転用-競売取消し.0"
            End With
            '/*20161130 不許可処理の追加*/
            With .Nodes.Add("不許可")
                .Name = "不許可.0"
                .Nodes.Add("農地法３条").Name = "３条不許可.0"
                .Nodes.Add("農地法3条1項の届出").Name = "農地法3条1項不許可.0"

                .Nodes.Add("農地法４条").Name = "４条不許可.0"
                .Nodes.Add("農地法５条").Name = "５条不許可.0"
                .Nodes.Add("18条解約").Name = "18条解約不許可.0"

                .Nodes.Add("基盤強化法所有権移転").Name = "所有権移転不許可.0"
                .Nodes.Add("基盤強化法利用権設定").Name = "利用権設定不許可.0"
                .Nodes.Add("基盤強化法利用権移転").Name = "利用権移転不許可.0"
                .Nodes.Add("合意解約").Name = "合意解約不許可.0"

                .Nodes.Add("あっせん申出(渡)").Name = "あっせん申出渡不許可.0"
                .Nodes.Add("あっせん申出(受)").Name = "あっせん申出受不許可.0"

                .Nodes.Add("農地改良届").Name = "農地改良届不許可.0"
                .Nodes.Add("農地利用目的変更").Name = "農地利用目的変更不許可.0"
                .Nodes.Add("農用地利用計画変更").Name = "農用地利用計画変更不許可.0"
                .Nodes.Add("非農地証明願い").Name = "非農地証明願い不許可.0"
                .Nodes.Add("事業計画変更").Name = "事業計画変更不許可.0"
                '
                .Nodes.Add("買受適格-耕作-公売").Name = "買受適格-耕作-公売不許可.0"
                .Nodes.Add("買受適格-耕作-競売").Name = "買受適格-耕作-競売不許可.0"
                .Nodes.Add("買受適格-転用-公売").Name = "買受適格-転用-公売不許可.0"
                .Nodes.Add("買受適格-転用-競売").Name = "買受適格-転用-競売不許可.0"
            End With
        End With

        With mvarTreeView.Nodes.Add("終期管理")
            With .Nodes.Add("農地法終期台帳")
                .Name = "農地法終期台帳.0"
            End With

            With .Nodes.Add("利用権終期台帳")
                .Name = "利用権終期台帳.0"
            End With
            With .Nodes.Add("一時転用終期台帳")
                .Name = "一時転用終期台帳.0"
            End With
        End With
        With mvarTreeView.Nodes.Add("中間管理機構")
            .Name = "中間管理機構.0"
            If Val(SysAD.DB(sLRDB).DBProperty("中間管理機構ID")) <> 0 Then
                'With mvarTreeView.Nodes.Add("中間管理機構あっせん希望")
                '    .Name = "中間管理機構あっせん希望.0"
                'End With

                .Nodes.Add("中間管理機構を介した貸借農地").Name = "中間管理機構を介した貸借農地.1"
            End If
        End With

        With mvarTreeView.Nodes.Add("認定農業者・担い手農家管理")
            .Name = "認定・担い手管理.0"
            .Nodes.Add("認定農業者管理").Name = "農業改善計画認定.1"
            .Nodes.Add("担い手農家管理").Name = "農業改善計画認定.2"
            .Nodes.Add("農業生産法人").Name = "農業改善計画認定.3"
            .Nodes.Add("認定農業者かつ担い手農家").Name = "農業改善計画認定.4"
        End With

        With mvarTreeView.Nodes.Add("農家の取り組み管理")
            .Name = "農家の取り組み管理.0"
            .Nodes.Add("家族経営協定").Name = "家族経営協定管理.0"
            .Nodes.Add("パソコン研修").Name = "パソコン研修管理.0"
            .Nodes.Add("農業法人研修希望").Name = "農業法人研修希望.0"
        End With

        With mvarTreeView.Nodes.Add("あっせん希望")
            .Name = "あっせん希望管理.0"
            .Nodes.Add("買受希望者リスト").Name = "買受希望者リスト.0"
            .Nodes.Add("買受希望農地リスト").Name = "買受希望農地リスト.0"
            .Nodes.Add("貸付希望農地リスト").Name = "貸付希望農地リスト.0"
            .Nodes.Add("売渡希望農地リスト").Name = "売渡希望農地リスト.0"
        End With

        With New TreeNodeX("農業委員管理.0", "農業委員管理", GetType(CObj各種), mvarTreeView.Nodes)
            Dim pView As New DataView(App農地基本台帳.DataMaster.Body, "[Class]='期別農業委員会'", "ID", DataViewRowState.CurrentRows)
            For Each pRowV As DataRowView In pView
                .Nodes.Add(New TreeNodeX("期別農業委員会." & pRowV.Item("ID"), pRowV.Item("名称"), GetType(CObj各種)))
            Next
        End With

        With mvarTreeView.Nodes.Add("年金関連")
            .Nodes.Add("年金対象者一覧").Name = "年金対象者一覧.0"
            '旧制度加入者,旧制度受給者,新制度加入者,新制度受給者
            With New TreeNodeX("年金加入者.0", "年金加入者", GetType(CObj各種集計), .Nodes)
                .Nodes.Add("農業者旧制度年金加入者").Name = "農業者年金一覧.1"
                .Nodes.Add("農業者新制度年金加入者").Name = "農業者年金一覧.3"
            End With


            With New TreeNodeX("年金受給者.0", "年金受給者", GetType(CObj各種), .Nodes)
                .Nodes.Add("農業者旧制度年金受給者").Name = "農業者年金一覧.2"
                .Nodes.Add("農業者新制度年金受給者").Name = "農業者年金一覧.4"

                .Nodes.Add("経営移譲年金受給者").Name = "経営移譲年金.0"
                .Nodes.Add("老齢年金受給者").Name = "農業者老齢年金.0"
            End With
        End With

        If App農地基本台帳.TBL農地.Columns.Contains("特例農地区分") Then
            With New TreeNodeX("特例農地.0", "特例農地", GetType(CObj各種), mvarTreeView.Nodes)
                .Nodes.Add("特例対象農地").Name = "特例対象農地.0"
            End With
        End If

        With New TreeNodeX("集計機能.0", "集計機能", GetType(CObj各種集計), mvarTreeView.Nodes)
            With New TreeNodeX("農地基本.0", "農地基本", GetType(CObj各種集計), .Nodes)
                With New TreeNodeX("基本集計.0", "基本集計", GetType(CObj各種集計), .Nodes)
                    .Nodes.Add(New TreeNodeX("農家世帯数.0", "経営農家件数", GetType(CObj各種集計)))
                    .Nodes.Add(New TreeNodeX("農業者数.0", "経営農業者件数", GetType(CObj各種集計)))
                End With
                .Nodes.Add(New TreeNodeX("世帯経営面積一覧.0", "世帯経営面積一覧", GetType(CObj各種集計)))
            End With

            With New TreeNodeX("農家情報関連.0", "農家情報関連", GetType(CObj各種集計), .Nodes)
                .Nodes.Add(New TreeNodeX("経営農家一覧.0", "経営農家一覧", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("農家別面積集計.0", "農家別面積集計", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("60歳以上の農業従事者.0", "60歳以上の農業従事者", GetType(CObj各種集計)))
            End With

            With New TreeNodeX("世帯情報関連.0", "世帯情報関連", GetType(CObj各種集計), .Nodes)
                .Nodes.Add(New TreeNodeX("経営世帯一覧.0", "経営世帯一覧", GetType(CObj各種集計)))
                With New TreeNodeX("地区別世帯集計.0", "地区別世帯集計", GetType(CObj各種集計), .Nodes)
                    .Nodes.Add(New TreeNodeX("地区別世帯件数集計.0", "地区別世帯件数集計", GetType(CObj各種集計)))
                    .Nodes.Add(New TreeNodeX("地区別貸付世帯件数集計.0", "地区別貸付世帯件数集計", GetType(CObj各種集計)))
                End With
                .Nodes.Add(New TreeNodeX("集落別耕作世帯集計.0", "集落別耕作世帯集計", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("集落別農地関係世帯集計.0", "集落別農地関係世帯集計", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("経営規模別世帯集計.0", "経営規模別世帯集計", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("後継者状況別世帯集計.0", "後継者状況別世帯集計", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("規模拡大希望世帯一覧.0", "規模拡大希望世帯一覧", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("法人化希望世帯一覧.0", "法人化希望世帯一覧", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("都道府県別世帯一覧.0", "都道府県別世帯一覧", GetType(CObj各種集計)))
            End With

            With New TreeNodeX("農地情報関連.0", "農地情報関連", GetType(CObj各種集計), .Nodes)
                .Nodes.Add(New TreeNodeX("管内農地一覧.0", "管内農地一覧", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("大字別面積集計.0", "大字別面積集計", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("小字別面積集計.0", "小字別面積集計", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("登記地目別面積集計.0", "登記地目別集計", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("現況地目別面積集計.0", "現況地目別集計", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("現況地目別管内面積集計.0", "現況地目別管内面積集計", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("年齢別耕作農地一覧.0", "年齢別耕作農地一覧", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("未世帯農地一覧.0", "未世帯農地一覧", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("台帳内作成農地一覧.0", "台帳内作成農地一覧", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("登記地目が農地以外一覧.0", "登記地目が農地以外一覧", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("町外・法人所有情報.0", "町外・法人所有情報", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("死亡所有農地一覧.0", "死亡所有農地一覧", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("死亡名義人農地一覧.0", "死亡名義人農地一覧", GetType(CObj各種集計)))
            End With
            With New TreeNodeX("転用(非農地)農地情報関連.0", "転用(非農地)農地情報関連", GetType(CObj各種集計), .Nodes)
                .Nodes.Add(New TreeNodeX("転用(非農地)農地一覧.0", "転用(非農地)農地一覧", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("転用農地申請履歴一覧.0", "転用農地申請履歴一覧", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("申請無転用農地一覧.0", "申請無転用農地一覧", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("削除農地一覧.0", "削除農地一覧", GetType(CObj各種集計)))
            End With
            With New TreeNodeX("経営面積集計.0", "経営面積集計", GetType(CObj各種集計), .Nodes)
                .Nodes.Add(New TreeNodeX("経営面積農地地目別集計.0", "経営面積農地地目別集計", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("面積別集計.0", "面積別集計", GetType(CObj各種集計)))
            End With
            With New TreeNodeX("申請許可関連.0", "申請許可関連", GetType(CObj各種集計), .Nodes)
                .Nodes.Add(New TreeNodeX("年度別許可済み申請集計.0", "年度別許可済み申請集計", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("農地異動申請別一覧.0", "農地異動申請別一覧", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("農振地内異動済み農地集計.0", "農振地内異動済み農地集計", GetType(CObj各種集計)))
            End With
            With New TreeNodeX("利用権関連.0", "利用権関連", GetType(CObj各種集計), .Nodes)
                .Nodes.Add(New TreeNodeX("利用権設定始期期間別一覧.0", "利用権設定始期期間別一覧", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("利用権設定終期期間別一覧.0", "利用権設定終期期間別一覧", GetType(CObj各種集計)))
                With New TreeNodeX("申請時農地一覧.0", "申請時農地一覧", GetType(CObj各種集計), .Nodes)
                    .Nodes.Add(New TreeNodeX("申請時始期期間別一覧.0", "申請時始期期間別一覧", GetType(CObj各種集計)))
                    .Nodes.Add(New TreeNodeX("申請時終期期間別一覧.0", "申請時終期期間別一覧", GetType(CObj各種集計)))
                End With
                With New TreeNodeX("利用権設定面積集計.0", "利用権設定面積集計", GetType(CObj各種集計), .Nodes)
                    .Nodes.Add(New TreeNodeX("利用権設定大字別面積集計.0", "利用権設定大字別面積集計", GetType(CObj各種集計)))
                    With New TreeNodeX("利用権設定集落別面積集計.0", "利用権設定集落別面積集計", GetType(CObj各種集計), .Nodes)
                        .Nodes.Add(New TreeNodeX("利用権設定をした集落別面積集計.0", "利用権設定をした集落別面積集計", GetType(CObj各種集計)))
                        .Nodes.Add(New TreeNodeX("利用権設定を受けた集落別面積集計.0", "利用権設定を受けた集落別面積集計", GetType(CObj各種集計)))
                    End With

                End With
                .Nodes.Add(New TreeNodeX("大字年度毎10a当賃借料集計.0", "大字-年度毎10a当賃借料集計", GetType(CObj各種集計)))
            End With
            With New TreeNodeX("貸借関連.0", "貸借関連", GetType(CObj各種集計), .Nodes)
                .Nodes.Add(New TreeNodeX("貸借集計.0", "貸借集計", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("貸借の形態別集計.0", "貸借の形態別集計", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("貸借農地一覧.0", "貸借農地一覧", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("貸付希望地一覧.0", "貸付希望地一覧", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("売渡希望地一覧.0", "売渡希望地一覧", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("死亡借受農地一覧.0", "死亡借受農地一覧", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("貸借地所有集計.0", "貸借地所有集計", GetType(CObj各種集計)))
            End With
            With New TreeNodeX("認定農業者関連.0", "認定農業者関連", GetType(CObj各種集計), .Nodes)
                .Nodes.Add(New TreeNodeX("認定農業者期間別貸借一覧.0", "認定農業者期間別貸借一覧", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("認定農業者農地明細.0", "認定農業者農地明細", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("認定農業者別経営面積.0", "認定農業者別経営面積", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("担い手区分別経営面積.0", "担い手区分別経営面積", GetType(CObj各種集計)))
            End With
            With New TreeNodeX("選挙関連.0", "選挙関連", GetType(CObj各種集計), .Nodes)
                .Nodes.Add(New TreeNodeX("集落別選挙有資格世帯数.0", "集落別選挙有資格世帯数", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("集落別選挙有資格者数.0", "集落別選挙有資格者数", GetType(CObj各種集計)))
            End With
            With New TreeNodeX("遊休農地関連.0", "遊休農地関連", GetType(CObj各種集計), .Nodes)
                With New TreeNodeX("遊休農地一覧.0", "遊休農地一覧", GetType(CObj各種集計), .Nodes)
                    .Nodes.Add(New TreeNodeX("遊休A分類.0", "A分類", GetType(CObj各種集計)))
                    .Nodes.Add(New TreeNodeX("遊休B分類.0", "B分類", GetType(CObj各種集計)))
                    .Nodes.Add(New TreeNodeX("遊休その他.0", "その他", GetType(CObj各種集計)))
                End With
                .Nodes.Add(New TreeNodeX("遊休農地調査年度指定.0", "遊休農地調査年度指定", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("農地状況別農地一覧.0", "農地状況別農地一覧", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("大字/地区別遊休地集計.0", "大字/地区別遊休地集計", GetType(CObj各種集計)))
            End With

            With New TreeNodeX("農振農用地内一覧.0", "農振農用地内一覧", GetType(CObj各種集計), .Nodes)
                .Nodes.Add(New TreeNodeX("農用地集計.0", "農用地集計", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("農用地外集計.0", "農用地外集計", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("農振外集計.0", "農振外集計", GetType(CObj各種集計)))
            End With

            With New TreeNodeX("都市計画法関連.0", "都市計画法関連", GetType(CObj各種集計), .Nodes)
                .Nodes.Add(New TreeNodeX("都計外集計.0", "都計外集計", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("都計内集計.0", "都計内集計", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("用途地域内集計.0", "用途地域内集計", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("調整区域内集計.0", "調整区域内集計", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("市街化区域内集計.0", "市街化区域内集計", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("都市計画白地集計.0", "都市計画白地集計", GetType(CObj各種集計)))
            End With

            With New TreeNodeX("エラーチェック関連.0", "エラーチェック関連", GetType(CObj各種集計), .Nodes)
                .Nodes.Add(New TreeNodeX("不整合貸借農地一覧.0", "不整合貸借農地一覧", GetType(CObj各種集計)))
                .Nodes.Add(New TreeNodeX("面積未設定エラー.0", "面積未設定エラー", GetType(CObj各種集計)))
                With New TreeNodeX("世帯と世帯主の情報エラー.0", "世帯と世帯主の情報エラー", GetType(CObj各種集計), .Nodes)
                    .Nodes.Add(New TreeNodeX("世帯主未設定エラー.0", "世帯主未設定エラー", GetType(CObj各種集計)))
                    .Nodes.Add(New TreeNodeX("世帯未設定エラー.0", "世帯未設定エラー", GetType(CObj各種集計)))
                End With
            End With
        End With

        With mvarTreeView.Nodes.Add("プロパティ")
            .Name = "農地台帳.0"
        End With


        With mvarTreeView.Nodes.Add("マスタ管理")
            .Name = "マスタ管理.0"

            Dim result = From c In App農地基本台帳.DataMaster.Body.AsEnumerable Group By ClassName = c.Item("Class") Into ID = Min(CInt(c.Item("ID"))) Select ClassName
            For Each sName As String In result
                .Nodes.Add(New TreeNodeX("マスタクラス.""" & sName, sName, GetType(CObj各種)))
            Next
        End With
    End Sub

    Public Sub Set過去履歴(ByVal pNode As TreeNode, ByVal sKey As String, ByVal ChildKeyHead As String, n法令() As enum法令, Optional ByVal s許可 As String = "許可")
        Dim sWhere As String = ""
        For Each e法令 As Integer In n法令
            sWhere &= "," & e法令
        Next


        With pNode
            .Name = sKey
            With .Nodes.Add("過年度")
                .Name = ChildKeyHead & "過年度.0"
                Dim pView年 As New DataView(許可済み申請TBL, "[法令] IN (" & Mid(sWhere, 2) & ") AND [年]<" & Year(Now), "[年],[月]", DataViewRowState.CurrentRows)
                Dim List年 As New List(Of Integer)
                For Each pRow As DataRowView In pView年
                    If Not List年.Contains(pRow.Item("年")) Then
                        .Nodes.Add(pRow.Item("年") & "年" & s許可 & "済み").Name = ChildKeyHead & "年別." & Strings.Right("0000" & pRow.Item("年"), 4) & "00"
                        List年.Add(pRow.Item("年"))
                    End If
                Next
            End With
            Dim pView月 As New DataView(許可済み申請TBL, "[法令] IN (" & Mid(sWhere, 2) & ") AND [年]=" & Year(Now), "[月]", DataViewRowState.CurrentRows)
            For Each pRow As DataRowView In pView月
                If Not .Nodes.ContainsKey(ChildKeyHead & "月別." & Strings.Right("0000" & Now.Year, 4) & Strings.Right("00" & pRow.Item("月"), 2)) Then
                    .Nodes.Add(pRow.Item("月") & "月" & s許可 & "済み").Name = ChildKeyHead & "月別." & Strings.Right("0000" & Now.Year, 4) & Strings.Right("00" & pRow.Item("月"), 2)
                End If
            Next
        End With
    End Sub

    Public Overrides Function GetFolderObject(ByVal sKey As String) As Object
        Select Case GetKeyHead(sKey)
            '/*受付*/
            Case "受付中"
            Case "３条受付中", "４条受付中", "５条受付中" : Return New C各種申請管理(sKey)
            Case "農地法3条1項の届出" : Return New C各種申請管理(sKey)
            Case "所有権移転受付中" : Return New C各種申請管理(sKey)
            Case "利用権設定受付中", "利用権移転受付中", "18条解約受付中", "合意解約受付中" : Return New C各種申請管理(sKey)
            Case "あっせん申出渡受付中" : Return New C各種申請管理(sKey)
            Case "あっせん申出受受付中" : Return New C各種申請管理(sKey)
            Case "事業計画変更受付中" : Return New C各種申請管理(sKey)
            Case "非農地証明願い受付中" : Return New C各種申請管理(sKey)
            Case "農地改良届受付中" : Return New C各種申請管理(sKey)
            Case "農地利用目的変更受付中" : Return New C各種申請管理(sKey)
            Case "農用地利用計画変更受付中" : Return New C各種申請管理(sKey)
            Case "買受適格-耕作-公売受付中" : Return New C各種申請管理(sKey)
            Case "買受適格-耕作-競売受付中" : Return New C各種申請管理(sKey)
            Case "買受適格-転用-公売受付中" : Return New C各種申請管理(sKey)
            Case "買受適格-転用-競売受付中" : Return New C各種申請管理(sKey)
            Case "市外管理"
                '/*許可・承認*/
            Case "３条許可済", "４条許可済", "５条許可済", "18条承認済" : Return New C各種申請管理(sKey)
            Case "３条許可済年別", "３条許可済月別" : Return New C転用許可済み(sKey)
            Case "３条１項処理済年別", "３条１項処理済月別" : Return New C転用許可済み(sKey)
            Case "４条許可済年別", "４条許可済月別" : Return New C転用許可済み(sKey)
            Case "５条許可済年別", "５条許可済月別" : Return New C転用許可済み(sKey)
            Case "18条承認済年別", "18条承認済月別" : Return New C各種申請管理(sKey)
            Case "利用権設定承認済年別", "利用権設定承認済月別" : Return New C各種申請管理(sKey)
            Case "所有権移転承認済年別", "所有権移転承認済月別" : Return New C各種申請管理(sKey)
            Case "利用権移転承認済年別", "利用権移転承認済月別" : Return New C各種申請管理(sKey)
            Case "合意解約承認済年別", "合意解約承認済月別" : Return New C各種申請管理(sKey)
            Case "あっせん申出渡承認済年別", "あっせん申出渡承認済月別" : Return New C各種申請管理(sKey)
            Case "あっせん申出受承認済年別", "あっせん申出受承認済月別" : Return New C各種申請管理(sKey)
            Case "農地改良届承認済年別", "農地改良届承認済月別" : Return New C各種申請管理(sKey)
            Case "農地利用目的変更承認済年別", "農地利用目的変更承認済月別" : Return New C各種申請管理(sKey)
            Case "農用地利用計画変更承認済年別", "農用地利用計画変更承認済月別" : Return New C各種申請管理(sKey)
            Case "非農地証明願済年別", "非農地証明願済月別" : Return New C各種申請管理(sKey)
            Case "事業計画変更承認済年別", "事業計画変更承認済月別" : Return New C各種申請管理(sKey)
            Case "買受適格-耕作-公売承認済年別", "買受適格-耕作-公売承認済月別" : Return New C各種申請管理(sKey)
            Case "買受適格-耕作-競売承認済年別", "買受適格-耕作-競売承認済月別" : Return New C各種申請管理(sKey)
            Case "買受適格-転用-公売承認済年別", "買受適格-転用-公売承認済月別" : Return New C各種申請管理(sKey)
            Case "買受適格-転用-競売承認済年別", "買受適格-転用-競売承認済月別" : Return New C各種申請管理(sKey)
            Case "公共機関" : Return New CObj各種(Me.FindTree(sKey))

                '/*審査中*/
            Case "３条審査中", "４条審査中", "５条審査中" : Return New C各種申請管理(sKey)
            Case "所有権移転審査中", "利用権設定審査中", "利用権移転審査中" : Return New C各種申請管理(sKey)
            Case "あっせん申出渡審査中" : Return New C各種申請管理(sKey)
            Case "あっせん申出受審査中" : Return New C各種申請管理(sKey)
            Case "農地利用目的変更審査中" : Return New C各種申請管理(sKey)
            Case "農用地利用計画変更審査中" : Return New C各種申請管理(sKey)
            Case "事業計画変更審査中" : Return New C各種申請管理(sKey)

                '/*取り下げ*/
            Case "３条取下げ" : Return New C各種申請管理(sKey)
            Case "３条１項取下げ" : Return New C各種申請管理(sKey)
            Case "４条取下げ" : Return New C各種申請管理(sKey)
            Case "５条取下げ" : Return New C各種申請管理(sKey)
            Case "18条解約取下げ" : Return New C各種申請管理(sKey)
            Case "所有権移転取下げ" : Return New C各種申請管理(sKey)
            Case "利用権設定取下げ" : Return New C各種申請管理(sKey)
            Case "利用権移転取下げ" : Return New C各種申請管理(sKey)
            Case "合意解約取下げ" : Return New C各種申請管理(sKey)
            Case "あっせん申出渡取下げ" : Return New C各種申請管理(sKey)
            Case "あっせん申出受取下げ" : Return New C各種申請管理(sKey)
            Case "農地改良届取下げ" : Return New C各種申請管理(sKey)
            Case "農地利用目的変更取下げ" : Return New C各種申請管理(sKey)
            Case "農用地利用計画変更取下げ" : Return New C各種申請管理(sKey)
            Case "非農地証明願い取下げ" : Return New C各種申請管理(sKey)
            Case "事業計画変更取下げ" : Return New C各種申請管理(sKey)
            Case "買受適格-耕作-公売取下げ" : Return New C各種申請管理(sKey)
            Case "買受適格-耕作-競売取下げ" : Return New C各種申請管理(sKey)
            Case "買受適格-転用-公売取下げ" : Return New C各種申請管理(sKey)
            Case "買受適格-転用-競売取下げ" : Return New C各種申請管理(sKey)

                '/*取り消し*/
            Case "３条取消し" : Return New C各種申請管理(sKey)
            Case "３条１項取消し" : Return New C各種申請管理(sKey)
            Case "４条取消し" : Return New C各種申請管理(sKey)
            Case "５条取消し" : Return New C各種申請管理(sKey)
            Case "18条解約取消し" : Return New C各種申請管理(sKey)
            Case "所有権移転取消し" : Return New C各種申請管理(sKey)
            Case "利用権設定取消し" : Return New C各種申請管理(sKey)
            Case "利用権移転取消し" : Return New C各種申請管理(sKey)
            Case "合意解約取消し" : Return New C各種申請管理(sKey)
            Case "あっせん申出渡取消し" : Return New C各種申請管理(sKey)
            Case "あっせん申出受取消し" : Return New C各種申請管理(sKey)
            Case "農地改良届取消し" : Return New C各種申請管理(sKey)
            Case "農地利用目的変更取消し" : Return New C各種申請管理(sKey)
            Case "農用地利用計画変更取消し" : Return New C各種申請管理(sKey)
            Case "非農地証明願い取消し" : Return New C各種申請管理(sKey)
            Case "事業計画変更取消し" : Return New C各種申請管理(sKey)
            Case "買受適格-耕作-公売取消し" : Return New C各種申請管理(sKey)
            Case "買受適格-耕作-競売取消し" : Return New C各種申請管理(sKey)
            Case "買受適格-転用-公売取消し" : Return New C各種申請管理(sKey)
            Case "買受適格-転用-競売取消し" : Return New C各種申請管理(sKey)

                '/*20161130 不許可処理の追加*/
            Case "３条不許可" : Return New C各種申請管理(sKey)
            Case "農地法3条1項不許可" : Return New C各種申請管理(sKey)
            Case "４条不許可" : Return New C各種申請管理(sKey)
            Case "５条不許可" : Return New C各種申請管理(sKey)
            Case "18条解約不許可" : Return New C各種申請管理(sKey)
            Case "所有権移転不許可" : Return New C各種申請管理(sKey)
            Case "利用権設定不許可" : Return New C各種申請管理(sKey)
            Case "利用権移転不許可" : Return New C各種申請管理(sKey)
            Case "合意解約不許可" : Return New C各種申請管理(sKey)
            Case "あっせん申出渡不許可" : Return New C各種申請管理(sKey)
            Case "あっせん申出受不許可" : Return New C各種申請管理(sKey)
            Case "農地改良届不許可" : Return New C各種申請管理(sKey)
            Case "農地利用目的変更不許可" : Return New C各種申請管理(sKey)
            Case "農用地利用計画変更不許可" : Return New C各種申請管理(sKey)
            Case "非農地証明願い不許可" : Return New C各種申請管理(sKey)
            Case "事業計画変更不許可" : Return New C各種申請管理(sKey)

            Case "買受適格-耕作-公売不許可" : Return New C各種申請管理(sKey)
            Case "買受適格-耕作-競売不許可" : Return New C各種申請管理(sKey)
            Case "買受適格-転用-公売不許可" : Return New C各種申請管理(sKey)
            Case "買受適格-転用-公売不許可" : Return New C各種申請管理(sKey)

            Case "行政区" : Return ObjectMan.GetObject(sKey)
            Case "あっせん希望管理", "認定・担い手管理" : Return Nothing
            Case Else
                If Not SysAD.IsClickOnceDeployed Then
                    'Stop
                End If
                Return ObjectMan.GetObject(sKey)
        End Select
        Return Nothing
    End Function

    Public Overrides ReadOnly Property ObjectMan As HimTools2012.TargetSystem.CObjectManSK
        Get
            Return modCommon.ObjectMan
        End Get
    End Property
End Class




