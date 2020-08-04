
Imports System.ComponentModel
Imports HimTools2012.controls
Imports HimTools2012.controls.DVCtrlCommonBase

Public Class DataViewNext農地
    Inherits CDataViewPanel農地台帳


    Private WithEvents Cmb自小作別 As ComboBoxPlus
    Private WithEvents Cmb人農地区分 As ComboBoxPlus
    Private WithEvents Cmb人農地貸付 As ComboBoxPlus
    Private mvarGroup As HimTools2012.controls.GroupBoxPlus
    Private mvarGroup転貸 As HimTools2012.controls.GroupBoxPlus
    Private mvarGroup転用 As HimTools2012.controls.GroupBoxPlus

    Public Sub New(ByVal pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget, App農地基本台帳.TBL農地, SysAD.page農家世帯.DataViewCollection, True, True)

        'LoadXML(My.Resources.Resource1._Interface, "D:農地Info", App農地基本台帳.DSet)
        Dim nID As Integer = pTarget.ID
        Me.SetButtons(CreateButton("所有者を呼ぶ", "所有者を呼ぶ"), New ToolStripSeparator)

        'Dim b新項目保存 As Boolean = Not pTarget.Row.Body.Table.Columns.Contains("耕地番号作成日")

        Dim nHeight As Integer = 0
        Panel.FlowDirection = FlowDirection.LeftToRight
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("管理情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            'pTBL.Rows.Add(.Text, .Expanded, .Body)
            nHeight = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsReadOnly, pTarget, "ID", , 80), "ID").Height
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsReadOnly, pTarget, "一筆コード", , 80), "一筆コード", em改行.改行なし)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "耕地番号", , 80), "耕地番号", em改行.改行なし)
            .Panel.AddCtrl(New DateTimePickerPlus(b新項目保存(pTarget, "耕地番号作成日"), pTarget, "耕地番号作成日"), "耕地番号作成日", em改行.改行あり)
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("先行異動", "Value"), "あり", "なし"), "先行異動")
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "先行異動日", 150), "異動日")
            .Panel.AddCtrl(New DateTimePickerPlus(True, pTarget, "更新日", 150), "更新日", em改行.改行あり)
            .Panel.FitLabelWidth()
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("(1・2・3)農地基本", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            'pTBL.Rows.Add(.Text, .Expanded, .Body)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "市町村ID", , 80), "市町村コード", em改行.改行あり)
            SetTextComboWithMaster(.Panel, "大字", "大字ID", "大　字", em改行.改行なし)
            SetTextComboWithMaster(.Panel, "小字", "小字ID", "小　字", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "所在", , 200), "所在(町外)", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "地番", , 100), "地  番")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "一部現況", , 40), "一部現況", em改行.改行なし)
            If App農地基本台帳.TBL農地.Columns.Contains("特例農地区分") Then
                .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("特例農地区分", "Value"), "有", "無"), "特例農地区分", em改行.改行なし)
                .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, pTarget, "特例農地指定日", 150), "特例農地指定日", em改行.改行なし)
                .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, pTarget, "特例農地指定除外日", 150), "特例農地指定除外日", em改行.改行あり)
            End If
            SetTextComboWithMaster(.Panel, "地目", "登記簿地目", "登記簿地目", em改行.改行なし)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Real, emRO.IsCanEdit, pTarget, "登記簿面積", , 100), "登記簿面積", em改行.改行なし)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "共有持分分子", , 40), "共有持分")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "共有持分分母", , 40), "/", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("課税地目"), "名称", "ID",
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "現況地目", , 60), "現況地目")
            ))
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Real, emRO.IsCanEdit, pTarget, "実面積", , 100), "現況面積", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Real, emRO.IsCanEdit, pTarget, "本地面積", , 100), "本地面積", em改行.改行なし)
            .Panel.AddCtrl(New DateTimePickerPlus(b新項目保存(pTarget, "本地面積作成日"), pTarget, "本地面積作成日", 150), "本地面積作成日", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("農委地目"), "名称", "ID",
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "農委地目ID", , 60), "農委地目")
            ))
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "農委地目認定日", 150), "認定日", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Real, emRO.IsCanEdit, pTarget, "部分面積", , 100), "部分面積", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Real, emRO.IsCanEdit, pTarget, "田面積", , 100), "田面積")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Real, emRO.IsCanEdit, pTarget, "畑面積", , 100), "畑面積")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Real, emRO.IsCanEdit, pTarget, "樹園地", , 100), "樹園地")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Real, emRO.IsCanEdit, pTarget, "採草放牧面積", , 100), "採草放牧面積")
            .Panel.FitLabelWidth()
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("(4・5)農地区分情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            '20150130_CSV必須項目
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,農用地区域,農振地域,農振地域外,その他,調査中", ","), nHeight, pTarget, "農振法区分"), "農振法区分", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,市街化区域,市街化調整区域,非線引き都市計画区域の「用途地域」,都市計画区域外,その他,調査中,非線引き都市計画区域内", ","), nHeight, pTarget, "都市計画法区分"), "都市計画区分", em改行.改行あり)
            '20150302_入力画面追加
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("生産緑地法", "Value"), "あり", "なし"), "生産緑地法", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,新生産緑地指定,旧長期営農継続農地制度認定,旧第一種生産緑地指定,旧第二種生産緑地指定", ","), nHeight, pTarget, "生産緑地法種別"), "生産緑地法種別", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "生産緑地法指定日", 150), "緑地法指定日", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Real, b新項目保存(pTarget, "生産緑地法指定面積"), pTarget, "生産緑地法指定面積", , 100), "生産緑地法指定面積")
            .Panel.AddCtrl(New OptionButtonPlus(Split("区域外,区域内(整備済),区域内(整備中),換地,その他", ","), nHeight, pTarget, "土地改良法"), "土地改良区分", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("区域外,区域内(整備済),区域内(整備中),換地,その他", ","), nHeight, pTarget, "区画整理"), "区画整理", em改行.改行あり) '【保留】ReadOnly
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "等級", , 200), "農地区分", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,甲種,第１種,第２種,第３種", ","), nHeight, pTarget, "農地種別"), "農地種別", em改行.改行あり) '【保留】ReadOnly
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("農地状況"), "名称", "ID",
               .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "農地状況", , 40), "農地状況"), , 200
           ))
            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, pTarget, "無断転用調査日", 150), "調査日", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("耕作状況"), "名称", "ID",
               .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "耕作状況"), pTarget, "耕作状況", , 40), "耕作状況"), b新項目保存(pTarget, "耕作状況"), 200
          ))
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "利用状況", , 200), "利用状況", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "作付け作物", , 200), "作付け作物", em改行.改行あり)
            .Panel.FitLabelWidth()
        End With

        '20150305_入力画面追加
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("(6)所有情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            'pTBL.Rows.Add(.Text, .Expanded, .Body)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "所有世帯ID", , 100), "所有世帯")
            SetInputIDandText(.Panel, "所有者名", "所有者ID", "所有者氏名", emDrop可能.可能, "所有者参照", "所有者を呼ぶ", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "管理世帯ID", , 100), "農地所有世帯")
            SetInputIDandText(.Panel, "農地所有者/管理者", "管理者ID", "管理者氏名", emDrop可能.可能, "農地所有者", "管理者を呼ぶ", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,管理人,代理人,変更済", ","), nHeight, pTarget, "農地所有内訳"), "農地所有者/管理者内訳", em改行.改行あり)
            '耕作者整理番号も↓に
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "登記名義人ID", , 100, , True), "登記名義", em改行.改行なし)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "名義人氏名", , 250), "", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "登記名義人氏名", , 250), "登記名義直接入力", em改行.改行あり)

            '20150130_CSV必須項目
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,所有権移転,貸し付け,人・農地プランへの位置づけ,農地中間管理機構への貸付の申出,自ら耕作,その他,調査中", ","), nHeight, pTarget, "所有者農地意向"), "所有者農地意向", em改行.改行あり)
            '↑今回　↓前回　変換必要
            '.Panel.AddCtrl(New OptionButtonPlus(Split("-,所有権移転,貸付,人・農地プランへの位置づけ,農地中間管理機構への貸付,その他,自ら耕作する", ","), nHeight, pTarget, "所有者農地意向"), "所有者農地意向", em改行.改行あり)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("所有者農地意向その他", "Text"), 400), "「その他」の内訳", em改行.改行あり) '【保留】ReadOnly
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("農地法第52公表同意", "Value"), "あり", "なし"), "農52条公表同意")
            .Panel.FitLabelWidth()
        End With

        '20150302_入力画面追加
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("(7)共有農地情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            'pTBL.Rows.Add(.Text, .Expanded, .Body)
            '.Panel.AddCtrl(New OptionButtonPlus(Split("-,個人所有地,共有地,その他", ","), nHeight, pTarget, "共有地区分", "Value"), "共有地判定", em改行.改行あり)
            '.Panel.AddCtrl(New  TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "共有者ID", , 100, , True), "共有者名", em改行.改行なし)
            '.Panel.AddCtrl(New  TextBoxPlus(TextBoxMode.Free, emRO.IsReadOnly, pTarget, "共有者氏名", , 200, , True), "", em改行.改行なし)
            '.Panel.AddCtrl(New  TextBoxPlus(TextBoxMode.Free, emRO.IsReadOnly, pTarget, "共有者住所", , 250), "", em改行.改行あり)
        End With

        Cmb自小作別 = Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("自小作区分"), "名称", "ID",
            Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "自小作別", , 60), "自小作の別")
        ), , em改行.改行あり)

        mvarGroup = CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("(8・9・10)借入地の状況", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
        With mvarGroup
            'pTBL.Rows.Add(.Text, .Expanded, .Body)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "借受世帯ID", , 100), "借受世帯番号")
            SetInputIDandText(.Panel, "借受者名", "借受人ID", "借受人氏名", emDrop可能.可能, "借受者参照", "借受者を呼ぶ", em改行.改行あり)
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("農業生産法人経由貸借", "Value"), "あり", "なし"), "法人経由の貸借")
            SetInputIDandText(.Panel, "経由法人名", "経由農業生産法人ID", "経由農業生産法人名", emDrop可能.可能, "法人参照", "経由法人を呼ぶ", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("適用法令"), "名称", "ID",
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "小作地適用法", , 60), "適用法令")
            ), , em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("小作形態"), "名称", "ID",
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "小作形態", , 60), "形態")
            ), , em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "小作料", , 100), "貸借料")
            .Panel.AddCtrl(New ComboList(ListResource.S_Data, "小作料単位", mvarTarget, "小作料単位", "Text", ComboBoxStyle.DropDown), "貸借料単位", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "10a賃借料", , 100), "10アール貸借料", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, b新項目保存(pTarget, "物納"), pTarget, "物納", , 100), "物納")
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("物納単位"), "名称", "ID",
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "物納単位"), pTarget, "物納単位", , 60), "物納単位", b新項目保存(pTarget, "物納単位"))
            ), , em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "小作開始年月日", 150), "貸借期間")
            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("～", "期間設定", True))
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "小作終了年月日", 150), "", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(b新項目保存(pTarget, "再設定終了年月日"), pTarget, "再設定終了年月日", 150), "再設定終期年月日", em改行.改行あり)
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("解除条件付きの農地の貸借", "Value"), "あり", "なし"), "条件付き貸借")
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("前払いの有無", "Value"), "あり", "なし"), "賃借料一括前払い") '【保留】ReadOnly

            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "利用集積計画番号"), pTarget, "利用集積計画番号", , 100), "利用集積計画番号", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(b新項目保存(pTarget, "利用集積公告日"), pTarget, "利用集積公告日", 150), "利用集積公告日", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,田,畑,その他", ","), nHeight, mvarTarget, "利用目的"), "利用目的", em改行.改行あり) '【保留】ReadOnly
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("利用目的備考", "Text"), 400), "利用目的備考", em改行.改行あり) '【保留】ReadOnly
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,新規設定,再設定", ","), nHeight, mvarTarget, "利用権設定区分"), "利用権設定区分", em改行.改行あり) '【保留】ReadOnly
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,交付金対象,交付金対象外", ","), nHeight, mvarTarget, "交付金判定"), "交付金判定", em改行.改行あり) '【保留】ReadOnly
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "交付金対象額"), pTarget, "交付金対象額", , 100), "交付金対象額", em改行.改行あり)

            mvarGroup転貸 = CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("転貸地の状況", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            With mvarGroup転貸
                .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("適用法令"), "名称", "ID",
                    .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "転貸適用法"), pTarget, "転貸適用法", , 60), "適用法令")
                ), , em改行.改行あり)
                .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("小作形態"), "名称", "ID",
                    .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "転貸形態"), pTarget, "転貸形態", , 60), "転貸権利の種類")
                ), , em改行.改行あり)
                .Panel.AddCtrl(New DateTimePickerPlus(b新項目保存(pTarget, "転貸始期年月日"), pTarget, "転貸始期年月日", 150), "貸借期間")
                .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("～", "転貸期間設定", True))
                .Panel.AddCtrl(New DateTimePickerPlus(b新項目保存(pTarget, "転貸終期年月日"), pTarget, "転貸終期年月日", 150), "", em改行.改行あり)
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "転貸料"), pTarget, "転貸料", , 100), "借賃額")
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "転貸10a転貸料"), pTarget, "転貸10a転貸料", , 100), "転貸10a転貸料", em改行.改行あり)
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, b新項目保存(pTarget, "転貸物納"), pTarget, "転貸物納", , 100), "物納")
                .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("物納単位"), "名称", "ID",
                    .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "転貸料単位"), pTarget, "転貸料単位", , 60), "物納単位")
                ), , em改行.改行あり)
            End With

            mvarGroup転用 = CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("転用状況", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            With mvarGroup転用
                .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("転用適用法"), "名称", "ID",
                    .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "転用適用法"), pTarget, "転用適用法", , 60), "転用適用法")
                ), , em改行.改行あり)
                .Panel.AddCtrl(New OptionButtonPlus(Split("-,転用,一時転用,非農地判断,その他", ","), nHeight, mvarTarget, "転用形態"), "転用形態", em改行.改行あり) '【保留】ReadOnly
                .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("調査転用用途"), "名称", "ID",
                    .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "転用用途"), pTarget, "転用用途", , 60), "転用用途")
                ), , em改行.改行あり)
                .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("転用換地有無", "Value"), "あり", "なし"), "転用換地の有無", em改行.改行あり) '【保留】ReadOnly
                .Panel.AddCtrl(New DateTimePickerPlus(b新項目保存(pTarget, "転用始期年月日"), pTarget, "転用始期年月日", 150), "貸借期間")
                .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("～", "転用期間設定", True))
                .Panel.AddCtrl(New DateTimePickerPlus(b新項目保存(pTarget, "転用終期年月日"), pTarget, "転用終期年月日", 150), "", em改行.改行あり)
            End With

            .Panel.FitLabelWidth()
        End With

        '20150304_入力画面追加
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("(11)特定作業受委託", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            '.Panel.AddCtrl(New  TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "特定作業者ID", , 100, , True), "特定作業者名", em改行.改行なし)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "特定作業者名", , 200, , True), "特定作業者名", em改行.改行なし)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "特定作業者住所", , 250), "特定作業者住所", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "特定作業作目種別", , 200), "特定作業作物")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "特定作業内容", , 200), "特定作業内容")
            .Panel.FitLabelWidth()
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("(12・13)利用状況報告", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            'pTBL.Rows.Add(.Text, .Expanded, .Body)
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("利用状況報告対象", "Value"), "あり", "なし"), "利用状況報告")
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "利用状況報告年月日", 150), "報告年月日", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "是正勧告日", 150), "勧告年月日", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "是正内容", , 200), "是正内容", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "是正期限", 150), "是正期限", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,１号,２号,３号", ","), nHeight, mvarTarget, "根拠条件農地法"), "根拠農地法", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,１号,２号,３号", ","), nHeight, mvarTarget, "根拠条件基盤強化法"), "根拠基盤法", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "是正状況", , 200), "是正状況", em改行.改行なし)
            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, pTarget, "是正確認", 150), "是正確認日", em改行.改行あり)

        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("(14)許可の取消に関する事項", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, pTarget, "取消年月日", 150), "取消年月日", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "取消事由", , 200), "取消事由", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,１号,２号,３号", ","), nHeight, mvarTarget, "取消条件農地法"), "取消農地法", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,１号,２号,３号", ","), nHeight, mvarTarget, "取消条件基盤強化法"), "取消基盤法", em改行.改行あり)
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("(15)相続等の届出", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "届出年月日", 150), "届出年月日", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("届出事由"), "名称", "ID",
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "届出事由", , 60), "届出事由")
            ), , em改行.改行あり)

            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, b新項目保存(pTarget, "相続届出者ID"), pTarget, "相続届出者ID", , 60), "権利取得者", em改行.改行なし)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "届出者氏名", , 200), "", em改行.改行あり)

            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("相続登記の有無", "Value"), "あり", "なし"), "相続登記の有無", em改行.改行あり) '【保留】ReadOnly
            SetInputIDandText(.Panel, "耕作しているであろう者", "推測耕作者ID", "相続者名", emDrop可能.可能, , , em改行.改行あり)
            .Panel.FitLabelWidth()
        End With


        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("(16)農地の利用状況調査", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New DateTimePickerPlus(b新項目保存(pTarget, "利用状況調査日"), pTarget, "利用状況調査日", 150), "調査年月日", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,１号,２号,遊休農地でない,その他(非農地判定予定等),立入困難等外因的理由で調査不可,調査中", ","), nHeight, mvarTarget, "利用状況調査農地法"), "農法第32条第1項", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(b新項目保存(pTarget, "利用状況調査不可判断日"), pTarget, "利用状況調査不可判断日", 150), "調査不可判断日", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,被災して農業利用が出来ない,災害や草木類の繁茂等により進入路が荒廃,その他", ","), nHeight, mvarTarget, "利用状況調査不可判断理由"), "調査不可判断理由", em改行.改行あり) '【保留】ReadOnly
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("利用状況調査不可判断その他理由", "Text"), 400), "判断理由「その他」内訳", em改行.改行あり) '【保留】ReadOnly
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,A分類,B分類,調査中", ","), nHeight, mvarTarget, "利用状況調査荒廃"), "荒廃農地調査", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,山林,原野,宅地,雑種地,その他", ","), nHeight, mvarTarget, "利用状況調査荒廃内訳"), "B分類内訳", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("農業委員", "([ParentKey]='" & SysAD.DB(sLRDB).DBProperty("今期農業委員会Key", "") & "' Or [ID]=0 Or [ID]=1)", "[ID] "), "名称", "ID",
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "利用状況調査委員ID"), mvarTarget, "利用状況調査委員ID"), "調査委員名"), False
            ), "", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, b新項目保存(pTarget, "利用状況耕作放棄地通し番号"), pTarget, "利用状況耕作放棄地通し番号", , 100), "耕作放棄地通し番号", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,一時転用,無断転用,違反転用", ","), nHeight, mvarTarget, "利用状況調査転用"), "調査結果転用", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,砂利採取,太陽光発電施設,その他,調査中", ","), nHeight, mvarTarget, "利用状況一時転用区分"), "一時転用区分", em改行.改行あり) '【保留】ReadOnly
            .Panel.FitLabelWidth()
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("(17)農地の利用意向調査", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "利用意向調査日", 150), "調査年月日", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,農地法第32条第１項,農地法第32条第４項,農地法第33条第１項,調査中", ","), nHeight, mvarTarget, "利用意向根拠条項"), "根拠条項", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "利用意向意思表明日", 150), "意思表明年月日", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,自ら耕作,機構事業,所有者代理事業,権利設定または転移,その他,調査中", ","), nHeight, mvarTarget, "利用意向意向内容区分"), "調査結果", em改行.改行あり)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("利用意向調査結果その他理由", "Text"), 400), "調査結果「その他」内訳", em改行.改行あり) '【保留】ReadOnly
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("利用意向措置実施状況", "Text"), 400), "措置実施状況", em改行.改行あり) '【保留】ReadOnly
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,対象外,調査中,調査済み", ","), nHeight, mvarTarget, "利用意向権利関係調査区分"), "所有者不明権利関係調査", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(b新項目保存(pTarget, "利用意向調査不可年月日"), pTarget, "利用意向調査不可年月日", 150), "所有者不明調査結果年月日", em改行.改行なし)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("所有者不明結果"), "名称", "ID",
                    .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "利用意向調査不可結果"), pTarget, "利用意向調査不可結果", , 60), "所有者不明調査結果"), , 200
                ), , em改行.改行あり)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("利用意向権利関係調査記録", "Text"), 400), "調査結果「その他」内訳", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "利用意向公示年月日", 150), "農32条公示日", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "利用意向通知年月日", 150), "農43条通知日", em改行.改行あり)
            .Panel.FitLabelWidth()
        End With

        '20150130_CSV必須項目
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("(18)農地中間管理機構との協議", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "農地法35の1通知日", 150), "農35条の1通知日", em改行.改行なし)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "農地法35の2通知日", 150), "農35条の2通知日", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(b新項目保存(pTarget, "農地法35の2申入日"), pTarget, "農地法35の2申入日", 150), "農35条の2申入日", em改行.改行なし)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "農地法35の3通知日", 150), "農35条の3通知日", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "勧告年月日", 150), "農36条の1勧告日", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,１号,２号,３号,４号,５号,調査中", ","), nHeight, mvarTarget, "勧告内容"), "勧告内容", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "中間管理勧告日", 150), "勧告通知日", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,機構法第20条,農地法第35条,農地法第37条,災害,農地法第34条,その他,調査中", ","), nHeight, mvarTarget, "再生利用困難農地"), "再生困難農地", em改行.改行あり)
        End With
        '20150130_CSV必須項目
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("(19・20)裁定－措置命令", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            'pTBL.Rows.Add(.Text, .Expanded, .Body)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "農地法40裁定公告日", 150), "農40条裁定日", em改行.改行なし)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "農地法43裁定公告日", 150), "農43条裁定日", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "農地法44の1裁定公告日", 150), "農44条の1裁定日", em改行.改行なし)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "農地法44の3裁定公告日", 150), "農44条の3裁定日", em改行.改行あり)
            .Panel.FitLabelWidth()
        End With
        '20150130_CSV必須項目
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("(21)農地中間管理と農用地利用配分計画", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            'pTBL.Rows.Add(.Text, .Expanded, .Body)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "中間管理権取得日", 150), "中間管理取得日", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "意見回答日", 150), "意見回答年月日", em改行.改行なし)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "知事公告日", 150), "知事公告年月日", em改行.改行なし)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "認可通知日", 150), "認可通知年月日", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,使用貸借権,賃貸借権", ","), nHeight, mvarTarget, "権利設定内容"), "権利設定内容", em改行.改行あり)
            '      .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "利用配分設定期間",, 150), "利用配分計画", em改行.改行なし)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "利用配分計画始期日", 150), "存続期間", em改行.改行なし)
            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("～", "利用配分期間設定", True))
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "利用配分計画終期日", 150), "", em改行.改行あり)
            '20150303_入力画面追加
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "利用配分計画借賃額", , 200), "1年間の借賃額", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "利用配分計画10a賃借料", , 200), "10a当り借賃額", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "貸借契約解除年月日", 150), "貸借解除年月日", em改行.改行あり)
            .Panel.FitLabelWidth()
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("(22)納税猶予の適用状況", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            'pTBL.Rows.Add(.Text, .Expanded, .Body)
            .Panel.AddCtrl(New OptionButtonPlus(Split("対象外,生前贈与の納税贈与,相続税の猶予,調査中", ","), nHeight, pTarget, "納税猶予対象農地"), "納税猶予対象", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,特例適用外,対象外,調査中", ","), nHeight, mvarTarget, "納税猶予種別"), "納税猶予種別", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "納税猶予相続日", 150), "納税猶予相続日", em改行.改行なし)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "納税猶予適用日", 150), "納税猶予適用日", em改行.改行なし)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "納税猶予継続日", 150), "納税猶予継続日", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(b新項目保存(pTarget, "納税猶予確認日"), pTarget, "納税猶予確認日", 150), "納税猶予確認日", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("対象外,農地取得資金,農地等購入資金", ","), nHeight, pTarget, "融資対象農地"), "融資対象農地", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,１号,２号,３号,調査中", ","), nHeight, pTarget, "租税処置法"), "特定貸付根拠条項", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("対象外,対象,調査中", ","), nHeight, pTarget, "営農困難"), "営農困難貸付", em改行.改行あり)
            '.Panel.AddCtrl(New  TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "営農困難", , 200), "営農困難貸付", em改行.改行あり)
            .Panel.FitLabelWidth()
        End With


        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("(23)仮登記の設定状況", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "仮登記日", 150), "仮登記日", em改行.改行あり)

            .Panel.AddCtrl(New ComboBoxPlus(New DataView(App農地基本台帳.TBL個人.Body, "[ID]=" & Val(pTarget.Row.Item("相続届出者ID").ToString), "[ID]", DataViewRowState.CurrentRows), "氏名", "ID",
                 .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "仮登記者ID"), pTarget, "仮登記者ID", , 60), "仮登記者")
            ), , em改行.改行あり).Width = 200

            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "仮登記住所", , 200), "仮登記者住所", em改行.改行あり)
        End With

        '20150304_入力画面追加
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("(24)各種交付金、補助金の支援状況", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            'pTBL.Rows.Add(.Text, .Expanded, .Body)
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("環境保全交付金", "Value"), "あり", "なし"), "環境保全交付金")
            .Panel.AddCtrl(New DateTimePickerPlus(b新項目保存(pTarget, "環境保全交付基準日"), pTarget, "環境保全交付基準日", 150), "環境保全交付基準日", em改行.改行あり)
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("農地維持交付金", "Value"), "あり", "なし"), "農地維持交付金")
            .Panel.AddCtrl(New DateTimePickerPlus(b新項目保存(pTarget, "農地維持交付基準日"), pTarget, "農地維持交付基準日", 150), "農地維持交付基準日", em改行.改行あり)
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("資源向上交付金", "Value"), "あり", "なし"), "資源向上交付金")
            .Panel.AddCtrl(New DateTimePickerPlus(b新項目保存(pTarget, "資源向上交付基準日"), pTarget, "資源向上交付基準日", 150), "資源向上交付基準日", em改行.改行あり)
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("中山間直接支払", "Value"), "あり", "なし"), "中山間直接支払")
            .Panel.AddCtrl(New DateTimePickerPlus(b新項目保存(pTarget, "中山間直接支払基準日"), pTarget, "中山間直接支払基準日", 150), "中山間直接支払基準日", em改行.改行あり)
            .Panel.FitLabelWidth()
        End With
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("(25)特定処分対象農地", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("特定処分対象農地等", "Value"), "あり", "なし"), "特定処分対象農地")
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,第1種加算,第2種加算,第3種加算", ","), nHeight, pTarget, "農業者年金処分対象農地"), "農年処分対象農地", em改行.改行あり) '【保留】ReadOnly
            .Panel.AddCtrl(New DateTimePickerPlus(b新項目保存(pTarget, "農業者年金処分適用日"), pTarget, "農業者年金処分適用日", 150), "農年処分適用日", em改行.改行あり)
            .Panel.FitLabelWidth()
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("その他", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            'pTBL.Rows.Add(.Text, .Expanded, .Body)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("備考", "Text"), 400), "備考", em改行.改行あり)

            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("あっせん希望", "Value"), "あり", "なし"), "あっせん希望")
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("貸付希望", "Value"), "あり", "なし"), "貸付希望")
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("売渡希望", "Value"), "あり", "なし"), "売渡希望", em改行.改行あり)
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("指定の有無", "Value"), "あり", "なし"), "一時利用地指定", em改行.改行あり) '【保留】ReadOnly
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("土地利用計画区域地目区分"), "名称", "ID",
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "調査土地利用区域地目", , 60), "土地利用区域地目")
            ), , em改行.改行あり)

            Cmb人農地区分 = .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("人農地プラン区分"), "名称", "ID",
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "人農地プラン区分", , 60), "人農地プラン区分")
        ), , em改行.改行なし)

            Cmb人農地貸付 = .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("人農地プラン貸付内訳"), "名称", "ID",
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "人農地プラン貸付内訳", , 60), "人農地プラン貸付内訳")
        ), , em改行.改行あり)
            Cmb人農地貸付.Enabled = False

            .Panel.FitLabelWidth()
        End With
    End Sub

    Private Sub Cmb自小作別_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Cmb自小作別.SelectedValueChanged
        Select Case Cmb自小作別.Text
            Case "自作"
                mvarGroup.Visible = False
                mvarGroup転貸.Visible = False
                mvarGroup転用.Visible = False
            Case Else
                mvarGroup.Visible = True
                mvarGroup転貸.Visible = True
                mvarGroup転用.Visible = True
        End Select
    End Sub

    Private Sub Cmb人農地区分_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Cmb人農地区分.SelectedValueChanged
        Select Case Cmb人農地区分.Text
            Case "貸付"
                Cmb人農地貸付.Enabled = True
            Case Else
                Cmb人農地貸付.Enabled = False
        End Select
    End Sub
End Class

Public Class ToolStripCheckCombo
    Inherits ToolStripControlHost

    Public Sub New()
        MyBase.New(New CheckedComboBox())
    End Sub

    Public Property Datasource() As Object
        Get
            Return DirectCast(Control, CheckedComboBox).ComboSource
        End Get
        Set(value As Object)
            DirectCast(Control, CheckedComboBox).ComboSource = value
        End Set
    End Property

    Public Property DisplayMember As String
        Get
            Return DirectCast(Control, CheckedComboBox).DisplayMember
        End Get
        Set(value As String)
            DirectCast(Control, CheckedComboBox).DisplayMember = value
        End Set
    End Property

    Public ReadOnly Property Items As CheckedListBox.ObjectCollection
        Get
            Return DirectCast(Control, CheckedComboBox).Items
        End Get
    End Property
    Public Property SelectValueMember As String
        Get
            Return DirectCast(Control, CheckedComboBox).SelectedValue
        End Get
        Set(value As String)
            DirectCast(Control, CheckedComboBox).SelectedValue = value
        End Set
    End Property
    Public Property ValueMember As String
        Get
            Return DirectCast(Control, CheckedComboBox).ValueMember
        End Get
        Set(value As String)
            DirectCast(Control, CheckedComboBox).ValueMember = value
        End Set
    End Property
End Class

''' <summary>
''' チェック付きコンボボックス
''' </summary>
''' <remarks></remarks>
Public Class CheckedComboBox
    Inherits ComboBox

    'Imports System.Collections
    'Imports System.Collections.Generic
    'Imports System.ComponentModel
    'Imports System.Text
    'Imports System.Windows.Forms
    'Imports System.Drawing
    'Imports System.Diagnostics

#Region "内部クラス"
    ''' <summary>
    ''' <see cref="CheckedComboBox">CheckedComboBox</see> のドロップダウンを表す内部クラス
    ''' </summary>
    Protected Friend Class CheckedComboBoxDropdown
        Inherits Form

#Region "内部クラス"
        ''' <summary>
        ''' カスタムチェックリストボックス
        ''' </summary>
        Protected Friend Class CustomCheckedListBox
            Inherits CheckedListBox

            Private m_curSelIndex As Integer = -1

            ''' <summary>
            ''' コンストラクタ
            ''' </summary>
            Public Sub New()
                MyBase.New()
                Me.SelectionMode = SelectionMode.One
                Me.HorizontalScrollbar = True
            End Sub

            ''' <summary>
            ''' キーダウン時のイベントハンドラ
            ''' </summary>
            Protected Overrides Sub OnKeyDown(ByVal e As KeyEventArgs)
                Select Case e.KeyCode

                    Case Keys.Enter
                        DirectCast(Parent, CheckedComboBoxDropdown).OnDeactivate(New CheckedComboBoxEventArgs(EventArgs.Empty, True))
                        e.Handled = True
                    Case Keys.Escape
                        DirectCast(Parent, CheckedComboBoxDropdown).OnDeactivate(New CheckedComboBoxEventArgs(EventArgs.Empty, False))
                        e.Handled = True
                    Case Keys.Delete
                        ' Delete は全てのチェックを解除, [Shift + Delete] は全てチェックします。

                        For i As Integer = 0 To Items.Count - 1
                            Me.SetItemChecked(i, e.Shift)
                        Next

                        e.Handled = True
                End Select

                MyBase.OnKeyDown(e)
            End Sub

            ''' <summary>
            ''' マウス移動時のイベントハンドラ
            ''' </summary>
            Protected Overrides Sub OnMouseMove(ByVal e As MouseEventArgs)
                MyBase.OnMouseMove(e)
                Dim index As Integer = Me.IndexFromPoint(e.Location)

                If (index >= 0) AndAlso (index <> m_curSelIndex) Then
                    m_curSelIndex = index
                    Me.SetSelected(index, True)
                End If
            End Sub
        End Class
#End Region

#Region "フィールド"
        Private m_cclb As CustomCheckedListBox
        Private m_checkedStateArr() As Boolean
        Private m_dropdownClosed As Boolean = True
        Private m_oldStrValue As String = String.Empty
        Private m_parent As CheckedComboBox
#End Region

#Region "コンストラクタ"

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="parent">
        ''' <seealso cref="CheckedComboBox">CheckedComboBox</seealso>オブジェクト
        ''' </param>
        Public Sub New(ByVal parent As CheckedComboBox)
            Me.m_parent = parent
            Me.InitializeComponent()

            Me.ShowInTaskbar = False
            ' イベントハンドラを設定します。

            AddHandler Me.m_cclb.ItemCheck, AddressOf Me.M_cclb_ItemCheck

            m_cclb.Items.Clear()
            For Each pRow As DataRow In CType(parent.ComboSource, DataTable).Rows

                Dim n As Integer = m_cclb.Items.Add(pRow.Item(parent.DisplayMember))
                If pRow.Item(parent.SelectValueMember) Then
                    'm_cclb.CheckedItems..Items(n)
                End If
            Next
        End Sub
#End Region

#Region "プロパティ"

        ''' <summary>
        ''' ドロップダウンに関連付けられている
        ''' <see cref="CheckedListBox">CheckedListBox</see>オブジェクトを取得します。
        ''' </summary>
        Public ReadOnly Property List() As CheckedListBox
            Get
                Return m_cclb
            End Get
        End Property

        ''' <summary>
        ''' ドロップダウンで現在選択された項目のインデックスを取得・設定します。 
        ''' </summary>
        Public Property SelectedIndex() As Integer
            Get
                Return m_cclb.SelectedIndex
            End Get

            Set(ByVal value As Integer)
                If (m_cclb IsNot Nothing) Then
                    m_cclb.SelectedIndex = value
                    If (m_cclb.Visible = False) Then
                        m_parent.Text = GetCheckedItemsStringValue()
                    End If
                End If
            End Set

        End Property

        ''' <summary>
        ''' 選択項目が変更されているか取得します。
        ''' </summary>
        Public ReadOnly Property ValueChanged() As Boolean
            Get
                Dim newStrValue As String = m_parent.Text
                If ((m_oldStrValue.Length > 0) AndAlso (newStrValue.Length > 0)) Then
                    Return (m_oldStrValue.CompareTo(newStrValue) <> 0)
                Else
                    Return (m_oldStrValue.Length <> newStrValue.Length)
                End If
            End Get
        End Property

#End Region

#Region "イベント"

        ''' <summary>
        ''' <see cref="CustomCheckedListBox.ItemCheck">
        ''' CustomCheckedListBox.ItemCheck</see>イベントを発生させます。 
        ''' </summary>
        Private Sub M_cclb_ItemCheck(ByVal sender As Object, ByVal e As ItemCheckEventArgs)
            If (m_parent.ItemChecked IsNot Nothing) Then
                m_parent.ItemChecked(sender, e)
            End If

        End Sub



        ''' <summary>
        ''' <see cref="Activated">Activated</see>イベントを発生させます。 
        ''' </summary>
        Protected Overrides Sub OnActivated(ByVal e As EventArgs)
            MyBase.OnActivated(e)
            m_dropdownClosed = False
            m_oldStrValue = m_parent.Text
            m_checkedStateArr = New Boolean(m_cclb.Items.Count - 1) {}
            For i As Integer = 0 To m_cclb.Items.Count - 1
                m_checkedStateArr(i) = m_cclb.GetItemChecked(i)
            Next
        End Sub



        ''' <summary>
        ''' <see cref="Deactivate">Deactivate</see>イベントを発生させます。 
        ''' </summary>
        Protected Overrides Sub OnDeactivate(ByVal e As EventArgs)
            MyBase.OnDeactivate(e)
            If (e IsNot Nothing) AndAlso (TypeOf e Is CheckedComboBoxEventArgs) Then
                CloseDropdown(DirectCast(e, CheckedComboBoxEventArgs).AssignValues)
            Else
                CloseDropdown(True)
            End If
        End Sub
#End Region

#Region "メソッド"

        ''' <summary>
        ''' コンポーネントを初期化します。
        ''' </summary>
        Private Sub InitializeComponent()
            Me.m_cclb = New CustomCheckedListBox()
            Me.SuspendLayout()

            ' 
            ' m_cclb
            ' 
            Me.m_cclb.BorderStyle = System.Windows.Forms.BorderStyle.None
            Me.m_cclb.Dock = System.Windows.Forms.DockStyle.Fill
            Me.m_cclb.FormattingEnabled = True
            Me.m_cclb.Location = New System.Drawing.Point(0, 0)
            Me.m_cclb.Name = "m_cclb"
            Me.m_cclb.Size = New System.Drawing.Size(47, 15)
            Me.m_cclb.Font = DirectCast(m_parent.Font.Clone(), Font)

            ' 
            ' Dropdown
            ' 
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0F, 13.0F)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.BackColor = System.Drawing.SystemColors.Menu
            Me.ClientSize = New System.Drawing.Size(47, 16)
            Me.ControlBox = False
            Me.Controls.Add(Me.m_cclb)
            Me.ForeColor = System.Drawing.SystemColors.ControlText
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
            Me.MinimizeBox = False

            Me.Name = "m_parent"
            Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
            Me.ResumeLayout(False)
        End Sub



        ''' <summary>
        ''' ドロップダウンを閉じ、指定されたブールパラメタに応じて変更します。
        ''' <para>
        ''' 呼び出し元が変更を求める可能性があるにもかかわらず、確定されていなければ、
        ''' これは必ずしもすべての変更が発生したわけではありません。<br/>
        ''' 発信者は、任意の実際の値の変更を決定するために 
        ''' <see cref="CheckedComboBox">CheckedComboBox</see>（ドロップダウンが閉じている）の 
        ''' <see cref="ValueChanged">ValueChanged</see>プロパティを確認してください。
        ''' </para>
        ''' </summary>
        ''' <param name="enactChanges">変更確定フラグ</param>
        Public Sub CloseDropdown(ByVal enactChanges As Boolean)
            If (m_dropdownClosed) Then
                Return
            End If

            If (enactChanges) Then
                m_parent.SelectedIndex = -1
                m_parent.Text = GetCheckedItemsStringValue()
            Else
                For i As Integer = 0 To m_cclb.Items.Count - 1
                    m_cclb.SetItemChecked(i, m_checkedStateArr(i))
                Next
            End If

            m_dropdownClosed = True
            m_parent.Focus()
            m_parent.SelectionLength = 0
            Me.Hide()

            m_parent.OnDropDownClosed(New CheckedComboBoxEventArgs(EventArgs.Empty, False))
        End Sub

        ''' <summary>
        ''' チェック項目の文字列を結合して取得します。
        ''' </summary>
        Public Function GetCheckedItemsStringValue() As String
            Dim sb As New System.Text.StringBuilder()

            For i As Integer = 0 To m_cclb.CheckedItems.Count - 1
                sb.Append(m_cclb.GetItemText(m_cclb.CheckedItems(i))).Append(m_parent.Separator)
            Next

            If (sb.Length > 0) Then
                sb.Remove(sb.Length - m_parent.Separator.Length, m_parent.Separator.Length)
            End If

            Return sb.ToString()
        End Function

#End Region
    End Class
#End Region


#Region "イベントの宣言"
    Public ItemChecked As ItemCheckEventHandler
#End Region
#Region "フィールド"
    Private m_dropdown As CheckedComboBoxDropdown
    Private m_separator As String
    Private mvarDataSource As Object
#End Region

#Region "コンストラクタ"
    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    Public Sub New()
        MyBase.New()
        Me.DrawMode = DrawMode.OwnerDrawVariable
        Me.Separator = ", "
        Me.DropDownHeight = 1
        Me.DropDownStyle = ComboBoxStyle.DropDown
        Me.CheckOnClick = True

    End Sub

#End Region

#Region "プロパティ"
    ''' <summary>
    ''' このコントロール内でチェックされているインデックスのコレクションを取得します。
    ''' </summary>
    Public ReadOnly Property CheckedIndices() As CheckedListBox.CheckedIndexCollection
        Get
            Return m_dropdown.List.CheckedIndices
        End Get
    End Property


    ''' <summary>
    ''' このコントロール内でチェックされている項目のコレクションを取得します。
    ''' </summary>
    Public ReadOnly Property CheckedItems() As CheckedListBox.CheckedItemCollection
        Get
            Return m_dropdown.List.CheckedItems
        End Get
    End Property

    ''' <summary>
    ''' 項目が選択されたときに、チェックボックスを切り替えるかどうかを示す値を取得または設定します。 
    ''' </summary>
    Public Property CheckOnClick() As Boolean

        Get
            If (m_dropdown IsNot Nothing) Then
                Return m_dropdown.List.CheckOnClick
            Else
                Return Nothing
            End If
        End Get

        Set(ByVal value As Boolean)
            If (m_dropdown IsNot Nothing) Then m_dropdown.List.CheckOnClick = value
        End Set

    End Property

    ''' <summary>
    ''' このコントロールのデータソースを取得または設定します。 
    ''' </summary>
    ''' <remarks>このプロパティは使用しないでください。実装すると 
    ''' <see cref="NotImplementedException">NotImplementedException</see>
    '''  例外が発生します。</remarks>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden), Browsable(False)>
    Private Shadows Property DataSource() As Object
        Get
            Throw New NotImplementedException("このプロパティは使用できません。実装しないでください。")
        End Get
        Set(ByVal value As Object)
            Throw New NotImplementedException("このプロパティは使用できません。実装しないでください。")
        End Set
    End Property

    Public Shadows Property ComboSource() As Object
        Get
            Return mvarDataSource
        End Get
        Set(ByVal value As Object)
            mvarDataSource = value
        End Set
    End Property

    Private mvarSelectValueMember As String = ""
    ''' <summary>
    ''' このコントロールの選択に関連づけられたプロパティを取得または設定します。 
    ''' </summary>
    Public Shadows Property SelectValueMember() As String
        Get
            Return mvarSelectValueMember
        End Get
        Set(ByVal value As String)
            mvarSelectValueMember = value
        End Set
    End Property



    ''' <summary>
    ''' このコントロールに表示するプロパティを取得または設定します。 
    ''' </summary>
    Public Shadows Property DisplayMember() As String
        Get
            Return MyBase.DisplayMember
        End Get
        Set(ByVal value As String)
            MyBase.DisplayMember = value
        End Set
    End Property

    ''' <summary>
    ''' このコントロール内の項目のコレクションを取得します。
    ''' </summary>
    Public Shadows ReadOnly Property Items() As CheckedListBox.ObjectCollection
        Get
            Return m_dropdown.List.Items
        End Get
    End Property


    ''' <summary>
    ''' 現在選択されている項目を指定しているインデックスを取得または設定します。 
    ''' </summary>
    Public Overrides Property SelectedIndex() As Integer
        Get
            If (m_dropdown IsNot Nothing) Then
                Return m_dropdown.SelectedIndex
            Else
                Return -1
            End If
        End Get

        Set(ByVal value As Integer)
            If (m_dropdown Is Nothing) Then Return
            m_dropdown.SelectedIndex = value
        End Set
    End Property

    ''' <summary>
    ''' <see cref="Text">Text</see> に表示される項目間のセパレータ文字を取得または設定します。
    ''' </summary>
    Public Property Separator() As String
        Get
            Return m_separator
        End Get
        Set(ByVal value As String)
            m_separator = value
        End Set
    End Property

    ''' <summary>
    ''' このコントロールに関連付けられているテキストを取得または設定します。
    ''' </summary>
    Public Overrides Property Text() As String
        Get
            If (MyBase.Items.Count = 0) Then
                Return String.Empty
            End If
            Return MyBase.Text
        End Get
        Set(ByVal value As String)
            Try
                MyBase.Text = value
            Catch ex As ArgumentOutOfRangeException
                ' この例外は意図的にスルーする。
            Catch ex As Exception
                Throw

            End Try
        End Set
    End Property

    ''' <summary>
    ''' コントロールの値が変更されたか取得します。
    ''' </summary>
    Public ReadOnly Property ValueChanged() As Boolean
        Get
            Return m_dropdown.ValueChanged
        End Get
    End Property

    ''' <summary>
    ''' コントロール内の項目の実際の値として使用するプロパティを取得または設定します。
    ''' </summary>
    Public Shadows Property ValueMember() As String
        Get
            If (m_dropdown IsNot Nothing) Then
                Return MyBase.ValueMember
            Else
                Return MyBase.ValueMember
            End If
        End Get
        Set(ByVal value As String)
            MyBase.ValueMember = value
        End Set
    End Property
#End Region
#Region "イベント"
    ''' <summary>
    ''' レイアウト変更時のイベントを発生させます。
    ''' </summary>
    Protected Overrides Sub OnLayout(ByVal levent As LayoutEventArgs)
        MyBase.OnLayout(levent)
        If (m_dropdown Is Nothing) Then
            m_dropdown = New CheckedComboBoxDropdown(Me)
        End If
    End Sub

    ''' <summary>
    ''' フォント変更時のイベントを発生させます。
    ''' </summary>
    Protected Overrides Sub OnFontChanged(ByVal e As EventArgs)
        MyBase.OnFontChanged(e)
        Dim font As Font = DirectCast(Me.Font.Clone(), Font)
        m_dropdown.Font = font
        m_dropdown.List.Font = font
    End Sub

    ''' <summary>
    ''' ドロップダウン時のイベントを発生させます。
    ''' </summary>
    Protected Overrides Sub OnDropDown(ByVal e As EventArgs)
        MyBase.OnDropDown(e)
        Me.DoDropDown()
    End Sub



    ''' <summary>
    ''' ドロップダウンを閉じた時のイベントを発生させます。
    ''' </summary>
    Protected Overrides Sub OnDropDownClosed(ByVal e As EventArgs)

        If (TypeOf e Is CheckedComboBoxEventArgs) Then

            MyBase.OnDropDownClosed(e)

        End If

    End Sub



    ''' <summary>
    ''' キーダウン時のイベントを発生させます。
    ''' </summary>
    Protected Overrides Sub OnKeyDown(ByVal e As KeyEventArgs)

        If (e.KeyCode = Keys.Down) Then
            OnDropDown(EventArgs.Empty)
        End If

        ' 特定のキーまたは組合せが妨げられないことを確認します。
        e.Handled = Not e.Alt AndAlso Not (e.KeyCode = Keys.Tab) AndAlso Not ((e.KeyCode = Keys.Left) OrElse (e.KeyCode = Keys.Right) OrElse (e.KeyCode = Keys.Home) OrElse (e.KeyCode = Keys.End))

        MyBase.OnKeyDown(e)
    End Sub



    ''' <summary>
    ''' キーが押された時のイベントを発生させます。
    ''' </summary>
    Protected Overrides Sub OnKeyPress(ByVal e As KeyPressEventArgs)
        e.Handled = True
        MyBase.OnKeyPress(e)

    End Sub

#End Region

#Region "メソッド"

    ''' <summary>
    ''' ドロップダウン処理を行います。
    ''' </summary>
    Private Sub DoDropDown()
        If (Not m_dropdown.Visible) Then
            Dim rect As Rectangle = RectangleToScreen(Me.ClientRectangle)

            m_dropdown.Location = New Point(rect.X, rect.Y + Me.Size.Height)
            Dim count As Integer = m_dropdown.List.Items.Count
            If (count > Me.MaxDropDownItems) Then
                count = Me.MaxDropDownItems
            ElseIf (count = 0) Then
                count = 1
            End If

            m_dropdown.Size = New Size(Me.Size.Width, (m_dropdown.List.ItemHeight) * count + 2)
            m_dropdown.Show(Me)

        End If

    End Sub

    ''' <summary>
    ''' 項目のチェック状況を取得します。
    ''' </summary>
    ''' <param name="index">インデックス</param>
    ''' <returns>チェックされていたら True</returns>
    ''' <remarks>index がリストの範囲外の場合、
    ''' <see cref="ArgumentOutOfRangeException">ArgumentOutOfRangeException</see>
    '''  例外が発生します。</remarks>
    Public Function GetItemChecked(ByVal index As Integer) As Boolean
        If (index < 0 OrElse index > Items.Count) Then
            Throw New ArgumentOutOfRangeException("index", "範囲外の値が渡されました。")
        Else
            Return m_dropdown.List.GetItemChecked(index)
        End If
    End Function

    ''' <summary>
    ''' 指定したインデックスの項目のチェック状況を調べます。
    ''' </summary>
    ''' <param name="index">インデックス</param>
    ''' <returns></returns>
    ''' <remarks>index がリストの範囲外の場合、
    ''' <see cref="ArgumentOutOfRangeException">ArgumentOutOfRangeException</see>
    '''  例外が発生します。</remarks>
    Public Function GetItemCheckState(ByVal index As Integer) As CheckState

        If (index < 0 OrElse index > Items.Count) Then

            Throw New ArgumentOutOfRangeException("index", "範囲外の値が渡されました。")

        Else

            Return m_dropdown.List.GetItemCheckState(index)

        End If

    End Function

    ''' <summary>
    ''' 指定したインデックスの項目をチェックします。
    ''' </summary>
    ''' <param name="index">インデックス</param>
    ''' <param name="isChecked"></param>
    ''' <remarks>index がリストの範囲外の場合、
    ''' <see cref="ArgumentOutOfRangeException">ArgumentOutOfRangeException</see>
    '''  例外が発生します。</remarks>
    Public Sub SetItemChecked(ByVal index As Integer, ByVal isChecked As Boolean)

        If (index < 0 OrElse index > Items.Count) Then

            Throw New ArgumentOutOfRangeException("index", "範囲外の値が渡されました。")

        Else

            m_dropdown.List.SetItemChecked(index, isChecked)

            ' Text の更新に必要です

            Me.Text = m_dropdown.GetCheckedItemsStringValue()

            Me.SelectionLength = 0

        End If

    End Sub

    ''' <summary>
    ''' 指定したインデックスの項目のチェック状況を設定します。
    ''' </summary>
    ''' <param name="index">インデックス</param>
    ''' <param name="state"></param>
    ''' <remarks>index がリストの範囲外の場合、
    ''' <see cref="ArgumentOutOfRangeException">ArgumentOutOfRangeException</see>
    '''  例外が発生します。</remarks>
    Public Sub SetItemCheckState(ByVal index As Integer, ByVal state As CheckState)

        If (index < 0 OrElse index > Items.Count) Then

            Throw New ArgumentOutOfRangeException("index", "範囲外の値が渡されました。")

        Else

            m_dropdown.List.SetItemCheckState(index, state)

            ' Text の更新に必要です

            Me.Text = m_dropdown.GetCheckedItemsStringValue()

            Me.SelectionLength = 0

        End If

    End Sub

#End Region

End Class

''' <summary>
''' <see cref="CheckedComboBox">CheckedComboBox</see> 用イベントパラメータクラス
''' </summary>
Friend Class CheckedComboBoxEventArgs

    Inherits EventArgs



#Region "フィールド"

    Private m_assignValues As Boolean = False

    Private m_event As EventArgs

#End Region



#Region "コンストラクタ"

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    Public Sub New(ByVal e As EventArgs, ByVal assignValues As Boolean)

        MyBase.New()

        m_event = e

        m_assignValues = assignValues

    End Sub

#End Region



#Region "プロパティ"

    ''' <summary>
    ''' 値を割り当てているか取得・設定します。
    ''' </summary>
    ''' <value>値を割り当てていれば True</value>
    Public Property AssignValues() As Boolean

        Get

            Return m_assignValues

        End Get

        Set(ByVal value As Boolean)

            m_assignValues = value

        End Set

    End Property

#End Region



End Class

