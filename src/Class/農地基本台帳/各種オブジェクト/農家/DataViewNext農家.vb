
Imports HimTools2012.controls
Imports HimTools2012.controls.DVCtrlCommonBase

Public Class DataViewNext農家
    Inherits CDataViewPanel農地台帳


    Public Sub New(pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget, App農地基本台帳.TBL世帯, SysAD.page農家世帯.DataViewCollection, True, True)
        SetButtons(New ToolStripSeparator, CreateButton("経営農地", "経営農地の一覧"), CreateButton("基本台帳印刷", "基本台帳印刷"))
        '"AddTButton","耕作証明ボタン","VB.Commandbutton;Caption=面積証明;Height=300;NewLine;","耕作面積証明印刷",""
        '"AddTButton","多筆証明ボタン","VB.Commandbutton;Caption=多筆証明;Height=300;NewLine;","耕作証明多筆型",""
        Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsReadOnly, pTarget, "ID"), "ID")
        Panel.AddCtrl(New DateTimePickerPlus(True, pTarget, "更新日"), "更新日", em改行.改行あり)

        Dim b新項 As Boolean = b新項目保存(pTarget, "支店等住所")

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("世帯基本", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New ComboBoxPlus(New DataView(App農地基本台帳.TBL個人.Body, "[世帯ID]=" & pTarget.ID, "", DataViewRowState.CurrentRows), "氏名", "ID",
                        .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "世帯主ID"), "世帯主")
            ))

            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("呼び出す", "世帯主を呼ぶ"), "", True)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(HorizontalAlignment.Left, False, pTarget, "市町村ID", , 200), "市町村ID", em改行.改行あり)
            If Not b新項 Then SetTextComboWithMaster(.Panel, "大字", "大字ID", "大　字", em改行.改行なし) '【保留】ReadOnly
            If Not b新項 Then SetTextComboWithMaster(.Panel, "小字", "小字ID", "小　字", em改行.改行あり) '【保留】ReadOnly
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(HorizontalAlignment.Left, True, pTarget, "住所", , 600), "住所", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(HorizontalAlignment.Left, b新項目保存(pTarget, "支店等住所"), pTarget, "支店等住所", , 600), "支店等住所", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("行政区"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsReadOnly, pTarget, "世帯主行政区ID"), "行政区"), True
            ), "", True)

            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsReadOnly, pTarget, "世帯主郵便番号"), "郵便番号")
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsReadOnly, pTarget, "世帯主電話番号"), "電話番号", True)
            .Panel.AddCtrl(New OptionButtonPlus(Split("設定なし,農業専業,農業を主,専業が主,農業非従事,不明", ","), 20, pTarget, "就業状況"), "就業状況", em改行.改行あり) '【保留】ReadOnly
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("所有区分"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "農地所有区分"), pTarget, "農地所有区分", , 60), "農地所有区分"), b新項目保存(pTarget, "農地所有区分")
            ), "", True, 200)
            .Panel.AddCtrl(New ComboBoxPlus(New DataView(App農地基本台帳.TBL個人.Body, "[ID] = " & mvarTarget.ID, "[ID]", DataViewRowState.CurrentRows), "氏名", "ID",
             .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "農事組合ID"), mvarTarget, "農事組合ID", , 60), "農事組合")
           ), , em改行.改行あり).Width = 200
            .Panel.AddCtrl(New ComboBoxPlus(New DataView(App農地基本台帳.TBL個人.Body, "[ID] = " & mvarTarget.ID, "[ID]", DataViewRowState.CurrentRows), "氏名", "ID",
             .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "所属農協ID"), mvarTarget, "所属農協ID", , 60), "所属農協")
           ), , em改行.改行あり).Width = 200

            'SetInputIDandText(.Panel, "所有者氏名", "農事組合ID", "所有者氏名")

            '認定情報を引っ張る？
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Real, b新項目保存(pTarget, "認定時面積"), pTarget, "認定時面積"), "認定時面積")
            .Panel.FitLabelWidth()
        End With
        Dim nHeight As Integer
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("農家情報", Me), "", True), HimTools2012.controls.GroupBoxPlus)
            nHeight = .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "農家番号"), "農家番号", True).Height

            .Panel.AddCtrl(New OptionButtonPlus(Split("-,中心経営体,中心経営体ではない,調査中", ","), nHeight, pTarget, "人農地プラン中心経営体区分"), "人農地プラン中心経営体区分", em改行.改行あり) '【保留】ReadOnly
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("家族経営協定", "Value"), "希望する", "希望しない"), "家族経営協定")
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("規模拡大希望", "Value"), "希望する", "希望しない"), "規模拡大希望", em改行.改行あり)
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("あっせん希望", "Value"), "希望する", "希望しない"), "あっせん希望")
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("あっせん希望種別"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "あっせん希望種別"), "あっせん種別"), False
            ), "", True)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,申出,調査中", ","), nHeight, pTarget, "農地移動適正化あっせん事業"), "農地移動適正化あっせん事業", em改行.改行あり) '【保留】ReadOnly
            .Panel.AddCtrl(New DateTimePickerPlus(b新項目保存(pTarget, "あっせん登録日"), pTarget, "あっせん登録日"), "あっせん登録日", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Real, b新項目保存(pTarget, "あっせん登録時面積"), pTarget, "あっせん登録時面積"), "あっせん登録時面積", em改行.改行あり)

            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("法人化希望", "Value"), "希望する", "希望しない"), "法人化希望")
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("農業法人研修希望", "Value"), "希望する", "希望しない"), "農業法人研修希望", True)
            .Panel.AddCtrl(New OptionButtonPlus(Split("未登録,青色申告,白色申告,その他", ","), nHeight, pTarget, "青色申告"), "申告方法")
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "青色申告開始日"), "青色申告開始日", True)
            .Panel.FitLabelWidth()
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("組織等への参加状態", Me), "", True), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New OptionButtonPlus(Split("参加,不参加,調整中", ","), nHeight, pTarget, "農用地改善団体参加"), "農用地改善団体参加", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("非構成員,構成員,調整中", ","), nHeight, pTarget, "地域農業集団構成員"), "地域農業集団", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "その他の団体への参加", , 400), "その他の団体", em改行.改行あり)

            .Panel.FitLabelWidth()
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("農家世帯の状態", Me), "", True), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("合併異動", "Value"), "あり", "なし"), "合併異動")
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "合併異動日"), "合併異動日", em改行.改行あり)
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("先行異動", "Value"), "あり", "なし"), "先行異動")
            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, pTarget, "先行異動日"), "更新日", em改行.改行あり)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("備考", "Text"), 400), "備考", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, pTarget, "確認日時"), "確認日時", em改行.改行あり)

            .Panel.FitLabelWidth()
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("主要農機具および農業用施設", Me), "", True), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "トラクター台数"), pTarget, "トラクター台数", , 100), "トラクター台数(台)", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "耕運機台数"), pTarget, "耕運機台数", , 100), "耕運機台数(台)", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "田植機台数"), pTarget, "田植機台数", , 100), "田植機台数(台)", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "コンバイン台数"), pTarget, "コンバイン台数", , 100), "コンバイン台数(台)", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "乾燥機台数"), pTarget, "乾燥機台数", , 100), "乾燥機台数(台)", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "噴霧器台数"), pTarget, "噴霧器台数", , 100), "噴霧器台数(台)", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "その他機具台数"), pTarget, "その他機具台数", , 100), "その他機具台数(台)", em改行.改行なし)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, b新項目保存(pTarget, "その他機具内訳"), pTarget, "その他機具内訳", , 200), "その他機具内訳", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "畜舎規模"), pTarget, "畜舎規模", , 100), "畜舎規模(㎡)", em改行.改行なし)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "畜舎数"), pTarget, "畜舎数", , 100), "畜舎数(棟)", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "温室規模"), pTarget, "温室規模", , 100), "温室規模(㎡)", em改行.改行なし)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "温室数"), pTarget, "温室数", , 100), "温室数(棟)", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "その他施設規模"), pTarget, "その他施設規模", , 100), "その他施設規模(㎡)", em改行.改行なし)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "その他施設数"), pTarget, "その他施設数", , 100), "その他施設数(棟)", em改行.改行なし)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, b新項目保存(pTarget, "その他施設内訳"), pTarget, "その他施設内訳", , 200), "その他施設内訳", em改行.改行あり)
            .Panel.FitLabelWidth()
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("主な販売収入", Me), "", True), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,米,畜産,果樹,そさい,養蚕,その他", ","), nHeight, pTarget, "販売収入1位"), "販売収入1位", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,米,畜産,果樹,そさい,養蚕,その他", ","), nHeight, pTarget, "販売収入2位"), "販売収入2位", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,米,畜産,果樹,そさい,養蚕,その他", ","), nHeight, pTarget, "販売収入3位"), "販売収入3位", em改行.改行あり)
            .Panel.FitLabelWidth()
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("主要作目", Me), "", True), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("耕作状況"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "主要作物1"), pTarget, "主要作物1", , 60), "主要作物1"), b新項目保存(pTarget, "主要作物1")
            ), "", em改行.改行なし)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "主要作物規模1"), pTarget, "主要作物規模1", , 100), "主要作物規模1(㎡)", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("耕作状況"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "主要作物2"), pTarget, "主要作物2", , 60), "主要作物2"), b新項目保存(pTarget, "主要作物2")
            ), "", em改行.改行なし)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "主要作物規模2"), pTarget, "主要作物規模2", , 100), "主要作物規模2(㎡)", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("耕作状況"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "主要作物3"), pTarget, "主要作物3", , 60), "主要作物3"), b新項目保存(pTarget, "主要作物3")
            ), "", em改行.改行なし)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "主要作物規模3"), pTarget, "主要作物規模3", , 100), "主要作物規模3(㎡)", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("耕作状況"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "主要作物4"), pTarget, "主要作物4", , 60), "主要作物4"), b新項目保存(pTarget, "主要作物4")
            ), "", em改行.改行なし)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "主要作物規模4"), pTarget, "主要作物規模4", , 100), "主要作物規模4(㎡)", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("耕作状況"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "主要作物5"), pTarget, "主要作物5", , 60), "主要作物5"), b新項目保存(pTarget, "主要作物5")
            ), "", em改行.改行なし)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "主要作物規模5"), pTarget, "主要作物規模5", , 100), "主要作物規模5(㎡)", em改行.改行あり)
            .Panel.FitLabelWidth()
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("主要家畜", Me), "", True), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "肉用牛頭数"), pTarget, "肉用牛頭数", , 100), "肉用牛頭数", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "乳牛頭数"), pTarget, "乳牛頭数", , 100), "乳牛頭数", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "豚頭数"), pTarget, "豚頭数", , 100), "豚頭数", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "採卵用鶏羽数"), pTarget, "採卵用鶏羽数", , 100), "採卵用鶏羽数", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "ブロイラー羽数"), pTarget, "ブロイラー羽数", , 100), "ブロイラー羽数", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "その他家畜頭数"), pTarget, "その他家畜頭数", , 100), "その他家畜頭数", em改行.改行なし)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, b新項目保存(pTarget, "その他家畜内訳"), pTarget, "その他家畜内訳", , 200), "その他家畜内訳", em改行.改行あり)
            .Panel.FitLabelWidth()
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("制度資金等利用状況", Me), "", True), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, b新項目保存(pTarget, "制度資金種別1"), pTarget, "制度資金種別1", , 200), "制度資金種類1", em改行.改行なし)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "制度資金西暦1"), pTarget, "制度資金西暦1", , 100), "制度資金年次(西暦)1", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, b新項目保存(pTarget, "制度資金種別2"), pTarget, "制度資金種別2", , 200), "制度資金種類2", em改行.改行なし)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "制度資金西暦2"), pTarget, "制度資金西暦2", , 100), "制度資金年次(西暦)2", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, b新項目保存(pTarget, "制度資金種別3"), pTarget, "制度資金種別3", , 200), "制度資金種類3", em改行.改行なし)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "制度資金西暦3"), pTarget, "制度資金西暦3", , 100), "制度資金年次(西暦)3", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, b新項目保存(pTarget, "制度資金種別4"), pTarget, "制度資金種別4", , 200), "制度資金種類4", em改行.改行なし)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "制度資金西暦4"), pTarget, "制度資金西暦4", , 100), "制度資金年次(西暦)4", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, b新項目保存(pTarget, "制度資金種別5"), pTarget, "制度資金種別5", , 200), "制度資金種類5", em改行.改行なし)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "制度資金西暦5"), pTarget, "制度資金西暦5", , 100), "制度資金年次(西暦)5", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, b新項目保存(pTarget, "制度資金種別6"), pTarget, "制度資金種別6", , 200), "制度資金種類6", em改行.改行なし)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "制度資金西暦6"), pTarget, "制度資金西暦6", , 100), "制度資金年次(西暦)6", em改行.改行あり)
            .Panel.FitLabelWidth()
        End With

        If Not b新項 Then
            With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("農家分類", Me), "", True), HimTools2012.controls.GroupBoxPlus)
                .Panel.AddCtrl(New OptionButtonPlus(Split("-,専業,兼業(恒常的勤務),兼業(出稼),兼業(臨時),兼業(自営)", ","), nHeight, pTarget, "農家分類専業形態"), "専業形態", em改行.改行あり)
                .Panel.AddCtrl(New OptionButtonPlus(Split("-,「農業あとつぎ」がいる,「農業あとつぎ予定者」がいる,その他", ","), nHeight, pTarget, "農家分類あとつぎ"), "あとつぎ", em改行.改行あり)
                .Panel.AddCtrl(New OptionButtonPlus(Split("-,経営規模を拡大する,現状維持する,経営規模を縮小する,農業をやめたい", ","), nHeight, pTarget, "農家分類規模拡大志向"), "規模拡大志向", em改行.改行あり)
                .Panel.FitLabelWidth()
            End With
        End If

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("その他営農の状況", Me), "", True), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "団地数"), pTarget, "団地数", , 200), "団地数(筆)", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "基盤整備実施済筆数"), pTarget, "基盤整備実施済筆数", , 200), "基盤整備実施済筆数", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Real, b新項目保存(pTarget, "基盤整備実施済面積"), pTarget, "基盤整備実施済面積", , 200), "基盤整備実施済面積", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "転作筆数"), pTarget, "転作筆数", , 200), "転作筆数", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Real, b新項目保存(pTarget, "転作面積"), pTarget, "転作面積", , 200), "転作面積", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "裏作利用筆数"), pTarget, "裏作利用筆数", , 200), "裏作利用筆数", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Real, b新項目保存(pTarget, "裏作利用面積"), pTarget, "裏作利用面積", , 200), "裏作利用面積", em改行.改行あり)
            .Panel.FitLabelWidth()
        End With

        If Not b新項 Then
            With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("経営意向等", Me), "", True), HimTools2012.controls.GroupBoxPlus)
                .Panel.AddCtrl(New DateTimePickerPlus(b新項目保存(pTarget, "経営意向等調査日"), pTarget, "経営意向等調査日"), "経営意向等調査日", em改行.改行あり)
                .Panel.AddCtrl(New OptionButtonPlus(Split("-,農業だけでやる,農業中心でやる,兼業中心でやる,農業をやめたい,その他,調査中", ","), nHeight, pTarget, "経営意向等農業志向"), "農業志向", em改行.改行あり)
                .Panel.AddCtrl(New OptionButtonPlus(Split("-,経営規模拡大,現状維持,経営規模縮小,その他,調査中", ","), nHeight, pTarget, "経営意向等経営計画"), "経営計画", em改行.改行あり)

                With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("経営計画部門1", Me), "", True), HimTools2012.controls.GroupBoxPlus)
                    .Panel.AddCtrl(New OptionButtonPlus(Split("-,米,畜産,果樹,そさい,養蚕,その他", ","), nHeight, pTarget, "経営部門1"), "経営計画部門", em改行.改行あり)
                    .Panel.AddCtrl(New OptionButtonPlus(Split("-,拡大,縮小", ","), nHeight, pTarget, "経営部門1拡大縮小"), "拡大縮小", em改行.改行あり)
                    .Panel.AddCtrl(New OptionButtonPlus(Split("-,売買,貸借,作業受委託,その他", ","), nHeight, pTarget, "経営部門1拡大縮小方法"), "拡大縮小方法", em改行.改行あり)
                    .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Real, b新項目保存(pTarget, "経営部門1拡大縮小面積"), pTarget, "経営部門1拡大縮小面積", , 200), "拡大縮小面積", em改行.改行あり)
                    .Panel.FitLabelWidth()
                End With

                With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("経営計画部門2", Me), "", True), HimTools2012.controls.GroupBoxPlus)
                    .Panel.AddCtrl(New OptionButtonPlus(Split("-,米,畜産,果樹,そさい,養蚕,その他", ","), nHeight, pTarget, "経営部門2"), "経営計画部門", em改行.改行あり)
                    .Panel.AddCtrl(New OptionButtonPlus(Split("-,拡大,縮小", ","), nHeight, pTarget, "経営部門2拡大縮小"), "拡大縮小", em改行.改行あり)
                    .Panel.AddCtrl(New OptionButtonPlus(Split("-,売買,貸借,作業受委託,その他", ","), nHeight, pTarget, "経営部門2拡大縮小方法"), "拡大縮小方法", em改行.改行あり)
                    .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Real, b新項目保存(pTarget, "経営部門2拡大縮小面積"), pTarget, "経営部門2拡大縮小面積", , 200), "拡大縮小面積", em改行.改行あり)
                    .Panel.FitLabelWidth()
                End With

                With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("経営計画部門3", Me), "", True), HimTools2012.controls.GroupBoxPlus)
                    .Panel.AddCtrl(New OptionButtonPlus(Split("-,米,畜産,果樹,そさい,養蚕,その他", ","), nHeight, pTarget, "経営部門3"), "経営計画部門", em改行.改行あり)
                    .Panel.AddCtrl(New OptionButtonPlus(Split("-,拡大,縮小", ","), nHeight, pTarget, "経営部門3拡大縮小"), "拡大縮小", em改行.改行あり)
                    .Panel.AddCtrl(New OptionButtonPlus(Split("-,売買,貸借,作業受委託,その他", ","), nHeight, pTarget, "経営部門3拡大縮小方法"), "拡大縮小方法", em改行.改行あり)
                    .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Real, b新項目保存(pTarget, "経営部門3拡大縮小面積"), pTarget, "経営部門3拡大縮小面積", , 200), "拡大縮小面積", em改行.改行あり)
                    .Panel.FitLabelWidth()
                End With

                .Panel.FitLabelWidth()
            End With
        End If

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("法人等管理項目", Me), "", True), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("法人格"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, b新項目保存(pTarget, "法人格"), pTarget, "法人格"), "法人格"), b新項目保存(pTarget, "法人格")
            ), "", em改行.改行あり, 200)
            .Panel.AddCtrl(New DateTimePickerPlus(b新項目保存(pTarget, "法人格設立日"), pTarget, "法人格設立日"), "設立日", em改行.改行なし)
            .Panel.AddCtrl(New DateTimePickerPlus(b新項目保存(pTarget, "法人格初回許可日"), pTarget, "法人格初回許可日"), "最初の許可年月日", em改行.改行あり)
            .Panel.FitLabelWidth()
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("選挙関連", Me), "", True), HimTools2012.controls.GroupBoxPlus)

            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "選挙連番"), "選挙連番")
            .Panel.FitLabelWidth()
        End With

        'With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("家族一覧", Me), "", True), HimTools2012.controls.GroupBoxPlus)
        '    App農地基本台帳.TBL個人.Find("[世帯ID]=" & pTarget.ID)
        '    .Panel.AddCtrl(New GridViewNext(New DataView(App農地基本台帳.TBL個人, "[世帯ID]=" & pTarget.ID, "", DataViewRowState.CurrentRows),
        '                       DesignName.DN農家個人, 600), "家族一覧")
        '    .Panel.FitLabelWidth()
        'End With

    End Sub

End Class
