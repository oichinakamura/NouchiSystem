
Imports HimTools2012.controls
Imports HimTools2012.controls.DVCtrlCommonBase

Public Class CInterface利用権設定
    Inherits DataViewNext申請Type1

    Public Sub New(ByVal pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget)

    End Sub

    Public Overrides Sub SetInterface(ByVal pPanel As HimTools2012.controls.FlowLayoutPanelPlus)
        Dim nHeight As Integer = 申請基本1("承認")
        set譲受人(pPanel, emRO.IsReadOnly)

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("法人経由", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "経由法人ID", , 60, , True), "経由法人", em改行.改行なし)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsReadOnly, mvarTarget, "経由法人名", , 240, , True), , em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(False, Me.Target, "機構配分計画中間管理権取得日", 150), "中間管理取得日", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(False, Me.Target, "機構配分計画意見回答日", 150), "意見回答年月日", em改行.改行なし)
            .Panel.AddCtrl(New DateTimePickerPlus(False, Me.Target, "機構配分計画知事公告日", 150), "知事公告年月日", em改行.改行なし)
            .Panel.AddCtrl(New DateTimePickerPlus(False, Me.Target, "機構配分計画認可通知日", 150), "認可通知年月日", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,使用貸借権,賃貸借権", ","), nHeight, mvarTarget, "機構配分計画権利設定内容"), "権利設定内容", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(False, Me.Target, "機構配分計画利用配分計画始期日", 150), "存続期間", em改行.改行なし)
            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("～", "経由法人期間設定", True))
            .Panel.AddCtrl(New DateTimePickerPlus(False, Me.Target, "機構配分計画利用配分計画終期日", 150), "", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, Me.Target, "機構配分計画利用配分計画10a賃借料", , 200), "10a当り借賃額(円)", em改行.改行あり)
            .Panel.FitLabelWidth()
        End With

        set譲渡人(pPanel, "譲渡", True, False, True)

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請地情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            Set関連農地一覧(.Body, False, "農地リスト", "賃借料1円10a当たり:String", "申請部分面積:Decimal")
            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("全選択", "全選択"), "", em改行.改行あり)

            .Panel.AddCtrl(New ComboList(ListResource.S_Data, "譲渡申請理由", mvarTarget, "申請理由A", "Text", ComboBoxStyle.DropDown), "貸人事由", em改行.改行なし, 250)
            .Panel.AddCtrl(New ComboList(ListResource.S_Data, "譲受申請理由", mvarTarget, "申請理由B", "Text", ComboBoxStyle.DropDown), "借人事由", em改行.改行あり, 250)

            ' "AddItem","chk条件付き貸借","台帳管理EXE.CheckCtrl;WithLabel=条件付き貸借;Height=330;","解除条件付きの農地の貸借","",""
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("再設定", "Value"), "あり", "なし"), "再設定", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("小作形態"), "名称", "ID",
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "権利種類", , 60), "形態")
            ), "", em改行.改行あり)
            .Panel.AddCtrl(New ComboList(ListResource.S_Data, "利用権内容", mvarTarget, "利用権内容", "Text", ComboBoxStyle.DropDown), "利用権内容", em改行.改行あり)

            sub小作料設定(.Panel, "小作料", "単位", em改行.改行あり)
            .Panel.AddCtrl(New ComboList(ListResource.S_Data, "支払方法", mvarTarget, "支払方法", "Text", ComboBoxStyle.DropDown), "支払方法", em改行.改行あり)

            sub期間設定(.Panel, "始期", "開始年月日", "終期", "終了年月日", "期間設定", em改行.改行なし)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "期間", , 80), "期間", em改行.改行あり)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("申請地目安", "Text"), 600), "申請地の目安", em改行.改行あり)
        End With

        Set申請人世帯営農状況(pPanel, "貸　人", "借　人")

        'With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("営農情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)

        '    'TODO "IF","申請時担い手農家","AddItem","Chk担い手農家","VB.CheckBox;Caption=担い手農家;Width=1500;Height=330;","担い手農家"
        '    'TODO "IF","申請時認定農業者","AddItem","Chk認定農業者","VB.CheckBox;Caption=認定農業者;Width=1500;Height=330;","認定農業者"
        'End With
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("管理情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, mvarTarget, "公告年月日"), "公告年月日", em改行.改行あり)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("備考", "Text"), 400), "備考", em改行.改行あり)
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("同意書", "Value"), "有", "無"), "同意書の有無")
        End With

        sub権利移動借賃等調査_様式1()
    End Sub
End Class
