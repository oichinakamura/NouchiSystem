Imports HimTools2012.controls
Imports HimTools2012.controls.DVCtrlCommonBase

Public Class CInterface基盤法所有権設定
    Inherits DataViewNext申請Type1

    Public Sub New(ByVal pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget)
    End Sub

    Public Overrides Sub SetInterface(ByVal pPanel As HimTools2012.controls.FlowLayoutPanelPlus)
        申請基本1("承認")
        set譲受人(pPanel, emRO.IsReadOnly)
        set譲渡人(pPanel, "譲渡", True, False, True)

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請地情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            Set関連農地一覧(.Body, False, "農地リスト")
            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("全選択", "全選択"), "", em改行.改行あり)

            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("部分設定", "部分設定"), "", em改行.改行あり)
            .Panel.AddCtrl(New ComboList(ListResource.S_Data, "譲渡申請理由", mvarTarget, "申請理由A", "Text", ComboBoxStyle.DropDown), "譲渡人事由", em改行.改行あり)
            .Panel.AddCtrl(New ComboList(ListResource.S_Data, "譲受申請理由", mvarTarget, "申請理由B", "Text", ComboBoxStyle.DropDown), "譲受人事由", em改行.改行あり)

            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "小作料"), "対価")
            .Panel.AddCtrl(New ComboList(ListResource.S_Data, "小作料単位", mvarTarget, "小作料単位", "Text", ComboBoxStyle.DropDown), "単位", em改行.改行あり)

            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, mvarTarget, "対価支払年月日"), "対価支払日", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, mvarTarget, "時期"), "引渡時期", em改行.改行あり)

            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("申請地目安", "Text"), 600), "申請地の目安", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Kanje, False, mvarTarget, "調査員B"), "担当委員")

        End With
        Set申請人世帯営農状況(pPanel, "譲渡人", "譲受人")
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("管理情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsReadOnly, mvarTarget, "公告年月日"), "公告年月日", em改行.改行あり)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("備考", "Text"), 400), "備考", em改行.改行あり)
        End With
        sub権利移動借賃等調査_様式1()
    End Sub
End Class
