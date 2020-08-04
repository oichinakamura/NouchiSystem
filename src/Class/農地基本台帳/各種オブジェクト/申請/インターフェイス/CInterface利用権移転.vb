
Imports HimTools2012.controls
Imports HimTools2012.controls.DVCtrlCommonBase

Public Class CInterface利用権移転
    Inherits DataViewNext申請Type1

    Public Sub New(ByVal pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget)
    End Sub

    Public Overrides Sub SetInterface(ByVal pPanel As HimTools2012.controls.FlowLayoutPanelPlus)
        申請基本1("承認")
        set譲受人(pPanel, emRO.IsReadOnly)
        set譲渡人(pPanel, "譲渡", True, False, True)

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("出し手情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "申請世帯C"), "出手世帯")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "申請者C"), "出し手")

            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, True, mvarTarget, "氏名C", , 250))
            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("参照", "申請人Cを呼ぶ"), "", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, True, mvarTarget, "住所C", , 400), "住所", em改行.改行あり)
            .Panel.AddCtrl(New ComboList(ListResource.S_Data, "職業", mvarTarget, "職業C", "Text", ComboBoxStyle.DropDown), "職業")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "年齢C"), "年齢")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Kanje, False, mvarTarget, "集落C"), "集落")
        End With
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請地情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            Set関連農地一覧(.Body, False, "農地リスト", "賃借料1円10a当たり:String", "申請部分面積:Decimal")
            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("全選択", "全選択"), "", em改行.改行あり)

            .Panel.AddCtrl(New ComboList(ListResource.S_Data, "譲渡申請理由", mvarTarget, "申請理由A", "Text", ComboBoxStyle.DropDown), "旧受手事由")
            .Panel.AddCtrl(New ComboList(ListResource.S_Data, "譲受申請理由", mvarTarget, "申請理由B", "Text", ComboBoxStyle.DropDown), "新受手事由", em改行.改行あり)

            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("小作形態"), "名称", "ID",
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "権利種類", , 60), "形態")
            ))
            .Panel.AddCtrl(New ComboList(ListResource.S_Data, "利用権内容", mvarTarget, "利用権内容", "Text", ComboBoxStyle.DropDown), "利用権内容", em改行.改行あり)

            sub小作料設定(.Panel, "小作料", "単位", em改行.改行あり)

            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, mvarTarget, "始期"), "開始年月日")
            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("～", "期間設定", True))
            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, mvarTarget, "終期"), "終了年月日")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "期間", , 120), "期間", em改行.改行あり)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("申請地目安", "Text"), 600), "申請地の目安", em改行.改行あり)


            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("同意書", "Value"), "有", "無"), "同意書の有無")
        End With
        Set申請人世帯営農状況(pPanel, "旧受手", "新受手")

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("管理情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, mvarTarget, "公告年月日"), "公告年月日", em改行.改行あり)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("備考", "Text"), 400), "備考", em改行.改行あり)
        End With

        sub権利移動借賃等調査_様式1()
    End Sub
End Class
