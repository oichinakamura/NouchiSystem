
Imports HimTools2012.controls
Imports HimTools2012.controls.DVCtrlCommonBase

Public Class CInterface3条耕作権設定
    Inherits DataViewNext申請Type1

    Public Sub New(pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget)
    End Sub

    Public Overrides Sub SetInterface(pPanel As HimTools2012.controls.FlowLayoutPanelPlus)
        申請基本1("許可")
        set譲受人(pPanel, emRO.IsReadOnly)
        set譲渡人(pPanel, "譲渡", True, True, True)

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            Set関連農地一覧(.Body, False, "農地リスト", "賃借料1円10a当たり:String", "申請部分面積:Decimal")
            Set申請理由A(.Panel, "貸人事由", "譲渡申請理由", em改行.改行あり)
            Set申請理由B(.Panel, "借人事由", "譲受申請理由", em改行.改行あり)

            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("小作形態"), "名称", "ID",
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "権利種類", , 60), "形態")
            ))

            sub小作料設定(.Panel, "小作料", "単位", em改行.改行あり)
            sub期間設定(.Panel, "始期", "開始年月日", "終期", "終了年月日", "期間設定", em改行.改行なし, True, True)

            .Panel.AddCtrl(New ComboList(ListResource.S_Data, "利用権内容", mvarTarget, "利用権内容", "Text", ComboBoxStyle.DropDown), "利用権内容", em改行.改行あり)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("申請地目安", "Text"), 600), "申請地の目安", em改行.改行あり)

            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("区分地上権", "Value"), "あり", "なし"), "区分地上権", em改行.改行あり)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("区分地上権内容", "Text"), 600), "区分地上権内容", em改行.改行あり)
        End With

        Set申請人世帯営農状況(pPanel, "貸　人", "借　人")

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("3条条件", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("条件A", "Text"), 400), "条件出手", em改行.改行あり)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("条件B", "Text"), 400), "条件受手", em改行.改行あり)
        End With
        sub農地法管理情報(True)
        sub権利移動借賃等調査_様式1()
    End Sub

End Class
