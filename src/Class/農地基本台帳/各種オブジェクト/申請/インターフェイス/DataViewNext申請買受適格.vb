
Imports HimTools2012.controls
Imports HimTools2012.controls.DVCtrlCommonBase

Public MustInherit Class DataViewNext申請買受適格
    Inherits DataViewNext申請


    Public Sub New(ByVal pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget)
    End Sub

    Public Sub SetInterface買受適格(pPanel As HimTools2012.controls.FlowLayoutPanelPlus)
    End Sub

End Class

Public Class CInterface買受適格耕公
    Inherits DataViewNext申請

    Public Sub New(ByVal pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget)
    End Sub
    Public Overrides Sub SetInterface(ByVal pPanel As HimTools2012.controls.FlowLayoutPanelPlus)
        申請基本1("許可")

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請人情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            Set対象者一覧(.Body)
            '"AddItem","TX関連農業者数","VB.TextBox;WithLabel=農業従事者;Alignment=1;Width=1000;NewLine;","関連農業者数"
        End With
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            Set関連農地一覧(.Body, False, "農地リスト")
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("申請地目安", "Text"), 600), "申請地の目安", em改行.改行あり)

            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Kanje, False, mvarTarget, "調査員B"), "担当委員")

            ' "AddItem","TX物件","VB.TextBox;WithLabel=物件番号;Alignment=2;Height=300;","物件番号"
            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, mvarTarget, "入札期間A"), "入札期間", em改行.改行あり)

            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, mvarTarget, "売却決定日"), "売却決定期日", em改行.改行あり)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("事件等", "Text"), 400), "備　考", em改行.改行あり)
        End With
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("管理情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
        End With
    End Sub
End Class

Public Class CInterface買受適格耕競
    Inherits DataViewNext申請買受適格

    Public Sub New(ByVal pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget)
    End Sub


    Public Overrides Sub SetInterface(ByVal pPanel As HimTools2012.controls.FlowLayoutPanelPlus)
        申請基本1("許可")

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請地情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            Set関連農地一覧(.Body, False, "農地リスト")

            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("申請地目安", "Text"), 600), "申請地の目安", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Kanje, emRO.IsCanEdit, mvarTarget, "調査員B"), "担当委員")
        End With
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請人情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            Set対象者一覧(.Body)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "関連農業者数", , 67), "農業従事者", em改行.改行あり)
        End With
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "物件番号"), "物件番号")

            sub期間設定(.Panel, "入札期間A", "入札期間", "入札期間B", "", "入札期間設定", em改行.改行あり)

            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, mvarTarget, "入札期日"), "入札期日", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, mvarTarget, "売却決定日", 120), "売却決定期日", em改行.改行あり)

            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("所轄", "Text"), 400), "所轄裁判所名", em改行.改行あり)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("事件等", "Text"), 400), "競売事件名", em改行.改行あり)
        End With
    End Sub

End Class
Public Class CInterface買受適格転公
    Inherits DataViewNext申請買受適格

    Public Sub New(ByVal pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget)
    End Sub

    Public Overrides Sub SetInterface(ByVal pPanel As HimTools2012.controls.FlowLayoutPanelPlus)
        申請基本1("許可")

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請地情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            Set関連農地一覧(.Body, True, "農地リスト")

            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("申請地目安", "Text"), 600), "申請地の目安", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Kanje, False, mvarTarget, "調査員B"), "担当委員")
        End With
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請人情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            Set対象者一覧(.Body)
        End With
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Kanje, False, mvarTarget, "申請理由A"), "転用目的", em改行.改行あり)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("備考", "Text"), 400), "備考", em改行.改行あり)

            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Kanje, False, mvarTarget, "原因"), "公売原因", em改行.改行あり)

            sub期間設定(.Panel, "入札期間A", "売却実施期間", "入札期間B", "", "入札期間設定", em改行.改行あり)


            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("売買の場所", "Text"), 400), "売買の場所", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, mvarTarget, "売却決定日"), "売却決定の日", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Kanje, False, mvarTarget, "所轄"), "所轄国税局")
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("事件等", "Text"), 400), "売却区分番号", em改行.改行あり)
        End With
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
        End With
    End Sub

End Class
Public Class CInterface買受適格転競
    Inherits DataViewNext申請買受適格


    Public Sub New(ByVal pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget)
    End Sub

    Public Overrides Sub SetInterface(pPanel As HimTools2012.controls.FlowLayoutPanelPlus)
        申請基本1("許可")

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請地情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            Set関連農地一覧(.Body, True, "農地リスト")
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("申請地目安", "Text"), 600), "申請地の目安", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Kanje, False, mvarTarget, "調査員B"), "担当委員")
        End With
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請人情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            Set対象者一覧(.Body)
        End With
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Kanje, False, mvarTarget, "申請理由A"), "転用目的", em改行.改行あり)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("備考", "Text"), 400), "備考", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Kanje, False, mvarTarget, "原因"), "競売原因", em改行.改行あり)

            sub期間設定(.Panel, "入札期間A", "入札期間", "入札期間B", "", "入札期間設定", em改行.改行あり)

            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, mvarTarget, "入札期日"), "開札期日", em改行.改行あり)

            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, mvarTarget, "売却決定日"), "売却決定期日", em改行.改行なし)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Kanje, False, mvarTarget, "所轄"), "所轄裁判所名")

            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("事件等", "Text"), 400), "競売事件名", em改行.改行あり)
        End With
    End Sub

End Class