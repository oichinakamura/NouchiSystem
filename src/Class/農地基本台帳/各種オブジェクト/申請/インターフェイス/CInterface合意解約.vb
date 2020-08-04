
Imports HimTools2012.controls
Imports HimTools2012.controls.DVCtrlCommonBase

Public Class CInterface合意解約
    Inherits DataViewNext申請Type2

    Public Sub New(ByVal pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget)
    End Sub

    Public Overrides Sub SetInterface(ByVal pPanel As HimTools2012.controls.FlowLayoutPanelPlus)
        申請基本1("承認")

        set譲受人(pPanel, emRO.IsReadOnly)
        set譲渡人(pPanel, "譲渡", True, True, True)

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請地情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            Set関連農地一覧(.Body, False, "農地リスト")

            .Panel.AddCtrl(New ComboList(ListResource.S_Data, "解約の理由", mvarTarget, "申請理由A", "Text", ComboBoxStyle.DropDown), "解約の理由")
            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, mvarTarget, "解約年月日"), "解約年月日", em改行.改行あり)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("申請理由B", "Text"), 267, 67), "条件出手", em改行.改行あり)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("備考", "Text"), 400), "備考", em改行.改行あり)
        End With
        sub権利移動借賃等調査_様式2()
    End Sub
End Class
