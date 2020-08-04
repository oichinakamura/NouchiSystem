
Imports HimTools2012.controls
Imports HimTools2012.controls.DVCtrlCommonBase

Public Class CInterface3条の3第1項農地取得
    Inherits DataViewNext申請Type1

    Public Sub New(pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget)
    End Sub



    Public Overrides Sub SetInterface(pPanel As HimTools2012.controls.FlowLayoutPanelPlus)
        申請基本1("許可")
        set譲受人(pPanel, emRO.IsReadOnly)
        set譲渡人(pPanel, "譲渡", True, True, True)

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            Set関連農地一覧(.Body, False, "農地リスト")

            '.Panel.AddCtrl(New ComboList(ListResource.S_Data, "譲渡申請理由", mvarTarget, "申請理由A", "Text", ComboBoxStyle.DropDown), "譲渡人事由")
            '.Panel.AddCtrl(New ComboList(ListResource.S_Data, "譲受申請理由", mvarTarget, "申請理由B", "Text", ComboBoxStyle.DropDown), "譲受人事由", em改行.改行あり)

            '.Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("所有権移転の種類"), "名称", "ID",
            '            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "所有権移転の種類"), "種類"), False
            '), "", em改行.改行あり)

            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("備考", "Text"), 400), "備考", em改行.改行あり) '//追加
        End With

        sub農地法管理情報(False)
        sub権利移動借賃等調査_様式1()
    End Sub
End Class
