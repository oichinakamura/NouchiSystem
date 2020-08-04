
Imports HimTools2012.controls
Imports HimTools2012.controls.DVCtrlCommonBase

Public Class CInterfaceあっせん申出受
    Inherits DataViewNext申請

    Public Sub New(ByVal pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget)
    End Sub


    Public Overrides Sub SetInterface(ByVal pPanel As HimTools2012.controls.FlowLayoutPanelPlus)
        申請基本1("確定")
        set申請者(pPanel, False)

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("経営概況", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("営農類型"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "営農類型", , 60), "営農類型")
            ), , em改行.改行あり)
            '
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "経営面積A", , 200), "作付面積", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "経営面積B", , 200), "経営面積", em改行.改行あり)
            '
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "働手男数B", , 200), "働手男数", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "働手女数B", , 200), "働手女数", em改行.改行あり)

        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)

            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("あっせん区分"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "区分", , 60), "内　容")
            ), , em改行.改行あり)

            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Kanje, False, mvarTarget, "用途"), "希望地目")
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "数量", , 100), "希望面積", em改行.改行あり)

            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("条件A", "Text"), 400), "条　件", em改行.改行あり)
        End With
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("管理情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("備考", "Text"), 400), "備考", em改行.改行あり)
        End With
    End Sub
End Class
