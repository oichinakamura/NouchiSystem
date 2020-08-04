
Imports HimTools2012.controls
Imports HimTools2012.controls.DVCtrlCommonBase

Public Class CInterface農地改良届
    Inherits DataViewNext申請

    Public Sub New(ByVal pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget)
    End Sub

    Public Overrides Sub SetInterface(ByVal pPanel As HimTools2012.controls.FlowLayoutPanelPlus)
        申請基本1("許可")
        set申請者(pPanel, False)
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請地情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            Set関連農地一覧(.Body, False, "農地リスト")
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請内容", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New ComboBoxPlus(New DataView(App農地基本台帳.DataMaster.Body, "[Class]='農委地目'", "ID", DataViewRowState.CurrentRows), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "区分", , 30), "変更後地目")
            ), , em改行.改行あり)

            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Kanje, False, mvarTarget, "用途", , 120), "用途", em改行.改行あり)

            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("予備1", "Text"), 400), "工事内容", em改行.改行あり)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("意見", "Text"), 400), "意見", em改行.改行あり)
        End With

        sub申請地区分()
        sub管理情報()
    End Sub
End Class
