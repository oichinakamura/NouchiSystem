
Imports HimTools2012.controls
Imports HimTools2012.controls.DVCtrlCommonBase

Public Class CInterface非農地証明願
    Inherits DataViewNext申請Type3
    Public Sub New(ByVal pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget)
    End Sub
    Public Overrides Sub SetInterface(ByVal pPanel As HimTools2012.controls.FlowLayoutPanelPlus)
        申請基本1("許可")
        set申請者(pPanel, True)

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請地情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            Set関連農地一覧(.Body, True, "農地リスト")
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請内容", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("申請理由A", "Text"), 400, 80), "申請理由", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, False, mvarTarget, "変更年月日TXT", , 120), "変更年月日", em改行.改行あり)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("意見", "Text"), 400, 80), "現況(意見)", em改行.改行あり)
        End With
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("農地区分", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            set始末書付き農地区分(.Body)
        End With

        MyBase.SetInterface転用(pPanel, False)
    End Sub
End Class
