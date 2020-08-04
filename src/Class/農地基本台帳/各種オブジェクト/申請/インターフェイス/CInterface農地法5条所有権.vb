'20160411霧島

Imports HimTools2012.controls
Imports HimTools2012.controls.DVCtrlCommonBase

Public Class CInterface農地法5条所有権
    Inherits DataViewNext申請Type3
    Public Sub New(ByVal pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget)
    End Sub

    Public Overrides Sub SetInterface(ByVal pPanel As HimTools2012.controls.FlowLayoutPanelPlus)
        Dim nHeight As Integer = 申請基本1("許可")
        set譲受人(pPanel, emRO.IsCanEdit)
        set譲渡人(pPanel, "譲渡", False, True, True)

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請地情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            Set関連農地一覧(.Body, True, "農地リスト")
        End With
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請内容", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            Set申請理由A(.Panel, "転用目的", "転用目的", em改行.改行あり)

            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("所有権移転の種類"), "名称", "ID",
                     .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "所有権移転の種類"), "設定種類"), False
            ), "", em改行.改行あり)
            sub小作料設定(.Panel, "対価", "単位", em改行.改行あり)
            set転用申請情報(.Body, nHeight)
            .Panel.FitLabelWidth()
        End With
        MyBase.SetInterface転用(pPanel, True)
    End Sub

End Class
