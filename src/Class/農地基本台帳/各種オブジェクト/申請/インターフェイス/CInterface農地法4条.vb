
Imports HimTools2012.controls
Imports HimTools2012.controls.DVCtrlCommonBase

Public Class CInterface農地法4条
    Inherits DataViewNext申請Type3
    Public Sub New(ByVal pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget)
    End Sub
    Public Overrides Sub SetInterface(ByVal pPanel As HimTools2012.controls.FlowLayoutPanelPlus)
        Dim nHeight As Integer = 申請基本1("許可")
        set申請者(pPanel, True)

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請地情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            Set関連農地一覧(.Body, True, "農地リスト")
        End With
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請内容", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            Set申請理由A(.Panel, "転用目的", "転用目的", em改行.改行あり)
            set転用申請情報(.Body, nHeight)

            .Panel.FitLabelWidth()
        End With
        MyBase.SetInterface転用(pPanel, True)
    End Sub
End Class

Public Class CInterface農地法4条一時転用
    Inherits DataViewNext申請Type3
    Public Sub New(ByVal pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget)
    End Sub
    Public Overrides Sub SetInterface(ByVal pPanel As HimTools2012.controls.FlowLayoutPanelPlus)
        Dim nHeight As Integer = 申請基本1("許可")
        set申請者(pPanel, True)

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請地情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            Set関連農地一覧(.Body, True, "農地リスト")
        End With
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請内容", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            Set申請理由A(.Panel, "転用目的", "転用目的", em改行.改行あり)
            sub期間設定(.Panel, "始期", "開始年月日", "終期", "終了年月日", "期間設定", em改行.改行あり, True, True)
            set転用申請情報(.Body, nHeight)

            .Panel.FitLabelWidth()
        End With
        MyBase.SetInterface転用(pPanel, True)
    End Sub
End Class
