
Imports System.ComponentModel
Imports System.Drawing.Design
Imports System.Windows.Forms.Design
Imports HimTools2012.controls.DVCtrlCommonBase
Imports HimTools2012.controls

Public Class CObj営農情報
    Inherits CTargetObjWithView農地台帳

    Public Overrides Function ToString() As String
        Return Me.氏名 & "(営農情報)"
    End Function

    Public Sub New(ByVal pRow As DataRow, ByVal bAddNew As Boolean)
        MyBase.New(pRow, bAddNew, New HimTools2012.TargetSystem.DataKey("営農情報", pRow.Item("ID")), "D_世帯営農")
    End Sub

    Public Overrides Function CanDropKeyHead(ByVal sKey As String, ByVal sOption As String) As Boolean
        If Not SysAD.IsClickOnceDeployed Then
            Debug.Print(sKey)
            Stop
        End If
        Return False
    End Function
    Public Overrides ReadOnly Property DataTableWithUpdateList As HimTools2012.Data.DataTableWithUpdateList
        Get
            Return App農地基本台帳.TBL世帯営農
        End Get
    End Property

    Public ReadOnly Property 氏名() As String
        Get
            Return MyBase.GetStringValue("世帯営農氏名")
        End Get
    End Property
    Public Overrides Function DoCommand(ByVal sCommand As String, ByVal ParamArray sParams() As String) As Object
        Select Case sCommand
            Case "閉じる"
                If Me.DataViewPage IsNot Nothing Then
                    Me.DataViewPage.DoClose()
                End If
        End Select
        Return ""
    End Function

    Public Overrides Sub DropDown(ByVal sSourceList As String, Optional ByVal sOption As String = "")

    End Sub



    Public Overrides Function GetContextMenu(Optional ByVal pMenu As HimTools2012.controls.MenuItemEx = Nothing, Optional nDips As Integer = 1, Optional sParam() As String = Nothing) As HimTools2012.controls.MenuPlus
        Dim pMenuf As New HimTools2012.controls.ContextMenuEx(AddressOf ClickMenu)

        With pMenuf
            .AddMenu("閉じる", , AddressOf ClickMenu)
        End With

        Return pMenuf
    End Function

    Public Overrides Function GetProperty(ByVal sParam As String) As Object
        Return ""
    End Function


    Public Overrides Function InitDataViewNext(ByRef pDVCol As HimTools2012.TargetSystem.CDataViewCollection, Optional InterfaceName As String = "") As Boolean
        If Me.DataViewPage Is Nothing Then
            Me.DataViewPage = New DataViewNext営農情報(Me)
        End If
        Return True
    End Function

    Public Overrides Function SaveMyself() As Boolean
        Return False
    End Function

    Public Overrides Sub ValueChange(ByVal sField As String, ByVal pValue As Object)

    End Sub
End Class

Public Class DataViewNext営農情報
    Inherits CDataViewPanel農地台帳

    Private WithEvents cmb自小作別 As ComboBoxPlus
    Private mvarGroup As HimTools2012.controls.GroupBoxPlus

    Public Sub New(ByVal pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget, Nothing, SysAD.page農家世帯.DataViewCollection, True, True)
        'App農地基本台帳.TBL営農情報
        Dim nID As Integer = pTarget.ID

        Me.SetButtons(New ToolStripSeparator)

        Dim nHeight As Integer = 0
        Panel.FlowDirection = FlowDirection.LeftToRight
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("世帯情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            nHeight = .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsReadOnly, pTarget, "ID", , 80), "ID").Height
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsReadOnly, pTarget, "世帯営農氏名", , 100), "世帯主氏名", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsReadOnly, pTarget, "世帯営農住所", , 400), "住所", em改行.改行あり)
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("情報公開拒否", "Value"), "あり", "なし"), "情報公開拒否", em改行.改行あり)
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("営農情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New OptionButtonPlus(Split("不明,経営規模拡大,現状維持,経営規模縮小", ","), nHeight, pTarget, "経営計画"), "経営計画", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("不明,農業だけでやる,農業中心でやる,兼業中心でやる,農業をやめたい", ","), nHeight, pTarget, "経営意向"), "経営意向", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("拡大縮小方法"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "拡大縮小方法", , 60), "拡大縮小方法")
            ))
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "希望年数", , 100), "希望年数(年後)", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "希望面積", , 100), "希望面積(ha)", em改行.改行あり)
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("営農計画", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New OptionButtonPlus(Split("現状維持,拡大,縮小", ","), nHeight, pTarget, "経営計画米麦作"), "米麦作", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("現状維持,拡大,縮小", ","), nHeight, pTarget, "経営計画畜産"), "畜産", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("現状維持,拡大,縮小", ","), nHeight, pTarget, "経営計画果樹"), "果樹", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("現状維持,拡大,縮小", ","), nHeight, pTarget, "経営計画そさい"), "そさい", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("現状維持,拡大,縮小", ","), nHeight, pTarget, "経営計画養蚕"), "養蚕", em改行.改行あり)

        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("農機具・施設", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("農機具"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "農機具種類1", , 60), "農機具種類1")
            ))
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "農機具数量1", , 100), "農機具数量1", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("農機具"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "農機具種類2", , 60), "農機具種類2")
            ))
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "農機具数量2", , 100), "農機具数量2", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("農機具"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "農機具種類3", , 60), "農機具種類3")
            ))
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "農機具数量3", , 100), "農機具数量3", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("農機具"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "農機具種類4", , 60), "農機具種類4")
            ))
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "農機具数量4", , 100), "農機具数量4", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("農機具"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "農機具種類5", , 60), "農機具種類5")
            ))
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "農機具数量5", , 100), "農機具数量5", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("農機具"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "農機具種類6", , 60), "農機具種類6")
            ))
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "農機具数量6", , 100), "農機具数量6", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("農機具"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "農機具種類7", , 60), "農機具種類7")
            ))
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "農機具数量7", , 100), "農機具数量7", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("農機具"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "農機具種類8", , 60), "農機具種類8")
            ))
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "農機具数量8", , 100), "農機具数量8", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("農機具"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "農機具種類9", , 60), "農機具種類9")
            ))
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "農機具数量9", , 100), "農機具数量9", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("農機具"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "農機具種類10", , 60), "農機具種類10")
            ))
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "農機具数量10", , 100), "農機具数量10", em改行.改行あり)
            '.Panel.AddCtrl(New HimTools2012.controls.ButtonNext("農機具・施設追加", "農機具追加ボタン", True), "", True)
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("家畜", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("家畜"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "家畜種類1", , 60), "家畜種類1")
            ))
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "家畜数量1", , 100), "家畜数量1", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("家畜"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "家畜種類2", , 60), "家畜種類2")
            ))
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "家畜数量2", , 100), "家畜数量2", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("家畜"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "家畜種類3", , 60), "家畜種類3")
            ))
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "家畜数量3", , 100), "家畜数量3", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("家畜"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "家畜種類4", , 60), "家畜種類4")
            ))
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "家畜数量4", , 100), "家畜数量4", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("家畜"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "家畜種類5", , 60), "家畜種類5")
            ))
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "家畜数量5", , 100), "家畜数量5", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("家畜"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "家畜種類6", , 60), "家畜種類6")
            ))
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "家畜数量6", , 100), "家畜数量6", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("家畜"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "家畜種類7", , 60), "家畜種類7")
            ))
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "家畜数量7", , 100), "家畜数量7", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("家畜"), "名称", "ID",
                .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "家畜種類8", , 60), "家畜種類8")
            ))
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "家畜数量8", , 100), "家畜数量8", em改行.改行あり)
            '.Panel.AddCtrl(New HimTools2012.controls.ButtonNext("家畜追加", "家畜追加ボタン"), "", True)
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("販売収入順位", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "販売収入順位米", , 100), "米の順位", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "販売収入順位畜産", , 100), "畜産の順位", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "販売収入順位果樹", , 100), "果樹の順位", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "販売収入順位そさい", , 100), "そさいの順位", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "販売収入順位養蚕", , 100), "養蚕の順位", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "販売収入名称その他１", , 100), "販売収入名称他１", em改行.改行なし)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "販売収入順位その他１", , 100), "販売収入順位他１", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "販売収入名称その他２", , 100), "販売収入名称他２", em改行.改行なし)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "販売収入順位その他２", , 100), "販売収入順位他２", em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "販売収入名称その他３", , 100), "販売収入名称他３", em改行.改行なし)
            .Panel.AddCtrl(New HimTools2012.controls.DVCtrlCommonBase.TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "販売収入順位その他３", , 100), "販売収入順位他３", em改行.改行あり)
            '"RoundFrame","RFrame","Top=720;Left=8100;Width=1200;Height=500;"
            '.Panel.AddCtrl(New HimTools2012.controls.ButtonNext("世帯情報", "世帯呼出ボタン"), "", True)
        End With
    End Sub
End Class
