
Imports HimTools2012.controls
Imports HimTools2012.controls.DVCtrlCommonBase

Public Class CObj固定土地 : Inherits CTargetObjWithView農地台帳

    Public Sub New(ByVal pRow As DataRow, ByVal bAddNew As Boolean)
        MyBase.New(pRow, bAddNew, New HimTools2012.TargetSystem.DataKey("固定土地", pRow.Item("nID")), "M_固定情報")
        If pRow Is Nothing Then
            Stop
        End If
    End Sub

    Public Overrides Function InitDataViewNext(ByRef pDB As HimTools2012.TargetSystem.CDataViewCollection, Optional InterfaceName As String = "") As Boolean
        If Me.DataViewPage Is Nothing Then
            Me.DataViewPage = New DataViewNext固定(Me)
        End If
        Return True
    End Function

    Public Overrides Function CanDropKeyHead(ByVal sKey As String, ByVal sOption As String) As Boolean
        If Not SysAD.IsClickOnceDeployed Then
            Debug.Print(sKey)
            Stop
        End If
        Return False
    End Function
    Public Overrides ReadOnly Property DataTableWithUpdateList As HimTools2012.Data.DataTableWithUpdateList
        Get
            Return App農地基本台帳.TBL固定情報
        End Get
    End Property


    Public Overrides Function DoCommand(ByVal sCommand As String, ByVal ParamArray sParams() As String) As Object
        Return Nothing
    End Function

    Public Overrides Sub DropDown(ByVal sSourceList As String, Optional ByVal sOption As String = "")

    End Sub

    Public Overrides Function GetContextMenu(Optional ByVal pMenu As HimTools2012.controls.MenuItemEx = Nothing, Optional nDips As Integer = 1, Optional sParam() As String = Nothing) As HimTools2012.controls.MenuPlus
        Return Nothing
    End Function

    Public Overrides Function GetDataBaseView() As System.Data.DataView
        Return New DataView(App農地基本台帳.TBL固定情報, "[nID]=" & GetIntegerValue("ID"), "", DataViewRowState.CurrentRows)
    End Function

    Public Overrides Function GetProperty(ByVal sParam As String) As Object
        Return Nothing
    End Function



    Public Overrides Function SaveMyself() As Boolean
        Return False
    End Function

    Public Overrides Sub ValueChange(ByVal sField As String, ByVal pValue As Object)
        App農地基本台帳.TBL固定情報.AddUpdateListwithDataViewPage(Me, sField, pValue)
    End Sub
End Class


Public Class DataViewNext固定
    Inherits CDataViewPanel農地台帳


    Private WithEvents cmb自小作別 As ComboBoxPlus
    Private mvarGroup As HimTools2012.controls.GroupBoxPlus

    Public Sub New(pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget, App農地基本台帳.TBL固定情報, SysAD.page農家世帯.DataViewCollection, True, True)

        Dim nID As Integer = pTarget.ID

        Me.SetButtons(CreateButton("所有者を呼ぶ", "所有者を呼ぶ"), New ToolStripSeparator)

        Dim nHeight As Integer = 0
        Panel.FlowDirection = FlowDirection.LeftToRight
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("管理情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            nHeight = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsReadOnly, pTarget, "ID", , 80), "ID").Height
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsReadOnly, pTarget, "固定資産番号", , 80), "一筆コード", em改行.改行あり)
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("先行異動", "Value"), "あり", "なし"), "先行異動")
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "先行異動日", 100), "異動日")
            .Panel.AddCtrl(New DateTimePickerPlus(True, pTarget, "更新日", 100), "更新日", em改行.改行あり)
            .Panel.FitLabelWidth()
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("農地基本", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            SetTextComboWithMaster(.Panel, "大字", "大字ID", "大　字", em改行.改行なし)
            SetTextComboWithMaster(.Panel, "小字", "小字ID", "小　字", em改行.改行あり)

            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "所在", , 200), "所在(町外)", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "地番", , 100), "地  番")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "一部現況", , 40), "一部現況", em改行.改行あり)

            SetTextComboWithMaster(.Panel, "地目", "登記簿地目", "登記簿地目", em改行.改行あり)

            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "登記簿面積", , 100), "登記簿面積")

            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "共有持分分子", , 40), "共有持分")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "共有持分分母", , 40), "/", em改行.改行あり)

            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("課税地目"), "名称", "ID",
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "現況地目", , 60), "現況地目")
            ))
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "実面積", , 100), "現況面積", em改行.改行あり)

            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("農委地目"), "名称", "ID",
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "農委地目ID", , 60), "農委地目")
            ))
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "農委地目認定日", 100), "認定日", em改行.改行あり)

            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "部分面積", , 100), "部分面積", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "田面積", , 100), "田面積")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "畑面積", , 100), "畑面積")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "樹園地", , 100), "樹園地")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "採草放牧面積", , 100), "採草放牧面積")
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("所有情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "所有世帯ID", , 100, , True), "所有世帯番号")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "所有者ID", , 100), "所有者")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsReadOnly, pTarget, "所有者氏名", , 250))
            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("所有者", "所有者を呼ぶ"), , em改行.改行あり)

            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "管理世帯ID", , 100), "農地所有世帯")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "管理者ID", , 100), "農地所有者")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsReadOnly, pTarget, "管理者氏名", , 250))
            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("農地所有者", "管理者を呼ぶ"), , em改行.改行あり)

            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "登記名義人ID", , 100, , True), "登記名義")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsReadOnly, pTarget, "名義人氏名", , 250), "氏名", em改行.改行あり)
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("農地区分情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New OptionButtonPlus(Split("他…農用地外,内…農用地内,外…振興地域外", ","), nHeight, pTarget, "農業振興地域"), "農振区分", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("都計外,都計内,用途地域内,調整区域内,市街化区域内,都市計画白地", ","), nHeight, pTarget, "都市計画法"), "都市計画区分", em改行.改行あり)
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("生産緑地法", "Value"), "あり", "なし"), "生産緑地法")
            .Panel.AddCtrl(New OptionButtonPlus(Split("区域外,区域内(整備済),区域内(整備中)", ","), nHeight, pTarget, "土地改良法"), "土地改良区分", em改行.改行あり)

            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "等級", , 200), "農地区分", em改行.改行あり)
        End With

        cmb自小作別 = Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("自小作区分"), "名称", "ID",
            Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "自小作別", , 60), "自小作の別")
        ), , em改行.改行あり)

        mvarGroup = CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("自小作情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
        With mvarGroup

            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "借受世帯ID", , 100), "借受世帯番号")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "借受人ID", , 100), "借受者名")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsReadOnly, pTarget, "借受人氏名", , 250), , em改行.改行あり)
            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("借受者参照", "借受者を呼ぶ"), , em改行.改行あり)

            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("農業生産法人経由貸借", "Value"), "あり", "なし"), "法人経由の貸借")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsReadOnly, pTarget, "経由農業生産法人ID", , 100), "経由法人")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsReadOnly, pTarget, "経由農業生産法人名", , 250), "法人名", em改行.改行あり)

            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("適用法令"), "名称", "ID",
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "小作地適用法", , 60), "適用法令")
            ), , em改行.改行あり)

            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("小作形態"), "名称", "ID",
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "小作形態", , 60), "形態")
            ), , em改行.改行あり)

            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "小作料", , 100), "小作料")
            .Panel.AddCtrl(New ComboList(ListResource.S_Data, "小作料単位", mvarTarget, "小作料単位", "Text", ComboBoxStyle.DropDown), "小作料単位", em改行.改行あり)

            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "小作開始年月日", 150), "貸借期間")
            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("～", "期間設定", True))
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "小作終了年月日", 150), "", em改行.改行あり)
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("解除条件付きの農地の貸借", "Value"), "あり", "なし"), "条件付き貸借")
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("利用状況報告", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("利用状況報告対象", "Value"), "あり", "なし"), "利用状況報告")
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "利用状況報告年月日", 150), "報告年月日", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "是正勧告日", 150), "勧告年月日", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "是正内容", , 200), "是正内容", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "是正期限", 150), "是正期限", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,１号,２号,３号", ","), nHeight, mvarTarget, "根拠条件農地法"), "根拠農地法", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,１号,２号,３号", ","), nHeight, mvarTarget, "根拠条件基盤強化法"), "根拠基盤法", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "是正状況", , 200), "是正状況", em改行.改行なし)
            .Panel.AddCtrl(New DateTimePickerPlus(True, pTarget, "是正確認", 100), "是正確認日", em改行.改行あり)

            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "取消事由", , 200), "取消事由", em改行.改行あり)

            .Panel.AddCtrl(New DateTimePickerPlus(True, pTarget, "取消年月日", 100), "取消年月日", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,１号,２号", ","), nHeight, mvarTarget, "取消条件農地法"), "取消農地法", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,１号,２号", ","), nHeight, mvarTarget, "取消条件基盤強化法"), "取消基盤法", em改行.改行あり)
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("農地の利用状況調査", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("農地状況"), "名称", "ID",
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "農地状況", , 60), "農地状況")
            ))
            .Panel.AddCtrl(New DateTimePickerPlus(True, pTarget, "無断転用調査日", 100), "調査日", em改行.改行あり)

            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "利用状況", , 200), "利用状況", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "作付け作物", , 200), "作付け作物", em改行.改行あり)

            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("耕作放棄解消区分"), "名称", "ID",
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "耕作放棄解消区分", , 60), "解消区分")
            ), , em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(True, pTarget, "耕作放棄解消年月日", 100), "解消年月日", em改行.改行あり)
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("農地法３条３第１項に基づく届出", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New ComboList(ListResource.S_Data, "33届出事由", mvarTarget, "届出事由", "Text", ComboBoxStyle.DropDown), "届出事由", em改行.改行あり)

            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "届出年月日", 150), "届出年月日", em改行.改行あり)

            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "届出者氏名", , 200), "届出者氏名", em改行.改行あり)

            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("あっせん希望", "Value"), "あり", "なし"), "あっせん希望")
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("納税猶予の適用状況", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New OptionButtonPlus(Split("対象外,生前贈与の納税贈与,相続税の猶予", ","), nHeight, pTarget, "納税猶予対象農地"), "納税猶予対象", em改行.改行あり)

            .Panel.AddCtrl(New OptionButtonPlus(Split("対象外,農地取得資金,農地等購入資金", ","), nHeight, pTarget, "融資対象農地"), "融資対象農地", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,１号,２号,３号", ","), nHeight, pTarget, "租税処置法"), "処置法70条", em改行.改行あり)

            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "営農困難", , 200), "営農困難貸付", em改行.改行あり)
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("遊休化", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("遊休化", "Value"), "あり", "なし"), "遊休化")
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "遊休時期", 150), "調査日", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,１号,２号", ","), nHeight, pTarget, "調査結果農地法"), "農法30条3項", em改行.改行あり)
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("解消意向", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "解消意向", , 200), "解消意向", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "遊休指導通知期限", 150), "指導通知年月", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "遊休指導内容", , 200), "遊休指導内容", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "農業委員会の指導年月", 150), "指導年月日", em改行.改行あり)
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("遊休通知", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "市町村長の勧告年月", 150), "通知年月日", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,1号,2号,3号,ただし書", ","), nHeight, pTarget, "遊休利用増進農地法"), "農法32条", em改行.改行あり)
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("遊休解消通知", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "遊休解消届出", 150), "解消届出日", em改行.改行あり)
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("遊休解消斡旋", "Value"), "あり", "なし"), "あっせん希望")
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("遊休利用増進勧告", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "遊休利用増進勧告日", 150), "利用増進勧告", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "遊休利用増進是正期限", 150), "利用増進期限", em改行.改行あり)
            .Panel.AddCtrl(New OptionButtonPlus(Split("-,1号,2号,3号", ","), nHeight, pTarget, "遊休利用増進農地法"), "農法34条", em改行.改行あり)

            .Panel.AddCtrl(New ComboList(ListResource.S_Data, "是正勧告内容", mvarTarget, "是正勧告内容", "Text", ComboBoxStyle.DropDown), "是正勧告内容", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "遊休利用増進是正状況", , 200), "是正状況", em改行.改行あり)

            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "遊休利用増進是正確認", 150), "是正報告", em改行.改行あり)
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("所有権移転等の協議", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "利用増進協議者名", , 200), "協議者名", em改行.改行あり)

            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "利用増進通知日", 150), "協議通知", em改行.改行あり)
        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("仮登記", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New DateTimePickerPlus(False, pTarget, "仮登記日", 150), "仮登記日", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "仮登記氏名", , 200), "仮登記氏名", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "仮登記住所", , 200), "仮登記住所", em改行.改行あり)

        End With

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("その他", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("備考", "Text"), 400), "備考", em改行.改行あり)
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("貸付希望", "Value"), "あり", "なし"), "貸付希望")
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("売渡希望", "Value"), "あり", "なし"), "売渡希望")
        End With

    End Sub

    Private Sub cmb自小作別_SelectedValueChanged(sender As Object, e As System.EventArgs) Handles cmb自小作別.SelectedValueChanged
        Select Case cmb自小作別.Text
            Case "自作"
                mvarGroup.Visible = False
            Case Else
                mvarGroup.Visible = True
        End Select
    End Sub

End Class





