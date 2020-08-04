
Imports HimTools2012.controls
Imports HimTools2012.controls.DVCtrlCommonBase

Public Class CInterface事業計画変更
    Inherits DataViewNext申請

    Public Sub New(ByVal pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget)
    End Sub

    Public Overrides Sub SetInterface(ByVal pPanel As HimTools2012.controls.FlowLayoutPanelPlus)
        申請基本1("許可")

        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("変更前情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            '.Panel.AddCtrl(New HimTools2012.controls.ButtonNext("上をコピー", "変更前複写"), "", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "申請世帯A"), "申請世帯")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "申請者A", , , , True), "申請人")
            Dim p変更後氏名 As TextBoxPlus = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsReadOnly, mvarTarget, "氏名A", , 250, , True, Windows.Forms.ImeMode.Hiragana))
            AddHandler p変更後氏名.Validated, AddressOf CType(mvarTarget, CObj申請).名称変更
            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("参照", "申請人Aを呼ぶ"), "", em改行.改行なし)
            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("クリア", "申請人Aをクリア"), "", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, True, mvarTarget, "住所A", , 400), "住所", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Kanje, False, mvarTarget, "集落A"), "集落")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "年齢A"), "年齢", em改行.改行あり)
            .Panel.AddCtrl(New ComboList(ListResource.S_Data, "職業", mvarTarget, "職業A", "Text", ComboBoxStyle.DropDown), "職業", em改行.改行あり)

            .Panel.AddCtrl(New CListViewNext(mvarTarget, "予備1", "土地所在", "登記地目", "現況地目", "面積"), "変更前農地", em改行.改行あり, 500)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsReadOnly, mvarTarget, "予備3", , 200, 20, False, ImeMode.Hiragana), "許可日")
        End With
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("変更後情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "申請世帯C"), "申請世帯")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "申請者C", , , , True), "申請人")
            Dim p変更前氏名 As TextBoxPlus = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsReadOnly, mvarTarget, "氏名C", , 250, , True, Windows.Forms.ImeMode.Hiragana))
            AddHandler p変更前氏名.Validated, AddressOf CType(mvarTarget, CObj申請).名称変更
            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("参照", "申請人Cを呼ぶ"), "", em改行.改行なし)
            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("クリア", "申請人Cをクリア"), "", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Kanje, False, mvarTarget, "住所C", , 400), "住所", em改行.改行あり)
        End With
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請地情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            Set関連農地一覧(.Body, True, "農地リスト", "申請部分面積:Decimal")
        End With
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請内容", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsReadOnly, mvarTarget, "予備2", , 250, 20, False, ImeMode.Hiragana), "変更前転用目的", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, mvarTarget, "申請理由A", , 250, 20, False, ImeMode.Hiragana), "変更後転用目的", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "工事開始年1", , 60), "工事計画", em改行.改行なし)
            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("年", "", True))
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "工事開始月1", , 60), "", em改行.改行なし)
            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("月", "", True))

            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "工事終了年1", , 60), "～", em改行.改行なし)
            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("年", "", True))
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "工事終了月1", , 60), "", em改行.改行なし)
            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("月", "", True), , em改行.改行あり)

            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("申請理由B", "Text"), 333, 60), "理 由", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(New DataView(App農地基本台帳.DataMaster.Body, "[Class]='農地区分'", "ID", DataViewRowState.CurrentRows), "名称", "ID",
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "農地区分", , 30), "農地区分")
            ), , em改行.改行なし)
            .Panel.AddCtrl(New ComboList(ListResource.S_Data, "農地区分補足", mvarTarget, "農地区分補足", "Text", ComboBoxStyle.DropDown), "農地区分補足", em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(New DataView(App農地基本台帳.DataMaster.Body, "[Class]='申請時農振区分'", "ID", DataViewRowState.CurrentRows), "名称", "ID",
              .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "農振区分", , 30), "農振区分")
            ), , em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(New DataView(App農地基本台帳.DataMaster.Body, "[Class]='都市計画区分'", "ID", DataViewRowState.CurrentRows), "名称", "ID",
              .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "都市計画区分", , 30), "都市計画区分")
            ), , em改行.改行あり)

            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("申請地目安", "Text"), 600), "申請地の目安", em改行.改行あり)

            .Panel.FitLabelWidth()
        End With
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("管理情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("備考", "Text"), 400), "備考", em改行.改行あり)
        End With
    End Sub
    'Public Overrides Sub SetInterface(ByVal pPanel As HimTools2012.controls.FlowLayoutPanelPlus)
    '    申請基本1("許可")

    '    With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("変更前情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
    '        '.Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "申請者C", , , , True), "申請人")
    '        'Dim p変更前氏名 As TextBoxPlus = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsReadOnly, mvarTarget, "氏名C", , 250, , True, Windows.Forms.ImeMode.Hiragana))
    '        'AddHandler p変更前氏名.Validated, AddressOf CType(mvarTarget, CObj申請).名称変更


    '        .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "申請世帯C"), "申請世帯")
    '        .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "申請者C"), "申請人")
    '        .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Kanje, True, mvarTarget, "氏名C", , 250))
    '        .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("参照", "申請人Cを呼ぶ"), "", em改行.改行あり)
    '        .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Kanje, False, mvarTarget, "住所C", , 400), "住所", em改行.改行あり)
    '        .Panel.AddCtrl(New CListViewNext(mvarTarget, "予備1", "土地所在", "登記地目", "現況地目", "面積"), "変更前農地", em改行.改行あり)

    '        .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsReadOnly, mvarTarget, "予備3", , 200, 20, False, ImeMode.Hiragana), "許可日")
    '    End With
    '    With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("変更後情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
    '        .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("上をコピー", "変更前複写"), "", em改行.改行あり)
    '        .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "申請世帯"), "申請世帯")
    '        .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "申請者A", , , , True), "申請人")

    '        .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Kanje, False, mvarTarget, "氏名A", , 250))
    '        .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("参照", "申請人Aを呼ぶ"), "", em改行.改行あり)
    '        .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, True, mvarTarget, "住所A", , 400), "住所", em改行.改行あり)
    '        .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Kanje, False, mvarTarget, "集落A"), "集落")
    '        .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "年齢A"), "年齢", em改行.改行あり)
    '        .Panel.AddCtrl(New ComboList(ListResource.S_Data, "職業", mvarTarget, "職業A", "Text", ComboBoxStyle.DropDown), "職業")
    '    End With
    '    With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請地情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
    '        Set関連農地一覧(.Body, True, "農地リスト")
    '    End With
    '    With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請内容", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
    '        .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsReadOnly, mvarTarget, "予備2", , 250, 20, False, ImeMode.Hiragana), "変更前転用目的", em改行.改行あり)
    '        .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, mvarTarget, "申請理由A", , 250, 20, False, ImeMode.Hiragana), "変更後転用目的", em改行.改行あり)
    '        .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "工事開始年1", , 60), "工事計画", em改行.改行なし)
    '        .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("年", "", True))
    '        .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "工事開始月1", , 60), "", em改行.改行なし)
    '        .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("月", "", True))

    '        .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "工事終了年1", , 60), "～", em改行.改行なし)
    '        .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("年", "", True))
    '        .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "工事終了月1", , 60), "", em改行.改行なし)
    '        .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("月", "", True), , em改行.改行あり)

    '        .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("申請理由B", "Text"), 333, 60), "理 由", em改行.改行あり)
    '        .Panel.AddCtrl(New ComboBoxPlus(New DataView(App農地基本台帳.DataMaster.Body, "[Class]='農地区分'", "ID", DataViewRowState.CurrentRows), "名称", "ID",
    '            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "農地区分", , 30), "農地区分")
    '        ), , em改行.改行なし)
    '        .Panel.AddCtrl(New ComboList(ListResource.S_Data, "農地区分補足", mvarTarget, "農地区分補足", "Text", ComboBoxStyle.DropDown), "農地区分補足", em改行.改行あり)
    '        .Panel.AddCtrl(New ComboBoxPlus(New DataView(App農地基本台帳.DataMaster.Body, "[Class]='申請時農振区分'", "ID", DataViewRowState.CurrentRows), "名称", "ID",
    '          .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "農振区分", , 30), "農振区分")
    '        ), , em改行.改行あり)
    '        .Panel.AddCtrl(New ComboBoxPlus(New DataView(App農地基本台帳.DataMaster.Body, "[Class]='都市計画区分'", "ID", DataViewRowState.CurrentRows), "名称", "ID",
    '          .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "都市計画区分", , 30), "都市計画区分")
    '        ), , em改行.改行あり)

    '        .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("申請地目安", "Text"), 600), "申請地の目安", em改行.改行あり)

    '        .Panel.FitLabelWidth()
    '    End With
    '    With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("管理情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
    '        .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("備考", "Text"), 400), "備考", em改行.改行あり)
    '    End With
    'End Sub
End Class
