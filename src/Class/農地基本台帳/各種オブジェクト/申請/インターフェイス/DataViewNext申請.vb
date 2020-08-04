'20160401霧島

Imports System.ComponentModel
Imports HimTools2012.controls
Imports HimTools2012.controls.DVCtrlCommonBase

Public MustInherit Class DataViewNext申請
    Inherits CDataViewPanel農地台帳

    Public mvarGrid申請地一覧 As GridViewNext


    Public Property 申請地一覧() As GridViewNext
        Get
            Return mvarGrid申請地一覧
        End Get
        Set(value As GridViewNext)
            mvarGrid申請地一覧 = value
        End Set
    End Property


    Public Sub New(pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget, App農地基本台帳.TBL申請, SysAD.page農家世帯.DataViewCollection, True, True)
        SetButtons()
        Dim nID As Integer = pTarget.ID

        Panel.FlowDirection = FlowDirection.LeftToRight


    End Sub

    Public MustOverride Sub SetInterface(pPanel As HimTools2012.controls.FlowLayoutPanelPlus)

    Protected Sub sub申請地区分()
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請地区分", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New ComboBoxPlus(New DataView(App農地基本台帳.DataMaster.Body, "[Class]='農地区分'", "ID", DataViewRowState.CurrentRows), "名称", "ID",
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "農地区分", , 30), "農地区分")
            ), , em改行.改行あり)
            .Panel.AddCtrl(New ComboList(ListResource.S_Data, "農地区分補足", mvarTarget, "農地区分補足", "Text", ComboBoxStyle.DropDown), "農地区分補足", em改行.改行あり)

            .Panel.AddCtrl(New ComboBoxPlus(New DataView(App農地基本台帳.DataMaster.Body, "[Class]='申請時農振区分'", "ID", DataViewRowState.CurrentRows), "名称", "ID",
              .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "農振区分", , 30), "農振区分")
            ), , em改行.改行あり)

            .Panel.AddCtrl(New ComboBoxPlus(New DataView(App農地基本台帳.DataMaster.Body, "[Class]='都市計画区分'", "ID", DataViewRowState.CurrentRows), "名称", "ID",
               .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "都市計画区分", , 30), "都市計画区分")
             ), , em改行.改行あり)
        End With
    End Sub

    Protected Sub set始末書付き農地区分(pGroupBox As HimTools2012.controls.GroupBoxPlus)
        With pGroupBox
            .Panel.AddCtrl(New ComboBoxPlus(New DataView(App農地基本台帳.DataMaster.Body, "[Class]='農地区分'", "ID", DataViewRowState.CurrentRows), "名称", "ID",
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "農地区分", , 30), "農地区分")
            ), , em改行.改行あり)

            .Panel.AddCtrl(New ComboList(ListResource.S_Data, "農地区分補足", mvarTarget, "農地区分補足", "Text", ComboBoxStyle.DropDown), "農地区分補足", em改行.改行あり)
            .Panel.AddCtrl(New ComboList(ListResource.S_Data, "不許可の例外", mvarTarget, "不許可例外", "Text", ComboBoxStyle.DropDown), "不許可の例外", em改行.改行あり)

            .Panel.AddCtrl(New ComboBoxPlus(New DataView(App農地基本台帳.DataMaster.Body, "[Class]='申請時農振区分'", "ID", DataViewRowState.CurrentRows), "名称", "ID",
              .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "農振区分", , 30), "農振区分")
            ), , em改行.改行あり)
            .Panel.AddCtrl(New ComboBoxPlus(New DataView(App農地基本台帳.DataMaster.Body, "[Class]='都市計画区分'", "ID", DataViewRowState.CurrentRows), "名称", "ID",
              .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "都市計画区分", , 30), "都市計画区分")
            ), , em改行.改行あり)

            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("始末書", "Value"), "あり", "なし"), "始末書の有無")
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("同意書", "Value"), "あり", "なし"), "同意書の有無")
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("被害防除計画書", "Value"), "あり", "なし"), "被害防除計画書")
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("理由書", "Value"), "あり", "なし"), "面積超の理由書", em改行.改行あり)
        End With
    End Sub
    Private Sub SetInterface流動化奨励金(pPanel As HimTools2012.controls.FlowLayoutPanelPlus)
        申請基本1("確定")
        set申請者(pPanel, False)
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請地情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            Set関連農地一覧(.Body, False, "農地リスト")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "数量", , 100), "面積(a)")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "条件B"), "10a当金額")
        End With
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("管理情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, mvarTarget, "公告年月日"), "利用権設定", em改行.改行あり)
        End With
    End Sub


#Region "各種項目"
    Public Function Get農地条件(Optional ByVal s農地フィールド As String = "農地リスト") As String
        Dim nID As Decimal = Val(mvarTarget.Row.Body.Item("申請者A").ToString)
        Dim sWhere As String = ""
        Dim pRow As DataRow = Nothing
        If Not nID = 0 Then
            sWhere = "[所有者ID]=" & nID & " Or [管理者ID]=" & nID
            pRow = App農地基本台帳.TBL個人.FindRowByID(nID)
            If Not IsDBNull(pRow.Item("世帯ID")) AndAlso Not pRow.Item("世帯ID") = 0 Then
                sWhere = "(" & sWhere & " Or [所有世帯ID]=" & pRow.Item("世帯ID") & ")"
            End If
        End If

        If Not IsDBNull(mvarTarget.Row.Body.Item(s農地フィールド)) Then
            Dim St As String = mvarTarget.Row.Body.Item(s農地フィールド).ToString

            St = Replace(Replace(Replace(St, "転用", ""), "農地.", ""), ";", ",")
            If St.Length = 0 Then

            Else
                sWhere = sWhere & " OR [ID] IN (" & St & ")"
            End If
        End If
        Return sWhere

    End Function

    Protected Function 申請基本1(s許可タイトル As String) As Integer
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請基本", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            Dim nHeight As Integer = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, True, mvarTarget, "ID"), "ID").Height
            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsReadOnly, mvarTarget, "更新日"), "更新日", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "総会番号"), "総会番号")

            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("申請状況"), "名称", "ID",
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, True, mvarTarget, "状態"), "状態"), True
            ), "", em改行.改行あり)

            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, mvarTarget, "受付年月日"), "受付年月日")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "受付番号"), "受付番号")
            .Panel.AddCtrl(New ComboList(ListResource.S_Data, "受付補助記号リスト", mvarTarget, "受付補助記号", "Text", ComboBoxStyle.DropDown))

            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "通年受付番号"), "通年受付番号", em改行.改行あり)

            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, mvarTarget, "許可年月日"), s許可タイトル & "年月日")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "許可番号"), s許可タイトル & "番号")
            .Panel.AddCtrl(New ComboList(ListResource.S_Data, "許可補助記号リスト", mvarTarget, "許可補助記号", "Text", ComboBoxStyle.DropDownList), s許可タイトル & "補助記号", em改行.改行あり)
            .Panel.AddCtrl(New ComboList(ListResource.S_Data, "調査員", mvarTarget, "調査員A", "Text", ComboBoxStyle.DropDown), "調査員")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Kanje, False, mvarTarget, "調査員B"), "担当委員", em改行.改行あり)

            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, mvarTarget, "調査年月日"), "調査年月日", em改行.改行なし)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "現地調査番号"), "現地調査番号", em改行.改行あり)

            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("農業委員", "([ParentKey]='" & SysAD.DB(sLRDB).DBProperty("今期農業委員会Key", "") & "' Or [ID]=0 Or [ID]=1)", "[ID] "), "名称", "ID",
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "農業委員1"), "農業委員1"), False
            ), "", em改行.改行なし)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("農業委員", "([ParentKey]='" & SysAD.DB(sLRDB).DBProperty("今期農業委員会Key", "") & "' Or [ID]=0 Or [ID]=1)", "[ID]"), "名称", "ID",
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "農業委員2"), "農業委員2"), False
            ), "", em改行.改行なし)
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("農業委員", "([ParentKey]='" & SysAD.DB(sLRDB).DBProperty("今期農業委員会Key", "") & "' Or [ID]=0 Or [ID]=1)", "[ID]"), "名称", "ID",
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "農業委員3"), "農業委員3"), False
            ), "", em改行.改行あり)
            Dim p行政書士氏名 As TextBoxPlus = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, mvarTarget, "行政書士", , 250, , True, Windows.Forms.ImeMode.Hiragana), "行政書士")
            AddHandler p行政書士氏名.Validated, AddressOf CType(mvarTarget, CObj申請).名称変更

            Return nHeight
        End With
    End Function
    Protected Sub set譲受人(ByRef pPanel As HimTools2012.controls.FlowLayoutPanelPlus, ByVal bEditable As emRO)
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("譲受人情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "申請者B", , , , True), "譲受人")
            Dim p受け人名 As TextBoxPlus = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsReadOnly, mvarTarget, "氏名B", , 250, , True, Windows.Forms.ImeMode.Hiragana))
            AddHandler p受け人名.Validated, AddressOf CType(mvarTarget, CObj申請).名称変更

            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("参照", "申請人Bを呼ぶ"), "", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, mvarTarget, "住所B", , 400, , , Windows.Forms.ImeMode.Hiragana), "住所", em改行.改行あり)
            .Panel.AddCtrl(New ComboList(ListResource.S_Data, "職業", mvarTarget, "職業B", "Text", ComboBoxStyle.DropDown), "職業")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "年齢B"), "年齢")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Kanje, False, mvarTarget, "集落B"), "集落")
        End With
    End Sub

    Protected Sub set譲渡人(pPanel As HimTools2012.controls.FlowLayoutPanelPlus, s渡人タイトル As String, b年金関連 As Boolean, b複数渡人 As Boolean, b代理人 As Boolean)
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus(s渡人タイトル & "人情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "申請世帯A"), s渡人タイトル & "世帯", Not b年金関連)
            If b年金関連 Then .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("年金関連", "Value"), "あり", "なし"), "年金関連", em改行.改行あり)
            If b複数渡人 Then .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("複数申請人A", "Value"), "あり", "なし"), "譲渡人複数印刷", em改行.改行あり)

            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "申請者A", , , , True), s渡人タイトル & "人")
            Dim p渡し人名 As TextBoxPlus = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsReadOnly, mvarTarget, "氏名A", , 250, , True, Windows.Forms.ImeMode.Hiragana))
            AddHandler p渡し人名.Validated, AddressOf CType(mvarTarget, CObj申請).名称変更

            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("参照", "申請人Aを呼ぶ"), "", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "年齢A", ), "年齢")
            .Panel.AddCtrl(New ComboList(ListResource.S_Data, "職業", mvarTarget, "職業A", "Text", ComboBoxStyle.DropDown), "職業", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, mvarTarget, "住所A", , 400, , , Windows.Forms.ImeMode.Hiragana), "住所")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Kanje, False, mvarTarget, "集落A"), "集落", em改行.改行あり)

            If b代理人 Then
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "代理人A", , , , True, Windows.Forms.ImeMode.Off), "代理人")
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, mvarTarget, "代理人名", , , , True, Windows.Forms.ImeMode.Hiragana), "代理人名", em改行.改行あり)
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, mvarTarget, "代理人住所", , 400), "住所", em改行.改行あり)
            End If

        End With
    End Sub
    Protected Sub set申請者(pPanel As HimTools2012.controls.FlowLayoutPanelPlus, ByVal b代理人 As Boolean)
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請人情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)

            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "申請世帯A"), "申請世帯")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "申請者A", , , , True), "申請人")
            Dim p申請者名 As TextBoxPlus = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsReadOnly, mvarTarget, "氏名A", , 250, , True, Windows.Forms.ImeMode.Hiragana))
            AddHandler p申請者名.Validated, AddressOf CType(mvarTarget, CObj申請).名称変更

            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("参照", "申請人Aを呼ぶ"), "", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, True, mvarTarget, "住所A", , 400), "住所", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Kanje, False, mvarTarget, "集落A"), "集落")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "年齢A"), "年齢", em改行.改行あり)
            .Panel.AddCtrl(New ComboList(ListResource.S_Data, "職業", mvarTarget, "職業A", "Text", ComboBoxStyle.DropDown), "職業", em改行.改行あり).Width = 200

            If b代理人 Then
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "代理人A"), "代理人ID")
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, mvarTarget, "代理人名", , 200), , em改行.改行あり)
                .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, mvarTarget, "代理人住所", , 400), "住所", em改行.改行あり)
            End If
        End With
    End Sub
    Protected Sub set転用申請情報(pGroupBox As HimTools2012.controls.GroupBoxPlus, nHeight As Integer)
        With pGroupBox
            'Set申請理由A(.Panel, "用途", "用途", em改行.改行あり)
            .Panel.AddCtrl(New ComboList(ListResource.S_Data, "用途", mvarTarget, "用途", "Text", ComboBoxStyle.DropDown), "用途", em改行.改行あり, 300)
            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, mvarTarget, "時期"), "転用時期", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "工事開始年1", , 60), "工事計画", em改行.改行なし)
            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("年", "", True))
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "工事開始月1", , 60), "", em改行.改行なし)
            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("月", "", True))

            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "工事終了年1", , 60), "～", em改行.改行なし)
            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("年", "", True))
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "工事終了月1", , 60), "", em改行.改行なし)
            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("月", "", True), , em改行.改行あり)

            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Real, False, mvarTarget, "数量", , 60), "棟数など", em改行.改行なし)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Real, False, mvarTarget, "建築面積", , 60), "建築面積", em改行.改行なし)

            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, False, mvarTarget, "資金計画", , 200), "資金計画", em改行.改行あり)

            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("申請地目安", "Text"), 400), "周囲の状況", em改行.改行あり)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("申請理由B", "Text"), 400), "転用事由", em改行.改行あり)

            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("意見聴取案件", "Value"), "はい", "いいえ"), "意見聴取案件", em改行.改行あり)
        End With
        set始末書付き農地区分(pGroupBox)
    End Sub

    Protected Sub Set申請人世帯営農状況(pPanel As HimTools2012.controls.FlowLayoutPanelPlus, ByVal 申請人A As String, ByVal 申請人B As String, Optional ByVal 申請人C As String = "")
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("申請人営農状況", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            Dim nWidth As Integer = 150
            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext(申請人A, ""), "営農情報").Width = nWidth
            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext(申請人B, ""), "", IIf(申請人C.Length, em改行.改行なし, em改行.改行あり)).Width = nWidth
            '                If 申請人C.Length > 0 Then .Panel.AddCtrl(New HimTools2012.controls.ButtonNext(申請人C, ""), "", em改行.改行あり).Width = nWidth

            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "経営面積A"), "経営面積").Width = nWidth
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "経営面積B"), , em改行.改行あり).Width = nWidth
            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("", ""), "借入面積").Width = nWidth
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "借入面積B"), , em改行.改行あり).Width = nWidth

            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("", ""), "世帯員数").Width = nWidth
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "世帯員数B"), , em改行.改行あり).Width = nWidth

            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("", ""), "稼動人数").Width = nWidth
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "稼動人数"), , em改行.改行あり).Width = nWidth

            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("", ""), "働手男数").Width = nWidth
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "働手男数B"), , em改行.改行あり).Width = nWidth

            .Panel.AddCtrl(New HimTools2012.controls.ButtonNext("", ""), "働手女数").Width = nWidth
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "働手女数B"), , em改行.改行あり).Width = nWidth
        End With
    End Sub
    Protected Sub Set申請理由A(pPanel As HimTools2012.controls.FlowLayoutPanelPlus, ByVal sTitle As String, ByVal sList As String, Optional bBreakLine As em改行 = em改行.改行なし)
        pPanel.AddCtrl(New ComboList(ListResource.S_Data, sList, mvarTarget, "申請理由A", "Text", ComboBoxStyle.DropDown), sTitle, bBreakLine, 300)
    End Sub
    Protected Sub Set申請理由B(pPanel As HimTools2012.controls.FlowLayoutPanelPlus, ByVal sTitle As String, ByVal sList As String, Optional bBreakLine As em改行 = em改行.改行なし)
        pPanel.AddCtrl(New ComboList(ListResource.S_Data, "sList", mvarTarget, "申請理由B", "Text", ComboBoxStyle.DropDown), sTitle, bBreakLine)
    End Sub

    Protected Function Set関連農地一覧(pBox As HimTools2012.controls.GroupBoxPlus, ByVal b転用List As Boolean, ByVal s農地フィールド As String, ParamArray sパラメータフィールド() As String) As GridViewNext
        Dim sWhere As String = Get農地条件(s農地フィールド)
        Do While sWhere.StartsWith(" ")
            sWhere = Mid(sWhere, 2)
        Loop
        Do While sWhere.ToLower.StartsWith("or ")
            sWhere = Mid(sWhere, 4)
        Loop
        Do While sWhere.ToLower.StartsWith("and ")
            sWhere = Mid(sWhere, 5)
        Loop

        Dim pTBLEx As New C申請農地一覧TBL(sWhere, pBox.View, b転用List, sパラメータフィールド)
        mvarGrid申請地一覧 = New GridViewNext(Me, pTBLEx, 800, mvarTarget, "農地リスト")

        pTBLEx.DoStart(mvarGrid申請地一覧)

        Dim mvar農地View As DataView
        App農地基本台帳.TBL農地.FindRowBySQL(sWhere)
        mvar農地View = New DataView(App農地基本台帳.TBL農地.Body, sWhere, "", DataViewRowState.CurrentRows)

        mvarGrid申請地一覧.AllowDrop = True
        mvarGrid申請地一覧.CanDropKey = New String() {"農地"}

        AddHandler mvarGrid申請地一覧.ObjectContextMenu, AddressOf ObjectContextMenu
        pBox.Panel.AddCtrl(mvarGrid申請地一覧, "関連農地一覧", em改行.改行あり)
        Return mvarGrid申請地一覧
    End Function

    Protected Function Get個人条件() As String
        Dim sWhere As String = ""

        If Not IsDBNull(mvarTarget.Row.Body.Item("個人リスト")) Then
            Dim St As String = mvarTarget.Row.Body.Item("個人リスト")

            St = Replace(Replace(Replace(Replace(St, "個人.", ""), "対象者.", ""), "世帯員.", ""), ";", ",")

            sWhere = "[ID] IN (" & St & ")"
        End If
        Return sWhere

    End Function

    Private mvar対象者View As DataView
    Protected Sub Set対象者一覧(pBox As HimTools2012.controls.GroupBoxPlus)
        Dim sWhere As String = Get個人条件()
        'If sWhere.Length > 0 Then
        '    App農地基本台帳.TBL個人.Find(sWhere)

        '    mvar対象者View = New DataView(App農地基本台帳.TBL個人, sWhere, "", DataViewRowState.CurrentRows)
        '    Dim pGrid As New GridViewNext(Me, mvar対象者View, 600, mvarTarget, "個人リスト")
        '    pGrid.CanDropKey = New String() {"個人"}

        '    AddHandler pGrid.ObjectContextMenu, AddressOf ObjectContextMenu
        '    pBox.Panel.AddCtrl(pGrid, "対象者", em改行.改行あり)
        'Else
        '    mvar対象者View = New DataView(App農地基本台帳.TBL個人, "ID=0", "", DataViewRowState.CurrentRows)
        '    Dim pGrid As New GridViewNext(Me, mvar対象者View, 600, mvarTarget, "個人リスト")
        '    pGrid.CanDropKey = New String() {"個人"}

        '    AddHandler pGrid.ObjectContextMenu, AddressOf ObjectContextMenu
        '    pBox.Panel.AddCtrl(pGrid, "対象者", em改行.改行あり)
        'End If

    End Sub

    Protected Sub sub期間設定(ByRef pPanel As HimTools2012.controls.FlowLayoutPanelPlus,
                        ByVal sField開始 As String, ByVal sTitle開始 As String,
                        ByVal sField終期 As String, ByVal sTitle終期 As String,
                        ByVal s設定ボタンコマンド As String, Optional bBreakLine As em改行 = em改行.改行なし,
                        Optional ByVal b永年 As Boolean = False, Optional ByVal b期間 As Boolean = False)
        With pPanel
            .AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, mvarTarget, sField開始), sTitle開始)
            .AddCtrl(New HimTools2012.controls.ButtonNext("～", s設定ボタンコマンド, True))
            .AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, mvarTarget, sField終期), sTitle終期, em改行.改行あり AndAlso (Not b永年) AndAlso (Not b期間))

            If b永年 Then
                .AddCtrl(New CheckButtonPlus(Me.GetBindingValue("永久", "Value"), "あり", "なし"), "永久", bBreakLine AndAlso (Not b期間))
            End If
            If b期間 Then
                .AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "期間", , 120), "期間", bBreakLine)
            End If
        End With
    End Sub

    Protected Sub sub小作料設定(ByRef pPanel As HimTools2012.controls.FlowLayoutPanelPlus, ByVal s名称 As String, ByVal s単位 As String, Optional bBreakLine As em改行 = em改行.改行なし)
        pPanel.AddCtrl(New TextBoxPlus(TextBoxMode.Real, False, mvarTarget, "小作料"), s名称)
        pPanel.AddCtrl(New ComboList(ListResource.S_Data, "小作料単位", mvarTarget, "小作料単位", "Text", ComboBoxStyle.DropDown), s単位, em改行.改行あり)
    End Sub

    Protected Sub sub農地法管理情報(ByVal b不許可 As Boolean)
        With CType(Me.Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("管理情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("備考", "Text"), 400), "備考", em改行.改行あり)

            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, True, mvarTarget, "名称", , 200), "名称", em改行.改行あり)
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, True, mvarTarget, "通年受付番号"), "通年受付番号")
            .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, True, mvarTarget, "通年許可番号"), "通年許可番号", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, mvarTarget, "総会日"), "総会日", em改行.改行あり)

            If b不許可 Then
                .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("不許可理由", "Text"), 600), "不許可理由", em改行.改行あり)
            End If

            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, mvarTarget, "取下年月日"), "取下年月日", em改行.改行あり)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("予備1", "Text"), 400), "取下げ理由", em改行.改行あり)

            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, mvarTarget, "取消年月日"), "取消年月日", em改行.改行あり)

            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("予備3", "Text"), 400), "特記事項", em改行.改行あり)
        End With

    End Sub

    Protected Sub sub転用管理情報(ByRef pPanel As HimTools2012.controls.FlowLayoutPanelPlus)
        With CType(pPanel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("転用管理情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, mvarTarget, "進達年月日"), "進達年月日", em改行.改行あり)
            .Panel.AddCtrl(New DateTimePickerPlus(emRO.IsCanEdit, mvarTarget, "完了報告年月日"), "完了報告日", em改行.改行あり)
        End With
    End Sub

    Protected Sub sub農地区分補足()
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("農地区分補足", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New ComboBoxPlus(New DataView(App農地基本台帳.DataMaster.Body, "[Class]='農地の広がり'", "ID", DataViewRowState.CurrentRows), "名称", "ID",
              .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "農地の広がり", , 30), "農地の広がり")
            ), , em改行.改行あり)
            .Panel.AddCtrl(New CheckButtonPlus(Me.GetBindingValue("土地改良事業の有無", "Value"), "あり", "なし"), "土地改良事業の有無", em改行.改行あり)

            .Panel.AddCtrl(New ComboBoxPlus(New DataView(App農地基本台帳.DataMaster.Body, "[Class]='申請書_土地改良区の意見書について'", "ID", DataViewRowState.CurrentRows), "名称", "ID",
              .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "土地改良区の意見書の有不用", , 30), "土地改良区の意見書")
            ), , em改行.改行あり).Width = 200

            .Panel.AddCtrl(New ComboBoxPlus(New DataView(App農地基本台帳.DataMaster.Body, "[Class]='申請後農地分類'", "ID", DataViewRowState.CurrentRows), "名称", "ID",
              .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, False, mvarTarget, "申請後農地分類", , 30), "申請後農地分類")
            ), , em改行.改行あり)

            .Panel.FitLabelWidth()
        End With
    End Sub
    Protected Sub ObjectContextMenu(s As Object, e As ObjectEventArgs)
        CType(ObjectMan.GetObject(e.Key).GetContextMenu(Nothing), ContextMenuStrip).Show(s, e.Location)
    End Sub
    Public Sub sub管理情報()
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("管理情報", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            .Panel.AddCtrl(New RichTextBoxNext(Me.GetBindingValue("備考", "Text"), 400), "備考", em改行.改行あり)
        End With
    End Sub
#End Region


End Class

Public MustInherit Class DataViewNext申請Type1
    Inherits DataViewNext申請

    Public Sub New(ByVal pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget)
    End Sub

    Public Overrides Sub SetInterface(pPanel As HimTools2012.controls.FlowLayoutPanelPlus)
        sub権利移動借賃等調査_様式1()
    End Sub

#Region "権利移動・借賃等調査_様式1"
    Protected WithEvents txt調査適用法令 As TextBoxPlus
    Protected WithEvents cmb調査適用法令 As ComboBoxPlus

    Protected txt権利種類 As TextBoxPlus

    Protected txt下限該当 As TextBoxPlus
    Protected cmb下限該当 As ComboBoxPlus

    Protected txt不許可の例外該当 As TextBoxPlus
    Protected cmb不許可の例外該当 As ComboBoxPlus

    Protected txt個人法人の別A As TextBoxPlus
    Protected WithEvents cmb個人法人の別A As ComboBoxPlus

    Protected txt法人の形態別 As TextBoxPlus
    Protected cmb法人の形態別 As ComboBoxPlus


    Private Sub cmb調査適用法令_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmb調査適用法令.SelectedIndexChanged
        ChangeVisible()
    End Sub
    Private Sub ChangeVisible()
        txt下限該当.Visible = (Val(txt調査適用法令.Text) = 1)
        cmb下限該当.Visible = (Val(txt調査適用法令.Text) = 1)
        txt不許可の例外該当.Visible = (Val(txt調査適用法令.Text) = 1)
        cmb不許可の例外該当.Visible = (Val(txt調査適用法令.Text) = 1)
    End Sub

    Private Sub cmb個人法人の別A_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmb個人法人の別A.SelectedIndexChanged
        Try
            txt法人の形態別.Visible = (Val(txt個人法人の別A.Text) > 1)
            cmb法人の形態別.Visible = (Val(txt個人法人の別A.Text) > 1)
        Catch ex As Exception

        End Try
    End Sub
    Protected Sub sub権利移動借賃等調査_様式1()
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("権利移動・借賃等調査_様式1", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            'ｶ
            txt調査適用法令 = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "調査適用法令", , 40), "適用法令")
            cmb調査適用法令 = .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("調査適用法令", "[ID] In (0,1,2,3,4,5,6,7)"), "名称", "ID", txt調査適用法令, emRO.IsCanEdit), "", em改行.改行あり, 200)
            'ｻ
            txt権利種類 = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "調査権利の種類", , 40), "権利の種類")
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("調査権利種類", "[ID] In (0,1,2,3,4,5,6,7,8,9,10,11,12)"), "名称", "ID", txt権利種類, emRO.IsCanEdit), "", em改行.改行あり, 200)
            '1
            txt下限該当 = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "調査農法3条2項5号", , 40), "下限面積該当")
            cmb下限該当 = .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("調査下限面積該当"), "名称", "ID", txt下限該当, emRO.IsCanEdit), "", em改行.改行あり, 200)
            '2
            txt不許可の例外該当 = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "調査農法3条2項124号", , 40), "不許可の例外該当")
            cmb不許可の例外該当 = .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("調査不許可の例外該当"), "名称", "ID", txt不許可の例外該当, emRO.IsCanEdit), "", em改行.改行あり, 200)

            With CType(.Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("権利の設定・移転を受ける者（譲受人・借人）", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
                '6
                txt個人法人の別A = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "調査個人法人の別A", , 40), "個人法人の別")
                cmb個人法人の別A = .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("個人法人の別"), "名称", "ID", txt個人法人の別A, emRO.IsCanEdit), "", em改行.改行あり, 200)
                '7
                txt法人の形態別 = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "調査法人の形態別", , 40), "法人の形態別")
                cmb法人の形態別 = .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("法人の形態別"), "名称", "ID", txt法人の形態別, emRO.IsCanEdit), "", em改行.改行あり, 200)
                '8
                Dim txt経営改善計画 As TextBoxPlus = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "調査経営改善計画の有無", , 40), "経営改善計画認定の有無")
                .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("選択有無"), "名称", "ID", txt経営改善計画, emRO.IsCanEdit), "", em改行.改行あり, 200)

                .Panel.FitLabelWidth()
            End With

            With CType(.Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("権利の設定・移転をする者（譲渡人・貸人）", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
                '9
                Dim txt個人法人の別B As TextBoxPlus = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "調査個人法人の別B", , 40), "個人法人の別")
                .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("個人法人の別"), "名称", "ID", txt個人法人の別B, emRO.IsCanEdit), "", em改行.改行あり, 200)

                .Panel.FitLabelWidth()
            End With

            .Panel.FitLabelWidth()
        End With
    End Sub

#End Region
End Class

Public MustInherit Class DataViewNext申請Type2
    Inherits DataViewNext申請

    Public Sub New(ByVal pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget)
    End Sub

    Public Overrides Sub SetInterface(pPanel As HimTools2012.controls.FlowLayoutPanelPlus)
        sub権利移動借賃等調査_様式2()
    End Sub
#Region "権利移動・借賃等調査_様式2"
    Protected txt調査適用法令貸借終了 As TextBoxPlus
    Protected WithEvents cmb調査適用法令貸借終了 As ComboBoxPlus

    Protected txt個人法人の別A貸借終了 As TextBoxPlus
    Protected WithEvents cmb個人法人の別A貸借終了 As ComboBoxPlus

    Protected txt法人の形態別貸借終了 As TextBoxPlus
    Protected cmb法人の形態別貸借終了 As ComboBoxPlus

    Protected txt貸借終了根拠条項 As TextBoxPlus
    Protected cmb貸借終了根拠条項 As ComboBoxPlus

    Protected txt利用権終了後の農地状況 As TextBoxPlus
    Protected cmb利用権終了後の農地状況 As ComboBoxPlus

    Protected txt機構法終了後の農地状況 As TextBoxPlus
    Protected cmb機構法終了後の農地状況 As ComboBoxPlus
    Protected Sub sub権利移動借賃等調査_様式2()
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("権利移動・借賃等調査_様式2", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            'ｶ
            txt調査適用法令貸借終了 = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "調査適用法令", , 40), "適用法令")
            cmb調査適用法令貸借終了 = .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("調査適用法令", "[ID] In (0,10,11,12,13,14,15)"), "名称", "ID", txt調査適用法令貸借終了, emRO.IsCanEdit), "", em改行.改行あり, 200)
            'ｻ
            Dim txt権利種類2 As TextBoxPlus = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "調査権利の種類", , 40), "権利の種類")
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("調査権利種類", "[ID] In (0,21,22,23,24,25,26,27)"), "名称", "ID", txt権利種類2, emRO.IsCanEdit), "", em改行.改行あり, 200)

            With CType(.Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("返還する者（借人）", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
                '21
                txt個人法人の別A貸借終了 = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "調査個人法人の別A", , 40), "個人法人の別")
                cmb個人法人の別A貸借終了 = .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("個人法人の別"), "名称", "ID", txt個人法人の別A貸借終了, emRO.IsCanEdit), "", em改行.改行あり, 200)
                '22
                txt法人の形態別貸借終了 = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "調査法人の形態別", , 40), "法人の形態別")
                cmb法人の形態別貸借終了 = .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("法人の形態別"), "名称", "ID", txt法人の形態別貸借終了, emRO.IsCanEdit), "", em改行.改行あり, 200)

                .Panel.FitLabelWidth()
            End With

            With CType(.Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("権利の設定・移転をする者（譲渡人・貸人）", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
                '23
                Dim txt個人法人の別B As TextBoxPlus = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "調査個人法人の別B", , 40), "個人法人の別")
                .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("個人法人の別"), "名称", "ID", txt個人法人の別B, emRO.IsCanEdit), "", em改行.改行あり, 200)

                .Panel.FitLabelWidth()
            End With

            '24
            txt貸借終了根拠条項 = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "調査許可等の根拠条項", , 40), "貸借終了の根拠条項")
            cmb貸借終了根拠条項 = .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("調査貸借終了の根拠条項"), "名称", "ID", txt貸借終了根拠条項, emRO.IsCanEdit), "", em改行.改行あり, 200)
            '25
            txt利用権終了後の農地状況 = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "調査基盤法満了農地状況", , 40), "利用権終了後の農地状況")
            cmb利用権終了後の農地状況 = .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("調査基盤強化法利用権終了後の農地状況"), "名称", "ID", txt利用権終了後の農地状況, emRO.IsCanEdit), "", em改行.改行あり, 200)
            '26
            txt機構法終了後の農地状況 = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "調査中間管理事業法満了農地状況", , 40), "機構法終了後の農地状況")
            cmb機構法終了後の農地状況 = .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("調査機構法貸借終了後の農地状況"), "名称", "ID", txt機構法終了後の農地状況, emRO.IsCanEdit), "", em改行.改行あり, 200)

            .Panel.FitLabelWidth()
        End With
    End Sub

    Private Sub cmb調査適用法令貸借終了_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmb調査適用法令貸借終了.SelectedIndexChanged
        ChangeVisible貸借終了()
    End Sub

    Private Sub ChangeVisible貸借終了()
        txt貸借終了根拠条項.Visible = ({10, 11, 12, 13}.Contains(Val(txt調査適用法令貸借終了.Text)))
        cmb貸借終了根拠条項.Visible = ({10, 11, 12, 13}.Contains(Val(txt調査適用法令貸借終了.Text)))
        txt利用権終了後の農地状況.Visible = (Val(txt調査適用法令貸借終了.Text) = 14)
        cmb利用権終了後の農地状況.Visible = (Val(txt調査適用法令貸借終了.Text) = 14)
        txt機構法終了後の農地状況.Visible = (Val(txt調査適用法令貸借終了.Text) = 15)
        cmb機構法終了後の農地状況.Visible = (Val(txt調査適用法令貸借終了.Text) = 15)
    End Sub

    Private Sub cmb個人法人の別A貸借終了_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmb個人法人の別A貸借終了.SelectedIndexChanged
        Try
            txt法人の形態別貸借終了.Visible = (Val(txt個人法人の別A貸借終了.Text) > 1)
            cmb法人の形態別貸借終了.Visible = (Val(txt個人法人の別A貸借終了.Text) > 1)
        Catch ex As Exception

        End Try
    End Sub
#End Region
End Class


Public MustInherit Class DataViewNext申請Type3
    Inherits DataViewNext申請

    Public Sub New(ByVal pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(pTarget)
    End Sub

    Public Sub SetInterface転用(pPanel As HimTools2012.controls.FlowLayoutPanelPlus, ByVal b転用管理区分 As Boolean)
        sub農地区分補足()
        sub農地法管理情報(True)
        If b転用管理区分 Then sub転用管理情報(Me.Panel)
        sub権利移動借賃等調査_様式3()
    End Sub
#Region "権利移動・借賃等調査_様式3"
    Protected txt調査適用法令転用 As TextBoxPlus
    Protected WithEvents cmb調査適用法令転用 As ComboBoxPlus
    Protected txt権利種類 As TextBoxPlus

    Protected txt転用農地区分 As TextBoxPlus
    Protected WithEvents Cmb転用農地区分 As ComboBoxPlus

    Protected txt優良農地許可判断根拠 As TextBoxPlus
    Protected Cmb優良農地許可判断根拠 As ComboBoxPlus
    Protected Sub sub権利移動借賃等調査_様式3()
        With CType(Panel.AddCtrl(New HimTools2012.controls.GroupBoxPlus("権利移動・借賃等調査_様式3", Me), "", em改行.改行あり), HimTools2012.controls.GroupBoxPlus)
            'ｶ
            txt調査適用法令転用 = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "調査適用法令", , 40), "適用法令")
            cmb調査適用法令転用 = .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("調査適用法令", "[ID] In (0,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35)"), "名称", "ID", txt調査適用法令転用, emRO.IsCanEdit), "", em改行.改行あり, 200)
            'ｻ
            txt権利種類 = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "調査権利の種類", , 40), "権利の種類")
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("調査権利種類", "[ID] In (0,31,32,33,34,35)"), "名称", "ID", txt権利種類, emRO.IsCanEdit), "", em改行.改行あり, 200)
            '31
            Dim txt許可除外条項 As TextBoxPlus = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "調査許可等の除外条項", , 40), "許可等の除外条項")
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("調査転用許可・届出・協議・公告と除外事項"), "名称", "ID", txt許可除外条項, emRO.IsCanEdit), "", em改行.改行あり, 200)
            '32
            Dim txt土地利用計画区域 As TextBoxPlus = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "調査土地利用計画区域区分", , 40), "土地利用計画区域")
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("調査土地利用計画の区域区分"), "名称", "ID", txt土地利用計画区域, emRO.IsCanEdit), "", em改行.改行あり, 200)
            '33
            Dim txt農用地区域除外 As TextBoxPlus = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "調査転用に伴う農用地区域除外", , 40), "農用地区域除外")
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("選択有無"), "名称", "ID", txt農用地区域除外, emRO.IsCanEdit), "", em改行.改行あり, 200)
            '34
            Dim txt転用主体 As TextBoxPlus = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "調査転用主体", , 40), "転用主体")
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("調査転用主体"), "名称", "ID", txt転用主体, emRO.IsCanEdit), "", em改行.改行あり, 200)
            '35
            Dim txt転用用途 As TextBoxPlus = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "調査転用用途", , 40), "転用用途")
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("調査転用用途"), "名称", "ID", txt転用用途, emRO.IsCanEdit), "", em改行.改行あり, 200)
            '36
            Dim txt一時転用該当 As TextBoxPlus = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "調査一時転用該当有無", , 40), "一時転用該当")
            .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("選択有無"), "名称", "ID", txt一時転用該当, emRO.IsCanEdit), "", em改行.改行あり, 200)
            '37
            txt転用農地区分 = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "調査転用農地区分", , 40), "転用農地区分")
            Cmb転用農地区分 = .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("調査農地の区分"), "名称", "ID", txt転用農地区分, emRO.IsCanEdit), "", em改行.改行あり, 200)
            '38
            txt優良農地許可判断根拠 = .Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, mvarTarget, "調査優良農地許可判断根拠", , 40), "優良農地許可判断根拠")
            Cmb優良農地許可判断根拠 = .Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("調査優良農地の許可判断根拠"), "名称", "ID", txt優良農地許可判断根拠, emRO.IsCanEdit), "", em改行.改行あり, 200)

            .Panel.FitLabelWidth()
        End With
    End Sub

    Private Sub cmb調査適用法令転用_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmb調査適用法令転用.SelectedIndexChanged
        txt転用農地区分.Visible = ({20, 21, 22, 24, 25, 26, 27, 28, 29, 31, 32, 33}.Contains(Val(txt調査適用法令転用.Text)))
        Cmb転用農地区分.Visible = ({20, 21, 22, 24, 25, 26, 27, 28, 29, 31, 32, 33}.Contains(Val(txt調査適用法令転用.Text)))
    End Sub

    Private Sub Cmb転用農地区分_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Cmb転用農地区分.SelectedIndexChanged
        Try
            txt優良農地許可判断根拠.Visible = ({1, 11, 12, 13, 21, 22}.Contains(Val(txt転用農地区分.Text)))
            Cmb優良農地許可判断根拠.Visible = ({1, 11, 12, 13, 21, 22}.Contains(Val(txt転用農地区分.Text)))
        Catch ex As Exception

        End Try
    End Sub
#End Region
End Class
