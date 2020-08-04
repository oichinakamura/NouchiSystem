'20160314霧島

Imports HimTools2012
Imports HimTools2012.CommonFunc
Imports HimTools2012.DateFunctions
Imports HimTools2012.controls.DVCtrlCommonBase
Imports HimTools2012.TargetSystem

Public Enum enum申請状況
    受付 = 0
    審査 = 1
    許可_承認 = 2
    取下げ = 4
    取消し = 5
    不許可 = 42 'D_申請およびM_BASICALLに「42」で登録されているため
End Enum
Public Enum enum申請時農地区分
    なし = 0
    n1種農地 = 1
    n2種農地 = 2
    n3種農地 = 3
    甲種農地 = 4
    農用地区域内農地 = 5
End Enum

Public Enum enum転用目的種類
    未入力 = 0
    山林 = 1
    一般住宅 = 2
    農家住宅 = 3
    畜舎_堆肥舎 = 4
    倉庫_資材置場 = 5
    工場 = 6
    その他の住宅 = 7
    駐車場 = 8
    店舗_事務所 = 9
    その他 = 10
End Enum

Public Class CObj申請 : Inherits CTargetObjWithView農地台帳
    Private mvarParent As CTBL申請

    Public Sub New(ByVal pRow As DataRow, ByVal bAddNew As Boolean)
        MyBase.New(pRow, bAddNew, New HimTools2012.TargetSystem.DataKey("申請", pRow.Item("ID")), "D_申請")
    End Sub

#Region "プロパティ"
    Public Property 法令() As enum法令
        Get
            Return GetIntegerValue("法令")
        End Get
        Set(ByVal value As enum法令)
            ValueChange("法令", value)
        End Set
    End Property

    Public Property 名称() As String
        Get
            Return GetStringValue("名称")
        End Get
        Set(ByVal value As String)
            ValueChange("名称", value)
        End Set
    End Property

    Public Property 状態() As enum申請状況
        Get
            Return GetIntegerValue("状態")
        End Get
        Set(ByVal value As enum申請状況)
            ValueChange("状態", value)
        End Set
    End Property

    Public Property 受付番号() As Integer
        Get
            Return GetIntegerValue("受付番号")
        End Get
        Set(ByVal value As Integer)
            ValueChange("受付番号", value)
        End Set
    End Property

    Public Property 受付年月日() As DateTime
        Get
            Return GetDateValue("受付年月日")
        End Get
        Set(ByVal value As DateTime)
            ValueChange("受付年月日", value)
        End Set
    End Property

    Public Property 許可番号() As Integer
        Get
            Return GetIntegerValue("許可番号")
        End Get
        Set(ByVal value As Integer)
            ValueChange("許可番号", value)
        End Set
    End Property

    Public Property 許可年月日() As DateTime
        Get
            Return GetDateValue("許可年月日")
        End Get
        Set(ByVal value As DateTime)
            ValueChange("許可年月日", value)
        End Set
    End Property

    Public Property 受付補助記号() As String
        Get
            Return GetStringValue("受付補助記号")
        End Get
        Set(ByVal value As String)
            ValueChange("受付補助記号", value)
        End Set
    End Property

    Public ReadOnly Property 対価_賃借料() As String
        Get
            Return GetStringValue("小作料")
        End Get
    End Property
    Public ReadOnly Property 単位() As String
        Get
            Return GetStringValue("小作料単位")
        End Get
    End Property
    '"所有権移転の種類"
    Public ReadOnly Property 所有権移転の種類() As String
        Get
            Return GetStringValue("所有権移転の種類")
        End Get
    End Property

    Public Property 農地リスト() As String
        Get
            Return GetStringValue("農地リスト")
        End Get
        Set(ByVal value As String)
            ValueChange("農地リスト", value)
        End Set
    End Property

    Public Property 申請世帯A() As Integer
        Get
            Return GetIntegerValue("申請世帯A")
        End Get
        Set(ByVal value As Integer)
            ValueChange("申請世帯A", value)
        End Set
    End Property

    Public Property 申請者A() As Decimal
        Get
            Return GetDecimalValue("申請世帯A")
        End Get
        Set(ByVal value As Decimal)
            ValueChange("申請世帯A", value)
        End Set
    End Property

    Public Property 氏名A() As String
        Get
            Return GetStringValue("氏名A")
        End Get
        Set(ByVal value As String)
            ValueChange("氏名A", value)
        End Set
    End Property
    Public Property 職業A() As String
        Get
            Return GetStringValue("職業A")
        End Get
        Set(ByVal value As String)
            ValueChange("職業A", value)
        End Set
    End Property

    Public Property 経営面積A() As Decimal
        Get
            Return GetDecimalValue("経営面積A")
        End Get
        Set(ByVal value As Decimal)
            ValueChange("経営面積A", value)
        End Set
    End Property

    Public Property 住所A() As String
        Get
            Return GetStringValue("住所A")
        End Get
        Set(ByVal value As String)
            ValueChange("住所A", value)
        End Set
    End Property

    Public Property 氏名B() As String
        Get
            Return GetStringValue("氏名B")
        End Get
        Set(ByVal value As String)
            ValueChange("氏名B", value)
        End Set
    End Property
    Public Property 申請者B() As Integer
        Get
            Return GetIntegerValue("申請者B")
        End Get
        Set(ByVal value As Integer)
            ValueChange("申請者B", value)
        End Set
    End Property

    Public Property 申請者C() As Integer
        Get
            Return GetIntegerValue("申請者C")
        End Get
        Set(ByVal value As Integer)
            ValueChange("申請者C", value)
        End Set
    End Property

    Public Sub 名称変更()
        If Me.名称 <> Me.名称作成 Then
            ValueChange("名称", Me.名称作成)
            Me.SaveMyself()
        End If
    End Sub

    Public Property 権利種類() As Integer
        Get
            Return GetIntegerValue("調査権利の種類")
        End Get
        Set(ByVal value As Integer)
            ValueChange("調査権利の種類", value)
        End Set
    End Property
#End Region

    Public Overrides Function ToString() As String
        Return Me.名称
    End Function

    Public Overrides Function SaveMyself() As Boolean
        If Me.DataViewPage IsNot Nothing Then
            With CType(Me.DataViewPage, DataViewNext申請)
                'If .申請地一覧 IsNot Nothing Then
                '    If .申請地一覧.Columns.Contains() Then
                '    End If
            End With

        End If

        Return MyBase.SaveBase("D_申請")

    End Function
    Public Overrides ReadOnly Property DataTableWithUpdateList As HimTools2012.Data.DataTableWithUpdateList
        Get
            Return App農地基本台帳.TBL申請
        End Get
    End Property


    Public Overrides Function GetProperty(ByVal sParam As String) As Object
        Select Case sParam
            Case "土地IDリスト"
                Dim St As String = IIF(Not IsDBNull(mvarRow.Item("農地リスト")), mvarRow.Item("農地リスト").ToString, "NULL")
                Return Replace(Replace(Replace(St, "転用農地.", ""), "農地.", ""), ";", ",")

                'Case "名称"
                '    If Len(DVProperty.Rs.Value("名称")) Then
                '        CDataviewSK_GetProperty2 = DVProperty.Rs.Value("名称")
                '    Else
                '        Select Case DVProperty.Rs.Value("法令")
                '            Case state4条転用 : CDataviewSK_GetProperty2 = DVProperty.Rs.Value("氏名A") & "(４条転用)"
                '            Case Else
                '                Debug.Assert(False)
                '        End Select
                '        DVProperty.Rs.Update("名称", CDataviewSK_GetProperty2)
                '    End If
            Case "申請世帯A" : Return IIF(IsDBNull(mvarRow.Item("申請世帯A")), 0, mvarRow.Item("申請世帯A"))
            Case "Obj申請者A" : Return ObjectMan.GetObject("個人." & Val(GetDecimalValue("申請者A").ToString))
            Case "Obj申請者B" : Return ObjectMan.GetObject("個人." & Val(GetDecimalValue("申請者B").ToString))
                'Case "土地IDリスト"
                '    St = Fnc.NullCast(DVProperty.Rs.Value("農地リスト"), "NULL")
                '    CDataviewSK_GetProperty2 = Replace(Replace(Replace(St, "転用農地.", ""), "農地.", ""), ";", ",")

                'Case "個人リスト" : CDataviewSK_GetProperty2 = Replace(Replace(DVProperty.Rs.Value("個人リスト"), "対象者.", ""), ";", ",")

            Case ""
            Case Else

                Return mvarRow.Item(sParam).ToString
        End Select

        Return ""
    End Function


    Public Overrides Function GetContextMenu(Optional ByVal pMenu As HimTools2012.controls.MenuItemEx = Nothing, Optional nDips As Integer = 1, Optional sParam() As String = Nothing) As HimTools2012.controls.MenuPlus
        Dim pMenuPlus As HimTools2012.controls.MenuPlus = CreateMenu(pMenu)

        Select Case Me.状態
            Case enum申請状況.受付 : SetContextMenu受付中(pMenuPlus)
            Case enum申請状況.審査 : SetContextMenu審査中(pMenuPlus)
            Case enum申請状況.許可_承認 : SetContextMenu許可済(pMenuPlus)
            Case enum申請状況.取下げ : SetContextMenu取下げ(pMenuPlus)
            Case enum申請状況.取消し : SetContextMenu取消し(pMenuPlus)
            Case enum申請状況.不許可 : SetContextMenu不許可(pMenuPlus)
        End Select

        With pMenuPlus
            '申請農地を地図に表示;
            '名称の修正
        End With

        SetDVMenu(pMenuPlus, pMenu)
        Return pMenuPlus
    End Function

    Public Sub SetContextMenu受付中(ByRef pMenu As HimTools2012.controls.MenuPlus)
        Dim bEdit As Boolean = (SysAD.SystemInfo.ユーザー.n権利 > 0)

        With pMenu
            .AddMenu("開く", , AddressOf ClickMenu)
            .InsertSeparator()
            If Me.法令 <= enum法令.農地法5条貸借 Or enum法令.農地法5条一時転用 Or Me.法令 = enum法令.非農地証明願 Or Me.法令 = enum法令.農地利用目的変更 Or Me.法令 = enum法令.事業計画変更 Or (Me.法令 >= enum法令.買受適格耕公 And Me.法令 <= enum法令.買受適格転競) Then
                .AddMenu("受付・交付簿", , AddressOf ClickMenu)
            End If

            .AddMenu("受理証明書発行", , AddressOf ClickMenu)
            pMenu.InsertSeparator()

            Select Case Me.法令
                Case enum法令.農地法3条所有権, enum法令.農地法3条耕作権
                    .AddMenuByText({"許可番号の設定", "-", "許可する", "審査にする", "取下げ", "不許可", "履歴だけ作成して許可状態"},
                                   AddressOf ClickMenu, bEdit)
                Case enum法令.農地法3条の3第1項
                    .AddMenuByText({"処理する", "不許可"},
                                   AddressOf ClickMenu, bEdit)
                Case enum法令.農地法4条, enum法令.農地法4条一時転用
                    .AddMenuByText({"許可番号の設定", "-", "許可する", "審査にする", "取下げ", "不許可", "履歴だけ作成して許可状態", "-", "４条：申請者を呼ぶ", "４条：転用済み農地の一覧", "-", "申請情報をCSVファイルで出力"},
                                   AddressOf ClickMenu, bEdit)
                Case enum法令.農地法5条所有権, enum法令.農地法5条貸借, enum法令.農地法5条一時転用
                    .AddMenuByText({"許可番号の設定", "-", "許可する", "審査にする", "取下げ", "不許可", "-", "申請情報をCSVファイルで出力"},
                                   AddressOf ClickMenu, bEdit)
                Case enum法令.利用権設定, enum法令.基盤強化法所有権, enum法令.利用権移転
                    .AddMenuByText({"承認する", "審査にする", "不許可", "取下げ", "履歴だけ作成して許可状態", "-", "申請情報をCSVファイルで出力"},
                                   AddressOf ClickMenu, bEdit)
                Case enum法令.農地法18条解約, enum法令.農地法20条解約, enum法令.合意解約, enum法令.中間管理機構へ農地の返還
                    .AddMenuByText({"処理する", "不許可", "取下げ", "-", "通知書の発行"},
                                   AddressOf ClickMenu, bEdit)
                Case enum法令.奨励金交付A, enum法令.奨励金交付B
                    .AddMenuByText({"許可する", "不許可", "取下げ", "-", "申請書一括印刷", "受付設定"},
                                   AddressOf ClickMenu, bEdit)
                Case enum法令.事業計画変更, enum法令.農地利用目的変更, enum法令.農用地計画変更
                    .AddMenuByText({"承認する", "不許可", "取下げ", "審査にする"},
                                   AddressOf ClickMenu, bEdit)
                Case enum法令.あっせん出手, enum法令.あっせん受手
                    .AddMenuByText({"承認する", "不許可", "取下げ", "審査にする"},
                                   AddressOf ClickMenu, bEdit)
                Case enum法令.非農地証明願
                    .AddMenuByText({"承認する", "不許可", "取下げ", "-"},
                                   AddressOf ClickMenu, bEdit)
                Case enum法令.買受適格耕競, enum法令.買受適格耕公, enum法令.買受適格転競, enum法令.買受適格転公
                    .AddMenuByText({"承認する", "不許可", "取下げ", "-"},
                                   AddressOf ClickMenu, bEdit)
            End Select
            pMenu.InsertSeparator()
            Context申請者A(pMenu)
            Context申請者B(pMenu, Me.法令)
            pMenu.AddMenu("関連ファイルのリンク", , AddressOf ClickMenu)
            pMenu.InsertSeparator()

            pMenu.AddMenu("削除", , AddressOf ClickMenu)
        End With
    End Sub

    Public Sub SetContextMenu審査中(ByRef pMenu As HimTools2012.controls.ContextMenuEx)
        Dim bEdit As Boolean = (SysAD.SystemInfo.ユーザー.n権利 > 0)

        With CType(pMenu, HimTools2012.controls.MenuPlus)
            .AddMenu("開く", , AddressOf ClickMenu)
            .InsertSeparator()
            Select Case Me.法令
                Case enum法令.農地法3条所有権, enum法令.農地法3条耕作権
                    .AddMenuByText({"許可番号の設定", "-", "許可する", "履歴だけ作成して許可状態", "-", "受付に戻す"},
                                   AddressOf ClickMenu, bEdit)
                Case enum法令.農地法4条, enum法令.農地法4条一時転用
                    .AddMenuByText({"許可番号の設定", "-", "許可する", "履歴だけ作成して許可状態", "-", "受付に戻す"},
                                   AddressOf ClickMenu, bEdit)
                Case enum法令.農地法5条所有権, enum法令.農地法5条貸借, enum法令.農地法5条一時転用
                    .AddMenuByText({"許可番号の設定", "-", "許可する", "履歴だけ作成して許可状態", "-", "受付に戻す"},
                                   AddressOf ClickMenu, bEdit)
                Case enum法令.利用権設定, enum法令.基盤強化法所有権, enum法令.利用権移転
                    .AddMenuByText({"許可番号の設定", "-", "承認する", "履歴だけ作成して許可状態", "-", "受付に戻す"},
                                   AddressOf ClickMenu, bEdit)
                Case enum法令.事業計画変更, enum法令.農地利用目的変更, enum法令.農用地計画変更
                    .AddMenuByText({"許可番号の設定", "-", "承認する", "履歴だけ作成して許可状態", "-", "受付に戻す"},
                                   AddressOf ClickMenu, bEdit)
                Case enum法令.あっせん出手, enum法令.あっせん受手
                    .AddMenuByText({"許可番号の設定", "-", "承認する", "履歴だけ作成して許可状態", "-", "受付に戻す"},
                                   AddressOf ClickMenu, bEdit)
                    '            Case enum法令.利用権設定
                    '                St = St & n & "承認する;" & n & "受付に戻す;"
                    '            Case state事業計画変更, state農用地計画変更, state農地利用変更
                    '                St = St & n & "承認する;" & n & "受付に戻す;"
                    '            Case stateあっせん出手, stateあっせん受手 : St = St & n & "受付に戻す;"
                    '            Case Else
                    '                St = St & n & "許可する;" & n & "受付に戻す;"
            End Select
        End With
    End Sub

    Public Sub SetContextMenu許可済(ByRef pMenu As HimTools2012.controls.ContextMenuEx)
        Dim bEdit As Boolean = (SysAD.SystemInfo.ユーザー.n権利 > 0)
        With CType(pMenu, HimTools2012.controls.MenuPlus)
            .AddMenu("開く", , AddressOf ClickMenu)
            .InsertSeparator()
            Select Case Me.法令
                Case enum法令.農地法3条所有権, enum法令.農地法3条耕作権
                    .AddMenuByText({"受付に戻す", "取消し", "-", "許可書再発行"},
                                   AddressOf ClickMenu, bEdit)
                Case enum法令.農地法3条の3第1項
                    .AddMenuByText({"受付に戻す", "取消し", "-", "削除"},
                                   AddressOf ClickMenu, bEdit)
                Case enum法令.農地法4条, enum法令.農地法4条一時転用
                    .AddMenuByText({"受付に戻す", "取消し", "-", "事業計画変更申請", "工事進捗状況報告書", "-", "許可書の再発行"},
                                   AddressOf ClickMenu, bEdit) ' ４条：転用済み農地の一覧;
                Case enum法令.農地法5条所有権, enum法令.農地法5条貸借, enum法令.農地法5条一時転用
                    .AddMenuByText({"受付に戻す", "取消し", "-", "事業計画変更申請", "工事進捗状況報告書", "-", "許可書の再発行"},
                                   AddressOf ClickMenu, bEdit)
                Case enum法令.農地法18条解約, enum法令.合意解約, enum法令.中間管理機構へ農地の返還
                    .AddMenuByText({"受付に戻す", "取消し"},
                                   AddressOf ClickMenu, bEdit)
                Case enum法令.基盤強化法所有権, enum法令.利用権移転
                    .AddMenuByText({"受付に戻す", "取消し"},
                                   AddressOf ClickMenu, bEdit)
                Case enum法令.利用権設定
                    .AddMenuByText({"受付に戻す", "取消し", "-", "決定通知書"},
                                   AddressOf ClickMenu, bEdit)
                Case enum法令.あっせん出手, enum法令.あっせん受手
                    .AddMenuByText({"受付に戻す", "取消し"},
                                   AddressOf ClickMenu, bEdit)
                Case enum法令.農地改良届, enum法令.農地利用目的変更, enum法令.農用地計画変更
                    .AddMenuByText({"受付に戻す", "取消し"},
                                   AddressOf ClickMenu, bEdit)
                Case enum法令.非農地証明願
                    .AddMenuByText({"受付に戻す", "取消し"},
                                   AddressOf ClickMenu, bEdit)
                Case enum法令.事業計画変更
                    .AddMenuByText({"受付に戻す", "取消し", "工事進捗状況報告書", "許可書の再発行"},
                                   AddressOf ClickMenu, bEdit)
                Case enum法令.買受適格耕競, enum法令.買受適格耕公, enum法令.買受適格転競, enum法令.買受適格転公
                    .AddMenuByText({"受付に戻す", "取消し"},
                                   AddressOf ClickMenu, bEdit)
                Case Else
                    If Not SysAD.IsClickOnceDeployed Then
                        Stop
                    End If
                    ' Case state奨励金交付A, state奨励金交付B : St = St & n & "審査に戻す;-;交付設定;交付決定通知書;"
            End Select
            .InsertSeparator()
            Context申請者A(pMenu)
            Context申請者B(pMenu, Me.法令)
        End With
    End Sub

    Public Sub SetContextMenu取下げ(ByRef pMenu As HimTools2012.controls.ContextMenuEx)
        With pMenu
            .AddMenu("開く", , AddressOf ClickMenu)
            .AddMenu("受付に戻す", , AddressOf ClickMenu)
            .InsertSeparator()
        End With
        Context申請者A(pMenu)
        Context申請者B(pMenu, Me.法令)
    End Sub
    Public Sub SetContextMenu取消し(ByRef pMenu As HimTools2012.controls.ContextMenuEx)
        With pMenu
            .AddMenu("開く", , AddressOf ClickMenu)
            .AddMenu("許可に戻す", , AddressOf ClickMenu)
            .InsertSeparator()
        End With
        Context申請者A(pMenu)
        Context申請者B(pMenu, Me.法令)
    End Sub

    Public Sub SetContextMenu不許可(ByRef pMenu As HimTools2012.controls.ContextMenuEx)
        With pMenu
            .AddMenu("開く", , AddressOf ClickMenu)
            .AddMenu("受付に戻す", , AddressOf ClickMenu)
            .InsertSeparator()
        End With
        Context申請者A(pMenu)
        Context申請者B(pMenu, Me.法令)
    End Sub

    Private Sub Context申請者A(ByRef pMenu As HimTools2012.controls.MenuPlus, Optional ByVal nDips As Integer = 2)
        If Not Row.IsZero("申請者A") Then
            pMenu.AddSubMenu("申請者A", nDips, ObjectMan.GetObject("個人." & Me.GetDecimalValue("申請者A")), My.Resources.Resource1.個人.ToBitmap)
        End If
    End Sub
    Private Sub Context申請者B(ByRef pMenu As HimTools2012.controls.MenuPlus, ByVal n法令 As enum法令, Optional ByVal nDips As Integer = 2)
        If Not Row.IsZero("申請者B") Then
            pMenu.AddSubMenu("申請者B", nDips, ObjectMan.GetObject("個人." & Me.GetDecimalValue("申請者B")), My.Resources.Resource1.個人.ToBitmap)
        End If
    End Sub

    Public Sub RentEnd(ByVal sKey As String, ByVal dt解約日 As Date, ByVal st解約 As String)
        'Dim Rs As NK97.RecordsetEx
        Dim p申請Row As DataRow = App農地基本台帳.TBL申請.FindRowByID(GetKeyCode(sKey))
        If p申請Row IsNot Nothing Then
            Dim p申請 As New CObj申請(p申請Row, False)

            Dim StL As String = p申請Row.Item("農地リスト").ToString
            Dim Cn() As String = Split(StL, ";")

            For K As Integer = LBound(Cn) To UBound(Cn)
                Dim St As New System.Text.StringBuilder

                If Len(Cn(K)) = 0 Then
                ElseIf GetKeyHead(Cn(K)) = "農地" Then
                    Dim p農地 As CObj農地 = ObjectMan.GetObject(Cn(K))
                    If p農地 IsNot Nothing Then
                        Dim p出し手 As CObj個人
                        If st解約 = "農地返還" Then
                            p出し手 = ObjectMan.GetObject("個人." & p農地.経由法人ID)
                        Else
                            p出し手 = ObjectMan.GetObject("個人." & p農地.所有者ID)
                        End If
                        Dim p受け手 As CObj個人 = ObjectMan.GetObject("個人." & p農地.借受人ID)
                        St.Append(p出し手.氏名 & "→" & p受け手.氏名 & "の貸借の解約・終了")

                        If Not IsDBNull(p農地.小作開始年月日) Then St.Append(vbCrLf & " 設定期間:[" & 和暦Format(p農地.小作開始年月日) & "]") Else St.Append(vbCrLf & " 設定期間:[??/??/??]")
                        If Not IsDBNull(p農地.小作終了年月日) Then St.Append("～" & "[" & 和暦Format(p農地.小作終了年月日) & "]") Else St.Append("～" & "[??/??/??]")

                        If st解約 = "農地返還" Then
                            p農地.ValueChange("借受人ID", p農地.経由法人ID)
                            p農地.ValueChange("農業生産法人経由貸借", False)
                            p農地.ValueChange("経由農業生産法人ID", 0)
                            'SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [借受人ID]={1}, [D:農地Info].農業生産法人経由貸借 = False, [D:農地Info].経由農業生産法人ID = 0 WHERE [ID]={0}", p農地.ID, p農地.経由法人ID)
                        Else
                            p農地.自小作別 = 0
                            p農地.ValueChange("自小作別", 0)
                            'SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [自小作別]=0 WHERE [ID]={0}", p農地.ID)
                        End If
                        p農地.SaveMyself()

                        Try
                            Make農地履歴(p農地.ID, Now.Date, dt解約日, 10210, p申請.法令, St.ToString, p申請.申請者A, p申請.申請者B, Val(GetKeyCode(sKey)))
                        Catch ex As Exception
                            MsgBox("履歴の作成に失敗しました。内容：" & ex.Message)
                        End Try

                        Me.DoCommand("閉じる")
                    End If
                End If
            Next

            Me.SetItem("状態", 2)
            Me.SetItem("許可年月日", dt解約日)

            Me.SaveMyself()
        End If

    End Sub


    Public Function 名称作成() As String
        Select Case Me.法令
            Case enum法令.農地法5条所有権 : Return GetStringValue("氏名A") & "(5条所有権)→" & GetStringValue("氏名B")
            Case enum法令.農地法5条貸借 : Return GetStringValue("氏名A") & "(5条貸借)→" & GetStringValue("氏名B")
            Case enum法令.農地法5条一時転用 : Return GetStringValue("氏名A") & "(5条貸借一時)→" & GetStringValue("氏名B")
            Case Else
                Return Me.名称
        End Select
        Return ""
    End Function

    Public Overrides Function DoCommand(ByVal sCommand As String, ByVal ParamArray sParams() As String) As Object
        Select Case sCommand
            Case "開く", "データを開く"
                If SysAD.page農家世帯 Is Nothing Then
                    SysAD.page農家世帯 = New classPage農家世帯
                    SysAD.MainForm.MainTabCtrl.TabPages.Add(SysAD.page農家世帯)
                    SysAD.page農家世帯.Active()
                ElseIf Not SysAD.MainForm.MainPage.TabCtrl.TabPages.ContainsKey(SysAD.page農家世帯.Name) Then
                    SysAD.MainForm.MainTabCtrl.TabPages.Add(SysAD.page農家世帯)
                    SysAD.page農家世帯.Active()
                End If

                Return Me.OpenDataViewNext(SysAD.page農家世帯.DataViewCollection, ObjectMan)
            Case "名称" : Return GetItem("名称").ToString
            Case "申請人Aを呼ぶ" : CType(ObjectMan.GetObject("個人." & GetItem("申請者A")), CTargetObjWithView).OpenDataViewNext(SysAD.page農家世帯.DataViewCollection) : Return ""
            Case "申請人Bを呼ぶ" : CType(ObjectMan.GetObject("個人." & GetItem("申請者B")), CTargetObjWithView).OpenDataViewNext(SysAD.page農家世帯.DataViewCollection) : Return ""
            Case "申請人Cを呼ぶ" : CType(ObjectMan.GetObject("個人." & GetItem("申請者C")), CTargetObjWithView).OpenDataViewNext(SysAD.page農家世帯.DataViewCollection) : Return ""
            Case "申請人Aをクリア" : SubClear("A")
            Case "申請人Bをクリア" : SubClear("B")
            Case "申請人Cをクリア" : SubClear("C")
            Case "取下げ" : 農地法取下げ(Me)
            Case "取消し" : 農地法取消し(Me)
            Case "許可に戻す"
                ValueChange("状態", enum申請状況.許可_承認)
                ValueChange("取下年月日", DBNull.Value)
                SaveMyself()
            Case "関連ファイルのリンク"
                ' 関連ファイルのファイリング

            Case "受理申請書発行"
                MsgBox("受理申請書を登録してください")
                ' 受理申請書発行
            Case "許可する" : sub許可(許可区分.許可)
            Case "承認する" : sub許可(許可区分.承認)
            Case "処理する" : sub許可(許可区分.処理)
            Case "工事進捗状況報告書"
                Dim pDlg As New CPrint工事進捗状況報告書(Me)
            Case "許可番号の設定"
                Dim sNo As String = InputBox("許可番号を入力してください", "許可番号の設定")
                If Val(sNo) > 0 Then
                    SetItem("許可番号", Val(sNo))
                    SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D_申請] SET [許可番号]={0} WHERE [ID]=" & Me.ID, sNo)
                End If
            Case "受付に戻す"
                If MsgBox(String.Format("申請[{0}]を受付に戻しますか", Me.名称), vbYesNo) = vbYes Then
                    ValueChange("状態", 0)
                    Me.SaveMyself()
                    Try
                        App農地基本台帳.TBL土地履歴.履歴消去("[申請ID]=" & Me.ID)
                        SysAD.DB(sLRDB).ExecuteSQL("DELETE FROM [D_土地履歴] WHERE [申請ID]=" & Me.ID)
                    Catch ex As Exception

                    End Try
                End If
            Case "審査に戻す"

            Case "申請農地を地図に表示"
                '    St = Replace(Replace(Replace(DVProperty.Rs.Value("農地リスト"), "転用", ""), "農地.", ""), ";", ",")
                '    SysAD.SendMessTo地図(St)
                'Case "４条：転用済み農地の一覧"
                '    If Not DVProperty.Rs.IsNull("申請者A") Then view転用農地List("WHERE ([所有者ID]=" & DVProperty.Rs.Value("申請者A") & ")", DVProperty.Name & ":転用農地一覧")
                'Case "４条：申請者を呼ぶ" : mvarPDW.DataviewCol.Add(ADApp.ObjectMan.GetObject("個人." & DVProperty.Rs.Value("申請者A")))
                'Case "部分設定"
                '    Dim sName As String, ID As Long, nArea As Double

                '    St = DVProperty.Controls.SelectedList("土地一覧Lst")
                '    If Len(St) Then
                '        nArea = Val(DVProperty.Controls.GetSubItem("土地一覧Lst", "部分面積"))
                '        ID = FncNet.GetKeyCode(St)
                '        If nArea Then
                '            SysAD.DB(sLRDB).ExecuteSQL("DELETE FROM [D_部分申請] WHERE [申請ID]=" & DVProperty.ID & " AND [農地ID]=" & ID & " AND [部分面積]=" & nArea)
                '        Else
                '            sName = SysAD.DB(sLRDB).GetDirectData("V_農地", "土地所在", ID)

                '            St = Fnc.InputText("部分面積", "[" & sName & "]に設定する部分面積を入力してください", 0, 1, 3)
                '            If Val(St) > 0 Then
                '                SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO D_部分申請 ( 申請ID, 農地ID, 部分面積 ) VALUES(" & DVProperty.ID & "," & ID & "," & Val(St) & ");")
                '            End If
                '        End If
                '        DVProperty.Controls.Refresh("土地一覧Lst")
                '    End If
                'Case "土地一覧Lst_OLEDragDropKeyList"
                '    DVProperty.Rs.Update("農地リスト", IIf(IsNull(DVProperty.Rs.Value("農地リスト")), sParam, IIf(Len(DVProperty.Rs.Value("農地リスト")) = 0, sParam, DVProperty.Rs.Value("農地リスト") & ";" & sParam)))
                '    DVProperty.Controls.Value("土地一覧Lst") = DVProperty.Rs.Value("農地リスト")
                '    DVProperty.Controls.Refresh("土地一覧Lst")
                'Case "土地一覧Lst_ListRefresh" : DVProperty.Controls.Refresh("土地一覧Lst")
                'Case "対象者一覧Lst_OLEDragDropKeyList"
                '    If Fnc.GetKeyHead(sParam) = "個人" Then
                '        sParam = Replace(sParam, "個人", "対象者")
                '        DVProperty.Controls.Value("対象者一覧Lst") = DVProperty.Controls.Value("対象者一覧Lst") & ";" & sParam
                '        If Left$(DVProperty.Controls.Value("対象者一覧Lst"), 1) = ";" Then DVProperty.Controls.Value("対象者一覧Lst") = Mid$(DVProperty.Controls.Value("対象者一覧Lst"), 2)
                '        DVProperty.Controls.Refresh("対象者一覧Lst")
                '    End If
                'Case "対象者一覧Lst_ListRefresh" : DVProperty.Controls.Refresh("対象者一覧Lst")
            Case "許可書再発行", "許可書の再発行"
                Dim pDate As DateTime = GetDateValue("許可年月日")
                Dim sFolder As String = SysAD.OutputFolder & String.Format("\許可書{0}_{1}", pDate.Year, pDate.Month)

                If HimTools2012.FileManager.CheckAndCleateDirectory(sFolder) Then
                    Select Case Me.法令
                        Case enum法令.農地法3条所有権
                            Dim sDate As String = InputBox("発効日を入力してください", "発効日", Now.Date.ToString.Replace(" 0:00:00", ""))
                            sub異動所有権移転(pDate, sDate, Me.Row, sFolder, False)
                        Case enum法令.農地法3条耕作権
                            Dim sDate As String = InputBox("発効日を入力してください", "発効日", Now.Date.ToString.Replace(" 0:00:00", ""))
                            fnc設置利用権(Me.Key.KeyValue, Me.法令, sFolder, False, pDate, sDate)
                        Case enum法令.農地法4条, enum法令.農地法4条一時転用, enum法令.農地法5条所有権, enum法令.農地法5条貸借, enum法令.農地法5条一時転用
                            sub農地転用(Me.Row, 0, sFolder, True, pDate)
                        Case enum法令.非農地証明願, 600
                            sub非農地(Me.Row.Body, 0, sFolder, True, pDate)
                        Case enum法令.事業計画変更
                            Dim sDate As String = InputBox("発効日を入力してください", "発効日", Now.Date.ToString.Replace(" 0:00:00", ""))
                            sub農地転用事業計画(pDate, sDate, Me.Row, sFolder, False)
                    End Select
                Else
                    MsgBox("デスクトップへ出力先フォルダが作成できませんでした。", MsgBoxStyle.Critical)
                End If

            Case "更新"
                Me.SaveMyself()
            Case "決定通知書"
                Dim pDate As DateTime = GetDateValue("許可年月日")
                Dim sFolder As String = SysAD.OutputFolder & String.Format("\許可書{0}_{1}", pDate.Year, pDate.Month)

                If HimTools2012.FileManager.CheckAndCleateDirectory(sFolder) Then
                    Select Case Me.法令
                        Case enum法令.利用権設定
                            Dim sDate As String = InputBox("発効日を入力してください", "発効日", Now.Date.ToString.Replace(" 0:00:00", ""))
                            fnc設置利用権(Me.Key.KeyValue, 10000 + GetIntegerValue("法令"), sFolder, True, GetDateValue("許可年月日"), sDate)
                    End Select
                Else
                    MsgBox("デスクトップへ出力先フォルダが作成できませんでした。", MsgBoxStyle.Critical)
                End If
            Case "削除"
                If MsgBox("申請情報を本当に削除しますよろしいですか", vbYesNo) = vbYes Then
                    Me.DoCommand("閉じる")

                    SysAD.DB(sLRDB).ExecuteSQL("Delete * FROM [D_申請] WHERE [ID]=" & Me.ID)
                    Dim pRow As DataRow = App農地基本台帳.TBL申請.Rows.Find(Me.ID)
                    App農地基本台帳.TBL申請.Rows.Remove(pRow)
                End If
            Case "受付・交付簿"
                '    pDic = Fnc.InputMulti("受付交付簿", "対象となる総会番号と受付範囲を入力してください", "総会番号;数値;" & DVProperty.Rs.Value("総会番号") & ";|受付範囲;日付;" & ADApp.Common.AppData("受付範囲") & ";")
                '    If pDic Is Nothing Then
                '    ElseIf pDic.Item("総会番号") = 0 Then
                '        MsgBox("総会番号が入力されてません。")
                '    Else
                '        n = pDic.Item("総会番号")
                '        ADApp.Common.AppData("受付範囲") = pDic.Item("受付範囲")
                '        Select Case DVProperty.Rs.Value("法令")
                '            Case enum法令.農地法3条所有権, state3条賃借権, state県3条所有権, state県3条賃借権
                '                mvarPDW.PrintGo(ObjectMan.GetObject("受付交付簿３条"), n & ";" & ADApp.Common.AppData("受付範囲"))
                '            Case state4条転用 : mvarPDW.PrintGo(ObjectMan.GetObject("受付交付簿４条"), n & ";" & ADApp.Common.AppData("受付範囲"))
                '            Case enum法令.農地法5条所有権, enum法令.農地法5条貸借 : mvarPDW.PrintGo(ObjectMan.GetObject("受付交付簿５条"), n & ";" & ADApp.Common.AppData("受付範囲"))
                '            Case state非農地証明 : mvarPDW.PrintGo(ObjectMan.GetObject("受付交付簿非農地"), n & ";" & ADApp.Common.AppData("受付範囲"))
                '            Case state農地利用変更 : mvarPDW.PrintGo(ObjectMan.GetObject("受付交付簿農地変更"), n & ";" & ADApp.Common.AppData("受付範囲"))
                '            Case state事業計画変更 : mvarPDW.PrintGo(ObjectMan.GetObject("受付交付簿事業計画"), n & ";" & ADApp.Common.AppData("受付範囲"))
                '            Case state買受適格耕公, state買受適格耕競 : mvarPDW.PrintGo(ObjectMan.GetObject("受付交付簿買受適格"), n & ";1" & ";" & ADApp.Common.AppData("受付範囲"))
                '            Case state買受適格転公, state買受適格転競 : mvarPDW.PrintGo(ObjectMan.GetObject("受付交付簿買受適格"), n & ";2" & ";" & ADApp.Common.AppData("受付範囲"))
                '        End Select
                '    End If
            Case "受理証明書発行"

            Case "期間設定"
                Dim pDT As Object = mvarRow.Item("始期")
                If Not IsDBNull(pDT) AndAlso pDT IsNot Nothing AndAlso IsDate(pDT) Then
                    With New dlgInputBWDate(pDT)
                        If .ShowDialog() = DialogResult.OK Then
                            ValueChange("終期", .RestultDate)
                            ValueChange("期間", .Restult年数)

                            Me.DataViewPage.UpdateRefresh()
                        End If
                    End With
                Else
                    MsgBox("貸借開始年月日を入力してください", vbCritical)
                End If
            Case "経由法人期間設定"
                Dim pDT As Object = mvarRow.Item("機構配分計画利用配分計画始期日")
                If pDT IsNot Nothing AndAlso IsDate(pDT) Then
                    With New dlgInputBWDate(pDT)
                        If .ShowDialog() = DialogResult.OK Then
                            ValueChange("機構配分計画利用配分計画終期日", .RestultDate)

                            If IsDBNull(mvarRow.Item("始期")) AndAlso IsDBNull(mvarRow.Item("終期")) Then
                                ValueChange("始期", mvarRow.Item("機構配分計画利用配分計画始期日"))
                                ValueChange("終期", mvarRow.Item("機構配分計画利用配分計画終期日"))
                                ValueChange("期間", .Restult年数)
                            End If

                            Me.DataViewPage.UpdateRefresh()
                        End If
                    End With
                Else
                    MsgBox("貸借開始年月日を入力してください", vbCritical)
                End If
                'Case "申請書一括印刷" : mvarPDW.PrintGo(New CPrint奨励金交付申請書, "総会番号=" & SysAD.DB(sLRDB).DBProperty("今月総会番号"))
                'Case "受付設定"
                '    Set pDic = SysAD.MDIForm.InputMulti("受付設定", "以下の項目を入力してください", "受付番号;数値;" & mod農家台帳.Get受付番号Ex("990,991", DVProperty.Rs.Value("総会番号")) & ";|受付年月日;日付;" & Date & ";")
                '    If Not pDic Is Nothing Then
                '        If IsDate(pDic.Item("受付年月日")) Then
                '            DVProperty.Rs.Update("受付年月日", CStr(CDate(pDic.Item("受付年月日"))))
                '            DVProperty.Rs.Update("受付番号", Val(pDic.Item("受付番号")))
                '            mvarPDW.SQLListview.Refresh()
                '        End If
                '    End If
                'Case "交付設定"
                '    pDic = SysAD.MDIForm.InputMulti("交付設定", "以下の項目を入力してください", "交付決定番号;数値;" & DVProperty.Rs.Value("許可番号") & ";|交付決定年月日;日付;" & DVProperty.Rs.Value("許可年月日") & ";")
                '    If Not pDic Is Nothing Then
                '        If IsDate(pDic.Item("交付決定年月日")) Then
                '            DVProperty.Rs.Update("許可年月日", CStr(CDate(pDic.Item("交付決定年月日"))))
                '            DVProperty.Rs.Update("許可番号", Val(pDic.Item("交付決定番号")))
                '            mvarPDW.SQLListview.Refresh()
                '        End If
                '    End If
            Case "履歴だけ作成して許可状態"
                ValueChange("状態", enum申請状況.許可_承認)
                Me.SaveMyself()
                Make農地履歴(Me.ID, Now, Now, 土地異動事由.その他, mvarRow.Item("法令"), "職権による申請の許可・承認")
                'sub履歴のみ()
            Case "事業計画変更申請"
                mod申請データ作成処理.事業計画変更(Me, True)
                'Return New C申請データ作成("転用農地法4条の受付", Me.Key.KeyValue, Nothing)
            Case "変更前複写"
            Case "審査に戻す", "審査にする"
                ValueChange("状態", enum申請状況.審査)
                Me.SaveMyself()
            Case "４条：申請者を呼ぶ"
            Case "申請情報をCSVファイルで出力"
                OutPutCSV申請情報()
            Case "不許可"
                If MsgBox(String.Format("申請[{0}]を不許可にしますか？", Me.名称), vbYesNo) = vbYes Then
                    ValueChange("状態", enum申請状況.不許可)
                    Me.SaveMyself()
                    Make農地履歴(Me.ID, Now, Now, 土地異動事由.その他, mvarRow.Item("法令"), "申請を不許可として処理しました。")
                End If
            Case "通知書の発行" : sub通知書発行(許可区分.承認)
            Case "全選択"
                If MsgBox("全選択を行ってもよろしいですか？", vbOKCancel) = Microsoft.VisualBasic.MsgBoxResult.Ok Then
                    Dim list As New List(Of String)
                    Dim pGrid = CType(Me.DataViewPage, DataViewNext申請).mvarGrid申請地一覧
                    For Each pRow As DataRow In pGrid.DataSource.Rows
                        list.Add("農地." & pRow.Item("ID"))
                        pRow.Item("選択") = True
                    Next
                    Dim St As String = ""
                    For Each s As String In list
                        If St = "" Then
                            St = s
                        Else
                            St = St & ";" & s
                        End If
                    Next
                    ValueChange("農地リスト", St)
                    Me.SaveMyself()
                End If

            Case Else
                Return MyBase.DoCommand(sCommand, sParams)
        End Select
        Return ""
    End Function

    Private Sub OutPutCSV申請情報()
        Dim pTBL As DataTable = New DataView(App農地基本台帳.TBL申請.Body, "[ID] = " & Me.ID, "", DataViewRowState.CurrentRows).ToTable
        Dim sCSV As New System.Text.StringBuilder
        Dim pLineRow As New System.Text.StringBuilder

        Dim LastColumn As Integer = pTBL.Columns.Count - 1
        For i = 0 To LastColumn
            Dim field As String = pTBL.Columns(i).Caption 'ヘッダの取得

            field = """" & field & """" '"で囲む
            pLineRow.Append(field) 'フィールドを書き込む

            If LastColumn > i Then
                pLineRow.Append(","c) 'カンマを書き込む
            End If
        Next
        sCSV.AppendLine(pLineRow.ToString)
        pLineRow.Clear()

        For Each pRow As DataRow In pTBL.Rows
            For i = 0 To LastColumn
                Dim field As String = pRow(i).ToString() 'フィールドの取得
                Debug.Print(field)
                field = """" & field & """" '"で囲む
                pLineRow.Append(field) 'フィールドを書き込む

                If LastColumn > i Then
                    pLineRow.Append(","c) 'カンマを書き込む
                End If
            Next
        Next
        sCSV.AppendLine(pLineRow.ToString)
        名前を付けて保存(sCSV, String.Format("申請エラー({0})", Format(Now, "yyyyMMdd")), True, True)
    End Sub

    Private sPath As String = ""
    Private Sub 名前を付けて保存(ByVal sCSV As System.Text.StringBuilder, ByVal SaveFileName As String, Optional ByVal OpenDialog As Boolean = False, Optional ByVal OpenFolder As Boolean = False)
        '/***名前を付けて保存***/
        If OpenDialog = True Then
            With New SaveFileDialog
                .FileName = String.Format("{0}.csv", SaveFileName)
                .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
                .Filter = "CSVファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*"

                If .ShowDialog = DialogResult.OK Then
                    sPath = .FileName
                End If
            End With
        End If

        Dim ArSavePath As Object = Split(sPath, "\")
        For n As Integer = 0 To UBound(ArSavePath)
            If n = 0 Then : sPath = ArSavePath(0)
            ElseIf n = UBound(ArSavePath) Then : sPath = sPath & "\" & String.Format("{0}.csv", SaveFileName)
            Else : sPath = sPath & "\" & ArSavePath(n)
            End If
        Next

        Dim CSVText As New System.IO.StreamWriter(sPath, False, System.Text.Encoding.GetEncoding(932))
        CSVText.Write(sCSV.ToString)
        CSVText.Dispose()

        If OpenFolder = True Then
            MsgBox("CSVの出力が完了しました。")
            SysAD.ShowFolder(System.IO.Directory.GetParent(sPath).ToString)
        End If
    End Sub

    Private Sub SubClear(ByVal pKey As String)
        If MsgBox("申請人情報をクリアしてもよろしいですか？", vbOKCancel) = vbOK Then
            ValueChange("申請世帯" & pKey, 0)
            ValueChange("申請者" & pKey, 0)
            ValueChange("氏名" & pKey, "")
            ValueChange("職業" & pKey, "")
            ValueChange("住所" & pKey, "")
            ValueChange("集落" & pKey, "")
            ValueChange("年齢" & pKey, 0)
            ValueChange("経営面積" & pKey, 0)
            'ValueChange("名称", 0)
            SaveMyself()
        End If
    End Sub

#Region "議案書・許可書置き直し"
#Region "Replace申請者A"
    Public Sub Replace申請者A(ByRef sXML As String)
        sXML = Replace(sXML, "{郵便番号A}", "〒" & Me.GetItem("申請者情報郵便番号A").ToString)
        sXML = Replace(sXML, "{郵便番号}", "〒" & Me.GetItem("申請者情報郵便番号A").ToString)
        sXML = Replace(sXML, "{住所A}", Me.GetItem("住所A").ToString)
        sXML = Replace(sXML, "{氏名A}", Me.GetItem("氏名A").ToString)
        sXML = Replace(sXML, "{申請者Ａ住所}", Me.GetItem("住所A").ToString)
        sXML = Replace(sXML, "{申請者Ａ氏名}", Me.GetItem("氏名A").ToString)
    End Sub
    Public Sub Replace申請者A(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet)
        pSheet.ValueReplace("{郵便番号A}", GetStringValue("申請者情報郵便番号A"))
        pSheet.ValueReplace("{郵便番号}", GetStringValue("申請者情報郵便番号A"))
        pSheet.ValueReplace("{申請者Ａ氏名}", GetStringValue("氏名A"))
        pSheet.ValueReplace("{申請者Ａ住所}", GetStringValue("住所A"))
        pSheet.ValueReplace("{氏名A}", GetStringValue("氏名A"))
        pSheet.ValueReplace("{住所A}", GetStringValue("住所A"))
    End Sub
#End Region

#Region "Replace申請者B"
    Public Sub Replace申請者B(ByRef sXML As String)
        sXML = Replace(sXML, "{郵便番号B}", "〒" & Me.GetItem("申請者情報郵便番号B").ToString)
        sXML = Replace(sXML, "{住所B}", Me.GetItem("住所B").ToString)
        sXML = Replace(sXML, "{氏名B}", Me.GetItem("氏名B").ToString)
    End Sub
    Public Sub Replace申請者B(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet)
        pSheet.ValueReplace("{郵便番号B}", GetStringValue("申請者情報郵便番号B"))
        pSheet.ValueReplace("{氏名B}", GetStringValue("氏名B"))
        pSheet.ValueReplace("{住所B}", GetStringValue("住所B"))
    End Sub

#End Region
#End Region

    Public Overrides Sub DropDown(ByVal sSourceList As String, Optional ByVal sOption As String = "")
        Select Case sOption
            Case "代理人A", "代理人名"
                If MsgBox("代理人を変更しますか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    Dim pPerson As HimTools2012.TargetSystem.CTargetObjectBase = ObjectMan.GetObject(sSourceList)
                    Me.Row.NullCheck("代理人A")
                    Select Case pPerson.Key.DataClass
                        Case "農家"
                            ValueChange("代理人A", pPerson.GetProperty("世帯主ID"))
                            ValueChange("代理人名", pPerson.GetProperty("世帯主氏名"))
                            'ValueChange("代理人住所", pPerson.GetProperty("世帯主住所"))
                        Case "個人"
                            ValueChange("代理人A", pPerson.ID)
                            ValueChange("代理人名", pPerson.GetProperty("氏名"))
                            ValueChange("代理人住所", pPerson.GetProperty("住所"))
                    End Select
                End If
            Case "個人リスト"
                If GetKeyHead(sSourceList) = "個人" Then
                    sSourceList = Replace(sSourceList, "個人", "対象者")
                    Dim sList As String = GetStringValue("個人リスト") & ";" & sSourceList
                    If sList.StartsWith(";") Then Strings.Mid(sList, 2)

                    ValueChange("代理人A", sList)


                    ValueChange("個人リスト", sList & ";" & sSourceList)
                    ValueChange("個人リスト", sList)
                End If
            Case "経由法人ID", "経由法人名"
                If MsgBox("経由法人を変更しますか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    Me.Row.NullCheck("経由法人ID")
                    Select Case GetKeyHead(sSourceList)
                        Case "個人" : ValueChange("経由法人ID", GetKeyCode(sSourceList))
                    End Select
                End If
            Case "農地リスト"
                Select Case GetKeyHead(sSourceList)
                    Case "農地"
                        ValueChange("農地リスト", GetStringValue("農地リスト") & ";" & sSourceList)
                        With CType(Me.DataViewPage, DataViewNext申請)

                            Dim sWhere As String = .Get農地条件()
                            App農地基本台帳.TBL農地.FindRowBySQL(sWhere)
                            Dim mvar農地View As New DataView(App農地基本台帳.TBL農地.Body, sWhere, "", DataViewRowState.CurrentRows)
                            Dim pTable As DataTable = mvar農地View.ToTable("農地", False, New String() {"Key", "ID", "大字", "小字", "地番", "登記簿面積", "登記簿地目名", "現況地目名", "自小作", "借受人氏名"})

                            With CType(.Controls.Find("農地リスト", True)(0), GridViewNext)
                                CType(.DataSource, DataTable).Merge(pTable, False, MissingSchemaAction.AddWithKey)
                                .Value = GetStringValue("農地リスト")
                            End With

                        End With

                End Select
            Case "申請者A", "氏名A"
                Select Case GetKeyHead(sSourceList)
                    Case "個人"
                        Dim p渡手 As CObj個人 = ObjectMan.GetObject(sSourceList)
                        If MsgBox("申請人を変更しますか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                            Me.ValueChange("申請世帯A", p渡手.世帯ID)
                            Me.ValueChange("申請者A", p渡手.ID)
                            Me.ValueChange("氏名A", p渡手.氏名)
                            Me.ValueChange("住所A", p渡手.住所)
                            Me.ValueChange("職業A", "")
                            Me.ValueChange("年齢A", 0)
                            Me.ValueChange("集落A", "")
                            Me.DataViewPage.UpdateRefresh()
                            Dim s名称 As String = GetStringValue("名称")
                            If InStr(s名称, "→") > 0 Then
                                s名称 = p渡手.氏名 & HimTools2012.StringF.Mid(s名称, InStr(s名称, "("))
                                ValueChange("名称", s名称)
                            ElseIf InStr(s名称, "(") > 0 Then
                                s名称 = p渡手.氏名 & HimTools2012.StringF.Mid(s名称, InStr(s名称, "("))
                                ValueChange("名称", s名称)
                            End If
                        End If
                End Select
            Case "申請者B", "氏名B"
                Select Case GetKeyHead(sSourceList)
                    Case "個人"
                        Dim p受手 As CObj個人 = ObjectMan.GetObject(sSourceList)
                        If MsgBox("譲受人を変更しますか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                            Me.ValueChange("申請世帯B", p受手.世帯ID)
                            Me.ValueChange("申請者B", p受手.ID)
                            Me.ValueChange("氏名B", p受手.氏名)
                            Me.ValueChange("住所B", p受手.住所)
                            Me.ValueChange("職業B", "")
                            Me.ValueChange("年齢B", 0)
                            Me.ValueChange("集落B", "")
                            Me.DataViewPage.UpdateRefresh()
                            Dim s名称 As String = GetStringValue("名称")
                            If InStr(s名称, "→") > 0 Then
                                s名称 = HimTools2012.StringF.Left(s名称, InStr(s名称, "→")) & p受手.氏名
                                ValueChange("名称", s名称)
                            End If
                        End If
                End Select
            Case "申請者C", "氏名C"
                Select Case GetKeyHead(sSourceList)
                    Case "個人"
                        Dim p渡手 As CObj個人 = ObjectMan.GetObject(sSourceList)
                        If MsgBox("譲渡人を変更しますか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                            Me.ValueChange("申請世帯C", p渡手.世帯ID)
                            Me.ValueChange("申請者C", p渡手.ID)
                            Me.ValueChange("氏名C", p渡手.氏名)
                            Me.ValueChange("住所C", p渡手.住所)
                            Me.ValueChange("職業C", "")
                            Me.ValueChange("年齢C", 0)
                            Me.ValueChange("集落C", "")
                            Me.DataViewPage.UpdateRefresh()
                            Dim s名称 As String = GetStringValue("名称")
                            If InStr(s名称, "→") > 0 Then
                                s名称 = HimTools2012.StringF.Left(s名称, InStr(s名称, "→")) & p渡手.氏名
                                ValueChange("名称", s名称)
                            End If
                        End If
                End Select
            Case "行政書士"
                If MsgBox("行政書士を変更しますか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    Dim pPerson As HimTools2012.TargetSystem.CTargetObjectBase = ObjectMan.GetObject(sSourceList)
                    Me.Row.NullCheck("行政書士")
                    Select Case pPerson.Key.DataClass
                        Case "個人"
                            ValueChange("行政書士", pPerson.GetProperty("氏名"))
                    End Select
                End If
            Case "農業委員1", "農業委員2", "農業委員3"
                If MsgBox("農業委員を変更しますか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    Dim pPerson As HimTools2012.TargetSystem.CTargetObjectBase = ObjectMan.GetObject(sSourceList)
                    Me.Row.NullCheck(sOption)
                    Select Case pPerson.Key.DataClass
                        Case "個人"
                            ValueChange(sOption, pPerson.ID)
                    End Select
                End If
            Case ""
                '    pPerson = ADApp.ObjectMan.GetObject(sSourceList)
                '    Select Case Fnc.GetKeyHead(pPerson.Key)
                '        Case "申請"
                '            If DVProperty.Rs.Value("法令") = state事業計画変更 Then
                '                If pPerson.GetProperty2("法令") = state4条転用 Or pPerson.GetProperty2("法令") = enum法令.農地法5条所有権 Or pPerson.GetProperty2("法令") = enum法令.農地法5条貸借 Then
                '                    Dim Rs As NK97.RecordsetEx
                '                    Dim sSQL As String
                '                    Dim n As Long
                '                    Dim i As Long
                '                    Dim St As String
                '                    Dim sTable As String

                '                    Select Case pPerson.GetProperty2("法令")
                '                        Case state4条転用
                '                            .Value("変更前世帯ID") = pPerson.GetProperty2("申請世帯A")
                '                            .Value("変更前申請人ID") = pPerson.GetProperty2("申請者A")
                '                            .Value("変更前申請人名") = pPerson.GetProperty2("氏名A")
                '                            .Value("変更前住所") = pPerson.GetProperty2("住所A")
                '                        Case enum法令.農地法5条所有権, enum法令.農地法5条貸借
                '                            .Value("変更前世帯ID") = pPerson.GetProperty2("申請世帯B")
                '                            .Value("変更前申請人ID") = pPerson.GetProperty2("申請者B")
                '                            .Value("変更前申請人名") = pPerson.GetProperty2("氏名B")
                '                            .Value("変更前住所") = pPerson.GetProperty2("住所B")
                '                    End Select
                '                    .Value("変更前転用目的TX") = pPerson.GetProperty2("申請理由A")
                '                    If pPerson.GetProperty2("状態") = 2 Then
                '                        .Value("変更前許可日TX") = Format(pPerson.GetProperty2("許可年月日"), "GEE/MM/DD")
                '                    Else
                '                        .Value("変更前許可日TX") = ""
                '                    End If
                '                    If IsNull(pPerson.GetProperty2("農地リスト")) Then
                '                        .Value("変更前土地一覧Grid") = ""
                '                    Else
                '                        If pPerson.GetProperty2("状態") = 2 Then
                '                            St = Replace(Replace(pPerson.GetProperty2("農地リスト"), "転用農地.", ""), ";", ",")
                '                            sTable = "V_転用農地"
                '                        Else
                '                            St = Replace(Replace(pPerson.GetProperty2("農地リスト"), "農地.", ""), ";", ",")
                '                            sTable = "V_農地"
                '                        End If

                '                        sSQL = "SELECT [" & sTable & "].[土地所在],[" & sTable & "].[実面積],[" & sTable & "].[田面積],[" & sTable & "].[畑面積],[V_地目].名称 AS 登記地目,[地目2].名称 AS [現況],[V_農委地目].名称 AS [農委地目]" & _
                '                        " FROM ((([" & sTable & "] LEFT JOIN [V_地目] ON [" & sTable & "].[登記簿地目]=[V_地目].ID) LEFT JOIN V_地目 AS 地目2 ON [" & sTable & "].[現況地目]=[地目2].ID) LEFT JOIN [V_農委地目] ON [" & sTable & "].[農委地目ID]=[V_農委地目].ID) " & _
                '                        " WHERE [" & sTable & "].[ID] IN (" & St & ") ORDER BY [" & sTable & "].[大字ID],[" & sTable & "].[小字ID],val(IIf(IsNull([地番]),0,IIf(InStr([地番],'-')=0,Val([地番]),val(Left([地番],InStr([地番],'-')-1)))) *1000 + Val(IIf(IsNull([地番]),0,IIf(InStr([地番],'-')=0,0,val(Mid([地番],InStr([地番],'-')+1))))))"
                '                        Rs = SysAD.DB(sLRDB).GetRecordsetEx(sSQL, 1, , "事業計画変更申請")
                '                        St = "" : i = 0
                '                        Do Until Rs.EOF
                '                            n = Rs.NullCast("田面積", 0) + Rs.NullCast("畑面積", 0)
                '                            If n = 0 Then n = Rs.NullCast("実面積", 0)
                '                            If i > 0 Then St = St & vbCrLf
                '                            St = St & Rs.V("土地所在") & ";" & Rs.V("登記地目") & ";" & Rs.V("現況") & ";" & n
                '                            i = i + 1
                '                            Rs.MoveNext()
                '                        Loop
                '                        SysAD.DB(sLRDB).CloseRs(Rs)
                '                        .Value("変更前土地一覧Grid") = St
                '                    End If
                '                End If
                '            End If
                '    End Select
            Case Else
#If DEBUG Then
                Stop
#End If

        End Select
    End Sub

    Private Enum 許可区分
        許可
        承認
        処理
    End Enum

    Private Sub sub許可(ByVal n許可区分 As 許可区分, Optional ByVal bCont As Boolean = False)
        Dim s区分 As String = n許可区分.ToString

        If Not IsDate(mvarRow.Item("受付年月日")) Then
            MsgBox("受付年月日が設定されていません。", vbCritical)
            Exit Sub
        ElseIf IsDBNull(mvarRow.Item("農地リスト")) Then
            MsgBox("農地が登録されていません。", vbCritical)
            Exit Sub
        End If

        With New HimTools2012.PropertyGridDialog(New C許可入力支援(Me), Me.名称 & " 許可処理")

            If .ShowDialog = DialogResult.OK Then
                Dim sFolder As String = SysAD.OutputFolder & String.Format("\許可書{0}_{1}", Now.Year, Now.Month)

                If Not IO.Directory.Exists(sFolder) Then
                    IO.Directory.CreateDirectory(sFolder)
                End If


                With CType(.ResultProperty, C許可入力支援)
                    Dim sDate = .許可_処理年月日

                    Select Case mvarRow.Item("法令")
                        Case enum法令.農地法4条, enum法令.農地法4条一時転用 : sub農地転用(Me.Row, 10040, sFolder, False, CDate(sDate))
                        Case enum法令.農地法5条所有権 : sub農地転用(Me.Row, 10050, sFolder, False, CDate(sDate))
                        Case enum法令.農地法5条貸借, enum法令.農地法5条一時転用 : sub農地転用(Me.Row, 10051, sFolder, False, CDate(sDate))
                        Case enum法令.農地法3条所有権, enum法令.基盤強化法所有権
                            Dim _result As String = InputBox("発効日を入力してください", "発効日", Now.Date.ToString.Replace(" 0:00:00", ""))
                            sub異動所有権移転(CDate(sDate), sDate, Me.Row, sFolder, False)
                        Case enum法令.農地法3条耕作権
                            Dim _result As String = InputBox("発効日を入力してください", "発効日", Now.Date.ToString.Replace(" 0:00:00", ""))
                            fnc設置利用権(Me.Key.KeyValue, Me.法令, sFolder, False, CDate(sDate), _result)
                        Case enum法令.利用権設定 : 終了メッセージ(fnc設置利用権(Me.Key.KeyValue, 10000 + Me.法令, sFolder, False, CDate(sDate), Nothing))
                        Case enum法令.農地法3条の3第1項 : sub異動所有権移転(CDate(sDate), Now.Date, Me.Row, sFolder, False)
                        Case enum法令.非農地証明願, 600 : sub非農地(Me.Row.Body, 18099, sFolder, False, CDate(sDate))
                        Case enum法令.利用権移転
                            Dim _result As String = InputBox("通知書の発行日を入力してください", "許可処理", Now.Date.ToString.Replace(" 0:00:00", ""))
                            fnc設置利用権(Me.Key.KeyValue, 10000 + Me.法令, sFolder, False, CDate(sDate), _result)
                        Case enum法令.事業計画変更
                            Dim _result As String = InputBox("発効日を入力してください", "発効日", Now.Date.ToString.Replace(" 0:00:00", ""))
                            sub農地転用事業計画(CDate(sDate), _result, Me.Row, sFolder, False)
                        Case enum法令.農地法18条解約 : RentEnd(Me.Key.KeyValue, CDate(sDate), "18条解約")
                        Case enum法令.合意解約 : RentEnd(Me.Key.KeyValue, CDate(sDate), "20条解約")
                        Case enum法令.中間管理機構へ農地の返還 : RentEnd(Me.Key.KeyValue, CDate(sDate), "農地返還")
                        Case enum法令.農地改良届 : sub許可承認("農地改良届")
                        Case enum法令.農地利用目的変更 : sub許可承認("農地利用目的変更")
                        Case enum法令.農用地計画変更 : sub許可承認("農用地利用計画変更")
                        Case enum法令.あっせん出手, enum法令.あっせん受手 : sub許可承認("あっせん申出")
                        Case enum法令.買受適格耕競, enum法令.買受適格耕公, enum法令.買受適格転競, enum法令.買受適格転公 : sub許可承認("買受適格証明")
                        Case Else
                            Stop
                    End Select
                End With

            Else
                Exit Sub
            End If
        End With

    End Sub

    Private Sub sub許可承認(ByVal 履歴内容 As String)
        If MsgBox(String.Format("申請[{0}]を承認しますか", Me.名称), vbYesNo) = vbYes Then
            ValueChange("状態", 2)
            Me.SaveMyself()
            Make農地履歴(Me.ID, Now, Now, 土地異動事由.その他, mvarRow.Item("法令"), 履歴内容 & "の承認")
        End If
    End Sub

    Private Sub sub通知書発行(ByVal n許可区分 As 許可区分, Optional ByVal bCont As Boolean = False)
        Dim s区分 As String = n許可区分.ToString

        If IsDBNull(mvarRow.Item("農地リスト")) Then
            MsgBox("農地が登録されていません。", vbCritical)
            Exit Sub
        End If

        If Not IsDate(mvarRow.Item("受付年月日").ToString) Then
            If MsgBox("受付年月日が設定されていません。発行を続行しますか？", vbYesNo) = vbYes Then
            Else
                Exit Sub
            End If
        End If

        Dim sFolder As String = SysAD.OutputFolder & String.Format("\通知書{0}_{1}", Now.Year, Now.Month)

        If Not IO.Directory.Exists(sFolder) Then
            IO.Directory.CreateDirectory(sFolder)
        End If

        Select Case mvarRow.Item("法令")
            Case enum法令.農地法18条解約, enum法令.合意解約, enum法令.中間管理機構へ農地の返還 : fnc通知書発行(Me.Key.KeyValue, 10000 + Me.法令, sFolder, False, CDate(Now), Nothing)
            Case Else
                Stop
        End Select
    End Sub

    Private Sub 終了メッセージ(ByVal b As Boolean)
        If b Then
            MsgBox("処理が完了しました。")
        Else

        End If
    End Sub
    Private Sub sub履歴のみ(Optional ByVal bCont As Boolean = False)
        'Dim St As String
        'Dim StX As String
        'Dim pObj As Object
        'Dim SD As String
        'Dim pDic As CDataDictionary
        'Dim n As Long
        'Dim Rs As NK97.RecordsetEx
        'Dim s法令 As String

        'If Not IsDate(DVProperty.Rs.Value("受付年月日")) Then
        '    MsgBox("受付年月日が設定されていません。", vbCritical)
        '    Exit Sub
        'ElseIf IsNull(DVProperty.Rs.Value("農地リスト")) Then
        '    MsgBox("農地が登録されていません。", vbCritical)
        '    Exit Sub
        'End If

        'CDataviewSK_DoCommand2("更新")
        'If bCont Then
        '    SD = CStr(DateSerial(Year(DVProperty.Rs.Value("受付年月日")), Month(DVProperty.Rs.Value("受付年月日")), 28))
        'Else
        '    Rs = SysAD.DB(sLRDB).GetRecordsetEx("SELECT * FROM [D_PSTRING] WHERE [開始日]<=#" & DVProperty.Rs.Value("受付年月日") & "# AND [終了日]>=#" & DVProperty.Rs.Value("受付年月日") & "#", 0, , Me)
        '    If Not Rs.EOF Then
        '        St = CStr(Rs.NullCast("総会日", ""))
        '    End If
        '    Rs.CloseRs()
        '    If Len(St) = 0 Then
        '        St = CStr(DateSerial(DVProperty.Rs.Year("受付年月日"), DVProperty.Rs.Month("受付年月日"), 28))
        '    End If

        '    '            Select Case DVProperty.Rs.Value("法令")
        '    '                Case state奨励金交付A, state奨励金交付B:
        '    '                    Set pDic = SysAD.MDIForm.InputMulti("交付決定", "以下の項目を入力してください", "交付決定番号;数値;" & DVProperty.Rs.Value("受付番号") & ";|交付決定年月日;日付;" & St & ";")
        '    '                    If Not pDic Is Nothing Then
        '    '                        If IsDate(pDic.Item("交付決定年月日")) Then
        '    '                            DVProperty.Rs.Update "許可番号", Val(pDic.Item("交付決定番号"))
        '    '                            SD = pDic.Item("交付決定年月日")
        '    '                        End If
        '    '                    End If
        '    '                Case Else
        '    SD = Fnc.InputText("許可年月日", "許可年月日を入力してください", St, 0, 0)
        '    '            End Select
        'End If

        'If IsDate(SD) Then
        '    Select Case DVProperty.Rs.Value("法令")
        '        Case state3条所有権, state県3条所有権 : sub異動from農地法3条(CDate(SD), DVProperty.Rs.Value("農地リスト"), SysAD.DB(sLRDB).GetDirectData("D:個人Info", "世帯ID", DVProperty.Rs.Value("申請者B")), DVProperty.Rs.Value("申請者B"), DVProperty.ID, True) : s法令 = "3条"
        '        Case state3条賃借権, state県3条賃借権 : LandRent(CDate(SD), DVProperty.Rs.Value("法令"), True) : s法令 = "3条"
        '        Case state4条転用 : mod農家台帳.sub農地転用(DVProperty.Key, 10040, True) : s法令 = "4条"
        '        Case enum法令.農地法5条所有権 : mod農家台帳.sub農地転用(DVProperty.Key, 10050, True) : s法令 = "5条"
        '        Case enum法令.農地法5条貸借 : mod農家台帳.sub農地転用(DVProperty.Key, 10050, True) : s法令 = "5条"
        '        Case state基盤所有権設定 : sub異動from基盤強化法(CDate(SD), DVProperty.Rs.Value("農地リスト"), Val(SysAD.DB(sLRDB).GetDirectData("D:個人Info", "世帯ID", DVProperty.Rs.Value("申請者B"))), DVProperty.Rs.Value("申請者B"), DVProperty.ID)
        '        Case state基盤利用権移転, enum法令.利用権設定 : LandRent(CDate(SD), DVProperty.Rs.Value("法令"), True)
        '            '                Case state20条解約: RentEnd CDate(SD), "20条解約"
        '            '                Case state合意解約: RentEnd CDate(SD), "合意解約"
        '            '                Case state農地利用変更, state農用地計画変更: UseChange CDate(SD), DVProperty.Rs.Value("法令")
        '            '                Case state事業計画変更:
        '        Case Else
        '            Debug.Assert(CaseAssertPrint(DVProperty.Rs.Value("法令"), ""))
        '    End Select

        '    DVProperty.Rs.Update("許可年月日", CStr(CDate(SD)))
        '    DVProperty.Rs.Update("許可補助記号", IIf(IsNull(ADApp.User.Propertyv支所符号), Null, ADApp.User.Propertyv支所符号))

        '    DVProperty.Rs.Update("状態", 2)
        '    mvarState = 2
        'End If
        'CDataviewSK_DoCommand2("閉じる")

    End Sub

    Private Sub sub不許可(Optional ByVal bCont As Boolean = False)
        'Dim St As String
        'Dim pObj As Object
        'Dim SD As String
        'Dim n As Long

        'If bCont Then
        '    SD = CStr(DateSerial(Year(DVProperty.Rs.Value("受付年月日")), Month(DVProperty.Rs.Value("受付年月日")), 28))
        'Else
        '    SD = Fnc.InputText("不許可決定の年月日", "不許可を決定した年月日を入力してください", CStr(DateSerial(Year(DVProperty.Rs.Value("受付年月日")), Month(DVProperty.Rs.Value("受付年月日")), 28)), 0, 0)
        'End If

        'If IsDate(SD) Then

        '    DVProperty.Rs.Update("許可年月日", CStr(CDate(SD)))
        '    DVProperty.Rs.Update("許可補助記号", IIf(IsNull(ADApp.User.Propertyv支所符号), Null, ADApp.User.Propertyv支所符号))

        '    If Not (DVProperty.Rs.Value("法令") = 30 Or DVProperty.Rs.Value("法令") = 31) Then
        '    ElseIf bCont Then
        '    ElseIf MsgBox("不許可に関する書類を発行しますか", vbYesNo) = vbYes Then
        '        n = DVProperty.Rs.Value("受付番号")
        '        With Fnc.InputData("不許可番号", "不許可の発行番号を入力してください", Me, GetArray("発行番号;数値;" & n, "発効日;日付;" & Now()))
        '            If .IsObtained Then
        '                If .Exists("発行番号") Then DVProperty.Rs.Update("許可番号", Val(.Item("発行番号")))
        '                SysAD.ApplicationDLL.Form.PrintGo(SysAD.ApplicationDLL.ObjectMan.GetObject("許可書"), "3条." & FncNet.GetKeyCode(DVProperty.Key) & "不許可" & vbFormFeed & .Item("発効日"))
        '            End If
        '        End With
        '    End If

        '    DVProperty.Rs.Update("状態", 42)
        '    mvarState = 42
        'End If
        Me.DoCommand("閉じる")

    End Sub


    Public Overrides Function InitDataViewNext(ByRef pDB As HimTools2012.TargetSystem.CDataViewCollection, Optional InterfaceName As String = "") As Boolean
        If Me.DataViewPage Is Nothing Then
            Select Case CType(Me.GetIntegerValue("法令"), enum法令)
                Case enum法令.農地法3条所有権 : Me.DataViewPage = New CInterface3条所有権設定(Me)
                Case enum法令.農地法3条耕作権 : Me.DataViewPage = New CInterface3条耕作権設定(Me)
                Case enum法令.農地法3条の3第1項 : Me.DataViewPage = New CInterface3条の3第1項農地取得(Me)

                Case enum法令.農地法4条 : Me.DataViewPage = New CInterface農地法4条(Me)
                Case enum法令.農地法4条一時転用 : Me.DataViewPage = New CInterface農地法4条一時転用(Me)
                Case enum法令.農地法5条所有権 : Me.DataViewPage = New CInterface農地法5条所有権(Me)
                Case enum法令.農地法5条貸借 : Me.DataViewPage = New CInterface農地法5条賃借権(Me)
                Case enum法令.農地法5条一時転用 : Me.DataViewPage = New CInterface農地法5条一時転用(Me)

                Case enum法令.非農地証明願, 600 : Me.DataViewPage = New CInterface非農地証明願(Me)
                    'MyBase.InitDataViewNext(pDB, "申請非農地証明願")    Return True
                Case enum法令.農地法18条解約, enum法令.農地法20条解約, enum法令.中間管理機構へ農地の返還 : Me.DataViewPage = New CInterface18条解約(Me)
                Case enum法令.合意解約, enum法令.中間管理機構へ農地の返還 : Me.DataViewPage = New CInterface合意解約(Me)

                Case enum法令.基盤強化法所有権 : Me.DataViewPage = New CInterface基盤法所有権設定(Me)
                Case enum法令.利用権設定 : Me.DataViewPage = New CInterface利用権設定(Me)
                Case enum法令.利用権移転 : Me.DataViewPage = New CInterface利用権移転(Me)

                Case enum法令.農用地計画変更 : Me.DataViewPage = New CInterface農振地整備計画変更(Me)
                Case enum法令.農地利用目的変更 : Me.DataViewPage = New CInterface農地利用目的変更(Me)
                Case enum法令.農地改良届 : Me.DataViewPage = New CInterface農地改良届(Me)
                Case enum法令.事業計画変更 : Me.DataViewPage = New CInterface事業計画変更(Me)

                Case enum法令.買受適格耕公 : Me.DataViewPage = New CInterface買受適格耕公(Me)
                Case enum法令.買受適格耕競 : Me.DataViewPage = New CInterface買受適格耕競(Me)
                Case enum法令.買受適格転公 : Me.DataViewPage = New CInterface買受適格転公(Me)
                Case enum法令.買受適格転競 : Me.DataViewPage = New CInterface買受適格転競(Me)
                Case enum法令.あっせん出手 : Me.DataViewPage = New CInterfaceあっせん申出渡(Me)
                Case enum法令.あっせん受手 : Me.DataViewPage = New CInterfaceあっせん申出受(Me)
            End Select
        End If
        With CType(Me.DataViewPage, DataViewNext申請)
            .SetInterface(.Panel)
        End With
        Return True
    End Function
    Public Overloads Overrides Function CanDropKeyHead(ByVal sKey As String, ByVal sOption As String) As Boolean
        Select Case sKey
            Case "個人"
                Select Case sOption
                    Case "経由法人ID", "経由法人名" : Return True
                    Case "申請者A", "氏名A" : Return True
                    Case "申請者B", "氏名B" : Return True
                    Case "申請者C", "氏名C" : Return True
                    Case "代理人A", "代理人名" : Return True
                    Case "行政書士" : Return True
                    Case "農業委員1", "農業委員2", "農業委員3" : Return True
                    Case "" : Return False
                    Case Else
                        If Not SysAD.IsClickOnceDeployed Then
                            CasePrint(sOption)
                            Stop
                        End If
                        Return False

                End Select
            Case "" : Return False
            Case "農地"
                Return False
            Case Else
                If Not SysAD.IsClickOnceDeployed Then
                    Debug.Print(sKey)
                    Stop
                End If
                Return False
        End Select
    End Function

    Public Overrides Sub ValueChange(ByVal sField As String, ByVal pValue As Object)
    

        App農地基本台帳.TBL申請.AddUpdateListwithDataViewPage(Me, sField, pValue)
    End Sub
End Class
