

Public Class C大崎町
    Inherits C市町村別

    ''' <summary>
    ''' \\ibmserver\農政Server\Avail\システム配置\農委大崎町
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        MyBase.New("大崎町")
    End Sub
    Public Overrides Function Get地区情報(ByVal s住所 As String) As String
        Return "大崎"
    End Function
    Public Overrides Function Get選挙世帯一覧() As System.Data.DataTable
        Return Nothing
    End Function

    Public Overrides Sub InitLocalData()
        With New dlgLoginForm()

            If Not .ShowDialog() = Windows.Forms.DialogResult.OK Then
                Try
                    End
                Catch ex As Exception

                End Try
            End If
        End With

        Dim sPath As String = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\農地台帳LocaloData"
        If Not IO.Directory.Exists(sPath) Then
            MsgBox("=>" & sPath & "を作成します。")
            Try
                IO.Directory.CreateDirectory(sPath)
            Catch ex As Exception
                MsgBox("失敗しました:" & ex.Message)
            End Try
        End If
        SysAD.SystemInfo.LocalDataPath = sPath
        SysAD.SystemInfo.XMLDataPath = sPath

        sub農地期間満了の終了()

    End Sub
    Public Overrides ReadOnly Property 旧農振都市計画使用 As Boolean
        Get
            Return True
        End Get
    End Property
    Public Overrides Sub InitMenu(pMain As HimTools2012.SystemWindows.CMainPageSK)
        With pMain
            .ListView.Clear()
            .ListView.ItemAdd("農家検索", "農家検索", "閲覧・検索", "閲覧・検索", AddressOf 農家一覧)
            .ListView.ItemAdd("農地台帳一括印刷", "農地台帳一括印刷", "印刷", "印刷", AddressOf ClickMenu)
            .ListView.ItemAdd("利用権終期台帳", "利用権終期台帳", "集計一覧", "印刷", AddressOf sub利用権終期管理)
            '.ListView.ItemAdd("町村会固定取り込み", "町村会固定取り込み",ImageKey.作業, "操作", AddressOf sub鹿児島県町村会固定資産取り込み)

            .ListView.ItemAdd("総会資料作成", "総会資料作成", ImageKey.作業, "操作", AddressOf ClickMenu)
            .ListView.ItemAdd("固定資産比較", "固定資産比較", "他システム連携", "操作", AddressOf 固定資産比較)

            .ListView.ItemAdd("CSVto農地", "CSVto農地", ImageKey.作業, "操作", AddressOf CSVto農地)

            MyBase.InitMenu(pMain)
        End With
    End Sub

    Public Overrides ReadOnly Property Get総会締日() As Integer
        Get
            Return 31
        End Get
    End Property

    'Private Sub sub鹿児島県町村会固定資産取り込み()
    '    If Not SysAD.Form.MainTabCtrl.TabPages.ContainsKey("固定資産照合") Then
    '        With New OpenFileDialog
    '            .Filter = "テキストファイル|*.*;*.CSV"
    '            .Title = "固定資産ファイルの取り込み"
    '            If .ShowDialog = DialogResult.OK Then
    '                With New C鹿児島県町村会固定資産取り込み(.FileName)
    '                    .Dialog.StartProc(True, True)

    '                    If .Dialog._objException IsNot Nothing Then
    '                        If .Dialog._objException.Message = "Cancel" Then
    '                            MsgBox("処理を中止しました。　", , "処理中止")
    '                        Else
    '                            'Throw objDlg._objException
    '                        End If
    '                    Else
    '                        Dim pNT As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE [固定照合]=0")
    '                        Dim pPage As New C固定資産照合(.固定TBL, pNT)
    '                        SysAD.Form.MainTabCtrl.TabPages.Add(pPage)
    '                    End If
    '                End With
    '            Else
    '                Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [M_固定情報] WHERE [農地ID]=0")
    '                Dim pNT As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE [固定照合]=0 ORDER BY [地番]")
    '                Dim pPage As New C固定資産照合(pTBL, pNT)
    '                SysAD.Form.MainTabCtrl.TabPages.Add(pPage)
    '            End If
    '        End With
    '    End If


    '    CType(SysAD.Form.MainTabCtrl.TabPages("固定資産照合"), HimTools2012.controls.CTabPageWithToolStrip).Active()
    'End Sub
    Public Overrides ReadOnly Property 市町村別現況地目CD(nType As C市町村別.地目Type) As Integer()
        Get
            Select Case nType
                Case 地目Type.田地目 : Return {10, 11}
                Case 地目Type.畑地目 : Return {20, 21}
                Case 地目Type.農地地目 : Return {10, 11, 20, 21}
                Case 地目Type.その他地目
                    Return MyBase.Make市町村別現況地目コード(nType)
                Case Else
                    Return MyBase.市町村別現況地目CD(nType)
            End Select

            Return MyBase.市町村別現況地目CD(nType)
        End Get
    End Property
End Class

Public Class C鹿児島県町村会固定資産取り込み
    Inherits HimTools2012.clsAccessor
    Private mvarFileName As String
    Public 固定TBL As DataTable

    Public Sub New(ByVal sFileName As String)
        mvarFileName = sFileName

    End Sub

    Public Overrides Sub Execute()
        Message = "CSVファイルを解析しています。"
        Dim sCSV2TB As New HimTools2012.Data.CSV2Table(mvarFileName, System.Text.Encoding.GetEncoding("shift-jis"))
        Message = "固定資産テーブルを読み込んでいます。"

        固定TBL = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [M_固定情報]")
        固定TBL.Columns.Add("チェック済み", GetType(Boolean))
        固定TBL.PrimaryKey = New DataColumn() {固定TBL.Columns("nID")}

        Dim pView As New DataView(sCSV2TB, "[閉鎖フラグ]=-1", "", DataViewRowState.CurrentRows)
        Message = "CSVをデータベースに反映します。"
        Maximum = pView.Count
        Value = 0
        Dim sB As New System.Text.StringBuilder

        For Each pRow As DataRowView In pView
            Value += 1
            Message = "CSVをデータベースに反映します。(" & Value & "/" & Maximum & ")"
            Dim p固定 As C市町村別固定資産 = Nothing
            Select Case SysAD.市町村.市町村名
                Case "大崎町" : p固定 = New C大崎町村会固定(pRow.Row)
                Case "阿久根市" : p固定 = New C阿久根市町村会固定(pRow.Row)
                Case Else
                    Stop
            End Select

            Dim pNewRow As DataRow = 固定TBL.Rows.Find(p固定.ID)

            If pNewRow Is Nothing Then
                Dim sRet As String = p固定.AddNewRow
                If Len(sRet) Then
                    sB.Append(IIf(sB.Length > 0, vbCrLf, "") & sRet)
                    pNewRow = 固定TBL.NewRow
                    pNewRow.Item("ID") = p固定.ID
                    pNewRow.Item("nID") = p固定.ID
                    pNewRow.Item("大字ID") = p固定.大字
                    pNewRow.Item("小字ID") = p固定.小字

                    Dim SST As String = StrConv(Replace(Replace(pRow.Item("地番名称"), "の", "-"), Chr(34), ""), VbStrConv.Narrow)
                    Do Until IsNumeric(SST.Substring(0, 1))
                        If Len(SST) = 1 Then Exit Do
                        SST = SST.Substring(1)
                    Loop
                    pNewRow.Item("地番") = SST
                    ' 
                    pNewRow.Item("一部現況") = pRow.Item("共有区分")
                    pNewRow.Item("登記面積") = pRow.Item("登記地積")
                    pNewRow.Item("登記地目") = pRow.Item("登記地目")

                    pNewRow.Item("現況面積") = pRow.Item("課税地積")
                    pNewRow.Item("現況地目") = pRow.Item("現況地目")

                    pNewRow.Item("所有者ID") = Val(pRow.Item("所有者番号").ToString)
                    If IsDate(Replace(pRow.Item("登記異動年月日"), Chr(34), "")) Then
                        pNewRow.Item("異動年月日") = CDate(Replace(pRow.Item("登記異動年月日"), Chr(34), ""))
                    End If

                    pNewRow.Item("チェック済み") = True
                    固定TBL.Rows.Add(pNewRow)
                End If
            Else
                Dim pDB固定 As New M固定資産(pNewRow)
                If pDB固定.CheckData(p固定) Then
                    pNewRow.Item("チェック済み") = True
                Else
                    Dim sRet As String = pDB固定.Update(p固定)
                    If Len(sRet) Then
                        sB.Append(IIf(sB.Length > 0, vbCrLf, "") & sRet)
                        pNewRow.Item("チェック済み") = True
                    Else
                        Stop
                    End If
                End If
            End If

            If sB.Length > 500 Then
                Debug.Print(sB.ToString)
                Dim sRet2 As String = SysAD.DB(sLRDB).ExecuteSQL(sB.ToString)

                sB.Clear()
            End If

        Next

        If sB.Length > 0 Then
            Debug.Print(sB.ToString)
            Dim sRet2 As String = SysAD.DB(sLRDB).ExecuteSQL(sB.ToString)

            sB.Clear()
        End If

        Dim pView2 As New DataView(固定TBL, "[チェック済み] Is Null Or [チェック済み]=False", "", DataViewRowState.CurrentRows)
        Maximum = pView2.Count
        Value = 0

        For Each pRowV As DataRowView In pView2
            Value += 1
            Message = "消除済み固定資産を処理します。(" & Value & "/" & Maximum & ")"

            If IsDBNull(pRowV.Item("チェック済み")) OrElse pRowV.Item("チェック済み") = False Then
                Dim SR As String = SysAD.DB(sLRDB).ExecuteSQL("DELETE FROM [M_固定情報] WHERE [nID]=" & pRowV.Item("ID"))

                If SR.Length > 0 AndAlso Not SR = "OK" Then
                    Stop
                End If
            Else
                Stop
            End If
        Next
    End Sub
End Class

Public Class M固定資産
    Private mvarRow As DataRow
    Public Sub New(ByRef pRow As DataRow)
        mvarRow = pRow
    End Sub

    Public ReadOnly Property 大字() As Integer
        Get
            Return Val(mvarRow.Item("大字ID").ToString)
        End Get
    End Property
    Public ReadOnly Property 小字() As Integer
        Get
            Return Val(mvarRow.Item("小字ID").ToString)
        End Get
    End Property
    Public ReadOnly Property 地番() As String
        Get
            Return mvarRow.Item("地番").ToString
        End Get
    End Property

    Public ReadOnly Property 一部現況() As Integer
        Get
            Return Val(mvarRow.Item("一部現況").ToString)
        End Get
    End Property


    Public ReadOnly Property 所有者ID() As Integer
        Get
            Return Val(mvarRow.Item("所有者ID").ToString)
        End Get
    End Property

    Public ReadOnly Property 登記地目() As Integer
        Get
            Return Val(mvarRow.Item("登記地目").ToString)
        End Get
    End Property
    Public ReadOnly Property 現況地目() As Integer
        Get
            Return Val(mvarRow.Item("現況地目").ToString)
        End Get
    End Property


    Public ReadOnly Property 登記面積() As Decimal
        Get
            Return Val(mvarRow.Item("登記面積").ToString)
        End Get
    End Property
    Public ReadOnly Property 現況面積() As Decimal
        Get
            Return Val(mvarRow.Item("現況面積").ToString)
        End Get
    End Property


    Public Function CheckData(obj As C市町村別固定資産) As Boolean



        If Not Me.大字 = obj.大字 Then
            Return False
        ElseIf Not Me.小字 = obj.小字 AndAlso Not (Me.小字 = -1 And obj.小字 = 0) Then
            Return False
        ElseIf Not Me.地番 = obj.地番 Then
            Return False
    

        ElseIf Not Me.登記面積 = obj.登記面積 Then
            Return False
        ElseIf Not Me.登記地目 = obj.登記地目 Then
            Return False
        ElseIf Not Me.現況面積 = obj.現況面積 Then
            Return False
        ElseIf Not Me.現況地目 = obj.現況地目 Then
            Return False
        ElseIf Not Me.所有者ID = obj.所有者ID Then
            Return False
        End If



        Return True
    End Function

    Public Function Update(obj As C市町村別固定資産) As String
        Dim sB As New System.Text.StringBuilder

        If Not Me.大字 = obj.大字 Then
            sB.Append(IIf(sB.Length > 0, ", ", "") & "[大字ID]=" & obj.大字)
        End If
        If Not Me.小字 = obj.小字 Then
            sB.Append(IIf(sB.Length > 0, ", ", "") & "[小字ID]=" & obj.小字)
        End If
        If Not Me.地番 = obj.地番 Then
            sB.Append(IIf(sB.Length > 0, ", ", "") & "[地番]='" & obj.地番 & "'")
        End If
   
        If Not Me.登記面積 = obj.登記面積 Then
            sB.Append(IIf(sB.Length > 0, ", ", "") & "[登記面積]=" & obj.登記面積)
        End If
        If Not Me.登記地目 = obj.登記地目 Then
            sB.Append(IIf(sB.Length > 0, ", ", "") & "[登記地目]=" & obj.登記地目)
        End If
        If Not Me.現況面積 = obj.現況面積 Then
            sB.Append(IIf(sB.Length > 0, ", ", "") & "[現況面積]=" & obj.現況面積)
        End If
        If Not Me.現況地目 = obj.現況地目 Then
            sB.Append(IIf(sB.Length > 0, ", ", "") & "[現況地目]=" & obj.現況地目)
        End If
        If Not Me.所有者ID = obj.所有者ID Then
            sB.Append(IIf(sB.Length > 0, ", ", "") & "[所有者ID]=" & obj.所有者ID)
        End If

        If IsDate(obj.異動年月日) Then
            With CDate(obj.異動年月日)
                sB.Append(IIf(sB.Length > 0, ", ", "") & String.Format("[異動年月日]=#{0}/{1}/{2}#", .Month, .Day, .Year))
            End With
        End If

        If sB.Length Then
            Return "UPDATE [M_固定情報] SET " & sB.ToString & " WHERE [nID]=" & Me.mvarRow.Item("ID")
        Else
            Return ""
        End If
    End Function

   
End Class


Public MustInherit Class C市町村別固定資産
    Protected mvarRow As DataRow
    MustOverride ReadOnly Property ID() As Integer
    MustOverride ReadOnly Property 大字() As Integer
    MustOverride ReadOnly Property 小字() As Integer
    MustOverride ReadOnly Property 地番() As String
    MustOverride ReadOnly Property 一部現況() As Integer
    MustOverride ReadOnly Property 登記面積() As Decimal
    MustOverride ReadOnly Property 登記地目() As Integer

    MustOverride ReadOnly Property 現況面積() As Decimal
    MustOverride ReadOnly Property 現況地目() As Integer

    MustOverride ReadOnly Property 所有者ID() As Integer
    MustOverride ReadOnly Property 異動年月日 As Object

    Public Sub New(ByRef pRow As DataRow)
        mvarRow = pRow
    End Sub

    Public Function AddNewRow() As String
        Return String.Format("INSERT INTO [M_固定情報]([ID],[nID],[大字ID],[小字ID],[地番],[登記面積],[現況面積],[登記地目],[現況地目],[所有者ID],[異動年月日]) VALUES({0},{0},{1},{2},'{3}',{4},{5},{6},{7},{8},{9})",
            Me.ID,
            Me.大字,
            Me.小字,
            Me.地番,
            Me.登記面積,
            Me.現況面積,
            Me.登記地目,
            Me.現況地目,
            Me.所有者ID,
            IIf(IsDate(Me.異動年月日), "Null", HimTools2012.StringF.Toリテラル日付(Me.異動年月日.ToString)))
    End Function
End Class

Public Class C大崎町村会固定
    Inherits C市町村別固定資産

    Public Sub New(ByRef pRow As DataRow)
        MyBase.New(pRow)
    End Sub


    Public Overrides ReadOnly Property 大字 As Integer
        Get
            Return Val(mvarRow.Item("大字コード").ToString)
        End Get
    End Property
    Public Overrides ReadOnly Property 小字 As Integer
        Get
            If Val(mvarRow.Item("小字コード").ToString) < 0 Then
                Return 0
            Else
                Return Val(mvarRow.Item("大字コード").ToString) * 1000 + Val(mvarRow.Item("小字コード").ToString)
            End If
        End Get
    End Property

    Public Overrides ReadOnly Property 地番 As String
        Get
            Dim St As String = Replace(Replace(mvarRow.Item("地番名称").ToString, """", ""), "の", "-")
            St = StrConv(St, VbStrConv.Narrow)
            If InStr(St, "(") Then
                St = Left$(St, InStr(St, "(") - 1)
            End If

            Return St
        End Get
    End Property

    Public Overrides ReadOnly Property 一部現況 As Integer
        Get
            If Val(mvarRow.Item("共有区分").ToString) = -1 Then
                Return 0
            Else
                Return Val(mvarRow.Item("共有区分").ToString)
            End If
        End Get
    End Property

    Public Overrides ReadOnly Property 異動年月日 As Object
        Get
            If IsDate(Replace(mvarRow.Item("登記異動年月日"), Chr(34), "")) Then
                Return CDate(Replace(mvarRow.Item("登記異動年月日"), Chr(34), ""))
            Else
                Return DBNull.Value
            End If
        End Get
    End Property

    Public Overrides ReadOnly Property 現況地目 As Integer
        Get
            Return Val(mvarRow.Item("現況地目").ToString)
        End Get
    End Property

    Public Overrides ReadOnly Property 現況面積 As Decimal
        Get
            Return Val(mvarRow.Item("課税地積").ToString)
        End Get
    End Property

    Public Overrides ReadOnly Property 所有者ID As Integer
        Get
            Return Val(mvarRow.Item("所有者番号").ToString)
        End Get
    End Property

    Public Overrides ReadOnly Property 登記地目 As Integer
        Get
            Return Val(mvarRow.Item("登記地目").ToString)
        End Get
    End Property

    Public Overrides ReadOnly Property 登記面積 As Decimal
        Get
            Return Val(mvarRow.Item("登記地積").ToString)
        End Get
    End Property

    Public Overrides ReadOnly Property ID As Integer
        Get
            Return Val(mvarRow.Item("物件番号").ToString)
        End Get
    End Property
End Class


