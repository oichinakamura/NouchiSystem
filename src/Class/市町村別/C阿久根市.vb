
Public Class C阿久根市
    Inherits C市町村別

    ''' <summary>
    ''' 目次: \\tisekikk\MAPDRIVE\システム配置\阿久根市農委\
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        MyBase.New("阿久根市")

    End Sub

    Public Overrides Sub InitMenu(pMain As HimTools2012.SystemWindows.CMainPageSK)
        With pMain
            .ListView.Clear()

            .ListView.ItemAdd("農家検索", "農地・農家検索", "閲覧・検索", "閲覧・検索", AddressOf 農家一覧)
            .ListView.ItemAdd("利用権終期台帳", "利用権終期台帳", "印刷", "印刷", AddressOf sub利用権終期管理)
            .ListView.ItemAdd("固定資産比較", "固定資産比較", "他システム連携", "操作", AddressOf 固定資産比較)

            .ListView.ItemAdd("現地調査表作成", "現地調査表作成", ImageKey.作業, "操作", AddressOf 現地調査)
            .ListView.ItemAdd("受付公布簿", "受付公布簿", "印刷", "印刷", Sub(s, e) SysAD.MainForm.MainTabCtrl.ExistPage("受付公布簿", True, GetType(CTabPage受付公布簿)))

            MyBase.InitMenu(pMain)
        End With
    End Sub

    Public Overrides ReadOnly Property 旧農振都市計画使用 As Boolean
        Get
            Return True
        End Get
    End Property

    Public Overrides Function Get選挙世帯一覧() As System.Data.DataTable
        Return Nothing
    End Function

    Public Overrides Sub InitLocalData()
        sub農地期間満了の終了()
    End Sub

End Class

Public Class C阿久根市町村会固定
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

            If St.EndsWith("番") Then
                St = Replace(St, "番", "")
            End If
            If InStr(St, "番") Then
                St = Replace(St, "番", "-")
            End If

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
