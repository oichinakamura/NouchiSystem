Imports System.Drawing.Design
Imports System.Windows.Forms.Design
Imports System.Threading
Imports HimTools2012.CommonFunc

#Region "列挙"

''' <summary></summary>
''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:51]</remarks>
Public Enum enum自小作別
    ''' <summary></summary>
    自作 = 0
    ''' <summary></summary>
    小作 = 1
    ''' <summary></summary>
    農年 = 2
    ''' <summary></summary>
    やみ小作 = -1
    ''' <summary></summary>
    未入力 = -2
End Enum

''' <summary></summary>
''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:51]</remarks>
Public Enum enum農業振興地域
    ''' <summary></summary>
    農用地外 = 0
    ''' <summary></summary>
    農用地内 = 1
    ''' <summary></summary>
    振興地域外 = 2
End Enum

Public Enum enum農振法区分
    不明 = 0
    農用地区域 = 1
    農振地域 = 2
    農振地域外 = 3
    その他 = 4
    調査中 = 5
End Enum

Public Enum enum都市計画法
    ''' <summary></summary>
    都市計画法外 = 0
    ''' <summary></summary>
    都市計画法内 = 1
    ''' <summary></summary>
    用途地域内 = 2
    ''' <summary></summary>
    調整区域内 = 3
    ''' <summary></summary>
    市街化区域内 = 4
    ''' <summary></summary>
    都市計画白地 = 5
End Enum

Public Enum enum都市計画法区分
    不明 = 0
    市街化区域 = 1
    市街化調整区域 = 2
    非線引き都市計画区域の用途地域 = 3
    都市計画区域外 = 4
    その他 = 5
    調査中 = 6
    非線引き都市計画区域内 = 7
End Enum

Public Enum enum土地改良法
    ''' <summary></summary>
    区域外 = 0
    ''' <summary></summary>
    区域内_整備済 = 1
    ''' <summary></summary>
    区域内_整備中 = 2
End Enum
''' <summary></summary>
''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:51]</remarks>
Public Enum enum小作地適用法
    ''' <summary></summary>
    不明 = 0
    ''' <summary></summary>
    農地法 = 1
    ''' <summary></summary>
    基盤法 = 2
    ''' <summary></summary>
    特定農地貸付法 = 3
    ''' <summary></summary>
    その他 = 4
End Enum

''' <summary></summary>
''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:51]</remarks>
Public Enum enum小作形態
    ''' <summary></summary>
    不明 = 0
    ''' <summary></summary>
    賃貸借 = 1
    ''' <summary></summary>
    使用貸借 = 2
    ''' <summary></summary>
    その他 = 3
    ''' <summary></summary>
    地上権 = 4
    ''' <summary></summary>
    永小作権 = 5
    ''' <summary></summary>
    質権 = 6
    ''' <summary></summary>
    期間借地 = 7
    ''' <summary></summary>
    残存小作地 = 8
    ''' <summary></summary>
    使用賃借 = 9
End Enum


''' <summary></summary>
''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:51]</remarks>
Public Enum enum様式1
    ''' <summary></summary>
    不明 = 0
    ''' <summary></summary>
    農地法1号 = 1
    ''' <summary></summary>
    農地法2号 = 2
    ''' <summary></summary>
    農地法3号 = 3
End Enum
''' <summary></summary>
''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:51]</remarks>
Public Enum enum様式2
    ''' <summary></summary>
    不明 = 0
    ''' <summary></summary>
    農地法1号 = 1
    ''' <summary></summary>
    農地法2号 = 2
    ''' <summary></summary>
    遊休農地でない = 3
End Enum
''' <summary></summary>
''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:51]</remarks>
Public Enum enum様式3
    ''' <summary></summary>
    不明 = 0
    ''' <summary></summary>
    農地法1号 = 1
    ''' <summary></summary>
    農地法2号 = 2
    ''' <summary></summary>
    農地法3号 = 3
    ''' <summary></summary>
    ただし書 = 4
End Enum

''' <summary></summary>
''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:51]</remarks>
Public Enum enum租税処置法
    ''' <summary></summary>
    不明 = 0
    ''' <summary></summary>
    租税処置法1号 = 1
    ''' <summary></summary>
    租税処置法2号 = 2
    ''' <summary></summary>
    租税処置法3号 = 3
End Enum
''' <summary></summary>
''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
Public Enum enum納税猶予
    ''' <summary></summary>
    不明 = 0
    ''' <summary></summary>
    贈与税 = 1
    ''' <summary></summary>
    相続税 = 2
End Enum
''' <summary></summary>
''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
Public Enum 農振農用地区分
    ''' <summary></summary>
    他 = 0
    ''' <summary></summary>
    内 = 1
    ''' <summary></summary>
    外 = 2
End Enum

''' <summary></summary>
''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
Public Enum enum利用状況調査農地法
    ''' <summary></summary>
    不明 = 0
    ''' <summary></summary>
    農法32条1項1号 = 1
    ''' <summary></summary>
    農法32条1項2号 = 2
    ''' <summary></summary>
    遊休農地でない = 3
    ''' <summary></summary>
    その他 = 4
    ''' <summary></summary>
    調査不可 = 5
    ''' <summary></summary>
    調査中 = 6
End Enum

''' <summary></summary>
''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
Public Enum enum利用状況調査荒廃
    ''' <summary></summary>
    不明 = 0
    ''' <summary></summary>
    A分類 = 1
    ''' <summary></summary>
    B分類 = 2
    ''' <summary></summary>
    調査中 = 3
End Enum
''' <summary></summary>
''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
Public Enum enum利用状況調査荒廃内訳
    ''' <summary></summary>
    不明 = 0
    ''' <summary></summary>
    山林 = 1
    ''' <summary></summary>
    原野 = 2
    ''' <summary></summary>
    宅地 = 3
    ''' <summary></summary>
    雑種地 = 4
    ''' <summary></summary>
    その他 = 5
End Enum
''' <summary></summary>
''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
Public Enum enum利用状況調査転用
    ''' <summary></summary>
    不明 = 0
    ''' <summary></summary>
    一時転用 = 1
    ''' <summary></summary>
    無断転用 = 2
    ''' <summary></summary>
    違反転用 = 3
End Enum
''' <summary></summary>
''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
Public Enum enum利用意向根拠条項
    ''' <summary></summary>
    不明 = 0
    ''' <summary></summary>
    農地法第32条第1項 = 1
    ''' <summary></summary>
    農地法第32条第4項 = 2
    ''' <summary></summary>
    農地法第33条第1項 = 3
End Enum
''' <summary></summary>
''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
Public Enum enum利用意向内容区分
    ''' <summary></summary>
    不明 = 0
    ''' <summary></summary>
    自ら耕作 = 1
    ''' <summary></summary>
    機構事業 = 2
    ''' <summary></summary>
    所有者代理事業 = 3
    ''' <summary></summary>
    権利設定または移転 = 4
    ''' <summary></summary>
    その他 = 5
End Enum
''' <summary></summary>
''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
Public Enum enum利用意向権利関係調査区分
    ''' <summary></summary>
    不明 = 0
    ''' <summary></summary>
    対象外 = 1
    ''' <summary></summary>
    調査中 = 2
    ''' <summary></summary>
    調査済み = 3
End Enum
#End Region

''' <summary></summary>
''' <remarks>
''' 未検証 件数91
''' </remarks>
Public Class CObj農地
    Inherits CTargetObjWithView農地台帳

    ''' <summary></summary>
    ''' <returns>[Private][class:System.String]</returns>
    ''' <remarks>Verified [中村 雄一 date：2016/10/18 16:52]</remarks>
    ''' <作業履歴>
    '''  <作業内容 Date="2016/9/16" 作業者="中村 雄一">サンプル01。</作業内容>
    '''  <作業内容 Date="2016/9/17" 作業者="中村 雄一">サンプル02。</作業内容>
    ''' </作業履歴>
    Public Overrides Function ToString() As String
        Return Me.土地所在
    End Function

    ''' <summary>CObj農地のコンストラクタ</summary>
    ''' <param name="pRow">class：System.Data.DataRow</param>
    ''' <param name="bAddNew">class：System.Boolean</param>
    ''' <remarks>Verified [中村 雄一 date：2016/9/29 16:1]</remarks>
    Public Sub New(ByVal pRow As DataRow, ByVal bAddNew As Boolean)
        MyBase.New(pRow, bAddNew, New HimTools2012.TargetSystem.DataKey("農地", pRow.Item("ID")), "D:農地Info")
        If pRow Is Nothing Then
            Stop
        End If
    End Sub

    ''' <summary></summary>
    ''' <param name="sParam">class：System.String</param>
    ''' <returns>[Private][class:System.Object]</returns>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
    Public Overrides Function GetProperty(ByVal sParam As String) As Object
        Try
            Select Case sParam
                Case "借受人ID"
                    Select Case Me.Row.Body.Item("自小作別")
                        Case 0 : Return 0
                        Case Else
                            Return Me.GetItem("借受人ID", 0)
                    End Select
                Case "管理者ID" : Return Me.GetItem("管理者ID", 0)
                Case "管理者"
                    Dim nID As Decimal = Me.GetProperty("管理者ID")
                    If nID <> 0 Then
                        Dim pRow As DataRow = App農地基本台帳.TBL個人.FindRowByID(nID)
                        If pRow Is Nothing Then
                            Return Nothing
                        Else
                            Return New CObj個人(pRow, False)
                        End If
                    Else
                        Return Me.GetProperty("所有者")
                    End If
                Case "土地所在"
                    Return Me.土地所在
                Case "経由農業生産法人ID" : Return Me.GetItem("経由農業生産法人ID", 0)
                Case "Obj経由農業生産法人" : Return ObjectMan.GetObjectDB("個人." & Me.GetProperty("経由農業生産法人ID"), App農地基本台帳.TBL個人.FindRowByID(Me.GetProperty("経由農業生産法人ID")), GetType(CObj個人), True)
                Case "名称" : Return Me.Row.Body.Item("土地所在").ToString
                Case "Obj貸人" : Return Me.GetProperty("管理者")
                Case "Obj受人" : Return ObjectMan.GetObjectDB("個人." & Me.借受人ID, App農地基本台帳.TBL個人.FindRowByID(Me.借受人ID), GetType(CObj個人), True)
                Case "所有者" : Return ObjectMan.GetObject("個人." & Val(Me.GetItem("所有者ID").ToString))
                Case "所有者名" : Return CType(Me.GetProperty("所有者"), CObj個人).氏名
                Case "Obj貸人"
                    If IsDBNull(Me.Row.Body.Item("所有者ID")) Then
                        Return Nothing
                    Else
                        Dim pRow貸人 As DataRow = App農地基本台帳.TBL個人.FindRowByID(Me.GetItem("所有者ID"))
                        If pRow貸人 IsNot Nothing Then
                            Return New CObj個人(pRow貸人, False)
                        Else
                            Return Nothing
                        End If
                    End If
                Case "Obj受人"
                    Select Case Me.Row.Body.Item("自小作別")
                        Case 0 : Return 0
                        Case Else
                            If IsDBNull(Me.Row.Body.Item("借受人ID")) Then
                                Return Nothing
                            Else
                                Dim pRow受人 As DataRow = App農地基本台帳.TBL個人.FindRowByID(Me.Row.Body.Item("借受人ID"))
                                If pRow受人 IsNot Nothing Then
                                    Return New CObj個人(pRow受人, False)
                                Else
                                    Return Nothing
                                End If
                            End If
                    End Select
                Case Else

                    Return Me.Row.Body.Item(sParam)
            End Select
            Return ""
        Catch ex As Exception
        End Try
        Return ""
    End Function

    ''' <summary></summary>
    ''' <param name="NewID">class：System.Nullable(Of System.Int64)</param>
    ''' <returns>[Private][class:HimTools2012.TargetSystem.CTargetObjectBase]</returns>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
    Public Function CopyObject(Optional ByVal NewID As Long? = Nothing) As HimTools2012.TargetSystem.CTargetObjectBase
        Dim sKeyIP As String = ""
        Dim adrList As System.Net.IPAddress() = SysAD.IPAddressList()
        If adrList.Length > 0 Then
            For Each padr As System.Net.IPAddress In adrList
                sKeyIP = Replace(padr.ToString, ".", "")
                If sKeyIP.IndexOf(":") = -1 AndAlso sKeyIP.Length > 4 Then
                    Exit For
                End If
            Next
        End If

        If NewID Is Nothing Then
            Dim pTBLMin As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Min([D:農地Info].ID) AS IDMin FROM [D:農地Info];")
            If pTBLMin.Rows(0).Item("IDMin") = 0 Then
                NewID = -1
            Else
                NewID = CLng(pTBLMin.Rows(0).Item("IDMin") - 1)
            End If
        End If

        Try
            SysAD.DB(sLRDB).ExecuteSQL("SELECT * INTO [農地追加{0}] FROM [D:農地Info] WHERE [ID]={1};", sKeyIP, Me.ID)
            SysAD.DB(sLRDB).ExecuteSQL("UPDATE [農地追加{0}] SET [ID]={2} WHERE [ID]={1}", sKeyIP, Me.ID, NewID)
            SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D:農地Info] SELECT * FROM 農地追加{0}", sKeyIP)
            SysAD.DB(sLRDB).ExecuteSQL("DROP TABLE [農地追加{0}]", sKeyIP)
        Catch ex As Exception
            Return Nothing
        End Try

        Return ObjectMan.GetObject("農地." & NewID)
    End Function


    ''' <summary></summary>
    ''' <value>[Public][ReadOnly]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
    Public Overrides ReadOnly Property DataTableWithUpdateList As HimTools2012.Data.DataTableWithUpdateList
        Get
            Return App農地基本台帳.TBL農地
        End Get
    End Property

#Region "プロパティ"

    '経営農地の筆別表-1
#Region "01_基本情報"

    ''' <summary></summary>
    ''' <value>[Public][ReadOnly]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
    Public ReadOnly Property 土地所在() As String
        Get
            If MyBase.GetStringValue("所在").Length > 0 Then
                Return MyBase.GetStringValue("所在") & MyBase.GetStringValue("地番")
            Else
                Dim sB As New System.Text.StringBuilder
                If Not IsDBNull(Me.Row.Body("大字ID")) AndAlso Me.Row.Body("大字ID") <> 0 Then
                    Dim p大字() As DataRowView = SysAD.MasterView("大字").FindRows(Me.Row.Body("大字ID"))
                    If p大字 IsNot Nothing Then
                        sB.Append(p大字(0).Item("名称"))
                    End If
                End If

                If Not Me.Row.IsZero("小字ID") Then
                    Dim p小字() As DataRowView = SysAD.MasterView("小字").FindRows(Me.Row.Body("小字ID"))
                    If p小字 IsNot Nothing AndAlso p小字.Length > 0 Then
                        Dim s小字 As String = p小字(0).Item("名称").ToString
                        If s小字.Length > 0 AndAlso Replace(s小字, "-", "").Length > 0 Then

                            sB.Append("字" & p小字(0).Item("名称"))
                        End If
                    End If
                End If

                Return sB.ToString & MyBase.GetStringValue("地番")
            End If

            Return MyBase.GetStringValue("所在")
        End Get
    End Property


    ''' <summary></summary>
    ''' <value>[Public][ReadOnly]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
    Public ReadOnly Property 大字 As String
        Get
            Return GetStringValue("大字")
        End Get
    End Property

    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
    Public Property 大字ID As Integer
        Get
            Return GetIntegerValue("大字ID")
        End Get
        Set(ByVal value As Integer)
            ValueChange("大字ID", value)
        End Set
    End Property

    ''' <summary></summary>
    ''' <value>[Public][ReadOnly]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
    Public ReadOnly Property 小字 As String
        Get
            Return GetStringValue("小字")
        End Get
    End Property

    ''' <summary></summary>
    ''' <value>[Public][ReadOnly]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
    Public ReadOnly Property 小字ID As Integer
        Get
            Return GetIntegerValue("小字ID")
        End Get
    End Property


    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
    Public Property 所在 As String
        Get
            Return GetStringValue("所在")
        End Get
        Set(ByVal value As String)
            ValueChange("所在", value)
        End Set
    End Property



    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
    Public Property 地番 As String
        Get
            Return GetStringValue("地番")
        End Get
        Set(ByVal value As String)
            ValueChange("地番", value)
        End Set
    End Property

    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
    Public Property 耕地番号 As Integer
        Get
            Return GetIntegerValue("耕地番号")
        End Get
        Set(ByVal value As Integer)
            ValueChange("耕地番号", value)
        End Set
    End Property
#End Region

#Region "02_地目"

    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
    Public Property 登記簿地目 As Integer
        Get
            Return GetIntegerValue("登記簿地目")
        End Get
        Set(ByVal value As Integer)
            ValueChange("登記簿地目", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
    Public Property 現況地目 As Integer
        Get
            Return GetIntegerValue("現況地目")
        End Get
        Set(ByVal value As Integer)
            ValueChange("現況地目", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
    Public Property 農委地目 As Integer
        Get
            Return GetIntegerValue("農委地目ID")
        End Get
        Set(ByVal value As Integer)
            ValueChange("農委地目ID", value)
        End Set
    End Property
#End Region

#Region "03_面積情報"
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
    Public Property 登記簿面積 As Decimal
        Get
            Return GetDecimalValue("登記簿面積")
        End Get
        Set(ByVal value As Decimal)
            ValueChange("登記簿面積", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
    Public Property 実面積 As Decimal
        Get
            Return GetDecimalValue("実面積")
        End Get
        Set(ByVal value As Decimal)
            ValueChange("実面積", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
    Public Property 本地面積 As Integer
        Get
            Return GetIntegerValue("本地面積")
        End Get
        Set(ByVal value As Integer)
            ValueChange("本地面積", value)
        End Set
    End Property

    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
    Public Property 田面積 As Decimal
        Get
            Return GetDecimalValue("田面積")
        End Get
        Set(ByVal value As Decimal)
            ValueChange("田面積", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
    Public Property 畑面積 As Decimal
        Get
            Return GetDecimalValue("畑面積")
        End Get
        Set(ByVal value As Decimal)
            ValueChange("畑面積", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
    Public Property 樹園地 As Decimal
        Get
            Return GetDecimalValue("樹園地")
        End Get
        Set(ByVal value As Decimal)
            ValueChange("樹園地", value)
        End Set
    End Property

    Public Property 採草放牧地 As Decimal
        Get
            Return GetDecimalValue("採草放牧面積")
        End Get
        Set(ByVal value As Decimal)
            ValueChange("採草放牧面積", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public][ReadOnly]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
    Public ReadOnly Property 耕地計 As Decimal
        Get
            Return Me.田面積 + Me.畑面積 + Me.樹園地
        End Get
    End Property
#End Region

#Region "04_農地区分"
    ''' <summary></summary>
    Public Const n農地区分 = 15
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
    Public Property 旧農振区分 As enum農業振興地域
        Get
            Return GetIntegerValue("農業振興地域")
        End Get
        Set(ByVal value As enum農業振興地域)
            ValueChange("農業振興地域", value)
        End Set
    End Property

    Public Property 農振法区分 As enum農振法区分
        Get
            Return GetIntegerValue("農振法区分")
        End Get
        Set(ByVal value As enum農振法区分)
            ValueChange("農振法区分", value)
        End Set
    End Property

    Public Property 都市計画法 As enum都市計画法
        Get
            Return GetIntegerValue("都市計画法")
        End Get
        Set(ByVal value As enum都市計画法)
            ValueChange("都市計画法", value)
        End Set
    End Property

    Public Property 都市計画法区分 As enum都市計画法区分
        Get
            Return GetIntegerValue("都市計画法区分")
        End Get
        Set(ByVal value As enum都市計画法区分)
            ValueChange("都市計画法区分", value)
        End Set
    End Property
    Public Property 土地改良法 As enum土地改良法
        Get
            Return GetIntegerValue("土地改良法")
        End Get
        Set(ByVal value As enum土地改良法)
            ValueChange("土地改良法", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
    Public Property 生産緑地法 As enum有無
        Get
            Return GetIntegerValue("生産緑地法")
        End Get
        Set(ByVal value As enum有無)
            ValueChange("生産緑地法", value)
        End Set
    End Property
#End Region

#Region "06_所有者情報"
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
    Public Property 所有者ID As Long
        Get
            Return GetLongIntValue("所有者ID")
        End Get
        Set(ByVal value As Long)
            ValueChange("所有者ID", value)
        End Set
    End Property

    ''' <summary></summary>
    ''' <value>[Public][ReadOnly]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
    Public ReadOnly Property 所有者氏名 As String
        Get
            Return GetStringValue("所有者氏名")
        End Get
    End Property

    ''' <summary></summary>
    ''' <value>[Public][ReadOnly]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
    Public ReadOnly Property 所有者住所 As String
        Get
            Return GetStringValue("所有者住所")
        End Get
    End Property

#End Region

#Region "09_貸借情報"
    ''' <summary></summary>
    Public Const n賃借情報 = 700
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:52]</remarks>
    Public Property 自小作別 As enum自小作別
        Get
            Return GetIntegerValue("自小作別")
        End Get
        Set(ByVal value As enum自小作別)
            ValueChange("自小作別", value)
        End Set
    End Property


    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 借受世帯ID As Long
        Get
            Return GetLongIntValue("借受世帯ID")
        End Get
        Set(ByVal value As Long)
            ValueChange("借受世帯ID", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 借受人ID As Long
        Get
            Return GetLongIntValue("借受人ID")
        End Get
        Set(ByVal value As Long)
            ValueChange("借受人ID", value)
        End Set
    End Property

    ''' <summary></summary>
    ''' <value>[Public][ReadOnly]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public ReadOnly Property 借受人氏名 As String
        Get
            Return GetStringValue("借受人氏名")
        End Get
    End Property

    ''' <summary></summary>
    ''' <value>[Public][ReadOnly]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public ReadOnly Property 借受人住所 As String
        Get
            Return GetStringValue("適用法令")
        End Get
    End Property

    ''' <summary></summary>
    ''' <value>[Public][ReadOnly]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public ReadOnly Property 適用法令 As String
        Get
            Return GetStringValue("適用法令")
        End Get
    End Property


    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 小作地適用法ID As Integer
        Get
            Return GetIntegerValue("小作地適用法")
        End Get
        Set(ByVal value As Integer)
            ValueChange("小作地適用法", value)
        End Set
    End Property

    ''' <summary></summary>
    ''' <value>[Public][ReadOnly]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public ReadOnly Property 小作形態種別 As String
        Get
            Return GetStringValue("小作形態種別")
        End Get
    End Property

    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 小作形態 As Integer
        Get
            Return GetIntegerValue("小作形態")
        End Get
        Set(ByVal value As Integer)
            ValueChange("小作形態", value)
        End Set
    End Property


    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 貸借始期 As DateTime
        Get
            Return GetDateValue("貸借始期")
        End Get
        Set(ByVal value As DateTime)
            ValueChange("貸借始期", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 貸借終期 As DateTime
        Get
            Return GetDateValue("貸借終期")
        End Get
        Set(ByVal value As DateTime)
            ValueChange("貸借終期", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 小作開始年月日 As DateTime
        Get
            Return GetDateValue("小作開始年月日")
        End Get
        Set(ByVal value As DateTime)
            ValueChange("小作開始年月日", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 小作終了年月日 As DateTime
        Get
            Return GetDateValue("小作終了年月日")
        End Get
        Set(ByVal value As DateTime)
            ValueChange("小作終了年月日", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 小作料 As Integer
        Get
            Return GetIntegerValue("小作料")
        End Get
        Set(ByVal value As Integer)
            ValueChange("小作料", value)
        End Set
    End Property

    Public Property 物納 As String
        Get
            Return GetStringValue("物納")
        End Get
        Set(ByVal value As String)
            ValueChange("物納", value)
        End Set
    End Property

    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 小作料単位 As String
        Get
            Return GetStringValue("小作料単位")
        End Get
        Set(ByVal value As String)
            ValueChange("小作料単位", value)
        End Set
    End Property

    Public Property 物納単位 As String
        Get
            Return GetStringValue("物納単位")
        End Get
        Set(ByVal value As String)
            ValueChange("物納単位", value)
        End Set
    End Property

    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 賃借料10a As Decimal
        Get
            Return GetDecimalValue("10a賃借料")
        End Get
        Set(ByVal value As Decimal)
            ValueChange("10a賃借料", value)
        End Set
    End Property

    ''' <summary></summary>
    ''' <value>[Public][ReadOnly]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public ReadOnly Property 小作料表示 As String
        Get
            Dim sB As New System.Text.StringBuilder


            If Not IsDBNull(Row.Item("小作料")) Then
                sB.Append(Row.Item("小作料").ToString)
            End If

            If Not Me.小作料単位.Length = 0 Then
                sB.Append(Me.小作料単位)
            End If

            Return sB.ToString
        End Get
    End Property
    Public ReadOnly Property 物納表示 As String
        Get
            Dim sB As New System.Text.StringBuilder

            If Not IsDBNull(Row.Item("物納")) Then
                sB.Append(Row.Item("物納").ToString)
            End If

            If Not Me.物納単位.Length = 0 Then
                sB.Append(Me.物納単位)
            End If

            Return sB.ToString
        End Get
    End Property

    Public Property 経由法人ID As Decimal
        Get
            Return GetLongIntValue("経由農業生産法人ID")
        End Get
        Set(ByVal value As Decimal)
            ValueChange("経由農業生産法人ID", value)
        End Set
    End Property

    Public Property 機構契約開始年月日 As Date
        Get
            Return GetDateValue("利用配分計画始期日")
        End Get
        Set(ByVal value As Date)
            ValueChange("利用配分計画始期日", value)
        End Set
    End Property

    Public Property 機構契約終了年月日 As Date
        Get
            Return GetDateValue("利用配分計画終期日")
        End Get
        Set(ByVal value As Date)
            ValueChange("利用配分計画終期日", value)
        End Set
    End Property
#End Region

#Region "農地等の利用状況"

    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 農地状況 As Integer
        Get
            Return GetIntegerValue("農地状況")
        End Get
        Set(ByVal value As Integer)
            ValueChange("農地状況", value)
        End Set
    End Property

    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 利用状況報告対象 As enum有無
        Get
            Return GetIntegerValue("利用状況報告対象")
        End Get
        Set(ByVal value As enum有無)
            ValueChange("利用状況報告対象", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 利用状況報告年月日 As DateTime
        Get
            Return GetDateValue("利用状況報告年月日")
        End Get
        Set(ByVal value As DateTime)
            ValueChange("利用状況報告年月日", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 是正勧告日 As DateTime
        Get
            Return GetDateValue("是正勧告日")
        End Get
        Set(ByVal value As DateTime)
            ValueChange("是正勧告日", value)
        End Set
    End Property

#End Region

    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 是正内容 As String
        Get
            Return GetStringValue("是正内容")
        End Get
        Set(ByVal value As String)
            ValueChange("是正内容", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 是正期限 As DateTime
        Get
            Return GetDateValue("是正期限")
        End Get
        Set(ByVal value As DateTime)
            ValueChange("是正期限", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 根拠条件農地法 As enum様式1
        Get
            Return GetIntegerValue("根拠条件農地法")
        End Get
        Set(ByVal value As enum様式1)
            ValueChange("根拠条件農地法", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 根拠条件基盤強化法 As enum様式1
        Get
            Return GetIntegerValue("根拠条件基盤強化法")
        End Get
        Set(ByVal value As enum様式1)
            ValueChange("根拠条件基盤強化法", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 是正確認 As DateTime
        Get
            Return GetDateValue("是正確認")
        End Get
        Set(ByVal value As DateTime)
            ValueChange("是正確認", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 是正状況 As String
        Get
            Return GetStringValue("是正状況")
        End Get
        Set(ByVal value As String)
            ValueChange("是正状況", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 取消年月日 As DateTime
        Get
            Return GetDateValue("取消年月日")
        End Get
        Set(ByVal value As DateTime)
            ValueChange("取消年月日", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 取消事由 As String
        Get
            Return GetStringValue("取消事由")
        End Get
        Set(ByVal value As String)
            ValueChange("取消事由", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 取消条件農地法 As enum様式2
        Get
            Return GetIntegerValue("取消条件農地法")
        End Get
        Set(ByVal value As enum様式2)
            ValueChange("取消条件農地法", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 取消条件基盤強化法 As enum様式2
        Get
            Return GetIntegerValue("取消条件基盤強化法")
        End Get
        Set(ByVal value As enum様式2)
            ValueChange("取消条件基盤強化法", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 届出年月日 As DateTime
        Get
            Return GetDateValue("届出年月日")
        End Get
        Set(ByVal value As DateTime)
            ValueChange("届出年月日", value)
        End Set
    End Property


#Region "16_農地の利用状況調査"
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 利用状況調査日 As DateTime
        Get
            Return GetDateValue("利用状況調査日")
        End Get
        Set(ByVal value As DateTime)
            ValueChange("利用状況調査日", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 利用状況調査農地法 As enum利用状況調査農地法
        Get
            Return GetIntegerValue("利用状況調査農地法")
        End Get
        Set(ByVal value As enum利用状況調査農地法)
            ValueChange("利用状況調査農地法", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 利用状況調査荒廃 As enum利用状況調査荒廃
        Get
            Return GetIntegerValue("利用状況調査荒廃")
        End Get
        Set(ByVal value As enum利用状況調査荒廃)
            ValueChange("利用状況調査荒廃", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 利用状況調査荒廃内訳 As enum利用状況調査荒廃内訳
        Get
            Return GetIntegerValue("利用状況調査荒廃内訳")
        End Get
        Set(ByVal value As enum利用状況調査荒廃内訳)
            ValueChange("利用状況調査荒廃内訳", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 利用状況調査転用 As enum利用状況調査転用
        Get
            Return GetIntegerValue("利用状況調査転用")
        End Get
        Set(ByVal value As enum利用状況調査転用)
            ValueChange("利用状況調査転用", value)
        End Set
    End Property
#End Region

#Region "17_農地の利用意向調査"
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 利用意向調査日 As DateTime
        Get
            Return GetDateValue("利用意向調査日")
        End Get
        Set(ByVal value As DateTime)
            ValueChange("利用意向調査日", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 利用意向根拠条項 As enum利用意向根拠条項
        Get
            Return GetIntegerValue("利用意向根拠条項")
        End Get
        Set(ByVal value As enum利用意向根拠条項)
            ValueChange("利用意向根拠条項", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 利用意向意思表明日 As DateTime
        Get
            Return GetDateValue("利用意向意思表明日")
        End Get
        Set(ByVal value As DateTime)
            ValueChange("利用意向意思表明日", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 利用意向意向内容区分 As enum利用意向内容区分
        Get
            Return GetIntegerValue("利用意向意向内容区分")
        End Get
        Set(ByVal value As enum利用意向内容区分)
            ValueChange("利用意向意向内容区分", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 利用意向権利関係調査区分 As enum利用意向権利関係調査区分
        Get
            Return GetIntegerValue("利用意向権利関係調査区分")
        End Get
        Set(ByVal value As enum利用意向権利関係調査区分)
            ValueChange("利用意向権利関係調査区分", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 利用意向権利関係調査記録 As String
        Get
            Return GetStringValue("利用意向権利関係調査記録")
        End Get
        Set(ByVal value As String)
            ValueChange("利用意向権利関係調査記録", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 利用意向公示年月日 As DateTime
        Get
            Return GetDateValue("利用意向公示年月日")
        End Get
        Set(ByVal value As DateTime)
            ValueChange("利用意向公示年月日", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 利用意向通知年月日 As DateTime
        Get
            Return GetDateValue("利用意向通知年月日")
        End Get
        Set(ByVal value As DateTime)
            ValueChange("利用意向通知年月日", value)
        End Set
    End Property

    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 利用配分計画借賃額 As Decimal
        Get
            Return GetDecimalValue("利用配分計画借賃額")
        End Get
        Set(ByVal value As Decimal)
            ValueChange("利用配分計画借賃額", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 利用配分計画10a賃借料 As Decimal
        Get
            Return GetDecimalValue("利用配分計画10a賃借料")
        End Get
        Set(ByVal value As Decimal)
            ValueChange("利用配分計画10a賃借料", value)
        End Set
    End Property
#End Region


    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 備考 As String
        Get
            Return GetStringValue("備考")
        End Get
        Set(ByVal value As String)
            ValueChange("備考", value)
        End Set
    End Property

    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 固定資産番号 As Integer
        Get
            Return GetIntegerValue("固定資産番号")
        End Get
        Set(ByVal value As Integer)
            ValueChange("固定資産番号", value)
        End Set
    End Property
    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Property 市町村ID As Integer
        Get
            Return GetIntegerValue("市町村ID")
        End Get
        Set(ByVal value As Integer)
            ValueChange("市町村ID", value)
        End Set
    End Property
#End Region
    ''' <summary></summary>
    Public 構成点 As New List(Of Point)

    ''' <summary></summary>
    ''' <param name="pMenu">class：HimTools2012.controls.MenuItemEX</param>
    ''' <param name="nDips">class：System.Int32</param>
    ''' <param name="sParam">class：</param>
    ''' <returns>[Private][class:HimTools2012.controls.MenuPlus]</returns>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Overrides Function GetContextMenu(Optional ByVal pMenu As HimTools2012.controls.MenuItemEX = Nothing, Optional ByVal nDips As Integer = 1, Optional ByVal sParam() As String = Nothing) As HimTools2012.controls.MenuPlus
        Dim pMenuf As New HimTools2012.controls.ContextMenuEX(AddressOf ClickMenu)
        Dim bEdit As Boolean = (SysAD.SystemInfo.ユーザー.n権利 > 0)

        With pMenuf
            .AddMenu("開く", , AddressOf ClickMenu, , (Me.DataViewPage Is Nothing))
            .AddMenu("履歴一覧", , Sub(s, e) SysAD.page農家世帯.土地履歴リスト.検索開始("[LID]=" & Me.ID, "[LID]=" & Me.ID))
            .AddMenu("履歴の追加", , AddressOf ClickMenu, , bEdit)
            .InsertSeparator()

            Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("Select * From [D_土地系図] Where [自ID]=" & Me.ID)
            If pTBL.Rows.Count > 0 Then
                With .AddMenu("異動前農地")
                    For Each pRow As DataRow In pTBL.Rows
                        .AddSubMenu(pRow.Item("元土地所在"), 2, ObjectMan.GetObject("削除農地." & pRow.Item("元ID")))

                        'Dim p元TBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("Select * From [D_土地系図] Where [自ID]=" & pRow.Item("元ID"))
                        'If p元TBL.Rows.Count > 0 Then
                        '    With .AddMenu("異動前農地")
                        '        For Each p元Row As DataRow In p元TBL.Rows
                        '            .AddSubMenu(p元Row.Item("元土地所在"), 2, ObjectMan.GetObject("削除農地." & p元Row.Item("元ID")))
                        '        Next
                        '    End With
                        'End If
                    Next

                End With
            End If

            .InsertSeparator()
            .AddMenuByText({"閲覧用農地台帳印刷", "農地台帳要約書印刷", "-", "関連申請"}, AddressOf ClickMenu, bEdit)

            If Me.GetDecimalValue("管理世帯ID") <> 0 AndAlso Me.GetDecimalValue("管理世帯ID") <> Me.GetDecimalValue("所有世帯ID") Then .AddMenu("管理世帯を呼ぶ", , AddressOf ClickMenu)
            If Me.GetDecimalValue("所有世帯ID") <> 0 Then .AddSubMenu("所有世帯", 2, ObjectMan.GetObject("農家." & Me.GetDecimalValue("所有世帯ID")))

            .AddSubMenu("所有者", 2, ObjectMan.GetObject("個人." & Me.GetDecimalValue("所有者ID")))
            If Me.自小作別 <> enum自小作別.自作 Then
                .InsertSeparator()
                If Me.GetDecimalValue("借受世帯ID") <> 0 Then .AddSubMenu("借受世帯", 2, ObjectMan.GetObject("農家." & Me.GetDecimalValue("借受世帯ID")))
                If Me.GetDecimalValue("借受人ID") <> 0 Then .AddSubMenu("借受人", 2, ObjectMan.GetObject("個人." & Me.GetDecimalValue("借受人ID")))

                With .AddMenu("解約／再設定")
                    .AddMenuByText({"議案書付解約", "議案書無し合意解約", "職権解約", "-"}, AddressOf ClickMenu, bEdit)
                    .AddMenu("期間満了の終了", , AddressOf ClickMenu)
                    Select Case Val(Row.Body.Item("小作地適用法").ToString)
                        Case 1
                            With .AddMenu("期間の延長")
                                .AddMenu("関連申請の延長", , AddressOf ClickMenu)
                                .AddMenu("選択農地のみ延長", , AddressOf ClickMenu)
                            End With
                        Case 2 : .AddMenu("貸借の再設定", , AddressOf ClickMenu)
                    End Select
                    If Val(Row.Body.Item("経由農業生産法人ID").ToString) <> 0 AndAlso Row.Body.Item("農業生産法人経由貸借") = True Then
                        .AddMenuByText({"-", "中間管理機構へ農地の返還"}, AddressOf ClickMenu, bEdit)
                    End If

                End With
            End If
            .InsertSeparator()
            .AddMenu("分筆", , AddressOf ClickMenu, , bEdit)
            .InsertSeparator()

            If Not IsDBNull(Row.Item("一部現況")) AndAlso Row.Item("一部現況") > 0 Then
                .AddMenu("部分農地の結合", , AddressOf ClickMenu, , bEdit)
                .AddMenu("部分農地の再分割", , AddressOf ClickMenu, , bEdit)
            Else
                .AddMenu("部分農地の分割", , AddressOf ClickMenu, , bEdit)
            End If
            .AddMenu("共有持分分母", , AddressOf ClickMenu, , bEdit)

            .AddMenu("換地処理", , AddressOf ClickMenu, , bEdit)
            .AddMenu("合筆処理", , AddressOf ClickMenu, , bEdit)
            If SysAD.MapConnection.HasMap Then
                .InsertSeparator()
                .AddMenu("地図を呼ぶ", , AddressOf sub地図を呼ぶ)
            End If
            If SysAD.地図有無 Then
                Dim pTBLX As DataTable = SysAD.DB(s地図情報).GetTableBySqlSelect("SELECT * FROM [D:LotProperty] WHERE [OAZA]=" & Me.大字ID & " AND [Name]='" & Me.地番 & "'")
                If pTBLX.Rows.Count > 0 Then
                    App農地基本台帳.TBL筆情報.MergePlus(pTBLX)
                    .InsertSeparator()
                    .AddMenu("地図表示", , AddressOf ClickMenu)
                End If
            End If
            .InsertSeparator()
            With .AddMenu("転用")
                With .AddMenu("４条転用関連")
                    .AddMenu("４条申請", , AddressOf ClickMenu, , bEdit)
                    .AddMenu("４条申請一時転用", , AddressOf ClickMenu, , bEdit)
                End With
                .AddMenu("５条所有権申請", , AddressOf ClickMenu, , bEdit)
                .AddMenu("５条貸借申請", , AddressOf ClickMenu, , bEdit)
                With .AddMenu("直接転用")
                    .AddMenu("４条転用", , AddressOf ClickMenu, , bEdit)
                    .AddMenu("５条転用", , AddressOf ClickMenu, , bEdit)
                End With
            End With

            .InsertSeparator()
            .AddMenu("農地改良届", , Sub(s, e) C申請データ作成.農地改良届(Me, True), , bEdit)
            .AddMenu("農用地利用計画変更", , AddressOf ClickMenu, , bEdit)
            .AddMenu("農地利用目的変更", , AddressOf ClickMenu, , bEdit)
            .InsertSeparator()
            .AddMenu("あっせん申出", , AddressOf ClickMenu, , bEdit)

            With .AddMenu("非農地関連")
                .AddMenu("非農地証明願(届出)", , AddressOf ClickMenu, , bEdit)
                .AddMenu("非農地設定", , AddressOf ClickMenu, , bEdit)
            End With

            .InsertSeparator()

            With .AddMenu("買受適格申請")
                .AddMenu("耕作目的－公売", , AddressOf ClickMenu, , bEdit)
                .AddMenu("耕作目的－競売", , AddressOf ClickMenu, , bEdit)
                .AddMenu("転用目的－公売", , AddressOf ClickMenu, , bEdit)
                .AddMenu("転用目的－競売", , AddressOf ClickMenu, , bEdit)
            End With

        End With
        Return GetCommonMenu(pMenuf, pMenu, bEdit)
    End Function

    ''' <summary></summary>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Sub sub地図を呼ぶ()
        Dim sB As New System.Text.StringBuilder

        sB.AppendLine("Clear:0")
        sB.AppendLine("LogicMode:2")
        sB.AppendLine("PaintMode:1")
        sB.AppendLine("LotIDP:" & Me.ID.ToString & ",1")

        SysAD.MapConnection.SelectMap(sB.ToString)
    End Sub

    ''' <summary></summary>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Sub List関連申請()
        Dim nID As Long = Me.ID
        Dim St As New System.Text.StringBuilder
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D_申請] WHERE [農地リスト] LIKE '%" & nID & "%'")

        For Each pRow As DataRow In pTBL.Rows
            If InStr(pRow.Item("農地リスト").ToString & ";", "." & nID & ";") Then
                St.Append(IIf(St.Length > 0, ",", "") & pRow.Item("ID"))
            End If
        Next

        If St.Length > 0 Then
            Dim pList As C申請リスト
            Dim sTitle As String = Me.土地所在 & "の関連申請"
            If Not SysAD.page農家世帯.TabPageContainKey(sTitle) Then
                pList = New C申請リスト(SysAD.page農家世帯, sTitle, sTitle)
                pList.Name = sTitle
                SysAD.page農家世帯.中央Tab.AddPage(pList)
            Else
                pList = SysAD.page農家世帯.GetItem(sTitle)
            End If

            Dim sWhere As String = "[ID] IN (" & St.ToString & ")"
            pList.検索開始(sWhere, sWhere)
        Else
            MsgBox("該当する申請がありません", vbInformation, "農地に関連する申請")
        End If
    End Sub

    ''' <summary></summary>
    ''' <param name="sCommand">class：System.String</param>
    ''' <param name="sParams">class：</param>
    ''' <returns>[Private][class:System.Object]</returns>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Overrides Function DoCommand(ByVal sCommand As String, ByVal ParamArray sParams() As String) As Object
        Select Case sCommand
            Case "開く"
                Return Me.OpenDataViewNext(SysAD.page農家世帯.DataViewCollection, ObjectMan)
            Case "部分農地の分割", "部分農地の再分割" : Do一部現況()
            Case "分筆" : Do分筆()
            Case "関連申請" : List関連申請()
            Case "換地処理" : mod農地異動関連.農地換地処理(Me, "換地", sParams)
            Case "合筆処理" : mod農地異動関連.農地換地処理(Me, "合筆", sParams)
            Case "部分農地の結合"
                Me.ClosePage()

                SysAD.page農家世帯.中央Tab.ExistPage("部分結合." & Me.ID, True, GetType(CTabPage部分結合), {Me})
            Case "地図表示"
                Dim pPage As CadastralMaps
                If Not SysAD.page農家世帯.TabPageContainKey("地籍図") Then
                    pPage = New CadastralMaps
                    CType(SysAD.page農家世帯.TabCtrls("TabC"), TabControl).TabPages.Add(pPage)
                Else
                    pPage = SysAD.page農家世帯.GetItem("地籍図")
                    pPage.Active()
                    SysAD.page農家世帯.TabCtrls("TabC").SelectedItem = pPage
                End If

                pPage.DrawLandBoundary(Me)
            Case "４条転用"
                If MsgBox("転用しますか", vbYesNo) = vbYes Then
                    Me.DoCommand("閉じる")
                    Make農地履歴(Me.ID, Now, Now, 土地異動事由.農地法第4条による異動, enum法令.農地法4条, "職権による転用修正")
                    Dim St As String = SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO D_転用農地 SELECT [D:農地Info].* FROM [D:農地Info] WHERE ((([D:農地Info].ID)=" & Me.ID & "));")

                    If InStr(St, "OK") > 0 Then
                        SysAD.DB(sLRDB).ExecuteSQL("DELETE [D:農地Info].ID FROM [D:農地Info] WHERE [D:農地Info].ID=" & Me.ID)
                        App農地基本台帳.TBL農地.Rows.Remove(Me.Row.Body)
                    Else
                        MsgBox("データの転送に失敗しました", MsgBoxStyle.Critical)
                    End If
                End If
            Case "５条転用"
                If MsgBox("転用しますか", vbYesNo) = vbYes Then
                    Me.DoCommand("閉じる")
                    Make農地履歴(Me.ID, Now, Now, 土地異動事由.農地法５条による転用, enum法令.農地法5条所有権, "職権による転用修正")
                    Dim St As String = SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO D_転用農地 SELECT [D:農地Info].* FROM [D:農地Info] WHERE ((([D:農地Info].ID)=" & Me.ID & "));")

                    If InStr(St, "OK") > 0 Then
                        SysAD.DB(sLRDB).ExecuteSQL("DELETE [D:農地Info].ID FROM [D:農地Info] WHERE [D:農地Info].ID=" & Me.ID)
                        App農地基本台帳.TBL農地.Rows.Remove(Me.Row.Body)
                    Else
                        MsgBox("データの転送に失敗しました", MsgBoxStyle.Critical)
                    End If
                End If
            Case "非農地設定"
                If System.Windows.MessageBox.Show("非農地設定しますか", "非農地設定", Windows.MessageBoxButton.YesNo) = Windows.MessageBoxResult.Yes Then
                    Me.DoCommand("閉じる")

                    With New HimTools2012.PropertyGridDialog(New C異動日入力(), "異動日入力", "異動日を入力してください。")
                        If .ShowDialog = DialogResult.OK Then
                            Dim dt異動日 As DateTime = CType(.ResultProperty, C異動日入力).異動日

                            Make農地履歴(Me.ID, Now, dt異動日, 100004, enum法令.非農地証明願, String.Format("非農地判断 [{0}]", 和暦Format(dt異動日)))
                            SysAD.DB(sLRDB).ExecuteSQL("DELETE FROM [D_転用農地] WHERE [ID]={0}", Me.ID)

                            Dim St As String = SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO D_転用農地 SELECT [D:農地Info].* FROM [D:農地Info] WHERE ((([D:農地Info].ID)=" & Me.ID & "));")

                            If St = "" OrElse InStr(St, "OK") > 0 Then
                                SysAD.DB(sLRDB).ExecuteSQL("DELETE [D:農地Info].ID FROM [D:農地Info] WHERE [D:農地Info].ID=" & Me.ID)
                                App農地基本台帳.TBL農地.Rows.Remove(Me.Row.Body)
                            Else
                                MsgBox("データの転送に失敗しました", MsgBoxStyle.Critical)
                            End If
                        End If
                    End With

                End If
            Case "データを削除"
                If (sParams.Length > 0 AndAlso sParams(0) = "NQ") OrElse MsgBox("本当に削除しますか", vbOKCancel) = MsgBoxResult.Ok Then
                    農地削除(App農地基本台帳.TBL農地.GetDataView("[ID]=" & Me.ID, "[ID]=" & Me.ID, "").ToTable, 土地異動事由.農地の削除, C農地削除.enum転送先.削除農地, "職権による農地の削除", Now.Date)
                End If
            Case "管理世帯を呼ぶ" : Open世帯(Val(mvarRow.Item("管理世帯ID").ToString), "指定された管理番号は見つかりませんでした")
            Case "管理者を呼ぶ" : Open個人(mvarRow.Item("管理者ID"), "指定された管理者が見つかりませんでした。")
            Case "所有者を呼ぶ" : Open個人(mvarRow.Item("所有者ID"), "指定された所有者が見つかりませんでした。")
            Case "経由法人を呼ぶ" : Open個人(mvarRow.Item("経由農業生産法人ID"), "指定された経由法人が見つかりませんでした。")
            Case "借受者を呼ぶ" : Open個人(mvarRow.Item("借受人ID"), "指定された借受者が見つかりませんでした。")
            Case "関連申請" : List関連申請()
            Case "同所有者の所有地検索" : SysAD.page農家世帯.農地リスト.検索開始("[所有者ID]=" & Me.Row.Body.Item("所有者ID"), "[所有者ID]=" & Me.Row.Body.Item("所有者ID"))

                '/******************************* mod申請データ作成処理 ***************************************/
            Case "４条申請" : Return New C申請データ作成("転用農地法4条の受付", Me.Key.KeyValue, Nothing)
            Case "４条申請一時転用" : Return New C申請データ作成("4条一時転用の申請受付", Me.Key.KeyValue, Nothing)
            Case "５条所有権申請" : Return New C申請データ作成("転用を伴う所有権移転(5条)の申請受付", Me.Key.KeyValue, Nothing)
            Case "議案書付解約" : C申請データ作成.解約申請(Me.Key.KeyValue, True)
            Case "貸借の再設定" : C申請データ作成.経営基盤法利用権設定(Me.GetProperty("管理者"), Me.Key.KeyValue, Me.GetProperty("Obj受人"), "再設定")
            Case "農地利用目的変更" : C申請データ作成.農地利用目的変更(Me, True)
            Case "耕作目的－公売", "耕作目的－競売", "転用目的－公売", "転用目的－競売" : C申請データ作成.買受適格(Me, sCommand, True)
            Case "非農地証明願(届出)" : C申請データ作成.非農地証明願(Me)
            Case "農用地利用計画変更" : C申請データ作成.農用地利用計画変更(Me, True)
            Case "あっせん申出" : C申請データ作成.あっせん申出渡(Me)
            Case "期間満了の終了" : sub農地期間満了の終了(Me.ID)
            Case "関連申請の延長" : sub関連申請の延長(Me.ID)
            Case "選択農地のみ延長" : sub貸借期間の延長(Me.ID)
            Case "エクスポート(XML)"
                With New SaveFileDialog
                    .Filter = "データファイル(*.XML)|*.XML"
                    If .ShowDialog = DialogResult.OK Then
                        Dim pSaveTable As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE [ID]=" & Me.ID)
                        pSaveTable.TableName = "農地Info"
                        pSaveTable.WriteXml(.FileName, XmlWriteMode.WriteSchema)
                        MsgBox("保存しました")
                    End If
                End With
            Case "履歴追加", "履歴の追加" : 農地履歴手動追加(Me)
            Case "議案書無し合意解約", "職権解約" : 解約()
            Case "中間管理機構へ農地の返還" : C申請データ作成.返還申請(Me.Key.KeyValue, True)
            Case "閲覧用農地台帳印刷" : mod農地基本台帳.農地台帳印刷(Me.Key.KeyValue, ExcelViewMode.Preview, "閲覧用")
            Case "農地台帳要約書印刷" : mod農地基本台帳.農地台帳印刷(Me.Key.KeyValue, ExcelViewMode.Preview, "要約書")
            Case "期間設定"
                Dim pDT As Object = mvarRow.Item("小作開始年月日")
                If pDT IsNot Nothing AndAlso IsDate(pDT) Then
                    With New dlgInputBWDate(pDT)
                        If .ShowDialog() = DialogResult.OK Then
                            ValueChange("小作終了年月日", .RestultDate)
                        End If
                    End With
                Else
                    MsgBox("貸借開始年月日を入力してください", vbCritical)
                End If
            Case "転貸期間設定"
                Dim pDT As Object = mvarRow.Item("転貸始期年月日")
                If pDT IsNot Nothing AndAlso IsDate(pDT) Then
                    With New dlgInputBWDate(pDT)
                        If .ShowDialog() = DialogResult.OK Then
                            ValueChange("転貸終期年月日", .RestultDate)
                        End If
                    End With
                Else
                    MsgBox("転貸開始年月日を入力してください", vbCritical)
                End If
            Case "転用期間設定"
                Dim pDT As Object = mvarRow.Item("転用始期年月日")
                If pDT IsNot Nothing AndAlso IsDate(pDT) Then
                    With New dlgInputBWDate(pDT)
                        If .ShowDialog() = DialogResult.OK Then
                            ValueChange("転用終期年月日", .RestultDate)
                        End If
                    End With
                Else
                    MsgBox("転用開始年月日を入力してください", vbCritical)
                End If
            Case "利用配分期間設定"
                Dim pDT As Object = mvarRow.Item("利用配分計画始期日")
                If pDT IsNot Nothing AndAlso IsDate(pDT) Then
                    With New dlgInputBWDate(pDT)
                        If .ShowDialog() = DialogResult.OK Then
                            ValueChange("利用配分計画終期日", .RestultDate)
                        End If
                    End With
                Else
                    MsgBox("利用配分開始年月日を入力してください", vbCritical)
                End If
            Case "所有者", "所有世帯", "転用"
            Case "共有持分の分割"
                '    Dim pFrm As New frm共有分割
                '    If pFrm.分割(DVProperty.ID) Then
                '        CDataviewSK_DoCommand2("閉じる")
                '    End If
            Case "更新" : Me.SaveMyself()
            Case Else
                Return MyBase.DoCommand(sCommand, sParams)
        End Select
        Return ""
    End Function

    ''' <summary></summary>
    ''' <param name="sSourceList">class：System.String</param>
    ''' <param name="sOption">class：System.String</param>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Overrides Sub DropDown(ByVal sSourceList As String, Optional ByVal sOption As String = "")
        Select Case sOption
            Case "所有世帯ID", "所有者ID", "所有者氏名"
                Select Case GetKeyHead(sSourceList)
                    Case "農家"
                        If MsgBox("所有世帯を変更しますか", vbYesNo) = vbYes Then
                            Dim p農家 As CObj農家 = ObjectMan.GetObject(sSourceList)
                            ValueChange("所有世帯ID", p農家.ID)
                            ValueChange("所有者ID", p農家.世帯主ID)
                        End If
                    Case "個人"
                        If MsgBox("所有者を変更しますか", vbYesNo) = vbYes Then
                            Dim p個人 As CObj個人 = ObjectMan.GetObject(sSourceList)
                            If p個人 IsNot Nothing Then
                                ValueChange("所有世帯ID", p個人.世帯ID)
                                ValueChange("所有者ID", p個人.ID)
                            End If
                        End If
                End Select
            Case "管理世帯ID", "管理者ID"
                Select Case GetKeyHead(sSourceList)
                    Case "農家"
                        If MsgBox("管理世帯を変更しますか", vbYesNo) = vbYes Then
                            Dim p農家 As CObj農家 = ObjectMan.GetObject(sSourceList)
                            ValueChange("管理世帯ID", p農家.ID)
                            ValueChange("管理者ID", p農家.世帯主ID)
                        End If
                    Case "個人"
                        If MsgBox("管理者を変更しますか", vbYesNo) = vbYes Then
                            Dim p個人 As CObj個人 = ObjectMan.GetObject(sSourceList)
                            If p個人 IsNot Nothing Then
                                ValueChange("管理世帯ID", p個人.世帯ID)
                                ValueChange("管理者ID", p個人.ID)
                            End If
                        End If
                End Select
            Case "特定作業者ID"
                Select Case GetKeyHead(sSourceList)
                    Case "個人"
                        Dim St対象項目 As String = ""
                        Select Case sOption
                            Case "特定作業者ID"
                                St対象項目 = "特定作業者"
                        End Select
                        If MsgBox(St対象項目 & "を変更しますか", vbYesNo) = vbYes Then
                            Dim p個人 As CObj個人 = ObjectMan.GetObject(sSourceList)
                            If p個人 IsNot Nothing Then
                                ValueChange(St対象項目 & "ID", p個人.ID)
                                ValueChange(St対象項目 & "名", p個人.氏名)
                                ValueChange(St対象項目 & "住所", p個人.住所)
                                Me.SaveMyself()
                            End If
                        End If
                End Select

            Case "借受人氏名", "TX小作者", "TX小作者ID", "借受人ID"
                Select Case GetKeyHead(sSourceList)
                    Case "農家"
                        If MsgBox("借受世帯を変更しますか", vbYesNo) = vbYes Then
                            Dim p農家 As CObj農家 = ObjectMan.GetObject(sSourceList)
                            ValueChange("借受世帯ID", p農家.ID)
                            ValueChange("借受人ID", p農家.世帯主ID)
                        End If
                    Case "個人"
                        If MsgBox("借受者を変更しますか", vbYesNo) = vbYes Then
                            Dim p個人 As CObj個人 = ObjectMan.GetObject(sSourceList)
                            ValueChange("借受世帯ID", p個人.世帯ID)
                            ValueChange("借受人ID", p個人.ID)
                        End If
                End Select
            Case "登記名義人ID"
                Select Case GetKeyHead(sSourceList)
                    Case "個人"
                        If MsgBox("登記名義人を変更しますか", vbYesNo) = vbYes Then
                            Dim p個人 As CObj個人 = ObjectMan.GetObject(sSourceList)
                            ValueChange("登記名義人ID", p個人.ID)
                        End If
                End Select
            Case "経由農業生産法人ID"
                Select Case GetKeyHead(sSourceList)
                    Case "個人"
                        If MsgBox("経由者を変更しますか", vbYesNo) = vbYes Then
                            Dim p個人 As CObj個人 = ObjectMan.GetObject(sSourceList)
                            ValueChange("経由農業生産法人ID", p個人.ID)
                        End If
                End Select
            Case "相続届出者ID"
                Select Case GetKeyHead(sSourceList)
                    Case "個人"
                        If MsgBox("相続届出者を変更しますか", vbYesNo) = vbYes Then
                            Dim p個人 As CObj個人 = ObjectMan.GetObject(sSourceList)
                            ValueChange("相続届出者ID", p個人.ID)
                        End If
                End Select
            Case "推測耕作者ID"
                Select Case GetKeyHead(sSourceList)
                    Case "個人"
                        If MsgBox("耕作しているであろう者を変更しますか", vbYesNo) = vbYes Then
                            Dim p個人 As CObj個人 = ObjectMan.GetObject(sSourceList)
                            ValueChange("推測耕作者ID", p個人.ID)
                        End If
                End Select
            Case ""
                Select Case GetKeyHead(sSourceList)
                    Case "農地"
                        Dim sSelect As String = "換地前→換地後の異動処理;同一地の結合"
                        Dim St As String = OptionSelect(sSelect, "設定する内容を選択してください。")
                        Select Case St
                            Case "同一地の結合"
                                mod農地異動関連.同一地の結合(sSourceList, Me)
                            Case "換地前→換地後の異動処理"
                                mod農地異動関連.換地前後関連付け処理(sSourceList, Me)
                            Case ""
                            Case Else
                                MsgBox("現在この機能は使われていません", MsgBoxStyle.Critical)
                        End Select
                End Select
            Case Else
                CasePrint(sOption)
        End Select
    End Sub

    ''' <summary></summary>
    ''' <param name="sField">class：System.String</param>
    ''' <param name="pValue">class：System.Object</param>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Overrides Sub ValueChange(ByVal sField As String, ByVal pValue As Object)
        App農地基本台帳.TBL農地.AddUpdateListwithDataViewPage(Me, sField, pValue)
    End Sub

    ''' <summary></summary>
    ''' <param name="nID">class：System.Int64</param>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Private Sub sub農地期間満了の終了(ByVal nID As Long)
        App農地基本台帳.TBL農地.MergePlus(SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE [ID]=" & nID))
        Dim pRow As DataRow = App農地基本台帳.TBL農地.FindRowByID(nID)

        If pRow.Item("自小作別") > 0 AndAlso IsDate(pRow.Item("小作終了年月日")) Then
            If pRow.Item("小作終了年月日") <= Now() Then
                If MsgBox("貸借を期間満了で終了しますか？", vbYesNo) = vbYes Then
                    SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO D_土地履歴 ( LID, 内容, 異動日,更新日, 入力日, 関係者A, 関係者B, 異動事由 ) SELECT [D:農地Info_1].ID, [氏名] & 'への利用権設定[' & Format$([D:農地Info].[小作開始年月日],'gggee\年mm\月dd\日') & ']～[' & Format$([D:農地Info].[小作終了年月日],'gggee\年mm\月dd\日') & ']を期間満了で終了。' AS 式1, Date() AS 式2, Date() AS 式3, Date() AS 式4, [D:農地Info_1].所有者ID, [D:農地Info].借受人ID, 10201 AS 式5 FROM ([D:農地Info] INNER JOIN [D:農地Info] AS [D:農地Info_1] ON ([D:農地Info].小作終了年月日 = [D:農地Info_1].小作終了年月日) AND ([D:農地Info].小作開始年月日 = [D:農地Info_1].小作開始年月日) AND ([D:農地Info].小作地適用法 = [D:農地Info_1].小作地適用法) AND ([D:農地Info].借受人ID = [D:農地Info_1].借受人ID) AND ([D:農地Info].自小作別 = [D:農地Info_1].自小作別)) INNER JOIN [D:個人Info] ON [D:農地Info_1].借受人ID = [D:個人Info].ID WHERE ((([D:農地Info].ID)=" & nID & ") AND (([D:農地Info].自小作別)<>0));")
                    SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] INNER JOIN [D:農地Info] AS [D:農地Info_1] ON ([D:農地Info].小作終了年月日 = [D:農地Info_1].小作終了年月日) AND ([D:農地Info].小作開始年月日 = [D:農地Info_1].小作開始年月日) AND ([D:農地Info].小作地適用法 = [D:農地Info_1].小作地適用法) AND ([D:農地Info].借受人ID = [D:農地Info_1].借受人ID) AND ([D:農地Info].自小作別 = [D:農地Info_1].自小作別) SET [D:農地Info_1].自小作別 = 0 WHERE ((([D:農地Info].ID)=" & nID & ") AND (([D:農地Info].自小作別)<>0));")

                    App農地基本台帳.TBL農地.MergePlus(SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE [ID]=" & nID))
                End If
            Else
                If MsgBox("期限は過ぎていませんが、貸借を期間満了で終了しますか？", vbYesNo) = vbYes Then
                    SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO D_土地履歴 ( LID, 内容, 異動日,更新日, 入力日, 関係者A, 関係者B, 異動事由 ) SELECT [D:農地Info_1].ID, [氏名] & 'への利用権設定[' & Format$([D:農地Info].[小作開始年月日],'gggee\年mm\月dd\日') & ']～[' & Format$([D:農地Info].[小作終了年月日],'gggee\年mm\月dd\日') & ']を期間満了で終了。' AS 式1, Date() AS 式2, Date() AS 式3, Date() AS 式4, [D:農地Info_1].所有者ID, [D:農地Info].借受人ID, 10201 AS 式5 FROM ([D:農地Info] INNER JOIN [D:農地Info] AS [D:農地Info_1] ON ([D:農地Info].小作終了年月日 = [D:農地Info_1].小作終了年月日) AND ([D:農地Info].小作開始年月日 = [D:農地Info_1].小作開始年月日) AND ([D:農地Info].小作地適用法 = [D:農地Info_1].小作地適用法) AND ([D:農地Info].借受人ID = [D:農地Info_1].借受人ID) AND ([D:農地Info].自小作別 = [D:農地Info_1].自小作別)) INNER JOIN [D:個人Info] ON [D:農地Info_1].借受人ID = [D:個人Info].ID WHERE ((([D:農地Info].ID)=" & nID & ") AND (([D:農地Info].自小作別)<>0));")
                    SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] INNER JOIN [D:農地Info] AS [D:農地Info_1] ON ([D:農地Info].小作終了年月日 = [D:農地Info_1].小作終了年月日) AND ([D:農地Info].小作開始年月日 = [D:農地Info_1].小作開始年月日) AND ([D:農地Info].小作地適用法 = [D:農地Info_1].小作地適用法) AND ([D:農地Info].借受人ID = [D:農地Info_1].借受人ID) AND ([D:農地Info].自小作別 = [D:農地Info_1].自小作別) SET [D:農地Info_1].自小作別 = 0 WHERE ((([D:農地Info].ID)=" & nID & ") AND (([D:農地Info].自小作別)<>0));")

                    App農地基本台帳.TBL農地.MergePlus(SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE [ID]=" & nID))
                End If

            End If
        Else

        End If

    End Sub

    Private Sub sub関連申請の延長(ByVal nID As Long)
        MsgBox("現在実装中です。「選択農地のみ延長」より変更をお願い致します。")
    End Sub

    Private Sub sub貸借期間の延長(ByVal nID As Long)
        Dim p農地 As Object = Nothing
        p農地 = ObjectMan.GetObject("農地." & nID)

        If p農地 Is Nothing OrElse p農地.Row Is Nothing Then
        Else
            If p農地.GetIntegerValue("自小作別") > 0 AndAlso IsDate(p農地.GetDateValue("小作終了年月日")) Then
                If MsgBox("貸借期間を延長しますか？", vbYesNo) = vbYes Then
                    Dim sResult As DateTime = InputBox("延長後の終了年月日を入力してください", "終了年月日", p農地.GetDateValue("小作終了年月日"))
                    If IsDate(sResult) Then
                        Make農地履歴(nID, Now, Now, 100005, enum法令.職権異動, "貸借期間「" & 和暦Format(p農地.GetDateValue("小作開始年月日")) & "～" & 和暦Format(p農地.GetDateValue("小作終了年月日")) & "」を「" & 和暦Format(p農地.GetDateValue("小作開始年月日")) & "～" & 和暦Format(sResult) & "」に変更しました。", , 0)
                    Else
                        MsgBox("指定した日付は存在しません。再入力をお願いします。")
                        sub貸借期間の延長(nID)
                    End If

                    p農地.ValueChange("小作終了年月日", sResult)
                    p農地.SaveMyself()
                End If
            End If
        End If

    End Sub

    ''' <summary></summary>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Private Sub 解約()
        Dim pDate As Date = Now
        Dim St As New System.Text.StringBuilder
        With New HimTools2012.PropertyGridDialog(New C異動日入力(), "解約の実行", "履歴のみで受付情報は記録しません")
            If .ShowDialog = DialogResult.OK Then
                With CType(.ResultProperty, C異動日入力)
                    Dim p出し手 As CObj個人 = ObjectMan.GetObject("個人." & Me.GetDecimalValue("所有者ID"))
                    Dim p受け手 As CObj個人 = ObjectMan.GetObject("個人." & Me.GetDecimalValue("借受人ID"))
                    St.Append(p出し手.氏名 & "→" & p受け手.氏名 & "の貸借の解約・終了")

                    If Not IsDBNull(Me.Row.Body.Item("小作開始年月日")) Then St.Append(vbCrLf & " 設定期間:[" & 和暦Format(Me.GetDateValue("小作開始年月日")) & "]") Else St.Append(vbCrLf & " 設定期間:[??/??/??]")
                    If Not IsDBNull(Me.Row.Body.Item("小作終了年月日")) Then St.Append("～" & "[" & 和暦Format(Me.GetDateValue("小作終了年月日")) & "]") Else St.Append("～" & "[??/??/??]")

                    Me.DoCommand("閉じる")
                    Me.SetIntegerValue("自小作別", 0)

                    SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [自小作別]=0 WHERE [ID]={0}", Me.ID)
                    SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO D_土地履歴 ( LID, 更新日, 異動日, 入力日, 異動事由, 行政区, 内容,関係者A,関係者B) VALUES({0},{1},{2},{3}, 10210,0, '{4}',{5},{6});", Me.ID, HimTools2012.StringF.Toリテラル日付(Now.Date, True), HimTools2012.StringF.Toリテラル日付(.異動日.Date, True), HimTools2012.StringF.Toリテラル日付(Now.Date, True), St.ToString, p出し手.ID, p受け手.ID)
                End With
            End If
        End With
    End Sub

    ''' <summary></summary>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Sub Do分筆()
        Me.DoCommand("閉じる")
        SysAD.page農家世帯.中央Tab.ExistPage("分筆処理" & Me.ID, True, GetType(CTabPage分筆処理), {Me})
    End Sub
    ''' <summary></summary>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Sub Do一部現況()
        Me.DoCommand("閉じる")
        SysAD.page農家世帯.中央Tab.ExistPage("分割処理." & Me.ID, True, GetType(CTabPage分割処理), {Me})
    End Sub

    ''' <summary></summary>
    ''' <param name="pDB">class：HimTools2012.TargetSystem.CDataViewCollection</param>
    ''' <param name="InterfaceName">class：System.String</param>
    ''' <returns>[Private][class:System.Boolean]</returns>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Overrides Function InitDataViewNext(ByRef pDB As HimTools2012.TargetSystem.CDataViewCollection, Optional ByVal InterfaceName As String = "") As Boolean
        If Me.DataViewPage Is Nothing Then
            Me.DataViewPage = New DataViewNext農地(Me)
        End If
        Return True
    End Function

    ''' <summary></summary>
    ''' <param name="sKey">class：System.String</param>
    ''' <param name="sOption">class：System.String</param>
    ''' <returns>[Private][class:System.Boolean]</returns>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/29 15:53]</remarks>
    Public Overloads Overrides Function CanDropKeyHead(ByVal sKey As String, ByVal sOption As String) As Boolean
        Select Case GetKeyHead(sKey) & "-" & sOption
            Case "世帯-所有世帯ID"
                Return True
            Case "個人-借受人ID", "個人-借受人氏名", "個人-所有者ID", "個人-所有者氏名"
                Return True
            Case "個人-特定作業者ID"
                Return True
            Case "農地-借受人氏名"
                Return False
            Case "農地-"
                Return True
            Case "個人-"
                Return True
            Case "個人-経由農業生産法人ID"
                Return True
            Case "個人-管理者ID"
                Return True
            Case "個人-登記名義人ID"
                Return True
            Case "個人-相続届出者ID"
                Return True
            Case "個人-推測耕作者ID"
                Return True
            Case "-経由農業生産法人ID", "-経由農業生産法人名"
                Return False
            Case Else
                CasePrint(GetKeyHead(sKey) & "-" & sOption)
                Return False
        End Select
    End Function
End Class

