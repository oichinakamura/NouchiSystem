
Imports System.ComponentModel
Imports HimTools2012.controls.PropertyGridSupport
Imports HimTools2012.TypeConverterCustom

<TypeConverter(GetType(PropertyOrderConverter))>
Public Class CInput年度設定
    Inherits HimTools2012.InputSupport.CInputSupport

    Public Sub New()
        MyBase.New(農地基本台帳.App農地基本台帳.DataMaster.Body)

    End Sub
    <Category("年度")>
    Public Property 年度 As 平成
End Class

<TypeConverter(GetType(PropertyOrderConverter))>
Public Class C申請入力支援
    Inherits HimTools2012.InputSupport.CInputSupport

    Public Sub New()
        MyBase.New(農地基本台帳.App農地基本台帳.DataMaster.Body)
    End Sub

    <PropertyOrderAttribute(3)> <Category("02 申請情報")> <DefaultValue("")> <Description("受付年月日")>
    Public Property 受付年月日 As DateTime
    <PropertyOrderAttribute(4)> <Category("02 申請情報")> <DefaultValue("")> <Description("受付番号を月毎に")>
    Public Property 受付番号 As Integer
    <PropertyOrderAttribute(5)> <Category("02 申請情報")> <DefaultValue("")> <Description("受付番号を")>
    Public Property 受付通年番号 As Integer

    Public Function Get受付番号MAX(ByVal n法令 As Integer) As Integer
        Try
            Dim n申請締切 As Integer = Val(SysAD.DB(sLRDB).DBProperty("申請締切", "15"))
            Dim s期間 As String = SysAD.DB(sLRDB).DBProperty("申請締切", "15")
            Dim nYear As Integer = Now.Year
            Dim nMonth As Integer = Now.Month
            Dim nDay As Integer = n申請締切

            Do Until IsDate(String.Format("{0}/{1}/{2}", nMonth, nDay, nYear))
                If nDay > 28 Then
                    nDay -= 1
                ElseIf nDay < 1 Then

                    nDay = 31
                End If
            Loop

            Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Max([D_申請].受付番号) AS 最大番号 FROM [D_申請] WHERE ([D_申請].法令 In ( " & n法令 & " )) AND (DatePart('yyyy',[受付年月日])=DatePart('yyyy',Date()))")
            If pTBL.Rows.Count = 1 Then
                Return Val(pTBL.Rows(0).Item("最大番号").ToString) + 1
            Else
                Return 1
            End If
        Catch ex As Exception
            Return 1
        End Try
    End Function
    Public Function Get受付通年番号MAX(ByVal n法令 As Integer) As Integer
        Dim pTable As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Max([D_申請].通年受付番号) AS 最大番号 FROM [D_申請] WHERE ([D_申請].法令 In ( " & n法令 & " )) AND (DatePart('yyyy',[受付年月日])=DatePart('yyyy',Date()))")
        If pTable.Rows.Count = 1 Then
            Return Val(pTable.Rows(0).Item("最大番号").ToString) + 1
        Else
            Return 1
        End If
    End Function
End Class

'------------------------ 申請人１人
Public Class C申請入力一人称
    Inherits C申請入力支援


    Private mvar申請者 As CObj個人

    Public Sub New(ByRef p申請者 As CObj個人, ByVal n法令 As enum法令)
        MyBase.New()

        mvar申請者 = p申請者
        受付年月日 = Now.Date

        受付番号 = Get受付番号MAX(n法令)
        受付通年番号 = Get受付通年番号MAX(n法令)
    End Sub


    <PropertyOrderAttribute(2)> <Category("01 基本情報")> <DefaultValue("")> <Description("申請者")>
    Public ReadOnly Property 申請者() As String
        Get
            Return mvar申請者.ToString
        End Get
    End Property

End Class
'------------------------ 申請人２人
<TypeConverter(GetType(PropertyOrderConverter))>
Public Class C申請入力二人称
    Inherits C申請入力支援

    Private mvar出し手 As CObj個人
    Private mvar受け手 As CObj個人

    Public Sub New(ByRef p所有者 As HimTools2012.TargetSystem.CTargetObjWithView, ByRef p受け手 As HimTools2012.TargetSystem.CTargetObjWithView, ByVal n法令 As enum法令, Optional p農地() As CObj農地 = Nothing)
        MyBase.New()

        Select Case p所有者.Key.DataClass
            Case "農家" : mvar出し手 = p所有者.GetProperty("世帯主")
            Case "個人" : mvar出し手 = p所有者
        End Select
        If p受け手 IsNot Nothing Then
            Select Case p受け手.Key.DataClass
                Case "農家" : mvar受け手 = p受け手.GetProperty("世帯主")
                Case "個人" : mvar受け手 = p受け手
            End Select
            If TypeOf p受け手 Is CObj農家 Then
                mvar受け手 = CType(p受け手, CObj農家).世帯主
            ElseIf TypeOf p受け手 Is CObj個人 Then
                mvar受け手 = p受け手
            End If
        End If


        受付年月日 = Now.Date

        受付番号 = Get受付番号MAX(n法令)
        受付通年番号 = Get受付通年番号MAX(n法令)
    End Sub

    <PropertyOrderAttribute(1)> <Category("01 基本情報")> <DefaultValue("")> <Description("譲渡・貸し手")>
    Public ReadOnly Property 譲渡_貸し手() As String
        Get
            If mvar出し手 IsNot Nothing Then
                Return mvar出し手.ToString
            Else
                Return ""
            End If
        End Get
    End Property
    <PropertyOrderAttribute(2)> <Category("01 基本情報")> <DefaultValue("")> <Description("譲受・借り手")>
    Public ReadOnly Property 譲受_借り手() As String
        Get
            If mvar受け手 IsNot Nothing Then
                Return mvar受け手.ToString
            Else
                Return ""
            End If
        End Get
    End Property
End Class


<TypeConverter(GetType(PropertyOrderConverter))>
Public Class C借受者変更
    Inherits HimTools2012.InputSupport.CInputSupport

    Private mvar出し手 As CObj個人
    Private mvar前借受者 As CObj個人
    Private mvar後借受者 As CObj個人

    Public Sub New(ByRef p所有者 As HimTools2012.TargetSystem.CTargetObjWithView,
                   ByRef p前借受者 As HimTools2012.TargetSystem.CTargetObjWithView,
                   ByRef p新借受者 As HimTools2012.TargetSystem.CTargetObjWithView,
                   ByVal n法令 As enum法令, Optional p農地() As CObj農地 = Nothing)
        MyBase.New(農地基本台帳.App農地基本台帳.DataMaster.Body)

        Select Case p所有者.Key.DataClass
            Case "農家" : mvar出し手 = p所有者.GetProperty("世帯主")
            Case "個人" : mvar出し手 = p所有者
        End Select
        If p前借受者 IsNot Nothing Then
            Select Case p前借受者.Key.DataClass
                Case "農家" : mvar前借受者 = p前借受者.GetProperty("世帯主")
                Case "個人" : mvar前借受者 = p前借受者
            End Select
        End If
        If p新借受者 IsNot Nothing Then
            Select Case p新借受者.Key.DataClass
                Case "農家" : mvar後借受者 = p新借受者.GetProperty("世帯主")
                Case "個人" : mvar後借受者 = p新借受者
            End Select
        End If

    End Sub
    Private mvar変更を受けた日 As DateTime
    <PropertyOrderAttribute(1)> <Category("01 基本情報")> <DefaultValue("")> <Description("変更を受けた日")>
    Public Property 変更を受けた日() As DateTime
        Get
            Return mvar変更を受けた日
        End Get
        Set(value As DateTime)
            mvar変更を受けた日 = value
        End Set
    End Property


    <PropertyOrderAttribute(2)> <Category("01 基本情報")> <DefaultValue("")> <Description("所有者")>
    Public ReadOnly Property 所有者() As String
        Get
            If mvar出し手 IsNot Nothing Then
                Return mvar出し手.ToString
            Else
                Return ""
            End If
        End Get
    End Property
    <PropertyOrderAttribute(3)> <Category("01 基本情報")> <DefaultValue("")> <Description("前の借受者")>
    Public ReadOnly Property 前の借受者() As String
        Get
            If mvar前借受者 IsNot Nothing Then
                Return mvar前借受者.ToString
            Else
                Return ""
            End If
        End Get
    End Property
    <PropertyOrderAttribute(4)> <Category("01 基本情報")> <DefaultValue("")> <Description("新しい借受者")>
    Public ReadOnly Property 新しい借受者() As String
        Get
            If mvar後借受者 IsNot Nothing Then
                Return mvar後借受者.ToString
            Else
                Return ""
            End If
        End Get
    End Property
    Private mvar前貸借終了日 As DateTime
    Private mvar新貸借開始日 As DateTime
    Private mvar新貸借終了日 As DateTime
    Private mvar小作形態 As enum小作形態
    Private mvar小作料 As Integer

    <PropertyOrderAttribute(5)> <Category("02 貸借の終了")> <DefaultValue("")> <Description("前の貸借の終了日")>
    Public Property 前の貸借の終了日() As DateTime
        Get
            Return mvar前貸借終了日
        End Get
        Set(value As DateTime)
            mvar前貸借終了日 = value
        End Set
    End Property
    <PropertyOrderAttribute(6)> <Category("03 新しい貸借の条件")> <DefaultValue("")> <Description("新しい貸借の開始日")>
    Public Property 新しい貸借の開始日() As DateTime
        Get
            Return mvar新貸借開始日
        End Get
        Set(value As DateTime)
            mvar新貸借開始日 = value
        End Set
    End Property
    <PropertyOrderAttribute(7)> <Category("03 新しい貸借の条件")> <DefaultValue("")> <Description("新しい貸借の終了日")>
    Public Property 新しい貸借の終了日() As DateTime
        Get
            Return mvar新貸借終了日
        End Get
        Set(value As DateTime)
            mvar新貸借終了日 = value
        End Set
    End Property
    <PropertyOrderAttribute(8)> <Category("03 新しい貸借の条件")> <DefaultValue("")> <Description("新しい貸借の形態")>
    Public Property 新しい貸借の形態() As enum小作形態
        Get
            Return mvar小作形態
        End Get
        Set(value As enum小作形態)
            mvar小作形態 = value
        End Set
    End Property
    '<PropertyOrderAttribute(9)> <Category("03 新しい貸借の条件")> <DefaultValue("")> <Description("小作料")>
    'Public Property 小作料() As Decimal
    '    Get
    '        Return mvar小作料
    '    End Get
    '    Set(value As Decimal)
    '        mvar小作料 = value
    '    End Set
    'End Property
    '<PropertyOrderAttribute(10)> <Category("03 新しい貸借の条件")> <DefaultValue("")> <Description("小作料単位")>
    'Public Property 小作料単位() As Integer
    '    Get
    '        Return mvar小作料
    '    End Get
    '    Set(value As Integer)
    '        mvar小作料 = value
    '    End Set
    'End Property
    <PropertyOrderAttribute(11)> <Category("03 新しい貸借の条件")> <DefaultValue("")> <Description("新しい10a当たりの小作料")>
    Public Property 新しい10a当たりの小作料() As Integer
        Get
            Return mvar小作料
        End Get
        Set(value As Integer)
            mvar小作料 = value
        End Set
    End Property

    Public Overrides Function DataCompleate() As Boolean
        With Me
            Dim sMessage As New System.Text.StringBuilder
            If IsDBNull(.変更を受けた日) OrElse Year(.変更を受けた日) < 2000 Then sMessage.AppendLine("「変更を受けた日」が未入力か不正です。")
            If IsDBNull(.前の貸借の終了日) OrElse Year(.前の貸借の終了日) < 2000 Then sMessage.AppendLine("「前の貸借の終了日」が未入力か不正です。")
            If IsDBNull(.新しい貸借の開始日) OrElse Year(.新しい貸借の開始日) < 2000 Then sMessage.AppendLine("「新しい貸借の開始日」が未入力か不正です。")
            If IsDBNull(.新しい貸借の終了日) OrElse Year(.新しい貸借の終了日) < 2000 Then sMessage.AppendLine("「新しい貸借の終了日」が未入力か不正です。")

            Select Case .新しい貸借の形態
                Case enum小作形態.使用貸借
                Case enum小作形態.使用賃借
                Case enum小作形態.賃貸借
                    If IsDBNull(.新しい10a当たりの小作料) OrElse .新しい10a当たりの小作料 = 0 Then sMessage.AppendLine("「新しい10a当たりの小作料」が未入力か不正です。")
                Case Else
                    sMessage.AppendLine("「新しい貸借の形態」が未入力か不正です。")
            End Select
            If sMessage.Length = 0 Then
                Return True

            Else
                MsgBox(sMessage.ToString)
                Return False
            End If
        End With
    End Function


End Class

Public Class C許可入力支援
    Inherits HimTools2012.InputSupport.CInputSupport
    Dim mvar申請 As CObj申請

    Public Sub New(ByRef p申請 As CObj申請)
        MyBase.New(農地基本台帳.App農地基本台帳.DataMaster.Body)
        mvar申請 = p申請
        Dim p許可日 As Object = p申請.GetDateValue("許可年月日")
        If p許可日 IsNot Nothing AndAlso p許可日 > #1990/01/01# Then
            許可_処理年月日 = p許可日
        Else
            許可_処理年月日 = Now.Date
        End If

        If p申請.許可番号 > 0 Then
            Me.許可_処理番号 = p申請.許可番号
        Else
            Me.許可_処理番号 = Get許可番号MAX(mvar申請.法令)
            'Me.許可_処理番号 = Get許可番号MAX(mvar申請.法令)
        End If

    End Sub

    Public ReadOnly Property 受付年月日 As DateTime
        Get
            Return mvar申請.受付年月日
        End Get
    End Property

    Public Property 許可_処理年月日 As DateTime
        Get
            Return mvar申請.GetDateValue("許可年月日")
        End Get
        Set(value As DateTime)

            mvar申請.SetDateValue("許可年月日", value)
        End Set
    End Property

    Public Property 許可_処理番号 As Integer
        Get
            Return mvar申請.GetIntegerValue("許可番号")
        End Get
        Set(value As Integer)
            mvar申請.SetIntegerValue("許可番号", value)
        End Set
    End Property

    Private Function Get許可番号MAX(ByVal n法令 As enum法令) As Integer
        Dim s法令 As String = ""
        Select Case n法令
            Case enum法令.農地法5条貸借, enum法令.農地法5条所有権, enum法令.農地法5条一時転用 : s法令 = enum法令.農地法5条所有権 & "," & enum法令.農地法5条貸借 & "," & enum法令.農地法5条一時転用
            Case Else
                s法令 = n法令
        End Select
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Max([D_申請].許可番号) AS 最大番号 FROM [D_申請] WHERE ([D_申請].法令 In ( " & s法令 & " )) AND (DatePart('yyyy',[許可年月日])=DatePart('yyyy',Date()))")
        If pTBL.Rows.Count = 1 Then
            Return Val(pTBL.Rows(0).Item("最大番号").ToString) + 1
        Else
            Return 1
        End If
    End Function

    Public Function Get許可通年番号MAX(ByVal n法令 As Integer) As Integer
        Dim s法令 As String = ""
        Select Case n法令
            Case enum法令.農地法5条貸借, enum法令.農地法5条所有権, enum法令.農地法5条一時転用 : s法令 = enum法令.農地法5条所有権 & "," & enum法令.農地法5条貸借 & "," & enum法令.農地法5条一時転用
            Case Else
                s法令 = n法令
        End Select
        Dim pTable As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Max([D_申請].通年許可番号) AS 最大番号 FROM [D_申請] WHERE ([D_申請].法令 In ( " & n法令 & " )) AND (DatePart('yyyy',[許可年月日])=DatePart('yyyy',Date()))")
        If pTable.Rows.Count = 1 Then
            Return Val(pTable.Rows(0).Item("最大番号").ToString) + 1
        Else
            Return 1
        End If
    End Function
End Class

'------------------------ 申請取下げ
<TypeConverter(GetType(PropertyOrderConverter))>
Public Class C申請取下げ
    Inherits HimTools2012.InputSupport.CInputSupport

    Private mvar出し手 As CObj個人
    Private mvar受け手 As CObj個人
    Friend 取下げ理由 As Object
    Friend 取下げ年月日 As Date

    Public Sub New(ByRef p所有者 As HimTools2012.TargetSystem.CTargetObjWithView, ByRef p受け手 As HimTools2012.TargetSystem.CTargetObjWithView, ByVal n法令 As enum法令, Optional p農地() As CObj農地 = Nothing)
        MyBase.New(農地基本台帳.App農地基本台帳.DataMaster.Body)

        If Not p所有者 Is Nothing Then
            Select Case p所有者.Key.DataClass
                Case "農家" : mvar出し手 = p所有者.GetProperty("世帯主")
                Case "個人" : mvar出し手 = p所有者
            End Select
        End If

        If Not p受け手 Is Nothing Then
            Select Case p受け手.Key.DataClass
                Case "農家" : mvar受け手 = p受け手.GetProperty("世帯主")
                Case "個人" : mvar受け手 = p受け手
            End Select
        End If


        mvar受け手 = p受け手

    End Sub

    <PropertyOrderAttribute(1)> <Category("01 基本情報")> <DefaultValue("")> <Description("譲渡・貸し手")>
    Public ReadOnly Property 譲渡_貸し手() As String
        Get
            If mvar出し手 IsNot Nothing Then
                Return mvar出し手.ToString
            Else
                Return ""
            End If
        End Get
    End Property
    <PropertyOrderAttribute(2)> <Category("01 基本情報")> <DefaultValue("")> <Description("譲受・借り手")>
    Public ReadOnly Property 譲受_借り手() As String
        Get
            If mvar受け手 IsNot Nothing Then
                Return mvar受け手.ToString
            Else
                Return ""
            End If
        End Get
    End Property


    <PropertyOrderAttribute(3)> <Category("02 取下げ・取り消し内容")> <DefaultValue("")> <Description("届出の年月日")>
    Public Property 届出年月日 As DateTime = Now.Date

    <PropertyOrderAttribute(4)> <Category("02 取下げ・取り消し内容")> <DefaultValue("")> <Description("届出の理由")> <EditorAttribute(GetType(HimTools2012.TypeConverterCustom.LongTextUIEditor), GetType(System.Drawing.Design.UITypeEditor))>
    Public Property 理由 As String = ""

End Class

<TypeConverter(GetType(PropertyOrderConverter))>
Public Class C異動日入力
    Inherits HimTools2012.InputSupport.CInputSupport

    Public Sub New()
        MyBase.New(農地基本台帳.App農地基本台帳.DataMaster.Body)
    End Sub

    <Category("異動日")>
    Public Property 異動日 As DateTime

    Public Overrides Function DataCompleate() As Boolean
        If Not 異動日.Year > 1900 Then
            異動日 = Now
        End If

        Return IsDate(異動日)
    End Function
End Class

'------------------------ 範囲入力
