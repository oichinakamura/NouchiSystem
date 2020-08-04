
Imports System.ComponentModel
Imports System.Drawing.Design
Imports System.Windows.Forms.Design
Imports HimTools2012.controls.PropertyGridSupport
Imports HimTools2012.TypeConverterCustom

<TypeConverter(GetType(PropertyOrderConverter))>
Public Class C農地検索条件
    Inherits Common検索条件
    Public Enum enum有無効選択
        有効 = True
        無効 = False
    End Enum
    Public Enum enum荒廃状況分類
        条件なし = -1
        荒廃していない = 0
        A分類 = 1
        B分類 = 2
        調査中 = 3
    End Enum
    Public Enum enum利用状況調査結果
        条件なし = -1
        未設定 = 0
        農法32条1項1号 = 1
        農法32条1項2号 = 2
        遊休農地でない = 3
        その他 = 4
        調査不可 = 5
        調査中 = 6
    End Enum

    <Category("01.農地条件")> <PropertyOrderAttribute(0)> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Off)>
    Public Property ID As Integer = Nothing

    Private mvar大字 As Integer = 0
    <Category("01.農地条件")> <PropertyOrderAttribute(1)> <TypeConverter(GetType(大字ComboListConverter))> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Off)>
    Public Property 大字 As String
        Get
            If IsDBNull(mvar大字) Then
                Return "0000000:-"
            Else
                Dim pView As New DataView(App農地基本台帳.DataMaster.Body, "[Class]='大字' AND [ID]=" & mvar大字, "", DataViewRowState.CurrentRows)
                If pView.Count > 0 Then
                    Return String.Format("{0:D7}:{1}", mvar大字, pView.Item(0).Item("名称"))
                Else
                    Return "0000000:-"
                End If
            End If
        End Get
        Set(ByVal value As String)
            mvar大字 = Val(value)
            検索Common.mvar大字Code = Val(value)
            検索Common.mvar検索小字View = New DataView(App農地基本台帳.DataMaster.Body, "Class='小字' AND [nParam]=" & 検索Input.mvar大字Code, "ID", DataViewRowState.CurrentRows)
        End Set
    End Property
    Private mvar小字 As Integer
    <Category("01.農地条件")> <PropertyOrderAttribute(2)> <TypeConverter(GetType(小字ComboListConverter))> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Off)>
    Public Property 小字 As String
        Get
            If IsDBNull(mvar小字) Then
                Return "0000000:-"
            Else
                Dim pView As New DataView(App農地基本台帳.DataMaster.Body, "[Class]='小字' AND [ID]=" & mvar小字, "", DataViewRowState.CurrentRows)
                If pView.Count > 0 Then
                    Return String.Format("{0:D7}:{1}", mvar小字, pView.Item(0).Item("名称"))
                Else
                    Return "0000000:-"
                End If
            End If
        End Get
        Set(ByVal value As String)
            mvar小字 = Val(value)
        End Set
    End Property

    <Category("01.農地条件")> <PropertyOrderAttribute(3)> <TypeConverter(GetType(AutoCollectConverter))> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Off)> <Description("末尾に*を付けることであいまい検索に対応します。")>
    Public Property 地番 As String = ""



    Private mvar登記地目 As C市町村別.地目Type = C市町村別.地目Type.指定なし
    <Category("04.地目条件")> <PropertyOrderAttribute(4)> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Off)>
    Public Property 登記地目 As C市町村別.地目Type
        Get
            Return mvar登記地目
        End Get
        Set(ByVal value As C市町村別.地目Type)
            mvar登記地目 = value
        End Set
    End Property


    Private mvar現況地目 As C市町村別.地目Type = C市町村別.地目Type.指定なし
    <Category("04.地目条件")> <PropertyOrderAttribute(5)> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Off)>
    Public Property 現況地目 As C市町村別.地目Type
        Get
            Return mvar現況地目
        End Get
        Set(ByVal value As C市町村別.地目Type)
            mvar現況地目 = value
        End Set
    End Property


    Private mvar農委地目 As Integer = 0
    <Category("04.地目条件")> <PropertyOrderAttribute(6)> <TypeConverter(GetType(農委地目ComboListConverter))> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Off)>
    Public Property 農委地目 As String
        Get
            If IsDBNull(mvar農委地目) Then
                Return "0000000:-"
            Else
                Dim pView As New DataView(App農地基本台帳.DataMaster.Body, "[Class]='農委地目' AND [ID]=" & mvar農委地目, "", DataViewRowState.CurrentRows)
                If pView.Count > 0 Then
                    Return String.Format("{0:D7}:{1}", mvar農委地目, pView.Item(0).Item("名称"))
                Else
                    Return "0000000:-"
                End If
            End If
        End Get
        Set(ByVal value As String)
            mvar農委地目 = Val(value)
        End Set
    End Property


    Private mvar自小作別 As Integer = -2
    <Category("05.自小作条件")> <PropertyOrderAttribute(7)> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Off)>
    Public Property 自小作別 As enum自小作別
        Get
            Return mvar自小作別
        End Get
        Set(ByVal value As enum自小作別)
            mvar自小作別 = value
        End Set
    End Property

    <Category("16.農地の利用状況調査")> <PropertyOrderAttribute(8)>
    Public Property 荒廃状況分類 As enum荒廃状況分類 = -1

    <Category("16.農地の利用状況調査")>
    Public Property 利用状況調査結果 As enum利用状況調査結果 = -1

    Private mvar賃借料検索 As New 賃借料検索
    Private mvar賃借開始 As New 賃借開始
    <Category("72.賃借関連")> <PropertyOrderAttribute(9)>
    Public ReadOnly Property 賃借開始条件() As 賃借開始
        Get
            Return mvar賃借開始
        End Get
    End Property


    <Category("72.賃借関連")>
    Public ReadOnly Property 賃借料10a当たり() As 賃借料検索
        Get
            Return mvar賃借料検索
        End Get
    End Property


    Public Overrides Function ToString() As String
        Dim sB As New System.Text.StringBuilder
        Dim sAND As String = ""

        If IsNumeric(ID) AndAlso ID <> 0 Then
            sB.Append(sAND & String.Format("[ID] = {0}", ID))
            sAND = " AND "
        End If

        If Val(大字) > 0 Then
            sB.Append(sAND & String.Format("[大字ID] = {0}", Val(大字)))
            sAND = " AND "
        End If

        If Val(小字) <> 0 Then
            sB.Append(sAND & String.Format("[小字ID] ={0}", Val(小字)))
            sAND = " AND "
        End If

        If 地番.Length > 0 Then
            sB.Append(sAND & Getワイルドカード文字("地番", 地番).Replace("*", "%"))
            sAND = " AND "

            SysAD.検索関連.AutoCollectStr("地番") = Me.地番
        End If

        If 荒廃状況分類 > -1 Then
            sB.Append(sAND & "[利用状況調査荒廃]=" & 荒廃状況分類)
            sAND = " AND "
        End If

        If 利用状況調査結果 > -1 Then
            sB.Append(sAND & "[利用状況調査農地法]=" & 利用状況調査結果)
            sAND = " AND "
        End If

        If 登記地目 > -1 Then
            Dim s登記地目List As New System.Text.StringBuilder
            For Each n地目CD As Integer In CType(SysAD.市町村, C市町村別).市町村別登記地目CD(登記地目)
                s登記地目List.Append(IIf(s登記地目List.Length > 0, ",", "") & n地目CD)
            Next
            If s登記地目List.Length > 0 Then
                sB.Append(sAND & "[登記簿地目] IN (" & s登記地目List.ToString & ")")
            End If
            sAND = " AND "
        End If

        If mvar農委地目 <> 0 Then
            sB.Append(sAND & "[農委地目ID] =" & mvar農委地目)
            sAND = " AND "
        End If


        If 現況地目 > -1 Then
            Dim s現況地目List As New System.Text.StringBuilder
            For Each n地目CD As Integer In CType(SysAD.市町村, C市町村別).市町村別現況地目CD(現況地目)
                s現況地目List.Append(IIf(s現況地目List.Length > 0, ",", "") & n地目CD)
            Next
            If s現況地目List.Length > 0 Then
                sB.Append(sAND & "[現況地目] IN (" & s現況地目List.ToString & ")")
            End If
            sAND = " AND "
        End If

        Dim s賃借開始 As String = mvar賃借開始.ToSQL

        If 自小作別 > -2 Then
            sB.Append(sAND & String.Format("[自小作別] = {0}", mvar自小作別))
            sAND = " AND "
        ElseIf s賃借開始.Length > 0 Then
            sB.Append(sAND & "[自小作別] >0 AND " & s賃借開始)
            sAND = " AND "
        End If

        If mvar賃借料検索.ToSQL.Length > 0 Then
            sB.Append(sAND & mvar賃借料検索.ToSQL)
            sAND = " AND "

        End If

        Return sB.ToString()
    End Function

    Public Sub New()
        検索Common.mvar検索小字View = New DataView(App農地基本台帳.DataMaster.Body, "Class='小字'", "ID", DataViewRowState.CurrentRows)
    End Sub

    Public Overrides Function View検索条件() As String
        Return Me.ToString()
    End Function
End Class

<TypeConverter(GetType(ExpandableObjectConverter))>
Public Class 賃借開始
    Private mvar貸借始期開始日 As DateTime = Nothing
    Private mvar貸借始期終了日 As DateTime = Nothing

    Public Property いつから As DateTime
        Get
            Return mvar貸借始期開始日
        End Get
        Set(value As DateTime)
            mvar貸借始期開始日 = value
        End Set
    End Property

    Public Property いつまで As DateTime
        Get
            Return mvar貸借始期終了日
        End Get
        Set(value As DateTime)
            mvar貸借始期終了日 = value
        End Set
    End Property

    <Browsable(False)>
    Public Overrides Function ToString() As String
        Dim sB As New System.Text.StringBuilder
        If Not IsNothing(mvar貸借始期開始日) AndAlso IsDate(mvar貸借始期開始日) AndAlso mvar貸借始期開始日.Year > 1901 Then
            sB.Append(和暦Format(mvar貸借始期開始日, "gyy/M/d") & "～")
        End If
        If Not IsNothing(mvar貸借始期終了日) AndAlso IsDate(mvar貸借始期終了日) AndAlso mvar貸借始期終了日.Year > 1901 Then
            sB.Append("～" & 和暦Format(mvar貸借始期終了日, "gyy/M/d"))
        End If
        If sB.Length > 0 Then
            sB.Append("に開始した")
        End If

        Return Replace(sB.ToString(), "～～", "～")
    End Function

    <Browsable(False)>
    Public ReadOnly Property ToSQL() As String
        Get
            Dim sB As New System.Text.StringBuilder
            If Not IsNothing(mvar貸借始期開始日) AndAlso IsDate(mvar貸借始期開始日) AndAlso mvar貸借始期開始日.Year > 1901 Then
                sB.Append(String.Format("[小作開始年月日]>={0}", HimTools2012.StringF.Toリテラル日付(mvar貸借始期開始日)))
            End If
            If Not IsNothing(mvar貸借始期終了日) AndAlso IsDate(mvar貸借始期終了日) AndAlso mvar貸借始期終了日.Year > 1901 Then
                sB.Append(IIf(sB.Length > 0, " AND ", "") & String.Format("[小作開始年月日]<={0}", HimTools2012.StringF.Toリテラル日付(mvar貸借始期終了日)))
            End If

            Return sB.ToString()
        End Get
    End Property

End Class


<TypeConverter(GetType(ExpandableObjectConverter))>
Public Class 賃借料検索
    Private mvar最低額 As Decimal = 0
    Private mvar最高額 As Decimal = 0

    Public Property 最低額 As Decimal
        Get
            Return mvar最低額
        End Get
        Set(value As Decimal)
            mvar最低額 = value
        End Set
    End Property
    Public Property 最高額 As Decimal
        Get
            Return mvar最高額
        End Get
        Set(value As Decimal)
            mvar最高額 = value
        End Set
    End Property
    Public Overrides Function ToString() As String

        If mvar最低額 > 0 AndAlso mvar最高額 > 0 AndAlso mvar最低額 > mvar最高額 Then
            Dim pSw As Decimal = mvar最高額
            mvar最高額 = mvar最低額
            mvar最低額 = pSw
        End If

        If mvar最低額 > 0 AndAlso mvar最高額 > 0 Then
            Return "範囲(" & mvar最低額 & " - " & mvar最高額 & ")"
        ElseIf mvar最低額 = 0 AndAlso mvar最高額 > 0 Then
            Return "範囲( <= " & mvar最高額 & ")"
        ElseIf mvar最低額 > 0 AndAlso mvar最高額 = 0 Then
            Return "範囲(" & mvar最低額 & " <=)"
        Else
            Return "-"
        End If
        Return MyBase.ToString()
    End Function
    <Browsable(False)>
    Public ReadOnly Property ToSQL() As String
        Get
            If mvar最低額 > 0 AndAlso mvar最高額 > 0 Then
                Return "([10a賃借料]>=" & mvar最低額 & " AND [10a賃借料]<=" & mvar最高額 & ")"
            ElseIf mvar最低額 = 0 AndAlso mvar最高額 > 0 Then
                Return "[10a賃借料]<=" & mvar最高額
            ElseIf mvar最低額 > 0 AndAlso mvar最高額 = 0 Then
                Return "[10a賃借料]>=" & mvar最低額
            Else
                Return ""
            End If
        End Get
    End Property

End Class


<TypeConverter(GetType(PropertyOrderConverter))>
Public Class C固定資産農地検索
    Inherits Common検索条件

    Private mvar大字 As Integer = 0
    <Category("01.農地条件")> <PropertyOrderAttribute(1)> <TypeConverter(GetType(大字ComboListConverter))> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Off)>
    Public Property 大字 As String
        Get
            If IsDBNull(mvar大字) Then
                Return "0000000:-"
            Else
                Dim pView As New DataView(App農地基本台帳.DataMaster.Body, "[Class]='大字' AND [ID]=" & mvar大字, "", DataViewRowState.CurrentRows)
                If pView.Count > 0 Then
                    Return String.Format("{0:D7}:{1}", mvar大字, pView.Item(0).Item("名称"))
                Else
                    Return "0000000:-"
                End If
            End If
        End Get
        Set(ByVal value As String)
            mvar大字 = Val(value)
            検索Common.mvar大字Code = Val(value)
            検索Common.mvar検索小字View = New DataView(App農地基本台帳.DataMaster.Body, "Class='小字' AND [nParam]=" & 検索Input.mvar大字Code, "ID", DataViewRowState.CurrentRows)
        End Set
    End Property
    Private mvar小字 As Integer
    <Category("01.農地条件")> <PropertyOrderAttribute(2)> <TypeConverter(GetType(小字ComboListConverter))> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Off)>
    Public Property 小字 As String
        Get
            If IsDBNull(mvar小字) Then
                Return "0000000:-"
            Else
                Dim pView As New DataView(App農地基本台帳.DataMaster.Body, "[Class]='小字' AND [ID]=" & mvar小字, "", DataViewRowState.CurrentRows)
                If pView.Count > 0 Then
                    Return String.Format("{0:D7}:{1}", mvar小字, pView.Item(0).Item("名称"))
                Else
                    Return "0000000:-"
                End If
            End If
        End Get
        Set(ByVal value As String)
            mvar小字 = Val(value)
        End Set
    End Property

    <Category("01.農地条件")> <PropertyOrderAttribute(3)> <TypeConverter(GetType(AutoCollectConverter))> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Off)> <Description("末尾に*を付けることであいまい検索に対応します。")>
    Public Property 地番 As String = ""

    Public Overrides Function ToString() As String
        Dim sB As New System.Text.StringBuilder
        Dim sAND As String = ""


        If Val(大字) > 0 Then
            sB.Append(sAND & String.Format("[大字ID] = {0}", Val(大字)))
            sAND = " AND "
        End If

        If Val(小字) <> 0 Then
            sB.Append(sAND & String.Format("[小字ID] ={0}'", Val(小字)))
            sAND = " AND "
        End If

        If 地番.Length > 0 Then
            sB.Append(sAND & Getワイルドカード文字("地番", 地番).Replace("*", "%"))
            sAND = " AND "

            SysAD.検索関連.AutoCollectStr("地番") = Me.地番
        End If

        Return sB.ToString()
    End Function

    Public Overrides Function View検索条件() As String
        Return ToString()
    End Function
End Class

<TypeConverter(GetType(PropertyOrderConverter))>
Public Class C転用農地検索
    Inherits Common検索条件

    Private mvar大字 As Integer = 0
    <Category("01.農地条件")> <PropertyOrderAttribute(1)> <TypeConverter(GetType(大字ComboListConverter))> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Off)>
    Public Property 大字 As String
        Get
            If IsDBNull(mvar大字) Then
                Return "0000000:-"
            Else
                Dim pView As New DataView(App農地基本台帳.DataMaster.Body, "[Class]='大字' AND [ID]=" & mvar大字, "", DataViewRowState.CurrentRows)
                If pView.Count > 0 Then
                    Return String.Format("{0:D7}:{1}", mvar大字, pView.Item(0).Item("名称"))
                Else
                    Return "0000000:-"
                End If
            End If
        End Get
        Set(ByVal value As String)
            mvar大字 = Val(value)
            検索Common.mvar大字Code = Val(value)
            検索Common.mvar検索小字View = New DataView(App農地基本台帳.DataMaster.Body, "Class='小字' AND [nParam]=" & 検索Input.mvar大字Code, "ID", DataViewRowState.CurrentRows)
        End Set
    End Property
    Private mvar小字 As Integer
    <Category("01.農地条件")> <PropertyOrderAttribute(2)> <TypeConverter(GetType(小字ComboListConverter))> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Off)>
    Public Property 小字 As String
        Get
            If IsDBNull(mvar小字) Then
                Return "0000000:-"
            Else
                Dim pView As New DataView(App農地基本台帳.DataMaster.Body, "[Class]='小字' AND [ID]=" & mvar小字, "", DataViewRowState.CurrentRows)
                If pView.Count > 0 Then
                    Return String.Format("{0:D7}:{1}", mvar小字, pView.Item(0).Item("名称"))
                Else
                    Return "0000000:-"
                End If
            End If
        End Get
        Set(ByVal value As String)
            mvar小字 = Val(value)
        End Set
    End Property

    <Category("01.農地条件")> <PropertyOrderAttribute(3)> <TypeConverter(GetType(AutoCollectConverter))> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Off)> <Description("末尾に*を付けることであいまい検索に対応します。")>
    Public Property 地番 As String = ""

    Public Overrides Function ToString() As String
        Dim sB As New System.Text.StringBuilder
        Dim sAND As String = ""


        If Val(大字) > 0 Then
            sB.Append(sAND & String.Format("[大字ID] = {0}", Val(大字)))
            sAND = " AND "
        End If

        If Val(小字) <> 0 Then
            sB.Append(sAND & String.Format("[小字ID] ={0}", Val(小字)))
            sAND = " AND "
        End If

        If 地番.Length > 0 Then
            sB.Append(sAND & Getワイルドカード文字("地番", 地番).Replace("*", "%"))
            sAND = " AND "

            SysAD.検索関連.AutoCollectStr("地番") = Me.地番
        End If

        Return sB.ToString()
    End Function

    Public Overrides Function View検索条件() As String
        Return ToString()
    End Function
End Class

<TypeConverter(GetType(PropertyOrderConverter))>
Public Class C削除農地検索
    Inherits Common検索条件

    Private mvar大字 As Integer = 0
    <Category("01.農地条件")> <PropertyOrderAttribute(1)> <TypeConverter(GetType(大字ComboListConverter))> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Off)>
    Public Property 大字 As String
        Get
            If IsDBNull(mvar大字) Then
                Return "0000000:-"
            Else
                Dim pView As New DataView(App農地基本台帳.DataMaster.Body, "[Class]='大字' AND [ID]=" & mvar大字, "", DataViewRowState.CurrentRows)
                If pView.Count > 0 Then
                    Return String.Format("{0:D7}:{1}", mvar大字, pView.Item(0).Item("名称"))
                Else
                    Return "0000000:-"
                End If
            End If
        End Get
        Set(ByVal value As String)
            mvar大字 = Val(value)
            検索Common.mvar大字Code = Val(value)
            検索Common.mvar検索小字View = New DataView(App農地基本台帳.DataMaster.Body, "Class='小字' AND [nParam]=" & 検索Input.mvar大字Code, "ID", DataViewRowState.CurrentRows)
        End Set
    End Property
    Private mvar小字 As Integer
    <Category("01.農地条件")> <PropertyOrderAttribute(2)> <TypeConverter(GetType(小字ComboListConverter))> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Off)>
    Public Property 小字 As String
        Get
            If IsDBNull(mvar小字) Then
                Return "0000000:-"
            Else
                Dim pView As New DataView(App農地基本台帳.DataMaster.Body, "[Class]='小字' AND [ID]=" & mvar小字, "", DataViewRowState.CurrentRows)
                If pView.Count > 0 Then
                    Return String.Format("{0:D7}:{1}", mvar小字, pView.Item(0).Item("名称"))
                Else
                    Return "0000000:-"
                End If
            End If
        End Get
        Set(ByVal value As String)
            mvar小字 = Val(value)
        End Set
    End Property

    <Category("01.農地条件")> <PropertyOrderAttribute(3)> <TypeConverter(GetType(AutoCollectConverter))> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Off)> <Description("末尾に*を付けることであいまい検索に対応します。")>
    Public Property 地番 As String = ""

    Public Overrides Function ToString() As String
        Dim sB As New System.Text.StringBuilder
        Dim sAND As String = ""


        If Val(大字) > 0 Then
            sB.Append(sAND & String.Format("[大字ID] = {0}", Val(大字)))
            sAND = " AND "
        End If

        If Val(小字) <> 0 Then
            sB.Append(sAND & String.Format("[小字ID] ={0}'", Val(小字)))
            sAND = " AND "
        End If

        If 地番.Length > 0 Then
            sB.Append(sAND & Getワイルドカード文字("地番", 地番).Replace("*", "%"))
            sAND = " AND "

            SysAD.検索関連.AutoCollectStr("地番") = Me.地番
        End If

        Return sB.ToString()
    End Function

    Public Overrides Function View検索条件() As String
        Return ToString()
    End Function
End Class

Public Module 検索Common
    Public mvar大字Code As Integer = 0
    Public mvar検索小字View As DataView
End Module

Public Class 大字ComboListConverter
    Inherits HimTools2012.InputSupport.ComboListConverter

    Public Sub New()
        MyBase.New(New DataView(App農地基本台帳.DataMaster.Body, "Class='大字'", "ID", DataViewRowState.CurrentRows))
    End Sub

    Public Overrides Function GetFilter() As String
        Return "Class='大字'"
    End Function
End Class



Public Class 小字ComboListConverter
    Inherits HimTools2012.InputSupport.ComboListConverter

    Public Sub New()
        MyBase.New(検索Common.mvar検索小字View)
    End Sub

    Public Overrides Function GetFilter() As String
        If 検索Common.mvar大字Code <> 0 Then
            Return "Class='小字' AND [nParam]=" & 検索Common.mvar大字Code
        Else
            Return "Class='小字'"
        End If
    End Function
End Class

Public Class 農委地目ComboListConverter
    Inherits HimTools2012.InputSupport.ComboListConverter

    Public Sub New()
        MyBase.New(New DataView(App農地基本台帳.DataMaster.Body, "Class='農委地目'", "ID", DataViewRowState.CurrentRows))
    End Sub

    Public Overrides Function GetFilter() As String
        Return "Class='農委地目'"
    End Function
End Class
