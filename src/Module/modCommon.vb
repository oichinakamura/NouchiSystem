Imports System.ComponentModel
Imports System.ComponentModel.TypeConverter
Imports System.Globalization
Imports HimTools2012
Imports HimTools2012.controls
Imports HimTools2012.controls.DVCtrlCommonBase

''' <summary></summary>
''' <remarks>
''' 未検証 件数3
''' </remarks>
Public Class 議案書作成パラメータ


    ''' <summary></summary>
    ''' <value></value>
    ''' <remarks>Verified [中村 雄一 date：2016/9/14 15:6]</remarks>
    Public Property 総会日 As DateTime

    ''' <summary></summary>
    ''' <value></value>
    ''' <remarks>
    ''' Verified [中村 雄一 date：2016/9/14 15:6]
    ''' </remarks>
    Public Property 開始年月日 As DateTime

    ''' <summary></summary>
    ''' <value></value>
    ''' <remarks>
    ''' Verified [中村 雄一 date：2016/9/14 15:6]
    ''' </remarks>
    Public Property 終了年月日 As DateTime

    ''' <summary></summary>
    ''' <remarks></remarks>
    Public n対象年 As Integer

    ''' <summary></summary>
    ''' <remarks></remarks>
    Public n対象月 As Integer

    ''' <summary></summary>
    ''' <remarks>
    ''' Todo 検証してください！！未検証 コメント作成：[2016/09/14 14:45]
    ''' </remarks>
    Public Sub New()
        Dim skDT As String

        n対象年 = Now.Year
        n対象月 = Now.Month + (Now.Day < 15)
        Dim n締日 As Integer = CType(SysAD.市町村, C市町村別).Get総会締日()

        If n締日 >= Now.Day Then
        Else
            n対象月 = n対象月 + 1
            If n対象月 > 12 Then
                n対象年 += 1
                n対象月 = 1
            End If
        End If
        skDT = GetDt("総会日", HimTools2012.DateFunctions.GetMaxDay(n対象年, n対象月, 31))
        n締日 = HimTools2012.DateFunctions.GetMaxDay(n対象年, n対象月, n締日)


        終了年月日 = DateSerial(n対象年, n対象月, n締日)
        If n締日 = HimTools2012.DateFunctions.GetMaxDay(n対象年, n対象月, 31) Then
            開始年月日 = DateSerial(n対象年, n対象月, 1)
        Else
            開始年月日 = 終了年月日.AddMonths(-1).AddDays(1)
        End If

        総会日 = DateSerial(n対象年, n対象月, Val(skDT))

    End Sub

    ''' <summary></summary>
    ''' <param name="sKey"></param>
    ''' <param name="MaxDay"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Todo 検証してください！！未検証 コメント作成：[2016/09/14 14:34]
    ''' </remarks>
    Private Function GetDt(ByVal sKey As String, ByVal MaxDay As Long) As String
        Dim sDT As String

        sDT = SysAD.DB(sLRDB).DBProperty(sKey)
        Do Until Val(sDT) > 0 And Val(sDT) < MaxDay
            sDT = InputBox(sKey & "日を入力してください", sKey & "日", 15, 1)

            If Len(sDT) = 0 Then

            ElseIf Val(sDT) > 0 And Val(sDT) < MaxDay Then
                SysAD.DB(sLRDB).DBProperty(sKey) = Val(sDT)
            End If
        Loop
        Return sDT
    End Function

    ''' <summary></summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' Todo 検証してください！！未検証 コメント作成：[2016/09/14 14:34]
    ''' </remarks>
    Public Function ToKey() As String
        Return n対象年.ToString & HimTools2012.StringF.Right("00" & n対象月, 2)
    End Function
End Class


''' <summary></summary>
''' <remarks></remarks>
Public Enum 農地状況
    非農地 = 51
    一時転用中 = 1040
    転用許可済み = 1050
End Enum


''' <summary></summary>
Public Class ImageKey

    ''' <summary></summary>
    ''' <remarks></remarks>
    Public Const 終了 = "_Exit"

    ''' <summary></summary>
    ''' <remarks></remarks>
    Public Const メンテナンス = "Maintenance"

    ''' <summary></summary>
    ''' <remarks></remarks>
    Public Const 作業 = "申請"

    ''' <summary></summary>
    ''' <remarks></remarks>
    Public Const 他システム連携 = "他システム連携"

    ''' <summary></summary>
    ''' <remarks></remarks>
    Public Const 閲覧検索 = "List"

    ''' <summary></summary>
    ''' <remarks></remarks>
    Public Const 集計一覧 = "集計一覧"

    ''' <summary></summary>
    ''' <remarks></remarks>
    Public Const 選挙関連 = "選挙関連"

    ''' <summary></summary>
    ''' <remarks></remarks>
    Public Const 印刷 = "printer"

    ''' <summary></summary>
    ''' <remarks></remarks>
    Public Const ヘルプ = "Help"
End Class


''' <summary></summary>
''' <remarks></remarks>
Public Enum enum法令
    農地法3条所有権 = 30
    農地法3条耕作権 = 31
    農地法3条の3第1項 = 311
    農地法4条 = 40
    農地法4条一時転用 = 42

    農地法5条所有権 = 50
    農地法5条貸借 = 51
    農地法5条一時転用 = 52

    農地法18条解約 = 180
    農地法20条解約 = 200
    合意解約 = 210
    中間管理機構へ農地の返還 = 250

    基盤強化法所有権 = 60
    利用権設定 = 61
    利用権移転 = 62
    中間管理機構経由 = 65

    農地改良届 = 301
    農用地計画変更 = 302
    事業計画変更 = 303

    あっせん出手 = 400
    あっせん受手 = 401

    農地利用目的変更 = 500
    非農地証明願 = 602

    買受適格耕公 = 801
    買受適格耕競 = 802
    買受適格転公 = 803
    買受適格転競 = 804

    奨励金交付A = 990
    奨励金交付B = 991

    満期解約 = 1000

    機構を介した利用権設定の受け手変更 = 2100

    分筆登記 = 9801
    換地処理 = 9821
    その他分割処理 = 9910
    その他分割統合 = 9911
    職権異動 = 10002
End Enum


Public Enum 平成
    H20 = 2008
    H21 = 2009
    H22 = 2010
    H23 = 2011
    H24 = 2012
    H25 = 2013
    H26 = 2014
    H27 = 2015
    H28 = 2016
    H29 = 2017
    H30 = 2018
    H31 = 2019
    R02 = 2020
End Enum


Public Enum enum有無
    有 = -1
    無 = 0
End Enum


Public Enum enum農業改善計画認定
    認定農業者 = 1
    担い手農家 = 2
    農業生産法人 = 3
    認定農業者_担い手農家 = 4
    認定農業者_農業生産法人 = 5
End Enum


''' <summary></summary>
''' <remarks></remarks>
Public Enum 農地法
    '利用意向調査
    農地法32の1 = 32001 '利用状況調査の結果、耕作の目的に供されないまたは周囲より利用の程度が著しく劣る農地の所有者に意向調査を実施する
    農地法32の4 = 32004 '利用状況調査の結果、32条1項に該当する農地の権利者が複数ある場合、32条3項の公示を行い、それに対して期間内に申出の無い場合、意向調査を実施する
    農地法33の1 = 33001 '耕作の事業に従事するものが不在（又は予測される）場合、その農地の所有者に対して、意向調査を実施する
End Enum


Module modCommon
    ''' <summary></summary>
    ''' <remarks></remarks>
    Public SysAD As CSystem
    ''' <summary></summary>
    ''' <remarks></remarks>
    Public ObjectMan As CObjectMan
    ''' <summary></summary>
    ''' <remarks></remarks>
    Public App農地基本台帳 As C農地基本台帳
    ''' <summary></summary>
    ''' <remarks></remarks>
    Public 選択氏名StandardValuesCollection As StandardValuesCollection

    ''' <summary></summary>
    ''' <param name="sKey"></param>
    ''' <param name="sXX"></param>
    ''' <remarks>
    ''' Todo 検証してください！！未検証 コメント作成：[2016/09/14 14:42]
    ''' </remarks>
    Public Sub CasePrint(ByVal sKey As String, Optional ByVal sXX As String = "")
        Dim sParent As String = ""
        Try
            Dim pSF As New System.Diagnostics.StackFrame(1)

            sParent = pSF.GetFileName & "\" & pSF.GetMethod.Name
        Catch ex As Exception

        End Try

        Debug.Print(String.Format("case ""{0}"" :{1} '{2}", sKey, sXX, sParent))

        If Not SysAD.IsClickOnceDeployed Then
            Stop
        End If
    End Sub

    ''' <summary></summary>
    ''' <param name="sParam"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Todo 検証してください！！未検証 コメント作成：[2016/09/14 14:42]
    ''' </remarks>
    Public Function FncDebug(ParamArray sParam() As Object) As String
#If DEBUG Then
        Dim sB As New System.Text.StringBuilder
        With New StackFrame(1)
            sB.AppendLine("場所=" & .GetMethod.DeclaringType.Name & "." & .GetMethod.Name)
        End With

        For Each St As String In sParam
            sB.AppendLine(St)
        Next
        Debug.Print(sB.ToString)
        MsgBox(sB.ToString)
        Return sB.ToString
#Else
        return ""
#End If

    End Function

    Public Function 和暦Format(ByVal p年月日 As DateTime, Optional ByVal sFormat As String = "gggyy年M月d日", Optional ByVal sNullStr As String = "-") As String
        Try
            Dim dt年月日 As DateTime
            If IsDBNull(p年月日) = True OrElse Convert.IsDBNull(p年月日) OrElse IsDate(p年月日.ToString()) = False Then
                Return sNullStr
            Else
                dt年月日 = CDate(p年月日)
            End If

            Dim n年号Length As Integer = 0
            If (sFormat.ToLower().StartsWith("ggg")) Then n年号Length += 1
            If (sFormat.ToLower().StartsWith("gg")) Then n年号Length += 1
            If (sFormat.ToLower().StartsWith("g")) Then n年号Length += 1

            Dim culture As CultureInfo = New CultureInfo("ja-JP", True)
            culture.DateTimeFormat.Calendar = New JapaneseCalendar()

            Dim target As DateTime = dt年月日
            Dim result As String = target.ToString(sFormat, culture)
            If InStr(result, "平成") > 0 AndAlso CInt(target.ToString("yyyyMMdd")) >= 20190501 Then
                target = dt年月日.AddYears(-30)
                result = target.ToString(sFormat, culture).Replace("平成", "令和")
            End If

            Select Case n年号Length
                Case 1
                    Select Case StringF.Left(result, 2)
                        Case "明治" : result = StringF.Replace(result, "明治", "M")
                        Case "大正" : result = StringF.Replace(result, "大正", "T")
                        Case "昭和" : result = StringF.Replace(result, "昭和", "S")
                        Case "平成" : result = StringF.Replace(result, "平成", "H")
                        Case "令和" : result = StringF.Replace(result, "令和", "R")
                    End Select
                Case 2
                    Select Case StringF.Left(result, 2)
                        Case "明治" : result = StringF.Replace(result, "明治", "明")
                        Case "大正" : result = StringF.Replace(result, "大正", "大")
                        Case "昭和" : result = StringF.Replace(result, "昭和", "昭")
                        Case "平成" : result = StringF.Replace(result, "平成", "平")
                        Case "令和" : result = StringF.Replace(result, "令和", "令")
                    End Select
                Case Else
            End Select

            result = Replace(Replace(result, "令和1年", "令和元年"), "令和01年", "令和元年")
            result = Replace(Replace(result, "令1年", "令元年"), "令01年", "令元年")

            Return result
        Catch ex As Exception
            Return "変換できません"
        End Try
    End Function

    Public Class DataGridViewDateTimePickerColumn
        Inherits DataGridViewColumn

        Private mvarGroup As String = ""

        Public Sub New()
            MyBase.New(New CalendarCell)
        End Sub

        Public Sub New(ByVal sName As String, ByVal sColumnName As String, ByVal sHeaderText As String, ByVal sGroup As String, Optional ByVal sCustomFormat As String = "gyy/MM/dd")
            MyBase.New
            Me.Name = sName
            Me.HeaderText = sHeaderText
            Me.DataPropertyName = sColumnName
            Me.DefaultCellStyle.Format = sCustomFormat
            Me.SortMode = DataGridViewColumnSortMode.Automatic
            Me.mvarGroup = sGroup
        End Sub

        Public Property Format As String
            Get
                Return Me.DefaultCellStyle.Format
            End Get
            Set
                Me.DefaultCellStyle.Format = Value
            End Set
        End Property

        Public Overrides Property CellTemplate As DataGridViewCell
            Get
                Return MyBase.CellTemplate
            End Get
            Set
                If ((Not (Value) Is Nothing) _
                            AndAlso Not Value.GetType.IsAssignableFrom(GetType(CalendarCell))) Then
                    Throw New InvalidCastException("Must be a CalendarCell")
                End If

                MyBase.CellTemplate = Value
            End Set
        End Property

        Public ReadOnly Property Group As String
            Get
                Return Me.mvarGroup
            End Get
        End Property

        Public Overrides Function Clone() As Object
            Dim col As DataGridViewDateTimePickerColumn = CType(MyBase.Clone, DataGridViewDateTimePickerColumn)
            col.DefaultCellStyle.FormatProvider = Me.DefaultCellStyle.FormatProvider
            col.DefaultCellStyle.Format = Me.DefaultCellStyle.Format
            Return col
        End Function
    End Class

    Public Class DateTimePickerPlus2
        Inherits System.Windows.Forms.DateTimePicker
        Implements DVCtrlCommon

        Private mvarParams As CommonParamater = New CommonParamater

        Public ReadOnly Property Params As CommonParamater
            Get
                Return Me.mvarParams
            End Get
        End Property

        Private mvarDateStyle As DateStyle

        Private mvarBindStartValue As Object

        Public Sub New(ByVal bReadOnly As emRO, ByVal pTarget As TargetSystem.CTargetObjWithView, ByVal sFieldName As String, Optional ByVal nWidth As Integer = -1, Optional ByVal dumy As System.Data.DataTable = Nothing, Optional ByVal nDateStyle As DateStyle = DateStyle.typ和暦)
            Me.Params.BindingValue = New controls.CBindingValue(sFieldName, "Value", pTarget)
            Me.ShowCheckBox = True
            MyBase.Format = DateTimePickerFormat.Custom
            Me.Margin = New Padding(0, 0, 0, 0)
            Me.Enabled = (bReadOnly <> emRO.IsReadOnly)
            If (nWidth > 0) Then
                Me.Width = nWidth
            End If

            Me.mvarDateStyle = nDateStyle
            AddHandler ValueChanged, AddressOf Me.DateTimePickerPlus_ValueChanged
            AddHandler FormatChanged, AddressOf Me.DateTimePickerPlus_FormatChanged
            AddHandler MouseDown, AddressOf Me.DateTimePickerPlus_MouseDown
            AddHandler BindingContextChanged, AddressOf Me.DateTimePickerPlus_BindingContextChanged
        End Sub

        Private Sub DateTimePickerPlus_BindingContextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim pRow As System.Data.DataRow = CType(Me.DataBindings(0).DataSource, System.Data.DataView)(0).Row
            Me.Format = DateTimePickerFormat.Custom
            If Convert.IsDBNull(pRow(Me.DataBindings(0).BindingMemberInfo.BindingMember)) Then
                MyBase.Checked = False
                MyBase.CustomFormat = "-"
            ElseIf (CType(pRow(Me.DataBindings(0).BindingMemberInfo.BindingMember), DateTime) = New DateTime(1899, 12, 30)) Then
                MyBase.Checked = False
                MyBase.CustomFormat = "-"
            Else
                MyBase.Checked = True
                MyBase.CustomFormat = Me.FormatJPCalendar(MyBase.Value)
            End If

            Me.mvarBindStartValue = MyBase.Value
        End Sub

        Private Sub DateTimePickerPlus_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
            If (e.Button = System.Windows.Forms.MouseButtons.Right) Then
                Dim sText As String = CommonFunc.InpuText("入力してください", "日付直接入力", MyBase.Value.ToString)
                If DateFunctions.IsDate(sText) Then
                    Me.Value = DateTime.Parse(sText)
                    Me.DataBindings("Value").WriteValue()
                End If

            End If

        End Sub

        Private Sub DateTimePickerPlus_FormatChanged(ByVal sender As Object, ByVal e As System.EventArgs)
            Me.Refresh()
        End Sub

        Private Sub DateTimePickerPlus_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)
            Me.Format = DateTimePickerFormat.Custom
            If MyBase.Checked Then
                MyBase.CustomFormat = Me.FormatJPCalendar(MyBase.Value)
            Else
                MyBase.CustomFormat = "-"
            End If

            Me.Params.IsUpdate = True
            Me.DataBindings("Value").WriteValue()
        End Sub

        Public Shadows Property Value As Object
            Get
                If MyBase.Checked Then
                    Return MyBase.Value
                Else
                    Return Convert.DBNull
                End If

            End Get
            Set
                Try
                    MyBase.Format = DateTimePickerFormat.Custom
                    If (Convert.IsDBNull(Value) OrElse ((TypeOf Value Is DateTime AndAlso (CType(Value, DateTime) = New DateTime(1899, 12, 30)))) OrElse (DateFunctions.IsDate(Value.ToString) AndAlso (DateTime.Parse(CType(Value, String)).Year = 1))) Then
                        MyBase.CustomFormat = "-"
                        MyBase.Checked = False
                    Else
                        MyBase.Value = Convert.ToDateTime(Value)
                        MyBase.Checked = True
                        MyBase.CustomFormat = Me.FormatJPCalendar(MyBase.Value)
                    End If
                Catch ex As Exception
                    If ((Not (Value) Is Nothing) _
                            AndAlso Not Convert.IsDBNull(Value)) Then
                        Try
                            MyBase.Value = Convert.ToDateTime(Value)
                            MyBase.Checked = True
                        Catch
                            MyBase.Checked = False
                        End Try

                    End If
                End Try
            End Set
        End Property

        Private ReadOnly Property DVCtrlCommon_Params As CommonParamater Implements DVCtrlCommon.Params
            Get
                Throw New NotImplementedException()
            End Get
        End Property

        Public Sub CompleteBinding(ByVal s As Object, ByVal e As System.Windows.Forms.BindingCompleteEventArgs)
            Me.Params.IsInitialize = False
            Me.Params.BeforeValue = Me.Value
        End Sub

        Private Function FormatJPCalendar(ByVal tday As DateTime) As String
            Dim cal As JapaneseCalendar = New JapaneseCalendar
            Dim era As Integer = cal.GetEra(tday)
            Dim nengo() As String = New String() {"明治", "大正", "昭和", "平成", "令和"}
            Select Case (era)
                Case 1, 2, 3, 4, 5
                    Dim result = String.Format("{0}{1:00}年MM月dd日", nengo((era - 1)), cal.GetYear(tday))
                    If ((result.IndexOf("平成") > 0) _
                                AndAlso (Integer.Parse(tday.ToString("yyyyMMdd")) >= 20190501)) Then
                        Return String.Format("{0}{1:00}年MM月dd日", "令和", cal.GetYear(tday.AddYears(-30)))
                    End If

                    Return result
                Case Else
                    Return "yyyy/MM/dd"
            End Select

        End Function

        Private Sub DVCtrlCommon_CompleteBinding(s As Object, e As BindingCompleteEventArgs) Implements DVCtrlCommon.CompleteBinding
            Throw New NotImplementedException()
        End Sub
    End Class
End Module

Namespace 農地基本情報情報
    Public Module 個人履歴事由
        ''' <summary></summary>
        ''' <remarks></remarks>
        Public Const 選挙権確定 = 10444
        ''' <summary></summary>
        ''' <remarks></remarks>
        Public Const 選挙印刷強制 = 20111
        ''' <summary></summary>
        ''' <remarks></remarks>
        Public Const 選挙送付拒否 = 20112
    End Module

    Public Module 農地異動事由
        ''' <summary></summary>
        ''' <remarks></remarks>
        Public Const その他 = 99970
        ''' <summary></summary>
        ''' <remarks></remarks>
        Public Const 現地調査 = 99972
        ''' <summary></summary>
        ''' <remarks></remarks>
        Public Const 国土調査 = 99978
        ''' <summary></summary>
        ''' <remarks></remarks>
        Public Const 固定資産職権修正 = 99979
        ''' <summary></summary>
        ''' <remarks></remarks>
        Public Const 法務局照合 = 99980
        ''' <summary></summary>
        ''' <remarks></remarks>
        Public Const 売買 = 99982

        ''' <summary></summary>
        ''' <remarks></remarks>
        Public Const 分筆登記 = 99984
        ''' <summary></summary>
        ''' <remarks></remarks>
        Public Const 交換 = 99993
        ''' <summary></summary>
        ''' <remarks></remarks>
        Public Const 贈与 = 99994
        ''' <summary></summary>
        ''' <remarks></remarks>
        Public Const 相続移転 = 99995
        ''' <summary></summary>
        ''' <remarks></remarks>
        Public Const 職権による相続移転 = 99996
        ''' <summary></summary>
        ''' <remarks></remarks>
        Public Const 職権による所有権移転 = 99997
        ''' <summary></summary>
        ''' <remarks></remarks>
        Public Const 時効取得 = 99998
        ''' <summary></summary>
        ''' <remarks></remarks>
        Public Const 砂利採取工事完了 = 100000
        ''' <summary></summary>
        ''' <remarks></remarks>
        Public Const 遺贈 = 99993
    End Module
End Namespace

