
Imports System.ComponentModel
Imports HimTools2012.TypeConverterCustom
Imports HimTools2012.controls
Imports HimTools2012.controls.PropertyGridSupport

Public Enum enum有無選
    有り = 1
    無し = 2
    無効 = 0
End Enum

''' <summary>
''' 
''' </summary>
''' <remarks>
''' 未検証　件数[2]
''' </remarks>
Public MustInherit Class Common検索条件
    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:56]
    '''   
    '''  
    ''' </remarks>
    Public MustOverride Function View検索条件() As String

    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <param name="sField"></param>
    ''' <param name="St"></param>
    ''' <returns></returns>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:56]
    '''   
    '''  
    ''' </remarks>
    Public Function Getワイルドカード文字(sField As String, St As String) As String
        Dim X() As String = Split(St.Replace("**", "*"), "*")
        Dim sRet As New System.Text.StringBuilder
        If X.Length > 1 Then
            Dim sC As String = ""
            For i As Integer = 0 To X.Length - 1
                Dim sV As String = X(i)
                Dim sz As String = ""

                If i + 1 = X.Length Then
                    sz = ""
                ElseIf i + 1 < X.Length Then
                    sz = "*"
                End If


                If sV.Length = 0 Then
                    sC = "*"
                Else
                    If i > 0 Then
                        sC = "*"
                    End If
                    sRet.Append(IIf(sRet.Length > 0, " AND ", "") & String.Format("[" & sField & "] LIKE '{0}'", String.Format(sC & sV & sz)))
                    sC = ""
                    sz = ""
                End If
            Next
            If sRet.Length = 0 Then
                sRet.Append("[" & sField & "] LIKE '" & St & "'")
            End If
            Return sRet.ToString
        Else
            Return "[" & sField & "] LIKE '" & St & "'"
        End If
    End Function
End Class

''' <summary>
''' 
''' </summary>
''' <remarks>
''' 未検証　件数[12]
''' </remarks>
<TypeConverter(GetType(PropertyOrderConverter))>
Public Class C個人検索条件
    Inherits Common検索条件

    ''' <summary>
    ''' 
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''  未検証　コメント作成日[2016/9/12 17:2]
    '''  
    ''' </remarks>
    <Category("01_農家個人条件")> <PropertyOrderAttribute(0)> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Off)> <TypeConverter(GetType(AutoCollectConverter))>
    Public Property 住民番号 As String = ""

    ''' <summary>
    ''' 
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''  未検証　コメント作成日[2016/9/12 17:2]
    '''  
    ''' </remarks>
    <Category("01_農家個人条件")> <Description("最後の*は不要になりました。")> <PropertyOrderAttribute(1)> <TypeConverter(GetType(AutoCollectConverter))> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.KatakanaHalf)>
    Public Property フリガナ As String = ""

    ''' <summary>
    ''' 
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''  未検証　コメント作成日[2016/9/12 17:2]
    '''  
    ''' </remarks>
    <Category("01_農家個人条件")> <PropertyOrderAttribute(2)> <TypeConverter(GetType(AutoCollectConverter))> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Hiragana)>
    Public Property 氏名検索 As String = ""

    ''' <summary>
    ''' 
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''  未検証　コメント作成日[2016/9/12 17:2]
    '''  
    ''' </remarks>
    <Category("01_農家個人条件")> <PropertyOrderAttribute(3)> <TypeConverter(GetType(AutoCollectConverter))> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Hiragana)>
    Public Property 住所 As String = ""

    Private mvar行政区 As Integer = 0
    ''' <summary>
    ''' 
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''  未検証　コメント作成日[2016/9/12 17:7]
    '''  
    ''' </remarks>
    <Category("01_農家個人条件")> <PropertyOrderAttribute(4)> <TypeConverter(GetType(行政区ComboListConverter))> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Off)>
    Public Property 自治会 As String
        Get
            If IsDBNull(mvar行政区) Then
                Return "0000000:-"
            Else
                Dim pView As New DataView(App農地基本台帳.DataMaster.Body, "[Class]='行政区' AND [ID]=" & mvar行政区, "", DataViewRowState.CurrentRows)
                If pView.Count > 0 Then
                    Return String.Format("{0:D7}:{1}", mvar行政区, pView.Item(0).Item("名称"))
                Else
                    Return "0000000:-"
                End If
            End If
        End Get
        Set(ByVal value As String)
            mvar行政区 = Val(value)
        End Set
    End Property

    ''' <summary>
    ''' 
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''  未検証　コメント作成日[2016/9/12 17:7]
    '''  
    ''' </remarks>
    <Category("02_年金条件")> <PropertyOrderAttribute(5)>
    Public Property 農業者年金受給 As enum有無選 = 0
    ''' <summary>
    ''' 
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''  未検証　コメント作成日[2016/9/12 17:7]
    '''  
    ''' </remarks>
    <Category("02_年金条件")> <PropertyOrderAttribute(6)>
    Public Property 老齢年金受給 As enum有無選 = 0
    ''' <summary>
    ''' 
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''  未検証　コメント作成日[2016/9/12 17:7]
    '''  
    ''' </remarks>
    <Category("02_年金条件")> <PropertyOrderAttribute(7)>
    Public Property 経営移譲年金受給 As enum有無選 = 0


    ''' <summary>
    ''' 
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''  未検証　コメント作成日[2016/9/12 17:7]
    '''  
    ''' </remarks>
    <Category("03_経営条件")> <PropertyOrderAttribute(8)>
    Public Property 最低経営面積 As Decimal = 0


    ''' <summary>
    ''' 
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    '''  未検証　コメント作成日[2016/9/12 17:7]
    '''  
    ''' </remarks>

    Public Overrides Function ToString() As String
        Dim sB As New System.Text.StringBuilder
        Dim sAND As String = ""

        If IsNumeric(住民番号) AndAlso 住民番号 <> 0 Then
            sB.Append(sAND & String.Format("[ID] = {0}", 住民番号))
            SysAD.検索関連.AutoCollectStr("住民番号") = Me.住民番号.ToString
            sAND = " AND "
        End If

        If Trim(フリガナ).Length > 0 Then
            Dim sフリガナ検索 As String = Replace(Replace(Replace(Replace(Trim(Me.フリガナ), "　", " "), "  ", " ") & "*", "**", "*"), "*", "%")
            Dim sフリガナ検索2 As String = Replace(Replace(StrConv(sフリガナ検索, VbStrConv.Narrow), " ", ""), "ﾞ", "")

            sB.Append(sAND & String.Format("([フリガナ] LIKE '{0}' Or [検索フリガナ] LIKE '{1}')", sフリガナ検索, sフリガナ検索))

            SysAD.検索関連.AutoCollectStr("フリガナ") = Me.フリガナ
            sAND = " AND "
        End If

        If 氏名検索.Length > 0 Then
            sB.Append(sAND & Getワイルドカード文字("氏名", 氏名検索).Replace("*", "%"))
            SysAD.検索関連.AutoCollectStr("氏名検索") = Me.氏名検索
            sAND = " AND "
        End If

        If 住所.Length > 0 Then
            sB.Append(sAND & Getワイルドカード文字("住所", 住所).Replace("*", "%"))
            SysAD.検索関連.AutoCollectStr("住所") = Me.住所
            sAND = " AND "
        End If
        If mvar行政区 > 0 Then
            sB.Append(sAND & String.Format("[行政区ID] ={0}", mvar行政区))
            sAND = " AND "
        End If

        Select Case Me.農業者年金受給
            Case enum有無選.有り
                sB.Append(sAND & "[農年受給の有無] =True")
                sAND = " AND "
            Case enum有無選.無し
                sB.Append(sAND & "[農年受給の有無] =False")
                sAND = " AND "
        End Select

        Select Case Me.老齢年金受給
            Case enum有無選.有り
                sB.Append(sAND & "[老齢受給の有無] =True")
                sAND = " AND "
            Case enum有無選.無し
                sB.Append(sAND & "[老齢受給の有無] =False")
                sAND = " AND "
        End Select

        Select Case Me.経営移譲年金受給
            Case enum有無選.有り
                sB.Append(sAND & "[経営移譲の有無] =True")
                sAND = " AND "
            Case enum有無選.無し
                sB.Append(sAND & "[経営移譲の有無] =False")
                sAND = " AND "
        End Select

        If 最低経営面積 > 0 Then
            Dim objAcc As New 個人耕作面積修正()
            With objAcc
                .Dialog.StartProc(True, True)
                If .Dialog._objException Is Nothing = False Then
                    If .Dialog._objException.Message = "Cancel" Then
                        MsgBox("処理を中止しました。　", , "処理中止")
                    Else
                        'Throw objDlg._objException
                    End If
                End If
            End With

            '
            sB.Append(sAND & "[経営面積]>=" & 最低経営面積)
            sAND = " AND "
        End If

        Return Replace(sB.ToString(), "%%", "%")
    End Function

    ''' <summary>
    ''' 
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    '''  未検証　コメント作成日[2016/9/12 17:7]
    '''  
    ''' </remarks>
    Public Overrides Function View検索条件() As String
        Dim sB As New System.Text.StringBuilder
        Dim sAND As String = ""

        If IsNumeric(住民番号) AndAlso 住民番号 <> 0 Then
            sB.Append(sAND & String.Format("[ID] = {0}", 住民番号))
            sAND = " AND "
        End If

        If フリガナ = "*" Then

        ElseIf フリガナ.Length > 0 Then
            Dim sフリガナ検索 As String = Replace(Replace(Replace(Trim(Me.フリガナ), "　", " "), "  ", " ") & "*", "**", "*")
            Dim sフリガナ検索2 As String = Replace(Replace(StrConv(sフリガナ検索, VbStrConv.Narrow), " ", ""), "ﾞ", "")

            sB.Append(sAND & "(" & Getワイルドカード文字("フリガナ", sフリガナ検索) & " Or " & Getワイルドカード文字("検索フリガナ", sフリガナ検索) & " Or " & Getワイルドカード文字("検索フリガナ", sフリガナ検索2) & ")")
            sAND = " AND "
        End If

        If 氏名検索.Length > 0 Then
            sB.Append(sAND & String.Format("[氏名] LIKE '*{0}*'", 氏名検索))
            sAND = " AND "
        End If

        If 住所.Length > 0 Then
            sB.Append(sAND & Getワイルドカード文字("住所", 住所))
            sAND = " AND "
        End If
        If mvar行政区 > 0 Then
            sB.Append(sAND & String.Format("[行政区ID] ={0}", mvar行政区))
            sAND = " AND "
        End If

        Select Case Me.農業者年金受給
            Case enum有無選.有り
                sB.Append(sAND & "[農年受給の有無] =-1")
                sAND = " AND "
            Case enum有無選.無し
                sB.Append(sAND & "[農年受給の有無] =0")
                sAND = " AND "
        End Select

        Select Case Me.老齢年金受給
            Case enum有無選.有り
                sB.Append(sAND & "[老齢受給の有無] =-1")
                sAND = " AND "
            Case enum有無選.無し
                sB.Append(sAND & "[老齢受給の有無] =0")
                sAND = " AND "
        End Select

        Select Case Me.経営移譲年金受給
            Case enum有無選.有り
                sB.Append(sAND & "[経営移譲の有無] =-1")
                sAND = " AND "
            Case enum有無選.無し
                sB.Append(sAND & "[経営移譲の有無] =0")
                sAND = " AND "
        End Select

        If 最低経営面積 > 0 Then
            sB.Append(sAND & "[経営面積]>=" & 最低経営面積)
            sAND = " AND "
        End If


        Return Replace(sB.ToString(), "**", "*")
    End Function

    ''' <summary>
    ''' 
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <remarks>
    '''  未検証　コメント作成日[2016/9/12 17:7]
    '''  
    ''' </remarks>
    Public Sub New()

    End Sub
End Class

''' <summary>
''' 
''' </summary>
''' <remarks>
''' 未検証　件数[8]
''' </remarks>
<TypeConverter(GetType(PropertyOrderConverter))>
Public Class C農家検索条件
    Inherits Common検索条件

    ''' <summary>
    ''' 
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''  未検証　コメント作成日[2016/9/12 17:7]
    '''  
    ''' </remarks>
    <Category("01_世帯主条件")> <PropertyOrderAttribute(0)> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Off)> <TypeConverter(GetType(AutoCollectConverter))>
    Public Property 世帯番号 As Object = Nothing

    ''' <summary>
    ''' 
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''  未検証　コメント作成日[2016/9/12 17:7]
    '''  
    ''' </remarks>
    <Category("01_世帯主条件")> <PropertyOrderAttribute(1)> <TypeConverter(GetType(AutoCollectConverter))> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.KatakanaHalf)>
    Public Property 世帯主フリガナ As String = ""

    ''' <summary>
    ''' 
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''  未検証　コメント作成日[2016/9/12 17:7]
    '''  
    ''' </remarks>
    <Category("01_世帯主条件")> <PropertyOrderAttribute(2)> <TypeConverter(GetType(AutoCollectConverter))> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Hiragana)>
    Public Property 世帯主氏名検索 As String = ""

    ''' <summary>
    ''' 
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''  未検証　コメント作成日[2016/9/12 17:7]
    '''  
    ''' </remarks>
    <Category("01_世帯主条件")> <PropertyOrderAttribute(3)> <TypeConverter(GetType(AutoCollectConverter))> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Hiragana)>
    Public Property 世帯主住所 As String = ""

    ''' <summary>
    ''' 
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''  未検証　コメント作成日[2016/9/12 17:7]
    '''  
    ''' </remarks>
    <Category("03_経営条件")> <PropertyOrderAttribute(4)>
    Public Property 最低経営面積 As Decimal = 0

    ''' <summary>
    ''' 
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    '''  未検証　コメント作成日[2016/9/12 17:7]
    '''  
    ''' </remarks>

    Public Overrides Function ToString() As String
        Dim sB As New System.Text.StringBuilder
        Dim sAND As String = ""

        If 世帯番号 IsNot Nothing AndAlso IsNumeric(世帯番号) AndAlso 世帯番号 <> 0 Then
            sB.Append(sAND & String.Format("[D:世帯INFO].[ID] = {0}", 世帯番号))
            SysAD.検索関連.AutoCollectStr("世帯番号") = Me.世帯番号.ToString
            sAND = " AND "
        End If

        If 世帯主フリガナ.Length > 0 Then
            Dim sフリガナ検索 As String = Replace(Replace(Replace(Replace(Trim(Me.世帯主フリガナ), "　", " "), "  ", " ") & "*", "**", "*"), "*", "%")
            Dim sフリガナ検索2 As String = Replace(Replace(StrConv(sフリガナ検索, VbStrConv.Narrow), " ", ""), "ﾞ", "")

            sB.Append(sAND & String.Format("([フリガナ] LIKE '{0}' Or [検索フリガナ] LIKE '{1}' Or [検索フリガナ] LIKE '{2}')", sフリガナ検索, sフリガナ検索, sフリガナ検索2))

            SysAD.検索関連.AutoCollectStr("世帯主フリガナ") = Me.世帯主フリガナ
            sAND = " AND "
        End If

        If 世帯主氏名検索.Length > 0 Then
            sB.Append(sAND & Getワイルドカード文字("D:個人Info].[氏名", 世帯主氏名検索).Replace("*", "%"))
            SysAD.検索関連.AutoCollectStr("世帯主氏名検索") = Me.世帯主氏名検索
            sAND = " AND "
        End If

        If 世帯主住所.Length > 0 Then
            sB.Append(sAND & Getワイルドカード文字("D:個人Info].[住所", 世帯主住所).Replace("*", "%"))
            SysAD.検索関連.AutoCollectStr("世帯主住所") = Me.世帯主住所
            sAND = " AND "
        End If

        If 最低経営面積 > 0 Then
            Dim objAcc As New 世帯面積修正()
            With objAcc
                .Dialog.StartProc(True, True)
                If .Dialog._objException Is Nothing = False Then
                    If .Dialog._objException.Message = "Cancel" Then
                        MsgBox("処理を中止しました。　", , "処理中止")
                    Else
                        'Throw objDlg._objException
                    End If
                End If
            End With

            '
            sB.Append(sAND & "[総経営地]>=" & 最低経営面積)
            sAND = " AND "
        End If

        Return sB.ToString()
    End Function


    ''' <summary>
    ''' 
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <remarks>
    '''  未検証　コメント作成日[2016/9/12 17:7]
    '''  
    ''' </remarks>
    Public Sub New()
        'Me.フリガナ検索 = GetSetting("農地基本台帳", "農家検索条件", "フリガナ検索", "")
    End Sub

    ''' <summary>
    ''' 
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    '''  未検証　コメント作成日[2016/9/12 17:7]
    '''  
    ''' </remarks>
    Public Overrides Function View検索条件() As String
        Dim sB As New System.Text.StringBuilder
        Dim sAND As String = ""

        If IsNumeric(世帯番号) AndAlso 世帯番号 <> 0 Then
            sB.Append(sAND & String.Format("[ID] = {0}", 世帯番号))
            sAND = " AND "
        End If

        If 世帯主フリガナ.Length > 0 Then
            Dim sフリガナ検索 As String = Replace(Replace(Replace(Replace(Trim(Me.世帯主フリガナ), "　", " "), "  ", " ") & "*", "**", "*"), "*", "%")
            Dim sフリガナ検索2 As String = Replace(Replace(StrConv(sフリガナ検索, VbStrConv.Narrow), " ", ""), "ﾞ", "")
            sB.Append(sAND & "(" & Getワイルドカード文字("フリガナ", sフリガナ検索) & " Or " & Getワイルドカード文字("検索フリガナ", sフリガナ検索) & " Or " & Getワイルドカード文字("検索フリガナ", sフリガナ検索2) & ")")
            sAND = " AND "
        End If

        If 世帯主氏名検索.Length > 0 Then
            sB.Append(sAND & Getワイルドカード文字("世帯主氏名", 世帯主氏名検索))
            sAND = " AND "
        End If

        If 世帯主住所.Length > 0 Then
            sB.Append(sAND & Getワイルドカード文字("住所", 世帯主住所))

            sAND = " AND "
        End If

        If 最低経営面積 > 0 Then
            sB.Append(sAND & "[総経営地]>=" & 最低経営面積)
            sAND = " AND "
        End If

        Return sB.ToString()
    End Function
End Class

''' <summary>
''' 
''' </summary>
''' <remarks>
''' 未検証　件数[8]
''' </remarks>
<TypeConverter(GetType(PropertyOrderConverter))>
Public Class C住記録検索
    Inherits Common検索条件

    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:57]
    '''   
    '''  
    ''' </remarks>
    <Category("01_農家個人条件")> <PropertyOrderAttribute(0)> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Off)> <TypeConverter(GetType(AutoCollectConverter))>
    Public Property 住民番号 As String = ""

    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:57]
    '''   
    '''  
    ''' </remarks>
    <Category("01_農家個人条件")> <Description("最後の*は不要になりました。")> <PropertyOrderAttribute(1)> <TypeConverter(GetType(AutoCollectConverter))> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.KatakanaHalf)>
    Public Property フリガナ As String = ""

    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:57]
    '''   
    '''  
    ''' </remarks>
    <Category("01_農家個人条件")> <PropertyOrderAttribute(2)> <TypeConverter(GetType(AutoCollectConverter))> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Hiragana)>
    Public Property 氏名検索 As String = ""

    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:57]
    '''   
    '''  
    ''' </remarks>
    <Category("01_農家個人条件")> <PropertyOrderAttribute(3)> <TypeConverter(GetType(AutoCollectConverter))> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Hiragana)>
    Public Property 住所 As String = ""

    Private mvar行政区 As Integer = 0

    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:58]
    '''  
    ''' </remarks>
    <Category("01_農家個人条件")> <PropertyOrderAttribute(4)> <TypeConverter(GetType(行政区ComboListConverter))> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Off)>
    Public Property 自治会 As String
        Get
            If IsDBNull(mvar行政区) Then
                Return "0000000:-"
            Else
                Dim pView As New DataView(App農地基本台帳.DataMaster.Body, "[Class]='行政区' AND [ID]=" & mvar行政区, "", DataViewRowState.CurrentRows)
                If pView.Count > 0 Then
                    Return String.Format("{0:D7}:{1}", mvar行政区, pView.Item(0).Item("名称"))
                Else
                    Return "0000000:-"
                End If
            End If
        End Get
        Set(ByVal value As String)
            mvar行政区 = Val(value)
        End Set
    End Property



    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:58]
    '''   
    '''  
    ''' </remarks>
    Public Overrides Function ToString() As String
        Dim sB As New System.Text.StringBuilder
        Dim sAND As String = ""

        If IsNumeric(住民番号) AndAlso 住民番号 <> 0 Then
            sB.Append(sAND & String.Format("[ID] = {0}", 住民番号))
            SysAD.検索関連.AutoCollectStr("住民番号") = Me.住民番号.ToString
            sAND = " AND "
        End If

        If フリガナ.Length > 0 Then
            Dim sフリガナ検索 As String = Replace(Me.フリガナ, "　", " ")

            sB.Append(sAND & String.Format("[{0}] LIKE '{1}'", "フリガナ", Replace(sフリガナ検索 & "*", "**", "*")).Replace("*", "%"))

            SysAD.検索関連.AutoCollectStr("フリガナ") = Me.フリガナ
            sAND = " AND "
        End If

        If 氏名検索.Length > 0 Then
            sB.Append(sAND & Getワイルドカード文字("氏名", 氏名検索).Replace("*", "%"))
            SysAD.検索関連.AutoCollectStr("氏名検索") = Me.氏名検索
            sAND = " AND "
        End If

        If 住所.Length > 0 Then
            sB.Append(sAND & Getワイルドカード文字("住所", 住所).Replace("*", "%"))
            SysAD.検索関連.AutoCollectStr("住所") = Me.住所
            sAND = " AND "
        End If
        If mvar行政区 > 0 Then
            sB.Append(sAND & String.Format("[行政区ID] ={0}", mvar行政区))
            sAND = " AND "
        End If

        Return Replace(sB.ToString(), "%%", "%")
    End Function

    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:58]
    '''   
    '''  
    ''' </remarks>
    Public Overrides Function View検索条件() As String
        Dim sB As New System.Text.StringBuilder
        Dim sAND As String = ""

        If IsNumeric(住民番号) AndAlso 住民番号 <> 0 Then
            sB.Append(sAND & String.Format("[ID] = {0}", 住民番号))
            sAND = " AND "
        End If

        If フリガナ = "*" Then

        ElseIf フリガナ.Length > 0 Then
            Dim sフリガナ検索 As String = Replace(Me.フリガナ, "　", " ")

            sB.Append(sAND & Getワイルドカード文字("フリガナ", Replace(sフリガナ検索 & "*", "**", "*")))
            sAND = " AND "
        End If

        If 氏名検索.Length > 0 Then
            sB.Append(sAND & String.Format("[氏名] LIKE '*{0}*'", 氏名検索))
            sAND = " AND "
        End If

        If 住所.Length > 0 Then
            sB.Append(sAND & Getワイルドカード文字("住所", 住所))
            sAND = " AND "
        End If
        If mvar行政区 > 0 Then
            sB.Append(sAND & String.Format("[行政区ID] ={0}", mvar行政区))
            sAND = " AND "
        End If

        Return Replace(sB.ToString(), "**", "*")
    End Function

    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:58]
    '''   
    '''  
    ''' </remarks>
    Public Sub New()

    End Sub
End Class
''' <summary>
''' 
''' </summary>
''' <remarks>
''' 未検証　件数[7]
''' </remarks>
<TypeConverter(GetType(PropertyOrderConverter))>
Public Class C削除個人検索
    Inherits Common検索条件

    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:58]
    '''   
    '''  
    ''' </remarks>
    <Category("01_農家個人条件")> <PropertyOrderAttribute(0)> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Off)> <TypeConverter(GetType(AutoCollectConverter))>
    Public Property 住民番号 As String = ""

    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:58]
    '''   
    '''  
    ''' </remarks>
    <Category("01_農家個人条件")> <Description("最後の*は不要になりました。")> <PropertyOrderAttribute(1)> <TypeConverter(GetType(AutoCollectConverter))> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.KatakanaHalf)>
    Public Property フリガナ As String = ""

    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:58]
    '''   
    '''  
    ''' </remarks>
    <Category("01_農家個人条件")> <PropertyOrderAttribute(2)> <TypeConverter(GetType(AutoCollectConverter))> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Hiragana)>
    Public Property 氏名検索 As String = ""

    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:58]
    '''   
    '''  
    ''' </remarks>
    <Category("01_農家個人条件")> <PropertyOrderAttribute(3)> <TypeConverter(GetType(AutoCollectConverter))> <HimTools2012.controls.PropertyGridIMEAttribute(Windows.Forms.ImeMode.Hiragana)>
    Public Property 住所 As String = ""



    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:58]
    '''   
    '''  
    ''' </remarks>
    Public Overrides Function ToString() As String
        Dim sB As New System.Text.StringBuilder
        Dim sAND As String = ""

        If IsNumeric(住民番号) AndAlso 住民番号 <> 0 Then
            sB.Append(sAND & String.Format("[ID] = {0}", 住民番号))
            SysAD.検索関連.AutoCollectStr("住民番号") = Me.住民番号.ToString
            sAND = " AND "
        End If

        If フリガナ.Length > 0 Then
            Dim sフリガナ検索 As String = Replace(Me.フリガナ, "　", " ")

            sB.Append(sAND & String.Format("[{0}] LIKE '{1}'", "フリガナ", Replace(sフリガナ検索 & "*", "**", "*")).Replace("*", "%"))

            SysAD.検索関連.AutoCollectStr("フリガナ") = Me.フリガナ
            sAND = " AND "
        End If

        If 氏名検索.Length > 0 Then
            sB.Append(sAND & Getワイルドカード文字("氏名", 氏名検索).Replace("*", "%"))
            SysAD.検索関連.AutoCollectStr("氏名検索") = Me.氏名検索
            sAND = " AND "
        End If

        If 住所.Length > 0 Then
            sB.Append(sAND & Getワイルドカード文字("住所", 住所).Replace("*", "%"))
            SysAD.検索関連.AutoCollectStr("住所") = Me.住所
            sAND = " AND "
        End If

        Return Replace(sB.ToString(), "%%", "%")
    End Function

    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:58]
    '''   
    '''  
    ''' </remarks>
    Public Overrides Function View検索条件() As String
        Dim sB As New System.Text.StringBuilder
        Dim sAND As String = ""

        If IsNumeric(住民番号) AndAlso 住民番号 <> 0 Then
            sB.Append(sAND & String.Format("[ID] = {0}", 住民番号))
            sAND = " AND "
        End If

        If フリガナ = "*" Then

        ElseIf フリガナ.Length > 0 Then
            Dim sフリガナ検索 As String = Replace(Me.フリガナ, "　", " ")

            sB.Append(sAND & Getワイルドカード文字("フリガナ", Replace(sフリガナ検索 & "*", "**", "*")))
            sAND = " AND "
        End If

        If 氏名検索.Length > 0 Then
            sB.Append(sAND & String.Format("[氏名] LIKE '*{0}*'", 氏名検索))
            sAND = " AND "
        End If

        If 住所.Length > 0 Then
            sB.Append(sAND & Getワイルドカード文字("住所", 住所))
            sAND = " AND "
        End If

        Return Replace(sB.ToString(), "**", "*")
    End Function

    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:58]
    '''   
    '''  
    ''' </remarks>
    Public Sub New()

    End Sub
End Class

''' <summary>
''' 
''' </summary>
''' <remarks>
''' 未検証　件数[2]
''' </remarks>
Public Class 世帯面積修正
    Inherits HimTools2012.clsAccessor
    Private mvarNID As Integer = 0
    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <param name="nID"></param>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:58]
    '''   
    '''  
    ''' </remarks>
    Public Sub New(Optional nID As Object = Nothing)
        If nID IsNot Nothing Then
            mvarNID = nID
        End If
    End Sub

    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:59]
    '''   
    '''  
    ''' </remarks>
    Public Overrides Sub Execute()
        Message = "データ取り込み中.."
        Dim pTBL As DataTable
        Dim pArea As Integer = 0
        If mvarNID = 0 Then
            pTBL = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:世帯Info].ID, V_農地.耕作世帯ID, Sum([V_農地].[田面積]+[V_農地].[畑面積]) AS 農地計, [D:世帯Info].総経営地 FROM V_農地 RIGHT JOIN [D:世帯Info] ON V_農地.耕作世帯ID = [D:世帯Info].ID GROUP BY [D:世帯Info].ID, V_農地.耕作世帯ID, [D:世帯Info].総経営地 HAVING ((([D:世帯Info].ID)>0));")
        Else
            pTBL = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:世帯Info].ID, V_農地.耕作世帯ID, Sum([V_農地].[田面積]+[V_農地].[畑面積]) AS 農地計, [D:世帯Info].総経営地 FROM V_農地 RIGHT JOIN [D:世帯Info] ON V_農地.耕作世帯ID = [D:世帯Info].ID GROUP BY [D:世帯Info].ID, V_農地.耕作世帯ID, [D:世帯Info].総経営地 HAVING ((([D:世帯Info].ID)=" & mvarNID & "));")
        End If

        Message = "総経営地修正中.."
        Maximum = pTBL.Rows.Count
        Dim nCount As Integer = 0
        Dim sSQL As New System.Text.StringBuilder

        For Each pRow As DataRow In pTBL.Rows
            If Not Val(pRow.Item("農地計").ToString) = Val(pRow.Item("総経営地").ToString) Then
                sSQL.AppendLine("UPDATE [D:世帯Info] SET [総経営地]=" & Val(pRow.Item("農地計").ToString) & " WHERE [ID]=" & pRow.Item("ID"))
                pArea = Val(pRow.Item("農地計").ToString)
                If sSQL.Length > 1024 Then
                    SysAD.DB(sLRDB).ExecuteSQL(sSQL.ToString)
                    sSQL.Clear()
                End If
                Me.Value = nCount
                My.Application.DoEvents()
            End If
            nCount += 1
        Next

        If sSQL.Length > 0 Then
            SysAD.DB(sLRDB).ExecuteSQL(sSQL.ToString)
            sSQL.Clear()
        End If

        If mvarNID <> 0 Then
            Dim pRow As DataRow = App農地基本台帳.TBL世帯.FindRowByID(mvarNID)
            If pRow IsNot Nothing Then
                pRow.Item("総経営地") = 123
            End If
        End If

        Me.Maximum = 0
    End Sub
End Class

''' <summary>
''' 
''' </summary>
''' <remarks>
''' 未検証　件数[1]
''' </remarks>
Public Class 個人耕作面積修正
    Inherits HimTools2012.clsAccessor

    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:59]
    '''   
    '''  
    ''' </remarks>
    Public Overrides Sub Execute()
        Message = "データ取り込み中.."
        SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:個人Info] SET [経営面積]=0 WHERE [経営面積] Is Null")
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:個人Info].ID, Sum([V_農地].[田面積]+[V_農地].[畑面積]) AS 農地計, [D:個人Info].経営面積 FROM [D:個人Info] LEFT JOIN V_農地 ON [D:個人Info].ID = V_農地.耕作者ID GROUP BY [D:個人Info].ID, [D:個人Info].経営面積 HAVING ((([D:個人Info].ID)>0));")

        Message = "経営面積修正中.."
        Maximum = pTBL.Rows.Count
        Dim nCount As Integer = 0
        Dim sSQL As New System.Text.StringBuilder

        For Each pRow As DataRow In pTBL.Rows
            If Not CInt(Val(pRow.Item("農地計").ToString)) = Val(pRow.Item("経営面積").ToString) Then
                sSQL.AppendLine("UPDATE [D:個人Info] SET [経営面積]=" & CInt(Val(pRow.Item("農地計").ToString)) & " WHERE [ID]=" & pRow.Item("ID"))
                If sSQL.Length > 1024 Then
                    SysAD.DB(sLRDB).ExecuteSQL(sSQL.ToString)
                    sSQL.Clear()
                End If
                Me.Value = nCount
                My.Application.DoEvents()
            End If
            nCount += 1
        Next

        If sSQL.Length > 0 Then
            SysAD.DB(sLRDB).ExecuteSQL(sSQL.ToString)
            sSQL.Clear()
        End If

        Me.Maximum = 0
    End Sub
End Class

''' <summary>
''' 
''' </summary>
''' <remarks>
''' 未検証　件数[2]
''' </remarks>
Public Class 行政区ComboListConverter
    Inherits HimTools2012.InputSupport.ComboListConverter

    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:59]
    '''   
    '''  
    ''' </remarks>
    Public Sub New()
        MyBase.New(New DataView(App農地基本台帳.DataMaster.Body, "Class='行政区'", "ID", DataViewRowState.CurrentRows))
    End Sub

    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:59]
    '''   
    '''  
    ''' </remarks>
    Public Overrides Function GetFilter() As String
        Return "Class='行政区'"
    End Function
End Class

''' <summary>
''' 
''' </summary>
''' <remarks>
''' 未検証　件数[4]
''' </remarks>
Public Class C検索関連
    Private mvarMuster As DataTable

    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:59]
    '''   
    '''  
    ''' </remarks>
    Public Property AutoCollectStr(ByVal s検索名 As String) As String
        Get
            Dim pRow As DataRow = mvarMuster.Rows.Find(s検索名)
            If pRow IsNot Nothing AndAlso Not IsDBNull(pRow("検索文字列")) Then
                Return pRow("検索文字列").ToString
            Else
                Return ""
            End If
        End Get
        Set(value As String)
            If value IsNot Nothing AndAlso value.Length > 0 Then
                Dim pRow As DataRow = mvarMuster.Rows.Find(s検索名)
                If pRow IsNot Nothing AndAlso Not IsDBNull(pRow("検索文字列")) Then
                    pRow.Item("検索文字列") = MakeCompleteString(pRow.Item("検索文字列"), value)
                Else
                    pRow = mvarMuster.NewRow
                    pRow.Item("検索名") = s検索名
                    pRow.Item("検索文字列") = MakeCompleteString("", value)
                    mvarMuster.Rows.Add(pRow)
                End If

                SaveFile()
            End If
        End Set
    End Property

    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <param name="sList"></param>
    ''' <param name="sData"></param>
    ''' <returns></returns>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:59]
    '''   
    '''  
    ''' </remarks>
    Private Function MakeCompleteString(ByVal sList As String, ByVal sData As String) As String
        Dim sRet As String = (";" & sList & ";").Replace(";" & sData & ";", ";")

        If sRet.StartsWith(";") Then
            sRet = sRet.Substring(1)
        End If
        If sRet.EndsWith(";") Then
            sRet = sRet.Substring(0, sRet.Length - 1)
        End If
        Dim sRets() As String
        Do
            sRets = sRet.Split(";")
            If UBound(sRets) > 10 Then
                sRet = sRet.Substring(0, sRet.LastIndexOf(";") - 1)
                sRets = sRet.Split(";")
            End If
        Loop Until UBound(sRets) <= 10


        Return sData & IIf(sRet.Length > 0, ";" & sRet, "")
    End Function

    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <param name="pSys"></param>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:59]
    '''   
    '''  
    ''' </remarks>
    Public Sub New(pSys As HimTools2012.System管理.CSystemBase)
        Dim sFName As String = pSys.SystemInfo.LocalDataPath & "\検索Param.xml"
        If IO.File.Exists(sFName) Then
            Try
                mvarMuster = New DataTable("検索Param")
                mvarMuster.ReadXml(sFName)
            Catch ex As Exception
                Stop
            End Try
        Else
            mvarMuster = New DataTable("検索Param")
            mvarMuster.Columns.Add(New DataColumn("検索名", GetType(String)))
            mvarMuster.Columns.Add(New DataColumn("検索文字列", GetType(String)))
            mvarMuster.PrimaryKey = New DataColumn() {mvarMuster.Columns("検索名")}
            mvarMuster.WriteXml(sFName, System.Data.XmlWriteMode.WriteSchema)
        End If
    End Sub
    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:59]
    '''   
    '''  
    ''' </remarks>
    Public Sub SaveFile()
        Dim sFName As String = SysAD.SystemInfo.LocalDataPath & "\検索Param.xml"
        mvarMuster.WriteXml(sFName, System.Data.XmlWriteMode.WriteSchema)
    End Sub
End Class


''' <summary>
''' 
''' </summary>
''' <remarks>
''' 未検証　件数[6]
''' </remarks>
Public Class AutoCollectConverter
    Inherits StringConverter

    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:59]
    '''   
    '''  
    ''' </remarks>
    Public Sub New()

    End Sub

    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <param name="context"></param>
    ''' <param name="propertyValues"></param>
    ''' <returns></returns>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:59]
    '''   
    '''  
    ''' </remarks>
    Public Overrides Function CreateInstance(context As System.ComponentModel.ITypeDescriptorContext, propertyValues As System.Collections.IDictionary) As Object
        Dim pOBJ As Object = MyBase.CreateInstance(context, propertyValues)
        Return pOBJ
    End Function
    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <param name="context"></param>
    ''' <returns></returns>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:59]
    '''   
    '''  
    ''' </remarks>
    Public Overloads Overrides Function GetStandardValuesSupported(ByVal context As ITypeDescriptorContext) As Boolean
        Return True
    End Function

    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <param name="context"></param>
    ''' <returns></returns>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:59]
    '''   
    '''  
    ''' </remarks>
    Public Overloads Overrides Function GetStandardValues(ByVal context As ITypeDescriptorContext) As StandardValuesCollection

        Return New StandardValuesCollection(SysAD.検索関連.AutoCollectStr(context.PropertyDescriptor.DisplayName).Split(";"))
    End Function
    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <param name="context"></param>
    ''' <returns></returns>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:59]
    '''   
    '''  
    ''' </remarks>
    Public Overloads Overrides Function GetStandardValuesExclusive(ByVal context As ITypeDescriptorContext) As Boolean
        Return False
    End Function

    ''' <summary>
    '''  
    '''  
    ''' Todo 検証を完了してください
    ''' </summary>
    ''' <param name="context"></param>
    ''' <param name="culture"></param>
    ''' <param name="value"></param>
    ''' <param name="destinationType"></param>
    ''' <returns></returns>
    ''' <remarks>
    '''   未検証　コメント作成日[2016/9/12 16:59]
    '''   
    '''  
    ''' </remarks>
    Public Overrides Function ConvertTo(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal culture As System.Globalization.CultureInfo, ByVal value As Object, ByVal destinationType As System.Type) As Object
        Return MyBase.ConvertTo(context, culture, value, destinationType)
    End Function

End Class

