Imports System.ComponentModel

Public MustInherit Class CInputDate
    Public MustOverride ReadOnly Property DataValidate As Boolean

    Public Event ValueError(ByVal s As Object, ByVal sPropertyName As String, ByVal sMessage As String, ByVal NewValue As Object)

    Public Sub ErrorEvent(ByVal pObject As Object, ByVal sPropeertyName As String, ByVal sMessage As String, ByVal NewValue As Object)
        RaiseEvent ValueError(pObject, sPropeertyName, sMessage, NewValue)
    End Sub

    <TypeConverter(GetType(C範囲入力ClassConverter))>
    Public Class C範囲入力
        Private _範囲開始 As Date = Now.Date
        Private _範囲終了 As Date = Now.Date

        <ReadOnlyAttribute(False)>
        Public Property 範囲開始() As Date
            Get
                Return _範囲開始
            End Get
            Set(ByVal Value As Date)
                _範囲開始 = Value
            End Set
        End Property

        <ReadOnlyAttribute(False)>
        Public Property 範囲終了() As Date
            Get
                Return _範囲終了
            End Get
            Set(ByVal Value As Date)
                _範囲終了 = Value
            End Set
        End Property

        <ReadOnlyAttribute(True)>
        Public ReadOnly Property 日数() As Integer
            Get
                Try
                    Dim pBit As TimeSpan = _範囲終了.Subtract(_範囲開始)
                    Return pBit.TotalDays + 1
                Catch ex As Exception
                    Return 0
                End Try
            End Get
        End Property

        Public Overrides Function ToString() As String
            If _範囲開始.Year = _範囲終了.Year AndAlso _範囲開始.Month = _範囲終了.Month AndAlso _範囲開始.Day = _範囲終了.Day Then
                Return Strings.Format(_範囲開始, "MM月dd日")
            ElseIf _範囲開始.Year = _範囲終了.Year AndAlso _範囲開始.Month = _範囲終了.Month Then
                Return Strings.Format(_範囲開始, "MM月dd日") & "～" & Strings.Format(_範囲終了, "dd日")
            ElseIf _範囲開始.Year = _範囲終了.Year Then
                Return Strings.Format(_範囲開始, "MM月dd日") & "～" & Strings.Format(_範囲終了, "MM月dd日")
            Else
                Return Strings.Format(_範囲開始, "yyyy年MM月dd日") & "," & Strings.Format(_範囲終了, "yyyy年MM月dd日")
            End If
        End Function
    End Class

    Public Class C範囲入力ClassConverter
        Inherits ExpandableObjectConverter

        Public Overloads Overrides Function CanConvertTo(ByVal context As ITypeDescriptorContext, ByVal destinationType As Type) As Boolean
            If destinationType Is GetType(C範囲入力) Then
                Return True
            End If
            Return MyBase.CanConvertTo(context, destinationType)
        End Function

        Public Overloads Overrides Function ConvertTo(ByVal context As ITypeDescriptorContext, ByVal culture As System.Globalization.CultureInfo, ByVal value As Object, ByVal destinationType As Type) As Object
            If destinationType Is GetType(String) And TypeOf value Is C範囲入力 Then
                Dim cc As C範囲入力 = CType(value, C範囲入力)
                Return cc.ToString
            End If
            Return MyBase.ConvertTo(context, culture, value, destinationType)
        End Function

        Public Overloads Overrides Function CanConvertFrom(ByVal context As ITypeDescriptorContext, ByVal sourceType As Type) As Boolean
            If sourceType Is GetType(String) Then
                Return True
            End If
            Return MyBase.CanConvertFrom(context, sourceType)
        End Function

        Public Overloads Overrides Function ConvertFrom(ByVal context As ITypeDescriptorContext, ByVal culture As System.Globalization.CultureInfo, ByVal value As Object) As Object
            If TypeOf value Is String Then

                Try
                    Dim ss As String() = Split(value, "～")
                    Dim cc As New C範囲入力

                    Select Case UBound(ss)
                        Case 0

                    End Select

                    cc.範囲開始 = CDate(ss(0))
                    cc.範囲終了 = CDate(ss(1))
                    Return cc

                Catch ex As Exception
                    Return Nothing
                End Try
            End If
            Return MyBase.ConvertFrom(context, culture, value)
        End Function
    End Class

End Class
