
Imports System.ComponentModel
Imports HimTools2012.TypeConverterCustom
Imports HimTools2012.controls.PropertyGridSupport

Module mod個人
    Public Sub 世帯追加(p個人 As CObj個人)
        Dim nID As Integer = 0
        If p個人.GetIntegerValue("世帯ID") = 0 Then
            Dim nTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Min([ID]) AS [最小値] FROM [D:世帯Info]")
            nID = nTBL.Rows(0).Item("最小値") - 1
            Dim St As String = InputBox("世帯番号を入力してください", "世帯の追加", nID)

            nID = Val(St)
        Else
            nID = p個人.GetIntegerValue("世帯ID")
        End If

        If nID <> 0 Then
            Try
                Dim pRow As DataRow = App農地基本台帳.TBL世帯.FindRowByID(nID)

                If pRow Is Nothing Then
                    SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D:世帯Info]([ID],[世帯主ID]) VALUES(" & nID & "," & p個人.ID & ")")
                    Dim pRowN As DataRow = App農地基本台帳.TBL世帯.FindRowByID(nID)
                Else
                    MsgBox("指定された番号の世帯はすでに存在します。", vbExclamation)
                End If

                p個人.ValueChange("世帯ID", nID)
                p個人.SaveMyself()
                Open世帯(p個人.世帯ID, "")
            Catch ex As Exception
            End Try
        End If

    End Sub
    Public Sub 個人削除(ByRef p個人 As CObj個人)
        If MsgBox("本当に削除しますか", vbYesNo) = vbYes Then
            p個人.DoCommand("閉じる")
            SysAD.DB(sLRDB).ExecuteSQL("DELETE FROM [D_削除個人] WHERE [ID]=" & p個人.ID)
            SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO D_削除個人 SELECT [D:個人Info].* FROM [D:個人Info] WHERE ((([D:個人Info].ID)=" & p個人.ID & "));")
            SysAD.DB(sLRDB).ExecuteSQL("DELETE FROM [D:個人Info] WHERE [ID]=" & p個人.ID)
            App農地基本台帳.TBL個人.Rows.Remove(p個人.Row.Body)
            p個人.RowClear()
        End If
    End Sub
    Public Sub 個人復元(ByRef p削除個人 As CObj各種削除)
        If MsgBox("本当に削除しますか", vbYesNo) = vbYes Then
            p削除個人.DoCommand("閉じる")
            SysAD.DB(sLRDB).ExecuteSQL("DELETE FROM [D:個人Info] WHERE [ID]=" & p削除個人.ID)
            SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D:個人Info] SELECT [D_削除個人].* FROM [D_削除個人] WHERE ((([D_削除個人].ID)=" & p削除個人.ID & "));")
            SysAD.DB(sLRDB).ExecuteSQL("DELETE FROM [D_削除個人] WHERE [ID]=" & p削除個人.ID)
            App農地基本台帳.TBL削除個人.Rows.Remove(p削除個人.Row.Body)
            p削除個人.RowClear()
        End If
    End Sub

End Module


<TypeConverter(GetType(PropertyOrderConverter))>
Public Class 住民追加条件
    Inherits InputObjectParam

    Private mvarID As Integer


    Public Sub New()

    End Sub

    <Category("追加条件")> <PropertyOrderAttribute(0)>
    Public Property 住民番号 As Integer
        Get
            Return mvarID
        End Get
        Set(ByVal value As Integer)
            mvarID = value
        End Set
    End Property
    Private mvar氏名 As String
    <Category("追加条件")> <PropertyOrderAttribute(1)>
    Public Property 氏名 As String
        Get
            Return mvar氏名
        End Get
        Set(ByVal value As String)
            mvar氏名 = value
        End Set
    End Property
    Private mvarフリガナ As String
    <Category("追加条件")> <PropertyOrderAttribute(2)>
    Public Property フリガナ As String
        Get
            Return mvarフリガナ
        End Get
        Set(ByVal value As String)
            mvarフリガナ = value
        End Set
    End Property


    Public Overrides Function AddRecord() As Long
        Return 0
    End Function

    Public Overrides Function CheckValues() As Boolean
        If 住民番号 <> 0 AndAlso mvar氏名 IsNot Nothing AndAlso mvar氏名.Length > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Overrides Sub SetUpdateRow(ByRef pUpdateRow As HimTools2012.Data.UpdateRow)

    End Sub
End Class
