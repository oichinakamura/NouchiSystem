Imports System.ComponentModel

Imports System.Windows.Forms.Design
Imports System.Drawing.Design
Imports HimTools2012.TypeConverterCustom
Imports HimTools2012.controls.PropertyGridSupport


<TypeConverter(GetType(PropertyOrderConverter))>
Public Class CInput農地
    Inherits InputObjectParam

    Public Sub New()

        検索Input.mvar検索小字View = New DataView(App農地基本台帳.DataMaster.Body, "Class='小字' AND [nParam]=0", "ID", DataViewRowState.CurrentRows)
    End Sub

    Private mvar大字 As Integer = 0
    Private mvar小字 As Integer = 0

    <Category("農地条件")> <PropertyOrderAttribute(0)> <Editor(GetType(大字ComboItemsConverter), GetType(UITypeEditor))>
    Public Property 大字 As Object
        Get
            Return SysAD.MasterFind("大字", mvar大字)
        End Get
        Set(ByVal value As Object)
            mvar大字 = Val(value)
            If 検索Input.mvar大字Code <> mvar大字 Then
                検索Input.mvar大字Code = Val(value)
                検索Input.mvar検索小字View.RowFilter = "Class='小字' AND [nParam]=" & 検索Input.mvar大字Code
                For Each pRow As DataRowView In 検索Input.mvar検索小字View
                    If pRow.Item("ID") = mvar小字 Then
                        Return
                    End If
                Next
                mvar小字 = 0
            End If
        End Set
    End Property

    <Category("農地条件")> <PropertyOrderAttribute(1)> <Editor(GetType(小字ComboItemsConverter), GetType(UITypeEditor))>
    Public Property 小字 As Object
        Get
            Return SysAD.MasterFind("小字", mvar小字)
        End Get
        Set(ByVal value As Object)
            mvar小字 = Val(value)
        End Set
    End Property

    <Category("農地条件")> <PropertyOrderAttribute(2)>
    Public Property 市外地の所在 As String

    <Category("農地条件")> <PropertyOrderAttribute(3)>
    Public Property 地番 As String

    <Category("所有条件")> <ReadOnlyAttribute(True)> <PropertyOrderAttribute(4)>
    Public Property 所有世帯ID As Long
    <Category("所有条件")> <ReadOnlyAttribute(True)> <PropertyOrderAttribute(5)>
    Public Property 所有者ID As Long

    Public Overrides Function CheckValues() As Boolean
        If 地番 IsNot Nothing AndAlso 地番.Length > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Overrides Sub SetUpdateRow(ByRef pUpdateRow As HimTools2012.Data.UpdateRow)
        pUpdateRow.SetValue("大字ID", mvar大字)
        pUpdateRow.SetValue("小字ID", mvar小字)
        pUpdateRow.SetValue("所在", 市外地の所在)

        pUpdateRow.SetValue("地番", Me.地番)

        pUpdateRow.SetValue("所有世帯ID", Me.所有世帯ID)
        pUpdateRow.SetValue("所有者ID", Me.所有者ID)

        pUpdateRow.SetValue("自小作別", 0)
    End Sub

    Public Overrides Function AddRecord() As Long
        Dim pNewRow As DataRow = App農地基本台帳.TBL農地.NewRow
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Min([ID]) AS MinID FROM [D:農地Info];")
        Dim p転用 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Min([ID]) AS MinID FROM [D_転用農地];")
        Dim p削除 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Min([ID]) AS MinID FROM [D_削除農地];")

        If pTBL.Rows.Count > 0 Then
            Dim nID As Integer = Val(pTBL.Rows(0).Item("MinID").ToString) - 1
            If nID >= 0 Then
                nID = -1
            End If
            If p転用.Rows.Count > 0 Then
                Dim TID As Integer = Val(p転用.Rows(0).Item("MinID").ToString) - 1
                If nID > TID Then
                    nID = TID
                End If
            End If
            If p削除.Rows.Count > 0 Then
                Dim DID As Integer = Val(p削除.Rows(0).Item("MinID").ToString) - 1
                If nID > DID Then
                    nID = DID
                End If
            End If

            Do Until Replace(SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D:農地Info]([ID]) VALUES(" & nID & ")"), "OK", "") = ""
                nID = nID - 1
            Loop

            Dim mvarUpdateRow As New HimTools2012.Data.UpdateRow(pNewRow, HimTools2012.Data.UPDateMode.AutoUpdate)
            Me.SetUpdateRow(mvarUpdateRow)
            App農地基本台帳.TBL農地.Update(mvarUpdateRow, False)

            pNewRow.Item("ID") = nID
            App農地基本台帳.TBL農地.Rows.Add(pNewRow)
            Return nID
        End If
        Return 0
    End Function
End Class

#Region "CodeConverter"

Public Class 地目Converter
    Inherits HimTools2012.TypeConverterCustom.AddAbleEnumConverter
    Public Sub New()
        MyBase.New(New DataView(App農地基本台帳.DataMaster.Body, "Class='地目'", "ID", DataViewRowState.CurrentRows), "名称")
    End Sub
End Class
Public Class 現況地目Converter
    Inherits HimTools2012.TypeConverterCustom.AddAbleEnumConverter
    Public Sub New()
        MyBase.New(New DataView(App農地基本台帳.DataMaster.Body, "Class='課税地目'", "ID", DataViewRowState.CurrentRows), "名称")
    End Sub
End Class
Public Class 農委地目Converter
    Inherits HimTools2012.TypeConverterCustom.AddAbleEnumConverter
    Public Sub New()
        MyBase.New(New DataView(App農地基本台帳.DataMaster.Body, "Class='農委地目'", "ID", DataViewRowState.CurrentRows), "名称")
    End Sub
End Class
Public Class 農地状況Converter
    Inherits HimTools2012.TypeConverterCustom.AddAbleEnumConverter
    Public Sub New()
        MyBase.New(New DataView(App農地基本台帳.DataMaster.Body, "Class='農地状況'", "ID", DataViewRowState.CurrentRows), "名称")
    End Sub
End Class
Public Class 農振区分Converter
    Inherits HimTools2012.TypeConverterCustom.AddAbleEnumConverter
    Public Sub New()
        MyBase.New(New DataView(App農地基本台帳.DataMaster.Body, "Class='農振区分'", "ID", DataViewRowState.CurrentRows), "名称")
    End Sub
End Class


Public Class 大字Converter
    Inherits HimTools2012.TypeConverterCustom.AddAbleEnumConverter
    Public Sub New()
        MyBase.New(New DataView(App農地基本台帳.DataMaster.Body, "Class='大字'", "ID", DataViewRowState.CurrentRows), "名称")
    End Sub
End Class

Public Class 小字Converter
    Inherits HimTools2012.TypeConverterCustom.AddAbleEnumConverter
    Public Sub New()
        MyBase.New(New DataView(App農地基本台帳.DataMaster.Body, "Class='小字'", "ID", DataViewRowState.CurrentRows), "名称")
    End Sub
End Class

Public Module 検索Input
    Public mvar大字Code As Integer = 0
    Public mvar検索小字View As DataView
End Module

Public MustInherit Class InputObjectParam
    MustOverride Function CheckValues() As Boolean
    MustOverride Sub SetUpdateRow(ByRef pUpdateRow As HimTools2012.Data.UpdateRow)
    MustOverride Function AddRecord() As Long
End Class



Public Class 小字ComboItemsConverter
    Inherits ComboBoxEditor

    Public Sub New()
        MyBase.New(検索Input.mvar検索小字View)
    End Sub

End Class

Public Class 大字ComboItemsConverter
    Inherits ComboBoxEditor

    Public Sub New()
        MyBase.New(New DataView(App農地基本台帳.DataMaster.Body, "Class='大字'", "ID", DataViewRowState.CurrentRows))
    End Sub

End Class


<Security.Permissions.PermissionSet(Security.Permissions.SecurityAction.Demand, Name:="FullTrust")>
Public Class ComboBoxEditor
    Inherits UITypeEditor
    Private WithEvents _editorUi As ListBox = Nothing
    Dim editservice As IWindowsFormsEditorService

    Public Sub New(ByVal pView As DataView)
        _editorUi = New ListBox
        _editorUi.DataSource = pView
        _editorUi.DisplayMember = "名称"
        _editorUi.ValueMember = "ID"
    End Sub

    Public Overloads Overrides Function GetEditStyle(ByVal context As System.ComponentModel.ITypeDescriptorContext) As System.Drawing.Design.UITypeEditorEditStyle
        Return UITypeEditorEditStyle.DropDown
    End Function


    Public Overloads Overrides Function EditValue(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal provider As System.IServiceProvider, ByVal value As Object) As Object
        editservice = provider.GetService(GetType(IWindowsFormsEditorService))

        editservice.DropDownControl(_editorUi)

        If _editorUi.SelectedItem IsNot Nothing Then
            Return CType(_editorUi.SelectedItem, DataRowView).Item("ID")
        Else
            Return Nothing
        End If

    End Function

    Private Sub _editorUi_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles _editorUi.DoubleClick
        If editservice IsNot Nothing Then
            editservice.CloseDropDown()
        End If
    End Sub
End Class
#End Region

