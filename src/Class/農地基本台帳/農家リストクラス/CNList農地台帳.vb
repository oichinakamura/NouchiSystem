
Imports HimTools2012.TabPages

Public Class GridViewType
    Inherits System.Attribute
    Private mvarType As enmColumnType
    Public Sub New(pCtrl As enmColumnType)
        mvarType = pCtrl
    End Sub
    Public ReadOnly Property CellType As enmColumnType
        Get
            Return mvarType
        End Get
    End Property
End Class


Public Class CNList農地台帳
    Inherits HimTools2012.TabPages.NListSK
    Protected WithEvents mvar検索Page As CPage検索

    Public Overrides ReadOnly Property 検索Page As HimTools2012.TabPages.CPage検索SK
        Get
            Return mvar検索Page
        End Get
    End Property

    Public Sub New(ByVal sText As String, ByVal sName As String, Optional ByVal bCloseable As Boolean = False)
        MyBase.New(bCloseable, sName, sText, ObjectMan, SysAD.ImageList48, True)
    End Sub

    Public Sub New(ByRef pNode As Xml.XmlNode, ByRef pLayout As HimTools2012.controls.XMLLayout)
        MyBase.New(pNode, pLayout)
    End Sub
    Public Sub SetGridColumn(ByVal t As Type)

        mvarGrid.AddButtonColumn("更新", "更新", "更新")

        For Each p As System.Reflection.PropertyInfo In t.GetProperties
            Dim pDiff As New ColumnDef(p.Name, p.PropertyType)

            If p.Name = "更新" Then
            Else
                For Each pAttr As Object In p.GetCustomAttributes(False)
                    Select Case TypeName(pAttr)
                        Case "BrowsableAttribute" : pDiff.Browsable = CType(pAttr, System.ComponentModel.BrowsableAttribute).Browsable
                        Case "CategoryAttribute" : pDiff.CategoryAttribute = CType(pAttr, System.ComponentModel.CategoryAttribute).Category

                        Case "DisplayNameAttribute" : pDiff.DisplayName = CType(pAttr, System.ComponentModel.DisplayNameAttribute).DisplayName
                        Case "PropertyOrderAttribute"
                        Case "ReadOnlyAttribute" : pDiff.ReadOnlyAttr = CType(pAttr, System.ComponentModel.ReadOnlyAttribute).IsReadOnly
                        Case "PropertyGridIMEAttribute"
                        Case "TypeConverterAttribute"
                            Select Case pDiff.Name
                                Case "大字"
                                    pDiff.ColumnType = enmColumnType.Combo
                                    pDiff.DataSource = New DataView(App農地基本台帳.DataMaster.Body, "Class='大字'", "[ID]", DataViewRowState.CurrentRows)
                                Case "小字"
                                    pDiff.ColumnType = enmColumnType.Combo
                                    pDiff.DataSource = New DataView(App農地基本台帳.DataMaster.Body, "Class='小字'", "[ID]", DataViewRowState.CurrentRows)
                                Case "行政区"
                                    pDiff.ColumnType = enmColumnType.Combo
                                    pDiff.DataSource = New DataView(App農地基本台帳.DataMaster.Body, "Class='行政区'", "[ID]", DataViewRowState.CurrentRows)
                                Case "住民区分"
                                    pDiff.ColumnType = enmColumnType.Combo
                                    pDiff.DataSource = New DataView(App農地基本台帳.DataMaster.Body, "Class='住民区分'", "[ID]", DataViewRowState.CurrentRows)
                                Case "登記簿地目"
                                    pDiff.ColumnType = enmColumnType.Combo
                                    pDiff.DataSource = New DataView(App農地基本台帳.DataMaster.Body, "Class='地目'", "[ID]", DataViewRowState.CurrentRows)
                                Case "農委地目"
                                    pDiff.ColumnType = enmColumnType.Combo
                                    pDiff.DataSource = New DataView(App農地基本台帳.DataMaster.Body, "Class='農委地目'", "[ID]", DataViewRowState.CurrentRows)
                                Case "現況地目"
                                    pDiff.ColumnType = enmColumnType.Combo
                                    pDiff.DataSource = New DataView(App農地基本台帳.DataMaster.Body, "Class='課税地目'", "[ID]", DataViewRowState.CurrentRows)
                                Case "農地状況"
                                    pDiff.ColumnType = enmColumnType.Combo
                                    pDiff.DataSource = New DataView(App農地基本台帳.DataMaster.Body, "Class='農地状況'", "[ID]", DataViewRowState.CurrentRows)
                                Case "農振区分"
                                    pDiff.ColumnType = enmColumnType.Combo
                                    pDiff.DataSource = New DataView(App農地基本台帳.DataMaster.Body, "Class='農振区分'", "[ID]", DataViewRowState.CurrentRows)
                                Case "あっせん希望種別"
                                    pDiff.ColumnType = enmColumnType.Combo
                                    pDiff.DataSource = New DataView(App農地基本台帳.DataMaster.Body, "Class='あっせん希望種別'", "[ID]", DataViewRowState.CurrentRows)
                                Case Else
                                    Stop
                            End Select
                        Case "GridViewType"
                            pDiff.ColumnType = CType(pAttr, GridViewType).CellType
                        Case Else
                            Stop
                    End Select
                Next

                If pDiff.IsAddAble Then

                    mvarGrid.Columns.Add(pDiff.CreateColumn())
                End If

            End If

        Next
        Dim pKeyColumn As New DataGridViewTextBoxColumn
        pKeyColumn.DataPropertyName = "Key"
        pKeyColumn.Name = "Key"
        pKeyColumn.Visible = False
        mvarGrid.Columns.Add(pKeyColumn)

        Dim pIconColumn As New DataGridViewTextBoxColumn
        pIconColumn.DataPropertyName = "アイコン"
        pIconColumn.Name = "アイコン"
        pIconColumn.Visible = False
        mvarGrid.Columns.Add(pIconColumn)
    End Sub

    Public Overrides Property IconKey As String
        Get
            Return Nothing
        End Get
        Set(value As String)

        End Set
    End Property

    Public Overrides Sub 検索開始(sWhere As String, sViewWhere As String, Optional sOrderBy As String = "", Optional sColumnStyle As String = "")

    End Sub

    Public Class ColumnDef
        Public Name As String = ""
        Public Browsable As Boolean = False
        Public CategoryAttribute As String = ""
        Public DataName As String = ""
        Public DataType As System.Type
        Public DisplayName As String = ""
        Public Order As Integer = Nothing
        Public ReadOnlyAttr As Boolean = False
        Public ColumnType As enmColumnType = enmColumnType.Text
        Public DataSource As DataView

        Public Sub New(ByVal sName As String, ByVal nType As System.Type, Optional ByVal sDataName As String = "")
            Me.Name = sName
            Me.DisplayName = sName
            Me.DataType = nType

            If sDataName.Length > 0 Then
                Me.DataName = sDataName
            End If
            Select Case Me.DataType.FullName
                Case "System.Int32", "System.Decimal"

                Case "System.Date", "System.DateTime"
                    Me.ColumnType = enmColumnType.DateTime
                Case Else
            End Select
        End Sub

        Public Function IsAddAble() As Boolean
            Return Browsable AndAlso DataName.Length > 0
        End Function

        Public Function CreateColumn() As DataGridViewColumn

            Select Case ColumnType
                Case enmColumnType.Combo
                    Dim pCreateColumn As New DataGridViewComboBoxColumn
                    Select Case Me.DataType.FullName
                        Case Else
                            pCreateColumn.DataPropertyName = Me.Name
                            pCreateColumn.DataSource = Me.DataSource
                            pCreateColumn.ValueMember = "ID"
                            pCreateColumn.DisplayMember = "名称"
                    End Select
                    pCreateColumn.Name = Me.DataName
                    pCreateColumn.HeaderText = Me.DisplayName
                    pCreateColumn.DataPropertyName = Me.DataName

                    pCreateColumn.ToolTipText = Me.CategoryAttribute

                    pCreateColumn.DisplayStyleForCurrentCellOnly = True
                    If ReadOnlyAttr Then
                        pCreateColumn.ReadOnly = Me.ReadOnlyAttr
                        pCreateColumn.DefaultCellStyle.BackColor = Color.LightGoldenrodYellow
                    End If
                    Return pCreateColumn
                Case enmColumnType.EnableBtn
                    Dim pCreateButton As New HimTools2012.controls.DataGridViewDisableButtonColumn
                    With pCreateButton
                        .Name = Me.DataName
                        .HeaderText = Me.DisplayName
                        .DataPropertyName = Me.DataName
                    End With
                    pCreateButton.UseColumnTextForButtonValue = False
                    Return pCreateButton
                Case enmColumnType.CheckBox
                    Dim pCreateCheckBox As New DataGridViewCheckBoxColumn
                    With pCreateCheckBox
                        .Name = Me.DataName
                        .HeaderText = Me.DisplayName
                        .DataPropertyName = Me.DataName
                    End With
                    Return pCreateCheckBox
                Case Else
                    Dim pCreateColumn As DataGridViewColumn
                    Select Case Me.DataType.FullName
                        Case "System.Int32", "System.Decimal"
                            pCreateColumn = New DataGridViewTextBoxColumn
                            pCreateColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        Case "System.Date", "System.DateTime"
                            pCreateColumn = New DataGridViewDateTimePickerColumn '【保留】HimTools2012.controls.
                        Case Else
                            pCreateColumn = New DataGridViewTextBoxColumn
                    End Select

                    pCreateColumn.Name = Me.Name
                    pCreateColumn.HeaderText = Me.DisplayName
                    pCreateColumn.DataPropertyName = Me.DataName
                    pCreateColumn.ReadOnly = Me.ReadOnlyAttr
                    pCreateColumn.ToolTipText = Me.CategoryAttribute
                    If ReadOnlyAttr Then
                        pCreateColumn.DefaultCellStyle.BackColor = Color.LightGoldenrodYellow
                    End If

                    Return pCreateColumn
            End Select
        End Function
    End Class

End Class

