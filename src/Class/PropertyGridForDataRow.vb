
''' <summary></summary>
Public Class PropertyGridForDataRow
    Inherits DataGridView

    Private mvarDataSource As Object = Nothing
    Private mvarAutoGenerateColumns As Boolean = False

    ''' <summary>PropertyGridForDataRowのコンストラクタ</summary>
    ''' <remarks>Verified [中村 雄一 date：2016/9/29 18:57]</remarks>
    Public Sub New()
        Me.AllowUserToAddRows = False
        Me.AllowUserToDeleteRows = False
        Me.AllowUserToOrderColumns = False

        Me.Columns.Clear()
        Dim pTitle As New DataGridViewTextBoxColumn
        pTitle.HeaderText = "項目名"
        Me.Columns.Add(pTitle)
        Dim pData As New DataGridViewTextBoxColumn
        pData.HeaderText = "データ"
        Me.Columns.Add(pData)

    End Sub

    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Verified [中村 雄一 date：2016/9/29 18:58]</remarks>
    Public Shadows Property DataSource() As Object
        Get
            Return mvarDataSource
        End Get
        Set(value As Object)
            mvarDataSource = value
            If TypeOf mvarDataSource Is DataTable Then
                Me.Rows.Clear()
                With CType(mvarDataSource, DataTable)
                    For Each pCol As DataColumn In CType(mvarDataSource, DataTable).Columns
                        Select Case pCol.DataType.FullName
                            Case "System.Int32", "System.Int16", "System.Double", "System.Single", "System.Decimal"
                                Me.Rows.Add({pCol.ColumnName, .Rows(0).Item(pCol.ColumnName)})

                            Case "System.String"
                                Me.Rows.Add({pCol.ColumnName, .Rows(0).Item(pCol.ColumnName)})
                            Case "System.DateTime"
                                Dim pCell As New HimTools2012.controls.CalendarCell
                                pCell.Value = .Rows(0).Item(pCol.ColumnName)
                                Dim n As Integer = Me.Rows.Add({pCol.ColumnName, .Rows(0).Item(pCol.ColumnName)})
                                Me.Rows(n).Cells(1) = pCell
                            Case "System.Boolean"
                                Dim pCell As New DataGridViewCheckBoxCell()

                                pCell.Value = .Rows(0).Item(pCol.ColumnName)
                                Dim n As Integer = Me.Rows.Add({pCol.ColumnName, .Rows(0).Item(pCol.ColumnName)})
                                Me.Rows(n).Cells(1) = pCell
                            Case Else
                                Stop
                                Me.Rows.Add({pCol.ColumnName, .Rows(0).Item(pCol.ColumnName)})
                        End Select
                    Next
                End With
            End If

        End Set
    End Property

    ''' <summary></summary>
    ''' <value>[Public]</value>
    ''' <remarks>Verified [中村 雄一 date：2016/9/29 18:58]</remarks>
    Public Shadows Property AutoGenerateColumns As Boolean
        Get
            Return mvarAutoGenerateColumns
        End Get
        Set(value As Boolean)
            mvarAutoGenerateColumns = value
        End Set
    End Property

End Class
