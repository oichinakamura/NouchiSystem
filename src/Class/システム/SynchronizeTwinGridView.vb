Imports System.Windows.Forms

Public Class SynchronizeTwinGridView
    Inherits SplitContainer
    Public WithEvents Grid01 As HimTools2012.controls.DataGridViewWithDataView
    Public WithEvents Grid02 As HimTools2012.controls.DataGridViewWithDataView

    Public Sub New(ByRef G1 As HimTools2012.controls.DataGridViewWithDataView, ByRef G2 As HimTools2012.controls.DataGridViewWithDataView, Optional FixedPanelNumber As Integer = 0)
        MyBase.New()
        Grid01 = G1
        Grid02 = G2
        Panel1.Controls.Add(Grid01)
        Panel2.Controls.Add(Grid02)
        Select Case FixedPanelNumber
            Case 0
            Case 1
                Grid01.ScrollBars = ScrollBars.None
                Me.FixedPanel = FixedPanel.Panel1
            Case 2
                Grid02.ScrollBars = ScrollBars.None
                Me.FixedPanel = FixedPanel.Panel2
        End Select

    End Sub

    Public Property FixedPanelNumber() As Integer
        Get
            If Me.FixedPanel = FixedPanel.Panel1 Then
                Return 1
            ElseIf Me.FixedPanel = FixedPanel.Panel2 Then
                Return 2
            Else
                Return 0
            End If
        End Get
        Set(value As Integer)
            Select Case value
                Case 0
                Case 1
                    Grid01.ScrollBars = ScrollBars.None
                    Grid02.ScrollBars = ScrollBars.Both
                    Me.FixedPanel = FixedPanel.Panel1
                Case 2
                    Grid01.ScrollBars = ScrollBars.Both
                    Grid02.ScrollBars = ScrollBars.None
                    Me.FixedPanel = FixedPanel.Panel2
            End Select
        End Set
    End Property


    Private Sub grid02_ColumnWidthChanged(sender As Object, e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles Grid02.ColumnWidthChanged
        If Me.Orientation = Windows.Forms.Orientation.Horizontal AndAlso Not Grid01.Columns(e.Column.Name).Width = e.Column.Width Then
            Grid01.Columns(e.Column.Name).Width = e.Column.Width
        End If
    End Sub

    Private Sub grid01_ColumnWidthChanged(sender As Object, e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles Grid01.ColumnWidthChanged
        If Me.Orientation = Windows.Forms.Orientation.Horizontal AndAlso Not Grid02.Columns(e.Column.Name).Width = e.Column.Width Then
            Grid02.Columns(e.Column.Name).Width = e.Column.Width
        End If
    End Sub

    Private Sub grid01_Scroll(sender As Object, e As System.Windows.Forms.ScrollEventArgs) Handles Grid01.Scroll
        Select Case e.ScrollOrientation
            Case ScrollOrientation.HorizontalScroll
                Grid02.HorizontalScrollingOffset = e.NewValue
        End Select
    End Sub

    Private Sub grid02_Scroll(sender As Object, e As System.Windows.Forms.ScrollEventArgs) Handles Grid02.Scroll
        Select Case e.ScrollOrientation
            Case ScrollOrientation.HorizontalScroll
                Grid01.HorizontalScrollingOffset = e.NewValue
        End Select
    End Sub
End Class
