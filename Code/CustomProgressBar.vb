Imports System
Imports System.Drawing
Imports System.Drawing.Drawing2D

Public Class CustomProgressBar

    Inherits UserControl

    Public Property Maximum As Integer
    Public Property Value As Integer

    Protected _nAnimation As Integer
    Protected WithEvents _timer As Timer 


    Public Sub New()

        SetStyle(ControlStyles.OptimizedDoubleBuffer Or ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint, True)

        _nAnimation = 0

        Me.Value = 0
        Me.Maximum = 10

        _timer = New Timer()
        _timer.Interval = 30

    End Sub

    Protected Sub timer_Tick(ByVal sender As Object, ByVal e As EventArgs) Handles _timer.Tick

        'PerformStep()

        _nAnimation = _nAnimation + 2
        If (_nAnimation > 1024) Then
            _nAnimation = 0
        End If

        Me.Invalidate()

    End Sub

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)

        If _timer.Enabled = False Then
            _timer.Start()
        End If

        Dim ptTop = New Point(e.ClipRectangle.X, e.ClipRectangle.Y)
        Dim ptBottom = New Point(e.ClipRectangle.X, e.ClipRectangle.Bottom)

        Dim brYellow As LinearGradientBrush = New LinearGradientBrush(ptTop, ptBottom, Color.FromArgb(255, 254, 231, 0), Color.FromArgb(255, 254, 231, 0))
        Dim brBlack As LinearGradientBrush = New LinearGradientBrush(ptTop, ptBottom, Color.FromArgb(255, 48, 36, 0), Color.FromArgb(255, 0, 0, 0))

        'Dim brYellow As SolidBrush = New SolidBrush(Color.FromArgb(255, 254, 231, 0))
        'Dim brBlack As SolidBrush = New SolidBrush(Color.FromArgb(255, 48, 36, 0))

        e.Graphics.FillRectangle(brYellow, e.ClipRectangle)

        Dim pen As Pen = New Pen(Color.Black, 6)

        Dim path As GraphicsPath

        Dim arrPoints(5) As Point
        Dim arrTypes(5) As Byte

        e.Graphics.FillRectangle(brYellow, e.ClipRectangle)

        Dim nPercent As Integer = (Me.Value / Me.Maximum) * 100
        Dim nLimit As Integer = (e.ClipRectangle.Width / 100) * nPercent

        Dim nStart As Integer = (e.ClipRectangle.Left - 1024) + _nAnimation

        Dim rcOriginalClip As RectangleF = e.Graphics.ClipBounds
        Dim rcProgressClip As RectangleF = New RectangleF(rcOriginalClip.X, rcOriginalClip.Y, rcOriginalClip.Width, rcOriginalClip.Height)
        rcProgressClip.Width = rcProgressClip.X + (e.ClipRectangle.Right - e.ClipRectangle.Width) + nLimit + 5

        e.Graphics.SetClip(rcProgressClip)

        For X = nStart To e.ClipRectangle.Right Step 20

            arrTypes(1) = PathPointType.Start
            arrTypes(2) = PathPointType.Line
            arrTypes(3) = PathPointType.Line
            arrTypes(4) = PathPointType.Line
            arrTypes(5) = PathPointType.Line

            arrPoints(1) = New Point(X, e.ClipRectangle.Y)
            arrPoints(2) = New Point(X + 20, e.ClipRectangle.Y)
            arrPoints(3) = New Point(X, e.ClipRectangle.Bottom)
            arrPoints(4) = New Point(X - 20, e.ClipRectangle.Bottom)
            arrPoints(5) = New Point(X, e.ClipRectangle.Y)

            path = New GraphicsPath(arrPoints, arrTypes)

            path.CloseFigure()

            e.Graphics.DrawPath(pen, path)

        Next

        e.Graphics.SetClip(rcOriginalClip)

        pen = New Pen(Color.Black, 3)
        e.Graphics.DrawRectangle(pen, e.ClipRectangle)

    End Sub

    Protected Overrides Sub OnResize(ByVal e As EventArgs)
        Me.Invalidate()
    End Sub

    Public Sub PerformStep()

        Me.Value = Me.Value + 1
        Me.Invalidate()

    End Sub

End Class
