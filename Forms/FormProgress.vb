Option Explicit On

Imports System.IO
Imports System.Drawing
Imports System.Security
Imports System.Reflection
Imports System.Windows.Forms
Imports System.ComponentModel
Imports System.Drawing.Drawing2D
Imports System.Collections.Generic
Imports System.Text.RegularExpressions

Imports System.Runtime.InteropServices

Imports Microsoft.Win32

Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Interop.Excel

Imports Word = Microsoft.Office.Interop.Word
Imports Excel = Microsoft.Office.Interop.Excel

Public Class FormProgress

#Region "Data members"

    Private allowCoolMove As Boolean = False
    Private dx, dy As Integer 'I used this two integers as I could use the function new POint due to the Import of the Excel

    Public _bStop As Boolean

#End Region

#Region "Events"

    Public Sub New()

        InitializeComponent()

        _bStop = False

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load

        Me.Top += 25

    End Sub

    Private Sub Form1_FormClosing(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.FormClosing

    End Sub

    Private Sub btnStop_Click(sender As System.Object, e As System.EventArgs) Handles btnStop.Click

        _bStop = True

        Me.DialogResult = DialogResult.None

    End Sub

    Private Sub Form1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown
        allowCoolMove = True
        dx = Cursor.Position.X - Me.Location.X '// get coordinates.
        dy = Cursor.Position.Y - Me.Location.Y '// get coordinates.
    End Sub

    Private Sub Form1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseMove
        If allowCoolMove = True Then
            Dim cls As New ClassMain
            Me.Location = cls.Pts(Cursor.Position.X - dx, Cursor.Position.Y - dy) '// set coordinates.
        End If
    End Sub

    Private Sub Form1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseUp
        allowCoolMove = False
    End Sub

#End Region

    Private Sub FormProgress_Paint(sender As System.Object, e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint

        Try

            Dim clrTop As Color = SystemColors.ButtonFace
            Dim clrBottom As Color = SystemColors.ControlDark

            Dim rcTop As System.Drawing.Rectangle = e.ClipRectangle
            rcTop.Height = rcTop.Top + (rcTop.Height / 2)

            Dim rcBottom As System.Drawing.Rectangle = e.ClipRectangle
            rcTop.Height = rcTop.Top + (rcTop.Height / 2)
            rcTop.Y = rcTop.Y + rcTop.Height

            '///////////////////////////////////////////////////////////////////////////////////////////////////////

            Dim brTop As New LinearGradientBrush(rcTop, clrTop, clrBottom, LinearGradientMode.Vertical)
            Dim brBottom As New LinearGradientBrush(rcBottom, clrTop, clrBottom, LinearGradientMode.Vertical)

            e.Graphics.FillRectangle(brTop, rcTop)
            e.Graphics.FillRectangle(brBottom, rcBottom)

            '///////////////////////////////////////////////////////////////////////////////////////////////////////

            'Dim rcBorder As System.Drawing.Rectangle = e.ClipRectangle
            'rcBorder.Width -= 1
            'rcBorder.Height -= 1

            'Dim penBorder As Pen = New Pen(Color.Black, 1)

            'e.Graphics.DrawRectangle(penBorder, rcBorder)

            Dim brText As SolidBrush = New SolidBrush(Color.Black)

            '///////////////////////////////////////////////////////////////////////////////////////////////////////

            Dim font As System.Drawing.Font = New System.Drawing.Font(Me.Font.FontFamily, 14, FontStyle.Bold)

            '///////////////////////////////////////////////////////////////////////////////////////////////////////

            Dim ptText As System.Drawing.Point = New System.Drawing.Point(e.ClipRectangle.Left, e.ClipRectangle.Top + 10)

            '///////////////////////////////////////////////////////////////////////////////////////////////////////

            Dim strText As String = "Recovering..."

            '///////////////////////////////////////////////////////////////////////////////////////////////////////

            Dim sizeText As SizeF = e.Graphics.MeasureString(strText, font)
            sizeText.Width += 10

            Dim rcText As System.Drawing.Rectangle = New System.Drawing.Rectangle()

            rcText.X = (e.ClipRectangle.Width / 2) - (sizeText.Width / 2)
            rcText.Y = e.ClipRectangle.Y + 20

            rcText.Width = sizeText.Width
            rcText.Height = sizeText.Height

            '///////////////////////////////////////////////////////////////////////////////////////////////////////

            e.Graphics.DrawString(strText, font, brText, rcText)

        Catch
        End Try

    End Sub

End Class
