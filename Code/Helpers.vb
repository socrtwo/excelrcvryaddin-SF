Imports System
Imports System.Drawing

Imports stdole

Module Helpers

    Public Class ImageHelper

        Inherits AxHost

        Public Sub New()

            MyBase.New(Nothing)

        End Sub

        Public Shared Function GetIPicture(ByVal img As Image) As IPictureDisp

            Dim objPicture As IPictureDisp = Nothing

            Try
                objPicture = AxHost.GetIPictureDispFromPicture(img)
            Catch
            End Try

            GetIPicture = objPicture

        End Function

    End Class


    Public Class ClassMain

        Public Function Pts(ByVal x As Integer, ByVal y As Integer) As Point
            Pts = New Point(x, y)
        End Function

    End Class

End Module
