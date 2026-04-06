Imports System.IO
Imports System.Threading
Imports System.Collections.Generic
Imports System.Runtime.InteropServices

Imports Microsoft.Win32

Imports Extensibility
Imports stdole

Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Drawing

Module ThisModule
    Public ThisAddin As Connect
End Module

<GuidAttribute("55F3C181-4519-4408-A164-B13F0362C422"), ProgIdAttribute("ExcelRecoveryAddin.Connect_2007")> _
Public Class Connect

#Region "Declarations"

    Implements Extensibility.IDTExtensibility2
    Implements Microsoft.Office.Core.IRibbonExtensibility

#End Region

#Region "Data members"

    Private _objExcelApp As Excel.Application

    Private _arrFiles As ArrayList
    Private _objRibbonUI As IRibbonUI
    Private _strCurrentFile As String

    Private _strClearList As String = "<< Clear list >>"

#End Region

#Region "IDTExtensibility2 Implementation"

    Public Sub OnBeginShutdown(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnBeginShutdown
    End Sub

    Public Sub OnAddInsUpdate(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnAddInsUpdate
    End Sub

    Public Sub OnStartupComplete(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnStartupComplete
    End Sub

    Public Sub OnDisconnection(ByVal RemoveMode As Extensibility.ext_DisconnectMode, ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnDisconnection

    End Sub

    Public Sub OnConnection(ByVal objApplication As Object, ByVal connectMode As Extensibility.ext_ConnectMode, ByVal objAddin As Object, ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnConnection

        Try
            _objExcelApp = objApplication
            ThisModule.ThisAddin = Me

            System.Windows.Forms.Application.EnableVisualStyles()

            LoadFilesList()

        Catch ex As System.Exception

            MsgBox(ex.Message)

        End Try

    End Sub

#End Region

#Region "IRibbonExtensibility Implementation"

    Public Function GetCustomUI(ByVal strRibbonID As String) As String Implements IRibbonExtensibility.GetCustomUI

        Dim strRibbonXML As String = ""

        If strRibbonID = "Microsoft.Excel.Workbook" Then

            strRibbonXML = My.Resources.Excel_Main

        End If

        GetCustomUI = strRibbonXML

    End Function

    Public Sub Ribbon_OnLoad(ByVal piRibbonUI As IRibbonUI)

        _objRibbonUI = piRibbonUI

    End Sub

#End Region

#Region "Ribbon Button Events"

    Public Sub On_Ribbon_Button_SaveAsSYLK_Clicked(ByVal piRibbonCtrl As IRibbonControl)

        Dim worker As New Worker()
        Dim strFile = GetCurrentFile()
        If String.IsNullOrEmpty(strFile) = False Then

            Dim thread As New Thread(New ParameterizedThreadStart(AddressOf worker.Impl_SaveAsSYLK))
            thread.Start(strFile)

        End If

    End Sub

    Public Sub On_Ribbon_Button_Manual_Calculations_Clicked(ByVal piRibbonCtrl As IRibbonControl)

        Dim worker As New Worker()
        Dim strFile = GetCurrentFile()
        If String.IsNullOrEmpty(strFile) = False Then

            Dim thread As New Thread(New ParameterizedThreadStart(AddressOf worker.Impl_ManualCalculations))
            thread.Start(strFile)

        End If

    End Sub

    Public Sub On_Ribbon_Button_External_References_Clicked(ByVal piRibbonCtrl As IRibbonControl)

        Dim worker As New Worker()
        Dim strFile = GetCurrentFile()
        If String.IsNullOrEmpty(strFile) = False Then

            Dim thread As New Thread(New ParameterizedThreadStart(AddressOf worker.Impl_ExternalReferences))
            thread.Start(strFile)

        End If

    End Sub

    Public Sub On_Ribbon_Button_Safe_Mode_Clicked(ByVal piRibbonCtrl As IRibbonControl)

        Dim worker As New Worker()
        Dim strFile = GetCurrentFile()
        If String.IsNullOrEmpty(strFile) = False Then

            Dim thread As New Thread(New ParameterizedThreadStart(AddressOf worker.Impl_SafeMode))
            thread.Start(strFile)

        End If

    End Sub

    Public Sub On_Ribbon_Button_SaveAsHTML_Clicked(ByVal piRibbonCtrl As IRibbonControl)

        Dim worker As New Worker()
        Dim strFile = GetCurrentFile()
        If String.IsNullOrEmpty(strFile) = False Then

            Dim thread As New Thread(New ParameterizedThreadStart(AddressOf worker.Impl_SaveAsHTML))
            thread.Start(strFile)

        End If

    End Sub

    Public Sub On_Ribbon_Button_Open_With_WordPad_Clicked(ByVal piRibbonCtrl As IRibbonControl)

        Dim worker As New Worker()
        Dim strFile = GetCurrentFile()
        If String.IsNullOrEmpty(strFile) = False Then

            Dim thread As New Thread(New ParameterizedThreadStart(AddressOf worker.Impl_OpenWithWordPad))
            thread.Start(strFile)

        End If

    End Sub

    Public Sub On_Ribbon_Button_Open_With_Excel_Viewer_Clicked(ByVal piRibbonCtrl As IRibbonControl)

        Dim worker As New Worker()
        Dim strFile = GetCurrentFile()
        If String.IsNullOrEmpty(strFile) = False Then

            Dim thread As New Thread(New ParameterizedThreadStart(AddressOf worker.Impl_OpenWithExcelViewer))
            thread.Start(strFile)

        End If

    End Sub

    Public Sub On_Ribbon_Button_MS_Open_And_Repair_Clicked(ByVal piRibbonCtrl As IRibbonControl)

        Dim worker As New Worker()
        Dim strFile = GetCurrentFile()
        If String.IsNullOrEmpty(strFile) = False Then

            Dim thread As New Thread(New ParameterizedThreadStart(AddressOf worker.Impl_MSOpenAndRepair))
            thread.Start(strFile)

        End If

    End Sub

    Public Sub On_Ribbon_Button_MS_Open_And_Extract_Data_Clicked(ByVal piRibbonCtrl As IRibbonControl)

        Dim worker As New Worker()
        Dim strFile = GetCurrentFile()
        If String.IsNullOrEmpty(strFile) = False Then

            Dim thread As New Thread(New ParameterizedThreadStart(AddressOf worker.Impl_MSOpenAndExtractData))
            thread.Start(strFile)

        End If

    End Sub

    Public Sub On_Ribbon_Button_Non_MS_XLSX_Extract_Data_Clicked(ByVal piRibbonCtrl As IRibbonControl)

        If Worker.IsRunningAsAdmin() = False Then

            If Worker.RequestPriviledges() = True Then

                _objExcelApp.Quit()

                Dim proc As New Process()
                proc.StartInfo.UseShellExecute = True
                proc.StartInfo.Verb = "runas"
                proc.StartInfo.FileName = Process.GetCurrentProcess().MainModule.FileName
                proc.Start()

            End If

            Exit Sub

        End If

        Dim worker_ As New Worker()
        Dim strFile = GetCurrentFile()
        If String.IsNullOrEmpty(strFile) = False Then

            Dim thread As New Thread(New ParameterizedThreadStart(AddressOf worker_.Impl_NonMSXlsxExtractData))
            thread.Start(strFile)

        End If

    End Sub

    Public Sub On_Ribbon_Button_Non_MS_XLSX_Extract_Data_2_Clicked(ByVal piRibbonCtrl As IRibbonControl)

        Dim worker As New Worker()
        Dim strFile = GetCurrentFile()
        If String.IsNullOrEmpty(strFile) = False Then

            Dim thread As New Thread(New ParameterizedThreadStart(AddressOf worker.Impl_NonMSXlsxExtractData2))
            thread.Start(strFile)

        End If

    End Sub

    Public Sub On_Ribbon_Button_Zip_Repair_Try_Clicked(ByVal piRibbonCtrl As IRibbonControl)

        Dim worker As New Worker()
        Dim strFile = GetCurrentFile()
        If String.IsNullOrEmpty(strFile) = False Then

            Dim thread As New Thread(New ParameterizedThreadStart(AddressOf worker.Impl_ZipRepairTry))
            thread.Start(strFile)

        End If

    End Sub

    Public Sub On_Ribbon_Button_Excel_Fix_Clicked(ByVal piRibbonCtrl As IRibbonControl)

        Dim worker As New Worker()
        Dim strFile = GetCurrentFile()
        If String.IsNullOrEmpty(strFile) = False Then

            Dim thread As New Thread(New ParameterizedThreadStart(AddressOf worker.Impl_ExcelFix))
            thread.Start(strFile)

        End If

    End Sub

    Public Sub On_Ribbon_Button_SelectFile_Clicked(ByVal piRibbonCtrl As IRibbonControl)

        SelectFile()

    End Sub

    Public Sub On_Ribbon_Button_PreviousVersion_Clicked(ByVal piRibbonCtrl As IRibbonControl)

        FindPreviousVersions()

    End Sub

#End Region

#Region "Ribbon Images Events"

    Public Function On_Ribbon_Button_SaveAsSYLK_GetImage(ByVal piRibbonCtrl As IRibbonControl) As IPictureDisp

        On_Ribbon_Button_SaveAsSYLK_GetImage = ImageHelper.GetIPicture(My.Resources.SaveAsSYLK)

    End Function

    Public Function On_Ribbon_Button_Manual_Calculations_GetImage(ByVal piRibbonCtrl As IRibbonControl) As IPictureDisp

        On_Ribbon_Button_Manual_Calculations_GetImage = ImageHelper.GetIPicture(My.Resources.ManualCalculations)

    End Function

    Public Function On_Ribbon_Button_External_References_GetImage(ByVal piRibbonCtrl As IRibbonControl) As IPictureDisp

        On_Ribbon_Button_External_References_GetImage = ImageHelper.GetIPicture(My.Resources.ExternalReferences)

    End Function

    Public Function On_Ribbon_Button_Safe_Mode_GetImage(ByVal piRibbonCtrl As IRibbonControl) As IPictureDisp

        On_Ribbon_Button_Safe_Mode_GetImage = ImageHelper.GetIPicture(My.Resources.OpenInSafeMode)

    End Function

    Public Function On_Ribbon_Button_SaveAsHTML_GetImage(ByVal piRibbonCtrl As IRibbonControl) As IPictureDisp

        On_Ribbon_Button_SaveAsHTML_GetImage = ImageHelper.GetIPicture(My.Resources.SaveAsHTML)

    End Function

    Public Function On_Ribbon_Button_Open_With_WordPad_GetImage(ByVal piRibbonCtrl As IRibbonControl) As IPictureDisp

        On_Ribbon_Button_Open_With_WordPad_GetImage = ImageHelper.GetIPicture(My.Resources.OpenInWordPad)

    End Function

    Public Function On_Ribbon_Button_Open_With_Excel_Viewer_GetImage(ByVal piRibbonCtrl As IRibbonControl) As IPictureDisp

        On_Ribbon_Button_Open_With_Excel_Viewer_GetImage = ImageHelper.GetIPicture(My.Resources.OpenInExcel)

    End Function

    Public Function On_Ribbon_Button_MS_Open_And_Repair_GetImage(ByVal piRibbonCtrl As IRibbonControl) As IPictureDisp

        On_Ribbon_Button_MS_Open_And_Repair_GetImage = ImageHelper.GetIPicture(My.Resources.MSOpenAndRepair)

    End Function

    Public Function On_Ribbon_Button_MS_Open_And_Extract_Data_GetImage(ByVal piRibbonCtrl As IRibbonControl) As IPictureDisp

        On_Ribbon_Button_MS_Open_And_Extract_Data_GetImage = ImageHelper.GetIPicture(My.Resources.MSOpenAndExtract)

    End Function

    Public Function On_Ribbon_Button_Non_MS_XLSX_Extract_Data_GetImage(ByVal piRibbonCtrl As IRibbonControl) As IPictureDisp

        On_Ribbon_Button_Non_MS_XLSX_Extract_Data_GetImage = ImageHelper.GetIPicture(My.Resources.NonMSDataExtract)

    End Function

    Public Function On_Ribbon_Button_Non_MS_XLSX_Extract_Data_2_GetImage(ByVal piRibbonCtrl As IRibbonControl) As IPictureDisp

        On_Ribbon_Button_Non_MS_XLSX_Extract_Data_2_GetImage = ImageHelper.GetIPicture(My.Resources.NonMSDataExtract2)

    End Function

    Public Function On_Ribbon_Button_Zip_Repair_Try_GetImage(ByVal piRibbonCtrl As IRibbonControl) As IPictureDisp

        On_Ribbon_Button_Zip_Repair_Try_GetImage = ImageHelper.GetIPicture(My.Resources.ZipRepairTry)

    End Function

    Public Function On_Ribbon_Button_Excel_Fix_GetImage(ByVal piRibbonCtrl As IRibbonControl) As IPictureDisp

        On_Ribbon_Button_Excel_Fix_GetImage = ImageHelper.GetIPicture(My.Resources.ExcelFix)

    End Function

    Public Function On_Ribbon_Button_SelectFile_GetImage(ByVal piRibbonCtrl As IRibbonControl) As IPictureDisp

        On_Ribbon_Button_SelectFile_GetImage = ImageHelper.GetIPicture(My.Resources.SelectFile)

    End Function

    Public Function On_Ribbon_Button_PreviousVersion_GetImage(ByVal piRibbonCtrl As IRibbonControl) As IPictureDisp

        On_Ribbon_Button_PreviousVersion_GetImage = ImageHelper.GetIPicture(My.Resources.FindVersions)

    End Function

#End Region

#Region "Combobox Events"

    Public Function On_Ribbon_Combobox_CurrentFile_GetItemCount(control As IRibbonControl) As Integer

        Return _arrFiles.Count

    End Function

    Public Function On_Ribbon_Combobox_CurrentFile_GetItemLabel(control As IRibbonControl, index As Integer) As String

        Dim strResult As String = ""

        Try
            strResult = CType(_arrFiles(index), String)

            strResult = Path.GetFileName(strResult)
        Catch
        End Try

        Return strResult

    End Function

    Public Function On_Ribbon_Combobox_CurrentFile_GetText(control As IRibbonControl) As String

        Dim strResult As String = ""

        Try
            strResult = Path.GetFileName(_strCurrentFile)
        Catch
        End Try

        Return strResult

    End Function

    Public Sub On_Ribbon_Combobox_CurrentFile_OnChange(control As IRibbonControl, text As String)

        If text = _strClearList Then

            _arrFiles.Clear()
            _arrFiles.Add(_strClearList)

            _strCurrentFile = ""

        Else

            Dim bExists As Boolean = False

            For Each strFullPath In _arrFiles

                If strFullPath.EndsWith(text) = True Then

                    _strCurrentFile = strFullPath
                    bExists = True

                    Exit For

                End If

            Next

            If bExists = False Then

                _strCurrentFile = text

            End If

        End If

        SaveFilesList()

        _objRibbonUI.Invalidate()

    End Sub

#End Region

#Region "General Implementation"

    Private Sub LoadFilesList()

        Try
            _arrFiles = New ArrayList()

            Dim key As RegistryKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\ExcelRecovery\RecentFiles")
            If key IsNot Nothing Then

                Dim arrValues As String() = key.GetValueNames()

                Dim strValue As String

                For Each strValueName As String In arrValues
                    strValue = key.GetValue(strValueName)
                    If String.IsNullOrEmpty(strValue) = False Then
                        _arrFiles.Add(strValue)
                    End If
                Next

                key.Close()

            End If

            _arrFiles.Add(_strClearList)

            If _arrFiles.Count > 1 Then
                _strCurrentFile = CType(_arrFiles(0), String)
            End If

        Catch
        End Try

    End Sub

    Private Sub SaveFilesList()

        Try

            Try
                Registry.CurrentUser.DeleteSubKeyTree("SOFTWARE\ExcelRecovery\RecentFiles")
            Catch
            End Try

            Dim key As RegistryKey = Registry.CurrentUser.CreateSubKey("SOFTWARE\ExcelRecovery\RecentFiles")
            If key IsNot Nothing Then

                Dim nCounter As Integer = 1

                For Each strValue As String In _arrFiles

                    If strValue <> _strClearList Then

                        key.SetValue("File" + nCounter.ToString(), strValue)

                        nCounter = nCounter + 1

                    End If

                Next

                key.Close()

            End If

        Catch
        End Try

    End Sub

    Private Function GetCurrentFile() As String

        If String.IsNullOrEmpty(_strCurrentFile) = True Then
            _strCurrentFile = SelectFile()
        End If

        Return _strCurrentFile
    End Function

    Private Function SelectFile() As String

        Dim strResult As String = ""

        Dim dlgOpenFile As New OpenFileDialog()

        dlgOpenFile.Filter = "All files (*.*)|*.*||"

        If dlgOpenFile.ShowDialog() = DialogResult.OK Then

            _strCurrentFile = dlgOpenFile.FileName

            Dim bExists As Boolean = False

            For Each strFullPath As String In _arrFiles

                If String.Compare(strFullPath, _strCurrentFile, True) = 0 Then
                    bExists = True
                    Exit For
                End If

            Next

            If bExists = False Then

                _arrFiles.Insert(0, dlgOpenFile.FileName)

            End If

            SaveFilesList()

            _objRibbonUI.Invalidate()

            strResult = dlgOpenFile.FileName

        End If

        Return strResult

    End Function

    Private Sub FindPreviousVersions()

        Dim formMain As New FormMain()
        formMain.objExcel = _objExcelApp
        formMain.PathTb.Text = _strCurrentFile
        formMain.ShowDialog()

    End Sub

#End Region

End Class
