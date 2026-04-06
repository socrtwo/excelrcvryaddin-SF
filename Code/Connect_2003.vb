Imports System.IO
Imports System.Drawing
Imports System.Threading
Imports System.Collections.Generic
Imports System.Runtime.InteropServices

Imports Microsoft.Win32

Imports stdole
Imports Extensibility

Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Reflection

Module ThisModule
    Public ThisAddin As Connect
End Module

<GuidAttribute("9A3A30C0-D000-4FF6-B490-A227E0D363F6"), ProgIdAttribute("ExcelRecoveryAddin.Connect_2003")> _
Public Class Connect

#Region "Declarations"

    Implements Extensibility.IDTExtensibility2

#End Region

#Region "Data members"

    Private _objExcelApp As Excel.Application

    Private _arrFiles As ArrayList
    Private _strCurrentFile As String

    Private _cmb As CommandBarComboBox

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

            Dim objExcelToolbar As CommandBar = FindCommandBar(_objExcelApp.CommandBars, "ExcelRecoveryAddin")
            If objExcelToolbar IsNot Nothing Then
                objExcelToolbar.Delete()
            End If

            Dim objExcelMenuBar As CommandBar = _objExcelApp.CommandBars(1)

            Dim objMenuItem As CommandBarPopup

            Dim objControls As CommandBarControls = objExcelMenuBar.Controls

            objMenuItem = objControls.Add(MsoControlType.msoControlPopup, Missing.Value, Missing.Value, objControls.Count + 1, True)
            objMenuItem.Caption = "Recovery Add-in"
            objMenuItem.Tag = "ExcelRecoveryAddin"

            objControls = objMenuItem.Controls

            If objControls IsNot Nothing Then

                Dim btn As CommandBarButton

                '----------------------------------------------------------------------------------------------------------------

                btn = objControls.Add(MsoControlType.msoControlButton, Missing.Value, Missing.Value, objControls.Count + 1, True)
                If btn IsNot Nothing Then

                    btn.Caption = "Select File"
                    btn.Style = MsoButtonStyle.msoButtonIconAndCaption
                    btn.Picture = On_CommandBar_Button_SelectFile_GetImage()
                    'btn.Mask = On_CommandBar_Button_SelectFile_GetImage()
                    btn.Tag = btn.Caption
                    btn.BeginGroup = False

                    AddHandler btn.Click, AddressOf On_CommandBar_Button_SelectFile_Clicked

                End If

                '----------------------------------------------------------------------------------------------------------------

                btn = objControls.Add(MsoControlType.msoControlButton, Missing.Value, Missing.Value, objControls.Count + 1, True)
                If btn IsNot Nothing Then

                    btn.Caption = "Find Previous Version"
                    btn.Style = MsoButtonStyle.msoButtonIconAndCaption
                    btn.Picture = On_CommandBar_Button_PreviousVersion_GetImage()
                    'btn.Mask = On_CommandBar_Button_PreviousVersion_GetImage()
                    btn.Tag = btn.Caption
                    btn.BeginGroup = False

                    AddHandler btn.Click, AddressOf On_CommandBar_Button_PreviousVersion_Clicked

                End If

                '----------------------------------------------------------------------------------------------------------------

                _cmb = objControls.Add(MsoControlType.msoControlComboBox, Missing.Value, Missing.Value, objControls.Count + 1, True)
                If _cmb IsNot Nothing Then

                    _cmb.Caption = "Current file:"
                    _cmb.BeginGroup = True
                    _cmb.Width = 170

                    AddHandler _cmb.Change, AddressOf On_CommandBar_Combobox_CurrentFile_Change

                End If

                '----------------------------------------------------------------------------------------------------------------

                btn = objControls.Add(MsoControlType.msoControlButton, Missing.Value, Missing.Value, objControls.Count + 1, True)
                If btn IsNot Nothing Then

                    btn.Caption = "MS Open and Repair"
                    btn.Style = MsoButtonStyle.msoButtonIconAndCaption
                    btn.Picture = On_CommandBar_Button_MS_Open_And_Repair_GetImage()
                    btn.Mask = On_CommandBar_Button_MS_Open_And_Repair_GetImage()
                    btn.Tag = btn.Caption
                    btn.BeginGroup = True

                    AddHandler btn.Click, AddressOf On_CommandBar_Button_MS_Open_And_Repair_Clicked

                End If

                '----------------------------------------------------------------------------------------------------------------

                btn = objControls.Add(MsoControlType.msoControlButton, Missing.Value, Missing.Value, objControls.Count + 1, True)
                If btn IsNot Nothing Then

                    btn.Caption = "MS Open and Extract Data"
                    btn.Style = MsoButtonStyle.msoButtonIconAndCaption
                    btn.Picture = On_CommandBar_Button_MS_Open_And_Extract_Data_GetImage()
                    btn.Mask = On_CommandBar_Button_MS_Open_And_Extract_Data_GetImage()
                    btn.Tag = btn.Caption
                    btn.BeginGroup = False

                    AddHandler btn.Click, AddressOf On_CommandBar_Button_MS_Open_And_Extract_Data_Clicked

                End If

                '----------------------------------------------------------------------------------------------------------------

                btn = objControls.Add(MsoControlType.msoControlButton, Missing.Value, Missing.Value, objControls.Count + 1, True)
                If btn IsNot Nothing Then

                    btn.Caption = "Non-MS XLSX Data Extract I"
                    btn.Style = MsoButtonStyle.msoButtonIconAndCaption
                    btn.Picture = On_CommandBar_Button_Non_MS_XLSX_Extract_Data_GetImage()
                    btn.Mask = On_CommandBar_Button_Non_MS_XLSX_Extract_Data_GetImage()
                    btn.Tag = btn.Caption
                    btn.BeginGroup = False

                    AddHandler btn.Click, AddressOf On_CommandBar_Button_Non_MS_XLSX_Extract_Data_Clicked

                End If

                '----------------------------------------------------------------------------------------------------------------

                btn = objControls.Add(MsoControlType.msoControlButton, Missing.Value, Missing.Value, objControls.Count + 1, True)
                If btn IsNot Nothing Then

                    btn.Caption = "Non-MS XLSX Data Extract II"
                    btn.Style = MsoButtonStyle.msoButtonIconAndCaption
                    btn.Picture = On_CommandBar_Button_Non_MS_XLSX_Extract_Data_2_GetImage()
                    btn.Mask = On_CommandBar_Button_Non_MS_XLSX_Extract_Data_2_GetImage()
                    btn.Tag = btn.Caption
                    btn.BeginGroup = False

                    AddHandler btn.Click, AddressOf On_CommandBar_Button_Non_MS_XLSX_Extract_Data_2_Clicked

                End If

                '----------------------------------------------------------------------------------------------------------------

                btn = objControls.Add(MsoControlType.msoControlButton, Missing.Value, Missing.Value, objControls.Count + 1, True)
                If btn IsNot Nothing Then

                    btn.Caption = "ZipRepair-Try for XLS and XLSX"
                    btn.Style = MsoButtonStyle.msoButtonIconAndCaption
                    btn.Picture = On_CommandBar_Button_Zip_Repair_Try_GetImage()
                    btn.Mask = On_CommandBar_Button_Zip_Repair_Try_GetImage()
                    btn.Tag = btn.Caption
                    btn.BeginGroup = False

                    AddHandler btn.Click, AddressOf On_CommandBar_Button_Zip_Repair_Try_Clicked

                End If

                '----------------------------------------------------------------------------------------------------------------

                btn = objControls.Add(MsoControlType.msoControlButton, Missing.Value, Missing.Value, objControls.Count + 1, True)
                If btn IsNot Nothing Then

                    btn.Caption = "Cimaware ExcelFix"
                    btn.Style = MsoButtonStyle.msoButtonIconAndCaption
                    btn.Picture = On_CommandBar_Button_Excel_Fix_GetImage()
                    'btn.Mask = On_CommandBar_Button_Excel_Fix_GetImage()
                    btn.Tag = btn.Caption
                    btn.BeginGroup = False

                    AddHandler btn.Click, AddressOf On_CommandBar_Button_Excel_Fix_Clicked

                End If

                '----------------------------------------------------------------------------------------------------------------

                btn = objControls.Add(MsoControlType.msoControlButton, Missing.Value, Missing.Value, objControls.Count + 1, True)
                If btn IsNot Nothing Then

                    btn.Caption = "Save as SYLK"
                    btn.Style = MsoButtonStyle.msoButtonIconAndCaption
                    btn.Picture = On_CommandBar_Button_SaveAsSYLK_GetImage()
                    btn.Mask = On_CommandBar_Button_SaveAsSYLK_GetImage()
                    btn.Tag = btn.Caption
                    btn.BeginGroup = True

                    AddHandler btn.Click, AddressOf On_CommandBar_Button_SaveAsSYLK_Clicked

                End If

                '----------------------------------------------------------------------------------------------------------------

                btn = objControls.Add(MsoControlType.msoControlButton, Missing.Value, Missing.Value, objControls.Count + 1, True)
                If btn IsNot Nothing Then

                    btn.Caption = "Save as HTML"
                    btn.Style = MsoButtonStyle.msoButtonIconAndCaption
                    btn.Picture = On_CommandBar_Button_Save_As_HTML_GetImage()
                    btn.Mask = On_CommandBar_Button_Save_As_HTML_GetImage()
                    btn.Tag = btn.Caption
                    btn.BeginGroup = False

                    AddHandler btn.Click, AddressOf On_CommandBar_Button_Save_As_HTML_Clicked

                End If

                '----------------------------------------------------------------------------------------------------------------

                btn = objControls.Add(MsoControlType.msoControlButton, Missing.Value, Missing.Value, objControls.Count + 1, True)
                If btn IsNot Nothing Then

                    btn.Caption = "Manual Calculations"
                    btn.Style = MsoButtonStyle.msoButtonIconAndCaption
                    btn.Picture = On_CommandBar_Button_Manual_Calculations_GetImage()
                    btn.Mask = On_CommandBar_Button_Manual_Calculations_GetImage()
                    btn.Tag = btn.Caption
                    btn.BeginGroup = False

                    AddHandler btn.Click, AddressOf On_CommandBar_Button_Manual_Calculations_Clicked

                End If

                '----------------------------------------------------------------------------------------------------------------

                btn = objControls.Add(MsoControlType.msoControlButton, Missing.Value, Missing.Value, objControls.Count + 1, True)
                If btn IsNot Nothing Then

                    btn.Caption = "External References"
                    btn.Style = MsoButtonStyle.msoButtonIconAndCaption
                    btn.Picture = On_CommandBar_Button_External_References_GetImage()
                    btn.Mask = On_CommandBar_Button_External_References_GetImage()
                    btn.Tag = btn.Caption
                    btn.BeginGroup = False

                    AddHandler btn.Click, AddressOf On_CommandBar_Button_External_References_Clicked

                End If

                '----------------------------------------------------------------------------------------------------------------

                btn = objControls.Add(MsoControlType.msoControlButton, Missing.Value, Missing.Value, objControls.Count + 1, True)
                If btn IsNot Nothing Then

                    btn.Caption = "Open in Safe Mode"
                    btn.Style = MsoButtonStyle.msoButtonIconAndCaption
                    btn.Picture = On_CommandBar_Button_Safe_Mode_GetImage()
                    btn.Mask = On_CommandBar_Button_Safe_Mode_GetImage()
                    btn.Tag = btn.Caption
                    btn.BeginGroup = False

                    AddHandler btn.Click, AddressOf On_CommandBar_Button_Safe_Mode_Clicked

                End If

                '----------------------------------------------------------------------------------------------------------------


                btn = objControls.Add(MsoControlType.msoControlButton, Missing.Value, Missing.Value, objControls.Count + 1, True)
                If btn IsNot Nothing Then

                    btn.Caption = "Open with WordPad"
                    btn.Style = MsoButtonStyle.msoButtonIconAndCaption
                    btn.Picture = On_CommandBar_Button_Open_With_WordPad_GetImage()
                    btn.Mask = On_CommandBar_Button_Open_With_WordPad_GetImage()
                    btn.Tag = btn.Caption
                    btn.BeginGroup = True

                    AddHandler btn.Click, AddressOf On_CommandBar_Button_Open_With_WordPad_Clicked

                End If

                '----------------------------------------------------------------------------------------------------------------

                btn = objControls.Add(MsoControlType.msoControlButton, Missing.Value, Missing.Value, objControls.Count + 1, True)
                If btn IsNot Nothing Then

                    btn.Caption = "Open with Excel Viewer"
                    btn.Style = MsoButtonStyle.msoButtonIconAndCaption
                    btn.Picture = On_CommandBar_Button_Open_With_Excel_Viewer_GetImage()
                    btn.Mask = On_CommandBar_Button_Open_With_Excel_Viewer_GetImage()
                    btn.Tag = btn.Caption
                    btn.BeginGroup = False

                    AddHandler btn.Click, AddressOf On_CommandBar_Button_Open_With_Excel_Viewer_Clicked

                End If

                '----------------------------------------------------------------------------------------------------------------

            End If

            LoadFilesList()

        Catch ex As System.Exception

            MsgBox(ex.Message)

        End Try

    End Sub

#End Region

#Region "CommandBar Button Events"

    Public Sub On_CommandBar_Button_SaveAsSYLK_Clicked(objButton As CommandBarButton, ByRef bCancelDefault As Boolean)

        Dim worker As New Worker()
        Dim strFile = GetCurrentFile()
        If String.IsNullOrEmpty(strFile) = False Then

            Dim thread As New Thread(New ParameterizedThreadStart(AddressOf worker.Impl_SaveAsSYLK))
            thread.Start(strFile)

        End If

    End Sub

    Public Sub On_CommandBar_Button_Manual_Calculations_Clicked(objButton As CommandBarButton, ByRef bCancelDefault As Boolean)

        Dim worker As New Worker()
        Dim strFile = GetCurrentFile()
        If String.IsNullOrEmpty(strFile) = False Then

            Dim thread As New Thread(New ParameterizedThreadStart(AddressOf worker.Impl_ManualCalculations))
            thread.Start(strFile)

        End If

    End Sub

    Public Sub On_CommandBar_Button_External_References_Clicked(objButton As CommandBarButton, ByRef bCancelDefault As Boolean)

        Dim worker As New Worker()
        Dim strFile = GetCurrentFile()
        If String.IsNullOrEmpty(strFile) = False Then

            Dim thread As New Thread(New ParameterizedThreadStart(AddressOf worker.Impl_ExternalReferences))
            thread.Start(strFile)

        End If

    End Sub

    Public Sub On_CommandBar_Button_Safe_Mode_Clicked(objButton As CommandBarButton, ByRef bCancelDefault As Boolean)

        Dim worker As New Worker()
        Dim strFile = GetCurrentFile()
        If String.IsNullOrEmpty(strFile) = False Then

            Dim thread As New Thread(New ParameterizedThreadStart(AddressOf worker.Impl_SafeMode))
            thread.Start(strFile)

        End If

    End Sub

    Public Sub On_CommandBar_Button_Save_As_HTML_Clicked(objButton As CommandBarButton, ByRef bCancelDefault As Boolean)

        Dim worker As New Worker()
        Dim strFile = GetCurrentFile()
        If String.IsNullOrEmpty(strFile) = False Then

            Dim thread As New Thread(New ParameterizedThreadStart(AddressOf worker.Impl_SaveAsHTML))
            thread.Start(strFile)

        End If

    End Sub

    Public Sub On_CommandBar_Button_Open_With_WordPad_Clicked(objButton As CommandBarButton, ByRef bCancelDefault As Boolean)

        Dim worker As New Worker()
        Dim strFile = GetCurrentFile()
        If String.IsNullOrEmpty(strFile) = False Then

            Dim thread As New Thread(New ParameterizedThreadStart(AddressOf worker.Impl_OpenWithWordPad))
            thread.Start(strFile)

        End If

    End Sub

    Public Sub On_CommandBar_Button_Open_With_Excel_Viewer_Clicked(objButton As CommandBarButton, ByRef bCancelDefault As Boolean)

        Dim worker As New Worker()
        Dim strFile = GetCurrentFile()
        If String.IsNullOrEmpty(strFile) = False Then

            Dim thread As New Thread(New ParameterizedThreadStart(AddressOf worker.Impl_OpenWithExcelViewer))
            thread.Start(strFile)

        End If

    End Sub

    Public Sub On_CommandBar_Button_MS_Open_And_Repair_Clicked(objButton As CommandBarButton, ByRef bCancelDefault As Boolean)

        Dim worker As New Worker()
        Dim strFile = GetCurrentFile()
        If String.IsNullOrEmpty(strFile) = False Then

            Dim thread As New Thread(New ParameterizedThreadStart(AddressOf worker.Impl_MSOpenAndRepair))
            thread.Start(strFile)

        End If

    End Sub

    Public Sub On_CommandBar_Button_MS_Open_And_Extract_Data_Clicked(objButton As CommandBarButton, ByRef bCancelDefault As Boolean)

        Dim worker As New Worker()
        Dim strFile = GetCurrentFile()
        If String.IsNullOrEmpty(strFile) = False Then

            Dim thread As New Thread(New ParameterizedThreadStart(AddressOf worker.Impl_MSOpenAndExtractData))
            thread.Start(strFile)

        End If

    End Sub

    Public Sub On_CommandBar_Button_Non_MS_XLSX_Extract_Data_Clicked(objButton As CommandBarButton, ByRef bCancelDefault As Boolean)

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

    Public Sub On_CommandBar_Button_Non_MS_XLSX_Extract_Data_2_Clicked(objButton As CommandBarButton, ByRef bCancelDefault As Boolean)

        Dim worker As New Worker()
        Dim strFile = GetCurrentFile()
        If String.IsNullOrEmpty(strFile) = False Then

            Dim thread As New Thread(New ParameterizedThreadStart(AddressOf worker.Impl_NonMSXlsxExtractData2))
            thread.Start(strFile)

        End If

    End Sub

    Public Sub On_CommandBar_Button_Zip_Repair_Try_Clicked(objButton As CommandBarButton, ByRef bCancelDefault As Boolean)

        Dim worker As New Worker()
        Dim strFile = GetCurrentFile()
        If String.IsNullOrEmpty(strFile) = False Then

            Dim thread As New Thread(New ParameterizedThreadStart(AddressOf worker.Impl_ZipRepairTry))
            thread.Start(strFile)

        End If

    End Sub

    Public Sub On_CommandBar_Button_Excel_Fix_Clicked(objButton As CommandBarButton, ByRef bCancelDefault As Boolean)

        Dim worker As New Worker()
        Dim strFile = GetCurrentFile()
        If String.IsNullOrEmpty(strFile) = False Then

            Dim thread As New Thread(New ParameterizedThreadStart(AddressOf worker.Impl_ExcelFix))
            thread.Start(strFile)

        End If

    End Sub

    Public Sub On_CommandBar_Button_SelectFile_Clicked(objButton As CommandBarButton, ByRef bCancelDefault As Boolean)

        SelectFile()

    End Sub

    Public Sub On_CommandBar_Button_PreviousVersion_Clicked(objButton As CommandBarButton, ByRef bCancelDefault As Boolean)

        FindPreviousVersions()

    End Sub

#End Region

#Region "CommandBar Images Events"

    Public Function On_CommandBar_Button_SaveAsSYLK_GetImage() As IPictureDisp

        On_CommandBar_Button_SaveAsSYLK_GetImage = ImageHelper.GetIPicture(My.Resources.SaveAsSYLK)

    End Function

    Public Function On_CommandBar_Button_Manual_Calculations_GetImage() As IPictureDisp

        On_CommandBar_Button_Manual_Calculations_GetImage = ImageHelper.GetIPicture(My.Resources.ManualCalculations)

    End Function

    Public Function On_CommandBar_Button_External_References_GetImage() As IPictureDisp

        On_CommandBar_Button_External_References_GetImage = ImageHelper.GetIPicture(My.Resources.ExternalReferences)

    End Function

    Public Function On_CommandBar_Button_Safe_Mode_GetImage() As IPictureDisp

        On_CommandBar_Button_Safe_Mode_GetImage = ImageHelper.GetIPicture(My.Resources.OpenInSafeMode)

    End Function

    Public Function On_CommandBar_Button_Save_As_HTML_GetImage() As IPictureDisp

        On_CommandBar_Button_Save_As_HTML_GetImage = ImageHelper.GetIPicture(My.Resources.SaveAsHTML)

    End Function

    Public Function On_CommandBar_Button_Open_With_WordPad_GetImage() As IPictureDisp

        On_CommandBar_Button_Open_With_WordPad_GetImage = ImageHelper.GetIPicture(My.Resources.OpenInWordPad)

    End Function

    Public Function On_CommandBar_Button_Open_With_Excel_Viewer_GetImage() As IPictureDisp

        On_CommandBar_Button_Open_With_Excel_Viewer_GetImage = ImageHelper.GetIPicture(My.Resources.OpenInExcel)

    End Function

    Public Function On_CommandBar_Button_MS_Open_And_Repair_GetImage() As IPictureDisp

        On_CommandBar_Button_MS_Open_And_Repair_GetImage = ImageHelper.GetIPicture(My.Resources.MSOpenAndRepair)

    End Function

    Public Function On_CommandBar_Button_MS_Open_And_Extract_Data_GetImage() As IPictureDisp

        On_CommandBar_Button_MS_Open_And_Extract_Data_GetImage = ImageHelper.GetIPicture(My.Resources.MSOpenAndExtract)

    End Function

    Public Function On_CommandBar_Button_Non_MS_XLSX_Extract_Data_GetImage() As IPictureDisp

        On_CommandBar_Button_Non_MS_XLSX_Extract_Data_GetImage = ImageHelper.GetIPicture(My.Resources.NonMSDataExtract)

    End Function

    Public Function On_CommandBar_Button_Non_MS_XLSX_Extract_Data_2_GetImage() As IPictureDisp

        On_CommandBar_Button_Non_MS_XLSX_Extract_Data_2_GetImage = ImageHelper.GetIPicture(My.Resources.NonMSDataExtract2)

    End Function

    Public Function On_CommandBar_Button_Zip_Repair_Try_GetImage() As IPictureDisp

        On_CommandBar_Button_Zip_Repair_Try_GetImage = ImageHelper.GetIPicture(My.Resources.ZipRepairTry)

    End Function

    Public Function On_CommandBar_Button_Excel_Fix_GetImage() As IPictureDisp

        On_CommandBar_Button_Excel_Fix_GetImage = ImageHelper.GetIPicture(My.Resources.ExcelFix)

    End Function

    Public Function On_CommandBar_Button_SelectFile_GetImage() As IPictureDisp

        On_CommandBar_Button_SelectFile_GetImage = ImageHelper.GetIPicture(My.Resources.SelectFile)

    End Function

    Public Function On_CommandBar_Button_PreviousVersion_GetImage() As IPictureDisp

        On_CommandBar_Button_PreviousVersion_GetImage = ImageHelper.GetIPicture(My.Resources.FindVersions)

    End Function

#End Region

#Region "CommandBar Combobox Events"

    Public Sub On_CommandBar_Combobox_CurrentFile_Change(objCombobox As CommandBarComboBox)

        Try

            If objCombobox.Text = _strClearList Then

                _arrFiles.Clear()

                _strCurrentFile = ""
                objCombobox.Text = ""

                SaveFilesList()
                LoadFilesList()

            Else

                Dim bExists As Boolean = False

                For Each strFullPath In _arrFiles

                    If strFullPath.EndsWith(objCombobox.Text) = True Then

                        _strCurrentFile = strFullPath
                        bExists = True

                        Exit For

                    End If

                Next

                If bExists = False Then

                    _strCurrentFile = objCombobox.Text

                End If

            End If

        Catch
        End Try

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

            If _arrFiles.Count > 0 Then
                _strCurrentFile = CType(_arrFiles(0), String)
            End If

            '----------------------------------------------------------------------------------------------------------------

            _cmb.Clear()

            For Each strFilePath In _arrFiles

                Try
                    _cmb.AddItem(Path.GetFileName(strFilePath))
                Catch ex As Exception

                End Try

            Next

            _arrFiles.Add(_strClearList)

            _cmb.AddItem(_strClearList)
            _cmb.Text = Path.GetFileName(_strCurrentFile)

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

        If (String.IsNullOrEmpty(_strCurrentFile) = True) Or (_strCurrentFile = _strClearList) Then
            _strCurrentFile = SelectFile()
            _cmb.Text = Path.GetFileName(_strCurrentFile)
        End If

        Return _strCurrentFile

    End Function

    Private Function SelectFile() As String

        Dim strResult As String = ""

        Dim dlgOpenFile As New OpenFileDialog()

        dlgOpenFile.Filter = "All files (*.*)|*.*||"

        If dlgOpenFile.ShowDialog() = DialogResult.OK Then

            _strCurrentFile = Path.GetFileName(dlgOpenFile.FileName)
            _cmb.Text = Path.GetFileName(_strCurrentFile)

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
            LoadFilesList()

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

    Protected Function FindCommandBar(objCommandBars As CommandBars, strTargetName As String) As CommandBar

        Dim objContextMenu As CommandBar = Nothing

        Try

            If objCommandBars IsNot Nothing Then

                Dim strName As String

                strTargetName = strTargetName.ToLower()

                objContextMenu = Nothing

                For Each objBar As CommandBar In objCommandBars

                    If String.IsNullOrEmpty(objBar.Name) = False Then

                        strName = objBar.Name.ToLower()

                        If strName = strTargetName Then

                            objContextMenu = objBar
                            Exit For

                        End If

                    End If

                Next

            End If

        Catch
        End Try

        Return objContextMenu

    End Function

#End Region

End Class
