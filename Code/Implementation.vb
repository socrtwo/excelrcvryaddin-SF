Option Explicit On

Imports System.IO
Imports System.Drawing
Imports System.Threading
Imports System.Windows.Forms
Imports System.ComponentModel
Imports System.Security.Principal
Imports System.Collections.Generic
Imports System.Text.RegularExpressions

Imports System.Runtime.InteropServices

Imports Microsoft.Win32

Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Interop.Excel

Imports Word = Microsoft.Office.Interop.Word
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Reflection

Module Implementation

    Public Class Worker

#Region "Implementation"

        Public Sub Impl_SaveAsHTML(sFile As String)

            Dim formProgress As FormProgress = InitProgressForm(4)

            Dim hWnd As IntPtr = FindExcelWindow()

            Try

                Dim TargetKey As RegistryKey
                TargetKey = Registry.ClassesRoot.OpenSubKey("excel.application")
                'Look to see if Excel is installed.
                If TargetKey Is Nothing Then
                    MsgBox("The save to SYLK format method requires Excel, however it does not appear to be installed.")
                    Exit Sub

                Else

                    If PerformProgressStep(formProgress) Then

                        'Key is found
                        TargetKey.Close()
                        Dim oExcel As New Excel.Application
                        Dim oBooks As Excel.Workbooks = Nothing
                        Dim oBook As Workbook = Nothing
                        Dim oWSheet As Worksheet = Nothing
                        Dim sFileName As String = Path.GetFileNameWithoutExtension(sFile)
                        Dim sDirName As String = Path.GetDirectoryName(sFile)
                        Dim sFileHTMLName As String = sDirName & "\" & sFileName & ".html"

                        If PerformProgressStep(formProgress) Then

                            'Start Excel and open the workbook. Then save to SYLK format.
                            oExcel.Visible = False
                            oBooks = oExcel.Workbooks
                            oBook = oBooks.Open(Filename:=sFile)
                            oWSheet = oBook.ActiveSheet()
                            oWSheet.SaveAs(Filename:=sFileHTMLName, FileFormat:=Excel.XlFileFormat.xlHtml)

                            If PerformProgressStep(formProgress) Then

                                'Open SYLK formatted file.
                                If File.Exists(sFileHTMLName) Then
                                    oBook.Close()
                                    oExcel.Visible = True
                                    oBook = oBooks.Open(Filename:=sFileHTMLName)

                                    Try
                                        oBook.Activate()
                                        Thread.Sleep(500)
                                        DeactivateExcel(hWnd)
                                    Catch
                                    End Try

                                Else
                                    MessageBox.Show("Failed to create " & sFileHTMLName)
                                End If

                                PerformProgressStep(formProgress)

                            End If

                        End If

                    End If

                End If

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            CleanupProgressForm(formProgress)

        End Sub

        Public Sub Impl_SafeMode(sFile As String)

            Dim formProgress As FormProgress = InitProgressForm(2)

            Try
                Dim excelPath As String = _
                    Registry.GetValue( _
                    "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe", _
                    "Path", "Key does not exist")

                Dim TargetKey As RegistryKey
                TargetKey = Registry.ClassesRoot.OpenSubKey("excel.application")

                If TargetKey Is Nothing Then
                    MsgBox("The Open in Safe Mode method requires Excel, however it does not appear to be installed.")
                    Exit Sub

                Else    'key is found
                    TargetKey.Close()

                    If PerformProgressStep(formProgress) Then

                        Shell(excelPath & "excel.exe /s /r " & """" & sFile & """", AppWinStyle.NormalFocus)

                        PerformProgressStep(formProgress)

                    End If

                End If

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            CleanupProgressForm(formProgress)

        End Sub

        Public Sub Impl_OpenWithExcelViewer(sFile As String)

            Dim formProgress As FormProgress = InitProgressForm(1)

            Try

                If File.Exists("C:\Program Files\Microsoft Office\Office12\XLVIEW.EXE") Then
                    Shell("C:\Program Files\Microsoft Office\Office12\XLVIEW.EXE /s /r " & """" & sFile & """", AppWinStyle.NormalFocus)

                Else

                    MsgBox("You have not downloaded the most recent Microsoft Excel Viewer. " _
                          & "Please install after downloading and try clicking this button again.")

                    System.Diagnostics.Process.Start("http://www.microsoft.com/download/en/details.aspx?id=10")

                End If

                PerformProgressStep(formProgress)

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            CleanupProgressForm(formProgress)

        End Sub

        Public Sub Impl_MSOpenAndExtractData(sFile As String)

            Dim formProgress As FormProgress = InitProgressForm(3)
            Dim hWnd As IntPtr = FindExcelWindow()

            Try
                Dim TargetKey As RegistryKey
                TargetKey = Registry.ClassesRoot.OpenSubKey("excel.application")

                If TargetKey Is Nothing Then
                    MsgBox("The Open and Extract Data method requires Excel, however it does not appear to be installed.")
                    Exit Sub

                Else    'key is found
                    TargetKey.Close()

                    If PerformProgressStep(formProgress) Then

                        Dim oExcel As New Excel.Application
                        Dim oBooks As Excel.Workbooks = Nothing
                        Dim oWB As Workbook = Nothing

                        oExcel.Visible = True

                        If PerformProgressStep(formProgress) Then

                            'Start Excel and open the workbook.

                            oBooks = oExcel.Workbooks
                            oWB = oBooks.Open(Filename:=sFile, CorruptLoad:=XlCorruptLoad.xlExtractData)

                            Try
                                DeactivateExcel(hWnd)
                                oWB.Activate()
                            Catch
                            End Try

                            PerformProgressStep(formProgress)

                        End If

                    End If

                End If

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            CleanupProgressForm(formProgress)

        End Sub

        Public Sub Impl_ManualCalculations(sFile As String)

            Dim formProgress As FormProgress = InitProgressForm(4)
            Dim hWnd As IntPtr = FindExcelWindow()

            Try
                Dim TargetKey As RegistryKey
                TargetKey = Registry.ClassesRoot.OpenSubKey("excel.application")

                If TargetKey Is Nothing Then
                    MsgBox("Opening the Excel file with calculations set to manual method requires Excel, however it does not appear to be installed.")
                    Exit Sub

                Else

                    If PerformProgressStep(formProgress) Then

                        'key is found
                        TargetKey.Close()
                        Dim oExcel As Excel.Application
                        Dim oBook As Excel.Workbook
                        Dim oBooks As Excel.Workbooks
                        oExcel = CreateObject("Excel.application")
                        oExcel.Visible = True

                        If PerformProgressStep(formProgress) Then

                            oBooks = oExcel.Workbooks
                            oBook = oBooks.Open(sFile)
                            oExcel.Calculation = Excel.XlCalculation.xlCalculationManual

                            Try
                                DeactivateExcel(hWnd)
                                oBook.Activate()
                            Catch
                            End Try

                            PerformProgressStep(formProgress)

                        End If

                    End If

                End If

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            CleanupProgressForm(formProgress)

        End Sub

        Public Sub Impl_MSOpenAndRepair(sFile As String)

            Dim formProgress As FormProgress = InitProgressForm(4)
            Dim hWnd As IntPtr = FindExcelWindow()

            Try
                Dim TargetKey As RegistryKey
                TargetKey = Registry.ClassesRoot.OpenSubKey("excel.application")

                If TargetKey Is Nothing Then
                    MsgBox("The Open and Repair method requires Excel, however it does not appear to be installed.")
                    Exit Sub

                Else
                    'key is found
                    TargetKey.Close()

                    If PerformProgressStep(formProgress) Then

                        Dim oExcel As New Excel.Application
                        Dim oBooks As Excel.Workbooks = Nothing
                        Dim oWB As Workbook = Nothing

                        If PerformProgressStep(formProgress) Then

                            'Start Excel and open the workbook.
                            If Path.GetExtension(sFile) = ".xls" Then
                                oExcel.Visible = True
                                oBooks = oExcel.Workbooks
                                oWB = oBooks.Open(Filename:=sFile, CorruptLoad:=XlCorruptLoad.xlRepairFile)
                                MsgBox("Excel completed file level validation and repair. Some " _
                                & "parts of this workbook may have been repaired or discarded.")

                            Else
                                oExcel.Visible = True
                                oBooks = oExcel.Workbooks
                                oWB = oBooks.Open(Filename:=sFile, CorruptLoad:=XlCorruptLoad.xlRepairFile)

                            End If

                            Try
                                DeactivateExcel(hWnd)
                                oWB.Activate()
                            Catch
                            End Try

                            PerformProgressStep(formProgress)

                        End If

                    End If

                End If

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            CleanupProgressForm(formProgress)

        End Sub

        Public Sub Impl_ExternalReferences(sFile As String)

            Dim formProgress As FormProgress = InitProgressForm(4)
            Dim hWnd As IntPtr = FindExcelWindow()

            Try
                Dim TargetKey As RegistryKey
                TargetKey = Registry.ClassesRoot.OpenSubKey("excel.application")

                If TargetKey Is Nothing Then
                    MsgBox("The External References method requires Excel, however it does not appear to be installed.")
                    Exit Sub

                Else
                    'key is found
                    TargetKey.Close()

                    If PerformProgressStep(formProgress) Then

                        Dim oExcel As Excel.Application
                        Dim oBook As Excel.Workbook

                        oExcel = CreateObject("Excel.application")
                        oBook = oExcel.Workbooks.Add
                        oExcel.Visible = True

                        Try
                            DeactivateExcel(hWnd)
                            oBook.Activate()
                        Catch
                        End Try

                        If PerformProgressStep(formProgress) Then

                            oExcel.Calculation = Excel.XlCalculation.xlCalculationAutomatic
                            oExcel.Range("A1").Value = "=" & "'" & sFile & "'" & "!A1"

                            Dim rRange As Excel.Range
                            oExcel.DisplayAlerts = False
                            rRange = oExcel.InputBox(Prompt:= _
                                "Please select a range similar in size to your corrupt data " _
                                & "that you wish to recover.", Title:="SPECIFY RANGE", Type:=8)
                            oExcel.DisplayAlerts = True

                            Try
                                DeactivateExcel(hWnd)
                                oBook.Activate()
                            Catch
                            End Try

                            If PerformProgressStep(formProgress) Then

                                If rRange Is Nothing Then

                                    Exit Sub

                                Else
                                    oExcel.Range("A1").Copy()
                                    oBook.ActiveSheet.Paste(rRange)

                                    DeactivateExcel(hWnd)
                                    oBook.Activate()

                                End If

                                PerformProgressStep(formProgress)

                            End If
                        End If

                    End If

                End If

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            CleanupProgressForm(formProgress)

        End Sub

        Public Sub Impl_SaveAsSYLK(sFile As String)

            Dim formProgress As FormProgress = InitProgressForm(4)

            Dim hWnd As IntPtr = FindExcelWindow()

            Try
                Dim TargetKey As RegistryKey
                TargetKey = Registry.ClassesRoot.OpenSubKey("excel.application")

                If TargetKey Is Nothing Then
                    MsgBox("The save to SYLK format method requires Excel, however it does not appear to be installed.")
                    Exit Sub

                Else
                    'key is found
                    TargetKey.Close()

                    If PerformProgressStep(formProgress) Then

                        Dim oExcel As New Excel.Application
                        Dim oBooks As Excel.Workbooks = Nothing
                        Dim oBook As Workbook = Nothing
                        Dim oWSheet As Worksheet = Nothing
                        Dim sFileName As String = Path.GetFileNameWithoutExtension(sFile)
                        Dim sDirName As String = Path.GetDirectoryName(sFile)
                        Dim sFileSylkName As String = sDirName & "\" & sFileName & ".slk"

                        'Start Excel and open the workbook.
                        oExcel.Visible = False

                        If PerformProgressStep(formProgress) Then

                            oBooks = oExcel.Workbooks
                            oBook = oBooks.Open(Filename:=sFile)
                            oWSheet = oBook.ActiveSheet()
                            oWSheet.SaveAs(Filename:=sFileSylkName, FileFormat:=Excel.XlFileFormat.xlSYLK)

                            If PerformProgressStep(formProgress) Then

                                If File.Exists(sFileSylkName) Then
                                    oBook.Close(False)
                                    oExcel.Visible = True
                                    oBook = oBooks.Open(Filename:=sFileSylkName)
                                Else
                                    MessageBox.Show("Failed to create " & sFileSylkName)
                                End If

                                Try
                                    DeactivateExcel(hWnd)
                                    oBook.Activate()
                                Catch
                                End Try

                                PerformProgressStep(formProgress)

                            End If

                        End If

                    End If

                End If

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            CleanupProgressForm(formProgress)

        End Sub

        Public Sub Impl_OpenWithWordPad(sFile As String)

            Dim formProgress As FormProgress = InitProgressForm(2)

            Try
                Dim regVersion As Microsoft.Win32.RegistryKey
                regVersion = Microsoft.Win32.Registry.LocalMachine.OpenSubKey( _
                    "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\wordpad.exe", False)
                If regVersion IsNot Nothing Then
                    Dim proc As New Process
                    With proc.StartInfo
                        .FileName = regVersion.GetValue("").ToString
                        .Arguments = Chr(34) + sFile + Chr(34)
                    End With

                    If PerformProgressStep(formProgress) Then

                        proc.Start()

                        PerformProgressStep(formProgress)

                    End If

                End If

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            CleanupProgressForm(formProgress)

        End Sub

        Public Sub Impl_NonMSXlsxExtractData(sFile As String)

            Dim formProgress As FormProgress = InitProgressForm(8)
            Dim hWnd As IntPtr = FindExcelWindow()

            Try

                Dim extractCMD As New Process()
                Dim myPercent As Char

                myPercent = Chr(37)
                Dim myZipCommand As String = """no-frills.exe " & Chr(37) _
                                & "a " & Chr(37) & "d " & Chr(37) & "f"""

                Dim strAddinPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)

                Dim strTempFileName As String = "ExcelRecoveryAddin.tmp"
                Dim strTempFilePath As String = Path.Combine(strAddinPath, strTempFileName)

                If File.Exists(strTempFilePath) Then
                    File.Delete(strTempFilePath)
                End If

                If PerformProgressStep(formProgress) Then

                    File.Copy(sFile, strTempFilePath)

                    If PerformProgressStep(formProgress) Then

                        extractCMD.StartInfo.FileName = Path.Combine(strAddinPath, "doctotext.exe")
                        extractCMD.StartInfo.Arguments = "--fix-xml --unzip-cmd=" & myZipCommand & " """ & strTempFileName & """"
                        extractCMD.StartInfo.UseShellExecute = False
                        extractCMD.StartInfo.RedirectStandardOutput = True
                        extractCMD.StartInfo.CreateNoWindow = True
                        extractCMD.StartInfo.WorkingDirectory = strAddinPath
                        extractCMD.Start()

                        If PerformProgressStep(formProgress) Then

                            Dim sFileText As String = sFile & ".txt"
                            Dim doctotextOutput As String = extractCMD.StandardOutput.ReadToEnd()

                            If PerformProgressStep(formProgress) Then

                                Dim sErr As String = ""
                                'Save to different file
                                Dim bAns As String = SaveTextToFile(doctotextOutput, sFileText, sErr)
                                If bAns Then

                                    If PerformProgressStep(formProgress) Then

                                        MsgBox("Please note: all successfully extracted worksheets will appear on just one worksheet traveling vertically down.")

                                        Dim oExcel As New Excel.Application
                                        oExcel.Visible = True

                                        If PerformProgressStep(formProgress) Then

                                            Dim oBooks As Workbooks = oExcel.Workbooks
                                            Dim oBook As Workbook = oBooks.Open(sFileText)

                                            DeactivateExcel(hWnd)

                                            If oBook IsNot Nothing Then
                                                Try
                                                    oBook.Activate()
                                                Catch
                                                End Try
                                            End If

                                            PerformProgressStep(formProgress)

                                            Marshal.FinalReleaseComObject(oBooks)
                                            Marshal.FinalReleaseComObject(oExcel)

                                            PerformProgressStep(formProgress)

                                        Else
                                            MsgBox("Error extracting file: " & sErr)
                                        End If

                                        extractCMD.Close()

                                        File.Delete(strTempFilePath)

                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            CleanupProgressForm(formProgress)

        End Sub

        Public Sub Impl_NonMSXlsxExtractData2(sFile As String)

            Dim formProgress As FormProgress = InitProgressForm(3)

            Try

                If sFile.EndsWith("xls", StringComparison.InvariantCultureIgnoreCase) = False Then

                    MsgBox("If successful, each Worksheet will be saved as separate " _
                           & "CSV files in the same directory as your corrupt file.")

                    Dim OFlD As New FolderBrowserDialog

                    Dim coffecCMD As New Process()

                    Dim strFullPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "coffec.exe")

                    coffecCMD.StartInfo.FileName = strFullPath
                    coffecCMD.StartInfo.Arguments = "-t """ & sFile & """"
                    coffecCMD.StartInfo.UseShellExecute = True
                    coffecCMD.StartInfo.CreateNoWindow = True
                    coffecCMD.Start()

                    If PerformProgressStep(formProgress) Then

                        coffecCMD.WaitForExit()
                        coffecCMD.Close()

                        If PerformProgressStep(formProgress) Then

                            Dim sFileInfo As New FileInfo(sFile)
                            Dim sFileName As String = sFileInfo.Name
                            Dim sFilePath As String = DelFromRight(sFileName, sFile)
                            Process.Start("explorer.exe", sFilePath)

                            PerformProgressStep(formProgress)

                        End If

                    End If

                Else

                    MsgBox("This tool doesn't support Excel 97-2003 workbooks")

                End If

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            CleanupProgressForm(formProgress)

        End Sub

        Public Sub Impl_ZipRepairTry(sFile As String)

            Dim formProgress As FormProgress = InitProgressForm(6)
            Dim hWnd As IntPtr = FindExcelWindow()

            Try
                Dim repairZip As New Process()
                Dim sFileZip As String = sFile & ".zip"
                Dim sFileInfo As New FileInfo(sFile)
                Dim sFileName As String = sFileInfo.Name
                Dim zipRepairedsFileName As String = "zipRepaired" & sFileName & ".zip"
                Dim sFileBasePath As String = DelFromRight(sFileName, sFile)
                Dim zipRepairedFullPathFileName As String = sFileBasePath & zipRepairedsFileName

                If File.Exists(sFileZip) Then
                    File.Delete(sFileZip)
                End If

                FileCopy(sFile, sFileZip)

                If PerformProgressStep(formProgress) Then

                    Dim strFullPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "zip.exe")

                    repairZip.StartInfo.FileName = strFullPath
                    repairZip.StartInfo.Arguments = "-FF """ & sFileZip & """ --out " & Chr(34) & zipRepairedFullPathFileName & Chr(34)
                    repairZip.StartInfo.UseShellExecute = False
                    repairZip.StartInfo.RedirectStandardOutput = True
                    repairZip.StartInfo.CreateNoWindow = True
                    repairZip.Start()

                    If PerformProgressStep(formProgress) Then

                        Dim repairZipReader As StreamReader = repairZip.StandardOutput
                        Dim repairZipCompOut As String = repairZipReader.ReadToEnd

                        If PerformProgressStep(formProgress) Then

                            repairZipReader.Close()
                            repairZip.WaitForExit()
                            repairZip.Close()

                            If PerformProgressStep(formProgress) Then

                                Dim zipRepairedFullPathXlsxName As String = DelFromRight(".zip", zipRepairedFullPathFileName)
                                Dim oExcel As New Excel.Application
                                Dim oBooks As Excel.Workbooks = Nothing
                                Dim oBook As Workbook = Nothing

                                If File.Exists(zipRepairedFullPathXlsxName) Then
                                    File.Delete(zipRepairedFullPathXlsxName)
                                End If

                                Rename(zipRepairedFullPathFileName, zipRepairedFullPathXlsxName)

                                oExcel.Visible = True

                                If PerformProgressStep(formProgress) Then

                                    oBooks = oExcel.Workbooks
                                    oBook = oBooks.Open(Filename:=zipRepairedFullPathXlsxName)

                                    Try
                                        DeactivateExcel(hWnd)
                                        oBook.Activate()
                                    Catch
                                    End Try

                                    PerformProgressStep(formProgress)

                                End If
                            End If
                        End If
                    End If
                End If

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            CleanupProgressForm(formProgress)

        End Sub

        Public Sub Impl_ExcelFix(sFile As String)

            Try
                Dim proc As New Process()
                proc.StartInfo.FileName = "http://www.cimaware.com/main/products/excelfix.php"
                proc.StartInfo.Arguments = ""
                proc.StartInfo.UseShellExecute = True
                proc.Start()
            Catch
            End Try

        End Sub

#End Region

#Region "Helpers"

        Public Function SaveTextToFile(ByVal strData As String, _
         ByVal FullPath As String, _
           Optional ByVal ErrInfo As String = "") As Boolean


            Dim bAns As Boolean = False
            Dim objReader As StreamWriter
            Try


                objReader = New StreamWriter(FullPath)
                objReader.Write(strData)
                objReader.Close()
                bAns = True
            Catch Ex As Exception
                ErrInfo = Ex.Message

            End Try
            Return bAns
        End Function

        Public Function DelFromRight(ByVal sChars As String, ByVal sLine As String) As String
            'Removes unwanted characters from right of given string
            ' EXAMPLE
            '  MsgBox DelFromRight(" TEST", "THIS IS A TEST")
            'displays "THIS IS A"



            sLine = ReverseString(sLine)
            sChars = ReverseString(sChars)
            sLine = DelFromLeft(sChars, sLine)
            DelFromRight = ReverseString(sLine)
            Exit Function


        End Function

        Public Function DelFromLeft(ByVal sChars As String, _
                ByVal sLine As String) As String

            ' Removes unwanted characters from left of given string
            '  EXAMPLE
            '      MsgBox DelFromLeft("THIS", "THIS IS A TEST")
            '        displays  "IS A TEST"


            Dim iCount As Integer
            Dim sChar As String

            DelFromLeft = ""
            ' Remove unwanted characters to left of folder name
            If InStr(sLine, sChars) > 0 Then
                For iCount = 1 To Len(sChars)
                    ' Retrieve character from start string to 
                    'look for in folder string (sLine)
                    sChar = Mid$(sChars, iCount, 1)
                    ' Remove all characters to left of found string
                    sLine = Mid$(sLine, InStr(sLine, sChar) + 1)

                Next iCount
            End If
            DelFromLeft = sLine
            Exit Function

        End Function

        Public Function ReverseString(ByVal InputString As String) _
          As String

            'If you have vb6, you can use
            'StrReverse instead of this function

            Dim lLen As Long, lCtr As Long
            Dim sChar As String
            Dim sAns As String = ""

            lLen = Len(InputString)
            For lCtr = lLen To 1 Step -1
                sChar = Mid(InputString, lCtr, 1)
                sAns = sAns & sChar
            Next

            ReverseString = sAns

        End Function

        Private Function ReadExeFromResources(ByVal filename As String) As Byte()
            Dim CurrentAssembly As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
            Dim Resource As String = String.Empty
            Dim ArrResources As String() = CurrentAssembly.GetManifestResourceNames()
            For Each Resource In ArrResources
                If Resource.IndexOf(filename) > -1 Then Exit For
            Next
            Dim ResourceStream As IO.Stream = CurrentAssembly.GetManifestResourceStream(Resource)
            If ResourceStream Is Nothing Then
                Return Nothing
            End If
            Dim ResourcesBuffer(CInt(ResourceStream.Length) - 1) As Byte
            ResourceStream.Read(ResourcesBuffer, 0, ResourcesBuffer.Length)
            ResourceStream.Close()
            Return ResourcesBuffer
        End Function

        Public Shared Function IsRunningAsAdmin() As Boolean

            Dim bIsAdmin As Boolean = False

            Try
                Dim identity As WindowsIdentity = WindowsIdentity.GetCurrent()

                If identity IsNot Nothing Then

                    Dim pricipal As WindowsPrincipal = New WindowsPrincipal(identity)

                    bIsAdmin = pricipal.IsInRole(WindowsBuiltInRole.Administrator)

                    pricipal = Nothing

                End If
            Catch
            End Try

            Return bIsAdmin

        End Function

        Public Shared Function RequestPriviledges() As Boolean

            Dim bResult As Boolean

            Dim result As DialogResult = MsgBox("This feature requires elevated privileges. Do you want to rerun Excel as Administrator?", MessageBoxButtons.YesNo)

            If result = DialogResult.Yes Then
                bResult = True
            Else
                bResult = False
            End If

            Return bResult

        End Function

        Public Function InitProgressForm(ByVal nSteps As Integer) As FormProgress

            Dim formProgress As New FormProgress()

            formProgress.TopMost = True
            'formProgress.progress.Maximum = nSteps
            formProgress.Show()

            System.Windows.Forms.Application.DoEvents()

            Return formProgress

        End Function

        Public Function PerformProgressStep(ByVal formProgress As FormProgress) As Boolean

            System.Windows.Forms.Application.DoEvents()

            If (formProgress._bStop = False) Then

                'formProgress.progress.PerformStep()

                System.Windows.Forms.Application.DoEvents()

                Return True

            Else

                Return False

            End If

        End Function

        Public Sub CleanupProgressForm(ByVal formProgress As FormProgress)

            formProgress.Close()

        End Sub

        Public Function FindExcelWindow() As IntPtr

            Dim hWnd As IntPtr = IntPtr.Zero

            Try
                hWnd = FindWindowByClass("XLMAIN", IntPtr.Zero)
            Catch
            End Try

            Return hWnd

        End Function

        Public Sub DeactivateExcel(hWnd As IntPtr)

            Try

                If hWnd <> IntPtr.Zero Then

                    ShowWindow(hWnd, 6)

                End If

            Catch ex As Exception

            End Try

        End Sub

        <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
        Private Shared Function ShowWindow(ByVal hwnd As IntPtr, ByVal nCmdShow As Integer) As Boolean
        End Function

        <DllImport("user32.dll", EntryPoint:="FindWindow", SetLastError:=True, CharSet:=CharSet.Auto)> _
        Private Shared Function FindWindowByClass(ByVal lpClassName As String, ByVal zero As IntPtr) As IntPtr
        End Function

#End Region

    End Class

End Module
