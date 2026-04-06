Option Explicit On

Imports System.IO
Imports System.Drawing
Imports System.Security
Imports System.Reflection
Imports System.Windows.Forms
Imports System.ComponentModel
Imports System.Collections.Generic
Imports System.Text.RegularExpressions

Imports System.Runtime.InteropServices

Imports Microsoft.Win32

Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Interop.Excel

Imports Word = Microsoft.Office.Interop.Word
Imports Excel = Microsoft.Office.Interop.Excel

Public Class FormMain

#Region "Data members"

    Dim filename As String

    Dim WithEvents _timer As Timer

    Dim counterVariable As Integer
    Dim previousVersionCounterVariable As Integer
    Dim sFileShadowPath As String
    Dim sFileShadowName As String
    Dim sFileShadowPathDate As String
    Dim matchCount As Integer
    Dim shadowLinkFolderName As New List(Of String)
    Dim nonErrorShadowPathList As New List(Of String)
    Dim comboBoxIndex As Integer = 0
    Dim preVersionHashTable As New Hashtable

    Private allowCoolMove As Boolean = False
    Private dx, dy As Integer 'I used this two integers as I could use the function new POint due to the Import of the Excel

    Dim saveShadowPath As String
    Dim selectedsFileShadowPathDate As String
    Dim selectedPreviousVersion As String

    Public objExcel As Excel.Application

#End Region

#Region "Events"

    Public Sub New()

        InitializeComponent()

        _timer = New Timer()
        _timer.Interval = 50

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load

        If String.IsNullOrEmpty(PathTb.Text) = False Then

            _timer.Start()

        End If

    End Sub

    Private Sub Form1_FormClosing(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.FormClosing

        Try

            Dim x As Integer

            For x = 0 To shadowLinkFolderName.Count - 1

                'Starts a command line in the background and removes the temporary folders mapped to the restore point snapshots.

                Dim myProcess As New Process()
                myProcess.StartInfo.FileName = "cmd.exe"
                myProcess.StartInfo.UseShellExecute = False
                myProcess.StartInfo.RedirectStandardInput = True
                myProcess.StartInfo.RedirectStandardOutput = True
                myProcess.StartInfo.CreateNoWindow = True
                myProcess.Start()

                Dim myStreamWriter As StreamWriter = myProcess.StandardInput

                myStreamWriter.WriteLine("rmdir " & shadowLinkFolderName(x))
                myStreamWriter.Close()
                myProcess.WaitForExit()
                myProcess.Close()

            Next

        Catch
        End Try

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

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim sFile As String = PathTb.Text

            Dim oExcel As New Excel.Application
            Dim oBooks As Excel.Workbooks = Nothing
            Dim oWB As Workbook = Nothing

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


        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim x As Integer

            For x = 0 To shadowLinkFolderName.Count - 1

                Dim myProcess As New Process()
                myProcess.StartInfo.FileName = "cmd.exe"
                myProcess.StartInfo.UseShellExecute = False
                myProcess.StartInfo.RedirectStandardInput = True
                myProcess.StartInfo.RedirectStandardOutput = True
                myProcess.StartInfo.CreateNoWindow = True
                myProcess.Start()
                Dim myStreamWriter As StreamWriter = myProcess.StandardInput

                myStreamWriter.WriteLine("rmdir " & shadowLinkFolderName(x))
                myStreamWriter.Close()
                myProcess.WaitForExit()
                myProcess.Close()

            Next

            Me.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Public Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim oExcel As Excel.Application
        Dim oBook As Excel.Workbook
        Dim oBooks As Excel.Workbooks

        Dim sFile As String = PathTb.Text

        'Start Excel and open the workbook.

        oExcel = CreateObject("Excel.application")
        oExcel.Visible = True

        oBooks = oExcel.Workbooks
        oBook = oBooks.Open(sFile)

        oExcel.Run("OpenAndRepairWorkbook")

        ' Clean-up: Close the workbook and quit Excel.
        oBook.Close(False)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook)
        oBook = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oBooks)
        oBooks = Nothing
        oExcel.Quit()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel)
        oExcel = Nothing

        GC.Collect()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Try
            Dim OFD As New OpenFileDialog
            If OFD.ShowDialog() = DialogResult.OK Then

                filename = OFD.FileName
                PathTb.Text = OFD.FileName

                FindPreviousVersions(PathTb.Text)

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Label7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub PictureBox12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim TargetKey As RegistryKey
            TargetKey = Registry.ClassesRoot.OpenSubKey("excel.application")

            If TargetKey Is Nothing Then
                MsgBox("The Macro Graph Data Recovery method requires Excel, however it does not appear to be installed.")
                Exit Sub

            Else
                'key is found
                TargetKey.Close()
                Dim oExcel As New Excel.Application
                Dim oBooks As Excel.Workbooks = Nothing
                Dim oBook As Workbook = Nothing
                Dim oWSheet As Worksheet = Nothing
                Dim oChart As Chart = Nothing
                Dim sFile As String = PathTb.Text
                Dim NumberOfRows As Integer
                Dim X As Object
                Dim Counter As Integer = 2

                oExcel.Visible = True
                oBooks = oExcel.Workbooks
                oBook = oBooks.Open(Filename:=sFile)
                oWSheet = oBook.Worksheets.Add()
                oWSheet.Name = "ChartData"
                oWSheet.Activate()
                MsgBox("Select the chart you wish to extract data from.")
                oChart = oBook.ActiveChart

                ' Calculate the number of rows of data. 
                NumberOfRows = UBound(oChart.SeriesCollection(1).Values)
                oWSheet.Cells(1, 1) = "X Values"

                ' Write x-axis values to worksheet. 
                With oWSheet
                    .Range(.Cells(2, 1), _
                    .Cells(NumberOfRows + 1, 1)).Value = _
                    oExcel.WorksheetFunction.Transpose(oChart.SeriesCollection(1).XValues)
                End With

                ' Loop through all series in the chart   
                ' and write their values to the worksheet. 
                For Each X In oChart.SeriesCollection
                    oWSheet.Cells(1, Counter) = X.Name

                    With oWSheet
                        .Range(.Cells(2, Counter), _
                        .Cells(NumberOfRows + 1, Counter)).Value = _
                        oExcel.WorksheetFunction.Transpose(X.Values)
                    End With

                    Counter = Counter + 1
                Next
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub Label11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button1_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Dim saveFileDialog1 As New SaveFileDialog()
            saveFileDialog1.Filter = "Excel Files (*.xls;*.xlsx;*.xlt;*.xla;" _
                & "*.xlsm;*.xltx;*.xltm;*.xlsb;*.xlam)|*.xls;*.xlsx;*.xlt;*.xla;" _
                & "*.xlsm;*.xltx;*.xltm;*.xlsb;*.xlam|All Files (*.*)|*.*"
            saveFileDialog1.FilterIndex = 1
            saveFileDialog1.RestoreDirectory = True

            If saveFileDialog1.ShowDialog() = DialogResult.OK Then
                saveShadowPath = saveFileDialog1.FileName

                If System.IO.File.Exists(selectedPreviousVersion) = True Then
                    System.IO.File.Copy(selectedPreviousVersion, saveShadowPath, True)
                    MsgBox(("The Previous version of " & sFileShadowName _
                                  & " last modified on " & selectedsFileShadowPathDate _
                                  & " was saved to a new location: " & saveShadowPath) & ".")
                Else
                    MsgBox("Can't connect to previous version file.")
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            Dim sFileShadowPathInfo As New FileInfo(sFileShadowPath)
            sFileShadowPathDate = sFileShadowPathInfo.LastWriteTime
            sFileShadowName = sFileShadowPathInfo.Name
            selectedsFileShadowPathDate = DelFromLeft("File Name: " & sFileShadowName _
                            & " Last Modified: ", ComboBox1.Text.ToString)
            selectedPreviousVersion = preVersionHashTable.Item(selectedsFileShadowPathDate)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Label12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub SaveFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs)

    End Sub

    Private Sub FolderBrowserDialog1_HelpRequest(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub PathTb_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PathTb.TextChanged

    End Sub

    Private Sub timer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _timer.Tick

        _timer.Stop()

        FindPreviousVersions(PathTb.Text)

    End Sub


#End Region

#Region "Implementation"

    Private Sub FindPreviousVersions(sFile As String)

        Try

            If Worker.IsRunningAsAdmin = False Then

                If Worker.RequestPriviledges() = True Then

                    objExcel.Quit()

                    Dim proc As New Process()
                    proc.StartInfo.UseShellExecute = True
                    proc.StartInfo.Verb = "runas"
                    proc.StartInfo.FileName = Process.GetCurrentProcess().MainModule.FileName
                    proc.Start()

                End If

                Me.Close()
                Exit Sub

            End If

            shadowLinkFolderName.Clear()
            nonErrorShadowPathList.Clear()
            ComboBox1.Items.Clear()
            'Find out the number of vss shadow snapshots (restore 
            'points). All shadows apparently have a linkable path 
            '\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy#,
            'where # is a simple one or two or three digit integer.

            Dim strPassword As New SecureString()

            Dim objProcess As New Process()
            objProcess.StartInfo.UseShellExecute = False
            objProcess.StartInfo.RedirectStandardOutput = True
            objProcess.StartInfo.CreateNoWindow = True
            objProcess.StartInfo.RedirectStandardError = True
            objProcess.StartInfo.FileName() = "vssadmin"
            objProcess.StartInfo.Arguments() = "List Shadows"
            objProcess.Start()

            Dim vssadminOutput As String = objProcess.StandardOutput.ReadToEnd
            Dim strError As String = objProcess.StandardError.ReadToEnd()
            objProcess.WaitForExit()

            ' Call Regex.Matches method.
            Dim matches As MatchCollection = Regex.Matches(vssadminOutput, _
                "\\\\\?\\GLOBALROOT\\Device\\HarddiskVolumeShadowCopy[0-9]+")
            counterVariable = 0
            matchCount = matches.Count
            MsgBox("Please wait while Excel Recovery searches for previous versions of your file.")
            ' Loop over matches.
            For Each m As Match In matches
                Dim driveLetter As String = sFile.Substring(0, 2)
                shadowLinkFolderName.Add("C:\" & DelFromLeft( _
                    "\\?\GLOBALROOT\Device\HarddiskVolume", (m.ToString())))
                sFileShadowPath = (shadowLinkFolderName(counterVariable) & DelFromLeft( _
                    driveLetter, sFile))

                'Here I create temporary folders off the C: 
                'drive which are mapped to each snapshot.
                Dim myProcess As New Process()
                myProcess.StartInfo.FileName = "cmd.exe"
                myProcess.StartInfo.UseShellExecute = False
                myProcess.StartInfo.RedirectStandardInput = True
                myProcess.StartInfo.RedirectStandardOutput = True
                myProcess.StartInfo.CreateNoWindow = True
                myProcess.Start()
                Dim myStreamWriter As StreamWriter = myProcess.StandardInput

                myStreamWriter.WriteLine("mklink /d " & (shadowLinkFolderName(counterVariable).ToString) _
                                        & " " & (m.ToString()) & "\")
                myStreamWriter.Close()
                myProcess.WaitForExit()
                myProcess.Close()

                'Here I compare our recovery target file against the shadow 
                'copies. One shadow file copy is compared for each iteration 
                'of the loop. If the string "no difference encountered is found" 
                'then I know this shadow copy of the file is not worth looking 
                'at, as it is the same as the recovery target.
                Dim fileCompare As New Process()
                fileCompare.StartInfo.FileName = "cmd.exe"
                fileCompare.StartInfo.UseShellExecute = False
                fileCompare.StartInfo.RedirectStandardInput = True
                fileCompare.StartInfo.RedirectStandardOutput = True
                fileCompare.StartInfo.CreateNoWindow = True
                fileCompare.Start()
                Dim fileCompareWriter As StreamWriter = fileCompare.StandardInput

                fileCompareWriter.WriteLine("fc """ & sFile & """ """ _
                                & sFileShadowPath & """")
                fileCompareWriter.Dispose()
                fileCompareWriter.Close()
                Dim fileCompareReader As StreamReader = fileCompare.StandardOutput
                Dim fileCompareOut As String = fileCompareReader.ReadToEnd
                fileCompareReader.Close()
                fileCompare.WaitForExit()
                fileCompare.Close()
                Dim fileCompareBoolean As Boolean = fileCompareOut.Contains( _
                                "no differences encountered").ToString
                Dim fileCompBooleanError As Boolean = fileCompareOut.Contains( _
                                "FC: cannot open").ToString
                If fileCompBooleanError = "True" Then
                    counterVariable = counterVariable + 1
                    Continue For
                End If

                If fileCompareBoolean = "True" Then
                    counterVariable = counterVariable + 1
                    Continue For
                End If
                'Here I take a positive result of a file difference between
                'the target and the shadow copy, and I write it out to a combo 
                'box on the form, so it can be chosen. I also only keep the 
                'first instance of a different shadow file as the others are 
                'identical. I distinguish if they are the same by date.


                Dim sFileShadowPathInfo As New FileInfo(sFileShadowPath)
                sFileShadowPathDate = sFileShadowPathInfo.LastWriteTime
                sFileShadowName = sFileShadowPathInfo.Name

                If ComboBox1.Items.Count = 0 Then
                    ComboBox1.Items.Add("File Name: " & sFileShadowName _
                                                            & " Last Modified: " & sFileShadowPathDate)
                    nonErrorShadowPathList.Add(sFileShadowPath)
                    preVersionHashTable.Add(sFileShadowPathDate, sFileShadowPath)
                    counterVariable = counterVariable + 1
                    Continue For
                End If

                previousVersionCounterVariable = ComboBox1.Items.Count - 1
                Dim previoussFileShadowPath As String
                previoussFileShadowPath = nonErrorShadowPathList(previousVersionCounterVariable)
                Dim prevsFileShadowPathInfo As New FileInfo(previoussFileShadowPath)
                Dim prevsFileShadowPathDate As String = prevsFileShadowPathInfo.LastWriteTime

                If String.Equals(sFileShadowPathDate, prevsFileShadowPathDate) Then
                    counterVariable = counterVariable + 1
                    Continue For
                Else
                    ComboBox1.Items.Add("File Name: " & sFileShadowName _
                                                            & " Last Modified: " & sFileShadowPathDate)
                    nonErrorShadowPathList.Add(sFileShadowPath)
                    preVersionHashTable.Add(sFileShadowPathDate, sFileShadowPath)
                    counterVariable = counterVariable + 1
                    Continue For
                End If

            Next m
            MsgBox("Processing has finished and should have returned previous versions, if they exist.")

        Catch ex As Exception
            MessageBox.Show(ex.Message)
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

#End Region

End Class
