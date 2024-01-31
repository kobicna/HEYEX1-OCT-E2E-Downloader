Attribute VB_Name = "ExportE2EDataOut"
Declare PtrSafe Function ActivateKeyboardLayout Lib "user32.dll" (ByVal myLanguage As Long, Flag As Boolean) As Long
Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)





Sub ExportE2EData()
    Dim strPath As String 'Filesystem object
    Dim intCountRows As Integer
    Dim counter As Integer
    Dim folderPath As String
    Dim EventHappend As Integer
    Dim fso As New FileSystemObject
    Dim Timein As Date
    Dim Timeout As Date
    Dim HasData As Boolean
    Dim DownloadedE2E As Integer
    HasData = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    DownloadedE2E = 0
    counter = 0
    Call ActivateKeyboardLayout(1033, 0)
    Application.FileDialog(msoFileDialogFolderPicker).Title = "Select a Path"  'the dialog is displayed to the user
    intResult = Application.FileDialog(msoFileDialogFolderPicker).Show
    'checks if user has cancled the dialog
    If intResult <> 0 And CheckIfWindow("Heidelberg Eye Explorer") = 1 Then 'checks if "Heidelberg Eye Explorer" is open
        

        strPath = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
        MsgBox strPath  'show folder path were will E2E will be exported to
        intCountRows = Sheets("ToDownload").Range("$A$2:" & Cells(Sheets("ToDownload").Cells(Rows.Count, 1).End(xlUp).Row, 1).Address).Cells.Count
        encryptDOB = "01/01/1900" ' encrypted date of birth
        

        'Sheets("ToDownload").Range("A2") = Sheets("ToDownload").Range("A2")
        Do While Sheets("ToDownload").Range("$A$2").Value <> "" Or counter < intCountRows ' A while loop to work till all rows were deleted (moved to "Downloaded" sheet) or counter got to the number of rows
            If CheckIfWindow("Heidelberg Eye Explorer") <> 1 Then 'for each loop checks again if "Heidelberg Eye Explorer" is open, if not, exits the sub
                MsgBox "Heidelberg Eye Explorer was closed - program terminates"
                ActiveWorkbook.Save
                Exit Sub
            End If
                       
            ActiveWorkbook.Save ' Save the workbook for every run.
            cellID = Sheets("ToDownload").Range("A2").Value
            encryptVal = Sheets("ToDownload").Range("A2").Offset(0, 1).Value ' copy the encrypted name in order to fill the form of extracted data
            Call MouseMoveA2B(0, 0, 68, 101) ' moves the mouse to Search patient name text box and clicks
            Call MouseCureent2a(68, 101)
            Call MouseCureent2a(68, 101)
            Call SingleLeftClick
            Call SingleLeftClick
            Application.SendKeys "^a" ' Ctrl+A in order to mark all data in search patient
            Application.SendKeys Format(Sheets("ToDownload").Range("A2").Value, "$0000000000")    'Insert the new patient to the Search patient name text box
            Application.Wait Now() + TimeSerial(0, 0, 1)
            Application.SendKeys "~" ' Click keyboard Enter to make sure name is being searched
            Application.Wait Now() + TimeSerial(0, 0, 5) ' Waits a reasnabole time to see of the patient names appear
            Call MouseCureent2a(68, 170)  'Move from Name to Patient General where If the patient exists in the database should apear
            Call SingleLeftClick
            Call SingleLeftClick    ' double click on Patient name in oreder to load data into the remote "Heidelberg Eye Explorer"
            i = 0
            Do 'wait till patient apear in loaded data at the "Heidelberg Eye Explorer" or 45sec has passed (it time passes skip to next patient)
                DoEvents
                Application.Wait Now() + TimeSerial(0, 0, 1)
                i = i + 1
            Loop Until (getScreenPixel(991, 170) = 0 Or i = 45)
            i = 0
            If getScreenPixel(991, 170) = 0 Then ' check if the patient, check if the patient was loaded, if not, passes to next one
                 'Create a folder if it does not already exist, if it does, do nothing
                folderPath = strPath & "\" & encryptVal 'Check if the folder exists
                HasData = True
                If Dir(folderPath, vbDirectory) = "" Then
                    'Folder does not exist, so create it
                    'MsgBox folderPath & " does not exist. and was created"
                    MkDir folderPath
                'Else
                    'MsgBox folderPath & " exists."
                End If
                Timein = Now ' time download has started
                Sheets("ToDownload").Range("A2").Offset(0, 3).Value = Format(Timein, "mm/dd/yyyy HH:mm:ss")
                Call MouseCureent2a(1012, 170)  'Move from Patient General to Patient Specific
                Application.Wait Now() + TimeSerial(0, 0, 5)
                Call SingleRightClick
                'Application.Wait Now() + TimeSerial(0, 0, 5)
                'Call waitForPixelToEqual(1053, 188, 16579836) ' waif for right click to open menu
                i = 0
                Do 'wait till patient apear in loaded data at the "Heidelberg Eye Explorer" or 45sec has passed (it time passes skip to next patient)
                    DoEvents
                    Application.Wait Now() + TimeSerial(0, 0, 1)
                    i = i + 1
                Loop Until (getScreenPixel(1030, 250) = 15790320 Or i = 7)
                If Not getScreenPixel(1030, 250) = 15790320 Then
                    GoTo nextPatient
                    Sheets("ToDownload").Range("A2").Offset(0, 2).Value = "Skipped"
                End If
                i = 0
                Call MouseCureent2a(1090, 254)  'Move from Patient Specific  to Export Button
                Application.Wait Now() + TimeSerial(0, 0, 1)
                Call MouseCureent2a(1255, 254)  'Move from Export Button to E2E Butoon
                Call SingleLeftClick
                'Application.Wait Now() + TimeSerial(0, 0, 30)
                Call waitForWindow("Export Options") ' wait for Export options to open
                Call MouseCureent2a(796, 542) 'Move from E2E Butoon to Anonymize Data Button
                Call SingleLeftClick
                Application.Wait Now() + TimeSerial(0, 0, 10)
                Call MouseCureent2a(900, 471)  'Move to Patient ID
                Application.Wait Now() + TimeSerial(0, 0, 3)
                Call SingleLeftClick
                Call SingleLeftClick
                Application.Wait Now() + TimeSerial(0, 0, 1)
                Application.SendKeys encryptVal
                Call MouseCureent2a(885, 495)  'Move from Patient ID to PatientDOB
                Call SingleLeftClick
                Call SingleLeftClick
                Application.Wait Now() + TimeSerial(0, 0, 1)
                Application.SendKeys encryptDOB
                Call MouseCureent2a(796, 542)  'Move from PatientDOB to Anonymize Data Button
                Call SingleLeftClick
                Application.Wait Now() + TimeSerial(0, 0, 1)
                Call MouseCureent2a(890, 300)  'Move from  Anonymize Data Button to Export Path
                Call SingleLeftClick
                Call SingleLeftClick
                Application.SendKeys "^a"
                Application.Wait Now() + TimeSerial(0, 0, 1)
                Application.SendKeys folderPath & "\" & encryptVal & ".E2E"
                Application.Wait Now() + TimeSerial(0, 0, 1)
                Call MouseCureent2a(914, 778)  'Move from Export Path to Ok Button
                Application.Wait Now() + TimeSerial(0, 0, 1)
                Call SingleLeftClick
                Application.Wait Now() + TimeSerial(0, 0, 2)
                Call MouseCureent2a(971, 571)
                Call SingleLeftClick
                Application.Wait Now() + TimeSerial(0, 0, 2)
eent:
                i = 0
                EventHappend = 0
                Do
                DoEvents
                
                    'Application.StatusBar = "EventHappend=" & EventHappend
                    If CheckIfWindow("Heidelberg Eye Explorer") <> 1 Then
                        MsgBox "Heidelberg Eye Explorer was closed - program terminates"
                        'Application.ScreenUpdating = True
                        ActiveWorkbook.Save
                        Exit Sub
                    End If
                    'Dealing with Errors or Warnings , Finiding pixel of sign and trying to click OK button
                    If CheckIfWindow("Error") = 1 Then
                        Application.SendKeys "~"
                    End If
                    If getScreenPixel(793, 490) = 57852 And getScreenPixel(770, 460) = 16777215 Then
                            Call MouseCureent2a(1029, 597)
                            Call SingleLeftClick
                    End If
                    If getScreenPixel(832, 496) = 16777215 And getScreenPixel(997, 597) = 14803425 And getScreenPixel(768, 464) = 10724259 Then
                            Call MouseCureent2a(1029, 597)
                            Call SingleLeftClick
                    End If
                    If getScreenPixel(805, 494) = 57852 And getScreenPixel(1000, 600) = 14803425 Then ' Warning occured and button at 1000, 600 is at color 14803425 click button

                            Call MouseCureent2a(1029, 597)
                            Call SingleLeftClick
                    End If
                    If EventHappend = 0 And CheckIfWindow("Export E2E Files") = 1 Then 'Multiple files are downloading, moving to EventHappend = 2 (wait for download to finish and click OK button)
                            EventHappend = 2
                    ElseIf EventHappend = 2 And getScreenPixel(1029, 686) = 14803425 Then 'If download finished -> click OK button
                        Call MouseCureent2a(1029, 686)
                        Call SingleLeftClick
                        EventHappend = 100
                    ElseIf getScreenPixel(880, 531) = 14120960 And getScreenPixel(880, 458) = 15790320 Then 'The program thinks only 1 file is  downloading, moving to EventHappend = 3 (check if download finished or if it is multiple files download)
                        EventHappend = 3
                    ElseIf EventHappend = 3 Then ' If program thought it was 1 file to download but eventually had been multiple go back to multiple file download event
                        If CheckIfWindow("Export E2E Files") = 1 Then
                            GoTo eent
                        End If
                        If getScreenPixel(880, 531) = 16777215 And CheckIfWindow("Export E2E Files") <> 1 Then ' 1 file finished -> click OK button)
                            EventHappend = 100
                        End If
                    ElseIf EventHappend = 0 And i < 3 Then ' In case no screen was opened and nothing is happening wair for 45 sec (15sec * 3)
                        Call MouseCureent2a(200, 2)
                        Call SingleLeftClick
                        i = i + 1
                        GetWindows2
                        Application.Wait Now() + TimeSerial(0, 0, 15)
                        If i = 3 Then
                            EventHappend = 100
                        End If
                    End If
                    Call MouseCureent2a(200, 2)
                    Call SingleLeftClick
                    Application.Wait Now() + TimeSerial(0, 0, 15)
                    Call MouseCureent2a(400, 2)
                    Application.Wait Now() + TimeSerial(0, 0, 15)
                    Call SingleLeftClick
                    Loop Until EventHappend = 100 ' finished download -> next patient
                

                

                Sheets("ToDownload").Range("A2").Offset(0, 1).Interior.Color = RGB(241, 175, 90)
                Sheets("ToDownload").Range("A2").Offset(0, 2).Value = folderPath ' & "\" & encryptVal
                
                
                    
            
            End If
nextPatient:
            
            Call MouseCureent2a(1012, 170) 'Move from Patient General to Patient Specific
            Application.Wait Now() + TimeSerial(0, 0, 1)
            Call SingleRightClick
            Call waitForPixelToEqual(1029, 68, 15790320)
            Call MouseCureent2a(1056, 189)  'Move from Patient Specific to Unload
            Call SingleLeftClick
            If HasData = True Then
                Timeout = Now
                Sheets("ToDownload").Range("A2").Offset(0, 4).Value = Format(Timeout, "mm/dd/yyyy HH:mm:ss")
                Sheets("ToDownload").Range("A2").Offset(0, 5).Value = Format(Timeout - Timein, "HH:mm:ss")
    
                Sheets("ToDownload").Range("A2").Offset(0, 6).Value = CountFiles_FolderAndSubFolders(folderPath)
                Sheets("ToDownload").Range("A2").Offset(0, 7).Value = FolderSize(folderPath)
                DownloadedE2E = DownloadedE2E + 1
            End If
            Sheets("ToDownload").Range(Sheets("ToDownload").Range("A2").Offset(0, 1), Sheets("ToDownload").Range("A2").Offset(0, 7)).Copy
            Sheets("Downloaded").Cells(Rows.Count, "A").End(xlUp).Offset(1).PasteSpecial xlPasteValuesAndNumberFormats
            Sheets("ToDownload").Range("A2").EntireRow.Delete Shift:=xlUp
            HasData = False
            counter = counter + 1
            Application.StatusBar = "Downloading #" & DownloadedE2E & " out of " & intCountRows - (counter - DownloadedE2E) & " Patients data, " & intCountRows - counter & " left| Skipped (Had no OCT): " & (counter - DownloadedE2E)
            Application.CutCopyMode = False
        Loop 'Next Sheets("ToDownload").Range("A2")

        MsgBox DownloadedE2E & " of " & intCountRows & " ID's with oct were found, Skipped (Had no OCT): " & (counter - DownloadedE2E)
    End If

               
End Sub
