VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "RSC EZ-VIEW Builder Version 1.01"
   ClientHeight    =   9855
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   14235
   LinkTopic       =   "Form1"
   ScaleHeight     =   9855
   ScaleWidth      =   14235
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   22
      Top             =   9450
      Width           =   14235
      _ExtentX        =   25109
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   21167
            MinWidth        =   21167
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   9345
      Left            =   2910
      TabIndex        =   1
      Top             =   30
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   16484
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Data Source"
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "UsrData"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "EZ-VIEW CD"
      TabPicture(1)   =   "frmMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdEZVIEW(0)"
      Tab(1).Control(1)=   "fraOptions"
      Tab(1).Control(2)=   "fraTarget"
      Tab(1).Control(3)=   "fraSource"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "PSPA Year Book CD"
      TabPicture(2)   =   "frmMain.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "chkFieldCopy"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "dirCopy"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "drvCopy"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "txtNewCopyDir"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "chkCopy"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cmdTools(0)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      Begin VB.CommandButton cmdTools 
         Height          =   705
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Align images to database"
         Top             =   8520
         Width           =   825
      End
      Begin VB.CheckBox chkCopy 
         BackColor       =   &H80000004&
         Caption         =   "Create CD Direcotries in the following location"
         Height          =   315
         Left            =   150
         TabIndex        =   20
         Top             =   780
         Value           =   1  'Checked
         Width           =   3945
      End
      Begin VB.TextBox txtNewCopyDir 
         Height          =   345
         Left            =   420
         TabIndex        =   19
         Top             =   4680
         Width           =   5835
      End
      Begin VB.DriveListBox drvCopy 
         Height          =   315
         Left            =   420
         TabIndex        =   18
         Top             =   1110
         Width           =   5850
      End
      Begin VB.DirListBox dirCopy 
         Height          =   3240
         Left            =   420
         TabIndex        =   17
         Top             =   1410
         Width           =   5835
      End
      Begin VB.CheckBox chkFieldCopy 
         BackColor       =   &H80000004&
         Caption         =   "Create README.TXT file"
         Height          =   345
         Left            =   150
         TabIndex        =   16
         Top             =   480
         Value           =   1  'Checked
         Width           =   4065
      End
      Begin VB.CommandButton cmdEZVIEW 
         Height          =   705
         Index           =   0
         Left            =   -74910
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Batch Process Images"
         Top             =   8490
         Width           =   825
      End
      Begin VB.Frame fraOptions 
         BackColor       =   &H80000004&
         Caption         =   "Options"
         ForeColor       =   &H00C00000&
         Height          =   1185
         Left            =   -74880
         TabIndex        =   11
         Top             =   7110
         Width           =   10875
         Begin VB.CheckBox chkOptionValidate 
            BackColor       =   &H80000004&
            Caption         =   "Validate Image Files"
            Height          =   315
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   3705
         End
         Begin VB.CheckBox chkOptionID 
            BackColor       =   &H80000004&
            Caption         =   "Copy Subject ID to Student ID"
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Top             =   510
            Value           =   1  'Checked
            Width           =   3705
         End
         Begin VB.CheckBox chkOptionReference 
            BackColor       =   &H80000004&
            Caption         =   "Create Cross Reference Text File"
            Height          =   285
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Value           =   1  'Checked
            Width           =   2745
         End
      End
      Begin VB.Frame fraTarget 
         BackColor       =   &H80000004&
         Caption         =   "Target Location"
         ForeColor       =   &H00C00000&
         Height          =   6645
         Left            =   -69420
         TabIndex        =   7
         Top             =   420
         Width           =   5415
         Begin VB.DriveListBox drvTarget 
            Height          =   315
            Left            =   60
            TabIndex        =   10
            Top             =   210
            Width           =   5280
         End
         Begin VB.FileListBox filTarget 
            Height          =   1650
            Left            =   60
            TabIndex        =   9
            Top             =   4920
            Width           =   5295
         End
         Begin VB.DirListBox dirTarget 
            Height          =   4365
            Left            =   60
            TabIndex        =   8
            Top             =   540
            Width           =   5265
         End
      End
      Begin VB.Frame fraSource 
         BackColor       =   &H80000004&
         Caption         =   "EZ-VIEW Source Location"
         ForeColor       =   &H00C00000&
         Height          =   6645
         Left            =   -74880
         TabIndex        =   3
         Top             =   420
         Width           =   5415
         Begin VB.DirListBox dirSource 
            Height          =   4365
            Left            =   60
            TabIndex        =   6
            Top             =   540
            Width           =   5265
         End
         Begin VB.FileListBox filSource 
            Height          =   1650
            Left            =   60
            TabIndex        =   5
            Top             =   4950
            Width           =   5295
         End
         Begin VB.DriveListBox drvSource 
            Height          =   315
            Left            =   60
            TabIndex        =   4
            Top             =   210
            Width           =   5280
         End
      End
      Begin EzView_Builder.ctlData UsrData 
         Height          =   8775
         Left            =   -74940
         TabIndex        =   2
         Top             =   360
         Width           =   11145
         _extentx        =   19659
         _extenty        =   15478
      End
   End
   Begin EzView_Builder.ctlImage UsrImage 
      Height          =   9345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2865
      _extentx        =   5054
      _extenty        =   16484
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   0
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Help"
      Index           =   1
      Begin VB.Menu mnuHelp 
         Caption         =   "&About EZ-VIEW Builder"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------
'
' Project......: RSC EZ-VIEW(r) Builder
'
' Component....: frmMain
'
' Procedure....: (Declarations)
'
' Description..: RSC EZ-VIEW Plug-In
'
' Author.......: Ronald D. Redmer
'
' History......: 07-01-97 RDR Designed and Programmed
'
' (c) 1997-1999 Redmer Software Company, Inc.
' All Rights Reserved
'----------------------------------------------------------------------------
Option Explicit                                             'Require explicit variable declaration
Private Const conProcessButton = 0                          'Index of processing images button
Private Const conRSCIcon = 101                              'Resource ID of RSC Icon
Private Const conStopButtonIcon = 110                       'Resource ID of Stop Icon
Private Const conProcessButtonIcon = 105                    'Resource ID of Process Icon
Private Const conExitButtonIcon = 112                       'Resource ID of Exit Button
Private Const conReportFolder = "\REPORTS\"                 'EZ-VIEW standard report folder
Private Const conDataFolder = "\DATA\"                      'EZ-VIEW standard report folder
Private Const conSetupFolder = "\SETUP\"                    'EZ-VIEW standard setup folder
Private Const conReferenceFolder = "\DATAMAC"               'EZ-VIEW standard reference folder for SASI
Private Const conSASIfile = "XREFPICT.TXT"                  'EZ-VIEW standard SASI file
Private Const conFoxProDSN = "DSN=Visual FoxPro Tables;UID=;PWD=;SourceType=DBF;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=No;Deleted=Yes;SourceDB="
Private bCancel As Boolean                                  'Cancel flag
Private bProcessing As Boolean                              'Set active when image processing loop running
Private sFileName As String                                 'Name of the currently selected image file
Private Sub Form_Load()                                     'Load the form and initialize controls
    If Not EZ_DEBUG Then On Error Resume Next
    Me.Icon = LoadResPicture(conRSCIcon, vbResIcon)         'Load the window icon
    tabMain.Tab = 0
    cmdEZVIEW(conProcessButton).Picture = LoadResPicture(conProcessButtonIcon, vbResBitmap)
    cmdTools(conProcessButton).Picture = LoadResPicture(conProcessButtonIcon, vbResBitmap)
    bCancel = False                                         'Initialize batch process cancel flag to false
End Sub

Private Sub mnuFile_Click(Index As Integer)
    If Not EZ_DEBUG Then On Error Resume Next
    If MsgBox("Are you sure?", vbApplicationModal + vbYesNo + vbQuestion, "Exit") = vbYes Then
        End
    End If
End Sub

Private Sub mnuHelp_Click(Index As Integer)
    If Not EZ_DEBUG Then On Error Resume Next
    frmSplash.Show vbModal
End Sub

Private Sub cmdEZVIEW_Click(Index As Integer)               'Process command button clicks
    If Not EZ_DEBUG Then On Error Resume Next
    Select Case Index                                       'Select on the index of the button clicked
        Case conProcessButton                               'Bacth process images button
            If Not bProcessing Then                         'Toggle Icons
                If MsgBox("Create EZ-VIEW Distribution?", vbApplicationModal + vbQuestion + vbYesNo, EZ_CAPTION) = vbYes Then
                    cmdEZVIEW(conProcessButton).Picture = LoadResPicture(conStopButtonIcon, vbResIcon)
                    CopySetup                               'Build the EZ-VIEW CD-ROM
                    CopyData                                'Copy data from source table to students table
                    cmdEZVIEW(conProcessButton).Picture = LoadResPicture(conProcessButtonIcon, vbResBitmap)
                End If
            Else
                bCancel = True
            End If
    End Select
End Sub
Private Sub CopySetup()                                     'Copy setup folders from source to target
    If Not EZ_DEBUG Then On Error Resume Next
    Dim fs As FileSystemObject                              'FileSystem object provided by "Microsoft Scripting Runtime"
    sbMain.Panels(0).Text = "Copying files..."                   'Update status
    Set fs = New FileSystemObject                           'Instantiate the object
    DoEvents
    fs.CopyFolder dirSource.Path, dirTarget.Path, True      'Copy folders (with overwrite set)
    DoEvents
    Set fs = Nothing                                        'Clear the filesystem object
End Sub
Private Sub CopyData()
    If Not EZ_DEBUG Then On Error Resume Next
    Dim cnnEZ As ADODB.Connection                           'Connection to EZ-VIEW Database
    Dim rsStudents As ADODB.Recordset                       'Students recordset
    Dim fRef As FileSystemObject                            'Cross reference file system handle
    Dim fl As TextStream
    sbMain.Panels(0).Text = "Appending Data..."                   'Update status
    Set cnnEZ = New ADODB.Connection                        'Instantiate the connection object
    Set rsStudents = New ADODB.Recordset                    'Instantiate the recordset object
    Set fRef = New FileSystemObject
    cnnEZ.Open conFoxProDSN & dirTarget.Path & conDataFolder & ";"  'Open the target database using DSN-less ODBC connection
    rsStudents.Open "SELECT * FROM STUDENTS", cnnEZ, adOpenDynamic, adLockOptimistic
    UsrData.rsRecords.MoveFirst                               'Move to first record in records table
    If chkOptionReference.Value = 1 Then                    'If build reference file is checked
        fRef.CreateFolder dirTarget.Path & conReferenceFolder
        Set fl = fRef.OpenTextFile(dirTarget.Path & conReferenceFolder & "\" & conSASIfile, ForAppending, True, TristateFalse)
    End If
    Do While Not UsrData.rsRecords.EOF                        'Loop for each record in the recordset
        If chkOptionID.Value = 1 Then                       'If copy subject to student id is checked
            If Len(Trim$(UsrData.rsRecords("STUDENTID").Value)) = 0 Then      'Copy the subject field to the student id
                UsrData.rsRecords("STUDENTID").Value = Trim$(UsrData.rsRecords("SUBJECT").Value)
                UsrData.rsRecords.Update
            End If
        End If
        rsStudents.AddNew                                   'Add a new record to the EZ-VIEW Students table
        rsStudents("STUDENTID").Value = Trim$(UsrData.rsRecords("STUDENTID").Value)
        rsStudents("FIRST_NAME").Value = Trim$(UsrData.rsRecords("FIRST_NAME").Value)
        rsStudents("LAST_NAME").Value = Trim$(UsrData.rsRecords("LAST_NAME").Value)
        rsStudents("GRADE").Value = Trim$(UsrData.rsRecords("GRADE").Value)
        rsStudents("TEACHER").Value = Trim$(UsrData.rsRecords("TEACHER").Value)
        rsStudents("HOMEROOM").Value = Trim$(UsrData.rsRecords("HOMEROOM").Value)
        rsStudents("BOX").Value = Trim$(UsrData.rsRecords("BOX").Value)
        rsStudents("ADDRESS1").Value = Trim$(UsrData.rsRecords("ADDRESS1").Value)
        rsStudents("ADDRESS2").Value = Trim$(UsrData.rsRecords("ADDRESS2").Value)
        rsStudents("CITY").Value = Trim$(UsrData.rsRecords("CITY").Value)
        rsStudents("ZIP_CODE").Value = Trim$(UsrData.rsRecords("ZIP_CODE").Value)
        rsStudents("PHONE1").Value = Trim$(UsrData.rsRecords("HOME_PHONE").Value)
        rsStudents("GENDER").Value = Trim$(UsrData.rsRecords("GENDER").Value)
        rsStudents.Update                                   'Update the EZ-VIEW table with new values
        If chkOptionReference.Value = 1 Then                'If build reference file is checked
            fl.WriteLine Chr$(34) & Format$(rsStudents("STUDENTID").Value, "0000000000") & Chr$(34) & "," & Chr$(34) & Trim$(rsStudents("STUDENTID").Value) & ".PCT" & Chr$(34)
        End If
        UsrData.rsRecords.MoveNext                            'Move to the next record in the source recordset
        DoEvents
    Loop
    If chkOptionReference.Value = 1 Then                    'If build reference file is checked
        fl.Close                                            'Close the reference file
        Set fl = Nothing                                    'Release the reference file memory
        Set fRef = Nothing                                  'Release the filesystem memory
    End If
    rsStudents.Close                                        'Close the students table
    cnnEZ.Close                                             'Close the connection
    Set rsStudents = Nothing                                'Release students object
    Set cnnEZ = Nothing                                     'Release connection object
End Sub
Private Sub drvSource_change()
    If Not EZ_DEBUG Then On Error Resume Next
    dirSource.Path = drvSource.Drive
End Sub
Private Sub dirSource_Change()
    If Not EZ_DEBUG Then On Error Resume Next
    filSource.Path = dirSource.Path
End Sub
Private Sub drvTarget_Change()
    If Not EZ_DEBUG Then On Error Resume Next
    dirTarget.Path = drvTarget.Drive
ErrorHandler:
    Resume Next
End Sub
Private Sub dirTarget_Change()
    If Not EZ_DEBUG Then On Error Resume Next
    filTarget.Path = dirTarget.Path
ErrorHandler:
    Resume Next
End Sub

Private Sub cmdTools_Click(Index As Integer)
    If Not EZ_DEBUG Then On Error Resume Next
    OutputImages
End Sub

Private Sub drvCopy_Change()
    If Not EZ_DEBUG Then On Error Resume Next
    dirCopy.Path = drvCopy.Drive
End Sub

Private Sub txtNewCopyDir_LostFocus()
    If Not EZ_DEBUG Then On Error Resume Next
    Dim sFolder As String
    If Trim$(txtNewCopyDir.Text) <> "" Then
        sFolder = dirCopy.Path & "\" & Trim$(txtNewCopyDir.Text)
        MkDir sFolder
        txtNewCopyDir.Text = ""
        dirCopy.Refresh
    End If
    Exit Sub
ErrorHandler:
    MsgBox "Could not create folder.", vbOKOnly + vbApplicationModal + vbInformation, EZ_CAPTION
    Resume Next
End Sub

Private Sub OutputImages()                                   'Create PSPA CD
    If Not EZ_DEBUG Then On Error GoTo ErrorHandler
    Dim sSrcName As String
    Dim sTargetDir As String
    Dim sDstName As String
    Dim sExt As String
    Dim sTxt As String
    Dim fRef As FileSystemObject                            'Cross reference file system handle
    Dim fl As TextStream
    Dim fIndex As TextStream
    Dim iFolder As Integer
    Dim iNumFolders As Integer
    Dim iImage As Integer
    
    Set fRef = New FileSystemObject
    Set fl = fRef.OpenTextFile(dirCopy.Path & "\LOG.TXT", ForWriting, True, TristateFalse)
    Set fIndex = fRef.OpenTextFile(dirCopy.Path & "\INDEX.TXT", ForWriting, True, TristateFalse)
    
    sExt = UsrImage.ImageExtension                          'Get the extension on the currently selected image file
        
    If MsgBox("Produce CD folders from [" & Trim$(Str$(UsrData.rsRecords.RecordCount)) & "] images?", vbYesNo + vbApplicationModal + vbQuestion, EZ_CAPTION) = vbNo Then
        Exit Sub                                            'Exit this routine
    End If
        
    '---- Create folders for copy
    sbMain.Panels(1).Text = "Creating folders..."                   'Update status
    DoEvents
    iNumFolders = (UsrData.rsRecords.RecordCount / 200)       '200 images max per folder
    If iNumFolders < 1 Then
        iNumFolders = 1
    End If
    For iFolder = 1 To iNumFolders                          'Loop for each folder to be created
        sTargetDir = dirCopy.Path & "\FOLDER" & Trim$(Str(iFolder))
        
       ' MsgBox sTargetDir
        
        If Not fRef.FolderExists(sTargetDir) Then
            fRef.CreateFolder sTargetDir
        End If
    Next
        
    '---- Copy images to folders
    sbMain.Panels(1).Text = "Copying Images..."                   'Update status
    DoEvents
    iFolder = 1
    iImage = 1
    UsrData.rsRecords.MoveFirst                           'Move to first record in data set
    Do While Not UsrData.rsRecords.EOF                    'Loop for each record in data set
         
        If Len(Trim(UsrData.rsRecords(Trim(UsrData.ImageTag)).Value)) > 0 Then
            sSrcName = Trim(UsrImage.ImagePath) & Trim(UsrData.rsRecords(Trim(UsrData.ImageTag)).Value) & sExt
            sDstName = Trim(dirCopy.Path) & "\FOLDER" & Trim$(Str(iFolder)) & "\" & Trim(UsrData.rsRecords(Trim(UsrData.ImageTag)).Value) & sExt
            If fRef.FileExists(sSrcName) Then
                fRef.CopyFile sSrcName, sDstName, True
                If Not fRef.FileExists(sDstName) Then
                    fl.WriteLine "*** Error copying [" & sSrcName & "] to [" & sDstName & "] filesystem error."
                Else
                    fl.WriteLine "*** Copied [" & sSrcName & "] to [" & sDstName & "] OK."
                    
                    sTxt = "VOL1" & vbTab & "FOLDER" & Trim(Str(iFolder)) & vbTab
                    sTxt = sTxt & Trim(UsrData.rsRecords(Trim(UsrData.ImageTag)).Value) & sExt & vbTab
                    sTxt = sTxt & Trim(UsrData.rsRecords("GRADE").Value & "") & vbTab
                    sTxt = sTxt & Trim(UsrData.rsRecords("LAST_NAME").Value & "") & vbTab
                    sTxt = sTxt & Trim(UsrData.rsRecords("FIRST_NAME").Value & "") & vbTab
                    sTxt = sTxt & Trim(UsrData.rsRecords("HOMEROOM").Value & "") & vbTab & vbTab
                    sTxt = sTxt & Trim(UsrData.rsRecords("TEACHER").Value & "") & vbTab
                    
                    fIndex.WriteLine sTxt
                    
                    iImage = iImage + 1
                    If iImage > 200 Then
                        iFolder = iFolder + 1
                        iImage = 1
                    End If
                    
                End If
            Else
                fl.WriteLine "*** Error copying [" & sSrcName & "] to [" & sDstName & "] Source not found."
            End If
        End If
        UsrData.rsRecords.MoveNext
    
    Loop
    
    fl.Close
    fIndex.Close
    UsrData.rsRecords.MoveFirst                               'Move to first record in data set
    MsgBox "CD image creation complete.", vbOKOnly + vbApplicationModal + vbInformation, EZ_CAPTION
    Exit Sub
ErrorHandler:
    MsgBox "Error in CD creation: #[" & Str(Err.Number) & "][" & Err.Description & "]"
    Resume Next
End Sub

