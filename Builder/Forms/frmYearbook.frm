VERSION 5.00
Begin VB.Form frmYearbook 
   Caption         =   "Build PSPA Yearbook"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraYearbook 
      BackColor       =   &H80000004&
      Caption         =   "PSPA Standard Yearbook CD"
      ForeColor       =   &H00C00000&
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5955
   End
End
Attribute VB_Name = "frmYearbook"
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
Private Const conExitButton = 0                             'Index of exit button
Private Const conProcessButton = 1                          'Index of processing images button
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
    On Error GoTo ErrorHandler                              'Set error handler
    Me.Icon = LoadResPicture(conRSCIcon, vbResIcon)         'Load the window icon
    cmdTools(conExitButton).Picture = LoadResPicture(conExitButtonIcon, vbResIcon)
    bCancel = False                                         'Initialize batch process cancel flag to false
    Exit Sub                                                'Exit this routine
ErrorHandler:                                               'Error handling code
    Resume Next                                             'Simply resume next line of code
End Sub


