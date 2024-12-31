*-----------------------------------------------------------------------------
*
*  System.....: RSC EZ-VIEW(R) School Image Management Software
*
*  Module.....: EZ-VIEW.PRG 
*
*  Description: Main application module.
*
*  History....:
*
*   v1.00r001 10-20598 RDR Ported to Microsoft Windows NT(r)/Visual FoxPro 5.0SR3.
*
*  (c) 1992-2000 Redmer Software Company.  All rights reserved.
*-----------------------------------------------------------------------------

	*--- process command line parameters
	PARAMETERS gcCommandLine
	IF TYPE("gcCommandLine") <> "C"
		gcCommandLine = ""
	ENDIF

	*--- Setup the FoxPro environment
	CLEAR
	SET TALK OFF
	SET NOTIFY OFF
	SET MESSAGE TO 0
	SET STATUS BAR OFF
	SET DELETED ON
	SET CONFIRM OFF
	SET SAFETY OFF
	SET ESCAPE OFF
	SET EXCLUSIVE OFF
	SET EXACT ON
	SET LIBRARY TO
	SET STATUS OFF
	SET RESOURCE OFF
	CLOSE ALL
	ON KEY
	SET REPROCESS TO -1
	SET SYSMENU OFF

	*--- Define application constants
	#define APP_NAME		"RSC EZ-VIEW/2000"				&& Application name used on forms
	#define APP_ID			"EZVIEW"						&& Application id used in registry
	#define STATION_ID		"EZVIEW01"						&& Default station id
	#define APP_VERSION		"Release 003"					&& Application version.

	*--- Declare PUBLIC (global) DEBUG variable - this should be the only global non-object variable in the application
	PUBLIC glDebug												&& Global debug flag (for developer testing)
	STORE .F. TO glDebug										&& Initialize debug flag to false
	IF "DEBUG" $ gcCommandLine
		STORE .T. TO glDebug
	ENDIF
	
	*--- Set path to application directory
	LOCAL lcSys16, lcProgram
	lcSys16 = SYS(16)
	lcProgram = SUBSTR(lcSys16, AT(":", lcSys16) - 1)
	CD LEFT(lcProgram, RAT("\", lcProgram))
	SET PATH TO ;DATA;LOCALFILES;INCLUDE;FORMS;SCREENS;GRAPHICS;HELP;LIBS;MENUS;PROGS;REPORTS

	*--- Maximize the Visual FoxPro window to full screen & set caption
	_SCREEN.WindowState = 2
	_SCREEN.Caption = APP_NAME + " " + APP_VERSION + IIF(glDebug, "  *** DEBUG MODE *** ", "" )
	_SCREEN.Closable = .F.
	_SCREEN.Icon = "RSC_CompanyLogo.ico"

	=GL_LOG_INITIALIZE()
	=GL_LOG_PROC( PROGRAM() )

	*-- DECLARE DLL statements for reading/writing to private INI files
	=GL_LOG_ACTION( "Declaring WIN32 API Functions" )
	DECLARE INTEGER GetPrivateProfileString IN Win32API  AS GetPrivStr ;
	  String cSection, String cKey, String cDefault, String @cBuffer, ;
	  Integer nBufferSize, String cINIFile
	DECLARE INTEGER WritePrivateProfileString IN Win32API AS WritePrivStr ;
	  String cSection, String cKey, String cValue, String cINIFile

	*--- Instantiate global objects
	RELEASE goSQL, goRegistry, goMain, goSplash, goError, goPreferences
	PUBLIC  goSQL, goRegistry, goMain, goSplash, goError, goPreferences

	SET CLASSLIB TO EZ_MAIN ADDITIVE

	*--- Instantiate and display the splash screen object
	=GL_LOG_ACTION( "Initializing class libraries: vc_splash_screen" )
	goSplash = CREATEOBJECT("vc_Splash_Screen")					&& Create the splash screen object
	goSplash.SetPreferences("RSC EZ-VIEW(R)", APP_NAME, APP_VERSION)	&& Set screen labels to application-specific values
	goSplash.Show(2)											&& Display as non-modal form - show(2)

	DoEvents													&& Allow Windows to process events

	*--- Instantiate the registry object, used to store and retrieve application settings.
	=GL_LOG_ACTION( "Initializing class libraries: vc_registry" )
	goRegistry = CREATEOBJECT("vc_Registry")					&& Create the registry object

	*--- Instantiate the database object
	=GL_LOG_ACTION( "Initializing class libraries: vc_sql_foxpro" )
	goSQL = CREATEOBJECT("vc_SQL_FoxPro")						&& Create the global foxpro EZLAB database driver
	goSQL.db_OpenTables											&& Open the FoxPro tables
	
	*--- Hide the splash screen
	goSplash.Hide

	*--- Instantiate the application security interface
	=GL_LOG_ACTION( "Initializing class libraries: vc_Application_Security" )
	goSecurity = CREATEOBJECT("vc_Application_Security")
	goSecurity.SetPreferences("RSC EZ-VIEW(R)", APP_NAME, APP_VERSION)
	goSecurity.Show(1)											&& Perform login (if security configured).

	*--- Instantiate the menu object
	=GL_LOG_ACTION( "Initializing main menu" )
	DO ez_view.mpr

	*--- Instantiate the main application form
	=GL_LOG_ACTION( "Initializing class libraries: vc_Student_Maintenance" )
	goMain = CREATEOBJECT( "vc_Student_Maintenance" )
	goMain.Show
	READ EVENTS													&& Process Application Events, exit code in menu File\Exit

	=GL_LOG_ACTION( "******* PROCESSED PAST READ EVENTS!  EXITING EVENT LOOP. *******" )

RETURN	


FUNCTION GL_LOG_INITIALIZE
	LOCAL tmpname
	PUBLIC gnLogFile
	
	IF (glDebug)
		*--- CREATE LOG FILE AND WRITE HEADER INFORMATION
		gnLogFile = FCREATE( "LOGFILE.TXT", 0 )
		=FPUTS( gnLogFile, _SCREEN.Caption )
		=FPUTS( gnLogFile, REPLICATE("-", 80 ))
		=FPUTS( gnLogFile, "Log Initialized:  " + DTOC(DATE()) + " At " + TIME() )
		=FPUTS( gnLogFile, VERSION(1) )
		=FPUTS( gnLogFile, IIF(VERSION(2)=0,"Runtime",IIF(VERSION(2)=1,"Standard","Professional")) + " version.  Localization=" + VERSION(3))
		=FPUTS( gnLogFile, REPLICATE("-", 80 ))
		tmpname = SYS(0)
		=FPUTS( gnLogFile, "NETWORK MACHINE INFO: " + IIF(TYPE("tmpname")=="C",tmpname,"") )
		RETURN .T.
	ELSE
		*--- REMOVE THE LOG FILE TO AVOID CONFUSION WITH PRIOR BUILD
		IF FILE("LOGFILE.TXT")
			DELETE FILE LOGFILE.TXT
		ENDIF
	ENDIF
	
RETURN .F.

FUNCTION GL_LOG_PROC
	LPARAMETERS lpcName
	IF (glDebug)
		*--- LOG PROCEDURE NAME TO LOG FILE
		=FPUTS( gnLogFile, REPLICATE("-", 80) )
		=FPUTS( gnLogFile, "PROC=" + lpcName )
	ENDIF
RETURN .T.

FUNCTION GL_LOG_ACTION
	LPARAMETERS lpcAction
	IF (glDebug)
		*--- LOG ACTION TO LOG FILE WITH TIME, SECONDS (MILLISECONDS), FREE MEM, & USER MEM
		=FPUTS( gnLogFile, "TM=" + TIME() + "|SE=" + STR(SECONDS(),6,0) + "|FM=" + SYS(1001) + "|UM=" + SYS(1016) + "|" + lpcAction )
	ENDIF
RETURN .T.
