'----- License Stuff
'This file is part of SPFLite.

'    SPFLite is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    SPFLite is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with SPFLite.  If not, see <https://www.gnu.org/licenses/>.

'----- Compiler stuff
#COMPILE EXE "SPFLite10"
#DIM ALL
#STACK 2000000
#DEBUG DISPLAY OFF
#DEBUG ERROR OFF
#TOOLS OFF
#OPTIMIZE CODE ON
%USEMACROS = 1

'----- Bring in some of the bits and pieces
#INCLUDE ONCE "Win32Api.inc"                                      ' Windows standard stuff
#INCLUDE ONCE "WinIOCtl.inc"                                      ' Windows IOCtl stuff
#INCLUDE ONCE "ComDlg32.inc"                                      ' Common Dialog stuff
#INCLUDE ONCE "commCtrl.inc"                                      ' Common Controls stuff
#INCLUDE ONCE "ShlWAPI.inc"                                       ' Utility API's
#INCLUDE ONCE "WinINET.inc"                                       ' Internet stuff
#INCLUDE ONCE "RICHEDIT.INC"                                      ' RichEdit stuff
#INCLUDE ONCE "HtmlHelp.inc"                                      ' HTML Help stuff
#INCLUDE ONCE "_Types.inc"                                        ' Application TYPE definitions & Constant Equates
#INCLUDE ONCE "_AsmData.inc"                                      ' ASM Data tables
#INCLUDE ONCE "_Resource.inc"                                     ' Icons, KBMaster etc. tables
#INCLUDE ONCE "_Version.inc"                                      ' Version, Build number
#INCLUDE ONCE "_DialogEquates.inc"                                ' Dialogue equates (Control ID's etc.)
#INCLUDE ONCE "_HnD.inc"                                          ' Help Topic Equates
#INCLUDE ONCE "_PCRE.inc"                                         ' PCRE Regex stuff
#INCLUDE ONCE "_thinCore.INC"                                     ' thinBasic interface definitions
#RESOURCE RES "WinVerManifest.res"                                ' Needed for valid result with Windows 8.1 +

'----------
' Global Data
'----------

'----- Define the Dialog Tabs control data

GLOBAL Tabs()              AS iObjTabData                         ' Table of active tabs
GLOBAL TabsNum             AS LONG                                ' No. of entries in Tabs()
GLOBAL TabUnique           AS LONG                                ' Unique ID for the session
GLOBAL TP                  AS iObjTabData                         ' Pointer to current tab data

'----- Dialog handles

GLOBAL hInstance           AS DOUBLE                              ' Handle for my own instance
GLOBAL hWnd                AS LONG                                ' Main    Window   handle
GLOBAL hTab                AS LONG                                ' Tab     Control  handle
GLOBAL hStatusBar          AS LONG                                ' Status  Bar      handle
GLOBAL hPrf                AS LONG                                ' PROFILE Display  handle
GLOBAL hPage1              AS LONG                                ' Profile Page 1   handle
GLOBAL hPage2              AS LONG                                ' Profile Page 2   handle
GLOBAL hWel                AS LONG                                ' Welcome Display  handle
GLOBAL hKey                AS LONG                                ' KEYMAP  Display  handle
GLOBAL hOpt                AS LONG                                ' OPTIONS Display  handle
GLOBAL hGeneral            AS LONG                                ' OPTIONS General  handle
GLOBAL hSubmit             AS LONG                                ' OPTIONS Submit   handle
GLOBAL hFManager           AS LONG                                ' OPTIONS FManager handle
GLOBAL hScreen             AS LONG                                ' OPTIONS Screen   handle
GLOBAL hKeyboard           AS LONG                                ' OPTIONS Keyboard handle
GLOBAL hSBar               AS LONG                                ' OPTIONS SBar     handle
GLOBAL hScheme             AS LONG                                ' OPTIONS Scheme   handle
GLOBAL hHiLites            AS LONG                                ' OPTIONS HiLites  handle
GLOBAL hDef                AS LONG                                ' Default Display  handle
GLOBAL hPrt                AS LONG                                ' PRINT   Display  handle
GLOBAL hANSI               AS LONG                                ' ANSI    Display  handle
GLOBAL hIntr               AS LONG                                ' FF Interupt      handle
GLOBAL hSplash             AS LONG                                ' Splash screen    handle
GLOBAL hMsg                AS LONG                                ' Help Level2      handle
GLOBAL hMsgHnd             AS LONG                                ' Help Level2 Send handle
GLOBAL hUpdate             AS LONG                                ' Update  Display  handle
GLOBAL hRich               AS LONG                                ' Update  URL      handle
GLOBAL hToolTips           AS LONG                                ' Tooltip creation handle
GLOBAL hKbrdHook           AS LONG                                ' Keyboard hook
GLOBAL hBoldFont           AS LONG                                ' Font handles
GLOBAL hFixedFont          AS LONG                                '
GLOBAL hScrFont            AS LONG                                '
GLOBAL hScrFontUnd         AS LONG                                '
GLOBAL hSBFont             AS LONG                                '
GLOBAL hSBFontB            AS LONG                                '
GLOBAL gSBWidth            AS LONG                                '
GLOBAL gSBHeight           AS LONG                                '
GLOBAL gTabRC              AS RECT                                '
GLOBAL gTabHdrRC           AS RECT                                '
GLOBAL gResizeActive       AS LONG                                '

'----- Global Colors (Set from ENV.Scheme variables)
GLOBAL CCustFG             AS LONG                                ' Keep all these color variable in order
GLOBAL CCustBG1            AS LONG                                ' as they're also accessed as a table
GLOBAL CCustBG2            AS LONG                                '
GLOBAL cTxtLoFG            AS LONG                                '
GLOBAL cTxtLoBG1           AS LONG                                '
GLOBAL cTxtLoBG2           AS LONG                                '
GLOBAL cTxtHiFG            AS LONG                                '
GLOBAL cTxtHiBG1           AS LONG                                '
GLOBAL cTxtHiBG2           AS LONG                                '
GLOBAL cLNoHiFG            AS LONG                                '
GLOBAL cLNoHiBG1           AS LONG                                '
GLOBAL cLNoHiBG2           AS LONG                                '
GLOBAL cLNoLoFG            AS LONG                                '
GLOBAL cLNoLoBG1           AS LONG                                '
GLOBAL cLNoLoBG2           AS LONG                                '
GLOBAL cATabModFG          AS LONG                                '
GLOBAL cATabModBG1         AS LONG                                '
GLOBAL cATabModBG2         AS LONG                                '
GLOBAL cATabNModFG         AS LONG                                '
GLOBAL cATabNModBG1        AS LONG                                '
GLOBAL cATabNModBG2        AS LONG                                '
GLOBAL cITabModFG          AS LONG                                '
GLOBAL cITabModBG1         AS LONG                                '
GLOBAL cITabModBG2         AS LONG                                '
GLOBAL cITabNModFG         AS LONG                                '
GLOBAL cITabNModBG1        AS LONG                                '
GLOBAL cITabNModBG2        AS LONG                                '
GLOBAL cPFKFG              AS LONG                                '
GLOBAL cPFKBG1             AS LONG                                '
GLOBAL cPFKBG2             AS LONG                                '
GLOBAL cStatFG             AS LONG                                '
GLOBAL cStatBG1            AS LONG                                '
GLOBAL cStatBG2            AS LONG                                '
GLOBAL cFMToolFG           AS LONG                                '
GLOBAL cFMToolBG1          AS LONG                                '
GLOBAL cFMToolBG2          AS LONG                                '
GLOBAL cErrorFG            AS LONG                                '
GLOBAL cErrorBG1           AS LONG                                '
GLOBAL cErrorBG2           AS LONG                                '
GLOBAL cSelectedFG         AS LONG                                '
GLOBAL cSelectedBG1        AS LONG                                '
GLOBAL cSelectedBG2        AS LONG                                '
GLOBAL cRsvd12FG           AS LONG                                '
GLOBAL cRsvd12BG1          AS LONG                                '
GLOBAL cRsvd12BG2          AS LONG                                '
GLOBAL cRsvd13FG           AS LONG                                '
GLOBAL cRsvd13BG1          AS LONG                                '
GLOBAL cRsvd13BG2          AS LONG                                '
GLOBAL cRsvd14FG           AS LONG                                '
GLOBAL cRsvd14BG1          AS LONG                                '
GLOBAL cRsvd14BG2          AS LONG                                '
GLOBAL cRsvd15FG           AS LONG                                '
GLOBAL cRsvd15BG1          AS LONG                                '
GLOBAL cRsvd15BG2          AS LONG                                '
GLOBAL cRsvd16FG           AS LONG                                '
GLOBAL cRsvd16BG1          AS LONG                                '
GLOBAL cRsvd16BG2          AS LONG                                '
GLOBAL cBandBG             AS LONG                                ' Active BG for Banding needs

'----- Global copy of HiLite names
GLOBAL nHiLites()          AS STRING                              ' Table of valid color names
GLOBAL nHiLitesChrs        AS STRING                              ' String of valid 1st characters

'----- Global Variables for Printing Routines

GLOBAL gPrinterOpen        AS INTEGER                             '
GLOBAL gPFontHndl          AS LONG                                '
GLOBAL gPCharWidth         AS LONG                                '
GLOBAL gPCharHeight        AS LONG                                '
GLOBAL gPPageWidth         AS LONG                                '
GLOBAL gPPageHeight        AS LONG                                '
GLOBAL gPCpl               AS LONG                                '
GLOBAL gPLpp               AS LONG                                '
GLOBAL gPLFill             AS LONG                                '
GLOBAL gPRFill             AS LONG                                '
GLOBAL gPTFill             AS LONG                                '
GLOBAL gPColor             AS LONG

'----- Global Variables during execution
                                                                  '
GLOBAL gKbdRecFlag         AS LONG                                ' KB recording is active
GLOBAL gKbdRecTxtFlag      AS LONG                                ' KB doing a text string
GLOBAL gKbdRecording       AS STRING                              ' The recording data
GLOBAL gDateActive         AS STRING                              ' The Date/Time session starte
GLOBAL gDateActive1        AS STRING                              ' The Date/Time session started - 1 hr
GLOBAL gDateActive8        AS STRING                              ' The Date/Time session started - 8 hr
GLOBAL gDateActive24       AS STRING                              ' The Date/Time session started - 24 hr
GLOBAL gDateActive48       AS STRING                              ' The Date/Time session started - 48 hr
GLOBAL gPTbl()             AS PTypeTable                          ' PowerType table
GLOBAL gPTblCount          AS LONG                                ' Number in PT
GLOBAL gFQ()               AS WatchData                           ' List of Open Filenames
GLOBAL gfDoingMsg          AS LONG                                ' Doing a MSGBOX, OPEN Dialog, etc. (active counter)
GLOBAL gfDialogDone        AS LONG                                ' Result of Dialog display
GLOBAL gfTermFlag          AS LONG                                ' Set for full termination
GLOBAL gfXRebuild          AS LONG                                ' Exclude lines rebuild needed
GLOBAL gfEndAll            AS LONG                                ' ENDALL in progress
GLOBAL gfInterrupt         AS LONG                                ' Interrupt flag
GLOBAL gCaretCtr           AS LONG                                ' Count of Caret Show/Hide
GLOBAL gShutFlag           AS LONG                                ' Stacked =X or EXIT command
GLOBAL gLastKBTime         AS DWORD                               ' Time of last KB operation.
GLOBAL gGlblMessage        AS STRING                              ' Message from a tab to next visible tab
GLOBAL gfLeftDown          AS LONG                                ' To track Mouse button
GLOBAL gfMiddleDown        AS LONG                                ' To track Middle mouse button
GLOBAL gfActive            AS LONG                                ' Are we active or not
GLOBAL gfOptCancel         AS LONG                                ' OPTION was cancelled
GLOBAL gCmdRtrev()         AS STRING                              ' Hold area for Retrieve commands
GLOBAL gCmdRtrevIX         AS LONG                                ' Pointer to next retrieve item
GLOBAL gCmdRtrevMsg        AS STRING                              ' Retrieve SB message
GLOBAL gCmdRtrevLast       AS STRING                              ' Last retrieve operation
GLOBAL gCmdList()          AS STRING                              ' Command list for DO processing
GLOBAL gWinList()          AS WININFOTYPE                         '
GLOBAL gWinListPos         AS LONG                                '
GLOBAL gDataLen            AS LONG                                ' Length of Txt data portion of screen
GLOBAL gwScrHeight         AS LONG                                ' Working screen height
GLOBAL gEOLFlagList()      AS STRING                              '
GLOBAL gRECFMList()        AS STRING                              '
GLOBAL gPrimTable()        AS STRING                              '
GLOBAL gTabDelList()       AS LONG                                ' Tabs to be deleted this interrupt
GLOBAL gTabDelCtr          AS LONG                                ' Count of entries
GLOBAL gTabDelMsg          AS STRING                              ' Message for next tab
GLOBAL gTabSwitch          AS LONG                                ' Tab to be switched to
GLOBAL gTabSwitchMsg       AS STRING                              ' Message to be issued in the tab
GLOBAL gTabSwitchCmd       AS STRING                              ' Command to be issued in the tab
GLOBAL gTabStack()         AS LONG                                ' Stack of Tab activity
GLOBAL gTabStackNum        AS LONG                                ' Number in stack
GLOBAL gFontHeight         AS INTEGER                             ' Current font height
GLOBAL gFontWidth          AS INTEGER                             ' Current font width
GLOBAL gFontScale          AS SINGLE                              ' Scale factor for dialog fonts
GLOBAL gLNPadCol           AS LONG                                ' Location of Pad column
GLOBAL gLNData1            AS LONG                                ' Location of Data column 1
GLOBAL gKeyChr             AS STRING                              ' Current KB key entered
GLOBAL gKeyPrimOper        AS STRING                              ' Current Primitive Operand
GLOBAL gSBTable()          AS SBarEntry                           ' Global version of SB table
GLOBAL gSubmitFile         AS STRING                              ' Submit filename
GLOBAL gResultFile         AS STRING                              ' Result (~R) filename
GLOBAL gSubmitType         AS STRING                              ' "J" or "S"
GLOBAL gJobID              AS STRING                              ' Current job number as JOB12345
GLOBAL gDosOperands        AS STRING                              ' User supplied DOS operands
GLOBAL gSetKey()           AS STRING                              ' SET Key table
GLOBAL gSetData()          AS STRING                              ' SET Data table
GLOBAL gSetCount           AS LONG                                ' No. of items in SET table
GLOBAL gSetClipB           AS STRING                              ' Saved clipboard
GLOBAL gPageNumber         AS STRING                              ' Print Page number for ~#
GLOBAL gPrtRaw             AS LONG                                ' Print Raw mode
GLOBAL gPrtPaper           AS STRING                              ' Current paper in printer
GLOBAL gEnumWith           AS LONG                                ' ENUM incr value
GLOBAL gLnoTextType()      AS LONG                                ' Table of Line Number display types
GLOBAL gLnoTextTxt()       AS STRING                              ' Table of Line Number display text
GLOBAL gUpper              AS STRING                              ' Upper list with Int. chars if needed
GLOBAL gLower              AS STRING                              ' Lower list with Int. chars if needed
GLOBAL gHelpKey()          AS STRING                              ' Help Index key
GLOBAL gHelpMapid          AS LONG                                ' Help MapID
GLOBAL gHelpCount          AS LONG                                ' Help Count
GLOBAL gInternalCB         AS STRING                              ' Secret internal Clipboard
GLOBAL gDefaultAnswer      AS STRING                              ' Answer from the DispDefault popup
GLOBAL gLoopFlag           AS LONG                                ' Count transaction loop
GLOBAL gLoopCtr            AS LONG                                ' # seconds doing this transaction
GLOBAL gNoVersion          AS LONG                                ' Suppress version from Windows title bar
GLOBAL gCrashList()        AS STRING                              ' Module trace
GLOBAL gCrashCtr           AS LONG                                ' Module trace Index
GLOBAL gCrashLastPCmd      AS STRING                              ' Last Primary command
GLOBAL gCrashLastLCmd      AS STRING                              ' Last Line command
GLOBAL gCrashLastPrim      AS STRING                              ' Last Primitive

'----- Line Control handling

GLOBAL LTblAIX             AS LONG                                ' Count of LTblA entries
GLOBAL LTblA()             AS LCtlScan                            ' Initial line scan table
GLOBAL LTblBIX             AS LONG                                ' Count of LTblB entries
GLOBAL LTblB()             AS LCtlCmd                             ' Complete line command list
GLOBAL LTblRange           AS LONG                                ' A line control range is available
GLOBAL LTblSCmd            AS STRING * 8                          ' LLCtl Srce Line command
GLOBAL LTblSFrom           AS LONG                                ' LLCtl Srce Cmd FromIX
GLOBAL LTblSTo             AS LONG                                ' LLCtl Srce Cmd ToIX
GLOBAL LTblSRpt            AS LONG                                ' LLCtl Srce Cmd Repeat
GLOBAL LTblSFlag           AS LONG                                ' LLCtl Srce Cmd Flag
GLOBAL LTblDCmd            AS STRING * 8                          ' LLCtl Dest Line command
GLOBAL LTblDFrom           AS LONG                                ' LLCtl Dest Cmd FromIX
GLOBAL LTblDTo             AS LONG                                ' LLCtl Dest Cmd ToIX
GLOBAL LTblDRpt            AS LONG                                ' LLCtl Dest Cmd Repeat
GLOBAL LTblDFlag           AS LONG                                ' LLCtl Dest Cmd Flag

'----- Command Parse Data

GLOBAL pCmdLen             AS INTEGER                             ' Length of pCommand on screen
GLOBAL pCmdNumOps          AS LONG                                ' Number of Operands
GLOBAL pCmdOps()           AS STRING                              ' Parsed out operands
GLOBAL pCmdRaw()           AS STRING                              ' Raw operand
GLOBAL pCmdOpsType()       AS LONG                                ' Type of operand

'----- CRTParse output data

GLOBAL CRTFlag             AS QUAD                                ' Contains %CRT.... status bits
GLOBAL CRTFCol             AS LONG                                ' From column
GLOBAL CRTTCol             AS LONG                                ' To   column
GLOBAL CRTL1               AS STRING                              ' Literal 1
GLOBAL CRTL2               AS STRING                              ' Literal 2
GLOBAL CRTL1RData          AS STRING                              ' Literal 1 massaged data
GLOBAL CRTL2RData          AS STRING                              ' Literal 2 massaged data
GLOBAL CRTL1Raw            AS STRING                              ' Literal 1 Raw
GLOBAL CRTL2Raw            AS STRING                              ' Literal 2 Raw
GLOBAL CRTHiLiteClr        AS LONG                                ' HiLite search color
GLOBAL CRTHiLiteOff        AS LONG                                ' HiLite Off color
GLOBAL CRTHiLiteOn         AS LONG                                ' HiLite ON  color

'----- Block cut / paste control

GLOBAL SlecthDC            AS DWORD                               ' For handling Mouse text frame select

'----- Global Object hooks
GLOBAL PCmdT               AS iPCmdTable                          ' Primary command table
GLOBAL LCmdT               AS iLCmdTable                          ' Line command table
GLOBAL PrimT               AS iPrimTable                          ' Primitive table
GLOBAL ENV                 AS iENVariables                        ' Environment stuff
GLOBAL KbdT                AS iKbdTable                           ' Keyboard table
GLOBAL KwdT                AS iKwdTable                           ' Keyword tables
GLOBAL IO                  AS iIO                                 ' I/O handler

'----- File Manager Data Areas
GLOBAL FM_File_Size        AS LONG                                ' Screen size for FileName
GLOBAL FM_Crit_Size        AS LONG                                ' Screen size for FM Criteria lines
GLOBAL FM_Note_Size        AS LONG
GLOBAL FM_Quick_Line_1     AS LONG
GLOBAL FM_Quick_Pos_1      AS LONG
GLOBAL FM_Quick_Pos_2      AS LONG
GLOBAL FM_Quick_Pos_3      AS LONG
GLOBAL FM_Quick_Pos_4      AS LONG
GLOBAL FM_Quick_Pos_5      AS LONG
GLOBAL FM_Quick_Pos_6      AS LONG
GLOBAL FM_Quick_Pos_7      AS LONG
GLOBAL FM_Quick_Pos_8      AS LONG
GLOBAL FM_Quick_Pos_9      AS LONG
GLOBAL FM_Path_Line        AS LONG
GLOBAL FM_Path_Left        AS LONG
GLOBAL FM_Mask_Line        AS LONG
GLOBAL FM_Mask_Left        AS LONG
GLOBAL FM_Misc_Length      AS LONG
GLOBAL FM_Head_Line        AS LONG
GLOBAL FM_Head_Name_Left   AS LONG
GLOBAL FM_Head_Raw_Left    AS LONG
GLOBAL FM_Head_Ext_Left    AS LONG
GLOBAL FM_Head_Size_Left   AS LONG
GLOBAL FM_Head_Date_Left   AS LONG
GLOBAL FM_Head_Time_Left   AS LONG
GLOBAL FM_Head_Lines_Left  AS LONG
GLOBAL FM_Head_Note_Left   AS LONG
GLOBAL FM_Top_File_Line    AS LONG
GLOBAL FM_List_Height      AS LONG
GLOBAL Rightmost           AS LONG

'---------- RegEx areas
GLOBAL PCRE_Regex_Str2     AS STRING                              ' String to build pseudo ASCIIZ string
GLOBAL PCRE_ErrPtr         AS ASCIIZ PTR                          ' Pointer to pcre_compile error message string
GLOBAL PCRE_ErrOffsetPtr   AS DWORD                               ' Pointer to offset in Regex string where error was detected
GLOBAL PCRE_Options        AS LONG                                ' Options passed to pcre
GLOBAL PCRE_Offsets()      AS LONG                                ' Table of offsets to located strings
GLOBAL hLib_PCRE           AS LONG                                ' Handle of PCRE3.dll library
GLOBAL hProc_PCRE_Compile  AS LONG                                ' Handle to Compile function
GLOBAL hProc_PCRE_Exec     AS LONG                                ' Handle to Exec function
GLOBAL hProc_PCRE_Free     AS LONG                                ' Handle to Free function
GLOBAL hProc_PCRE_Free_Ptr AS LONG                                ' Handle to Real Free function


'----- Macro Basic data
GLOBAL gMacroMode          AS LONG                                ' A Macro is running
GLOBAL gMacroName          AS STRING                              ' The macro name
GLOBAL gMacLines()         AS STRING                              ' The macro code
GLOBAL gMacThread          AS DWORD                               ' Macro thread handle
GLOBAL gStrVar             AS IPOWERCOLLECTION                    ' STR variable pool
GLOBAL gNumVar             AS IPOWERCOLLECTION                    ' NUM variable pool
GLOBAL gParseTbl           AS IPOWERCOLLECTION                    ' Parse table storage
GLOBAL gMacNoLoopTest      AS LONG                                ' User wants no Loop checking
GLOBAL gMacTempVars        AS LONG                                ' Temp variables need cleanup
GLOBAL gMacroTrace         AS LONG                                ' Trace functions 0 = No, 1 = Yes, 3 = Error only
GLOBAL gMacroTHeader       AS STRING                              ' Trace function name, params
GLOBAL gMacroRC            AS LONG                                ' Final RC
GLOBAL gMacroMsg           AS STRING                              ' Final Msg
GLOBAL gMacroFile          AS STRING                              ' Full macro filename
GLOBAL gMacRange           AS LONG                                ' Line control range at Macro start
GLOBAL gMacSCmd            AS STRING                              ' LLCtl Srce Line command
GLOBAL gMacSFrom           AS LONG                                ' LLCtl Srce Cmd FromIX
GLOBAL gMacSTo             AS LONG                                ' LLCtl Srce Cmd ToIX
GLOBAL gMacSRpt            AS LONG                                ' LLCtl Srce Cmd Repeat
GLOBAL gMacSFlag           AS LONG                                ' LLCtl Srce Cmd Flag
GLOBAL gMacDCmd            AS STRING                              ' LLCtl Dest Line command
GLOBAL gMacDFrom           AS LONG                                ' LLCtl Dest Cmd FromIX
GLOBAL gMacDTo             AS LONG                                ' LLCtl Dest Cmd ToIX
GLOBAL gMacDRpt            AS LONG                                ' LLCtl Dest Cmd Repeat
GLOBAL gMacDFlag           AS LONG                                ' LLCtl Dest Cmd Flag
GLOBAL gMacCore            AS DWORD                               ' Address of thinBasic_Core
GLOBAL gMacRelease         AS DWORD                               ' Address of thinBasic_Release
GLOBAL gMacFString         AS STRING                              ' Find string for EXEC Macro / MapStr
GLOBAL gMacCString         AS STRING                              ' Change string returned from EXEC MACRO / MapStr
GLOBAL gMacErrString       AS STRING                              ' Error message string returned from MapStr
GLOBAL mgotlit, mgotnum, mgotlptr, mgottag AS LONG                ' Parse gotten operand counts
GLOBAL mOprands() AS STRING

'----- Include code stored in INC files

#INCLUDE ONCE "_ObjENV.inc"                                       ' Environment stuff
#INCLUDE ONCE "_ObjIO.inc"                                        ' I/O handler
#INCLUDE ONCE "_ObjPCmdT.inc"                                     ' Primary Command Table Object
#INCLUDE ONCE "_ObjLCmdT.inc"                                     ' Line Command Table Object
#INCLUDE ONCE "_ObjKbdT.inc"                                      ' Keyboard tables
#INCLUDE ONCE "_ObjKwdT.inc"                                      ' Keyword tables
#INCLUDE ONCE "_ObjPrimT.inc"                                     ' Primitive table
#INCLUDE ONCE "_ObjProf.inc"                                      ' Profile Data
#INCLUDE ONCE "_TabData.inc"                                      ' CLASS for Tab Data
#INCLUDE ONCE "_PCRE.inc"                                         ' PCRE RegEx stuff


'==========
' Away we go!!!
'==========

FUNCTION WINMAIN (BYVAL hInst AS LONG, BYVAL hPrevInst AS DWORD, BYVAL lpszCmdLine AS WSTRINGZ PTR, BYVAL nCmdShow  AS LONG ) AS LONG
   hInstance = hInst                                              ' Save my Instance
   RealPBMain                                                     ' Call The real guy
'  profile "D:\Documents\SPFLite10\SPFLiteProfile.txt"            ' Uncomment to create a profile
END FUNCTION

FUNCTION RealPBMAIN AS LONG
LOCAL i, j, MeditOnce, cTab1, cTab2, FNum, x, y, nx, ny, h AS LONG, fn, fn2, cmd, t, u, Prof AS STRING
LOCAL tHndl, tResult, MRFOpen AS LONG, parmp, lptr AS LONG POINTER, lFD AS DIRDATA

'---------- Initialize the Global arrays

DIM Tabs(1)                        AS iObjTabData                 ' Dim the Tabs table
DIM gCmdRtrev(1 TO 50)             AS GLOBAL STRING               ' Retrievable Command line(s)
DIM pCmdOps(0 TO 500)              AS GLOBAL STRING               ' Command operands
DIM pCmdRaw(0 TO 500)              AS GLOBAL STRING               ' Raw operands
DIM pCmdOpsType(0 TO 500)          AS GLOBAL LONG                 ' Command operand types
DIM gHelpKey(300)                  AS GLOBAL STRING               ' Help Keyword    |__ DIM'ed the same
DIM gHelpMapid(300)                AS GLOBAL LONG                 ' Help MapID      |
DIM gPTbl(1 TO 500)                AS GLOBAL PTypeTable           ' PowerType table
DIM gFQ(1 TO 1000)                 AS GLOBAL WatchData            ' Global File Queue
DIM TTbl(1 TO 1000)                AS GLOBAL TouchEntry           ' Touch Table
DIM LTblA(500)                     AS GLOBAL LCtlScan             ' Initial line scan table
DIM LTblB(500)                     AS GLOBAL LCtlCmd              ' Complete line command list
DIM gLnoTextType(1 TO 15)          AS GLOBAL LONG                 ' Table of Line Number display types
DIM gLnoTextTxt(1 TO 15)           AS GLOBAL STRING               ' Table of Line Number display text
DIM gWinList(1 TO 500)             AS GLOBAL WININFOTYPE          ' Windows list
DIM gCrashList(0 TO 200)           AS GLOBAL STRING               ' Module trace
DIM mOprands(1 TO 50)              AS GLOBAL STRING               ' MACRO operands
DIM gTabDelList(1 TO 100)          AS GLOBAL LONG                 ' Tabs to be deleted this interrupt
DIM gTabDelNext(1 TO 100)          AS GLOBAL LONG                 ' Tabs to be switched to
DIM gTabStack(1 TO 100)            AS GLOBAL LONG                 ' Active tab stack
DIM PCRE_Offsets(12)               AS GLOBAL LONG                 ' Table of offsets to PCRE located strings
'----- Build some text arrays
DIM gEOLFlagList(7)                AS GLOBAL STRING               ' Pulldown menu for EOLFlag
DIM gRECFMList(5)                  AS GLOBAL STRING               '     "      "   "  RECFM
DIM gCmdList(1 TO 10)              AS GLOBAL STRING               ' Cmd list for DO processing
DIM nHiLites(1 TO 15)              AS GLOBAL STRING               ' HiLite color names
DIM gSBTable(1 TO 15)              AS GLOBAL SBarEntry            ' Status Bar table
ARRAY ASSIGN gEOLFlagList() = "CRLF", "LF", "CR", "NL", "AUTO", "AUTONL", "NONE", " "
ARRAY ASSIGN gRECFMList() = "U", "F", "V", "VBI", "VLI"
ARRAY ASSIGN nHiLites() = "BLUE", "GREEN", "YELLOW", "RED", "BLACK", "NAVY", "TEAL", "VIOLET", _
                          "ORANGE", "GRAY", "LIME", "CYAN", "PINK", "MAGENTA", "WHITE"
   nHiLitesChrs = "BGYRKNTVOALCPMW"                               ' 1 char matching IDs
   MEntry
   sAdjustFontSizes                                               ' Go Adjust font sizes

'----- Establish some Global Objects
LET gStrVar   = CLASS "PowerCollection"                           ' Macro global string variables
LET gNumVar   = CLASS "PowerCollection"                           ' Macro global numeriv variables
LET gParseTbl = CLASS "PowerCollection"                           ' Macro parse
LET ENV       = CLASS "cENVariables"                              ' Environment stuff
LET PCmdt     = CLASS "cPCmdTable"                                ' Primary Command Table
LET LCmdt     = CLASS "cLCmdTable"                                ' Line Command Table
LET Primt     = CLASS "cPrimTable"                                ' Primitive Table
LET KbdT      = CLASS "cKbdTable"                                 ' Keyboard tables
LET KwdT      = CLASS "cKwdTable"                                 ' Keyword tables
LET IO        = CLASS "cIO"                                       ' I/O Object

   '----- Actually start doing something
   SETUNHANDLEDEXCEPTIONFILTER CODEPTR(sLoopHandler)              ' Trap my exceptions
   RANDOMIZE TIMER                                                ' Good practice
   SetProcessDPIAware                                             ' So we can test DPI properly

   '---------- Get PCRE ready to use
   hLib_PCRE = LoadLibraryA( BYCOPY "PCRE3.Dll" )                 ' Get handle to the DLL
   IF hLib_PCRE THEN                                              ' PCRE exists

      '---Try to load the functions
      hProc_PCRE_Compile          = GetProcAddress(hLib_PCRE, BYCOPY "pcre_compile")
      hProc_PCRE_Exec             = GetProcAddress(hLib_PCRE, BYCOPY "pcre_exec")
      hProc_PCRE_Free             = GetProcAddress(hLib_PCRE, BYCOPY "pcre_free")
      lptr = hProc_PCRE_Free                                      ' Free returns a POINTER to the real free routine
      hProc_PCRE_Free_Ptr = @lptr                                 ' so chain to it as the REAL entry point
      '---If all went fine ...
      IF hProc_PCRE_Compile        AND _                          ' All three better be non-zero
         hProc_PCRE_Exec           AND _                          '
         hProc_PCRE_Free_Ptr       THEN                           '
         ' All is well                                            '
      ELSE
         MSGBOX "Internal PCRE functions not found"               ' Error
         EXIT FUNCTION                                            '
      END IF                                                      '

   ELSE                                                           '
      MSGBOX "PCRE DLL does not appear to be installed"           ' Say why we didn't do it
      EXIT FUNCTION                                               '
   END IF                                                         '
   InitLocalTables                                                ' Initialize some more Local tables

   InitLNText                                                     ' Go initialize line number text constants

   InitASynchStuff                                                ' Initialize the low priority stuff

   '----- Get a valid TP environment real early
   INCR TabsNum: INCR TabUnique                                   ' Add a Tab
   LET Tabs(TabsNum) = CLASS "cObjTabData"                        ' Build the Class entry
   TP = Tabs(tabsNum)                                             ' Make the new entry the active tab class
   TP.PgNumber = TabsNum                                          ' Save Page Number
   TP.LInitTxtData("")                                            ' Initialize the Editor text data

   '----- If first time user, let them do some settings
   IF ENV.FirstTime THEN                                          '
      DispOptions(1)                                              ' Go let user play with them
   END IF                                                         '


   '----- Display emergency KEYMAP if asked for
   IF IsENVKeyMap THEN KbdT.DispKeyMap(0)                         ' Recovery mode KEYMAP?

   IF InitSeeUnique THEN MExitFunc                                ' Go see if we should be unique


   '----- See if DEFAULT.INI exists
   IF ISFALSE ISFILE(ENV.PROFPath + "DEFAULT.INI") THEN           ' If no, then get one now
      DO                                                          ' Get a Profile created
         sDoMsgBox("The |KDEFAULT.INI |Bprofile is missing, or has not yet been created," + $CRLF +  _
                   "the Profile edit dialog will be displayed so you can choose the" + $CRLF +   _
                   "initial default values.   These can be altered at any time by using a" + $CRLF + _
                   "'|KPROFILE EDIT DEFAULT|B' command.", %MB_OK OR %MB_USERICON, "DEFAULT Profile Warning")
         TP.pCmdPROFILE("PROFILE NEW DEFAULT")                    ' display the dialog
      LOOP WHILE ISFALSE ISFILE(ENV.PROFPath + "DEFAULT.INI")     ' Loop in case they CANCEL out of the Dialog
   END IF


   t = sINIGetString("General", "MRFList", "")                    ' Get the MRF list
   IF ENV.InitString = "" AND t <> "" AND ISTRUE ENV.ReOpenLast THEN ' If no Cmd line and an existing MRF list
      ENV.InitString = t                                          ' Set MRF into the InitString
      MRFOpen = %True                                             ' Set MRF mode
      sIniSetString("General", "MRFList", "")                     ' Kill list so it is only used once
   END IF                                                         '

   sTabAddFManager                                                ' Add the FM tab

   '----- Create a dummy invisible basic screen to get font sizes

   DIALOG NEW PIXELS, 0, "SPFLite", ENV.LastScreenX, ENV.LastScreenY, 200, 200, _
          %WS_CAPTION OR %WS_THICKFRAME OR %WS_MINIMIZEBOX OR %WS_SYSMENU OR %WS_OVERLAPPEDWINDOW OR %WS_CLIPCHILDREN OR _
          %WS_DISABLED, 0 TO hWnd

   DIALOG SET ICON hWnd, "A"                                      ' Set TitleBaR Icon
   DIALOG SET COLOR hWnd, %RGB_WHITE, %RGB_DIMGRAY                ' Default color it
   DragAcceptFiles hWnd, %TRUE                                    ' Turn on Drag/Drop

   '----- Add the StatusBar
   CONTROL ADD STATUSBAR, hWnd, %IDC_StatusBar, "", 0, 0, 0, 0, %CCS_BOTTOM, %WS_EX_WINDOWEDGE    ' Add the Status Bar
   CONTROL HANDLE hWnd, %IDC_StatusBar TO hStatusBar              ' Save its handle
   CONTROL SET FONT   hWnd, %IDC_StatusBar, hSBFont               ' Set font
   CONTROL GET SIZE   hWnd, %IDC_StatusBar TO gSBWidth, gSBHeight ' Get the SB size
   gSBHeight += 4                                                 ' Make it a bit bigger

   '----- Create a dummy Graphic window just to get Font sizes
   CONTROL ADD GRAPHIC, hWnd, %IDC_SPFLiteWindow, "", 0, 0, 100, 100
   GRAPHIC ATTACH hWnd, %IDC_SPFLiteWindow                        ' Attach it
   FONT NEW ENV.FontName, ENV.FontPitch, ENV.FontStyle, 1, 1 TO hScrFont      ' Get the basic font
   FONT NEW ENV.FontName, ENV.FontPitch, ENV.FontStyle + 4, 1, 1 TO hScrFontUnd   ' Get the underline version of the basic font
   GRAPHIC SET FONT hScrFont                                      ' Set the desired font

   GRAPHIC CELL SIZE TO gFontWidth, gFontHeight                   ' Get size of a character
   CONTROL KILL hWnd, %IDC_SPFLiteWindow                          ' Kill dummy graphic area

   '----- Now calc our needed Graphic Window size
   x = (ENV.ScrWidth * gFontWidth) + %GLM + %GRM                  ' + the LM and RM pad values
   y = (ENV.ScrHeight * gFontHeight)                              ' Y

   '----- Now add a Tab control with no sizes yet
   CONTROL ADD TAB, hWnd, %IDC_SPFLiteTAB, "", 0, 0, 0, 0, %TCS_OWNERDRAWFIXED
   CONTROL HANDLE hWnd, %IDC_SPFLiteTab TO hTab                   ' Get handle for Tab
   CONTROL SET FONT hWnd, %IDC_SPFLiteTAB, hSBFont                '
   CONTROL SET COLOR hWnd, %IDC_SPFLiteTAB, cStatFG, cStatBG1     ' Default color it

   '----- Set Rect to our needed Graphic size and then set the Tab to that size
   gTabRC.nTop = 0:     gTabRC.nLeft = 0                          ' Init Tab Rect
   gTabRC.nRight = x: gTabRC.nBottom = y                          '
   TabCtrl_AdjustRect hTab, 1, gTabRC                             ' Set Tab display to suit graphic
   TabCtrl_GetItemRect hTab, 1, gTabHdrRC                         ' Get Tab title dimensions .nBottom = height
   CONTROL SET SIZE hWnd, %IDC_SPFLiteTAB, gTabRC.nRight - gTabRC.nLeft, gTabRC.nBottom - gTabRC.nTop + gTabHdrRC.nBottom

   '----- Insert a page (will be used for the FM)
   TAB INSERT PAGE hWnd, %IDC_SPFLiteTAB, TP.PgNumber, 0, $Empty, CALL DlgCallBack TO h
   TP.PgHandle = h                                                ' Save the handle

   '----- Add our Graphic window to the new Page
   CONTROL ADD GRAPHIC, TP.PgHandle, TP.WindowID, "", 0, 0, x, y  '
   CONTROL HANDLE TP.PgHandle, TP.WindowID TO h                   ' Save handle to graphic window
   TP.gHandle = h                                                 '
   GRAPHIC ATTACH TP.PgHandle, TP.WindowID                        ' Set as the default graphic area
   GRAPHIC CLEAR cTxtLoBG1                                        ' Clear the background
   TP.cCurrent = %False                                           ' Set cursor state
   GRAPHIC SET FONT hScrFont                                      ' Set the font
   TP.ScreenDim(ENV.ScrHeight, ENV.ScrWidth)                      ' Redim the Screen shadow copy

   '----- Now get the actual Tab size
   CONTROL GET SIZE hWnd, %IDC_SPFLiteTAB TO nx, ny               ' Now get the tab size
   DIALOG SET CLIENT hWnd, nx, ny + gSBHeight                     ' Resize the whole dialog, allowing for the headers

   '----- Calc some dependent values based on the screen size
   gDataLen  = ENV.ScrWidth - gLNPadCol                           ' Calc derived values
   pCmdLen   = ENV.ScrWidth - 24                                  '
   gwScrHeight = ENV.ScrHeight - ENV.PFKShow                      ' Shrink data area by PFK Show area
   InitFMLayout                                                   ' Adjust FM area
   sSetupSB                                                       ' Setup the status bar

   '----- Pretty up the display
   TP.WindowTitle                                                 ' Alter window/Tab titles
   SetCmd                                                         ' Put cursor to command line
   TAB SELECT hWnd, %IDC_SPFLiteTAB, TP.PgNumber                  ' Select the FM Tab
   GRAPHIC REDRAW                                                 '
   sDoCursor                                                      '

   IF ENV.Maximized THEN DIALOG MAXIMIZE hWnd                     ' Were we maximized? Return to that state

   '----- Open either CLIP mode or the InitFiles default
   IF IsENVClip THEN                                              ' -CLIP mode
      sTabAdd("", "DEFAULT")                                      ' Go do it
   ELSEIF ISNOTNULL(ENV.InitString) THEN                          ' Got an Init value?
      GOSUB InitFiles                                             ' Go open the tabs for the file (list)
   END IF                                                         '
   ENV.PMode = 0                                                  ' Clear PMode flags now

   '----- Start the loop watching thread
   IF ISFALSE IsPBDebuggerOn AND _                                ' Only if not in debugger
      ISFALSE ENV.NoLoopMode THEN                                 ' and not NOLOOP mode
      THREAD CREATE sLoopThread(parmp) 65536, TO tHndl            ' Fire up the thread
      SLEEP 50                                                    ' Wait a bit
      THREAD STATUS tHndl TO tResult                              ' See if running OK
      THREAD CLOSE tHndl TO tResult                               ' Free up our handle
   END IF                                                         '

   '----- Finally show the Dialog and let it run
   ENABLEWINDOW(hWnd, %True)                                      ' Remove disabled status
   DIALOG SHOW MODAL hWnd CALL DlgCallback                        ' Fire up the main Dialog
   SETUNHANDLEDEXCEPTIONFILTER %NULL                              ' Kill exception trap
   TRACE OFF
   MExitFunc

'----- Open either the $CMDLINE filename or MRF list

InitFiles:
   TP.CmdClear = %True                                            ' Clear command line next swap
   IF PARSECOUNT(ENV.InitString, "?") = 1 AND ENV.InitProfile <> "" THEN ' A cmd line .Profile
      Prof = ENV.InitProfile                                      ' Setup to use it
   END IF                                                         '
   FOR i = 1 TO PARSECOUNT(ENV.InitString, "?")                   ' Loop for each file in list
      fn = PARSE$(ENV.InitString, "?", i)                         ' Extract the next name

      '----- Open a normal type filename list
      IF ISNULL(fn) THEN ITERATE FOR                              ' ignore null strings
      IF INSTR(fn, "|") = 0 THEN                                  ' Just a normal filename?
         IF LEFT$(fn, 3) = "(e)" OR LEFT$(fn, 3) = "(b)" OR LEFT$(fn, 3) = "(v)" THEN  cTab1 = %True ' If lowercase, this was the active tab
         IF fn <> "NEW" THEN                                      ' If not the NEW quirk
            IF IsEQ(LEFT$(fn, 3), "(E)") THEN                     ' (E)?
               fn = MID$(fn, 4)                                   ' Strip off (E)
               ENV.PMode = ENV.PMode AND (&HFFFFFFFF - %MBrowse - %MView) ' Clear Browse and View
               ENV.PMode = ENV.PMode OR %MEdit                    ' Set Edit
            ELSEIF IsEQ(LEFT$(fn, 3), "(B)") THEN                 ' (B)?
               fn = MID$(fn, 4)                                   ' Just strip off (B)
               ENV.PMode = ENV.PMode AND (&HFFFFFFFF - %MEdit - %MView) ' Clear Edit and View
               ENV.PMode = ENV.PMode OR %MBrowse                  ' Remember it
            ELSEIF IsEQ(LEFT$(fn, 3), "(V)") THEN                 ' (V)?
               fn = MID$(fn, 4)                                   ' Just strip off (V)
               ENV.PMode = ENV.PMode AND (&HFFFFFFFF - %MBrowse - %MEdit) ' Clear Browse and Edit
               ENV.PMode = ENV.PMode OR %MView                    ' Remember it
            END IF                                                '
            IF ISNULL(PATHSCAN$(FULL, fn)) THEN                   ' Valid name?
               IF MRFOpen THEN                                    ' Re-Open style?
                  sDoMsgBox "While re-opening previous files, the following file cannot be found:" + $CRLF + $CRLF + _
                  "|K" + fn, %MB_OK OR %MB_USERICON, "SPFLite - Missing File"                    '
                  ITERATE FOR                                     '
               ELSE                                               ' A command line OPEN
                  sMakeNullFile(fn)                               ' Create it as an empty file
                  sDoMsgBox "The specified file (|K" + fn + "|B), cannot be found:" + $CRLF + _
                            "Created as a Zero length file", %MB_OK OR %MB_USERICON, "SPFLite - Missing File"                    '
               END IF
            END IF                                                '
            IF ISFALSE IsEnvBrowse AND ISFALSE IsEnvView THEN     ' Opening for EDIT?
               t = DIR$(fn, TO lFD)                               ' Get all the DIR info
               IF ISNOTNULL(t) AND (lFD.FileAttributes AND %FILE_ATTRIBUTE_READONLY) = %FILE_ATTRIBUTE_READONLY THEN
                  ENV.PMode = ENV.PMode AND (&HFFFFFFFF - %MEdit) ' Switch to View
                  ENV.PMode = ENV.PMode OR %MView                 '
                  sDoMsgBox "The specified Edit file (|K" + fn + "|B), is now Read-Only." + $CRLF + _
                            "It will be opened in |KView", %MB_OK OR %MB_USERICON, "SPFLite - Read-Only File"
               END IF
            END IF                                                '
         ELSE                                                     '
            fn = ""                                               ' Null for a NEW request
         END IF                                                   '
         fn2 = fn                                                 ' Copy for SetFMCrit
         GOSUB SetFMCrit                                          ' Go set FM c riteria

         IF IsEnvBrowse THEN                                      ' Opening in Browse mode?
            TP.CmdClear = %True                                   ' Clear command line next swap
            pCmdBrowse("BROWSE " + $DQ + fn + $DQ + " " + Prof)   ' Go do it
            Prof = ""                                             ' Prof use only once
         ELSEIF IsEnvView THEN                                    ' Opening in View mode?
            TP.CmdClear = %True                                   ' Clear command line next swap
            pCmdView("VIEW " + $DQ + fn + $DQ + " " + Prof)       ' Go do it
            Prof = ""                                             ' Prof use only once
         ELSE                                                     '
            TP.CmdClear = %True                                   ' Clear command line next swap
            pCmdEdit("EDIT " + $DQ + fn + $DQ + " " + Prof)       ' Go do it
            Prof = ""                                             ' Prof use only once
         END IF                                                   '
         IF cTab1 THEN cTab2 = TabsNum: cTab1 = %False            ' If that was the active tab, remember the tab number

      '----- Open a MEDIT session again

      ELSE                                                        ' A MEdit string
         cmd = "MEDIT "                                           ' Init command
         FOR j = PARSECOUNT(fn, "|") TO 1 STEP -1                 ' Loop for each file in list
            fn2 = PARSE$(fn, "|", j)                              ' Extract the next MEdit name
            IF LEFT$(fn2, 3) = "(e)" THEN cTab1 = %True           ' Remember if this was the active tab
            IF IsEQ(LEFT$(fn2, 3), "(E)") THEN fn2 = MID$(fn2, 4) ' Strip off the (E)
            IF ISNULL(TRIM$(fn2)) THEN ITERATE FOR                '
            cmd += $DQ + fn2 + $DQ + " "                          '
            IF ISFALSE MeditOnce THEN                             ' Just once
               MeditOnce = %True                                  '
               IF TP.LastLine > 2 THEN cmd += "NEW "              ' If we can't use tab 1
               GOSUB SetFMCrit                                    ' Go set FM criteria
            END IF                                                '
         NEXT j                                                   '
         pCmdMEdit(cmd)                                           '
         IF cTab1 THEN cTab2 = TabsNum: cTab1 = %False            ' If that was the active tab, remember the tab number
      END IF                                                      '
   NEXT i                                                         ' Loop-de-loop
   ENV.PMode = ENV.PMODE AND (&HFFFFFFFF - %MEdit - %MBrowse - %MView) ' Clear EDIT, BROWSE and VIEW settings
   IF cTab2 > 0 THEN
      TP = Tabs(CTab2)                                            ' Switch to correct tab
      TAB SELECT hWnd, %IDC_SPFLiteTAB, TP.PgNumber               ' Select the new tab
      TP.WindowTitle                                              ' Alter window title
      sDoCursor                                                   ' Get cursor going
   END IF                                                         '
   gTabSwitch = 0                                                 ' Clear the normal tab switcher, we've done it
   RETURN                                                         '

SetFMCrit:
   IF TP.LastLine > 2 THEN                                        ' If 1st tab already in use
      ENV.FMPath = LEFT$(fn2, INSTR(-1, fn2, "\"))                ' Pass FM variables to the called command
      ENV.FMMask = "*"                                            '
      ENV.FMFileList = ""                                         '
   ELSE                                                           ' Else we'll be using this one
      TP.OFrmFPath = LEFT$(fn2, INSTR(-1, fn2, "\"))              ' Save any where we were started from values
      TP.OFrmFMask = "*"                                          '
      TP.OFrmFileL = ""                                           '
   END IF                                                         '
   RETURN
END FUNCTION

#INCLUDE ONCE "_InitRoutines.inc"                                 ' Initialization routines
#INCLUDE ONCE "_DialogStuff.inc"                                  ' Dialog Build and Callbacks
#INCLUDE ONCE "_Mainline.inc"                                     ' Mainline routine
#INCLUDE ONCE "_BMacro.inc"                                       ' thinBasic Macro Support
#INCLUDE ONCE "_mapping.inc"                                      ' CHANGE M' literal support
#INCLUDE ONCE "_mapping_calc.inc"                                 ' Calc support for M' literals


SUB      DEBUG (st AS STRING)
'---------- Print stuff to a console
STATIC Consl AS LONG
LOCAL szConsole AS ASCIIZ * 1024
   '----- Allocate a console if we haven't already done so
   IF Consl = 0 THEN
      AllocConsole
      SetConsoleTitle "PB Diagnostic Console"
      Consl = GetStdHandle(%STD_OUTPUT_HANDLE)
      SetConsoleTextAttribute Hwnd&, %FOREGROUND_RED OR _
                                     %FOREGROUND_GREEN OR _
                                     %FOREGROUND_BLUE
   END IF

   '----- print the line
   IF Consl > 0 THEN
      szConsole = st & $CRLF
      WriteConsole Consl, szConsole, LEN(szConsole), %NULL, %NULL
    END IF
END SUB

SUB      DEBUGw (st AS WSTRING)
'---------- Print stuff to a console
STATIC Consl AS LONG
LOCAL szConsole AS ASCIIZ * 1024, i AS LONG
   '----- Allocate a console if we haven't already done so
   IF Consl = 0 THEN
      AllocConsole
      SetConsoleTitle "PB Diagnostic Console"
      Consl = GetStdHandle(%STD_OUTPUT_HANDLE)
      SetConsoleTextAttribute Hwnd&, %FOREGROUND_RED OR _
                                     %FOREGROUND_GREEN OR _
                                     %FOREGROUND_BLUE
   END IF

   '----- print the line
   IF Consl > 0 THEN
      FOR i = 1 TO LEN(st)
         szConsole = HEX$(ASC(MID$(st, i)), 4) + " "
         WriteConsole Consl, szConsole, LEN(szConsole), %NULL, %NULL
      NEXT i
      szConsole = $CRLF
      WriteConsole Consl, szConsole, LEN(szConsole), %NULL, %NULL
    END IF
END SUB

SUB      DEBUGAT (ctr AS LONG, st AS STRING)
'---------- Print stuff to a console
STATIC Consl AS LONG
LOCAL szConsole AS ASCIIZ * 1024
STATIC yet AS LONG

   INCR yet
   IF yet < ctr THEN EXIT SUB
   IF yet > ctr + 25  THEN EXIT SUB
   '----- Allocate a console if we haven't already done so
   IF Consl = 0 THEN
      AllocConsole
      SetConsoleTitle "PB Diagnostic Console"
      Consl = GetStdHandle(%STD_OUTPUT_HANDLE)
      SetConsoleTextAttribute Hwnd&, %FOREGROUND_RED OR _
                                     %FOREGROUND_GREEN OR _
                                     %FOREGROUND_BLUE
   END IF

   '----- print the line
   IF Consl > 0 THEN
      szConsole = st & $CRLF
      WriteConsole Consl, szConsole, LEN(szConsole), %NULL, %NULL
   END IF
END SUB


SUB      DEBUGLog (st AS STRING)
'---------- Print stuff to a file log
STATIC FNum, ctr AS LONG, fn AS STRING
   '----- Allocate a file if we haven't already done so
   IF FNum = 0 THEN
      fn = "C:\SPFLite\SPFLite.log"
      FNum = FREEFILE
      TRY
         IF ISFALSE ISFOLDER("C:\SPFLite") THEN MKDIR "C:\SPFLite"
      CATCH
         MSGBOX "Can't create C:\SPFLite folder"
      END TRY
      TRY
         OPEN fn FOR OUTPUT AS # FNum
      CATCH
         MSGBOX "Can't Open C:\SPFLite\SPFLite.log"
      END TRY
   END IF

   IF FNum > 0 THEN
      PRINT #FNum, st
      FLUSH #FNum
   END IF
END SUB

SUB      DEBUGLog2 (st AS STRING)
'---------- Print stuff to a file log
STATIC FNum2, ctr AS LONG, fn2 AS STRING
   '----- Allocate a file if we haven't already done so
   IF FNum2 = 0 THEN
      fn2 = "D:\Documents\SPFLite\SPFLite.Log2.log"
      FNum2 = FREEFILE
      OPEN fn2 FOR OUTPUT AS # FNum2
   END IF

   IF FNum2 > 0 THEN
      PRINT #FNum2, st
   END IF
END SUB

SUB      DEBUGStack()
REGISTER i AS LONG
   DEBUG "--- Stack ---"                                          ' Dump the stack
   FOR i = gCrashCtr - 1 TO 0 STEP -1
         Debug FORMAT$(i, "00") + " | " + gCrashList(i)
   NEXT i
END SUB

THREAD FUNCTION DispANSI2 (BYVAL colMode AS LONG) AS LONG         ' Start the ANSI window
'---------- Display the ANSI table
' (CharSet)    passes colMode = %False via METHOD DispANSI in _TabData
' (CharSetCol) passes colMode = %True  via METHOD DispANSI in _TabData

LOCAL i, j AS LONG, TX, Char, txChar AS STRING

THREADED tls_hANSI AS DWORD
THREADED tls_colMode AS LONG
THREADED tls_PrfGetX2APtr AS LONG                                 ' for either COLLATE or SOURCE tran table
THREADED tls_PrfGetA2XPtr AS LONG                                 ' for either COLLATE or SOURCE tran table
THREADED tls_ansiSource AS STRINGZ * 32                           ' a name like ANSI or EBCDIC

   MEntry

   tls_colMode = colMode

   IF TP.PrfCollateXlate THEN
      tls_ansiSource = TP.PrfPCollate                             ' a name like EBCDIC
      tls_PrfGetX2APtr = TP.PrfGetCS2APtr                         ' pointer to COLLATE tran table
      tls_PrfGetA2XPtr = TP.PrfGetCA2SPtr                         ' pointer to COLLATE tran table

   ELSEIF TP.PrfSrceXlate THEN
      tls_ansiSource = TP.PrfPSource                              ' a name like EBCDIC
      tls_PrfGetX2APtr = TP.PrfGetSS2APtr                         ' pointer to SOURCE  tran table
      tls_PrfGetA2XPtr = TP.PrfGetSA2SPtr                         ' pointer to SOURCE  tran table

   ELSE
      tls_ansiSource = "ANSI"
      tls_PrfGetX2APtr = 0                                        ' no active translation table
   END IF

   ' *** test for hANSI being active was removed. these dialogs can be concurrent

   DIALOG DEFAULT FONT ENV.FontName, ENV.FontPitch, ENV.FontStyle, 0

   DIALOG NEW PIXELS, 0, tls_ansiSource + " - Left Click replaces Clipboard - Right Click adds to Clipboard",,, _
                      54 * gFontWidth, 23 * gFontHeight, _
                      %WS_CAPTION OR %WS_SYSMENU OR %WS_MINIMIZEBOX _
                      TO tls_hANSI

   CONTROL ADD GRAPHIC, tls_hANSI, 5001, "", 0, 0, 54 * gFontWidth, 24 * gFontHeight
   GRAPHIC ATTACH tls_hANSI, 5001                                 ' Set as the default graphic area
   GRAPHIC SET FONT hScrFont                                      ' Set the desired font

   GRAPHIC CLEAR cTxtLoBG1                                        ' Clear it
   GRAPHIC COLOR cTxtLoFG, cTxtLoBg1                              '

   sWinclip_get(TX)                                               ' Get any current text to start

   TX = PARSE$(TX, $CRLF, 1)                                      ' Get 1st 'line'

   '----- Finally draw the chart
   GRAPHIC SET POS (1, 1)                                         ' Starting position

   GOSUB ShowHorizontalCaption

   FOR i = 1 TO 16                                                ' Display the table
      GRAPHIC SET POS (1, gFontHeight + i * gFontHeight)          ' Set position for one row

      GOSUB ShowVerticalCaption                                   ' Print the left hand heading

      FOR j = 1 TO 16                                             ' Now a row of characters
         GRAPHIC SET POS (4 * gFontWidth + ((j - 1) * (gFontWidth * 3)), gFontHeight + i * gFontHeight) ' Set position for Print

         IF tls_colMode THEN
            char = CHR$(CHR$(((j-1) * 16) + (i-1)))               ' Create character - column mode order
         ELSE
            char = CHR$(CHR$(((i-1) * 16) + (j-1)))               ' Create character - row mode order
         END IF

         IF tls_PrfGetX2APtr THEN                                 ' If not ANSI
            TP.Translate(Char, tls_PrfGetX2APtr)                  ' Get the ANSI equivalent
         END IF

         GRAPHIC PRINT Char                                       ' Print the next ANSI char
      NEXT j

      GRAPHIC SET POS (5 * gFontWidth + (16 * (gFontWidth * 3) - gFontWidth), gFontHeight + i * gFontHeight)
      GOSUB ShowVerticalCaption                                   ' Print the right hand heading

   NEXT i
   GRAPHIC SET POS (1, 19 * gFontHeight)                          ' Starting position

   GOSUB ShowHorizontalCaption

   '  New format on bottom of ANSI popup, with new line added:
   '
   '  Dec Hex Chr Len Current ANSI Clipboard:
   '  051  33 '3'  6  "ABC123"
   '                  |
   '  1234567890123456|

   GRAPHIC SET POS (1, 21 * gFontHeight)                          ' Set position for last row
   GRAPHIC PRINT "Dec Hex Chr Len Current " + tls_ansiSource + " Clipboard:"  ' Clipboard data caption

   ' we will make the initial display be consistant as if we just typed the last character
   ' in the clipboard.  if the clipboard's length is zero, we pretend they entered a NUL,
   ' which they can't really do.

   IF LEN(TX) = 0 THEN
      Char = $NUL
      txChar = " "
   ELSE
      Char = RIGHT$(TX, 1)
      txChar = Char
   END IF

   IF tls_PrfGetX2APtr THEN                                       ' If not ANSI
      TP.Translate(Char, tls_PrfGetA2XPtr)                        ' Get the non-ANSI equivalent
   END IF

   GRAPHIC SET POS (1, 22 * gFontHeight)                          ' Set position for last row

   GRAPHIC PRINT FORMAT$(ASC(Char), "000") + "  " + HEX$(ASC(Char), 2) _
      + " '" + txChar + "'" + DEC$(LEN(TX), -3) + "  " + $DQ + TX + $DQ + SPACE$(50)

   GRAPHIC REDRAW

   '/ store pointers to tls data so dialog can grab them

   DIALOG SET USER tls_hANSI, %uda_colMode,      tls_colMode
   DIALOG SET USER tls_hANSI, %uda_PrfGetX2APtr, tls_PrfGetX2APtr
   DIALOG SET USER tls_hANSI, %uda_ansiSource,   VARPTR (tls_ansiSource)

   DIALOG SHOW MODAL tls_hANSI CALL DlgANSICallback               ' Display ANSI screen

   MExitFunc

ShowHorizontalCaption:
   IF tls_colMode THEN
      GRAPHIC PRINT "   00 10 20 30 40 50 60 70 80 90 A0 B0 C0 D0 E0 F0" ' Print heading
   ELSE
      GRAPHIC PRINT "   00 01 02 03 04 05 06 07 08 09 0A 0B 0C 0D 0E 0F" ' Print heading
   END IF
   RETURN

ShowVerticalCaption:
   IF tls_colMode THEN
      GRAPHIC PRINT "0" + HEX$(i - 1, 1)                          ' Print left/right heading as 0x
   ELSE
      GRAPHIC PRINT       HEX$(i - 1, 1) + "0"                    ' Print left/right heading as x0
   END IF
   RETURN

END FUNCTION

FUNCTION PCRE_Regex_Compile(str1 AS STRING) AS STRING
'---------- Compile / Test a Regex string
LOCAL txtP AS ASCIIZ PTR, lclPCRE AS DWORD

   MEntry
   lclPCRE = TP.hPCRE                                             ' Get PCRE handle
   IF lclPCRE <> 0 THEN                                           ' Any prior area
      CALL DWORD hProc_PCRE_Free_Ptr USING pcre_free( _           ' Free it
                 lclPCRE)                                         ' Compile handle
      TP.hPCRE = 0                                                '
   END IF                                                         '
   PCRE_Options = IIF(TP.PrfPCase = "C", 0, %PCRE_CASELESS)       ' Set Options to match CASE
   PCRE_Regex_Str2 = str1 + CHR$(0)                               ' Make into pseudo ASCIIZ
   CALL DWORD hProc_PCRE_Compile USING pcre_compile( _            ' Try the compile
              BYVAL STRPTR(PCRE_Regex_Str2), _                    ' Regex string
              BYVAL PCRE_Options,            _                    ' Options
              BYVAL VARPTR(PCRE_ErrPtr),     _                    ' Pointer to error string
              BYVAL VARPTR(PCRE_ErrOffsetPtr),_                   ' Error offset
              BYVAL &0) _                                         ' Character tables
              TO lclPCRE                                          ' Answer area
   IF lclPCRE = 0 THEN                                            ' OK?
      txtp = PCRE_ErrPtr                                          ' No, get error message
      FUNCTION = "at Col: " + FORMAT$(PCRE_ErrOffsetPtr + 1) + " : " +  @Txtp
      MExitFunc                                                   '
   END IF                                                         '
   FUNCTION = ""                                                  ' Return null to indicate OK
   TP.hPCRE = lclPCRE                                             ' Save PCRE area for this tab
   MExit
END FUNCTION

SUB PCRE_Regex_Test(str1 AS STRING, scol AS LONG, fcol AS LONG, flen AS LONG)
'---------- PCRE regex test
LOCAL strt, RC AS LONG, lclPCRE AS DWORD
LOCAL optr AS BYTE PTR
   MEntry
   lclPCRE = TP.hPCRE                                             ' Get Compile area
   strt = scol - 1                                                ' Calc start column
   optr = VARPTR(PCRE_Offsets(0))                                 ' Point at answer array
   CALL DWORD hProc_PCRE_Exec USING pcre_exec( _                  ' Call for the test
              lclPCRE,                         _                  ' Compile handle
              &0,                              _                  ' extra-data
              STRPTR(str1),                    _                  ' Test-string
              LEN(str1),                       _                  ' length of Test-srtring
              strt,                            _                  ' Starting position
              &0,                              _                  ' Options
              optr,                            _                  ' PCRE_Offsets array
              &12)                             _                  ' Size of offsets array
              TO RC                                               ' Answer area

   IF RC < 1 THEN                                                 ' How'd search go?
      fcol = 0: flen = 0                                          ' Not found
   ELSE                                                           '
      fcol = PCRE_Offsets(0) + 1                                  ' Pass back column found in
      flen = PCRE_Offsets(1) - PCRE_Offsets(0)                    ' And length
   END IF                                                         '
   MExit
END SUB

FUNCTION sAdd128(str AS STRING) AS STRING
'---------- Shift a string upward by 128
REGISTER i AS LONG
LOCAL t AS STRING
LOCAL letteri, lettero AS BYTE POINTER
   MEntry
   t = SPACE$(LEN(str))                                           ' Make string as long as input
   letteri = STRPTR(str): lettero = STRPTR(t)                     ' Point at 1st byte of each
   FOR i = 1 TO LEN(str)                                          ' Loop through string
      @lettero = @letteri + 128                                   ' Process a byte
      INCR letteri: INCR lettero                                  ' Bump pointers
   NEXT i                                                         '
   FUNCTION = t                                                   ' Pass back result
   MExit
END FUNCTION

FUNCTION sString_is_proper(BYVAL test_line AS STRING, BYVAL i AS LONG, BYVAL check_escape AS LONG) AS LONG
'/-----------------------------------------------------------------------------
'/ string_is_proper
'/
'/ test_line at position 'i' should contain a quote
'/ if there is at least one non-escaped quote of the same type after
'/ position 'i', then the string at 'i' is a proper string; otherwise it is
'/ unclosed.  if the string is proper, return TRUE else FALSE
'/-----------------------------------------------------------------------------
LOCAL n AS LONG, quote AS STRING
   MEntry
   n = LEN (test_line)

   '/ if 'i' is at the last position of test_line or beyond, there can't
   '/ possibly be a proper string

   IF n = 0 OR i < 1 OR i >= n  THEN
      FUNCTION = 0
      MExitFunc
   END IF

   quote = MID$ (test_line, i, 1)
   IF quote <> $DQ AND quote <> "'" THEN
      FUNCTION = 0          '/ there was supposed to be a quote at position 'i'
      MExitFunc
   END IF

   '/ scan remainder of line looking for close quote.
   '/ don't look at position 'i' in the loop because we already did that above.

   DO WHILE i < n
      i += 1
      IF MID$ (test_line, i, 1) = quote THEN
         FUNCTION = 1                                    '/ found closing quote
         MExitFunc
      END IF
      IF check_escape = 1 THEN
         IF MID$ (test_line, i, 1) = "\" THEN
            '/ skip over the escape, and next DO loop will skip over escaped char
            i += 1
         END IF
      END IF
   LOOP
   FUNCTION = 0                                    '/ never found closing quote
   MExit
END FUNCTION ' string_is_proper

FUNCTION sSub128(str AS STRING) AS STRING
'---------- Shift a string upward by 128
REGISTER i AS LONG
LOCAL t AS STRING
   MEntry
   FOR i = 1 TO LEN(str)                                          ' Loop through string
      t += IIF$(ASC(str, i) > 127, CHR$(ASC(str, i) - 128), CHR$(ASC(str, i))) '
   NEXT i                                                         '
   FUNCTION = t                                                   ' Pass back result
   MExit
END FUNCTION

SUB      sAddCheck(wnd AS LONG, ID AS LONG, Value AS LONG, x AS LONG, y AS LONG, lgth AS LONG, Txt1 AS STRING)
'---------- Add a checkbox to a Tab
   MEntry
   CONTROL ADD CHECKBOX, wnd, ID, Txt1, x, y, lgth, 12
   CONTROL SET COLOR     wnd, ID, %WHITE, -2
   CONTROL SET CHECK     wnd, ID, Value
   MExit
END SUB

SUB sAdjustFontSizes()
'----- See if a -SCALE entered
LOCAL hDC, LPIy, factor AS LONG
   hDC = GetDC(%HWND_DESKTOP)                                     ' Get Desktop handle
   LPIy = GetDeviceCaps(hDC, %LOGPIXELSY)                         ' Get pixels / inch vertically
   ReleaseDC %HWND_DESKTOP, hDC                                   ' Free hDC
   factor = (LPiY/96) * 100                                       ' Calc % font size
   gFontScale = factor / 100                                      ' Create a factor out of it
   '--------------------------------------------------------------+
   ' Get the Fonts created                                        |
   '--------------------------------------------------------------+
   FONT NEW "Arial",       10 / gFontScale, 1, 1, 1 TO hBoldFont  ' Build font for our Preferences Dialog
   FONT NEW "Courier New", 10 / gFontScale, 1, 1, 1 TO hFixedFont ' Build font for our Preferences Dialog
   FONT NEW "Tahoma",      10 / gFontScale, 0, 1, 1 TO hSBFont    ' Build font for the Status Bar
   FONT NEW "Tahoma",      10 / gFontScale, 1, 1, 1 TO hSBFontB   ' Build font for the Status Bar

END SUB


FUNCTION sAutoMask(Fn AS STRING) AS STRING
'---------- See if filename matches an AUTOFAV mask
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL Masks() AS ASCIIZ * 50, k, NumMask AS LONG
LOCAL FavNames() AS STRING, MPtr, FPtr, TPtr AS BYTE POINTER
LOCAL FName AS ASCIIZ * %MAX_PATH, tmask AS ASCIIZ * 50
DIM Masks(1 TO 50) AS ASCIIZ * 50
DIM FavNames(1 TO 50) AS STRING
   MEntry
   IF gSetCount > 0 THEN                                          ' Scan SET table
      FOR i = 1 TO gSetCount                                      '
         IF IsEQ(LEFT$(gSetKey(i), 8), "AUTOFAV.") THEN           ' is this an AUTOFAV. entry?
            INCR NumMask                                          ' We have one
            IF NumMask > UBOUND(Masks()) THEN                     ' Table exceeded
               REDIM PRESERVE Masks(1 TO 2 * NumMask) AS ASCIIZ * 50 ' Enlarge it
               REDIM PRESERVE FavNames(1 TO 2 * NumMask) AS STRING'
            END IF                                                '
            Masks(NumMask) = UUCASE(MID$(gSetKey(i), 9))          ' Copy the mask
            '----- Each SET item could be a Stack
            k = PARSECOUNT(gSetData(i), BINARY)                   ' Get count of number in stack
            REDIM SetVar(1 TO k) AS STRING                        ' Dim variable table
            PARSE gSetData(i), SetVar(), BINARY                   ' Extract the table
            FavNames(NumMask) = SetVar(1)                         ' Copy the FAV name
         END IF                                                   '
      NEXT i                                                      '
   END IF                                                         '
   IF NumMask = 0 THEN FUNCTION = "": MExitFunc                   ' No masks, return null

   '----- We have some masks, test them
   FUNCTION = ""                                                  ' Say we didn't find it
   FName = UUCASE(fn)                                             '
   FName = PATHNAME$(NAMEX, FName)                                ' Just use the filename portion

   FOR j = 1 TO NumMask                                           ' Now for each mask
      tmask = UUCASE(Masks(j))                                    ' Copy it

      MPtr = VARPTR(tmask)                                        ' Point at mask string
      FPtr = VARPTR(FName)                                        '
      DO WHILE @Mptr AND @FPtr                                    ' While each point at something
         SELECT CASE @MPtr                                        ' Work through the mask characters
            CASE 63                                               ' ? = Match any character
               INCR MPtr: INCR FPtr                               ' Step each pointer
            CASE 42                                               ' * = Match any string
               TPtr = MPtr + 1                                    ' Point at next mask char
               IF ISFALSE @TPtr THEN GOSUB PassName: MExitFunc    ' No more mask, so this mask is a win
               DO WHILE @TPtr <> @FPtr AND @FPtr                  ' Scan filename looking for a match
                  INCR FPtr                                       '
               LOOP                                               '
               IF ISFALSE @FPtr THEN EXIT DO                      ' Char not found
               INCR MPtr                                          ' Matched step over them
            CASE ELSE                                             ' Any other char must match
               IF @MPtr <> @FPtr THEN EXIT DO                     ' No, skip this mask
               INCR MPtr: INCR FPtr                               ' Yes, continue
         END SELECT                                               '
         IF ISFALSE @FPtr AND @MPtr = 42 THEN GOSUB PassName: MExitFunc
      LOOP                                                        '
      IF ISFALSE @MPtr AND ISFALSE @FPtr THEN GOSUB PassName: MExitFunc
   NEXT j
   MExitFunc

PassName:
   IF IsEQ(FavNames(j), "FAV") THEN FUNCTION = "Favorite Files": RETURN ' Handle aliases
   IF IsEQ(FavNames(j), "FAVORITE") THEN FUNCTION = "Favorite Files": RETURN
   IF IsEQ(FavNames(j), "FAVOURITE") THEN FUNCTION = "Favorite Files": RETURN
   FUNCTION = FavNames(j)                                         ' Just return it
   RETURN                                                         '
END FUNCTION

FUNCTION sBinSort(p1 AS TouchEntry, p2 AS TouchEntry) AS LONG
'---------- Support sort of the FMFiles array
LOCAL P1Fn, P2Fn AS STRING
   MEntry
   IF p1.LinNo < p2.LinNo THEN FUNCTION = -1: MExitFunc
   IF p1.LinNo > p2.LinNo THEN FUNCTION = +1: MExitFunc
   FUNCTION = 0
   MExit
END FUNCTION

SUB      sCalcEditBG(lno AS LONG)
'------- Calc bandinng value
   cBandBG = %False                                               ' Start as not needed
   IF lno <= (3 + TP.TPPrfGetCols) THEN EXIT SUB                  ' Off if in Title area
   IF ISTRUE ENV.Banding THEN                                     ' Are we banding?
      cBandBG = IIF((lno) \ 3 MOD 2 = 0, %True, %False)           ' Setup based on line number
   END IF                                                         '
END SUB

SUB      sCalcFMBG(lno AS LONG)
'------- Calc bandinng value
   cBandBG = %False                                               ' Start as not needed
   IF lno <= FM_Top_File_Line THEN EXIT SUB                       ' Off if in Title area
   IF ISTRUE ENV.Banding THEN                                     ' Are we banding?
      cBandBG = IIF((lno + 2) \ 3 MOD 2 = 0, %True, %False)       ' Setup based on line number
   END IF                                                         '
END SUB

SUB      sCallTabCmd(tn AS LONG, tcmd AS STRING)
'---------- Execute a command in another tab
LOCAL OurTab AS LONG
   MEntry
   OurTab = TP.PgNumber                                           ' Save our page number
   TP = Tabs(tn)                                                  ' Switch to the tab
   TP.pCommand = tcmd                                             ' Stuff in the command
   TP.AttnDo =  (TP.AttnDo OR %Attention)                         ' Request Attention
   TP.PostKeyboard                                                ' Go try it
   TP = Tabs(OurTab)                                              ' Switch back to original tab
   sDoPendingTabDels                                              ' In case a Del is needed
   MExit                                                          ' We're done
END SUB

SUB      scError(sev AS LONG, pMsg AS STRING)
'---------- Handle command error
   IF sev > TP.errFlag OR ISNULL(TP.ErrMsg) THEN                  '
      TP.ErrMsg = pMsg                                            ' Stuff in the message text
      TP.errFlag = sev                                            ' Set the Error severity level
   END IF                                                         '
   TP.ErrMAdd(pMsg, sev)                                          ' Go Queue the message text
END SUB

SUB      sCloneEdit(fn AS STRING)
'---------- Start another SPFLite process edit a file
LOCAL lclDrive, lclPath, lclCmd, lclMode, lclPCmd AS STRING
LOCAL RetC AS LONG

   MEntry
   '----- Build a command to start another SPFLite Instance
   lclPCmd = TP.CurrPCmd                                          ' Get current command name
   lclMode = SWITCH$(lclPCmd = "OPEN", " -OPEN ", lclPCmd = "OPENV", " -OPENV ", lclPCmd = "OPENB", " -OPENB ")
   lclPath = CURDIR$                                              ' Locate where we are
   IF MID$(lclPath, 2, 1) = ":" THEN _                            ' Extract Drive if present
      lclDrive = LEFT$(lclPath, 2)                                ' and save for restore
   IF MID$(ENV.EXEPath, 2, 1) = ":" THEN _                        ' See if gEXEPath has drive
      CHDRIVE LEFT$(ENV.EXEPath, 2)                               ' If so, go to it
   CHDIR ENV.EXEPath                                              ' Switch to EXE path
   lclCmd = $DQ + EXE.FULL$ + $DQ + lclMode + $DQ + fn + $DQ      ' Build command string

   '----- Issue the command
   RetC = SHELL(lclCmd,1)                                         '
   IF ERR THEN                                                    ' Tell user result
      scError(%eFail, "Error issuing START command, " + ERROR$)   '
   ELSE                                                           '
      scError(0, "New instance of SPFLite started")               '
   END IF                                                         '
   IF ISNOTNULL(lclDrive) THEN CHDRIVE lclDrive                   ' Switch drive if needed
   CHDIR lclPath                                                  ' put back the original path
   MExit
END SUB

FUNCTION sDOMacGet(MacName AS STRING) AS LONG
'---------- Get gCmdList() setup
LOCAL i, j, k AS LONG, Txt1 AS STRING
LOCAL CmdIO AS iIO                                                ' For our I/O stuff

   '----- OK, try and read it
   IF ISFALSE ISFILE(ENV.MacrosPath + MacName + ".DO") THEN       ' See if a file exists
      FUNCTION = 0: EXIT FUNCTION                                 ' No file, then return error
   END IF                                                         '
   LET CmdIO = CLASS "cIO"                                        ' Init for I/O
   CmdIO.Setup("IE", "", "", ENV.MacrosPath + MacName + ".DO")    ' Tell IO what we're opening
   CmdIO.EXEC                                                     '
   FILESCAN # CmdIO.FNum, RECORDS TO j                            ' Get # records in file
   REDIM gCmdList(1 TO j) AS GLOBAL STRING                        ' Redim gCmdList() to correct size
   k = 0                                                          '                                                    '
   DO WHILE ISFALSE EOF(CmdIO.FNum)                               ' Read the data
      LINE INPUT # CmdIO.FNum, Txt1                               ' Get a line
      INCR k                                                      ' Count it
      gCmdList(k) = Txt1                                          ' Store it in gCmdList()
   LOOP                                                           '
   CmdIO.Close()                                                  ' Close the file
   FUNCTION = j                                                   ' Pass back # of entries
END FUNCTION

FUNCTION sCmdOpType() AS LONG
'---------- Preprocess Cmd operands
LOCAL i, j AS LONG, t, p, dlm AS STRING
   MEntry
   IF pCmdNumOps = 0 THEN MExitFunc                               ' Exit if nothing to do
   FOR i = 1 TO pCmdNumOps                                        ' Something there, loop through them
      t = pCmdOps(i)                                              ' Get local copy
      pCmdRaw(i) = pCmdOps(i)                                     ' Save Raw form
      '----- Handle the [   ] bracketed operands
      IF IsNE(pCmdOps(0), "LINE") THEN                            ' Don't do this for LINE commands
         IF LEFT$(t, 1) = "[" THEN                                ' [ Operand
            pCmdOps(i) = MID$(t, 2)                               ' Copy remainder
            pCmdOpsType(i) = %OpSqb                               ' Set type
            ITERATE FOR                                           ' We're done
         END IF                                                   '
         IF RIGHT$(t, 1) = "]" THEN                               ' ] Operand
            pCmdOps(i) = LEFT$(t, LEN(t) - 1)                     ' Copy remainder
            pCmdOpsType(i) = %OpSqb                               ' Set type
            ITERATE FOR                                           ' We're done
         END IF                                                   '
      END IF                                                      '

      '----- Check for last char being a quote
      dlm = RIGHT$(t, 1)                                          ' Get last char
      IF dlm = $DQ OR dlm = CHR$(96) OR dlm = $SQ THEN            ' See if this is a quoted string
         IF LEFT$(t,1) = dlm THEN                                 ' See if a normal quoted type
            t = MID$(t, 2, LEN(t) - 2)                            ' Strip off quotes
            IF INSTR(t, dlm) THEN                                 ' Extraneous quotes?
               scError(%eFail, "Extraneous quotes detected in Operand #" + FORMAT$(i)) ' Issue error
               FUNCTION = %True: MExitFunc                        '
            END IF                                                '
            pCmdOps(i) = t                                        ' Save back without quotes
            pCmdOpsType(i) = %OpQStr                              ' Remember it was a quoted string
            ITERATE FOR                                           ' We're done this operand
         END IF                                                   '

         IF MID$(t,2,1) = dlm THEN                                ' See if quotes in 2nd and last position
            p = UUCASE(LEFT$(t,1))                                ' Pick up 1st character
            j = INSTR("CPRXTFME", p)                              ' A C"xx", P'xx', R'xx' X'xx' T'xx', F'xx', M'xx' or E'xx'  type?
            IF j THEN                                             ' If so
               t = MID$(t, 3, LEN(t) - 3)                         ' Strip to unquoted version
               IF INSTR(t, dlm) THEN                              ' Extraneous quotes?
                  scError(%eFail, "Extraneous quotes detected in Operand #" + FORMAT$(i)) ' Issue error
                  FUNCTION = %True: MExitFunc                     '
               END IF                                             '
               pCmdOps(i) = t                                     ' Save back without quotes
               pCmdOpsType(i) = CHOOSE(j, %OpCStr, %OpPStr, %OpRStr, %OpXStr, %OpTStr, %OpFStr, %OpMStr, %OpEStr) ' Remember its type
               ITERATE FOR                                        ' We're done this operand
            END IF                                                '
         END IF                                                   '
      END IF                                                      '

      '----- Now try the 2nd last position
      dlm = MID$(t, LEN(t) - 1, 1)                                ' Nor try the 'ABC'R format
      IF dlm = $DQ OR dlm = CHR$(96) OR dlm = $SQ THEN            ' See if this is a quoted string
         IF LEFT$(t, 1) = dlm THEN                                ' See if quotes in 1st position
            p = UUCASE(RIGHT$(t, 1))                              ' Pick up the last character
            j = INSTR("CPRXTFME", p)                              ' A C"xx", P'xx', R'xx', X'xx', T'xx', F'xx', M'xx' or E'xx' type?
            IF j THEN                                             ' Yes
               t = MID$(t, 2, LEN(t) - 3)                         ' Strip to unquoted version
               IF INSTR(t, dlm) THEN                              ' Extraneous quotes?
                  scError(%eFail, "Extraneous quotes detected in Operand #" + FORMAT$(i)) ' Issue error
                  FUNCTION = %True: MExitFunc                     '
               END IF                                             '
               pCmdOps(i) = t                                     ' Save back without quotes
               pCmdOpsType(i) = CHOOSE(j, %OpCStr, %OpPStr, %OpRStr, %OpXStr, %OpTStr, %OpFStr, %OpMStr, %OpEStr) ' Remember its type
               pCmdRaw(i) = p + "'" + t + "'"                     ' Fudge Raw back to orig format (P'zzz' processing needs it)
               ITERATE FOR                                        ' We're done this operand
            END IF                                                '
         END IF                                                   '

         scError(%eFail, "Extraneous quotes detected in Operand #" + FORMAT$(i)) ' Issue error
         FUNCTION = %True: MExitFunc                              '
      END IF                                                      '

      '----- See if a dotted operand (a label)
      IF LEFT$(t,1) = "." THEN                                    ' See if dotted (a label)
         pCmdOpsType(i) = %OpDotd                                 ' Remember its type
         ITERATE FOR                                              ' We're done this operand
      END IF                                                      '

      '----- See if an LPtr operand
      IF LEFT$(t,1) = "!" THEN                                    ' See if ! prefix
         IF VERIFY(MID$(t, 2), $Numeric) = 0 THEN                 ' Remainder numeric?
            pCmdOpsType(i) = %OpLPtr                              ' Remember its type
            ITERATE FOR                                           ' We're done this operand
         END IF                                                   '
      END IF                                                      '

      '----- See if a tag type operand
      IF LEFT$(t,1) = ":" THEN                                    ' See if : (a tag)
         IF LEN(t) > 8 THEN                                       ' Too long?
            scError(%eFail, "Operand #" + FORMAT$(i) + " is too long for a tag") ' Issue error
            FUNCTION = %True: MExitFunc                           '
         END IF                                                   '
         pCmdOpsType(i) = %OpTag                                  ' Remember its type
         ITERATE FOR                                              ' We're done this operand
      END IF                                                      '

      '----- See if a valid keyword
      j = Kwdt.KWLookup(t)                                        ' Can we find it as a KW?
      IF j THEN                                                   ' Got it
         pCmdOpsType(i) = j                                       ' Copy the unique KW code
         ITERATE FOR                                              ' We're done this operand
      END IF                                                      '

      '----- See if a simple numeric
      IF VERIFY(t, $Numeric) = 0 THEN                             ' All numeric?
         pCmdOpsType(i) = %OpNum                                  ' Remember its type
         ITERATE FOR                                              ' We're done this operand
      END IF                                                      '

      '----- Bitch if still got embedded quotes
      IF INSTR(pCmdOps(i), ANY $DQ+$SQ+CHR$(96)) THEN             ' Embedded quote?
         scError(%eFail, "Operand - " + pCmdOps(i) + " - contains embedded quotes") ' Issue error
         FUNCTION = %True: MExitFunc                              ' Exit
      END IF                                                      '

      '---- Mark it as a simple string
      pCmdOpsType(i) = %OpStr                                     ' Else normal string (for now)
   NEXT i                                                         '
   MExit
END FUNCTION

SUB      sCaretCreate()
'---------- Create caret based on current settings
   IF ISTRUE IsTPNSrtFlag THEN                                    ' Want an Insert mode cursor?
      IF ISTRUE ENV.VertInsCurs THEN                              ' Want a Vertical Insert cursor?
         CreateCaret(TP.gHandle, 0, 2, gFontHeight)               ' Create a vertical bar
      ELSE                                                        ' No, normal 'blob' cursor
         CreateCaret(TP.gHandle, 0, gFontWidth, INT((gFontHeight * ENV.CursInsert / 100))) ' Create an Insert blob
      END IF                                                      '
   ELSE                                                           ' No, we want an ordinary cursor
      CreateCaret(TP.gHandle, 0, gFontWidth, INT((gFontHeight * ENV.CursNormal / 100))) ' Create an Insert blob
   END IF                                                         '
END SUB

SUB      sCaretDestroy()
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL lclBG AS LONG
'---------- Destroy the Caret
   DestroyCaret()                                                 ' Destroy it
   gCaretCtr = 0                                                  ' Clear counter
END SUB

SUB      sCaretSet(lx AS LONG, ly AS LONG)
'---------- Set Caret to current cursor location
   IF ISFALSE ENV.VertInsCurs THEN                                ' If normal blob cursors
      IF ISFALSE IsTPNSrtFlag THEN                                ' And not insert mode
         SetCaretPos (lx, ly - INT((gFontHeight * ENV.CursNormal / 100))) ' Just set a normal position
      ELSE
         SetCaretPos (lx, ly - INT((gFontHeight * ENV.CursInsert / 100))) ' Just set a normal position
      END IF
   ELSE                                                           ' Maybe a vertical cursor
      IF ISFALSE IsTPNSrtFlag THEN                                ' Doing non insert cursor
         SetCaretPos (lx, ly - INT((gFontHeight * ENV.CursNormal / 100))) ' Just set a normal position
      ELSE                                                        ' Doing vertical Insert cursor
         SetCaretPos (lx - 1, ly - gFontHeight)                   ' Position for it
      END IF                                                      '
   END IF                                                         '
END SUB

SUB      sCaretShow()
'---------- Show the Caret
   IF gCaretCtr = 0 THEN                                          ' Only do it once
      ShowCaret(TP.gHandle)                                       ' Show the caret
      INCR gCaretCtr                                              ' Count it
   END IF                                                         '
END SUB

SUB      sCaretHide()
'---------- Hide the Caret
   DO WHILE gCaretCtr > 0                                         ' Only while ctr is > 0
      HideCaret(TP.gHandle)                                       ' Hide the caret
      DECR gCaretCtr                                              ' Reduce count
   LOOP                                                           '
END SUB

FUNCTION sCRTParse(valid AS STRING) AS LONG
'---------- Parse out the basic Criteria values
LOCAL ALL, Subset, Direct, Modifier, Exclude, Negative, CShift, Cols, L1, L2, lclTag, RetCode AS LONG
LOCAL LR, Pen, CPen, UserL, i, j, MaybeL1, FindChange, ClrCount1, ClrCount2 AS LONG
LOCAL ValidKW, Rx, IRx, T1, FCList, t AS STRING

   MEntry
   '----- See if one of the FIND/CHANGE command types, remember it
   FCList = "CHANGE  C       CHA     CHG     RCHANGE F       FIND    " + _
            "RFIND   NFIND   RNFIND  DELETE  DEL     NDELETE NDEL    " + _
            "EXCLUDE EX      EXC     X       NEXCLUDENX      FLIP    " + _
            "NFLIP   SHOW    NSHOW   FF      ULINE   REVERT  NULINE  " + _
            "NREVERT VV      UU      "
   t1 = UUCASE(LSET$(TP.CurrPCmd, 8))                             ' Make a temp fixed length field
   IF INSTR(FCList, t1) <> 0 THEN FindChange = %True              '

   '----- Set the basic operand group flags based on the validation string passed
   ALL      = INSTR(valid, $CRTAllOK)                             ' Setup local flags for allowable operands
   SubSet   = INSTR(valid, $CRTSubSet)                            '
   Direct   = INSTR(valid, $CRTDirect)                            '
   Modifier = INSTR(valid, $CRTModifier)                          '
   Exclude  = INSTR(valid, $CRTExclude)                           '
   Negative = INSTR(valid, $CRTNegative)                          '
   Cols     = INSTR(valid, $CRTCols)                              '
   CShift   = INSTR(valid, $CRTShift)                             '
   L1       = INSTR(valid, $CRTL1)                                '
   L2       = INSTR(valid, $CRTL2)                                '
   lclTag   = INSTR(valid, $CRTTag)                               '
   LR       = INSTR(valid, $CRTLR)                                '
   Pen      = INSTR(valid, $CRTPen)                               '
   CPen     = INSTR(valid, $CRTCPen)                              '
   UserL    = INSTR(valid, $CRTUser)                              '

   '----- Start by resetting everything, exit if nothing to do
   RESET CRTFlag, CRTFCol, CRTTCol, CRTL1, CRTL2, CRTL1RData, CRTL2RData
   CRTHiLiteClr = 0: CRTHiLiteOff = 0: CRTHiLiteOn = 0            '
   IF pCmdNumOps = 0 OR ISNULL(valid) THEN                        ' Nothing to process
      FUNCTION = %False: MExitFunc                                ' Exit if nothing to do
   END IF                                                         '

   '----- Build valid KW string based on the Operand group flags set above
   ValidKW += IIF$(ALL, CHR$(%KWTOP, %KWALL), "")                 ' Build list of valid KWs for this Parse
   ValidKW += IIF$(SubSet, CHR$(%KWEXCLUDE, %KWNX), "")           '
   ValidKW += IIF$(Direct, CHR$(%KWFIRST, %KWLAST, %KWPREV, %KWNEXT), "")
   ValidKW += IIF$(Modifier, CHR$(%KWWORD, %KWCHARS, %KWPREFIX, %KWSUFFIX, %KWLM, %KWRM), "")
   ValidKW += IIF$(Exclude, CHR$(%KWMX, %KWDX), "")               '
   ValidKW += IIF$(CShift, CHR$(%KWCS, %KWDS), "")                '
   ValidKW += IIF$(Negative, CHR$(%KWNF), "")                     '
   ValidKW += IIF$(lclTag, CHR$(%KWON, %KWOFF, %KWTOGGLE, %KWASSERT, %KWSET), "")
   ValidKW += IIF$(LR, CHR$(%KWLEFT, %KWRIGHT), "")               '
   ValidKW += IIF$(L2, CHR$(%KWTRUNC), "")                        '
   ValidKW += IIF$(Pen, CHR$(%KWSTD, %KWMSTD, %KWSOLID, %KWMSOLID), "")
   ValidKW += IIF$(CPen, CHR$(%KWPSTD), "")                       '
   ValidKW += IIF$(UserL, CHR$(%KWU, %KWNU), "")                  '

   '----- Now loop through the operands
   FOR i = 1 TO pCmdNumOps                                        ' Got some, loop through them
      IF INSTR(ValidKW, CHR$(pCmdOpsType(i))) = 0 THEN            ' It's not one of the valid keywords

         '----- It's not a KW and it IS numeric
         IF pCmdOpsType(i) = %OpNum THEN                          ' It's a number

            '----- Figure out if its a COL operand or maybe a Literal 1 operand
            IF Cols THEN                                          ' And Cols are allowed
               IF ISFALSE BIT(CrtFlag, %CrtFCol) THEN             ' Have we seen FCol yet?

                  IF VAL(pCmdOps(i)) < 1 THEN                     ' Catch zero or negative
                     scError(%eFail, "Invalid From Column operand - " + pCmdOps(i)) ' Issue error
                     FUNCTION = %True: MExitFunc                  '

                  ELSEIF VAL(pCmdOps(i)) > TP.MaxLength THEN      '

                     IF L1 AND ISFALSE BIT(CrtFlag, %CrtLit1) THEN' No L1 yet?
                        CrtL1 = pCmdOps(i)                        ' Set this as L1 then
                        CrtL1Raw = pCmdRaw(i)                     ' Save Raw form
                        CrtL1RData = pCmdRaw(i)                   ' Save Raw form
                        IF TP.PrfPCase = "C" THEN BIT SET CrtFlag, %CRTL1CaseComp   ' Set the comparison default
                        IF TP.PrfPCase = "T" THEN BIT SET CrtFlag, %CRTL1CaseInComp
                        BIT SET CRTFlag, %CRTLit1                 ' Remember we saw it
                        ITERATE FOR                               '
                     END IF                                       '
                  END IF
                  CrtFCol = VAL(pCmdOps(i))                       ' No? Set this as FCol then
                  MaybeL1 = i                                     ' Save Ops index in case Num Oper 1 fudge
                  BIT SET CRTFlag, %CRTFCol                       ' Remember we saw it
                  ITERATE FOR                                     '

               ELSEIF ISFALSE BIT(CrtFlag, %CrtTCol) THEN         ' Try TCol then

                  IF VAL(pCmdOps(i)) < CrtFCol THEN               '
                     scError(%eFail, "To Col: " + pCmdOps(i) + " is not >= From Col: " + FORMAT$(CrtFCol)) ' Issue error
                     FUNCTION = %True: MExitFunc                  '
                  END IF
                  CrtTCol = VAL(pCmdOps(i))                       ' No? Set this as TCol then
                  BIT SET CRTFlag, %CRTTCol                       ' Remember we saw it
                  ITERATE FOR                                     '

               ELSE                                               '

                  scError(%eFail, "Extra numeric operand detected - " + pCmdOps(i)) ' Issue error
                  FUNCTION = %True: MExitFunc                     '
               END IF                                             '
            END IF                                                '
         END IF                                                   '

         '----- Not Numeric, see if a Color operand
         IF LEFT$(pCmdOps(i), 1) <> "+" AND LEFT$(pCmdOps(i), 1) <> "-" THEN ' A possible 'normal' color name?
            ARRAY SCAN nHiLites(), COLLATE UCASE, = pCmdOps(i), TO j ' See if a HiLite Name
            IF j THEN                                             ' A winner
               IF CRTHiLiteClr <> 0 THEN                          ' Already have one?
                  scError(%eFail, "Extra color name operand detected - " + pCmdOps(i)) ' Issue error
                  FUNCTION = %True: MExitFunc                     '
               END IF                                             '
               CRTHiLiteClr = j                                   ' Save the color hilite index
               BIT SET CRTFlag, %CRTHiClr                         ' Remember we saw it
               ITERATE FOR                                        ' We're done
            END IF

         ELSE                                                     ' A +/- color operand
            ARRAY SCAN nHiLites(), COLLATE UCASE, = MID$(pCmdOps(i), 2), TO j ' See if remainder is a HiLite Name
            IF j THEN                                             ' A winner
               IF LEFT$(pCmdOps(i), 1) = "+" THEN                 ' The + variation?
                  IF CRTHiLiteOn <> 0 THEN                        ' Already have one?
                     scError(%eFail, "Extra +color name operand detected - " + pCmdOps(i)) ' Issue error
                     FUNCTION = %True: MExitFunc                  '
                  END IF                                          '
                  CRTHiLiteOn = j                                 ' Save the color hilite number
                  BIT SET CRTFlag, %CRTHiOn                       ' Remember we saw it
                  ITERATE FOR                                     ' We're done
               ELSE                                               ' Must be - (Off)
                  IF CRTHiLiteOff <> 0 THEN                       ' Already have one?
                     scError(%eFail, "Extra -color name operand detected - " + pCmdOps(i)) ' Issue error
                     FUNCTION = %True: MExitFunc                  '
                  END IF                                          '
                  CRTHiLiteOff = j                                ' Save the color hilite number
                  BIT SET CRTFlag, %CRTHiOff                      ' Remember we saw it
                  ITERATE FOR                                     ' We're done
               END IF                                             '
            END IF                                                '
         END IF                                                   '

         '----- Not a color operand, maybe literal 1
         IF L1 THEN                                               ' Literal1 allowed?
            IF ISFALSE BIT(CrtFlag, %CrtLit1) THEN                ' Have we seen Literal1 yet?
               IF ISNULL(pCmdOps(i)) THEN                         ' No null search strings (yet)
                  scError(%eFail, "Search literal may not be Null") ' Issue error
                  FUNCTION = %True: MExitFunc                     '
               END IF                                             '
               IF pCmdOpsType(i) = %OpPStr THEN                   ' Need to convert a Picture String?
                  CrtL1 = pCmdOps(i)                              ' Save string
                  CrtL1RData = pCmdOps(i)                         ' Swap in the converted string
                  CrtL1Raw = pCmdRaw(i)                           ' Save Raw form
                  BIT SET CRTFlag, %CRTLit1                       ' Remember we saw it
                  BIT SET CrtFlag, %CRTL1Picture                  ' Remember it was a Picture string

                  '----- Fudge the [ and ] characters
                  IF CrtL1RData = "[]" OR CrtL1RData = "[" OR CrtL1RData = "]"  OR _ ' Invalid string
                     CrtL1RData = "{}" OR CrtL1RData = "{" OR CrtL1RData = "}"  THEN '
                     scError(%eFail, CrtL1RData + " is an invalid Left/Right boundary literal") ' Issue error
                     FUNCTION = %True: MExitFunc                  '
                  END IF
                  IF LEFT$(CrtL1RData, 1) = "{" THEN BIT SET CrtFlag, %CrtLM  ' If we have a { modifier, set LM
                  IF RIGHT$(CrtL1RData, 1) = "}" THEN BIT SET CrtFlag, %CrtRM  ' If we have a } modifier, set RM

                  IF LEFT$(CrtL1RData, 1) = "[" THEN              ' If we have a [ modifier
                     CrtL1RData = MID$(CrtL1RData, 2)             ' Strip off the [
                     BIT SET CrtFlag, %CrtLM                      ' Set the LM flag
                  END IF                                          '
                  IF RIGHT$(CrtL1RData, 1) = "]" THEN             ' Got a ] ?
                     IF MID$(CrtL1RData, LEN(CrtL1RData) - 1, 1) <> "\" THEN ' If we have an unescaped ] modifier
                        CrtL1RData = LEFT$(CrtL1RData, LEN(CrtL1RData) - 1) ' Strip off the ]
                        BIT SET CrtFlag, %CrtRM                   ' Set the RM flag
                     END IF                                       '
                     j = INSTR(CrtL1RData, "\")                   ' Handle escaped chars
                     DO WHILE j                                   '
                        IF MID$(CrtL1RData, j + 1, 1) = "[" OR _  ' Trailing [ ]
                           MID$(CrtL1RData, j + 1, 1) = "]" THEN  '
                           CrtL1RData = LEFT$(CrtL1RData, j - 1) + MID$(CrtL1RData, j + 1)
                        END IF                                    '
                        j = INSTR(j + 1, CrtL1RData, "\")         ' Continue scan
                     LOOP                                         '
                  END IF                                          '

               ELSEIF pCmdOpsType(i) = %OpXStr THEN               ' Need to convert a Hex String?
                  iRx = pCmdOps(i)                                ' Copy to work field
                  GOSUB iRxHex                                    ' Go convert it
                  CrtL1 = pCmdOps(i)                              ' Save string
                  CrtL1RData = rx                                 ' Swap in the converted string
                  CrtL1Raw = pCmdRaw(i)                           ' Save Raw form
                  BIT SET CRTFlag, %CRTLit1                       ' Remember we saw it
                  BIT SET CrtFlag, %CRTL1Hex                      ' Remember it was a Hex string

               ELSEIF pCmdOpsType(i) = %OpRStr THEN               ' Need to handle a RegEx String?
                  t = PCRE_Regex_Compile(pCmdOps(i))              ' Let pcre_compile test the string
                  IF ISNOTNULL(t) THEN                            ' We get an error message back?
                     scError(%eFail, "Regex error " + t)          ' Issue error
                     FUNCTION = %True: MExitFunc                  '
                  END IF                                          '
                  CrtL1 = pCmdOps(i)                              ' Save string
                  CrtL1RData = pCmdOps(i)                         ' Twice
                  CrtL1Raw = pCmdRaw(i)                           ' Save Raw form
                  BIT SET CRTFlag, %CRTLit1                       ' Remember we saw it
                  BIT SET CrtFlag, %CRTL1RegEx                    ' Remember it was a RegEx

                  '----- Fudge the line start/end characters
                  IF LEFT$(CrtL1RData, 1) = "^" THEN              ' If we have a ^ modifier
                     BIT SET CrtFlag, %CrtLM                      ' Set the LM flag
                  END IF                                          '
                  IF RIGHT$(CrtL1RData, 1) = "$" THEN             ' Got a $ ?
                     IF MID$(CrtL1RData, LEN(CrtL1RData) - 1, 1) <> "\" THEN ' If we have an unescaped $ modifier
                        BIT SET CrtFlag, %CrtRM                   ' Set the RM flag
                     END IF                                       '
                  END IF                                          '

               ELSEIF pCmdOpsType(i) = %OpCStr THEN               ' C'...' type literal (Case sensitive)
                  CrtL1 = pCmdOps(i)                              ' Set this as L1 then
                  CrtL1Raw = pCmdRaw(i)                           ' Save Raw form
                  CrtL1RData = CrtL1                              ' Save also as raw dataRaw form
                  BIT SET CRTFlag, %CRTLit1                       ' Remember we saw it
                  BIT SET CrtFlag, %CRTL1CaseComp                 ' Remember it was a Case sensitive
               ELSEIF pCmdOpsType(i) = %OpTStr THEN               ' T'...' type literal (Case IN sensitive)
                  CrtL1 = pCmdOps(i)                              ' Set this as L1 then
                  CrtL1Raw = pCmdRaw(i)                           ' Save Raw form
                  CrtL1RData = CrtL1                              ' Save also as raw dataRaw form
                  BIT SET CRTFlag, %CRTLit1                       ' Remember we saw it
                  BIT SET CrtFlag, %CRTL1CaseInComp               ' Remember it was a Case in sensitive
               ELSEIF pCmdOpsType(i) = %OpFStr THEN               ' F'...' type literal (Format string)
                  scError(%eFail, "Search literal may not use F'xxx' type literal - " + pCmdOps(i)) ' Issue error
                  FUNCTION = %True: MExitFunc                     '
               ELSE                                               ' Everything else
                  CrtL1 = pCmdOps(i)                              ' Set this as L1 then
                  CrtL1Raw = pCmdRaw(i)                           ' Save Raw form
                  IF TP.PrfPCase = "C" THEN BIT SET CrtFlag, %CRTL1CaseComp   ' Set the comarison default
                  IF TP.PrfPCase = "T" THEN BIT SET CrtFlag, %CRTL1CaseInComp
                  BIT SET CRTFlag, %CRTLit1                       ' Remember we saw it
               END IF                                             '
               ITERATE FOR                                        '
            END IF                                                '
         END IF                                                   '

         '----- Wasn't Col or Literal 1, maybe make it literal 2
         IF L2 THEN                                               ' Literal2 allowed?
            IF ISFALSE BIT(CrtFlag, %CrtLit2) THEN                ' Try Literal2 then
               IF pCmdOpsType(i) = %OpPStr THEN                   ' Need to convert a Picture String?

                  CrtL2 = pCmdOps(i)                              ' Save string
                  CrtL2RData = pCmdOps(i)                         ' Swap in the converted string
                  CrtL2Raw = pCmdRaw(i)                           ' Save Raw form
                  BIT SET CRTFlag, %CRTLit2                       ' Remember we saw it
                  BIT SET CrtFlag, %CRTL2Picture                  ' Remember it was a Picture string

                  GOSUB FudgeSplit                                ' Fudge out the | split character
                  GOSUB FudgeUCLC                                 ' Fudge out the !< and !>  character
                  IF LEN(CrtL1RData) < TALLY(CrtL2RData, ANY "=<>~") THEN ' Picture literals have to be the same length
                     scError(%eFail, "CHANGE chars =<>~ greater than Find picture length") ' Issue error
                     FUNCTION = %True: MExitFunc                  '
                  END IF                                          '
                  IF LEN(CrtL2Rdata) > LEN(CrtL1RData) THEN       ' If change string longer than search string
                     IF INSTR(MID$(CrtL2RData, LEN(CrtL1RData) + 1), ANY "=<>~") THEN ' Special chars past the end?
                        scError(%eFail, "CHANGE chars =<>~ appear past the length of the Find string") ' Issue error
                        FUNCTION = %True: MExitFunc               '
                     END IF                                       '
                  END IF                                          '
                  IF ISTRUE INSTR(CrtL2, "{") AND ISFALSE INSTR(CrtL1, "{") THEN ' If change has {, then find must have {
                     scError(%eFail, "CHANGE string has {, FIND string must also have {") ' Issue error
                     FUNCTION = %True: MExitFunc                  '
                  END IF                                          '
                  IF ISTRUE INSTR(CrtL2, "}") AND ISFALSE INSTR(CrtL1, "}") THEN ' If change has }, then find must have }
                     scError(%eFail, "CHANGE string has }, FIND string must also have }") ' Issue error
                     FUNCTION = %True: MExitFunc                  '
                  END IF                                          '

               ELSEIF pCmdOpsType(i) = %OpFStr THEN               ' Need to convert a Format String?
                  IF LEN(CrtL1) < TALLY(pCmdOps(i), ANY "=<>~") THEN ' Picture literals have to be the same length
                     scError(%eFail, "CHANGE chars =<> greater than Find picture length") ' Issue error
                     FUNCTION = %True: MExitFunc                  '
                  END IF                                          '
                  CrtL2 = pCmdOps(i)                              ' Save string
                  CrtL2RData = pCmdOps(i)                         ' Swap in the converted string
                  CrtL2Raw = pCmdRaw(i)                           ' Save Raw form
                  BIT SET CRTFlag, %CRTLit2                       ' Remember we saw it
                  BIT SET CrtFlag, %CRTL2Format                   ' Remember it was a Picture string
                  GOSUB FudgeSplit                                ' Fudge out the | split character
                  GOSUB FudgeUCLC                                 ' Fudge out the !< and !>  character

               ELSEIF pCmdOpsType(i) = %OpMStr THEN               ' Need to convert a Mapping String?
                  CrtL2 = pCmdOps(i)                              ' Save string
                  CrtL2RData = pCmdOps(i)                         ' Swap in the converted string
                  CrtL2Raw = pCmdRaw(i)                           ' Save Raw form
                  BIT SET CRTFlag, %CRTLit2                       ' Remember we saw it
                  BIT SET CrtFlag, %CRTL2Map                      ' Remember it was a Map string

                  '----- Let MapStr do it's own validation
                  gMacFString = ""                                ' Setup null FIND string to indicate validate
                  gMacCString = ""                                ' Setup default answer to null
                  gMacErrString = "DIAG"                          ' Setup for a validate call

                  RetCode = mapstr_process(gMacFString, CRTL2, gMacCString, gMacErrString)

                  IF RetCode <> 0 THEN                            ' Better say OK
                     scError(%eFail, "CHANGE Map string error: " + gMacErrString) ' Issue error returned by MapStr
                     FUNCTION = %True: MExitFunc                  '
                  END IF                                          '


               ELSEIF pCmdOpsType(i) = %OpEStr THEN               ' Need to convert a Exec String?
                  CrtL2 = pCmdOps(i)                              ' Save string
                  CrtL2RData = pCmdOps(i)                         ' Swap in the converted string
                  CrtL2Raw = pCmdRaw(i)                           ' Save Raw form

                  '----- OK, validation will be done later after the loop
                  BIT SET CRTFlag, %CRTLit2                       ' Remember we saw it
                  BIT SET CrtFlag, %CRTL2Exec                     ' Remember it was an EXEC string

               ELSEIF pCmdOpsType(i) = %OpXStr THEN               ' Need to convert a Hex String?
                  iRx = pCmdOps(i)                                ' Copy to work field
                  GOSUB iRxHex                                    ' Go convert it
                  CrtL2 = pCmdOps(i)                              ' Save string
                  CrtL2RData = rx                                 ' Swap in the converted string
                  CrtL2Raw = pCmdRaw(i)                           ' Save Raw form
                  BIT SET CRTFlag, %CRTLit2                       ' Remember we saw it
                  BIT SET CrtFlag, %CRTL2Hex                      ' Remember it was a Hex string

               ELSEIF pCmdOpsType(i) = %OpCStr THEN               ' C'...' type literal (Case sensitive)
                  CrtL2 = pCmdOps(i)                              ' Set this as L2 then
                  CrtL2Raw = pCmdRaw(i)                           ' Save Raw form
                  BIT SET CRTFlag, %CRTLit2                       ' Remember we saw it
                  BIT SET CrtFlag, %CRTL2CaseComp                 ' Remember it was a Case sensitive

               ELSEIF pCmdOpsType(i) = %OpTStr THEN               ' T'...' type literal (Case IN sensitive)
                  CrtL2 = pCmdOps(i)                              ' Set this as L2 then
                  CrtL2Raw = pCmdRaw(i)                           ' Save Raw form
                  BIT SET CRTFlag, %CRTLit2                       ' Remember we saw it
                  BIT SET CrtFlag, %CRTL2CaseInComp               ' Remember it was a Case in sensitive

               ELSE                                               ' Everything else
                  CrtL2 = pCmdOps(i)                              ' Set this as L2 then
                  CrtL2Raw = pCmdRaw(i)                           ' Save Raw form
                  BIT SET CRTFlag, %CRTLit2                       ' Remember we saw it
               END IF                                             '
               ITERATE FOR                                        '
            END IF                                                '
         END IF                                                   '

         '----- If E"macname" operand present, validate it last else pCmd parsing data gets corrupted
         IF ISTRUE BIT(CrtFlag, %CrtL2Exec) THEN                  ' Try L2 EXEC type then
            t = sGetWord(CrtL2, %NoStrip, %QuoteNotSig)           ' Get the macroname

            '----- See if a valid macro
            IF ISFALSE ISFILE(ENV.MacrosPath + t + ".MACRO") THEN ' See if it exists
               scError(%eFail, "CHANGE EXEC macro: " + t + " does not exist") ' Issue error
               FUNCTION = %True: MExitFunc                        '
            END IF                                                '

            '----- Let macro do it's own validation
            gMacFString = ""                                      ' Setup null FIND string to indicate validate
            gMacCString = ""                                      ' Setup default answer to null
            pCmdMacro(CrtL2RData)                                 ' Let MACRO command handle it
            IF IsNE(gMacroMsg, "OK") THEN                         ' Better say OK
               scError(%eFail, "CHANGE EXEC macro: " + t + " " + IIF$(gMacCString = "", "failed validation", gMacCstring)) ' Issue error returned by macro
               FUNCTION = %True: MExitFunc                        '
            END IF                                                '
         END IF                                                   '

         '----- Didn't get chosen above, call it an error
         scError(%eFail, "Unknown operand detected - " + pCmdOps(i)) ' Issue error
         FUNCTION = %True: MExitFunc                              '

      '----- A recognizable KW was detected
      ELSE                                                        ' It's a recognizable KW

         '----- If it's one of ours, set it's flag, else reject it.
         SELECT CASE AS LONG pCmdOpsType(i)                       ' Split by KW
            CASE %KWTOP:     IF ISTRUE ALL THEN BIT SET CRTFlag, %CRTTop:        ITERATE FOR
            CASE %KWALL:     IF ISTRUE ALL THEN BIT SET CRTFlag, %CRTAll:        ITERATE FOR
            CASE %KWEXCLUDE: IF ISTRUE SubSet THEN BIT SET CRTFlag, %CRTX:       ITERATE FOR
            CASE %KWNX:      IF ISTRUE SubSet THEN BIT SET CRTFlag, %CRTNX:      ITERATE FOR
            CASE %KWFIRST:   IF ISTRUE Direct THEN BIT SET CRTFlag, %CRTFirst:   ITERATE FOR
            CASE %KWLAST:    IF ISTRUE Direct THEN BIT SET CRTFlag, %CRTLast:    ITERATE FOR
            CASE %KWPREV:    IF ISTRUE Direct THEN BIT SET CRTFlag, %CRTPrev:    ITERATE FOR
            CASE %KWNEXT:    IF ISTRUE Direct THEN BIT SET CRTFlag, %CRTNext:    ITERATE FOR
            CASE %KWWORD:    IF ISTRUE Modifier THEN BIT SET CRTFlag, %CRTWord:  ITERATE FOR
            CASE %KWCHARS:   IF ISTRUE Modifier THEN BIT SET CRTFlag, %CRTChars: ITERATE FOR
            CASE %KWPREFIX:  IF ISTRUE Modifier THEN BIT SET CRTFlag, %CRTPrefix:ITERATE FOR
            CASE %KWSUFFIX:  IF ISTRUE Modifier THEN BIT SET CRTFlag, %CRTSuffix:ITERATE FOR
            CASE %KWLM:      IF ISTRUE Modifier THEN BIT SET CRTFlag, %CRTLM:    ITERATE FOR
            CASE %KWRM:      IF ISTRUE Modifier THEN BIT SET CRTFlag, %CRTRM:    ITERATE FOR
            CASE %KWMX:      IF ISTRUE Exclude THEN BIT SET CRTFlag, %CRTMX:     ITERATE FOR
            CASE %KWDX:      IF ISTRUE Exclude THEN BIT SET CRTFlag, %CRTDX:     ITERATE FOR
            CASE %KWCS:      IF ISTRUE CShift THEN BIT SET CRTFlag, %CRTCS:      ITERATE FOR
            CASE %KWDS:      IF ISTRUE CShift THEN BIT SET CRTFlag, %CRTDS:      ITERATE FOR
            CASE %KWLEFT:    IF ISTRUE LR THEN BIT SET CRTFlag, %CRTLEFT:        ITERATE FOR
            CASE %KWRIGHT:   IF ISTRUE LR THEN BIT SET CRTFlag, %CRTRIGHT:       ITERATE FOR
            CASE %KWNF:      IF ISTRUE Negative THEN BIT SET CRTFlag, %CRTNF:    ITERATE FOR
            CASE %KWON:      IF ISTRUE lclTag THEN BIT SET CRTFlag, %CRTON:      ITERATE FOR
            CASE %KWOFF:     IF ISTRUE lclTag THEN BIT SET CRTFlag, %CRTOFF:     ITERATE FOR
            CASE %KWTOGGLE:  IF ISTRUE lclTag THEN BIT SET CRTFlag, %CRTTOGGLE:  ITERATE FOR
            CASE %KWASSERT:  IF ISTRUE lclTag THEN BIT SET CRTFlag, %CRTASSERT:  ITERATE FOR
            CASE %KWSET:     IF ISTRUE lclTag THEN BIT SET CRTFlag, %CRTSET:     ITERATE FOR
            CASE %KWTRUNC:   IF ISTRUE L2  THEN BIT SET CRTFlag, %CRTL2TRUNC:    ITERATE FOR
            CASE %KWSTD:     IF ISTRUE Pen THEN BIT SET CRTFlag, %CRTStd:        ITERATE FOR
            CASE %KWMSTD:    IF ISTRUE Pen THEN BIT SET CRTFlag, %CRTMStd:       ITERATE FOR
            CASE %KWPSTD:    IF ISTRUE Pen THEN BIT SET CRTFlag, %CRTPStd:       ITERATE FOR
            CASE %KWSOLID:   IF ISTRUE Pen THEN BIT SET CRTFlag, %CRTSolid:      ITERATE FOR
            CASE %KWMSOLID:  IF ISTRUE Pen THEN BIT SET CRTFlag, %CRTMSolid:     ITERATE FOR
            CASE %KWU:       IF ISTRUE UserL THEN BIT SET CRTFlag, %CRTU:        ITERATE FOR
            CASE %KWNU:      IF ISTRUE UserL THEN BIT SET CRTFlag, %CRTNU:       ITERATE FOR
            CASE ELSE
               scError(%eFail, "Keyword not allowed on the " + pCmdOps(0) + " command - " + pCmdOps(1)) ' Issue error
               FUNCTION = %True: MExitFunc                        '
         END SELECT                                               '
      END IF                                                      '
   NEXT i                                                         '

   '----- Now check for Conflicts between KW and other operands
   IF BIT(CrtFlag, %CRTCS) AND BIT(CrtFlag, %CRTDS) THEN _
      scError(%eFail, "Both CS and DS not allowed on the " + pCmdOps(0) + " command."): FUNCTION = %True: MExitFunc

   IF BIT(CrtFlag, %CRTX) AND BIT(CrtFlag, %CRTNX) THEN _
      scError(%eFail, "Both X and NX not allowed on the " + pCmdOps(0) + " command."): FUNCTION = %True: MExitFunc

   IF BIT(CrtFlag, %CRTU) AND BIT(CrtFlag, %CRTNU) THEN _
      scError(%eFail, "Both U and NU not allowed on the " + pCmdOps(0) + " command."): FUNCTION = %True: MExitFunc

   IF BIT(CrtFlag, %CRTFirst) AND (BIT(CrtFlag, %CRTLast) OR BIT(CrtFlag, %CRTNext) OR BIT(CrtFlag, %CRTPrev)) THEN _
      scError(%eFail, "FIRST not allowed with LAST, NEXT or PREV"): FUNCTION = %True: MExitFunc

   IF BIT(CrtFlag, %CRTLast) AND (BIT(CrtFlag, %CRTNext) OR BIT(CrtFlag, %CRTPrev)) THEN _
      scError(%eFail, "LAST not allowed with NEXT or PREV"): FUNCTION = %True: MExitFunc

   IF BIT(CrtFlag, %CRTNext) AND BIT(CrtFlag, %CRTPrev) THEN _
      scError(%eFail, "NEXT not allowed with PREV"): FUNCTION = %True: MExitFunc

   IF BIT(CrtFlag, %CRTWord) AND (BIT(CrtFlag, %CRTPrefix) OR BIT(CrtFlag, %CRTSuffix)) THEN _
      scError(%eFail, "WORD not allowed with PREFIX or SUFFIX"): FUNCTION = %True: MExitFunc

   IF BIT(CrtFlag, %CRTWord) AND BIT(CrtFlag, %CRTChars) THEN _
      scError(%eFail, "WORD not allowed with CHARS"): FUNCTION = %True: MExitFunc

   IF BIT(CrtFlag, %CRTPrefix) AND BIT(CrtFlag, %CRTSuffix) THEN _
      scError(%eFail, "PREFIX not allowed with SUFFIX"): FUNCTION = %True: MExitFunc

   IF BIT(CrtFlag, %CRTLeft) AND BIT(CrtFlag, %CRTRight) THEN _
      scError(%eFail, "Both RIGHT and LEFT not allowed"): FUNCTION = %True: MExitFunc

   IF BIT(CrtFlag, %CRTMX) AND BIT(CrtFlag, %CRTDX) THEN _
      scError(%eFail, "Both MX and DX not allowed on the " + pCmdOps(0) + " command."): FUNCTION = %True: MExitFunc

   IF (BIT(CrtFlag, %CRTDX) OR BIT(CrtFlag, %CRTMX)) AND IsTPHideFlag AND ISFALSE BIT(CrtFlag, %CRTALL) THEN _
      scError(%eFail, "HIDE mode and DX/MX require ALL keyword as well"): FUNCTION = %True: MExitFunc

   IF (BIT(CrtFlag, %CRTLeft) OR BIT(CrtFlag, %CRTRight)) AND BIT(CrtFlag, %CRTL1RegEx) THEN _
      scError(%eFail, "LEFT/RIGHT cannot be used with RegEX literals"): FUNCTION = %True: MExitFunc

   IF BIT(CrtFlag, %CRTLM) AND (BIT(CrtFlag, %CRTFCol) OR BIT(CrtFlag, %CRTTCol)) THEN _
      scError(%eFail, "LM or [ Left bound in Picture cannot be used with column operands"): FUNCTION = %True: MExitFunc

   IF BIT(CrtFlag, %CRTRM) AND (BIT(CrtFlag, %CRTFCol) OR BIT(CrtFlag, %CRTTCol)) THEN _
      scError(%eFail, "RM or ] Right bound in Picture cannot be used with column operands"): FUNCTION = %True: MExitFunc

   IF ISFALSE BIT(CrtFlag, %CRTWord) AND ISFALSE BIT(CrtFlag, %CRTChars) AND _ ' Apply global default
      ISFALSE BIT(CrtFlag, %CRTPrefix) AND ISFALSE BIT(CrtFlag, %CRTSuffix) AND _
      ISTRUE FindChange AND ISTRUE BIT(CrtFlag, %CrtLit1) THEN    ' Have we seen Literal1 and no Word/Chars default?
      IF TP.FindWord THEN BIT SET CRTFlag, %CRTWord               ' Set WORD if that's the default
   END IF                                                         '

   '----- Verify all the color operands
   IF BIT(CrtFlag, %CRTPSTD) THEN INCR ClrCount1                  '
   IF BIT(CrtFlag, %CRTHiOn) THEN INCR ClrCount1                  '
   IF ClrCount1 > 1 THEN _
      scError(%eFail, "Cannot specify multiple change +color values"): FUNCTION = %True: MExitFunc

   ClrCount1 = 0                                                  ' Reset counter
   IF BIT(CrtFlag, %CRTPSTD) THEN INCR ClrCount1                  '
   IF BIT(CrtFlag, %CRTHiOn) THEN INCR ClrCount1                  '

   IF BIT(CrtFlag, %CRTMSTD) THEN INCR ClrCount2                  '
   IF BIT(CrtFlag, %CRTHiOff) THEN INCR ClrCount2                 '

   IF ClrCount1 > 0 AND ClrCount2 > 0 THEN _                      ' Do we have both positive and negative?
      scError(%eFail, "Cannot specify both positive and negative color changes"): FUNCTION = %True: MExitFunc

   '----- Kludge for a single numeric operand
   IF L1 AND ISFALSE L2 THEN                                      ' L1 allowed but not L2, Col1 got but not Col2
      IF BIT(CrtFlag, %CRTLit1) = 0 AND ISTRUE BIT(CrtFlag, %CRTFCol) AND ISFALSE BIT(CrtFlag, %CRTTCol) THEN
         BIT RESET CRTFlag, %CRTFCol: BIT SET CRTFlag, %CRTLit1   ' Adjust the bit flags
         CRTL1 = pCmdOps(MaybeL1): CRTL1Raw = CRTL1: CRTL1RData = CRTL1
      END IF                                                      '
   END IF                                                         '

   '----- See if maybe global BNDS should be applied
   IF Cols AND BIT(CrtFlag, %CrtLit1) THEN                        ' If Cols are allowed and we have a literal
      IF ISFALSE BIT(CrtFlag, %CRTFCol) THEN                      ' And no column operand specified
         IF TP.PrfBndLeft > 1 OR TP.PrfBndRight > 0 THEN          ' And Global BNDS are active
            BIT SET CrtFlag, %CrtFCol: BIT SET CrtFlag, %CrtTCol  ' Inherit the global bounds
            CrtFCol = TP.PrfBndLeft: CrtTCol = TP.PrfBndRight     '
         END IF                                                   '
      END IF                                                      '
   END IF                                                         '
   MExitFunc

'----- Process a possible HEX operand
iRxHex:
   iRx = UUCASE(iRx)                                              ' Uppercase it
   IF (LEN(iRx) MOD 2) = 1 THEN _                                 ' Better be even # chars
      scError(%eFail, "Invalid Hex literal length"): FUNCTION = %True: MExitFunc
   IF VERIFY(iRx, $Hex) <> 0 THEN _                               ' All valid Hex chars?
      scError(%eFail, "Restricted character in Hex literal"): FUNCTION = %True: MExitFunc
   rx = ""                                                        ' Convert to hex
   FOR j = 1 TO LEN(iRx) STEP 2                                   '
      rx += CHR$(VAL("&H" + MID$(iRx, j, 2)))                     '
   NEXT j                                                         '
   IF TP.PrfSrceXlate THEN TP.Translate(rx, TP.TPPrfGetSS2APtr)   ' Translate SOURCE to ANSI
   RETURN                                                         '

FudgeSplit:
   '----- Fudge the | split character
   IF FindChange THEN RETURN                                      ' Ignore for normal F/C operands
   j = INSTR(CrtL2RData, "|")                                     ' Any | characters?
   DO WHILE j                                                     ' Process the | characters
      IF j = 1 OR MID$(CrtL2RData, j - 1, 1) <> "\" THEN          ' We have an unescaped | char.
         IF TP.cfSplitPt2 <> 0 THEN                               ' Oops, already have a Split Point
            scError(%eFail, "Multiple | SPLIT points present")    ' Issue error
            FUNCTION = %True: MExitFunc                           '
         END IF                                                   '
         TP.cfSplitPt2 = j                                        ' Save it
         CrtL2RData = LEFT$(CrtL2RData, j - 1) + MID$(CrtL2RData, j + 1) ' Shrink the literal
      END IF                                                      '
      j = INSTR(j + 1, CrtL2RData, "|")                           ' Any MORE | characters?
   LOOP                                                           '
   j = INSTR(CrtL2RData, "\")                                     ' Handle escaped chars
   DO WHILE j                                                     '
      IF MID$(CrtL2RData, j + 1, 1) = "|" THEN                    ' Trailing |
         CrtL2RData = LEFT$(CrtL2RData, j - 1) + MID$(CrtL2RData, j + 1)
         IF j < TP.cfSplitPt2 THEN TP.cfSplitPt2 = TP.cfSplitPt2 - 1
      END IF                                                      '
      j = INSTR(j + 1, CrtL2RData, "\")                           ' Continue scan
   LOOP                                                           '
   RETURN

FudgeUCLC:
   '----- Fudge the UC/LC escaped versions into « » characters

   j = INSTR(CrtL2RData, "!")                                     ' Any ! characters?
   DO WHILE j                                                     ' Process the ! characters
      IF j = 1 OR MID$(CrtL2RData, j - 1, 1) <> "\" THEN          ' We have an unescaped ! char.
         IF MID$(CrtL2RData, j + 1, 1) = "<" THEN _               ' A lowercase request
            CrtL2RData = LEFT$(CrtL2RData, j - 1) + "«" + MID$(CrtL2RData, j + 2) ' Shrink the literal
         IF MID$(CrtL2RData, j + 1, 1) = ">" THEN _               ' An uppercase request
            CrtL2RData = LEFT$(CrtL2RData, j - 1) + "»" + MID$(CrtL2RData, j + 2) ' Shrink the literal
         IF MID$(CrtL2RData, j + 1, 1) = "=" THEN _               ' A sort-of escaped version
            CrtL2RData = LEFT$(CrtL2RData, j) + MID$(CrtL2RData, j + 2) ' Shrink the literal
      END IF                                                      '
      j = INSTR(j + 1, CrtL2RData, "!")                           ' Any MORE ! characters?
   LOOP                                                           '
   RETURN
END FUNCTION

FUNCTION sDate() AS STRING
'---------- Return Date in Windows format
LOCAL st AS SYSTEMTIME
LOCAL szDate AS ASCIIZ * 64, Locale, Format AS LONG
LOCAL szFormat AS ASCIIZ * 2, rc AS LONG
LOCAL szSep AS ASCIIZ * 2, sep, MyDate AS STRING
   MEntry
   GetLocalTime st
   GetDateFormat %LOCALE_SYSTEM_DEFAULT, %DATE_SHORTDATE, st, BYVAL %NULL, szDate, SIZEOF(szDate)
   rc = GetLocaleInfo(Locale, %LOCALE_IDATE,  szFormat, SIZEOF(szFormat))
   Format = VAL(szFormat)
   rc = GetLocaleInfo(Locale, %LOCALE_SDATE, szSep, SIZEOF(szSep))
   sep = IIF$(rc = 0, "/", szSep)
   MyDate = DATE$
   REPLACE "-" WITH sep IN MyDate
   SELECT CASE AS LONG format
      CASE 0: FUNCTION = MyDate                                   ' mmddyyyy
      CASE 1: FUNCTION = MID$(MyDate, 4, 3) & MID$(MyDate, 1, 3) & MID$(MyDate, 7) ' ddmmyyyy
      CASE 2: FUNCTION = MID$(MyDate, 7) & sep & MID$(MyDate, 4, 3) & MID$(MyDate, 1, 2) ' yyyymmdd
  END SELECT
  MExit
END FUNCTION

SUB      sDoBeep()
'---------- Issue BEEP if allowed
REGISTER lx AS LONG
REGISTER i AS LONG
   MEntry
   IF ISTRUE ENV.ABeepFlag THEN PlaySound("SystemAsterisk", %NULL, %SND_ASYNC)
   IF ISFALSE ENV.VBeepFlag OR ISTRUE gMacroMode THEN MExitSub    ' Exit if no VBeep
   lx = (7 * gFontWidth) + 1                                      ' Get width of "Command"
   GRAPHIC ATTACH TP.PgHandle, TP.WindowID                        ' Set as the default graphic area
   GRAPHIC SET MIX %MIX_NOT                                       ' Set MIX
   FOR i = 1 TO 4                                                 ' Invert it 4 times
      GRAPHIC BOX (2, 2) - (lx, gFontHeight), 0, cTxtHiFG, cTxtHiFG '
      SLEEP 80                                                    ' Delay between flips
   NEXT i                                                         '
   MExit
END SUB

SUB      sDoCursor ()
'---------- Handle the cursor on the graphic screen
LOCAL lx, ly AS LONG
LOCAL lr, lRow, lclCol, i, j, cHeight, lclBG, DidErase, DoRedraw AS LONG, char AS STRING
LOCAL cLoc, Stat AS STRING
STATIC px, py AS INTEGER
   MEntry
   IF gMacroMode OR gfTermFlag THEN EXIT SUB
   DidErase = %False: DoRedraw = %False
   IF TabsNum = 0 THEN MexitSub                                   ' Prevent Memory Access during shutdown
   lRow = TP.CsrRow: lclCol = TP.CsrCol                           ' Get local copies
   sDoStatusBar($SBMode+$SBLinNo+$SBLines+$SBCols+$SBMisc+$SBEOL+$SBState) ' re-Do some StatusBar boxes

   '---------- Finally do the actual cursor
   lx = (lclCol - 1) * gFontWidth + %GLM                          ' Include LM pad
   ly = lRow  * gFontHeight + 1
   sCaretSet(lx, ly)                                              ' Go set the caret and show it

   '----- Do other 'cursor' if in PTYPE mode
   IF ISTRUE IsTPPTypeMode THEN                                   ' In PTYPE mode?
      GRAPHIC ATTACH TP.PgHandle, TP.WindowID, REDRAW             ' Set as the default graphic area
      DoRedraw = %True                                            ' Do the redraw
      GRAPHIC SET MIX %MIX_COPYSRC                                '
      IF TP.LastPTCurs <> 0 AND TP.LastPTCurs <> TP.CsrCol THEN   ' Previous to erase AND a new column?
         i = TP.LastPTCurs                                        ' Working copy
         IF ISFALSE ENV.Banding THEN                              '
            GRAPHIC LINE ((i - 1) * gFontWidth + %GLM - 1, (TP.PTFDisp) * gFontHeight) - ((i - 1) * gFontWidth + %GLM - 1, TP.PTLDisp * gFontHeight), cTxtLoBG1 ' Draw the line
            GRAPHIC LINE (i * gFontWidth + %GLM - 1, (TP.PTFDisp) * gFontHeight) - (i * gFontWidth + %GLM - 1, TP.PTLDisp * gFontHeight), cTxtLoBG1 ' Draw the line
         ELSE                                                     ' Much harder now, we're in Banding mode
            FOR j = TP.PTFDisp + 1 TO TP.PTLDisp                  '
               GRAPHIC LINE ((i - 1) * gFontWidth + %GLM - 1, ((j - 1) * gFontHeight) + 1) - ((i - 1) * gFontWidth + %GLM - 1, (j * gFontHeight) + 1), cTxtLoBG1 ' Draw the line
               GRAPHIC LINE (i * gFontWidth + %GLM - 1, (j - 1) * gFontHeight) - (i * gFontWidth + %GLM - 1, j * gFontHeight), cTxtLoBG1 ' Draw the line
            NEXT j                                                '
         END IF                                                   '
         TP.LastPTCurs = 0                                        '
      END IF                                                      '
      IF TP.LastPTCurs <> TP.CsrCol THEN                          ' New cursor to draw?
         i = TP.CsrCol                                            ' Get cursor column
         TP.LastPTCurs = i                                        ' Save it
         GRAPHIC LINE ((i - 1) * gFontWidth + %GLM - 1, (TP.PTFDisp) * gFontHeight) - ((i - 1) * gFontWidth + %GLM - 1, TP.PTLDisp * gFontHeight), ENV.cMarkLine ' Draw the line
         GRAPHIC LINE (i * gFontWidth + %GLM - 1, (TP.PTFDisp) * gFontHeight) - (i * gFontWidth + %GLM - 1, TP.PTLDisp * gFontHeight), ENV.cMarkLine ' Draw the line
      END IF                                                      '

   ELSEIF ISTRUE ENV.HRuler OR ISTRUE ENV.VRuler THEN             ' Vertical or horizontal ruler mode?
      GRAPHIC ATTACH TP.PgHandle, TP.WindowID, REDRAW             ' Set as the default graphic area
      DoRedraw = %True                                            ' Do the redraw
      GRAPHIC SET MIX %MIX_COPYSRC                                '

      '----- Split now for FM / Non-FM
      IF ISFALSE IsFMTab THEN                                     ' The Edit tab?

         '----- Erase previous lines if needed
         IF ISTRUE ENV.VRuler THEN                                ' Vertical?
            IF TP.LastRulCol <> 0 AND TP.LastRulCol <> TP.CsrCol THEN ' Previous to erase?
               i = TP.LastRulCol                                  ' Working copy
               IF ISTRUE TP.PrfPMark AND MID$(TP.PrfMarkWorking, i + TP.Offset - gLNPadCol, 1) = "*" THEN ' MARKing and also a MARK column?
                  GRAPHIC LINE ((i - 1) * gFontWidth + %GLM - 1, 1) - ((i - 1) * gFontWidth + %GLM - 1, (2 + TP.PrfCols) * gFontHeight), cTxtLoBG1 ' Draw the line
               ELSE                                               ' Else do a full line erase
                  DidErase = %True                                '
                  IF ISFALSE ENV.Banding THEN                     '
                     GRAPHIC LINE ((i - 1) * gFontWidth + %GLM - 1, 1) - ((i - 1) * gFontWidth + %GLM - 1, (ENV.ScrHeight - ENV.PFKShow) * gFontHeight), cTxtLoBG1 ' Draw the line
                  ELSE                                            ' Much harder now, we're in Banding mode
                     FOR j = 1 TO (ENV.ScrHeight - ENV.PFKShow)   '
                        sCalcEditBG(j)                            ' Calc Banding
                        lclBG = IIF(cBandBG, ENV.GetClr(%SCTxtLo, %SCBG2), ENV.GetClr(%SCTxtLo, %SCBG1))  ' Chose Scheme's BG color
                        GRAPHIC LINE ((i - 1) * gFontWidth + %GLM - 1, ((j - 1) * gFontHeight) + 1) - ((i - 1) * gFontWidth + %GLM - 1, (j * gFontHeight) + 1), lclBG ' Draw the line
                     NEXT j                                       '
                  END IF                                          '
               END IF                                             '
               TP.LastRulCol = 0                                  '
            END IF                                                '
         END IF                                                   '

         IF ISTRUE ENV.HRuler THEN                                ' Horizontal?
            IF 1 = 1 THEN                                         '
               IF TP.LastRulRow <> 0 AND TP.LastRulRow <> TP.CsrRow THEN ' Previous to erase AND a new row?
                  i = TP.LastRulRow                               ' Working copy
                  sCalcEditBG(i)                                  ' Calc Banding
                  lclBG = IIF(cBandBG, ENV.GetClr(%SCTxtLo, %SCBG2), ENV.GetClr(%SCTxtLo, %SCBG1))  ' Chose Scheme's BG color
                  DidErase = %True                                '
                  GRAPHIC LINE (1, i * gFontHeight) - (ENV.ScrWidth * gFontWidth + %GLM, i * gFontHeight), lclBG ' Draw the line
                  TP.LastRulRow = 0                               '
                  IF TP.PrfHexMode = &1 THEN                      ' If not HEX mode, do some more
                     IF sGetIX(j) > 0 THEN                        ' If a text line
                        TP.DispLine(sGetix(j), j)                 ' Re disp it
                     END IF                                       '
                  END IF                                          '
                  TP.DoMarkLines                                  ' Redraw the MARK lines
               END IF                                             '
            END IF                                                '
         END IF                                                   '

         IF ISTRUE ENV.VRuler THEN                                ' Vertical?
            IF (1 = 1  AND TP.CsrCol <> TP.LastRulCol) OR DidErase THEN
               i = TP.CsrCol                                      ' Get cursor column
               TP.LastRulCol = i                                  ' Save it
               GRAPHIC LINE ((i - 1) * gFontWidth + %GLM - 1, 1) - ((i - 1) * gFontWidth + %GLM - 1, (ENV.ScrHeight - ENV.PFKShow) * gFontHeight), ENV.cMarkLine ' Draw the line
            END IF                                                '
         END IF                                                   '

         IF ISTRUE ENV.HRuler THEN                                ' Horizontal?
            IF (1 = 1  AND TP.CsrRow <> TP.LastRulRow) OR DidErase THEN
               i = TP.CsrRow                                      ' Get cursor row
               TP.LastRulRow = i                                  ' Save it
               GRAPHIC LINE (1, i * gFontHeight) - (ENV.ScrWidth * gFontWidth + %GLM, i * gFontHeight), ENV.cMarkLine ' Draw the line
            END IF                                                '
         END IF                                                   '

      '----- Do the FM variety
      ELSE
         '----- Erase previous lines if needed
         IF ISTRUE ENV.VRuler THEN                                ' Vertical?
            IF TP.LastRulCol <> 0 AND TP.LastRulCol <> TP.CsrCol THEN ' Previous to erase?
               i = TP.LastRulCol                                  ' Working copy
               DidErase = %True                                   '
               FOR j = 1 TO (ENV.ScrHeight - IIF(ENV.FMHelpFlag, 3, 0))   '
                  sCalcFMBG(j)                                    ' Calc Banding
                  lclBG = IIF(cBandBG, ENV.GetClr(%SCTxtLo, %SCBG2), ENV.GetClr(%SCTxtLo, %SCBG1))  ' Chose Scheme's BG color
                  GRAPHIC LINE ((i - 1) * gFontWidth + %GLM - 1, ((j - 1) * gFontHeight)) - ((i - 1) * gFontWidth + %GLM - 1, (j * gFontHeight)), IIF(j = 3 OR j = 6, cFMToolBG1, lclBG) ' Draw the line
               NEXT j                                             '
               TP.LastRulCol = 0                                  '
            END IF                                                '
         END IF                                                   '

         IF ISTRUE ENV.HRuler THEN                                ' Horizontal?
            IF 1 = 1 THEN                                         '
               IF TP.LastRulRow <> 0 AND TP.LastRulRow <> TP.CsrRow THEN ' Previous to erase AND a new row?
                  i = TP.LastRulRow                               ' Working copy
                  sCalcFMBG(i)                                    ' Calc Banding
                  lclBG = IIF(cBandBG, ENV.GetClr(%SCTxtLo, %SCBG2), ENV.GetClr(%SCTxtLo, %SCBG1))  ' Chose Scheme's BG color
                  DidErase = %True                                '
                  GRAPHIC LINE (1, i * gFontHeight) - (ENV.ScrWidth * gFontWidth + %GLM, i * gFontHeight), lclBG ' Draw the line
                  TP.LastRulRow = 0                               '
               END IF                                             '
            END IF                                                '
         END IF                                                   '

         IF ISTRUE ENV.VRuler THEN                                ' Vertical?
            IF (1 = 1  AND TP.CsrCol <> TP.LastRulCol) OR DidErase THEN
               i = TP.CsrCol                                      ' Get cursor column
               TP.LastRulCol = i                                  ' Save it
               GRAPHIC LINE ((i - 1) * gFontWidth + %GLM - 1, 1) - ((i - 1) * gFontWidth + %GLM - 1, (ENV.ScrHeight - IIF(ENV.FMHelpFlag, 3, 0)) * gFontHeight), ENV.cMarkLine ' Draw the line
            END IF                                                '
         END IF                                                   '

         IF ISTRUE ENV.HRuler THEN                                ' Horizontal?
            IF (1 = 1  AND TP.CsrRow <> TP.LastRulRow) OR DidErase THEN
               i = TP.CsrRow                                      ' Get cursor row
               TP.LastRulRow = i                                  ' Save it
               GRAPHIC LINE (1, i * gFontHeight) - (ENV.ScrWidth * gFontWidth + %GLM, i * gFontHeight), ENV.cMarkLine ' Draw the line
            END IF                                                '
         END IF                                                   '
      END IF
   END IF                                                         '
   IF DoRedraw THEN GRAPHIC REDRAW                                ' Redraw it if we need to
   MExitSub
END SUB

SUB     sDoENDAll(pCmd AS STRING)
'---------- Terminate all the tabs
LOCAL i, j, cTab, noreopen AS LONG, shutcmd, mnames, mrf, TabType, MSG AS STRING
   MEntry
   ON ERROR GOTO Whoops                                           '
   IF TP.CmdParse(pCmd) THEN GOTO Exit2                           ' Do basic parsing, exit if errors
   FOR i = 1 TO pCmdNumOps                                        ' Scan operands
      SELECT CASE AS CONST$ UUCASE(pCmdOps(i))                    '
         CASE "NOREOPEN"      : noreopen = %True                  ' NOREOPEN
         CASE "END"                                               ' END
            IF shutcmd = "" THEN
               shutcmd = "END "                                   '
            ELSE                                                  ' Oops
               ScError(nMac(%eFail), "Multiple END/CANCEL operands"): MExitSub
            END IF                                                '
         CASE "CANCEL", "CAN"                                     ' CANCEL
            IF shutcmd = "" THEN                                  '
               shutcmd = "CANCEL "                                '
            ELSE                                                  ' Oops
               ScError(nMac(%eFail), "Multiple END/CANCEL operands"): MExitSub
            END IF                                                '
         CASE "PURGE", "PUR"                                      ' PURGE
            IF shutcmd = "CANCEL " THEN                           '
               shutcmd += "PURGE "                                '
            ELSE                                                  '
               ScError(nMac(%eFail), "PURGE allowed only with CANCEL"): MExitSub
            END IF                                                '
         CASE "DELETE", "DEL"                                     ' DELETE
            IF shutcmd = "CANCEL " THEN                           '
               shutcmd += "DELETE "                               '
            ELSE                                                  '
               ScError(nMac(%eFail), "DELETE allowed only with CANCEL"): MExitSub
            END IF                                                '
         CASE ELSE
            ScError(nMac(%eFail), "Unknown EXIT operand: " + pCmdOps(i)): MExitSub
      END SELECT                                                  '
   NEXT i                                                         '
   IF shutcmd = "" THEN shutcmd = "END "                          ' Provide default if no operands

   gfEndAll = %True                                               ' Remember us in global memory
   cTab = TP.PgNumber                                             ' Remember active tab
   FOR i = TabsNum TO 1 STEP -1                                   ' Do for each tab
      TP = Tabs(i)                                                ' Pick the Tab
      IF TP.LastLine < 3 OR TP.PgNumber = 1 THEN ITERATE FOR      ' Ignore empty tabs
      IF ISFALSE IsClip AND ISFALSE IsSetEdit THEN                ' Are we not in ClipBoard/SetEdit Mode?
         TabType = IIF$(IsBrowse , "(B)", IIF$(IsView, "(V)", "(E)"))  ' Set basic Tab Type
         IF cTab = TP.PgNumber THEN TabType = LLCASE(TabType)     ' Lowercase the active tab
      END IF                                                      '

      IF ISFALSE IsMedit AND ISFALSE IsClip AND ISFALSE IsSetEdit THEN ' If not oddball type
         IF TP.TIPFilePath <> $Empty THEN _                       '
            mrf = TabType + TP.TIPFilePath + IIF$(ISNULL(mrf), "", "?") + mrf ' Add the open filename

      ELSEIF IsMEdit THEN                                         ' It's a MEdit session
         FOR j = 1 TO TP.MEditCount                               ' Put all the names in a string
            mnames = TP.MEditListGet(j) + IIF$(ISNULL(mnames), "", "|") + mnames
         NEXT j                                                   '
         IF ISNOTNULL(mnames) THEN _                              '
            mrf = TabType + mnames + IIF$(ISNULL(mrf), "", "?") + mrf' Add the mnames string
      END IF                                                      '
      TAB SELECT hWnd, %IDC_SPFLiteTAB, TP.PgNumber               ' Select the tab
      IF LEFT$(shutcmd, 3) = "END" THEN                           ' Which type
         pCmdEND(shutcmd)                                         ' Let END have a go
      ELSE                                                        '
         pCmdCANCEL(shutcmd)                                      ' else let CANCEL have a go
      END IF                                                      '
      IF ISFALSE gfEndAll THEN GOTO Exit2                         ' Somebody selected CANCEL, bail out
   NEXT i                                                         '

   TP = Tabs(cTab)                                                ' Put back the visible tab

   '----- Save FM stuff
   TP = Tabs(1)                                                   ' Point at the FM tab
   sIniSetString("FManager", "Recall", TP.FileListNm)             '
   sIniSetString("FManager", "DefDir1", TP.FPath)                 '
   sIniSetString("FManager", "DefTypes", TP.FMask)                '


Exit1:
   IF noreopen THEN mrf = ""                                      ' If NOREOPEN, null the list
   sIniSetString("General", "MRFList", mrf)                       ' Save it
   sRetrSave                                                      ' Go save the Retrieve stack
   gfTermFlag = %True                                             ' So we don't loop


   DIALOG END hWnd                                                ' Kill our dialog
   TabsNum = 0                                                    ' No more tabs
   gTabDelCtr = 0                                                 ' No more deletes
   gTabSwitch = 0                                                 ' No more switches

Exit2:
   ON ERROR GOTO 0                                                ' Turn trap off
   MExitSub

'---- Just in case
Whoops:
   IF ISFALSE gfEndAll THEN RESUME Exit2                          ' Somebody said CANCEL, don't save MRF
   RESUME Exit1                                                   '
END SUB

SUB      sDoMEDIT(pCmd AS STRING)
'---------- Multi-Edit startup
LOCAL lclCmd, fn, fn2 AS STRING, i, j, op, NewTab, fMIX, WatchOff AS LONG
LOCAL DoMIO AS iIO                                                ' For our I/O stuff
LOCAL DoPrf AS iProf                                              ' For local Profile area
   MEntry
   IF TP.CmdParse(pCmd) THEN MExitSub                             ' Do basic parsing, exit if errors
   IF pCmdNumOps >= 2 AND pCmdOpsType(2) = %KWNew THEN NewTab = %True ' Set NewTab based on whether 2nd param says NEW
   IF pCmdNumOps = 1 AND pCmdOpsType(1) = %KWNew THEN pCmdNumOps = 0 ' Fudge entry from FM

   IF pCmdNumOps = 0 THEN                                         ' No Operands
      fn = sDoOpenFile("Specify file to Edit")                    ' Go get a filename
      IF ISNULL(fn) THEN _                                        '
         ScError(nMac(%eNone), "File selection cancelled"): MExitSub ' No selection?   Bail out
      pCmdOps(1) = TRIM$(fn)                                      ' Trim the filename
      pCmdNumOps = 1                                              ' Fudge one operand so we can use loop
   END IF                                                         '

   '----- See if all files Exist and not Open already
   FOR op = 1 TO pCmdNumOps                                       ' Loop through operands
      IF op = 2 AND pCmdOpsType(op) = %KWNew THEN ITERATE FOR     ' Skip NEW operand if the 2nd operand
      fn = pCmdOps(op)                                            ' Get the filename
      fn2 = PATHSCAN$(FULL, fn)                                   ' See if file exists and get full name
      IF ISNOTNULL(fn2) THEN                                      ' A real file, see if in use
         i = VAL(sFileQueue("S", " ", fn2))                       ' Returns tab number if open, else zero
         IF i > 0 THEN                                            ' Tab number?
            TP = Tabs(i)                                          ' Switch to found tab
            scError(%eFail, "File already open in this tab")      ' Set message
            gTabSwitch = i                                        '
            MExitSub                                              '
         END IF                                                   '
      ELSE                                                        '
         scError(nMac(%eFail), fn + " does not exist"): MExitSub  '
      END IF                                                      '
   NEXT op                                                        '

   '----- See if a huge # of files
   IF pCmdNumOps > 100 THEN                                       ' More than 100 files?
      j = sDoMsgBox("You are loading more than 100 files. This may trigger crashes due to" + $CRLF + _
                    "exhausted Stack space. Turning off File-Watch will prevent this crash." + $CRLF + _
                    "Click |KYES |Bto turn off File-Watch, or |KNO |Bto attempt to continue.", _
                    %MB_YESNO OR %MB_USERICON, "SPFLite File-Watch")
      WatchOff = IIF(j = %IDYES, %True, %False)                   ' Set chosen Suppress File-Watch
   END IF                                                         '

   '----- Now do the Opens
   FOR op = 1 TO pCmdNumOps                                       ' Loop through operands
      IF op = 2 AND pCmdOpsType(op) = %KWNew THEN ITERATE FOR     ' Skip NEW operand if the 2nd operand
      LET DoMIO = CLASS "cIO"                                     '
      LET DoPrf = CLASS "cProf"                                   '
      fn = pCmdOps(op)                                            ' Get the filename
      fn2 = PATHSCAN$(FULL, fn)                                   ' See if file exists and get full name
      IF NewTab THEN                                              ' Open this in a new TAB?
         NewTab = %False                                          ' Just NEW once
         TP.ErrFlag = %eNone                                      ' Say we're OK
         ENV.PMode = %MMedit                                      ' Say we're starting an MEdit tab
         DoMIO.Setup("ER", "", "", fn2)                           ' See if it exists
         IF DoMIO.EXEC THEN                                       ' Go validate
            scError(%eFail, fn2 + " ignored - " + DoMIO.ResultMsg)' Oops?  Bail out
            IF op = 1 THEN NewTab = %True                         ' Re-establish NewTab since we're skipping sTabAdd
         ELSE                                                     '
            sTabAdd(fn2, "")                                      ' Yes, let sTabAdd do the work
            IF gTabSwitch <> 0 THEN _                             ' If switch
               TP = Tabs(gTabSwitch)                              ' Switch to the tab just added
         END IF                                                   '
      ELSE                                                        '
         IF TP.MeditTbl("S", fn2) > 0 THEN _                      ' Already here in MEdit mode?
            ScError(%eFail, "File already Open in this MEdit session"): MExitSub
         DoMIO.Setup("ER", "", "", fn2)                           ' Set for Exist ROTest
         IF DoMIO.EXEC THEN                                       ' Go validate
            scError(%eFail, fn2 + " ignored - " + DoMIO.ResultMsg)' Oops?  Bail out
         ELSE                                                     '
            IF TP.MEditCount = 0 AND TP.LastLine > 2 THEN         ' If not yet Medit mode
               FOR i = 2 TO TP.LastLine - 1                       ' Mark all existing lines as being file 1 lines
                  TP.LMixSet(i, 1)                                '
               NEXT i                                             '
               TP.LInsertEmpty(1, 1, %File)                       ' Insert for the =FILE> line
               TP.LMixSet(2, 1)                                   ' Mark as File 1
               TP.LTxtSet(2, TP.TIPFilePath)                      ' Stuff existing filename in as the text
               TP.MeditTbl("A", TP.TIPFilePath)                   ' Add to Medit table
               TP.MEditFlagSet(1, IIF(IsTPModdFlag, %True, %False))' Copy modified status
               TP.UpdLControl(2)                                  ' Setup LLCtl
            END IF                                                '
            TP.TMode = %MMEdit                                    ' Set just in case
            fMIX = TP.MEditTbl("A", fn2)                          ' Go get MIX value
            j = TP.LastLine - 1                                   ' Point at last data line
            TP.LInsertEmpty(j, 1, %File)                          ' Insert for the =FILE> line
            TP.LTxtSet(j + 1, fn2)                                ' Stuff filename in as the text
            TP.LMIXSet(j + 1, fMIX)                               ' Mark with MIX index
            TP.UpdLControl(j + 1)                                 ' Setup LLCtl
            TP.CopyAFile(j + 1, DoMIO, DoPrf, 0, 0, 0, %False)    ' Go load the data (afterline, DoMIO, DoPrf, PMFlag, fromline, toline, not quick)
            IF TP.errFlag <> %eNone THEN MExitSub                 ' If errors, bail out
            sFileQueue("A", " ", fn2)                             ' Add to Open queue
            IF ISFALSE WatchOff THEN                              ' Do FileWatch if not suppressed
               IF TP.FileWatch(fn2, %WatchStart) THEN             ' Establish the watch
                  scError(0, "File watch could not be established")  '
               END IF                                             '
            END IF                                                '
         END IF                                                   '
      END IF                                                      '
      LET DoMIO = NOTHING                                         '
      LET DoPrf = NOTHING                                         '
   NEXT op                                                        '
   TP.WindowTitle                                                 ' Alter window/Tab titles
   TP.ErrFlag = %eNone                                            ' Say we're OK
   TP.UndoSave()                                                  ' Take an initial one
   gTabSwitch = TP.PgNumber                                       ' Set to 'switch' here
   MExit
END SUB

FUNCTION sDoInputBox(iText AS STRING, iTitle AS STRING, iDefault AS STRING) AS STRING
'---------- Get an InputBox answer
   sPopReady                                                      ' Ready for pop-up
   FUNCTION = TRIM$(INPUTBOX$(iText, iTitle, iDefault))           ' Issue it, return trimmed answer
   sPopReset                                                      ' Reset popup state
END FUNCTION

FUNCTION      sDoMsgBox(mTxt AS STRING, mFlags AS LONG, title AS STRING, OPT mPitch AS LONG) AS LONG
'---------- Issue a MSGBOX
   IF ENV.InitDone THEN sPopReady                                 ' Ready for pop-up
   IF ISMISSING(mPitch) THEN                                      ' Pitch provided?
      FUNCTION = MyMsgBox(mTxt, mFlags, title)                    ' Pass on without pitch
   ELSE                                                           '
      FUNCTION = MyMsgBox(mTxt, mFlags, title, mPitch)            ' Pass the pitch onward
   END IF                                                         '
   IF ENV.InitDone THEN sPopReset                                 ' Reset popup state
END FUNCTION

FUNCTION sDoOpenFile(iPrompt AS STRING) AS STRING
'---------- Prompt for a filename
LOCAL fn AS STRING
   sPopReady                                                      ' Ready for pop-up
   DISPLAY OPENFILE hWnd, , , iPrompt, sGetDefDir, CHR$("All Files", 0, "*.*", 0), "", "", _
                    %OFN_ENABLESIZING TO fn                       '
   sPopReset                                                      ' Reset popup state
   FUNCTION = TRIM$(fn)                                           ' Return answer
END FUNCTION

SUB      sDoPendingTabDels
'---------- Do any penfing Tab Deletes
LOCAL i, j, k AS LONG
   MEntry
   IF gTabDelCtr = 0 THEN MExitSub                                ' Nothing, exit quickly
   ARRAY SORT gTabDelList() FOR gTabDelCtr, TAGARRAY gTabDelNext(), DESCEND
   FOR i = 1 TO gTabDelCtr                                        ' Do the tab deletes
      TP = Tabs(gTabDelList(i))                                   ' Switch to the tab to be deleted
      k = gTabDelList(i)                                          ' Save the last tab deleted
      sTabStackDel(TP.PgNumber)                                   ' Remove from Tab Stack
      j = GTabStack(1)                                            ' Get tab to switch to
      sTabDel                                                     ' Go delete it
   NEXT                                                           '
   RESET gTabDelCtr, gTabDelList(), gTabDelNext()                 ' Clear out the table

   '----- Figure out what other tab to switch to (or shut down)

   IF TabsNum = 1 AND ENV.FMCloseFlag THEN                        ' All that's left is FM and we should close it
      TabsNum = 0                                                 ' Say no more active tabs
      gfTermFlag = %True                                          ' So we don't loop
      GOSUB SaveFMStuff                                           ' Go save the FM status
      DIALOG END hWnd                                             ' Kill our dialog
   END IF                                                         '

   IF j <> 0 THEN                                                 ' End up with a tab number
      TP = Tabs(j)                                                ' Switch to it
      TP.ErrMsg = gTabDelMsg                                      ' Issue any message
      TAB SELECT hWnd, %IDC_SPFLiteTAB, j                         '
   ELSE                                                           '
      IF TabsNum > 1 THEN                                         ' Switch to another tab
         IF k > 1 THEN                                            '
            TP = Tabs(k - 1)                                      ' Switch TP right away
            TP.ErrMsg = gTabDelMsg                                ' Issue any message
            TAB SELECT hWnd, %IDC_SPFLiteTAB, k - 1               ' Go Left
         ELSE                                                     '
            TP = Tabs(1)                                          ' Switch TP to FM
            TP.ErrMsg = gTabDelMsg                                ' Issue any message
            TAB SELECT hWnd, %IDC_SPFLiteTAB, 1                   ' Go to 1st tab
         END IF                                                   '
      ELSEIF TabsNum = 1 AND ISFALSE ENV.FMCloseFlag THEN         ' If just 1 and File Manager and we're not to autoclose it
         TP = Tabs(1)                                             ' Switch TP right away
         TP.ErrMsg = gTabDelMsg                                   ' Issue any message
         TAB SELECT hWnd, %IDC_SPFLiteTAB, 1                      ' Go to 1st tab
      ELSEIF gfEndAll THEN                                        '
         gfTermFlag = %True                                       ' So we don't loop
         GOSUB SaveFMStuff                                        ' Go save the FM status
         DIALOG END hWnd                                          ' Kill our dialog
      ELSE                                                        ' ELSE (we're all done)
         gfTermFlag = %True                                       ' So we don't loop
         GOSUB SaveFMStuff                                        ' Go save the FM status
         DIALOG END hWnd                                          ' Kill our dialog
      END IF                                                      '
   END IF                                                         '
   RESET gTabDelCtr, gTabDelMsg, gTabDelList(), gTabDelNext()     ' Clear out the table
   MExitSub                                                       '

SaveFMStuff:
   sRetrSave                                                      ' Save retrieve stack
   TP = Tabs(1)                                                   ' Point at the FM tab
   sIniSetString("FManager", "Recall", TP.FileListNm)             '
   sIniSetString("FManager", "DefDir1", TP.FPath)                 '
   sIniSetString("FManager", "DefTypes", TP.FMask)                '
   RETURN
END SUB

SUB      sDoPendingTabSwitch
'---------- Do a penfing Tab Switch
LOCAL i, j, k AS LONG
   MEntry
   IF gTabSwitch = 0 THEN MExitSub                                ' If no switch, just exit
   TP = Tabs(gTabSwitch): k = %True                               ' Switch to the tab to be switched to
   IF gTabSwitchMsg <> "" THEN                                    ' An associated message?
      TP.ErrMsg = gTabSwitchMsg                                   ' Set it
      TP.AttnDo = (TP.AttnDo OR %Refresh)                         ' Have it looked at
   END IF                                                         '
   IF gTabSwitchCmd <> "" THEN                                    ' An associated message?
      TP.pCommand = gTabSwitchCmd                                 ' Set it
      TP.AttnDo = (TP.AttnDo OR %Attention)                       ' Have it looked at
      TP.PostKeyBoard                                             '
   END IF                                                         '
   RESET gTabSwitch, gTabSwitchMsg, gTabSwitchCmd                 ' Clear things
   IF k THEN                                                      ' If we switched
      TAB SELECT hWnd, %IDC_SPFLiteTAB, TP.PgNumber               ' Select the new tab
'      GRAPHIC ATTACH TP.PgHandle, TP.WindowID                    ' Swap the default graphic area
      CONTROL SET FOCUS hWnd, %IDC_SPFLiteTAB                     ' Set focus
      sCaretDestroy                                               '
      sCaretCreate                                                '
      sDoCursor                                                   '
      sCaretShow                                                  '
   END IF                                                         '
   MExit                                                          '
END SUB

FUNCTION sDoQuoteString(x AS STRING) AS STRING
'---------- Wrap quotes around a string, or convert to hex if needed
LOCAL c, result AS STRING, n AS LONG
   MEntry
   IF INSTR(x, $DQ) = 0 THEN                                      ' See if safe to use $DQ
      result = BUILD$($DQ, x, $DQ)                                ' Yes, do so
   ELSEIF INSTR(x, $SQ) = 0 THEN                                  ' No? Maybe $SQ
      result = BUILD$($SQ, x, $SQ)                                ' Yes, do do
   ELSEIF TALLY(x, "`") = 0 THEN                                  ' No? Back-quote?
      result = BUILD$("`", x, "`")                                ' Yes, do so
   ELSE                                                           ' Wow! All 3 quotes in use

      '----- string uses all three quotes, force into hex mode
      result = "X'"                                               ' Build a hex literal
      FOR n = 1 TO LEN(x)                                         ' Loop through it
         c = MID$(x, n, 1)                                        '
         result += HEX$(ASC(c), 2)                                '
      NEXT                                                        '
      result += "'"
   END IF                                                '
   FUNCTION = result
   Mexit
END FUNCTION

SUB      sDoStatusBar(which AS STRING)
'---------- Do one or more StatusBar boxes
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL lr, lRow, lclCol, k, mx, my AS LONG
LOCAL cLoc, char, LFMask, t, lclwhich AS STRING, cloc2 AS ASCIIZ * 200
   MEntry
   IF ISFALSE ENV.InitDone OR ISTRUE gMacroMode THEN MExitSub     ' If INIT not done, or MacroMode, exit
   lclwhich = which + "O"                                         ' Always add PAD
   LFMask = REPEAT$(ENV.LinNoSize, "0")
   lRow = TP.CsrRow: lclCol = TP.CsrCol                           ' Get local copies

   IF lclwhich = "REFRESHO" THEN GOTO JustRefresh                 ' Don't waste time building things on a REFRESH
   '----- Do the requested StatusBar boxes for FM
   IF IsFMTab THEN                                                ' If Tab mode
      '----- Setup the SB boxes
      lclwhich = $AllStatusBarBoxes + "O"                         ' FM always does them all
      TP.SBSetText(%SBMode, UUCASE(TP.DefCommand))                ' Setup mode
      i = lrow - FM_Top_File_Line + TP.TopScrn                    ' Calc AFList index
      IF lRow < FM_Top_File_Line OR _                             ' If not in linenum area
         ISFALSE TP.AFIsFileDir(i) OR _                           ' or not a File Dir
         ISFALSE TP.AFIsReadOnly(i) THEN                          ' or not ReadOnly
         TP.SBSetText(%SBLinNo, " ")                              ' Setup LinNo
      ELSE                                                        '
         TP.SBSetText(%SBLinNo, " Read-Only File / Directory")    ' Setup LinNo
      END IF                                                      '
      TP.SBSetText(%SBInsOvr,   IIF$(IsTPNsrtFlag, IIF$(IsTPNsrtData, "ins", "INS"), "OVR")) ' InsOvr
      TP.SBSetText(%SBCaseWord, TP.PrfPCase + IIF$(TP.FindWord, "  W", ""))                  ' CaseWord
      TP.SBSetOvScheme(%SBMisc, IIF(ISTRUE gKbdRecFlag, %SCHiRed, 0)) ' Set color properly
      TP.SBSetText(%SBMisc,     IIF$(ISTRUE gKbdRecFlag, "KB Recording", " "))               ' Misc

   '----- Do the requested StatusBar boxes for non-FM
   ELSE
      FOR i = 1 TO LEN(lclwhich)                                  ' Loop through Box requests
         SELECT CASE AS CONST$ MID$(lclwhich, i, 1)               ' See which are called for
            '----- Mode box
            CASE $SBMode                                          '
               TP.SBSetOvScheme(%SBMode, 0)                       ' Start with default colors
               cLoc = SWITCH$(IsBrowse, "Browse", IsView, "View", IsClip, "Clip", IsSetEdit, "SET Edit", %True, "Edit")
               IF cLoc = "View" AND TP.TIPROStat THEN cLoc = "RdOnly"         '
               IF (cLoc ="Browse" OR cLoc = "View" OR cLoc = "RdOnly") AND IsTPModdFlag THEN
                  TP.SBSetOvScheme(%SBMode, %SCHiRed)             ' Make it WHITE on RED
               END IF                                             '
               IF IsMedit THEN                                    ' If MEdit, re-do it all
                  cloc = FORMAT$(TP.MEditCount) + " Edit"         ' Build basic saying how many files
                  FOR k = 1 TO TP.MEditCount                      ' See how many modified
                     IF TP.MEditFlagGet(k) THEN INCR j            ' Count modified
                  NEXT k                                          '
                  IF j THEN cloc += " " + FORMAT$(j) + "*"        ' If any modified, show the count
               ELSE                                               '
                  cLoc += IIF$(IsTPModdFlag, " *", "")            ' Do simple modified display
               END IF                                             '
               TP.SBSetText(%SBMode, cloc)                        ' Filler box

            '----- LinNo box
            CASE $SBLinNo                                         '
               TP.SBSetOvScheme(%SBLinNo, 0)                      ' Start with default colors
               IF TP.CursData THEN                                ' If in the data area
                  lr = sGetIX(lRow)                               ' Add in the cursor position
                  IF lr = -1 OR lr = -2 THEN lr = sGetIX(lRow - ABS(lr))      '
                  IF lr > 0 THEN                                  '
                     IF TP.LFlagData(lr) THEN                     '
                        cLoc = "L " + FORMAT$(VAL(TP.LLNumGet(lr)), LFMask) + "  C " + FORMAT$(lclCol - gLNPadCol + TP.Offset, 6)
                        GOSUB AddLabels                           ' Add Labels/Tags
                     ELSEIF ISTRUE (TP.LFlagGet(lr) AND (%Tabs OR %Mark OR %Mask OR %Note OR %Cols OR %Bounds)) THEN
                        cLoc = "L ---  C " + FORMAT$(lclCol - gLNPadCol + TP.Offset, 6)
                        GOSUB AddLabels                           ' Add Labels/Tags
                     ELSEIF ISTRUE TP.LFlagXclude(lr) AND TP.CsrLinDx > 0 THEN
                        cLoc = "L " + FORMAT$(TP.CsrLinDX, LFMask) + "  C " + FORMAT$(lclCol - gLNPadCol + TP.Offset, 6)
                        TP.SBSetOvScheme(%SBLinNo, %SCHiGreen)    ' Make it WHITE on GREEN
                        GOSUB AddLabels                           ' Add Labels/Tags
                     ELSEIF ISTRUE TP.LFlagXclude(lr) THEN        ' Simple Exclude
                        IF TP.LWrk1Get(lr) = 1 THEN               ' Just a single line?
                           INCR lr                                ' Point at the real line
                           cLoc = "L " + FORMAT$(VAL(TP.LLNumGet(lr)), LFMask) + "  C " + FORMAT$(lclCol - gLNPadCol + TP.Offset, 6)
                           GOSUB AddLabels                        ' Add Labels/Tags
                        ELSE                                      '
                           cLoc = "L " + FORMAT$(VAL(TP.LLNumGet(lr + 1)), LFMask) + " - " + FORMAT$(VAL(TP.LLNumGet(lr + TP.LWrk1Get(lr))), LFMask)
                        END IF                                    '
                     ELSE                                         '
                        cloc = " "                                ' Blank box
                     END IF                                       '
                     TP.SBSetText(%SBLinNo, cloc)                 ' LinNo box
                  END IF                                          '
               ELSEIF TP.CursLinN THEN                            ' If in the Line Number area
                  lr = sGetIX(lRow)                               ' Add in the cursor position
                  IF lr = -1 OR lr = -2 THEN lr = sGetIX(lRow - ABS(lr))      '
                  IF lr > 0 THEN                                  '
                     IF TP.LFlagData(lr) THEN                     '
                        cLoc = "  L " + FORMAT$(VAL(TP.LLNumGet(lr)), LFMask)
                        GOSUB AddLabels                           ' Add Labels/Tags
                     ELSEIF ISTRUE TP.LFlagXclude(lr) THEN        ' Simple Exclude
                        IF TP.LWrk1Get(lr) = 1 THEN               ' Just a single line?
                           INCR lr                                ' Point at the real line
                           cLoc = "  L " + FORMAT$(VAL(TP.LLNumGet(lr)), LFMask)
                           GOSUB AddLabels                        ' Add Labels/Tags
                        ELSE                                      '
                           cLoc = "  L " + FORMAT$(VAL(TP.LLNumGet(lr + 1)), LFMask) + " - " + FORMAT$(VAL(TP.LLNumGet(lr + TP.LWrk1Get(lr))), LFMask)
                        END IF                                    '
                     ELSE                                         '
                        cloc = " "                                '
                     END IF                                       '
                     TP.SBSetText(%SBLinNo, cloc)                 ' LinNo box
                  END IF                                          '
               ELSEIF TP.CursCmnd THEN                            ' If in the command area
                  IF ISNOTNULL(gCmdRtrevMsg) THEN                 ' Retrieve message?
                     TP.SBSetText(%SBLinNo, gCmdRtrevMsg)         ' LinNo box
                  ELSEIF ISFALSE IsClip AND ISFALSE IsSetEdit AND ISFALSE IsMEdit AND TP.TIPDate <> "" THEN                                '
                     TP.SBSetText(%SBLinNo, "  " + TP.TIPDate + "  " + TP.TIPTime) ' LinNo box
                  ELSE                                            '
                     TP.SBSetText(%SBLinNo, " ")                  ' LinNo box
                  END IF                                          '
               END IF                                             '

               IF IsMedit THEN                                    ' If MEdit, do Title bar fudge
                  lr = sGetIX(lRow)                               ' Add in the cursor position
                  IF lr = -1 OR lr = -2 THEN lr = sGetIX(lRow - ABS(lr))
                  IF lr = 0 THEN                                  ' Wasn't on a data line
                     FOR mx = 3 + TP.PrfCols TO gwScrHeight       ' Find 1st data line on screen
                        my = sGetIX(mx)                           ' Get a line reference
                        IF my = -1 OR my = -2 THEN my = sGetIX(mx - ABS(my))
                        IF my AND TP.LFlagData(my) THEN           ' A data line?
                           lr = my                                ' Lets use it
                           EXIT FOR                               ' We've found one
                        END IF                                    '
                     NEXT j                                       '
                  END IF                                          '
                  IF lr > 0 THEN                                  '
                     lr = TP.LMixGet(lr)                          ' Get the MIX index
                     IF lr > 0 THEN                               ' Valid?
                        cloc = TP.MEditListGet(lr)                ' Get the filename
                        cloc = MID$(cloc, INSTR(-1, cloc, "\") + 1)  ' Strip off path
                        cloc2 = cloc + " - Multi-Edit - SPFLite" + "(v" + ENV.PgmVers + ")"
                        SetWindowText(hWnd, cloc2)                ' Alter window title
                     END IF                                       '
                  END IF                                          '
               END IF                                             '

            '----- Lines box
            CASE $SBLines                                         '
               TP.SBSetText(%SBLines, "Lines: " + FORMAT$(TP.LastReal)) ' Lines:

            '----- Cols box
            CASE $SBCols                                          '
               TP.SBSetText(%SBCols, "Cols " + FORMAT$((TP.Offset + 1), "#####") + " to " + FORMAT$((TP.Offset + gDataLen), "#####")) 'Cols

            '----- Bounds box
            CASE $SBBnds                                          '
               TP.SBSetOvScheme(%SBBnds, 0)                       ' Start with default colors
               cLoc = "Bnds: "                                    ' Add Bnds
               IF TP.PrfBndLeft = 1 AND TP.PrfBndRight = 0 THEN   ' Entire line bounds?
                  cLoc += "MAX"                                   ' Say MAX
               ELSEIF TP.PrfBndRight = 0 THEN                     ' RB to Max?
                  cLoc += FORMAT$(TP.PrfBndLeft) + " to MAX"      ' Setup display
                  TP.SBSetOvScheme(%SBBnds, %SCHiRed)             ' Set WHITE on RED
               ELSEIF TP.PrfBndRight > 0 THEN                     ' LB to RB?
                  cLoc += FORMAT$(TP.PrfBndLeft) + " to " + FORMAT$(TP.PrfBndRight)' Setup display
                  TP.SBSetOvScheme(%SBBnds, %SCHiRed)             ' Set WHITE on RED
               END IF                                             '
               TP.SBSetText(%SBBnds, cLoc)                        '

            '----- InsOvr box
            CASE $SBInsOvr                                        '
               TP.SBSetOvScheme(%SBInsOvr, 0)                     ' Start with default colors
               TP.SBSetText(%SBInsOvr, IIF$(IsTPNsrtFlag, "INS", "OVR")) ' Add Insert Mode
               t = IIF$(IsTPNsrtFlag, "INS", "OVR") + " " + FORMAT$(TP.Flag2)
               IF IsTPNsrtFlag AND IsTPNsrtData THEN              '
                  TP.SBSetOvScheme(%SBInsOvr, %SCHiGreen)         ' Set WHITE on GREEN
               END IF                                             '

            '----- CaseWord box
            CASE $SBCaseWord                                      '
               TP.SBSetText(%SBCaseWord, TP.PrfPCase + IIF$(TP.FindWord, " W", "")) ' Add default literal Case and Find Word modes

            '----- Change box
            CASE $SBChange                                        '
               TP.SBSetText(%SBChange, IIF$(TP.PrfChangeMode = "D", "DS", "CS")) ' Add Change setting

            '----- State box
            CASE $SBState                                         '
               TP.SBSetText(%SBState, "S" + IIF$(IsTPStateExist, "+", "-")) ' Add STATE staus

            '----- Misc box
            CASE $SBMisc                                          '
               TP.SBSetOvScheme(%SBMisc, 0)                       ' Start with default colors
               IF ISFALSE IsTPMarkDrawn THEN                      ' If nothing hi-lighted
                  IF IsTPSwapDrawn THEN                           ' Insert the Swap Pending
                     cloc = "Swap Pending"                        '
                     TP.SBSetOvScheme(%SBMisc, %SCHiRed)          ' Set WHITE on RED
                  ELSE                                            '
                     IF IsTPPTypeMode THEN                        ' If PowerType
                        cLoc = "PowerType"                        ' Add to status bar
                        TP.SBSetOvScheme(%SBMisc, %SCHiRed)       ' Set WHITE on RED
                     ELSE                                         '
                        IF gKbdRecFlag THEN                       ' Multi use, pick one
                           cLoc = "KB Recording"                  '
                           TP.SBSetOvScheme(%SBMisc, %SCHiRed)    ' Set WHITE on RED
                        ELSEIF TP.PrfAutoSave = 0 THEN            '
                           cLoc = "AUTOSAVE OFF"                  '
                           TP.SBSetOvScheme(%SBMisc, %SCHiRed)    ' Set WHITE on RED
                        ELSEIF TP.TIPProfile <> TP.TIPExtn AND lrow = 1 THEN  '
                           cLoc = "Profile: " + TP.TIPProfile     '
                        ELSE                                      '
                           cLoc =  " "                            '
                        END IF                                    '
                        IF cLoc = " " THEN                        ' Result of prev tests a blank?
                           lr = sGetIX(lRow)                      ' Get the line number
                           IF lr = -1 OR lr = -2 THEN lr = sGetIX(lRow - ABS(lr))
                           IF lr > 0 THEN                         '
                              IF TP.LFlagData(lr) THEN            '
                                 cLoc = "Line Len " + FORMAT$(TP.LTxtLen(lr), "0000")
                              END IF                              '
                           END IF                                 '
                        END IF                                    '
                     END IF                                       '
                  END IF                                          '
                  TP.SBSetText(%SBMisc, cLoc)                     '
               ELSE                                               '
                  IF IsTPSwapDrawn THEN                           ' Insert the Swap Pending
                     cloc = "Swap Pending"                        '
                     TP.SBSetOvScheme(%SBMisc, %SCHiRed)          ' Set WHITE on RED
                  ELSE                                            '
                     IF TP.MarkELin <> TP.MarkSLin  OR TP.MarkECol <> TP.MarkSCol OR _
                        TP.MarkSCol > TP.LTxtLen(TP.MarkSLin) THEN'
                        cLoc = "Len "                             ' Build selected message
                        cloc += FORMAT$(TP.MarkECol - TP.MarkSCol + 1)        '
                        IF TP.MarkELin <> TP.MarkSLin THEN        '
                           k = 0                                  '
                           FOR j = TP.MarkSLin TO TP.MarkELin     ' Count lines in group
                              IF TP.LFlagData(j) THEN INCR k      '
                           NEXT j                                 '
                           IF k > 0 THEN cLoc += " x " + FORMAT$(k)
                        END IF                                    '
                     ELSE                                         '
                        char = MID$(TP.LTxtGet(TP.MarkSLin), TP.MarkSCol, 1)  ' Get char at cursor
                        IF TP.PrfSrceXlate THEN _                 '
                           TP.Translate(char, TP.TPPrfGetSA2SPtr) ' Translate ANSI to SOURCE
                        cLoc = "X'" + HEX$(ASC(char), 2) + "' = " + FORMAT$(ASC(char)) ' Build rest of message
                     END IF                                       '
                  END IF                                          '
                  IF IsTPPTypeMode THEN cloc = "PT  " + cloc      '
                  TP.SBSetText(%SBMisc, cLoc)                     '
               END IF                                             '

            '----- Select box
            CASE $SBSelect                                        '
               IF IsTPSlecSet THEN                                ' If nothing set, do nothing
                  IF IsTPSlecActive THEN                          ' Active?
                     cloc = "@ " + FORMAT$(TP.SlecSCol) + " " + FORMAT$(TP.SlecECol) + "  # " + _
                            "." + FORMAT$(TP.SlecSLin) + " ." + FORMAT$(TP.SlecELin)
                     TP.SBSetText(%SBSelect,  cloc)               '
                  ELSE                                            '
                     TP.SBSetText(%SBSelect, "Select")            '
                  END IF                                          '
               ELSE                                               '
                  TP.SBSetText(%SBSelect, " ")                    '
               END IF                                             '

            '----- Caps box
            CASE $SBCaps                                          '
               cLoc = ""                                          '
               IF TP.PrfFold THEN cLoc = "FOLD | "                ' Insert FOLD if active
               IF TP.PrfCapsDesired <> 2 THEN                     ' Add CAPS
                  cLoc += IIF$(TP.PrfCapsDesired, "CAPS ON", "CAPS OFF")      '
               ELSE                                               '
                  cLoc += "CAPS AUTO:" + IIF$(TP.PrfCapsActual, "on", "off")  '
               END IF                                             '
               TP.SBSetText(%SBCaps, cLoc)                        '

            '----- Source box
            CASE $SBSource                                        '
               TP.SBSetText(%SBSource, TP.PrfPSource)             ' Add Source Encoding type

            '----- EOL box
            CASE $SBEOL                                           '
               cLoc = TP.PrfEOL                                   ' Add EOL setting
               IF TP.PrfPageFlag <> 1 OR LEFT$(TP.PrfEOL, 4) <> "AUTO" THEN   ' If not in PAGE mode
                  TP.SBSetText(%SBEOL, cLoc)                      '
               ELSE                                               ' We're in PAGE mode
                  IF TP.TopScrn = 0 THEN TP.SBSetText(%SBEOL, cLoc): GOTO ExitEOL  ' Escape hatch
                  lr = sGetIX(lRow)                               ' Get line number of the cursor position
                  IF lr = 0 THEN lr = TP.TopScrn                  ' Or top-of-screen
                  IF lr = 1 THEN lr = 2                           ' Fudge over top of dataset line
                  DO WHILE ISFALSE TP.LFlagData(lr) AND ISFALSE TP.LFlagBottom(lr) ' Step over any special lines
                     INCR lr                                      '
                  LOOP                                            '
                  DO WHILE ISFALSE TP.LFlagPage(lr) AND (TP.LFlagData(lr) OR TP.LFlagBottom(lr))  ' Look for start of page
                     DECR lr                                      '
                  LOOP                                            '
                  IF ISFALSE TP.LFlagPage(lr) THEN TP.SBSetText(%SBEOL, cLoc): GOTO ExitEOL   ' Escape hatch
                  cLoc = "Pg: " + FORMAT$(TP.LWrk2Get(lr) + TP.PrfPageOffset, "#;-#")   ' Build msg
                  lr = TP.LastLine                                '
                  DO WHILE ISFALSE TP.LFlagPage(lr) OR lr = 1     ' Look for start of page
                     DECR lr                                      '
                  LOOP                                            '
                  IF TP.PrfPageOffset = 0 THEN                    ' Add 'of' only with no offset
                     cLoc +=  " of " + FORMAT$(TP.LWrk2Get(lr))   ' Complete message
                  END IF                                          '
                  TP.SBSetText(%SBEOL, cLoc)                      ' Display it
                  ExitEOL:                                        ' Escape point
               END IF                                                         '

         END SELECT                                               '
      NEXT i
      TP.SBSetText(%SBPad, " ")                                   ' Filler box
   END IF                                                         '

JustRefresh:
   '----- Now output the selected boxes
   FOR i = 1 TO TP.SBCount                                        ' Loop through
      j = TP.SBXrefGet(i)                                         ' Get the real SB item number
      IF TP.SBGetActive(j) = "Y" AND INSTR(lclwhich, TP.SBGetID(j)) > 0 THEN
         SENDMESSAGE (hStatusBar, %SB_SETTEXT, (i - 1) OR %SBT_OWNERDRAW, BYVAL TP.SBGetMySelfP(j))
      END IF                                                      '
   NEXT i                                                         '
   MExitSub                                                       ' We're done

AddLabels:
   IF TP.LTagGet(lr) <> $BlankLNo  AND TP.LLblGet(lr) <> $BlankLNo THEN
      cloc += "  " + TRIM$(TP.LTagGet(lr))
   END IF
   RETURN

END SUB

FUNCTION sEnumWindowsProc(BYVAL lHandle AS LONG, BYVAL lNotUsed AS LONG) AS LONG
'---------- Get list of windows
LOCAL wTitle AS ASCIIZ * 256, sTitle AS STRING
LOCAL lPos   AS LONG
   GetWindowText lHandle, wTitle, 255                             '
   sTitle = TRIM$(wTitle)                                         '
   IF LEN(sTitle) THEN                                            '
      INCR gWinListPos                                            '
      IF gWinListPos > UBOUND(gWinList()) THEN _                  '
         REDIM PRESERVE gWinList(1 TO gWinListPos * 2) AS GLOBAL WININFOTYPE ' Expand if needed
      gWinList(gWinListPos).WinTitle = sTitle                     '
      gWinList(gWinListPos).WinHandle = lHandle                   '
   END IF                                                         '
   FUNCTION = 1                                                   '
END FUNCTION

SUB      sFileChanged()
'---------- Handle the notification that a file has changed
LOCAL i, ThisID, hForeWnd, ForeID, NewStart, lclMode AS LONG, lclFn, MorD AS STRING
STATIC Pending AS LONG

   MEntry
   IF Pending THEN MExitSub                                       '
   FOR i = 1 TO UBOUND(gFQ())                                     ' Search for it
      IF ISTRUE gFQ(i).gInUse AND gFQ(i).gPgNumber = TP.PgNumber THEN ' Is this an entry for this tab?
         IF gFQ(i).gFlag <> " " THEN                              ' Been flagged?
            lclFn = gFQ(i).gWatchFile                             ' Yes, get the filename involved
            MorD = gFQ(i).gFlag                                   ' And the type of change flag
            gFQ(i).gFlag = " "                                    ' Reset the flag
            EXIT FOR                                              ' Done searching
         END IF                                                   '
      END IF                                                      '
   NEXT i                                                         '
   IF ISNULL(lclFn) THEN TRACE OFF: MExitSub                      ' Oops! Shouldn't happen, but ...

   '----- Check the user's Notify options
   IF ENV.NotifyLevelT = 0 THEN MexitSub                          ' Never notify, just exit
   IF ENV.NotifyLevelT = 1 THEN                                   ' Just EDIT tabs?
      IF IsBrowse OR IsView THEN MexitSub                         ' But this is BROWSE/VIEW, just exit
   END IF                                                         ' Following then is for Notify = 2 (All)

   '----- See if file was deleted
   IF MorD = "D" THEN                                             ' File has been deleted
      TP.WatchFlag = " "                                          ' Reset the WatchFlag
      Pending = %True                                             ' Say we've got a prompt out.
      sDoMsgBox "The Current file:" + $CRLF + $CRLF + _
             "|K" + lclFn + $CRLF + $CRLF + _
             "|Bmay have been modified or Deleted. SPFLite cannot determine which." + $CRLF + $CRLF + _
             "You should save this session under a temporary name" + $CRLF + _
             "(|KOTHER THAN |Bit's true name) using the |KCREATE |Bcommand" + $CRLF + _
             "for your own protection before investigating" + $CRLF + _
             "the cause of this alteration.", %MB_OK OR %MB_USERICON, "SPFLite"
      Pending = %False                                            ' Kill outstanding prompt
      TRACE OFF

   '----- File has been Modified
   ELSEIF MorD = "M" THEN                                         ' Should be the only thing left but ...
      TP.WatchFlag = " "                                          ' Reset the WatchFlag
      IF IsMEdit THEN                                             ' Whoops! a MEdit session
         Pending = %True                                          ' Say we've got a prompt out.
         sDoMsgBox "The Current file:" + $CRLF + $CRLF + _
                "|K" + lclFn + $CRLF + $CRLF + _
                "|Bhas been Modified and is part of this Multi-Edit session." + $CRLF + $CRLF + _
                "You should evaluate the changes you may have made in this session and" + $CRLF + _
                "determine the best action. If suitable, a RELOAD command should" + $CRLF + _
                "be issued to re-load all the files for this session; but this will lose" + $CRLF + _
                "all current changes." + $CRLF + $CRLF + _
                "Or you could evaluate what the external change was and ignore the RELOAD.", %MB_OK OR %MB_USERICON, "SPFLite"
         Pending = %False                                         ' Kill outstanding prompt
         MExitSub
      END IF                                                      '
      Pending = %True                                             ' Say we've got a prompt out.
      i = sDoMsgBox( "The Current file:" + $CRLF + $CRLF + _
                "|K" + lclFn + $CRLF + $CRLF + _
                "|BHas been modified elsewhere" + $CRLF + _
                "Do you want to load the modified version?", %MB_YESNO OR %MB_USERICON, "SPFLite")
      Pending = %False                                            ' Kill outstanding prompt
      IF i = %IDYES THEN                                          ' Reply was YES, reload it
         lclMode = TP.TMode                                       ' Save tab mode
         TP.MarkKill                                              ' Kill any active block select
         TP.SwapKill                                              ' Kill any active Swap select
         i = TP.FileWatch("", %WatchEnd)                          ' Kill any prior Watch
         IF ISFALSE IsTPModdFlag AND TP.PrfStart = "NEW" THEN     ' If not modified and START NEW
            NewStart = TP.LastLine - 1                            ' Save where .START should go
         END IF                                                   '
         TP.LInitTxtData(lclFn)                                   ' Wipe everything out then
         TP.InitaFile(%False)                                     ' Initialize file stuff again
         IF NewStart <> 0 AND NewStart < TP.LastLine THEN         ' Is the save NewStart valid?
            TP.TopScrn = NewStart                                 '
            TP.LLblSet(TP.LastLine - 1, ".START")                 ' Add the Label
            TP.UpdLControl(TP.LastLine - 1)                       ' Put back the line number
         END IF                                                   '
         TP.TMode = lclMode                                       ' Restore Mode
         TP.UndoInit                                              ' Init the Undo file names
         TP.UndoSave()                                            ' Take an initial one
         TP.TIPTimeDateRefresh                                    ' Get refreshed Date/Time
         TP.WindowTitle                                           ' Alter window title
         scError(0, "File data has been re-loaded")               ' Issue error
      END IF                                                      '
   END IF                                                         '
   MExit
END SUB

FUNCTION sFileListAdd(FileList AS STRING, NewFile AS STRING) AS STRING
'---------- Add/Update a FileList
LOCAL FNum, i, j, k AS LONG
LOCAL fn, t, tlist(), lclfile, MSG, lclMask, tNormList, tNormFile AS STRING
LOCAL RQPath, RQMask, RQFlags, RQNote AS STRING
   MEntry

   '----- Get the existing FILELIST
   fn = ENV.FileListPath + FileList + ".FLIST"                    ' Build the filename
   FNum = FREEFILE                                                ' Load the table
   IF ISFILE(fn) THEN                                             ' Open file if it exists
      OPEN fn FOR INPUT AS #FNum                                  ' Open the FILELIST File
      FILESCAN #FNum, RECORDS TO i                                ' Get the number of records
      REDIM tlist(1 TO MAX(i, 1)) AS STRING                       ' Redim array to match save data
      LINE INPUT #FNum, tlist() TO j                              ' Read it all
      CLOSE #FNum                                                 ' Close it
      FOR i = j TO 1 STEP -1                                      ' Massage it
         tlist(i) = TRIM$(tlist(i))                               ' Clean it up
         IF ISNULL(tlist(i)) THEN                                 ' A null entry?
            ARRAY DELETE tlist(i)                                 ' Remove it from the list
            DECR j                                                ' Adjust count
         END IF                                                   '
         IF INSTR(tlist(i), "|") = 0 THEN                         ' If not yet in the new | format
            REPLACE ANY "," WITH "|" IN tlist(i)                  ' Swap commas to |
         END IF                                                   '
      NEXT i                                                      '
   ELSE                                                           '
      MSG = "File added to New " + FileList + ".FLIST"            ' Build message
      REDIM tlist(1 TO 1) AS STRING                               ' Recreate a table
      tlist(1) = ""                                               ' Add dummy entry
      j = 1                                                       '
   END IF                                                         '

   FOR i = 1 TO MAX(j, 1)                                         ' First see if the file is already here
      IF ISNULL(tlist(i)) THEN                                    ' Dummy marker?
         tlist(i) = NewFile + "|||"                               ' Swap in provided entry
         GOSUB UpdateFile                                         ' Go update the file
         FUNCTION = MSG                                           ' Return Msg
         MExitFunc                                                ' We're done
      END IF                                                      '

      TP.RQSplit(tlist(i), RQPath, RQMask, RQFlags, RQNote)       ' Split entry
      tNormList = UUCASE(RQPath): REPLACE "/" WITH "\" IN tNormList ' Normalize / and \
      tNormFile = UUCASE(NewFile): REPLACE "/" WITH "\" IN tNormFile'
      IF tNormList = tNormFile THEN                               ' Same file?

         IF ISNULL(MSG) THEN MSG = "File previously added to " + FileList + ".FLIST" ' Build message
         IF i <> 1 AND (FileList = "Recent Files" OR FileList = "Recent Paths") THEN ' RECENT/PATHS and not the 1st
            IF FileList = "Recent Paths" THEN RQMask = "*"        ' Normalize mask for Path entries
            ARRAY DELETE tlist(i)                                 ' Delete it at this position
            ARRAY INSERT tlist(), BUILD$(RQPath, "|", RQMask, "|", RQFlags, "|", RQNote) ' Re-insert at the beginning
         END IF                                                   '
         GOSUB UpdateFile                                         ' Go update the file
         FUNCTION = MSG                                           ' Return Msg
         MExitFunc                                                ' We're done
      END IF                                                      '
   NEXT i                                                         '

   INCR j                                                         ' J = count in array
   REDIM PRESERVE tlist(1 TO j) AS STRING                         ' Expand the list by 1
   ARRAY INSERT tlist(), BUILD$(NewFile, "|*")                    ' Insert at the top
   IF ISNULL(MSG) THEN MSG = "File added to " + FileList + ".FLIST"' Build message
   IF (FileList = "Recent Files" OR FileList = "Recent Paths") AND j > ENV.RecentCtr THEN ' Time to trim the RECENT/PATHS table?
      REDIM PRESERVE tlist(1 TO ENV.RecentCtr) AS STRING          ' Redim to the max size
      j = ENV.RecentCtr                                           ' Set J correctly
   END IF                                                         '
   GOSUB UpdateFile:                                              ' Update things
   FUNCTION = MSG                                                 ' Return Msg
   MExitFunc

UpdateFile:
   FNum = FREEFILE
   OPEN fn FOR OUTPUT AS #FNum                                    ' Open the output File
   FOR i = 1 TO MAX(j, 1)                                         ' Write the array back out
      IF ISNOTNULL(tlist(i)) THEN PRINT #FNum, tlist(i)           ' Just non-Null entries
   NEXT i                                                         '
   SETEOF #FNum                                                   '
   CLOSE #FNum                                                    '
   RETURN
END FUNCTION

SUB      sFileListDel(list AS STRING, file AS STRING)
'---------- Delete an entry from a list
LOCAL FNum, c, i, j, k AS LONG
LOCAL t, fn, tlist(), RQPath, RQMask, RQFlags, RQNote AS STRING
   MEntry
   fn = ENV.FileListPath + list + ".FLIST"                        ' Build the filename
   FNum = FREEFILE                                                ' Load the FILELIST
   IF ISFALSE ISFILE(fn) THEN MExitSub                            ' Exit if file not found.  ???
   OPEN fn FOR INPUT AS #FNum                                     ' Open the FILELIST File
   FILESCAN #FNum, RECORDS TO i                                   ' Get the number of records
   REDIM tlist(1 TO i) AS STRING                                  ' Redim array to match save data
   LINE INPUT #FNum, tlist() TO j                                 ' Read it all
   CLOSE #FNum                                                    ' Close it
   c = j                                                          ' Save original number
   FOR i = j TO 1 STEP -1                                         ' Look for file to delete
      tlist(i) = TRIM$(tlist(i))                                  ' Clean it up
      IF ISNULL(tlist(i)) OR _                                    ' A null entry?
         IsEQ(LEFT$(tlist(i), 9), "FILESRCH:") THEN               '
         ARRAY DELETE tlist(i)                                    ' Remove it from the list
         DECR j                                                   ' Adjust count
      END IF                                                      '
      IF INSTR(tlist(i), "|") = 0 THEN                            ' If not yet in the new | format
         REPLACE ANY "," WITH "|" IN tlist(i)                     ' Swap commas to |
      END IF                                                      '
      TP.RQSplit(tlist(i), RQPath, RQMask, RQFlags, RQNote)       ' Split operands
      IF IsEQ(RQPath, file) THEN                                  ' The one to delete?
         ARRAY DELETE tlist(i)                                    ' Delete it
         DECR j                                                   ' Adjust count
         EXIT FOR                                                 ' Exit loop
      END IF                                                      '
   NEXT i                                                         '
   IF c <> j THEN                                                 ' Did we delete something?
      IF j = 0 THEN                                               ' End up empty?
         i = sRecycleBin(fn, "D")                                 ' Delete empty files
         TP.FileListNm = ""                                       ' Just turn it off
      ELSE
         FNum = FREEFILE                                          ' Rewrite the file
         OPEN fn FOR OUTPUT AS #FNum                              ' Open the output File
         FOR i = 1 TO j                                           ' Write the array back out
            PRINT #FNum, tlist(i)                                 '
         NEXT i                                                   '
         SETEOF #FNum                                             '
         CLOSE #FNum                                              '
      END IF                                                      '
   END IF                                                         '
   MExit
END SUB

SUB      sFileListRename(oname AS STRING, nname AS STRING)
'---------- Rename throughout the FILELIST set
LOCAL tfn, fn AS STRING, FD AS DIRDATA, FNum, i, j, OurTab, Found AS LONG, tlist() AS STRING
LOCAL RQPath, RQMask, RQFlags, RQNote AS STRING
   OurTab = TP.PgNumber                                           ' Save our page number
   TP = Tabs(OurTab)                                              ' Switch back to original tab
   MEntry
   tfn = DIR$(ENV.FileListPath + "*.FLIST" TO FD)                 ' Look for our FILELIST files

   '----- Loop through the FILELIST files
   DO WHILE ISNOTNULL(tfn)                                        ' While we're getting entries
      IF (FD.FileAttributes AND %FILE_ATTRIBUTE_DIRECTORY) <> %FILE_ATTRIBUTE_DIRECTORY THEN

         '----- For each FILELIST file look at it's data
         fn = ENV.FileListPath + TRIM$(FD.FileName)               ' Get the filename
         FNum = FREEFILE                                          ' Load the Data
         OPEN fn FOR INPUT AS #FNum                               ' Open the FILELIST File
         FILESCAN #FNum, RECORDS TO i                             ' Get the number of records
         IF i > 0 THEN                                            ' Some records?
            REDIM tlist(1 TO i) AS STRING                         ' Redim array to match save data
            LINE INPUT #FNum, tlist() TO j                        ' Read it all
         END IF                                                   '
         CLOSE #FNum                                              ' Close it
         Found = %False                                           '

         '----- Scan the lines in the file for our rename
         IF j THEN                                                ' Did we get some records?
            FOR i = 1 TO j                                        ' OK, let's scan the records
               IF INSTR(tlist(i), "|") = 0 THEN                   ' If not yet in the new | format
                  REPLACE ANY "," WITH "|" IN tlist(i)            ' Swap commas to |
               END IF                                             '
               TP.RQSplit(tlist(i), RQPath, RQMask, RQFlags, RQNote) ' Split out operands
               IF IsEQ(RQPath, oname) THEN                        ' Found it
                  '----- Found our rename, swap in the new name
                  RQPath = nname                                  ' Swap in the new name
                  tlist(i) = BUILD$(RQPath, "|", RQMask, "|", RQFlags, "|", RQNote)
                  Found = %True                                   ' Say to save it
               END IF                                             '
            NEXT i                                                '
            IF Found THEN                                         ' If we found it
               '----- Re-Write the FILELIST back out
               FNum = FREEFILE                                    ' Write the file back out
               OPEN fn FOR OUTPUT AS #FNum                        ' Open the output File
               PRINT #FNum, tlist()                               ' Dump it back out
               SETEOF #FNum                                       '
               CLOSE #FNum                                        '
            END IF                                                '
         END IF                                                   '
      END IF                                                      '
      tfn = DIR$(NEXT, TO FD)                                     ' Get next FILELIST entry
   LOOP                                                           '
   OurTab = TP.PgNumber                                           ' Save our page number
   TP = Tabs(1)                                                   ' Switch to FM tab
   TP.AttnDo = TP.AttnDo OR %LoadReq                              ' Request reload any FileList
   TP = Tabs(OurTab)                                              ' Switch back to original tab
   MExit                                                          '
END SUB


FUNCTION sFileQueue(func AS STRING, pflag AS STRING, fn AS STRING) AS STRING
'---------- Manage the FileQueue
LOCAL i, j AS LONG
LOCAL lfull AS ASCIIZ * %MAX_PATH
LOCAL lpath AS ASCIIZ * %MAX_PATH
   MEntry
   lfull = fn                                                     ' To fixed length string
   lpath = PATHNAME$(PATH, fn): lpath = LEFT$(lpath, LEN(TRIM$(lpath)) - 1)
   FUNCTION = ""

   GOSUB CheckReset                                               ' See if table can be reset

   SELECT CASE AS CONST$ func                                     ' Split by transaction type
      CASE "A"                                                    ' Add
         FOR i = 1 TO UBOUND(gFQ())                               ' Search for an empty slot
            IF ISFALSE gFQ(i).gInUse THEN                         ' Is this available?
               gFQ(i).gWatchDir = lpath                           ' Add the entry
               gFQ(i).gWatchFile = lfull                          '
               gFQ(i).gPgNumber = TP.PgNumber                     '
               gFQ(i).gInUse = %True                              '
               gFQ(i).gFlag = " "                                 '
               MExitFunc                                          ' We're done
            END IF                                                '
         NEXT i                                                   '

      CASE "D"                                                    ' Del
         FOR i = 1 TO UBOUND(gFQ())                               ' Search for it
            IF ISTRUE gFQ(i).gInUse AND gFQ(i).gWatchFile = lfull THEN ' Is this the one?
               gFQ(i).gInUse = %False                             ' Mark it no longer in use
               MExitFunc                                          ' We're done
            END IF                                                '
         NEXT i                                                   '

      CASE "F"                                                    ' Flag
         FOR i = 1 TO UBOUND(gFQ())                               ' Search for it
            IF ISTRUE gFQ(i).gInUse AND gFQ(i).gWatchFile = lfull THEN ' Is this the one?
               IF gFQ(i).gflag = " " THEN                         ' Still blank?
                  gFQ(i).gflag = pflag                            ' Yes, update it
                  FUNCTION = pflag                                ' Return the flag as an OK
                  MExitFunc                                       ' We're done
               ELSE                                               ' Hmmm, already flagged?
                  MExitFunc                                       ' Exit with the default ""
               END IF                                             '
            END IF                                                '
         NEXT i                                                   '

      CASE "G"                                                    ' Get Flag
         FOR i = 1 TO UBOUND(gFQ())                               ' Search for it
            IF ISTRUE gFQ(i).gInUse AND gFQ(i).gWatchFile = lfull THEN ' Is this the one?
               FUNCTION = gFQ(i).gflag                            ' Yes, pass back the flag
               MExitFunc                                          ' We're done
            END IF                                                '
         NEXT i                                                   '

      CASE "I"                                                    ' Search for index
         FOR i = 1 TO UBOUND(gFQ())                               ' Search for it
            IF ISTRUE gFQ(i).gInUse AND gFQ(i).gWatchFile = lfull THEN ' Is this the one?
               FUNCTION = FORMAT$(i)                              ' Yes, pass back the index
               MExitFunc                                          ' We're done
            END IF                                                '
            FUNCTION = "0"                                        ' Return zero (fail)
         NEXT i                                                   '

      CASE "S"                                                    ' Search
         FUNCTION = "0"                                           ' Set not found return
         FOR i = 1 TO UBOUND(gFQ())                               ' Search for it
            IF ISTRUE gFQ(i).gInUse AND UUCASE(gFQ(i).gWatchFile) = UUCASE(lfull) THEN ' Is this the one?
               FUNCTION = FORMAT$(gFQ(i).gPgNumber)               ' Yes, pass back the tab numberflag
               MExitFunc                                          ' We're done
            END IF                                                '
         NEXT i                                                   '

      CASE "T"                                                    ' Adjust Tab numbers
         j = VAL(pflag)                                           ' Get deleted tab # from flag parameter
         FOR i = 1 TO UBOUND(gFQ())                               ' Search for it
            IF ISTRUE gFQ(i).gInUse AND gFQ(i).gPgNumber >= j THEN' In a tab higher than the deleted one?
               gFQ(i).gPgNumber = gFQ(i).gPgNumber - 1            ' Yes, reduce it's number by one
            END IF                                                '
         NEXT i                                                   '
   END SELECT
   MExitFunc

CheckReset:
   FOR i = 1 TO UBOUND(gFQ())                                     ' Search for it
      IF ISTRUE gFQ(i).gInUse THEN RETURN                         ' Any active entry, just return
   NEXT i                                                         '

   '----- There are no active entries, we can do a reset
   RESET gFQ()                                                    ' Clear the FQ table
   RETURN
END FUNCTION

THREAD FUNCTION sFileWatchThread(BYVAL tData AS LONG) AS LONG     ' Monitor a directory, stop when caller tells us via Event
THREADED T1pData AS WatchData POINTER, T1hSearch, T1Posted, T1i AS LONG
THREADED T1FD         AS DIRDATA, ttxt AS STRING                  '
THREADED T1KMsg AS kbMsg                                          '
DIM WaitList(0 TO 1) AS THREADED LONG
THREADED ActiveIn, ActiveOut AS LONG                              '
   T1pData = tData                                                ' Get address of our parameters
   @T1pData.gChanged = %False                                     ' Start off as %False
   @T1pData.gActive = %False                                      ' Start off as %False
   ActiveIn = VARPTR(@T1pData.gActive)
   WaitList(0) = @t1pData.gEvent                                  ' Copy Event address to our WaitList(0)
   WaitList(1) = FindFirstChangeNotification(@T1pData.gWatchDir, 0, _ ' Put FindFCN into WaitList(1)
                     %FILE_NOTIFY_CHANGE_FILE_NAME   OR _         '
                     %FILE_NOTIFY_CHANGE_ATTRIBUTES  OR _         '
                     %FILE_NOTIFY_CHANGE_SIZE        OR _         '
                     %FILE_NOTIFY_CHANGE_LAST_WRITE)              '
   IF WaitList(1) = %INVALID_HANDLE_VALUE THEN                    ' Should never happen, but ...
      FUNCTION = 8: EXIT FUNCTION                                 ' Tell mainline we couldn't start
   END IF                                                         '
   @T1pData.gActive = %True                                       ' Say we're active

   DO
      T1Posted = WaitForMultipleObjects(2, WaitList(0), 0, %INFINITE)  ' Sleep till Windows Posts us
      SELECT CASE T1Posted                                        ' See who Posted us
         CASE %WAIT_OBJECT_0                                      ' Being told to stop by the MainLine?
            CloseHandle(WaitList(0))                              ' Close the Event handle
            FindCloseChangeNotification(WaitList(1))              ' Close the notification object
            EXIT LOOP                                             ' Fall out the bottom

         CASE %WAIT_OBJECT_0 + 1                                  ' Is this the FxCN?
            SLEEP 100                                             ' Wait 100 ms for Windows multiple timestamps to be done
            GOSUB CheckFileDates                                  ' See if our file's status has changed
            IF ISFALSE FindNextChangeNotification(WaitList(1)) THEN ' Tell FNCN to keep looking
               FUNCTION = 8: @T1pData.gActive = %False: EXIT FUNCTION ' Some kind of error, bail out to mainline
            END IF                                                '

         CASE ELSE                                                ' Anything else should not happen
            FUNCTION = 8: @T1pData.gActive = %False: EXIT FUNCTION' Some kind of error, bail out to mainline
      END SELECT                                                  '
   LOOP                                                           '
   ActiveOut = VARPTR(@T1pData.gActive)
   @T1pData.gActive = %False                                      ' Say we're no longer active
   @T1pData.gInUse = %False                                       ' Say the FQ entry is no longer in use
   FUNCTION = 0                                                   '
   EXIT FUNCTION                                                  '

CheckFileDates:
   T1hSearch = GetFileAttributesEx(@T1pData.gWatchFile, &H0, BYREF T1FD)  ' Search for the filename
   IF T1hSearch = 0 THEN                                          ' No file, error
      @T1pData.gChanged = 1                                       ' Say file has been deleted
      GOSUB NotifyMainLine                                        ' Send back a message
   ELSE                                                           ' Else file still exists
      IF @T1pData.gFileLWTime <> T1FD.LastWriteTime OR _          ' See if anything has changed
         @T1pData.gFileCRTime <> T1FD.CreationTime OR _           '
         @T1pData.gFileSizeHigh <> T1FD.FileSizeHigh OR _         '
         @T1pData.gFileSizeLow <> T1FD.FileSizeLow OR _           '
         @T1pData.gFileAttrib <> T1FD.FileAttributes THEN         '
         @T1pData.gChanged = -1                                   ' Say data has changed
         @T1pData.gFileLWTime = T1FD.LastWriteTime                ' Save for the next time
         @T1pData.gFileCRTime = T1FD.CreationTime                 '
         @T1pData.gFileSizeHigh = T1FD.FileSizeHigh               '
         @T1pData.gFileSizeLow = T1FD.FileSizeLow                 '
         @T1pData.gFileAttrib = T1FD.FileAttributes               '
         GOSUB NotifyMainLine                                     ' Send back a message
      END IF                                                      '
   END IF                                                         '
   RETURN                                                         '

NotifyMainLine:
   MID$(T1KMsg.kbString, 1, 1) = CHR$(2)                          ' Flag 1st byte as Hex 2
   MID$(T1KMsg.kbString, 2, 1) = IIF$(@T1pData.gChanged = 1, "D", "M") ' Flag 2nd byte as "D" or "M"
   t1KMsg.kbInt(1) = @t1pData.gPgNumber                           ' Copy Tab # to be notified
   ttxt = TRIM$(@T1pData.gWatchFile)                              ' Convert to normal string
   IF sFileQueue("F", IIF$(@T1pData.gChanged = 1, "D", "M"), ttxt) = "" THEN RETURN ' Set flag in FQ area, exit if already posted
   t1i = PostMessage(hWnd, %WM_USER, t1KMsg.MsgwParam, 0)         ' To the mainline Callback routine
   RETURN

END FUNCTION


'/-----------------------------------------------------------------------------
'/ FUNCTION sFindLineNum(ln AS LONG) AS LONG
'/ '---------- Return IX of a specified visible line number, 0 if not found
'/ REGISTER i AS LONG
'/ LOCAL t AS STRING
'/    MEntry
'/    t = FORMAT$(ln, "00000000")                                    ' Create search arg as Text
'/    IF t > TP.LLNumGet(TP.LastLine - 1) THEN                       ' If a super big number
'/       sFindLineNum = 0 - (TP.LastLine - 1)                        ' Pass back the last line as negative
'/       MExitFunc                                                   ' and return
'/    END IF
'/    i = TP.LLNumScan(t)                                            '
'/    sFindLineNum = IIF(i = 0, 0, i - 1)                            ' Pass back result
'/    MExit
'/ END FUNCTION
'/-----------------------------------------------------------------------------


FUNCTION sFindLineNum(ln AS LONG) AS LONG
'---------- Return IX of a specified visible line number, 0 if not found
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL t AS STRING
   MEntry
   IF TP.LastLine = 2 THEN sFindLineNum = 0: MExitFunc            ' If no lines, then not found
   j = TP.LastLine - 1                                            ' Get Last data line
   DO WHILE j > 1 AND ISFALSE TP.LFlagData(j)                     ' If last line isn't data line, find it
      DECR j                                                      ' Backup
   LOOP                                                           '
   IF j = 1 THEN sFindLineNum = 0: MExitFunc                      ' If just special lines, then not found

   t = FORMAT$(ln, "00000000")                                    ' Create search arg as Text
   IF t > TP.LLNumGet(j) THEN                                     ' If a super big number
      sFindLineNum = 0 - (j)                                      ' Pass back the last line as negative
      MExitFunc                                                   ' and return
   END IF
   i = TP.LLNumScan(t)                                            '
   sFindLineNum = IIF(i = 0, 0, i - 1)                            ' Pass back result
   MExit
END FUNCTION


FUNCTION sFindWindow(BYVAL lTitle AS STRING) AS LONG
'---------- Find a window by title
REGISTER Ctr AS LONG
   gWinListPos = 0                                                ' Init
   EnumWindows CODEPTR(sEnumWindowsProc), 0                       ' Count windows
   lTitle = UUCASE(lTitle)                                        ' Uppercase request
   FOR Ctr = 1 TO gWinListPos                                     ' Loop through list
      IF IsWindowVisible(gWinList(Ctr).WinHandle) = 1 THEN        ' Visible one?
         IF INSTR(UUCASE(gWinList(Ctr).WinTitle), lTitle) THEN    ' Yes, contain our string?
            FUNCTION = gWinList(Ctr).WinHandle                    ' Yes, pass back its handle
            EXIT FOR                                              '
         END IF                                                   '
      END IF                                                      '
   NEXT Ctr                                                       ' Loop back
END FUNCTION

FUNCTION sFMSortFlag(P1Flag AS LONG, P2Flag AS LONG) AS LONG
'---------- Handle the Dir sorting fudge
   IF TP.DirSort = "Dir+" THEN                                    ' Old style?
      IF P1Flag < P2Flag THEN FUNCTION = -1: EXIT FUNCTION        ' Do old style compare
      IF P1Flag > P2Flag THEN FUNCTION = +1: EXIT FUNCTION        '

   ELSEIF TP.DirSort = "Dir-" THEN                                ' If put Dirs last?
      IF P1Flag = %FDirDown THEN P1Flag = %FDirLow                ' Swap to bottom keys
      IF P2Flag = %FDirDown THEN P2Flag = %FDirLow                '
      IF P1Flag < P2Flag THEN FUNCTION = -1: EXIT FUNCTION        ' Do old style compare
      IF P1Flag > P2Flag THEN FUNCTION = +1: EXIT FUNCTION        '

   ELSE                                                           ' Else in-line
      IF P1Flag = %FDirUp THEN FUNCTION = -1: EXIT FUNCTION       ' Force Dir up to the top
      IF P1Flag = %FDirDown AND P2Flag >= %FEntry AND P2Flag <= %FProfile OR _
         P2Flag = %FDirDown AND P1Flag >= %FEntry AND P1Flag <= %FProfile THEN
         ' Treat as equal
      ELSE
         IF P1Flag < P2Flag THEN FUNCTION = -1: EXIT FUNCTION     ' Do old style compare
         IF P1Flag > P2Flag THEN FUNCTION = +1: EXIT FUNCTION     '
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION sFMSortDateUp(p1 AS FMFList, p2 AS FMFList) AS LONG
'---------- Support sort of the FMFiles array
LOCAL P1Flag, P2Flag, FTest AS LONG
   P1Flag = p1.Flag: P2Flag = p2.Flag                             ' Get a working copy of flags
   FTest = sFMSortFlag(P1Flag, P2Flag)                            ' Do test in common code
   IF FTest <> 0 THEN FUNCTION = FTest:EXIT FUNCTION              ' If < or > result, pass it back

   IF P1Flag = %FDirDown THEN
      IF UUCASE(p1.FD.Filename) < UUCASE(p2.FD.FileName) THEN FUNCTION = +1: EXIT FUNCTION
      IF UUCASE(p1.FD.Filename) > UUCASE(p2.FD.FileName) THEN FUNCTION = -1: EXIT FUNCTION
      FUNCTION = 0
      EXIT FUNCTION
   END IF
   IF UUCASE(p1.LWTime) > UUCASE(p2.LWTime) THEN FUNCTION = +1: EXIT FUNCTION
   IF UUCASE(p1.LWTime) < UUCASE(p2.LWTime) THEN FUNCTION = -1: EXIT FUNCTION
   FUNCTION = 0
END FUNCTION

FUNCTION sFMSortDateDown(p1 AS FMFList, p2 AS FMFList) AS LONG
'---------- Support sort of the FMFiles array
LOCAL P1Flag, P2Flag, FTest AS LONG
   P1Flag = p1.Flag: P2Flag = p2.Flag                             ' Get a working copy of flags
   FTest = sFMSortFlag(P1Flag, P2Flag)                            ' Do test in common code
   IF FTest <> 0 THEN FUNCTION = FTest:EXIT FUNCTION              ' If < or > result, pass it back

   IF P1Flag = %FDirDown THEN
      IF UUCASE(p1.FD.Filename) < UUCASE(p2.FD.FileName) THEN FUNCTION = +1: EXIT FUNCTION
      IF UUCASE(p1.FD.Filename) > UUCASE(p2.FD.FileName) THEN FUNCTION = -1: EXIT FUNCTION
      FUNCTION = 0
      EXIT FUNCTION
   END IF
   IF UUCASE(p1.LWTime) < UUCASE(p2.LWTime) THEN FUNCTION = +1: EXIT FUNCTION
   IF UUCASE(p1.LWTime) > UUCASE(p2.LWTime) THEN FUNCTION = -1: EXIT FUNCTION
   FUNCTION = 0
END FUNCTION

FUNCTION sFMSortExtUp(p1 AS FMFList, p2 AS FMFList) AS LONG
'---------- Support sort of the FMFiles array
LOCAL P1Flag, P2Flag, FTest AS LONG
   P1Flag = p1.Flag: P2Flag = p2.Flag                             ' Get a working copy of flags
   FTest = sFMSortFlag(P1Flag, P2Flag)                            ' Do test in common code
   IF FTest <> 0 THEN FUNCTION = FTest:EXIT FUNCTION              ' If < or > result, pass it back

   IF p1.Ext < p2.Ext THEN FUNCTION = -1: EXIT FUNCTION
   IF p1.Ext > p2.Ext THEN FUNCTION = +1: EXIT FUNCTION
   IF UUCASE(p1.FD.Filename) < UUCASE(p2.FD.FileName) THEN FUNCTION = -1: EXIT FUNCTION
   IF UUCASE(p1.FD.Filename) > UUCASE(p2.FD.FileName) THEN FUNCTION = +1: EXIT FUNCTION
   FUNCTION = 0
END FUNCTION

FUNCTION sFMSortExtDown(p1 AS FMFList, p2 AS FMFList) AS LONG
'---------- Support sort of the FMFiles array
LOCAL P1Flag, P2Flag, FTest AS LONG
   P1Flag = p1.Flag: P2Flag = p2.Flag                             ' Get a working copy of flags
   FTest = sFMSortFlag(P1Flag, P2Flag)                            ' Do test in common code
   IF FTest <> 0 THEN FUNCTION = FTest:EXIT FUNCTION              ' If < or > result, pass it back

   IF p1.Ext < p2.Ext THEN FUNCTION = +1: EXIT FUNCTION
   IF p1.Ext > p2.Ext THEN FUNCTION = -1: EXIT FUNCTION
   IF UUCASE(p1.FD.Filename) < UUCASE(p2.FD.FileName) THEN FUNCTION = -1: EXIT FUNCTION
   IF UUCASE(p1.FD.Filename) > UUCASE(p2.FD.FileName) THEN FUNCTION = +1: EXIT FUNCTION
   FUNCTION = 0
END FUNCTION

FUNCTION sFMSortFileUp(p1 AS FMFList, p2 AS FMFList) AS LONG
'---------- Support sort of the FMFiles array
LOCAL P1Fn, P2Fn AS STRING
LOCAL P1Flag, P2Flag, FTest AS LONG
   P1Flag = p1.Flag: P2Flag = p2.Flag                             ' Get a working copy of flags
   FTest = sFMSortFlag(P1Flag, P2Flag)                            ' Do test in common code
   IF FTest <> 0 THEN FUNCTION = FTest:EXIT FUNCTION              ' If < or > result, pass it back

   IF TP.FileListNm = "" THEN
      P1Fn = TRIM$(UUCASE(p1.Path+p1.FD.Filename)): P2Fn = TRIM$(UUCASE(p2.Path+p2.FD.FileName))
   ELSE
      P1Fn = TRIM$(UUCASE(p1.FD.Filename)): P2Fn = TRIM$(UUCASE(p2.FD.FileName))
   END IF
   IF RIGHT$(P1Fn, 1) = "\" THEN P1Fn = LEFT$(P1Fn, LEN(P1Fn) - 1)
   IF RIGHT$(P2Fn, 1) = "\" THEN P2Fn = LEFT$(P2Fn, LEN(P2Fn) - 1)
   IF P1Fn < P2Fn THEN FUNCTION = -1: EXIT FUNCTION
   IF P1Fn > P2Fn THEN FUNCTION = +1: EXIT FUNCTION
   FUNCTION = 0
END FUNCTION

FUNCTION sFMSortFileDown(p1 AS FMFList, p2 AS FMFList) AS LONG
'---------- Support sort of the FMFiles array
LOCAL P1Fn, P2Fn AS STRING
LOCAL P1Flag, P2Flag, FTest AS LONG
   P1Flag = p1.Flag: P2Flag = p2.Flag                             ' Get a working copy of flags
   FTest = sFMSortFlag(P1Flag, P2Flag)                            ' Do test in common code
   IF FTest <> 0 THEN FUNCTION = FTest:EXIT FUNCTION              ' If < or > result, pass it back

   IF TP.FileListNm = "" THEN
      P1Fn = TRIM$(UUCASE(p1.Path+p1.FD.Filename)): P2Fn = TRIM$(UUCASE(p2.Path+p2.FD.FileName))
   ELSE
      P1Fn = TRIM$(UUCASE(p1.FD.Filename)): P2Fn = TRIM$(UUCASE(p2.FD.FileName))
   END IF
   IF RIGHT$(P1Fn, 1) = "\" THEN P1Fn = LEFT$(P1Fn, LEN(P1Fn) - 1)
   IF RIGHT$(P2Fn, 1) = "\" THEN P2Fn = LEFT$(P2Fn, LEN(P2Fn) - 1)
   IF P1Fn < P2Fn THEN FUNCTION = +1: EXIT FUNCTION
   IF P1Fn > P2Fn THEN FUNCTION = -1: EXIT FUNCTION
   FUNCTION = 0
END FUNCTION

FUNCTION sFMSortSizeUp(p1 AS FMFList, p2 AS FMFList) AS LONG
'---------- Support sort of the FMFiles array
LOCAL P1Flag, P2Flag, FTest AS LONG
   P1Flag = p1.Flag: P2Flag = p2.Flag                             ' Get a working copy of flags
   FTest = sFMSortFlag(P1Flag, P2Flag)                            ' Do test in common code
   IF FTest <> 0 THEN FUNCTION = FTest:EXIT FUNCTION              ' If < or > result, pass it back

   IF P1Flag = %FDirDown THEN
      IF UUCASE(p1.FD.Filename) < UUCASE(p2.FD.FileName) THEN FUNCTION = -1: EXIT FUNCTION
      IF UUCASE(p1.FD.Filename) > UUCASE(p2.FD.FileName) THEN FUNCTION = +1: EXIT FUNCTION
      FUNCTION = 0
      EXIT FUNCTION
   END IF
   IF MAK(QUAD, p1.FD.FileSizeLow, p1.FD.FileSizeHigh) < MAK(QUAD, p2.FD.FileSizeLow, p2.FD.FileSizeHigh) THEN FUNCTION = -1: EXIT FUNCTION
   IF MAK(QUAD, p1.FD.FileSizeLow, p1.FD.FileSizeHigh) > MAK(QUAD, p2.FD.FileSizeLow, p2.FD.FileSizeHigh) THEN FUNCTION = +1: EXIT FUNCTION
   FUNCTION = 0
END FUNCTION

FUNCTION sFMSortSizeDown(p1 AS FMFList, p2 AS FMFList) AS LONG
'---------- Support sort of the FMFiles array
LOCAL P1Flag, P2Flag, FTest AS LONG
   P1Flag = p1.Flag: P2Flag = p2.Flag                             ' Get a working copy of flags
   FTest = sFMSortFlag(P1Flag, P2Flag)                            ' Do test in common code
   IF FTest <> 0 THEN FUNCTION = FTest:EXIT FUNCTION              ' If < or > result, pass it back

   IF P1Flag = %FDirDown THEN
      IF UUCASE(p1.FD.Filename) < UUCASE(p2.FD.FileName) THEN FUNCTION = -1: EXIT FUNCTION
      IF UUCASE(p1.FD.Filename) > UUCASE(p2.FD.FileName) THEN FUNCTION = +1: EXIT FUNCTION
      FUNCTION = 0
      EXIT FUNCTION
   END IF
   IF MAK(QUAD, p1.FD.FileSizeLow, p1.FD.FileSizeHigh) < MAK(QUAD, p2.FD.FileSizeLow, p2.FD.FileSizeHigh) THEN FUNCTION = +1: EXIT FUNCTION
   IF MAK(QUAD, p1.FD.FileSizeLow, p1.FD.FileSizeHigh) > MAK(QUAD, p2.FD.FileSizeLow, p2.FD.FileSizeHigh) THEN FUNCTION = -1: EXIT FUNCTION
   FUNCTION = 0
END FUNCTION

FUNCTION sFMSortLinesUp(p1 AS FMFList, p2 AS FMFList) AS LONG
'---------- Support sort of the FMFiles array
LOCAL P1Flag, P2Flag, FTest AS LONG
   P1Flag = p1.Flag: P2Flag = p2.Flag                             ' Get a working copy of flags
   FTest = sFMSortFlag(P1Flag, P2Flag)                            ' Do test in common code
   IF FTest <> 0 THEN FUNCTION = FTest:EXIT FUNCTION              ' If < or > result, pass it back

   IF P1Flag = %FDirDown THEN
      IF UUCASE(p1.FD.Filename) < UUCASE(p2.FD.FileName) THEN FUNCTION = -1: EXIT FUNCTION
      IF UUCASE(p1.FD.Filename) > UUCASE(p2.FD.FileName) THEN FUNCTION = +1: EXIT FUNCTION
      FUNCTION = 0
      EXIT FUNCTION
   END IF
   IF p1.LinesInt < p2.LinesInt THEN FUNCTION = -1: EXIT FUNCTION
   IF p1.LinesInt > p2.LinesInt THEN FUNCTION = +1: EXIT FUNCTION
   FUNCTION = 0
END FUNCTION

FUNCTION sFMSortLinesDown(p1 AS FMFList, p2 AS FMFList) AS LONG
'---------- Support sort of the FMFiles array
LOCAL P1Flag, P2Flag, FTest AS LONG
   P1Flag = p1.Flag: P2Flag = p2.Flag                             ' Get a working copy of flags
   FTest = sFMSortFlag(P1Flag, P2Flag)                            ' Do test in common code
   IF FTest <> 0 THEN FUNCTION = FTest:EXIT FUNCTION              ' If < or > result, pass it back

   IF P1Flag = %FDirDown THEN
      IF UUCASE(p1.FD.Filename) < UUCASE(p2.FD.FileName) THEN FUNCTION = -1: EXIT FUNCTION
      IF UUCASE(p1.FD.Filename) > UUCASE(p2.FD.FileName) THEN FUNCTION = +1: EXIT FUNCTION
      FUNCTION = 0
      EXIT FUNCTION
   END IF
   IF p1.LinesInt < p2.LinesInt THEN FUNCTION = +1: EXIT FUNCTION
   IF p1.LinesInt > p2.LinesInt THEN FUNCTION = -1: EXIT FUNCTION
   FUNCTION = 0
END FUNCTION

FUNCTION sFMSortNoteUp(p1 AS FMFList, p2 AS FMFList) AS LONG
'---------- Support sort of the FMFiles array
LOCAL P1Flag, P2Flag, FTest AS LONG
   P1Flag = p1.Flag: P2Flag = p2.Flag                             ' Get a working copy of flags
   FTest = sFMSortFlag(P1Flag, P2Flag)                            ' Do test in common code
   IF FTest <> 0 THEN FUNCTION = FTest:EXIT FUNCTION              ' If < or > result, pass it back

   IF p1.Note < p2.Note THEN FUNCTION = -1: EXIT FUNCTION
   IF p1.Note > p2.Note THEN FUNCTION = +1: EXIT FUNCTION
   IF UUCASE(p1.FD.Filename) < UUCASE(p2.FD.FileName) THEN FUNCTION = -1: EXIT FUNCTION
   IF UUCASE(p1.FD.Filename) > UUCASE(p2.FD.FileName) THEN FUNCTION = +1: EXIT FUNCTION
   FUNCTION = 0
END FUNCTION

FUNCTION sFMSortNoteDown(p1 AS FMFList, p2 AS FMFList) AS LONG
'---------- Support sort of the FMFiles array
LOCAL P1Flag, P2Flag, FTest AS LONG
   P1Flag = p1.Flag: P2Flag = p2.Flag                             ' Get a working copy of flags
   FTest = sFMSortFlag(P1Flag, P2Flag)                            ' Do test in common code
   IF FTest <> 0 THEN FUNCTION = FTest:EXIT FUNCTION              ' If < or > result, pass it back

   IF p1.Note < p2.Note THEN FUNCTION = +1: EXIT FUNCTION
   IF p1.Note > p2.Note THEN FUNCTION = -1: EXIT FUNCTION
   IF UUCASE(p1.FD.Filename) < UUCASE(p2.FD.FileName) THEN FUNCTION = -1: EXIT FUNCTION
   IF UUCASE(p1.FD.Filename) > UUCASE(p2.FD.FileName) THEN FUNCTION = +1: EXIT FUNCTION
   FUNCTION = 0
END FUNCTION

SUB      sGblOptSet()
'---------- Correct all tabs after an OPTION command
LOCAL i, j AS LONG, t AS STRING
   MEntry
   j = TP.PgNumber                                                ' Save where we are
   FOR i = 1 TO TabsNum                                           '
      TP = Tabs(i)                                                ' Swap to tabs data area
      TP.PrfWordVal()                                             ' Process WordInput into Word again
      TP.PicSetAll                                                ' Say total initialization Word has changed
      TP.PicInit                                                  ' Re-Initialize Picture control area
      sSetupSB                                                    '
   NEXT i                                                         '
   TP = Tabs(j)                                                   ' Go back to initial tab
   MExit
END SUB

SUB      sGblSaveAll(Cond AS LONG)
'---------- Do the save all
LOCAL CurrTab, i AS LONG
   CurrTab = TP.PgNumber                                          ' Save where we are
   FOR i = 1 TO TabsNum                                           ' Do for each tab
      TP = Tabs(i)                                                ' Pick the Tab
      IF ISFALSE IsFMTab THEN                                     ' If not FM
         IF (TP.TMode AND (%MClip OR %MSetEdit OR %MBrowse OR %MView)) THEN ITERATE FOR  ' Don't do certain types
         IF Cond THEN                                             ' Conditional?
            IF IsTPModdFlag THEN                                  ' If modified
               TP.TabTitleSet(%True)                              '
               IF IsMEdit THEN                                    '
                  pCmdSave("SAVE MEditOnly ")                     '
               ELSE                                               '
                  pCmdSave("SAVE")                                '
               END IF                                             '
               TP.DispScreen                                      '
            END IF                                                '
         ELSE                                                     '
            TP.TabTitleSet(%True)                                 '
            pCmdSave("SAVE")                                      '
            TP.DispScreen                                         '
         END IF                                                   '
      END IF                                                      '
   NEXT i                                                         '
   TP = Tabs(CurrTab)                                             ' Go back to initial tab
   IF ISFALSE gMacroMode THEN                                     ' If not macro mode
      TAB SELECT hWnd, %IDC_SPFLiteTAB, TP.PgNumber               ' Select the new tab
   END IF                                                         '
END SUB

FUNCTION sGetDefDir() AS STRING
'---------- Return the default Directory to use
LOCAL sDir AS STRING
   MEntry
   IF TP.TIPFilePath = $Empty OR INSTR(TP.TIPFilePath, "\") = 0 THEN ' If there's no active file with a path
      sDir = TP.FPath                                             ' Yes, pass back what FM is using.
   ELSE                                                           ' We have an active file
      sDir = TP.TIPPath                                           ' So pass back its path
   END IF                                                         '
   IF ISNULL(sDir) THEN sDir = CURDIR$ + "\"                      ' Just in case
   FUNCTION = sDir                                                ' Pass back the string
   MExit
END FUNCTION

FUNCTION sGetDropFiles(BYVAL hDropParam AS DWORD) AS STRING
'---------- Return list of Drag/Drop filenames
LOCAL sDropFiles AS STRING, sText AS STRING, i AS LONG
   MEntry
   FOR i = 0 TO DragQueryFile(hDropParam, &HFFFFFFFF&, "", 0) - 1 ' Loop through passed entries
      sText = SPACE$(DragQueryFile(hDropParam, i, "", 0) + 1)     '
      DragQueryFile hDropParam, i, BYVAL STRPTR(sText), LEN(sText)'
      sText = LEFT$(sText, LEN(sText) - 1)                        ' Drop last byte
      IF IsEQ(RIGHT$(sText, 4), ".LNK") THEN sText = sLNKConvert(sText) ' Convert any .LNKs to filenames
      IF ISNOTNULL(sText) THEN sDropFiles = sDropFiles + sText + "|"   ' Add filename now to string
   NEXT i                                                         '
   DragFinish hDropParam                                          ' Finish off the Drag param
   FUNCTION = sDropFiles                                          ' Pass back the string
   MExit
END FUNCTION

FUNCTION sGetFnClipboard() AS STRING
'---------- Get a filename from the Clipboard
LOCAL fn AS STRING
   MEntry

   sWinclip_get(fn)                                               ' Get any current text to start

   fn = PARSE$(fn, $CRLF, 1)                                      ' Get 1st 'line' from the data
   sUnQuote(fn)                                                   ' Strip off any quotes if present
   FUNCTION = fn                                                  ' Strip off any quotes if present
END FUNCTION

FUNCTION sGetIX(lrow AS LONG) AS LONG
'---------- Get and validate an S() array pointer
REGISTER i AS LONG
   i = TP.SGet(lrow)                                              ' Get the L() index from S()
   FUNCTION = IIF(i > TP.LastLine, 0, i)                          ' If not in range, return zero
END FUNCTION

FUNCTION sGetLines(fName AS STRING) AS LONG
'---------- Get # lines via the STATE data
LOCAL pName, lclFn, sLine AS STRING, fNum AS LONG

   '----- Get extension from the filename
   pName = sParseProfile(fName)                                   ' Extract the extension

   '----- See if STATE active for this Profile
   IF ISFALSE sProfState("FETCH", pName) THEN                     ' If STATE=OFF for this Profile
      FUNCTION = -1: EXIT FUNCTION                                ' Then pass back -1 for N/A
   END IF

   '----- See if a STATE file exists
   lclFn = fName + ".STATE"                                       ' Temp copy with .STATE on the end
   REPLACE ANY ":\/" WITH "```" IN lclFN                          ' Make : / and \ into `
   lclFn = ENV.StatePath + lclFn                                  ' Add our STATE folder
   IF ISFALSE ISFILE(lclFn) THEN                                  ' If no STATE file
      FUNCTION = -1: EXIT FUNCTION                                ' Then pass back -1 for N/A
   END IF                                                         '
   FNum = FREEFILE                                                ' Open the file
   OPEN lclFn FOR INPUT AS #FNum                                  ' Open the STATE File
   LINE INPUT #FNum, sLine                                        ' Read 1st line
   CLOSE #FNum                                                    ' Close it

   '----- Get line count from header line
   IF LEFT$(sLine, 2) <> "#1" THEN                                ' Better be the header line
      FUNCTION = -1: EXIT FUNCTION                                ' If not, then pass back -1 for N/A
   END IF                                                         '
   FUNCTION = VAL(PARSE$(sLine, ",", 2))                          ' Extract the # lines to return
END FUNCTION

FUNCTION sGetNewTempFile(pfx AS STRING) AS STRING
'---------- Get a Temp file allocated
LOCAL lpTempFileName    AS ASCIIZ * %MAX_PATH
LOCAL lpShortPathName   AS ASCIIZ * %MAX_PATH
LOCAL lpszPrefix        AS ASCIIZ * 20
LOCAL r                 AS LONG, fn AS STRING
   MEntry
   lpszPrefix = pfx                                               ' Get Prefix to ASCIIZ
   r = GetTempFileName(sGetWindowsTempDir, lpszPrefix, 0, lpTempFileName)
   IF r = 0 THEN                                                  ' If Failed, then try current folder
      r = GetTempFileName(".", lpszPrefix, 0, lpTempFileName)     '
      IF r = 0 OR ISFALSE ISFILE(lpTempFileName) THEN             ' If failed again or can't find the file            '
         lpTempFileName = pfx + FORMAT$(TIMER) + ".tmp"           ' Adjust filename ourselves
         fn = lpTempFileName                                      ' Make normal string for sMakeNullFile
         sMakeNullFile(fn)                                        ' Try one  last time
         MSGBOX "GetTempFileName failed twice, allocated manually"' Tell user
      END IF                                                      '
   END IF                                                         '
   fn = sGetShortName(lpTempFileName)                             ' Shorten the name
   FUNCTION = fn                                                  '
   MExit
END FUNCTION

FUNCTION sGet_Set_number(BYVAL var_name AS STRING) AS LONG
'/-----------------------------------------------------------------------------
'/ get_Set_number
'/
'/ fetch the value of named SET symbol, and convert to numeric.
'/ if the name is undefined, return 0.
'/-----------------------------------------------------------------------------
LOCAL set_symbol AS STRING
   set_symbol = sSetTable ("GET", var_name)
   FUNCTION = IIF(LEN(set_symbol) > 1 AND LEFT$(set_symbol, 1) = "0", VAL(MID$(set_symbol, 2)), 0)
END FUNCTION ' get_Set_number

FUNCTION sGetShortName(BYREF longname AS ASCIIZ) AS STRING
'---------- Convert long name to short name
LOCAL shortname AS ASCIIZ * %MAX_PATH
   MEntry
   IF GetShortPathName(longname, shortname, %MAX_PATH) THEN       ' Shorten the name
      FUNCTION = shortname                                        ' If OK, pass back the short name
   ELSE                                                           '
      FUNCTION = longname                                         ' Else pass the long name
   END IF                                                         '
   MExit
END FUNCTION

FUNCTION sGetVerFile() AS STRING
'---------- Get SPFLite.VER file from the website
LOCAL VerData AS STRING, URLPath1, URLPath2 AS ASCIIZ * %MAX_PATH, LocalPath AS ASCIIZ * %MAX_PATH
LOCAL iResult AS LONG
   MEntry
   URLPath1 = "http://www.spflite.com/files/SPFLite.ver"
   LocalPath = ENV.INIPath + "SPFLite.ver"                        '

   '----- Check 1st URL address
   iResult = DeleteURLCacheEntry(URLPath1)                        ' 1 = success  clear the cache
   IF URLDownloadToFile (BYVAL 0, URLPath1, LocalPath, 0, 0)  THEN
      GOSUB VerError                                              ' Error Bail out
   END IF                                                         '
   OPEN LocalPath FOR INPUT AS #1                                 ' Get version on the server
   LINE INPUT #1, VerData                                         '
   CLOSE #1                                                       '
   IF LEFT$(VerData, 1) <> "<" THEN                               ' Look OK?
      sGetVerFile = LEFT$(VerData, INSTR(VerData, " ") - 1)       ' Return 1st 'word' from line
      MExitFunc                                                   '
   END IF                                                         '

   GOSUB VerError                                                 '
   MExitFunc

VerError:
   sDoMsgBox "Can't access SPFLite web site for Update check", %MB_OK OR %MB_USERICON, "Version Check"
   sGetVerFile = ""                                               ' Return null
   MExitFunc                                                      '
   RETURN                                                         '
END FUNCTION

FUNCTION sGetWindowsTempDir() AS STRING
'---------- Get windows temp dir name
LOCAL lResult AS LONG
LOCAL buff AS ASCIIZ * %MAX_PATH
   lResult = GetTempPath(BYVAL SIZEOF(buff), Buff)                '
   FUNCTION = TRIM$(buff)                                         '
END FUNCTION

FUNCTION sGetWord(BYREF Sent AS STRING, DoStrip AS INTEGER, QuoteSig AS INTEGER) AS STRING
'---------- Get Next "Word" from passed string, return null if error
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL OrigStr AS STRING
LOCAL Ltr AS BYTE POINTER
   MEntry
   OrigStr = Sent                                                 ' Save original Input in case NOStrip
   Sent = TRIM$(Sent)                                             ' Clean up
   IF ISNULL(Sent) THEN                                           ' If nothing passed
      sGetWord = ""                                               ' Return NULL
      IF DoStrip THEN Sent = "" ELSE Sent = OrigStr               ' If DoStrip, make Sent also null, else leave alone
      MExitFunc                                                   ' And exit
   END IF                                                         '
   Sent = Sent + " "                                              ' Ensure trailing blank
   Ltr = STRPTR(Sent)                                             ' Create pointer to input string
   FOR i = 1 TO LEN(Sent)                                         ' OK, lets scan the string
      j = 0                                                       ' Indicate no quote skipping done
      SELECT CASE @Ltr                                            ' The characters we need to inlude in a word
         CASE 32                                                  ' We have the end of the 'word'
            EXIT FOR                                              ' We can leave loop now and process it, i is set
         CASE 34                                                  ' Quotes -> "
            j = INSTR(i + 1, Sent, $DQ)                           ' Find closing quote
            IF j THEN                                             ' Skip over literal
               i = j: Ltr = STRPTR(Sent) + j - 1                  ' Skip I over to closing quote
            ELSE                                                  ' Whoops, unmatched quotes
               '
            END IF                                                '
         CASE 39                                                  ' Single quote -> '
            j = INSTR(i + 1, Sent, CHR$(39))                      ' Find closing quote
            IF j THEN                                             ' Skip over literal
               i = j: Ltr = STRPTR(Sent) + j - 1                  ' Skip I over to closing quote
            ELSE                                                  ' Whoops, unmatched quotes
               '
            END IF                                                '
         CASE 96                                                  ' Back quote -> `
            j = INSTR(i + 1, Sent, CHR$(96))                      ' Find closing quote
            IF j THEN                                             ' Skip over literal
               i = j: Ltr = STRPTR(Sent) + j - 1                  ' Skip I over to closing quote
            ELSE                                                  ' Whoops, unmatched quotes
               '
            END IF                                                '
      END SELECT                                                  '
      INCR Ltr                                                    ' Onward
   NEXT I                                                         ' Loop till done
   sGetWord = LEFT$(Sent, i - 1)                                  ' Extract the word
   IF DoStrip THEN                                                '
      sent = MID$(Sent, i + 1, LEN(Sent) - i - 1)                 ' Remove it from source string if requested
   ELSE                                                           '
      Sent = OrigStr                                              ' Else restore original
   END IF                                                         '
   MExit
END FUNCTION

FUNCTION sHash(sTxt AS STRING, iHash AS DWORD) AS DWORD
'---------- Create DWORD Hash for a string (Now used only for KBD file)
REGISTER i AS LONG, lhash AS DWORD
LOCAL cptr AS BYTE POINTER
   MEntry
   lHash = iHash
   cptr = STRPTR(sTxt)
   FOR i = 1 TO LEN(sTxt)
      lhash = lhash + @cptr: ROTATE LEFT lhash, 1
      lhash = ABS(lhash)
'      lhash and= &H7FFFFFFF
      INCR cptr
   NEXT i
   FUNCTION = lhash
   MExit
END FUNCTION

FUNCTION sHelpIndex(topic AS STRING) AS LONG
'---------- Lookup help request and convert to mapid index for HH.EXE
LOCAL i AS LONG, key, t AS STRING
STATIC HelpBuilt AS LONG

   MEntry
   IF ISFALSE HelpBuilt THEN                                      ' Build table only once
      HelpBuilt = %True                                           ' Flip once switch
      sHelpInitA                                                  ' Do a chunk
   END IF
   key = TRIM$(UUCASE(topic))                                     ' Trim and UC it
   FOR i = 1 TO gHelpCount                                        ' Loop through Help Table
      IF key = gHelpKey(i) THEN                                   '
         sHelpIndex = gHelpMapid(i)                               '
         MExitFunc                                                ' Found it, exit
      END IF                                                      '
   NEXT i                                                         '
   sHelpIndex = 0                                                 ' Not found, return null
   MExit
END FUNCTION

SUB      sHelpInitA()
'---------- Add the HELP operand keywords
   sHelpInit("A",            %HELP_A)
   sHelpInit("AA",           %HELP_AA)
   sHelpInit("Abbrev",       %HELP_Abbreviations)
   sHelpInit("ADD",          %HELP_ADD)
   sHelpInit("APPEND",       %HELP_APPEND)
   sHelpInit("Appendix",     %HELP_Appendix)
   sHelpInit("AUTOBKUP",     %HELP_AUTOBKUP)
   sHelpInit("AUTONUM",      %HELP_AUTONUM)
   sHelpInit("AUTOFAV",      %HELP_UsingAUTOFAVtoaddtoFileLists)
   sHelpInit("AUTOCAPS",     %HELP_AUTOCAPS)
   sHelpInit("AUTOSAVE",     %HELP_AUTOSAVE)
   sHelpInit("B",            %HELP_B)
   sHelpInit("BasicEdit",    %HELP_BasicEdit)
   sHelpInit("BB",           %HELP_BB)
   sHelpInit("BNDS",         %HELP_BNDS)
   sHelpInit("BOTTOM",       %HELP_BOTTOM)
   sHelpInit("BOUNDS",       %HELP_BOUNDS)
   sHelpInit("BROWSE",       %HELP_BROWSE)
   sHelpInit("C",            %HELP_C)
   sHelpInit("CC",           %HELP_C)
   sHelpInit("CANCEL",       %HELP_CANCEL)
   sHelpInit("CAPS",         %HELP_CAPS)
   sHelpInit("CASE",         %HELP_CASE)
   sHelpInit("CHANGE",       %HELP_CHANGE)
   sHelpInit("CLIP",         %HELP_CLIP)
   sHelpInit("ClipBoard",    %HELP_Clipboard)
   sHelpInit("CLONE",        %HELP_CLONE)
   sHelpInit("CMD",          %HELP_CMD)
   sHelpInit("COLLATE",      %HELP_COLLATE)
   sHelpInit("Colorize",     %HELP_Colorize)
   sHelpInit("ColorSel",     %HELP_ColorSelectionCriteria)
   sHelpInit("COLSL",        %HELP_COLSL)
   sHelpInit("COLSP",        %HELP_COLSP)
   sHelpInit("COMPRESS",     %HELP_COMPRESS)
   sHelpInit("COPY",         %HELP_COPY)
   sHelpInit("CREATE",       %HELP_CREATE)
   sHelpInit("CREATEUse",    %HELP_CreateReplace)
   sHelpInit("CRETRIEV",     %HELP_CRETRIEV)
   sHelpInit("Customize",    %HELP_Customize)
   sHelpInit("CUT",          %HELP_CUT)
   sHelpInit("D",            %HELP_D)
   sHelpInit("DD",           %HELP_D)
   sHelpInit("DEFAULTKeys",  %HELP_DefaultKeys)
   sHelpInit("DELETE",       %HELP_DELETE)
   sHelpInit("Differences",  %HELP_Differences)
   sHelpInit("DIR",          %HELP_DIR)
   sHelpInit("DROP",         %HELP_DROP)
   sHelpInit("EDIT",         %HELP_EDIT)
   sHelpInit("EDITBnds",     %HELP_EditBoundaries)
   sHelpInit("END",          %HELP_END)
   sHelpInit("Enum",         %HELP_Enumerating)
   sHelpInit("ENUMWITH",     %HELP_ENUMWITH)
   sHelpInit("EOL",          %HELP_EOL)
   sHelpInit("EXCLUDE",      %HELP_EXCLUDEEdit)
   sHelpInit("EXCLUDEUse",   %HELP_Excludedlines)
   sHelpInit("EXIT",         %HELP_EXIT)
   sHelpInit("ExtFile",      %HELP_ExternalFileChanges)
   sHelpInit("F",            %HELP_F)
   sHelpInit("FAVORITE",     %HELP_FAVORITE)
   sHelpInit("Features",     %HELP_Features)
   sHelpInit("FF",           %HELP_FF)
   sHelpInit("Filelists",    %HELP_WorkingWithFileLists)
   sHelpInit("FileMngr",     %HELP_FileManager)
   sHelpInit("FileProfs",    %HELP_Workingwithfileprofiles)
   sHelpInit("FIND",         %HELP_FIND)
   sHelpInit("FindChange",   %HELP_FindingandChangingData)
   sHelpInit("FindFM",       %HELP_FINDFM)
   sHelpInit("FLIP",         %HELP_FLIP)
   sHelpInit("FOLD",         %HELP_FOLD)
   sHelpInit("Fonts",        %HELP_FONTS)
   sHelpInit("G",            %HELP_G)
   sHelpInit("GG",           %HELP_G)
   sHelpInit("GLUEWITH",     %HELP_GLUEWITH)
   sHelpInit("H",            %HELP_H)
   sHelpInit("HH",           %HELP_H)
   sHelpInit("HELP",         %HELP_HELP)
   sHelpInit("HEX",          %HELP_HEX)
   sHelpInit("HIDE",         %HELP_HIDE)
   sHelpInit("HILITE",       %HELP_HILITE)
   sHelpInit("I",            %HELP_I)
   sHelpInit("Include",      %HELP_Include)
   sHelpInit("[",            %HELP_IndentShiftLeft)
   sHelpInit("[[",           %HELP_IndentShiftLeft)
   sHelpInit("]",            %HELP_IndentShiftRight)
   sHelpInit("]]",           %HELP_IndentShiftRight)
   sHelpInit("PrimIndex",    %HELP_IndexToKeyboardPrimitives)
   sHelpInit("Install",      %HELP_Installing)
   sHelpInit("Intro",        %HELP_Welcometospflite)
   sHelpInit("J",            %HELP_J)
   sHelpInit("JJ",           %HELP_J)
   sHelpInit("JOIN",         %HELP_JOIN)
   sHelpInit("KEEP",         %HELP_KEEP)
   sHelpInit("KBCustomize",  %HELP_KeyboardCustomize)
   sHelpInit("KBMacros",     %HELP_KeyBoardMacros)
   sHelpInit("Primitives",   %HELP_KeyboardPrimitives)
   sHelpInit("AppPrims",     %HELP_KeyboardPrimitives1)
   sHelpInit("KEYMAP",       %HELP_KEYMAP)
   sHelpInit("KMDialog",     %HELP_KeyMapDialog)
   sHelpInit("KMOverview",   %HELP_KeyMappingOverview)
   sHelpInit("L",            %HELP_L)
   sHelpInit("Labels",       %HELP_Labels)
   sHelpInit("LC",           %HELP_LC)
   sHelpInit("LCC",          %HELP_LC)
   sHelpInit("LCP",          %HELP_LCP)
   sHelpInit("LCSyntax",     %HELP_LCSyntax)
   sHelpInit("Line",         %HELP_LINE)
   sHelpInit("LCExtensions", %HELP_LineCommandExtensions)
   sHelpInit("LCommands",    %HELP_LineCommands)
   sHelpInit("LineLengths",  %HELP_LineLengths)
   sHelpInit("LOCATE",       %HELP_LOCATE)
   sHelpInit("LRECL",        %HELP_LRECL)
   sHelpInit("M",            %HELP_M)
   sHelpInit("MM",           %HELP_M)
   sHelpInit("Macros",       %HELP_Macros)
   sHelpInit("MAKELIST",     %HELP_MAKELIST)
   sHelpInit("MAPPING",      %HELP_WorkingMappingStrings)
   sHelpInit("MARK",         %HELP_MARK)
   sHelpInit("MARKP",        %HELP_MARKP)
   sHelpInit("MASK",         %HELP_MASK)
   sHelpInit("MD",           %HELP_MD)
   sHelpInit("MDD",          %HELP_MD)
   sHelpInit("MEDIT",        %HELP_MEDIT)
   sHelpInit("MINLEN",       %HELP_MINLEN)
   sHelpInit("MN",           %HELP_MN)
   sHelpInit("MNN",          %HELP_MN)
   sHelpInit("Mouse",        %HELP_Mouse)
   sHelpInit("MultiEdit",    %HELP_MultiEdit)
   sHelpInit("N",            %HELP_N)
   sHelpInit("NDELETE",      %HELP_NDELETE)
   sHelpInit("NEXCLUDE",     %HELP_NEXCLUDE)
   sHelpInit("NFIND",        %HELP_NFIND)
   sHelpInit("NFLIP",        %HELP_NFLIP)
   sHelpInit("NonWindow",    %HELP_NonWindowsFiles)
   sHelpInit("NOTE",         %HELP_NOTE)
   sHelpInit("Notes",        %HELP_NOTEs)
   sHelpInit("NOTIFY",       %HELP_NOTIFY)
   sHelpInit("NREVERT",      %HELP_NREVERT)
   sHelpInit("NSHOW",        %HELP_NSHOW)
   sHelpInit("NULINE",       %HELP_NULINE)
   sHelpInit("NUM",          %HELP_NUMBER)
   sHelpInit("NUMB",         %HELP_NUMBER)
   sHelpInit("NUMBER",       %HELP_NUMBER)
   sHelpInit("NUMTYPE",      %HELP_NUMTYPE)
   sHelpInit("O",            %HELP_O)
   sHelpInit("OO",           %HELP_O)
   sHelpInit("OPEN",         %HELP_OPEN)
   sHelpInit("OPTIONS",      %HELP_OPTIONS)
   sHelpInit("OptFM",        %HELP_OptionsFManager)
   sHelpInit("OptGeneral",   %HELP_OptionsGeneral)
   sHelpInit("OptKB",        %HELP_OptionsKeyboard)
   sHelpInit("OptMouse",     %HELP_OptionsMouse)
   sHelpInit("OptScreen",    %HELP_OptionsScreen)
   sHelpInit("OprSubmit",    %HELP_OptionsSubmit)
   sHelpInit("OR",           %HELP_OR)
   sHelpInit("ORR",          %HELP_OR)
   sHelpInit("ORDER",        %HELP_ORDER)
   sHelpInit("PAGE",         %HELP_PAGE)
   sHelpInit("PASTE",        %HELP_PASTE)
   sHelpInit("PCClrChg",     %HELP_ColorChangeRequest)
   sHelpInit("PCLinRange",   %HELP_PCLineRange)
   sHelpInit("PCSyntax",     %HELP_PCSyntax)
   sHelpInit("PL",           %HELP_PL)
   sHelpInit("PLL",          %HELP_PL)
   sHelpInit("Portable",     %HELP_Portable)
   sHelpInit("PowerType",    %HELP_WorkingwithPowerTypingMode)
   sHelpInit("PREPEND",      %HELP_PREPEND)
   sHelpInit("PRESERVE",     %HELP_PRESERVE)
   sHelpInit("PCommands",    %HELP_PrimaryCommands)
   sHelpInit("PRINT",        %HELP_PRINT)
   sHelpInit("PrtScreen",    %HELP_PrintScreen)
   sHelpInit("PROFILE",      %HELP_PROFILE)
   sHelpInit("PTYPE",        %HELP_PTYPE)
   sHelpInit("QUERY",        %HELP_QUERY)
   sHelpInit("R",            %HELP_R)
   sHelpInit("RR",           %HELP_R)
   sHelpInit("RCHANGE",      %HELP_RCHANGE)
   sHelpInit("ReadOnly",     %HELP_ReadOnlyFiles)
   sHelpInit("RECALL",       %HELP_RECALL)
   sHelpInit("RECFM",        %HELP_RECFM)
   sHelpInit("REDO",         %HELP_REDO)
   sHelpInit("RegEx",        %HELP_SpecifyingaRegularExpression)
   sHelpInit("RELOAD",       %HELP_RELOAD)
   sHelpInit("RENAME",       %HELP_RENAME)
   sHelpInit("RENUM",        %HELP_RENUM)
   sHelpInit("REPLACE",      %HELP_REPLACE)
   sHelpInit("RESET",        %HELP_RESETEdit)
   sHelpInit("RETF",         %HELP_RETF)
   sHelpInit("RETRIEVE",     %HELP_RETRIEVE)
   sHelpInit("REVERT",       %HELP_REVERT)
   sHelpInit("RFIND",        %HELP_RFIND)
   sHelpInit("RLOC",         %HELP_RLOC)
   sHelpInit("RLOCFIND",     %HELP_RLOCFIND)
   sHelpInit("RUN",          %HELP_RUN)
   sHelpInit("S",            %HELP_S)
   sHelpInit("SS",           %HELP_S)
   sHelpInit("SAVE",         %HELP_SAVE)
   sHelpInit("SAVEALL",      %HELP_SAVEALL)
   sHelpInit("SAVEAS",       %HELP_SAVEAS)
   sHelpInit("SC",           %HELP_SC)
   sHelpInit("SCC",          %HELP_SC)
   sHelpInit("SCP",          %HELP_SCP)
   sHelpInit("SCROLL",       %HELP_SCROLL)
   sHelpInit("Scrolling",    %HELP_Scrolling)
   sHelpInit("SET",          %HELP_SET)
   sHelpInit("SETUNDO",      %HELP_SETUNDO)
   sHelpInit("<",            %HELP_ShiftDataLeft)
   sHelpInit("<<",           %HELP_ShiftDataLeft)
   sHelpInit(">",            %HELP_ShiftDataRight)
   sHelpInit(">>",           %HELP_ShiftDataRight)
   sHelpInit("Shifting",     %HELP_Shifting)
   sHelpInit("(",            %HELP_ShiftLeft)
   sHelpInit("((",           %HELP_ShiftLeft)
   sHelpInit(")",            %HELP_ShiftRight)
   sHelpInit("))",           %HELP_ShiftRight)
   sHelpInit("SHOW",         %HELP_SHOW)
   sHelpInit("SI",           %HELP_SI)
   sHelpInit("SORT",         %HELP_SORT)
   sHelpInit("SOURCE",       %HELP_SOURCE)
   sHelpInit("Pictures",     %HELP_SpecifyingaPictureorFormatString)
   sHelpInit("SPLIT",        %HELP_SPLIT)
   sHelpInit("SplitJoin",    %HELP_WorkingwithSPLITandJOINCommands)
   sHelpInit("START",        %HELP_START)
   sHelpInit("Starting",     %HELP_StartingandEndingSPFLite)
   sHelpInit("STATE",        %HELP_STATE)
   sHelpInit("STATEUse",     %HELP_STATEConsiderations)
   sHelpInit("StatusBar",    %HELP_StatusBar)
   sHelpInit("SUBARG",       %HELP_SUBARG)
   sHelpInit("SUBCMD",       %HELP_SUBCMD)
   sHelpInit("SUBMIT",       %HELP_SUBMIT)
   sHelpInit("SUBMITUse",    %HELP_SubmitUsage)
   sHelpInit("Substitute",   %HELP_Substitution)
   sHelpInit("Support",      %HELP_Support)
   sHelpInit("SWAP",         %HELP_SWAP)
   sHelpInit("T",            %HELP_T)
   sHelpInit("TT",           %HELP_T)
   sHelpInit("TabPages",     %HELP_TabPages)
   sHelpInit("TabColumns",   %HELP_TabsColumns)
   sHelpInit("TABSL",        %HELP_TABSL)
   sHelpInit("TABSP",        %HELP_TABSP)
   sHelpInit("TAG",          %HELP_TAG)
   sHelpInit("Tags",         %HELP_Tags)
   sHelpInit("TB",           %HELP_TB)
   sHelpInit("TBB",          %HELP_TB)
   sHelpInit("TC",           %HELP_TC)
   sHelpInit("TCC",          %HELP_TC)
   sHelpInit("TCP",          %HELP_TCP)
   sHelpInit("TF",           %HELP_TF)
   sHelpInit("TFF",          %HELP_TF)
   sHelpInit("TG",           %HELP_TG)
   sHelpInit("TGG",          %HELP_TG)
   sHelpInit("TJ",           %HELP_TJ)
   sHelpInit("TJJ",          %HELP_TJ)
   sHelpInit("TL",           %HELP_TL)
   sHelpInit("TLL",          %HELP_TL)
   sHelpInit("TM",           %HELP_TM)
   sHelpInit("TMM",          %HELP_TM)
   sHelpInit("TOP",          %HELP_TOP)
   sHelpInit("TR",           %HELP_TR)
   sHelpInit("TRR",          %HELP_TR)
   sHelpInit("TS",           %HELP_TS)
   sHelpInit("TU",           %HELP_TU)
   sHelpInit("TUU",          %HELP_TU)
   sHelpInit("TX",           %HELP_TX)
   sHelpInit("TXX",          %HELP_TX)
   sHelpInit("U",            %HELP_U)
   sHelpInit("UU",           %HELP_U)
   sHelpInit("UC",           %HELP_UC)
   sHelpInit("UCC",          %HELP_UC)
   sHelpInit("UCP",          %HELP_UCP)
   sHelpInit("ULINE",        %HELP_ULINE)
   sHelpInit("UNDO",         %HELP_UNDO)
   sHelpInit("UNNUMBER",     %HELP_UNNUMBER)
   sHelpInit("UNNUM",        %HELP_UNNUMBER)
   sHelpInit("V",            %HELP_V)
   sHelpInit("VV",           %HELP_V)
   sHelpInit("VIEW",         %HELP_VIEW)
   sHelpInit("VHiliting",    %HELP_VirtualHighlighting)
   sHelpInit("VSAVE",        %HELP_VSAVE)
   sHelpInit("W",            %HELP_W)
   sHelpInit("WW",           %HELP_W)
   sHelpInit("WDIR",         %HELP_WDIR)
   sHelpInit("WORD",         %HELP_WORD)
   sHelpInit("WProcess",     %HELP_WordProcessing)
   sHelpInit("Working",      %HELP_WorkingWith)
   sHelpInit("LINEUse",      %HELP_WorkingwithLINE)
   sHelpInit("ULINEUse",     %HELP_WorkingwithUserlines)
   sHelpInit("WORDUse",      %HELP_WorkingwithWordandDelimiterChara)
   sHelpInit("X",            %HELP_X)
   sHelpInit("XX",           %HELP_X)
   sHelpInit("XSUBMIT",      %HELP_XSUBMIT)
   sHelpInit("XTABS",        %HELP_XTABS)
END SUB

SUB      sHelpInit(kw AS STRING, eq AS LONG)
'---------- Add a single Help Index to the table
   INCR gHelpCount                                                ' Bump global count
   gHelpKey(gHelpCount) = kw                                      ' Store the Key
   gHelpMapid(gHelpCount) = eq                                    ' and the MapID
END SUB

FUNCTION sHex2Str(HexStr AS STRING) AS STRING
'---------- Convert a Hex string to its proper self
LOCAL Temp AS STRING, I AS LONG
   MEntry
   Temp = SPACE$(LEN(HexStr) \ 2)
   FOR i = 1 TO LEN(HexStr) \ 2
      MID$(Temp, i, 1) = CHR$(VAL("&H" & MID$(HexStr, i * 2 - 1, 2)))
   NEXT I
   FUNCTION = Temp
   MExit
END FUNCTION

FUNCTION sHexLower(Lower AS STRING, Orig AS STRING) AS STRING
'---------- Modify the lower half of a hex char
LOCAL OrigHex, NewHex AS STRING                                   '
   MEntry
   IF TP.PrfSrceXlate THEN                                        ' If not ANSI, must double translate
      OrigHex = Orig                                              ' Get Orig in equivalent SOURCE hex
      TP.Translate(OrigHex, TP.TPPrfGetSA2SPtr)                   '
      OrigHex = HEX$(ASC(OrigHex), 2)                             '
      NewHex = CHR$(VAL("&H" + LEFT$(OrigHex, 1) + Lower))        ' Alter the Hex now
      TP.Translate(NewHex, TP.TPPrfGetSS2APtr)                    '
      sHexLower = NewHex                                          ' Pass back character in ANSI for screen display
   ELSE                                                           ' Else normal ANSI
      OrigHex = HEX$(ASC(Orig), 2)                                ' Make it hex
      sHexLower = CHR$(VAL("&H" + LEFT$(OrigHex, 1) + Lower))     ' Alter and pass back the result
   END IF                                                         '
   MExit
END FUNCTION

FUNCTION sHexUpper(Upper AS STRING, Orig AS STRING) AS STRING
'---------- Modify the upper half of a hex char
LOCAL OrigHex, NewHex AS STRING                                   '
   MEntry
   IF TP.PrfSrceXlate THEN                                        ' If not ANSI, must double translate
      OrigHex = Orig                                              ' Get Orig in equivalent SOURCE hex
      TP.Translate(OrigHex, TP.TPPrfGetSA2SPtr)                   '
      OrigHex = HEX$(ASC(OrigHex), 2)                             '
      NewHex = CHR$(VAL("&H" + Upper + RIGHT$(OrigHex, 1)))       ' Alter the Hex now
      TP.Translate(NewHex, TP.TPPrfGetSS2APtr)                    ' Pass back character in ANSI for screen display
      sHexUpper = NewHex                                          '
   ELSE                                                           ' Else normal ANSI
      OrigHex = HEX$(ASC(Orig), 2)                                ' Make it hex
      sHexUpper = CHR$(VAL("&H" + Upper + RIGHT$(OrigHex, 1)))    '
   END IF                                                         '
  MExit
END FUNCTION                                                      '

FUNCTION sINIGetString(BYVAL sSection AS STRING, BYVAL sKey AS STRING, BYVAL sDefault AS STRING) AS STRING
'---------- Get string from ini file
LOCAL RetVal AS LONG, zResult AS ASCIIZ * 32768
LOCAL zSection AS ASCIIZ * %MAX_PATH, zKey AS ASCIIZ * %MAX_PATH
LOCAL zDefault AS ASCIIZ * %MAX_PATH
LOCAL t AS STRING
   t = ENV.INIFileName
   zSection = sSection : zKey  = sKey                             '
   zDefault = sDefault                                            '
   RetVal = GetPrivateProfileString(zSection, zKey, zDefault, zResult, SIZEOF(zResult), ENV.INIFileName)
   IF RetVal THEN FUNCTION = LEFT$(zResult, RetVal)               '
END FUNCTION

FUNCTION sINISetString(BYVAL sSection AS STRING, BYVAL sKey AS STRING, _
                      BYVAL sStr AS STRING) AS LONG
'---------- Set string to ini file                                '
LOCAL zSection AS ASCIIZ * %MAX_PATH, zKey AS ASCIIZ * %MAX_PATH  '
LOCAL zStr AS ASCIIZ * 32768, lStr AS STRING                      '
   lStr = sStr                                                    '
   IF LEFT$(lStr, 1) = $DQ AND RIGHT$(lStr, 1) = $DQ THEN         ' Is this a quoted string?
      lStr = $DQ + lStr + $DQ                                     ' Yes, double the quotes
   END IF                                                         '
   zSection = sSection : zKey  = sKey : zStr = lStr               '
   FUNCTION = WritePrivateProfileString(zSection, zKey, zStr, ENV.INIFileName)
END FUNCTION

SUB sInit_CodePage (CP AS CodePage_CP_T)
'/-----------------------------------------------------------------------------/
'/  Init_CodePage ()                                                           /
'/                                                                             /
'/  Perform CONSTRUCTOR-Like initialization of a CodePage_CP_T STRUCTURE        /
'/-----------------------------------------------------------------------------/
LOCAL I, T AS LONG
    CP.CP_LineNo            = 0
    CP.CP_Errors            = 0
    CP.CP_Reason            = ""
    CP.TT.TT_Errors         = 0
    CP.TT.TT_Reason         = ""
    CP.TT.TT_Author         = ""               '/ Creator of table
    CP.TT.TT_GenDate        = ""               '/ "2002-12-03 00:00:00"
    CP.TT.TT_Mode           = ""               '/ RT/ES round trip/subset
    CP.TT.TT_Name           = ""               '/ Name of translation table
    CP.TT.TT_Title          = ""               '/ Title of translation table
    CP.TT.TT_Other          = ""               '/ Creator of table

    FOR I = 1 TO %TX_Max                       '/ From ASCII TO EBCDIC
        CP.TX(I).TX_Defined = 0
        CP.TX(I).TX_Errors  = 0
        CP.TX(I).TX_Values  = 0                '/ Number of values stored
        CP.TX(I).TX_Reason  = ""               '/ Reason for reported error
        FOR T = 0 TO 15
            CP.TX(I).TX_Entry (T) = 0          '/ Flage for 0_ TO F_ lines
        NEXT
        CP.TX(I).TX_CCSID   = ""               '/ Pieces of 'NUMBER' as Int
        CP.TX(I).TX_CGCSGID = ""               '/ CGCSGID   "00695"
        CP.TX(I).TX_CodeSet = ""               '/ Full CodeSet name
        CP.TX(I).TX_CPGID   = ""               '/ CPGID  "01140"
        CP.TX(I).TX_Euro    = ""               '/ Value of Euro OR 00
        CP.TX(I).TX_Number  = ""               '/ "1140", "8859_1", ETC.
        CP.TX(I).TX_Origin  = ""               '/ "IBM", "ISO" etc.
        CP.TX(I).TX_Related = ""               '/ Related Euro CCSID "-37"
        CP.TX(I).TX_Scheme  = ""               '/ Encoding scheme "1100"
        CP.TX(I).TX_Size    = ""               '/ Num of defined chars
        CP.TX(I).TX_Sub     = ""               '/ Substitution char
        CP.TX(I).TX_Type    = ""               '/ "ASCII", "EBCDIC"
        CP.TX(I).TX_UCM     = ""               '/ .UCM file name
        CP.TX(I).TX_UCMDate = ""               '/ "2002-12-03 00:00:00"
        CP.TX(I).TX_Version = ""               '/ "2.3.3", "1995", ETC.
        CP.TX(I).TX_Other   = ""               '/ Unknown keyword value
        FOR T = 0 TO 255
            CP.TX(I).TX_Table (T) = 0          '/ Final translation table
        NEXT
    NEXT
END SUB ' sInit_CodePage

FUNCTION sInvert(t AS STRING) AS STRING
'---------- Invert a string for sorting
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL nt AS STRING
   MEntry
   nt = ""                                                        '
   FOR i = 1 TO LEN(t)                                            ' Loop through string
      j = ASC(MID$(t, i, 1))                                      ' Get the ASCII value
      j = j XOR 255                                               ' Invert it
      nt = nt + CHR$(j)                                           ' Stuff it back
   NEXT i                                                         '
   sInvert = nt                                                   ' Pass back the answer
   MExit
END FUNCTION

FUNCTION sIsNullFile(fn AS STRING) AS LONG
'---------- Return true if a file exists AND it is zero length
LOCAL FD AS DIRDATA, fn2 AS STRING
   fn2 = DIR$(fn, TO FD)                                          ' Get the FD data if it exists
   IF ISNULL(fn2) THEN EXIT FUNCTION                              ' No file?  Then return false.
   IF FD.FileSizeHigh > 0 OR FD.FileSizeLow > 0 THEN EXIT FUNCTION' Something in the size, exit false
   FUNCTION = %True                                               ' Else we have a true nullfile
END FUNCTION                                                      '

FUNCTION sKbdSortEx(p1 AS STRING * 32, BYREF p2 AS STRING * 32) AS LONG
'---------- Custom sort key for the PFShow Help data
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL Key1, Key2, t, tt, p AS STRING
   t = sSub128(TRIM$(p1))                                         ' Back to normal ASCII
   p = LEFT$(t, 1): t = MID$(t, 2)                                ' Extract chord sequence
   i = INSTR(t, "-")                                              ' Look for prefix
   j = INSTR(t, "=")                                              ' Look for the =
   tt = IIF$(i <> 0, MID$(t, i + 1 TO j - 1), LEFT$(t, j - 1))    ' Extract the Key name
   IF LEN(tt) = 2 AND LEFT$(tt, 1) = "F" THEN                     ' Fn name?
      tt = "F0" + MID$(tt, 2)                                     ' Normalize it to Fnn
   END IF                                                         '
   IF i = 0 THEN                                                  ' If no - then
      Key1 = tt                                                   ' Key is everything up to the =
   ELSE                                                           '
      Key1 = tt + p                                               ' Base followed by chord priority
   END IF

   t = sSub128(TRIM$(p2))                                         ' Back to normal ASCII
   p = LEFT$(t, 1): t = MID$(t, 2)                                ' Extract chord sequence
   i = INSTR(t, "-")                                              ' Look for prefix
   j = INSTR(t, "=")                                              ' Look for the =
   tt = IIF$(i <> 0, MID$(t, i + 1 TO j - 1), LEFT$(t, j - 1))    ' Extract the Key name
   IF LEN(tt) = 2 AND LEFT$(tt, 1) = "F" THEN                     ' Fn name?
      tt = "F0" + MID$(tt, 2)                                     ' Normalize it to Fnn
   END IF                                                         '
   IF i = 0 THEN                                                  ' If no - then
      Key2 = tt                                                   ' Key is everything up to the =
   ELSE                                                           '
      Key2 = tt + p                                               ' Base followed by chord priority
   END IF
   FUNCTION = StrCmpr(Key1, Key2)
END FUNCTION

FUNCTION sKMacro(macline AS STRING) AS LONG
'---------- Handle all the difficult ~K(...) substitutions
LOCAL i, j, k AS LONG, Kline, Kkey1, KKey2, KData, KeyMode AS STRING
LOCAL KeyName AS STRING * 12
   MEntry
   KLine = macline
   i = INSTR(UUCASE(Kline), "~K("): IF i = 0 THEN i = INSTR(UUCASE(KLine), "^K(") ' Find next ~K(
   DO WHILE i > 0                                                 ' Loop-de-loop
      k = INSTR(i, KLine, ")")                                    ' Look for closing bracket
      IF k = 0 OR k = i + 3 THEN FUNCTION = %True: MExitFunc      ' No?? or null (), Error return
      Kkey1 = MID$(Kline, i, k - i + 1)                           ' Extract K key1 ~K(xxx)
      Kkey2 = MID$(Kline, i + 3, k - i - 3)                       ' Extract K key2 xxx
      SELECT CASE AS CONST$ UUCASE(Kkey2)                         ' OK then, what've we got
         CASE "DATE"                                              ' DATE
            REPLACE KKey1 WITH sDate() IN KLine                   '
         CASE "ISODATE"                                           ' ISODate
            KData = DATE$                                         ' Go get the date
            KData = MID$(KData, 7) + "-" + LEFT$(KData, 2) + "-" + MID$(KData, 4, 2)   ' Reformat it to ISO standard
            REPLACE KKey1 WITH KData IN KLine                     '
         CASE "ISOTIME"                                           ' ISOTime
            REPLACE KKey1 WITH TIME$ IN KLine                     '
         CASE ELSE                                                ' We're left with a Keyname style
            KeyMode = ""                                          ' Default mode
            KKey2 = UUCASE(KKey2)                                 '
            KeyName = KKey2                                       '
            j = INSTR(KKey2, "-")                                 ' Look for a dash
            IF j THEN                                             ' A dash, complicate it
               IF INSTR(LEFT$(KKey2, j - 1), "S") THEN KeyMode += "S"
               IF INSTR(LEFT$(KKey2, j - 1), "C") THEN KeyMode += "C"
               IF INSTR(LEFT$(KKey2, j - 1), "A") THEN KeyMode += "A"
               KeyName = MID$(KKey2, j + 1)                       ' Separate keyname
            END IF                                                '
            FOR j = 1 TO 104                                      ' Search the key master table
               IF KbdT.Labl(j) = KeyName THEN                     ' Find the key?
                  IF ISNULL(Keymode) THEN KData = KbdT.NData(j)   ' Select correct text string
                  IF Keymode = "S"   THEN KData = KbdT.SData(j)   '
                  IF Keymode = "C"   THEN KData = KbdT.CData(j)   '
                  IF Keymode = "A"   THEN KData = KbdT.AData(j)   '
                  IF Keymode = "SC"  THEN KData = KbdT.SCData(j)  '
                  IF Keymode = "SA"  THEN KData = KbdT.SAData(j)  '
                  IF Keymode = "SCA" THEN KData = KbdT.SCAData(j)
                  IF Keymode = "CA"  THEN KData = KbdT.CAData(j)  '
                  IF LEFT$(KData, 1) = "[" THEN                   ' Emitted text?
                     DO WHILE INSTR(KData, "[[")                  ' Clean up pairs
                        REPLACE "[[" WITH "[" IN KData            '
                     LOOP                                         '
                     DO WHILE INSTR(KData, "]]")                  ' Clean up pairs
                        REPLACE "]]" WITH "]" IN KData            '
                     LOOP                                         '
                     KData = MID$(KData, 2): KData = LEFT$(KData, LEN(KData) - 1)
                     REPLACE KKey1 WITH KData IN KLine            '
                     EXIT SELECT                                  '
                  ELSEIF LEFT$(Kdata, 1) = "(" OR LEFT$(Kdata, 1) = "{" THEN
                     FUNCTION = %True: MExitFunc                  '
                  ELSE                                            ' Non-bracketed string
                     REPLACE KKey1 WITH KData IN KLine            '
                     EXIT SELECT                                  '
                  END IF                                          '
               END IF                                             '
            NEXT j                                                '
            FUNCTION = %True: MExitFunc                           '
      END SELECT                                                  '
      i = INSTR(UUCASE(Kline), "~K("): IF i = 0 THEN i = INSTR(UUCASE(KLine), "^K(") ' Find next ~K(
   LOOP                                                           '
   macline = KLine
   FUNCTION = %False                                              ' All is well
   MExit
END FUNCTION

SUB sLCmdDataShiftLeft(BYREF orig_line         AS STRING, _
                       BYREF orig_attr         AS WSTRING, _
                       BYVAL shift_amount      AS LONG, _
                       BYVAL left_bound        AS LONG, _
                       BYVAL right_bound       AS LONG, _
                       BYREF shift_error       AS LONG, _
                       BYREF shift_change      AS LONG)
'/-----------------------------------------------------------------------------
'/ lCmdDataShiftLeft
'/
'/ Accepts data line, bounds, and left-shift amount.
'/
'/ Returns shifted data as result, plus flag if "Data shifting incomplete"
'/ has occurred.  A revised color line is also returned.
'/-----------------------------------------------------------------------------
DIM seg (0 TO 1)        AS DataShift_Segment_t           '/ REDIM'd by Setup
LOCAL space_width, delta, max_delta, s, seg_count  AS LONG
   MEntry
   shift_error = 0
   shift_change = 0

   '/--------------------------------------------------------------------------
   '/ perform setup common to all < > << >> shifts
   '/ if Setup fails, pass original line back as is without reporting an error
   '/--------------------------------------------------------------------------
   seg_count = sLCmdDataShiftSetup                                            _
   (  orig_line                                                               _
   ,  shift_amount                                                            _
   ,  left_bound                                                              _
   ,  right_bound                                                             _
   ,  space_width                                                             _
   ,  seg ()                                                                  _
      )                                                                     '''

   IF seg_count < 1 THEN
      '/ line is zero-length, entirely blank, or data configuration will not
      '/ permit a shift to take place

      MExitSub
   END IF

   IF seg(1).gap_len <= space_width THEN
      '/ first segment is already as far left as it can go, can't shift it
      shift_error = 1                               '/ Data shifting incomplete
      MExitSub
   END IF

   '/ max_delta is the maximum possible left-shift for segment 1
   max_delta = seg(1).gap_len - space_width

   '/--------------------------------------------------------------------------
   '/ if shift_amount does not exceed the max_delta, then shift will be
   '/ successful, otherwise only shift by the max_delta and set the error flag
   '/ because we were not able to shift as much as was requested
   '/--------------------------------------------------------------------------
   IF shift_amount <= max_delta THEN
      delta = shift_amount                    '/ shift for the amount requested
   ELSE
      shift_error = 1                       '/ can't shift as much as requested
      delta = max_delta                       '/ only shift as much as possible
   END IF
   seg(1).gap_len -= delta                       '/ slide segment 1 to the left

   '/--------------------------------------------------------------------------
   '/ locate a segment, starting with segment 2, where the number of leading
   '/ spaces is greater than space_width that also contains data.  when found,
   '/ that segment gets additional leading spaces to match the amount shifted
   '/ left by segment 1.
   '/--------------------------------------------------------------------------
   IF delta > 0 THEN
      FOR s = 2 TO seg_count
         IF seg(s).gap_len > space_width AND seg(s).data_len > 0 THEN
            seg(s).gap_len += delta
            delta = 0
            EXIT FOR
         END IF
      NEXT
   END IF

   '/--------------------------------------------------------------------------
   '/ if the process above couldn't assign the blanks anywhere, put the blanks
   '/ on the tail segment regardless of any leading blanks IT may have, as long
   '/ as there is any data on it.  otherwise, just discard the extra blanks,
   '/ and the line will end up being shorter.
   '/--------------------------------------------------------------------------
   IF delta > 0 THEN
      IF seg(seg_count+1).data_len > 0 THEN
         seg(seg_count+1).gap_len += delta
      END IF
   END IF

   '/ after seg table has been adjusted, reconstruct and return a new data line
   sLCmdDataShiftResult (  _
      orig_line,           _
      orig_attr,           _
      seg (),              _
      seg_count,           _
      shift_error,         _
      shift_change)
   MExit
END SUB ' slCmdDataShiftLeft

SUB sLCmdDataShiftResult(BYREF orig_line         AS STRING, _
                         BYREF orig_attr         AS WSTRING, _
                         BYREF seg ()            AS DataShift_Segment_t, _
                         BYVAL seg_count         AS LONG, _
                         BYREF shift_error       AS LONG, _
                         BYREF shift_change      AS LONG)
'/-----------------------------------------------------------------------------
'/ lCmdDataShiftResult
'/
'/ using guidance from the segment table, transform orig_line string into a
'/ new string, in which the spacing between segments has been modified.
'/ build a string with modified spacing and copy back to orig_line.
'/-----------------------------------------------------------------------------
LOCAL c, work_line                                   AS STRING
LOCAL work_attr, Attr_Data, color_attr               AS WSTRING
LOCAL work_size, work_pos, work_len, change, i       AS LONG
   MEntry
   '/ determine new string size.  "label" if any is located in segment 0, and
   '/ "tail" segment is in segment (seg_count + 1).

   work_size = 0
   FOR i = 0 TO seg_count + 1
      work_size += seg(i).gap_len + seg(i).data_len
   NEXT

   work_line  = SPACE$ (work_size)
   work_attr = REPEAT$(work_size, CHR$$(0))

   '/ copy data positions into work_line.  since work_line already has blanks,
   '/ we don't need to copy the spaces - just skip over them.

   '/ if a color line was defined, create a new color line.  in case the
   '/ original color line was short, the new one will have the same size as
   '/ the new data line.
   work_pos = 1
   FOR i = 0 TO seg_count + 1

      '/ if a color line is defined, we have to copy the original color
      '/ attributes, if any, for blank characters as well as nonblanks.
      '/ if the size of a blank span has changed, the color attributes are
      '/ truncated or propagated on the right

      '/ the target location for the attribues of blanks is at 'work_pos'
      IF  LEN(orig_Attr)  > 0  _                          '/ colors are defined
      AND seg(i).gap_len   > 0 _
      AND seg(i).blank_pos > 0 _
      AND seg(i).blank_len > 0 THEN            '/ need to copy blank attributes
         attr_data = MID$ (orig_attr, seg(i).blank_pos, seg(i).blank_len)

         color_attr = RIGHT$ (attr_data, 1)
         MID$ (work_attr, work_pos, seg(i).gap_len) = _
            LSET$ (attr_data, seg(i).gap_len USING color_attr)

      END IF

      work_pos += seg(i).gap_len
      work_len =  seg(i).data_len

      IF work_len > 0 THEN

         MID$ (work_line, work_pos, work_len) = _
            MID$ (orig_line, seg(i).data_pos, work_len)

         c = MID$ (orig_attr, seg(i).data_pos, work_len)
         c = LSET$ (c, work_len)                    '/ force matching length
         MID$ (work_attr, work_pos, work_len) = c      '/ update color line
         work_pos += work_len
      END IF
   NEXT

   IF orig_line = work_line THEN
      shift_error = 1                     '/ shift did not result in any change
      shift_change = 0                     '/ caller should not modify the line

   ELSE
      '/ leave shift_error as-is, in case it was set on elsewhere

      shift_change = 1                         '/ caller should modify the line
      orig_line = work_line
      orig_attr = work_attr
   END IF
   MExit
END SUB ' slCmdDataShiftResult

SUB sLCmdDataShiftRight(BYREF orig_line         AS STRING, _
                        BYREF orig_attr         AS WSTRING, _
                        BYVAL shift_amount      AS LONG, _
                        BYVAL left_bound        AS LONG, _
                        BYVAL right_bound       AS LONG, _
                        BYREF shift_error       AS LONG, _
                        BYREF shift_change      AS LONG)

'/-----------------------------------------------------------------------------
'/ lCmdDataShiftRight
'/
'/ Accepts data line, bounds, and right-shift amount.
'/
'/ Returns shifted data as result, plus flag if "Data shifting incomplete"
'/ has occurred.  A revised color line is also returned.
'/-----------------------------------------------------------------------------
DIM seg (0 TO 1)        AS DataShift_Segment_t           '/ REDIM'd by Setup

LOCAL right_edge, space_width, delta, remaining_delta, remaining_room AS LONG
LOCAL right_bound_temp, s, seg_count  AS LONG
   MEntry
   shift_error = 0
   shift_change = 0

   '/--------------------------------------------------------------------------
   '/ for a right data-shift, if we are running in BOUND MAX mode, add the
   '/ shift amount to the current length so we can make the line longer.
   '/ if we shift-right by n, the new line can't be any longer than 'n'
   '/ more than it already is, but it could possibly increase in size by
   '/ less than that (including zero, if the shift is 'taken up' by other
   '/ segments within the line.
   '/ we check shift_amount > 0 to be sure we don't mask any parameter errors
   '/--------------------------------------------------------------------------

   IF TP.PrfBndRight = 0 AND shift_amount > 0 THEN
      right_bound_temp = LEN (orig_line) + shift_amount
   ELSE
      right_bound_temp = right_bound
   END IF

   '/--------------------------------------------------------------------------
   '/ perform setup common to all < > << >> shifts
   '/ if Setup fails, pass original line back as is without reporting an error
   '/--------------------------------------------------------------------------
   seg_count = sLCmdDataShiftSetup                                            _
   (  orig_line                                                               _
   ,  shift_amount                                                            _
   ,  left_bound                                                              _
   ,  right_bound_temp                                                        _
   ,  space_width                                                             _
   ,  seg ()                                                                  _
      )                                                                     '''

   IF seg_count < 1 THEN
      '/ line is zero-length, entirely blank, or data configuration will not
      '/ permit a shift to take place
      MExitSub
   END IF

   '/--------------------------------------------------------------------------
   '/ for all segments after the first one, attempt to collapse as many blanks
   '/ as possible, as long as doing to does not reduce the leading spaces by
   '/ less than space_width, nor removing more total spaces than then number of
   '/ columns in the shift_amount.  the process continues until the complete
   '/ number of columns have been shifted, or there are no more segments with
   '/ which to apply them.
   '/--------------------------------------------------------------------------
   remaining_delta = shift_amount

   FOR s = 2 TO seg_count
      IF remaining_delta = 0 THEN EXIT FOR

      '/-----------------------------------------------------------------------
      '/ the maximum possible number of leading blanks we can remove cannot
      '/ exceed space_width, unless the segment is a final one that contains
      '/ only blanks and no data.  when that exists, all the blanks can be
      '/ taken.
      '/-----------------------------------------------------------------------
      IF s = seg_count        _
      AND seg(s).gap_len > 0 _
      AND seg(s).data_len = 0 THEN
         '/ this is the last segment within bounds, and it's all blanks
         delta = seg(s).gap_len
      ELSE
         '/ this is a non-last segment, so don't take more than space_width
         delta = seg(s).gap_len - space_width
      END IF

      '/ unless the potential blanks is > 0 we cannot alter this segment
      IF delta > 0 THEN
         '/--------------------------------------------------------------------
         '/ if there are more leading blanks available than we need, take only
         '/ as many as needed; otherwise only take the possible amount
         '/--------------------------------------------------------------------
         IF delta > remaining_delta THEN
            delta = remaining_delta
         END IF

         '/--------------------------------------------------------------------
         '/ move this segment closer to the one preceding it, while at the same
         '/ time moving segment 1 to the right by the same amount
         '/--------------------------------------------------------------------
         seg(s).gap_len -= delta                               '/ rob Peter ...
         seg(1).gap_len += delta                             '/ to pay Paul ...
         remaining_delta -= delta
      END IF
   NEXT

   '/--------------------------------------------------------------------------
   '/ after distributing the spaces as best we can, if there are any spaces
   '/ not distributed, it is a potential shift error.  the "tail" segment of
   '/ the line is at seg(seg_count+1). if this seg has a data length, then
   '/ the tail is not movable, and we have a shift error.  if seg(seg_count+1)
   '/ has a data len of 0, it was at the end of the bounds or end of line.
   '/
   '/ if that is so, and we are running with BOUNDS MAX, then we will assume
   '/ that the line is extensible, and all remaining blanks will be applied to
   '/ the first segment; otherwise a shift error is reported, and we have
   '/ shifted the line as much as it's ever going to get shifted.
   '/--------------------------------------------------------------------------
   IF remaining_delta > 0 THEN                           '/ shift is incomplete
      right_edge = 0
      FOR s = 0 TO seg_count + 1
         right_edge += seg(s).gap_len + seg(s).data_len
      NEXT

      IF seg(seg_count+1).data_len > 0 THEN             '/ tail set not movable
         shift_error = 1

      ELSEIF right_edge < right_bound THEN
         '/ line is shorter than the right bound.  if the difference between
         '/ the two is enough for the remaining shift amount, use it and shift
         '/ is successful.  if not, use as much as is available and report that
         '/ the shift if incomplete
         remaining_room = right_bound - right_edge

         IF remaining_room >= remaining_delta THEN
            seg(1).gap_len += remaining_delta
         ELSE
            seg(1).gap_len += remaining_room
            shift_error = 1
         END IF

      ELSEIF TP.PrfBndRight > 0 THEN    '/ BOUNDS MAX not in effect
         shift_error = 1                '/ tail is null but line not extensible
      ELSE
         '/ first segment gets all remaining blanks, and pushes line to right
         seg(1).gap_len += remaining_delta
      END IF
   END IF

   '/ after seg table has been adjusted, reconstruct and return a new data line
   sLCmdDataShiftResult (  _
      orig_line,           _
      orig_attr,           _
      seg (),              _
      seg_count,           _
      shift_error,         _
      shift_change)
   MExit
END SUB ' slCmdDataShiftRight

FUNCTION sLCmdDataShiftSetup(BYVAL orig_line         AS STRING, _
                            BYVAL shift_amount      AS LONG, _
                            BYVAL left_bound        AS LONG, _
                            BYVAL right_bound       AS LONG, _
                            BYREF space_width       AS LONG, _
                            BYREF seg ()            AS DataShift_Segment_t)  AS LONG
'/-----------------------------------------------------------------------------
'/ lCmdDataShiftSetup
'/
'/ Perform processing common to < > << and >> commands
'/
'/ If errors, return 0 else return number of segments found
'/ segment 0 will contain the number of trailing blanks
'/-----------------------------------------------------------------------------
LOCAL i, s, n, seg_count, left_edge, right_edge     AS LONG
LOCAL test_line, c, quote               AS STRING
LOCAL gap_len, data_pos, data_end, data_len, span   AS LONG
LOCAL blank_pos, blank_len AS LONG
LOCAL check_escape  AS LONG
   MEntry
   '/--------------------------------------------------------------------------
   '/ validate parameters.  part of this is to ensure calling code is correct,
   '/ and part is to filter out situations where we just can't shift anything.
   '/--------------------------------------------------------------------------
   n = LEN (orig_line)
   IF n = 0 OR _                             '/ can't shift-left a null line
      shift_amount < 1 OR _                  '/ shift amount is illegal
      left_bound < 1 OR _                    '/ bound location doesn't make sense
      left_bound >= right_bound OR _         '/ bound location doesn't make sense
      n < left_bound THEN                    '/ data is not in the bounded area
      FUNCTION = 0                           '/ unable to perform data shift
      MExitFunc
   END IF

   right_edge = MIN(n, right_bound)
   IF VERIFY (MID$ (orig_line, left_bound TO right_edge), " ") = 0 THEN
      '/ entire line within eligible area is blank
      FUNCTION = 0                              '/ unable to perform data shift
      MExitFunc
   END IF

   '/--------------------------------------------------------------------------
   '/ fetch overriding size of gap, and escape usage, if SET names defined
   '/--------------------------------------------------------------------------
   space_width = sGet_Set_number ("OPT.DS.MINSIZE")
   IF space_width < 1 OR space_width > 9 THEN space_width = 1
   check_escape = sGet_Set_number ("OPT.DS.ESCAPE")
   IF check_escape <> 0 THEN check_escape = 1

   '/--------------------------------------------------------------------------
   '/ copy orig_line to test_line, changing unusable blanks into underscores
   '/ the test line is just used for calculating segment positioning, while
   '/ the orig_line remains the correct source of data
   '/
   '/ when spans of blanks are inside quoted strings, they are unusable.
   '/ however, according to ISPF logic, if a quoted string is never properly
   '/ closed, the (leading) quote is simply ignored.  so, we have to do a
   '/ lookahead to make sure the opening quote has a trailing quote, otherwise
   '/ the string is not closed and we have to treat the quote as ordinary data.
   '/--------------------------------------------------------------------------
   test_line = orig_line
   i = 1
   quote = " "

   DO WHILE i <= n
      c = MID$(test_line, i, 1)
      IF quote = " " THEN                      '/ not currently inside a string
         IF c = $DQ OR c = "'" THEN
            IF sString_is_proper (test_line, i, check_escape) THEN
               quote = c
            END IF

         ELSEIF c = " " THEN
            IF i < left_bound OR i > right_bound THEN
               MID$(test_line, i, 1) = "_"             '/ blanks outside bounds
            END IF
         END IF

      ELSE                                                    '/ inside a quote
         IF c = $DQ OR c = "'" THEN
            quote = " "                                '/ quote has been closed
         ELSEIF c = " " THEN
            MID$(test_line, i, 1) = "_"       '/ quoted blanks are not eligible
         ELSEIF check_escape = 1 AND c = "\" THEN  '/ escaped char, maybe quote
            i += 1                                          '/ skip over escape
            IF i <= n THEN
               IF MID$(test_line, i, 1) = " " THEN _
                  MID$(test_line, i, 1) = "_"       '/ ignore any escaped blank
            END IF
         END IF
      END IF
      i += 1
   LOOP

   '/--------------------------------------------------------------------------
   '/ spans of blanks that are still eligible from the logic above, but are
   '/ shorter than space_width are treated as if not blank at all.  this only
   '/ applies when space_width > 1; otherwise all blanks spans are eligible.
   '/
   '/ if there is no label at the left_bound, but there is a span of spaces
   '/ but (possibly) less than the space_width, that initial part of the line
   '/ can be detached during a shift right. so, any initial leading blanks are
   '/ skipped, even if shorter than shift_width.
   '/
   '/ these added considerations don't exist in ISPF, because they only have a
   '/ fixed space-width of 1, while this code supports a variable space_width.
   '/
   '/ if there is a short span starting at the left bound, allow it.  this
   '/ makes it possible to detach from the left boundary during a right shift
   '/ if the space_width is > 1.
   '/
   '/ columns to the left of left_bound have already been handled.
   '/--------------------------------------------------------------------------
   IF space_width > 1 THEN
      i = left_bound

      '/ skip over any initial non-label blanks, regardless of size
      DO WHILE i <= n AND MID$(test_line, i, 1) = " "
         i += 1
      LOOP

      DO WHILE i <= n
         IF MID$(test_line, i, 1) = " " THEN
            span = sLeadingBlanks (test_line, i)
            IF span < 1 THEN EXIT DO                                '/ failsafe
            IF span < space_width THEN                            '/ short span

               '/ short spans are ineligible, except if starting at left bound
               IF i > left_bound THEN
                  MID$(test_line, i, span) = STRING$ (span, "_")
               END IF
            END IF
            i += span                    '/ skip over span whether short or not

         ELSE
            i += 1
         END IF
      LOOP
   END IF

   '/--------------------------------------------------------------------------
   '/ count number of segments, and allocate seg table.
   '/--------------------------------------------------------------------------
   i = 1
   s = 0

   DO WHILE i <= right_edge AND MID$(test_line, i, 1) <> " "
      i += 1                                             '/ skip over nonblanks
   LOOP

   DO WHILE i <= right_edge
      span = 0
      DO WHILE i <= right_edge AND MID$(test_line, i, 1) = " "
         i += 1                                             '/ skip over blanks
         span = 1
      LOOP

      DO WHILE i <= right_edge AND MID$(test_line, i, 1) <> " "
         i += 1                                          '/ skip over nonblanks
         span = 1
      LOOP

      IF data_pos > 0 THEN
         data_len = data_end - data_pos + 1
      END IF

      IF span = 1 THEN
         s += 1
      END IF
   LOOP

   '/--------------------------------------------------------------------------
   '/ allocate the segment table.  the +3 accounts for the "tail" segment, plus
   '/ a couple extra, just in case (being paranoid) so we don't overrun the
   '/ table with any bad math ... and yes, it's a fudge factor.
   '/--------------------------------------------------------------------------
   REDIM seg (0 TO s + 3) AS DataShift_Segment_t

   '/--------------------------------------------------------------------------
   '/ determine "label" prefix.  if found, store as segment 0 which cannot be
   '/ moved.
   '/--------------------------------------------------------------------------
   seg(0).gap_len = 0                '/ segment 0 will never have leading space
   seg(0).blank_len = 0
   seg(0).blank_pos = 0
   left_edge = left_bound

   '/--------------------------------------------------------------------------
   '/ the "label" ends at the first blank
   '/ because of preprocessing done previously, only spans of blanks that are
   '/ big enough will have true blanks in the test_line.
   '/--------------------------------------------------------------------------
   DO WHILE left_edge <= right_edge AND MID$(test_line, left_edge, 1) <> " "
      left_edge += 1
   LOOP

   IF left_edge >= right_edge THEN         '/ label fills up entire bounds area
      FUNCTION = 0                              '/ unable to perform data shift
      MExitFunc

   ELSEIF left_edge > left_bound THEN                '/ a label found in bounds
      seg(0).data_pos = 1
      seg(0).data_len = left_edge - 1

   ELSEIF left_bound > 1 THEN              '/ unmovable segment precedes bounds
      seg(0).data_pos = 1
      seg(0).data_len = left_bound - 1

   ELSE                                         '/ there is no unmovable prefix
      seg(0).data_pos = 0
      seg(0).data_len = 0

   END IF

   '/--------------------------------------------------------------------------
   '/ parse the remaining line, breaking it up into spans of blanks followed by
   '/ non_blanks, where the non_blanks are counted first.  the code above will
   '/ ensure that left_edge starts out pointing to a blank.
   '/--------------------------------------------------------------------------
   i = left_edge
   seg_count = 0

   DO WHILE i <= right_edge
      gap_len = 0
      data_pos = 0
      data_len = 0
      blank_pos = 0

      DO WHILE i <= right_edge AND MID$(test_line, i, 1) = " "
         IF blank_pos = 0 THEN blank_pos = i
         gap_len += 1
         i += 1                                             '/ skip over blanks
      LOOP

      DO WHILE i <= right_edge AND MID$(test_line, i, 1) <> " "
         data_end = i
         IF data_pos = 0 THEN data_pos = i
         i += 1                                          '/ skip over nonblanks
      LOOP

      IF data_pos > 0 THEN
         data_len = data_end - data_pos + 1
      END IF

      IF gap_len > 0 OR data_len > 0 THEN
         seg_count += 1

         seg(seg_count).gap_len  = gap_len
         seg(seg_count).data_pos = data_pos
         seg(seg_count).data_len = data_len

         '/ we need original pos/len of blanks to propagate color info
         seg(seg_count).blank_len = gap_len
         seg(seg_count).blank_pos = blank_pos
      END IF
   LOOP

   '/--------------------------------------------------------------------------
   '/ if there was a non-default bounds and line extends to the right of it,
   '/ it defines a 'tail' segment that is unmovable.  if so, define it here,
   '/ otherwise create a null tail segment.  we disregard any blanks in the
   '/ tail, and simply record the data area extent of the tail.
   '/--------------------------------------------------------------------------
   seg(seg_count+1).gap_len   = 0
   seg(seg_count+1).blank_len = 0
   seg(seg_count+1).blank_pos = 0

   IF n > right_edge THEN
      seg(seg_count+1).data_pos = right_edge + 1
      seg(seg_count+1).data_len = n - right_edge
   ELSE
      seg(seg_count+1).data_pos = 0
      seg(seg_count+1).data_len = 0
   END IF
   FUNCTION = seg_count                                        '/ normal return
   MExit
END FUNCTION ' slCmdDataShiftSetup

FUNCTION sLeadingBlanks(BYVAL input_str AS STRING, BYVAL start_pos AS LONG) AS LONG                                          '''
'/-----------------------------------------------------------------------------
'/ leadingBlanks
'/
'/ tally input_str from start_pos to end for leading blanks
'/-----------------------------------------------------------------------------

LOCAL i, n, result AS LONG
   MEntry
   n = LEN(input_str)
   IF n = 0 OR start_pos > n OR start_pos < 1 THEN
      FUNCTION = 0
      MExitFunc
   END IF
   result = 0
   i = start_pos
   DO WHILE i <= n AND MID$(input_str, i, 1) = " "
      result += 1
      i += 1
   LOOP
   FUNCTION = result
   MExit
END FUNCTION

FUNCTION sLNKConvert(fname AS STRING) AS STRING
'---------- Convert a filename for LNK lookup
LOCAL CLSID_ShellLink, IID_IShellLink AS GUIDAPI
LOCAL CLSCTX_INPROC_SERVER, Flags, lResult AS DWORD
LOCAL FileData AS WIN32_FIND_DATA
LOCAL IID_Persist AS STRING * 16, pp, ppf, psl AS DWORD PTR
LOCAL outvalue, TmpAsciiz AS ASCIIZ * %MAX_PATH
LOCAL TmpWide AS ASCIIZ * (2 * %MAX_PATH)

   MEntry
   POKE$ VARPTR(CLSID_ShellLink), MKL$(&H00021401) + CHR$(0, 0, 0, 0, &HC0, 0, 0, 0, 0, 0, 0, &H46)
   POKE$ VARPTR(IID_IShellLink), MKL$(&H000214EE) + CHR$(0, 0, 0, 0, &HC0, 0, 0, 0, 0, 0, 0, &H46)
   IID_Persist = MKL$(&H0000010B) + CHR$(0, 0, 0, 0, &HC0, 0, 0, 0, 0, 0, 0, &H46)
   CLSCTX_INPROC_SERVER = 1

   IF ISFALSE (CoCreateInstance(CLSID_ShellLink, BYVAL %NULL, CLSCTX_INPROC_SERVER, IID_IShellLink, psl)) THEN
      pp = @psl: CALL DWORD @pp USING Sub3(BYVAL psl, IID_Persist, ppf) TO lResult
      TmpAsciiz = fname
      MultiByteToWideChar %CP_ACP, 0, TmpAsciiz, %MAX_PATH, BYVAL VARPTR(TmpWide), 2 * %MAX_PATH
      pp = @ppf + 20: CALL DWORD @pp USING Sub3(BYVAL ppf, TmpWide, BYVAL %TRUE)
      pp = @psl + 12: CALL DWORD @pp USING Sub5(BYVAL psl, outvalue, BYVAL %MAX_PATH, FileData, Flags)  'GetFilePath
      pp = @ppf + 8: CALL DWORD @pp USING Sub1(BYVAL ppf)         'Release the persistant file
      pp = @psl + 8: CALL DWORD @pp USING Sub1(BYVAL psl)         'Unbind the shell link object from the persistent file
      FUNCTION = outvalue
   END IF
   MExit
END FUNCTION

FUNCTION sLoopHandler(BYREF lpEP AS MY_EXCEPTION_POINTERS) AS LONG
'----- Handle Loop trap exception
STATIC TerminateInProgress AS LONG
LOCAL ErrorRecord AS EXCEPTION_RECORD POINTER
LOCAL ErrorCode AS LONG POINTER
LOCAL MSG, fn  AS STRING, i, FNm AS LONG
   ErrorRecord =  lpEP.ExceptionRecord                            ' Get address of Exception record
   ErrorCode = @ErrorRecord.ExceptionCode                         ' Get the exception code
   IF TerminateInProgress THEN                                    ' Oops, been here already
       MSGBOX "A second serious error has occured, no recovery is possible", _
              %MB_ICONERROR OR %MB_TASKMODAL OR %MB_DEFBUTTON1, "SPFLite Crash Intercept"
      SETUNHANDLEDEXCEPTIONFILTER %NULL                           ' Deactivate our handler
   ELSE                                                           ' Lets see if it's ours
      FNm = FREEFILE                                              ' Get file number
      fn = ENV.INIPath + "SPFLiteCrash."                          ' Build a filename
      fn += RIGHT$(DATE$, 4) + LEFT$(DATE$, 2) + MID$(DATE$, 4, 2)' Add the date
      fn += LEFT$(TIME$, 2) + MID$(TIME$, 4, 2) + RIGHT$(TIME$, 2) + RIGHT$(FORMAT$(sOneSecondTimer), 2) ' Add the time
      fn += ".txt"                                                ' For saving this message
      MSG = "SPFLite has " + IIF$(Errorcode = 123456789, "detected an internal loop ", "encountered an execution exception (" + HEX$(Errorcode, 8) + ") ")
      MSG += $CRLF + $CRLF                                        '
      MSG += "Last Interactions were:" + $CRLF                    '
      MSG += "  KB     Primitive: " + gCrashLastPrim + $CRLF      '
      MSG += "  Line       Cmnd: " + gCrashLastLCmd + $CRLF       '
      MSG += "  Primary Cmnd: " + gCrashLastPCmd + $CRLF + $CRLF  '
      MSG += "Module Back Trace:" + $CRLF                         '
      FOR i = gCrashCtr - 1 TO 0 STEP -1
            MSG += " " + FORMAT$(i, "00") + " | " + gCrashList(i) + $CRLF ' Add to message
      NEXT i                                                      '
      MSG += $CRLF + "     Note: A copy of this message is in: " + fn + $CRLF + $CRLF
      IF Errorcode = 123456789 THEN                               ' The Loop prompt
         MSG += "Select |KOK|B, to ignore loop and continue execution, or" + $CRLF + _
                "            Cancel to terminate the program"     '
         i = sDoMsgBox(MSG, %MB_USERICON OR %MB_OKCANCEL, "SPFLite Loop Intercept")
         IF i = %IDOK THEN                                        ' Continue?
            gLoopCtr = 0                                          ' Reset loop counter
            FUNCTION = %EXCEPTION_CONTINUE_EXECUTION              '
         ELSE                                                     '
            TerminateInProgress = %True                           ' Once only please
            FUNCTION = %EXCEPTION_EXECUTE_HANDLER                 ' Continue termination
         END IF                                                   '
      ELSE                                                        '
         sDoMsgBox MSG, %MB_USERICON OR %MB_OK, "SPFLite Crash Intercept"
         TerminateInProgress = %True                              ' Once only please
         FUNCTION = %EXCEPTION_EXECUTE_HANDLER                    ' Continue termination
      END IF                                                      '

      '----- Write the crash text file
      OPEN fn FOR OUTPUT AS #FNm                                  ' Open the output File
      PRINT #FNm, MSG                                             ' Write the msg
      SETEOF #FNm                                                 '
      CLOSE #FNm                                                  '
   END IF                                                         '
END FUNCTION

THREAD FUNCTION sLoopThread(BYVAL dummy AS LONG POINTER) AS LONG
'---------- See if it looks like we're looping
   DO
      SLEEP 1000                                                  ' Wait 1 second
      IF ISFALSE gLoopFlag OR ISTRUE gfDoingMsg THEN              ' No transaction in progress or external Dialog?
         RESET gLoopCtr                                           ' Reset the counter
      ELSEIF gLoopCtr = -1 THEN                                   ' Suppressed?
         ' Do nothing                                             ' Do nothing
      ELSE                                                        '
         INCR gLoopCtr                                            ' Count another, transaction still running
         IF gLoopCtr > 10 + (TP.LastLine / 750000) THEN           ' Do we consider this a loop? (10 seconds + 1 per 750,000 lines)
            RaiseException 123456789, %NULL, %NULL, %NULL         ' Trigger a unique error
         END IF                                                   '
      END IF                                                      '
   LOOP                                                           ' On and on
END FUNCTION

SUB slPrintUnicodeInit (BYREF u AS utab_t)
'---------- Initialize "utab" by finding the .Unicode file and reading it
LOCAL fileName AS STRING
   fileName = sUnicodeGetTableName (ENV.INIPath, "SPFLite", "Print")
   sUnicodeGetTable (fileName, u)
END SUB

SUB      sMakeNullFile(fn AS STRING)
'---------- Create an empty physical file
LOCAL FNm AS LONG
   FNm = FREEFILE                                                 ' Create it as an empty file
   OPEN fn FOR OUTPUT AS #FNm                                     ' Open the File
   SETEOF #FNm                                                    ' Set EOF
   CLOSE #FNm                                                     ' Close it
END SUB

FUNCTION sMakePrettySizeLarge(vl AS QUAD) AS STRING
'---------- Make a 15 char 'pretty' size field
   FUNCTION = RSET$(FORMAT$(vl, "0,"), 15)                        '
END FUNCTION

FUNCTION sMakePrettySizeSmall(vl AS QUAD) AS STRING
'---------- Make a 6 char 'pretty' size field
LOCAL t AS STRING
   SELECT CASE AS vl                                              ' Select which format
      CASE < 1000:          t = FORMAT$(vl, "* ####"): FUNCTION = IIF$(t = "     0", "      ", t)
      CASE < 1048576:       FUNCTION = FORMAT$(vl / 1024, "* #.0\K")
      CASE < 1073741824:    FUNCTION = FORMAT$(vl / 1048576, "* #.0\M")
      CASE ELSE:            FUNCTION = FORMAT$(vl / 1073741824, "* #.0\G")
   END SELECT                                                     '
END FUNCTION

FUNCTION sMakePrettyTime(vl AS QUAD) AS STRING
'---------- Make a 'pretty' TimeStamp
LOCAL LTime AS IPOWERTIME                                         ' Create a PowerTime object
LET LTime = CLASS "PowerTime"                                     '
   LTime.FileTime = vl                                            ' Assign the passed FILETIME QUAD to it
   LTime.ToLocalTime                                              '
   FUNCTION = FORMAT$(LTime.year(), "0000") + "-" + FORMAT$(LTime.Month(), "00") + "-" + FORMAT$(LTime.day(), "00") + "  " + _
            FORMAT$(LTime.Hour(), "00") + ":" + FORMAT$(LTime.Minute(), "00") + ":" + FORMAT$(LTime.Second(), "00")
END FUNCTION                                                      '

FUNCTION sMarkSimple(mk AS STRING) AS STRING
'---------- Simplify the new Mark string
LOCAL nmk AS STRING, i, j AS LONG
   MEntry
   nmk = SPACE$(LEN(mk) + 1)                                      ' Get working copy of blanks
   j = 1
   i = INSTR(1, mk, ANY "<*>")                                    ' Look for special chars
   DO WHILE i                                                     ' While a special char left
      SELECT CASE AS CONST$ MID$(mk, i, 1)                        ' Which one?
         CASE "<", "*"                                            ' < or *
            MID$(nmk, i, 1) = "*": j = i + 1                      ' Just transfer it
         CASE ">"                                                 ' >
            MID$(nmk, i + 1, 1) = "*": j = i + 1                  ' Put * in next column
      END SELECT                                                  '
      i = INSTR(j, mk, ANY "<*>")                                 ' Look for special chars
   LOOP                                                           '
   FUNCTION = nmk                                                 ' Pass back the answer
   MExit
END FUNCTION

FUNCTION sNumDaysSince(OldDate AS STRING) AS LONG
'---------- Calc # days since provided date (mm-dd-yyyy)
LOCAL d, m, y, day1, day2 AS LONG, yy AS DOUBLE
   MEntry
   d = VAL(MID$(OldDate, 4, 2))                                   ' Convert passed date
   m = VAL(LEFT$(OldDate, 2))                                     '
   y = VAL(MID$(OldDate, 7, 4))                                   '
   yy = y + (m - 2.85) / 12                                       ' Convert to AstroDay
   Day1 = INT(INT(INT(367 * yy) - 1.75 * INT(yy) + d) -0.75 * INT(0.01 * yy)) + 1721119

   d = VAL(MID$(DATE$, 4, 2))                                     ' Convert today
   m = VAL(LEFT$(DATE$, 2))                                       '
   y = VAL(MID$(DATE$, 7, 4))                                     '
   yy = y + (m - 2.85) / 12                                       ' Convert to AstroDay
   Day2 = INT(INT(INT(367 * yy) - 1.75 * INT(yy) + d) -0.75 * INT(0.01 * yy)) + 1721119
   sNumDaysSince = Day2 - Day1                                    ' Pass back difference
   MExit
END FUNCTION

FUNCTION sOneSecondTimer  AS DWORD
'---------- Get time in hundredths of a second
STATIC sdwLastTime AS DWORD
STATIC slRollOvers AS LONG
STATIC dwTimeNow   AS DWORD
   dwTimeNow = GetTickCount
   IF dwTimeNow < sdwLastTime THEN INCR slRollOvers               ' GetTickCount has rolled over at 49.710 days.
   sdwLastTime = dwTimeNow
   FUNCTION = (dwTimeNow + (slRollOvers * (%MAXDWORD + 1))) \ 10&&
   'Change divisor to 100&& to return tenths of seconds.
   'Change divisor to 10&& to return hundredths of seconds.
END FUNCTION

FUNCTION sOnOff(prm AS LONG) AS LONG
'---------- Convert On/Off to True/False
   MEntry
   IF pCmdNumOps < prm THEN FUNCTION = 1: MExitFunc               ' No parm exists? = ON
   IF INSTR(CHR$(%KWON, %KWOFF),CHR$(pCmdOpsType(prm))) = 0 THEN  ' Invalid?
      FUNCTION = -1                                               ' Flag error
      scError(%eFail,"Unknown OFF / ON operand - " + pCmdOps(prm))' Issue error message
      MExitFunc                                                   '
   END IF                                                         '
   FUNCTION = IIF(pCmdOpsType(prm) = %KWON, 1, 0)                 ' Pass back the On/Off
   MExit
END FUNCTION

FUNCTION sOpenPrinter(Setup AS STRING) AS INTEGER
'---------- Open the Default Printer
LOCAL i, lclppix, lclppiy AS LONG, lcllm, lclrm, lcltm, lclbm AS SINGLE
LOCAL fIndex AS LONG, fList AS STRING, t, fTable() AS STRING
   MEntry
   IF Setup = "SETUP" THEN                                        ' SETUP requested?
      DispPrint()                                                 ' Go let user set them
      FUNCTION = %True                                            '
      MExitFunc                                                   '
   END IF                                                         '

   IF ISNULL(ENV.PrtName) THEN                                    '
      scError(%eFail, "Printer SETUP has not been completed yet") ' Tell user to do SETUP
      gPrinterOpen = %False                                       '
      FUNCTION = %False                                           ' Set a no-go default
      MExitFunc                                                   '
   END IF                                                         '

   gLoopCtr = - 1                                                 ' Prevent loop detection
   XPRINT CANCEL                                                  ' Just in case?
   XPRINT CLOSE                                                   '
   XPRINT ATTACH ENV.PrtName, "SPFLite Print"                     ' Attach the printer
   IF ERR = 0 AND LEN(XPRINT$) > 0 THEN                           ' OK?
      XPRINT GET DUPLEX TO i                                      ' Duplex supportable?
      IF ISTRUE i THEN XPRINT SET DUPLEX ENV.PrtDuplex            ' Set duplex
      XPRINT GET DUPLEX TO i                                      ' Get it again
      ENV.PrtDuplex = i                                           '
      IF ENV.PrtDuplex = 0 THEN ENV.PrtDuplex = 1                 ' Eliminate zero case
      XPRINT SET ORIENTATION ENV.PrtOrient                        '
      XPRINT GET MARGIN TO lcllm, lcltm, lclrm, lclbm
      XPRINT GET PPI TO lclppix, lclppiy
      FONT NEW ENV.PrtFontName, VAL(ENV.PrtFontPitch), VAL(ENV.PrtFontStyle), 1, 1 TO gPFontHndl ' Create the font
      XPRINT SET FONT gPFontHndl                                  ' Set it
      XPRINT CHR SIZE TO gPCharWidth, gPCharHeight                ' Get size of a character
      XPRINT GET CLIENT TO gPPageWidth, gPPageHeight              ' Get page size
      gPCpl = gPPageWidth \ gPCharWidth                           ' Calc characters per line
      gPLpp = gPPageHeight \ gPCharHeight                         ' Calc lines per page
      gPTFill = ((ENV.PrtTMargin * lclppiy) - lcltm) \ gPCharHeight' Filler lines at top
      gPLpp -= gPTFill                                            ' Adjust line count
      gPLpp -= ((ENV.PrtBMargin * lclppiy) - lclbm) \ gPCharHeight' Again for bottom margin
      gPLFill = ((ENV.PrtLMargin * lclppix) - lcllm) \ gPCharWidth' Filler chars at left
      gPCpl -= gPLFill                                            ' Adjust line length
      gPCpl -= ((ENV.PrtRMargin * lclppix) - lclRm) \ gPCharWidth ' Again for right margin
      gPRFill = ((ENV.PrtRMargin * lclppix) - lclRm) \ gPCharWidth' Filler chars at right
      gPColor = ENV.PrtPColor                                     ' Printing in color?
      XPRINT SCALE (0, 0) - (gPCpl, gPLpp)                        ' Set page scale to characters
      XPRINT SET PAPER ENV.PrtPaper                               ' Set Paper type
      XPRINT GET PAPERS TO fList                                  ' Get string of forms available
      REDIM fTable(1 TO PARSECOUNT(fList)) AS STRING              ' Build a table
      PARSE fList, fTable()                                       '
      FOR i = 1 TO UBOUND(fTable) STEP 2                          '
         IF VAL(fTable(i)) = ENV.PrtPaper THEN                    ' Found our entry?
            gPrtPaper = fTable(i + 1): EXIT FOR                   ' Save paper name
         END IF                                                   '
      NEXT i                                                      '
      gPrinterOpen = %True                                        ' Set the internal flag showing printer is open
      FUNCTION = %True                                            '
   ELSE                                                           '
      gPrinterOpen = %False                                       '
      FUNCTION = %False                                           ' Set a no-go default
   END IF                                                         '
   MExit
END FUNCTION

FUNCTION sParseProfile(fn AS STRING) AS STRING
'---------- Return Profile for a filename
LOCAL i, j AS LONG, pname AS STRING
   MEntry
   i = INSTR(-1, fn, ".")                                         ' Get last . in filename
   j = INSTR(-1, fn, "\")                                         ' Get last \ in filename
   IF j > i THEN i = 0                                            ' if \ right of . then no . (e.g. clist.s\fname)
   IF i = 0 THEN                                                  ' If no extension
      IF ISFALSE ENV.DirProfFlag THEN                             ' May we use the DIR name for profile?
         FUNCTION = "DEFAULT"                                     ' No, Request the default
      ELSE                                                        '
         j = INSTR(-1, fn, "\")                                   ' Get last \ in filename
         IF j = 0 THEN j = INSTR(-1, fn, ":")                     ' 2nd chance, try for a ':'
         IF j THEN                                                ' We got a Dir level
            pname = MID$(fn, j + 1)                               ' Extract last DIR level
            j = INSTR(pname, " ")                                 ' Better not be multiple words words
            IF j THEN                                             ' It is?
               IF j <> 1 THEN                                     ' and not in col 1
                  FUNCTION = LEFT$(pname, j - 1)                  ' Return the first word
               ELSE                                               ' Something weird here
                  FUNCTION = "DEFAULT"                            ' Go back to the default
               END IF                                             '
            END IF                                                '
         ELSE                                                     ' That was the last chance
            FUNCTION = "DEFAULT"                                  ' Use the default
         END IF                                                   '
      END IF                                                      '
   ELSE                                                           ' Else there IS an extension
      FUNCTION = MID$(fn, i + 1)                                  ' Return what's there
   END IF                                                         '
   MExit
END FUNCTION

SUB sParse_Keyword_Value_Data(KV AS KEYWORD_VALUE_DATA_T)
'/-----------------------------------------------------------------------------/
'/  Parse_Keyword_Value_Data                                                   /
'/                                                                             /
'/  Keyword_Value_Data_T.KV_Data contains one or more pairs of                 /
'/  Keyword=Value strings.  we scan left to tight looking for the first such   /
'/  pair, split the keyword and value into separate strings and store them.    /
'/  Once a Keyword=Value is found, the KV_Data string is truncated on the left /
'/  and this continues until the entire KV_Data string is consumed.            /
'/-----------------------------------------------------------------------------/
LOCAL C, QUOTE              AS STRING
LOCAL I, LEN_IN_DATA        AS LONG
LOCAL START_KEYWORD, START_VALUE, END_VALUE, NEXT_KEYWORD AS LONG

    LEN_IN_DATA = LEN(KV.KV_DATA)

    '/ Initialize returned structure
    KV.KV_KEYWORD           = ""
    KV.KV_VALUE             = ""
    KV.KV_Reason            = ""
    KV.KV_STATUS            = 0         '/ 0 means no more values left to return

    '/ Find start of keyword
    C = " "                                                        '/ Initialize
    FOR I = 1 TO LEN_IN_DATA
        C = UUCASE(MID$(KV.KV_DATA, I, 1))                        '/ Uses UUCASE
        IF  C = " " THEN ITERATE FOR                      '/ Skip leading blanks
        START_KEYWORD = I
        EXIT FOR
    NEXT

    '/ If line is blank or starts with an * as an end of line
    '/ comment, we have reached and of buffer; parsing complete
    IF  (START_KEYWORD = 0) OR (C = "*") THEN
        KV.KV_STATUS = 0                                     '/ Parsing complete
        EXIT SUB
    END IF

    '/ At this point, C should be the first letter of the keyword
    '/ which must be letter, otherwise keyword format is invalid
    '/ keywords are forced to uppercaseE.
    IF  (C < "A") OR (C > "Z") THEN
        KV.KV_Reason = "KEYWORD STARTS WITH INVALID CHARACTER: " + C
        KV.KV_STATUS = -1                                  '/ Bad keyword format
        EXIT SUB
    END IF

    '/ Store characters of keyword until '=' found
    FOR I = START_KEYWORD TO LEN_IN_DATA
        C = UUCASE(MID$(KV.KV_DATA, I, 1))                        '/ Uses UUCASE
        IF  ( (C >= "A") AND (C <= "Z") ) _
        OR  ( (C >= "0") AND (C <= "9") ) _
        OR    (C  = "_")                  _
        OR    (C  = ".") THEN
            KV.KV_KEYWORD += C                  '/ Append char to keyword string
        ELSEIF C = "=" THEN
            START_VALUE = I + 1                    '/ Value starts after the '='
            EXIT FOR
        END IF
    NEXT

    IF  START_VALUE = 0 _
    OR  START_VALUE >= LEN_IN_DATA THEN
        KV.KV_Reason = "KEYWORD " + KV.KV_KEYWORD + " NOT FOLLOWED BY = SIGN"
        KV.KV_STATUS = -2                            '/ Value is missing or null
        EXIT SUB
    END IF

    C = MID$(KV.KV_DATA, START_VALUE, 1)
    IF  C <= " " THEN                                 '/ Nonblank value expected
        KV.KV_Reason = "KEYWORD " + KV.KV_KEYWORD + "= NOT FOLLOWED BY VALUE"
        KV.KV_STATUS = -2                            '/ Value is missing or null
        EXIT SUB
    END IF

    '/-------------------------------------------------------------------------/
    '/ Values can be quoted or unquoted.                                       /
    '/ Unquoted values are upper-cased, and end with space or end of string    /
    '/ Quoted values are not upper-cased, and end with matching quote          /
    '/-------------------------------------------------------------------------/
    IF  (C = $SQ) OR (C = $DQ) OR (C = "`") THEN
        QUOTE = C
        START_VALUE += 1

        '/---------------------------------------------------------------------/
        '/  For quoted string, Start_Value should point to first char of       /
        '/  value.  This means there must be at least 2 positions left: 1 for  /
        '/  the data and 1 for the ending quote.  Be sure there is enough      /
        '/  data left for them                                                 /
        '/---------------------------------------------------------------------/
        IF  START_VALUE >= LEN_IN_DATA THEN
            KV.KV_Reason = "KEYWORD " + KV.KV_KEYWORD                          _
                + " QUOTED VALUE MALFORMED"
            KV.KV_STATUS = -3                           '/ Error in quoted value
            EXIT SUB
        END IF
    ELSE
        QUOTE = " "
    END IF

    '/ Accumulate value string
    IF  QUOTE = " " THEN                          '/ Accumulate non-quoted value
        END_VALUE = LEN_IN_DATA                     '/ Default end if last value
        FOR I = START_VALUE TO LEN_IN_DATA
            C = UUCASE(MID$(KV.KV_DATA, I, 1))                    '/ Uses UUCASE
            IF  C = " " THEN
                END_VALUE = I                '/ A non-quoted value ends at space
                EXIT FOR
            END IF
            KV.KV_VALUE += C                      '/ Append char to value string
        NEXT
        NEXT_KEYWORD = END_VALUE + 1
    ELSE                                              '/ Accumulate quoted value
        FOR I = START_VALUE TO LEN_IN_DATA
            C = MID$(KV.KV_DATA, I, 1)                         '/ Without UUCASE
            IF  C = QUOTE THEN
                END_VALUE = I                    '/ A quoted value ends at quote
                EXIT FOR
            END IF
            KV.KV_VALUE += C                      '/ Append char to value string
        NEXT

        IF  END_VALUE = 0 THEN                        '/ Close quote never found
            KV.KV_Reason = "KEYWORD " + KV.KV_KEYWORD                          _
                + " QUOTED VALUE NOT CLOSED"
            KV.KV_STATUS = -3                           '/ Error in quoted value
            EXIT SUB
        END IF
        NEXT_KEYWORD = END_VALUE + 1

        '/ Ending quote must have one space after, unless last
        IF  NEXT_KEYWORD <= LEN_IN_DATA THEN
            C = MID$(KV.KV_DATA, NEXT_KEYWORD, 1)
            IF  C = " " THEN
                NEXT_KEYWORD += 1                    '/ Skip over trailing space
            ELSE
                KV.KV_Reason = "KEYWORD " + KV.KV_KEYWORD                      _
                    + " CLOSE QUOTE NOT FOLLOED BY SPACE"
                KV.KV_STATUS = -3                       '/ Error in quoted value
                EXIT SUB
            END IF
        END IF
    END IF

    '/ Chop off data value for next time
    IF  NEXT_KEYWORD > LEN_IN_DATA THEN
        KV.KV_DATA = ""                              '/ Data string was consumed
    ELSE
        KV.KV_DATA = MID$(KV.KV_DATA, NEXT_KEYWORD)
    END IF
    KV.KV_STATUS = 1                          '/ One Keyword-Value pair returned
END SUB ' sParse_Keyword_Value_Data

SUB      sPopReady()
'---------- Ready cursor etc. for a popup dialog
   IF gfDoingMsg = 0 THEN                                         ' If first time
      sCaretHide                                                  ' Get rid of cursor
      sCaretDestroy                                               '
   END IF                                                         '
   INCR gfDoingMsg                                                ' Tell KB hook to ignore things
END SUB

SUB      sPopReset()
'---------- Reset from a PopReady
   IF gfDoingMsg = 0 THEN EXIT SUB                                ' Nothing to do
   DECR gfDoingMsg                                                ' Decr count
   IF gfDoingMsg = 0 THEN                                         ' Gone to zero?
      sCaretCreate                                                ' Create it
      sDoCursor                                                   ' Position it
      sCaretShow                                                  ' Show the caret
   END IF
END SUB

SUB      sPrint(sTxt AS STRING, sAttr AS WSTRING, sRow AS LONG, sCol AS LONG, OPT sOffset AS LONG, OPT sPad AS LONG)
'---------- Print a string at a specific row/col using the Attr definition
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL lclScheme, lclUC, lclUL, lclBG, cPtr, oPtr, tLen AS LONG
LOCAL AttrAsc, AttrHiLite AS WORD
LOCAL lTxt, cTbl, t AS STRING, pTxt AS STRING POINTER

   TP.ScreenRep(sRow, sCol, sTxt)                                 ' Update Text Image copy
   IF gMacroMode THEN EXIT SUB                                    ' If macro mode, exit

   '----- Setup possible translated string if a data text string
   pTxt = VARPTR(sTxt)                                            ' Point at passed string to start
   IF ISFALSE ISMISSING(sOffset) THEN                             ' Only do this for text data lines
      cTbl = ENV.Charset + " "                                    ' Get local valid character table
      i = VERIFY(sTxt, cTbl)                                      ' Get location of any unprintable chars
      IF i THEN                                                   ' Got some
         lTxt = sTxt: pTxt = VARPTR(lTxt)                         ' Copy and force use of the copied version
         cTbl = ENV.Charset + " "                                 ' Get local valid character table
         t = ENV.InvChar
         DO WHILE i                                               ' Make them all the chosen
            MID$(lTxt, i, 1) = ENV.InvChar                        ' Invalid character substitute
            i = VERIFY(lTxt, cTbl)                                ' Get location of any further unprintable chars
         LOOP                                                     '
      END IF                                                      '
   END IF                                                         '

   '----- Setup for line scans
   oPtr = 1                                                       '
   IF ISMISSING(sOffset) THEN                                     ' If no Offset or Pad
      cPtr = 1                                                    ' Start at left end of sTxt
      tLen = LEN(sTxt)                                            ' Do length of sTxt
   ELSE                                                           '
      cPtr = sOffset + 1                                          ' Start at Offset location
      tLen = sPad                                                 ' and do sPad # characters
   END IF                                                         '

   GRAPHIC ATTACH TP.PgHandle, TP.WindowID                        ' Using the live window
   GRAPHIC SET POS ((sCol - 1) * gFontWidth + %GLM, (sRow - 1) * gFontHeight + 0) ' Set position for Print
   DO WHILE cPtr <= LEN(@pTxt)                                    ' While still stuff in @pTxt
      GOSUB SetlclScheme                                          ' Setup lclScheme then
      GRAPHIC COLOR ENV.GetClr(lclScheme, %SCFG), lclBG           ' Set the FG/BG colors
      IF lclUL THEN GRAPHIC SET FONT hScrFontUnd                  ' Switch font if underlined

      i = VERIFY(cPtr + 1, sAttr, CHR$$(AttrAsc))                 ' Find next different Attr character
      j = IIF(i, i - cPtr, MIN(LEN(@pTxt) - cPtr + 1, tLen - oPtr))' Size of chunk
      IF i = 0 THEN                                               ' Remainder of attributes are the same
         IF lclUC THEN                                            '
            GRAPHIC PRINT UUCase$(MID$(@pTxt, cPtr));             ' Print rest of text uppercased
         ELSE                                                     '
            GRAPHIC PRINT MID$(@pTxt, cPtr);                      ' Print rest of text
         END IF                                                   '
         cPtr = LEN(@pTxt) + 1: oPtr += j                         ' Force loop exit
      ELSE                                                        ' There's another Scheme value coming up
         IF lclUC THEN                                            ' Handle UC if needed
            GRAPHIC PRINT UUCase$(MID$(@pTxt, cPtr, i - cPtr));   ' Print this segment uppercased
         ELSE                                                     '
            GRAPHIC PRINT MID$(@pTxt, cPtr, i - cPtr);            ' Print this segment
         END IF                                                   '
         cPtr += j: oPtr += j                                     ' Adjust pointers
      END IF                                                      '
      IF lclUL THEN GRAPHIC SET FONT hScrFont                     ' Put font back to normal
   LOOP                                                           '
   IF oPtr < tLen THEN                                            ' Pad needed
      AttrAsc = %SCTxtLo                                          ' Get colors back to normal text
      GOSUB SetlclScheme2                                         '
      GRAPHIC COLOR ENV.GetClr(lclScheme, %SCFG), lclBG           ' Set the FG/BG colors
      GRAPHIC PRINT SPACE$(tLen - oPtr + 1);                          '
   END IF                                                         '
   EXIT SUB                                                       '

   SetlclScheme:
      Attrasc = ASC(MID$(sAttr, cPtr, 1))                         ' Get Attribute byte
   SetlclScheme2:                                                 ' Alternate entry point
      AttrHiLite = (AttrAsc AND %AttrHiLite)                      ' Isolate the hi-lite color
      SHIFT RIGHT AttrHiLite, 8                                   '
      lclUC = IIF((Attrasc AND %AttrUC) <> 0, %True, %False)      ' Setup the UC flag
      lclUL = IIF((Attrasc AND %AttrUL) <> 0, %True, %False)      ' Setup the UL flag
      lclScheme = AttrAsc AND %AttrScheme                         ' Get scheme number
      IF AttrHiLite THEN lclScheme = AttrHiLite + 31              ' A color hilight? Adjust to a scheme number
      IF (AttrAsc AND %AttrInv) <> 0 THEN                         ' Invert request?
         cCustFG  = ENV.GetClr(lclScheme, %SCBG1)                 ' Setup custom request
         cCustBG1 = ENV.GetClr(lclScheme, %SCFG)                  '
         cCustBG2 = ENV.GetClr(lclScheme, %SCFG)                  '
         lclScheme = %SCCust                                      '
      END IF                                                      '
      lclBG = IIF(cBandBG, ENV.GetClr(lclScheme, %SCBG2), ENV.GetClr(lclScheme, %SCBG1))  ' Chose Scheme's BG color
      RETURN                                                      '

END SUB

SUB      sPrtPrint(sMode AS LONG, sTxt AS STRING, sAttr AS WSTRING, Number AS LONG)
'---------- Print a string on the printer
REGISTER i AS LONG
REGISTER j AS LONG
STATIC LineCtr, PageCtr, lgth, lgth2, posX, posY AS LONG
STATIC utab AS utab_t, utabdone AS LONG
LOCAL k, lclScheme, lclBG, cPtr, oPtr, tLen, DoingClose AS LONG, lAttr, lTxt AS WSTRING, xTxt, hText, subst AS STRING
LOCAL AttrAsc AS WORD
   '----- Setup Unicode table if needed
   IF ISFALSE utabdone THEN                                       ' Not done yet?
      slPrintUnicodeInit (utab)                                   ' initialize Unicode table
      utabdone = %True                                            ' remember we did it
   END IF                                                         '
   lAttr = sAttr: GOSUB UniCodeTxt                                ' Get working copies

   '----- Split by Mode
   SELECT CASE AS LONG sMode                                      ' Why called?
      CASE %PRTReset                                              ' Reset
         LineCtr = 0: PageCtr = 0                                 ' Line and Page counters

      CASE %PRTLine                                               ' Print some text
         GOSUB DoALine                                            ' Go do it

      CASE %PRTNewLine                                            ' Move to a new line
         XPRINT PRINT                                             '
         IF LineCtr > 0 THEN INCR LineCtr                         '

      CASE %PRTNewPage                                            ' Move to a new Page
         DO WHILE LineCtr <> 0                                    ' Loop to fill page
            lTxt = " ": lAttr = $$TxtLo                           ' Dummy values                                           '
            INCR Linectr                                          '
            GOSUB DoALine                                         '
         LOOP                                                     '

      CASE %PRTFlushClose                                         ' Shut things down
         DoingClose = %True                                       ' To avoid extra blank page at end
         DO WHILE LineCtr <> 0                                    ' Loop to fill page
            lTxt = " ": lAttr = $$TxtLo                           ' Dummy values
            INCR LineCtr                                          '
            GOSUB DoALine                                         '
            XPRINT PRINT                                          '
         LOOP                                                     '
         XPRINT CLOSE                                             ' End the document
         gPrinterOpen = %False                                    '
         FONT END gPFontHndl                                      ' Delete the font we created
         LineCtr = 0: PageCtr = 0                                 ' Reset Counters to 0
   END SELECT                                                     '
   EXIT SUB                                                       ' We're all done


   DoALine:
      IF LineCtr = 0 THEN                                         ' Page heading time?
         IF PageCtr <> 0 THEN XPRINT FORMFEED                     ' Yes, start a new page (other than 1st page)
         GOSUB Bandit                                             ' Add bands
         FOR i = 1 TO gPTFill: XPRINT " ": INCR LineCtr: NEXT i   ' Add top fill lines
         INCR PageCtr                               '             ' Bump page count
         gPageNumber = FORMAT$(PageCtr, "###")                    ' Make available for ~# substitution
         IF ENV.PrtHeader AND ISFALSE gPrtRaw THEN                ' Only if we're doing headers
            XPRINT COLOR ENV.GetClr(%SCTxtHi, %SCFG), -1          ' Print headings in Hi-Intensity
            hText = SPACE$(gPCpl)                                 ' Make Heading 1
            subst = ENV.PrtHeaderLeft                             ' Get format string
            TP.MacSubst(subst)                                    ' Do substitution
            LSET ABS hText = subst                                ' Insert it
            subst = ENV.PrtHeaderRight                            ' Get format string
            TP.MacSubst(subst)                                    ' Do substitution
            RSET ABS hText = subst                                ' Insert it
            subst = ENV.PrtHeaderCenter                           ' Get format string
            TP.MacSubst(subst)                                    ' Do substitution
            lgth = LEN(subst)                                     ' Get length of center part
            lgth2 = (gPCpl - lgth) / 2
            hText = LEFT$(hText, lgth2) + subst + RIGHT$(hText, gPCpl - lgth - lgth2)
            xTxt = SPACE$(gPLFill) + hText                        ' Add left fill
            XPRINT PRINT xTxt                                     ' Print it
            INCR LineCtr                                          '
            xTxt = SPACE$(gPLFill) + REPEAT$(gPCpl, "—")          ' Make Heading 2
            XPRINT PRINT xTxt                                     ' Print it
            INCR LineCtr                                          '
         END IF                                                   '
      END IF                                                      '

      '----- Process a text string
      cPtr = 1                                                    ' Start at left end of lTxt
      tLen = LEN(lTxt)                                            ' Do length of lTxt
      IF sAttr <> " " THEN XPRINT PRINT SPACE$(gPLFill);          ' Insert left fill

      DO WHILE cPtr <= LEN(lTxt)                                  ' While still stuff in lTxt
         Attrasc = ASC(MID$(sAttr, cPtr, 1))                      ' Get Attribute byte
         lclScheme = AttrAsc AND %AttrScheme                      ' Get scheme number
         IF ISFALSE ENV.PRTPColor OR ISTRUE gPrtRaw THEN lclScheme = %SCTxtHi       ' Force TxtHi of not color printing
         XPRINT COLOR ENV.GetClr(lclScheme, %SCFG), -1            ' Set the FG, BG = Transparent

         i = VERIFY(cPtr + 1, sAttr, CHR$$(AttrAsc))              ' Find next different Attr character
         j = IIF(i, i - cPtr, MIN(LEN(sTxt) - cPtr + 1, tLen - oPtr))' Size of chunk
         IF j + XPRINT(POS.X) > gPCpl THEN                        ' Will this go over a line?
            k = gPCpl - XPRINT(POS.X)                             ' Calc how much will fit
            xTxt = MID$(lTxt, cPtr, k)                            ' Print whatever that is
            XPRINT PRINT xTxt                                     '
            INCR LineCtr                                          '
            cPtr += k: oPtr += k                                  ' Adjust pointers
            IF sAttr <> " " THEN XPRINT PRINT SPACE$(gPLFill);    ' Insert left fill
            IF Number THEN XPRINT PRINT SPACE$(ENV.LinNoSize + 1); '
         ELSE                                                     '
            IF i = 0 THEN                                         ' Remainder of attributes are the same
               xTxt = MID$(lTxt, cPtr)                            ' Print rest of text
               XPRINT PRINT xTxt;                                 '
               cPtr = LEN(lTxt) + 1: oPtr += j                    ' Force loop exit
            ELSE                                                  ' There's another Scheme value coming up
               xTxt = MID$(lTxt, cPtr, i - cPtr)                  ' Print this segment
               XPRINT PRINT xTxt;                                 '
               cPtr += j: oPtr += j                               ' Adjust pointers
            END IF                                                '
         END IF                                                   '
      LOOP                                                        '

      IF LineCtr >= IIF(ENV.PrtFooter AND ISFALSE gPrtRaw, gPLpp - 2, gPLpp) THEN ' Time for footer?
         IF ENV.PrtFooter AND ISFALSE gPrtRaw THEN                ' Doing footers?
            XPRINT PRINT                                          ' End the last line
            XPRINT COLOR ENV.GetClr(%SCTxtHi, %SCFG), -1          ' Print headings in Hi-Intensity
            xTxt = SPACE$(gPLFill) + REPEAT$(gPCpl, "—")          ' Make Footer 1
            XPRINT PRINT xTxt                                     ' Print it

            hText = SPACE$(gPCpl)                                 ' Make Footer 2
            subst = ENV.PrtFooterLeft                             ' Get format string
            TP.MacSubst(subst)                                    ' Do substitution
            LSET ABS hText = subst                                ' Insert it
            subst = ENV.PrtFooterRight                            ' Get format string
            TP.MacSubst(subst)                                    ' Do substitution
            RSET ABS hText = subst                                ' Insert it
            subst = ENV.PrtFooterCenter                           ' Get format string
            TP.MacSubst(subst)                                    ' Do substitution
            lgth = LEN(subst)                                     ' Get length of center part
            lgth2 = (gPCpl - lgth) / 2                            '
            hText = LEFT$(hText, lgth2) + subst + RIGHT$(hText, gPCpl - lgth - lgth2)
            xTxt = SPACE$(gPLFill) + hText                        ' Left fill it
            XPRINT PRINT xTxt                                     ' Print it
         END IF                                                   '
         LineCtr = 0                                              ' Start next page
      END IF                                                      '
      RETURN                                                      ' Back now

   Bandit:

      IF ISFALSE ENV.PrtBanding AND ISFALSE ENV.PrtBandLines THEN RETURN
      XPRINT GET POS TO posX, posY                                ' Save current POS
      XPRINT SCALE PIXELS                                         ' Switch to Pixel mode

      FOR i = 1 + gPTFill + IIF(ENV.PrtHeader AND ISFALSE gPrtRaw, 2, 0) TO gPLpp - ABS((2 * ENV.PrtFooter)) _
          STEP IIF(ENV.PrtBanding, 6, 3)                          ' For vertical # lines
         IF ENV.PrtBanding THEN                                   ' If banding
            XPRINT BOX (1 + (gPLFill * gPCharWidth), (i * gPCharHeight) - gPCharHeight)  - _
                       (gPPageWidth - (gPRFill * gPCharWidth), ((i+3) * gPCharHeight) - gPCharHeight), _
                       0, ENV.PrtBandColor, ENV.PrtBandColor, 0   '
         ELSEIF ENV.PrtBandLines THEN                             ' Else, is it the Line version?
            XPRINT LINE (1 + (gPLFill * gPCharWidth), (1 + (i+3) * gPCharHeight) - gPCharHeight)  - _
                       (gPPageWidth - (gPRFill * gPCharWidth), (1 + (i+3) * gPCharHeight) - gPCharHeight), ENV.PrtBandColor
         END IF                                                   '
      NEXT i                                                      '
      XPRINT SCALE (0, 0) - (gPCpl, gPLpp)                        ' Set page scale back to characters
      XPRINT SET POS (posX, posY)                                 ' Restore current POS
      RETURN

   UniCodeTxt:
      lTxt = sTxt                                                 ' Do default Unicode translation
      IF utab.valid <> %false AND LEN (sTxt) > 0 THEN             ' If we have a table and some data
         lTxt = ""                                                ' Init string
         FOR k = 1 TO LEN(sTxt)                                   ' Translate the string
            lTxt += utab.uchar (ASC(MID$ (sTxt, k, 1)))           ' build line w/private Unicode translation
         NEXT k                                                   '
      END IF                                                      '
   RETURN                                                         '

END SUB

SUB      sPrtHelp(iTxt AS STRING, iRow AS LONG)
'---------- Print a Help string with selective underlining
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL k AS LONG
   IF gMacroMode THEN EXIT SUB                                    ' If macro mode, exit
   i = 1: j = 1                                                   ' Init for loop
   DO WHILE i <= LEN(iTxt)                                        '
      k = INSTR(i, iTxt, ANY $UpperSpec)                          ' Look for any > 128 chars
      IF k = 0 THEN sPrint (MID$(iTxt, i), $$PFK, iRow, j): EXIT SUB ' No more, remainder is normal
      IF k = i THEN                                               ' Found at start column?
         sPrint (CHR$(ASC(MID$(iTxt, i, 1)) - 128), $$PFKUL, iRow, j)
         INCR i: INCR j                                           ' Step over
      ELSE                                                        '
         sPrint (MID$(iTxt, i, k - i), $$PFK, iRow, j)            '
         j += (k - i): i = k                                      ' Adjust continue values
      END IF                                                      '
   LOOP                                                           '
END SUB

SUB sProcess_CodePage_Data_Line(TX AS CODEPAGE_TX_T, BUF AS STRING, CP_LineNo AS LONG)

'/-----------------------------------------------------------------------------/
'/  Process_Codepage_Data_Line                                                 /
'/                                                                             /
'/  Process tran table data entries like that are prefixed like "3_"           /
'/  the prefix defines one of it sectors where the data is stored              /
'/  there must be exactly 16 entries, and each can/must be used only once      /
'/  here is a sample AE line referenced by the code below.                     /
'/                                                                             /
'/  3_ F0 F1 F2 F3 F4 F5 F6 F7 F8 F9 7A 5E 4C 7E 6E 6F 3_                      /
'/  1234                                                                       /
'/  After the 16th entry on the line, the remainder is treated as comments     /
'/-----------------------------------------------------------------------------/
LOCAL C, NUM, SECTOR_CODE  AS STRING
LOCAL SECTOR_NUM, SECTOR_NDX, CHAR_NDX, CHAR_VALUE AS LONG
REGISTER I AS LONG

    IF  MID$(BUF, 2, 2) <> "_ " THEN
        TX.TX_Errors += 1                              '/ A parse error occurred
        TX.TX_Reason = "LINE " + TRIM$(CP_LineNo) +                            _
            ": DATA FORMAT ERROR AT: " + BUF
        EXIT SUB
    END IF

    SECTOR_CODE = UUCASE(LEFT$(BUF, 1))
    IF  VERIFY (SECTOR_CODE, "0123456789ABCDEF") <> 0 THEN
        TX.TX_Errors += 1                              '/ A parse error occurred
        TX.TX_Reason = "1: " + SECTOR_CODE + "LINE " + TRIM$(CP_LineNo)        _
            + ": DATA FORMAT ERROR AT: " + BUF
        EXIT SUB
    END IF

    '/ Create sector number and index.   3_ becomes &H30 = 48
    '/ we use &H0 in the VAL() to ensure an unsigned conversion happens
    SECTOR_NUM = VAL ("&H0" & SECTOR_CODE)                           '/ 3_ --> 3
    SECTOR_NDX = SECTOR_NUM * 16                         '/ 3 * 16 --> &H30 = 48
    CHAR_NDX = 0
    I = 3
    DO ' LOOP                                   '/ Start with space after prefix
        IF  I > LEN(BUF) THEN EXIT DO                   '/ Reached end of buffer
        C = MID$(BUF, I, 1)
        IF  C = " " THEN                                   '/ Inter-number space
            I += 1                                            '/ Skip over space
            ITERATE DO
        END IF

        IF  C = "*" THEN EXIT DO                          '/ End of line comment

        '/  3_ F0 F1 F2 F3 <== CURR BUF
        '/  ......0123      <-- Value of I + ?
        C = MID$(BUF, I+2, 1)
        IF (C <> " ") AND (C <> "*") THEN                       '/ Bad delimiter
            TX.TX_Errors += 1                          '/ A parse error occurred
            TX.TX_Reason = "LINE " + TRIM$(CP_LineNo)                          _
                + ": DELIMITER ERROR IN: " + BUF
            EXIT SUB
        END IF

        NUM = UUCASE(MID$(BUF, I, 2))             '/ Grab 2 digits like F1 above
        I += 2                                           '/ Consume the 2 digits

        IF  VERIFY (NUM, "0123456789ABCDEF") <> 0 THEN       '/ Invalid hex code
            TX.TX_Errors += 1                          '/ A parse error occurred
            TX.TX_Reason = "2: " + NUM + "LINE " + TRIM$(CP_LineNo)            _
                + ": DATA FORMAT ERROR AT: " + BUF
            EXIT DO
        END IF

        CHAR_VALUE = VAL ("&H0" & NUM)                  '/ Like F1 above --> 241
        TX.TX_Table (SECTOR_NDX + CHAR_NDX) = CHAR_VALUE
        TX.TX_Entry (SECTOR_NUM) += 1          '/ Count the number of 3_ entries
        TX.TX_Values += 1                             '/ Number of values stored
        TX.TX_Defined = 1                          '/ At least one entry defined
        CHAR_NDX += 1
        IF  CHAR_NDX >= 16 THEN EXIT DO
    LOOP

    IF  CHAR_NDX <> 16 THEN                         '/ Didn't get 16 good values
        TX.TX_Errors += 1                              '/ A parse error occurred
        TX.TX_Reason = "LINE " + TRIM$(CP_LineNo)                              _
            + ": DATA FORMAT ERROR AT: " + BUF
        EXIT SUB
    END IF
END SUB ' sProcess_CodePage_Data_Line

SUB sProcess_CodePage_Source_Line(CP AS CodePage_CP_T, BUF AS STRING)
'/-----------------------------------------------------------------------------/
'/  Process_CodePage_Source_Line                                               /
'/                                                                             /
'/  This routine examines the leading two character action code, and routes    /
'/  the buffer to the appropriate handler.                                     /
'/  Action codes are: TT, TA, TE, AE AND EA                                    /
'/  The prefixes on data lines like "4_" are alco considered action codes      /
'/  Blank lines, full-line comments, and the // EOF mark are handled elsewhere /
'/-----------------------------------------------------------------------------/
STATIC AE_EA_Mode       AS LONG
LOCAL ACTION, C         AS STRING
LOCAL I, TT_INDEX       AS LONG
    ACTION = UUCASE(LEFT$((BUF + "     "), 3))
    TT_INDEX = 0
    IF     ACTION = "TA " THEN
        TT_INDEX =%TA_Index
    ELSEIF ACTION = "TE " THEN
        TT_INDEX =%TE_Index
    ELSEIF ACTION = "AE " THEN
        AE_EA_Mode = %AE_Mode                      '/ Save static state of AE/EA
        EXIT SUB
    ELSEIF ACTION = "EA " THEN
        AE_EA_Mode = %EA_Mode                      '/ Save static state of AE/EA
        EXIT SUB
    END IF

    BUF = BUF + " "                   '/ Processing routines need trailing space

    IF  TT_INDEX <> 0 THEN                       '/ BUF has TA or TE action code
        AE_EA_Mode = 0
        sProcess_CodePage_TX_Line (CP.TX (TT_INDEX), BUF, CP.CP_LineNo)
        EXIT SUB
    END IF

    IF  ACTION = "TT " THEN                            '/ BUF has TT action code
        AE_EA_Mode = 0
        sProcess_CodePage_TT_Line (CP.TT, BUF, CP.CP_LineNo)
        EXIT SUB
    END IF

    '/-------------------------------------------------------------------------/
    '/  If current line is an AE/EA data entry, store it                       /
    '/  lines are only recognized in AE/EA mode                                /
    '/  unless we got a prior AE/EA action code, the lines are out of order    /
    '/-------------------------------------------------------------------------/
    IF  AE_EA_Mode <> 0 THEN             '/ This mode state is a static varioable
        IF  MID$(BUF, 2, 1) = "_"  THEN               '/ Entry looks like A "0_"

            '/ We let Process_CodePage_Data_Line() validate the line
            sProcess_CodePage_Data_Line (CP.TX (AE_EA_Mode), BUF, CP.CP_LineNo)
            EXIT SUB
        END IF
    END IF

    '/ At this point, the line has an invalid action code
    AE_EA_Mode = 0
    CP.CP_Errors += 1
    CP.CP_Reason = "LINE " + TRIM$(CP.CP_LineNo) + ": "                        _
        + "ACTION CODE UNDEFINED: " + ACTION
END SUB ' sProcess_CodePage_Source_Line

SUB sProcess_CodePage_TT_Line(TT AS CODEPAGE_TT_T, BUF AS STRING, CP_LineNo AS LONG)
'/-----------------------------------------------------------------------------/
'/  Process_CodePage_TT_Line                                                   /
'/                                                                             /
'/  Process tran table TT attributes.  These are general descriptive           /
'/  attributes to help identify the source and nature of the codepage          /
'/-----------------------------------------------------------------------------/
LOCAL KV                  AS KEYWORD_VALUE_DATA_T
LOCAL FailSafe            AS LONG
    KV.KV_DATA = MID$(BUF, 3)                                '/ Skip the TT code
    FailSafe = LEN(BUF)                  '/ A rule-of-thumb upper limit to parse

    '/ There can't be more tokens in a buffer than the number of characters.
    '/ The FailSafe is to protect the logic from locking up if the parse fails
    DO  UNTIL FailSafe <= 0
        FailSafe -= 1
        sParse_Keyword_Value_Data (KV)
        IF  KV.KV_STATUS = 0 THEN EXIT DO    '/ Buffer consumed, parse completed
        IF  KV.KV_STATUS = 1 THEN                '/ One keyword value pair found
            SELECT CASE KV.KV_KEYWORD
                CASE "AUTHOR"   : TT.TT_Author  = KV.KV_VALUE
                CASE "GENDATE"  : TT.TT_GenDate = KV.KV_VALUE
                CASE "MODE"     : TT.TT_Mode    = KV.KV_VALUE
                CASE "NAME"     : TT.TT_Name    = KV.KV_VALUE
                CASE "TITLE"    : TT.TT_Title   = KV.KV_VALUE

                '/ If we get an unknown keyword, just store keyword/value as is
                '/ maybe someone is trying to convey a 'note' we didn't
                '/ account for  so we don't treat this as a fatal error.
                CASE ELSE:     TT.TT_Other = KV.KV_KEYWORD + "=" + KV.KV_VALUE
            END SELECT

        ELSE
            '/ At this point, some kind of parse error occurred
            '/ KV.KV_Data will have what is left of the parse buffer, which
            '/ should be at the beginning of where the problem is
            TT.TT_Errors += 1                          '/ A parse error occurred
            TT.TT_Reason = "LINE " + TRIM$(CP_LineNo) + ": " + KV.KV_Reason
            EXIT SUB
        END IF
    LOOP

END SUB ' sProcess_CodePage_TT_Line

SUB sProcess_CodePage_TX_Line(TX AS CODEPAGE_TX_T, BUF AS STRING, CP_LineNo AS LONG)
'/-----------------------------------------------------------------------------/
'/  Process_CodePage_TX_Line                                                   /
'/                                                                             /
'/  This routine parses and stores all the keyword/value pairs associated with /
'/  the TA/TE lines.  Their format is the same, so we process both kinds in    /
'/  the same routine.  We get passed the appropriate TA/TE structure           /
'/-----------------------------------------------------------------------------/
    '/ Process tran table TA/TE attributes
LOCAL KV                  AS KEYWORD_VALUE_DATA_T
LOCAL FailSafe            AS LONG

    KV.KV_DATA = MID$(BUF, 3)                             '/ Skip the TA/TE code
    FailSafe = LEN(BUF)                  '/ A rule-of-thumb upper limit to parse

    '/ There can't be more tokens in a buffer that the number of characters
    '/ THE FailSafe is to protect the logic from locking up if the parse fails
    DO  UNTIL FailSafe <= 0
        FailSafe -= 1
        sParse_Keyword_Value_Data (KV)
        IF  KV.KV_STATUS = 0 THEN EXIT DO    '/ Buffer consumed, parse completed
        IF  KV.KV_STATUS = 1 THEN                '/ One keyword value pair found
            SELECT CASE KV.KV_KEYWORD
                CASE "CCSID"    : TX.TX_CCSID   = KV.KV_VALUE
                CASE "CGCSGID"  : TX.TX_CGCSGID = KV.KV_VALUE
                CASE "CODESET"  : TX.TX_CodeSet = KV.KV_VALUE
                CASE "CPGID"    : TX.TX_CPGID   = KV.KV_VALUE
                CASE "EURO"     : TX.TX_Euro    = KV.KV_VALUE
                CASE "NUMBER"   : TX.TX_Number  = KV.KV_VALUE
                CASE "ORIGIN"   : TX.TX_Origin  = KV.KV_VALUE
                CASE "RELATED"  : TX.TX_Related = KV.KV_VALUE
                CASE "SCHEME"   : TX.TX_SCHEME  = KV.KV_VALUE
                CASE "SIZE"     : TX.TX_Size    = KV.KV_VALUE
                CASE "SUB"      : TX.TX_Sub     = KV.KV_VALUE
                CASE "TYPE"     : TX.TX_Type    = KV.KV_VALUE
                CASE "UCM"      : TX.TX_UCM     = KV.KV_VALUE
                CASE "UCMDATE"  : TX.TX_UCMDate = KV.KV_VALUE
                CASE "VERSION"  : TX.TX_Version = KV.KV_VALUE
                '/ If we get an unknown keyword, just store keyword/value as is
                '/ Maybe someone is trying to convey a 'note' we didn't
                '/ account for, so we don't treat this as a fatal error
                CASE ELSE:     TX.TX_Other = KV.KV_KEYWORD + "=" + KV.KV_VALUE
            END SELECT

        ELSE
            '/ At this point, some kind of parse error occurred
            '/ KV.KV_Data will have what is left of the parse buffer, which
            '/ should be at the beginning of where the problem is
            TX.TX_Errors += 1                          '/ A parse error occurred
            TX.TX_Reason = "LINE " + TRIM$(CP_LineNo) + ": " + KV.KV_Reason
            EXIT SUB
        END IF
    LOOP
END SUB ' sProcess_CodePage_TX_Line

FUNCTION sProfState(verb AS STRING, OPTIONAL pName AS STRING) AS LONG
'---------- Handle Profile STATE table
STATIC ProfTable() AS STRING
DIM ProfTable(1 TO 300) AS STATIC STRING
STATIC ProfNumber AS LONG
LOCAL i, j AS LONG, wProf, usename AS STRING
LOCAL RetVal AS LONG, zResult AS ASCIIZ * 2000
LOCAL zSection AS ASCIIZ * %MAX_PATH, zKey AS ASCIIZ * %MAX_PATH
LOCAL zDefault AS ASCIIZ * %MAX_PATH, ININamez AS ASCIIZ * %MAX_PATH

   IF ISFALSE ISMISSING(pName) THEN                               ' Only if optional parameter
      wProf = UCASE$(pName)                                       ' Create our working name
      zSection = "File": zDefault = "": usename = ""              ' Always the "File" INI section, always "" default
   END IF                                                         '

   SELECT CASE AS CONST$ verb                                     ' Why were we called?

      CASE "RESET": ProfNumber = 0: FUNCTION = 0                  ' RESET

      '----- Return STATE setting
      CASE "FETCH"                                                ' FETCH
         '----- See if using the DEFAULT PROF list
         IF INSTR(UUCASE("," + ENV.DefaultShr), "," + wProf + ",") THEN  ' Is it in the Default Prof list?
            wProf = "DEFAULT"                                     ' Swap in the DEFAULT
         END IF                                                   '

         '----- See if in the table
         IF ProfNumber > 0 THEN                                   ' Got some cached?
            ARRAY SCAN ProfTable() FOR ProfNumber, FROM 2 TO 100, = wProf, TO i
            IF i THEN FUNCTION = VAL(LEFT$(ProfTable(i), 1)): EXIT FUNCTION
         END IF                                                   '

         '----- Not cached, see if the PROFILE exiats
         IF ISFALSE ISFILE(ENV.PROFPath + wProf + ".INI") THEN    ' If it doesn't exist, create an OFF entry
            INCR ProfNumber                                       ' Bump number saved
            ProfTable(ProfNumber) = "0" + wProf                   ' Save an OFF entry
            FUNCTION = %False: EXIT FUNCTION                      ' Return False
         END IF                                                   '

         ' Profile exists, have a look at it
         ININamez = ENV.PROFPath + wProf + ".INI"                 ' Setup the INI filename

         '----- See if an active USING
         zKey  = "ProfUsing"                                      ' Get the USING value
         RetVal = GetPrivateProfileString(zSection, zKey, zDefault, zResult, SIZEOF(zResult), ININamez)
         usename = IIF$(RetVal, LEFT$(zResult, RetVal), "")       ' Get usename
         IF usename <> "" THEN                                    ' If USING active
            ININamez = ENV.PROFPath + usename + ".INI"            ' Redo the INI name
         END IF                                                   '

         '----- Finally, get the STATE setting
         zKey  = "StateFlag"                                      ' Get the STATE value
         RetVal = GetPrivateProfileString(zSection, zKey, zDefault, zResult, SIZEOF(zResult), ININamez)
         j = VAL(IIF$(RetVal, LEFT$(zResult, RetVal), ""))        ' Get STATE
         INCR ProfNumber                                          ' Bump number saved
         ProfTable(ProfNumber) = FORMAT$(j) + wProf               ' Save the entry
         IF usename <> "" THEN                                    ' Add using entry if being used
            INCR ProfNumber                                       ' Bump number saved
            ProfTable(ProfNumber) = FORMAT$(j) + usename          ' Save the entry
         END IF
         FUNCTION = j                                             ' Pass back True/False

      CASE ELSE                                                   ' ??
         FUNCTION = %False                                        '
   END SELECT                                                     '
END FUNCTION

SUB      sPrtScreen
'---------- Handle the Print Screen Requests
LOCAL f, i, j AS LONG, fn, CBD, d, MSG AS STRING
   '----- Split off based on which type called for
   MEntry
   SELECT CASE AS CONST$ gKeyChr                                  ' Which flavour of PRT did we get
      CASE "PRTSCRNCLIPBOARD", "PRTTEXTCLIPBOARD"                 ' Plain Print or Data Print to Clipboard
         GOSUB PrtClipBoard                                       ' Go do it
      CASE "PRTSCRNPRINTER"                                       ' Print to default printer
         GOSUB PrtPrinter                                         ' Go do it
      CASE "PRTSCRNLOG"                                           ' Print to SPFLite.LOG (append)
         GOSUB PrtLog                                             '
   END SELECT                                                     '
   MExitSub

'----- Print to the Clipboard
PRTClipBoard:
   '----- Build the Clipboard string
   CBD = ""                                                       ' Start as ""
   IF gKeyChr = "PRTSCRNCLIPBOARD" THEN                           ' Do full screen dump
      FOR i = 1 TO gwScrHeight + ENV.PFKShow                      ' Loop through the screen image
         CBD += TP.ScreenGet(i) + $CR + $LF                       ' Add each line with CR/LF
      NEXT x                                                      '
   ELSE                                                           ' Do data lines only
      FOR i = 3 + TP.PrfCols TO gwScrHeight                       ' Loop through the screen data lines
         CBD += MID$(TP.ScreenGet(i), 8) + $CR + $LF              ' Add each line with CR/LF
      NEXT x                                                      '
   END IF                                                         '
   CBD += $NUL                                                    ' ASCIIZ terminate it

   j = sWinclip_set(CBD)                                          ' Send print data to the Clipboard

   IF ISTRUE j THEN                                               ' OK?
      IF gKeyChr = "PRTTEXTCLIPBOARD" THEN                        ' Issue appropriate message
         MSG = "Data lines sent to Clipboard"                     '
         GOSUB PrtErrMsg                                          ' Issue locally since we're not in KBAttn mode
      ELSE                                                        '
         MSG = "Print Screen Image sent to Clipboard"             '
         GOSUB PrtErrMsg                                          ' Issue locally since we're not in KBAttn mode
      END IF                                                      '
   ELSE                                                           '
      MSG = "Print Screen failed"                                 '
      GOSUB PrtErrMsg                                             ' Issue locally since we're not in KBAttn mode
   END IF                                                         '
   RETURN

'----- Print to the Printer
PRTPrinter:
   '-----
   IF ISFALSE sOpenPrinter("") THEN                               ' Get printer ready if not already
      MSG = "OPEN of Printer failed"                              ' Oops!
      GOSUB PrtErrMsg                                             ' Issue locally since we're not in KBAttn mode
   ELSE                                                           '
      gPrtRaw = %True                                             ' Say no headings etc.
      sPRTPrint(%PRTReset, " ", " ", %False)                      ' Tell sPRTPrint to reset
      FOR i = 1 TO gwScrHeight + ENV.PFKShow                      ' Loop through the screen image
         sPRTPrint(%PRTLine, TP.ScreenGet(i), $$TxtHi, %False)    ' Print each screen line
         sPRTPrint(%PRTNewLine, " ", " ", %False)                 ' New Line
      NEXT i                                                      '
      sPRTPrint(%PRTFlushClose, " ", ",", %False)                 ' Tell sPRTPrint to flush page
      gPrtRaw = %False                                            ' Turn off Raw mode print
      MSG = "Screen Image sent to Default Printer"                '
      GOSUB PrtErrMsg                                             ' Issue locally since we're not in KBAttn mode
   END IF                                                         '
   RETURN

'----- Print to the LOG file
PRTLog:
   '-----
   f = FREEFILE                                                   ' Get a free file number
   fn = ENV.INIPath + "SPFLiteScrPrt.LOG"                         ' Make full name of the log file
   OPEN fn FOR APPEND AS #f                                       ' Go open it
   FOR i = 1 TO gwScrHeight + ENV.PFKShow                         ' Loop through the screen image
      PRINT#f, TP.ScreenGet(i)                                    ' Print each screen line
   NEXT i                                                         '
   CLOSE#f                                                        ' Close the file
   MSG = "Screen Image sent to SPFLiteScrPrt.LOG"                 '
   GOSUB PrtErrMsg                                                ' Issue locally since we're not in KBAttn mode
   RETURN                                                         '

PRTErrMsg:
      d = STRING$(ENV.ScrWidth - LEN(MSG), "_")                   ' Build LH part of dash line
      sPrint (d, $$TxtLo, 2, 1)                                   ' Print LH part of line 2
      sPrint (MSG, $$Error, 2, ENV.ScrWidth - LEN(MSG) + 1)       ' Print rest of line
   RETURN
END SUB

FUNCTION sQBColor(BYVAL N AS LONG) AS LONG
'---------- Convert QB color value
LOCAL cc() AS LONG
DIM cc(16)
'                            0         1         2         3         4         5         6         7
   FUNCTION = CHOOSE&(n + 1, 0,        &h800000, &h008000, &h808000, &h0000C4, &h800080, &h004080, &hC4C4C4, _
                             &h808080, &hFF0000, &h00FF00, &hFFFF00, &h0000FF, &hFF00FF, &h00FFFF, &hFFFFFF ELSE &hFFFFFF)
'                            8         9         10        11        12        13        14        15
END FUNCTION

SUB   sQDir2Array(mask AS STRING, dList() AS STRING, dNum AS LONG)
'---------- Return a DIR filename list
LOCAL tDir AS STRING
   dNum = 0                                                       ' Clear count
   tDir = DIR$(mask)                                              ' Get first entry
   WHILE LEN(tDir)                                                ' While we got something
      INCR dNum                                                   ' Bump count
      IF dNum > UBOUND(dList()) THEN _                            ' If needed, expand array
         REDIM PRESERVE dList(UBOUND(dList()) + 500) AS STRING    ' By 500
      dList(dNum) = tDir                                          ' Save answer
      tDir = DIR$(NEXT)                                           ' Try for another
   WEND                                                            '
END SUB

FUNCTION sReadClipboard(CBData AS STRING, dlm AS STRING, EraseIt AS LONG) AS LONG
'---------- Read Clipboard for the caller
LOCAL lclPrimOper, CBError AS STRING
LOCAL cbIO AS iIO                                                 ' For our I/O stuff

   MEntry
   LET cbIO = CLASS "cIO"                                         '
   lclPrimOper = gKeyPrimOper
   IF LEFT$(lclPrimOper, 4) = "$RAW" OR LEFT$(lclPrimOper, 4) = "$RAA" THEN _
      lclPrimOper = MID$(lclPrimOper, 5)
   dlm = $CRLF                                                    ' Use a std delimiter
   CBData = ""                                                    ' Start as null

   '----- See which Clipboard to read
   IF ISNULL(lclPrimOper) THEN                                    ' If no operand, use normal Clipboard
      sWinclip_get(CBData)                                        ' Read from Win Clipboard; data required


   ELSEIF lclPrimOper = "|InternalCB" THEN                        ' The internal CBD?
      CBData = gInternalCB                                        ' Yes, return it

   ELSE                                                           '
      cbIO.Setup("BE", "", "", ENV.CLIPPath + lclPrimOper + ".CLIP") ' Set filename
      IF cbIO.EXEC THEN                                           ' Go OPEN the file
         FUNCTION = %True: MExitFunc                              ' Oops?  Bail out
      END IF                                                      '
      GET$ # cbIO.FNum, LOF(cbIO.FNum), CBData                    ' Get whole file in one gulp
      cbIO.Close                                                  ' Close the FBO
      IF EraseIt THEN sRecycleBin(ENV.CLIPPath + lclPrimOper + ".CLIP", "D")  ' If ERASE then delete the CLIP file
   END IF                                                         '

   '----- Figure out delimiters; if none of these found, 'dlm' retains value set by caller

   IF INSTR(CBData, $CRLF) THEN                                   ' See what kind of delimiters
      dlm = $CRLF                                                 ' CRLF
   ELSEIF INSTR(CBData, $CR) THEN                                 '
      dlm = $CR                                                   ' CR
   ELSEIF INSTR(CBData, $LF) THEN                                 '
      dlm = $LF                                                   ' LF
   END IF                                                         '
   MExitFunc                                                          '

END FUNCTION                                                      '

FUNCTION sRead_CodePage_Source_File(CP AS CodePage_CP_T, SOURCE_ID AS STRING) AS LONG
'/-----------------------------------------------------------------------------/
'/  Read_CodePage_Source_File                                                  /
'/                                                                             /
'/  Locate and read a .SOURCE file; Source_ID is base name like "EBCDIC"       /
'/  Return 1 if read successful, else 0                                        /
'/-----------------------------------------------------------------------------/
LOCAL BUF, SOURCE_FILE_Name AS STRING
LOCAL SOURCE_FILE, RETCODE  AS LONG
   SOURCE_FILE_Name = ENV.INIPath + SOURCE_ID + ".SOURCE"
   IF ISFILE (SOURCE_FILE_Name) THEN
      sInit_CodePage (CP)
      SOURCE_FILE = FREEFILE
      OPEN SOURCE_FILE_Name FOR INPUT ACCESS READ AS # SOURCE_FILE
      DO WHILE ISFALSE EOF (SOURCE_FILE)
         CP.CP_LineNo += 1
         LINE INPUT # SOURCE_FILE, BUF
         BUF = TRIM$(BUF)
         IF BUF = "" THEN ITERATE DO                              '/ Line is blank
         IF LEFT$(BUF, 1) = "*" THEN ITERATE DO                   '/ Line is comment
         IF MID$(BUF, 2, 1) = "*" THEN ITERATE DO                 '/ Line is comment

         '/-----------------------------------------------------------------/
         '/  When the logical-EOF action code // is read, we stop reading   /
         '/  the file.  The rest of the XX line is ignored, as is any other /
         '/  data following the XX line, which can be used for comments.    /
         '/-----------------------------------------------------------------/
         IF LEFT$(BUF, 2) = "//" THEN                             '/ Logical EOF reached on file
            EXIT DO
         END IF
         sProcess_CodePage_Source_Line (CP, BUF)
      LOOP
      CLOSE # SOURCE_FILE
      RETCODE = sValidate_CodePage_Data (CP)                      '/ Return 1 if valid, else 0
      FUNCTION = RETCODE
      EXIT FUNCTION
   END IF
   FUNCTION = 0                                                   '/ Failure processing source file
END FUNCTION ' sRead_CodePage_Source_File

SUB      sRecentAdd(file AS STRING)
'---------- Add/Update an item in the RECENT list
   MEntry
   IF IsNE(RIGHT$(file, 6), ".FLIST") THEN                        ' If not a FILELIST itself
      sFileListAdd("Recent Files", file)                          ' Call common routine with added unique parameters
      sFileListAdd("Recent Paths", LEFT$(file, INSTR(-1, file, "\")))  ' Ditto for PATHS
   END IF                                                         '
   MExit
END SUB

FUNCTION sRecycleBin(FilNam AS STRING, which AS STRING) AS LONG
'---------- Send file to the Recycle bin
LOCAL shfo AS SHFILEOPSTRUCT                                      ' Predefined structure
LOCAL szSource AS ASCIIZ * %MAX_PATH
LOCAL dummy AS LONG
   MEntry
   '----- Setup the parameter list
   szSource = FilNam + CHR$(0, 0)                                 ' Convert to ASCIIZ
   shfo.wFunc = %FO_DELETE                                        ' Function delete file
   shfo.pFrom = VARPTR(szSource)                                  ' Pointer to file

   '----- Set ALLOWUNDO based on ENV.UseRecycle and specific request
   IF ENV.UseRecycle THEN                                         ' Set correct flags
      IF which = "D" THEN                                         ' Normal Delete?
         shfo.fFlags = %FOF_ALLOWUNDO OR %FOF_NOCONFIRMATION OR %FOF_NOERRORUI OR %FOF_NO_UI ' Enable undo / no confirm
      ELSE                                                        '
         shfo.fFlags = %FOF_NOCONFIRMATION OR %FOF_NOERRORUI OR %FOF_NO_UI  ' Must be "K"ill, or P"U"rge
      END IF
   ELSE                                                           '
      shfo.fFlags = %FOF_NOCONFIRMATION OR %FOF_NOERRORUI OR %FOF_NO_UI ' Enable no confirm
   END IF                                                         '

   '----- Tell system to do it and return result
   dummy = SHFileOperation(shfo)                                  ' Call funtion
   FUNCTION = shfo.fAnyOperationsAborted                          ' Return value, either 0 or non-zero
   MExit
END FUNCTION

FUNCTION sRegGet(BYVAL sSubKeys AS STRING, BYVAL sValueName AS STRING, BYVAL sDefault AS STRING) AS STRING
'---------- Get a key from the Registry
LOCAL lKey AS DWORD, zRegVal AS ASCIIZ * 1024, dwType AS DWORD, dwSize AS DWORD
   zRegVal = sDefault                                             '
   IF (RegOpenKeyEx(%HKEY_CURRENT_USER, TRIM$(sSubKeys, "\"), 0, %KEY_READ, lKey) = %ERROR_SUCCESS) THEN
      dwType = %REG_SZ                                            '
      dwSize = SIZEOF(zRegVal)                                    '
      RegQueryValueEx(lKey, BYCOPY sValueName, 0, dwType, zRegVal, dwSize)
      RegCloseKey lKey                                            '
   END IF                                                         '
   FUNCTION = zRegVal                                             '
END FUNCTION

FUNCTION sRegSet(BYVAL sSubKeys AS STRING, BYVAL sValueName AS STRING, BYVAL sData AS STRING) AS LONG
'---------- Set a key into the Registry
LOCAL lKey AS DWORD, zRegName AS ASCIIZ * 1024, zRegVal AS ASCIIZ * 1024, dwType AS DWORD, dwSize AS DWORD
   zRegVal = sData                                                '
   zRegName = sValueName                                          '
   IF RegCreateKeyEx(%HKEY_CURRENT_USER, TRIM$(sSubKeys, "\"), 0, "", 0, %KEY_WRITE, BYVAL %Null, _
                     lKey, BYVAL %Null) = %ERROR_SUCCESS THEN     '
      dwSize = SIZEOF(zRegVal)                                    '
      dwType = %REG_SZ                                            '
      IF RegSetValueEx(lKey, zRegName, 0, dwType, zRegVal, dwSize) = %ERROR_SUCCESS THEN FUNCTION = %True
      RegCloseKey lKey                                            '
   END IF                                                         '
END FUNCTION


SUB      sResize(sMax AS STRING)
'---------- Do a Window resize based on user drag
LOCAL x, y, ix, iy, w, h, mx, my AS LONG
STATIC LastX, LastY AS LONG
   MEntry
   IF ISFALSE gResizeActive THEN MExitSub

   OffTPMarkActive                                                ' Kill any Active Mark area

   '----- Get the new size of the Window
   DIALOG GET SIZE hWnd TO mx, my                                 '
   DIALOG GET CLIENT hWnd TO ix, iy                               '

   IF ABS(LastX - ix) < 4 AND ABS(LastY - iy) < 4 THEN MExitSub   ' Do nothing until we change a lot
   LastX = ix: LastY = iy                                         ' Save the last x,y we actually process

   x = ix - %GLM - %GRM - 2                                       ' Usable x (minus LM and RM pad
   y = iy - gTabHdrRC.nBottom - gSBHeight                         ' Usable y = Client size minus Tab Header size - SB height

   '----- Convert to Text width and height, see what's changed
   w = FIX(x \ gFontWidth): h = FIX(y \ gFontHeight)              '
   IF w < 30 OR h < 10 THEN                                       ' Stupid user?
      sDoBeep                                                     '
      w = MAX(30, w): h = MAX(10, h)                              '
   END IF                                                         '

   ENV.ScrWidth = w: ENV.ScrHeight = h                            '
   sIniSetString("Screen", "ScrWidth", FORMAT$(ENV.ScrWidth))     ' Save the new values
   sIniSetString("Screen", "ScrHeight", FORMAT$(ENV.ScrHeight))   '

   IF sMax = "M" THEN
      sResizeWindow(mx, my)                                       ' Go apply the new size values
   ELSE                                                           '
      sResizeWindow(0, 0)                                         ' Go apply the new size values
   END IF                                                         '
   MExit
END SUB

SUB      sResizeWindow(mx AS LONG, my AS LONG)
LOCAL h, i, x, y, nx, ny, PgNo AS LONG
LOCAL JustOnce AS LONG
   MEntry
   IF ISFALSE gResizeActive THEN MExitSub
   IF TabsNum > 0 THEN                                            ' Initialized already?
      PgNo = TP.PgNumber                                          ' Save what page we're on for later
      GRAPHIC ATTACH TP.PgHandle, TP.WindowID                     ' Using the live window

      '----- Set the new fonts and get their sizes
      TRY                                                         ' Just in case
         IF hScrFont THEN FONT END hScrFont                       ' Free any prior Font
      CATCH                                                       '
         EXIT TRY                                                 ' Ignore error
      END TRY                                                     '
      TRY                                                         ' Just in case
         IF hScrFontUnd THEN FONT END hScrFontUnd                 '
      CATCH                                                       '
         EXIT TRY                                                 ' Ignore error
      END TRY                                                     '
      FONT NEW ENV.FontName, ENV.FontPitch, ENV.FontStyle, 1, 1 TO hScrFont   ' Get the basic font
      FONT NEW ENV.FontName, ENV.FontPitch, ENV.FontStyle + 4, 1, 1 TO hScrFontUnd ' Get the underline version of the basic font
      GRAPHIC SET FONT hScrFont                                   ' Set the desired font

      GRAPHIC CELL SIZE TO gFontWidth, gFontHeight                ' Get size of a character (done in PBMain for Init)
   END IF

   '----- We now have fonts created and gotten their sizes
   gDataLen  = ENV.ScrWidth - gLNPadCol                           ' Calc derived values
   pCmdLen   = ENV.ScrWidth - 24                                  ' |Command > | and | Scroll > | and 4 char scroll field = 24
   gwScrHeight = ENV.ScrHeight - ENV.PFKShow                      ' Shrink data area by PFK Show area

   x = (ENV.ScrWidth * gFontWidth) + %GLM + %GRM: y = (ENV.ScrHeight * gFontHeight) ' Calc basic x,y (PLUS lm AND rm PADS)

   '----- Re-size things

   '----- Set Rect to our needed Graphic size and then set the Tab to that size
   gTabRC.nTop = 0:     gTabRC.nLeft = 0                          ' Init Tab Rect
   gTabRC.nRight = x: gTabRC.nBottom = y                          '
   TabCtrl_AdjustRect hTab, 1, gTabRC                             ' Set Tab display to suit graphic
   TabCtrl_GetItemRect hTab, 1, gTabHdrRC                         ' Get Tab title dimensions .nBottom = height
   CONTROL SET SIZE hWnd, %IDC_SPFLiteTAB, gTabRC.nRight - gTabRC.nLeft, gTabRC.nBottom - gTabRC.nTop + gTabHdrRC.nBottom
   CONTROL SET COLOR hWnd, %IDC_SPFLiteTAB, cStatFG, cStatBG1     ' Default color it

   '----- Now get the actual Tab size
   CONTROL GET SIZE hWnd, %IDC_SPFLiteTAB TO nx, ny               ' Now get the tab size
   DIALOG SET CLIENT hWnd, nx, ny + gSBHeight                     ' Resize the whole dialog, allowing for the headers

   '----- Re-size things
   IF mx OR my THEN                                               ' Passed MAX size?
      DIALOG SET SIZE hWnd, mx, my                                ' Set to the Max values
      DIALOG SET COLOR hWnd, cStatFG, cStatBG1                    ' Default color it
      gSBWidth = mx                                               ' Save as SB width
   ELSE
      DIALOG SET CLIENT hWnd, nx, ny                              ' Set to calculated values
      DIALOG SET COLOR hWnd, cStatFG, cStatBG1                    ' Default color it
      gSBWidth = nx                                               ' Save as SB width
   END IF                                                         '

   sSetupSB                                                       ' Go set up the Status Bar

   InitFMLayout                                                   ' Adjust FM area

   '----- Re-do the existing tabs
   IF TabsNum > 0 THEN                                            ' If at least one tab page
      FOR i = 1 TO TabsNum                                        ' Establish the tabs (again)
         TP = Tabs(i)                                             ' Get tab data addressable
         TP.CsrRow = 1: TP.CsrCol = 11                            ' Force cursor to command line
         TP.MarkedLine = 0: TP.SwapSLin = 0                       ' Clear marked lines for Edit and FM
         CONTROL SET SIZE TP.PgHandle, TP.WindowID, x, y          ' Resize the graphic
         GRAPHIC ATTACH TP.PgHandle, TP.WindowID                  ' Set as the default graphic area
         GRAPHIC SET FONT hScrFont                                ' Set the font
         GRAPHIC CLEAR cTxtLoBG1                                  ' Clear it
         TP.ScreenDim(ENV.ScrHeight + 1, ENV.ScrWidth)            ' Redim the Screen shadow copy
         TP.DispScreen                                            ' Re-display stuff
      NEXT i                                                      '
      TP = Tabs(PgNo)                                             ' Put back starting page number
      TAB SELECT hWnd, %IDC_SPFLiteTAB, TP.PgNumber               ' Select its tab
      TP.WindowTitle                                              ' Alter window title
      sDoCursor                                                   ' Activate cursor
   END IF                                                         '

   MExit
END SUB

THREAD FUNCTION sResultFileWatchThread(BYVAL rfile AS LONG) AS LONG   ' Monitor a directory, stop when file disappears
'---------- Watch a SUBMIT result file
THREADED hSearch, Posted, i, FNm AS LONG                          '
THREADED tfd AS DIRDATA                                           '
THREADED Fn, passfile, L1, L2, L3, t AS STRING                    '
THREADED sfiletime, sfilesize AS QUAD                             '
THREADED watchdir AS ASCIIZ * %MAX_PATH                           '
THREADED WaitEvent AS LONG
THREADED fptr AS STRING POINTER                                   '

   MEntry
   '----- Use the file's data to establish a Watch event
   fptr = rfile                                                   ' Get file parameter
   passfile = gResultFile                                         '
   fn = DIR$(passfile, TO tfd)                                    ' Get current file info
   IF ISNULL(fn) THEN FUNCTION = 8: MExitFunc                     ' Whoops! Should never happen
   sfiletime = tfd.LastWriteTime                                  ' Save original file timestamp
   sfilesize = MAK(QUAD, tFD.FileSizeLow , tFD.FileSizeHigh)      ' Save original file size
   watchdir = LEFT$(passfile, INSTR(-1, passfile, "\") - 1)       ' Extract the Dir path
   WaitEvent = FindFirstChangeNotification(watchDir, 0, _         ' Put FindFCN into WaitEvent
               %FILE_NOTIFY_CHANGE_FILE_NAME   OR _               '
               %FILE_NOTIFY_CHANGE_ATTRIBUTES  OR _               '
               %FILE_NOTIFY_CHANGE_SIZE        OR _               '
               %FILE_NOTIFY_CHANGE_LAST_WRITE)                    '
   IF WaitEvent = %INVALID_HANDLE_VALUE THEN _                    ' Should never happen, but ...
      FUNCTION = 8: MExitFunc                                     ' Tell mainline we couldn't start

   SLEEP 3000                                                     ' Let things settle down
   '----- Go to sleep now until something happens
   DO
      Posted = WaitForSingleObject(WaitEvent, %INFINITE)          ' Sleep till Windows Posts us
      SLEEP 100                                                   ' Wait 100 ms for Windows multiple timestamps to be done

      '----- Go see if we care about this, else sleep again
      GOSUB CheckFileDate                                         ' See if our file's status has changed
      IF ISFALSE FindNextChangeNotification(WaitEvent) THEN _     ' Tell FNCN to keep looking
         FUNCTION = 8: MExitFunc                                  ' Some kind of error, bail out to mainline
   LOOP                                                           '
   FUNCTION = 0                                                   '
   MExitFunc

'----- See if we should wake up the user
CheckFileDate:
   fn = DIR$(passfile, TO tfd)                                    ' Get current file info

   '----- If file's been deleted, just shut down
   IF ISNULL(fn) THEN                                             ' No file?
      FUNCTION = 0: MExitFunc                                     ' Just terminate
   END IF                                                         '

   '----- If file timestamp has changed, if so, tell user
   IF sfiletime <> tFD.LastWriteTime OR _                         ' See if file has changed
      sfilesize <> MAK(QUAD, tFD.FileSizeLow , tFD.FileSizeHigh) THEN ' or size
      sfiletime = tFD.LastWriteTime                               ' Save for the next time
      sfilesize = MAK(QUAD, tFD.FileSizeLow , tFD.FileSizeHigh)   '

      '----- Show user up to the first three lines of the file
      Fnm = FREEFILE                                              ' Get a file number
      OPEN ENV.SubmitDir + "\" + fn FOR INPUT ACCESS READ LOCK SHARED AS #FNm ' Open the Result file
      FILESCAN #FNm, RECORDS TO i                                 ' Get the number of records
      IF i > 0 THEN LINE INPUT #FNm, L1                           ' Some records?
      IF i > 1 THEN LINE INPUT #FNm, L2                           '
      IF i > 2 THEN LINE INPUT #FNm, L3                           '
      CLOSE #FNm                                                  ' Close it
      t = MID$(fn, LEN(fn) - 11, 8)                               ' extract the Jobnnnnn
      sDoMsgBox L1 + $CRLF + L2 + $CRLF + L3, %MB_OK OR %MB_USERICON, t
   END IF                                                         '
   RETURN                                                         '

END FUNCTION

SUB      sRetrLoad()
'---------- Load the Retrieve stack
LOCAL FNum AS LONG, fn, buf AS STRING
   MEntry
   ON ERROR GOTO RetBail
   FNum = FREEFILE: fn = ENV.INIPath + "SPFLite.SPR"              ' Get ready for open
   OPEN fn FOR BINARY AS #FNum                                    ' Open it
   IF LOF(#FNum) > 0 THEN                                         ' Something in it?
      GET$ #FNum, LOF(#FNum), buf                                 ' Read the file
      PARSE buf, gCmdRtrev(), $CRLF                               ' Assign back to the array
   ELSE                                                           ' If no data
      RESET gCmdRtrev()                                           ' Empty the array
   END IF                                                         '
   CLOSE #FNum                                                    '
   MExitSub                                                       '
RetCont:
   RESET gCmdRtrev()                                              ' Empty the array
   MExitSub

RetBail:                                                          '
   RESUME RetCont                                                 '
END SUB                                                           '

SUB      sRetrSave()
'---------- Save the Retrieve stack and the Private Clipboard
LOCAL FNum AS LONG, buf, fn AS STRING
   MEntry
   buf = JOIN$(gCmdRtrev(), $CRLF)                                ' Build the buffer
   FNum= FREEFILE: : fn = ENV.INIPath + "SPFLite.SPR"             ' Get ready to OPEN
   OPEN fn FOR BINARY AS #FNum                                    ' And finally save it all to disk
   PUT$ #FNum, buf                                                '
   SETEOF #FNum                                                   '
   CLOSE #FNum                                                    '
   MExit
END SUB

SUB      sSBStatusbarDrawItem (BYVAL HDlg AS  DWORD, BYVAL LPARAM AS LONG)
'---------- Ownerdraw for StatusBar
LOCAL lpdis AS DRAWITEMSTRUCT PTR, rc AS RECT
LOCAL zp AS ASCIIZ PTR, stxt AS ASCIIZ * 64
LOCAL Brush AS DWORD
LOCAL ALIGN, SBI, lclScheme, i, j, k, BClr, fg, bg AS LONG

   MEntry
   '----- Get access to what's going on
   lpdis = LPARAM
   rc = @lpdis.rcItem                                             ' Get box RECT
   zp = @lpdis.itemData                                           ' Get text addressable
   SBI = VAL(LEFT$(@zp, 2))                                       ' Get index to the SBTable entry
   lclScheme = TP.SBGetDfScheme(SBI)                              ' Get the Scheme number
   IF TP.SBGetOvScheme(SBI) <> 0 THEN lclScheme = TP.SBGetOvScheme(SBI) ' Override it if requested
   fg = ENV.GetClr(lclScheme, %SCFG)                              ' Extract requested FG color
   bg = ENV.GetClr(lclScheme, %SCBG1)                             ' Extract requested BG color
   stxt = TRIM$(TP.SBGetText(SBI))                                ' Extract the real text

   '----- Create the BG and position bar brushes
   Brush = CreateSolidBrush(bg)                                   ' Brush for the background
   FillRect @lpdis.hDC, rc, Brush                                 ' Fill the background
   DeleteObject Brush                                             '
   SetBkColor @lpdis.hDC, bg                                      ' Set BG for text printing
   SetTextColor @lpdis.hDC, fg                                    ' Set the text color

   '----- Set alignment differently for some boxes
   ALIGN = %DT_CENTER
   IF TP.SBGetAlign(SBI) = "L" THEN                               ' If not CENTER then
     rc.left = rc.left + 5                                        ' Pad the left side a bit
     ALIGN = %DT_LEFT                                             ' Left align some fields
   END IF

   '----- Finally draw the text
   DrawText @lpdis.hDC, stxt, LEN(stxt), rc, %DT_SINGLELINE OR ALIGN
   IF ISFALSE IsFMTab THEN                                        ' If not FM
      IF TP.SBGetPosBar(SBI) = "Y" AND TP.LastLine > gwScrHeight - 3 THEN ' Is this the Misc box and # lines > screen height?
         Brush = CreateSolidBrush(bg XOR &H00FFFFFF)              ' Brush for position bar
         k = rc.nRight - rc.nLeft                                 ' Get width of box
         i = MAX(INT((gwScrHeight / MAX(TP.LastLine, gwScrHeight)) * k), 5)    ' Length of bar (min 5)
         j = MIN(INT((TP.TopScrn / MAX(TP.LastLine, gwScrHeight)) * k), k - i) ' Starting pos
         rc.nLeft += j: rc.nRight = rc.nLeft + i: rc.nTop = rc.nBottom - 2     ' 2 pixels at the bottom
         FillRect @lpdis.hDC, rc, Brush                           ' Draw line for relative file position
         DeleteObject Brush                                       '
      END IF                                                      '
   END IF                                                         '
   MExit
END SUB

FUNCTION sSelectColor(BYVAL hParent AS LONG, BYVAL iStartColor AS LONG) AS LONG
'---------- Use common Dialog to get a colour
LOCAL cca AS ChooseColorApi, tt AS STRING
LOCAL ccTemp AS CustColor
LOCAL rc, i AS LONG
   MEntry
   tt = sINIGetString("General",  "CustomClr", "0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0")
   FOR i = 0 TO 15                                                ' Fetch our custom colours
      ccTemp.cc(i) = VAL(PARSE$(tt, i + 1))                       '
   NEXT i                                                         '
   sPopReady                                                      ' Ready for pop-up
   DISPLAY COLOR hWnd, , , iStartColor, ccTemp, %CC_FULLOPEN TO rc
   sPopReset                                                      ' Reset popup state
   tt = ""                                                        '
   FOR i = 0 TO 15                                                ' Store our custom colours
      tt = tt + FORMAT$(ccTemp.cc(i)) + ","                       '
   NEXT i                                                         '
   tt = LEFT$(tt, LEN(tt) - 1)                                    '
   sINISetString("General",  "CustomClr", tt)                     '
   FUNCTION = IIF(rc = -1, iStartColor, rc)                       ' Return new if selected or the original
   MExit
END FUNCTION

FASTPROC SetCmd()                                                 ' Put cursor at Command line
   TP.CsrRow = 1: TP.CsrCol = 11                                  '
END FASTPROC                                                      '

FASTPROC SetScrl()                                                ' Put cursor at Scroll Amount
   TP.CsrRow = 1: TP.CsrCol = 21 + pCmdLen                        '
END FASTPROC                                                      '

FUNCTION sFCS32Update (fcs AS DWORD, BYVAL pBuffer AS DWORD, BYVAL bufSize AS DWORD) AS DWORD
'---------- Compute FCS hash for a buffer
   #REGISTER NONE
   ! mov esi, fcs
   ! mov ebx, bufSize
   ! XOR edi, edi
   ! jmp EndOfByte
   NextByte:
   ! mov edx, [esi]
   ! mov ecx, edx
   ! shr ecx, 8
   ! mov eax, pBuffer
   ! movzx eax, BYTE PTR [eax+edi]
   ! XOR edx, eax
   ! AND edx, &h0FF
   ! XOR ecx, DWORD PTR FCS32table[edx*4]
   ! mov [esi], ecx
   ! inc edi
   EndOfByte:
   ! cmp edi, ebx
   ! jb NextByte
   ! mov FUNCTION, ecx
EXIT FUNCTION

END FUNCTION

SUB sFCS32Init(fcs AS DWORD)
'---------- Initialize an FCS 32 accumulator
   fcs = &h0FFFFFFFF
END SUB

SUB sFCS32Final(fcs AS DWORD)
'---------- Fianalize an FCS 32 accumulator
   fcs = (NOT fcs)
END SUB

FUNCTION sSetTable(cmd AS STRING, operand AS STRING) AS STRING
'---------- Handle updates to the SET table
LOCAL fn, key, tdata AS STRING, i, j, k, fnum AS LONG
LOCAL SetVar() AS STRING, SetVarCtr AS LONG, SetKey AS STRING
   MEntry
   SELECT CASE AS CONST$ cmd                                      ' Which kind do we have

      CASE "GET"                                                  ' GET (fetch)
         IF gSetCount > 0 THEN                                    ' Anything in table?
            ARRAY SCAN gSetKey(), COLLATE UCASE, =operand, TO i   ' Can we find the key?
            IF i = 0 THEN FUNCTION = "8Unknown SET variable": MExitFunc
            REDIM SetVar(1 TO PARSECOUNT(gSetData(i), BINARY))    ' Dim variable table
            PARSE gSetData(i), SetVar(), BINARY                   ' Extract the table
            FUNCTION = "0" + TRIM$(SetVar(1)): MExitFunc          ' Return the first item (top of stack)
         ELSE                                                     '
            FUNCTION = "8Unknown SET variable": MExitFunc         '
         END IF                                                   '

      CASE "DEL"                                                  ' DEL
         IF gSetCount > 0 THEN                                    ' Anything in table?
            ARRAY SCAN gSetKey() FOR gSetCount, COLLATE UCASE, =operand, TO i ' Can we find the key?
            IF i = 0 THEN FUNCTION = "8Unknown SET variable": MExitFunc
            ARRAY DELETE gSetKey(i)                               ' Delete it
            ARRAY DELETE gSetData(i)                              '
            DECR gSetCount                                        '
            sUpdSetTable                                          ' Write the table
            FUNCTION = "0SET variable removed": MExitFunc         '
         ELSE                                                     '
            FUNCTION = "8Unknown SET variable": MExitFunc         '
         END IF                                                   '

      CASE "SET"                                                  ' SET
         key = LEFT$(operand, INSTR(operand, " ") - 1)            ' Separate key and data
         tdata = MID$(operand, INSTR(operand, " ") + 1)           '
         IF VERIFY(key, $AlphaNum + ".?*_") <> 0 THEN _           ' Only reasonable characters, plus '.?*_'
            FUNCTION = "8Invalid characters in SET variable name.": MExitFunc
         IF VERIFY(LEFT$(key, 1), $Numeric + ".?*_") = 0 THEN _   ' No leading numbers
            FUNCTION = "8Invalid characters in SET variable name.": MExitFunc
         ARRAY SCAN gSetKey(), COLLATE UCASE, =key, TO i          ' Can we find the key?
         IF i > 0 THEN                                            ' Already exist?
            REDIM SetVar(1 TO PARSECOUNT(gSetData(i), BINARY)) AS STRING   ' Dim variable table
            PARSE gSetData(i), SetVar(), BINARY                   ' Extract the table
            SetVar(1) = tdata                                     ' Swap new data into top slot
            gSetData(i) = JOIN$(SetVar(), BINARY)                 '
         ELSE                                                     '
            INCR gSetCount                                        ' Bump
            IF gSetCount > UBOUND(gSetKey()) THEN                 ' Add space if needed
               REDIM PRESERVE gSetKey(1 TO UBOUND(gSetKey()) * 2) AS STRING
               REDIM PRESERVE gSetData(1 TO UBOUND(gSetData()) * 2) AS STRING
            END IF                                                '
            REDIM SetVar(1 TO 1) AS STRING                        ' Dim variable table
            SetVar(1) = tdata                                     ' Swap new data into top slot
            gSetKey(gSetCount) = key                              ' Doesn't exist, build a new item
            gSetData(gSetCount) = JOIN$(SetVar(), BINARY)         '
         END IF                                                   '
         ARRAY SORT gSetKey() FOR gSetCount, TAGARRAY gSetData()  '
         sUpdSetTable                                             ' Write the table
         FUNCTION = "0SET variable stored"                        ' Setup return

      CASE "POP"                                                  ' POP (fetch)
         IF gSetCount > 0 THEN                                    ' Anything in table?
            ARRAY SCAN gSetKey(), COLLATE UCASE, =operand, TO i   ' Can we find the key?
            IF i = 0 THEN FUNCTION = "8Unknown SET variable": MExitFunc
            IF PARSECOUNT(gSetData(i), BINARY) = 1 THEN _         ' Just one entry?
               FUNCTION = "8No additional values to POP for: " + operand: MExitFunc ' Say no can do
            REDIM SetVar(1 TO PARSECOUNT(gSetData(i), BINARY))    ' Dim variable table
            PARSE gSetData(i), SetVar(), BINARY                   ' Extract the table
            ARRAY DELETE SetVar(1)                                ' Delete it
            REDIM PRESERVE SetVar(1 TO UBOUND(SetVar()) - 1)      ' Shrink the stack table
            gSetData(i) = JOIN$(SetVar(), BINARY)                 '
            FUNCTION = "0" + TRIM$(SetVar(1))
            sUpdSetTable                                          ' Write the table
            MexitFunc                                             ' Return the first item (top of stack)
         ELSE                                                     '
            FUNCTION = "8Unknown SET variable": MExitFunc         '
         END IF                                                   '

      CASE "PUSH"                                                 ' PUSH
         IF INSTR(operand, " ") THEN                              ' Got both key and data?
            key = LEFT$(operand, INSTR(operand, " ") - 1)         ' Separate key and data
            tdata = MID$(operand, INSTR(operand, " ") + 1)        '
         ELSE                                                     ' Just the key
            key = TRIM$(operand)                                  '
            tdata = ""                                            '
         END IF                                                   '
         ARRAY SCAN gSetKey(), COLLATE UCASE, =key, TO i          ' Can we find the key?
         IF i = 0 THEN FUNCTION = "8Unknown SET variable": MExitFunc
         REDIM SetVar(1 TO PARSECOUNT(gSetData(i), BINARY)) AS STRING   ' Dim variable table
         PARSE gSetData(i), SetVar(), BINARY                      ' Extract the table
         IF ISNULL(tdata) THEN tdata = SetVar(1)                  ' If no value, Dup the top value
         REDIM PRESERVE SetVar(1 TO UBOUND(SetVar()) + 1)         ' Expand the stack table
         ARRAY INSERT SetVar(1), tdata                            ' Insert the value
         gSetData(i) = JOIN$(SetVar(), BINARY)                    '
         ARRAY SORT gSetKey() FOR gSetCount, TAGARRAY gSetData()  '
         sUpdSetTable                                             ' Write the table
         FUNCTION = "0" + TRIM$(tdata): MExitFunc                 ' Return the first item (top of stack)

   END SELECT                                                     '
   MExit
END FUNCTION

SUB sSetupSB()
REGISTER i AS LONG
REGISTER j AS LONG
   i = TP.SBCount
   CONTROL SET SIZE hWnd, %IDC_StatusBar,  gSBWidth, gSBHeight' Set the SB size
   SELECT CASE i                                                  ' Pick the one that matches the count
      CASE 1: STATUSBAR SET PARTS hWnd, %IDC_STATUSBAR, TP.SBGetXrWidth(1), 9999
      CASE 2: STATUSBAR SET PARTS hWnd, %IDC_STATUSBAR, TP.SBGetXrWidth(1), _
                                                        TP.SBGetXrWidth(2), 9999
      CASE 3: STATUSBAR SET PARTS hWnd, %IDC_STATUSBAR, TP.SBGetXrWidth(1), _
                                                        TP.SBGetXrWidth(2), _
                                                        TP.SBGetXrWidth(3), 9999
      CASE 4: STATUSBAR SET PARTS hWnd, %IDC_STATUSBAR, TP.SBGetXrWidth(1), _
                                                        TP.SBGetXrWidth(2), _
                                                        TP.SBGetXrWidth(3), _
                                                        TP.SBGetXrWidth(4), 9999
      CASE 5: STATUSBAR SET PARTS hWnd, %IDC_STATUSBAR, TP.SBGetXrWidth(1), _
                                                        TP.SBGetXrWidth(2), _
                                                        TP.SBGetXrWidth(3), _
                                                        TP.SBGetXrWidth(4), _
                                                        TP.SBGetXrWidth(5), 9999
      CASE 6: STATUSBAR SET PARTS hWnd, %IDC_STATUSBAR, TP.SBGetXrWidth(1), _
                                                        TP.SBGetXrWidth(2), _
                                                        TP.SBGetXrWidth(3), _
                                                        TP.SBGetXrWidth(4), _
                                                        TP.SBGetXrWidth(5), _
                                                        TP.SBGetXrWidth(6), 9999
      CASE 7: STATUSBAR SET PARTS hWnd, %IDC_STATUSBAR, TP.SBGetXrWidth(1), _
                                                        TP.SBGetXrWidth(2), _
                                                        TP.SBGetXrWidth(3), _
                                                        TP.SBGetXrWidth(4), _
                                                        TP.SBGetXrWidth(5), _
                                                        TP.SBGetXrWidth(6), _
                                                        TP.SBGetXrWidth(7), 9999
      CASE 8: STATUSBAR SET PARTS hWnd, %IDC_STATUSBAR, TP.SBGetXrWidth(1), _
                                                        TP.SBGetXrWidth(2), _
                                                        TP.SBGetXrWidth(3), _
                                                        TP.SBGetXrWidth(4), _
                                                        TP.SBGetXrWidth(5), _
                                                        TP.SBGetXrWidth(6), _
                                                        TP.SBGetXrWidth(7), _
                                                        TP.SBGetXrWidth(8), 9999
      CASE 9: STATUSBAR SET PARTS hWnd, %IDC_STATUSBAR, TP.SBGetXrWidth(1), _
                                                        TP.SBGetXrWidth(2), _
                                                        TP.SBGetXrWidth(3), _
                                                        TP.SBGetXrWidth(4), _
                                                        TP.SBGetXrWidth(5), _
                                                        TP.SBGetXrWidth(6), _
                                                        TP.SBGetXrWidth(7), _
                                                        TP.SBGetXrWidth(8), _
                                                        TP.SBGetXrWidth(9), 9999
      CASE 10:STATUSBAR SET PARTS hWnd, %IDC_STATUSBAR, TP.SBGetXrWidth(1), _
                                                        TP.SBGetXrWidth(2), _
                                                        TP.SBGetXrWidth(3), _
                                                        TP.SBGetXrWidth(4), _
                                                        TP.SBGetXrWidth(5), _
                                                        TP.SBGetXrWidth(6), _
                                                        TP.SBGetXrWidth(7), _
                                                        TP.SBGetXrWidth(8), _
                                                        TP.SBGetXrWidth(9), _
                                                        TP.SBGetXrWidth(10), 9999
      CASE 11:STATUSBAR SET PARTS hWnd, %IDC_STATUSBAR, TP.SBGetXrWidth(1), _
                                                        TP.SBGetXrWidth(2), _
                                                        TP.SBGetXrWidth(3), _
                                                        TP.SBGetXrWidth(4), _
                                                        TP.SBGetXrWidth(5), _
                                                        TP.SBGetXrWidth(6), _
                                                        TP.SBGetXrWidth(7), _
                                                        TP.SBGetXrWidth(8), _
                                                        TP.SBGetXrWidth(9), _
                                                        TP.SBGetXrWidth(10), _
                                                        TP.SBGetXrWidth(11), 9999
      CASE 12:STATUSBAR SET PARTS hWnd, %IDC_STATUSBAR, TP.SBGetXrWidth(1), _
                                                        TP.SBGetXrWidth(2), _
                                                        TP.SBGetXrWidth(3), _
                                                        TP.SBGetXrWidth(4), _
                                                        TP.SBGetXrWidth(5), _
                                                        TP.SBGetXrWidth(6), _
                                                        TP.SBGetXrWidth(7), _
                                                        TP.SBGetXrWidth(8), _
                                                        TP.SBGetXrWidth(9), _
                                                        TP.SBGetXrWidth(10), _
                                                        TP.SBGetXrWidth(11), _
                                                        TP.SBGetXrWidth(12), 9999
      CASE 13:STATUSBAR SET PARTS hWnd, %IDC_STATUSBAR, TP.SBGetXrWidth(1), _
                                                        TP.SBGetXrWidth(2), _
                                                        TP.SBGetXrWidth(3), _
                                                        TP.SBGetXrWidth(4), _
                                                        TP.SBGetXrWidth(5), _
                                                        TP.SBGetXrWidth(6), _
                                                        TP.SBGetXrWidth(7), _
                                                        TP.SBGetXrWidth(8), _
                                                        TP.SBGetXrWidth(9), _
                                                        TP.SBGetXrWidth(10), _
                                                        TP.SBGetXrWidth(11), _
                                                        TP.SBGetXrWidth(12), _
                                                        TP.SBGetXrWidth(13), 9999
      CASE 14:STATUSBAR SET PARTS hWnd, %IDC_STATUSBAR, TP.SBGetXrWidth(1), _
                                                        TP.SBGetXrWidth(2), _
                                                        TP.SBGetXrWidth(3), _
                                                        TP.SBGetXrWidth(4), _
                                                        TP.SBGetXrWidth(5), _
                                                        TP.SBGetXrWidth(6), _
                                                        TP.SBGetXrWidth(7), _
                                                        TP.SBGetXrWidth(8), _
                                                        TP.SBGetXrWidth(9), _
                                                        TP.SBGetXrWidth(10), _
                                                        TP.SBGetXrWidth(11), _
                                                        TP.SBGetXrWidth(12), _
                                                        TP.SBGetXrWidth(13), _
                                                        TP.SBGetXrWidth(14), 9999
      CASE 15:STATUSBAR SET PARTS hWnd, %IDC_STATUSBAR, TP.SBGetXrWidth(1), _
                                                        TP.SBGetXrWidth(2), _
                                                        TP.SBGetXrWidth(3), _
                                                        TP.SBGetXrWidth(4), _
                                                        TP.SBGetXrWidth(5), _
                                                        TP.SBGetXrWidth(6), _
                                                        TP.SBGetXrWidth(7), _
                                                        TP.SBGetXrWidth(8), _
                                                        TP.SBGetXrWidth(9), _
                                                        TP.SBGetXrWidth(10), _
                                                        TP.SBGetXrWidth(11), _
                                                        TP.SBGetXrWidth(12), _
                                                        TP.SBGetXrWidth(13), _
                                                        TP.SBGetXrWidth(14), _
                                                        TP.SBGetXrWidth(15), 9999
   END SELECT
END SUB

FUNCTION sSourceLoad(TblName AS STRING, A2S AS STRING, S2A AS STRING) AS LONG
'---------- Load a specified SOURCE table
LOCAL CP AS CODEPAGE_CP_T                                         ' CODEPAGE data
LOCAL TX_TABLE_PTR AS STRING PTR * 256
LOCAL Fn AS STRING
   MEntry

   Fn = ENV.INIPath + TblName + ".SOURCE"                         ' Build the filename
   FUNCTION = %True                                               ' Start off as all is well
   IF ISFILE(Fn) THEN                                             ' See if the Customized file exists
      IF ISFALSE sRead_CodePage_Source_File(CP, TblName) THEN     ' Go read the SOURCE file
         sDoMsgBox "|K" + Fn + "|B failed validation, Null translate tables will be used", %MB_OK OR %MB_USERICON, "SPFLite"
         FUNCTION = %False: MExitFunc                             '
      END IF                                                      '

      TX_TABLE_PTR = VARPTR (CP.TX(%AE_MODE).TX_TABLE (0))        ' Pick up pointer to table
      A2S = @TX_TABLE_PTR                                         ' Pass it back
      TX_TABLE_PTR = VARPTR (CP.TX(%EA_MODE).TX_TABLE (0))        ' Now the reverse
      S2A = @TX_TABLE_PTR                                         '

   ELSE                                                           ' Else build two 'do nothing' translate tables
      sDoMsgBox "|K" + Fn + "|B was not found, Null translate tables will be used", %MB_OK OR %MB_USERICON, "SPFLite"
      FUNCTION = %False                                           '
   END IF                                                         '
   MExit                                                          ' We're done
END FUNCTION


FUNCTION sSSet(macline AS STRING) AS LONG
'---------- Handle all the difficult ~S(...) substitutions
LOCAL i, j, k AS LONG, Sline, Skey1, SKey2, SData AS STRING
LOCAL KeyName AS STRING * 12
   MEntry
   SLine = macline
   i = INSTR(UUCASE(Sline), "~S("): IF i = 0 THEN i = INSTR(UUCASE(SLine), "^S(") ' Find next ~S(
   DO WHILE i > 0                                                 ' Loop-de-loop
      k = INSTR(i, SLine, ")")                                    ' Look for closing bracket
      IF k = 0 OR k = i + 3 THEN FUNCTION = %True: MExitFunc      ' No?? or null (), Error return
      Skey1 = MID$(Sline, i, k - i + 1)                           ' Extract S key1 ~S(xxx)
      Skey2 = MID$(Sline, i + 3, k - i - 3)                       ' Extract S key2 xxx
      SData = sSetTable("GET", SKey2)                             ' Retrieve substitution value
      IF VAL(LEFT$(SData, 1)) > 0 THEN                            ' Not found?
         REPLACE Skey1 WITH "" IN SLine                           ' Substitute ""
         macline = SLine                                          '
         FUNCTION = %True: MExitFunc                              ' No?? Error return
      END IF                                                      '
      REPLACE Skey1 WITH MID$(SData, 2) IN SLine                  ' Substitute
      i = INSTR(UUCASE(Sline), "~S("): IF i = 0 THEN i = INSTR(UUCASE(SLine), "^S(") ' Find next ~S(
   LOOP                                                           '
   macline = SLine                                                '
   FUNCTION = %False                                              ' All is well
   MExit
END FUNCTION

FUNCTION sStr2Hex(fstr AS STRING) AS STRING
'---------- Convert a string to display hex
REGISTER i AS LONG
LOCAL fstr2 AS STRING
   fstr2 = "X'"                                                   ' Initialize fstr2
   FOR i = 1 TO LEN(fstr)                                         ' Convert it to hex
      fstr2 += HEX$(ASC(MID$(fstr, i, 1)), 2)                     '
   NEXT i                                                         '
   FUNCTION = fstr2 + "'"                                         '
END FUNCTION

SUB      sTabAdd (fn AS STRING, SetProf AS STRING)
'---------- Add a new TAB to the window
LOCAL x, y, ox, oy, h, i AS LONG                                  '
   MEntry

   '---------- Expand TABS table if needed                        '
   IF (TabsNum + 1) > UBOUND(Tabs) THEN                           ' Need to resize?
      REDIM PRESERVE Tabs(TabsNum + 1) AS iObjTabData             ' Yes, bump it up
   END IF                                                         '

   '---------- Create a new Tabs entry and associated storage CLASS
   INCR TabsNum: INCR TabUnique                                   ' Add a Tab
   LET Tabs(TabsNum) = CLASS "cObjTabData"                           ' Build the Class entry
   TP = Tabs(tabsNum)                                             ' Make the new entry the active tab class
   TP.PgNumber = TabsNum                                          ' Save Page Number
   TP.PrfReadAll(%True)                                           ' Go read the DEFAULT Profile if it exixts
   TP.PicInit                                                     ' Initialize Picture control area
   TP.ActionCtr = 0                                               ' Reset ActionCtr
   '---------- Build the Page and fire it up                      '
   TP.WindowID = %IDC_SPFLiteWindow + TabUnique                   ' Create the Dialog ID
   x = (ENV.ScrWidth * gFontWidth) + 1                            ' X
   y = (ENV.ScrHeight * gFontHeight)                              ' Y
   gwScrHeight = ENV.ScrHeight - ENV.PFKShow                      ' Shrink data area by PFK Show area

   TAB INSERT PAGE hWnd, %IDC_SPFLiteTAB, TP.PgNumber, 0, $Empty, CALL DlgCallBack TO h
   TP.PgHandle = h                                                ' Save the handle
   CONTROL ADD GRAPHIC, TP.PgHandle, TP.WindowID, "", 0, 0, x, y
   GRAPHIC ATTACH TP.PgHandle, TP.WindowID                        ' Set as the default graphic area
   CONTROL HANDLE TP.PgHandle, TP.WindowID TO h                   ' Save handle to graphic window
   TP.gHandle = h                                                 '
   GRAPHIC CLEAR cTxtLoBG1                                        ' Clear the background
   TP.cCurrent = %False                                           ' Set cursor state
   GRAPHIC SET FONT hScrFont                                      ' Set the font
   TP.ScreenDim(ENV.ScrHeight, ENV.ScrWidth)                      ' Redim the Screen shadow copy

   '---------- Initialize Data Areas                              '
   TP.OFrmFPath = ENV.FMPath                                      ' Save any where we were started from values
   TP.OFrmFMask = ENV.FMMask                                      '
   TP.OFrmFileL = ENV.FMFileList                                  '
   TP.LInitTxtData(fn)                                            ' Initialize our Text area
   TP.TMode = (TP.TMode AND %MFMTab)                              ' Kill all but FMTab
   TP.TMode = (TP.TMode OR ENV.PMode)                             ' Copy the global requests on top

   '----- Get IO set up
   IF ISFALSE IsClip AND ISFALSE IsSetEdit THEN                   ' If not a clipboard type startup
      TP.TIPSetup("E" + IIF$(TP.CurrPCmd = "EDIT" OR TP.CurrPCmd = "MEDIT", "R", ""), SetProf, "", fn) ' Exit + ROTest if needed
      IF TP.TIPEXEC THEN _                                        ' Go check Exist and RO
         scError(%eFail, TP.TIPResultMsg): MExitSub               ' Issue the error message
   END IF                                                         '
   TP.InitaFile(%False)                                           ' Initialize file stuff
   TP.UndoInit                                                    ' Init the Undo file names
   TP.UndoSave()                                                  ' Take an initial one
   TP.SetStart()                                                  ' Do initial positioning
   TP.WindowTitle                                                 ' Alter window/Tab titles
   gTabSwitch = TP.PgNumber                                       ' Set to switch to this tab
   sTabStackAdd(TP.PgNumber)                                      '
   MExit
END SUB

SUB      sTabBGFill (BYVAL phDC AS DWORD)
'---------- Fill background of Dialog tab
LOCAL rectFill, rectClient AS RECT                                '
LOCAL hBrush AS DWORD                                             '
   MEntry
   GetClientRect WindowFromDC(phDC), rectClient                   '
   SetRect rectFill, 0, 0, rectClient.nright + 1, rectClient.nbottom
   hBrush = CreateSolidBrush(RGB(100, 100, 100))                  '
   Fillrect phDC, rectFill, hBrush                                '
   DeleteObject hBrush                                            '
   MExit
END SUB

SUB      sTabDel()
'---------- Delete a tab
LOCAL DelTab, i, j, lx, SaveTabsNum AS LONG, Retr, MSG AS STRING
   MEntry
   ON ERROR GOTO EndBail
   '----- Followed by an EXIT?
   MSG = UUCASE(TP.CmdStackNext)                                  ' See if any pending commands
   IF MSG = "=X" OR MSG = "EXIT" THEN gShutFlag = %True           ' Remember we have to do this

   '----- Shut a TAB down
   SCaretDestroy                                                  '
   '----- Tell everyone to avoid tab structures
   SaveTabsNum = TabsNum                                          '

   '----- OK, finally destroy the tab data areas
   DelTab = TP.PgNumber                                           ' Save PageNum
   DECR TabsNum                                                   ' Reduce active count

   '----- Remove from the Dialog
   TAB DELETE hWnd, %IDC_SPFLiteTAB, TP.PgNumber                  ' Delete the tab

   '----- Free the Tabs data areas, don't understand it all, but it works
   TP = NOTHING                                                   '
   FOR lx = 1 TO SaveTabsNum                                      ' Now lets delete all its data
      IF Tabs(lx).PgNumber = DelTab THEN EXIT FOR                 ' Find the Tabs() entry we're deleting
   NEXT lx                                                        '
   j = UBOUND(Tabs)                                               ' Get the UBOUND
   Tabs(lx) = NOTHING                                             ' Free the Instance data
   FOR i = lx TO j - 1                                            ' Shift down the table
      Tabs(i) = Tabs(i + 1)                                       '
   NEXT                                                           '
   Tabs(j) = NOTHING                                              ' Make last entry empty
   IF j > 0 THEN                                                  ' Resize the table
      REDIM PRESERVE Tabs(0 TO j - 1) AS GLOBAL IObjTabData       '
   ELSE                                                           '
      ERASE Tabs()                                                '
   END IF                                                         '

   '----- Adjust things for the deleted tab
   IF TabsNum > 0 THEN                                            ' If we're still active
      FOR i = 1 TO TabsNum                                        ' Now lets reset the page numbers
         Tabs(i).PgNumber = i                                     '
      NEXT i                                                      '
   END IF                                                         '
   sFileQueue("T", FORMAT$(DelTab), "")                           ' Go adjust FileQueue for deleted tab
EndResume:
   ON ERROR GOTO 0
   MExitSub

EndBail:
   RESUME EndResume
END SUB

SUB      sTabAddFManager ()
'---------- Add the File Manager Tab data areas
LOCAL i AS LONG
   MEntry
   TP.PgNumber = TabsNum                                          ' Save Page Number
   TP.WindowID = %IDC_SPFLiteWindow + TabUnique                   ' Create the Dialog ID
   TP.PrfReset                                                    ' Reset Profile variables to the default
   TP.PrfSetProfName("Default")                                   ' Setup as default
   TP.PrfReadAll(%True)                                           ' Go read the DEFAULT Profile if it exixts
   TP.PicInit                                                     ' Initialize Picture control area
   TP.ActionCtr = 0                                               ' Reset ActionCtr

   '---------- Initialize non-INI file stuff                      '
   gDataLen  = ENV.ScrWidth - gLNPadCol                           ' Calc derived values
   pCmdLen   = ENV.ScrWidth - 24                                  ' Check sWindowCmd lengths as well
   TP.TMode = TP.TMode OR %MFMTab                                 ' Flag this as a File Manager Tab
   TP.TabTitleSet(%True)                                          '
   gwScrHeight = ENV.ScrHeight - ENV.PFKShow                      ' Shrink data area by PFK Show area

   TP.FPath = ENV.FMPath                                          ' Establish FM startup values
   IF ISNULL(TP.FPath) THEN TP.FPath = "C:\"                      '
   TP.FMask = ENV.FMMask                                          '
   IF ISNULL(TP.FMask) THEN TP.FMask = "*"                        '
   TP.FileListNm = ENV.FMFileList                                 '
   IF TP.FileListNm = "$Null$" THEN TP.FileListNm = ""            '
   TP.LFPath      = TP.FPath                                      ' Save as 'previous' values
   TP.LFMask      = TP.FMask                                      '
   TP.LFileListNm = TP.FileListNm                                 '
   TP.LastLine = 3                                                ' Prevent Editor opens
   ENV.FMPath = "":  ENV.FMMask = "": ENV.FMFileList = ""         ' Reset the Global fields
   TP.ScrlAmtC    = ENV.FMScrlAmt                                 '
   TP.AttnDo = (TP.AttnDo OR %LoadReq)                            ' Trigger initial Req and Data load
   TP.SetupFMSBXref                                               ' Setup correct SB fields for FM
   sTabStackAdd(1)
   MExit                                                          ' Done Here
END SUB

SUB      sTabHighLight(HDLG AS LONG, wParm AS LONG, lParm AS LONG)
'---------- Highlight the selected Dialog tab
LOCAL lDISPtr AS DRAWITEMSTRUCT PTR, zCap AS ASCIIZ * 50
LOCAL ti AS TC_ITEM
LOCAL hBrush, active, modified, BGColor AS LONG
   MEntry
   lDisPtr = lparm                                                '
   ti.mask = %TCIF_TEXT                                           '
   ti.pszText = VARPTR(zCap)                                      '
   ti.cchTextMax = SIZEOF(zCap)                                   '
   TabCtrl_GetItem(GetDlgItem(HDLG, wParm), @lDisptr.itemID, ti)  '
   @lDisptr.rcItem.nTop = @lDisptr.rcItem.nTop + 2                '
   IF @lDisPtr.ItemState = %ODS_SELECTED THEN active = %True      '
   IF ASC(LEFT$(zCap, 1)) > 127 THEN                              ' Set Modified/NonModified color
      zCap = CHR$(ASC(LEFT$(zCap, 1)) - 128) + MID$(zCap, 2)      '
      modified = %True                                            '
   END IF                                                         '
   IF active THEN                                                 ' Set text color based on active
      SetTextColor @lDisPtr.hDc, IIF(modified, cATabModFG, cATabNModFG)
      BGColor = IIF(modified, cATabModBG1, cATabNModBG1)          '
   ELSE                                                           '
      SetTextColor @lDisPtr.hDc, IIF(modified, cITabModFG, cITabNModFG)                             '
      BGColor = IIF(modified, cITabModBG1, cITabNModBG1)          '
   END IF                                                         '

   SetBkColor @lDisPtr.hDc, BGColor                               '
   hBrush = CreateSolidBrush(BGColor)                             '
   SelectObject @lDisptr.hDc, hBrush                              '
   FillRect @lDisptr.hDc, @lDisptr.rcItem, hBrush                 '
   SelectObject @lDisPtr.hDc, hSBFont                             '
   DrawText @lDisptr.hDc, zCap, LEN(zCap), @lDisptr.rcItem, %DT_SINGLELINE OR %DT_CENTER
   DeleteObject hBrush                                            '
   MExit
END SUB

SUB      sTabStackAdd(tbnum AS LONG)
'---------- Add tab number to the stack
REGISTER i AS LONG
   IF tbnum = gTabStack(1) THEN EXIT SUB                          ' Exit quickly for the commonest case

   FOR i = 1 TO gTabStackNum                                      ' See if already in the list
      IF gTabStack(i) = tbnum THEN                                ' Found it lower down?
         ARRAY DELETE gTabStack(i)                                ' Remove it from the old location
         ARRAY INSERT gtabStack(1), tbnum                         ' Add it back at the top
         EXIT SUB                                                 ' We're done
      END IF                                                      '
   NEXT i                                                         '

   '----- Not found in the stack, add it
   IF gTabStackNum + 1 > UBOUND(gTabStack()) THEN REDIM PRESERVE gTabStack(1 TO 2* gTabStackNum) AS GLOBAL LONG
   ARRAY INSERT gTabStack(1), tbnum                               ' Insert this at the top
   INCR gTabStackNum                                              ' count it
END SUB

SUB      sTabStackDel(tbnum AS LONG)
'---------- Remove tab number to the stack
REGISTER i AS LONG
REGISTER j AS LONG

   FOR i = 1 TO gTabStackNum                                      ' See if already in the list
      IF gTabStack(i) = tbnum THEN                                ' Found it lower down?
         j = i                                                    ' Save the one to be deleted
      ELSE                                                        ' Not the one being deleted
         IF gTabStack(i) > tbnum THEN                             ' If a higher tab number
            DECR gTabStack(i)                                     ' Reduve by the deleted tab
         END IF                                                   '
      END IF                                                      '
   NEXT i                                                         '
   ARRAY DELETE gTabStack(j)                                      ' Remove it from the old location
   DECR gTabStackNum                                              ' Adjust count
END SUB

FUNCTION sTagVal(ptag AS STRING, lookup AS LONG) AS LONG
'---------- Validate a basic Tag operand
LOCAL t AS STRING
   MEntry
   t = UUCASE(ptag)                                               ' Get an uppercase copy
   FUNCTION = %True                                               ' Start as failure
   IF LEFT$(ptag, 1) <> ":" THEN MExitFunc                        ' No starting :
   IF VERIFY(2, t, "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789") <> 0 THEN MExitFunc ' Valid characters?
   IF LEN(t) > 8 THEN MExitFunc                                   ' Valid length?
   IF t = ":ZALL" THEN FUNCTION = %False: MExitFunc               '
   t = LSET$(t, 8)                                                ' Make it 8
   IF ISTRUE lookup THEN                                          ' Validate the tagname
      IF ISFALSE TP.LLTagScan(t) THEN MExitFunc                   ' Better exist
   END IF                                                         '
   FUNCTION = %False                                              ' Say it's OK
   MExit
END FUNCTION

FUNCTION sTime() AS STRING
'---------- Return Time
LOCAL MyTime AS STRING, hh AS LONG
   MEntry
   MyTime = TIME$ + " AM"                                         ' Get the time  (hh:mm:ss)
   hh = VAL(LEFT$(MyTime, 2))                                     ' Make it a 12 hr clock
   IF hh > 12 THEN                                                '
      hh -= 12                                                    '
      MID$(MyTime, 1, 2) = FORMAT$(hh, "00")                      '
      MID$(MyTime, 10, 2) = "PM"                                  '
   END IF                                                         '
   FUNCTION = MyTime                                              '
   MExit
END FUNCTION

FUNCTION sToolTipCreate (BYVAL Wnd AS LONG) AS LONG
'---------- Create tooltips control if needed.                    '
   MEntry
   IF hToolTips = 0 THEN                                          '
      IF Wnd = 0 THEN Wnd = GetActiveWindow()                     '
      IF Wnd = 0 THEN MExitFunc                                   '
      InitCommonControls                                          '
      hToolTips = CreateWindowEx(0, "tooltips_class32", "", %TTS_ALWAYSTIP OR %TTS_BALLOON, _
             0, 0, 0, 0, Wnd, BYVAL 0&, GetModuleHandle(""), BYVAL %NULL)
   END IF                                                         '
   FUNCTION = hToolTips                                           '
   MExit
END FUNCTION

FUNCTION sToolTipSet (BYVAL Wnd AS LONG, BYVAL Txt1 AS STRING) AS LONG
'---------- Add a tooltip to a window/control
LOCAL ti AS TOOLINFO                                              '
   MEntry                                                         '
   IF ENV.WineMode THEN MExitFunc                                 ' Skip this under WINE
   IF sToolTipCreate(GetParent(Wnd)) = 0 THEN MExitFunc           ' Ensure creation
   ti.cbSize   = LEN(ti)                                          '
   ti.uFlags   = %TTF_SUBCLASS OR %TTF_IDISHWND                   '
   ti.hWnd     = GetParent(Wnd)                                   '
   ti.uId      = hWnd                                             '

   '---------- Remove existing tooltip                            '
   IF SENDMESSAGE (hToolTips, %TTM_GETTOOLINFO, 0, BYVAL VARPTR(ti)) THEN
      SENDMESSAGE hToolTips, %TTM_DELTOOL, 0, BYVAL VARPTR(ti)    '
   END IF                                                         '
   ti.cbSize   = LEN(ti)                                          '
   ti.uFlags   = %TTF_SUBCLASS OR %TTF_IDISHWND                   '
   ti.hWnd     = GetParent(Wnd)                                   '
   ti.uId      = Wnd                                              '
   ti.lpszText = STRPTR(Txt1)                                     '
   FUNCTION = SENDMESSAGE(hToolTips, %TTM_ADDTOOL, 0, BYVAL VARPTR(ti)) 'add tooltip
   MExit
END FUNCTION

FUNCTION StrCmpr(str1 AS STRING, str2 AS STRING) AS LONG
'---------- Case insensitive string compare
LOCAL S3, s4 AS DWORD, match AS LONG
' --------------------------------------------------
' compare two basic dynamic strings case insensitive
' Return values.
' -1 -- Str1 is low
'  0 -- Str1 = Str2
'  1 -- Str2 is low
' --------------------------------------------------
    #REGISTER NONE
   s3 = LEN(str1): s4 = LEN(str2)
    PREFIX "! "
    mov   esi, str1                                               ' Copy text pointers into register
    mov   edi, str2                                               '
    mov   esi, [esi]                                              ' Dereference it to get text address
    mov   edi, [edi]                                              '

    mov   ecx,s3                                                  ' Get the length of 1st string in ecx
    mov   edx,s4                                                  ' Get the lenght of 2nd string in edx
    cmp   ecx,edx                                                 ' Compare the two lengths
    pushf                                                         ' Save the current status on the stack
    jbe   LgthSet                                                 ' Put shorter length in ecx
    mov   ecx,edx                                                 '

  LgthSet:
    SUB   edx,edx                                                 ' Clear counter

  MainTest:
    AND   ecx,ecx                                                 ' Any length left?
    jz    AllEqual                                                ' No? We're done byte comparisons
    movzx eax, BYTE PTR [esi+edx]                                 ' Get Str1 byte into eax
    movzx eax, BYTE PTR Cmpi_tbl[eax]                             ' Get translated Str1 byte into eax
    movzx ebx, BYTE PTR [edi+edx]                                 ' Get Str2 byte into ebx
    movzx ebx, BYTE PTR Cmpi_tbl[ebx]                             ' Get translated Str1 byte into eax
    inc   edx                                                     ' Bump index
    dec   ecx                                                     ' Decr length left
    cmp   al, bl                                                  ' Compare Str1 translated to Str2 translated
    je    MainTest                                                ' Loop till done
    pop   ax                                                      ' Flush the length test held status from the stack

  RealMismatch:
    jb    Str1Low                                                 ' If 1st string lowest, we are done
    inc   match                                                   ' Else 2nd string lower, set Match = 1
    jmp   short SetRC                                             ' Exit

   Str1Low:
    dec   match                                                   ' 1st string low, set match to -1
    jmp   short SetRC                                             ' Exit

   AllEqual:                                                      ' All matched chars agree
    popf                                                          ' So use the result of the original length compare
    jnz   RealMismatch                                            ' If not equal, go set match status

   SetRC:

   END PREFIX
   FUNCTION = match                                               ' Pass back our answer
   EXIT FUNCTION
END FUNCTION

THREAD FUNCTION sUNDOSaveThread(BYVAL pData AS LONG) AS LONG      ' Save the UNDO data asynchronously
LOCAL lclTIDX AS STRING, lclL() AS DataLine
LOCAL USFNum, UIX AS LONG, fnU, FnT, FnTW, FnIX, buf AS STRING, bufw AS WSTRING
LOCAL supData AS UNDOType POINTER
REGISTER USi AS LONG
REGISTER USj AS LONG
   supData = pData                                                ' Copy pointer locally
   DIM lclL(@supData.UBoundL) AS DataLine

   FOR USi = 0 TO @supData.UBoundT                                ' first calculate whole size
      USj += LEN(@supData.@pT[USi]) + 2                           ' plus 2 for each line feed
   NEXT
   buf = SPACE$(USj) : USj = 0                                    ' allocate enough space

   FOR USi = 0 TO @supData.UBoundT                                ' first calculate whole size
      USj += LEN(@supData.@pTW[USi]) + 2                          ' plus 2 for each line feed
   NEXT
   bufw = SPACE$(USj) : USj = 1                                    ' allocate enough space

   FOR USi = 0 TO @supData.UBoundT                                ' then place array into string
      MID$(buf, USj) = @supData.@pT[USi]                          ' array element goes here
      USj += LEN(@supData.@pT[USi])                               ' move position ahead
      MID$(buf, USj) = $CRLF                                      ' line feed goes here
      USj += 2                                                    ' move position ahead
   NEXT
   USj = 1
   FOR USi = 0 TO @supData.UBoundT                                ' then place array into string
      MID$(bufw, USj) = @supData.@pTW[USi]                        ' array element goes here
      USj += LEN(@supData.@pTW[USi])                              ' move position ahead
      MID$(bufw, USj) = $CRLF                                     ' line feed goes here
      USj += 2                                                    ' move position ahead
   NEXT

   '----- Make copies of the data to be written
   lclTIDX = @supData.@pTIDX                                      ' Make copies of the data

   MEMORY COPY @supData.pL, VARPTR(lclL(0)), @supData.UBoundL * SIZEOF(DataLine) ' Copy the L() table

   '----- Tell mainline the copies are done, they can resume
   FnIX = @supData.IXFn                                           ' Get the filenames before
   FnU = @supData.UFn                                             ' Letting mainline resume
   FnT = @supData.TFn                                             ' So TP pointer doesn't get swapped under us.
   FnTW = @supData.TWFn                                           '
   @supData.mCpyBusy = %False                                     ' Say copies are done

   '----- Now do the actual I/O at our leisure

   USFNum = FREEFILE
   OPEN FnT FOR BINARY AS #USFNum                                 ' Write the string data
   PUT$ #USFNum, buf                                              '
   SETEOF #USFNum                                                 '
   CLOSE #USFNum                                                  '
   buf = ""                                                       ' free the text strings

   USFNum = FREEFILE
   OPEN FnTW FOR BINARY AS #USFNum                                ' Write the wstring data
   PUT$$ #USFNum, bufw                                            '
   SETEOF #USFNum                                                 '
   CLOSE #USFNum                                                  '
   bufw = ""                                                      ' free the wtext strings

   USFNum = FREEFILE
   OPEN FnIX FOR BINARY AS #USFNum                                ' Write the Alloc string
   PUT$ #USFNum, lclTIDX                                          '
   SETEOF #USFNum                                                 '
   CLOSE #USFNum                                                  '
   lclTIDX = ""                                                   ' free the big string

   USFNum = FREEFILE
   OPEN FnU FOR BINARY AS #USFNum                                 ' Write the L() array
   PUT #USFNum, ,lclL()                                           '
   SETEOF #USFNum                                                 '
   CLOSE #USFNum                                                  '
   ERASE lclL()                                                   ' free the line table
   @supData.Busy = %False                                         ' Clear this slot's Busy indicator
END FUNCTION

SUB sUnicodeGetTable (fileName AS STRING, BYREF u AS utab_t)
'---------- Load a 'utab' with code definitions from .Unicode file, if present
'           if there is no file found, create a default translation table, but mark it
'           as 'not valid'.  the table is actually 'valid' but when the flag is false
'           it's simply not needed.
LOCAL fileNum  AS LONG
LOCAL n        AS LONG
LOCAL uError   AS LONG
LOCAL lineNum  AS LONG
LOCAL anum     AS DWORD       '/ holds value of 2 digit ANSI code
LOCAL unum     AS DWORD       '/ holds value of 4 digit Unicode
LOCAL uLine AS STRING
LOCAL uWork AS STRING
LOCAL astring  AS STRING      '/ ANSI hex str
LOCAL ustring  AS STRING      '/ Unicode hex str
LOCAL uReason  AS STRING      '/ reason text for msgbox

   MEntry

   u.valid = %false                                   '/ false means not needed
   uError = %false
   lineNum = 0

   '/ convert ANSI 00 to FF to Unicode
   '/ PB converts this to Unicode with its own internal translation table

   FOR n = 0 TO 255                          '/ create default Unicode mappings
      u.uchar(n) = CHR$(n)  '/ translation happens; it's not just an assignment
   NEXT

   IF fileName = "" THEN                '/ a /Unicode file was not found before
      MExitSub
   END IF

   fileNum = FREEFILE

   TRY                                 '/ we already confirmed file was present
      OPEN fileName FOR INPUT AS # fileNum
   CATCH                               '/ this code is just to be extra careful
      sDoMsgBox "Can't Open |K" & fileName, %MB_OK OR %MB_USERICON, "Load UNICODE Table"  '/  should not occur
      MExitSub
   END TRY

   u.valid = %true                             '/ assume unless an error found

   DO
      LINE INPUT # fileNum, uLine
      IF EOF (fileNum) THEN EXIT DO   '/ an empty .Unicode file is not an error
      INCR lineNum

      uWork = UUCASE(REMOVE$(uLine, " "))

      IF uWork = "" THEN ITERATE                               '/ line is blank
      IF LEFT$ (uWork, 1) = "[" THEN ITERATE         '/ ignore [Unicode] header
      IF LEFT$ (uWork, 1) = ";" THEN ITERATE                    '/ comment line
      uWork &= ";"

      '/ X00=U0020; ' '    null as space      '/ this is the format of the data
      '/ 1234567890                         '/ 2 digits for ANSI, 4 for Unicode

      IF LEN (uWork) < 10            _                  '/ reject short records
      OR MID$(uWork, 1, 1) <> "X" _              '/ X delimiter must be present
      OR MID$(uWork, 4, 2) <> "=U" _             '/ = delimiter must be present
      OR MID$(uWork,10, 1) <> ";" THEN           '/ ; delimiter must be present
         u.valid = %false
         uError = %true
         uReason = "invalid delimiters"
        EXIT DO
      END IF

      astring = MID$(uWork, 2, 2)
      ustring = MID$(uWork, 6, 4)

      IF VERIFY (astring, "0123456789ABCDEF") <> 0 _       '/ must be valid hex
      OR VERIFY (ustring, "0123456789ABCDEF") <> 0 THEN
         u.valid = %false
         uError = %true
         uReason = "invalid hex values: " & astring & " " & ustring
         EXIT DO
      END IF

      anum = VAL ("&H" & astring)             '/ convert hex values to integers
      unum = VAL ("&H" & ustring)             '/ anum must be in range 00 to FF
      u.uchar(anum) = CHR$$(unum)                 '/ store entry in unicode tab
   LOOP

   IF uError THEN
      sDoMsgBox "Unicode file format invalid: " & uReason & $CRLF _
         & " File: |K" & fileName & $CRLF _
         & "|B Line |K" & DEC$(lineNum) & ": " & uLine, %MB_OK OR %MB_USERICON, "Load UNICODE File"
   END IF

   CLOSE # fileNum

   MExit

END SUB

FUNCTION sUnicodeGetTableName(BYVAL argPathname AS STRING, _
                              tQualifier        AS STRING, _
                              tType             AS STRING) AS STRING
'---------- Locate a .Unicode file
'           if a type-specific table exists (like Print or Display) use it
'           otherwise, look for a generic .Unicode file, in case a single table is
'           being used for both purposes.  function returns the file name
'           if neither version of file is found, return null string

LOCAL fileName AS STRING

   MEntry

   '/ if table type is "Print" then file will be SPFLite.Print.Unicode

   fileName = argPathname & tQualifier & "." & tType & ".Unicode"
   '/ use type-specific table

   IF ISFILE(fileName) THEN
      FUNCTION = fileName
   ELSE
      fileName = argPathname & tQualifier & ".Unicode"
      '/ use common table

      IF ISFILE(fileName) THEN
         FUNCTION = fileName
      ELSE
         FUNCTION = ""
      END IF
   END IF

   MExit

END FUNCTION

SUB      sUnQuote(str AS STRING)
'---------- Remove quotes from a string
   MEntry
   IF INSTR($Quotes, LEFT$(str, 1)) <> 0 AND INSTR($Quotes, RIGHT$(str, 1)) <> 0 THEN _
      str = MID$(str, 2, LEN(str) - 2)                            ' Remove quotes if present
   MExit
END SUB

SUB      sUpdSetTable()
'---------- Re-write the SET table
LOCAL FNum, i, j, k AS LONG, fn AS STRING
LOCAL SetVar() AS STRING, SetVarCtr AS LONG, SetKey AS STRING
   MEntry
   FNum = FREEFILE: fn = ENV.INIPath + "SPFLite.SPS"              ' Re-write the table
   OPEN fn FOR OUTPUT AS #FNum                                    ' Open the SPS File
   IF gSetCount > 0 THEN                                          ' Only if there are entries
      FOR i = 1 TO gSetCount                                      '

         '----- Each SET item could be a Stack
         k = PARSECOUNT(gSetData(i), BINARY)                      ' Get count of number in stack
         REDIM SetVar(1 TO k) AS STRING                           ' Dim variable table
         PARSE gSetData(i), SetVar(), BINARY                      ' Extract the table

         '----- 1st item geta Key=value; remainder get just =value style
         FOR j = 1 TO k                                           ' For eack stack item
            IF j = 1 THEN
               PRINT #FNum, gSetKey(i) + "=" + SetVar(1)          ' Write the Data
            ELSE                                                  '
               PRINT #FNum, "=" + SetVar(j)                       '
            END IF
         NEXT j                                                   '
      NEXT i                                                      '
   END IF
   SETEOF #FNum                                                   '
   CLOSE #FNum                                                    '
   MExit
END SUB

FUNCTION sUtf8FromAnsi (BYREF a_str AS STRING) AS STRING
'---------- Convert an ANSI string to UTF8
'           it is necessary to convert ANSI to Unicode first
'           then convert Unicode to UTF-8
LOCAL u_str                   AS STRING                           ' output UTF-8 string
LOCAL u_num1                  AS LONG                             ' first  UTF-8 byte
LOCAL u_num2                  AS LONG                             ' second UTF-8 byte
LOCAL u_num3                  AS LONG                             ' third  UTF-8 byte
LOCAL u_len                   AS LONG                             ' length of output string

LOCAL w_num                   AS LONG                             ' binary value of Unicode char
LOCAL w_char                  AS WSTRING * 1                      ' one Unicode char

LOCAL a_ndx                   AS LONG                             ' index into input string
LOCAL a_char                  AS STRING * 1                       ' one Ansi char

   MEntry

   IF LEN (a_str) = 0 THEN FUNCTION = "": MExitFunc               ' supplied Ansi string was null, return null

   u_len = 0                                                      ' where new string written

   '----- estimated size must take into account that Unicode translations of Ansi may take 2 or 3 bytes in UTF-8

   u_str = SPACE$ (LEN (a_str) * 3)                               ' allocate result string max length, may need less
   FOR a_ndx = 1 TO LEN (a_str)                                   ' examine each character in the input string

      a_char = MID$ (a_str, a_ndx, 1)                             ' grab one Ansi char
      w_char = a_char                                             ' convert to Unicode

      w_num = ASC(w_char) AND &H0FFFF                             ' get binary value of Unicode

      IF (w_num <= &H07F) AND (w_num > 0) THEN                    ' ASCII is a 1-byte sequence
         u_len += 1                                               ' get next position in output string
         MID$ (u_str, u_len, 1) = a_char                          ' copy Ansi to output as is

      '----- 2-byte UTF-8 has up 11 significant bits, divided into 5 and 6

      ELSEIF (w_num <= &H07FF) OR (w_num = 0) THEN                ' data requires a 2-byte sequence
         u_num2 = ((w_num AND &H03F) OR &H080)                    ' byte 2 of UTF-8 gets 6 lower bits, plus x'80' marker
         SHIFT RIGHT w_num, 6
         u_num1 = ((w_num AND &H01F) OR &H0C0)                    ' byte 1 of UTF-8 gets 5 upper bits, plus x'C0' marker

         u_len += 1                                               ' get next position in output string
         MID$ (u_str, u_len, 1) = CHR$ (u_num1)
         u_len += 1                                               ' get next position in output string
         MID$ (u_str, u_len, 1) = CHR$ (u_num2)

      '----- 3-byte UTF-8 has up 16 significant bits, divided into 4, 6 and 6

      ELSE                                                        ' data requires a 3-byte sequence
         u_num3 = ((w_num AND &H03F) OR &H080)                    ' byte 3 of UTF-8 gets 6 lower bits, plus x'80' marker
         SHIFT RIGHT w_num, 6

         u_num2 = ((w_num AND &H03F) OR &H080)                    ' byte 2 of UTF-8 gets next 6 bits, plus x'80' marker
         SHIFT RIGHT w_num, 6

         u_num1 = ((w_num AND &H0F) OR &H0E0)                     ' byte 1 of UTF-8 gets 4 upper bits, plus x'E0' marker

         u_len += 1                                               ' get next position in output string
         MID$ (u_str, u_len, 1) = CHR$ (u_num1)
         u_len += 1                                               ' get next position in output string
         MID$ (u_str, u_len, 1) = CHR$ (u_num2)
         u_len += 1                                               ' get next position in output string
         MID$ (u_str, u_len, 1) = CHR$ (u_num3)

      END IF
   NEXT
   FUNCTION = LEFT$ (u_str, u_len)                                ' return the encoded UTF-8 string
   MExit
END FUNCTION

FUNCTION sUtf8ToAnsi (BYREF u_str AS STRING, err_flag AS LONG) AS STRING
'---------- Converts an UTF-8 encoded string to an Ansi string
'
'  err_flag 1 = value found outside Ansi 8-bit range
'  err_flag 2 = UTF-8 value is malformed, errors replaced with X'A4' substitution
'  err_flag 3 = errors 1 and 2 both occurred

'  if valid UTF-8 converts to a Unicode value that is not Ansi, we can't use it

STATIC st_valid_unicode_set   AS LONG
STATIC st_valid_unicode ()    AS WSTRING * 1

LOCAL a_str                   AS STRING                           ' output Ansi string
LOCAL a_char                  AS STRING * 1                       ' an Ansi char

LOCAL w_str                   AS WSTRING                          ' intermediate Unicode string
LOCAL w_len                   AS LONG                             ' length of output string
LOCAL w_num                   AS LONG                             ' accumulated Unicode char value
LOCAL w_char                  AS WSTRING * 1                      ' a Unicode char

LOCAL u_ndx                   AS LONG                             ' index into input string
LOCAL u_word                  AS LONG                             ' UTF accumulation value
LOCAL u_byte                  AS LONG                             ' individual UTF value in binary
LOCAL u_char                  AS STRING * 1                       ' a UTF-8 char

LOCAL format_err              AS LONG                             ' something wrong with UTF format
LOCAL data_err                AS LONG                             ' may be valid UTF but is out of range

LOCAL one_bits                AS LONG                             ' number of leading 1 bits in  a UTF byte
LOCAL select_mask             AS LONG                             ' used to select bits for testing
LOCAL test_mask               AS LONG                             ' used to compare bits after being selected
LOCAL data_mask               AS LONG                             ' used to mask off data portion of byte
LOCAL avail_len               AS LONG                             ' how many bytes left in the buffer
LOCAL actual_len              AS LONG                             ' byte length used to decode UTF data
LOCAL n                       AS LONG

   MEntry

   '----- first time we are called, define the Unicode validation table

   IF st_valid_unicode_set = 0 THEN
      st_valid_unicode_set = 1
      DIM st_valid_unicode (0 TO 255) AS STATIC WSTRING * 1

      FOR n = 0 TO 255
         st_valid_unicode (n) = CHR$ (n)                          ' the assignment converts Ansi to Unicode here
      NEXT
   END IF


   err_flag = 0

   IF LEN (u_str) = 0 THEN
      FUNCTION = ""                                               ' nothing to convert
      MExit
      EXIT FUNCTION
   END IF

   w_len = 0                                                      ' where new string written
   w_str = SPACE$ (LEN (u_str))                                   ' allocate result string max length, may need less

   u_ndx = 0                                                      ' examine each character in the UTF-8 string
   DO WHILE (u_ndx < LEN (u_str))

      u_ndx += 1                                                  ' get next position in input string
      u_char = MID$ (u_str, u_ndx, 1)                             ' grab first UTF-8 byte
      u_byte = (ASC(u_char) AND &H0FF)                            ' form binary of UTF value

      '----- determine the number of leading 1 bits

      one_bits = 0
      select_mask = &H0100
      data_mask   = &H00FF

      '----- scan the value left to right, stopping when first 0 bit is found
      '      data mask is 1 bit to right of select_mask

      FOR n = 1 TO 8
         SHIFT RIGHT select_mask, 1
         SHIFT RIGHT data_mask, 1
         IF ((u_byte AND select_mask) = 0) THEN EXIT FOR
         one_bits = n
      NEXT

      '----- when one_bits = 0 is it ASCII
      '      1 leading 1-bit is a continuation byte without a lead byte
      '      leading 1 bits of 2-4 are UTF-8, any more is invalid

      IF one_bits = 0 THEN                                        ' UTF byte is 7-bit ASCII
         w_len += 1                                               ' get next position in buffer
         MID$ (w_str, w_len, 1) = CHR$$ (u_byte)                  ' copy non-encoded Ansi to intermeidate buffer
         ITERATE LOOP

      ELSEIF (one_bits = 1) OR (one_bits > 4) THEN
         err_flag = (err_flag OR 2)                               ' make note of format error
         w_len += 1                                               ' get next position in output string
         MID$ (w_str, w_len, 1) = CHR$$ (&H0A4)                   ' substitute char is logenze/square/currency symbol
         ITERATE LOOP

      END IF

      '----- at this point, the UTF-8 value is 2, 3 or 4 bytes long, indicated by number of 1 bits
      '
      '      there are two "format issues" to contend with.
      '      first, there may not be that many bytes left in the buffer (it's "truncated")
      '      second, all of the extra bytes must be proper continuation bytes, with '10' bits on the left
      '      if either of these requirements are not met, the UTF-8 is malformed
      '
      '      there is also a "data issue".  a properly formatted UTF-8 string may be encoding value
      '      that is outside the range of Ansi.  that is nothing wrong with the UTF-8.  it is simply
      '      a limitation of SPFLite, which can't handle true Unicode

      format_err = 0                                              ' assume no error for now
      avail_len = LEN (u_str) - u_ndx + 1                         ' how many bytes left including curr one

      IF one_bits > avail_len THEN
         format_err = 1
         actual_len = avail_len
      ELSE
         actual_len = one_bits
      END IF

      w_num = (u_byte AND data_mask)                              ' holds accumulated Unicode value

      '----- get remaining bytes of UTF-8.  these must all be continuation bytes

      FOR n = 2 TO actual_len

         u_ndx += 1                                               ' get next position in input string
         u_char = MID$ (u_str, u_ndx, 1)                          ' grab next UTF-8 byte
         u_byte = (ASC(u_char) AND &H0FF)                         ' form binary of Unicode value

         IF ((u_byte AND &H0C0) <> &H080) THEN                    ' this is not a continuation byte - FORMAT ERROR
            format_err = 1
            u_ndx -= 1                                            ' we cannot use this byte - BACK OFF THE SCAN HERE
            EXIT FOR                                              ' no point in trying to converting any more of this
         END IF

         SHIFT LEFT w_num, 6                                      ' make room for next 6 bits of data
         w_num += (u_byte AND &H03F)
      NEXT

      '----- see if the accumulated Unicode value can be converted to Ansi

      IF w_num > &H0FFFF THEN
         data_err = 1                                             ' no Ansi value has a Unicode value > FFFF

      ELSEIF w_num > &H07F THEN
         data_err = 1                                             ' assume it's bad unless found in the table
         w_char = CHR$$ (w_num)

         FOR n = 128 TO 255
            IF w_char = st_valid_unicode (n) THEN
               data_err = 0                                       ' if found it table, it can be converted
               EXIT FOR
            END IF
         NEXT

      ELSE
         data_err = 0                                             ' data is not out of range
      END IF


      IF format_err THEN
         err_flag = (err_flag OR 2)                               ' make note of format error
         w_char = CHR$$ (&H0A4)                                   ' substitute char is logenze/square/currency symbol

      ELSEIF data_err THEN
         err_flag = (err_flag OR 1)                               ' make note of data error
         w_char = CHR$$ (&H0A4)                                   ' substitute char is logenze/square/currency symbol

      ELSE
         w_char = CHR$$ (w_num)

      END IF

      w_len += 1                                                  ' get next position in intermediate buffer
      MID$ (w_str, w_len, 1) = w_char                             ' store one Unicode char

   LOOP

   a_str = LEFT$ (w_str, w_len)                                   ' convert Unicode to Ansi
   FUNCTION = a_str                                               ' return the Ansi string
   MExit
END FUNCTION

FUNCTION sValidate_CodePage_Data(CP AS CodePage_CP_T) AS LONG
'/-----------------------------------------------------------------------------/
'/  Validate_CodePage_Data                                                     /
'/                                                                             /
'/  The complete .SOURCE file has been read, parsed and stored into the Code   /
'/  Page data structures.  Wenow validate thet we have a usable definition.    /
'/  In addition, if one of the two tables was omitted, assuming that Roundtrip /
'/  Mode is in effect, we synthesize the missing table from the one that is    /
'/  present, since the two tables are symmetrical.                             /
'/                                                                             /
'/  Return 1 if valid, else store error message in CP_Reason and Return 0      /
'/                                                                             /
'/  Method: If initial storage of CodePage_CP_T created any errors, stop now.  /
'/  Otherwise, validate each component, and return errors if these components  /
'/  have any problems.  If all the components report valid status, so do we.   /
'/-----------------------------------------------------------------------------/
LOCAL RETCODE, I, RT_Mode, AE_INDEX, AE_VALUE, EA_INDEX, EA_VALUE AS LONG
LOCAL D_CHAR, D, U AS LONG
    FUNCTION = 0                                         '/ Default as error
    IF  CP.CP_Errors > 0 THEN EXIT FUNCTION              '/ Errors found earlier
    RETCODE = sValidate_CodePage_TT_Data (CP.TT)

    IF  RETCODE = 0 THEN                                 '/ Problem with TT data
        CP.CP_Reason = CP.TT.TT_Reason
        EXIT FUNCTION
    END IF

    IF  CP.TT.TT_Mode  = "RT" THEN
        RT_Mode = 1
    ELSE
        RT_Mode = 0
    END IF

    IF  CP.TX(%AE_Mode).TX_Defined = 0 _
    AND CP.TX(%EA_Mode).TX_Defined = 0 THEN
        CP.CP_Reason = "AE/EA TABLES ARE NOT DEFINED"
        EXIT FUNCTION
    END IF

    '/-------------------------------------------------------------------------/
    '/  If round-trip mode, and only one table is defined, invert the table to /
    '/  create the missing one.                                                /
    '/-------------------------------------------------------------------------/
    IF  RT_Mode = 1 THEN
        IF  CP.TX(%AE_Mode).TX_Defined <> CP.TX(%EA_Mode).TX_Defined THEN

            '/ Figure out which one is defined and set indexes to make the
            '/ logic easier to deal with
            IF  CP.TX(%AE_Mode).TX_Defined = 1 THEN
                D = %AE_Mode                                   '/ AE was defined
                U = %EA_Mode                                 '/ EA was undefined
            ELSE
                D = %EA_Mode                                   '/ EA was defined
                U = %AE_Mode                                 '/ AE was undefined
            END IF
            IF  CP.TX(D).TX_Values = 256 _
            AND CP.TX(U).TX_Values = 0   THEN
                '/ Copy error count, whatever it is
                CP.TX(U).TX_Errors  = CP.TX(D).TX_Errors
                CP.TX(U).TX_Values  = 256              '/ Force to correct value
                CP.TX(U).TX_Defined = 1                '/ Force to correct value

                '/ Copy and invert the characters
                FOR I = 0 TO 255
                    D_CHAR = CP.TX(D).TX_Table (I)
                    CP.TX(U).TX_Table (D_CHAR) = I
                NEXT

                '/ Replicate the TX_Entry list
                FOR I = 0 TO 15
                    CP.TX(U).TX_Entry (I) = CP.TX(D).TX_Entry (I)
                NEXT
            END IF
        END IF ' TX_Defined values differ
    END IF ' RT_Mode = 1

    '/-------------------------------------------------------------------------/
    '/  Now that inverted table was created, if necessary, finish the table    /
    '/  validation                                                             /
    '/-------------------------------------------------------------------------/
    FOR I = %TX_ASCII TO %TX_EBCDIC
        RETCODE = sValidate_CodePage_TX_Data (CP.TX(I), I, RT_Mode)
        IF  RETCODE = 0 THEN                             '/ Problem with TX data
            CP.CP_Reason = CP.TX(I).TX_Reason
            EXIT FUNCTION
        END IF
    NEXT

    '/-------------------------------------------------------------------------/
    '/  Cross-validate round-trap CodePage tables.  If RT_Mode, verify that    /
    '/  every byte value is double-translated bacl to itself                   /
    '/                                                                         /
    '/  Example case: AE[39] = F9, so EA[F9] must = 39                         /
    '/-------------------------------------------------------------------------/
    IF  RT_Mode = 1 THEN
        FOR AE_INDEX = 0 TO 255                                 '/ AE_Index = 39

            AE_VALUE = CP.TX (%AE_Mode).TX_Table (AE_INDEX)  ' --> AE_Value = F9
            EA_INDEX = AE_VALUE                              ' --> EA_Index = F9

            EA_VALUE = CP.TX (%EA_Mode).TX_Table (EA_INDEX)     '/ EA_Value = 39

            '/ We forced AE_Value = EA_Index at "-->" now do inverse test
            IF  AE_INDEX <> EA_VALUE THEN
                CP.CP_Reason = "MODE=RT BUT AE/EA TABLES ARE NOT " +           _
                    "MUTUALLY ROUND-TRIP"
                EXIT FUNCTION
            END IF
        NEXT
    END IF
    FUNCTION = 1                                       '/ Data structures are OK
END FUNCTION ' sValidate_CodePage_Data

FUNCTION sValidate_CodePage_TT_Data(TT AS CODEPAGE_TT_T)  AS LONG
'/-----------------------------------------------------------------------------/
'/  Validate_CodePage_TT_Data                                                  /
'/                                                                             /
'/  Fields validated: AE, EA and MODE.                                         /
'/  AE and EA must be 0, 1 or blank; BLANK defaults to 1.                      /
'/  REMAINING FIELDS IGNORED                                                   /
'/                                                                             /
'/  If MODE is RT, one of AE or EA can be omitted.                             /
'/                                                                             /
'/  Return 1 if valid, else return 0                                           /
'/-----------------------------------------------------------------------------/
    FUNCTION = 0                                         '/ Default as error
    IF  TT.TT_Errors > 0 THEN EXIT FUNCTION              '/ Errors found earlier

    '/-------------------------------------------------------------------------/
    '/  MODE value must be 'RT', 'ES' or BLANK; BLANK defaults to RT           /
    '/-------------------------------------------------------------------------/
    IF  TRIM$(TT.TT_Mode) = "" THEN TT.TT_Mode = "RT"
    TT.TT_Mode = UUCASE(TT.TT_Mode)
    IF  TT.TT_Mode <> "RT"  _
    AND TT.TT_Mode <> "ES"  THEN
        TT.TT_Errors += 1
        TT.TT_Reason = "INVALID: MODE=" + TT.TT_Mode
        EXIT FUNCTION
    END IF
    FUNCTION = 1                                    '/ All required fields valid
END FUNCTION ' sValidate_CodePage_TT_Data

FUNCTION  sValidate_CodePage_TX_Data(TX AS CODEPAGE_TX_T, TX_Index AS LONG, RT_Mode AS LONG) AS LONG
'/-----------------------------------------------------------------------------/
'/  Validate_CodePage_TX_Data                                                  /
'/                                                                             /
'/  Fields validated: TX_Type                                                  /
'/                                                                             /
'/  Return 1 if valid, else return 0                                           /
'/-----------------------------------------------------------------------------/
LOCAL Deduced_Type        AS STRING
LOCAL TX_ID               AS STRING
DIM   Test (0 TO 255)     AS BYTE
LOCAL Byte_Index          AS BYTE
REGISTER I                AS LONG
    FUNCTION = 0                                         '/ Set failure code
    '/ TX_Index is either %TX_ASCII or %TX_EBCDIC
    IF  TX.TX_Errors > 0 THEN EXIT FUNCTION              '/ Errors found earlier

    IF  TX_Index = %TA_Index THEN
        TX_ID = "AE"
    ELSEIF TX_Index = %TE_Index THEN
        TX_ID = "EA"
    ELSE
        TX_ID = "??"                                         '/ Should not occur
    END IF

    '/-------------------------------------------------------------------------/
    '/  Deduce type of CodePage from condition of digit '9'                    /
    '/  ASCII-to-ASCII is possible but unusual; EBCDIC-EBCDIC is invalid       /
    '/  we have to be able to deduce which kind of table we are dealing with   /
    '/-------------------------------------------------------------------------/
    IF     TX.TX_Table(&H039) = &H0F9 THEN           '/ Table is ASCII to EBCDIC
        Deduced_Type = "ASCII"
    ELSEIF TX.TX_Table(&H0F9) = &H039 THEN           '/ Table is EBCDIC to ASCII
        Deduced_Type = "EBCDIC"
    ELSEIF TX.TX_Table(&H039) = &H039 THEN            '/ Table is ASCII to ASCII
        Deduced_Type = "ASCII"

    '/ if we have a 'null' table that translates to itself, it's possible that
    '/ the entry at [F9] will equal F9.  it's not necessarily an error

    ELSEIF TX.TX_Table(&H0F9) = &H0F9 THEN          '/ Table is EBCDIC to EBCDIC
        Deduced_Type = "ASCII"

    '/  TX.TX_Errors += 1
    '/  TX.TX_Reason = TX_ID + " INVALID: TABLE IS EBCDIC/EBCDIC"
    '/  EXIT FUNCTION

    ELSE
        TX.TX_Errors += 1
        TX.TX_Reason = TX_ID + " INVALID: [39] = " + HEX$(TX.TX_Table(&H39))   _
            + " [F9] = " + HEX$(TX.TX_Table(&HF9))
        EXIT FUNCTION
    END IF

    '/-------------------------------------------------------------------------/
    '/  Make initial assumption about type based on TX index                   /
    '/-------------------------------------------------------------------------/
    TX.TX_Type = UUCASE(TRIM$(TX.TX_Type))
    IF  TX.TX_Type = "" THEN
        IF TX_Index = %TX_ASCII THEN
            TX.TX_Type = "ASCII"
        ELSEIF TX_Index = %TX_EBCDIC THEN
            TX.TX_Type = "EBCDIC"
        END IF
    END IF

 '/ '/-------------------------------------------------------------------------/
 '/ '/  Type must be 'ASCII' or 'EBCDIC'                                       /
 '/ '/-------------------------------------------------------------------------/
 '/ IF  TX.TX_Type <> "ASCII"  _
 '/ AND TX.TX_Type <> "EBCDIC" THEN
 '/     TX.TX_Errors += 1
 '/     TX.TX_Reason = "INVALID: TYPE=" + TX.TX_Type
 '/     EXIT FUNCTION
 '/ END IF
 '/
 '/ '/-------------------------------------------------------------------------/
 '/ '/  'ASCII' or 'EBCDIC' type must agree with contents of table             /
 '/ '/-------------------------------------------------------------------------/
 '/ IF  TX.TX_Type <> Deduced_Type THEN
 '/     TX.TX_Errors += 1
 '/     TX.TX_Reason = "TYPE=" + TX.TX_Type + " INCONSISTENT WITH CODE DATA"
 '/     EXIT FUNCTION
 '/ END IF

    '/-------------------------------------------------------------------------/
    '/  A Code Page must have 256 values                                       /
    '/-------------------------------------------------------------------------/
    IF  TX.TX_Values <> 256 THEN
        TX.TX_Errors += 1
        TX.TX_Reason = "TYPE=" + TX.TX_Type + " HAS " + TRIM$(TX.TX_Values)    _
            + ", EXPECTING 256"

        '/ If the number of values is low, see if we can find a missing row
        FOR I = 0 TO 15
            IF  TX.TX_Entry (I) = 0 THEN
                TX.TX_Reason += ", " + TX_ID + " ROW " + HEX$ (I) + "_ MISSING"
                EXIT FOR
            END IF
        NEXT
        EXIT FUNCTION
    END IF

    '/-------------------------------------------------------------------------/
    '/  Each TX_Entry must have 16 values CP.TX(I).TX_Entry (T) = 0            /
    '/-------------------------------------------------------------------------/
    FOR I = 0 TO 15
        IF  TX.TX_Entry (I) <> 16 THEN
            TX.TX_Errors += 1
            TX.TX_Reason = TX_ID + " ROW " + HEX$ (I) + "_ HAS "               _
                + TRIM$(TX.TX_Entry (I)) + " VALUES, EXPECTING 16"
            EXIT FUNCTION
        END IF
    NEXT

    '/-------------------------------------------------------------------------/
    '/  If round-trip mode, verify 100% coverage                               /
    '/-------------------------------------------------------------------------/
    IF  RT_Mode = 1 THEN
        FOR I = 0 TO 255
            Test (I) = 0                                '/ Initialize test array
        NEXT
        FOR I = 0 TO 255
            Byte_Index = TX.TX_Table (I)
            IF  Test (Byte_Index) = 1 THEN
                TX.TX_Errors += 1
                TX.TX_Reason = TX_ID + " TABLE INCONSISTENT WITH MODE=RT"
                EXIT FUNCTION
            END IF
            Test (Byte_Index) = 1
        NEXT
    END IF
    FUNCTION = 1                                    '/ all required fields valid
END FUNCTION ' sValidate_CodePage_TX_Data

FUNCTION sVSet(macline AS STRING) AS LONG
'---------- Handle all the difficult ~V(...) substitutions
LOCAL i, j, k AS LONG, Sline, Skey1, SKey2, SData AS STRING
LOCAL KeyName AS STRING * 12
   MEntry
   SLine = macline
   i = INSTR(UUCASE(Sline), "~V("): IF i = 0 THEN i = INSTR(UUCASE(SLine), "^V(") ' Find next ~V(
   DO WHILE i > 0                                                 ' Loop-de-loop
      k = INSTR(i, SLine, ")")                                    ' Look for closing bracket
      IF k = 0 OR k = i + 3 THEN FUNCTION = %True: MExitFunc      ' No?? or null (), Error return
      Skey1 = MID$(Sline, i, k - i + 1)                           ' Extract S key1 ~S(xxx)
      Skey2 = MID$(Sline, i + 3, k - i - 3)                       ' Extract S key2 xxx
      SData = ENVIRON$(SKey2)                                     ' Fetch environ string
      REPLACE Skey1 WITH SData IN SLine                           ' Substitute
      i = INSTR(UUCASE(Sline), "~V("): IF i = 0 THEN i = INSTR(UUCASE(SLine), "^V(") ' Find next ~V(
   LOOP                                                           '
   macline = SLine                                                '
   FUNCTION = %False                                              ' All is well
   MExit
END FUNCTION

FUNCTION sWinErrorMsg(BYVAL ErrorCode AS DWORD) AS STRING
'----- Return a formatted Windows error message
LOCAL pzError  AS ASCIIZ POINTER
LOCAL ErrorLen AS DWORD

   ErrorLen = FormatMessage(%FORMAT_MESSAGE_FROM_SYSTEM OR %FORMAT_MESSAGE_ALLOCATE_BUFFER, _
                            BYVAL %NULL, ErrorCode, %NULL, BYVAL VARPTR(pzError), %NULL, BYVAL %NULL)
   IF ErrorLen THEN
      FUNCTION = "Error" & STR$(ErrorCode) & " (0x" & HEX$(ErrorCode) & ") : " & @pzError
      LocalFree(pzError)
   ELSE
      FUNCTION = "Unknown error" & STR$(ErrorCode) & " (0x" & HEX$(ErrorCode) & ")"
   END IF
END FUNCTION

FUNCTION sWriteClipboard(CBData AS STRING) AS LONG
'---------- Write string to normal or private clipboard
LOCAL i AS LONG, CBError AS STRING
LOCAL WCIO AS iIO                                                 ' For our I/O stuff

   MEntry
   IF INSTR(CBData, CHR$(0)) THEN _                               '
      scError(%eFail, "Data contains a NULL X'00' character, not allowed"): MExitFunc ' Oops?  Bail out

   LET WCIO = CLASS "cIO"                                         '
   '----- Stuff it in the clipboard
   IF ISNULL(gKeyPrimOper) THEN                                   ' If no operand, use normal Clipboard
      sWinclip_set(CBData)                                        ' Write to Windows clipboard, null string OK
      FUNCTION = %True                                            ' Pass back result

   ELSE                                                           '
      IF ISNOTNULL(CBData) THEN                                   ' If something to save
         WCIO.Setup("OR", "", "", ENV.CLIPPath + gKeyPrimOper + ".CLIP") ' Set filename
         IF WCIO.EXEC THEN                                        ' Go Open it
            scError(%eFail, WCIO.ResultMsg): MExitFunc            ' Oops?  Bail out
         END IF                                                   '
         PRINT # WCIO.FNum, CBData;                               '
         WCIO.Close                                               ' Close the FBO
      ELSE                                                        '
         sRecycleBin(ENV.CLIPPath + gKeyPrimOper + ".CLIP", "D")  ' Delete any CLIP file
      END IF                                                      '
      FUNCTION = %True                                            ' Say all is well
   END IF                                                         '
   MExitFunc

END FUNCTION

'/
'/ support code for English-only case conversions
'/
'/ ASCII hex:  A=&H41   Z=&H5A   a=&H61   z=&H7A   diff=&H20
'/

FUNCTION UUCASE (BYVAL sAscii AS STRING) AS STRING
'---------- Do a character translate from lower to upper case ASCII
REGISTER n   AS LONG
LOCAL pAscii AS BYTE PTR

   IF LEN(sAscii) = 0 THEN                           '/ handle null-string case
      FUNCTION = ""

   ELSEIF ENV.ENGchars THEN                           '/ do ENGLISH-only UUCASE
      pAscii = STRPTR(sAscii)                   '/ point to first char of value

      FOR n = 1 TO LEN(sAscii)                     '/ scan ASCII looking for LC
         IF  @pAscii >= &H61  _                                       '/ LC "a"
         AND @pAscii <= &H7A  THEN                                    '/ LC "z"
             @pAscii -= &H20              '/ shift LC down to where UC ASCII is
         END IF

         INCR pAscii                          '/ bump scan pointer to next char
      NEXT
      FUNCTION = sAscii            '/ return modified string argument as result

   ELSE                                              '/ do international UUCASE
      FUNCTION = UCASE$(sAscii)
   END IF

END FUNCTION

FUNCTION LLCASE (BYVAL sAscii AS STRING) AS STRING
'---------- Do a character translate from upper to lower case ASCII
REGISTER n   AS LONG
LOCAL pAscii AS BYTE PTR

   IF LEN(sAscii) = 0 THEN                           '/ handle null-string case
      FUNCTION = ""

   ELSEIF ENV.ENGchars THEN                           '/ do ENGLISH-only LLCASE
      pAscii = STRPTR(sAscii)                   '/ point to first char of value

      FOR n = 1 TO LEN(sAscii)                     '/ scan ASCII looking for UC
         IF  @pAscii >= &H41  _                                       '/ UC "A"
         AND @pAscii <= &H5A  THEN                                    '/ UC "Z"
             @pAscii += &H20                '/ shift UC up to where LC ASCII is
         END IF

         INCR pAscii                          '/ bump scan pointer to next char
      NEXT
      FUNCTION = sAscii            '/ return modified string argument as result

   ELSE                                              '/ do international LLCASE
      FUNCTION = LCASE$(sAscii)
   END IF

END FUNCTION

FUNCTION sWinclip_set(arg_varname AS STRING) AS LONG                                       '''
'---------- SET the Windows clipboard, with retry capability
LOCAL MemPointer AS ASCIIZ PTR
LOCAL hMem AS DWORD
LOCAL loc_retry, i AS LONG
LOCAL CBError AS STRING

   MEntry

   FOR loc_retry = 1 TO 20                                        ' wait up to 2 seconds

      IF OpenClipboard(0) THEN                                    ' Open the Clipboard

         IF EmptyClipboard THEN                                   ' Empty it

            '----- Build the CB contents in global memory
            hMem = GlobalAlloc(%GHND, LEN(arg_varname) + 1)       ' Get some global memory
            MemPointer = GlobalLock(hMem)                         ' Lock it
            @MemPointer = arg_varname                             ' Copy data to it
            GlobalUnlock hMem                                     ' Unlock it

            IF SetClipboardData(%CF_TEXT, hMem) THEN              ' Set the text in
               GlobalFree(hMem)                                   ' Free the global memory
               IF CloseClipboard THEN                             ' Close the clipboard
                  FUNCTION = %True: MExitFunc                     ' Exit if OK

               ELSE                                               ' CloseClipboard error
                  CBError = "CB CLOSE ": GOSUB CBFormatError      ' If error, Format an error message
                  FUNCTION = %False: MExitFunc                    ' and exit with error
               END IF                                             '

            ELSE                                                  ' Else SetClipboard Data error
               CBError = "CB SET ": GOSUB CBFormatError           ' If error, Format an error message
               GlobalFree(hMem)                                   ' Free the global memory
               FUNCTION = %False: MExitFunc                       ' and exit with error
            END IF                                                '

         ELSE                                                     ' Else EmptyClipboard error
            CBError = "CB EMPTY ": GOSUB CBFormatError            ' If error, Format an error message
            FUNCTION = %False: MExitFunc                          ' and exit with error
         END IF                                                   '

      ELSE                                                        ' OpenClipboard error
         CBError = "CB OPEN ": GOSUB CBFormatError                ' If error, Format an error message
         FUNCTION = %False: MExitFunc                             ' and exit with error
      END IF                                                      '
      SLEEP 100                                                   ' Sleep before retry
   NEXT loc_retry                                                 ' else loop back
   FUNCTION = %False                                              ' SET failed
   MExitFunc

CBFormatError:
   sDoMsgBox "SET Clipboard Error: " + CBError + "= " + sWinErrorMsg(GetLastError), %MB_OK OR %MB_USERICON, "Clipboard"
   RETURN

END FUNCTION

FUNCTION sWinclip_get(arg_varname AS STRING) AS LONG
'---------- GET the Windows clipboard, with retry capability
LOCAL loc_retry AS LONG
LOCAL CBError AS STRING
LOCAL CBhBuffer AS ASCIIZ PTR

   MEntry

   FOR loc_retry = 1 TO 20                                        ' wait up to 2 seconds

      IF OpenClipboard(0) THEN                                    ' Open the Clipboard

         CBhBuffer = GetClipboardData(%CF_TEXT)                   ' Get some text data

         IF CBhBuffer THEN                                        ' If we get a pointer
            arg_varname = @CBhBuffer                              ' Copy the data

            IF CloseClipboard THEN                                ' Close the Clipboard
               FUNCTION = %True: MExitFunc                        ' Exit - all is well

            ELSE                                                  ' CloseClipboard failed
               CBError = "CB CLOSE ": GOSUB CBFormatError         ' If error, Format an error message
               FUNCTION = %False: MExitFunc                       ' and exit with error
            END IF                                                '

         ELSE                                                     ' GetClipboardData failed
            CBError = "CB GET DATA ": GOSUB CBFormatError         ' Format an error message
            FUNCTION = %False: MExitFunc                          ' and exit with error
         END IF                                                   '

      ELSE                                                        ' OpenClipboard failed
         CBError = "CB OPEN ": GOSUB CBFormatError                ' If error, Format an error message
         FUNCTION = %False: MExitFunc                             ' and exit with error
      END IF                                                      '

      SLEEP 100
   NEXT loc_retry                                                 ' else loop back
   FUNCTION = %False                                              ' SET failed
   MExitFunc

CBFormatError:
   sDoMsgBox "GET Clipboard Error: " + CBError + "= " + sWinErrorMsg(GetLastError), %MB_OK OR %MB_USERICON, "Clipboard"
   RETURN
END FUNCTION
