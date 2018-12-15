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

FUNCTION DispDefault(ft AS STRING) AS STRING
'---------- Display the Welcome dialog
LOCAL tfn, INIs() AS STRING, INICtr AS LONG
DIM INIs(1 TO 100) AS STRING
   MEntry
   '----- Collect the existing INI file names
   INICtr = 1: INIs(INICtr) = "(Choose)"                          ' Put header in list
   tfn = DIR$(ENV.ProfPath + "*.INI")                             ' Look for INI files
   DO WHILE ISNOTNULL(tfn)                                        ' While we're getting entries
      INCR INICtr                                                 ' Count it
      IF INICtr > UBOUND(INIs()) THEN REDIM PRESERVE INIs(1 TO UBOUND(INIs()) * 2) AS STRING
      INIs(INICtr) = UUCASE(tfn)                                  ' Save it
      tfn = DIR$(NEXT)                                            ' Get next entry
   LOOP                                                           '
   gDefaultAnswer = ""                                            ' Clear answer
   '----- Build the dialog
   DIALOG FONT DEFAULT "Tahoma", 10 / gFontScale,0, 0
   DIALOG NEW HWnd, "Create new Profile INI",,, 225, IIF(IsFMTab, 114, 80), %WS_THICKFRAME OR %WS_CAPTION TO hDef
   DIALOG SET COLOR hDef, %RGB_GAINSBORO, %RGB_GAINSBORO
   CONTROL ADD LABEL,  hDef, %WELCOME_TEXT+0, "No profile currently exists for file-type " + ft, 1, 1, 200, 11
   CONTROL SET COLOR   hDef, %WELCOME_TEXT+0, %BLUE, %RGB_GAINSBORO
   CONTROL ADD LABEL,  hDef, %WELCOME_TEXT+1, "You may choose one of the following options:", 1, 11, 247, 11
   CONTROL SET COLOR   hDef, %WELCOME_TEXT+1, %BLUE, %RGB_GAINSBORO
   CONTROL ADD BUTTON, hDef, %WELCOME_OPTION1, "DEFAULT.INI", 1, 27, 60, 13, %WS_BORDER
   CONTROL ADD LABEL,  hDef, %WELCOME_TEXT+7, "Use DEFAULT.INI for this file-type", 70, 29, 200, 11
   CONTROL SET COLOR   hDef, %WELCOME_TEXT+7, %BLUE, %RGB_GAINSBORO
   CONTROL ADD BUTTON, hDef, %WELCOME_OPTION2, ft + ".INI", 1, 43, 60, 13, %WS_BORDER
   CONTROL ADD LABEL,  hDef, %WELCOME_TEXT+8, "Create a new " + ft + ".INI for this file-type", 70, 45, 200, 11
   CONTROL SET COLOR   hDef, %WELCOME_TEXT+8, %BLUE, %RGB_GAINSBORO
   CONTROL ADD COMBOBOX, hDef, %WELCOME_OPTION3, INIs(), 1, 59, 60, 60, %CBS_DROPDOWNLIST OR %WS_VSCROLL
   COMBOBOX SELECT hDef, %WELCOME_OPTION3, 1
   CONTROL ADD LABEL,  hDef, %WELCOME_TEXT+9, "Associate (USE) a current INI for this file-type", 70, 61, 200, 11
   CONTROL SET COLOR   hDef, %WELCOME_TEXT+9, %BLUE, %RGB_GAINSBORO
   IF IsFMTab THEN
      CONTROL ADD BUTTON, hDef, %WELCOME_OPTION4, "Skip " + ft, 1, 76, 60, 13, %WS_BORDER
      CONTROL ADD LABEL,  hDef, %WELCOME_TEXT+10, "Skip " + ft + " in this and future FF searches", 70, 78, 200, 11
      CONTROL SET COLOR   hDef, %WELCOME_TEXT+10, %BLUE, %RGB_GAINSBORO
      CONTROL ADD BUTTON, hDef, %WELCOME_OPTION5, "Cancel ", 1, 92, 60, 13, %WS_BORDER
      CONTROL ADD LABEL,  hDef, %WELCOME_TEXT+11, "Cancel this FF search", 70, 94, 200, 11
      CONTROL SET COLOR   hDef, %WELCOME_TEXT+11, %BLUE, %RGB_GAINSBORO
   END IF
   sPopReady                                                      ' Ready for pop-up
   DIALOG SHOW MODAL hDef CALL DlgDefaultCallback                 ' Display it all
   sPopReset                                                      ' Reset popup state
   FUNCTION = gDefaultAnswer                                      ' Pass back the answer
   MExit
END FUNCTION

SUB     DispInterrupt()
'---------- Display the Interrupt button
LOCAL i, j AS LONG
   DIALOG GET LOC hWnd TO i, j                                    ' Get location where window was
   DIALOG NEW PIXELS, 0, "Click to Interrupt Search", i + 15, j + 15, 300, 86, %WS_CAPTION OR %WS_SYSMENU TO hIntr
   DIALOG SET COLOR     hIntr, %RGB_GAINSBORO, %RGB_GAINSBORO
   CONTROL ADD LABEL,   hIntr, %INTERRUPT_TEXT1, "Processing file:", 4, 4, 100, 16
   CONTROL SET FONT     hIntr, %INTERRUPT_TEXT1, hFixedFont
   CONTROL SET COLOR    hIntr, %INTERRUPT_TEXT1, %BLACK, %RGB_GAINSBORO
   CONTROL ADD LABEL,   hIntr, %INTERRUPT_TEXT2, " ", 14, 20, 250, 16
   CONTROL SET FONT     hIntr, %INTERRUPT_TEXT2, hFixedFont
   CONTROL SET COLOR    hIntr, %INTERRUPT_TEXT2, %BLUE, %RGB_GAINSBORO
   CONTROL ADD LABEL,   hIntr, %INTERRUPT_TEXT3, " ", 14, 36, 250, 16
   CONTROL SET FONT     hIntr, %INTERRUPT_TEXT3, hFixedFont
   CONTROL SET COLOR    hIntr, %INTERRUPT_TEXT3, %BLUE, %RGB_GAINSBORO
   CONTROL ADD BUTTON,  hIntr, %INTERRUPT_BREAK, "Break", 115, 56, 75, 28, %WS_BORDER
   DIALOG SHOW MODELESS hIntr CALL DlgIntruptCallback             ' Display it all

END SUB


SUB      DispMessageList()
'---------- Display the ErrMsgHlp array
LOCAL t AS STRING, i AS LONG
   MEntry
   FOR i = TP.ErrMsgHlpC TO 1 STEP -1                             ' Copy the entries
      t += TP.ErrMHGet(i) + $CRLF                                 ' To one string
   NEXT i                                                         '

   '----- Build the dialog
   DIALOG NEW HWnd, "Messages",,, 400, 200, %WS_THICKFRAME OR %WS_CAPTION OR %WS_SYSMENU TO hMsg
   CONTROL ADD TEXTBOX, hMsg, %MSG_ListBox, t, 0, 0, 400, 200, %ES_AUTOHSCROLL OR %ES_MULTILINE OR %ES_AUTOVSCROLL OR _
                                                      %WS_HSCROLL OR %WS_VSCROLL, %WS_EX_CLIENTEDGE
   CONTROL SET FONT     hMsg, %MSG_ListBox, hFixedFont            '
   sPopReady                                                      ' Ready for pop-up
   DIALOG SHOW MODAL    hMsg, CALL DlgMessageCallback             '
   sPopReset                                                      ' Back from pop-up
   TP.ErrMsgHlpC = 0                                              ' reset it
   MExit
END SUB

SUB      DispOptions(StartPg AS LONG)
'---------- Build and display the Options Dialog
LOCAL i, j               AS LONG
DIM MinRetrieveList(9)   AS STRING                                '
DIM NotifyList(3)        AS STRING
DIM DefSortList(10)      AS STRING
DIM DirSortList(3)       AS STRING
DIM LastDirOpenList(2)   AS STRING
DIM ScrLabels(16 TO 25)  AS STRING

   MEntry
   ARRAY ASSIGN MinRetrieveList() = "1", "2", "3", "4", "5", "6", "7", "8", "9"
   ARRAY ASSIGN NotifyList() = "None", "Edit", "All"
   ARRAY ASSIGN DefSortList() = "Name+", "Name-", "Ext+", "Ext-", "Size+", "Size-", "Date+", "Date-", "Note+", "Note-"
   ARRAY ASSIGN DirSortList() = "Dir+", "Dir-", "Dir*"
   ARRAY ASSIGN LastDirOpenList() = "Working Dir", "Last Used Dir"
   ARRAY ASSIGN ScrLabels() = "Line Numbers High Intensity", "Line Numbers Low Intensity", _
                              "Active Tab Modified", "Active Tab Unmodified", _
                              "Inactive Tab Modified", "Inactive Tab Unmodified", _
                              "PFK Help Lines", "Status Line", "FM Tool Bar", "Error Messages"

   '----- Create the main dialogue
   DIALOG FONT DEFAULT "Tahoma", 13 / gFontScale,0, 0
   DIALOG NEW hWnd, "SPFLite Global Options", , , 350, 260, %WS_POPUP OR %WS_BORDER OR %WS_DLGFRAME OR %WS_CAPTION OR _
      %WS_CLIPSIBLINGS OR %WS_VISIBLE OR %DS_MODALFRAME OR %DS_3DLOOK OR %DS_NOFAILCREATE OR _
      %DS_SETFONT, %WS_EX_CONTROLPARENT OR %WS_EX_LEFT OR %WS_EX_LTRREADING OR %WS_EX_RIGHTSCROLLBAR, TO hOpt
   DIALOG SET COLOR hOpt, %WHITE, -2

   '----- Add the Tab Control to main dialog box
   CONTROL ADD TAB,     hOpt, %TAB_CONTROL, "", 1, 1, 348, 233, %TCS_OWNERDRAWFIXED
   CONTROL ADD IMAGEX,  hOpt, %GENERAL_GPLV3, "GPLV3", 2, 235, 66, 24, %SS_ICON
   CONTROL ADD LABEL,   hOpt, %DIALOG_INI_TEXT1, "INI File is:", 74, 240, 130, 12
   CONTROL SET COLOR    hOpt, %DIALOG_INI_TEXT1, %WHITE, -2
   CONTROL ADD TEXTBOX, hOpt, %DIALOG_INI_TEXT2, ENV.INIFileName, 110, 240, 190, 11, %ES_AUTOHSCROLL
   CONTROL ADD BUTTON,  hOpt, %DIALOG_CANCEL, "Cancel", 305, 234, 40, 11, %WS_BORDER OR %BS_VCENTER OR %BS_CENTER
   sToolTipSet (GetDlgItem(hOpt, %DIALOG_CANCEL), " Press Cancel to exit Options and make NO changes. ")
   CONTROL ADD BUTTON,  hOpt, %DIALOG_DONE, "Done", 305, 247, 40, 11, %WS_BORDER OR %BS_VCENTER OR %BS_CENTER
   sToolTipSet (GetDlgItem(hOpt, %DIALOG_DONE), " Press Done once you have made all desired changes. ")

   '----- Add the General Tab
   TAB INSERT PAGE hOpt, %TAB_CONTROL, 1, 0, "General", CALL DlgOptionsCallBack TO hGeneral
   sAddCheck(hGeneral, %GENERAL_INIABeepFlag, ENV.ABeepFlag, 4, 4, 115, "Audible BEEP on Errors?")
   sToolTipSet (GetDlgItem(hGeneral, %GENERAL_INIABeepFlag), " Make a BEEP on error messages or invalid keyboard entry? ")
   sAddCheck(hGeneral, %GENERAL_INIUniqueFlag, ENV.UniqueFlag, 200, 4, 115, "Only 1 SPFLite running?")
   sToolTipSet (GetDlgItem(hGeneral, %GENERAL_INIUniqueFlag), " Use 1 Window (multiple tabs) for all SPFLite Edit sessions? ")
   sAddCheck(hGeneral, %GENERAL_INIVBeepFlag, ENV.VBeepFlag, 4, 14, 115, "Visual BEEP on Errors?")
   sToolTipSet (GetDlgItem(hGeneral, %GENERAL_INIVBeepFlag), " Do a visual Blink to show a BEEP on error messages or invalid keyboard entry? ")
   sAddCheck(hGeneral, %GENERAL_INIReOpenLast, ENV.ReOpenLast, 200, 14, 115, "Re-Open last file(s) at start?")
   sToolTipSet (GetDlgItem(hGeneral, %GENERAL_INIReOpenLast), " Open last file used file set if no other file specified on the Cmd line? ")
   sAddCheck(hGeneral, %GENERAL_INIUseRecycle, ENV.UseRecycle, 4, 24, 115, "Delete to Recycle Bin?")
   sToolTipSet (GetDlgItem(hGeneral, %GENERAL_INIUseRecycle), " If Ticked, Deleted files will be moved to the Recycle Bin. ")
   sAddCheck(hGeneral, %GENERAL_INIViewWarn, ENV.ViewWarn, 200, 24, 115, "Warn on modified View file?")
   sToolTipSet (GetDlgItem(hGeneral, %GENERAL_INIViewWarn), " If Ticked, an END for a modified View tab will generate a warning prompt. ")
   sAddCheck(hGeneral, %GENERAL_INISplash, ENV.Splash, 200, 34, 135, "Display splash screen on startup?")
   sToolTipSet (GetDlgItem(hGeneral, %GENERAL_INISplash), " If Ticked, SPFLite will display it's Splash screen whenever it starts.")
   sAddCheck(hGeneral, %GENERAL_INIDirProfFlag, ENV.DirProfFlag, 4, 34, 200, "Use DIR name AS PROFILE when no File extension?")
   sToolTipSet (GetDlgItem(hGeneral, %GENERAL_INIDirProfFlag), " If Ticked, files with no extensions will use the DIR name as the Profile name")
   sAddCheck(hGeneral, %GENERAL_FINDWORD, ENV.FindWord, 4, 44, 200, "Use WORD as the default for FIND/CHANGE commands?")
   sToolTipSet (GetDlgItem(hGeneral, %GENERAL_FINDWORD), " If Ticked, FIND/CHANGE will assume WORD mode unless otherwise specified")
   sAddCheck(hGeneral, %GENERAL_ALLOW2D, ENV.Allow2DMouse, 4, 54, 200, "Allow 2-D mouse selection without Shift/Ctrl/Alt?")
   sToolTipSet (GetDlgItem(hGeneral, %GENERAL_ALLOW2D), " If Ticked, block mode mouse selection will not need Shift/Ctrl/Alt to allow vertical movement")
   sAddCheck(hGeneral, %GENERAL_INIENGChars, ENV.ENGChars, 4, 64, 250, "Only English letters A-Z and a-z are considered alphabetic")
   sToolTipSet (GetDlgItem(hGeneral, %GENERAL_INIENGChars), " If Ticked, only A-Z and a-z will be considered alphabatic.")
   sAddCheck(hGeneral, %GENERAL_INIResetU, ENV.ResetU, 4, 74, 250, "Default RESET will Revert User line status.")
   sToolTipSet (GetDlgItem(hGeneral, %GENERAL_INIResetU), " If Ticked, a RESET command will clear all User line status, just like Exclude status.")

   CONTROL ADD COMBOBOX, hGeneral, %GENERAL_INIMinRetrieve, MinRetrieveList(), 200, 85, 40, 40, %CBS_DROPDOWNLIST OR %WS_VSCROLL
   COMBOBOX SELECT hGeneral, %GENERAL_INIMinRetrieve, ENV.MinRetrieve
   sToolTipSet (GetDlgItem(hGeneral, %GENERAL_INIMinRetrieve), " Command lines whose length is less than this value will not be saved for RETRIEVE ")
   CONTROL ADD LABEL, hGeneral, %GENERAL_INIMinRetrieve_TEXT, "Minimum command length for RETRIEVE?", 4, 86, 150, 12
   CONTROL SET COLOR  hGeneral, %GENERAL_INIMinRetrieve_TEXT, %WHITE, -2

   CONTROL ADD TEXTBOX,   hGeneral, %GENERAL_INICmdChr, ENV.CmdChr, 200, 100, 20, 10, %ES_CENTER
   sToolTipSet (GetDlgItem(hGeneral, %GENERAL_INICmdChr), " Single character used to separate multiple commands on the Command Line ")
   CONTROL POST hGeneral, %GENERAL_INICmdChr, %EM_LIMITTEXT, 1, 0
   CONTROL ADD LABEL, hGeneral, %GENERAL_INICmdChr_TEXT, "Command separator char.", 4, 101, 150, 12
   CONTROL SET COLOR  hGeneral, %GENERAL_INICmdChr_TEXT, %WHITE, -2

   CONTROL ADD LABEL, hGeneral, %GENERAL_INIQuickRenum_TEXT, "Use fast renum if # lines > this value", 4, 115, 150, 12
   CONTROL SET COLOR  hGeneral, %GENERAL_INIQuickRenum_TEXT, %WHITE, -2
   CONTROL ADD TEXTBOX,   hGeneral, %GENERAL_INIQuickRenum, FORMAT$(ENV.QuickRenum), 200, 116, 30, 10, %ES_NUMBER
   sToolTipSet (GetDlgItem(hGeneral, %GENERAL_INIQuickRenum), " Normal renum will be used when # lines is less that this, Fast renum when over ")
   CONTROL POST hGeneral, %GENERAL_INIQuickRenum, %EM_LIMITTEXT, 7, 0

   CONTROL ADD LABEL, hGeneral, %GENERAL_INIDefDataShift_TEXT, "Default # columns for data shifts (Min 1)", 4, 129, 150, 12
   CONTROL SET COLOR  hGeneral, %GENERAL_INIDefDataShift_TEXT, %WHITE, -2
   CONTROL ADD TEXTBOX,   hGeneral, %GENERAL_INIDefDataShift, FORMAT$(ENV.DefDataShift), 200, 130, 30, 10, %ES_NUMBER
   sToolTipSet (GetDlgItem(hGeneral, %GENERAL_INIDefDataShift), " Number of columns to shift for )) >> ]] line commands ")
   CONTROL POST hGeneral, %GENERAL_INIDefDataShift, %EM_LIMITTEXT, 3, 0

   CONTROL ADD LABEL, hGeneral, %GENERAL_INIAutoScroll_TEXT, "Numer of columns/lines per Auto-Scroll", 4, 143, 150, 12
   CONTROL SET COLOR  hGeneral, %GENERAL_INIAutoScroll_TEXT, %WHITE, -2
   CONTROL ADD TEXTBOX, hGeneral, %GENERAL_INIAutoScroll, FORMAT$(ENV.AutoScroll), 200, 145, 30, 10, %ES_NUMBER
   sToolTipSet (GetDlgItem(hGeneral, %GENERAL_INIAutoScroll), " Number of columns/lines to scroll per scroll event ")
   CONTROL POST hGeneral, %GENERAL_INIAutoScroll, %EM_LIMITTEXT, 3, 0

   CONTROL ADD COMBOBOX, hGeneral, %GENERAL_ININotify, NotifyList(), 200, 158, 40, 40, %CBS_DROPDOWNLIST OR %WS_VSCROLL
   COMBOBOX SELECT hGeneral, %GENERAL_ININotify, SWITCH(ENV.NotifyLevel = 0, 1, ENV.NotifyLevel = 1, 2, ENV.NotifyLevel = 2, 3)
   sToolTipSet (GetDlgItem(hGeneral, %GENERAL_ININotify), " Select Tab types to be Notified on an external file change occurring ")
   CONTROL ADD LABEL, hGeneral, %GENERAL_ININotify_TEXT, "Notify tabs on external file change", 4, 158, 150, 12
   CONTROL SET COLOR  hGeneral, %GENERAL_ININotify_TEXT, %WHITE, -2

   CONTROL ADD LABEL, hGeneral, %GENERAL_INILinRepeatMax_TEXT, "Line Commands repeat limit (0=no limit)", 4, 172, 150, 12
   CONTROL SET COLOR  hGeneral, %GENERAL_INILinRepeatMax_TEXT, %WHITE, -2
   CONTROL ADD TEXTBOX,   hGeneral, %GENERAL_INILinRepeatMax, FORMAT$(ENV.LinRepeatMax), 200, 173, 30, 10, %ES_NUMBER
   sToolTipSet (GetDlgItem(hGeneral, %GENERAL_INILinRepeatMax), " Limit Line Command repeat value to this maximum. ")
   CONTROL POST hGeneral, %GENERAL_INILinRepeatMax, %EM_LIMITTEXT, 9999, 0

   CONTROL ADD LABEL, hGeneral, %GENERAL_INICharSet_TEXT, "Normal characters for screen display and for P'.' picture literals", 4, 186, 300, 12
   CONTROL SET COLOR  hGeneral, %GENERAL_INICharSet_TEXT, %WHITE, -2
   CONTROL ADD TEXTBOX,   hGeneral, %GENERAL_INICharSet, ENV.CharSet, 4, 195, 337, 10, %ES_AUTOHSCROLL
   sToolTipSet (GetDlgItem(hGeneral, %GENERAL_INICharSet), " Enter your 'Normally' displayable screen characters. Also used for P'.' literal searches. ")

   CONTROL ADD TEXTBOX,   hGeneral, %GENERAL_INIInvChar, ENV.InvChar, 200, 207, 20, 10, %ES_CENTER
   sToolTipSet (GetDlgItem(hGeneral, %GENERAL_INIInvChar), " Display character to use when the character is not in the above list. ")
   CONTROL POST hGeneral, %GENERAL_INIInvChar, %EM_LIMITTEXT, 1, 0
   CONTROL ADD LABEL, hGeneral, %GENERAL_INIInvChar_TEXT, "Display Char. for Invalid characters.", 4, 207, 150, 12
   CONTROL SET COLOR  hGeneral, %GENERAL_INIInvChar_TEXT, %WHITE, -2

   '----- Add the File Manager Tab
   TAB INSERT PAGE hOpt, %TAB_CONTROL, 2, 0, "FM", CALL DlgOptionsCallBack TO hFManager
   sAddCheck(hFManager, %FMANAGER_INIFMCloseFlag, ENV.FMCloseFlag, 4, 4, 115, "Close SPFLite with last file tab?")
   sToolTipSet (GetDlgItem(hFManager, %FMANAGER_INIFMCloseFlag), " If Ticked, SPFLite will close when last file data tab closes. ")
   sAddCheck(hFManager, %FMANAGER_INIConfDelFlag, ENV.ConfDelFlag, 4, 14, 115, "Confirm file Deletes?")
   sToolTipSet (GetDlgItem(hFManager, %FMANAGER_INIConfDelFlag), " If Ticked, you will be asked permission before each file delete. ")
   sAddCheck(hFManager, %FMANAGER_FMHLFLAG, ENV.FMHelpFlag, 4, 24, 115, "Display File Manager Help?")
   sToolTipSet (GetDlgItem(hFManager, %FMANAGER_FMHLFLAG), " If Ticked, FM will Display 3 lines of Help @ bottom of Screen")
   sAddCheck(hFManager, %FMANAGER_FMDATEHILITE, ENV.FMDateHiLite, 4, 34, 115, "Highlight Recent / Active dates?")
   sToolTipSet (GetDlgItem(hFManager, %FMANAGER_FMDATEHILITE), " If Ticked, FM will Highlight Recent / Active Dates in the File List")

   CONTROL ADD LABEL, hFManager, %FMANAGER_INIFMLayout_TEXT, "Enter the Optional desired columns and their order", 4, 45, 200, 12
   CONTROL SET COLOR  hFManager, %FMANAGER_INIFMLayout_TEXT, %WHITE, -2
   CONTROL ADD TEXTBOX,   hFManager, %FMANAGER_INIFMLayout, ENV.FMLayout, 4, 55, 337, 10, %ES_AUTOHSCROLL OR %ES_UPPERCASE
   sToolTipSet (GetDlgItem(hFManager, %FMANAGER_INIFMLayout), " Enter the FM columns in the desired left to right order. ")
   CONTROL ADD LABEL, hFManager, %FMANAGER_INIFMLayout_Help1_TEXT, "Valid column names to choose from are:", 4, 65, 200, 10
   CONTROL SET COLOR  hFManager, %FMANAGER_INIFMLayout_Help1_TEXT, %WHITE, -2
   CONTROL ADD LABEL, hFManager, %FMANAGER_INIFMLayout_Help2_TEXT, "  EXT, DATE, DATETIME, SIZESHORT, SIZELONG, LINES and NOTE", 4, 72, 225, 10
   CONTROL SET COLOR  hFManager, %FMANAGER_INIFMLayout_Help2_TEXT, %WHITE, -2

   CONTROL ADD LABEL, hFManager, %FMANAGER_INIFMFFList_TEXT, "Enter filename extensions considered to be Non-Text files", 4, 83, 200, 12
   CONTROL SET COLOR  hFManager, %FMANAGER_INIFMFFList_TEXT, %WHITE, -2
   CONTROL ADD TEXTBOX,   hFManager, %FMANAGER_INIFMFFList, ENV.FMFFList, 4, 92, 337, 10, %ES_AUTOHSCROLL
   sToolTipSet (GetDlgItem(hFManager, %FMANAGER_INIFMFFList), " Files in this list will be treated as Non-Text; e.g. will not be earched by " + $CRLF + _
                                                              "the FF (Find in Files) command. ")

   CONTROL ADD LABEL, hFManager, %FMANAGER_INIRecentCtr_TEXT, "Number of entries in recent lists", 4, 118, 150, 12
   CONTROL SET COLOR  hFManager, %FMANAGER_INIRecentCtr_TEXT, %WHITE, -2
   CONTROL ADD TEXTBOX,   hFManager, %FMANAGER_INIRecentCtr, FORMAT$(ENV.RecentCtr), 150, 118, 20, 10, %ES_NUMBER
   sToolTipSet (GetDlgItem(hFManager, %FMANAGER_INIRecentCtr), " Specify how many entries to retain in the recent file/path lists ")
   CONTROL POST hFManager, %FMANAGER_INIRecentCtr, %EM_LIMITTEXT, 2, 0

   CONTROL ADD COMBOBOX, hFManager, %FMANAGER_INILastDirOpen, LastDirOpenList(), 150, 132, 70, 40, %CBS_DROPDOWNLIST OR %WS_VSCROLL
   COMBOBOX SELECT    hFManager, %FMANAGER_INILastDirOpen, ENV.LastDirOpen
   sToolTipSet (GetDlgItem(hFManager, %FMANAGER_INILastDirOpen), " Default Directory in which to Open files. ")
   CONTROL ADD LABEL, hFManager, %FMANAGER_INILastDirOpen_TEXT, "Default File Open Directory", 4, 134, 150, 12
   CONTROL SET COLOR  hFManager, %FMANAGER_INILastDirOpen_TEXT, %WHITE, -2

   CONTROL ADD LABEL, hFManager, %FMANAGER_INIFMLCmdWidth_TEXT, "Width of Line Command area (5 to 20)", 4, 149, 150, 12
   CONTROL SET COLOR  hFManager, %FMANAGER_INIFMLCmdWidth_TEXT, %WHITE, -2
   CONTROL ADD TEXTBOX,   hFManager, %FMANAGER_INIFMLCmdWidth, FORMAT$(ENV.FMLCmdWidth), 150, 149, 20, 10, %ES_NUMBER
   sToolTipSet (GetDlgItem(hFManager, %FMANAGER_INIFMLCmdWidth), " Enter desired display width in characters of the FM Line command area. ")
   CONTROL POST hFManager, %FMANAGER_INIFMLCmdWidth, %EM_LIMITTEXT, 2, 0

   CONTROL ADD LABEL, hFManager, %FMANAGER_INIFMNoteWidth_TEXT, "Width of Note column area (5 to 20)", 4, 164, 150, 12
   CONTROL SET COLOR  hFManager, %FMANAGER_INIFMNoteWidth_TEXT, %WHITE, -2
   CONTROL ADD TEXTBOX,   hFManager, %FMANAGER_INIFMNoteWidth, FORMAT$(ENV.FMNoteWidth), 150, 164, 20, 10, %ES_NUMBER
   sToolTipSet (GetDlgItem(hFManager, %FMANAGER_INIFMNoteWidth), " Enter desired display width in characters of the FM Note column area. ")
   CONTROL POST hFManager, %FMANAGER_INIFMNoteWidth, %EM_LIMITTEXT, 2, 0

   '----- Add the Submit Tab
   TAB INSERT PAGE hOpt, %TAB_CONTROL, 3, 0, "Submit", CALL DlgOptionsCallBack TO hSubmit

   CONTROL ADD LABEL, hSubmit, %SUBMIT_SUBMIT_TEXT, "Prototype command line to be used by SUBMIT", 4, 4, 200, 12
   CONTROL SET COLOR  hSubmit, %SUBMIT_SUBMIT_TEXT, %WHITE, -2
   CONTROL ADD TEXTBOX,   hSubmit, %SUBMIT_SUBMIT, ENV.SubmitCmd, 4, 14, 337, 10, %ES_AUTOHSCROLL
   sToolTipSet (GetDlgItem(hSubmit, %SUBMIT_SUBMIT), " Enter your prototype SUBMIT command string using fixed and/or substitution strings as needed. ")

   CONTROL ADD LABEL, hSubmit, %SUBMIT_INISubmitDir_TEXT, "Working Directory to be used by SUBMIT", 4, 34, 200, 12
   CONTROL SET COLOR  hSubmit, %SUBMIT_INISubmitDir_TEXT, %WHITE, -2
   CONTROL ADD TEXTBOX,   hSubmit, %SUBMIT_INISubmitDir, ENV.SubmitDir, 4, 44, 337, 10, %ES_AUTOHSCROLL
   sToolTipSet (GetDlgItem(hSubmit, %SUBMIT_INISubmitDir), " Enter the working Directory where SUBMIT can place temporary work and result files. ")

   CONTROL ADD LABEL, hSubmit, %SUBMIT_INICMDFlags_TEXT, "CMD.EXE flags used by SPFLite CMD command", 4, 64, 200, 12
   CONTROL SET COLOR  hSubmit, %SUBMIT_INICMDFlags_TEXT, %WHITE, -2
   CONTROL ADD TEXTBOX,   hSubmit, %SUBMIT_INICMDFlags, ENV.CmdFlags, 4, 74, 337, 10, %ES_AUTOHSCROLL
   sToolTipSet (GetDlgItem(hSubmit, %SUBMIT_INICMDFlags), " Enter the CMD.EXE command line flags to be used by the SPFLite CMD command. ")

   CONTROL ADD LABEL, hSubmit, %SUBMIT_INIRUNFlags_TEXT, "CMD.EXE flags used by SPFLite RUN command", 4, 94, 200, 12
   CONTROL SET COLOR  hSubmit, %SUBMIT_INIRUNFlags_TEXT, %WHITE, -2
   CONTROL ADD TEXTBOX,   hSubmit, %SUBMIT_INIRUNFlags, ENV.RunFlags, 4, 104, 337, 10, %ES_AUTOHSCROLL
   sToolTipSet (GetDlgItem(hSubmit, %SUBMIT_INIRUNFlags), " Enter the CMD.EXE command line flags to be used by the SPFLite RUN command. ")

   CONTROL ADD LABEL, hSubmit, %SUBMIT_INIInclKey_TEXT, "Trigger key for the SUBMIT Include statement", 4, 124, 200, 12
   CONTROL SET COLOR  hSubmit, %SUBMIT_INIInclKey_TEXT, %WHITE, -2
   CONTROL ADD TEXTBOX,   hSubmit, %SUBMIT_INIInclKey, ENV.SubmitInclKey, 4, 134, 337, 10, %ES_AUTOHSCROLL
   sToolTipSet (GetDlgItem(hSubmit, %SUBMIT_INIInclKey), " Enter the special key that identifies the SUBMIT Include statement. ")

   CONTROL ADD LABEL, hSubmit, %SUBMIT_INIInclCol_TEXT, "Column number where above Include Key must appear", 4, 154, 200, 12
   CONTROL SET COLOR  hSubmit, %SUBMIT_INIInclCol_TEXT, %WHITE, -2
   CONTROL ADD TEXTBOX,   hSubmit, %SUBMIT_INIInclCol, FORMAT$(ENV.SubmitInclCol), 4, 164, 337, 10, %ES_AUTOHSCROLL
   sToolTipSet (GetDlgItem(hSubmit, %SUBMIT_INIInclCol), " Enter the specific column number where the Include Key must appear. ")

   '----- Add the Screen Tab
   TAB INSERT PAGE hOpt, %TAB_CONTROL, 4, 0, "Screen", CALL DlgOptionsCallBack TO hScreen

   sAddCheck(hScreen, %SCREEN_INIHRuler, ENV.HRuler, 184, 1, 150, "Horizontal cursor ruler?")
   sToolTipSet (GetDlgItem(hScreen, %SCREEN_INIHRuler), " Select to display a horizontal line at the cursor row. ")
   sAddCheck(hScreen, %SCREEN_INIVRuler, ENV.VRuler, 184, 13, 150, "Vertical cursor ruler?")
   sToolTipSet (GetDlgItem(hScreen, %SCREEN_INIVRuler), " Select to display a vertical line at the cursor row. ")
   sAddCheck(hScreen, %SCREEN_INIBanding, ENV.Banding, 184, 25, 150, "Alternating Background screen colors?")
   sToolTipSet (GetDlgItem(hScreen, %SCREEN_INIBanding), " Select to display alternating backgrond colors every 3 lines. ")
   sAddCheck(hScreen, %SCREEN_INICMDHILITE, ENV.CmdHiLite, 184, 37, 150, "Hilight Command line keywords?")
   sToolTipSet (GetDlgItem(hScreen, %SCREEN_INICMDHILITE), " Select to have Command line keywords highlighted. ")

   CONTROL ADD TEXTBOX, hScreen, %SCREEN_INIPFKShow, FORMAT$(ENV.PFKShow), 184, 49, 10, 10
   sToolTipSet (GetDlgItem(hScreen, %SCREEN_INIPFKShow), " Enter the # lines at bottom of screen reserved for Keyboard Help.")
   CONTROL ADD LABEL, hScreen, %SCREEN_INIPFKShow_TEXT, "# K-Board help lines to show?", 198, 50, 100, 10
   CONTROL SET COLOR  hScreen, %SCREEN_INIPFKShow_TEXT, %WHITE, -2

   CONTROL ADD TEXTBOX, hScreen, %SCREEN_INILNOSIZE, FORMAT$(ENV.LinNoSize), 184, 61, 10, 10
   sToolTipSet (GetDlgItem(hScreen, %SCREEN_INILNOSIZE), " Enter the display width of the Line numbers on the screen.")
   CONTROL ADD LABEL, hScreen, %SCREEN_INILNOSIZE_TEXT, "Width of Line Numbers (5-8)", 198, 62, 100, 10
   CONTROL SET COLOR  hScreen, %SCREEN_INILNOSIZE_TEXT, %WHITE, -2

   sAddCheck(hScreen, %SCREEN_INIVERTINSCURS, ENV.VertInsCurs, 4, 1, 120, "      Vertical cursor in Ins. Mode.")
   sToolTipSet (GetDlgItem(hScreen, %SCREEN_INIVERTINSCURS), " Use a vertical cursor rather than a blob when in Inset Mode. ")

   CONTROL ADD TEXTBOX,   hScreen, %SCREEN_CURSHGHTNORM, FORMAT$(ENV.CursNormal), 4, 13, 20, 10, %ES_NUMBER
   sToolTipSet (GetDlgItem(hScreen, %SCREEN_CURSHGHTNORM), " Specifies % Height (0-100) of a Normal cursor vs a Full Blob. ")
   CONTROL ADD LABEL, hScreen, %SCREEN_CURSHGHTNORM_TEXT, "% Height Normal Cursor.", 30, 13, 150, 10
   CONTROL SET COLOR  hScreen, %SCREEN_CURSHGHTNORM_TEXT, %WHITE, -2

   CONTROL ADD TEXTBOX,   hScreen, %SCREEN_CURSHGHTINS, FORMAT$(ENV.CursInsert), 4, 25, 20, 10, %ES_NUMBER
   sToolTipSet (GetDlgItem(hScreen, %SCREEN_CURSHGHTINS), " Specifies % Height (20-100) of an Insert cursor vs a Full Blob. ")
   CONTROL ADD LABEL, hScreen, %SCREEN_CURSHGHTINS_TEXT, "% Height Insert Cursor.", 30, 25, 150, 10
   CONTROL SET COLOR  hScreen, %SCREEN_CURSHGHTINS_TEXT, %WHITE, -2

   CONTROL ADD TEXTBOX,   hScreen, %SCREEN_INIFontName, ENV.FontName, 4, 37, 90, 10
   sToolTipSet (GetDlgItem(hScreen, %SCREEN_INIFontName), " Enter the name of the Font to use for the Edit text display. ")
   CONTROL ADD LABEL, hScreen, %SCREEN_INIFontName_TEXT, "Font Name", 100, 37, 50, 10
   CONTROL SET COLOR  hScreen, %SCREEN_INIFontName_TEXT, %WHITE, -2

   CONTROL ADD TEXTBOX,   hScreen, %SCREEN_INIFontPitch, FORMAT$(ENV.FontPitch), 4, 49, 20, 10
   sToolTipSet (GetDlgItem(hScreen, %SCREEN_INIFontPitch), " Enter the size of the Font (in points). ")
   CONTROL ADD LABEL, hScreen, %SCREEN_INIFontPitch_TEXT, "Font Pitch", 100, 49, 50, 10
   CONTROL SET COLOR  hScreen, %SCREEN_INIFontPitch_TEXT, %WHITE, -2

   CONTROL ADD BUTTON, hScreen, %SCREEN_FONT_CHOOSE, "Choose", 4, 61, 45, 10', , %WS_BORDER
   CONTROL SET COLOR  hScreen, %SCREEN_FONT_CHOOSE, %BLUE, %WHITE
   sToolTipSet (GetDlgItem(hScreen, %SCREEN_FONT_CHOOSE), " Click this button for a Font Selection Dialog which allows you to choose both Font and Pitch. ")
   CONTROL ADD LABEL, hScreen, %SCREEN_FONT_CHOOSE_TEXT, "(or Choose Font here)", 100, 61, 80, 10
   CONTROL SET COLOR  hScreen, %SCREEN_FONT_CHOOSE_TEXT, %WHITE, -2

   CONTROL ADD GRAPHIC, hScreen, %SCREEN_INIcMarkLine, "", 4, 73, 12, 12, %WS_BORDER OR %SS_NOTIFY
   GRAPHIC ATTACH hScreen, %SCREEN_INIcMarkLine, REDRAW
   GRAPHIC BOX (1, 1) - (12, 12), 20, %BLACK, ENV.cMarkLine, 0
   sToolTipSet (GetDlgItem(hScreen, %SCREEN_INIcMarkLine), " Press to choose the color for the Column Mark Lines. ")
   CONTROL ADD LABEL, hScreen, %SCREEN_INIcMarkLine_TEXT, "Column Marker Line Color", 30, 75, 150, 12
   CONTROL SET COLOR  hScreen, %SCREEN_INIcMarkLine_TEXT, %WHITE, -2

   '----- Now add the Screen component color choices
   CONTROL ADD LABEL, hScreen, %SCHEME_HEADING1, "FG BG1 BG2", 202 , 78, 50, 12
   CONTROL SET COLOR  hScreen, %SCHEME_HEADING1, %YELLOW, -2
   i = 0
   FOR j = 16 TO 25
      CONTROL ADD LABEL, hScreen, %SCREEN_16_TEXT + (i * 4), ScrLabels(j), 4 , 88 + (i * 11), 200, 10, %SS_CENTERIMAGE
      CONTROL SET COLOR  hScreen, %SCREEN_16_TEXT + (i * 4), ENV.GetClr(j, 1), ENV.GetClr(j, 2)
      CONTROL SET FONT   hScreen, %SCREEN_16_TEXT + (i * 4), hScrFont
      IF j = 23 THEN _
         CONTROL SET FONT   hScreen, %SCREEN_16_TEXT + (i * 4), hSBFont

      CONTROL ADD GRAPHIC, hScreen, %SCREEN_16_FG + (i * 4), "", 201, 87 + (i * 11), 12, 12, %WS_BORDER OR %SS_NOTIFY
      CONTROL SET USER   hScreen, %SCREEN_16_FG + (i * 4), 1, j
      CONTROL SET USER   hScreen, %SCREEN_16_FG + (i * 4), 2, 1
      GRAPHIC ATTACH hScreen, %SCREEN_16_FG + (i * 4), REDRAW
      GRAPHIC BOX (1, 1) - (12, 12), 20, %BLACK, ENV.GetClr(j, 1), 0
      sToolTipSet (GetDlgItem(hScreen, %SCREEN_16_FG + (i * 4)), " Click to choose " + ScrLabels(j) + " ForeGround color. ")

      CONTROL ADD GRAPHIC, hScreen, %SCREEN_16_FG + 1 + (i * 4), "", 214, 87 + (i * 11), 12, 12, %WS_BORDER OR %SS_NOTIFY
      CONTROL SET USER   hScreen, %SCREEN_16_FG + 1 + (i * 4), 1, j
      CONTROL SET USER   hScreen, %SCREEN_16_FG + 1 + (i * 4), 2, 2
      GRAPHIC ATTACH hScreen, %SCREEN_16_FG + 1 + (i * 4), REDRAW
      GRAPHIC BOX (1, 1) - (12, 12), 20, %BLACK, ENV.GetClr(j, 2), 0
      sToolTipSet (GetDlgItem(hScreen, %SCREEN_16_FG + 1 + (i * 4)), " Click to choose " + ScrLabels(j) + " BackGround 1 color. ")

      CONTROL ADD GRAPHIC, hScreen, %SCREEN_16_FG + 2 + (i * 4), "", 229, 87 + (i * 11), 12, 12, %WS_BORDER OR %SS_NOTIFY
      CONTROL SET USER   hScreen, %SCREEN_16_FG + 2 + (i * 4), 1, j
      CONTROL SET USER   hScreen, %SCREEN_16_FG + 2 + (i * 4), 2, 3
      GRAPHIC ATTACH hScreen, %SCREEN_16_FG + 2 + (i * 4), REDRAW
      GRAPHIC BOX (1, 1) - (12, 12), 20, %BLACK, ENV.GetClr(j, 3), 0
      sToolTipSet (GetDlgItem(hScreen, %SCREEN_16_FG + 2 + (i * 4)), " Click to choose " + ScrLabels(j) + " BackGround 2 color. ")
      INCR i
   NEXT i

   '----- Add the Keyboard Tab
   TAB INSERT PAGE hOpt, %TAB_CONTROL, 5, 0, "KBD", CALL DlgOptionsCallBack TO hKeyboard
   sAddCheck(hKeyboard, %KBOARD_KBSCROLLH, ENV.KBHorScroll, 4, 4, 150, "Scroll left/right with Arrow Keys?")
   sToolTipSet (GetDlgItem(hKeyBoard, %KBOARD_KBSCROLLH), _
                    " If selected, scrolling will occur at the edge of the text area, otherwise the cursor will 'wrap' to the opposite screen edge. ")

   sAddCheck(hKeyboard, %KBOARD_KBSCROLLV, ENV.KBVerScroll, 4, 18, 150, "Scroll up/down with Arrow Keys?")
   sToolTipSet (GetDlgItem(hKeyBoard, %KBOARD_KBSCROLLV), _
                    " If selected, scrolling will occur at the edge of the text area, otherwise the cursor will 'wrap' to the opposite screen edge. ")

   sAddCheck(hKeyBoard, %KBOARD_INIInsMode, ENV.InsMode, 4, 32, 150, "Keyboard starts in INSERT mode?")
   sToolTipSet (GetDlgItem(hKeyBoard, %KBOARD_INIInsMode), " Default Cursor mode to be used. ")

   sAddCheck(hKeyBoard, %KBOARD_INIInsReset, ENV.InsReset, 4, 46, 150, "Reset INSERT mode on Attention?")
   sToolTipSet (GetDlgItem(hKeyBoard, %KBOARD_INIInsReset), " Select if Cursor should return to the Default Cursor mode following a Keyboard Attention key (Enter, PFK, etc.).")

   '----- Add the Status Bar Tab
   TAB INSERT PAGE hOpt, %TAB_CONTROL, 6, 0, "Status", CALL DlgOptionsCallBack TO hSBar

   CONTROL ADD LABEL, hSBar, %SBAR_Help1, "You may customize the contents of the Status Bar here.", 4, 5, 200, 10
   CONTROL SET COLOR  hSBar, %SBAR_Help1, %YELLOW, -2
   CONTROL ADD LABEL, hSBar, %SBAR_Help2, "Simply choose below which fields you want to appear.", 4, 20, 225, 10
   CONTROL SET COLOR  hSBar, %SBAR_Help2, %YELLOW, -2
   CONTROL ADD LABEL, hSBar, %SBAR_INISBOrder_TEXT, "Enter your desired Status Bar fields from Left to Right", 4, 45, 200, 12
   CONTROL SET COLOR  hSBar, %SBAR_INISBOrder_TEXT, %WHITE, -2
   CONTROL ADD TEXTBOX,   hSBar, %SBAR_INISBOrder, ENV.SBLayout, 4, 55, 337, 25, %ES_UPPERCASE OR %ES_MULTILINE
   sToolTipSet (GetDlgItem(hSBar, %SBAR_INISBOrder), " Enter desired Status Bar fields in order from left to right. ")
   CONTROL ADD LABEL, hSBar, %SBAR_Help3, "Valid field names to choose from are:", 4, 80, 200, 10
   CONTROL SET COLOR  hSBar, %SBAR_Help3, %WHITE, -2
   CONTROL ADD LABEL, hSBar, %SBAR_Help4, "  BNDS, CAPS, CASE, CHANGE, COLS, EOL, INSOVR, LINES,", 4, 90, 225, 10
   CONTROL SET COLOR  hSBar, %SBAR_Help4, %WHITE, -2
   CONTROL ADD LABEL, hSBar, %SBAR_Help5, "  LINNO, MISC, MODE, SELECT, SOURCE and STATE ", 4, 100, 225, 10
   CONTROL SET COLOR  hSBar, %SBAR_Help5, %WHITE, -2
   CONTROL ADD LABEL, hSBar, %SBAR_Help6, "  Refer to Help => Working with SPFLite => Status Bar Contents for more info.", 4, 125, 337, 30
   CONTROL SET COLOR  hSBar, %SBAR_Help6, %WHITE, -2

   '----- Add the Scheme Tab
   TAB INSERT PAGE hOpt, %TAB_CONTROL, 7, 0, "Schemes", CALL DlgOptionsCallBack TO hScheme
   CONTROL ADD LABEL, hScheme, %SCHEME_HEADING1, "Scheme                                                                            FG   BG1 BG2", 2 , 2, 250, 12
   CONTROL SET COLOR  hScheme, %SCHEME_HEADING1, %YELLOW, -2
   FOR i = 0 TO 15
      IF i = 0 THEN                                               ' Handle 1st two specially
         CONTROL ADD LABEL, hScheme, %SCHEME_00_TEXT + (i * 4), "Normal Text Lo Intensity",  4, 177, 185, 10, %SS_CENTERIMAGE
         CONTROL SET COLOR  hScheme, %SCHEME_00_TEXT + (i * 4), ENV.GetClr(i, 1), ENV.GetClr(i, 2)
         CONTROL SET FONT   hScheme, %SCHEME_00_TEXT + (i * 4), hScrFont
         CONTROL ADD GRAPHIC, hScheme, %SCHEME_00_FG + (i * 4), "", 198, 176, 12, 12, %WS_BORDER OR %SS_NOTIFY
         CONTROL SET USER   hScheme, %SCHEME_00_FG + (i * 4), 1, i
         CONTROL SET USER   hScheme, %SCHEME_00_FG + (i * 4), 2, 1
         GRAPHIC ATTACH hScheme, %SCHEME_00_FG + (i * 4), REDRAW
         GRAPHIC BOX (1, 1) - (12, 12), 20, %BLACK, ENV.GetClr(i, 1), 0
         sToolTipSet (GetDlgItem(hScheme, %SCHEME_00_FG + (i * 4)), " Click to choose SCHEME ForeGround color. ")

         CONTROL ADD GRAPHIC, hScheme, %SCHEME_00_FG + 1 + (i * 4), "", 213, 177, 12, 12, %WS_BORDER OR %SS_NOTIFY
         CONTROL SET USER   hScheme, %SCHEME_00_FG + 1 + (i * 4), 1, i
         CONTROL SET USER   hScheme, %SCHEME_00_FG + 1 + (i * 4), 2, 2
         GRAPHIC ATTACH hScheme, %SCHEME_00_FG + 1 + (i * 4), REDRAW
         GRAPHIC BOX (1, 1) - (12, 12), 20, %BLACK, ENV.GetClr(i, 2), 0
         sToolTipSet (GetDlgItem(hScheme, %SCHEME_00_FG + 1 + (i * 4)), " Click to choose SCHEME BackGround 1 color. ")

         CONTROL ADD GRAPHIC, hScheme, %SCHEME_00_FG + 2 + (i * 4), "", 229, 177, 12, 12, %WS_BORDER OR %SS_NOTIFY
         CONTROL SET USER   hScheme, %SCHEME_00_FG + 2 + (i * 4), 1, i
         CONTROL SET USER   hScheme, %SCHEME_00_FG + 2 + (i * 4), 2, 3
         GRAPHIC ATTACH hScheme, %SCHEME_00_FG + 2 + (i * 4), REDRAW
         GRAPHIC BOX (1, 1) - (12, 12), 20, %BLACK, ENV.GetClr(i, 3), 0
         sToolTipSet (GetDlgItem(hScheme, %SCHEME_00_FG + 2 + (i * 4)), " Click to choose SCHEME BackGround 2 color. ")
      ELSEIF i = 1 THEN                                           '
         CONTROL ADD LABEL, hScheme, %SCHEME_00_TEXT + (i * 4), "Normal Text Hi Intensity", 4, 188, 185, 10, %SS_CENTERIMAGE
         CONTROL SET COLOR  hScheme, %SCHEME_00_TEXT + (i * 4), ENV.GetClr(i, 1), ENV.GetClr(i, 2)
         CONTROL SET FONT   hScheme, %SCHEME_00_TEXT + (i * 4), hScrFont
         CONTROL ADD GRAPHIC, hScheme, %SCHEME_00_FG + (i * 4), "", 198, 187, 12, 12, %WS_BORDER OR %SS_NOTIFY
         CONTROL SET USER   hScheme, %SCHEME_00_FG + (i * 4), 1, i
         CONTROL SET USER   hScheme, %SCHEME_00_FG + (i * 4), 2, 1
         GRAPHIC ATTACH hScheme, %SCHEME_00_FG + (i * 4), REDRAW
         GRAPHIC BOX (1, 1) - (12, 12), 20, %BLACK, ENV.GetClr(i, 1), 0
         sToolTipSet (GetDlgItem(hScheme, %SCHEME_00_FG + (i * 4)), " Click to choose SCHEME ForeGround color. ")

         CONTROL ADD GRAPHIC, hScheme, %SCHEME_00_FG + 1 + (i * 4), "", 213, 187, 12, 12, %WS_BORDER OR %SS_NOTIFY
         CONTROL SET USER   hScheme, %SCHEME_00_FG + 1 + (i * 4), 1, i
         CONTROL SET USER   hScheme, %SCHEME_00_FG + 1 + (i * 4), 2, 2
         GRAPHIC ATTACH hScheme, %SCHEME_00_FG + 1 + (i * 4), REDRAW
         GRAPHIC BOX (1, 1) - (12, 12), 20, %BLACK, ENV.GetClr(i, 2), 0
         sToolTipSet (GetDlgItem(hScheme, %SCHEME_00_FG + 1 + (i * 4)), " Click to choose SCHEME BackGround 1 color. ")

         CONTROL ADD GRAPHIC, hScheme, %SCHEME_00_FG + 2 + (i * 4), "", 229, 187, 12, 12, %WS_BORDER OR %SS_NOTIFY
         CONTROL SET USER   hScheme, %SCHEME_00_FG + 2 + (i * 4), 1, i
         CONTROL SET USER   hScheme, %SCHEME_00_FG + 2 + (i * 4), 2, 3
         GRAPHIC ATTACH hScheme, %SCHEME_00_FG + 2 + (i * 4), REDRAW
         GRAPHIC BOX (1, 1) - (12, 12), 20, %BLACK, ENV.GetClr(i, 3), 0
         sToolTipSet (GetDlgItem(hScheme, %SCHEME_00_FG + 2 + (i * 4)), " Click to choose SCHEME BackGround 2 color. ")
      ELSE                                                        '
         CONTROL ADD LABEL, hScheme, %SCHEME_00_TEXT + (i * 4), "AUTO Colorization Scheme " + FORMAT$(i - 1), 4, 23 + ((i-2) * 11), 185, 10, %SS_CENTERIMAGE
         CONTROL SET COLOR  hScheme, %SCHEME_00_TEXT + (i * 4), ENV.GetClr(i, 1), ENV.GetClr(i, 2)
         CONTROL SET FONT   hScheme, %SCHEME_00_TEXT + (i * 4), hScrFont
         CONTROL ADD GRAPHIC, hScheme, %SCHEME_00_FG + (i * 4), "", 198, 22 + ((i-2) * 11), 12, 12, %WS_BORDER OR %SS_NOTIFY
         CONTROL SET USER   hScheme, %SCHEME_00_FG + (i * 4), 1, i
         CONTROL SET USER   hScheme, %SCHEME_00_FG + (i * 4), 2, 1
         GRAPHIC ATTACH hScheme, %SCHEME_00_FG + (i * 4), REDRAW
         GRAPHIC BOX (1, 1) - (12, 12), 20, %BLACK, ENV.GetClr(i, 1), 0
         sToolTipSet (GetDlgItem(hScheme, %SCHEME_00_FG + (i * 4)), " Click to choose SCHEME ForeGround color. ")

         CONTROL ADD GRAPHIC, hScheme, %SCHEME_00_FG + 1 + (i * 4), "", 213, 22 + ((i-2) * 11), 12, 12, %WS_BORDER OR %SS_NOTIFY
         CONTROL SET USER   hScheme, %SCHEME_00_FG + 1 + (i * 4), 1, i
         CONTROL SET USER   hScheme, %SCHEME_00_FG + 1 + (i * 4), 2, 2
         GRAPHIC ATTACH hScheme, %SCHEME_00_FG + 1 + (i * 4), REDRAW
         GRAPHIC BOX (1, 1) - (12, 12), 20, %BLACK, ENV.GetClr(i, 2), 0
         sToolTipSet (GetDlgItem(hScheme, %SCHEME_00_FG + 1 + (i * 4)), " Click to choose SCHEME BackGround 1 color. ")

         CONTROL ADD GRAPHIC, hScheme, %SCHEME_00_FG + 2 + (i * 4), "", 229, 22 + ((i-2) * 11), 12, 12, %WS_BORDER OR %SS_NOTIFY
         CONTROL SET USER   hScheme, %SCHEME_00_FG + 2 + (i * 4), 1, i
         CONTROL SET USER   hScheme, %SCHEME_00_FG + 2 + (i * 4), 2, 3
         GRAPHIC ATTACH hScheme, %SCHEME_00_FG + 2 + (i * 4), REDRAW
         GRAPHIC BOX (1, 1) - (12, 12), 20, %BLACK, ENV.GetClr(i, 3), 0
         sToolTipSet (GetDlgItem(hScheme, %SCHEME_00_FG + 2 + (i * 4)), " Click to choose SCHEME BackGround 2 color. ")
      END IF                                                      '

   NEXT i

   '----- Add the HiLites Tab
   TAB INSERT PAGE hOpt, %TAB_CONTROL, 8, 0, "HiLites", CALL DlgOptionsCallBack TO hHiLites
   CONTROL ADD LABEL, hHiLites, %HILITES_HEADING1, "        HiLite Color Name         FG   BG1 BG2", 2 , 2, 250, 12
   CONTROL SET COLOR  hHiLites, %HILITES_HEADING1, %YELLOW, -2
   CONTROL ADD LABEL, hHiLites, %HILITES_HEADING2, "The letters in parens are used as color-setting codes in SPFLite macros.", 2 , 195, 250, 12
   CONTROL SET COLOR  hHiLites, %HILITES_HEADING2, %YELLOW, -2
   FOR i = 0 TO 14
      CONTROL ADD LABEL, hHiLites, %HILITES_01_TEXT + (i * 1), "(" + MID$(nHiLitesChrs, (i + 1), 1) + ") " + nHiLites(i + 1), 17, 24 + (i * 11), 70, 10
      CONTROL SET COLOR  hHiLites, %HILITES_01_TEXT + (i * 1), ENV.GetClr(i + 32, 1), ENV.GetClr(i + 32, 2)
      CONTROL SET FONT   hHiLites, %HILITES_01_TEXT + (i * 1), hScrFont

      CONTROL ADD GRAPHIC, hHiLites, %HILITES_01_FG + (i * 3), "", 98, 22 + (i * 11), 12, 12, %WS_BORDER OR %SS_NOTIFY
      CONTROL SET USER   hHiLites, %HILITES_01_FG + (i * 3), 1, i + 32 ' Pass HiLite Scheme number
      CONTROL SET USER   hHiLites, %HILITES_01_FG + (i * 3), 2, 1      ' Pass FG/BG1/BG2
      CONTROL SET USER   hHiLites, %HILITES_01_FG + (i * 3), 3, %HILITES_01_TEXT + (i * 1) ' Pass back Text control ID
      GRAPHIC ATTACH hHiLites, %HILITES_01_FG + (i * 3), REDRAW
      GRAPHIC BOX (1, 1) - (12, 12), 20, %BLACK, ENV.GetClr(i + 32, 1), 0
      sToolTipSet (GetDlgItem(hHiLites, %HILITES_01_FG + (i * 3)), " Click to choose HILITE ForeGround color. ")

      CONTROL ADD GRAPHIC, hHiLites, %HILITES_01_FG + 1 + (i * 3), "", 113, 22 + (i * 11), 12, 12, %WS_BORDER OR %SS_NOTIFY
      CONTROL SET USER   hHiLites, %HILITES_01_FG + 1 + (i * 3), 1, i + 32 ' Pass HiLite Scheme number
      CONTROL SET USER   hHiLites, %HILITES_01_FG + 1 + (i * 3), 2, 2      ' Pass FG/BG1/BG2
      CONTROL SET USER   hHiLites, %HILITES_01_FG + 1 + (i * 3), 3, %HILITES_01_TEXT + (i * 1) ' Pass back Text control ID
      GRAPHIC ATTACH hHiLites, %HILITES_01_FG + 1 + (i * 3), REDRAW
      GRAPHIC BOX (1, 1) - (12, 12), 20, %BLACK, ENV.GetClr(i + 32, 2), 0
      sToolTipSet (GetDlgItem(hHiLites, %HILITES_01_FG + 1 + (i * 3)), " Click to choose HILITE BackGround 1 color. ")

      CONTROL ADD GRAPHIC, hHiLites, %HILITES_01_FG + 2 + (i * 3), "", 129, 22 + (i * 11), 12, 12, %WS_BORDER OR %SS_NOTIFY
      CONTROL SET USER   hHiLites, %HILITES_01_FG + 2 + (i * 3), 1, i + 32 ' Pass HiLite Scheme number
      CONTROL SET USER   hHiLites, %HILITES_01_FG + 2 + (i * 3), 2, 3      ' Pass FG/BG1/BG2
      CONTROL SET USER   hHiLites, %HILITES_01_FG + 2 + (i * 3), 3, %HILITES_01_TEXT + (i * 1) ' Pass back Text control ID
      GRAPHIC ATTACH hHiLites, %HILITES_01_FG + 2 + (i * 3), REDRAW
      GRAPHIC BOX (1, 1) - (12, 12), 20, %BLACK, ENV.GetClr(i + 32, 3), 0
      sToolTipSet (GetDlgItem(hHiLites, %HILITES_01_FG + 2 + (i * 3)), " Click to choose HILITE BackGround 2 color. ")

   NEXT i

   '----- DISPLAY the Options DIALOG BOX
   TAB SELECT hOpt, %TAB_CONTROL, StartPg                         '
   GRAPHIC REDRAW                                                 '
   IF ISFALSE ENV.FirstTime THEN sPopReady                        ' Ready for pop-up
   DIALOG SHOW MODAL hOpt CALL DlgOptionsCallBack                 '
   IF ISFALSE ENV.FirstTime THEN sPopReset                        ' Reset popup state
   MExit
END SUB

SUB      DispPrint()
'---------- Build and display the PRINT FORMAT dialog
LOCAL i, j AS LONG
   MEntry
   '----- Get printer attached
   IF ISFALSE sOpenPrinter("") THEN                               ' Get printer open
      SDoMsgBox("Can't attach default printer, current printer info may be incorrect.", %MB_OK OR %MB_USERICON, "Print Setup")  '
   END IF                                                         '
   XPRINT CANCEL                                                  ' Cancel
   XPRINT CLOSE                                                   ' and close it

   '----- Build the Dialog
   DIALOG FONT DEFAULT "Tahoma", 10 / gFontScale, 0, 0
   DIALOG NEW PIXELS, hWnd, "SPFLite Print Setup", , , 700, 430, %WS_POPUP OR %WS_BORDER OR %WS_DLGFRAME OR %WS_CAPTION OR _
                            %WS_CLIPSIBLINGS OR %WS_VISIBLE OR %DS_MODALFRAME OR %DS_3DLOOK OR %DS_NOFAILCREATE OR _
                            %DS_SETFONT, %WS_EX_CONTROLPARENT OR %WS_EX_LEFT OR %WS_EX_LTRREADING OR %WS_EX_RIGHTSCROLLBAR, TO hPrt
   DIALOG SET COLOR hPrt, %RGB_GAINSBORO, %RGB_GAINSBORO

   '----- Add the DONE and CANCEL buttons
   CONTROL ADD BUTTON,  hPrt, %PRINT_DONE, "Done", 600, 400, 75, 24, %WS_BORDER
   sToolTipSet (GetDlgItem(hPrt, %PRINT_DONE), " Press Done once you have made all desired changes. ")

   '----- Add the printer information
   CONTROL ADD LABEL,    hPrt, %PRINT_PTEXT1, "Current Printer", 4, 10, 150, 16
   CONTROL SET COLOR     hPrt, %PRINT_PTEXT1, %RGB_DARKBLUE, -2
   CONTROL ADD LABEL,    hPrt, %PRINT_INIPrtName_TEXT, "Printer: " + ENV.PrtName, 10, 35, 250, 16
   CONTROL SET COLOR     hPrt, %PRINT_INIPrtName_TEXT, %BLACK, %RGB_GAINSBORO
   CONTROL SET FONT      hPrt, %PRINT_INIPrtName_TEXT, hFixedFont
   CONTROL ADD LABEL,    hPrt, %PRINT_INIPrtOrient_TEXT, "Orientation: " + IIF$(ENV.PrtOrient = 1, "Portrait", "Landscape"), 10, 60, 250, 16
   CONTROL SET COLOR     hPrt, %PRINT_INIPrtOrient_TEXT, %BLACK, %RGB_GAINSBORO
   CONTROL SET FONT      hPrt, %PRINT_INIPrtOrient_TEXT, hFixedFont
   CONTROL ADD LABEL,    hPrt, %PRINT_INIPrtPaper_TEXT, "Paper: " + gPrtPaper, 10, 85, 250, 16
   CONTROL SET COLOR     hPrt, %PRINT_INIPrtPaper_TEXT, %BLACK, %RGB_GAINSBORO
   CONTROL SET FONT      hPrt, %PRINT_INIPrtPaper_TEXT, hFixedFont
   CONTROL ADD BUTTON,   hPrt, %PRINT_PCHOOSE, "Choose Printer", 40, 120, 100, 24, %WS_BORDER
   sToolTipSet (GetDlgItem(hPrt, %PRINT_PCHOOSE), " Press to change your printer choice. ")

   CONTROL ADD LINE, hPrt, %PRINT_VLine, "", 350, 2, 3, 148

   '----- Add the Font data
   CONTROL ADD LABEL,    hPrt, %PRINT_FTEXT1, "Current Font", 354, 10, 150, 16
   CONTROL SET COLOR     hPrt, %PRINT_FTEXT1, %RGB_DARKBLUE, %RGB_GAINSBORO
   CONTROL ADD LABEL,    hPrt, %PRINT_FNAME_TEXT, "Fontname: " + ENV.PrtFontName, 360, 35, 250, 16
   CONTROL SET COLOR     hPrt, %PRINT_FNAME_TEXT, %BLACK, %RGB_GAINSBORO
   CONTROL SET FONT      hPrt, %PRINT_FNAME_TEXT, hFixedFont
   CONTROL ADD LABEL,    hPrt, %PRINT_FPOINT_TEXT, "Pitch: " + ENV.PrtFontPitch, 360, 60, 250, 16
   CONTROL SET COLOR     hPrt, %PRINT_FPOINT_TEXT, %BLACK, %RGB_GAINSBORO
   CONTROL SET FONT      hPrt, %PRINT_FPOINT_TEXT, hFixedFont
   CONTROL ADD LABEL,    hPrt, %PRINT_FBOLD_TEXT, "Bold: " + IIF$(ENV.PrtFontStyle = "0", "Off", "On"), 360, 85, 250, 16
   CONTROL SET COLOR     hPrt, %PRINT_FBOLD_TEXT, %BLACK, %RGB_GAINSBORO
   CONTROL SET FONT      hPrt, %PRINT_FBOLD_TEXT, hFixedFont
   CONTROL ADD BUTTON,   hPrt, %PRINT_FCHOOSE, "Choose Font", 390, 120, 100, 24, %WS_BORDER
   sToolTipSet (GetDlgItem(hPrt, %PRINT_FCHOOSE), " Press to change your Font choice. ")

   CONTROL ADD LINE, hPrt, %PRINT_HLine, "", 2, 148, 698, 3

   '----- Add margin values
   CONTROL ADD LABEL,    hPrt, %PRINT_MTEXT1, "Page margins", 4, 152, 75, 16
   CONTROL SET COLOR     hPrt, %PRINT_MTEXT1, %RGB_DARKBLUE, %RGB_GAINSBORO
   CONTROL ADD CHECKBOX, hPrt, %PRINT_METRIC, "Use Metric (mm) instead of inches? ", 85, 152, 350, 16, %BS_LEFT OR %BS_VCENTER
   CONTROL SET CHECK     hPrt, %PRINT_METRIC, ENV.PrtMetric
   CONTROL SET COLOR     hPrt, %PRINT_METRIC, %RGB_DARKBLUE, %RGB_GAINSBORO
   sToolTipSet (GetDlgItem(hPrt, %PRINT_METRIC), " Select if you prefer margins in millimeters, else settings are in inches? ")

   CONTROL ADD TEXTBOX,  hPrt, %PRINT_TM, FORMAT$(IIF(ENV.PrtMetric, ENV.PrtTMargin * 25.4, ENV.PrtTMargin), "0.0"), 354, 175, 50, 20
   CONTROL SET FONT      hPrt, %PRINT_TM, hFixedFont
   sToolTipSet (GetDlgItem(hPrt, %PRINT_TM), " Enter the page top margin space. ")
   CONTROL ADD LABEL,    hPrt, %PRINT_TM_TEXT, "Page top margin " + IIF$(ENV.PrtMetric, "(mm)", "(in)"), 415, 175, 175, 16
   CONTROL SET COLOR     hPrt, %PRINT_TM_TEXT, %BLACK, %RGB_GAINSBORO

   CONTROL ADD TEXTBOX,  hPrt, %PRINT_BM, FORMAT$(IIF(ENV.PrtMetric, ENV.PrtBMargin * 25.4, ENV.PrtBMargin), "0.0"), 4, 175, 50, 20
   CONTROL SET FONT      hPrt, %PRINT_BM, hFixedFont
   sToolTipSet (GetDlgItem(hPrt, %PRINT_BM), " Enter the page bottom margin space. ")
   CONTROL ADD LABEL,    hPrt, %PRINT_BM_TEXT, "Page bottom margin " + IIF$(ENV.PrtMetric, "(mm)", "(in)"), 60, 175, 175, 16
   CONTROL SET COLOR     hPrt, %PRINT_BM_TEXT, %BLACK, %RGB_GAINSBORO

   CONTROL ADD TEXTBOX,  hPrt, %PRINT_RM, FORMAT$(IIF(ENV.PrtMetric, ENV.PrtRMargin * 25.4, ENV.PrtRMargin), "0.0"), 354, 201, 50, 20
   CONTROL SET FONT      hPrt, %PRINT_RM, hFixedFont
   sToolTipSet (GetDlgItem(hPrt, %PRINT_RM * 25.4), " Enter the page right margin space. ")
   CONTROL ADD LABEL,    hPrt, %PRINT_RM_TEXT, "Page right margin " + IIF$(ENV.PrtMetric, "(mm)", "(in)"), 415, 201, 175, 16
   CONTROL SET COLOR     hPrt, %PRINT_RM_TEXT, %BLACK, %RGB_GAINSBORO

   CONTROL ADD TEXTBOX,  hPrt, %PRINT_LM, FORMAT$(IIF(ENV.PrtMetric, ENV.PrtLMargin * 25.4, ENV.PrtLMargin), "0.0"), 4, 201, 50, 20
   CONTROL SET FONT      hPrt, %PRINT_LM, hFixedFont
   sToolTipSet (GetDlgItem(hPrt, %PRINT_LM), " Enter the page left margin space. ")
   CONTROL ADD LABEL,    hPrt, %PRINT_LM_TEXT, "Page left margin " + IIF$(ENV.PrtMetric, "(mm)", "(in)"), 60, 201, 175, 16
   CONTROL SET COLOR     hPrt, %PRINT_LM_TEXT, %BLACK, %RGB_GAINSBORO

   CONTROL ADD CHECKBOX, hPrt, %PRINT_BANDING, "Print band stripes? ", 4, 228, 130, 16, %BS_LEFT OR %BS_VCENTER
   CONTROL SET CHECK     hPrt, %PRINT_BANDING, ENV.PrtBanding
   CONTROL SET COLOR     hPrt, %PRINT_BANDING, %BLACK, %RGB_GAINSBORO
   sToolTipSet (GetDlgItem(hPrt, %PRINT_BANDING), " Should horizontal band stripes be added to the page for readability? ")

   CONTROL ADD CHECKBOX, hPrt, %PRINT_BANDLINES, "Print band lines? ", 140, 228, 130, 16, %BS_LEFT OR %BS_VCENTER
   CONTROL SET CHECK     hPrt, %PRINT_BANDLINES, ENV.PrtBandLines
   CONTROL SET COLOR     hPrt, %PRINT_BANDLINES, %BLACK, %RGB_GAINSBORO
   sToolTipSet (GetDlgItem(hPrt, %PRINT_BANDLINES), " Should horizontal lines be added to the page for readability? ")

   CONTROL ADD GRAPHIC, hPrt, %PRINT_BandingColor, "", 275, 228, 18, 18, %WS_BORDER OR %SS_NOTIFY
   GRAPHIC ATTACH hPrt, %PRINT_BandingColor, REDRAW
   GRAPHIC BOX (1, 1) - (16, 16), 20, %BLACK, ENV.PrtBandColor, 0
   sToolTipSet (GetDlgItem(hPrt, %PRINT_BandingColor), " Press here to choose the Print Band Stripe/Line color. ")
   CONTROL ADD LABEL, hPrt, %PRINT_BandingColor_TEXT, "Print Band Stripe/Line color", 305, 228, 250, 16
   CONTROL SET COLOR  hPrt, %PRINT_BandingColor_TEXT, %BLACK, %RGB_GAINSBORO

   CONTROL ADD CHECKBOX, hPrt, %PRINT_PCOLOR, "Print in Color? ", 475, 228, 130, 16, %BS_LEFT OR %BS_VCENTER
   CONTROL SET CHECK     hPrt, %PRINT_PCOLOR, ENV.PrtPColor
   CONTROL SET COLOR     hPrt, %PRINT_PCOLOR, %BLACK, %RGB_GAINSBORO
   sToolTipSet (GetDlgItem(hPrt, %PRINT_PCOLOR), " Should the Print be in Color?")

   CONTROL ADD LINE, hPrt, %PRINT_HLine + 1, "", 2, 248, 698, 3

   '----- Add page heading and footer stuff
   CONTROL ADD CHECKBOX, hPrt, %PRINT_HEADER, "Print page headers? ", 4, 255, 160, 16, %BS_LEFT OR %BS_VCENTER
   CONTROL SET CHECK     hPrt, %PRINT_HEADER, ENV.PrtHeader
   CONTROL SET COLOR     hPrt, %PRINT_HEADER, %RGB_DARKBLUE, %RGB_GAINSBORO
   sToolTipSet (GetDlgItem(hPrt, %PRINT_HEADER), " Should page headers as described below be printed? ")

   CONTROL ADD LABEL,    hPrt, %PRINT_HEADER_LEFT_TEXT, "Left-hand header", 4, 275, 100, 16
   CONTROL SET COLOR     hPrt, %PRINT_HEADER_LEFT_TEXT, %BLACK, %RGB_GAINSBORO
   CONTROL ADD TEXTBOX,  hPrt, %PRINT_HEADER_LEFT, ENV.PrtHeaderLeft, 4, 295, 225, 20
   CONTROL SET FONT      hPrt, %PRINT_HEADER_LEFT, hFixedFont
   sToolTipSet (GetDlgItem(hPrt, %PRINT_HEADER_LEFT), " Enter the format string for the left side of the header line. ")

   CONTROL ADD LABEL,    hPrt, %PRINT_HEADER_CENTER_TEXT, "Center header", 235, 275, 200, 16
   CONTROL SET COLOR     hPrt, %PRINT_HEADER_CENTER_TEXT, %BLACK, %RGB_GAINSBORO
   CONTROL ADD TEXTBOX,  hPrt, %PRINT_HEADER_CENTER, ENV.PrtHeaderCenter, 235, 295, 225, 20
   CONTROL SET FONT      hPrt, %PRINT_HEADER_CENTER, hFixedFont
   sToolTipSet (GetDlgItem(hPrt, %PRINT_HEADER_CENTER), " Enter the format string for the center of the header line. ")

   CONTROL ADD LABEL,    hPrt, %PRINT_HEADER_RIGHT_TEXT, "Right-hand header", 465, 275, 200, 16
   CONTROL SET COLOR     hPrt, %PRINT_HEADER_RIGHT_TEXT, %BLACK, %RGB_GAINSBORO
   CONTROL ADD TEXTBOX,  hPrt, %PRINT_HEADER_RIGHT, ENV.PrtHeaderRight, 465, 295, 225, 20
   CONTROL SET FONT      hPrt, %PRINT_HEADER_RIGHT, hFixedFont
   sToolTipSet (GetDlgItem(hPrt, %PRINT_HEADER_RIGHT), " Enter the format string for the right hand of the header line. ")

   CONTROL ADD CHECKBOX, hPrt, %PRINT_FOOTER, "Print page footers? ", 4, 330, 160, 16, %BS_LEFT OR %BS_VCENTER
   CONTROL SET CHECK     hPrt, %PRINT_FOOTER, ENV.PrtFooter
   CONTROL SET COLOR     hPrt, %PRINT_FOOTER, %RGB_DARKBLUE, %RGB_GAINSBORO
   sToolTipSet (GetDlgItem(hPrt, %PRINT_FOOTER), " Should page footers as described below be printed? ")

   CONTROL ADD LABEL,    hPrt, %PRINT_FOOTER_LEFT_TEXT, "Left-hand footer", 4, 345, 100, 16
   CONTROL SET COLOR     hPrt, %PRINT_FOOTER_LEFT_TEXT, %BLACK, %RGB_GAINSBORO
   CONTROL ADD TEXTBOX,  hPrt, %PRINT_FOOTER_LEFT, ENV.PrtFooterLeft, 4, 365, 225, 20
   CONTROL SET FONT      hPrt, %PRINT_FOOTER_LEFT, hFixedFont
   sToolTipSet (GetDlgItem(hPrt, %PRINT_FOOTER_LEFT), " Enter the format string for the left side of the footer line. ")

   CONTROL ADD LABEL,    hPrt, %PRINT_FOOTER_CENTER_TEXT, "Center footer", 235, 345, 200, 16
   CONTROL SET COLOR     hPrt, %PRINT_FOOTER_CENTER_TEXT, %BLACK, %RGB_GAINSBORO
   CONTROL ADD TEXTBOX,  hPrt, %PRINT_FOOTER_CENTER, ENV.PrtFooterCenter, 235, 365, 225, 20
   CONTROL SET FONT      hPrt, %PRINT_FOOTER_CENTER, hFixedFont
   sToolTipSet (GetDlgItem(hPrt, %PRINT_FOOTER_CENTER), " Enter the format string for the center of the footer line. ")

   CONTROL ADD LABEL,    hPrt, %PRINT_FOOTER_RIGHT_TEXT, "Right-hand footer", 465, 345, 200, 16
   CONTROL SET COLOR     hPrt, %PRINT_FOOTER_RIGHT_TEXT, %BLACK, %RGB_GAINSBORO
   CONTROL ADD TEXTBOX,  hPrt, %PRINT_FOOTER_RIGHT, ENV.PrtFooterRight, 465, 365, 225, 20
   CONTROL SET FONT      hPrt, %PRINT_FOOTER_RIGHT, hFixedFont
   sToolTipSet (GetDlgItem(hPrt, %PRINT_FOOTER_RIGHT), " Enter the format string for the right hand of the footer line. ")

   '----- DISPLAY the PRINT FORMAT DIALOG BOX
   GRAPHIC REDRAW                                                 '
   sPopReady                                                      ' Ready for pop-up
   DIALOG SHOW MODAL hPrt, CALL DlgPrintCallBack                  '
   sPopReset                                                      ' Reset popup state
   TAB SELECT hWnd, %IDC_SPFLiteTAB, TP.PgNumber                  ' Select its tab
   MExit
END SUB

CALLBACK FUNCTION DlgANSICallBack () AS LONG

'---------- Callback function used by the PRINT FORMAT Dialog
LOCAL Txt1, txtBin, txtAnsi AS STRING, cRow, cCol AS LONG, MSG AS ASCIIZ * 100
LOCAL xRow AS LONG, xCol AS LONG

LOCAL dlg_PrfGetX2APtr AS LONG                                    ' ptr to translation table
LOCAL dlg_colMode AS LONG
LOCAL dlg_hANSI AS DWORD

LOCAL ptr_ansiSource AS DWORD
DIM   dlg_ansiSource() AS STRINGZ * 32 AT ptr_ansiSource

   dlg_hANSI = CB.HNDL

   DIALOG GET USER dlg_hANSI, %uda_colMode TO dlg_colMode
   DIALOG GET USER dlg_hANSI, %uda_PrfGetX2APtr TO dlg_PrfGetX2APtr

   DIALOG GET USER dlg_hANSI, %uda_ansiSource TO ptr_ansiSource
   REDIM dlg_ansiSource(0) AS STRINGZ * 32 AT ptr_ansiSource

   ' in non-ANSI mode, we "lie" and say that an EBCDIC (or other) char is put in clipboard
   ' in reality, we ONLY put ANSI data in clipboard, because that is the only kind that it can have

   SELECT CASE AS LONG CB.MSG

      '----- Left Button, put selected chr in clipboard
      CASE %WM_LBUTTONDOWN
         Txt1 = ""                                                ' Null for left click
         GOSUB getchar                                            ' Go get character


      '----- Right Button, ADD selected chr to clipboard
      CASE %WM_RBUTTONUP                                          '
         sWinclip_get(Txt1)                                       ' Get any current text to start
         Txt1 = PARSE$(Txt1, $CRLF, 1)                            ' Get 1st 'line'
         GOSUB getchar                                            '

      '----- Final Dialog termination
      CASE %WM_DESTROY                                            ' Sent when the dialog is being destroyed
         RESET dlg_hANSI                                          ' Kill hook

   END SELECT                                                     '

   EXIT FUNCTION                                                  '

'----- Figure out which chr pointed to, add to clipboard
getchar:
   IF LO(WORD,CB.LPARAM) = 0 THEN EXIT FUNCTION                   ' If invalid location, ignore
   IF HI(WORD,CB.LPARAM) = 0 THEN EXIT FUNCTION                   '
   cCol = CEIL(LO(WORD,CB.LPARAM) / gFontWidth)                   ' Calc where click was
   cRow = CEIL(HI(WORD,CB.LPARAM) / gFontHeight)                  '

   xRow = cRow - 3                                                ' form relative-0 row value
   xCol = (cCol - 4) \ 3                                          ' form relative-0 col value

   IF xRow < 0 OR xRow > 15 THEN EXIT FUNCTION                    ' Ignore headings
   IF xCol < 0 OR xCol > 15 THEN EXIT FUNCTION                    ' and clipboard line

   IF dlg_colMode THEN
      txtBin = CHR$(xCol * 16 + xRow)                             ' Get character in col mode
   ELSE
      txtBin = CHR$(xRow * 16 + xCol)                             ' Get character in row mode
   END IF

   '/ at this point, 'txtBin' is the NATIVE binary value, which is either ANSI or EBCDIC

   txtAnsi = txtBin                                               ' Save ANSI version

   IF xRow + xCol = 0 THEN EXIT FUNCTION                          ' Don't store NUL

   '/ if we are in ANSI mode, ANSI and SOURCE are the same thing
   '/ if not in ANSI mode, txtAnsi actually has a 'binary' value (like EBCDIC)
   '/ so we have to convert it from its non-ANSI source to an ANSI one

   '/ the ANSI mode we are checking is the mode in effect when the popup was launched
   '/ we can't check gfANSIMode, because that changes each time the current edit
   '/ screen is changed to a different file, which may have a different source attribute

   IF dlg_PrfGetX2APtr <> 0 THEN                                  ' If not ANSI mode
      TP.Translate(txtAnsi, dlg_PrfGetX2APtr)                     ' Get the ANSI equivalent
   END IF

   '  New format on bottom of ANSI popup, with new line added:
   '
   '  Dec Hex Chr Len Current ANSI Clipboard:
   '  051  33 '3'  6  "ABC123"
   '                  |
   '  1234567890123456|

   Txt1 += txtAnsi                                                ' Add to string

   MSG = FORMAT$(ASC(txtBin), "000") + "  " + HEX$(ASC(txtBin), 2) _
      + " '" + txtAnsi + "'" + DEC$(LEN(Txt1), -3) + " " + $DQ

   sWinclip_set(Txt1)                                             ' Send it to the Clipboard

   GRAPHIC ATTACH dlg_hANSI, 5001                                 ' Set as the default graphic area
   GRAPHIC SET FONT hScrFont                                      ' Set the desired font

   GRAPHIC COLOR cTxtLoFG, cTxtLoBG1                              '

   ' the caption "Dec Hex Chr Len Current ANSI Clipboard:" does not need to be redrawn here

   GRAPHIC SET POS (1, 22 * gFontHeight)                          ' Set position for last row

   GRAPHIC PRINT MSG + Txt1 + $DQ + SPACE$(50)                    ' Print the next ANSI char

   GRAPHIC REDRAW                                                 '
   RETURN
END FUNCTION


CALLBACK FUNCTION DlgCallback()
'---------- Main Dialog's Callback Procedure
LOCAL x, y, PageNo, CurrTab AS LONG              '
LOCAL MSG AS kbMsg, pDrop AS LONG, sFiles, fn, Factor, tt AS STRING
LOCAL ptnmhdr AS NMHDR PTR                                        ' Notify msg header
LOCAL lpnm AS NMMOUSE PTR
LOCAL oldTP  AS iObjTabData                                       ' Save Pointer to current tab data
LOCAL tRow, tCol, i, j AS LONG
LOCAL lpdis AS DRAWITEMSTRUCT PTR, rc AS RECT
DIM KeyStateArray(0 TO 255) AS BYTE                               ' Keyboard stats
STATIC idEvent    AS LONG                                         ' Keep SetTimer's result in a static variable
STATIC pDataToGet AS COPYDATASTRUCT POINTER                       ' WM_COPYDATA data
STATIC pBuffer    AS ASCIIZ PTR                                   '
STATIC aRow, aCol, WentMin AS LONG                                '
STATIC StartRow, StartCol, LastX, LastY, DblClick AS LONG         ' For mouse select
STATIC LDownTime AS DOUBLE                                        ' To time LB clicks
STATIC dbi, SizeParam AS LONG
STATIC BlinkCtr AS INTEGER
STATIC t AS ASCIIZ * 200

   IF CB.HNDL <> hWnd THEN
      GOTO PageDialog
   END IF

   '----- Handle messages related to the main hWnd Dialog
   SELECT CASE AS LONG CB.MSG                                     '

      '----- Right Click on Tab Header
      CASE %WM_ContextMenu                                        '
         LOCAL HT AS TC_HitTestInfo, temp$                        '
         GetCursorPos HT.pt : ScreenToClient hTab, HT.pt          ' Tab coordinates of mouse
         PageNo = SendMessage(hTab, %TCM_HitTest, 0, VARPTR(HT)) + 1 ' Tab page under mouse at drop
         IF PageNo = 0 THEN GOTO DlgExit                          ' Leave if other than on tab header
         IF TabsNum = 0 THEN EXIT SELECT                          ' Prevent Memory Access violations at termination
         j = 0                                                    ' Clear answer index
         FOR i = 1 TO TabsNum                                     ' Search for correct TP entry
            IF ISOBJECT(Tabs(i)) THEN                             ' Is this still a valid entry?
               TP = Tabs(i)                                       ' Make it addressable
               IF PageNo = TP.PgNumber THEN                       ' This is the one, exit
                  j = i                                           ' Save the winner
                  EXIT FOR                                        '
               END IF                                             '
            END IF                                                '
         NEXT i                                                   '
         IF j = 0 THEN EXIT SELECT                                ' Bail out if we didn't find a valid Tab
         IF IsFMTab THEN                                          ' The FM Tab?
            IF (TP.AttnDo AND %SaveReq) THEN                      ' Anything pending?
               TP.SaveReq                                         ' Save Req list if anything pending
               TPDOset(%LoadReq)                                  ' Ensure it's re-loaded
            END IF                                                '
         END IF                                                   '
         sCallTabCmd(j, "END ")                                   ' Try a END command

      '----- Dialog is being initialized
      CASE %WM_INITDIALOG                                         '
         hKbrdHook = SetWindowsHookEx(%WH_KEYBOARD, CODEPTR(MainAKeyHook), _
                     0, GetCurrentThreadId)                       '
         IF IsTPNsrtFlag THEN OnTPNsrtSave ELSE OffTPNsrtSave     ' Set initial Insert status
         OffTPNsrtData                                            '
         sCaretDestroy                                            '
         sCaretCreate                                             ' Create the caret
         sDoCursor                                                ' Set Cursor
         sCaretShow                                               ' Show it
         ENV.InitDone = %True                                     ' Say all Initialization is done


      CASE %WM_SIZE
         IF TabsNum > 0 THEN                                      ' Initialized?
            IF CB.WPARAM = %SIZE_MAXIMIZED THEN                   ' Maximized?
               IF ISFALSE gResizeActive THEN                      ' One at a time
                  gResizeActive = %True                           ' Say Resize is active
                  sResize("M")                                    ' Go do what the user wants
                  gResizeActive = %False                          ' Say Resize done
                  ENV.Maximized = %True                           ' Remember we're maximixed
                  IF IsFMTab THEN                                 ' The FM Tab?
                     IF (TP.AttnDo AND %SaveReq) THEN             ' Anything pending?
                        TP.SaveReq                                ' Save Req list if anything pending
                        TPDOset(%LoadReq)                         ' Ensure it's re-loaded
                     END IF                                       '
                  END IF                                          '
                  sDoStatusBar($AllStatusBarBoxes)                ' re-Do the StatusBar box
               END IF                                             '
            ELSEIF CB.WPARAM = %SIZE_RESTORED THEN                ' Restored?
               IF ISFALSE gResizeActive THEN                      ' One at a time
                  gResizeActive = %True                           ' Say Resize is active
                  sResize("N")                                    ' Go do what the user wants
                  gResizeActive = %False                          ' Say Resize done
                  ENV.Maximized = %False                          ' Remember we're normal
                  sDoStatusBar($AllStatusBarBoxes)                ' re-Do the StatusBar box
               END IF                                             '
               IF WentMin THEN                                    ' Coming back from minimized?
                  WentMin = %False                                '
                  TP.DispScreen                                   ' Make sure we actually display something
               END IF                                             '
            ELSEIF CB.WPARAM = %SIZE_MINIMIZED THEN               ' Minimized
               IF IsFMTab THEN                                    ' The FM Tab?
                  IF (TP.AttnDo AND %SaveReq) THEN                ' Anything pending?
                     TP.SaveReq                                   ' Save Req list if anything pending
                     TPDOset(%LoadReq)                            ' Ensure it's re-loaded
                  END IF                                          '
               END IF                                             '
               WentMin = %True                                    '
            END IF                                                '

         END IF
         FUNCTION = 0

      '----- Our window is being activated/deactivated
      CASE %WM_ACTIVATEAPP
         IF TabsNum = 0 THEN FUNCTION = 0: GOTO DlgExit           ' Prevent Memory Access violations at termination
         IF gfDoingMsg THEN GOTO DlgExit                          ' Don't activate if doing a pop-up, let IT get control
         IF CB.WPARAM THEN                                        ' Being activated
            IF acol < 1 THEN acol = 11                            ' Just in case
            IF arow < 1 THEN arow = 1                             '
            TP.CsrCol = aCol: TP.CsrRow = aRow                    '
            gfActive = %True                                      ' Tell the world we're active
            GRAPHIC ATTACH TP.PgHandle, TP.WindowID               ' Set as the default graphic area
            GRAPHIC CLEAR cTxtLoBG1                               ' Clear
            CONTROL SET FOCUS hWnd, %IDC_SPFLiteTAB               ' Set focus
            IF IsTPPTypeMode THEN _                               ' Restore PTYPE message
               scError(%eNone, "Entering PowerType mode, Press Enter to exit")
            IF IsFMTab THEN TP.AttnDo = (TP.AttnDo OR %LoadData)  ' If an FM screen, trigger a re-load
            TP.LastRulCol = 0: TP.LastRulRow = 0: TP.LastPTCurs = 0 ' In case special cursors active
            TP.DispScreen                                         ' Refresh
            sCaretDestroy                                         '
            sCaretCreate                                          ' Create it
            sDoCursor                                             ' Position it
            sCaretShow                                            ' Show the caret
            FUNCTION = 0                                          ' We're done
         ELSE                                                     ' Being deactivated
            aRow = TP.CsrRow: aCol = TP.CsrCol                    '
            IF IsFMTab THEN                                       ' The FM Tab?
               IF (TP.AttnDo AND %SaveReq) THEN                   ' Anything pending?
                  TP.SaveReq                                      ' Save Req list if anything pending
                  TPDOset(%LoadReq)                               ' Ensure it's re-loaded
               END IF                                             '
            END IF                                                '
            sCaretDestroy                                         ' Destroy it
            gfActive = %False                                     ' Tell the world we're NOT active
            FUNCTION = 0                                          ' We're done
         END IF

      '----- A file has been dropped on our window
      CASE %WM_DROPFILES                                          ' Look for dropped files
         IF gfDoingMsg > 0 THEN GOTO DlgExit                      ' Busy doing pop-up, ignore
         TP.MarkKill                                              ' Kill any select box
         TP.SwapKill                                              ' Kill any Swap box
         pDrop = CBWPARAM                                         ' Copy param
         sFiles = sGetDropFiles(pDrop)                            ' Go get list
         DO WHILE ISNOTNULL(sFiles)                               ' Process the list
            fn = EXTRACT$(sFiles, "|")                            ' Get next filename
            sFiles = REMAIN$(sFiles, "|")                         ' Strip it out
            pCmdEdit("EDIT " + $DQ + fn + $DQ)                    ' Let Edit have a go
            IF gTabSwitch <> 0 THEN sDoPendingTabSwitch           '
            gCaretCtr = 0                                         ' Force caret out
            sCaretCreate                                          ' Create it
            sDoCursor                                             ' Position it
            sCaretShow                                            ' Show the caret
         LOOP                                                     '

      '----- Someone is sending us data
      CASE %WM_COPYDATA                                           ' We are receiving data
         IF gfDoingMsg > 0 THEN GOTO DlgExit                      ' Busy doing popup, ignore
         pDataToGet = CBLPARAM                                    ' Pointer is in CBLPARAM
         pBuffer    = @pDataToGet.lpData                          ' Pointer to data string
         IF ISNULL(TRIM$(TP.pCommand)) THEN                       ' Build command
            TP.pCommand = @pBuffer                                ' Pass the command onward

         ELSE                                                     '
            TP.pCommand = TRIM$(TP.pCommand) + ENV.CmdChr + @pBuffer
         END IF                                                   '
         gKeyChr = "ENTER"                                        ' Simulate an ENTER key
         MainDKeyProcess                                          ' Process the key event
         sDoPendingTabSwitch                                      ' Do any pending Tab switch
         FUNCTION = 1                                             ' Say we've handled it

      '----- A NOTIFY message coming in
      CASE %WM_NOTIFY                                             ' Notify message

         ptnmhdr = CBLPARAM                                       ' Point at notify header
         SELECT CASE @ptnmhdr.idFrom                              ' See where Notify from
            CASE %IDC_Statusbar                                   ' The statusbar
               IF @ptnmhdr.code = %NM_CLICK THEN                  ' A mouse click?
                  lpnm = CBLPARAM                                 ' Check the box number
                  IF @lpnm.dwItemSpec = 10 THEN                   ' Slec box?
                     krToggleSelect                               ' Toggle it
                  END IF                                          '
               END IF                                             '
         END SELECT                                               '

         SELECT CASE AS LONG CB.NMCODE                            ' What type?

            '----- A tab is getting control
            CASE %TCN_SELCHANGE                                   ' We're getting control
               IF TabsNum = 0 THEN EXIT SELECT                    ' Prevent Memory Access violations at termination
               CONTROL SEND CB.HNDL, %IDC_SPFLiteTAB, %TCM_GETCURSEL,0,0 TO PageNo
               TP = NOTHING
               FOR i = 1 TO TabsNum                               ' Search for correct TP entry
                  IF ISOBJECT(Tabs(i)) THEN                       ' Is this still a valid entry?
                     TP = Tabs(i)                                 ' Copy the TP address
                     IF PageNo + 1 = TP.PgNumber THEN             ' This is the one, exit
                        EXIT FOR                                  '
                     END IF                                       '
                  END IF                                          '
               NEXT i                                             '
               IF ISNOTHING(TP) THEN EXIT SELECT                  ' Bail out if we didn't set TP to a valid Tab

               GRAPHIC ATTACH TP.PgHandle, TP.WindowID            ' Swap the default graphic area
               CONTROL SET FOCUS hWnd, %IDC_SPFLiteTAB            ' Regain focus.
               TP.WindowTitle                                     ' Swap the Title bar
               TP.LastPTCurs = 0                                  '
               TP.LastRulCol = 0: TP.LastRulRow = 0               ' In case crosshair cursors active
               IF IsFMTab THEN TP.AttnDo = (TP.AttnDo OR %LoadData) ' If an FM screen, trigger a re-load
               TP.DispScreen                                      ' Refresh if the very first show
               sCaretDestroy                                      '
               sCaretCreate                                       ' Create it
               sDoCursor                                          ' Position it
               sCaretShow                                         ' Show the caret
               sSetupSB                                           ' Setup the status bar fields
               sDoStatusBar($AllStatusBarBoxes)                   ' Make sure they're all done
               sTabStackAdd(TP.PgNumber)                          '

            '----- A tab is losing control
            CASE %TCN_SELCHANGING                                 ' We're losing control
               IF IsTPPTypeMode THEN                              ' If PowerType mode
                  OffTPPTypeMode                                  ' Kill it
                  gPTblCount = 0                                  '
                  TP.ErrMsg = ""                                  '
               END IF                                             '
               IF TabsNum = 0 THEN EXIT SELECT                    ' Prevent Memory Access violations at termination
               IF IsFMTab THEN                                    ' The FM Tab?
                  IF (TP.AttnDo AND %SaveReq) THEN                ' Anything pending?
                     TP.SaveReq                                   ' Save Req list if anything pending
                     TPDOset(%LoadReq)                            ' Ensure it's re-loaded
                  END IF                                          '
               END IF                                             '
               CONTROL SEND CB.HNDL, %IDC_SPFLiteTAB, %TCM_GETCURSEL,0,0 TO PageNo
               TP = NOTHING
               FOR i = 1 TO TabsNum                               ' Search for correct TP entry
                  IF ISOBJECT(Tabs(i)) THEN                       ' Is this still a valid entry?
                     TP = Tabs(i)                                 ' Copy the TP address
                     IF PageNo + 1 = TP.PgNumber THEN EXIT FOR    ' This is the one, exit
                  END IF                                          '
               NEXT i                                             '
               IF ISNOTHING(TP) THEN EXIT SELECT                  ' Bail out if we didn't set TP to a valid Tab
               TP.MarkKill                                        ' Kill any marked text
               TP.SwapKill                                        ' Kill any swap text
               sCaretHide                                         '
               sCaretDestroy                                      '
               FUNCTION = 0                                       ' Allow tab swap to continue
         END SELECT                                               '

      '----- A user draw item request
      CASE %WM_DRAWITEM
         IF CBWPARAM = %IDC_STATUSBAR THEN                        ' Our statusbar redraw?
            sSBStatusbarDrawItem CBHNDL, CBLPARAM                 ' Do it
         END IF                                                   '
         IF CB.WPARAM = %IDC_SPFLiteTAB THEN                      '
            sTabHighLight CBHNDL, CBWPARAM, CBLPARAM              '
         END IF
         FUNCTION = 1                                             '

      '----- A USER message, either KB activity or FileWatch notify
      CASE %WM_USER                                               ' A User message coming in
         Msg.MsgWparam = CB.WPARAM                                ' Put wParam into the UNION
         gKeyChr = Msg.kbString                                   ' Move string into gKeyChr
         MainBUserMsg()                                           ' Go analyze what came in

      '----- A MouseWheel movement occurred
      CASE %WM_MOUSEWHEEL                                         ' Scroll wheel moved
         INCR dbi
         IF gfMiddleDown THEN GOTO DlgExit                        ' Ignore if Middle button down
         IF ENV.AutoScroll AND ISNULL(TRIM$(TP.pCommand)) THEN    ' Honor it at all? (i.e. no current pCommand + ENV.AutoScroll)
            i = CINT(HI(WORD, CBWPARAM))                          ' Get wheel direction and amount
            j = CINT(LO(WORD, CBWPARAM))                          ' Get Ctl keys
            Factor = FORMAT$(IIF((j AND %MK_CONTROL) = %MK_CONTROL, 4, 1)) + " M" ' Set Factor based on Ctrl key state
            IF (j AND %MK_SHIFT) <> %MK_SHIFT THEN                ' Shift key down?
               IF IsTPPTypeMode THEN GOTO DlgExit                 ' No Up/Down in PT mode
               IF i < 0 THEN                                      ' Backward rotation
                  OnTPInsClnSupp                                  ' Suppress Insert cleanup for 1 Attn pass
                  pCmdScroll("ISCROLL DOWN " + Factor)            ' Directly Call Scroll
                  TP.DispScreen                                   ' and display screen
                  sDoCursor                                       ' Set cursor
               ELSE                                               ' Forward
                  OnTPInsClnSupp                                  ' Suppress Insert cleanup for 1 Attn pass
                  pCmdScroll("ISCROLL UP " + Factor)              ' Directly call Scroll
                  TP.DispScreen                                   ' and display screen
                  sDoCursor                                       ' Set cursor
               END IF                                             '
            ELSE
               IF i < 0 THEN                                      ' Backward rotation
                  OnTPInsClnSupp                                  ' Suppress Insert cleanup for 1 Attn pass
                  pCmdScroll("ISCROLL RIGHT " + Factor)           ' Directly call Scroll
                  TP.DispScreen                                   ' and display screen
                  sDoCursor                                       ' Set cursor
               ELSE                                               ' Forward
                  OnTPInsClnSupp                                  ' Suppress Insert cleanup for 1 Attn pass
                  pCmdScroll("ISCROLL LEFT " + Factor)            ' Directly call Scroll
                  TP.DispScreen                                   ' and display screen
                  sDoCursor                                       ' Set cursor
               END IF                                             '
            END IF                                                '
         END IF                                                   '

      '----- SYSCOMMAND
      CASE %WM_SYSCOMMAND
         IF CB.HNDL <> hWnd THEN GOTO DlgExit                     '
         IF (CBWPARAM AND &HFFF0) = %SC_CLOSE THEN                ' Trap the [x] button and Alt-F4
            IF ISFALSE gfTermFlag THEN                            ' Been here before?
               sDoEndAll("EXIT")                                  ' No, Treat as an EXIT command
               IF ISFALSE gfEndAll THEN FUNCTION = 1 ELSE FUNCTION = 0
            ELSE                                                  '
               DIALOG GET LOC hWnd TO x, y                        ' Get location where window was
               ENV.LastScreenX = x: ENV.LastScreenY = y           ' Save them away
            END IF                                                '
         END IF                                                   '

      '----- CLOSE button pressed
      CASE %WM_CLOSE                                              ' Sent when the Close button is pressed
         IF CB.HNDL <> hWnd THEN GOTO DlgExit                     '
         DIALOG GET LOC hWnd TO x, y                              ' Get location where window was
         ENV.LastScreenX = x: ENV.LastScreenY = y                 ' Save them away

      '----- Final Dialog termination
      CASE %WM_DESTROY                                            ' Sent when the dialog is being destroyed
         UnhookWindowsHookEx hKbrdHook                            ' Kill the Keyboard hook
         HKbrdHook = 0                                            '

      CASE ELSE
         EXIT FUNCTION
   END SELECT

      GOTO DlgExit

   '----- Handle messages related to the main Tab Page

   PageDialog:

   SELECT CASE AS LONG CB.MSG

      '----- Left button press
      CASE %WM_LBUTTONDOWN                                        '
         GOSUB GetRowCol                                          ' Get Row/Col to tRow/tCol
         LastX = LO(WORD,CBLPARAM)                                ' Save last Mouse coordinate
         LastY = HI(WORD,CBLPARAM)                                '
         IF IsTPPTypeMode AND tRow <> gPTbl(1).sRow THEN GOTO DlgExit ' Must be same as PTmode Model line
         LDownTime = TIMER                                        ' Save button down time
         TP.MarkStart(trow, tcol,StartRow, StartCol)              ' See if we start a mark area

      '----- The mouse moved
      CASE %WM_MOUSEMOVE

         IF gfLeftDown THEN                                       ' If Mouse is down
            SlecthDC = GetDC(CBHNDL)                              '
            GOSUB GetRowCol                                       ' Get Row/Col to tRow/tCol
            IF ABS(LastX - LO(WORD,CBLPARAM)) < 3 AND _           ' See if minimum movement
               ABS(LastY - HI(WORD,CBLPARAM)) < 3 THEN GOTO DlgExit '
            TP.MarkExtend(trow, tcol,StartRow, StartCol)          ' Go see if we extend an area
         END IF                                                   '

      '----- Left Button Up
      CASE %WM_LBUTTONUP
         gfLeftDown = %False                                      '
         IF ISTRUE DblClick THEN                                  ' Is this the UP following a DblClick?
            DblClick = %False                                     ' Turn it off
            GOTO DlgExit                                          ' And ignore it
         END IF                                                   '
         GOSUB GetRowCol                                          ' Get Row/Col to tRow/tCol

         IF IsTPPTypeMode AND tRow <> gPTbl(1).sRow THEN GOTO DlgExit ' Must be same as PTmode Model line
         IF (TIMER - LDownTime) < 0.3 AND TP.MarkSCol = TP.MarkECol AND TP.MarkSLin = TP.MarkELin AND TP.MarkSCol > gLNPadCol THEN TP.MarkKill

         '----- If we're on the FM screen
         IF IsFMTab THEN                                          ' If we're on the FM screen
            GOSUB SetCurs                                         ' Go do it

            '----- If on column header line, adjust the data SORT criteria
            IF TP.CsrRow = FM_Head_Line THEN                      ' Column header line?
               SELECT CASE AS LONG TP.CsrCol                      ' Yes, now look at column

                  CASE FM_Head_Name_Left TO FM_Head_Name_Left + 3 ' Somewhere over Dir. ?
                     IF RIGHT$(TP.DirSort, 1) = "+" THEN          '
                        TP.DirSort = "Dir-"                       ' Rotate through the options
                     ELSEIF RIGHT$(TP.DirSort, 1) = "-" THEN      '
                        TP.DirSort = "Dir*"                       '
                     ELSEIF RIGHT$(TP.DirSort, 1) = "*" THEN      '
                        TP.DirSort = "Dir+"                       '
                     END IF                                       '
                     TP.LastTop = 0                               ' Kill position

                  CASE FM_Head_Name_Left + 6 TO FM_Head_Name_Left + FM_File_Size - 1
                     IF TP.ScreenChrHit(FM_Head_Line, TP.CsrCol) THEN ' And on a non-blank?
                        IF LEFT$(TP.DefSort, 4) = "Name" THEN     ' Already doing file sort?
                           TP.DefSort = IIF$(TP.DefSort = "Name+", "Name-", "Name+") ' Yes, flip direction
                        ELSE                                      '
                           TP.DefSort = "Name+"                   ' Swap to NAME+ sort
                        END IF                                    '
                        TP.LastTop = 0                            ' Kill position
                     END IF                                       '

                  CASE FM_Head_Ext_Left TO FM_Head_Ext_Left + %FMLExt - 1
                     IF FM_Head_Ext_Left > 0 THEN                 '  But only if EXT is being displayed
                        IF TP.ScreenChrHit(FM_Head_Line, TP.CsrCol) THEN ' And on a non-blank?
                           IF LEFT$(TP.DefSort, 4) = "Extn" THEN  ' Already doing EXT sort?
                              TP.DefSort = IIF$(TP.DefSort = "Extn+", "Extn-", "Extn+") ' Yes, flip direction
                           ELSE                                   '
                              TP.DefSort = "Extn+"                ' Swap to EXTN+ sort
                           END IF                                 '
                           TP.LastTop = 0                         ' Kill position
                        END IF                                    '
                     END IF                                       '

                  CASE FM_Head_Size_Left TO FM_Head_Size_Left + IIF(INSTR(ENV.FMLayout, "SIZELONG"), %FMLSizeL, %FMLSizeS) - 1
                     IF FM_Head_Size_Left > 0 THEN                '  But only if SIZE is being displayed
                        IF TP.ScreenChrHit(FM_Head_Line, TP.CsrCol) THEN ' And on a non-blank?
                           IF LEFT$(TP.DefSort, 4) = "Size" THEN  ' Already doing SIZE sort?
                              TP.DefSort = IIF$(TP.DefSort = "Size+", "Size-", "Size+") ' Yes, flip direction
                           ELSE                                   '
                              TP.DefSort = "Size-"                ' Swap to SIZE- sort
                           END IF                                 '
                           TP.LastTop = 0                         ' Kill position
                        END IF                                    '
                     END IF                                       '

                  CASE FM_Head_Lines_Left TO FM_Head_Lines_Left + %FMLLines - 1
                     IF FM_Head_Lines_Left > 0 THEN               '  But only if LINES is being displayed
                        IF TP.ScreenChrHit(FM_Head_Line, TP.CsrCol) THEN ' And on a non-blank?
                           IF LEFT$(TP.DefSort, 5) = "Lines" THEN ' Already doing LINES sort?
                              TP.DefSort = IIF$(TP.DefSort = "Lines+", "Lines-", "Lines+") ' Yes, flip direction
                           ELSE                                   '
                              TP.DefSort = "Lines-"               ' Swap to LINES- sort
                           END IF                                 '
                           TP.LastTop = 0                         ' Kill position
                        END IF                                    '
                     END IF                                       '

                  CASE FM_Head_Date_Left TO FM_Head_Date_Left + IIF(INSTR(ENV.FMLayout, "DATETIME"), %FMLDateT, %FMLDate) - 1
                     IF FM_Head_Date_Left > 0 THEN                '  But only if Date is being displayed
                        IF TP.ScreenChrHit(FM_Head_Line, TP.CsrCol) THEN ' And on a non-blank?
                           tt = TP.DefSort
                           IF LEFT$(TP.DefSort, 4) = "Date" THEN  ' Already doing Date sort?
                              TP.DefSort = IIF$(TP.DefSort = "Date+", "Date-", "Date+") ' Yes, flip direction
                              tt = TP.DefSort
                           ELSE                                   '
                              TP.DefSort = "Date-"                ' Swap to DATE- sort
                              tt = TP.DefSort
                           END IF                                 '
                           TP.LastTop = 0                         ' Kill position
                           tt = TP.DefSort
                        END IF                                    '
                     END IF                                       '

                  CASE FM_Head_Note_Left TO FM_Head_Note_Left + FM_Note_Size - 1
                     IF FM_Head_Note_Left > 0 THEN                '  But only if Note is being displayed
                        IF TP.ScreenChrHit(FM_Head_Line, TP.CsrCol) THEN ' And on a non-blank?
                           IF LEFT$(TP.DefSort, 4) = "Note" THEN  ' Already doing Note sort?
                              TP.DefSort = IIF$(TP.DefSort = "Note+", "Note-", "Note+") ' Yes, flip direction
                           ELSE                                   '
                              TP.DefSort = "Note+"                ' Swap to NOTE+ sort
                           END IF                                 '
                           TP.LastTop = 0                         ' Kill position
                        END IF                                    '
                     END IF                                       '

               END SELECT                                         '
               ENV.SetFMLast(TP.DirSort, TP.DefSort)              ' Save the new value
               TP.AttnDo = (TP.AttnDo OR %LoadData)               ' Request a data refresh
               sCaretDestroy                                      ' Hide cursor
               TP.DispScreen                                      '
               sCaretCreate                                       '
               sDoCursor                                          ' Turn Cursor back on
               sCaretShow                                         '
               GOTO DlgExit                                       '

            '----- Left up click in the FM quick row
            ELSEIF TP.CsrRow = FM_Quick_Line_1 THEN               ' Quick items?
               SELECT CASE AS LONG TP.CsrCol                      ' Yes, now look at column
                  CASE FM_Quick_Pos_1 TO FM_Quick_Pos_2 - 1       ' Within item?
                     IF TP.ScreenChrHit(TP.CsrRow, TP.CsrCol) THEN' And on a non-blank?
                        TP.CmdEmpty = "S"                         ' Create an "S"
                        GOSUB PostAttention                       ' Post an Attention request
                     END IF                                       '
                  CASE FM_Quick_Pos_2 TO FM_Quick_Pos_3 - 1       ' Within item?
                     IF TP.ScreenChrHit(TP.CsrRow, TP.CsrCol) THEN' And on a non-blank?
                        TP.CmdFilePath = "S"                      ' Create an "S"
                        GOSUB PostAttention                       ' Post an Attention request
                     END IF                                       '
                  CASE FM_Quick_Pos_3 TO FM_Quick_Pos_4 - 1       ' Within item?
                     IF TP.ScreenChrHit(TP.CsrRow, TP.CsrCol) THEN' And on a non-blank?
                        TP.CmdRecent = "S"                        ' Create an "S"
                        GOSUB PostAttention                       ' Post an Attention request
                     END IF                                       '
                  CASE FM_Quick_Pos_4 TO FM_Quick_Pos_5 - 1       ' Within item?
                     IF TP.ScreenChrHit(TP.CsrRow, TP.CsrCol) THEN' And on a non-blank?
                        TP.CmdFound = "S"                         ' Create an "S"
                        GOSUB PostAttention                       ' Post an Attention request
                     END IF                                       '
                  CASE FM_Quick_Pos_5 TO FM_Quick_Pos_6 - 1       ' Within item?
                     IF TP.ScreenChrHit(TP.CsrRow, TP.CsrCol) THEN' And on a non-blank?
                        TP.CmdOpen = "S"                          ' Create an "S"
                        GOSUB PostAttention                       ' Post an Attention request
                     END IF                                       '
                  CASE FM_Quick_Pos_6 TO FM_Quick_Pos_7 - 1       ' Within item?
                     IF TP.ScreenChrHit(TP.CsrRow, TP.CsrCol) THEN' And on a non-blank?
                        TP.CmdFavorite = "S"                      ' Create an "S"
                        GOSUB PostAttention                       ' Post an Attention request
                     END IF                                       '
                  CASE FM_Quick_Pos_7 TO FM_Quick_Pos_8 - 1       ' Within item?
                     IF TP.ScreenChrHit(TP.CsrRow, TP.CsrCol) THEN' And on a non-blank?
                        TP.CmdFileList = "S"                      ' Create an "S"
                        GOSUB PostAttention                       ' Post an Attention request
                     END IF                                       '
                  CASE FM_Quick_Pos_8 TO FM_Quick_Pos_9 - 1       ' Within item?
                     IF TP.ScreenChrHit(TP.CsrRow, TP.CsrCol) THEN' And on a non-blank?
                        TP.CmdRPaths = "S"                        ' Create an "S"
                        GOSUB PostAttention                       ' Post an Attention request
                     END IF                                       '
                  CASE >= FM_Quick_Pos_9                          ' Within item?
                     IF TP.ScreenChrHit(TP.CsrRow, TP.CsrCol) THEN' And on a non-blank?
                        TP.CmdProfiles = "S"                      ' Create an "S"
                        GOSUB PostAttention                       ' Post an Attention request
                     END IF                                       '
               END SELECT                                         '

            '----- Left up click in the FM file list area
            ELSEIF TP.CursLCmd  OR TP.CursLLin THEN               ' Was this a click in the list area?
               IF TP.CsrCol < FM_Head_Name_Left THEN              ' In the Cmd char area?
                  TP.CsrCol = 1                                   ' Just put cursor in col 1
               ELSE                                               '
                  IF TP.CsrCol >= FM_Head_Name_Left AND TP.CsrCol <= FM_Head_Name_Left + FM_File_Size - 1 AND _
                     TP.ScreenChrHit(TP.CsrRow, TP.CsrCol) THEN   ' And on a non-blank?
                     TP.AFLCmdSet(TP.CsrRow - FM_Top_File_Line + TP.TopScrn, TP.DefCommand)
                     TP.CsrCol = 1                                ' Yes, pretend cursor is in Column 1
                     GOSUB PostAttention                          ' Post an Attention request
                  END IF                                          '
               END IF                                             '
               GOTO DlgExit                                       '
            ELSE                                                  ' Not list area
               GOTO DlgExit                                       '
            END IF                                                '
         END IF                                                   ' End FM code

         j = CINT(LO(WORD, CBWPARAM))                             ' Get Ctl keys
         IF GetKeyState(%VK_MENU) < 0 THEN j = j OR %MK_ALT       ' Add in ALT key if pressed

         IF (j AND (%MK_SHIFT OR %MK_CONTROL OR %MK_ALT)) = (%MK_SHIFT OR %MK_CONTROL OR %MK_ALT) THEN  ' Shift-Ctrl-Alt Left-click
            IF KbdT.SCAUp(105) = "." THEN                         ' Set Cursor?
               GOSUB SetCurs                                      ' Go do it
            END IF                                                '
            IF UUCASE(KbdT.SCAData(105)) <> "(NULL)" AND UUCASE(KbdT.SCAData(105)) <> "(PASSTHRU)" THEN
               gKeyChr = CHR$(03) + KbdT.SCAData(105)             ' Pretend the selected key was hit
               MainBUserMsg()                                     ' Process the key event
            END IF                                                '

         ELSEIF (j AND (%MK_SHIFT OR %MK_ALT)) = (%MK_SHIFT OR %MK_ALT) THEN  ' Shift-Alt Left-click
            IF KbdT.SAUp(105) = "." THEN                          ' Set Cursor?
               GOSUB SetCurs                                      ' Go do it
            END IF                                                '
            IF UUCASE(KbdT.SAData(105)) <> "(NULL)" AND UUCASE(KbdT.SAData(105)) <> "(PASSTHRU)" THEN
               gKeyChr = CHR$(03) + KbdT.SAData(105)              ' Pretend the selected key was hit
               MainBUserMsg()                                     ' Process the key event
            END IF                                                '

         ELSEIF (j AND (%MK_CONTROL OR %MK_ALT)) = (%MK_CONTROL OR %MK_ALT) THEN  ' Control-Alt Left-click
            IF KbdT.CAUp(105) = "." THEN                          ' Set Cursor?
               GOSUB SetCurs                                      ' Go do it
            END IF                                                '
            IF UUCASE(KbdT.CAData(105)) <> "(NULL)" AND UUCASE(KbdT.CAData(105)) <> "(PASSTHRU)" THEN
               gKeyChr = CHR$(03) + KbdT.CAData(105)              ' Pretend the selected key was hit
               MainBUserMsg()                                     ' Process the key event
            END IF                                                '

         ELSEIF (j AND (%MK_SHIFT OR %MK_CONTROL)) = (%MK_SHIFT OR %MK_CONTROL) THEN  ' Shift-Control Left-click
            IF KbdT.SCUp(105) = "." THEN                          ' Set Cursor?
               GOSUB SetCurs                                      ' Go do it
            END IF                                                '
            IF UUCASE(KbdT.SCData(105)) <> "(NULL)" AND UUCASE(KbdT.SCData(105)) <> "(PASSTHRU)" THEN
               gKeyChr = CHR$(03) + KbdT.SCData(105)              ' Pretend the selected key was hit
               MainBUserMsg()                                     ' Process the key event
            END IF                                                '

         ELSEIF (j AND %MK_SHIFT) = %MK_SHIFT THEN                ' Shift Left-Click
            IF KbdT.SUp(105) = "." THEN                           ' Set Cursor?
               GOSUB SetCurs                                      ' Go do it
            END IF                                                '
            IF UUCASE(KbdT.SData(105)) <> "(NULL)" AND UUCASE(KbdT.SData(105)) <> "(PASSTHRU)" THEN
               gKeyChr = CHR$(03) + KbdT.SData(105)               ' Pretend the selected key was hit
               MainBUserMsg()                                     ' Process the key event
            END IF                                                '

         ELSEIF (j AND %MK_CONTROL) = %MK_CONTROL THEN            ' Control Left-click
            IF KbdT.CUp(105) = "." THEN                           ' Set Cursor?
               GOSUB SetCurs                                      ' Go do it
            END IF                                                '
            IF UUCASE(KbdT.CData(105)) <> "(NULL)" AND UUCASE(KbdT.CData(105)) <> "(PASSTHRU)" THEN
               gKeyChr = CHR$(03) + KbdT.CData(105)               ' Pretend the selected key was hit
               MainBUserMsg()                                     ' Process the key event
            END IF                                                '

         ELSEIF (j AND %MK_ALT) = %MK_ALT THEN                    ' Control Alt-click
            IF KbdT.AUp(105) = "." THEN                           ' Set Cursor?
               GOSUB SetCurs                                      ' Go do it
            END IF                                                '
            IF UUCASE(KbdT.AData(105)) <> "(NULL)" AND UUCASE(KbdT.AData(105)) <> "(PASSTHRU)" THEN
               gKeyChr = CHR$(03) + KbdT.AData(105)               ' Pretend the selected key was hit
               MainBUserMsg()                                     ' Process the key event
            END IF                                                '

         ELSE                                                     ' Normal Left-Click
            IF KbdT.NUp(105) = "." THEN                           ' Set Cursor?
               GOSUB SetCurs                                      ' Go do it
            END IF                                                '
            IF UUCASE(KbdT.NData(105)) <> "(NULL)" AND UUCASE(KbdT.NData(105)) <> "(PASSTHRU)" THEN
               gKeyChr = CHR$(03) + KbdT.NData(105)               ' Pretend the selected key was hit
               MainBUserMsg()                                     ' Process the key event
            END IF                                                '

         END IF                                                   '

      '----- Left Button Double click
      CASE %WM_LBUTTONDBLCLK
         DblClick = %True                                         ' To tell next UP to be ignored
         GOSUB GetRowCol                                          ' Get Row/Col to tRow/tCol
         IF IsTPPTypeMode AND tRow <> gPTbl(1).sRow THEN GOTO DlgExit ' Must be same as PTmode Model line
         IF IsTPMarkActive AND _                                  ' Mark active with a real range?
            (TP.MarkSCol <> TP.MarkECol) OR _
            (TP.MarkSLin <> TP.MarkELin) THEN GOTO DlgExit        '
         TP.CsrCol = tCol                                         ' Move cursor
         TP.CsrRow = tRow                                         '
         IF ISFALSE TP.CursData THEN GOTO DlgExit                 ' Exit if not in the data area

         '----- Continue only if we're in the Data area, mark a word if we're on one
         IF TP.CsrCol > gLNPadCol AND TP.CsrRow > 2 + TP.PrfCols AND _ ' Make sure its safe
            TP.CsrRow < gwScrHeight + 1 AND TP.PrfHexMode = &1 AND ISFALSE IsFMTab THEN ' In true text area?
            TP.CursWordLoc                                        ' Go see if we can locate a word
            IF IsTPMarkActive THEN                                ' If we marked it, try drawing a hi-lite
               sCaretDestroy                                      ' Turn cursor off
               TP.MarkScr                                         '
               sCaretCreate                                       '
               sDoCursor                                          ' Get cursor back
               sCaretShow                                         '
            END IF                                                '
         END IF

      '----- Middle Button Down
      CASE %WM_MBUTTONDOWN                                        '
         gfMiddleDown = %True                                     ' Middle down

      '----- Middle Button Up
      CASE %WM_MBUTTONUP                                          '
         gfMiddleDown = %False                                    ' Middle up
         GOSUB GetRowCol                                          ' Get Row/Col to tRow/tCol
         IF IsTPPTypeMode AND tRow <> gPTbl(1).sRow THEN GOTO DlgExit ' Must be same as PTmode Model line
         j = CINT(LO(WORD, CBWPARAM))                             ' Get Ctl keys
         IF GetKeyState(%VK_MENU) < 0 THEN j = j OR %MK_ALT       ' Add in ALT key if pressed

         IF (j AND (%MK_SHIFT OR %MK_CONTROL OR %MK_ALT)) = (%MK_SHIFT OR %MK_CONTROL OR %MK_ALT) THEN  ' Shift-Ctrl-Alt Middle-click
            IF KbdT.SCAUp(106) = "." THEN GOSUB SetCurs           ' Set Cursor?
            IF UUCASE(KbdT.SCAData(106)) <> "(NULL)" AND UUCASE(KbdT.SCAData(106)) <> "(PASSTHRU)" THEN
               gKeyChr = CHR$(03) + KbdT.SCAData(106)             ' Pretend the selected key was hit
               MainBUserMsg()                                     ' Process the key event
            END IF                                                '

         ELSEIF (j AND (%MK_SHIFT OR %MK_ALT)) = (%MK_SHIFT OR %MK_ALT) THEN  ' Shift-Alt Middle-click
            IF KbdT.SAUp(106) = "." THEN GOSUB SetCurs            ' Set Cursor?
            IF UUCASE(KbdT.SAData(106)) <> "(NULL)" AND UUCASE(KbdT.SAData(106)) <> "(PASSTHRU)" THEN
               gKeyChr = CHR$(03) + KbdT.SAData(106)              ' Pretend the selected key was hit
               MainBUserMsg()                                     ' Process the key event
            END IF                                                '

         ELSEIF (j AND (%MK_CONTROL OR %MK_ALT)) = (%MK_CONTROL OR %MK_ALT) THEN  ' Control-Alt Middle-click
            IF KbdT.CAUp(106)= "." THEN GOSUB SetCurs             ' Set Cursor?
            IF UUCASE(KbdT.CAData(106)) <> "(NULL)" AND UUCASE(KbdT.CAData(106)) <> "(PASSTHRU)" THEN
               gKeyChr = CHR$(03) + KbdT.CAData(106)              ' Pretend the selected key was hit
               MainBUserMsg()                                     ' Process the key event
            END IF                                                '

         ELSEIF (j AND (%MK_SHIFT OR %MK_CONTROL)) = (%MK_SHIFT OR %MK_CONTROL) THEN  ' Shift-Control Middle-click
            IF KbdT.SCUp(106)= "." THEN GOSUB SetCurs             ' Set Cursor?
            IF UUCASE(KbdT.SCData(106)) <> "(NULL)" AND UUCASE(KbdT.SCData(106)) <> "(PASSTHRU)" THEN
               gKeyChr = CHR$(03) + KbdT.SCData(106)              ' Pretend the selected key was hit
               MainBUserMsg()                                     ' Process the key event
            END IF                                                '

         ELSEIF (j AND %MK_SHIFT) = %MK_SHIFT THEN                ' Shift Middle-Click
            IF KbdT.SUp(106)= "." THEN GOSUB SetCurs              ' Set Cursor?
            IF UUCASE(KbdT.SData(106)) <> "(NULL)" AND UUCASE(KbdT.SData(106)) <> "(PASSTHRU)" THEN
               gKeyChr = CHR$(03) + KbdT.SData(106)               ' Pretend the selected key was hit
               MainBUserMsg()                                     ' Process the key event
            END IF                                                '

         ELSEIF (j AND %MK_CONTROL) = %MK_CONTROL THEN            ' Control Middle-click
            IF KbdT.CUp(106)= "." THEN GOSUB SetCurs              ' Set Cursor?
            IF UUCASE(KbdT.CData(106)) <> "(NULL)" AND UUCASE(KbdT.CData(106)) <> "(PASSTHRU)" THEN
               gKeyChr = CHR$(03) + KbdT.CData(106)               ' Pretend the selected key was hit
               MainBUserMsg()                                     ' Process the key event
            END IF                                                '

         ELSEIF (j AND %MK_ALT) = %MK_ALT THEN                    ' Alt Middle-click
            IF KbdT.AUp(106)= "." THEN GOSUB SetCurs              ' Set Cursor?
            IF UUCASE(KbdT.AData(106)) <> "(NULL)" AND UUCASE(KbdT.AData(106)) <> "(PASSTHRU)" THEN
               gKeyChr = CHR$(03) + KbdT.AData(106)               ' Pretend the selected key was hit
               MainBUserMsg()                                     ' Process the key event
            END IF                                                '

         ELSE                                                     ' Normal Middle-Click
            IF KbdT.NUp(106)= "." THEN GOSUB SetCurs              ' Set Cursor?
            IF UUCASE(KbdT.NData(106)) <> "(NULL)" AND UUCASE(KbdT.NData(106)) <> "(PASSTHRU)" THEN
               gKeyChr = CHR$(03) + KbdT.NData(106)               ' Pretend the selected key was hit
               MainBUserMsg()                                     ' Process the key event
            END IF                                                '
         END IF                                                   '

      '----- Right Button Down
      CASE %WM_RBUTTONDOWN                                        '

      '----- Right Button Up
      CASE %WM_RBUTTONUP                                          '
         GOSUB GetRowCol                                          ' Get Row/Col to tRow/tCol
         IF IsTPPTypeMode AND tRow <> gPTbl(1).sRow THEN GOTO DlgExit ' Must be same as PTmode Model line
         j = CINT(LO(WORD, CBWPARAM))                             ' Get Ctl keys
         IF GetKeyState(%VK_MENU) < 0 THEN j = j OR %MK_ALT       ' Add in ALT key if pressed

         IF (j AND (%MK_SHIFT OR %MK_CONTROL OR %MK_ALT)) = (%MK_SHIFT OR %MK_CONTROL OR %MK_ALT) THEN  ' Shift-Ctrl-Alt Right-click
            IF KbdT.SCAUp(107)= "." THEN GOSUB SetCurs            ' Set Cursor?
            IF UUCASE(KbdT.SCAData(107)) <> "(NULL)" AND UUCASE(KbdT.SCAData(107)) <> "(PASSTHRU)" THEN
               gKeyChr = CHR$(03) + KbdT.SCAData(107)             ' Pretend the selected key was hit
               MainBUserMsg()                                     ' Process the key event
            END IF                                                '

         ELSEIF (j AND (%MK_SHIFT OR %MK_ALT)) = (%MK_SHIFT OR %MK_ALT) THEN  ' Shift-Alt Right-click
            IF KbdT.SAUp(107)= "." THEN GOSUB SetCurs             ' Set Cursor?
            IF UUCASE(KbdT.SAData(107)) <> "(NULL)" AND UUCASE(KbdT.SAData(107)) <> "(PASSTHRU)" THEN
               gKeyChr = CHR$(03) + KbdT.SAData(107)              ' Pretend the selected key was hit
               MainBUserMsg()                                     ' Process the key event
            END IF                                                '

         ELSEIF (j AND (%MK_CONTROL OR %MK_ALT)) = (%MK_CONTROL OR %MK_ALT) THEN  ' Control-Alt Right-click
            IF KbdT.CAUp(107)= "." THEN GOSUB SetCurs             ' Set Cursor?
            IF UUCASE(KbdT.CAData(107)) <> "(NULL)" AND UUCASE(KbdT.CAData(107)) <> "(PASSTHRU)" THEN
               gKeyChr = CHR$(03) + KbdT.CAData(107)              ' Pretend the selected key was hit
               MainBUserMsg()                                     ' Process the key event
            END IF                                                '

         ELSEIF (j AND (%MK_SHIFT OR %MK_CONTROL)) = (%MK_SHIFT OR %MK_CONTROL) THEN  ' Shift-Control Right-click
            IF KbdT.SCUp(107)= "." THEN GOSUB SetCurs             ' Set Cursor?
            IF UUCASE(KbdT.SCData(107)) <> "(NULL)" AND UUCASE(KbdT.SCData(107)) <> "(PASSTHRU)" THEN
               gKeyChr = CHR$(03) + KbdT.SCData(107)              ' Pretend the selected key was hit
               MainBUserMsg()                                     ' Process the key event
            END IF                                                '

         ELSEIF (j AND %MK_SHIFT) = %MK_SHIFT THEN                ' Shift Right-Click
            IF KbdT.SUp(107)= "." THEN GOSUB SetCurs              ' Set Cursor?
            IF UUCASE(KbdT.SData(107)) <> "(NULL)" AND UUCASE(KbdT.SData(107)) <> "(PASSTHRU)" THEN
               gKeyChr = CHR$(03) + KbdT.SData(107)               ' Pretend the selected key was hit
               MainBUserMsg()                                     ' Process the key event
            END IF                                                '

         ELSEIF (j AND %MK_CONTROL) = %MK_CONTROL THEN            ' Control Right-click
            IF KbdT.CUp(107)= "." THEN GOSUB SetCurs              ' Set Cursor?
            IF UUCASE(KbdT.CData(107)) <> "(NULL)" AND UUCASE(KbdT.CData(107)) <> "(PASSTHRU)" THEN
               gKeyChr = CHR$(03) + KbdT.CData(107)               ' Pretend the selected key was hit
               MainBUserMsg()                                     ' Process the key event
            END IF                                                '

         ELSEIF (j AND %MK_ALT) = %MK_ALT THEN                    ' Alt Right-click
            IF KbdT.AUp(107)= "." THEN GOSUB SetCurs              ' Set Cursor?
            IF UUCASE(KbdT.AData(107)) <> "(NULL)" AND UUCASE(KbdT.AData(107)) <> "(PASSTHRU)" THEN
               gKeyChr = CHR$(03) + KbdT.AData(107)               ' Pretend the selected key was hit
               MainBUserMsg()                                     ' Process the key event
            END IF                                                '

         ELSE                                                     ' Normal Right-Click
            IF KbdT.NUp(107) = "." THEN GOSUB SetCurs             ' Set Cursor?
            IF UUCASE(KbdT.NData(107)) <> "(NULL)" AND UUCASE(KbdT.NData(107)) <> "(PASSTHRU)" THEN
               gKeyChr = CHR$(03) + KbdT.NData(107)               ' Pretend the selected key was hit
               MainBUserMsg()                                     ' Process the key event
            END IF                                                '

         END IF                                                   '

      '----- Throw all else away
      CASE ELSE
         GOTO DlgExit                                             '
   END SELECT                                                     '

DlgExit:                                                          '
   EXIT FUNCTION

SetCurs:
   GOSUB GetRowCol                                                ' Get Row/Col to tRow/tCol
   TP.CsrCol = tCol: TP.CsrRow = tRow                             ' Set where cursor goes
   IF ISFALSE IsFMTab AND TP.CsrCol = gLNPadCol AND TP.CsrRow > 2 THEN TP.CsrCol = TP.CsrCol + 1 ' Skip column gLNPadCol
   IF TP.CsrRow = 1 AND TP.CsrCol < 10 + pCmdLen AND ISNULL(TRIM$(TP.pCommand)) THEN  ' In Command area, and Cmd blank?
      TP.MarkKill                                                 '
      SetCmd                                                      ' Cursor to Cmd line
   END IF
   IF ISFALSE IsFMTab THEN                                        ' If not FM                                                           '
      '----- In the Line number area
      IF TP.CsrCol < gLNPadCol AND TP.CsrRow > 2 + TP.PrfCols AND _ ' Make sure its safe
         TP.CsrRow < gwScrHeight + 1 AND TP.PrfHexMode = &1 AND ISFALSE IsFMTab THEN ' In true text area?
         i = TP.sGet(TP.CsrRow)                                   '
         IF TP.LLCtlGet(i) = $BlankLNo OR TP.LLCtlGet(i) = "''''''''" OR _ ' Useless data?
            TP.LLCtlGet(i) = "--------" OR TP.LLCtlGet(i) = "********" OR _
            LEFT$(TP.LLCtlGet(i), 1) = "=" OR VERIFY(TP.LLCtlGet(i), $Numeric) = 0 THEN
            TP.CsrCol = 1                                         ' Put cursor in column 1
         END IF                                                   '
      END IF                                                      '
   END IF                                                         '
   TP.MarkLineNumb                                                ' Hilite the line
   sDoCursor                                                      ' Go move the cursor
   RETURN

GetRowCol:
   IF LO(WORD,CBLPARAM) = 0 THEN GOTO DlgExit                     ' If invalid location, ignore
   IF HI(WORD,CBLPARAM) = 0 THEN GOTO DlgExit                     '
   tCol = CEIL((LO(WORD,CBLPARAM) - %GLM + 2) / gFontWidth)       ' Calc where cursor is
   tRow = CEIL((HI(WORD,CBLPARAM) + 1) / gFontHeight)             '
   RETURN                                                         '

PostAttention:                                                    ' Post an Attention request
   Msg.kbString = CHR$(4, 0, 0, 0)                                ' Setup as an Attention request message
   PostMessage(hWnd, %WM_USER, Msg.MsgwParam, 0)                  ' To the mainline Callback routine
   RETURN                                                         '

END FUNCTION

CALLBACK FUNCTION DlgDefaultCallBack
'--------------------
' Callback function used by the Profile Default Dialog
'--------------------
LOCAL lclText AS STRING
   SELECT CASE AS LONG CB.MSG                                     '

      '----- Handle the various interactions

      '----- SYSCOMMAND
      CASE %WM_SYSCOMMAND
         IF CB.HNDL <> hDef THEN EXIT FUNCTION                    '
         IF (CBWPARAM AND &HFFF0) = %SC_CLOSE THEN                ' Trap the [x] button and Alt-F4
            FUNCTION = 1                                          ' Ignore it
         END IF                                                   '

      CASE %WM_COMMAND                                            '
         SELECT CASE AS LONG CB.CTL                               '
            CASE %WELCOME_OPTION1                                 '
               IF CB.CTLMSG = %BN_CLICKED THEN                    '
                  gDefaultAnswer = "DEF"                          ' Say to use the DEFAULT profile
                  RESET hToolTips                                 ' Clear so Tooltips get re-created correctly
                  DIALOG END hDef                                 ' Shut things down
               END IF                                             '

            CASE %WELCOME_OPTION2                                 '
               IF CB.CTLMSG = %BN_CLICKED THEN                    '
                  gDefaultAnswer = "NEW"                          ' Say to create a NEW  profile
                  RESET hToolTips                                 ' Clear so Tooltips get re-created correctly
                  DIALOG END hDef                                 ' Shut things down
               END IF                                             '

            CASE %WELCOME_OPTION3                                 '
               IF CB.CTLMSG = %CBN_SELENDOK THEN                  ' Something selected?
                  CONTROL GET TEXT hDef, %WELCOME_OPTION3 TO lclText '
                  IF lclText = "(Choose)" THEN EXIT FUNCTION      ' If nothing chosen, ignore
                  gDefaultAnswer = "USE " + lclText               ' Say to create a USE the profile
                  RESET hToolTips                                 ' Clear so Tooltips get re-created correctly
                  DIALOG END hDef                                 ' Shut things down
               END IF                                             '

            CASE %WELCOME_OPTION4                                 '
               IF CB.CTLMSG = %BN_CLICKED THEN                    '
                  gDefaultAnswer = "SKIP"                         ' Say to Skip this file type
                  RESET hToolTips                                 ' Clear so Tooltips get re-created correctly
                  DIALOG END hDef                                 ' Shut things down
               END IF                                             '

            CASE %WELCOME_OPTION5                                 '
               IF CB.CTLMSG = %BN_CLICKED THEN                    '
                  gDefaultAnswer = "ABORT"                        ' Say to Abort the FF
                  RESET hToolTips                                 ' Clear so Tooltips get re-created correctly
                  DIALOG END hDef                                 ' Shut things down
               END IF                                             '

         END SELECT                                               '
   END SELECT                                                     '
   EXIT FUNCTION
END FUNCTION

CALLBACK FUNCTION DlgIntruptCallBack
'--------------------
' Callback function used by the FF search Interrupt
'--------------------
   SELECT CASE AS LONG CB.MSG                                     '
      '----- Handle the FF Interrupt
      CASE %WM_COMMAND                                            '
         SELECT CASE AS LONG CB.CTL                               '
            CASE %INTERRUPT_BREAK                                 '
               IF CB.CTLMSG = %BN_CLICKED THEN                    ' Break?
                  gfInterrupt = %True                             ' Flag for mainline
                  DIALOG END hIntr                                ' End the window
               END IF                                             '
         END SELECT                                               '
   END SELECT
END FUNCTION

CALLBACK FUNCTION DlgMessageCallBack
'---------- Make sure Message box text starts off unselected
LOCAL mx, my AS LONG
   SELECT CASE CBMSG
      CASE %WM_INITDIALOG
         CONTROL HANDLE hMsg, %MSG_ListBox TO hMsgHnd
      CASE %WM_PAINT
         SendMessage hMsgHnd, %EM_SETSEL, -1, 0
      CASE %WM_SIZE
         DIALOG GET SIZE hMsg TO mx, my
         CONTROL SET SIZE hMsg, %MSG_ListBox, mx - 4, my - 30
         FUNCTION = 0
   END SELECT
END FUNCTION

CALLBACK FUNCTION DlgOptionsCallBack
'--------------------
' Callback function used by Options dialog
'--------------------
LOCAL LclKey, KeyMode, KeyCmd, ttxt, t, KeyMsg, Txt1, Txt2 AS STRING, lclScroll AS STRING * 4
LOCAL cTxt AS LONG POINTER
LOCAL i, j, k, m, CurrTab, hDC, HiLiteErr AS LONG
STATIC lclENGChars AS LONG
LOCAL lclFontname, VerData AS STRING, lclFontPitch, lclStyle AS INTEGER

   SELECT CASE AS LONG CB.MSG                                     '

      '----- Dialog is being initialized
      CASE %WM_INITDIALOG                                         '
         lclENGChars = ENV.ENGChars                               ' Save starting ENG Char value

      '----- Fill TAB background
      CASE %WM_ERASEBKGND                                         ' Time to fill Tab background
         hDC = CB.WPARAM                                          ' Get hDC
         sTabBGFill hDC                                           ' Pass to fill routine
         FUNCTION = 1                                             ' Say we did it

      '----- Draw Tab Item Text
      CASE %WM_DRAWITEM                                           ' draw tab item text in different color on select
         IF CB.WPARAM = %TAB_CONTROL THEN  sTabHighLight CBHNDL, CBWPARAM, CBLPARAM

      '----- Handle various interactions
      CASE %WM_COMMAND                                            '
         SELECT CASE AS LONG CB.CTL                               '

            '----- DONE button pressed
            CASE %DIALOG_DONE                                     '
               IF CB.CTLMSG = %BN_CLICKED THEN                    ' Done button pressed

                  '----- Validate Cmd Character
                  CONTROL GET TEXT hGeneral, %GENERAL_INICmdChr TO Txt1
                  Txt1 = LEFT$(Txt1, 1)              '
                  IF LEN(Txt1) = 0 OR Txt1 = " " THEN
                     sDoMsgBox("Command separator character cannot be null or blank.", %MB_OK OR %MB_USERICON, "Options")
                     EXIT FUNCTION                                '
                  END IF                                          '

                  ' Validate CMDFlags
                  CONTROL GET TEXT hSubmit, %SUBMIT_INICMDFlags TO ttxt
                  GOSUB FlagCheck
                  CONTROL SET TEXT hSubmit, %SUBMIT_INICMDFlags, ttxt

                  ' Validate RUNFlags
                  CONTROL GET TEXT hSubmit, %SUBMIT_INIRUNFlags TO ttxt
                  GOSUB FlagCheck
                  CONTROL SET TEXT hSubmit, %SUBMIT_INIRUNFlags, ttxt

                  ' Validate the FM Layout string
                  CONTROL GET TEXT hFManager, %FMANAGER_INIFMLayout TO ttxt
                  ttxt = UUCASE(REMOVE$(ttxt, " "))               ' Remove any blanks, uppercase it
                  ttxt = RTRIM$(ttxt, ",")                        ' Remove trailing comma
                  FOR i = 1 TO PARSECOUNT(ttxt, ",")              ' Check the operands
                     t = LSET$(PARSE$(ttxt, i), 9)                ' Get operand padded
                     IF INSTR("EXT      SIZESHORTSIZELONG NOTE     LINES    DATETIME DATE     ", t) = 0 THEN _
                        sDoMsgBox("FM Layout operand: |K" + TRIM$(t) + "|B is not recognixed.", %MB_OK OR %MB_USERICON, "Options"): EXIT FUNCTION
                  NEXT i                                          '
                  IF TALLY(ttxt, "EXT")   > 1 OR _                ' Look for duplicates
                     TALLY(ttxt, "SIZE")  > 1 OR _                '
                     TALLY(ttxt, "NOTE")  > 1 OR _                '
                     TALLY(ttxt, "DATE")  > 1 OR _                '
                     TALLY(ttxt, "LINES") > 1 OR _                '
                     TALLY(ttxt, "EXT")   > 1 THEN _              '
                        sDoMsgBox("FM Layout duplicate operand(s) detected.", %MB_OK OR %MB_USERICON, "Options"): EXIT FUNCTION
                  CONTROL SET TEXT hFManager, %FMANAGER_INIFMLayout,ttxt

                  ' Validate the SBAR Layout string
                  CONTROL GET TEXT hSBAR, %SBAR_INISBOrder TO ttxt
                  ttxt = UUCASE(REMOVE$(ttxt, " "))               ' Remove any blanks, uppercase it
                  ttxt = RTRIM$(ttxt, ",")                        ' Remove trailing comma
                  FOR i = 1 TO PARSECOUNT(ttxt, ",")              ' Check the operands
                     t = LSET$(PARSE$(ttxt, i), 6)                ' Get operand padded
                     IF INSTR("MODE  LINNO LINES COLS  BNDS  INSOVRCASE  CHANGESTATE MISC  SELECTCAPS  SOURCEEOL   ", t) = 0 THEN _
                        sDoMsgBox("Status Bar Layout operand: |K" + TRIM$(t) + "|B is not recognixed.", %MB_OK OR %MB_USERICON, "Options"): EXIT FUNCTION
                  NEXT i                                          '
                  IF TALLY(ttxt, "MODE")   > 1 OR _               ' Look for duplicates
                     TALLY(ttxt, "LINNO")  > 1 OR _               '
                     TALLY(ttxt, "LINES")  > 1 OR _               '
                     TALLY(ttxt, "COLS")   > 1 OR _               '
                     TALLY(ttxt, "BNDS")   > 1 OR _               '
                     TALLY(ttxt, "INSOVR") > 1 OR _               '
                     TALLY(ttxt, "CASE")   > 1 OR _               '
                     TALLY(ttxt, "CHANGE") > 1 OR _               '
                     TALLY(ttxt, "STATE")  > 1 OR _               '
                     TALLY(ttxt, "MISC")   > 1 OR _               '
                     TALLY(ttxt, "SELECT") > 1 OR _               '
                     TALLY(ttxt, "CAPS")   > 1 OR _               '
                     TALLY(ttxt, "SOURCE") > 1 OR _               '
                     TALLY(ttxt, "EOL")    > 1 THEN _             '
                        sDoMsgBox("Status Bar Layout duplicate operand(s) detected.", %MB_OK OR %MB_USERICON, "Options"): EXIT FUNCTION
                  CONTROL SET TEXT hSBAR, %SBAR_INISBOrder,ttxt   '
                  CONTROL GET TEXT hGeneral, %GENERAL_INICharset TO Txt2
                  CONTROL GET TEXT hGeneral, %GENERAL_INIInvChar TO Txt1
                  IF VERIFY(Txt1, Txt2) <> 0 THEN _               ' Must choose a visible character
                     sDoMsgBox("Chosen substitution character is not present in the valid character list. ", %MB_OK OR %MB_USERICON, "Options"): EXIT FUNCTION

                  '----- All seems OK, save all the option values
                  CONTROL GET CHECK hGeneral, %GENERAL_INIABeepFlag TO i                   : ENV.ABeepFlag = i
                  CONTROL GET CHECK hGeneral, %GENERAL_INIVBeepFlag TO i                   : ENV.VBeepFlag = i
                  CONTROL GET CHECK hGeneral, %GENERAL_INIUniqueFlag TO i                  : ENV.UniqueFlag = i
                  CONTROL GET CHECK hFManager, %FMANAGER_INIFMCloseFlag TO i               : ENV.FMCloseFlag = i
                  CONTROL GET CHECK hFManager, %FMANAGER_FMDATEHILITE TO i                 : ENV.FMDateHiLite = i
                  CONTROL GET CHECK hGeneral, %GENERAL_INIResetU TO i                      : ENV.ResetU = i
                  CONTROL GET CHECK hGeneral, %GENERAL_INIDirProfFlag TO i                 : ENV.DirProfFlag = i
                  CONTROL GET CHECK hFManager, %FMANAGER_INIConfDelFlag TO i               : ENV.ConfDelFlag = i
                  CONTROL GET CHECK hGeneral, %GENERAL_INIUseRecycle TO i                  : ENV.UseRecycle = i
                  CONTROL GET CHECK hGeneral, %GENERAL_INIViewWarn TO i                    : ENV.ViewWarn = i
                  CONTROL GET CHECK hGeneral, %GENERAL_INISplash TO i                      : ENV.Splash = i
                  CONTROL GET CHECK hFManager, %FMANAGER_FMHLFLAG TO i                     : ENV.FMHelpFlag = i
                  CONTROL GET TEXT hFManager, %FMANAGER_INIFMLAYOUT TO Txt1                : ENV.FMLayout = Txt1
                  CONTROL GET TEXT hFManager, %FMANAGER_INIFMFFLIST TO Txt1                : ENV.FMFFList = Txt1
                  CONTROL GET TEXT hSBAR, %SBAR_INISBOrder TO Txt1                         : ENV.SBLayout = Txt1
                  COMBOBOX GET SELECT hGeneral, %GENERAL_INIMinRetrieve TO i               : ENV.MinRetrieve = i
                  CONTROL GET TEXT hGeneral, %GENERAL_INICmdChr TO Txt1                    : ENV.CmdChr = LEFT$(Txt1, 1)
                  CONTROL GET CHECK hGeneral, %GENERAL_INIReOpenLast TO i                  : ENV.ReOpenLast = i
                  CONTROL GET CHECK hGeneral, %GENERAL_ALLOW2D TO i                        : ENV.Allow2DMouse = i

                  CONTROL GET TEXT hGeneral, %GENERAL_INICharSet TO Txt1                   : ENV.CharSet = Txt1
                  CONTROL GET CHECK hGeneral, %GENERAL_INIENGChars TO i                    : ENV.ENGChars = i

                  IF ENV.ENGChars THEN                            ' ENGLIGH ONLY ...
                     gUpper = $Upper: gLower = $Lower
                  ELSE                                            ' use international letters ...
                     gUpper = $Upper + $UpperNat: gLower = $Lower + $LowerNat
                  END IF

                  ENV.CharSet = REMOVE$(ENV.CharSet, " ")         ' Charset should not have blanks
                  Txt1 = ENV.CharSet                              ' Massage into RegEx string
                  REPLACE "\" WITH "\x5C" IN Txt1                 '
                  REPLACE "[" WITH "\x5B" IN Txt1                 '
                  REPLACE "]" WITH "\x5D" IN Txt1                 '
                  ENV.CharSetR = Txt1                             '
                  IF ISFALSE ENV.FirstTime THEN TP.PicSetWord(): TP.PicSetCharSet()' Reset Picture strings

                  CONTROL GET TEXT hGeneral, %GENERAL_INIInvChar TO Txt1
                  IF ISNULL(Txt1) THEN Txt1 = "."
                  ENV.InvChar = Txt1

                  COMBOBOX GET SELECT hGeneral, %GENERAL_ININotify TO i                    : ENV.NotifyLevel = i
                  ENV.NotifyLevel = SWITCH(ENV.NotifyLevel = 1, 0, ENV.NotifyLevel = 2, 1, ENV.NotifyLevel = 3, 2)
                  ENV.NotifyLevelT = ENV.NotifyLevel
                  CONTROL GET CHECK hGeneral, %GENERAL_FINDWORD TO i                       : ENV.FindWord = i
                  CONTROL GET TEXT hFManager, %FMANAGER_INIRecentCtr TO Txt1               : ENV.RecentCtr = VAL(Txt1)
                  CONTROL GET TEXT hFManager, %FMANAGER_INIFMLCmdWidth TO Txt1
                  i = VAL(Txt1): i = MAX(5, i): i = MIN(i, 20): Txt1 = FORMAT$(i)          : ENV.FMLCmdWidth = i
                  CONTROL GET TEXT hFManager, %FMANAGER_INIFMNoteWidth TO Txt1
                  i = VAL(Txt1): i = MAX(5, i): i = MIN(i, 20): Txt1 = FORMAT$(i)          : ENV.FMNoteWidth = i
                  CONTROL GET TEXT hGeneral, %GENERAL_INIQuickRenum TO Txt1                : ENV.QuickRenum = MAX(VAL(Txt1), 0)
                  CONTROL GET TEXT hGeneral, %GENERAL_INIDefDataShift TO Txt1              : ENV.DefDataShift = MAX(VAL(Txt1), 1)
                  CONTROL GET TEXT hGeneral, %GENERAL_INIAutoScroll TO Txt1
                  ENV.AutoScroll = VAL(Txt1): ENV.AutoScroll= MAX(0, ENV.AutoScroll)     ' Keep it zero or positive
                  CONTROL GET TEXT hGeneral, %GENERAL_INILinRepeatMax TO Txt1              : ENV.LinRepeatMax = VAL(Txt1)
                  CONTROL GET TEXT hSubmit, %SUBMIT_SUBMIT TO Txt1                         : ENV.SubmitCmd = Txt1
                  CONTROL GET TEXT hSubmit, %SUBMIT_INISubmitDir TO Txt1                   :ENV.SubmitDir = TRIM$(Txt1)
                  IF RIGHT$(ENV.SubmitDir, 1) = "\" THEN ENV.SubmitDir = LEFT$(ENV.SubmitDir, LEN(ENV.SubmitDir) - 1)
                  CONTROL GET TEXT hSubmit, %SUBMIT_INICMDFlags TO Txt1                    : ENV.CmdFlags = Txt1
                  CONTROL GET TEXT hSubmit, %SUBMIT_INIRUNFlags TO Txt1                    : ENV.RunFlags = Txt1
                  CONTROL GET TEXT hSubmit, %SUBMIT_INIInclKey TO Txt1                     : ENV.SubmitInclKey = TRIM$(Txt1)
                  CONTROL GET TEXT hSubmit, %SUBMIT_INIInclCol TO Txt1                     : ENV.SubmitInclCol = VAL(TRIM$(Txt1))
                  CONTROL GET CHECK hScreen, %SCREEN_INIHRuler TO i                        : ENV.HRuler = i
                  CONTROL GET CHECK hScreen, %SCREEN_INIVRuler TO i                        : ENV.VRuler = i
                  CONTROL GET CHECK hScreen, %SCREEN_INIBanding TO i                       : ENV.Banding = i
                  CONTROL GET CHECK hScreen, %SCREEN_INIVERTINSCURS TO i                   : ENV.VertInsCurs = i
                  CONTROL GET CHECK hScreen, %SCREEN_INICMDHILITE TO i                     : ENV.CmdHiLite = i
                  CONTROL GET TEXT hScreen, %SCREEN_CURSHGHTNORM TO Txt1
                  ENV.CursNormal = VAL(Txt1): ENV.CursNormal = MIN(ENV.CursNormal, 100)
                  CONTROL GET TEXT hScreen, %SCREEN_CURSHGHTINS TO Txt1
                  ENV.CursInsert = VAL(Txt1): ENV.CursInsert = MIN(ENV.CursInsert, 100)
                  CONTROL GET TEXT hScreen, %SCREEN_INIFontPitch TO Txt1
                  ENV.FontPitch = VAL(Txt1): ENV.FontPitch = MAX(ENV.FontPitch, 5): ENV.FontPitch = MIN(ENV.FontPitch, 24)
                  CONTROL GET TEXT hScreen, %SCREEN_INIFontName TO Txt1                    : ENV.FontName = Txt1
                  CONTROL GET TEXT hScreen, %SCREEN_INIPFKShow TO Txt1                     : ENV.PFKShow = MIN(VAL(Txt1), 5)
                  CONTROL GET TEXT hScreen, %SCREEN_INILNOSIZE TO Txt1
                  ENV.LinNoSize = MAX(VAL(Txt1), 5)
                  ENV.LinNoSize = MIN(ENV.LinNoSize, 8)
                  CONTROL GET CHECK hKeyBoard, %KBOARD_KBSCROLLH TO i                      : ENV.KBHorScroll = i
                  CONTROL GET CHECK hKeyBoard, %KBOARD_KBSCROLLV TO i                      : ENV.KBVerScroll = i
                  CONTROL GET CHECK hKeyBoard, %KBOARD_INIInsMode TO i                     : ENV.InsMode = i
                  CONTROL GET CHECK hKeyBoard, %KBOARD_INIInsReset TO i                    : ENV.InsReset = i
                  COMBOBOX GET SELECT hFManager, %FMANAGER_INILastDirOpen TO i             : ENV.LastDirOpen = i
                  ENV.SchemeSaveAll                               ' Save the Schemes and Hilites
                  '----- Change the running FM Tab values
                  CurrTab = TP.PgNumber                           ' Remember active tab
                  TP = Tabs(1)                                    ' Switch to the FM tab
                  TP.LastTop = 0                                  ' Kill position
                  TP = Tabs(CurrTab)                              ' Switch tab back
                  RESET hToolTips                                 ' Clear so Tooltips get re-created correctly
                  DIALOG END hOpt                                 ' Shut things down
               END IF                                             '

            '----- CANCEL button pressed
            CASE %DIALOG_CANCEL                                   '
               IF CB.CTLMSG = %BN_CLICKED THEN                    ' Cancel button pressed
                  gfOptCancel = %True                             ' Tell Options command we cancelled out
                  RESET hToolTips                                 ' Clear so Tooltips get re-created correctly
                  DIALOG END hOpt                                 ' Shut things down
               END IF                                             '


            '----- Font choose button pressed
            CASE %SCREEN_FONT_CHOOSE                              ' Manual choose font?
               IF CB.CTLMSG = %BN_CLICKED THEN                    ' Choose button pressed
                  sPopReady                                       ' Ready for pop-up
                  DISPLAY FONT hWnd, , , ENV.FontName, ENV.FontPitch, ENV.FontStyle, %CF_FIXEDPITCHONLY + %CF_SCREENFONTS _
                               TO lclFontName, lclFontPitch, lclStyle
                  sPopReset                                       ' Reset popup state
                  IF LEN(TRIM$(lclFontName)) <> 0 THEN            '
                      ENV.FontName = lclFontName + IIF$(lclStyle = 1, " Bold", "")
                      ENV.FontPitch = lclFontPitch: ENV.FontStyle = lclStyle ' Save used values
                      CONTROL SET TEXT hScreen, %SCREEN_INIFontName, ENV.FontName
                      CONTROL SET TEXT hScreen, %SCREEN_INIFontPitch, FORMAT$(ENV.FontPitch)
                  END IF                                          '
               END IF                                             '

            '----- Is it the Mark Line color
            CASE %SCREEN_INIcMarkLine
               IF CB.CTLMSG = %BN_CLICKED THEN                    ' Button pressed
                  k = sSelectColor(0, ENV.cMarkLine)              ' Get a new value
                  ENV.cMarkLine = k                               ' Save it
                  GRAPHIC ATTACH hScreen, %SCREEN_INIcMarkLine    '
                  GRAPHIC BOX (1, 1) - (12, 12), 20, %BLACK, k, 0 '
                  CONTROL REDRAW hScreen, %SCREEN_INIcMarkLine    '
               END IF                                             '

            '----- One of the Screen color boxes was clicked
            CASE %SCREEN_16_FG  TO %SCREEN_16_FG + 40
               IF CB.CTLMSG = %BN_CLICKED THEN                    ' Button pressed
                  CONTROL GET USER hScreen, CB.CTL, 1 TO i        ' Get the 1st index of the Color field
                  CONTROL GET USER hScreen, CB.CTL, 2 TO j        ' Get the 2nd index of the Color field
                  k = sSelectColor(0, ENV.GetClr(i, j))           ' Get a new value
                  ENV.SetClr(i, j, k)                             ' Save it
                  CONTROL SET COLOR  hScreen, %SCREEN_16_TEXT + ((i - 16) * 4), ENV.GetClr(i, 1), ENV.GetClr(i, 2)
                  CONTROL REDRAW hScreen, %SCREEN_16_TEXT + ((i - 16) * 4)
                  GRAPHIC ATTACH hScreen, %SCREEN_16_FG + (j-1) + ((i - 16) * 4)
                  GRAPHIC BOX (1, 1) - (12, 12), 20, %BLACK, ENV.GetClr(i, j), 0
                  CONTROL REDRAW hScreen, %SCREEN_16_FG + (j-1) + ((i - 16) * 4)
               END IF                                             '

            '----- One of the Scheme color boxes was clicked
            CASE %SCHEME_00_FG  TO (%SCHEME_00_FG + 128)
               IF CB.CTLMSG = %BN_CLICKED THEN                    ' Button pressed
                  CONTROL GET USER hScheme, CB.CTL, 1 TO i        ' Get the 1st index of the Color field
                  CONTROL GET USER hScheme, CB.CTL, 2 TO j        ' Get the 2nd index of the Color field
                  k = sSelectColor(0, ENV.GetClr(i, j))           ' Get a new value
                  ENV.SetClr(i, j, k)                             ' Save it
                  CONTROL SET COLOR  hScheme, %SCHEME_00_TEXT + (i * 4), ENV.GetClr(i, 1), ENV.GetClr(i, 2)
                  CONTROL REDRAW hScheme, %SCHEME_00_TEXT + (i * 4)
                  GRAPHIC ATTACH hScheme, %SCHEME_00_FG + (j-1) + (i * 4)
                  GRAPHIC BOX (1, 1) - (12, 12), 20, %BLACK, ENV.GetClr(i, j), 0
                  CONTROL REDRAW hScheme, %SCHEME_00_FG + (j-1) + (i * 4)
               END IF                                             '

            '----- One of the HiLites color boxes was clicked
            CASE %HILITES_01_FG TO (%HILITES_01_FG + 45)
               IF CB.CTLMSG = %BN_CLICKED THEN                    ' Button pressed
                  CONTROL GET USER hHiLites, CB.CTL, 1 TO i       ' Get the passed Scheme Number
                  CONTROL GET USER hHiLites, CB.CTL, 2 TO j       ' Get the passed FG/BG1/BG2 value
                  CONTROL GET USER hHiLites, CB.CTL, 3 TO k       ' Get the passed Text ID User value
                  m = sSelectColor(0, ENV.GetClr(i, j))           ' Get a new value
                  ENV.SetClr(i, j, m)                             ' Save it
                  CONTROL SET COLOR hHiLites, k, ENV.GetClr(i, 1), ENV.GetClr(i, 2)
                  CONTROL REDRAW hHiLites, k
                  GRAPHIC ATTACH hHiLites, CB.CTL
                  GRAPHIC BOX (1, 1) - (12, 12), 20, %BLACK, ENV.GetClr(i, j), 0
                  CONTROL REDRAW hHiLites, CB.CTL                 '
                  '----- If BG1, propagate to BG2 automatically
                  IF j = 2 THEN                                   ' If BG1, propagate to BG2
                     j = 3                                        ' Set to BG2
                     ENV.SetClr(i, j, m)                          ' Save BG1 color into BG2
                     GRAPHIC ATTACH hHiLites, CB.CTL + 1          ' Point at it
                     GRAPHIC BOX (1, 1) - (12, 12), 20, %BLACK, ENV.GetClr(i, j), 0
                     CONTROL REDRAW hHiLites, CB.CTL + 1          '
                  END IF                                          '
               END IF                                             '

            '----- We're in the ENGLISH-ONLY field
            CASE %GENERAL_INIENGChars                             ' ENGLISH-only checkbox ...
               IF CB.CTLMSG = %BN_CLICKED THEN                    ' Clicked?
                  GOSUB DoENGChars                                ' Go process it
               END IF                                             '

            '----- We're in the INI Def chars field
            CASE %GENERAL_INICharSet                              ' Def char field
               IF CB.CTLMSG = %EN_KILLFOCUS OR CB.CTLMSG = %BN_CLICKED THEN ' Losing focus or clicked?
                  GOSUB DoENGChars                                ' Go process it
               END IF                                             '
               IF CBCTLMSG = %EN_SETFOCUS THEN                    '
                  CONTROL SEND CBHNDL, CBCTL, %EM_SETSEL, 0, 0    '
               END IF                                             '


         END SELECT                                               '
   END SELECT                                                     '
   EXIT FUNCTION

DoENGChars:
   CONTROL GET CHECK hGeneral, %GENERAL_INIENGChars TO i          ' Get ENGLISH-only setting ...
   CONTROL GET TEXT hGeneral, %GENERAL_INICharSet TO tTxt         ' Get Charset
   IF tTxt = "" THEN                                              ' CharSet not yet defined?

      IF ENV.ENGChars THEN                                        ' Using ENGLISH only letters ?
         tTxt = $DefCharSet                                       ' Set it to include only ENGLISH
         gUpper = $Upper: gLower = $Lower                         '
      ELSE                                                        '
         tTxt = $DefCharSet + $UpperNat + $LowerNat               ' Set tTxt to include international chars
         gUpper = $Upper + $UpperNat: gLower = $Lower + $LowerNat '
      END IF                                                      '

      CONTROL SET TEXT hGeneral, %GENERAL_INICharSet, ttxt
   END IF

   IF lclENGChars AND ISFALSE i THEN                              ' ENGLISH-only setting is going ON  => OFF
      lclENGChars = i                                             ' Save the new 'last' value
      IF tTxt = $DefCharSet THEN                                  ' If we're still sitting at the default
         tTxt = $DefCharSet + $UpperNat + $LowerNat               ' Add the national to English
         gUpper = $Upper + $UpperNat: gLower = $Lower + $LowerNat '
         CONTROL SET TEXT hGeneral, %GENERAL_INICharSet, ttxt     '
      END IF                                                      '

   ELSEIF ISFALSE lclENGChars AND i THEN                          ' ENGLISH-only setting is going OFF => ON   gDefCharSet + $UpperNat + $LowerNat
      lclENGChars = i                                             ' Save the new 'last' value
      IF tTxt = $DefCharSet + $UpperNat + $LowerNat THEN          ' If we're still sitting at the stock Nat Char version
         tTxt = $DefCharSet                                       ' Remove the national, set to English only
         gUpper = $Upper: gLower = $Lower                         '
         CONTROL SET TEXT hGeneral, %GENERAL_INICharSet, ttxt     '
      END IF                                                      '

   END IF                                                         '

   RETURN

FlagCheck:
   ttxt = UUCASE(TRIM$(ttxt))                                     ' Get rid of white space
   i = INSTR(ttxt, "/K"): j = INSTR(ttxt, "/C")                   ' Test if /C or /K are there

   IF i <> 0 AND j <> 0 THEN                                      ' Both /C and /K ?
      ttxt = REMOVE$(ttxt, "/C")                                  ' Remove the /C
      i = INSTR(ttxt, "/K"): j = INSTR(ttxt, "/C")                ' Re-do the scan
   END IF                                                         '

   IF i = 0 AND j = 0 THEN                                        ' Neither?
      ttxt += " /K"                                               ' Default to /K
      ttxt = SHRINK$(ttxt)                                        ' remove extra spaces
      RETURN                                                      ' And we're done
   END IF                                                         '

   IF i <> 0 AND i <> LEN(ttxt) - 1 THEN                          ' /K and it's not at the end
      ttxt = REMOVE$(ttxt, "/K")                                  ' Remove the /K
      ttxt += " /K"                                               ' Add it at the end
   END IF

   IF j <> 0 AND j <> LEN(ttxt) - 1 THEN                          ' /C and it's not at the end
      ttxt = REMOVE$(ttxt, "/C")                                  ' Remove the /C
      ttxt += " /C"                                               ' Add it at the end
   END IF                                                         '
   ttxt = SHRINK$(ttxt)                                           ' Remove extra spaces
   RETURN

END FUNCTION

CALLBACK FUNCTION DlgPrintCallBack
'--------------------
' Callback function used by the PRINT FORMAT Dialog
'--------------------
LOCAL lclFontName, Txt1 AS STRING, lclFontPitch AS SINGLE, lclFontStyle AS LONG
LOCAL fIndex, i, j AS LONG, fList AS STRING, fTable() AS STRING
LOCAL cTxt AS LONG POINTER

   SELECT CASE AS LONG CB.MSG                                     '

      '----- Handle various interactions
      CASE %WM_COMMAND                                            '
         SELECT CASE AS LONG CB.CTL                               '

            '----- Choose Printer button pressed
            CASE %PRINT_PCHOOSE                                   '
               IF CB.CTLMSG = %BN_CLICKED THEN                    ' Choose Printer?
                  sPopReady                                       ' Ready for pop-up
                  XPRINT ATTACH CHOOSE, "SPFLite Printer Selection" ' Attach printer Choice menu
                  sPopReset                                       ' Reset popup state
                  IF ERR = 0 AND LEN(XPRINT$) > 0 THEN            ' OK?
                     XPRINT GET DUPLEX TO i                       ' Get duplex setting
                     ENV.PrtDuplex = i                            '
                     IF ENV.PrtDuplex = 0 THEN ENV.PrtDuplex = 1  ' Eliminate the 0 case
                     XPRINT GET ORIENTATION TO i                  ' Get orientation setting
                     ENV.PrtOrient = i                            '
                     ENV.PrtName = XPRINT$                        '
                     XPRINT GET PAPER TO i                        ' Get index to Paper type
                     ENV.PrtPaper = i                             '
                     XPRINT GET PAPERS TO fList                   ' Get string of forms available
                     REDIM fTable(1 TO PARSECOUNT(fList)) AS STRING ' Build a table
                     PARSE fList, fTable()                        '
                     FOR i = 1 TO UBOUND(fTable) STEP 2           '
                        IF VAL(fTable(i)) = ENV.PrtPaper THEN     ' Found our entry?
                           gPrtPaper = fTable(i + 1): EXIT FOR    ' Save paper name
                        END IF                                    '
                     NEXT i                                       '
                     CONTROL SET TEXT  hPrt, %PRINT_INIPrtName_TEXT, "Printer: " + ENV.PrtName
                     CONTROL SET TEXT  hPrt, %PRINT_INIPrtOrient_TEXT, "Orientation: " + IIF$(ENV.PrtOrient = 1, "Portrait", "Landscape")
                     CONTROL SET TEXT  hPrt, %PRINT_INIPrtPaper_TEXT, "Paper: " + gPrtPaper
                     XPRINT CANCEL                                '
                     XPRINT CLOSE                                 '
                  END IF                                          '
               END IF                                             '

            '----- Font choose button pressed
            CASE %PRINT_FCHOOSE                                   '
               IF CB.CTLMSG = %BN_CLICKED THEN                    ' Choose Font?
                  sPopReady                                       ' Ready for pop-up
                  DISPLAY FONT hWnd, , , ENV.PrtFontName, VAL(ENV.PrtFontPitch), VAL(ENV.PrtFontStyle), %CF_FIXEDPITCHONLY _
                          TO lclFontName, lclFontPitch, lclFontStyle
                  sPopReset                                       ' Reset popup state
                  IF LEN(TRIM$(lclFontName)) <> 0 THEN            '
                     ENV.PrtFontName = lclFontName + IIF$(lclFontStyle = 1, " BOLD", "") ' Swap
                     ENV.PrtFontPitch = FORMAT$(lclFontPitch, "00.0") '
                     ENV.PrtFontStyle = FORMAT$(lclFontStyle)     '
                     CONTROL SET TEXT  hPrt, %PRINT_FNAME_TEXT, "Fontname: " + ENV.PrtFontName
                     CONTROL SET TEXT  hPrt, %PRINT_FPOINT_TEXT, "Pitch: " + ENV.PrtFontPitch
                     CONTROL SET TEXT  hPrt, %PRINT_FBOLD_TEXT, "Bold: " + IIF$(ENV.PrtFontStyle = "0", "Off", "On")
                  END IF                                          '
               END IF                                             '

            '----- Metric toggle altered
            CASE %PRINT_METRIC
               IF CB.CTLMSG = %BN_CLICKED THEN                    ' Metric button pressed
                  CONTROL GET CHECK hPrt, %PRINT_METRIC TO i      '
                  ENV.PrtMetric = i                               '
                  CONTROL SET TEXT hPrt, %PRINT_LM_TEXT, "Page left margin " + IIF$(ENV.PrtMetric, "(mm)", "(in)")
                  CONTROL SET TEXT hPrt, %PRINT_RM_TEXT, "Page right margin " + IIF$(ENV.PrtMetric, "(mm)", "(in)")
                  CONTROL SET TEXT hPrt, %PRINT_TM_TEXT, "Page top margin " + IIF$(ENV.PrtMetric, "(mm)", "(in)")
                  CONTROL SET TEXT hPrt, %PRINT_BM_TEXT, "Page bottom margin " + IIF$(ENV.PrtMetric, "(mm)", "(in)")
                  CONTROL SET TEXT hPrt, %PRINT_LM, FORMAT$(IIF(ENV.PrtMetric, ENV.PrtLMargin * 25.4, ENV.PrtLMargin), "0.0")
                  CONTROL SET TEXT hPrt, %PRINT_RM, FORMAT$(IIF(ENV.PrtMetric, ENV.PrtRMargin * 25.4, ENV.PrtRMargin), "0.0")
                  CONTROL SET TEXT hPrt, %PRINT_TM, FORMAT$(IIF(ENV.PrtMetric, ENV.PrtTMargin * 25.4, ENV.PrtTMargin), "0.0")
                  CONTROL SET TEXT hPrt, %PRINT_BM, FORMAT$(IIF(ENV.PrtMetric, ENV.PrtBMargin * 25.4, ENV.PrtBMargin), "0.0")
               END IF                                             '

            '----- Banding color choice
            CASE %Print_BANDINGCOLOR                              ' A screen color choice
               IF CB.CTLMSG = %BN_CLICKED THEN                    ' Button pressed
                  j = ENV.PrtBandColor                            ' Get the current color
                  i = sSelectColor(0, j)                          ' Get a new value
                  IF ISFALSE ISWIN(hPrt) THEN EXIT FUNCTION       '
                  ENV.PrtBandColor = i                            ' Save it back
                  GRAPHIC ATTACH hPrt, CB.CTL                     ' Point at the control
                  GRAPHIC SET MIX %MIX_COPYSRC                    ' Set Mix mode
                  GRAPHIC BOX (1, 1) - (16, 16), 20, %BLACK, i, 0 ' Repaint it in the new color
               END IF                                             '

            '----- DONE button pressed
            CASE %PRINT_DONE                                      '
               IF CB.CTLMSG = %BN_CLICKED THEN                    ' Done button pressed
                  CONTROL GET CHECK hPrt, %PRINT_PCOLOR TO i
                  ENV.PrtPColor = i
                  CONTROL GET CHECK hPrt, %PRINT_BANDING TO i
                  ENV.PrtBanding = i
                  CONTROL GET CHECK hPrt, %PRINT_BANDLINES TO i
                  ENV.PrtBandLines = i
                  IF ENV.PrtBanding AND ENV.PrtBandLines THEN _
                        sDoMsgBox("Cannot have both Band stripes and Band Lines. ", %MB_OK OR %MB_USERICON, "Print Options"): EXIT FUNCTION
                  CONTROL GET CHECK hPrt, %PRINT_METRIC TO i      '
                  ENV.PrtMetric = i                               '
                  CONTROL GET TEXT hPrt, %PRINT_LM TO Txt1        '
                  ENV.PrtLMargin = IIF(ENV.PrtMetric, VAL(Txt1) / 25.4, VAL(Txt1)) '
                  CONTROL GET TEXT hPrt, %PRINT_RM TO Txt1        '
                  ENV.PrtRMargin = IIF(ENV.PrtMetric, VAL(Txt1) / 25.4, VAL(Txt1)) '
                  CONTROL GET TEXT hPrt, %PRINT_TM TO Txt1        '
                  ENV.PrtTMargin = IIF(ENV.PrtMetric, VAL(Txt1) / 25.4, VAL(Txt1)) '
                  CONTROL GET TEXT hPrt, %PRINT_BM TO Txt1        '
                  ENV.PrtBMargin = IIF(ENV.PrtMetric, VAL(Txt1) / 25.4, VAL(Txt1)) '
                  CONTROL GET CHECK hPrt, %PRINT_HEADER TO i
                  ENV.PrtHeader = i
                  CONTROL GET TEXT hPrt, %PRINT_HEADER_LEFT TO Txt1
                  ENV.PrtHeaderLeft = Txt1
                  CONTROL GET TEXT hPrt, %PRINT_HEADER_CENTER TO Txt1
                  ENV.PrtHeaderCenter = Txt1
                  CONTROL GET TEXT hPrt, %PRINT_HEADER_RIGHT TO Txt1
                  ENV.PrtHeaderRight = Txt1
                  CONTROL GET CHECK hPrt, %PRINT_FOOTER TO i
                  ENV.PrtFooter = i
                  CONTROL GET TEXT hPrt, %PRINT_FOOTER_LEFT TO Txt1
                  ENV.PrtFooterLeft = Txt1
                  CONTROL GET TEXT hPrt, %PRINT_FOOTER_CENTER TO Txt1
                  ENV.PrtFooterCenter = Txt1
                  CONTROL GET TEXT hPrt, %PRINT_FOOTER_RIGHT TO Txt1
                  ENV.PrtFooterRight = Txt1
                  RESET hToolTips                                 ' Clear so Tooltips get re-created correctly
                  DIALOG END hPrt                                 ' Shut things down
               END IF                                             '

         END SELECT                                               '
   END SELECT                                                     '
   EXIT FUNCTION
END FUNCTION

CALLBACK FUNCTION DlgProfileCallBack
'      (BYVAL hPrf AS DWORD,  _           ' CB.HNDL
'                                BYVAL wMsg AS LONG,   _           ' CB.MSG
'                                BYVAL wParam AS LONG, _           ' CB.WPARAM    CB.CTLMSG = HI(WORD, wParam) CB.CTL = LO(WORD, wParam)
'                                BYVAL lParam AS LONG) AS LONG     ' CB.LPARAM
'--------------------
' Callback function used by Profile Dialog
'--------------------
LOCAL LclKey, lclText AS STRING, lclScroll AS STRING * 4
LOCAL i, j, hDC AS LONG
LOCAL cTxt AS LONG PTR
LOCAL lclFontname AS STRING, lclFontPitch, lclStyle, lclLRECL AS INTEGER
LOCAL cData AS ClrCnvt                                            '

   SELECT CASE AS LONG CB.MSG                                     '

      '----- Fill TAB background
      CASE %WM_ERASEBKGND                                         ' Time to fill Tab background
         hDC = CB.WPARAM                                          ' Get hDC
         sTabBGFill hDC                                           ' Pass to fill routine
         FUNCTION = 1                                             ' Say we did it

      '----- Draw tab text
      CASE %WM_DRAWITEM                                           ' draw tab item text in different color on select
         IF CB.WPARAM = %TAB_CONTROL THEN  sTabHighLight CB.HNDL, CB.WPARAM, CB.LPARAM

      '----- Handle various interactions
      CASE %WM_COMMAND                                            '
         SELECT CASE AS CB.CTL                                    '

            '----- CANCEL pressed, bail out
            CASE %DIALOG_CANCEL                                   '
               IF CB.CTLMSG = %BN_CLICKED THEN                    ' Cancel button pressed
                  gfDialogDone = %True                            ' Say dialog was cancelled
                  RESET hToolTips                                 ' Clear so Tooltips get re-created correctly
                  DIALOG END hPrf                                 ' Shut things down
               END IF                                             '

            '----- DONE button pressed, validate and save
            CASE %DIALOG_DONE                                     '
               IF CB.CTLMSG = %BN_CLICKED THEN                    ' Done button pressed
                  IF TP.PrfDoneCheck() THEN                       ' Check validity, exit if errors
                     CONTROL SET FOCUS hPrf, %FILE_SCRLAMTC       ' Regain focus.
                     EXIT FUNCTION                                '
                  END IF                                          '

                  gfDialogDone = %False                           ' Say dialog ended normally
                  RESET hToolTips                                 ' Clear so Tooltips get re-created correctly
                  DIALOG END hPrf                                 ' Shut things down
               END IF                                             '

         END SELECT                                               '
   END SELECT                                                     '
END FUNCTION

FUNCTION MyMsgBox(mTxt AS STRING, OPT mStyle AS LONG, OPT mTitle AS STRING, OPT mpitch AS LONG, OPT mBG AS LONG) AS LONG
'---------- Private MSGBOX equivalent with color/Font size support
'mTxt     Text of the message
'         Default text color is %RGB_BLUE.
'         Change anywhere with embedded |x where
'         x can be: "K" - %RGB_BLACK       "G" - %RGB_GREEN       "B" - %RGB_BLUE
'                   "R" - %RGB_RED         "O" - %RGB_ORANGE      "I" - %RGB_INDIGO
'                   "P" - %RGB_DEEPPINK
'         e.g.  "This message has a |GGreen |Band |RRed |Bword in it"
'
'mStyle   Supports
'         %MB_OK                %MB_YESNO             %MB_OKCANCEL
'         %MB_RETRYCANCEL       %MB_YESNOCANCEL       %MB_ABORTRETRYIGNORE
'         %MB_CANCELTRYCONTINUE
'
'         %MB_ICONERROR         %MB_ICONSTOP          %MB_ICONHAND
'         %MB_ICONEXCLAMATION   %MB_ICONINFORMATION   %MB_ICONQUESTION
'         %MB_ICONWARNING       %MB_ICONASTERISK      %MB_USERICON
'
'         %MB_DEFBUTTON1        %MB_DEFBUTTON2        %MB_DEFBUTTON3
'mTitle   Text for popup box Title
'mPitch   Default size of Font for the message text
'
'         May be overridden FOR INDIVIDUAL lines by prefixing the line with |n
'             Where n = 1 - Standard default size
'                       2 - Standard * 1.25
'                       3 - Standard * 1.50
'                       4 - Standard * 1.75
'                       ...................
'                       9 - Standard * 3.00
'
'mBG      Color for message Background (Default is %RGB_BEIGE

REGISTER i AS LONG
REGISTER j AS LONG
LOCAL mlist(), t, t2, bTxt, lTitle AS STRING
LOCAL hmsg, hIcon, hFont() AS DWORD
LOCAL mlisth(), mlistf(), lResult, ChrX, ChrY, TxtX, TxtY, NumLines AS LONG
LOCAL x, y, w, h, clr, lStyle, lPitch, lBG, bID, bDef, bLoc, FontIX, yPos, yMax AS LONG
DIM hFont(10) AS DWORD

   '----- Get parameters and build fonts
   lStyle = IIF(ISMISSING(mStyle), %MB_OK OR %MB_USERICON, mStyle)' Get Style if provided
   lTitle = IIF$(ISMISSING(mTitle), "", mTitle)                   ' Get Title if provided
   lPitch = IIF(ISMISSING(mpitch), 11, mpitch)                    ' Get pitch if provided
   lBG    = IIF(ISMISSING(mBG), %RGB_BEIGE, mBG)                  ' Get BG color if provided
   FONT NEW "Segoe UI SemiBold", 12 TO hFont(0)                   ' Build our Font for button text
   FONT NEW "Segoe UI SemiBold", lPitch TO hFont(1)               ' Build our Base Font
   FontIX = 1                                                     ' Say Base font is current

   '----- Build an empty dialog to get font sizes
   DIALOG NEW PIXELS, 0, lTitle, 0, 0, 20, 20 TO hmsg             ' Create the Dialog
   DIALOG SET COLOR hMsg, %RGB_BLUE, lBG                          ' Set color
   CONTROL ADD GRAPHIC, hmsg, %MMB_msg, "", 0, 0, 20, 20          ' Add the graphic text area
   GRAPHIC ATTACH hmsg, %MMB_msg                                  ' Attach it
   GRAPHIC COLOR %RGB_BLUE, lBG                                   ' Set default color
   GRAPHIC SET FONT hfont(1)                                      ' Set it to our Base Font
   ChrX = GRAPHIC(CHR.SIZE.X)                                     ' Get the average character width
   ChrY = GRAPHIC(CHR.SIZE.Y)                                     ' Get the average character height

   '----- Calc width and line height based on the supplied text / Font
   NumLines = PARSECOUNT(mTxt, $CRLF)                             ' Get the Maximum line length
   REDIM mlist(NumLines): REDIM mlisth(NumLines): REDIM mlistf(NumLines)
   PARSE mTxt, mlist(), $CRLF                                     '
   FOR i = 0 TO NumLines                                          '
      t2 = LEFT$(mlist(i), 2)                                     'See if a Font request
      IF LEFT$(t2,1) = "|" AND INSTR("123456789", RIGHT$(t2,1)) <> 0 THEN ' Start with |n ?
         FontIX = VAL(RIGHT$(t2,1))                               ' Make it numeric
         mlistf(i) = FontIX                                       ' Save as FontIX for this line
         IF hFont(FontIX) = 0 THEN                                ' If we don't have it already
            x = lPitch * ((Fontix + 3) * .25)
            FONT NEW "Segoe UI SemiBold", lPitch * ((Fontix + 3) * .25) TO hFont(FontIX) ' Build a scaled version
         END IF                                                   '
         GRAPHIC SET FONT hfont(FontIX)                           ' Set it to the new font
         mlist(i) = MID$(mlist(i), 3)                             ' Remove /n
      END IF                                                      '
      j = GRAPHIC(TEXT.SIZE.X, mlist(i))                          '
      TxtX = MAX(j, TxtX)                                         '
      mlisth(i) = MAX(GRAPHIC(TEXT.SIZE.Y, mlist(i)), ChrY)       ' Save height of this line
      yMax += mlisth(i)                                           ' Accum total height
   NEXT i                                                         '

   TxtX += 80: TxtX = MAX(txtX, 340)                              ' Add 80 for Icon + Left/Right margins; ensure size for buttons
   TxtY = MAX(yMax + 10, 40)                                      ' Get total height for all lines - Minimum 40 for Icon

   '----- Resize now to actual size needed
   DIALOG SET CLIENT hmsg, TxtX, TxtY + 30                        ' Resize dialog (+ 30 for buttons)
   CONTROL SET SIZE hmsg, %MMB_msg, TxtX, TxtY                    ' Resize the text area

   '----- Print the text finally
   ypos = IIF(NumLines = 1, (40 - ChrY) / 2, 5)                   ' Set position of line 1
   FOR i = 0 TO NumLines                                          ' Print text now
      GRAPHIC SET POS (50, yPos)                                  ' Set position of this line
      GRAPHIC SET FONT hFont(mlistf(i))                           ' Set font for this line
      t = mlist(i)                                                ' Get a working copy
      IF LEN(t) = 0 THEN GRAPHIC PRINT: GOTO lExit                ' Skip null lines
      x = 1                                                       ' Start at the left
      DO WHILE x <= LEN(t)                                        ' While still stuff in t
         y = INSTR(x, t, "|")                                     ' Look for a color change
         IF y = 0 THEN GRAPHIC PRINT MID$(t, x);: GOTO lExit      ' No change, print remainder of line
         IF y > x THEN GRAPHIC PRINT MID$(t, x TO y - 1);         ' Print up to the |
         t2 = MID$(t, y + 1, 1)                                   ' Extract the color character
         clr = SWITCH(t2 = "O", %RGB_ORANGE, t2 = "B", %RGB_BLUE,     t2 = "R", %RGB_RED, _
                      t2 = "G", %RGB_GREEN,  t2 = "P", %RGB_DEEPPINK, t2 = "I", %RGB_INDIGO, _
                      t2 = "K", %RGB_BLACK)
         GRAPHIC COLOR clr, lBG                                   ' Set the new color
         x = y + 2                                                ' Step over |x
      LOOP                                                        ' Continue on this string
      lExit:
      yPos += mlisth(i)                                           ' Move Y down to nex line
   NEXT i                                                         ' Loop back for next line

   '----- Position the dialog
   DESKTOP GET CLIENT TO w, h                                     '
   x = (w - TxtX) / 2                                             '
   y = (h - (TxtY + 30)) / 2                                      '
   DIALOG SET LOC hMsg, x, y                                      '

   '----- Add whatever buttons are needed
   h = TxtY + 5
   i = lStyle AND &H0000000F                                      ' Isolate low nibble for button
   j = lStyle AND &H00000F00                                      ' Isolate 3rd nibble for button default
   IF i = %MB_YESNOCANCEL THEN
      IF j = %MB_DEFBUTTON1 OR j = 0 THEN                         ' Button 1 (or null request)
         bID = %MMB_Yes:      bTxt = "Yes":       bLoc = 10:  bDef = %True:  GOSUB BtnDef
         bID = %MMB_No:       bTxt = "No":        bLoc = 120: bDef = %False: GOSUB BtnDef
         bID = %MMB_Cancel:   bTxt = "Cancel":    bLoc = 230: bDef = %False: GOSUB BtnDef
      ELSEIF j = %MB_DEFBUTTON2 THEN                              ' Button 2
         bID = %MMB_Yes:      bTxt = "Yes":       bLoc = 10:  bDef = %False: GOSUB BtnDef
         bID = %MMB_No:       bTxt = "No":        bLoc = 120: bDef = %True:  GOSUB BtnDef
         bID = %MMB_Cancel:   bTxt = "Cancel":    bLoc = 230: bDef = %False: GOSUB BtnDef
      ELSE                                                        ' Button 3
         bID = %MMB_Yes:      bTxt = "Yes":       bLoc = 10:  bDef = %False: GOSUB BtnDef
         bID = %MMB_No:       bTxt = "No":        bLoc = 120: bDef = %False: GOSUB BtnDef
         bID = %MMB_Cancel:   bTxt = "Cancel":    bLoc = 230: bDef = %True:  GOSUB BtnDef
      END IF                                                      '
   ELSEIF i = %MB_ABORTRETRYIGNORE THEN
      IF j = %MB_DEFBUTTON1 OR j = 0 THEN                         ' Button 1 (or null request)
         bID = %MMB_Abort:    bTxt = "Abort":     bLoc = 10:  bDef = %True:  GOSUB BtnDef
         bID = %MMB_Retry:    bTxt = "Retry":     bLoc = 110: bDef = %False: GOSUB BtnDef
         bID = %MMB_Ignore:   bTxt = "Ignore":    bLoc = 230: bDef = %False: GOSUB BtnDef
      ELSEIF j = %MB_DEFBUTTON2 THEN                              ' Button 2
         bID = %MMB_Abort:    bTxt = "Abort":     bLoc = 10:  bDef = %False: GOSUB BtnDef
         bID = %MMB_Retry:    bTxt = "Retry":     bLoc = 110: bDef = %True:  GOSUB BtnDef
         bID = %MMB_Ignore:   bTxt = "Ignore":    bLoc = 230: bDef = %False: GOSUB BtnDef
      ELSE                                                        ' Button 3
         bID = %MMB_Abort:    bTxt = "Abort":     bLoc = 10:  bDef = %False: GOSUB BtnDef
         bID = %MMB_Retry:    bTxt = "Retry":     bLoc = 120: bDef = %False: GOSUB BtnDef
         bID = %MMB_Ignore:   bTxt = "Ignore":    bLoc = 230: bDef = %True:  GOSUB BtnDef
      END IF                                                      '
   ELSEIF i = %MB_CANCELTRYCONTINUE THEN
      IF j = %MB_DEFBUTTON1 OR j = 0 THEN                         ' Button 1 (or null request)
         bID = %MMB_Cancel:   bTxt = "Cancel":    bLoc = 10:  bDef = %True:  GOSUB BtnDef
         bID = %MMB_TryAgain: bTxt = "Try Again": bLoc = 120: bDef = %False: GOSUB BtnDef
         bID = %MMB_Continue: bTxt = "Continue":  bLoc = 230: bDef = %False: GOSUB BtnDef
      ELSEIF j = %MB_DEFBUTTON2 THEN                              ' Button 2
         bID = %MMB_Cancel:   bTxt = "Cancel":    bLoc = 10:  bDef = %False: GOSUB BtnDef
         bID = %MMB_TryAgain: bTxt = "Try Again": bLoc = 120: bDef = %True:  GOSUB BtnDef
         bID = %MMB_Continue: bTxt = "Continue":  bLoc = 230: bDef = %False: GOSUB BtnDef
      ELSE                                                        ' Button 3
         bID = %MMB_Cancel:   bTxt = "Cancel":    bLoc = 10:  bDef = %False: GOSUB BtnDef
         bID = %MMB_TryAgain: bTxt = "Try Again": bLoc = 120: bDef = %False: GOSUB BtnDef
         bID = %MMB_Continue: bTxt = "Continue":  bLoc = 230: bDef = %True:  GOSUB BtnDef
      END IF                                                      '
   ELSEIF i = %MB_YESNO THEN
      IF j = %MB_DEFBUTTON1 OR j = 0 THEN                         ' Button 1 (or null request)
         bID = %MMB_Yes:      bTxt = "Yes":       bLoc = 10:  bDef = %True:  GOSUB BtnDef
         bID = %MMB_No:       bTxt = "No":        bLoc = 120: bDef = %False: GOSUB BtnDef
      ELSEIF j = %MB_DEFBUTTON2 THEN                              ' Button 2
         bID = %MMB_Yes:      bTxt = "Yes":       bLoc = 10:  bDef = %False: GOSUB BtnDef
         bID = %MMB_No:       bTxt = "No":        bLoc = 120: bDef = %True:  GOSUB BtnDef
      END IF                                                      '
   ELSEIF i = %MB_RETRYCANCEL THEN
      IF j = %MB_DEFBUTTON1 OR j = 0 THEN                         ' Button 1 (or null request)
         bID = %MMB_Retry:    bTxt = "Retry":     bLoc = 10:  bDef = %True:  GOSUB BtnDef
         bID = %MMB_Cancel:   bTxt = "Cancel":    bLoc = 120: bDef = %False: GOSUB BtnDef
      ELSEIF j = %MB_DEFBUTTON2 THEN                              ' Button 2
         bID = %MMB_Retry:    bTxt = "Retry":     bLoc = 10:  bDef = %False: GOSUB BtnDef
         bID = %MMB_Cancel:   bTxt = "Cancel":    bLoc = 120: bDef = %True:  GOSUB BtnDef
      END IF                                                      '
   ELSEIF i = %MB_OKCANCEL THEN
      IF j = %MB_DEFBUTTON1 OR j = 0 THEN                         ' Button 1 (or null request)
         bID = %MMB_OK:       bTxt = "OK":        bLoc = 10:  bDef = %True:  GOSUB BtnDef
         bID = %MMB_Cancel:   bTxt = "Cancel":    bLoc = 120: bDef = %False: GOSUB BtnDef
      ELSEIF j = %MB_DEFBUTTON2 THEN                              ' Button 2
         bID = %MMB_OK:       bTxt = "OK":        bLoc = 10:  bDef = %False: GOSUB BtnDef
         bID = %MMB_Cancel:   bTxt = "Cancel":    bLoc = 120: bDef = %True:  GOSUB BtnDef
      END IF                                                      '
   ELSE
      bID = %MMB_OK:          bTxt = "OK":        bLoc = 30:  bDef = %True:  GOSUB BtnDef
   END IF
   DIALOG POST hMsg, %WM_CHANGEUISTATE, MAKLNG(%UIS_CLEAR, %UISF_HIDEFOCUS OR %UISF_HIDEACCEL), 0

   '----- Add icon
   h = lStyle AND &H000000F0                                      ' Isolate Icon portion
   IF h = 0 THEN h = %MB_USERICON                                 ' Provide a default
   IF h = %MB_USERICON THEN                                       ' The user version?
      CONTROL ADD IMAGEX, hMsg, %MMB_Icon, "a", 10, 8, 32, 32, %SS_ICON
   ELSE                                                           ' Do the standard Icons
      CONTROL ADD LABEL,  hMsg, %MMB_Icon, "",  10, 8, 0, 0, %SS_ICON
      i = 0: j = %MB_OK                                           '
      SELECT CASE AS LONG h                                       '
         CASE %MB_ICONERROR, %MB_ICONHAND, %MB_ICONSTOP           '
            i = 103: j = %MB_ICONHAND                             '
         CASE %MB_ICONQUESTION                                    '
            i = 102: j = %MB_ICONQUESTION                         '
         CASE %MB_ICONWARNING, %MB_ICONEXCLAMATION                '
            i = 101: j = %MB_ICONEXCLAMATION                      '
         CASE %MB_ICONINFORMATION, %MB_ICONASTERISK               '
            i = 104: j = %MB_ICONASTERISK                         '
      END SELECT                                                  '
      hIcon = LoadImage(BYVAL 0, BYVAL i, %IMAGE_ICON, 32, 32, 0)
      SendDlgItemMessage hMsg, %MMB_Icon, %STM_SETIMAGE, %IMAGE_ICON, hIcon
   END IF                                                         '

   '----- Finally we can display it
   DIALOG SHOW MODAL hmsg, CALL MyMsgBoxproc() TO lResult         ' Go display it
   FOR i = 0 TO 10
      IF hFont(i) <> 0 THEN FONT END hfont(i)                     ' Destroy our fonts
   NEXT i                                                         '
   FUNCTION = lResult                                             ' Pass back result
   EXIT FUNCTION

BtnDef:
   CONTROL ADD BUTTON, hMsg, bID, bTxt, bLoc, h, 90, 20, IIF(bDef, %BS_DEFAULT, 0)
   CONTROL SET FONT hMsg, bID, hFont(0)                           ' Set Button font
   IF bDef THEN CONTROL SET FOCUS hMsg, bID                       ' Set focus if default
   RETURN                                                         ' Done
END FUNCTION

CALLBACK FUNCTION MyMsgBoxProc() AS LONG
LOCAL RC AS LONG
   SELECT CASE CB.MSG
      CASE %WM_COMMAND
         IF CB.CTLMSG = %BN_CLICKED THEN                          ' If button clicked,
            SELECT CASE CB.CTL                                    ' Return with which one
               CASE %MMB_OK:      RC = %IDOK
               CASE %MMB_Yes:     RC = %IDYES
               CASE %MMB_No:      RC = %IDNO
               CASE %MMB_Cancel:  RC = %IDCANCEL
               CASE %MMB_Abort:   RC = %IDABORT
               CASE %MMB_Retry:   RC = %IDRETRY
               CASE %MMB_Ignore:  RC = %IDIGNORE
               CASE %MMB_TryAgain:RC = %IDTRYAGAIN
               CASE %MMB_Continue:RC = %IDCONTINUE
            END SELECT
            DIALOG END CB.HNDL, RC
         END IF
    END SELECT
END FUNCTION

FUNCTION DlgV10() AS STRING
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL dList(), t, t2, bTxt, lTitle AS STRING
LOCAL hAuto, hFont1, hFont2 AS DWORD
LOCAL dNum, lResult, x, y AS LONG
DIM dList(100) AS STRING
   '----- Get the currently available .AUTO files
   sQDir2Array(ENV.AutoPath + "*.AUTO", dList(), dNum)            ' Get current list of .AUTO files
   ARRAY DELETE dList(0)                                          ' Remove empty (0) entry

   '----- Get a font
   FONT NEW "Segoe UI SemiBold", 14 TO hFont1                     ' Build our Fonts
   FONT NEW "Segoe UI SemiBold", 18 TO hFont2                     '

   '----- Build a dialog to get font sizes
   DIALOG NEW PIXELS, 0, "Choose AUTO file to use as model", 0, 0, 700, 450 TO hAuto             ' Create the Dialog
   DIALOG SET COLOR hAuto, %RGB_BLUE, %RGB_BEIGE                  ' Set color
   CONTROL ADD LABEL, hAuto, %AUTO_T1, "AUTO file Colorization Option", 185, 5, 690, 45
   CONTROL SET FONT hAuto, %AUTO_T1, hFont2                       ' Set Font
   CONTROL SET COLOR hAuto, %AUTO_T1, %RGB_BLUE, %RGB_BEIGE       ' Set color
   CONTROL ADD LABEL, hAuto, %AUTO_T2, "SPFLite has converted your Basic Color settings to a new internal structure.", 5, 55, 690, 25
   CONTROL SET FONT hAuto, %AUTO_T2, hFont1                       ' Set Font
   CONTROL SET COLOR hAuto, %AUTO_T2, %RGB_BLUE, %RGB_BEIGE       ' Set color
   CONTROL ADD LABEL, hAuto, %AUTO_T3, "If you also use .AUTO colorize support, you may now choose to select one of your", 5, 85, 690, 25
   CONTROL SET FONT hAuto, %AUTO_T3, hFont1                       ' Set Font
   CONTROL SET COLOR hAuto, %AUTO_T3, %RGB_BLUE, %RGB_BEIGE       ' Set color
   CONTROL ADD LABEL, hAuto, %AUTO_T4, "your .AUTO files from which to build the default custom color schemes.", 5, 115, 690, 25
   CONTROL SET FONT hAuto, %AUTO_T4, hFont1                       ' Set Font
   CONTROL SET COLOR hAuto, %AUTO_T4, %RGB_BLUE, %RGB_BEIGE       ' Set color
   CONTROL ADD LABEL, hAuto, %AUTO_T5, "NOTE:", 275, 165, 690, 25
   CONTROL SET FONT hAuto, %AUTO_T5, hFont2                       ' Set Font
   CONTROL SET COLOR hAuto, %AUTO_T5, %RGB_BLUE, %RGB_BEIGE       ' Set color
   CONTROL ADD LABEL, hAuto, %AUTO_T6, "If you do NOT use AUTO Colorize support,", 275, 205, 690, 25
   CONTROL SET FONT hAuto, %AUTO_T6, hFont1                       ' Set Font
   CONTROL SET COLOR hAuto, %AUTO_T6, %RGB_BLUE, %RGB_BEIGE       ' Set color
   CONTROL ADD LABEL, hAuto, %AUTO_T7, "       simply click CANCEL below.", 275, 235, 690, 25
   CONTROL SET FONT hAuto, %AUTO_T7, hFont1                       ' Set Font
   CONTROL SET COLOR hAuto, %AUTO_T7, %RGB_BLUE, %RGB_BEIGE       ' Set color
   CONTROL ADD LABEL, hAuto, %AUTO_T8, "If you DO,", 275, 265, 690, 25
   CONTROL SET FONT hAuto, %AUTO_T8, hFont1                       ' Set Font
   CONTROL SET COLOR hAuto, %AUTO_T8, %RGB_BLUE, %RGB_BEIGE       ' Set color
   CONTROL ADD LABEL, hAuto, %AUTO_T9, "       select your desired 'model' AUTO file", 275, 295, 690, 25
   CONTROL SET FONT hAuto, %AUTO_T9, hFont1                       ' Set Font
   CONTROL SET COLOR hAuto, %AUTO_T9, %RGB_BLUE, %RGB_BEIGE       ' Set color
   CONTROL ADD LABEL, hAuto, %AUTO_T10, "       and click CONTINUE", 275, 325, 690, 25
   CONTROL SET FONT hAuto, %AUTO_T10, hFont1                       ' Set Font
   CONTROL SET COLOR hAuto, %AUTO_T10, %RGB_BLUE, %RGB_BEIGE       ' Set color

   '----- Now add the ListBox from which to select the AUTO file
   CONTROL ADD LISTBOX, hAuto, %AUTO_List, dList(), 25, 165, 200, 250
   CONTROL SET FONT hAuto, %AUTO_List, hBoldFont                  ' Set Font

   '----- Add the CANCEL and COntinue buttons
   CONTROL ADD BUTTON, hAuto, %AUTO_Continue, "Continue", 20, 400, 100, 40
   CONTROL SET FONT hAuto, %AUTO_Continue, hBoldFont              ' Set Font
   CONTROL ADD BUTTON, hAuto, %AUTO_Cancel, "Cancel", 175, 400, 100, 40
   CONTROL SET FONT hAuto, %AUTO_Cancel, hBoldFont                  ' Set Font


   '----- Finally we can display it
   DESKTOP GET CLIENT TO x, y                                     '
   DIALOG SET LOC hAuto, (x - 700) / 2 , (y - 450) / 2                                      '

   DIALOG SHOW MODAL hAuto, CALL DlgV10Proc() TO lResult          ' Go display it
   FONT END hfont1: FONT END hFont2                               ' Destroy our fonts
   FUNCTION = IIF$(lResult = 0, "", dList(lResult - 1))           ' Pass back result
   EXIT FUNCTION
END FUNCTION
'
CALLBACK FUNCTION DlgV10Proc() AS LONG
LOCAL i AS LONG
   SELECT CASE CB.MSG
      CASE %WM_COMMAND
         IF CB.CTLMSG = %BN_CLICKED THEN                          ' If button clicked,
            SELECT CASE CB.CTL                                    ' Return with which one
               CASE %AUTO_CANCEL                                  ' If CANCEL
                  DIALOG END CB.HNDL, 0
               CASE %AUTO_Continue                                ' If Continue
                  LISTBOX GET SELCOUNT CB.HNDL, %AUTO_List TO i   ' Get how many selected
                  IF i <> 1 THEN                                  ' Error if not 1
                     MSGBOX "You must select a single AUTO file", , "AUTO conversion error
                     EXIT FUNCTION                                ' Issue error and return
                  END IF                                          '
                  LISTBOX GET SELECT CB.HNDL, %AUTO_List TO i     ' Get which one
                  DIALOG END CB.HNDL, i                           ' Return it's index
            END SELECT
         END IF
    END SELECT
END FUNCTION

SUB     DispSplash()
'---------- Display the Splash screen
LOCAL hFont1, hFont2, hFont3 AS LONG, x, y AS LONG

   '----- Put out the Splash screen
   FONT NEW "Segoe Script Bold", 12 TO hFont1                   ' Build normal
   FONT NEW "Segoe UI SemiBold", 13 TO hFont2                     ' Build larger
   FONT NEW "Segoe UI", 10 TO hFont3                              ' Build smaller
   DIALOG NEW PIXELS, 0, "", 0, 0, 350, 350, %WS_POPUP, %WS_EX_TOPMOST TO hSplash
   DIALOG SET COLOR   hSplash, %RGB_MEDIUMBLUE, %RGB_GAINSBORO
   CONTROL ADD IMAGE, hSplash, %GENERAL_LOGO, "SPFLogo", 35, 0, 300, 50
   CONTROL ADD LABEL, hSplash, %GENERAL_T1, "A Windows Editor for SPF lovers", 35, 55, 290, 25
   CONTROL SET FONT   hSplash, %GENERAL_T1, hFont1
   CONTROL SET COLOR  hSplash, %GENERAL_T1, %BLACK, %RGB_GAINSBORO
   CONTROL ADD LABEL, hSplash, %GENERAL_T2, "SPFLite Copyright © 2005-2018", 45, 90, 290, 25
   CONTROL SET FONT   hSplash, %GENERAL_T2, hFont2
   CONTROL SET COLOR  hSplash, %GENERAL_T2, %RGB_MEDIUMBLUE, %RGB_GAINSBORO
   CONTROL ADD LABEL, hSplash, %GENERAL_T3, "G. D. Deluca and R. L. Hodge", 45, 115, 290, 25
   CONTROL SET FONT   hSplash, %GENERAL_T3, hFont2
   CONTROL SET COLOR  hSplash, %GENERAL_T3, %RGB_MEDIUMBLUE, %RGB_GAINSBORO
   CONTROL ADD LABEL, hSplash, %GENERAL_T4, "SPFLite is free software and comes with ABSOLUTELY", 1, 135, 350, 15, %SS_CENTER
   CONTROL SET FONT   hSplash, %GENERAL_T4, hFont3
   CONTROL SET COLOR  hSplash, %GENERAL_T4, %RGB_MEDIUMBLUE, %RGB_GAINSBORO
   CONTROL ADD LABEL, hSplash, %GENERAL_T5, "NO WARRANTY. You are welcome to redistribute it under", 1,  150, 350, 15, %SS_CENTER
   CONTROL SET FONT   hSplash, %GENERAL_T5, hFont3
   CONTROL SET COLOR  hSplash, %GENERAL_T5, %RGB_MEDIUMBLUE, %RGB_GAINSBORO
   CONTROL ADD LABEL, hSplash, %GENERAL_T6, "certain conditions. See 'License.txt' in the root install", 1,  165, 350, 15, %SS_CENTER
   CONTROL SET FONT   hSplash, %GENERAL_T6, hFont3
   CONTROL SET COLOR  hSplash, %GENERAL_T6, %RGB_MEDIUMBLUE, %RGB_GAINSBORO
   CONTROL ADD LABEL, hSplash, %GENERAL_T7, "folder for distribution details.", 1, 180, 350, 15, %SS_CENTER
   CONTROL SET FONT   hSplash, %GENERAL_T7, hFont3
   CONTROL SET COLOR  hSplash, %GENERAL_T7, %RGB_MEDIUMBLUE, %RGB_GAINSBORO
   CONTROL ADD LABEL, hSplash, %GENERAL_T8, "SPFLite uses other support tools:", 1,  200, 350, 15, %SS_CENTER
   CONTROL SET FONT   hSplash, %GENERAL_T8, hFont3
   CONTROL SET COLOR  hSplash, %GENERAL_T8, %RGB_MEDIUMBLUE, %RGB_GAINSBORO
   CONTROL ADD LABEL, hSplash, %GENERAL_T9, "thinBasic (C) 2015 Eros Olmi", 1,  215, 350, 15, %SS_CENTER
   CONTROL SET FONT   hSplash, %GENERAL_T9, hFont3
   CONTROL SET COLOR  hSplash, %GENERAL_T9, %RGB_MEDIUMBLUE, %RGB_GAINSBORO
   CONTROL ADD LABEL, hSplash, %GENERAL_T10, "PCRE (C) 2018 Philip Hazel", 1,  230, 350, 15, %SS_CENTER
   CONTROL SET FONT   hSplash, %GENERAL_T10, hFont3
   CONTROL SET COLOR  hSplash, %GENERAL_T10, %RGB_MEDIUMBLUE, %RGB_GAINSBORO

   CONTROL ADD IMAGE, hSplash, %GENERAL_GPLV3, "GPLV3L", 107, 260, 136, 68
   DESKTOP GET CLIENT TO x, y                                    '
   DIALOG SET LOC hSplash, (x - 350) / 2 , (y - 350) / 2         '
   DIALOG SHOW MODELESS hSplash CALL DlgSplashCallback           ' Display it

END SUB

CALLBACK FUNCTION DlgSplashCallback() AS LONG
   SELECT CASE CB.MSG
      CASE %WM_INITDIALOG
         SetTimer(hSplash, %ID_Timer, 4000, BYVAL %NULL)
      CASE %WM_TIMER
         DIALOG END hSplash
   END SELECT
END FUNCTION
