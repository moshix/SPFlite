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

CLASS cENVariables

'----- Following are the Global working variables, most built by looking around
   INSTANCE InitDone         AS LONG                              ' Set true when all initialization is complete
   INSTANCE AutoPath         AS STRING                            ' Path for \AUTO files
   INSTANCE ClipCmd          AS STRING                            ' CLIP operand
   INSTANCE ClipPath         AS STRING                            ' Path for \CLIP storage
   INSTANCE CmdLine          AS STRING                            ' Command line (Raw)
   INSTANCE EXEFullPath      AS STRING                            ' Full path to the EXE file
   INSTANCE EXEPath          AS STRING                            ' Path to the EXE folder
   INSTANCE FileListPath     AS STRING                            ' Path for \FILELIST files
   INSTANCE FirstTime        AS LONG                              ' This is a First Time run
   INSTANCE INIPath          AS STRING                            ' Path to the INI file
   INSTANCE INIFileName      AS STRING                            ' Full name of the INI file
   INSTANCE InitString       AS STRING                            ' Cmd line filename or MRF file list
   INSTANCE InitProfile      AS STRING                            ' Cmd line .profile
   INSTANCE InvChar          AS STRING                            ' Display char. for unprintable characters
   INSTANCE JobsPath         AS STRING                            ' Path for \JOBS
   INSTANCE MacrosPath       AS STRING                            ' Path for \MACROS files
   INSTANCE PMode            AS LONG                              ' Parameter options
   INSTANCE ProfPath         AS STRING                            ' Path for \Profiles
   INSTANCE PgmDate          AS STRING                            ' EXE creation date
   INSTANCE PgmVers          AS STRING                            ' EXE version
   INSTANCE Roaming          AS LONG                              ' Execution roaming mode
   INSTANCE RunPath          AS STRING                            ' Path for \RUN
   INSTANCE SETETab          AS LONG                              ' Tab # that issued SET Edit
   INSTANCE StatePath        AS STRING                            ' Path for \STATE files
   INSTANCE VSavePath        AS STRING                            ' Path for \VSAVE files
   INSTANCE WineMode         AS LONG                              ' WINE mode
   INSTANCE NoLoopMode       AS LONG                              ' NOLOOP mode
   INSTANCE WorkingDir       AS STRING                            ' Our working directory

'----- Following are the variables set from the INI file

   INSTANCE ABeepFlag        AS LONG                              ' Beep on errFlag=%eFail? True = yes
   INSTANCE Allow2DMouse     AS LONG                              ' Allow 2D mouse without mode key
   INSTANCE AutoScroll       AS LONG                              ' 0 = Off # = On and # of Lines/Cols to Scroll
   INSTANCE Banding          AS LONG                              ' Banding ON/OFF
   INSTANCE CharSet          AS STRING                            ' Character set
   INSTANCE CharSetR         AS STRING                            ' Character set for RegEx
   INSTANCE cMarkLine        AS LONG                              ' Color of MARK lines
   INSTANCE CmdChr           AS STRING                            ' Command delimiter
   INSTANCE CMDFlags         AS STRING                            ' CMD.EXE flags for CMD command
   INSTANCE CmdHiLite        AS LONG                              ' Cmd line KW hilite
   INSTANCE ConfDelFlag      AS LONG                              ' Confirm Deletes in File Manager?
   INSTANCE CursInsert       AS LONG                              ' Percent height of an Insert cursor
   INSTANCE CursNormal       AS LONG                              ' Percent height of a Normal cursor
   INSTANCE DefaultShr       AS STRING                            ' Profile names using the DEFAULT Profile
   INSTANCE DefDataShift     AS LONG                              ' Default shift for >> )) ]] line commands
   INSTANCE DirProfFlag      AS LONG                              ' DIR as Profile for no extensions files
   INSTANCE ENGChars         AS LONG                              ' English only A-Z and a-z characters
   INSTANCE FindWord         AS LONG                              ' Find default context is WORD
   INSTANCE FMCloseFlag      AS LONG                              ' Close FM with last tab?
   INSTANCE FMDateHiLite     AS LONG                              ' HiLite Recent/Active dates in the FM List
   INSTANCE FMDefMaster      AS STRING                            ' Master Default FM sort
   INSTANCE FMDefSort        AS STRING                            ' Default FM sort
   INSTANCE FMDirSort        AS STRING                            ' Default FM Dir sort
   INSTANCE FMFFList         AS STRING                            ' FM FF command exemption list
   INSTANCE FMFileList       AS STRING                            ' Initial FM FLIST name
   INSTANCE FMHelpFlag       AS LONG                              ' Display FM Help at bottom of screen?
   INSTANCE FMLastSectNo     AS LONG                              ' FM Index of last fetched section
   INSTANCE FMLayout         AS STRING                            ' FM Line layout
   INSTANCE FMLCmdWidth      AS LONG                              ' Width of FM Line Command  area
   INSTANCE FMMask           AS STRING                            ' Initial FM Mask
   INSTANCE FMNoteWidth      AS LONG                              ' Width of FM Note column  area
   INSTANCE FMPath           AS STRING                            ' Initial FM Path
   INSTANCE FMRecentSort     AS STRING                            ' Default FM Recent sort
   INSTANCE FMScrlAmt        AS STRING                            ' FM Scroll amount
   INSTANCE FontName         AS STRING                            ' Font name for screen
   INSTANCE FontPitch        AS LONG                              ' Font pitch for screen
   INSTANCE FontStyle        AS LONG                              ' Font Style for screen
   INSTANCE GlueWith         AS STRING                            ' GLUE join cgaracter string
   INSTANCE HRuler           AS LONG                              ' Horizontal ruler cursor
   INSTANCE InsMode          AS LONG                              ' Default Insert mode
   INSTANCE InsReset         AS LONG                              ' Insert mode Reset at each Attn key?
   INSTANCE JobNumber        AS LONG                              ' Submit job number
   INSTANCE KBHorScroll      AS LONG                              ' Scroll (Hor) screen when cursor hits edge?
   INSTANCE KBVerScroll      AS LONG                              ' Scroll (Ver) screen when cursor hits edge?
   INSTANCE LastDirOpen      AS LONG                              ' What directory to default OPEN to
   INSTANCE LastScreenX      AS LONG                              ' Last screen location X
   INSTANCE LastScreenY      AS LONG                              ' Last screen location Y
   INSTANCE LinNoSize        AS LONG                              ' Width of Line Number field
   INSTANCE LinRepeatMax     AS LONG                              ' Max Line Cmd rerpeat value
   INSTANCE MAXIMIZED        AS LONG                              ' Screen is in MAX mode
   INSTANCE MinRetrieve      AS LONG                              ' Min length retrievable command
   INSTANCE NotifyLevel      AS LONG                              ' Permanent Notify option
   INSTANCE NotifyLevelT     AS LONG                              ' Temporary (working) Notify option
   INSTANCE PFKShow          AS LONG                              ' Show PFKeys at bottom
   INSTANCE PrtBandColor     AS LONG                              ' Banding color
   INSTANCE PrtBanding       AS LONG                              ' Banding stripes
   INSTANCE PrtBandLines     AS LONG                              ' Banding lines
   INSTANCE PrtBMargin       AS SINGLE                            ' Printer bottom margin (in)
   INSTANCE PrtDuplex        AS LONG                              ' Printer duplex setting
   INSTANCE PrtFontName      AS STRING                            ' Printer font
   INSTANCE PrtFontPitch     AS STRING                            ' Printer pitch
   INSTANCE PrtFontStyle     AS STRING                            ' Printer style
   INSTANCE PrtFooter        AS LONG                              ' Printer footers 1 = yes
   INSTANCE PrtFooterCenter  AS STRING                            ' Printer center footers
   INSTANCE PrtFooterLeft    AS STRING                            ' Printer left footers
   INSTANCE PrtFooterRight   AS STRING                            ' Printer left footers
   INSTANCE PrtHeader        AS LONG                              ' Printer headers 1 = yes
   INSTANCE PrtHeaderCenter  AS STRING                            ' Printer center headers
   INSTANCE PrtHeaderLeft    AS STRING                            ' Printer left headers
   INSTANCE PrtHeaderRight   AS STRING                            ' Printer left headers
   INSTANCE PrtLMargin       AS SINGLE                            ' Printer left margin (in)
   INSTANCE PrtMetric        AS LONG                              ' Printer margins Eng/Metric
   INSTANCE PrtName          AS STRING                            ' Printer name
   INSTANCE PrtOrient        AS LONG                              ' Printer orientation, 1=Portrait 2=Landscape
   INSTANCE PrtPaper         AS LONG                              ' Paper index
   INSTANCE PrtRMargin       AS SINGLE                            ' Printer right margin (in)
   INSTANCE PrtTMargin       AS SINGLE                            ' Printer top margin (in)
   INSTANCE PrtPColor        AS LONG                              ' Print in color (%True or %False)
   INSTANCE QuickRenum       AS LONG                              ' Quick renum over this many lines
   INSTANCE RecentCtr        AS LONG                              ' How many recent files to keep
   INSTANCE ReOpenLast       AS LONG                              ' Re-Open last file(s)
   INSTANCE ResetU           AS LONG                              ' RESET clears U by default, like X
   INSTANCE RUNFlags         AS STRING                            ' CMD.EXE flags for RUN command
   INSTANCE SBLayout         AS STRING                            ' Status Bar layout
   INSTANCE Scheme()         AS LONG                              ' Scheme color value
   INSTANCE SchemeSave()     AS LONG                              ' Saved copy of Scheme
   INSTANCE ScrHeight        AS LONG                              ' Screen height
   INSTANCE ScrWidth         AS LONG                              ' Screen width
   INSTANCE Splash           AS LONG                              ' Display Splash screen (True/False)
   INSTANCE SubmitCmd        AS STRING                            ' Model command for SUBMIT
   INSTANCE SubmitDir        AS STRING                            ' Working DIR for SUBMIT
   INSTANCE SubmitInclKey    AS STRING                            ' Include key string for SUBMIT
   INSTANCE SubmitInclCol    AS LONG                              ' Include key string column for SUBMIT
   INSTANCE UniqueFlag       AS LONG                              ' Keep only 1 instance running
   INSTANCE UseRecycle       AS LONG                              ' Delete to Recycle bin?
   INSTANCE VBeepFlag        AS LONG                              ' Visual Beep on errFlag=%eFail? True = yes
   INSTANCE VertInsCurs      AS LONG                              ' Vertical Insert Cursor
   INSTANCE ViewWarn         AS LONG                              ' Warn on modified View?
   INSTANCE VRuler           AS LONG                              ' Vertical ruler cursor
   INSTANCE WelcomeOpt       AS LONG                              ' DispWelcome return option

   CLASS METHOD CREATE()                                          ' Constructor - Initialize Class stuff
   LOCAL i, rc, lpdwHandle AS LONG, t, BlockBuffer, sLangID, cword AS STRING
   LOCAL pLang AS LONG PTR
   LOCAL lclpValue AS ASCIIZ PTR
   LOCAL approvedlocation AS LONG
   LOCAL approvedpath  AS ASCIIZ * %MAX_PATH
   LOCAL subfolder     AS ASCIIZ * %MAX_PATH
   LOCAL tASCIIZ       AS ASCIIZ * %MAX_PATH
   LOCAL DriveHandle, zero, rlen AS DWORD
   LOCAL hp AS STORAGE_HOTPLUG_INFO
   DIM Scheme(0 TO 47, 1 TO 3) AS INSTANCE LONG
   DIM SchemeSave(0 TO 47, 1 TO 3) AS INSTANCE LONG

      '----- Fetch info about who we are, where we are etc.
      PMode = 0                                                   ' Start PMode as all off
      EXEFullPath = EXE.FULL$                                     ' Save where the EXE file is
      tASCIIZ = EXEFullPath
      i = GetFileVersionInfoSize(tASCIIZ, lpdwHandle)             ' Fetch size of area needed
      BlockBuffer = SPACE$(i)                                     ' Alloc a block
      rc = GetFileVersionInfo(BYCOPY EXEFullPath, lpdwHandle, i, BYVAL STRPTR(BlockBuffer))          ' Copy the EXE info block
      rc = VerQueryValue(BYVAL STRPTR(BlockBuffer), "\VarFileInfo\Translation", pLang, lpdwHandle)   ' Fetch language
      sLangID = IIF$(rc, HEX$(LO(WORD, @pLang), 4) + HEX$(HI(WORD, @pLang), 4), "040904E4")          ' Get lang. or use American English/ANSI

      rc = VerQueryValue(BYVAL STRPTR(BlockBuffer), "\StringFileInfo\" & sLangID & "\VersionDate", lclpValue, lpdwHandle) ' Fetch the version date
      PgmDate = @lclpValue                                        ' Save date
      rc = VerQueryValue(BYVAL STRPTR(BlockBuffer), "\StringFileInfo\" & sLangID & "\ProductVersion", lclpValue, lpdwHandle) ' Fetch the version
      PgmVers = @lclpValue                                        ' Save version

      '----- Save EXE and INI paths
      EXEPath = EXE.PATH$                                         ' Save the EXE path etc. we need gfRoaming setup real early
      approvedlocation = %csidl_personal                          ' Ask system for the correct INI path
      subfolder = "SPFLite"                                       '
      rc = shgetfolderpath (BYVAL %null, approvedlocation OR %CSIDL_FLAG_CREATE, -0&, %shgfp_type_current, approvedpath)
      INIPath = TRIM$(approvedpath) + "\SPFLite\"                 '

      '----- See if we're running from a removable drive, if so say we're in 'Roaming' mode
      tASCIIZ = "\\.\" + LEFT$(EXEPath, 2)                        ' Lets see where we're running from
      DriveHandle = CreateFile(tASCIIZ, 0, %FILE_SHARE_READ, BYVAL 0&, %OPEN_EXISTING, 0, BYVAL 0&)           ' Get a Handle for it
      DeviceIoControl(DriveHandle, %IOCTL_STORAGE_GET_HOTPLUG_INFO, Zero, 0&, BYREF hp, SIZEOF(hp), rlen, "") ' Get hotplug data
      CloseHandle(DriveHandle)                                    ' Done with handle now
      IF hp.devicehotplug THEN                                    ' Romoveable?
         Roaming = %True                                          ' Running from a removeable drive? Say we 're in Roaming mode
         INIPath = EXEPath + "CONFIG\"                            ' Swap INI path
      END IF
      IF ISFALSE ISFOLDER(INIPath) THEN MKDIR INIPath             ' Make sure INIPath exists

      '----- New see if the Cmd line overrides anything
      t = COMMAND$: CmdLine = t                                   ' Get Command$ and save it
      WorkingDir = CURDIR$                                        ' Get initial startup Working Directory
      PMode = %MEdit                                              ' Default to EDIT mode
      IF ISNOTNULL(t) THEN                                        ' Anything in COMMAND$

         '----- Process the operands
         DO WHILE ISNOTNULL(TRIM$(t))                             ' Loop through operands
            CWord = sGetWord(t, %Strip, %QuoteSig)                ' Get a word from the command
            IF LEFT$(CWord,1) = "-" THEN CWord = UUCASE(CWord)    ' If a -xxx operand, ucase it
            '----- CLIP
            IF cWord = LEFT$("-CLIP", LEN(CWord)) THEN            ' A CLIP request?
               PMode = PMode OR %MClip OR %MClipLoad OR %MClipSave' Remember CLIP mode active
               t = ""                                             ' Null anything else with CLIP
            '----- WINE
            ELSEIF cWord = LEFT$("-WINE", LEN(CWord)) THEN        ' A WINE request?
               WineMode = %True                                   ' Yes set flag
            '----- NOLOOP
            ELSEIF cWord = LEFT$("-NOLOOP", LEN(CWord)) THEN      ' A NOLOOP request?
               NoLoopMode = %True                                 ' Yes set flag
            '----- INIT
            ELSEIF cWord = LEFT$("-INIT", LEN(CWord)) THEN        ' An INIT request?
               INIFileName = sGetWord(t, %Strip, %QuoteSig)       ' Extract the INI filename
               sUnQuote(INIFileName)                              ' Remove quotes if needed
               IF UCASE$(RIGHT$(INIFileName, 4)) <> ".INI" THEN INIFileName += ".INI"
               IF INSTR(INIFileName, "\") = 0 THEN                ' Got a PATH?
                  INIFileName = INIPath + INIFileName             ' Insert the path
               END IF                                             '
               IF ISFALSE ISFILE(INIFileName) THEN                ' Better exist
                  sDoMsgBox "The specified INI file:|K" + $CRLF + INIFileName + $CRLF + _
                         "|Bdoes not exist, using the normal default INI", %MB_OK OR %MB_USERICON, "INI Load"
                  INIFileName = TRIM$(approvedpath) + "\SPFLite\SPFLite.INI"  '
               END IF
               INIPath = PATHNAME$(PATH, INIFileName)             ' And it's path
            '----- BROWSE
            ELSEIF cWord = LEFT$("-BROWSE", LEN(CWord)) THEN      ' A BROWSE request?
               PMode = PMode AND (&HFFFFFFFF - %MEdit)            ' Yes, Remove Edit Default
               PMode = PMode OR %MBrowse                          ' And set flag
            '----- VIEW
            ELSEIF cWord = LEFT$("-VIEW", LEN(CWord)) THEN        ' A VIEW request?
               PMode = PMode AND (&HFFFFFFFF - %MEdit)            ' Yes, Remove Edit Default
               PMode = PMode OR %MView                            ' And set flag
            '----- OPEN
            ELSEIF cWord = LEFT$("-OPEN", LEN(CWord)) THEN        ' An OPEN request?
               PMode = PMode OR %MOpen                            ' Yes set flag
            '----- OPENB
            ELSEIF cWord = LEFT$("-OPENB", LEN(CWord)) THEN       ' An OPEN Browse?
               PMode = PMode AND (&HFFFFFFFF - %MEdit)            ' Yes, Remove Edit Default
               PMode = PMode OR (%MBrowse + %MOpen)               ' And set flags
            '----- OPENV
            ELSEIF cWord = LEFT$("-OPENV", LEN(CWord)) THEN       ' An OPEN View?
               PMode = PMode AND (&HFFFFFFFF - %MEdit)            ' Yes, Remove Edit Default
               PMode = PMode OR (%MView + %MOpen)                 ' And set flags
            '----- KEYMAP
            ELSEIF cWord = LEFT$("-KEYMAP", LEN(CWord)) THEN      ' A KEYMAP request?
               PMode = PMode OR %MKeyMap                          ' Yes set flag
            '----- Maybe a .Profile name
            ELSEIF InitString <> "" AND LEFT$(cWord, 1) = "." THEN
               InitProfile = cWord                                ' Save as Profilename

            '----- Must be a filename or FM Path/Mask request then
            ELSE                                                  ' Anything else must be a filename
               InitString = CWord                                 ' Save it
               sUnQuote(InitString)                               ' Remove quotes
               IF LEFT$(InitString, 1) = "\" THEN InitString = MID$(InitString, 2) ' Strip off any leading \
               IF ISNOTNULL(InitString) THEN                      ' Adjust if something there
                  IF RIGHT$(InitString, 1) = ":" THEN InitString += "\" ' If simple trailing :, add a  \
                  IF RIGHT$(InitString, 1) = "\" THEN InitString += "*" ' If simple trailing \, add an *

                  '----- See if a simple FLIST name
                  IF INSTR(InitString, ANY ":\.*?") = 0 AND _     ' Just a simple non-filename format?
                     ISTRUE ISFILE(INIPath + "\FILELIST\" + InitString + ".FLIST") THEN ' And it exixts as a FLIST
                     FMFileList = InitString: InitString = ""     ' Make it the FM FileList name

                  '----- Any Wild card chars
                  ELSEIF INSTR(InitString, ANY "*?") = 0 THEN     ' No wild card characters?

                     '----- We have a simple filename
                     IF LEFT$(InitString, 2) = ".\" THEN InitString = WorkingDir + MID$(InitString, 2) ' Convert . to WorkingDir
                     IF INSTR(InitString, "\") = 0 THEN _         ' If not a full path specified
                        InitString = WorkingDir + "\" + InitString' Add the current directory
                     IF ISFOLDER(InitString) THEN
                        FMPath = InitString + "\"                 ' Use for the path
                        FMMask = "*"                              ' Use * for Mask
                        FMFileList = "$Null$"                     ' Tell InitFManager that FMRecent is really null
                        InitString = ""                           ' Kill the InitString
                     END IF

                  '----- Some kind of wild card request
                  ELSE
                     IF INSTR(InitString, ANY "\") > 0 THEN       ' Got a directory level?
                        FMPath = LEFT$(InitString, INSTR(-1, InitString, "\")) ' Extract path
                        FMMask = MID$(InitString, INSTR(-1, InitString, "\") + 1) ' Peel off Mask
                        IF FMMask = "" THEN FMMask = "*"          ' If no mask, use *
                        FMFileList = "$Null$"                     ' Tell InitFManager that FileListNm is really null
                        InitString = ""                           ' Kill the InitString
                     ELSE                                         '
                        FMMask = InitString                       ' Make what's left the FM File Mask
                        FMPath = CURDIR$ + "\"                    ' Use the CurDir for the path
                        FMFileList = "$Null$"                     ' Tell InitFManager that FMRecent is really null
                        InitString = ""                           ' Kill the InitString
                     END IF                                       '
                     IF INSTR(FMPath, ANY "*?") > 0 THEN          ' Invalid wild cards in FM File Path?
                        FMPath = "": FMMask = ""                  ' Yes, kill what we did, the path will fail
                        InitString = ""                           ' Kill the InitString
                     ELSE                                         '
                        InitString = ""                           ' Kill InitString
                     END IF                                       '
                  END IF                                          '
               ELSE                                               '
                  InitString = "NEW"                              '
               END IF                                             '
            END IF                                                '
         LOOP                                                     ' end of operand loop
      END IF                                                      '

      '----- Figure out what INI file to use
      IF ISFALSE ISFOLDER(INIPath) THEN MKDIR INIPath             ' Make sure INIPath exists
      IF ISNULL(INIFileName) THEN                                 ' Did user override the INI file name?
         INIFileName = INIPath + "SPFLite.INI"                    ' Create the full name of the INI file
      END IF                                                      '

      IF VAL(sINIGetString("General", "Splash", "0")) THEN DispSplash ' Display Splash as early as possible

      '----- ENVINIPath is set, create some others based on it
      AutoPath     = INIPath + "AUTO\"                            ' Create path for AUTO storage
      ClipPath     = INIPath + "CLIP\"                            ' Create path for CLIP storage
      FileListPath = INIPath + "FILELIST\"                        ' Create path for FILELIST storage
      JobsPath     = INIPath + "JOBS\"                            ' Create path for JOBS usage
      MacrosPath   = INIPath + "MACROS\"                          ' Create path for MACROS storage
      ProfPath     = INIPath + "PROFILES\"                        ' Create path for File Profiles
      RunPath      = INIPath + "RUN\"                             ' Create path for RUN usage
      StatePath    = INIPath + "STATE\"                           ' Create path for STATE storage
      VSavePath    = INIPath + "VSAVE\"                           ' Create path for VSave storega
      IF sINIGetString("System", "Build", "MISSING") = "MISSING" THEN ' Never been executed before?
         i = sRegSet("Software\SPFLite", "CurrVersion", PgmVers)  '
         i = sRegSet("Software\SPFLite", "CurrBuild", PgmDate)    '

         '----- Do some First Time stuff
         FirstTime = %True                                        ' Remember its the first time
         IF ISFALSE ISFOLDER(AutoPath) THEN MKDIR AutoPath        ' Ensure \AUTO exists
         IF ISFALSE ISFOLDER(ClipPath) THEN MKDIR ClipPath        ' Ensure \CLIP exists
         IF ISFALSE ISFOLDER(FileListPath) THEN MKDIR FileListPath' Ensure \FILELIST exists
         IF ISFALSE ISFOLDER(JobsPath) THEN MKDIR JobsPath        ' Ensure \JOBS exists
         IF ISFALSE ISFOLDER(MacrosPath) THEN MKDIR MacrosPath    ' Ensure \MACROS exists
         IF ISFALSE ISFOLDER(ProfPath) THEN MKDIR ProfPath        ' Ensure \PROFILE exists
         IF ISFALSE ISFOLDER(RunPath) THEN MKDIR RunPath          ' Ensure \RUN exists
         IF ISFALSE ISFOLDER(StatePath) THEN MKDIR StatePath      ' Ensure \STATE exists
         IF ISFALSE ISFOLDER(VSavePath) THEN MKDIR VSavePath      ' Ensure \VSAVE exists
         '----- Setup some default INI files if not present
         IF ISFALSE ISFILE(INIPath + "SPFLite.INI") THEN _        ' Create an INI if not present
            FILECOPY EXEPath + "Install SPFLite.INI", INIPath + "SPFLite.INI"
         IF ISFALSE ISFILE(INIPath + "SPFLite.KBD") THEN _        ' Create a KBD if not present
            FILECOPY EXEPath + "Install SPFLite.KBD", INIPath + "SPFLite.KBD"
         IF ISFALSE ISFILE(ProfPath + "DEFAULT.INI") THEN _       ' Create a DEFAULT Profile if not present
            FILECOPY EXEPath + "Install DEFAULT.INI", ProfPath + "DEFAULT.INI"
         Env.DispWelcome                                          ' Go display and get initial KB options
      END IF                                                      '
      i = sRegSet("Software\SPFLite", "CurrVersion", PgmVers)     ' Keep the version updated

      '----- Do special Version 10 startup
      t = sINIGetString("System", "Build", "MISSING")             ' Get the previous running Build level
      IF t = "MISSING" THEN                                       ' Totally new user, skip
         ' Do nothing                                             '
      ELSE                                                        ' Else, check the version
         t = LEFT$(t, INSTR(t, ".") - 1)                          ' Extract top level version
         IF VAL(t) = 0 THEN                                       ' Brand new install?
            ' Do nothing                                          '
         ELSEIF VAL(t) >= 10 THEN                                 ' Already at 10 or higher?
            ' Do nothing                                          '
         ELSE                                                     ' Save INI folder and force user to read something
            t = LEFT$(INIPath, LEN(INIPath) - 1)
            SHELL "XCOPY " + $DQ + t + $DQ + " " + $DQ + t + "_Pre_Version10" + $DQ + " /i /s /e /h /y /q", 0
            tASCIIZ = ENV.EXEPath + "SPFLiteVersion10.chm"
            rc = HtmlHelp(0, tASCIIZ, %HH_DISPLAY_TOC, 0&)        ' OPEN the CHM file up
            i = 1                                                 ' Setup the loop
            DO WHILE i <> 0                                       ' Loop till CHM closes
               SLEEP 3000                                         ' Give time between checks
               i = sFindWindow("SPFLite Version 10 Colorization") ' Go see if still open
            LOOP                                                  ' Loop back if still there
         END IF                                                   '
      END IF                                                      '
      ENV.LoadAllINI                                              ' Load all our INI values
      sIniSetString("System", "Build", PgmVers)                   ' Save latest version
      ENV.SetGlobalClrs                                           ' Copy all colors to Global

   END METHOD                                                     '

   INTERFACE iENVariables: INHERIT IUNKNOWN                       ' Define the interface

      '----- Following are the variables built by looking around
      GSProp(InitDone, LONG)                                      ' All Initialization done
      GSProp(AutoPath, STRING)                                    ' \AUTO path
      GSProp(ClipCmd, STRING)                                     ' CLIP operand
      GSProp(ClipPath, STRING)                                    ' Path for \CLIP storage
      GProp(CmdLine, STRING)                                      ' Command line
      GProp(EXEFullPath, STRING)                                  ' Full path to the EXE file
      GProp(EXEPath, STRING)                                      ' Path to the EXE folder
      GSProp(FileListPath, STRING)                                ' \FILELIST path
      GProp(FirstTime, LONG)                                      ' This is a First Time run
      GProp(INIPath, STRING)                                      ' Path to the INI file
      GProp(INIFileName, STRING)                                  ' Full name of the INI file
      GSProp(InitString, STRING)                                  ' Cmd line filename or Prev open MRF list
      GProp(InitProfile, STRING)                                  ' Cmd line .profile
      GProp(JobsPath, STRING)                                     ' \JOBS path
      GSProp(MacrosPath, STRING)                                  ' \MACROS path
      GSProp(PMode, LONG)                                         ' Parameter mode
      GProp(PgmDate, STRING)                                      ' EXE creation date
      GProp(ProfPath, STRING)                                     ' Path for \Profiles
      GProp(PgmVers, STRING)                                      ' EXE version
      GProp(Roaming, LONG)                                        ' Execution roaming mode
      GSProp(RunPath, STRING)                                     ' \RUN path
      GSProp(SETETab, LONG)                                       ' Tab # that issued SET Edit
      GSProp(StatePath, STRING)                                   ' \STATE path
      GSProp(VSavePath, STRING)                                   ' \VSAVE path
      GProp(WineMode, LONG)                                       ' WINE mode
      GProp(NoLoopMode, LONG)                                     ' NOLOOP mode
      GProp(WorkingDir, STRING)                                   ' Our working directory
      GSProp(WelcomeOpt, LONG)                                    ' DispWelcome return option

      '----- Following are the variables set from the INI file

      GProp(ABeepFlag, LONG)                                      ' Beep on errFlag=%eFail? True = yes
      PROPERTY SET ABeepFlag(BYVAL v AS LONG)
         ABeepFlag = v:
         sIniSetString("General", "BeepFlag", FORMAT$(ABeepFlag))
         me.SetINITimeStamp
      END PROPERTY

      GProp(Allow2DMouse, LONG)                                   ' Allow 2D mouse without mode key
      PROPERTY SET Allow2DMouse(BYVAL v AS LONG)
         Allow2DMouse = v:
         sIniSetString("General", "Allow2DMouse", FORMAT$(Allow2DMouse))
         me.SetINITimeStamp
      END PROPERTY

      GProp(AutoScroll, LONG)                                     ' 0 = Off # = On and # of Lines/Cols to Scroll
      PROPERTY SET AutoScroll(BYVAL v AS LONG)
         AutoScroll = v:
         sIniSetString("Mouse", "AutoScrl", FORMAT$(AutoScroll))
         me.SetINITimeStamp
      END PROPERTY

      GProp(Banding, LONG)                                        ' Banding ON/OFF
      PROPERTY SET Banding(BYVAL v AS LONG)
         Banding = v:
         sIniSetString("Screen", "Banding", FORMAT$(Banding))
         me.SetINITimeStamp
      END PROPERTY

      GProp(CharSet, STRING)                                      ' Character set
      PROPERTY SET CharSet(BYVAL v AS STRING)
         CharSet = v:
         CharSet = REMOVE$(CharSet, " ")                          ' Charset should not have blanks

         IF ISNULL(CharSet) THEN                                  ' Again if null
            IF ENGChars THEN                                      ' Default to use ENGLISH-ONLY characters
               CharSet = $DefCharSet                              ' Set it without Nat chars
            ELSE                                                  '
               CharSet = $DefCharSet + $UpperNat + $LowerNat      ' Set it that way
            END IF                                                '
         END IF                                                   '

         IF ENGChars THEN                                         ' Default to use ENGLISH-ONLY characters
            gUpper = $Upper: gLower = $Lower                      ' Set working $Upper and $Lower
         ELSE                                                     '
            gUpper = $Upper + $UpperNat: gLower = $Lower + $LowerNat ' Set working $Upper and $Lower
         END IF                                                   '

         CharSetR = CharSet                                       ' Massage into RegEx string
         REPLACE "\" WITH "\x5C" IN CharSetR                      '
         REPLACE "[" WITH "\x5B" IN CharSetR                      '
         REPLACE "]" WITH "\x5D" IN CharSetR                      '
         sIniSetString("General", "CharSet", CharSet)             '
         me.SetINITimeStamp
      END PROPERTY

      GSProp(CharSetR, STRING)                                    ' Character set for RegEx

      GProp(cMarkLine, LONG)                                      ' Color of MARK lines
      PROPERTY SET cMarkLine(BYVAL v AS LONG)
         cMarkLine = v:
         sIniSetString("Screen", "cMarkLine", FORMAT$(cMarkLine))
         me.SetINITimeStamp
      END PROPERTY

      GProp(CmdChr, STRING)                                       ' Command delimiter
      PROPERTY SET CmdChr(BYVAL v AS STRING)
         CmdChr = v:
         sIniSetString("General", "CmdChr", CmdChr)
         me.SetINITimeStamp
      END PROPERTY

      GProp(CMDFlags, STRING)                                     ' CMD.EXE flags for CMD command
      PROPERTY SET CMDFlags(BYVAL v AS STRING)
         CMDFlags = v:
         sIniSetString("General", "CMDFlags", CmdFlags)
         me.SetINITimeStamp
      END PROPERTY

      GProp(CmdHiLite, LONG)                                      ' Cmd line KW hilite
      PROPERTY SET CmdHiLite(BYVAL v AS LONG)
         CmdHiLite = v:
         sIniSetString("Screen", "CmdHiLite", FORMAT$(CmdHiLite))
         me.SetINITimeStamp
      END PROPERTY

      GProp(ConfDelFlag, LONG)                                    ' Confirm Deletes in File Manager?
      PROPERTY SET ConfDelFlag(BYVAL v AS LONG)
         ConfDelFlag = v:
         sIniSetString("General", "CdelFlag", FORMAT$(ConfDelFlag))
         me.SetINITimeStamp
      END PROPERTY

      GProp(CursInsert, LONG)                                     ' Percent height of an Insert cursor
      PROPERTY SET CursInsert(BYVAL v AS LONG)
         CursInsert = v:
         sIniSetString("Screen", "CursIns", FORMAT$(CursInsert))
         me.SetINITimeStamp
      END PROPERTY

      GProp(CursNormal, LONG)                                     ' Percent height of a Normal cursor
      PROPERTY SET CursNormal(BYVAL v AS LONG)
         CursNormal = v:
         sIniSetString("Screen", "CursNorm", FORMAT$(CursNormal))
         me.SetINITimeStamp
      END PROPERTY

      GProp(DefaultShr, STRING)                                   ' Profile names using the DEFAULT Profile
      PROPERTY SET DefaultShr(BYVAL v AS STRING)
         DefaultShr = v:
         sIniSetString("General", "DefaultShr", DefaultShr)
         me.SetINITimeStamp
      END PROPERTY

      GProp(DefDataShift, LONG)                                   ' Default shift for >>)) ]] line commands
      PROPERTY SET DefDataShift(BYVAL v AS LONG)
         DefDataShift = v:
         sIniSetString("General", "DefDataShift", FORMAT$(DefDataShift))
         me.SetINITimeStamp
      END PROPERTY

      GProp(DirProfFlag, LONG)                                    ' DIR as Profile for no extensions files
      PROPERTY SET DirProfFlag(BYVAL v AS LONG)
         DirProfFlag = v:
         sIniSetString("General", "DProfFlag", FORMAT$(DirProfFlag))
         me.SetINITimeStamp
      END PROPERTY

      GProp(FindWord, LONG)                                       ' Find default context is WORD
      PROPERTY SET FindWord(BYVAL v AS LONG)
         FindWord = v:
         sIniSetString("General", "FindWord", FORMAT$(FindWord))
         me.SetINITimeStamp
      END PROPERTY

      GProp(FMCloseFlag, LONG)                                    ' Close FM with last tab?
      PROPERTY SET FMCloseFlag(BYVAL v AS LONG)
         FMCloseFlag = v:
         sIniSetString("General", "FMOpFlag", FORMAT$(FMCloseFlag))
         me.SetINITimeStamp
      END PROPERTY

      GProp(FMDateHiLite, LONG)                                   ' HiLite Recent/Active dates in FM list
      PROPERTY SET FMDateHiLite(BYVAL v AS LONG)
         FMDateHiLite = v:
         sIniSetString("General", "FMDateHiLite", FORMAT$(FMDateHiLite))
         me.SetINITimeStamp
      END PROPERTY

      GProp(FMDefMaster, STRING)                                  ' Master Default FM sort
      PROPERTY SET FMDefMaster(BYVAL v AS STRING)
         FMDefMaster = v:
         sIniSetString("FManager", "FMDefMaster", FMDefMaster)
         me.SetINITimeStamp
      END PROPERTY

      GProp(FMDefSort, STRING)                                    ' Default FM sort

      GProp(FMDirSort, STRING)                                    ' Default FM Dir sort

      GProp(FMFFList, STRING)                                     ' FM FF command exemption list
      PROPERTY SET FMFFList(BYVAL v AS STRING)
         FMFFList = v:
         sIniSetString("FManager", "FMFFList", FMFFList)
         me.SetINITimeStamp
      END PROPERTY

      GProp(FMFileList, STRING)                                   ' FM filepath
      PROPERTY SET FMFileList(BYVAL v AS STRING)
         FMFileList = v:
         sIniSetString("FManager", "Recall", FMFileList)
         me.SetINITimeStamp
      END PROPERTY

      GProp(FMHelpFlag, LONG)                                     ' Display FM Help at bottom of screen?
      PROPERTY SET FMHelpFlag(BYVAL v AS LONG)
         FMHelpFlag = v:
         sIniSetString("General", "FMHelpFlag", FORMAT$(FMHelpFlag))
         me.SetINITimeStamp
      END PROPERTY

      GProp(FMLayout, STRING)                                     ' FM Line layout
      PROPERTY SET FMLayout(BYVAL v AS STRING)
         FMLayout = v:
         sIniSetString("General", "FMLayout", FMLayout)
         me.SetINITimeStamp
      END PROPERTY

      GProp(FMLCmdWidth, LONG)                                    ' Width of FM Line Command  area
      PROPERTY SET FMLCmdWidth(BYVAL v AS LONG)
         FMLCmdWidth = v:
         sIniSetString("General", "FMLCmdWidth", FORMAT$(FMLCmdWidth))
         me.SetINITimeStamp
      END PROPERTY

      GProp(FMMask, STRING)                                       ' FM Mask
      PROPERTY SET FMMask(BYVAL v AS STRING)
         FMMask = v:
         sIniSetString("FManager", "DefTypes", FMMask)
         me.SetINITimeStamp
      END PROPERTY

      GProp(FMNoteWidth, LONG)                                    ' Width of FM Note column  area
      PROPERTY SET FMNoteWidth(BYVAL v AS LONG)
         FMNoteWidth = v:
         sIniSetString("General", "FMNoteWidth", FORMAT$(FMNoteWidth))
         me.SetINITimeStamp
      END PROPERTY

      GProp(FMPath, STRING)                                       ' FM Path
      PROPERTY SET FMPath(BYVAL v AS STRING)
         FMPath = v:
         sIniSetString("FManager", "DefDir1", FMPath)
         me.SetINITimeStamp
      END PROPERTY

      GProp(FMRecentSort, STRING)                                 ' Default FM Recent sort
      PROPERTY SET FMRecentSort(BYVAL v AS STRING)
         FMRecentSort = v:
         sIniSetString("FManager", "FMRecentSort", FMRecentSort)
         me.SetINITimeStamp
      END PROPERTY

      GProp(FMScrlAmt, STRING)                                    ' FM Scroll ammount
      PROPERTY SET FMScrlAmt(BYVAL v AS STRING)
         FMScrlAmt = v:
         sIniSetString("FManager", "ScrlAmtc", TRIM$(FMScrlAmt))
         me.SetINITimeStamp
      END PROPERTY

      GProp(FontName, STRING)                                     ' Font name for screen
      PROPERTY SET FontName(BYVAL v AS STRING)
         FontName = v:
         sIniSetString("Screen", "FontName", FontName)
         me.SetINITimeStamp
      END PROPERTY

      GProp(FontPitch, LONG)                                      ' Font pitch for screen
      PROPERTY SET FontPitch(BYVAL v AS LONG)
         FontPitch = v:
         sIniSetString("Screen", "FontPitch", FORMAT$(FontPitch))
         me.SetINITimeStamp
      END PROPERTY

      GProp(FontStyle, LONG)                                      ' Font style for screen
      PROPERTY SET FontStyle(BYVAL v AS LONG)
         FontStyle = v:
         sIniSetString("Screen", "FontStyle", FORMAT$(FontStyle))
         me.SetINITimeStamp
      END PROPERTY

      GProp(GlueWith, STRING)                                     ' GLUE join cgaracter string
      PROPERTY SET GlueWith(BYVAL v AS STRING)
         GlueWith = v:
         sIniSetString("General", "GlueWith", GlueWith)
         me.SetINITimeStamp
      END PROPERTY

      GProp(HRuler, LONG)                                         ' Horizontal ruler cursor
      PROPERTY SET HRuler(BYVAL v AS LONG)
         HRuler = v:
         sIniSetString("Screen", "HRuler", FORMAT$(HRuler))
         me.SetINITimeStamp
      END PROPERTY

      GProp(InsMode, LONG)                                        ' Default Insert mode
      PROPERTY SET InsMode(BYVAL v AS LONG)
         InsMode = v:
         sIniSetString("KeyBoard", "InsMode", FORMAT$(InsMode))
         me.SetINITimeStamp
      END PROPERTY

      GProp(InsReset, LONG)                                       ' Insert mode Reset at each Attn key?
      PROPERTY SET InsReset(BYVAL v AS LONG)
         InsReset = v:
         sIniSetString("KeyBoard", "InsReset", FORMAT$(InsReset))
         me.SetINITimeStamp
      END PROPERTY

      GProp(InvChar, STRING)                                      ' Display invalid character
      PROPERTY SET InvChar(BYVAL v AS STRING)
         InvChar = v:
         sIniSetString("General", "InvChar", InvChar)
         me.SetINITimeStamp
      END PROPERTY

      GProp(JobNumber, LONG)                                      ' Submit job number
      PROPERTY SET JobNumber(BYVAL v AS LONG)
         JobNumber = v:
         sIniSetString("General", "JobNumber", FORMAT$(Jobnumber))
         me.SetINITimeStamp
      END PROPERTY

      GProp(KBHorScroll, LONG)                                    ' Scroll (Hor) screen when cursor hits edge?
      PROPERTY SET KBHorScroll(BYVAL v AS LONG)
         KBHorScroll = v:
         sIniSetString("Keyboard", "KBScrollH", FORMAT$(KBHorScroll))
         me.SetINITimeStamp
      END PROPERTY

      GProp(KBVerScroll, LONG)                                    ' Scroll (Ver) screen when cursor hits edge?
      PROPERTY SET KBVerScroll(BYVAL v AS LONG)
         KBVerScroll = v:
         sIniSetString("Keyboard", "KBScrollV", FORMAT$(KBVerScroll))
         me.SetINITimeStamp
      END PROPERTY

      GProp(LastDirOpen, LONG)                                    ' What directory to default OPEN to
      PROPERTY SET LastDirOpen(BYVAL v AS LONG)
         LastDirOpen = v:
         sIniSetString("General", "LastDirOpen", FORMAT$(LastDirOpen))
         me.SetINITimeStamp
      END PROPERTY

      GProp(LastScreenX, LONG)                                    ' Last screen location X
      PROPERTY SET LastScreenX(BYVAL v AS LONG)
         LastScreenX = v:
         sIniSetString("Screen", "LastScrX", FORMAT$(LastScreenX))
         me.SetINITimeStamp
      END PROPERTY

      GProp(LastScreenY, LONG)                                    ' Last screen location Y
      PROPERTY SET LastScreenY(BYVAL v AS LONG)
         LastScreenY = v:
         sIniSetString("Screen", "LastScrY", FORMAT$(LastScreenY))
         me.SetINITimeStamp
      END PROPERTY

      GProp(LinNoSize, LONG)                                      ' Width of Line Number field
      PROPERTY SET LinNoSize(BYVAL v AS LONG)
         LinNoSize = v:
         sIniSetString("Screen", "LNoSize", FORMAT$(LinNoSize))
         me.SetINITimeStamp
      END PROPERTY

      GProp(LinRepeatMax, LONG)                                   ' Max Line repeat
      PROPERTY SET LinRepeatMax(BYVAL v AS LONG)
         LinRepeatMax = v:
         sIniSetString("General", "LineRepeatMax", FORMAT$(LinRepeatMax))
         me.SetINITimeStamp
      END PROPERTY

      GProp(MAXIMIZED, LONG)                                      ' Screen is in Max mode
      PROPERTY SET MAXIMIZED(BYVAL v AS LONG)
         MAXIMIZED = v:
         sIniSetString("Screen", "Maximized", FORMAT$(MAXIMIZED))
         me.SetINITimeStamp
      END PROPERTY

      GProp(MinRetrieve, LONG)                                    ' Min length retrievable command
      PROPERTY SET MinRetrieve(BYVAL v AS LONG)
         MinRetrieve = v:
         sIniSetString("General", "MinRetrieve", FORMAT$(MinRetrieve))
         me.SetINITimeStamp
      END PROPERTY

      GProp(ENGChars, LONG)                                       ' English only A-Z and a-z
      PROPERTY SET ENGChars(BYVAL v AS LONG)
         ENGChars = v:
         sIniSetString("General", "ENGChars", FORMAT$(ENGChars))
         me.SetINITimeStamp
      END PROPERTY

      GProp(NotifyLevel, LONG)                                    ' Notify option
      PROPERTY SET NotifyLevel(BYVAL v AS LONG)
         NotifyLevel = v: NotifyLevelT = v
         sIniSetString("General", "Notify", FORMAT$(NotifyLevel))
         me.SetINITimeStamp
      END PROPERTY

      GSProp(NotifyLevelT, LONG)                                  ' Working Notify option

      GProp(PFKShow, LONG)                                        ' Show PFKeys at bottom
      PROPERTY SET PFKShow(BYVAL v AS LONG)
         PFKShow = v:
         sIniSetString("Screen", "KBHelp", FORMAT$(PFKShow))
         me.SetINITimeStamp
      END PROPERTY

      GProp(PrtBandColor, LONG)                                   ' Banding color
      PROPERTY SET PrtBandColor(BYVAL v AS LONG)
         PrtBandColor = v:
         sIniSetString("Printer", "BandingColor", FORMAT$(PrtBandColor))
         me.SetINITimeStamp
      END PROPERTY

      GProp(PrtBanding, LONG)                                     ' Banding stripes
      PROPERTY SET PrtBanding(BYVAL v AS LONG)
         PrtBanding = v:
         sIniSetString("Printer", "Banding", FORMAT$(PrtBanding))
         me.SetINITimeStamp
      END PROPERTY

      GProp(PrtBandLines, LONG)                                   ' Banding lines
      PROPERTY SET PrtBandLines(BYVAL v AS LONG)
         PrtBandLines = v:
         sIniSetString("Printer", "BandLines", FORMAT$(PrtBandLines))
         me.SetINITimeStamp
      END PROPERTY

      GProp(PrtBMargin, SINGLE)                                   ' Printer bottom margin (in)
      PROPERTY SET PrtBMargin(BYVAL v AS SINGLE)
         PrtBMargin = v:
         sIniSetString("Printer", "BM", FORMAT$(PrtBMargin))
         me.SetINITimeStamp
      END PROPERTY

      GProp(PrtDuplex, LONG)                                      ' Printer duplex setting
      PROPERTY SET PrtDuplex(BYVAL v AS LONG)
         PrtDuplex = v:
         sIniSetString("Printer", "Duplex", FORMAT$(PrtDuplex))
         me.SetINITimeStamp
      END PROPERTY

      GProp(PrtFontName, STRING)                                  ' Printer font
      PROPERTY SET PrtFontName(BYVAL v AS STRING)
         PrtFontName = v:
         sIniSetString("Printer", "Font", PrtFontName)
         me.SetINITimeStamp
      END PROPERTY

      GProp(PrtFontPitch, STRING)                                 ' Printer pitch
      PROPERTY SET PrtFontPitch(BYVAL v AS STRING)
         PrtFontPitch = v:
         sIniSetString("Printer", "Pitch", PrtFontPitch)
         me.SetINITimeStamp
      END PROPERTY

      GProp(PrtFontStyle, STRING)                                 ' Printer Style
      PROPERTY SET PrtFontStyle(BYVAL v AS STRING)
         PrtFontStyle = v:
         sIniSetString("Printer", "Style", PrtFontStyle)
         me.SetINITimeStamp
      END PROPERTY

      GProp(PrtFooter, LONG)                                      ' Printer footers 1 = yes
      PROPERTY SET PrtFooter(BYVAL v AS LONG)
         PrtFooter = v:
         sIniSetString("Printer", "Footers", FORMAT$(PrtFooter))
         me.SetINITimeStamp
      END PROPERTY

      GProp(PrtFooterCenter, STRING)                              ' Printer center footers
      PROPERTY SET PrtFooterCenter(BYVAL v AS STRING)
         PrtFooterCenter = v:
         sIniSetString("Printer", "FCenter", PrtFooterCenter)
         me.SetINITimeStamp
      END PROPERTY

      GProp(PrtFooterLeft, STRING)                                ' Printer left footers
      PROPERTY SET PrtFooterLeft(BYVAL v AS STRING)
         PrtFooterLeft = v:
         sIniSetString("Printer", "FLeft", PrtFooterLeft)
         me.SetINITimeStamp
      END PROPERTY

      GProp(PrtFooterRight, STRING)                               ' Printer right footers
      PROPERTY SET PrtFooterRight(BYVAL v AS STRING)
         PrtFooterRight = v:
         sIniSetString("Printer", "FRight", PrtFooterRight)
         me.SetINITimeStamp
      END PROPERTY

      GProp(PrtHeader, LONG)                                      ' Printer headers 1 = yes
      PROPERTY SET PrtHeader(BYVAL v AS LONG)
         PrtHeader = v:
         sIniSetString("Printer", "Headers", FORMAT$(PrtHeader))
         me.SetINITimeStamp
      END PROPERTY

      GProp(PrtHeaderCenter, STRING)                              ' Printer center headers
      PROPERTY SET PrtHeaderCenter(BYVAL v AS STRING)
         PrtHeaderCenter = v:
         sIniSetString("Printer", "HCenter", PrtHeaderCenter)
         me.SetINITimeStamp
      END PROPERTY

      GProp(PrtHeaderLeft, STRING)                                ' Printer left headers
      PROPERTY SET PrtHeaderLeft(BYVAL v AS STRING)
         PrtHeaderLeft = v:
         sIniSetString("Printer", "HLeft", PrtHeaderLeft)
         me.SetINITimeStamp
      END PROPERTY

      GProp(PrtHeaderRight, STRING)                               ' Printer left headers
      PROPERTY SET PrtHeaderRight(BYVAL v AS STRING)
         PrtHeaderRight = v:
         sIniSetString("Printer", "HRight", PrtHeaderRight)
         me.SetINITimeStamp
      END PROPERTY

      GProp(PrtLMargin, SINGLE)                                   ' Printer left margin (in)
      PROPERTY SET PrtLMargin(BYVAL v AS SINGLE)
         PrtLMargin = v:
         sIniSetString("Printer", "LM", FORMAT$(PrtLMargin))
         me.SetINITimeStamp
      END PROPERTY

      GProp(PrtMetric, LONG)                                      ' Printer margins Eng/Metric
      PROPERTY SET PrtMetric(BYVAL v AS LONG)
         PrtMetric = v:
         sIniSetString("Printer", "Metric", FORMAT$(PrtMetric))
         me.SetINITimeStamp
      END PROPERTY

      GProp(PrtName, STRING)                                      ' Printer name
      PROPERTY SET PrtName(BYVAL v AS STRING)
         PrtName = v:
         sIniSetString("Printer", "Name", PrtName)
         me.SetINITimeStamp
      END PROPERTY

      GProp(PrtOrient, LONG)                                      ' Printer orientation, 1=Portrait 2=Landscape
      PROPERTY SET PrtOrient(BYVAL v AS LONG)
         PrtOrient = v:
         sIniSetString("Printer", "Orient", FORMAT$(PrtOrient))
         me.SetINITimeStamp
      END PROPERTY

      GProp(PrtPaper, LONG)                                       ' Paper index
      PROPERTY SET PrtPaper(BYVAL v AS LONG)
         PrtPaper = v:
         sIniSetString("Printer", "Paper", FORMAT$(PrtPaper))
         me.SetINITimeStamp
      END PROPERTY

      GProp(PrtRMargin, SINGLE)                                   ' Printer right margin (in)
      PROPERTY SET PrtRMargin(BYVAL v AS SINGLE)
         PrtRMargin = v:
         sIniSetString("Printer", "RM", FORMAT$(PrtRMargin))
         me.SetINITimeStamp
      END PROPERTY

      GProp(PrtTMargin, SINGLE)                                   ' Printer top margin (in)
      PROPERTY SET PrtTMargin(BYVAL v AS SINGLE)
         PrtTMargin = v:
         sIniSetString("Printer", "TM", FORMAT$(PrtTMargin))
         me.SetINITimeStamp
      END PROPERTY

      GProp(PrtPColor, LONG)                                      ' Print in color (%True or %False)
      PROPERTY SET PrtPColor(BYVAL v AS LONG)
         PrtPColor = v:
         sIniSetString("Printer", "Color", FORMAT$(PrtPColor))
         me.SetINITimeStamp
      END PROPERTY

      GProp(QuickRenum, LONG)                                     ' Quick renum over this many lines
      PROPERTY SET QuickRenum(BYVAL v AS LONG)
         QuickRenum = v:
         sIniSetString("General", "QuickRenum", FORMAT$(QuickRenum))
         me.SetINITimeStamp
      END PROPERTY

      GProp(RecentCtr, LONG)                                      ' How many recent files to keep
      PROPERTY SET RecentCtr(BYVAL v AS LONG)
         RecentCtr = v:
         sIniSetString("General", "RecentCtr", FORMAT$(RecentCtr))
         me.SetINITimeStamp
      END PROPERTY

      GProp(ReOpenLast, LONG)                                     ' Re-Open last file(s)
      PROPERTY SET ReOpenLast(BYVAL v AS LONG)
         ReOpenLast = v:
         sIniSetString("General", "ROpenLast", FORMAT$(ReOpenLast))
         me.SetINITimeStamp
      END PROPERTY

      GProp(ResetU, LONG)                                         ' RESET clears U by default, like X
      PROPERTY SET ResetU(BYVAL v AS LONG)
         ResetU = v:
         sIniSetString("General", "ResetU", FORMAT$(ResetU))
         me.SetINITimeStamp
      END PROPERTY

      GProp(RUNFlags, STRING)                                     ' CMD.EXE flags for RUN command
      PROPERTY SET RUNFlags(BYVAL v AS STRING)
         RUNFlags = v:
         sIniSetString("General", "RUNFlags", RunFlags)
         me.SetINITimeStamp
      END PROPERTY

      GProp(SBLayout, STRING)                                     ' Status Bar layout
      PROPERTY SET SBLayout(BYVAL v AS STRING)
         SBLayout = v:
         sIniSetString("General", "SBLayout", SBLayout)
         me.SetINITimeStamp
      END PROPERTY

      GProp(ScrHeight, LONG)                                      ' Screen height
      PROPERTY SET ScrHeight(BYVAL v AS LONG)
         ScrHeight = v:
         sIniSetString("Screen", "ScrHeight", FORMAT$(ScrHeight))
         me.SetINITimeStamp
      END PROPERTY

      GProp(ScrWidth, LONG)                                       ' Screen width
      PROPERTY SET ScrWidth(BYVAL v AS LONG)
         ScrWidth = v:
         sIniSetString("Screen", "ScrWidth", FORMAT$(ScrWidth))
         me.SetINITimeStamp
      END PROPERTY

      GProp(Splash, LONG)                                         ' Splash
      PROPERTY SET Splash(BYVAL v AS LONG)
         Splash = v:
         sIniSetString("General", "Splash", FORMAT$(Splash))
         me.SetINITimeStamp
      END PROPERTY

      GProp(SubmitCmd, STRING)                                    ' Model command for SUBMIT
      PROPERTY SET SubmitCmd(BYVAL v AS STRING)
         SubmitCmd = v:
         sIniSetString("General", "SubmitCmd", SubmitCmd)
         me.SetINITimeStamp
      END PROPERTY

      GProp(SubmitDir, STRING)                                    ' Working DIR for SUBMIT
      PROPERTY SET SubmitDir(BYVAL v AS STRING)
         SubmitDir = v:
         sIniSetString("General", "SubmitDir", SubmitDir)
         me.SetINITimeStamp
      END PROPERTY

      GProp(SubmitInclKey, STRING)                                ' Include Key for SUBMIT
      PROPERTY SET SubmitInclKey(BYVAL v AS STRING)
         SubmitInclKey = v:
         sIniSetString("General", "SubmitInclKey", SubmitInclKey)
         me.SetINITimeStamp
      END PROPERTY

      GProp(SubmitInclCol, LONG)                                  ' Include Key Column for SUBMIT
      PROPERTY SET SubmitInclCol(BYVAL v AS LONG)
         SubmitInclCol = v:
         sIniSetString("General", "SubmitInclCol", FORMAT$(SubmitInclCol))
         me.SetINITimeStamp
      END PROPERTY

      GProp(UniqueFlag, LONG)                                     ' Keep only 1 instance running
      PROPERTY SET UniqueFlag(BYVAL v AS LONG)
         UniqueFlag = v:
         sIniSetString("General", "UniqFlag", FORMAT$(UniqueFlag))
         me.SetINITimeStamp
      END PROPERTY

      GProp(UseRecycle, LONG)                                     ' Delete to Recycle bin?
      PROPERTY SET UseRecycle(BYVAL v AS LONG)
         UseRecycle = v:
         sIniSetString("General", "UseRecycle", FORMAT$(UseRecycle))
         me.SetINITimeStamp
      END PROPERTY

      GProp(VBeepFlag, LONG)                                      ' Visual Beep on errFlag=%eFail? True = yes
      PROPERTY SET VBeepFlag(BYVAL v AS LONG)
         VBeepFlag = v:
         sIniSetString("General", "VBeepFlag", FORMAT$(VBeepFlag))
         me.SetINITimeStamp
      END PROPERTY

      GProp(VertInsCurs, LONG)                                    ' Vertical Insert Cursor
      PROPERTY SET VertInsCurs(BYVAL v AS LONG)
         VertInsCurs = v:
         sIniSetString("Screen", "CursInsVert", FORMAT$(VertInsCurs))
         me.SetINITimeStamp
      END PROPERTY

      GProp(ViewWarn, LONG)                                       ' Warn on modified View?
      PROPERTY SET ViewWarn(BYVAL v AS LONG)
         ViewWarn = v:
         sIniSetString("General", "BrowseWarn", FORMAT$(ViewWarn))
         me.SetINITimeStamp
      END PROPERTY

      GProp(VRuler, LONG)                                         ' Vertical ruler cursor
      PROPERTY SET VRuler(BYVAL v AS LONG)
         VRuler = v:
         sIniSetString("Screen", "VRuler", FORMAT$(VRuler))
         me.SetINITimeStamp
      END PROPERTY

      METHOD GetClr(RScheme AS LONG, which AS LONG) AS LONG
      '---------- Return color value for a particular scheme entry
         IF RScheme <> -1 THEN METHOD = Scheme(RScheme, which): EXIT METHOD ' Not %SCCust, Pass it back
         METHOD = CHOOSE(which, cCustFG, cCustBG1, cCustBG2)      ' Pass back a custom value
      END METHOD

      METHOD SetClr(RScheme AS LONG, which AS LONG, CValue AS LONG) AS LONG
      '---------- Save color value for a particular scheme entry
         Scheme(RScheme, which) = CValue                          ' Save it
      END METHOD

      METHOD SetGlobalClrs()
      '---------- Copy Scheme colors to Globals (easier access)
      REGISTER i AS LONG
      REGISTER j AS LONG
         i = VARPTR(cLNoHiFG)
         DIM GValues(1 TO 48) AS LONG AT i
         j = 1
         FOR i = 16 TO 31                                         ' Copy the Scheme 16 to 31 entries
            gValues(j)   = Scheme(i, 1)                           ' FG
            gValues(j+1) = Scheme(i, 2)                           ' BG1
            gValues(j+2) = Scheme(i, 3)                           ' BG2
            j += 3                                                '
         NEXT i                                                   '
         cTxtLoFG = scheme(0, 1): cTxtLoBG1 = scheme(0, 2): cTxtLoBG2 = scheme(0, 3)
         cTxtHiFG = scheme(1, 1): cTxtHiBG1 = scheme(1, 2): cTxtHiBG2 = scheme(1, 3)

      END METHOD

      METHOD SaveScheme()
         MAT SchemeSave() = Scheme()                              ' Save Scheme array 'as is' in case cancelled
      END METHOD

      METHOD LoadScheme()
         MAT Scheme() = SchemeSave()                              ' Reload saved Scheme array
      END METHOD

      METHOD GetFMSort(which AS STRING, part AS LONG) AS STRING
      '---------- Return one of the Master FM sort controls
      LOCAL section(), portion(), value AS STRING, i, j, k AS LONG
         DIM section(1 TO PARSECOUNT(FMDefMaster, " "))           ' Dim array to hold sections
         PARSE FMDefMaster, section(), " "                        ' Split out the sections
         FOR i = 1 TO UBOUND(section())                           ' Find the requested one
            IF LEFT$(section(i), LEN(which)) = which THEN         ' Is this the one
               FMLastSectNo = i                                   ' Save last section index
               j = INSTR(section(i), "(")                         ' Find operands
               k = INSTR(section(i), ")")                         '
               value = MID$(section(i), j + 1 TO k - 1)           ' Extract the (...) value
               DIM portion(1 TO PARSECOUNT(value, ","))           ' Dim the portion array
               PARSE value, portion(), ","                        ' Break it out
               METHOD = portion(part)                             ' Pass back the answer
               EXIT METHOD                                        '
            END IF                                                '
         NEXT i                                                   '
      END METHOD

      METHOD SchemeSaveAll()
      '---------- Save all the Scheme values
      LOCAL i AS LONG

         FOR i = 0 TO 47
            sIniSetString("Scheme", "Scheme" + FORMAT$(i, "00") + "FG", FORMAT$(Scheme(i, 1)))
            sIniSetString("Scheme", "Scheme" + FORMAT$(i, "00") + "BG1", FORMAT$(Scheme(i, 2)))
            sIniSetString("Scheme", "Scheme" + FORMAT$(i, "00") + "BG2", FORMAT$(Scheme(i, 3)))
         NEXT i
      END METHOD

      METHOD SetFMLast(nv1 AS STRING, nv2 AS STRING)
      '---------- Change the last used (current) master section of FM sorts
      LOCAL section(), portion(), value AS STRING, i, j, k AS LONG
         DIM section(1 TO PARSECOUNT(FMDefMaster, " "))           ' Dim array to hold sections
         PARSE FMDefMaster, section(), " "                        ' Split out the sections
         j = INSTR(section(FMLastSectNo), "(")                    ' Find operands
         k = INSTR(section(FMLastSectNo), ")")                    '
         value = nv1 + "," + nv2                                  ' Build the new value
         section(FMLastSectNo) = LEFT$(section(FMLastSectNo), j) + value + ")"    ' Rebuild section
         FMDefMaster = JOIN$(section(), " ")                      ' Rebuild FMDefMaster
         sIniSetString("FManager", "FMDefMaster", FMDefMaster)    ' Save it
         me.SetINITimeStamp                                       '
      END METHOD

      METHOD LoadAllINI()
      '---------- Get the Global Preferences INI values
      LOCAL i, j, ConvSave AS LONG, TX AS STRING

         '----- See if conversion to new Color Schemes has been done
         IF sINIGetString("Scheme", "Scheme00FG",  "MISSING") = "MISSING" THEN    ' See if SCTxtHi value is present
            me.INIColorCnvt                                       ' No?  Go convert existing colors
            ConvSave = %True                                      ' Do a save all later
         END IF                                                   '

         '----- Get Colour Schemes loaded
         Scheme(00, 1) = VAL(sINIGetString("Scheme", "Scheme00FG",  "&HFF6464"))   ' SCTxtLo
         Scheme(00, 2) = VAL(sINIGetString("Scheme", "Scheme00BG1", "&HFFFFFF"))
         Scheme(00, 3) = VAL(sINIGetString("Scheme", "Scheme00BG2", "&HD9FFD9"))
         Scheme(01, 1) = VAL(sINIGetString("Scheme", "Scheme01FG",  "&H000000"))   ' SCTxtHi
         Scheme(01, 2) = VAL(sINIGetString("Scheme", "Scheme01BG1", "&HFFFFFF"))
         Scheme(01, 3) = VAL(sINIGetString("Scheme", "Scheme01BG2", "&HD9FFD9"))

         FOR i = 2 TO 15                                          ' User colorize range
            Scheme(i, 1) = VAL(sINIGetString("Scheme", "Scheme" + FORMAT$(i, "00") + "FG", "&H000000"))
            Scheme(i, 2) = VAL(sINIGetString("Scheme", "Scheme" + FORMAT$(i, "00") + "BG1", "&HFFFFFF"))
            Scheme(i, 3) = VAL(sINIGetString("Scheme", "Scheme" + FORMAT$(i, "00") + "BG2", "&HD9FFD9"))
         NEXT i

         '----- Get 'Standard' Schemes loaded
         Scheme(16, 1) = VAL(sINIGetString("Scheme", "Scheme16FG",  "&HA00000"))   ' SCLnoHi
         Scheme(16, 2) = VAL(sINIGetString("Scheme", "Scheme16BG1", "&HFFFFFF"))
         Scheme(16, 3) = VAL(sINIGetString("Scheme", "Scheme16BG2", "&HD9FFD9"))
         Scheme(17, 1) = VAL(sINIGetString("Scheme", "Scheme17FG",  "&HB5B5B5"))   ' SCLnoLo
         Scheme(17, 2) = VAL(sINIGetString("Scheme", "Scheme17BG1", "&HFFFFFF"))
         Scheme(17, 3) = VAL(sINIGetString("Scheme", "Scheme17BG2", "&HD9FFD9"))
         Scheme(18, 1) = VAL(sINIGetString("Scheme", "Scheme18FG",  "&H0000FF"))   ' SCATabMod
         Scheme(18, 2) = VAL(sINIGetString("Scheme", "Scheme18BG1", "&H00FFFF"))
         Scheme(18, 3) = VAL(sINIGetString("Scheme", "Scheme18BG2", "&H00FFFF"))
         Scheme(19, 1) = VAL(sINIGetString("Scheme", "Scheme19FG",  "&HFF46A3"))   ' SCATabNMod
         Scheme(19, 2) = VAL(sINIGetString("Scheme", "Scheme19BG1", "&H00FFFF"))
         Scheme(19, 3) = VAL(sINIGetString("Scheme", "Scheme19BG2", "&H00FFFF"))
         Scheme(20, 1) = VAL(sINIGetString("Scheme", "Scheme20FG",  "&H0000FF"))   ' SCITabMod
         Scheme(20, 2) = VAL(sINIGetString("Scheme", "Scheme20BG1", "&HEFEFEF"))
         Scheme(20, 3) = VAL(sINIGetString("Scheme", "Scheme20BG2", "&HEFEFEF"))
         Scheme(21, 1) = VAL(sINIGetString("Scheme", "Scheme21FG",  "&HFF46A3"))   ' SCITabNMod
         Scheme(21, 2) = VAL(sINIGetString("Scheme", "Scheme21BG1", "&HEFEFEF"))
         Scheme(21, 3) = VAL(sINIGetString("Scheme", "Scheme21BG2", "&HEFEFEF"))
         Scheme(22, 1) = VAL(sINIGetString("Scheme", "Scheme22FG",  "&H757575"))   ' SCPFK
         Scheme(22, 2) = VAL(sINIGetString("Scheme", "Scheme22BG1", "&HE1DFCE"))
         Scheme(22, 3) = VAL(sINIGetString("Scheme", "Scheme22BG2", "&HE1DFCE"))
         Scheme(23, 1) = VAL(sINIGetString("Scheme", "Scheme23FG",  "&H000000"))   ' SCStatus
         Scheme(23, 2) = VAL(sINIGetString("Scheme", "Scheme23BG1", "&HE2E2E2"))
         Scheme(23, 3) = VAL(sINIGetString("Scheme", "Scheme23BG2", "&HE2E2E2"))
         Scheme(24, 1) = VAL(sINIGetString("Scheme", "Scheme24FG",  "&HFFFFFF"))   ' SCFMTool
         Scheme(24, 2) = VAL(sINIGetString("Scheme", "Scheme24BG1", "&HA39D6B"))
         Scheme(24, 3) = VAL(sINIGetString("Scheme", "Scheme24BG2", "&HA39D6B"))
         Scheme(25, 1) = VAL(sINIGetString("Scheme", "Scheme25FG",  "&H0000FF"))   ' SCError
         Scheme(25, 2) = VAL(sINIGetString("Scheme", "Scheme25BG1", "&HFFFFFF"))
         Scheme(25, 3) = VAL(sINIGetString("Scheme", "Scheme25BG2", "&HD9FFD9"))
         Scheme(26, 1) = VAL(sINIGetString("Scheme", "Scheme26FG2", "&H000000"))   ' Rsvrd6
         Scheme(26, 2) = VAL(sINIGetString("Scheme", "Scheme26BG1", "&HC5C5C5"))
         Scheme(26, 3) = VAL(sINIGetString("Scheme", "Scheme26BG2", "&HD9FFD9"))
         Scheme(27, 1) = VAL(sINIGetString("Scheme", "Scheme27FG",  "&H000000"))   ' Rsvrd5
         Scheme(27, 2) = VAL(sINIGetString("Scheme", "Scheme27BG1", "&HFFFFFF"))
         Scheme(27, 3) = VAL(sINIGetString("Scheme", "Scheme27BG2", "&HD9FFD9"))
         Scheme(28, 1) = VAL(sINIGetString("Scheme", "Scheme28FG",  "&H000000"))   ' Rsvrd4
         Scheme(28, 2) = VAL(sINIGetString("Scheme", "Scheme28BG1", "&HFFFFFF"))
         Scheme(28, 3) = VAL(sINIGetString("Scheme", "Scheme28BG2", "&HD9FFD9"))
         Scheme(29, 1) = VAL(sINIGetString("Scheme", "Scheme29FG",  "&H000000"))   ' Rsvrd3
         Scheme(29, 2) = VAL(sINIGetString("Scheme", "Scheme29BG1", "&HFFFFFF"))
         Scheme(29, 3) = VAL(sINIGetString("Scheme", "Scheme29BG2", "&HD9FFD9"))
         Scheme(30, 1) = VAL(sINIGetString("Scheme", "Scheme30FG",  "&H000000"))   ' Rsvrd2
         Scheme(30, 2) = VAL(sINIGetString("Scheme", "Scheme30BG1", "&HFFFFFF"))
         Scheme(30, 3) = VAL(sINIGetString("Scheme", "Scheme30BG2", "&HD9FFD9"))
         Scheme(31, 1) = VAL(sINIGetString("Scheme", "Scheme31FG",  "&H000000"))   ' Rsvrd1
         Scheme(31, 2) = VAL(sINIGetString("Scheme", "Scheme31BG1", "&HFFFFFF"))
         Scheme(31, 3) = VAL(sINIGetString("Scheme", "Scheme31BG2", "&HD9FFD9"))

         '----- Get Hi-Lite' Schemes loaded

         Scheme(32, 1) = VAL(sINIGetString("Scheme", "Scheme32FG",  "&HFFFFFF"))   ' SCBlue
         Scheme(32, 2) = VAL(sINIGetString("Scheme", "Scheme32BG1", "&HFF0000"))
         Scheme(32, 3) = VAL(sINIGetString("Scheme", "Scheme32BG2", "&HFF0000"))
         Scheme(33, 1) = VAL(sINIGetString("Scheme", "Scheme33FG",  "&HFFFFFF"))   ' SCGreen
         Scheme(33, 2) = VAL(sINIGetString("Scheme", "Scheme33BG1", "&H649632"))
         Scheme(33, 3) = VAL(sINIGetString("Scheme", "Scheme33BG2", "&H649632"))
         Scheme(34, 1) = VAL(sINIGetString("Scheme", "Scheme34FG",  "&H000000"))   ' SCYellow
         Scheme(34, 2) = VAL(sINIGetString("Scheme", "Scheme34BG1", "&H00FFFF"))
         Scheme(34, 3) = VAL(sINIGetString("Scheme", "Scheme34BG2", "&H00FFFF"))
         Scheme(35, 1) = VAL(sINIGetString("Scheme", "Scheme35FG",  "&HFFFFFF"))   ' SCRed
         Scheme(35, 2) = VAL(sINIGetString("Scheme", "Scheme35BG1", "&H0000F0"))
         Scheme(35, 3) = VAL(sINIGetString("Scheme", "Scheme35BG2", "&H0000F0"))
         Scheme(36, 1) = VAL(sINIGetString("Scheme", "Scheme36FG",  "&HFFFFFF"))   ' SCBlack
         Scheme(36, 2) = VAL(sINIGetString("Scheme", "Scheme36BG1", "&H000000"))
         Scheme(36, 3) = VAL(sINIGetString("Scheme", "Scheme36BG2", "&H000000"))
         Scheme(37, 1) = VAL(sINIGetString("Scheme", "Scheme37FG",  "&HFFFFFF"))   ' SCNavy
         Scheme(37, 2) = VAL(sINIGetString("Scheme", "Scheme37BG1", "&HD20000"))
         Scheme(37, 3) = VAL(sINIGetString("Scheme", "Scheme37BG2", "&HD20000"))
         Scheme(38, 1) = VAL(sINIGetString("Scheme", "Scheme38FG",  "&HFFFFFF"))   ' SCTeal
         Scheme(38, 2) = VAL(sINIGetString("Scheme", "Scheme38BG1", "&HC8C800"))
         Scheme(38, 3) = VAL(sINIGetString("Scheme", "Scheme38BG2", "&HC8C800"))
         Scheme(39, 1) = VAL(sINIGetString("Scheme", "Scheme39FG",  "&HFFFFFF"))   ' SCViolet
         Scheme(39, 2) = VAL(sINIGetString("Scheme", "Scheme39BG1", "&H8B008B"))
         Scheme(39, 3) = VAL(sINIGetString("Scheme", "Scheme39BG2", "&H8B008B"))
         Scheme(40, 1) = VAL(sINIGetString("Scheme", "Scheme40FG",  "&H000000"))   ' SCOrange
         Scheme(40, 2) = VAL(sINIGetString("Scheme", "Scheme40BG1", "&H00B4FF"))
         Scheme(40, 3) = VAL(sINIGetString("Scheme", "Scheme40BG2", "&H00B4FF"))
         Scheme(41, 1) = VAL(sINIGetString("Scheme", "Scheme41FG",  "&H000000"))   ' SCGray
         Scheme(41, 2) = VAL(sINIGetString("Scheme", "Scheme41BG1", "&HB4B4B4"))
         Scheme(41, 3) = VAL(sINIGetString("Scheme", "Scheme41BG2", "&HB4B4B4"))
         Scheme(42, 1) = VAL(sINIGetString("Scheme", "Scheme42FG",  "&H000000"))   ' SCLime
         Scheme(42, 2) = VAL(sINIGetString("Scheme", "Scheme42BG1", "&H00FF00"))
         Scheme(42, 3) = VAL(sINIGetString("Scheme", "Scheme42BG2", "&H00FF00"))
         Scheme(43, 1) = VAL(sINIGetString("Scheme", "Scheme43FG",  "&H000000"))   ' SCCyan
         Scheme(43, 2) = VAL(sINIGetString("Scheme", "Scheme43BG1", "&HFFFF00"))
         Scheme(43, 3) = VAL(sINIGetString("Scheme", "Scheme43BG2", "&HFFFF00"))
         Scheme(44, 1) = VAL(sINIGetString("Scheme", "Scheme44FG",  "&H000000"))   ' SCPink
         Scheme(44, 2) = VAL(sINIGetString("Scheme", "Scheme44BG1", "&HFFC8FF"))
         Scheme(44, 3) = VAL(sINIGetString("Scheme", "Scheme44BG2", "&HFFC8FF"))
         Scheme(45, 1) = VAL(sINIGetString("Scheme", "Scheme45FG",  "&HFFFFFF"))   ' SCMagenta
         Scheme(45, 2) = VAL(sINIGetString("Scheme", "Scheme45BG1", "&HFF00FF"))
         Scheme(45, 3) = VAL(sINIGetString("Scheme", "Scheme45BG2", "&HFF00FF"))
         Scheme(46, 1) = VAL(sINIGetString("Scheme", "Scheme46FG",  "&H000000"))   ' SCWhite
         Scheme(46, 2) = VAL(sINIGetString("Scheme", "Scheme46BG1", "&HFFFFFF"))
         Scheme(46, 3) = VAL(sINIGetString("Scheme", "Scheme46BG2", "&HFFFFFF"))

         ABeepFlag          = VAL(sINIGetString("General",  "BeepFlag",     "1"))
         Allow2DMouse       = VAL(sINIGetString("General",  "Allow2DMouse", "0"))
         AutoScroll         = VAL(sINIGetString("Mouse",    "AutoScrl",     "2"))
         Banding            = VAL(sINIGetString("Screen",   "Banding",      "0"))
         ENGChars           = VAL(sINIGetString("General",  "ENGChars",     "0"))        ' Must be before Charset
         CharSet            = sINIGetString("General",      "CharSet",      "**MISSING**")

         IF CharSet = "**MISSING**" THEN                          ' Initial setup?
            IF ENGChars THEN                                      ' Default to use ENGLISH-ONLY characters
               CharSet = $DefCharSet                              ' Set it with ENGLISH chars
            ELSE                                                  '
               CharSet = $DefCharSet + $UpperNat + $LowerNat      ' Set it with NAT
            END IF                                                '
         END IF                                                   '

         CharSet = REMOVE$(CharSet, " ")                          ' Charset should not have blanks

         IF ISNULL(CharSet) THEN                                  ' Again if null
            IF ENGChars THEN                                      ' Default to use ENGLISH-ONLY characters
               CharSet = $DefCharSet                              ' Set it with ENGLISH
            ELSE                                                  '
               CharSet = $DefCharSet + $UpperNat + $LowerNat      ' Set it with NAT
            END IF                                                '
         END IF                                                   '

         IF ENGChars THEN                                         ' Default to use ENGLISH characters
            gUpper = $Upper: gLower = $Lower                      ' Set working $Upper and $Lower
         ELSE                                                     '
            gUpper = $Upper + $UpperNat: gLower = $Lower + $LowerNat ' Set working $Upper and $Lower
         END IF                                                   '

         CharSetR = CharSet                                       ' Massage into RegEx string
         REPLACE "\" WITH "\x5C" IN CharSetR                      '
         REPLACE "[" WITH "\x5B" IN CharSetR                      '
         REPLACE "]" WITH "\x5D" IN CharSetR                      '
         cMarkLine          = VAL(sINIGetString("Screen",   "cMarkLine",    "&HC0C098"))
         CmdChr             = sINIGetString("General",      "CmdChr",       ";")
         CMDFlags           = sINIGetString("General",      "CMDFlags",     "/K")
         CmdHiLite          = VAL(sINIGetString("Screen",   "CmdHiLite",    "1"))
         ConfDelFlag        = VAL(sINIGetString("General",  "CDelFlag",     "1"))
         CursInsert         = VAL(sINIGetString("Screen",   "CursIns",      "100"))
         CursNormal         = VAL(sINIGetString("Screen",   "CursNorm",     "20"))
         DefaultShr         = sINIGetString("General",      "DefaultShr",   "")
         DefDataShift       = VAL(sINIGetString("General",  "DefDataShift", "2"))
         DirProfFlag        = VAL(sINIGetString("General",  "DProfFlag",    "0"))
         FindWord           = VAL(sINIGetString("General",  "FindWord",     "0"))
         FMCloseFlag        = VAL(sINIGetString("General",  "FMOpFlag",     "0"))
         FMDateHiLite       = VAL(sINIGetString("General",  "FMDateHiLite", "1"))

         '----- Build FMDefMaster if this is a first time thing
         FMDefMaster        = sINIGetString("FManager",     "FMDefMaster", "FILEPATH(Dir+,Name+) RECENT(Dir*,Name+) FOUND(Dir+,Name+) OPENED(Dir+,Name+) FAVORITES(Dir+,Name+) FLISTS(Dir+,Name+) PATHS(Dir*,Name+) PROFILES(Dir+,Name+)")
         FMFFList           = sINIGetString("FManager",     "FMFFList",     "EXE,DLL")
         IF ISNULL(FMFileList) THEN _                             ' Get FileList if not set by command line
            FMFileList      = sINIGetString("FManager",     "Recall",       "")
         FMHelpFlag         = VAL(sINIGetString("General",  "FMHelpFlag",   "1"))
         FMLayout           = sINIGetString("General",      "FMLayout",     "DATETIME,SIZELONG")
         FMLCmdWidth        = VAL(sINIGetString("General",  "FMLCmdWidth",  "5"))
         IF ISNULL(FMMask) THEN _                                 ' Get Mask if not set by command line
            FMMask          = sINIGetString("FManager",     "DefTypes",     "*")
         FMNoteWidth        = VAL(sINIGetString("General",  "FMNoteWidth",  "10"))
         IF ISNULL(FMPath) THEN _                                 ' Get Path if not set by command line
            FMPath          = sINIGetString("FManager",     "DefDir1",      "C:\")
         FMRecentSort       = sINIGetString("FManager",     "FMRecentSort", "NAMEUP")
         FMScrlAmt          = sINIGetString("FManager",     "ScrlAmtc",     "PAGE")
         FontName           = sINIGetString("Screen",       "FontName",     "Fixedsys")
         FontPitch          = VAL(sINIGetString("Screen",   "FontPitch",    "12"))
         FontStyle          = VAL(sINIGetString("Screen",   "FontStyle",    "0"))
         GlueWith           = sINIGetString("General",      "GlueWith",     "")
         HRuler             = VAL(sINIGetString("Screen",   "HRuler",       "0"))
         InsMode            = VAL(sINIGetString("KeyBoard", "InsMode",      "0"))
         InsReset           = VAL(sINIGetString("KeyBoard", "InsReset",     "1"))
         InvChar            = sINIGetString("General",      "InvChar",      ".")
         JobNumber          = VAL(sINIGetString("General",  "JobNumber",    "0"))
         KBHorScroll        = VAL(sINIGetString("KeyBoard", "KBScrollH",    "0"))
         KBVerScroll        = VAL(sINIGetString("KeyBoard", "KBScrollV",    "0"))
         LastDirOpen        = VAL(sINIGetString("General",  "LastDirOpen",  "2"))
         LastScreenX        = VAL(sINIGetString("Screen",   "LastScrX",     "5"))
         LastScreenY        = VAL(sINIGetString("Screen",   "LastScrY",     "5"))
         i = GetSystemMetrics(%SM_CXVirtualScreen)                ' Get width of virtual screen
         j = GetSystemMetrics(%SM_XVirtualScreen)                 ' Get left hand side of virtual screen
         IF i <> 0 THEN                                           ' Validate a bit
            IF LastScreenX < j OR LastScreenX > i+j THEN LastScreenX = 5 '
         END IF                                                   '
         i = GetSystemMetrics(%SM_CYVirtualScreen)                ' Get height of virtual screen
         j = GetSystemMetrics(%SM_YVirtualScreen)                 ' Get top of virtual screen
         IF i <> 0 THEN                                           ' Validate a bit
            IF LastScreenY < j OR LastScreenY > i+j THEN LastScreenY = 5 '
         END IF                                                   '
         LinNoSize          = VAL(sINIGetString("Screen",   "LNoSize",      "6"))
         LinRepeatMax       = VAL(sINIGetString("General",  "LineRepeatMax", "0"))
         MAXIMIZED          = VAL(sINIGetString("Screen",   "Maximized",    "0"))
         MinRetrieve        = VAL(sINIGetString("General",  "MinRetrieve",  "3"))
         NotifyLevel        = VAL(sINIGetString("General",  "Notify",       "2"))
         NotifyLevelT       = NotifyLevel
         PFKShow            = VAL(sINIGetString("Screen",   "KBHelp",       "0"))
         PrtBandColor       = VAL(sINIGetString("Printer",  "BandingColor", FORMAT$(%RGB_ANTIQUEWHITE)))
         PrtBanding         = VAL(sINIGetString("Printer",  "Banding",      "0"))     '
         PrtBandLines       = VAL(sINIGetString("Printer",  "BandLines",    "0"))  '
         PrtBMargin         = VAL(sINIGetString("Printer",  "BM",           "0.5"))
         PrtDuplex          = VAL(sINIGetString("Printer",  "Duplex",       "1"))
         IF PrtDuplex = 0 THEN PrtDuplex = 1
         PrtFontName        = sINIGetString("Printer",      "Font",         "Courier New")
         PrtFontPitch       = sINIGetString("Printer",      "Pitch",        "10")
         PrtFontStyle       = sINIGetString("Printer",      "Style",        "0")
         PrtFooter          = VAL(sINIGetString("Printer",  "Footers",      "1"))
         PrtFooterCenter    = sINIGetString("Printer",      "FCenter",      " ")
         PrtFooterLeft      = sINIGetString("Printer",      "FLeft",        "File Date: ~d ~t")
         PrtFooterRight     = sINIGetString("Printer",      "FRight",       "Print Date: ~K(ISODate) ~K(ISOTime)")
         PrtHeader          = VAL(sINIGetString("Printer",  "Headers",      "1"))
         PrtHeaderCenter    = sINIGetString("Printer",      "HCenter",      "SPFLite")
         PrtHeaderLeft      = sINIGetString("Printer",      "HLeft",        "~f.~x")
         PrtHeaderRight     = sINIGetString("Printer",      "HRight",       "Page ~#")
         PrtLMargin         = VAL(sINIGetString("Printer",  "LM",           "0.5"))
         PrtMetric          = VAL(sINIGetString("Printer",  "Metric",       "0"))
         PrtName            = sINIGetString("Printer",      "Name",         "DEFAULT")
         PrtOrient          = VAL(sINIGetString("Printer",  "Orient",       "1"))
         PrtPaper           = VAL(sINIGetString("Printer",  "Paper",        "1"))
         PrtRMargin         = VAL(sINIGetString("Printer",  "RM",           "0.5"))
         PrtTMargin         = VAL(sINIGetString("Printer",  "TM",           "0.5"))
         PrtPColor          = VAL(sINIGetString("Printer",  "Color",        FORMAT$(%True)))
         QuickRenum         = VAL(sINIGetString("General",  "QuickRenum",   "50000"))
         RecentCtr          = VAL(sINIGetString("General",  "RecentCtr",    "30"))
         RecentCtr          = MAX(RecentCtr, 2)
         TX                 = sINIGetString("General",      "ROpenLast",    "0")
         IF (TX = "NO" OR TX = "0") THEN ReOpenLast = 0 ELSE ReOpenLast = 1
         ResetU             = VAL(sINIGetString("General",  "ResetU",       "0"))
         RUNFlags           = sINIGetString("General",      "RUNFlags",     "/K")
         SBLayout           = sINIGetString("General",      "SBLayout",     "MODE,LINNO,LINES,COLS,BNDS,INSOVR,CASE,CHANGE,STATE,MISC,SELECT,CAPS,SOURCE,EOL")
         ScrHeight          = VAL(sINIGetString("Screen",   "ScrHeight",    "43"))
         ScrWidth           = VAL(sINIGetString("Screen",   "ScrWidth",     "90"))
         Splash             = VAL(sINIGetString("General",  "Splash",       "0"))
         SubmitCmd          = sINIGetString("General",      "SubmitCmd",    "SPFSUB.BAT ~S")
         SubmitDir          = sINIGetString("General",      "SubmitDir",    JobsPath)
         SubmitInclKey      = sINIGetString("General",      "SubmitInclKey", "#INCLUDE")
         SubmitInclCol      = VAL(sINIGetString("General",  "SubmitInclCol",  "1"))
         UniqueFlag         = VAL(sINIGetString("General",  "UniqFlag",     "0"))
         UseRecycle         = VAL(sINIGetString("General",  "UseRecycle",   "1"))
         VBeepFlag          = VAL(sINIGetString("General",  "VBeepFlag",    "0"))
         VertInsCurs        = VAL(sINIGetString("Screen",   "CursInsVert",  "1"))
         ViewWarn           = VAL(sINIGetString("General",  "BrowseWarn",   "0"))
         VRuler             = VAL(sINIGetString("Screen",   "VRuler",       "0"))

         '----- If we did a color conversion, save all the schemes
         IF ConvSave THEN me.SchemeSaveAll                        ' Save the new schemes right away
      END METHOD

      METHOD INIColorCnvt()
      '---------- Migrate old colors to new Schemes
      LOCAL FG, BG1, BG2, i, j, dNum, ClrOpTyp, ClrOpVal AS LONG, t, tx, dListTxt, ClrOpStr AS STRING
      DIM dList(1 TO 100) AS STRING
      LOCAL ClrIO AS iIO                                          ' For our I/O stuff
      DIM parsed(100) AS STRING
      DIM Numbers(1 TO 10) AS STRING


         '----- Set all to basic text
         FG  = VAL(sINIGetString("Screen",   "cTxtLo",       "&HFF6464"))
         BG1 = VAL(sINIGetString("Screen",   "cBGround",     "&HFFFFFF"))
         BG2 = VAL(sINIGetString("Screen",   "cBGround2",    "&HD9FFD9"))
         FOR i = 0 TO 31                                          ' Default all to this color combo
            sIniSetString("Scheme", "Scheme" + FORMAT$(i, "00") + "FG", FORMAT$(FG))
            sIniSetString("Scheme", "Scheme" + FORMAT$(i, "00") + "BG1", FORMAT$(BG1))
            sIniSetString("Scheme", "Scheme" + FORMAT$(i, "00") + "BG2", FORMAT$(BG2))
         NEXT i

         '----- Overlay now with the unique values from the old INI file
         i  = VAL(sINIGetString("Screen",   "cTxtHi",       "&H000000"))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCTxtHi,  "00") + "FG",  FORMAT$(i))

         i = VAL(sINIGetString("Screen",   "cATabBG",      "&H00FFFF"))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCATabMod,  "00") + "BG1", FORMAT$(i))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCATabMod,  "00") + "BG2", FORMAT$(i))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCATabNMod, "00") + "BG1", FORMAT$(i))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCATabNMod, "00") + "BG2", FORMAT$(i))

         i = VAL(sINIGetString("Screen",   "cBlueBG",      "&HFFB0B0"))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCHiBlue,  "00") + "BG1", FORMAT$(i))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCHiBlue,  "00") + "BG2", FORMAT$(i))

         i = VAL(sINIGetString("Screen",   "cBlueFG",      "&HFFFFFF"))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCHiBlue,  "00") + "FG",  FORMAT$(i))

         i = VAL(sINIGetString("Screen",   "cErrMsg",      "&H0000FF"))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCError,  "00") + "FG",  FORMAT$(i))

         i = VAL(sINIGetString("Screen",   "cGreenBG",     "&H008000"))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCHiGreen,  "00") + "BG1", FORMAT$(i))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCHiGreen,  "00") + "BG2", FORMAT$(i))

         i = VAL(sINIGetString("Screen",   "cGreenFG",     "&HFFFFFF"))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCHiGreen,  "00") + "FG",  FORMAT$(i))

         i = VAL(sINIGetString("Screen",   "cITabBG",      "&HEFEFEF"))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCITabMod,  "00") + "BG1", FORMAT$(i))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCITabMod,  "00") + "BG2", FORMAT$(i))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCITabNMod, "00") + "BG1", FORMAT$(i))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCITabNMod, "00") + "BG2", FORMAT$(i))

         i = VAL(sINIGetString("Screen",   "cLinNumHi",    "&HA00000"))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCLnoHi, "00") + "FG", FORMAT$(i))

         i = VAL(sINIGetString("Screen",   "cLinNumLo",    "&HB5B5B5"))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCLnoLo, "00") + "FG", FORMAT$(i))

         i = VAL(sINIGetString("Screen",   "cPFKBG",       "&HE1DFCE"))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCPFK, "00") + "BG1", FORMAT$(i))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCPFK, "00") + "BG2", FORMAT$(i))

         i = VAL(sINIGetString("Screen",   "cPFKFG",       "&H757575"))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCPFK, "00") + "FG", FORMAT$(i))

         i = VAL(sINIGetString("Screen",   "cRedBG",       "&H0000FF"))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCHiRed,  "00") + "BG1", FORMAT$(i))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCHiRed,  "00") + "BG2", FORMAT$(i))

         i = VAL(sINIGetString("Screen",   "cRedFG",      "&HFFFFFF"))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCHiRed,  "00") + "FG",  FORMAT$(i))

         i = VAL(sINIGetString("Screen",   "cStatBG",     "&HE2E2E2"))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCStatus, "00") + "BG1", FORMAT$(i))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCStatus, "00") + "BG2", FORMAT$(i))

         i = VAL(sINIGetString("Screen",   "cStatFG",     "&H000000"))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCStatus, "00") + "FG", FORMAT$(i))

         i = VAL(sINIGetString("Screen",   "cTabMod",     "&H0000FF"))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCATabMod, "00") + "FG", FORMAT$(i))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCITabMod, "00") + "FG", FORMAT$(i))

         i = VAL(sINIGetString("Screen",   "cTabNMod",     "&HFF46A3"))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCATabNMod, "00") + "FG", FORMAT$(i))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCITabNMod, "00") + "FG", FORMAT$(i))

         i = VAL(sINIGetString("Screen",   "cToolBG",      "&HA39D6B"))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCFMTool, "00") + "BG1", FORMAT$(i))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCFMTool, "00") + "BG2", FORMAT$(i))

         i = VAL(sINIGetString("Screen",   "cToolFG",      "&HFFFFFF"))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCFMTool, "00") + "FG", FORMAT$(i))

         i = VAL(sINIGetString("Screen",   "cTxtLo",       "&H000000"))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCTxtLo, "00") + "FG", FORMAT$(i))

         i = VAL(sINIGetString("Screen",   "cYellowBG",    "&H00FFFF"))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCHiYellow,  "00") + "BG1", FORMAT$(i))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCHiYellow,  "00") + "BG2", FORMAT$(i))

         i = VAL(sINIGetString("Screen",   "cYellowFG",    "&H000000"))
         sIniSetString("Scheme", "Scheme" + FORMAT$(%SCHiYellow,  "00") + "FG",  FORMAT$(i))

         '----- Now set some colorize Schemes based on a selected AUTO file
         t = DlgV10                                               ' Go get user's answer
         IF t = "" THEN EXIT METHOD                               ' Leave if SKIP/CANCEL selected

         '----- Now read the old scheme values
         LET ClrIO = CLASS "cIO"                                  '
         ClrIO.Setup("IE", "", "", ENV.AutoPath + t)              ' Tell IO what we're opening
         ClrIO.EXEC                                               '
         DO WHILE ISFALSE EOF(ClrIO.FNum)                         ' Read the data
            LINE INPUT # ClrIO.FNum, tx                           ' Get a line
            IF LEN(TRIM$(tx)) > 0 AND LEFT$(tx, 1) <> ";" THEN    ' Process the statement
               PARSE tx, parsed(), " "                            '
               '----- Since we use ' ' delimiters, eliminate multi-blank spans
               i = 0: j = 0                                       ' Shrink unused entries
               DO WHILE j <= 100                                  '
                  IF ISNULL(parsed(i)) THEN ARRAY DELETE parsed(i) ELSE INCR i
                  INCR j                                          '
               LOOP                                               '
               SELECT CASE AS CONST$ UUCASE(parsed(0))            ' Do different types

                  '----- SCHEME command
                  CASE "SCHEME"                                   ' SCHEME
                     IF parsed(3) = "" THEN ITERATE DO            ' Not Enough params? Ignore it
                     j = VAL(parsed(1))                           ' j = color index
                     IF j > 99 THEN                               ' If a 100+ value
                        j -= 100                                  ' Reduce to the normal value
                     END IF                                       '
                     IF j = 0 THEN ITERATE DO                     ' Ignore SCHEME 0
                     IF j > 14 THEN ITERATE DO                    ' Ignore if not within range
                     INCR j                                       ' Make j into the new Scheme number
                     ClrOpStr = parsed(2): GOSUB DoClrOp          '
                     IF ClrOpTyp  THEN                            ' A BGR value
                        FG = ClrOpVal                             ' Pick it up
                     ELSE                                         ' else
                        FG = sQBColor(ClrOpVal)                   ' Get QBColor version
                     END IF                                       '

                     ClrOpStr = parsed(3): GOSUB DoClrOp          '
                     IF ClrOpTyp THEN                             ' A BGR value
                        BG1 =  ClrOpVal                           ' Pick it up
                        BG2 = ClrOpVal                            ' Use as Alt. for now
                     ELSE                                         ' else
                        BG1 = sQBColor(ClrOpVal)                  ' Get QBColor version
                        BG2 = sQBColor(ClrOpVal)                  ' Use as Alt. for now
                     END IF                                       '
                     IF parsed(4) <> "" THEN                      ' Provided a 4th operand?
                        ClrOpStr = parsed(4): GOSUB DoClrOp       '
                        IF ClrOpTyp  THEN                         ' A BGR  value
                           BG2 = ClrOpVal                         ' Use as Alt
                        ELSE                                      ' else
                           BG2 = sQBColor(ClrOpVal)               ' Use as Alt.
                        END IF                                    '
                     END IF                                       '
                     sIniSetString("Scheme", "Scheme" + FORMAT$(j, "00") + "FG", FORMAT$(FG))
                     sIniSetString("Scheme", "Scheme" + FORMAT$(j, "00") + "BG1", FORMAT$(BG1))
                     sIniSetString("Scheme", "Scheme" + FORMAT$(j, "00") + "BG2", FORMAT$(BG2))

               END SELECT                                         '
            END IF                                                '
         LOOP                                                     '
         ClrIO.Close                                              ' Close file
         EXIT METHOD

      '---- Process a Color operand
      DoClrOp:
         ClrOpStr = UUCASE(ClrOpStr)                              ' Prevent Case problems
         IF LEFT$(ClrOpStr, 2) = "&H" THEN                        ' Old style &Hbbggrr value
            ClrOpVal = VAL(ClrOpStr)                              ' Setup answer area
            ClrOpTyp = %True                                      ' Say it's a BGR value color value

         ELSEIF LEFT$(ClrOpStr, 2) = "BX" THEN                    ' Bxbbggrr value
            MID$(ClrOpStr, 1, 2) = "&H"                           ' Make it look like &hbbggrr
            ClrOpVal = VAL(ClrOpStr)                              ' Setup answer area
            ClrOpTyp = %True                                      ' Say it's a BGR value color value

         ELSEIF LEFT$(ClrOpStr, 2) = "RX" THEN                    ' Rxrrggbb value
            MID$(ClrOpStr, 1, 2) = "&H"                           ' Make it look like &hrrggbb
            t = RIGHT$(ClrOpStr, 2)                               ' Save the bb
            MID$(ClrOpStr, 7, 2) = MID$(ClrOpStr, 3, 2)           ' Move the rr to the old bb position
            MID$(ClrOpStr, 3, 2) = t                              ' Stuff the bb into the old rr position
            ClrOpVal = VAL(ClrOpStr)                              ' Setup answer area now it's in &hbbggrr format
            ClrOpTyp = %True                                      ' Say it's a BGR value color value

         ELSEIF LEFT$(ClrOpStr, 2) = "B=" THEN                    ' B=bbb,ggg,rrr
            ClrOpStr = MID$(clrOpStr, 3)                          ' Strip off the B=
            PARSE ClrOpStr, numbers(), ","                        ' Parse out the nnn,nnn,nnn
            ClrOpStr = "&H"                                       ' Start building a new operand
            ClrOpStr += HEX$(VAL(numbers(1)), 2)                  ' Add the bb in hex
            ClrOpStr += HEX$(VAL(numbers(2)), 2)                  ' Add the gg in hex
            ClrOpStr += HEX$(VAL(numbers(3)), 2)                  ' Add the rr in hex
            ClrOpVal = VAL(ClrOpStr)                              ' Setup answer area now it's in &hbbggrr format
            ClrOpTyp = %True                                      ' Say it's a BGR value color value

         ELSEIF LEFT$(ClrOpStr, 2) = "R=" THEN                    ' R=rrr,ggg,bbb
            ClrOpStr = MID$(clrOpStr, 3)                          ' Strip off the R=
            PARSE ClrOpStr, numbers(), ","                        ' Parse out the nnn,nnn,nnn
            ClrOpStr = "&H"                                       ' Start building a new operand
            ClrOpStr += HEX$(VAL(numbers(3)), 2)                  ' Add the bb in hex
            ClrOpStr += HEX$(VAL(numbers(2)), 2)                  ' Add the gg in hex
            ClrOpStr += HEX$(VAL(numbers(1)), 2)                  ' Add the rr in hex
            ClrOpVal = VAL(ClrOpStr)                              ' Setup answer area now it's in &hbbggrr format
            ClrOpTyp = %True                                      ' Say it's a BGR value color value

         ELSEIF VAL(ClrOpStr) < 17 THEN                           ' A QB color number
            ClrOpVal = VAL(ClrOpStr)                              ' Setup answer area
            ClrOpTyp = %False                                     ' Say it's a QB color value
         END IF                                                   '
         RETURN                                                   '

      END METHOD

      METHOD SetINITimeStamp()
      '---------- Set a last update in the INI file
      LOCAL tx, Stamp AS STRING
         '----- Set the INI timestamp.
         TX = DATE$                                               ' Get date for timestamp
         Stamp = RIGHT$(TX, 2) + LEFT$(TX, 2) + MID$(TX, 4, 2) + "@" + TIME$
         Stamp = Stamp + "." + MID$($Numeric, RND(1, 10), 1) + MID$($Numeric, RND(1, 10), 1)
         sIniSetString("General", "$LastProfUpdate", Stamp)       ' Write it
      END METHOD

      METHOD DispWelcome()
      '---------- Display the Welcome dialog
      LOCAL rText AS STRING
         MEntry
         '----- Build the dialog
         DIALOG FONT DEFAULT "Tahoma", 11 / gFontScale, 0, 0
         DIALOG NEW 0, "SPFLite Welcome",,, 250, 140, %WS_THICKFRAME OR %WS_CAPTION TO hWel
         DIALOG SET COLOR hWel, %RGB_GAINSBORO, %RGB_GAINSBORO
         CONTROL ADD LABEL,  hWel, %WELCOME_TEXT+0, "Welcome to SPFLite Version " & ENV.PgmVers, 2, 5, 200, 14
         CONTROL SET COLOR   hWel, %WELCOME_TEXT+0, %BLUE, %RGB_GAINSBORO
         CONTROL ADD LABEL,  hWel, %WELCOME_TEXT+1, "This appears to be your first execution of SPFLite, so you will be", 2, 20, 247, 12
         CONTROL SET COLOR   hWel, %WELCOME_TEXT+1, %BLUE, %RGB_GAINSBORO
         CONTROL ADD LABEL,  hWel, %WELCOME_TEXT+2, "started in the Options dialog to perform your initial Options setup.", 2, 30, 247, 12
         CONTROL SET COLOR   hWel, %WELCOME_TEXT+2, %BLUE, %RGB_GAINSBORO
         CONTROL ADD LABEL,  hWel, %WELCOME_TEXT+4, "Choose one of the following settings for ENTER and NEWLINE", 2, 50, 300, 12
         CONTROL SET COLOR   hWel, %WELCOME_TEXT+4, %BLUE, %RGB_GAINSBORO
         CONTROL ADD BUTTON, hWel, %WELCOME_OPTION1, "(A)", 2, 65, 20, 13, %WS_BORDER
         CONTROL ADD LABEL,  hWel, %WELCOME_TEXT+6, "Enter key is ENTER,  Backpace key is NEWLINE", 30, 67, 200, 12
         CONTROL SET COLOR   hWel, %WELCOME_TEXT+6, %BLUE, %RGB_GAINSBORO
         CONTROL ADD BUTTON, hWel, %WELCOME_OPTION2, "(B)", 2, 85, 20, 13, %WS_BORDER
         CONTROL ADD LABEL,  hWel, %WELCOME_TEXT+7, "Enter key is ENTER,  Ctrl-Enter key is NEWLINE", 30, 87, 200, 12
         CONTROL SET COLOR   hWel, %WELCOME_TEXT+7, %BLUE, %RGB_GAINSBORO
         CONTROL ADD BUTTON, hWel, %WELCOME_OPTION3, "(C)", 2, 105, 20, 13, %WS_BORDER
         CONTROL ADD LABEL,  hWel, %WELCOME_TEXT+8, "Right Ctrl key is ENTER,  Enter key is NEWLINE", 30, 107, 200, 12
         CONTROL SET COLOR   hWel, %WELCOME_TEXT+8, %BLUE, %RGB_GAINSBORO
         CONTROL ADD LABEL,  hWel, %WELCOME_TEXT+5, "(You may alter these later via the KEYMAP command)", 2, 125, 300, 12
         CONTROL SET COLOR   hWel, %WELCOME_TEXT+5, %BLUE, %RGB_GAINSBORO
         DIALOG SHOW MODAL hWel CALL DlgWelcomeCallback                 ' Display it all
         MExit
      END METHOD

      METHOD DlgWelcomeCallBack(BYVAL hDlg AS DWORD,  _           ' CB.HNDL
                               BYVAL wMsg AS LONG,   _            ' CB.MSG
                               BYVAL WPARAM AS LONG, _            ' CB.WPARAM    CB.CTLMSG = HI(WORD, wParam) CB.CTL = LO(WORD, wParam)
                               BYVAL LPARAM AS LONG) AS LONG      ' CB.LPARAM
      '--------------------
      ' Callback function used by the Welcome Dialog
      '--------------------

         SELECT CASE AS LONG wMsg                                 '

            '----- SYSCOMMAND
            CASE %WM_SYSCOMMAND
               IF (WPARAM AND &HFFF0) = %SC_CLOSE THEN            ' Trap the [x] button and Alt-F4
                  METHOD = 1                                      ' Ignore it
               END IF                                             '

            '----- Handle the various interactions
            CASE %WM_COMMAND                                      '
               SELECT CASE AS LONG LO(WORD, WPARAM)               '
                  CASE %WELCOME_OPTION1                           '
                     IF HI(WORD, WPARAM) = %BN_CLICKED THEN       '
                        ENV.WelcomeOpt = 1                        '
                        RESET hToolTips                           ' Clear so Tooltips get re-created correctly
                        DIALOG END hDlg                           ' Shut things down
                     END IF                                       '

                  CASE %WELCOME_OPTION2                           '
                     IF HI(WORD, WPARAM) = %BN_CLICKED THEN       '
                        ENV.WelcomeOpt = 2                        '
                        RESET hToolTips                           ' Clear so Tooltips get re-created correctly
                        DIALOG END hDlg                           ' Shut things down
                     END IF                                       '

                  CASE %WELCOME_OPTION3                           '
                     IF HI(WORD, WPARAM) = %BN_CLICKED THEN       '
                        ENV.WelcomeOpt = 3                        '
                        RESET hToolTips                           ' Clear so Tooltips get re-created correctly
                        DIALOG END hDlg                           ' Shut things down
                     END IF                                       '
               END SELECT                                         '
         END SELECT                                               '
      END METHOD

   END INTERFACE

END CLASS

   CALLBACK FUNCTION DlgWelcomeCallBack()
      FUNCTION = Env.DlgWelcomeCallBack(CB.HNDL, CB.MSG, CB.WPARAM, CB.LPARAM)
   END FUNCTION
