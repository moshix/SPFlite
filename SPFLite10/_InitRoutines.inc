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


SUB      InitFMLayout()
'---------- Setup FM screen layout parameters
LOCAL lcmd, lext, lsize, ldate, ltime, lnote, nextpos, misc_lengths, i, j AS LONG
LOCAL Layout, lField AS STRING
LOCAL opers() AS STRING
DIM opers(1 TO 3) AS STRING
'----- Valid FMLayout KWs are: EXT, SIZESHORT, SIZELONG, DATE, DATETIME, LINES, NOTE
'_ New  _ FilePath  _ Recent  _ Found  _ Opened  _ Favorites  _ Flists  _ Paths  _ Profiles
'         1         2         3         4         5         6         7         8
'12345678901234567890123456789012345678901234567890123456789012345678901234567890
   Layout = ENV.FMLayout                                          ' Get working copy of layout
   FM_Quick_Line_1     = 3                                        ' Set top of screen area
   FM_Path_Line        = 4                                        '
   FM_Path_Left        = 17                                       '
   FM_Mask_Line        = 5                                        '
   FM_Mask_Left        = 17                                       '
   FM_Head_Line        = 6                                        '
   FM_Top_File_Line    = 7                                        '
   FM_Quick_Pos_1      = 1                                        '
   FM_Quick_Pos_2      = 8                                        '
   FM_Quick_Pos_3      = 20                                       '
   FM_Quick_Pos_4      = 30                                       '
   FM_Quick_Pos_5      = 39                                       '
   FM_Quick_Pos_6      = 49                                       '
   FM_Quick_Pos_7      = 62                                       '
   FM_Quick_Pos_8      = 72                                       '
   FM_Quick_Pos_9      = 81                                       '
   FM_Head_Name_Left   = ENV.FMLCmdWidth + 2                      '
   FM_Crit_Size =  ENV.ScrWidth - FM_Path_Left                    ' Calc available criteria size
   FM_Note_Size =  ENV.FMNoteWidth                                '

   '----- Calc length of Misc fields
   FM_Misc_Length = 0                                             ' Reset it
   DO WHILE ISNOTNULL(TRIM$(Layout))                              ' Step through the layout
      lField = EXTRACT$(Layout, ",")                              ' Get next operand
      Layout = REMAIN$(Layout, ",")                               ' Remove it from string
      SELECT CASE AS CONST$ lField                                ' See what field
         CASE "EXT":       FM_Misc_Length += %FMLExt + 2          ' EXT + pad
         CASE "SIZESHORT": FM_Misc_Length += %FMLSizeS + 2        ' Short size + pad
         CASE "SIZELONG":  FM_Misc_Length += %FMLSizeL + 2        ' Long size + pad
         CASE "DATE":      FM_Misc_Length += %FMLDate + 2         ' Date size + pad
         CASE "DATETIME":  FM_Misc_Length += %FMLDateT + 2        ' Date/Time size + pad
         CASE "LINES":     FM_Misc_Length += %FMLLines + 2        ' Lines size + pad
         CASE "NOTE":      FM_Misc_Length += FM_Note_Size + 2     ' Note size + pad
      END SELECT                                                  '
   LOOP                                                           '
   FM_Misc_Length -= 2                                            ' Don't need pad after last field

   '----- Now go back and adjust filename to use remaining space
   FM_File_Size =  ENV.ScrWidth - ENV.FMLCmdWidth - FM_Misc_Length - 3' Width - LCmd - Misc fields - 3
   FM_List_Height = ENV.ScrHeight - FM_Top_File_Line - (3 * ENV.FMHelpFlag) + 1 ' Calc size of scrolling area
END SUB

SUB      InitLNText
'---------- Init the Line Number text constants
LOCAL i, fnum AS LONG, txt1, txt2, fn AS STRING
   MEntry
   gLNPadCol = ENV.LinNoSize + 1                                  ' Pad column number
   gLNData1  = ENV.LinNoSize + 2                                  ' Data column 1 location
   ARRAY ASSIGN gLnoTextType() = %Top, %Bottom, %Tabs, %Bounds, %Cols, %Xclude, %InsertLine, %EQChange, %Word, %Mark, %Mask, %Prof, %File, %Note, %Page
   gLnoTextTxt(1) = REPEAT$(8, "*")                               ' Setup constants
   gLnoTextTxt(2) = REPEAT$(8, "*")                               '
   gLnoTextTxt(6) = REPEAT$(8, "-")                               '
   gLnoTextTxt(7) = REPEAT$(8, "'")                               '
   IF ENV.LinNoSize = 5 THEN                                      ' Set constants differently
      gLnoTextTxt(3)  = "TABS>   "                                '
      gLnoTextTxt(4)  = "BNDS>   "                                '
      gLnoTextTxt(5)  = "COLS>   "                                '
      gLnoTextTxt(8)  = "=CHG>   "                                '
      gLnoTextTxt(9)  = "WORD>   "                                '
      gLnoTextTxt(10) = "MARK>   "                                '
      gLnoTextTxt(11) = "MASK>   "                                '
      gLnoTextTxt(12) = "PROF>   "                                '
      gLnoTextTxt(13) = "FILE>   "                                '
      gLnoTextTxt(14) = "=##=>   "                                '
      gLnoTextTxt(15) = "PAGE>   "                                '
   ELSEIF ENV.LinNoSize = 6 THEN                                  '
      gLnoTextTxt(3)  = "=TABS>  "                                '
      gLnoTextTxt(4)  = "=BNDS>  "                                '
      gLnoTextTxt(5)  = "=COLS>  "                                '
      gLnoTextTxt(8)  = "==CHG>  "                                '
      gLnoTextTxt(9)  = "=WORD>  "                                '
      gLnoTextTxt(10) = "=MARK>  "                                '
      gLnoTextTxt(11) = "=MASK>  "                                '
      gLnoTextTxt(12) = "=PROF>  "                                '
      gLnoTextTxt(13) = "=FILE>  "                                '
      gLnoTextTxt(14) = "=NOTE>  "                                '
      gLnoTextTxt(15) = "=PAGE>  "                                '
   ELSEIF ENV.LinNoSize = 7 THEN                                  '
      gLnoTextTxt(3)  = "==TABS> "                                '
      gLnoTextTxt(4)  = "==BNDS> "                                '
      gLnoTextTxt(5)  = "==COLS> "                                '
      gLnoTextTxt(8)  = "===CHG> "                                '
      gLnoTextTxt(9)  = "==WORD> "                                '
      gLnoTextTxt(10) = "==MARK> "                                '
      gLnoTextTxt(11) = "==MASK> "                                '
      gLnoTextTxt(12) = "==PROF> "                                '
      gLnoTextTxt(13) = "==FILE> "                                '
      gLnoTextTxt(14) = "==NOTE> "                                '
      gLnoTextTxt(15) = "==PAGE> "                                '
   ELSEIF ENV.LinNoSize = 8 THEN                                  '
      gLnoTextTxt(3)  = "===TABS>"                                '
      gLnoTextTxt(4)  = "===BNDS>"                                '
      gLnoTextTxt(5)  = "===COLS>"                                '
      gLnoTextTxt(8)  = "====CHG>"                                '
      gLnoTextTxt(9)  = "===WORD>"                                '
      gLnoTextTxt(10) = "===MARK>"                                '
      gLnoTextTxt(11) = "===MASK>"                                '
      gLnoTextTxt(12) = "===PROF>"                                '
      gLnoTextTxt(13) = "===FILE>"                                '
      gLnoTextTxt(14) = "===NOTE>"                                '
      gLnoTextTxt(15) = "===PAGE>"                                '
   END IF                                                         '
   MExit                                                          '
END SUB

SUB      InitLocalTables()
'---------- Setup some misc. tables
LOCAL i AS LONG
LOCAL Retr, t AS STRING
LOCAL SetVar() AS STRING, SetVarCtr AS LONG, SetKey AS STRING
LOCAL SIO AS iIO                                                  ' For our I/O stuff
LOCAL LTime AS IPOWERTIME                                         ' Create a PowerTime object
   MEntry
   LET LTime = CLASS "PowerTime"                                  '
   LET SIO = CLASS "cIO"                                          '
   '---------- Setup the Retr Table
   sRetrLoad                                                      ' Go load the Retrieve stack
   gCmdRtrevIX = 1                                                '

   '----- Set up some other stuff
   gEnumWith = 1                                                  ' Global ENUM increment
                                                                  '

   '----- Setup gDateActive and gDateRecent
   LTime.Now                                                      ' Assign the current Date/Time to LTime
   gDateActive = FORMAT$(LTime.year(), "0000") + "-" + FORMAT$(LTime.Month(), "00") + "-" + FORMAT$(LTime.day(), "00") + "  " + _
            FORMAT$(LTime.Hour(), "00") + ":" + FORMAT$(LTime.Minute(), "00")
   LTime.AddHours(-1)                                             ' Minus 1 hour
   gDateActive1 = FORMAT$(LTime.year(), "0000") + "-" + FORMAT$(LTime.Month(), "00") + "-" + FORMAT$(LTime.day(), "00") + "  " + _
            FORMAT$(LTime.Hour(), "00") + ":" + FORMAT$(LTime.Minute(), "00")
   LTime.AddHours(-7)                                             ' Minus 7 more hours to make -8
   gDateActive8 = FORMAT$(LTime.year(), "0000") + "-" + FORMAT$(LTime.Month(), "00") + "-" + FORMAT$(LTime.day(), "00") + "  " + _
            FORMAT$(LTime.Hour(), "00") + ":" + FORMAT$(LTime.Minute(), "00")
   LTime.AddHours(-16)                                            ' Minus 16 more hours to make -24
   gDateActive24 = FORMAT$(LTime.year(), "0000") + "-" + FORMAT$(LTime.Month(), "00") + "-" + FORMAT$(LTime.day(), "00") + "  " + _
            FORMAT$(LTime.Hour(), "00") + ":" + FORMAT$(LTime.Minute(), "00")
   LTime.AddHours(-24)                                            ' Minus 24 more hours to make -48
   gDateActive48 = FORMAT$(LTime.year(), "0000") + "-" + FORMAT$(LTime.Month(), "00") + "-" + FORMAT$(LTime.day(), "00") + "  " + _
            FORMAT$(LTime.Hour(), "00") + ":" + FORMAT$(LTime.Minute(), "00")



   '----- Setup the SET Table
   DIM gSetKey(1 TO 100) AS STRING: DIM gSetData(1 TO 100) AS STRING: gSetCount = 0
   RESET SetVarCtr
   REDIM SetVar(1 TO 1) AS STRING                                 '
   IF ISFILE(ENV.INIPath + "SPFLite.SPS") THEN                    ' If it exists
      SIO.Setup("IE", "", "", ENV.INIPath + "SPFLite.SPS")        ' Tell IO what we're opening
      SIO.EXEC                                                    '
      DO WHILE ISFALSE EOF(SIO.FNum)                              ' Read the file
         LINE INPUT # SIO.FNum, t                                 ' Get a line
         i = INSTR(t, "=")                                        ' Locate =
         IF i THEN                                                ' Got one
            IF i <> 1 THEN                                        ' A normal key=value style
               IF SetVarCtr > 0 THEN GOSUB DumpSet                ' An item in hold
               SetKey = LEFT$(t, i - 1)                           ' Save the key
               INCR SetVarCtr                                     ' Bump var count
               IF SetVarCtr > UBOUND(SetVar()) THEN _             ' Time to expand?
                  REDIM PRESERVE SetVar(1 TO SetVarCtr) AS STRING '
               SetVar(SetVarCtr) = MID$(t, i + 1)                 ' Save in work array

            ELSE                                                  ' An =value stype
               INCR SetVarCtr                                     ' Bump var count
               IF SetVarCtr > UBOUND(SetVar()) THEN               ' Time to expand?
                  REDIM PRESERVE SetVar(1 TO SetVarCtr) AS STRING '
               END IF
               SetVar(SetVarCtr) = MID$(t, 2)                     ' Save in work array

            END IF                                                '
         END IF                                                   '
      LOOP                                                        '
      SIO.Close                                                   ' Close it
      IF SetVarCtr > 0 THEN GOSUB DumpSet                         ' An item in hold, dump it
   END IF                                                         '

   '----- Next setup the Status Bar Table
   '  '     Name       ID  Align PosBar Width  Def.Scheme  OV.Scheme  Text
   sSBInit("MODE    ", "A", "C",  "N",     80,  %SCStatus,        0,  "")
   sSBInit("LINNO   ", "B", "L",  "N",    175,  %SCStatus,        0,  "")
   sSBInit("LINES   ", "C", "C",  "N",    105,  %SCStatus,        0,  "")
   sSBInit("COLS    ", "D", "C",  "N",    135,  %SCStatus,        0,  "")
   sSBInit("BNDS    ", "E", "C",  "N",    135,  %SCStatus,        0,  "")
   sSBInit("INSOVR  ", "F", "C",  "N",     35,  %SCStatus,        0,  "")
   sSBInit("CASE    ", "G", "C",  "N",     35,  %SCStatus,        0,  "")
   sSBInit("CHANGE  ", "H", "C",  "N",     35,  %SCStatus,        0,  "")
   sSBInit("STATE   ", "I", "C",  "N",     35,  %SCStatus,        0,  "")
   sSBInit("MISC    ", "J", "C",  "Y",    125,  %SCStatus,        0,  "")
   sSBInit("SELECT  ", "K", "C",  "N",    150,  %SCStatus,        0,  "")
   sSBInit("CAPS    ", "L", "C",  "N",     90,  %SCStatus,        0,  "")
   sSBInit("SOURCE  ", "M", "C",  "N",     90,  %SCStatus,        0,  "")
   sSBInit("EOL     ", "N", "C",  "N",     90,  %SCStatus,        0,  "")
   sSBInit("PAD     ", "O", "C",  "N",   9999,  %SCStatus,        0,  "")
   MExitSub

DumpSet:
   INCR gSetCount                                                 ' Count entry
   IF gSetCount > UBOUND(gSetKey()) THEN                          ' Need a bigger table?
      REDIM PRESERVE gSetKey(1 TO UBOUND(gSetKey) * 2)            ' Yes, expand them
      REDIM PRESERVE gSetData(1 TO UBOUND(gSetData) * 2)          '
   END IF                                                         '
   gSetKey(gSetCount) = SetKey                                    ' Set key
   gSetData(gSetCount) = JOIN$(SetVar(), BINARY)                  ' Set Data = to SetVar array
   RESET SetVarCtr                                                '
   REDIM SetVar(1 TO 1) AS STRING                                 '
   RETURN
END SUB

SUB      sSBInit(SB1 AS STRING, SB2 AS STRING, SB3 AS STRING, SB4 AS STRING, SB5 AS LONG, SB6 AS LONG, SB7 AS LONG, SB8 AS STRING)
'---------- Add a single SB item to the table
STATIC SBI AS LONG
   INCR SBI                                                       ' Bump index
   gSBTable(SBI).SBName     = SB1                                 ' Name
   gSBTable(SBI).SBID       = SB2                                 ' ID
   gSBTable(SBI).SBActive   = " "                                 ' Clear Active flag
   gSBTable(SBI).SBAlign    = SB3                                 ' Alignment
   gSBTable(SBI).SBPosBar   = SB4                                 ' Position Bar box
   gSBTable(SBI).SBMySelf   = FORMAT$(SBI, "00")                  ' Myself in text form
   gSBTable(SBI).SBWidth    = SB5                                 ' Width
   gSBTable(SBI).SBDfScheme = SB6                                 ' Default Scheme
   gSBTable(SBI).SBOvScheme = SB7                                 ' Override Scheme
   gSBTable(SBI).SBText     = SB8                                 ' Message TextD

END SUB

SUB      InitASynchStuff
'---------- Initialize the ASynch startup things
STATIC TDummy AS LONG
STATIC PDummy AS LONG POINTER
LOCAL  TThread, tResult AS LONG
   MEntry
   pDummy = VARPTR(TDummy)                                        ' Setup a dummy pointer
   THREAD CREATE InitASynchStuffThread(pDummy) 65536, TO TThread  ' Fire up the thread
   IF TThread = 0 THEN MExitSub                                   ' Failed?
   THREAD STATUS TThread TO tResult                               ' See if running OK
   IF tResult <> 259 THEN                                         ' If running OK STATUS returns &H103 (See Help)
      THREAD CLOSE TThread TO tResult                             ' Free up our handle
   END IF                                                         '
   MExit                                                          ' Done
END SUB

THREAD FUNCTION InitASynchStuffThread(BYVAL pData AS LONG) AS LONG' Do the low priority startup stuff
REGISTER i AS LONG
REGISTER j AS LONG

LOCAL fn, kfn, txt1 AS STRING, FD AS DIRDATA, Expiry AS LONG
   '----- If only one instance running, do UNDO file cleanup

   IF ISFALSE sFindWindow("SPFLite(v" + LEFT$(ENV.PgmVers, 4)) THEN  ' If we're not running somewhere else
      fn = sGetWindowsTempDir()                                   ' Get the temp directory
      fn += "UND*.TMP"                                            ' Make a generic name
      IF DIR$(fn) <> "" THEN                                      ' Anything to delete?
         TRY                                                      '
            KILL fn                                               ' Wipe out any left over UNDO files
            EXIT TRY                                              '
         CATCH                                                    '
            EXIT TRY                                              '
         END TRY                                                  '
      END IF                                                      '
   END IF                                                         '

   '----- Make sure some directories exist, create them if not
   IF ISFALSE ISFOLDER(ENV.ClipPath) THEN MKDIR ENV.ClipPath      ' Ensure \CLIP exists
   IF ISFALSE ISFOLDER(ENV.AutoPath) THEN MKDIR ENV.AutoPath      ' Ensure \AUTO exists
   IF ISFALSE ISFOLDER(ENV.FileListPath) THEN MKDIR ENV.FileListPath  ' Ensure \FILELIST exists
   IF ISFALSE ISFOLDER(ENV.RunPath) THEN MKDIR ENV.RunPath        ' Ensure \RUN exists
   IF ISFALSE ISFOLDER(ENV.MacrosPath) THEN MKDIR ENV.MacrosPath  ' Ensure \MACROS exists
   IF ISFALSE ISFOLDER(ENV.JobsPath) THEN MKDIR ENV.JobsPath      ' Ensure \JOBS exists
   IF ISFALSE ISFOLDER(ENV.VSavePath) THEN MKDIR ENV.VSavePath    ' Ensure \VSAVE exists
   IF ISFALSE ISFOLDER(ENV.StatePath) THEN MKDIR ENV.StatePath    ' Ensure \STATE exists
   IF ISFALSE ISFOLDER(ENV.ProfPath) THEN MKDIR ENV.ProfPath      ' Ensure \PROFILE exists

   Expiry = 1                                                     ' Do the 1 day expiry guys
   '----- Clean up any old SUBMIT files
   fn = ENV.SubmitDir + "\SUB*.*"                                 ' Get SUBMIT named files
   GOSUB DoClean                                                  ' Clean them up

   '----- Clean up any old VSAVE files
   fn = ENV.VSavePath + "VSAVE.*"                                 ' Build VSave filename
   GOSUB DoClean                                                  ' Clean them up

   '----- Clean up any old RUN files
   fn = ENV.RunPath + "*.*"                                       ' Build an ALL mask
   GOSUB DoClean                                                  ' Clean them up

   Expiry = 180                                                   ' Do the 180 day expiry guys
   '----- Clean up any old STATE files
   fn = ENV.StatePath + "*.*"                                     ' Build an ALL mask
   GOSUB DoClean                                                  ' Clean them up


   MExitFunc                                                      ' We're done

DoClean:
   kfn = DIR$(fn, TO FD)                                          '
   DO WHILE ISNOTNULL(kfn)                                        ' See if any files to delete
      txt1 = sMakePrettyTime(FD.LastWriteTime)                    ' Get timestamp from file
      IF SNumDaysSince(MID$(txt1, 6, 2) + "-" + MID$(txt1, 9, 2) + "-" + LEFT$(txt1, 4)) > Expiry THEN ' If more than Expiry days old
         TRY                                                      '
            KILL fn                                               ' Delete it
            EXIT TRY                                              '
         CATCH                                                    '
            EXIT TRY                                              '
         END TRY                                                  '
      END IF                                                      '
      kfn = DIR$(NEXT, TO FD)                                     ' Try for another
   LOOP                                                           '
   RETURN

   EXIT FUNCTION                                                  ' We're done
END FUNCTION


FUNCTION InitSeeUnique() AS LONG
'---------- See if we should be running as one instance
LOCAL fWnd, MyThread, OtherThread AS LONG, t AS STRING
STATIC DataToSend AS COPYDATASTRUCT
STATIC pzBuffer   AS ASCIIZ PTR
STATIC zBuffer    AS STRING * %MAX_PATH + 50
   MEntry
   IF ISFALSE ENV.UniqueFlag THEN MExitFunc                       ' If user doesn't want Unique, then just exit

   '----- See if we're running somewhere else
   fWnd = sFindWindow("SPFLite(v" + LEFT$(ENV.PgmVers, 4))        ' Go see if we're running somewhere
   IF fWnd = 0 OR IsENVOpen THEN MExitFunc                        ' No handle returned, or forced by -OPEN then we're fine

   '----- Tell the running instance to 'take over' this request
   t = ENV.InitString                                             ' Get Init String
   sUnQuote(t): t = TRIM$(t)                                      ' Remove quotes
   IF isENVBrowse THEN                                            ' BROWSE requested?
      zBuffer = IIF$(ISNULL(t), "BROWSE ", "BROWSE " + $DQ + t + $DQ) ' Build the command to pass
   ELSEIF isENVView THEN                                          ' VIEW requested?
      zBuffer = IIF$(ISNULL(t), "VIEW ", "VIEW " + $DQ + t + $DQ) ' Build the command to pass
   ELSEIF IsENVClip THEN                                          ' CLIP requested?
      zBuffer = "CLIP "                                           '
   ELSE                                                           ' Else a simple edit
      zBuffer = "EDIT " + $DQ + t + $DQ                           '
   END IF                                                         '
   IF TRIM$(zBuffer) <> "EDIT " + CHR$(34,34) THEN                ' If we've actually got a requested file etc.
      DataToSend.lpData  = VARPTR(zBuffer)                        ' Build the message
      DataToSend.cbdata  = LEN(zBuffer) + 1                       '
      DataToSend.dwData  = 1                                      '
      SendMessage fWnd, %WM_COPYDATA, hWnd, VARPTR(DataToSend)    ' Send data to other instance
   END IF                                                         '

   '----- Tell our caller we passed it off, this instance will be shut down
   FUNCTION = %True                                               '

   '----- Make sure the other window is moved to the top to show that 'something' has happened.
   MyThread = GetWindowThreadProcessId(hWnd, 0)                   ' Do the fudge to put the other
   OtherThread = GetWindowThreadProcessId(fWnd, 0)                ' Window active in the foreground
   AttachThreadInput(MyThread, OtherThread, %TRUE)                '
   SetForegroundWindow(fWnd)                                      '
   AttachThreadInput(MyThread, OtherThread, %FALSE)               '
   IF IsIconic(fWnd) THEN                                         '
      ShowWindow(fWnd, %SW_RESTORE)                               '
   ELSEIF IsZoomed(fWnd) THEN                                     '
      ShowWindow(fWnd, %SW_SHOWMAXIMIZED)                         '
   ELSE                                                           '
      ShowWindow(fWnd, %SW_SHOW)                                  '
   END IF                                                         '
   MExit
END FUNCTION
