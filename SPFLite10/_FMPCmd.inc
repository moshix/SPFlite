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

'---------- File Manager Primary Commands

METHOD FMPCmdAll(pCmd AS STRING)
'---------- ALL FM Command
LOCAL lCmd, MEditCmd AS STRING, i, j AS LONG
LOCAL hFile AS LONG, hfile2 AS DWORD, fname AS WSTRINGZ * 260
LOCAL CreationTime, LastAccessTime, LastWriteTime AS FILETIME
LOCAL SystemTimeArea AS SYSTEMTIME
   MEntry
   lCmd = sGetWord(pCmd, %Strip, %QuoteNotSig)                    ' Strip off the ALL command
   lCmd = sGetWord(pCmd, %Strip, %QuoteNotSig)                    ' Strip off the command to be performed

   '----- Split based on type of requested command
   SELECT CASE AS CONST$ UUCASE(lCmd)                             '

      '----- ALL BROWSE
      CASE "B", "BRO", "BROWSE"
         FOR i = 1 TO AFCount                                     ' Lets look th the lines
            IF AFList(i).Flag = %FEntry OR AFList(i).Flag = %FFLEntry THEN ' Only look at real file entries
               me.DoBROWSE(i)                                     ' Go do Browse
            END IF                                                '
         NEXT i                                                   '

      '----- ALL CLONE
      CASE "C", "CLONE"
         FOR i = 1 TO AFCount                                     ' Lets look th the lines
            IF AFList(i).Flag = %FEntry OR AFList(i).Flag = %FFLEntry THEN ' Only look at real file entries
               me.CallTab("CLONE", $DQ + RTRIM$(AFList(i).Path) + RTRIM$(AFList(i).FD.FileName) + $DQ)
            END IF                                                '
         NEXT i                                                   '

      '----- ALL Edit
      CASE "E", "EDIT", "S", "SEL", "SELECT"
         FOR i = 1 TO AFCount                                     ' Lets look th the lines
            IF AFList(i).Flag = %FEntry OR AFList(i).Flag = %FFLEntry THEN ' Only look at real file entries
               me.DoEDIT(i)                                       ' Go do Browse
            END IF                                                '
         NEXT i                                                   '

      '----- ALL MEDIT
      CASE "M", "MEDIT"
         FOR i = 1 TO AFCount                                     ' Lets look asth the lines
            IF AFList(i).Flag = %FEntry OR AFList(i).Flag = %FFLEntry THEN ' Only look at real file entries
               IF ISNULL(MEditCmd) THEN                           ' First M command?
                  MEditCmd = $DQ + RTRIM$(AFList(i).Path) + RTRIM$(AFList(i).FD.FileName) + $DQ + " "
                  MEditCmd += " NEW "                             '
               ELSE                                               ' Else just add a filename on
                  MEditCmd += $DQ + RTRIM$(AFList(i).Path) + RTRIM$(AFList(i).FD.FileName) + $DQ + " "
               END IF                                             '
            END IF                                                '
         NEXT i                                                   '
         me.CallTab("MEDIT", MEditCmd)                            '

      '----- ALL PRINT
      CASE "P", "PRINT"
         FOR i = 1 TO AFCount                                     ' Lets look th the lines
            IF AFList(i).Flag = %FEntry OR AFList(i).Flag = %FFLEntry THEN ' Only look at real file entries
               me.FMLCmdPrintIt($DQ + RTRIM$(AFList(i).Path) + RTRIM$(AFList(i).FD.FileName) + $DQ, 0)            '
            END IF                                                '
         NEXT i                                                   '
         MErrExit(0, "All files Printed")                         '

      '----- ALL TOUCH
      CASE "T", "TOUCH"
         FOR i = 1 TO AFCount                                     ' Lets look th the lines
            IF AFList(i).Flag = %FEntry OR AFList(i).Flag = %FFLEntry THEN ' Only look at real file entries
               TRY                                                ' Ignore errors
                  hFile = FREEFILE                                ' Touch the file now
                  OPEN RTRIM$(AFList(i).Path) + RTRIM$(AFList(i).FD.FileName) FOR APPEND ACCESS READ WRITE LOCK SHARED AS #hFile
                  hFile2 = FILEATTR(hFile,2)                      ' Get Windows handle for the file
                  GetFileTime hFile2, CreationTime, LastAccessTime, LastWriteTime' Fetch timestamps
                  GetSystemTime SystemTimeArea                    ' Get the current time
                  SystemtimeToFileTime SystemTimeArea, LastWriteTime ' Put it in LastWriteTime
                  SetFileTime hFile2, CreationTime, LastAccessTime, LastWriteTime
               CATCH                                              '
               END TRY                                            '
               CLOSE #hFile                                       ' Close it
            END IF                                                '
         NEXT i                                                   '
         DoSet(%LoadData)                                         ' Ask for refresh
         MErrExit(0, "All files Touched")                         '

      '----- ALL VIEW
      CASE "V", "VIEW"
         FOR i = 1 TO AFCount                                     ' Lets look th the lines
            IF AFList(i).Flag = %FEntry OR AFList(i).Flag = %FFLEntry THEN ' Only look at real file entries
               me.DoVIEW(i)                                       ' Go do View
            END IF                                                '
         NEXT i                                                   '

      CASE ELSE
         MErrExit(%eFail, "Unsupported ALL command type: " + lcmd)
   END SELECT
   MExit                                                          '
END METHOD

METHOD  FMPCmdBottom(pCmd AS STRING)
'---------- BOTTOM in FM Mode
   MEntry
   LastTop = 0                                                    ' Allow movement
   TopScrn = MAX(1, MIN(AFCount, 9999999) - (ENV.ScrHeight + 1 - (3 * ENV.FMHelpFlag) - FM_Top_File_Line) + 1) ' Move TopScrn
   MExit
END METHOD

METHOD  FMPCmdBrowse(pCmd AS STRING)
'---------- BROWSE in FM Mode
LOCAL t AS STRING
   t = sGetWord(pCmd, %Strip, %QuoteSig)                          ' Strip off command
   me.CallTab("BROWSE", pCmd)                                     ' Do a normal Browse
END METHOD

METHOD  FMPCmdCase(pCmd AS STRING)
'---------- CASE in FM Mode
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN MErrExit(%eFail, "Missing CASE operand")'
   SELECT CASE AS LONG pCmdOpsType(1)                             ' Which?
      CASE %KWChange: Prf.PCase = "C"                             ' "C"
      CASE %KWT     : Prf.PCase = "T"                             ' "T"
      CASE ELSE     : MErrExit(%eFail, "Invalid CASE operand")    '
   END SELECT                                                     '
   SaveCase = Prf.PCase                                           ' Save the volatile ones
   scError(%eNone, "CASE temporarily set to: " + Prf.PCase)       '
   MExit                                                          '
END METHOD

METHOD  FMPCmdDefault(pCmd AS STRING)
'---------- Set the Default LMB command string
LOCAL Cmd AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps > 0 THEN DefCommand = pCmdOps(1) + " " + pCmdOps(2) + " " + pCmdOps(3)  ' Save it
   Cmd = UUCASE(LSET$(pCmdOps(1), 8))                             ' Get a fixed length version
   IF INSTR("D       DD      DEL     DELETE  PURGE   P       U       UU      PUR     ", Cmd) > 0  AND _
      ISFALSE ENV.ConfDelFlag THEN                                ' Dangerous command?
      sDoMsgBox "|KWARNING:|B The specified command can be dangerous as your" + $CRLF + _
                "options are not currently set to have deletions confirmed", %MB_OK OR %MB_USERICON, "SPFLite Warning"
   END IF                                                         '
   scError(%eNone, "DEFAULT set to " + DefCommand)                '
   sDoStatusBar($SBMode)                                          ' Correct Status Mode Bar
   MExit                                                          '
END METHOD

METHOD  FMPCmdDOWN(pCmd AS STRING)
'---------- DOWN in FM Mode
LOCAL thispage, onepage, csrpage, j AS LONG, u AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   onepage = (ENV.ScrHeight + 1 - (3 * ENV.FMHelpFlag) - FM_Top_File_Line)
   csrpage = IIF((CRow - FM_Top_File_Line) < 1, onepage, (CRow - FM_Top_File_Line))

   LastTop = 0                                                    ' Allow movement
   u = pCmdOps(1)                                                 ' Pick up operand
   IF pCmdNumOps = 0 THEN u = Prf.Scroll                          ' Default if none provided
   IF VAL(u) > 0 THEN                                             ' If numeric, use it
      thispage = VAL(u)                                           '
      j = 0                                                       '
   ELSE                                                           '
      u = UUCASE(LSET$(u, 4))                                     ' pad to 4
      j = INT(INSTR("    M   MAX P   PAGEF   FULLH   HALFD   DATAC   CSR ", u) / 4)
      IF j = 0 THEN MErrExit(%eFail,"Invalid Scroll amount")      '
      thispage = CHOOSE(j, 9999999,    9999999,      onepage,     onepage,     onepage, onepage, _ ' M/Max,  P/Page, F/Full
                           onepage / 2, onepage / 2, onepage - 1, onepage - 1, onepage, onepage)   ' H/Half, D/Data, C/Csr
   END IF                                                         '
   IF j = 1 OR j = 2 THEN                                         ' MAX
      TopScrn = MAX(1, MIN(AFCount, 9999999) - onepage + 1)       ' Move TopScrn
   ELSEIF j = 11 OR j = 12 THEN                                   ' CSR
      TopScrn += csrpage                                          ' Adjust top
      IF TopScrn > AFCount THEN TopScrn = MAX(1, MIN(AFCount, 9999999) - onepage + 1)
   ELSE                                                           '
      TopScrn = MIN(AFCount, TopScrn + thispage)                  ' Move TopScrn
   END IF                                                         '
   FMMarkdSLin = 0                                                ' Clear marked line
   MExit                                                          '
END METHOD

METHOD  FMPCmdEdit(pCmd AS STRING)
'---------- EDIT in FM Mode
LOCAL t AS STRING
   t = sGetWord(pCmd, %Strip, %QuoteSig)                          ' Strip off command
   me.CallTab("EDIT", pCmd)                                       ' Do a normal Edit
END METHOD

METHOD  FMPCmdEnd(pCmd AS STRING)
'---------- END in FM Mode
LOCAL J AS LONG
   MEntry
   IF gfEndAll THEN                                               ' Driven by EXIT?
      TMode = (TMode AND (&HFFFFFFFF - %MFMTab))                  ' Remove FM status
      me.TabTitleSet(%True)                                       '
      pCmdEnd("END")                                              ' Call END
      MExitMeth                                                   ' All done here
   END IF                                                         '

   IF ISNOTNULL(FileListNm) THEN                                  ' In filelist mode, Go back a nest level
      IF NestCtr > 0 THEN                                         '
         me.NestDel: MExitMeth                                    ' If one is available
      ELSE                                                        ' Else go back to File/Path mode
         FileListNm = "": DoSet(%LoadReq)                         '
      END IF                                                      '
   ELSE                                                           '
      IF NestCtr > 0 THEN                                         '
         me.NestDel: MExitMeth                                    ' If one is available
      ELSE                                                        '
         j = INSTR(-2, FPath, "\")                                ' Find last \
         IF j <> 0 THEN                                           ' Got one
            FPath = LEFT$(FPath, j)                               ' Cut off the last level
            sIniSetString("FManager", "DefDir1", FPath)           '
         END IF                                                   ' Else do nothing
         DoSet(%LoadReq)                                          ' Request refresh
      END IF                                                      '
   END IF                                                         '
   MExit                                                          '
END METHOD

METHOD  FMPCmdExclude(pCmd AS STRING)
'---------- EXCLUDE in FM Mode
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL t AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps <> 1 THEN MErrExit(%eFail, "EXCLUDE must have 1 single operand")
   IF pCmdOpsType(1) <> %OpStr AND pCmdOpsType(1) <> %OpQStr THEN _ ' Normal string?
      MErrExit(%eFail, "Unsupported EXCLUDE operand")             '
   FSearch = UUCASE(pCmdOps(1))                                   ' Get search operand
   LastTop = 0                                                    ' Allow movement
   FOR i = 1 TO AFCount                                           ' Lets do a search for matches
      IF AFList(i).Flag = %FEntry OR AFList(i).Flag = %FFLEntry OR AFList(i).Flag = %FFileList THEN
         IF ISTRUE me.TestMask(TRIM$(AFList(i).FD.FileName), FSearch) THEN ' If we pass the Mask test
            me.IncRXSize                                          ' Allocate an RX slot
            RXList(RXCount) = TRIM$(AFList(i).Path) + TRIM$(AFList(i).FD.FileName) ' Create an Exclude entry
            DoSet(%LoadData)                                      ' Have it take effect
            INCR j                                                ' Count items
         END IF                                                   '
      END IF                                                      '
   NEXT i                                                         '
   DoSet(%LoadData)
   scError(%eNone, IIF$(j = 0, "No entries excluded", FORMAT$(j) + " entries excluded"))
   MExit                                                          '
END METHOD

METHOD  FMpCmdFF(pCmd AS STRING)
'---------- FF (File Find) command
LOCAL action, lclWord, lclTag, valid, fn, MSG, d, fcmd, XList, X AS STRING
LOCAL i, j, AllCount, lclNF, FNum AS LONG, ff AS QUAD
DIM foundlist(1 TO 100) AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN MErrExit(%eFail, "Missing FF operands") ' Exit if nothing to do
   fcmd = pCommand                                                ' Save the whole FF command
   me.FindReset                                                   ' Reset the parse output fields
   IF sCRTParse("MNC1") THEN MExitMeth                            ' Scan (Modifier, Negative, Cols, Lit1)
   me.FindSetup                                                   ' Copy CRT parse's answers
   IF ISFALSE IsCfLit1 THEN MErrExit(%eFail, "No search literal specified") ' Better have a Lit
   IF IsCfNF THEN lclNF = %True: CfClear(%CrtNF)                  ' If NF, save flag and then clear it

   AllCount = 0                                                   ' Reset counter
'   GRAPHIC ATTACH PgHandle, WindowID                             ' Set as the default graphic area
   gfInterrupt = %False                                           '
   DispInterrupt                                                  ' Display the Interrupt popup

   FOR i = 1 TO AFCount                                           ' Lets see if any lines were selected
      IF gfInterrupt THEN EXIT FOR                                '
      IF AFList(i).Flag = %FEntry OR AFList(i).Flag = %FFLEntry THEN  ' Existing File?
         x = "," + UUCASE(MID$(RTRIM$(AFList(i).FD.FileName), INSTR(-1, RTRIM$(AFList(i).FD.FileName), ".") + 1)) + ","
         XList = "," + UUCASE(ENV.FMFFList) + ","                 ' Build XList here since a ProfSkip might have added an entry
         XList = REMOVE$(XList, " ")                              ' Eliminate any spaces
         IF INSTR(XList, x) THEN ITERATE FOR                      ' Skip if exempt file
         CONTROL SET TEXT hIntr, %INTERRUPT_TEXT2, RTRIM$(AFList(i).FD.FileName)
         CONTROL SET TEXT hIntr, %INTERRUPT_TEXT3, "Found: " + FORMAT$(AllCount)
         DIALOG DOEVENTS 0                                        ' Let others process
         RESET gLoopCtr                                           ' Reset LoopCtr to avoid treating as a loop
         me.LInitTxtData(RTRIM$(AFList(i).Path) + RTRIM$(AFList(i).FD.FileName)) ' Initialize our Text area
         TIP.Setup("E", "", "", RTRIM$(AFList(i).Path) + RTRIM$(AFList(i).FD.FileName)) ' Set filename
         TIP.EXEC                                                 ' Go do it
         me.InitaFile(%True)                                      ' Initialize file stuff
         j = me.Search(0, %CursBad)                               ' Normal version
         IF ISTRUE j AND ISFALSE lclNF OR _                       ' See if we have one
            ISFALSE j AND ISTRUE lclNF THEN                       '
            INCR AllCount                                         ' Count a found file
            IF AllCount > UBOUND(foundlist()) THEN _              ' Table still big enough?
               REDIM PRESERVE foundlist(1 TO 2 * UBOUND(foundlist())) ' No, enlarge it
            foundlist(AllCount) = RTRIM$(AFList(i).Path) + RTRIM$(AFList(i).FD.FileName)
         END IF                                                   '
         sFileQueue("D", " ", TIP.FilePath)                       ' Remove from Open queue
      END IF                                                      '
   NEXT i                                                         '
   ErrMsg = ""                                                    ' Clear any errors from Edit
   errFlag = 0                                                    '
   DIALOG END hIntr                                               ' End the window
   Prf.ReadAll(%True)                                             ' Overlay with DEFAULT if available
   Prf.PCase = SaveCase                                           ' Restore the volatile ones
   FindWord = SaveWord                                            '
   gfInterrupt = %False                                           '
   LastLine = 3                                                   ' Prevent Editor opens (if last file was empty or Skipped outouf)

   IF AllCount = 0 THEN                                           ' Issue approp. message
      MErrExit(%eFail, "No Files match search")                   '

   ELSE                                                           '
      scError( %eNone, FORMAT$(AllCount) + IIF$(AllCount > 1, " files", " file") + " satisfied search.")
      TOP.Setup("RECO", "", "", ENV.FileListPath + "Found Files.FLIST") ' Setup for IO
      IF TOP.EXEC THEN                                            ' Go OPEN the file
         MErrExit(%eFail, TOP.ResultMsg)                          ' Oops?  Bail out
      END IF                                                      '
      PRINT # TOP.FNum, "FileSrch: " + fCmd                       ' Save the search string
      FOR i = 1 TO AllCount                                       ' Write the array back out
         PRINT # TOP.FNum, foundlist(i)                           '
      NEXT i                                                      '
      TOP.Close                                                   ' Close the FBO
      FileListNm = "Found Files"                                  ' Set FOUND as the new filelist
      DoSet(%LoadReq OR %Refresh)                                 ' Refresh
   END IF                                                         '
   MExit                                                          '
END METHOD

METHOD  FMPCmdFind(pCmd AS STRING)
'---------- FIND in FM Mode
REGISTER i AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN MErrExit(%eFail, "Missing FIND operand")'
   IF pCmdNumOps = 1 AND pCmdOpsType(1) <> %OpQStr AND _          ' Single operand unquoted WORD?
      (IsEQ(pCmdOps(1), "WORD") OR IsEQ(pCmdOps(1), "WORDS")) THEN' or WORDS
      FindWord = %True                                            '
      MErrExit(%eNone, "FIND mode set temporarily to: WORD")      '
   END IF                                                         '
   IF pCmdNumOps = 1 AND pCmdOpsType(1) <> %OpQStr AND _          ' Single operand unquoted CHAR?
      (IsEQ(pCmdOps(1), "CHAR") OR IsEQ(pCmdOps(1), "CHARS")) THEN' or CHARS
      FindWord = %False                                           '
      MErrExit(%eNone, "FIND mode set temporarily to: CHARS")     '
   END IF                                                         '

   FSearch = UUCASE(pCmdOps(1)): LastLoc = TopScrn                ' Get search operand
   sUnQuote(FSearch)                                              ' Remove quotes
   LastTop = 0                                                    ' Allow movement
   FOR i = LastLoc TO AFCount                                     ' Lets see where it is
      IF INSTR(UUCASE(AFlist(i).FD.FileName), FSearch) > 0 THEN   ' Found it?
         TopScrn = i: LastLoc = i                                 ' Set top of screen here
         me.CsrRow = FM_Top_File_Line: me.CsrCol = 1: MExitMeth   '
      END IF                                                      '
   NEXT i                                                         '
   scError(%eNone, "Bottom of list reached")                      '
   LastLoc = 1                                                    '
   MExit                                                          '
END METHOD

METHOD  FMPCmdLocate(pCmd AS STRING)
'---------- LOCATE in FM Mode
LOCAL i AS LONG, u AS STRING
   MEntry
   IF LEFT$(TP.DefSort, 4) <> "Name" THEN _                       ' Only if sorted correctly
      MErrExit(%eFail,"LOCATE requires files in Name+ or Name- order")
   IF ME.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN MErrExit(%eFail,"LOCATE requires a search filename")
   u = UUCASE(pCmdOps(1))                                         ' Ucase the arg
   LastTop = 0                                                    ' Allow movement
   IF MID$(TP.DefSort, 5, 1) = "+" THEN                           ' Ascending sort?
      FOR i = 1 TO AFCount                                        ' Lets see where it is
         IF AFList(i).Flag = %FEntry OR AFList(i).Flag = %FFLEntry OR AFList(i).Flag = %FProfile THEN ' If a normal file entry
            IF u <= UUCASE(LEFT$(AFList(i).FD.FileName, LEN(u))) THEN ' Found our place?
               TopScrn = MAX(i - 1, 1): LastLoc = MAX(i - 1, 1)   ' Set top of screen here
               EXIT FOR                                           '
            END IF                                                '
         END IF                                                   '
      NEXT i                                                      '
   ELSE                                                           ' No, Descending
      FOR i = 1 TO AFCount                                        ' Lets see where it is
         IF AFList(i).Flag = %FEntry OR AFList(i).Flag = %FFLEntry OR AFList(i).Flag = %FProfile THEN ' If a normal file entry
            IF u >= UUCASE(LEFT$(AFList(i).FD.FileName, LEN(u))) THEN  ' Found our place?
               TopScrn = MAX(i - 1, 1): LastLoc = MAX(i - 1, 1)   ' Set top of screen here
               EXIT FOR                                           '
            END IF                                                '
         END IF                                                   '
      NEXT i                                                      '
   END IF
   MExit
END METHOD

METHOD  FMpCmdMAKELIST(pCmd AS STRING)
'---------- MAKELIST command for File Manager
LOCAL ListName, LN2, fn, MSG, d, t AS STRING
LOCAL i, j, k, SymFlag, RepFlag, AppFlag, ExistFlag, OldFLCtr, SymCtr AS LONG
DIM OldFL(1 TO 1) AS STRING
DIM SymEnt(1 TO 1) AS STRING
LOCAL MIO AS iIO                                                  ' For our I/O stuff
   MEntry
   '----- Validate operands
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN _                                       ' Exit if nothing to do
      MErrExit(%eFail, "Missing MAKELIST operands")               '
   LET MIO = CLASS "cIO"                                          '

   FOR i = 1 TO pCmdNumOps                                        ' Process each operand
      SELECT CASE AS CONST$ UUCASE(pCmdOps(i))                    '
         CASE "SYM"                      : SymFlag = %True        ' Set flags as needed
         CASE "REPLACE", "REPL", "REP"   : RepFlag = %True        '
         CASE "APPEND"                   : AppFlag = %True        '
         CASE ELSE                                                '
            IF ISNULL(ListName) THEN                              ' 1st time, save as ListName
               ListName = pCmdOps(i)                              ' Save the list name
            ELSE                                                  '
               MErrExit(%eFail, "Unknown MAKELIST operand: " + pCmdOps(i))
            END IF                                                '
      END SELECT                                                  '
   NEXT i                                                         '

   IF ISFILE(ENV.FileListPath + ListName + ".FLIST") THEN         ' Test if exists
      ExistFlag = %True                                           '
      IF ISFALSE RepFlag AND ISFALSE AppFlag THEN _               '
         MErrExit(%eFail, ListName + ".FLIST already exists, APPEND/REPLACE not specified")
   ELSE                                                           '
      IF ISTRUE AppFlag THEN _                                    '
         MErrExit(%eFail, ListName + ".FLIST doesn't exist, but APPEND requested")
   END IF                                                         '

   LN2 = UUCASE(ListName)                                         ' Uppercase it
   IF LN2 = "RECENT FILES" OR LN2 = "RECENT PATHS" OR _           ' Eliminate reserved ones
      LN2 = "FOUND FILES" OR LN2 = "FAVORITE FILES" OR _          '
      LN2 = "RECENT" OR LN2 = "PATHS" OR ln2 = "FOUND" OR _       '
      LN2 = "NAMED FAVORITES" OR LN2 = "OPEN Files" THEN _        '
         MErrExit(%eFail, "Reserved FILELIST name, cannot use")   '

   IF ISTRUE ExistFlag AND ISTRUE AppFlag THEN                    ' Are we appending?
      MIO.Setup("I", "", "", ENV.FileListPath + ListName + ".FLIST")  ' Tell IO what we're opening
      MIO.EXEC                                                    '
      FILESCAN # MIO.FNum, RECORDS TO i                           ' Get the number of records
      REDIM OldFL(1 TO i + AFCount) AS STRING                     ' Redim array to match save data
      LINE INPUT # MIO.FNum, OldFL() TO OldFLCtr                  ' Read it all
      MIO.Close                                                   ' Close it
   END IF                                                         '

   '----- SYM requested?
   IF SymFlag THEN                                                ' Is SYM requested?
      REDIM SymEnt(1 TO AFCount) AS STRING                        ' Set table size
      FOR i = 1 TO AFCount                                        ' Collect paths
         IF AFList(i).Flag = %FEntry OR AFList(i).Flag = %FFLEntry THEN
            INCR SymCtr: SymEnt(SymCtr) = AFList(i).Path          ' Save the path
         END IF                                                   '
      NEXT i                                                      '
      ARRAY SORT SymEnt() FOR SymCtr, COLLATE UCASE               ' Sort them
      FOR i = 1 TO SymCtr                                         ' Now remove duplicates
         IF i < SymCtr THEN                                       ' If not at the end
            IF IsEQ(SymEnt(i), SymEnt(i+1)) THEN                  ' We have a duplicate
               ARRAY DELETE SymEnt(i)                             ' Remove it
               DECR SymCtr: DECR i                                ' Adjust counters
            END IF                                                '
         END IF                                                   '
      NEXT i                                                      '

   '----- Do a full normal FileList type
   ELSE
      REDIM SymEnt(1 TO AFCount) AS STRING                        ' Set table size
      FOR i = 1 TO AFCount                                        ' Lets see if any lines were selected
         IF AFList(i).Flag = %FEntry OR AFList(i).Flag = %FFLEntry OR AFList(i).Flag = %FOpen THEN
            INCR SymCtr                                           ' Save the path
            SymEnt(SymCtr) = TRIM$(AFList(i).Path) + TRIM$(AFList(i).FD.FileName)
         END IF                                                   '
      NEXT i                                                      '
   END IF                                                         '

   '----- Merge in the OldFL if APPEND loaded it
   IF OldFLCtr > 0 THEN                                           ' Some already?
      FOR i = 1 TO OldFLCtr                                       ' See if already here
         FOR j = 1 TO SymCtr                                      ' See if its in the new items
            IF IsEQ(OldFl(i), SymEnt(j)) THEN                     '
               OldFl(i) = ""                                      ' Null the old one
            END IF                                                '
         NEXT j                                                   '
      NEXT i                                                      '
      FOR i = 1 TO OldFLCtr                                       ' See if already here
         IF OldFl(i) <> "" THEN                                   ' Add remaining to SymEnt
            INCR SymCtr                                           ' Save the path
            SymEnt(SymCtr) = OldFL(i)                             '
         END IF                                                   '
      NEXT i                                                      '
   END IF                                                         '

   '----- Write it now
   IF SymCtr = 0 THEN _                                           ' Did we end up with anything?
      MErrExit(%eNone, "MAKELIST skipped, no eligible file entries")
   MIO.Setup("OR", "", "", ENV.FileListPath + ListName + ".FLIST")' Tell IO what we're opening
   IF MIO.EXEC THEN _                                             '
      MErrExit(%eFail, MIO.ResultMsg)                             ' Oops?  Bail out
   FOR i = 1 TO SymCtr                                            '
      IF ISFALSE INSTR(SymEnt(i), ",") THEN                       ' If not an Old item with operands
         SymEnt(i) = SymEnt(i) + ",*"                             ' Add an * mask
      END IF                                                      '
      PRINT # MIO.FNum, SymEnt(i)                                 ' Write the entry out
   NEXT i                                                         '

   '----- Close up now
   MIO.Close                                                      ' Close the FBO
   scError(%eNone, ListName + ".FLIST " + IIF$(Appflag, "appended", IIF$(Repflag, " replaced", "created"))) '
   MExit
END METHOD

METHOD FMPCmdRecall(pCmd AS STRING)
'---------- RECALL in FM Mode
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN pCmdOps(1) = "Recent Files"             ' No ops, use RECENT
   SELECT CASE AS CONST$ UUCASE(pCmdOps(1))                       ' Correct into full name
      CASE "RECENT"                       : pCmdOps(1) = "Recent Files"
      CASE "FAV", "FAVOURITE", "FAVORITE" : pCmdOps(1) = "Favorite Files"
      CASE "FOUND"                        : pCmdOps(1) = "Found Files"
      CASE "PATHS"                        : pCmdOps(1) = "Recent Paths"
      CASE "LIST"                         : pCmdOps(1) = "Named Favorites"
      CASE "OPEN"                         : pCmdOps(1) = "Open Files"
      CASE ELSE                                                   '
         IF ISFALSE ISFILE(ENV.FileListPath + pCmdOps(1) + ".FLIST") THEN _  ' See if the file exists
            MErrExit(%eFail,"Recall failed - " + pCmdOps(1) + ".FLIST does not exist")
   END SELECT                                                     '
   FileListNm = pCmdOps(1)                                        ' Swap to FileListNm
   DoSet(%LoadReq)                                                ' Call for a LoadReq
   MExit                                                          '
END METHOD

METHOD  FMPCmdReset(pCmd AS STRING)
'---------- RESET in FM Mode
LOCAL fF, fX, fALL, i AS LONG, RQPath, RQMask, RQFlags, RQNote AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN                                         ' If no ops, then just do X
      fX = %True                                                  '
   ELSE                                                           '
      FOR i = 1 TO pCmdNumOps                                     ' Look at what's there
         SELECT CASE AS CONST$ UUCASE(pCmdOps(i))                 '
            CASE "X", "EX", "EXCL", "EXCLUDE"  : fX = %True       '
            CASE "F", "FOR", "FORGET", "FORGOT": fF = %True       '
            CASE "ALL"                         : fALL = %True     '
            CASE ELSE                                             '
               MErrExit(%eFail, "Unknown RESET operand: " + pCmdOps(i))
         END SELECT                                               '
      NEXT i                                                      '
   END IF                                                         '

   IF ISTRUE fALL AND ISFALSE fX AND ISFALSE fF THEN fX=%True: fF = %True ' If just ALL, set X and F

   '----- If fX it's simple, delete the X table stuff
   IF fX THEN RXCount = 0: RESET RXList()                         ' Wipe it out

   '----- OK, do Forget if needed
   IF fF THEN
      '----- OK, now loop through RQ table resetting things
      FOR i = RQCount TO 1 STEP -1                                ' Go backwards so deletes don't muck things up
         me.RQSplit(RQList(i), RQPath, RQMask, RQFlags, RQNote)   ' Split things out
         IF INSTR(RQFlags, "F") = 0 THEN ITERATE FOR              ' Not an F entry, just skip onward
         RQFlags = REMOVE$(RQFlags, "F")                          ' Remove the F status
         IF INSTR(RQFlags, "G") AND RQNote = "" THEN              ' No more reason to exist?
            ARRAY DELETE RQList(i): DECR RQCount                  ' Remove it
         ELSE                                                     ' Else
            RQList(i) = BUILD$(RQPath, "|", RQMask, "|", RQFlags, "|", RQNote) ' Update the entry
         END IF                                                   '
      NEXT i                                                      '
   END IF                                                         '
   scError(%eNone, "RESET complete")                              '
   DoSet(%SaveReq)                                                ' Request save of this now
   DoSet(%LoadReq)                                                ' Refresh all
   MExit                                                          '
END METHOD

METHOD  FMPCmdRFind(pCmd AS STRING)
'---------- RFIND in FM Mode
LOCAL i AS LONG, u AS STRING
   MEntry
   IF ISNULL(FSearch) OR LastLoc = 0 THEN _                       ' Previous success
      MErrExit(%eFail,"No prior search")                          '
   LastTop = 0                                                    ' Allow movement
   FOR i = LastLoc + 1 TO AFCount                                 ' Lets see where it is
      IF INSTR(UUCASE(AFList(i).FD.FileName), UUCASE(FSearch)) > 0 THEN ' Found it?
         TopScrn = i: LastLoc = i                                 ' Set top of screen here
         me.CsrRow = FM_Top_File_Line: me.CsrCol = 1: MExitMeth
      END IF                                                      '
   NEXT i                                                         '
   scError(%eNone, "Bottom of list reached")                      '
   LastLoc = 1                                                    '
   MExit                                                          '
END METHOD

METHOD  FMPCmdTop(pCmd AS STRING)
'---------- Top in FM Mode
   LastTop = 0                                                    ' Allow movement
   TopScrn = 1                                                    ' Move TopScrn
END METHOD

METHOD  FMPCmdView(pCmd AS STRING)
'---------- VIEW in FM Mode
LOCAL t AS STRING
   t = sGetWord(pCommand, %Strip, %QuoteSig)                      ' Strip off command
   me.CallTab("VIEW", pCommand)                                   ' Do a normal View
END METHOD

METHOD  FMPCmdUp(pCmd AS STRING)
'---------- UP in FM Mode
LOCAL thispage, onepage, csrpage, j AS LONG, u AS STRING
   MEntry                                                         '
   onepage = (ENV.ScrHeight + 1 - (3 * ENV.FMHelpFlag) - FM_Top_File_Line)
   csrpage = IIF(onepage - (CRow - FM_Top_File_Line) < 1, onepage, onepage - (CRow - FM_Top_File_Line) - 1)
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   LastTop = 0                                                    ' Allow movement
   u = pCmdOps(1)                                                 ' Pick up operand
   IF pCmdNumOps = 0 THEN u = Prf.Scroll                          ' Default if none provided
   IF VAL(u) > 0 THEN                                             ' If numeric, use it
      thispage = VAL(u)                                           '
      j = 0                                                       '
   ELSE                                                           '
      u = UUCASE(LSET$(u, 4))                                     ' pad to 4
      j = INT(INSTR("    M   MAX P   PAGEF   FULLH   HALFD   DATAC   CSR ", u) / 4)
      IF j = 0 THEN MErrExit(%eFail,"Invalid Scroll amount")      '
      thispage = CHOOSE(j, 9999999,     9999999,     onepage,     onepage,     onepage, onepage, _ ' M/Max,  P/Page, F/Full
                           onepage / 2, onepage / 2, onepage - 1, onepage - 1, csrpage, csrpage)   ' H/Half, D/Data, C/Csr
   END IF                                                         '
   TopScrn = MAX(1, TopScrn - thispage)                           ' Move TopScrn
   FMMarkdSLin = 0                                                ' Clear marked line
   MExit                                                          '
END METHOD
