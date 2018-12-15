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

'---------- Primary command Methods
METHOD  pCmdACTION(pCmd AS STRING)
'---------- Handle ACTION Command
LOCAL i, j AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN                                         ' No operands, just display it
      IF Prf.ActionSave = 0 THEN                                  '
         MErrExit(%eNone, "ACTION set to 0 (OFF)")                '
      ELSE                                                        '
         MErrExit(%eNone, "ACTION set to " + FORMAT$(Prf.ActionSave) + ", " + Prf.ActionVerb)
      END IF                                                      '
   END IF                                                         '

   IF pCmdOpsType(1) = %OpNum THEN                                ' Numeric?
      Prf.ActionSave = VAL(pCmdOps(1))                            ' Set it
   ELSE                                                           '
      MErrExit(%eFail, "Unknown ACTION operand - " + pCmdOps(1))  '
   END IF                                                         '

   IF pCmdNumOps = 2 THEN                                         ' Now do the verb
      IF IsNE(pCmdOps(2), "SAVE") AND IsNE(pCmdOps(2), "VSAVE") THEN _
         MErrExit(%eFail, "Unknown ACTION operand - " + pCmdOps(2))
      Prf.ActionVerb = UUCASE(pCmdOps(2))                         ' It's OK, save it
   END IF                                                         '

   IF Prf.ActionSave = 0 THEN                                     '
      scError(%eNone, "ACTION set to 0 (OFF)")                    '
   ELSE                                                           '
      scError(%eNone, "ACTION set to " + FORMAT$(Prf.ActionSave) + ", " + Prf.ActionVerb)
   END IF                                                         '
   ActionCtr = Prf.ActionSave                                     '
   MExit
END METHOD

METHOD  pCmdADD(pCmd AS STRING)
'---------- ADD a line to the file
LOCAL i AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors

   IF ISTRUE me.RangeSet("", %False, %False) THEN MExitMeth       ' Setup any range command operands
   IF ISFALSE DotSet THEN MErrExit(%eFail, "Missing line reference") ' Better be a line reference

   IF DotStart = 0 THEN MErrExit(%eFail, "Missing line reference")' Got first reference?
   IF DotEnd <> 0 AND DotStart <> DotEnd THEN MErrExit(%eFail, "Only a single line reference is supported") ' If 2nd, Better be the same

   IF pCmdNumOps <> 1 THEN MErrExit(%eFail, "Must have just one string operand") ' One operand then
   IF pCmdOpsType(1) <> %OpStr AND pCmdOpsType(1) <> %OpQStr THEN MErrExit(%eFail, "Must have just one string operand")

   '----- OK Lets insert the line
   me.ModSet(DotStart)                                            ' Remember we changed something
   me.LInsertEmpty(DotStart, 1, %Data)                            '
   me.LTxtSet(DotStart + 1, pCmdOps(1))                           ' Stuff in the text
   me.AttrScan(DotStart + 1)                                      ' Recolorize

   '----- If MEdit mode ensure the MIX value is correct
   IF IsMEdit THEN                                                ' In MEdit mode?
      i = me.MEditTbl("I", FORMAT$(DotStart + 1))                 ' Go get the MIX value of previous line
      LMixS(DotStart + 1) = i                                     ' Set into the new line
   END IF                                                         '
   me.CurSetReq(%Position, DotStart + 1, 1, %True)                ' Position
   gfXRebuild = %True                                             ' Force exclude rebuild
   MExit
END METHOD

METHOD  pCmdAUTOBKUP(pCmd AS STRING)
'---------- Turn AutoBkup ON or OFF
LOCAL nv AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps > 0 THEN                                         ' If an operand, examine it
      nv = sOnOff(1)                                              ' Go get True/False value for operand 1
      IF nv < 0 THEN MExitMeth                                    ' Error, exit, error msg already issued
      Prf.AutoBkup = nv                                           ' Set the new value
   END IF                                                         '
   scError(0, "AUTOBKUP set to " + IIF$(Prf.AutoBkup, "ON", "OFF")) ' Tell user the value
   MExit
END METHOD

METHOD  pCmdAUTOCAPS(pCmd AS STRING)
'---------- Turn AutoCaps ON or OFF
LOCAL nv AS LONG, cMsg AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps > 0 THEN                                         ' If an operand, examine it
      nv = sOnOff(1)                                              ' Go get True/False value for operand 1
      IF nv < 0 THEN MExitMeth                                    ' Error, exit, error msg already issued
      Prf.AutoCaps = nv                                           ' Set the new value
      IF Prf.PCase = "C" THEN                                     ' Do we have to force CASE T ?
         Prf.PCase = "T"                                          '
         cMsg = ", and CASE forced to 'T'"                        '
      END IF                                                      '
   END IF                                                         '
   scError(0, "AUTOCAPS set to " + IIF$(Prf.AutoCaps, "ON", "OFF") + cMsg)  ' Tell user the value
   MExit
END METHOD

METHOD  pCmdAUTONUM(pCmd AS STRING)
'---------- Turn AutoNum ON or OFF
LOCAL nv AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps > 0 THEN                                         ' If an operand, examine it
      nv = sOnOff(1)                                              ' Go get True/False value for operand 1
      IF nv < 0 THEN MExitMeth                                    ' Error, exit, error msg already issued
      Prf.AutoNum = nv                                            ' Set the new value
   END IF                                                         '
   scError(0, "AUTONUM set to " + IIF$(Prf.AutoNum, "ON", "OFF")) ' Tell user the value
   MExit
END METHOD

METHOD  pCmdAUTOSAVE(pCmd AS STRING)
'---------- Alter the AutoSave setting
LOCAL i AS LONG, lclMsg AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN                                         ' If no operands, treat as ON
      Prf.AutoSave = 1                                            ' No parms = ON

   ELSE                                                           '

      FOR i = 1 TO pCmdNumOps                                     ' Test operands for validity
         IF INSTR(CHR$(%KWON, %KWOFF, %KWPROMPT, %KWNOPROMPT),CHR$(pCmdOpsType(i))) = 0 THEN _ ' Ensure valid keywords?
            MErrExit(%eFail,"Unknown AUTOSAVE operand - " + pCmdOps(1)) ' Issue error message
      NEXT i                                                      '

      IF pCmdNumOps = 1 THEN                                      ' Just 1 operand?
         IF INSTR(CHR$(%KWON, %KWOFF),CHR$(pCmdOpsType(1))) = 0 THEN _ ' Ensure valid keywords?
            MErrExit(%eFail,"Unknown AUTOSAVE operand - " + pCmdOps(1)) ' Issue error message
         Prf.AutoSave = IIF(pCmdOpsType(1) = %KWON, 1, 2)         ' Set ON NOPROMPT or OFF PROMPT

      ELSE                                                        ' We have two operands
         IF pCmdOpsType(1) = %KWON AND pCmdOpsType(2) = %KWNOPROMPT THEN ' ON NOPROMPT
            Prf.AutoSave = 1                                      '
         ELSEIF pCmdOpsType(1) = %KWON AND pCmdOpsType(2) = %KWPROMPT THEN ' ON PROMPT
            Prf.AutoSave = 3                                      '
         ELSEIF pCmdOpsType(1) = %KWOFF AND pCmdOpsType(2) = %KWPROMPT THEN ' OFF PROMPT
            Prf.AutoSave = 2                                      '
         ELSEIF pCmdOpsType(1) = %KWOFF AND pCmdOpsType(2) = %KWNOPROMPT THEN ' OFF NOPROMPT
            Prf.AutoSave = 0                                      '
         END IF                                                   '
      END IF                                                      '
   END IF                                                         '

   lclMsg = IIF$((Prf.AutoSave AND 1) = 1, "ON ", "OFF ")         ' Build msg string
   lclMsg = lclMsg + IIF$((Prf.AutoSave AND 2) = 2, "PROMPT", "NOPROMPT") ' Build msg string
   scError(0, "AUTOSAVE set to " + lclMsg)                        '
   MExit                                                          '
END METHOD

METHOD  pCmdBOTTOM(pCmd AS STRING)
'---------- Alias for DOWN MAX
   me.pCmdDown("DOWN MAX ")                                       '
END METHOD

METHOD  pCmdBOUNDS(pCmd AS STRING)
'---------- Set BOUNDS via the command line
LOCAL t AS STRING, l, r, rmax AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0  OR (pCmdNumOps = 1 AND IsEQ(pCmdOps(1), "MAX")) THEN ' No Operands? Or just MAX?
      Prf.BndLeft = 1: Prf.BndRight = 0                           ' Yes, reset to defaults
      Prf.BuildBndText                                            ' Create the BNDS line
      sDoStatusBar($SBBnds)                                       ' re-Do the StatusBar Bnds box
      scError(%eNone, "BOUNDS reset to defaults")                 '
      MExitMeth                                                   '
   END IF                                                         '

   IF pCmdNumOps = 1 THEN                                         ' Fudge a single operand
      pCmdOps(2) = pCmdOps(1)                                     ' Shift it over and
      pCmdOps(1) = "1"                                            ' Default first to i
      pCmdNumOps = 2                                              ' Make it look like two operands
   END IF

   IF pCmdNumOps <> 2 THEN _                                      ' Must have two operands now
      MErrExit(%eFail, "BOUNDS has invalid operands")             '

   l = IIF(pCmdOps(1) <> "*", VAL(pCmdOps(1)), Prf.BndLeft)       ' Set left bound

   IF pCmdOps(2) = "*" THEN                                       ' Use prev value?
      r = Prf.BndRight                                            ' Use last bounds
   ELSEIF IsEQ(pCmdOps(2), "MAX") THEN                            ' MAX?
      r = 0                                                       ' Remember MAX
   ELSE                                                           '
      r = VAL(pCmdOps(2))                                         ' Use value
   END IF                                                         '

   IF r <> 0 AND l >= r THEN MErrExit(%eFail, "Illogical BOUNDS operands") ' Better be reasonable

   Prf.BndLeft = l: Prf.BndRight = r                              ' Save margins
   scError(0, IIF$(Prf.BndLeft = 1 AND Prf.BndRight = 0, "BOUNDS reset to defaults", "BOUNDS established")) ' Tell user OK
      Prf.BuildBndText                                            ' Create the BNDS line
   sDoStatusBar($SBBnds)                                          ' re-Do the StatusBar Bnds box
   MExit                                                          '
END METHOD

METHOD  pCmdBROWSE(pCmd AS STRING)
'---------- OPEN a new file in BROWSE mode
LOCAL fn, SetProf, MSG AS STRING
   MEntry
   '----- Do basic parsing
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   ENV.PMode = %MBrowse                                           ' Set to BROWSE mode

   fn = IIF$(pCmdOps(1) = "*", sGetFnClipboard, pCmdOps(1))       ' Get the Fn for Browse

   '----- See if a specific Profile name provided
   IF pCmdNumOps > 1 AND LEFT$(pCmdOps(2), 1) <> "." THEN _       ' A profile request?
      MErrExit(%eFail, "Unknown BROWSE operand: " + pCmdOps(2))   ' Issue error
   SetProf = IIF$(pCmdNumOps > 1, UUCASE(MID$(pCmdOps(2), 2)), "")' Extract the Profile name

   '----- Setup to validate it a bit
   TOP.Setup("AEP", SetProf, "Select file for Browse", fn)        ' See if Already Open and if it exists
   IF TOP.EXEC THEN
      IF TOP.RSTabFound THEN                                      ' Open elsewhere?
         MSG = TOP.ResultMsg                                      ' Save message arounf Tab switch
         TP = Tabs(TOP.TabFound)                                  ' Switch to it
         gTabSwitch = TOP.TabFound                                ' Set global switch
         MErrExit(%eFail, MSG)                                    ' Issue error and exit
      END IF
      MErrExit(%eFail, TOP.ResultMsg)                             ' Some other error, Issue the error message
   END IF                                                         '
   fn = TOP.FilePath: SetProf = TOP.Profile                       ' Copy fn/profile in case from Prompted filename
   me.LoadTheText(fn, SetProf)                                    ' Go load it in
   MExit
END METHOD

METHOD  pCmdCANCEL(pCmd AS STRING)
'---------- Cancel without saving anything
LOCAL i, j, RC AS LONG
LOCAL cData AS ClrCnvt, fn, DelFlag, tstyle, errbase AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                             ' Do basic parsing, exit if errors

   IF pCmdNumOps = 1 AND (IsEQ(pCmdOps(1), "DELETE") OR IsEQ(pCmdOps(1), "DEL") OR IsEQ(pCmdOps(1), "PURGE")) THEN ' DELETE or DEL?
      errbase = "DELETE/PURGE option cannot be used "             ' Setup basic message
      IF IsMEdit    THEN MErrExit(%eFail, errbase + "in MEdit mode")       ' Not for MEdit mode
      IF IsBrowse   THEN MErrExit(%eFail, errbase + "in BROWSE mode")      ' Not for Browse Mode
      IF TIP.ROStat THEN MErrExit(%eFail, errbase + "on a Read Only file") ' Nor read-only
      DelFlag = IIF$(IsEQ(pCmdOps(1), "PURGE"), "U", "D")         ' Remember what type of delete

   ELSEIF pCmdNumOps = 1 THEN                                     ' Else probably error
      MErrExit(%eFail,"Unknown CANCEL operand - " + pCmdOps(1))   ' Issue error message
   END IF                                                         '

   me.CleanTabData                                                ' Wipe everything out then
   IF TIP.FilePath <> $Empty THEN                                 ' If not an (Empty) file
      IF ISNOTNULL(DelFlag) THEN                                  ' DELETE asked for
         IF ENV.ConfDelFlag THEN                                  ' Need we prompt user for an OK?
            tstyle = "Purge: |K"                                    ' Pick message text
            IF ENV.UseRecycle AND delflag = "D" THEN tstyle = "Recycle: |K" '
            j =sDoMsgBox("OK to " + tstyle + TIP.FilePath, %MB_YESNO OR %MB_USERICON OR %MB_DEFBUTTON2, "SPFLite Delete")
            IF j = %IDYES THEN                                    ' OK?
               RC = sRecycleBin(TIP.FilePath, DelFlag)            '
            ELSE                                                  '
               MErrExit(%eNone,"CANCEL function terminated")      ' Bail right out
            END IF                                                '
         ELSE                                                     '
            RC = sRecycleBin(TIP.FilePath, DelFlag)               '
         END IF                                                   '
      END IF                                                      '
   END IF                                                         '

   '----- Request tab be deleted
   INCR gTabDelCtr                                                ' Bump table index
   gTabDelList(gTabDelCtr) = PgNumber                             ' Stuff in our page number
   CmdStackNum = 0                                                ' If a can, no more commands allowed
   MExit
END METHOD

METHOD  pCmdCAPS(pCmd AS STRING)
'---------- Turn CAPS ON, OFF or AUTO
LOCAL Wrd, lclCmd AS STRING, nv AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors

   IF pCmdNumOps = 0 THEN                                         ' If no operands, treat as ON
      Prf.CapsDesired = 1: Prf.CapsActual = 1                     '

   ELSE                                                           '
      IF INSTR(CHR$(%KWON, %KWOFF, %KWAUTO),CHR$(pCmdOpsType(1))) = 0 THEN _ ' Ensure valid keywords?
         MErrExit(%eFail,"Unknown CAPS operand - " + pCmdOps(1))  ' Issue error message

      IF pCmdOpsType(1) = %KWON THEN                              ' ON
         Prf.CapsDesired = 1: Prf.CapsActual = 1                  '
      ELSEIF pCmdOpsType(1) = %KWOFF THEN                         ' OFF
         Prf.CapsDesired = 0: Prf.CapsActual = 0                  '
      ELSEIF pCmdOpsType(1) = %KWAUTO THEN                        ' AUTO
         IF Prf.CapsDesired = 2 THEN                              ' Already in AUTO
            Prf.CapsActual = IIF(Prf.CapsActual, 0, 1)            ' Yes, then toggle active mode
         ELSE                                                     '
            Prf.CapsDesired = 2                                   '
         END IF                                                   '
      END IF                                                      '
   END IF                                                         '

   IF Prf.CapsDesired <> 2 THEN                                   ' If simple ON/OFF
      scError(0, "CAPS set to " + IIF$(Prf.CapsDesired, "ON", "OFF"))   ' Tell user new value
   ELSE                                                           '
      scError(0, "CAPS set to " + IIF$(Prf.CapsActual, "AUTO:on", "AUTO:off")) ' Tell user new value
   END IF

   sDoStatusBar($SBCaps)                                          ' re-Do the StatusBar Caps box
   MExit
END METHOD

METHOD  pCmdCASE(pCmd AS STRING)
'---------- Set CASE mode
LOCAL cMsg AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors

   IF pCmdNumOps = 0 THEN MErrExit(%eNone, "CASE set to " + UUCASE(Prf.PCase))
   SELECT CASE AS LONG pCmdOpsType(1)                             ' Which one
      CASE %KWCHANGE: Prf.PCase = "C"                             ' CASE C  ?
      CASE %KWT     : Prf.PCase = "T"                             ' CASE T  ?
      CASE ELSE     : MErrExit(%eFail, "Invalid CASE operand")    '
   END SELECT                                                     '
   IF Prf.AutoCaps AND Prf.PCase = "C" THEN _                     ' Is AUTOCAPS ON?
      cMsg = ", WARNING: AUTOCAPS=ON, this may cause problems"    '
   scError(0, "CASE set to " + UUCASE(Prf.PCase) + cMsg)          '
   sDoStatusBar($SBCaseWord)                                      ' re-Do the StatusBar Case box
   MExit
END METHOD

METHOD  pCmdCHANGE(pCmd AS STRING)
'---------- CHANGE command
LOCAL lclCmd, lclWord, modemsg, str, origChange AS STRING, lclCursMode AS INTEGER
LOCAL i, AllCount, AllLines, LLastLine AS LONG
LOCAL ec AS LONG, et AS STRING
   MEntry
   LocFind = "FIND"                                               ' Remember what Loc / FIND command was last
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 1 AND pCmdOpsType(1) = %KWCS THEN              ' Just a CS operand
      Prf.ChangeMode = "C"                                        ' Set to CS mode
      scError(0, "CHANGE set to CS")                              ' Tell user new value
      sDoStatusBar($SBChange)                                     ' Change the status Change bar
      MExitMeth                                                   '
   ELSEIF pCmdNumOps = 1 AND pCmdOpsType(1) = %KWDS THEN          ' Just a DS operand
      Prf.ChangeMode = "D"                                        ' Set to DS mode
      scError(0, "CHANGE set to DS")                              ' Tell user new value
      sDoStatusBar("H")                                           ' Change the status bar
      sDoStatusBar($SBChange)                                     ' Change the status Change bar
      MExitMeth                                                   '
   END IF

   me.ResetFunc(%ResetFind)                                       ' Clear previous FIND hilighting
   lclCursMode = Zone                                             ' Save CursMode at start of command
   IF ISTRUE me.RangeSet("", %False, %False) THEN MExitMeth       ' Setup any range command operands
   me.FindReset                                                   ' Reset the parse output fields
   IF sCRTParse("ASDMEC12LPQHU") THEN MExitMeth                   ' Scan (All, Subset, Direct, Modifier, Cols, Lit1, Lit2, Left/Right, Pen, CPen, Shift, User)
   me.FindSetup                                                   ' Copy CRT parse's answers

   IF ISFALSE IsCfNext AND ISFALSE IsCfPrev AND _                 '
      ISFALSE IsCfFirst AND ISFALSE IsCfLast THEN _               '
      CfSet(%CRTNext)                                             ' Set NEXT if no others specified


   IF ISFALSE IsCfLit1 THEN MErrExit(%eFail, "Missing CHANGE search string")    ' Lit1 present?
   IF ISFALSE IsCfLit2 THEN MErrExit(%eFail, "Missing CHANGE new value string") ' Lit2 present?
   str = IIF$(LEN(CrtL1Raw) < 16, CrtL1Raw, LEFT$(CrtL1Raw, 14) + "..")         ' Build msg literal
   modemsg = IIF$(IsCfWord, "WORD ", "CHARS ")  + str             ' Build mode part
   origChange = cfChange                                          ' Save in case Map style

   IF IsCfAll THEN                                                ' The ALL version?
      RESET AllCount, AllLines, lLastLine                         ' Reset counters
      i = me.Search(0, lclCursMode)                               ' Do the initial search
      lclCursMode = %CursBad                                      ' Make it bad for subsequent searches

      IF i THEN                                                   ' See if we have one
         INCR AllCount                                            ' Count successful find
         IF cfFLine <> lLastLine THEN INCR AllLines: lLastLine = cfFLine '
         cfChange = origChange                                    ' In case a mapping change
         me.Change()                                              ' Go do the Change
      END IF                                                      '

      DO WHILE i                                                  ' Do till not found
         i = me.Search(1, lclCursMode)                            ' Do re-Finds
         IF i THEN                                                '
            INCR AllCount                                         ' Count successful finds
            IF cfFLine <> lLastLine THEN INCR AllLines: lLastLine = cfFLine '
            cfChange = origChange                                 ' In case a mapping change
            me.Change()                                           '
         END IF                                                   '
      LOOP                                                        '

      ec = ErrFlag: et=Errmsg
      IF AllCount = 0 THEN                                        ' Issue approp. message
         scError(nMac(%eFail), modemsg + " Not found")            '
      ELSE                                                        '
         scError(0, modemsg + " triggered change processing " + FORMAT$(AllCount) + IIF$(AllCount > 1, " times", " time") + IIF$(AllCount = AllLines, "", " in " + FORMAT$(AllLines) + " lines"))
      END IF                                                      '

   ELSE                                                           ' The non-All version
      IF me.Search(0, lclCursMode) THEN                           ' Go see if we can find it
         me.Change()                                              ' Go do the change
      ELSE                                                        ' Not found, issue message
         scError(nMac(%eFail), IIF$(IsCfNext, "Bottom of data reached", "Top of data reached"))                             '
         cfNotFound = IIF(IsCfNext, 2, 1)                         ' Remember we failed
      END IF                                                      '
   END IF                                                         '
   MExit
END METHOD

METHOD  pCmdCLIP(pCmd AS STRING)
'---------- OPEN a new CLIP mode window
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   ENV.ClipCmd = IIF$(pCmdNumOps = 1, pCmdOps(1), "")             ' Setup clipboard name if provided
   ENV.PMode = %MClip OR %MClipLoad OR %MClipSave                 ' Set to CLIP mode
   sRecentAdd(ENV.CLIPPath + pCmdOps(1) + ".CLIP")                ' Add to recent path
   me.LoadTheText("", "DEFAULT")                                  ' Go load it in
   MExit
END METHOD

METHOD  pCmdCLONE(pCmd AS STRING)
'---------- CLONE (Copy) a file into a new tab
LOCAL fn1, fn2, SetProf AS STRING

   MEntry
   '----- Do basic parsing
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   ENV.PMode = %MEdit                                             ' Set to Edit mode
   SELECT CASE AS LONG pCmdNumOps                                 ' Split based on number of operands
      CASE 0                                                      ' No operands
         GOSUB Copyfn: SetProf = TIP.Profile                      ' Copy current fn
         IF fn1 = "" THEN                                         ' Not valid?
            MErrExit(%eFail, "Cannot Clone contents of this tab") '
         END IF                                                   '

      CASE 1
         IF pCmdOps(1) = "*" THEN                                 ' Get fn from the Clipboard
            fn1 = sGetFnClipboard: SetProf = sParseProfile(fn1)   '
         ELSEIF LEFT$(pCmdOps(1), 1) = "." THEN                   ' Just a Profile request?
            GOSUB CopyFn: SetProf = UUCASE(MID$(pCmdOps(1), 2))   ' Current file
         ELSE                                                     ' Else must be a filename
            fn1 = pCmdOps(1): SetProf = sParseProfile(fn1)        ' Get fn from Ops(1)
         END IF                                                   '

      CASE 2
         IF LEFT$(pCmdOps(2), 1) <> "." THEN MErrExit(%eFail, "Invalid operand format") ' Kill if not a profile
         fn1 = pCmdOps(1): SetProf = UUCASE(MID$(pCmdOps(2), 2))  ' Get fn and Profile from the operands

   END SELECT                                                     '

   '----- Validate the filename
   IF INSTR(fn1, "\") = 0 THEN fn1 = sGetDefDir + fn1             ' Add default dir if needed
   IF ISFALSE ISFILE(fn1) THEN _                                  ' See if file already exists
      MErrExit(%eFail, "Original Clone file cannot be found")     '

   '----- Get the new filename
   Fn2 = sDoInputBox("Enter the new FileName: ", "Clone Name", RTRIM$(Fn1))
   IF ISNULL(Fn2) OR IsEQ(Fn2, RTRIM$(Fn1)) THEN MErrExit(%eNone, "Clone cancelled")
   IF ISFILE(fn2) THEN _                                          ' See if file already exists
      MErrExit(%eFail, "Specified filename already exists")       '

   '----- Copy the file
   TRY
      FILECOPY fn1, fn2                                           ' COPY the file
      sRecentAdd(fn2)                                             ' Go add to RECENT list

   CATCH                                                          '
      MErrExit(%eFail, "Unable to create Cloned file (Authority?)") ' Kill if FILECOPY failed
   END TRY                                                        '

   SETATTR fn2, %NORMAL                                           ' Clear any RO bits

   '----- Open a new tab
   sTabAdd(fn2, SetProf)                                          ' Yes, let sTabAdd do the work

   MExitMeth                                                      ' Exit

   CopyFn:
   IF ISFALSE IsFMTab AND _                                       ' Not FM
      ISFALSE IsClip AND _                                        ' Not CLIP
      ISFALSE IsSetEdit AND _                                     ' Not Setedit
      TIP.FilePath <> $Empty THEN                                 ' Not (Empty)
      fn1 = TIP.FilePath                                          ' All looks well, use the current file
   ELSE                                                           ' Else null the fn
      fn1 = ""                                                    '
   END IF                                                         '
   RETURN
END METHOD

METHOD  pCmdCMD(pCmd AS STRING)
'---------- CMD - RUN an external command
LOCAL lclDrive, lclCmd, lclPath, newPath AS STRING
LOCAL RetC, i AS LONG
LOCAL lclCmd2 AS ASCIZ * 255
   MEntry
   lclCmd = TRIM$(pCommand)                                       ' Strip it clean
   sGetWord(lclCmd, %Strip, %QuoteNotSig)                         ' Strip off the command itself
   IF me.MacSubst(lclCmd) THEN _                                  ' Do substitution
      MErrExit(%eFail, "Malformed macro operand in command line")
   lclPath = CURDIR$                                              ' Save current path

   IF IsFMTab OR IsSetEdit OR IsClip THEN                         ' In FM or special Tab?
      newPath = FPath                                             ' Use the default Fm directory
   ELSE                                                           '
      newPath = IIF$(ISNULL(TIP.Path), CURDIR$, TIP.Path)         ' Setup something
   END IF                                                         '
   IF MID$(lclPath, 2, 1) = ":" THEN lclDrive = LEFT$(lclPath, 2) ' Extract Drive if present
   IF MID$(NewPath, 2, 1) = ":" THEN CHDRIVE LEFT$(NewPath, 2)    ' If NewPath has drive change to it
   CHDIR NewPath                                                  ' Switch to New path
   IF ISNULL(TRIM$(lclCmd)) THEN lclCmd = "CMD.EXE"               ' If no command, use CMD.EXE
   lclCmd2 = lclCmd                                               ' Swap to ASCIIZ string
   RetC = SHELL(ENVIRON$("COMSPEC") + " " + ENV.CmdFlags + lclCmd2)'
   IF ERR THEN                                                    ' Tell user result
      scError(%eFail, "Error issuing command " + ERROR$)          '
   ELSE                                                           '
      scError(0, "Command issued")                                '
   END IF                                                         '
   IF ISNOTNULL(lclDrive) THEN CHDRIVE lclDrive                   ' Switch drive if needed
   CHDIR lclPath                                                  ' put back the original path
   MExit
END METHOD

METHOD  pCmdCOLLATE(pCmd AS STRING)
'---------- Set COLLATE profile value
LOCAL Wrd1, t, A2S, S2A, NewName AS STRING, i AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN _                                       ' No Ops?
      MErrExit(%eNone, "SOURCE codepage set to " + Prf.PSource + ", COLLATE codepage set to " + Prf.PCollate)
   NewName = UUCASE(pCmdOps(1))                                   ' Get an uppercase version
   IF NewName <> "ANSI" AND _                                     ' Is SOURCE non-ANSI mode?
      NewName <> "UTF8" AND _                                     '
      NewName <> "UTF16" AND _                                    '
      NewName <> "UTF16LE" AND _                                  '
      NewName <> "UTF16BE" THEN                                   '
      IF ISFALSE ISFILE(ENV.INIPath + NewName + ".SOURCE") THEN _ ' See if the Custom SOURCE file exists
         MErrExit(%eFail, "COLLATE codepage file: " + ENV.INIPath + pCmdOps(1) + ".SOURCE is not found, command rejected")
   END IF

   Prf.PCollate = NewName                                         '
   sDoStatusBar($SBSource)                                        ' re-Do the StatusBar Source box
   scError(%eNone, "SOURCE codepage set to " + Prf.PSource + ", COLLATE codepage set to " + Prf.PCollate)
   MExit
END METHOD

METHOD  pCmdCOLS(pCmd AS STRING)
'---------- Turn Cols ON or OFF
LOCAL nv AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   nv = sOnOff(1)                                                 ' Go get True/False value for operand 1
   IF nv < 0 THEN MExitMeth                                       ' Error, exit, error msg already issued
   Prf.Cols = nv                                                  ' Set the new value
   scError(%eNone, "COLS set to " + IIF$(nv, "ON", "OFF"))        ' Tell user new value
   MExit
END METHOD

METHOD  pCmdCOMPRESS(pCmd AS STRING)
'---------- COMPRESS Command
LOCAL lclCmd, lclWord AS STRING
LOCAL i, j, k, l, x, y, Condense, lclAll AS LONG, AllCount, lLastLine, AllLines AS LONG, lclCursMode AS INTEGER
LOCAL fline AS LONG, ftxt, lclTxt, t AS STRING, aTxt AS WSTRING

   MEntry
   '----- Do the basic parsing stuff
   IF me.CmdParse(pCmd) THEN MExitmeth                            ' Do basic parsing, exit if errors
   lclCursMode = %CursBad                                         ' Make it look like Bad Cursor
   me.ResetFunc(%ResetFind)                                       ' Clear previous FIND hilighting

   '----- Extract the line range operands
   me.FindSave                                                    ' Save RFIND and Ranga data
   IF ISTRUE me.RangeSet("", %False, %False) THEN me.FindLoad: MExitMeth  ' Setup any range command operands

   '----- Scan out the standard search string stuff
   me.FindReset                                                   ' Reset the parse output fields
   IF sCRTParse("ASEC12PU") THEN me.FindLoad: MExitMeth           ' Scan (All, Subset, Exclude, Columns, Lit1, Lit2, Clr, User)
   me.FindSetup                                                   ' Copy CRT parse's answers
   IF ISFALSE IsCfNext AND ISFALSE IsCfPrev AND _                 '
      ISFALSE IsCfFirst AND ISFALSE IsCfLast THEN _               '
      CfSet(%CRTNext)                                             ' Set NEXT if no others specified
   IF DotSet OR _                                                 ' If a line range
      IsCfU OR IsCfNU OR IsCfX OR IsCfNX THEN                     ' Or U / NU / X / NX
      cfSet(%CRTALL)                                              ' Force ALL if no modifiers
   END IF                                                         '
   lclAll = %True                                                 ' Set the 'do at least 1 line' ALL flag

   '----- We'd better have something to work with
   IF ISFALSE IsCfLit1 THEN                                       ' Lit1 present?
      me.FindLoad                                                 ' Reload saved stuff
      MErrExit(%eFail, "Missing COMPRESS search string")          '
   END IF
   IF IsCfLit2 THEN                                               ' Lit2 present?
      IF LEN(cfChange) >= LEN(cfFind) THEN                        ' Lengths look OK?
         me.FindLoad                                              ' Reload saved stuff
         MErrExit(%eFail, "COMPRESS second literal must be shorter than first literal")
      END IF
      Condense = %True                                            ' Mark as Condense
   END IF                                                         '
   '----- OK lets do the search now
   RESET AllCount                                                 ' Reset variables

   '----- The ALL flavor search
   IF IsCfAll OR lclAll THEN                                      ' ALL?
      i = me.Search(0, lclCursMode)                               ' Do the initial search
      IF i THEN                                                   '
         i = sLine                                                ' Get found line
         GOSUB CompIt                                             ' Go compress the i line
      END IF                                                      '
      DO WHILE i                                                  ' Do till not found
         i = me.Search(1, lclCursMode)                            ' Do re-Finds
         IF i THEN                                                ' Found one
            i = sLine                                             ' Get found line
            GOSUB CompIt                                          ' Go compress the i line
         END IF                                                   '
      LOOP                                                        '

   '----- The single line search
   ELSE                                                           '
      i = me.Search(0, lclCursMode)                               ' Do the initial search
      IF i THEN                                                   '
         i = sLine                                                ' Get found line
         GOSUB CompIt                                             ' Go compress the i line
      END IF                                                      '
   END IF                                                         '

   '----- Search over, report the result
   IF AllCount = 0 THEN                                           ' Issue approp. message
      scError(nMac(%eFail), "No lines " + IIF$(Condense, "condensed", "compressed")) '
   ELSE                                                           '
      scError(0, IIF$(Condense, "Condensed ", "Compressed ") + FORMAT$(AllCount) + IIF$(AllCount > 1, " lines", " line"))
      me.CurSetReq(%Change, fline, 0, %True)                      ' Do it without hilite
   END IF                                                         '
   me.FindLoad                                                    ' Reload saved stuff
   MExitMeth

'----- Do the compress for one line
CompIt:
   IF Condense THEN GOSUB CondIt: RETURN                          ' Split off Condense here
   j = cfFCol                                                     ' Column where it was found
   k = cfFLen                                                     ' Length of found string (Use orig. length)
   l = LEN(L(i).@LTxt)                                            ' Length of text line
   lclTxt = LTxtG(i)                                              ' Get text line
   atxt = LAttrG(i)                                               ' Get Attr line
   ftxt = MID$(lclTxt, j, k)                                      ' Get found text
   x = j                                                          ' Where found
   DO                                                             ' Do forever
     IF x + k + k - 1 > LEN(lclTxt) THEN EXIT DO                  ' No more possibility of compressing, exit loop
     IF x + k + k - 1 > cfRMargin THEN EXIT DO                    '
     IF MID$(lclTxt, x + k, k) = ftxt THEN                        ' Two in a row?
        lclTxt = LEFT$(lclTxt, x + k - 1) + MID$(lclTxt, x + k + k) ' Compress
        atxt = LEFT$(atxt, x + k - 1) + MID$(atxt, x + k + k)     ' Compress Attr line in step
        IF i <> lLastLine THEN                                    ' Count actual changed lines
           INCR AllCount: lLastLine = i                           ' Count and save line #
        END IF                                                    '
        me.ModSet(i)                                              ' Remember we changed something
        LFOn(i, %EQChange)                                        ' Mark the line number
        me.UpdLControl(i)                                         ' Update LLCtl
        IF fline = 0 THEN fline = i                               ' Save 1st found line
     ELSE                                                         ' Not two in a row
        x = INSTR(x + k, lclTxt, ftxt)                            ' See if another string further along
        IF x = 0 THEN EXIT DO                                     ' No?  We're done
        IF x + k + k - 1 > cfRMargin THEN EXIT DO                 ' Yes, Within margins? No? We're done as well
     END IF                                                       '
   LOOP                                                           '

   me.LTxtSet(i, lclTxt)                                          ' Save it back
   LAttrS(i) = atxt                                               ' Save back Attr
   me.AttrScan(i)                                                 ' Recolorize
   RETURN

'----- Condense the line
CondIt:
   OnColSupp                                                      ' Tell sChange not to maintain columns
   IF cfFLine <> lLastLine THEN                                   '
      IF lLastLine <> 0 AND ISFALSE IsCfAll AND ISTRUE lclAll THEN' Just a lclALL type and we've done one line?
         i = 0: RETURN                                            ' Kill the searc loop and bail out
      END IF                                                      '
      INCR AllCount: lLastLine = cfFLine                          '
   END IF                                                         '
   IF AllCount = 1 THEN me.CurSetReq(%Position, cfFLine, 1, %True)
   t = LTxtG(cfFLine)                                             '
   me.Change()                                                    ' Go do the change
   sCol = 1                                                       ' Start at column 1 again
   sLine = sLine - 1                                              ' Backup sChanges restart point
   IF ISFALSE IsCfAll AND ISTRUE lclAll THEN                      ' Just a lclALL type?
   END IF
   RETURN
END METHOD

METHOD  pCmdCOPY(pCmd AS STRING)
'---------- Read file and insert after current row pointer
LOCAL i, j, x, ii, iLastLine, iCount, iPoint, NumAdj, NumInDest, pLastLine, SetCur AS LONG
LOCAL frmlin, tolin, before, after, ablabel, lclTop AS LONG
LOCAL fn, lclWord AS STRING

   MEntry
   '----- Do the basic command parsing
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors

   '----- Loop through the scanned operands
   FOR i = 1 TO pCmdNumOps                                        '
      IF pCmdOpsType(i) = %OpNum THEN                             ' Numeric?
         IF frmlin = 0 THEN                                       '
            frmlin = VAL(pCmdOps(i))                              ' Grab it as FrmLin
         ELSEIF tolin = 0 THEN                                    '
            tolin = VAL(pCmdOps(i))                               ' Grab it as ToLin
         ELSE                                                     '
            MErrExit(%eFail, "Too many numeric operands")         ' Issue error
         END IF                                                   '
         GOSUB DelOne                                             ' Remove from parse
      ELSEIF IsEQ(pCmdOps(i), "BEFORE") THEN                      ' Handle BEFORE
         before = %True                                           ' Remember it
         GOSUB DelOne                                             ' Remove from parse
      ELSEIF IsEQ(pCmdOps(i), "AFTER") THEN                       ' Handle AFTER
         after = %True                                            ' Remember it
         GOSUB DelOne                                             ' Remove from parse
      ELSEIF pCmdOpsType(i) = %OpDotd THEN                        ' A label?
         IF ablabel > 0 THEN MErrExit(%eFail, "Multiple line reference detected - " + pCmdOps(i))
         ablabel = me.LineNoRef(pCmdOps(i))                       ' Get the line number
         IF ablabel = -1 THEN MErrExit(%eFail, "Invalid line reference - " + pCmdOps(i)) ' Oops!
         GOSUB DelOne                                             ' Remove from parse
      ELSEIF pCmdOpsType(i) = %OpLPtr THEN                        ' A LPtr reference?
         IF TRIM$(pCmdOps(i)) = "!" THEN MErrExit(%eFail, "Line reference: " + pCmdOps(i) + " is invalid") ' Oops!
         IF ablabel > 0 THEN MErrExit(%eFail, "Multiple line reference detected - " + pCmdOps(i)) ' Oops!
         ablabel = VAL(MID$(pCmdOps(i), 2))                       ' Get the line number
         IF ablabel = 0 OR ablabel > LastLine THEN MErrExit(%eFail, "Line reference: " + pCmdOps(i) + " is invalid") ' Oops!
         GOSUB DelOne                                             ' Remove from parse
      END IF                                                      '
   NEXT i                                                         '

   '----- Validate what we scanned
   IF before AND after THEN MErrExit(%eFail, "Can't specify both BEFORE and AFTER")     ' Issue error
   IF (before OR after) AND ablabel = 0 THEN MErrExit(%eFail, "BEFORE/AFTER require a line label")       ' Issue error
   IF frmlin <> 0 AND tolin = 0 THEN MErrExit(%eFail, "Both start and end line numbers of copied file are required") ' Issue error
   IF frmlin > tolin THEN SWAP frmlin, tolin                      '

   '----- OK, a bit more validation and set the starting line for Insert
   IF LastLine <> 2 THEN                                          ' If we 're not empty
      IF before > 0 THEN                                          ' Stuff in A/B from cmd line
         i = ablabel - 1                                          '
         IF ISTRUE LTblRange THEN MErrExit(%eCmdPend, "Cannot specify BEFORE and also an A/B, AA/BB or C/CC line control")
      ELSEIF after > 0 THEN                                       '
         i = ablabel                                              '
         IF ISTRUE LTblRange THEN MErrExit(%eCmdPend, "Cannot specify AFTER and also an A/B, AA/BB or C/CC line control")
      ELSE                                                        '
         IF ISFALSE LTblRange THEN                                ' Better have a line range at this point
            MErrExit(%eCmdPend, "Specify A/B, AA/BB or H/HH line control, or BEFORE/AFTER for COPY")
         ELSE                                                     ' We have a line range
            IF LTblDCmd <> "H       " AND LTblDCmd <> "HH      " AND _ ' Only H/HH
               LTblDCmd <> "A       " AND LTblDCmd <> "B       " AND _ ' Or A/B
               LTblDCmd <> "AA      " AND LTblDCmd <> "BB      " THEN  ' or AA/BB
               MErrExit(%eCmdPend, "Specify A/B, AA/BB or H/HH line control, or BEFORE/AFTER for COPY")
            END IF                                                '
         END IF                                                   '
      END IF                                                      '
   ELSE                                                           '
      i = 1                                                       ' Else insert after top line
   END IF                                                         '
   me.RangeSet("RESET", 0, 0)                                     ' Clear any A/B
   lclTop = TopScrn                                               ' Save topscreen

   '----- Get the file to be copied
   IF pCmdNumOps > 1 THEN MErrExit(%eFail, "Extraneous operand detected - " + pCmdOps(2)) ' Oops!

   IF pCmdNumOps = 1 THEN                                         ' Just one operand left?
      IF pCmdOps(1) = "*" THEN                                    ' If 1st operand is a '*" then
         pCmdOps(1) = sGetFnClipboard                             ' Replace it with Clipboard contents
         pCmdOpsType(1) = %OpStr                                  ' Say it's a simple string
      END IF                                                      '
   END IF                                                         '

   errFlag = %eNone: ErrMsg = ""                                  ' Reset any Pending conflict flag

   fn = IIF$(pCmdNumOps = 0, "", pCmdOps(1))                      ' Get filename operand if present

   '----- Validate the filename
   TOP.Setup("PE","", "Select file to be Copied", fn)             ' Setup to test
   IF TOP.EXEC THEN                                               ' Go validate
      scError(%eFail, TOP.ResultMsg)                              ' Oops?  Bail out
      IF LTblDCmd = "A       " OR LTblDCmd = "B       " OR Before OR After THEN  ' Normal A / B ?
         IF ISFALSE Before AND ISFALSE After THEN                 '
            me.TTblDel(LTblDFrom)                                 ' Remove from Touched lines
            me.UpdLControl(LTblDFrom)                             ' Clear from the line
         END IF                                                   '
      ELSE                                                        ' Special kind then
         me.TTblDel(LTblDfrom)                                    ' Remove from Touched lines
         me.UpdLControl(LTblDfrom)                                ' Clear from the line
         me.TTblDel(LTblDTo)                                      ' Remove from Touched lines
         me.UpdLControl(LTblDTo)                                  ' Clear from the line
      END IF                                                      '
      MExitMeth                                                   ' Bail out
   END IF

   '----- OK, finally start doing something
   aPrf.SetProfName(TOP.Profile, %True)                           ' Get aPrf set to the file being copied
   aPrf.ReadAll(%True)                                            '

   '----- If the simple insert, do it
   IF i = 1 AND LTblDCmd <> "A       " AND LTblDCmd <> "B       " AND ISFALSE Before AND ISFALSE After THEN  ' Empty default?
      GOSUB FileInsert                                            ' Insert the file
      GOSUB FinishCopy                                            ' Finish off things

   ELSEIF LTblDCmd = "A       " OR LTblDCmd = "B       " OR Before OR After THEN  ' Normal A / B ?
      IF ISFALSE Before AND ISFALSE After THEN                    '
         i = IIF(LTblDCmd = "A       ", LTblDFrom, LTblDFrom - 1) ' Set Insert point
         me.TTblDel(LTblDFrom)                                    ' Remove from Touched lines
         me.UpdLControl(LTblDFrom)                                ' Clear from the line
      END IF                                                      '
      GOSUB FileInsert                                            ' Insert the file
      GOSUB FinishCopy                                            ' Finish off things

   '----- Insert in a replacement block
   ELSEIF LTblDCmd = "H       " OR LTblDCmd = "HH      " THEN     ' The H / HH ?

      '----- Remove the HH block range
      me.LStubDelRange(LTblDfrom, LTblDTo)                        ' Delete the range
      me.TTblDel(LTblDfrom)                                       ' Remove from Touched lines
      me.UpdLControl(LTblDfrom)                                   ' Clear from the line
      me.TTblDel(LTblDTo)                                         ' Remove from Touched lines
      me.UpdLControl(LTblDTo)                                     ' Clear from the line
      i = LTblDfrom - 1                                           ' Make previous line look like 'A'

      '----- And Insert thr file in it's place
      GOSUB FileInsert                                            ' Insert the file
      GOSUB FinishCopy                                            ' Finish off things

   '----- Insert repeatedly in an AA block
   ELSEIF LTblDCmd = "AA      " THEN                              ' The AA After mode?
      LTblDRpt = MAX(1, LTblDRpt)                                 ' Ensure at least count of 1
      me.TTblDel(LTblDfrom)                                       ' Remove from Touched lines
      me.UpdLControl(LTblDfrom)                                   ' Clear from the line
      me.TTblDel(LTblDTo)                                         ' Remove from Touched lines
      me.UpdLControl(LTblDTo)                                     ' Clear from the line

      '----- Determine size of block range
      FOR x = LTblDfrom TO LTblDTo                                ' Count the data lines
         IF IsLData(x) OR IsLTop(x) OR IsLFile(x) THEN            ' Just data lines
            INCR NumInDest                                        ' + 1
            GOSUB DoPM                                            '
         END IF                                                   '
      NEXT x                                                      '
      IF NumInDest <= LTblDRpt THEN MErrExit(%eFail, "AA number must be < range size") ' Dest count reasonable?

      '----- Setup and do the loop-de-loop
      i = LTblDfrom + LTblDRpt - 1                                ' Init for loop
      SetCur = LTblDfrom                                          ' Remember where cursor goes
      WHILE i <= LTblDTo + NumAdj                                 ' Big outer loop
         IF ISFALSE IsLData(i) AND ISFALSE IsLTop(i) AND ISFALSE IsLFile(i) THEN INCR i: ITERATE ' Ignore non Data lines
         IF SetCur > 0 THEN                                       ' Set cursor just once
            me.CurSetReq(%LineCmd, SetCur, 1, %False)             '
            SetCur = 0                                            '
         END IF                                                   '
         pLastLine = LastLine                                     ' Save so we can calc # lines inserted
         GOSUB FileInsert                                         ' Insert the file after 'i'
         NumAdj += (LastLine - pLastLine)                         ' Track # inserted
         i += LTblDRpt + (LastLine - pLastLine)                   ' Big outer loop
      LOOP                                                        ' next 'i'
      GOSUB FinishCopy                                            ' Finish off things

   '----- Insert repeatedly in an BB block
   ELSEIF LTblDCmd = "BB      " THEN                              ' The BB After mode?
      LTblDRpt = MAX(1, LTblDRpt)                                 ' Ensure at least count of 1
      me.TTblDel(LTblDfrom)                                       ' Remove from Touched lines
      me.UpdLControl(LTblDfrom)                                   ' Clear from the line
      me.TTblDel(LTblDTo)                                         ' Remove from Touched lines
      me.UpdLControl(LTblDTo)                                     ' Clear from the line

      '----- Determine size of block range
      FOR x = LTblDfrom TO LTblDTo                                ' Count the data lines
         IF IsLData(x) OR IsLBottom(x) OR IsLFile(x) THEN         ' Just data lines
            INCR NumInDest                                        ' + 1
            GOSUB DoPM                                            '
         END IF                                                   '
      NEXT x                                                      '
      IF NumInDest <= LTblDRpt THEN MErrExit(%eFail, "BB number must be < range size") ' Dest count reasonable?

      '----- Setup and do the loop-de-loop
      x = NumInDest \ LTblDRpt                                    ' x = # of insertion points
      i = LTblDTo - (x * LTblDRpt)                                ' Initial insertion point
      SetCur = LTblDfrom                                          ' Remember where cursor goes
      WHILE i < LTblDTo + NumAdj                                  ' Big outer loop
         IF ISFALSE IsLData(i) AND ISFALSE IsLTop(i) AND ISFALSE IsLFile(i) THEN
         INCR i: ITERATE                                          ' Ignore non Data lines
         END IF                                                   '
         IF SetCur > 0 THEN                                       ' Set cursor just once
            me.CurSetReq(%LineCmd, SetCur, 1, %False)             '
            SetCur = 0                                            '
         END IF                                                   '
         pLastLine = LastLine                                     ' Save so we can calc # lines inserted
         GOSUB FileInsert                                         ' Insert the file after 'i'
         NumAdj += (LastLine - pLastLine)                         ' Track # inserted
         i += LTblDRpt + (LastLine - pLastLine)                   ' Big outer loop
      LOOP                                                        ' next 'i'
      GOSUB FinishCopy                                            ' Finish off things

   END IF                                                         '
   TopScrn = lclTop                                               ' Restore topscreen
   MExitMeth

'----- Remove an entry from the parse array
DelOne:
   ARRAY DELETE pCmdOps(i): ARRAY DELETE pCmdRaw(i): ARRAY DELETE pCmdOpsType(i)
   DECR i: DECR pCmdNumOps                                        ' So we stay on the same index number
   RETURN

'----- Insert the file at the desired point
FileInsert:
   iPoint = i + 1                                                 ' Save location of 1st inserted line
   iLastLine = LastLine                                           ' Save so we can calc # lines inserted
   me.CopyAFile(i, TOP, APrf, LTblDFlag, frmlin, tolin, %False)   ' Go load the data
   me.ModSet(i)                                                   ' Remember we changed something
   IF errFlag <> %eNone THEN MExitMeth                            ' If errors, bail out
   iCount = (LastLine - iLastLine)                                ' Track # inserted
   FOR ii = iPoint TO iPoint + iCount - 1                         ' Loop marking the CmdIX
      LPCmdS(ii) = PCmdIX                                         ' Stuff in PCmdIX to make it :ZF
   NEXT ii                                                        '
   RETURN                                                         '

'----- Issue a done message
FinishCopy:
   scError(%eNone, "File copied")                                 ' Issue success message
   RETURN

'----- Handle the +/- involvement
DoPM:
   IF BIT(LTblDFlag, %lCmdX) THEN                                 ' Do the +/- processing for the Source
      LFOn(x, %Invisible)                                         ' If - make Invisible
      gfXRebuild = %True                                          ' Ask for Exclude processing
   ELSEIF BIT(LTblDFlag, %lCmdNX) THEN                            '
      LFOff(x, %Invisible)                                        ' If + make visible
      gfXRebuild = %True                                          ' Ask for Exclude processing
   END IF                                                         '
   RETURN

END METHOD

METHOD  pCmdCREATE(pCmd AS STRING)
'---------- CREATE a new filename
LOCAL lclCmd, lclWord, fn AS STRING, x, i, lclTop AS LONG

   MEntry
   '----- Do the basic parsing stuff
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   lclTop = TopScrn                                               ' Save where we are

   '----- If sGetFileAndRange is successful, we have something to do
   IF me.GetFileAndRange() THEN                                   ' Go get filename etc

      '----- Do a bit of validation
      IF ISFALSE IsCfALL THEN MErrExit(%eFail, "Pending line range") ' Some kind of line criteria?
      IF ISNOTNULL(PATHSCAN$(FULL, TOP.FilePath)) THEN MErrExit(%eFail, "File exists, use REPLACE to re-use it")

      '----- Write the file now
      TOP.Setup("OCR", "", "Select file for CREATE", TOP.FilePath)'
      IF me.WriteFile(%False) THEN                                '

         '----- If this is still an (Empty) tab make it a normal one now
         IF TIP.FilePath = $Empty THEN                            ' If still no current file
            TIP.Setup("*", TOP.Profile, "", TOP.FilePath)         ' Swap output filename into the current working file
            me.WindowTitle                                        ' Alter window/Tab titles
            Prf.SetProfName(TIP.Profile,%False)                   ' Build the File Profile INI name
            Prf.ReadAll(%True)                                    ' Go read this file type's INI values
            me.ClrLoad()                                          ' Get Attr loaded if it exists
            OffModdFlag: OnUndoFlag                               ' Turn off the Modd flag
            me.FileWatch("", %WatchEnd)                           ' Kill any prior Watch
            sFileQueue("A", " ", TIP.FilePath)                    ' Add to FQ
            IF me.FileWatch(TIP.FilePath, %WatchStart) THEN _     ' Establish the watch
               scError(0, "File watch could not be established")  '
         END IF                                                   '

         '----- If the line range was a MM (Move) then delete the range
         IF SavFilLCtl = "M" OR SavFilLCtl = "MM" THEN            ' If the line selection was a Take

            '----- But only if we're OK
            IF FileRecs = 0 THEN                                  '
               scError(%eFail, "No records written, File is empty, Record deletion skipped") ' Issue error
            ELSE                                                  '
               me.ModSet(DotStart)                                ' Remember we changed something
               me.LStubDelRange(DotStart, DotEnd)                 ' Delete the range
               me.AdjustPending(DotStart, DotEnd - DotStart + 1, 0) ' Adjust pending requests
            END IF                                                '
         END IF                                                   '

         '----- Make sure we're in the RECENT list
         sRecentAdd(TOP.FilePath)                                 ' Go add to RECENT list
         DoSet(%LoadReq)                                          ' Request refresh

      END IF                                                      '
   END IF                                                         '
   TopScrn = lclTop
   MExit
END METHOD

METHOD  pCmdCUT(pCmd AS STRING)
'---------- CUT data to the Windows Clipboard
LOCAL lclCmd, lclWord, CBD, cbname, MyDLM AS STRING, lTxtp AS STRING POINTER
LOCAL i, j, k AS LONG, lclCursMode AS INTEGER, fl AS QUAD
LOCAL NumDel, lclTop, AppendText, recs, Raw, DelLines AS LONG
DIM   DLines(1000) AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors

   lclCursMode = %CursBad                                         ' Pretend a bad cursor location
   me.FindSave                                                    ' Save any current RFIND  and Range data
   IF ISTRUE me.RangeSet("", %True, %False) THEN me.FindLoad: MExitMeth  ' Setup any range command operands
   IF pCmdNumOps > 0 THEN                                         ' If still more operands left
      FOR i = 1 TO pCmdNumOps                                     '
         IF pCmdOpsType(i) = %KWAPPEND OR _                       ' Handle APPEND
            pCmdOpsType(i) = %KWA THEN                            '
            AppendText = %True                                    '
            ARRAY DELETE pCmdOps(i): ARRAY DELETE pCmdRaw(i): ARRAY DELETE pCmdOpsType(i)
            DECR i: DECR pCmdNumOps                               ' So we stay on the same index number
         ELSEIF pCmdOpsType(i) = %KWREPLACE THEN                  ' Handle REPLACE noise word
            ARRAY DELETE pCmdOps(i): ARRAY DELETE pCmdRaw(i): ARRAY DELETE pCmdOpsType(i)
            DECR i: DECR pCmdNumOps                               ' So we stay on the same index number
         ELSEIF pCmdOpsType(i) = %KWDEL THEN                      ' Handle DELETE noise word
            DelLines = %True                                      ' Remember DELETE asked for
            ARRAY DELETE pCmdOps(i): ARRAY DELETE pCmdRaw(i): ARRAY DELETE pCmdOpsType(i)
            DECR i: DECR pCmdNumOps                               ' So we stay on the same index number
         ELSEIF pCmdOpsType(i) = %KWRAW THEN                      ' Handle RAW keyword
            Raw = %True                                           ' Remember this
            ARRAY DELETE pCmdOps(i): ARRAY DELETE pCmdRaw(i): ARRAY DELETE pCmdOpsType(i)
            DECR i: DECR pCmdNumOps                               ' So we stay on the same index number
         ELSEIF pCmdOpsType(i) = %OpStr THEN                      ' Simple string?
            IF INSTR($Numeric, LEFT$(pCmdOps(i) , 1)) = 0 THEN    ' Not a leading numeric?
               cbname = pCmdOps(i)                                ' Save clipboard name
               ARRAY DELETE pCmdOps(i): ARRAY DELETE pCmdRaw(i): ARRAY DELETE pCmdOpsType(i)
               DECR i: DECR pCmdNumOps                            ' So we stay on the same index number
            END IF                                                '
         END IF                                                   '
      NEXT i                                                      '
   END IF                                                         '
   IF pCmdNumOps = 0 AND ISFALSE DotSet THEN MErrExit(%ePending, "Pending line range") ' Better be something
   me.FindReset                                                   ' Reset the parse output fields
   IF sCRTParse("ASDU") THEN me.FindLoad: MExitMeth               ' Scan (All, Subset, Direction, User)
   me.FindSetup                                                   ' Copy CRT parse's answers
   CfSet(%CRTDX)                                                  ' Tell search hands off Invisible status
   fl = cfFlag
   IF DotSet OR _                                                 ' If a line range
      IsCfU OR IsCfNU OR IsCfX OR IsCfNX THEN                     ' Or U / NU / X / NX
      IF ISFALSE IsCfFirst AND ISFALSE IsCfLast AND _             ' And no modifiers
         ISFALSE IsCfNext  AND ISFALSE IsCfPrev THEN              '
         CfSet(%CRTALL)                                           ' Force ALL if no modifiers
      END IF                                                      '
   END IF                                                         '

   gKeyPrimOper = cbname: MyDLM = $CRLF                           ' Set clipboardname / delimiter
   errFlag = %eNone                                               ' Reset any Pending conflict flag
   lclTop = TopScrn                                               '
   CBD = ""                                                       ' Start as ""
   IF AppendText THEN sReadClipboard(CBD, MyDLM, %False)          ' Go get whatever's there, keep the file

   IF IsCfAll THEN                                                ' The ALL version?
      i = me.Search(0, lclCursMode)                               ' Do the initial search
      IF i THEN                                                   '
         GOSUB CutOne                                             ' Go handle one
      END IF                                                      '
      DO WHILE i                                                  ' Do till not found
         i = me.Search(1, lclCursMode)                            ' Do re-Finds
         IF i THEN                                                ' Found one
            GOSUB CutOne                                          ' Go handle one
         END IF                                                   '
      LOOP                                                        '
   ELSE                                                           ' The non-All version
      i = me.Search(0, lclCursMode)                               ' Go see if we can find it
      IF i THEN                                                   ' Got it?
         GOSUB CutOne                                             ' Go handle one
      END IF                                                      '
   END IF

   IF ISTRUE sWriteClipboard(CBD) THEN                            ' OK?
      IF j THEN                                                   ' Something to delete?
         FOR i = j - 1 TO 0 STEP -1                               ' Now delete the lines (backward)
            IF IsLData(dLines(i)) THEN me.ModSet(dLines(i))       ' Remember we changed something
            me.LTxtFree(dLines(i))                                ' Go free the dynamic string
            me.LEntDel(dLines(i))                                 ' Remove from the L() array
            INCR NumDel                                           '
            me.AdjustPending(dLines(i), -1, 0)                    ' Adjust pending requests
            DECR LastLine: DECR LastReal                          ' Adjust LastLine and LastReal
         NEXT i                                                   '
      END IF                                                      '
      TopScrn = lclTop                                            '
      IF recs > 0 THEN
         IF AppendText THEN                                       ' Issue appropriate message
            scError(0, FORMAT$(recs) + IIF$(recs > 1, " lines", " line") + " appended to " + IIF$(ISNULL(gKeyPrimOper), "Windows", gKeyPrimOper) + " clipboard")
         ELSE                                                     '
            scError(0, FORMAT$(recs) + IIF$(recs > 1, " lines", " line") + " cut to " + IIF$(ISNULL(gKeyPrimOper), "Windows", gKeyPrimOper) + " clipboard")
         END IF                                                   '
      ELSE                                                        '
         scError(%eFail, "Cut to clipboard failed")               '
      END IF                                                      '
   END IF                                                         '
   RESET gKeyPrimOper                                             ' Blank clipboardname
   me.FindLoad                                                    ' Reload RFIND data
   MExitMeth

CutOne:
   i = cfFLine                                                    ' Get the located line number
   IF IsLInsertLine(i) THEN RETURN                                ' Exempt Insert lines
   lTxtp = LTxtGP(i)                                              ' Get a pointer to the text
   IF ISFALSE Raw THEN                                            ' Normal mode
      CBD += @lTxtP + MyDLM                                       ' Add each line with delimiter
   ELSE                                                           ' Raw mode
      CBD += @lTxtP                                               ' Add each line without a delimiter
   END IF                                                         '
   INCR recs                                                      ' Count records
   IF LEFT$(DotMode, 1) = "M" OR DelLines THEN                    ' A TAKE like reference, or DELETE asked for?
      IF j + 1 > UBOUND(dlines) THEN                              ' Do we need to expand table?
         REDIM PRESERVE dLines(UBOUND(dLines) + 1000)             ' Yes, do it
      END IF                                                      '
      dLines(j) = i                                               ' Save the line number
      INCR j                                                      ' Count it
   END IF                                                         '
   RETURN                                                         '
END METHOD

METHOD  pCmdDELETE(pCmd AS STRING)
'---------- DELETE Specified lines
LOCAL Lasttxt, NoteOp AS STRING
LOCAL i, j, k, AllCount, lclCol, lDir, lsLine AS LONG, lclCursMode AS INTEGER
DIM dlines(1 TO 1000) AS LONG                                     ' To hold lines to be deleted
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   me.ResetFunc(%ResetFind)                                       ' Clear previous FIND hilighting
   IF ISTRUE me.RangeSet("", %True, %False) THEN MExitMeth        ' Setup any range command operands

   '----- Handle DUP here
   IF IsEQ(pCmdOps(1), "DUP") AND _                               ' DUP
      pCmdOpsType(1) <> %OpQStr THEN                              ' and it's not quoted
      IF pCmdNumOps <> 1 THEN MErrExit(%eFail, "DUP invalid operand detected") ' Eliminate > 1 operand

      FOR i = 1 TO LastLine                                       ' Lets scan things
         IF ISFALSE me.RangeVal(i) THEN ITERATE                   ' Honour range

         IF IsLData(i) THEN                                       ' Only data lines
            IF LTxtG(i) = Lasttxt THEN                            ' A DUP?
               GOSUB AddDel                                       ' Mark for deletion
            ELSE                                                  '
               Lasttxt = LTxtG(i)                                 ' Save as new DUP text data
            END IF                                                '
         END IF                                                   '
      NEXT i                                                      '

      '----- If we found any, delete them
      IF j THEN                                                   ' Any added to the Delete array?
         GOSUB DelThem                                            ' Go Delete them
         scError(0, FORMAT$(AllCount) + " duplicate " + IIF$(AllCount > 1, "lines", "line") + " removed")
      ELSE                                                        '
         scError(0, "No duplicate lines found")                   '
      END IF                                                      '
      MExitMeth                                                   '
   END IF                                                         '

   '----- Handle NOTE type here
   IF (IsEQ(pCmdOps(1), "NOTE") OR _                              ' NOTE or
       IsEQ(MID$(pCmdOps(1), 2), "NOTE")) AND _                   ' xNOTE
      pCmdOpsType(1) <> %OpQStr THEN                              ' and it's not quoted
      IF pCmdNumOps <> 1 THEN MErrExit(%eFail, "NOTE invalid operand detected") ' Eliminate > 1 operand

      IF VERIFY(LEFT$(UUCASE(pCmdOps(1)), 1), $Upper) <> 0 THEN MErrExit(%eFail, pCmdOps(1) + " is invalid")

      NoteOp = UUCASE(pCmdOps(1))                                 ' Save exact xNOTE request

      FOR i = 1 TO LastLine                                       ' Lets scan things
         IF ISFALSE me.RangeVal(i) THEN ITERATE                   ' Honour range
         IF IsLNote(i) THEN                                       ' Is this even a NOTE line?
            IF NoteOp = "NOTE" AND L(i).LWrk2 = 0 THEN            ' The simple NOTE type?
               GOSUB AddDel                                       ' Mark for deletion
            ELSEIF LEFT$(NoteOp, 1)  = "Z" AND L(i).LWrk2 <> 0 THEN ' Generic special note (Z) and this is a special?
               GOSUB AddDel                                       ' Mark for deletion
            ELSEIF L(i).LWrk2 = INSTR($Upper, LEFT$(NoteOp, 1)) THEN ' Specific NOTE type and a match?
               GOSUB AddDel                                       ' Mark for deletion
            END IF                                                '
         END IF                                                   '
      NEXT i                                                      '

      '----- If we found any, delete them
      IF j THEN                                                   ' Any added to the Delete array?
         GOSUB DelThem                                            ' Go Delete them
         scError(0, FORMAT$(AllCount) + " NOTE " + IIF$(AllCount > 1, "lines", "line") + " removed")
      ELSE                                                        '
         scError(0, "No NOTE lines found")                        '
      END IF                                                      '
      MExitMeth                                                   '
   END IF                                                         '

   '----- The 'normal' type delete
   '----- Parse the operands
   me.CmdParse(pCmd)                                              ' Reset the parse stuff
   lclCursMode = Zone                                             ' Save CursMode at start of command
   me.FindReset                                                   ' Reset the parse output fields
   IF ISTRUE me.RangeSet("", %False, %False) THEN MExitMeth       ' Setup any range command operands
   IF sCRTParse("ASDMC1PU") THEN MExitMeth                        ' Scan (All, Subset, Direct, Modifier, Cols, Lit1, Clr, User)
   me.FindSetup                                                   ' Copy CRT parse's answers
   IF ISFALSE IsCfNext AND ISFALSE IsCfPrev AND _                 '
      ISFALSE IsCfFirst AND ISFALSE IsCfLast THEN                 '
      CfSet(%CRTNext)                                             ' Set NEXT if no others specified
      IF DotSet THEN CfSet(%CRTAll)                               ' If a range from sRangeSet then set ALL
   END IF                                                         '

   CfSet(%CRTDX)                                                  ' Add the DX internally

   '----- Validate any weird combinations of operands
   IF ISFALSE IsCfLit1 THEN                                       ' No Lit1 present?
      IF ISFALSE IsCfAll AND ISFALSE DotSet AND _                 ' No Lit1, better have ALL or DotSet
         ISFALSE IsCfU AND ISFALSE IsCfNU AND _                   ' or U or NU
         ISFALSE IsCfX AND ISFALSE IsCfNX THEN _                  ' or X or NX
         MErrExit(%eFail, "No search string, but no ALL specified")
      IF IsCfWord OR IsCfPrefix OR IsCfSuffix THEN MErrExit(%eFail, "No search string, but string options are present") ' No extraneous operands
   ELSE
      IF ISFALSE IsCfAll THEN LocFind = "DELETE"                  ' If not ALL, remember which LOC / FIND command was last
   END IF                                                         '

   IF IsCfAll AND ISFALSE IsCfLit1 AND _                          ' Ending up with just ALL?
      ISFALSE DotSet AND ISFALSE IsCfX AND ISFALSE IsCfNX AND _   ' Nothing else?
      ISFALSE IsCfU AND ISFALSE IsCfNU THEN                       '
      IF IsNE(pCmdOps(0), "DELETE") THEN MErrExit(%eFail, "To delete entire file, fully spell out DELETE ALL")
   END IF                                                         '

   '----- Run the standard search procedure
   IF IsCfAll THEN                                                ' The ALL version?
      i = me.Search(0, lclCursMode)                               ' Do Find
      lclCursMode = %CursBad                                      ' Make it look like Bad Cursor
      IF i THEN                                                   '
         i = cfFLine                                              ' Get found line
         lclCol = sCol: lDir = sDir: lsLine = sLine               ' Save 1st located for RFIND stuff
         GOSUB AddDel                                             ' If found, Flag for delete
      END IF                                                      '
      DO WHILE i                                                  ' Do till not found
         i = me.Search(1, lclCursMode)                            ' Do re-Finds
         IF i THEN                                                ' Found one
            i = cfFLine                                           ' Get found line
            GOSUB AddDel                                          ' If found, Flag for delete
         END IF                                                   '
      LOOP                                                        '
      GOSUB DelThem                                               ' Go do the deletes
      IF AllCount = 0 THEN                                        ' Issue approp. message
         scError(nMac(%eFail), "No lines found")                  '
      ELSE                                                        '
         scError(0, "Deleted " + FORMAT$(AllCount) + IIF$(AllCount > 1, " lines", " line"))
         sCol = lclCol: sDir = lDir: sLine = lsLine               ' Restore to 1st one found
      END IF                                                      '

   ELSE                                                           ' The non-All version
      i = me.Search(0, lclCursMode)                               ' Do Find
      IF i = 0 THEN                                               ' Got it?
         scError(nMac(%eFail), IIF$(IsCfNext, "Bottom of data reached", "Top of data reached"))
         cfNotFound = IIF(IsCfNext, 2, 1)                         ' Remember we failed
      ELSE                                                        '
         i = sLine                                                ' Get found line
         GOSUB AddDel                                             ' If found, Flag for delete
         GOSUB DelThem                                            '
         scError(0, "Deleted " + FORMAT$(AllCount) + " line")     '
         ErrFlag = %eNone                                         ' Say we 're OK
      END IF                                                      '
   END IF                                                         '
   gfXRebuild = %True                                             ' Force exclude rebuild
   MExitMeth

AddDel:
   FOR k = 1 TO j                                                 '
      IF dLines(k) = i THEN RETURN                                ' Don't add line more than once
   NEXT k                                                         '
   INCR j
   IF J > UBOUND(dLines) THEN _                                   ' Expand table if needed
      REDIM PRESERVE dLines(1 TO UBOUND(dLines) * 2)              ' Double it
   dLines(j) = i                                                  ' Save the line number
   RETURN                                                         '

DelThem:
   Allcount = 0
   IF j = 0 THEN RETURN                                           ' Nothing to do
   FOR i = j TO 1 STEP -1                                         ' Now delete the lines (backward)
      me.ModSet(dLines(i))                                        ' Remember we changed something
      me.LTxtFree(dLines(i))                                      ' Go free the dynamic string
      me.LEntDel(dLines(i))                                       ' Remove from the L() array
      me.AdjustPending(dLines(i), -1, 0)                          ' Adjust pending requests
      DECR LastLine: DECR LastReal                                ' Adjust LastLine and LastReal
      INCR AllCount                                               ' Count real deletions
   NEXT i                                                         '
   RETURN

END METHOD

METHOD  pCmdDIR(pCmd AS STRING)
'---------- Open FM to current file's DIR
LOCAL lclPath, lclFn AS STRING, i AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF IsSetEdit OR IsClip THEN MErrExit(%eFail, "This tab has no implied DIR folder") ' Ineligible tab?
   lclPath = TIP.Path                                             ' Get the path

   IF IsMEdit THEN                                                ' For MEdit, make sure all paths are the same
      FOR i = 1 TO MEditCount                                     '
         lclFn = me.MEditListGet(i)                               ' Get filename from MEdit list
         IF lclpath <> LEFT$(lclFn, INSTR(-1, lclFn, "\") - 1) THEN MErrExit(%eFail, "All MEdit files are not from the same DIR folder")
      NEXT i                                                      '
   END IF                                                         '

   FMask = IIF$(pCmdNumOps = 1, pCmdOps(1), "*")                  ' Set the FileMask for FM
   FPath = lclPath                                                ' Set the path
   sIniSetString("FManager", "DefMask", FMask)                    '
   sIniSetString("FManager", "DefDir1", FPath)                    ' Save it

   TP = Tabs(1)                                                   ' Swap the TP pointer
   TP.FileListNm = ""                                             ' Null any FILELIST setting
   TP.FPath = sIniGetString("FManager", "DefDir1", "")            '
   TP.FMask = sIniGetString("FManager", "DefMask", "")            '
   TopScrn = 1                                                    ' From the top
   Lasttop = 0                                                    ' From the top
   TP.AttnDo = TP.AttnDo OR %LoadReq                              ' Refresh
   gTabSwitchCmd = ""                                             ' Pass command to the FM tab
   gTabSwitch = 1                                                 ' Switch to FM tab
   MExit
END METHOD

METHOD  pCmdDO(pCmd AS STRING)
'---------- DO a text block or a .CMD file
LOCAL i AS LONG
LOCAL DMsg AS KBMsg
   IF me.CmdParse(pCmd) THEN EXIT METHOD                          ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN                                         ' If no CMD file name, error
      scError(%eFail, "Missing DO file name")                     '
      EXIT METHOD                                                 '
   END IF                                                         '

   i = sDOMacGet(pCmdOps(1))                                      ' Get the specified DO file loaded to gCmdList()
   IF i = 0 THEN                                                  '
      scError(%eFail, "Error loading " + pCmdOps(1) + ".DO file")'
      EXIT METHOD                                                 '
   END IF                                                         '

   MID$(DMsg.kbString, 1, 1) = CHR$(5)                            ' Flag 1st byte as Hex 5 - DO execute
   MID$(DMsg.kbString, 2, 1) = CHR$(0)                            ' Flag 2nd byte as null
   DMsg.kbInt(1) = 0                                              ' Copy Int as zero
   i = PostMessage(hWnd, %WM_USER, DMsg.MsgwParam, 0)             ' To the mainline Callback routine
   EXIT METHOD
END METHOD

METHOD  pCmdDOWN(pCmd AS STRING)
'---------- Scroll Down
LOCAL ScrAmt, Lines2Page, PageOvrd  AS LONG
LOCAL ScrCmd, lclCmd AS STRING
   '----- Do basic parsing, assign a default scroll amount
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN                                         ' No Ops, fudge it
      me.CmdParse("DOWN " + Prf.Scroll)                           ' Re-parse it
   END IF
   ScrCmd = UUCASE(pCmdOps(1))                                    ' Pick up the operand
   IF pCmdOpsType(1) <> %OpNum AND INSTR(CHR$(%KWHALF, %KWPAGE, %KWFULL, %KWDATA, %KWMAX, %KWCSR), CHR$(pCmdOpsType(1))) = 0 THEN _
      MErrExit(%eFail, "Invalid scroll amount")                   '

   '----- If in the data set the current cursor location so it's preserved
   IF IsCData THEN                                                ' No, In data area?
      me.CurSetReq(%Position, CsrAbsLine, CCol - gLNPadCol + Offset, %False) ' Set cursor set attempt
   ELSEIF IsCLinN THEN                                            ' No, In Line Number area?
      me.CurSetReq(%Position, CsrAbsLine, 0, %False)              ' Set cursor set attempt
   END IF

   '----- Handle the various scroll types, set ScrAmt appropriately
   IF VAL(ScrCmd) <> 0 THEN                                       ' A Number?
      ScrAmt = VAL(ScrCmd)                                        ' Yes, use it
   ELSEIF pCmdOpsType(1) = %KWHALF THEN                           ' Half?
      ScrAmt = INT((gwScrHeight - 2) / 2)                         ' Yes, set it
      ScrAmt = ScrAmt \ Prf.HexMode                               ' Adjust for Hex
   ELSEIF pCmdOpsType(1) = %KWPAGE THEN                           ' Page?
      ScrAmt = INT((gwScrHeight - 2 - Prf.Cols))                  ' Yes, set it
      ScrAmt = ScrAmt \ Prf.HexMode                               ' Adjust for Hex
      Prf.ScrlPageSus = %False                                    ' Clear suspend PAGE mode
   ELSEIF pCmdOpsType(1) = %KWFULL THEN                           ' Full?
      ScrAmt = INT((gwScrHeight - 2 - Prf.Cols))                  ' Yes, set it
      ScrAmt = ScrAmt \ Prf.HexMode                               ' Adjust for Hex
   ELSEIF pCmdOpsType(1) = %KWDATA THEN                           ' Data?
      ScrAmt = INT((gwScrHeight - 2) - Prf.Cols) - 1              ' Yes, set it
      ScrAmt = ScrAmt \ Prf.HexMode                               ' Adjust for Hex
   ELSEIF pCmdOpsType(1) = %KWMAX THEN                            ' Max?
      ScrAmt = 9999999                                            ' Yes, set a BIG number
   ELSEIF IsCData THEN                                            ' No, In data area?
      IF TopScrn = CsrAbsLine THEN                                ' If Cursor at top of data area
         ScrAmt = gwScrHeight - 2 - Prf.Cols                      ' Do a full page
         ScrAmt = ScrAmt \ Prf.HexMode                            ' Adjust for Hex
      ELSE                                                        '
         TopScrn = CsrAbsLine                                     '
         ErrFlag = %eNone                                         ' Say we've done it
         MExitMeth                                                '
      END IF                                                      '
   ELSEIF IsCLinN THEN                                            ' No, In Line Number area?
      IF TopScrn = CsrAbsLine THEN                                ' If Cursor at top of data area
         ScrAmt = gwScrHeight - 2 - Prf.Cols                      ' Do a full page
         ScrAmt = ScrAmt \ Prf.HexMode                            ' Adjust for Hex
      ELSE                                                        '
         TopScrn = CsrAbsLine                                     '
         ErrFlag = %eNone                                         ' Say we've done it
         MExitMeth                                                '
      END IF                                                      '
   ELSE                                                           '
      ScrAmt = gwScrHeight - 2 - Prf.Cols                         ' None of above, Scroll forward one screen - 1 line
      ScrAmt = ScrAmt \ Prf.HexMode                               ' Adjust for Hex
   END IF                                                         '

   '----- Adjust the top of screen location now
   IF LEFT$(Prf.EOL, 4) = "AUTO" AND pCmdOpsType(1) = %KWPAGE THEN' Page in EOL=AUTO mode?
      IF Prf.PageFlag = 2 THEN                                    ' Page in EOL=AUTO mode and not overridden?
         Lines2Page = me.CRPFwd2Page                              ' Get # lines to next =PAGE> line (or eof)
         IF Lines2Page <= (gwScrHeight - Prf.Cols - 2) THEN       ' Did this page all fit on the screen?
            TopScrn = me.CRPFwd(%mPage, TopScrn, 1)               ' Move TopScrn by a PAGE
         ELSE                                                     '
            ScrAmt = INT((gwScrHeight - 2 - Prf.Cols))            ' Calc like FULL does
            ScrAmt = ScrAmt \ Prf.HexMode                         '
            TopScrn = me.CRPFwd(%mVisible, TopScrn, ScrAmt)       ' Move TopScrn by scroll amount
            Prf.ScrlPageSus = %False                              '
         END IF                                                   '
      ELSEIF Prf.PageFlag = 1 THEN                                ' Just PAGE ON?
         TopScrn = me.CRPFwd(%mPage, TopScrn, 1)                  ' Move TopScrn by a PAGE
      ELSE                                                        '
         TopScrn = me.CRPFwd(%mVisible, TopScrn, ScrAmt)          ' Move TopScrn by scroll amount
      END IF                                                      '
   ELSE                                                           ' Not AUTO PAGE
      TopScrn = me.CRPFwd(%mVisible, TopScrn, ScrAmt)             ' Move TopScrn by scroll amount
      IF ScrCmd = "M" OR ScrCmd = "MAX" THEN                      ' If down max, backup by one screen
         TopScrn = me.CRPBack(%mVisible, TopScrn, gwScrHeight - 3 - Prf.Cols)
      END IF                                                      '
   END IF                                                         '
   ErrFlag = %eNone                                               ' Say we 've done it
   MExit
END METHOD

METHOD  pCmdDROPKEEP(pCmd AS STRING)
'---------- Drop/Keep commands
LOCAL lclTag AS STRING * 8, lclFlag, workFlag AS QUAD, valid, which AS STRING
LOCAL i, j, k, NumDel, GrpCount, GrpS, GrpE, gi, gj, gk AS LONG
   MEntry
   '----- Do some basic parsing
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   which = UUCASE(pCmdOps(0))                                     ' Get KEEP or DROP
   IF pCmdNumOps < 2 THEN MErrExit(%eFail, "Missing " + which + " operand") ' Exit if nothing to do

   '----- Spin through the operands
   FOR i = 1 TO pCmdNumOps                                        ' See if we have a :tagname operand
      IF pCmdOpsType(i) = %OpTag THEN                             ' Possible
         IF IsEQ(pCmdOps(i), ":ZALL") THEN MErrExit(%eFail, "Invalid Tag operand") ' Eliminate ZALL here
         IF sTagVal(pCmdOps(i), %True) THEN MErrExit(%eFail, "Invalid Tag operand") ' See if 1st is a valid Tag
         lclTag = LSET$(UUCASE(pCmdOps(i)), 8)                    '
         ARRAY DELETE pCmdOps(i): ARRAY DELETE pCmdRaw(i): ARRAY DELETE pCmdOpsType(i)
         DECR pCmdNumOps                                          ' Strip off the Tag now
      END IF                                                      '
   NEXT i                                                         '
   valid = CHR$(%KWFIRST, %KWLAST)                                ' Build valid operand string
   IF which = "DROP" THEN valid += CHR$(%KWALL)                   ' Add ALL if the DROP command

   FOR i = 1 TO pCmdNumOps                                        ' Test operands for validity
      IF INSTR(valid, CHR$(pCmdOpsType(i))) = 0 THEN MErrExit(%eFail,"Unknown " + which + " operand - " + pCmdOps(i)) ' Ensure valid keywords?
      SELECT CASE pCmdOpsType(i)                                  '
         CASE %KWFIRST: BIT SET lclFlag, %CRTFirst                '
         CASE %KWLAST:  BIT SET lclFlag, %CRTLast                 '
         CASE %KWALL:   BIT SET lclFlag, %CRTAll                  '
      END SELECT                                                  '
   NEXT i                                                         '

   '----- Do some combo validation
   IF (BIT(lclFlag, %CRTFirst) AND (BIT(lclFlag, %CRTLast) OR BIT(lclFlag, %CRTAll))) OR _
      (BIT(lclFlag, %CRTLast) AND (BIT(lclFlag, %CRTAll))) THEN _ '
      MErrExit(%eFail, "Cannot combine FIRST, LAST, or ALL")      '
   me.ResetFunc(%ResetFind)                                       ' Clear previous FIND hilighting

   '----- OK, lets look for a TAG group
   FOR i = 1 TO LastLine                                          ' Lets look for 'proper' tag groups
      IF IsLData(i) THEN                                          ' Only Data lines matter
         IF LTagG(i) = lclTag THEN                                ' 1st line of Our Tag?
            GrpS = i: GrpCount = 1                                ' Save start line in j
            FOR k = GrpS + 1 TO LastLine                          ' Look for end of Tag group
               IF ISFALSE IsLData(k) AND ISFALSE IsLBottom(k) THEN ITERATE
               IF IsLBottom(k) OR LTagG(k) <> lclTag THEN EXIT FOR'
               INCR GrpCount: GrpE = k                            ' Count size of group, save as 'end' of group
            NEXT k                                                '
            IF GrpCount > 1 THEN                                  ' A Group?
               GOSUB DoDrop                                       ' Go do the DROP
               i += IIF(which = "KEEP", 1, (GrpCount - 1))        ' Adjust i to continue position
            END IF                                                '
         END IF                                                   '
      END IF                                                      '
   NEXT i                                                         '

   '----- Tell user the result
   IF NumDel = 0 THEN                                             '
      scError(0, "No eligible lines dropped")                     '
   ELSE                                                           '
      scError(0, "Dropped " + FORMAT$(NumDel) + IIF$(NumDel = 1, " line", " lines"))
   END IF                                                         '
   MExitMeth

'----- Actually do the work
DoDrop:
   IF which = "KEEP" THEN GOTO KeepVersion                        '

   '----- This is the DROP version
   FOR gi = GrpS TO GrpE                                          ' Loop through the located Tag Group
      IF IsLData(gi) THEN LPCmdS(gi) = (0 - PCmdIX)               ' Mark all lines as :ZNF
   NEXT gi

   IF BIT(lclFlag, %CRTFirst) THEN                                ' Just the 1st?
      me.ModSet(GrpS)                                             ' Remember we changed something
      me.LTxtFree(GrpS)                                           ' Go free the dynamic string
      me.LEntDel(GrpS)                                            ' Remove from the L() array
      me.AdjustPending(GrpS, 1, 0)                                ' Adjust pending requests
      DECR LastLine: DECR LastReal                                ' Adjust LastLine and LastReal
      INCR NumDel                                                 ' Count deletes
      RETURN                                                      '
   END IF                                                         '
   IF BIT(lclFlag, %CRTLast) THEN                                 ' Just the last?
      me.ModSet(GrpS)                                             ' Remember we changed something
      me.LTxtFree(GrpE)                                           ' Go free the dynamic string
      me.LEntDel(GrpE)                                            ' Remove from the L() array
      me.AdjustPending(GrpE, 1, 0)                                ' Adjust pending requests
      DECR LastLine: DECR LastReal                                ' Adjust LastLine and LastReal
      INCR NumDel                                                 ' Count deletes
      RETURN                                                      '
   END IF                                                         '
   FOR gi = GrpS TO GrpE                                          ' Loop through the located Tag Group
      IF IsLData(gi) THEN                                         ' Only Data lines matter
         me.LTxtFree(gi)                                          ' Go free the dynamic string
         me.LEntDel(gi)                                           ' Remove from the L() array
         me.AdjustPending(gi, 1, 0)                               ' Adjust pending requests
         DECR LastLine: DECR LastReal                             ' Adjust LastLine and LastReal
         INCR NumDel                                              ' Count deletes
         IF gi = GrpE THEN RETURN                                 '
         DECR GrpE: DECR gi                                       ' Adjust loop controls
      END IF                                                      '
   NEXT gi
   RETURN

KeepVersion:
   '----- This is the KEEP version
   me.ModSet(GrpS)                                                ' Remember we changed something
   FOR gi = GrpS TO GrpE                                          ' Loop through the located Tag Group
      IF IsLData(gi) THEN LPCmdS(gi) = PCmdIX                     ' Mark all lines as :ZF
   NEXT gi

   IF BIT(lclFlag, %CRTFirst) THEN INCR GrpS                      ' Just the 1st? Step over it
   IF BIT(lclFlag, %CRTLast)  THEN DECR GrpE                      ' Just the last? Back end by 1.
   FOR gi = GrpS TO GrpE                                          ' Loop through the located Tag Group
      IF IsLData(gi) THEN                                         ' Only Data lines matter
         me.LTxtFree(gi)                                          ' Go free the dynamic string
         me.LEntDel(gi)                                           ' Remove from the L() array
         me.AdjustPending(gi, 1, 0)                               ' Adjust pending requests
         DECR LastLine: DECR LastReal                             ' Adjust LastLine and LastReal
         INCR NumDel                                              ' Count deletes
         IF gi = GrpE THEN RETURN                                 '
         DECR GrpE: DECR gi                                       ' Adjust loop controls
      END IF                                                      '
   NEXT gi
   RETURN
END METHOD

METHOD  pCmdEDIT(pCmd AS STRING)
'---------- OPEN a new file into the work area
LOCAL lclMsg, SetProf, fn, MSG AS STRING

   MEntry
   '----- Do basic parsing
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   fn = IIF$(IsEQ(pCmdOps(1), "NEW") AND pCmdOpsType(1) <> %OpQStr, $Empty, pCmdOps(1))
   fn = IIF$(fn = "*", sGetFnClipboard, fn)                       ' Swap clipboard if asked for
   ENV.PMode = %MEdit                                             ' Set to Edit mode

   '----- See if a specific Profile name provided
   IF pCmdNumOps > 1 AND LEFT$(pCmdOps(2), 1) <> "." THEN _       ' A profile request?
      MErrExit(%eFail, "Unknown EDIT operand: " + pCmdOps(2))     ' Issue error
   SetProf = IIF$(pCmdNumOps > 1, UUCASE(MID$(pCmdOps(2), 2)), "")' Extract the Profile name

   '----- See if it exists already elsewhere
   TOP.Setup("APECR", SetProf, "Select file for Edit", fn)        ' Do some validation
   IF TOP.EXEC THEN                                               '
      IF TOP.RSTabFound THEN                                      ' Open elsewhere?
         MSG = TOP.ResultMsg                                      ' Save message around Tab switch
         TP = Tabs(TOP.TabFound)                                  ' Switch to it
         gTabSwitch = TOP.TabFound                                ' Set global switch
         MErrExit(%eFail, MSG)                                    ' Issue error and exit
      END IF
      scError(%eFail, TOP.ResultMsg)                              ' Issue error and exit
      IF TOP.RSUserCancel OR TOP.RSExistFail THEN MexitMeth       ' Cancelled out? Doesn't exist?
      ENV.PMode = %MEdit                                          ' Set to EDIT mode to start
      IF TOP.RSROnlyFail THEN                                     ' RO failure?
         ENV.PMode = ENV.PMode AND (&HFFFFFFFF - %MEdit)          ' Set to VIEW mode
         ENV.PMode = %MView                                       '
         gGlblMessage = "File is Read-Only, VIEW used instead of EDIT"
      END IF                                                      '
   END IF                                                         '
   fn = TOP.FilePath: SetProf = TOP.Profile                       ' Copy fn/profile in case from Prompted filename
   me.LoadTheText(fn, SetProf)                                    ' Go load it in

   MExit
END METHOD

METHOD  pCmdEND(pCmd AS STRING)
'---------- Terminate the whole thing now
LOCAL i, j, x, keep, SaveOK, fNum, lclModdFlag, fline, tline, MCount AS LONG, t, k, d, fn AS STRING
LOCAL SetVar() AS STRING, SetVarCtr AS LONG, SetKey AS STRING, FD AS DIRDATA, filesize AS QUAD
   MEntry
   ON ERROR GOTO EndBail                                          ' Set trap
   IF me.CmdParse(pCmd) THEN GOTO EndResume                       ' Do basic parsing, exit if errors
   CurrPcmd = "END"                                               ' Make getting here via =X or RETURN look the same as END

   IF IsClip THEN                                                 ' Are we in ClipBoard Mode?
      pCmdCUT("CUT .ZF .ZL " + ClipName)                          ' Go CUT all the data to the Clipboard
      GOTO DoDelExit                                              ' Go delete tab and exit
   END IF                                                         '

   IF IsSetEdit THEN                                              ' SetEdit?
      gSetCount = 0                                               '
      FOR x = 1 TO LastLine                                       ' Re-write the data
         IF ISFALSE IsLData(x) THEN ITERATE FOR                   ' Skip non Data lines
         t = LTxtG(x)                                             ' Get a line
         i = INSTR(t, "=")                                        ' Locate =
         IF i THEN                                                ' Got one
            IF i <> 1 THEN                                        ' A normal key=value style
               IF SetVarCtr > 0 THEN GOSUB DumpKey                ' An item in hold
               SetKey = LEFT$(t, i - 1)                           ' Save the key
               IF VERIFY(SetKey, $AlphaNum + ".?*_") <> 0 THEN _  ' Only reasonable characters, plus '.'
                  MErrExit(%eFail, "Invalid SET variablename: " + SetKey)
               IF VERIFY(LEFT$(SetKey, 1), $Numeric + ".?*_") = 0 THEN _ ' No leading numbers
                  MErrExit(%eFail, "Invalid SET variablename: " + SetKey)
               INCR SetVarCtr                                     ' Bump var count
               IF SetVarCtr > UBOUND(SetVar()) THEN _             ' Time to expand?
                  REDIM PRESERVE SetVar(1 TO SetVarCtr) AS STRING '
               SetVar(SetVarCtr) = MID$(t, i + 1)                 ' Save in work array

            ELSE                                                  ' An =value stype
               INCR SetVarCtr                                     ' Bump var count
               IF SetVarCtr > UBOUND(SetVar()) THEN               ' Time to expand?
                  REDIM PRESERVE SetVar(1 TO SetVarCtr) AS STRING '
               END IF                                             '
               SetVar(SetVarCtr) = MID$(t, 2)                     ' Save in work array

            END IF                                                '
         END IF                                                   '
      NEXT x                                                      '
      IF SetVarCtr > 0 THEN GOSUB DumpKey                         ' An item in hold
      ARRAY SORT gSetKey() FOR gSetCount, TAGARRAY gSetData()     '
      sUpdSetTable                                                ' Write the table
      '----- Request tab be deleted
      INCR gTabDelCtr                                             ' Bump table index
      gTabDelList(gTabDelCtr) = PgNumber                          ' Stuff in our page number

      GOTO EndResume                                              ' Exit
   END IF                                                         '

   IF IsView THEN                                                 ' Just Viewing?
      IF ISTRUE ENV.ViewWarn AND IsModdFlag THEN                  ' Was data modified and ENV.ViewWarn set?
         i =sDoMsgBox("Text in this View session has been modified," + $CRLF + _
                   "Do you want to continue |KEND|B without saving it?", %MB_YESNO + %MB_USERICON, "SPFLite")
         IF i = %IDNO THEN                                        '
            gfEndAll = %False                                     ' Kill the EndAll flag and exit
            GOTO EndResume                                        '
         END IF                                                   '
      END IF                                                      '
      fn = TIP.FilePath                                           ' Build full filename
      me.UnWatchQueue(fn)                                         ' Unwatch and dequeue it
      GOTO DoDelExit                                              ' Go delete tab and exit
   END IF                                                         '

   IF IsBrowse THEN                                               ' Just Browsing?
      fn = TIP.FilePath                                           ' Build full filename
      me.UnWatchQueue(fn)                                         ' Unwatch and dequeue it
      GOTO DoDelExit                                              ' Go delete tab and exit
   END IF                                                         '

   IF IsMEdit THEN                                                ' If a MEdit collect ModdFlag status
      FOR j = 1 TO MEditCount                                     ' Any MEdit files in modified state?
         IF me.MEditFlagGet(j) THEN lclModdFlag = %True: INCR MCount ' Turn on local ModdFlag, count modified
      NEXT j                                                      '
   ELSE                                                           ' Else the simple one
      IF IsModdFlag THEN lclModdFlag = %True                      ' Copy global to local ModdFlag
   END IF                                                         '

   IF Prf.AutoSave = 0 THEN                                       ' AutoSave OFF NOPROMPT?
      ' Do nothing                                                ' Nothing to do

   ELSEIF Prf.AutoSave = 1 THEN                                   ' AutoSave ON NOPROMPT??
      IF lclModdFlag THEN                                         ' Yes, was data modified?
         IF ISFALSE IsMEdit AND (TIP.FilePath = $Empty) THEN      ' Probable (Empty) or cloned file?
            pCmdSaveAs("SAVEAS")                                  ' Go do a SaveAs
            IF ErrFlag <> %eNone THEN                             ' Cancelled out?
               me.UnWatchQueue("")                                ' Unwatch and dequeue it
               GOTO DoDelExit                                     ' Go delete tab and exit
            ELSE                                                  '
               SaveOK = %True                                     ' Remember good save done
               lclmoddflag = %False                               ' No longer 'modified'
            END IF                                                '
         ELSE                                                     '
            pCmdSAVE("SAVE MEditOnly")                            ' Else a normal save
         END IF                                                   '
         IF errFlag = %eNone THEN                                 ' Save OK?
            SaveOK = %True                                        ' Remember good save done
            lclmoddflag = %False                                  ' No longer 'modified'
         END IF                                                   '
      END IF                                                      '

   ELSEIF (Prf.AutoSave AND 2) = 2 THEN                           ' AutoSave ON/OFF PROMPT
      IF lclModdFlag THEN                                         ' Yes, was data modified?
         IF ISFALSE keep THEN                                     ' Avoid prompting twice
            IF (Prf.AutoSave AND 1) = 1 THEN                      ' AutoSave ON PROMPT
               IF ISFALSE IsMEdit THEN                            ' Choose message format
                  i = sDoMsgBox("Save |K" & TIP.File + "|B before exiting?", %MB_YESNOCANCEL OR %MB_DEFBUTTON1, "SPFLite - File " + TIP.File + " has been modified")
               ELSE                                               ' Use the MEdit format
                  i = sDoMsgBox("Save modified files before exiting?", %MB_YESNOCANCEL OR %MB_DEFBUTTON1, "SPFLite - File " + TIP.File + " has been modified")
               END IF                                             '
            ELSE                                                  ' AutoSave OFF PROMPT
               IF ISFALSE IsMEdit THEN                            ' Choose message format
                  i = sDoMsgBox("Save |K" & TIP.File + "|B before exiting?", %MB_YESNOCANCEL OR %MB_DEFBUTTON2, "SPFLite - File " + TIP.File + " has been modified")
               ELSE                                               ' Use the MEdit format
                  i = sDoMsgBox("Save modified files before exiting?", %MB_YESNOCANCEL OR %MB_DEFBUTTON2, "SPFLite - File " + TIP.File + " has been modified")
               END IF                                             '
            END IF                                                '
         ELSE                                                     '
            i = %IDYES                                            '
         END IF                                                   '
         IF i = %IDYES THEN                                       '
            IF ISFALSE IsMEdit AND (TIP.FilePath = $Empty) THEN   ' Probable (Empty)
               pCmdSaveAs("SAVEAS")                               ' Go do a SaveAs
            ELSE                                                  '
               pCmdSAVE("SAVE MEditOnly")                         ' Else a normal save
            END IF                                                '
            IF errFlag = %eNone THEN                              ' If Save OK
               SaveOK = %True                                     ' Remember good save done
               lclmoddflag = %False                               ' No longer 'modified'
            END IF                                                '
         ELSEIF i = %IDCANCEL THEN                                ' CANCEL the END?
            gfEndAll = %False                                     ' Kill the EndAll flag and exit
            GOTO EndResume                                        '
         ELSEIF i = %IDNO THEN                                    ' NO
            ' Nothing                                             ' Do nothing
         END IF                                                   '
      END IF                                                      '
   END IF                                                         '

   IF IsMEdit THEN                                                ' If a MEdit clear each filename
      FOR j = MEditCount TO 1 STEP -1                             ' Loop in reverse (Medit table shrinks as deleted)
         fn = me.MEditListGet(j)                                  ' Get a filename
         IF ISFALSE me.MEditFlagGet(j) THEN                       ' If not modified
            RESET fline, tline                                    ' Zero our answer areas
            FOR i = 1 TO LastLine                                 ' Must search for range of this MEdit file
               IF IsLFile(i) THEN                                 ' Look for =FILE> lines
                  IF fn = TRIM$(LTxtG(i)) THEN                    ' Found start of file
                     fline = i + 1                                ' Save start line number
                     INCR i                                       '
                     DO WHILE ISFALSE IsLFile(i) AND ISFALSE IsLBottom(i)
                        tline = i                                 ' Save as hiwater last line
                        INCR i                                    '
                     LOOP                                         '
                     EXIT FOR                                     ' Done this looping crap
                  END IF                                          '
               END IF                                             '
            NEXT i                                                '
            me.StateSave(fn, fline, tline, %False)                ' Save the STATE data if unmodified, no filtering
         END IF                                                   '
      NEXT j                                                      '
   ELSE                                                           ' Else the simple one
      IF ISFALSE lclModdFlag THEN                                 ' If unmodified
         me.SetNewStart                                           ' Go move .START if needed
         me.StateSave(TIP.FilePath, 1, LastLine, %False)          ' Save the STATE data, no filtering
      END IF                                                      '
   END IF                                                         '

   IF SaveOK THEN                                                 ' Tell user Save was done
      IF ISFALSE IsMEdit THEN                                     '
         gGlblMessage = "File saved"                              '
      ELSE                                                        '
         gGlblMessage = FORMAT$(MCount) + " file" + IIF$(MCount = 1, "", "s") + " saved"
      END IF
   END IF                                                         '

DoDelExit:                                                        '
   ActionStop = %True                                             ' Kill ACTION SAVE processing
   me.CleanTabData                                                ' Wipe out data
   '----- Request tab be deleted
   INCR gTabDelCtr                                                ' Bump table index
   gTabDelList(gTabDelCtr) = PgNumber                             ' Stuff in our page number
   IF SaveOK THEN                                                 ' Tell user Save was done
      IF ISFALSE IsMEdit THEN                                     '
         gTabDelMsg = "File saved"                                '
      ELSE                                                        '
         gTabDelMsg = FORMAT$(MCount) + " file" + IIF$(MCount = 1, "", "s") + " saved"
      END IF
   END IF                                                         '

EndResume:                                                        '
   ON ERROR GOTO 0                                                '
   CmdStackNum = 0                                                ' If an END, no more commands allowed
   MExitMeth                                                      '

EndBail:                                                          '
   RESUME EndResume                                               '
   RETURN                                                         '

DumpKey:
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
END METHOD

METHOD  pCmdENUMWith(pCmd AS STRING)
'---------- Specify the ENUM increment value
LOCAL s AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN _                                       ' No ops, just display it
      MErrExit(%eNone, "ENUMWITH set to " + FORMAT$(gEnumWith))   '
   IF pCmdNumOps > 1 OR (pCmdOpsType(1) <> %OpNum AND pCmdOpsType(1) <> %OpXStr) THEN _
      MErrExit(%eFail, "ENUMWITH requires a single numeric or Hex operand")
   IF pCmdOpsType(1) = %OpNum THEN                                ' Convert simple numeric
      gEnumWith = VAL(pCmdOps(1))                                 ' Save it
   ELSE                                                           ' Must be Hex
      gEnumWith = VAL("&h" + pCmdOps(1))                          ' Save it
   END IF                                                         '
   IF gEnumWith = 0 THEN _                                        ' Better not be 0
      gEnumWith = 1: MErrExit(nMac(%eFail), "ENUMWITH cannot be Zero, reset to 1")
   scError(%eNone, "ENUMWITH set to " + IIF$(pCmdOpsType(1) = %OpNum, pCmdOps(1), "X'" + pCmdOps(1) + "'"))
   MExit
END METHOD

METHOD  pCmdEOL(pCmd AS STRING)
'---------- Alter the EOL setting
LOCAL i AS LONG, wrd1 AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN                                         ' None, error then
      Wrd1 = Prf.EOL                                              ' Display it
      IF VERIFY(Wrd1, $Hex) = 0 THEN Wrd1 = "X'" + Wrd1 + "'"     ' If hex operand, frame it for the message
      MErrExit(%eNone, "EOL set to " + UUCASE(Wrd1))              '
   END IF                                                         '

   Wrd1 = UUCASE(pCmdOps(1))                                      ' Uppercase and pad it
   ARRAY SCAN gEOLFlagList() FOR 7, = Wrd1, TO i                  ' One of the standard list items
   IF ISFALSE i THEN                                              ' Yes, it's OK
      Wrd1 = TRIM$(Wrd1)                                          ' Take off blanks
      IF (LEN(Wrd1) <> 2 AND LEN(Wrd1) <> 4) OR _                 ' See if valid hex
         VERIFY(Wrd1, $Hex) <> 0 THEN _                           ' Oops, Bad
         MErrExit(%eFail, "Invalid EOL operand")                  '
   END IF                                                         '
   IF IsEQ(Wrd1, "NONE") THEN                                     ' If NONE, check some more
      IF Prf.LRECL = 0 AND Prf.RECFM <> "V" THEN _                ' Better not be LRECL = 0 and RECFM <> "V"
         MErrExit(%eFail, "EOL = NONE not valid when LRECL = 0 and RECFM <> V") '
   ELSE                                                           ' EOL <> NONE
      IF Prf.RECFM = "V" THEN _                                   ' Better not be RECFM = "V"
         MErrExit(%eFail, "EOL = " + UUCASE(Wrd1) + " not valid when RECFM = V") '
   END IF                                                         '
   Prf.EOL = Wrd1                                                 ' Save it
   IF VERIFY(Wrd1, $Hex) = 0 THEN Wrd1 = "X'" + Wrd1 + "'"        ' If hex operand, frame it for the message
   scError(%eNone, "EOL set to " + UUCASE(Wrd1))                  '
   sDoStatusBar($SBEOL)                                           ' re-Do the StatusBar EOL box
   MExit
END METHOD

METHOD  pCmdEXCLFLIPSHOW(pCmd AS STRING, cmd AS STRING, msgname AS STRING)
'---------- handle the integrated Common Command code
LOCAL lclCmd, lclWord AS STRING
LOCAL i, j, AllCount AS LONG, lclCursMode AS INTEGER
LOCAL lclCol, ldir, lsline AS LONG                                ' Local saves for FIND ALL
DIM dlines(1 TO 1000) AS LONG                                     ' To hold lines to be deleted
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   me.ResetFunc(%ResetFind)                                       ' Clear previous FIND hilighting

   '----- Parse the operands
   lclCursMode = Zone                                             ' Save CursMode at start of command
   me.FindReset                                                   ' Reset the parse output fields
   IF ISTRUE me.RangeSet("", %False, %False) THEN MExitMeth       ' Setup any range command operands
   SELECT CASE AS CONST$ cmd                                      ' See which parse to use
      CASE "NDELETE"                                              ' NDELETE
         IF sCRTParse("ASDMC1PU") THEN MExitMeth                  ' Scan (All, Direct, Modifier, Cols, Lit1, Clr, User)
      CASE "APPEND", "PREPEND"                                    ' APPEND / PREPEND
         IF sCRTParse("ASM1PU") THEN MExitMeth                    ' Scan (All, Subset, Modifier, Lit1, Clr, User)
      CASE "ULINE", "REVERT", "NULINE", "NREVERT"                 ' ULINE / REVERT / NULINE / NREVERT
         IF sCRTParse("ASME1CP") THEN MExitMeth                   ' Scan (All, Subset, Exclude, Modifier, Lit1, Cols, Clr)
      CASE ELSE                                                   '
         IF sCRTParse("ADMC1PU") THEN MExitMeth                   ' Scan (All, Direct, Modifier, Cols, Lit1, Clr, User)
   END SELECT

   IF ISFALSE BIT(CrtFlag, %CRTNext) AND ISFALSE BIT(CrtFlag, %CRTPrev) AND _
      ISFALSE BIT(CrtFlag, %CRTFirst) AND ISFALSE BIT(CrtFlag, %CRTLast) THEN
      BIT SET CRTFlag, %CRTNext                                   ' Set NEXT if no others specified
      IF DotSet THEN _                                            ' If a range from sRangeSet
         BIT SET CRTFlag, %CRTAll                                 ' then set ALL
   END IF

   '----- Add selective options based of the major function
   SELECT CASE AS CONST$ cmd                                      ' See what extras based on the command
      CASE "NDELETE"                                              ' NDELETE
         BIT SET CRTFlag, %CRTDX                                  ' Add the DX internally
         BIT SET CRTFlag, %CRTNF                                  ' and NF
      CASE "EXCLUDE"                                              ' EXCLUDE
         BIT SET CRTFlag, %CRTMX                                  ' Add the MX internally
         BIT SET CRTFlag, %CRTNX                                  ' and NX
      CASE "NEXCLUDE"                                             ' NEXCLUDE
         BIT SET CRTFlag, %CRTMX                                  ' Add the MX internally
         BIT SET CRTFlag, %CRTNX                                  ' and NX
         BIT SET CRTFlag, %CRTNF                                  ' and NF
      CASE "SHOW"                                                 ' SHOW
         BIT SET CRTFlag, %CRTX                                   ' Add the X internally
      CASE "NSHOW"                                                ' NSHOW
         BIT SET CRTFlag, %CRTX                                   ' Add the X internally
         BIT SET CRTFlag, %CRTNF                                  ' and NF
      CASE "FLIP"                                                 ' FLIP
         BIT SET CRTFlag, %CRTDX                                  ' Add the DX internally
      CASE "NFLIP"                                                ' NFLIP
         BIT SET CRTFlag, %CRTDX                                  ' Add the DX internally
         BIT SET CRTFlag, %CRTNF                                  ' and NF
      CASE "NULINE", "NREVERT"                                    ' NULINE / NREVERT
         BIT SET CRTFlag, %CRTNF                                  ' Add NF
      CASE "APPEND", "PREPEND", "ULINE", "REVERT"                 ' APPEND / PREPEND / ULINE / REVERT
         'BIT SET CRTFlag, %CRTDX                                 ' Add the DX internally
   END SELECT
   me.FindSetup                                                   ' Copy CRT parse's answers

   '----- Negative stuff must have a literal
   SELECT CASE AS CONST$ cmd                                      ' See what extras based on the command
      CASE "NDELETE", "NEXCLUDE", "NSHOW", "NFLIP", "NULINE", "NREVERT"
         IF ISFALSE IsCfLit1 THEN _                               '
            MErrExit(%eFail, cmd + " must specify a search literal")
   END SELECT

   '----- Validate any weird combinations of operands
   IF cmd = "APPEND" OR cmd = "PREPEND"THEN                       ' APPEND/PREPEND must have a simple literal
      IF IsCfL1Picture OR IsCfL1RegEx THEN _                      '
         MErrExit(%eFail, cmd + " must use only a simple literal type")
      IF ISFALSE IsCfLit1 THEN _                                  '
         MErrExit(%eFail, cmd + " string is missing")             '
      IF ISFALSE DotSet AND ISFALSE IsCfAll THEN _                '
         MErrExit(%eFail, cmd + " has no line range specified or ALL")
      IF ISFALSE DotSet AND ISTRUE IsCfAll THEN _                 '
         DotSet = %True: DotStart = 1: DotEnd = LastLine          '
      CfClear(%CrtLit1)                                           ' Clear L1 so sStrSearch won't actually search 'literally'
   ELSE                                                           '
      IF ISFALSE IsCfLit1 THEN                                    ' Lit1 present?
         IF ISFALSE IsCfAll AND ISFALSE DotSet AND _              ' No Lit1, better have ALL or DotSet
            ISFALSE IsCfX AND ISFALSE IsCfNX THEN _               ' or X or NX
            MErrExit(%eFail, "No search string, but no ALL specified")
         IF IsCfWord OR IsCfPrefix OR IsCfSuffix THEN _           ' No extraneous operands
            MErrExit(%eFail, "No search string, but string options are present")
      END IF                                                      '

      IF IsCfAll AND ISFALSE IsCfLit1 AND _                       ' Ending up with just ALL?
         ISFALSE DotSet AND ISFALSE IsCfX AND ISFALSE IsCfNX THEN ' Nothing else?
         IF cmd = "DELETE" AND IsNE(pCmdOps(0), "DELETE") THEN _  '
            MErrExit(%eFail, "To delete entire file, fully spell out DELETE ALL")
      END IF                                                      '
   END IF                                                         '

   '----- Do the simple X/FLIP/SHOW ALL style
   IF IsCfAll AND ISFALSE IsCfLit1 AND _                          ' Ending up with just ALL?
      ISFALSE DotSet AND _                                        ' Nothing else
      ISFALSE IsCfU AND ISFALSE IsCfNU AND _                      '
      ISFALSE IsCfX AND ISFALSE IsCfNX THEN                       '
      IF cmd = "EXCLUDE" OR cmd = "FLIP" OR cmd = "SHOW" OR _     ' The simple ALL visibility ones?
         cmd = "ULINE" OR cmd = "REVERT" THEN                     '
         Topscrn = 1                                              '
         i = 2                                                    ' Spin through the file
         DO WHILE ISFALSE IsLBottom(i)                            '
            IF ISFALSE IsLFile(i) AND ISFALSE IsLXClude(i) THEN   ' Ignore MEdit FILE lines
               SELECT CASE AS CONST$ cmd                          ' Which one?
                  CASE "EXCLUDE": LFOn(i, %Invisible)             ' Set Invisible
                  CASE "FLIP"   : LFTog(i, %Invisible)            ' Invert Invisible
                  CASE "SHOW"   : LFOff(i, %Invisible)            ' Remove Invisible
                  CASE "ULINE"  : LFOn(i, %User)                  ' Add USER
                  CASE "REVERT" : LFOff(i, %User)                 ' Remove User
               END SELECT                                         '
            END IF                                                '
            INCR i                                                '
         LOOP                                                     '
         gfXRebuild = %True                                       ' Rebuild exclude stuff
         OnUndoFlag                                               ' Force a snapshot
         MExitMeth                                                ' We're done
      END IF                                                      '
   END IF                                                         '

   '----- Run the standard search procedure
   IF IsCfAll THEN                                                ' The ALL version?
      AllCount = 0                                                ' Reset variables
      IF IsCfNF THEN                                              ' Do appropriate test
         i = me.NFSearch(0, lclCursMode)                          ' Do the initial search
      ELSE                                                        '
         i = me.Search(0, lclCursMode)                            ' Do re-Finds
      END IF                                                      '
      lclCursMode = %CursBad                                      ' Make it look like Bad Cursor
      IF i THEN                                                   '
         i = cfFLine                                              ' Get found line
         INCR AllCount: lclCol = sCol: lDir = sDir: lsLine = sLine' Count and save 1st located for RFIND stuff
         IF cmd = "FLIP" OR cmd = "NFLIP" THEN GOSUB DoFlip       ' If found, Do the Flip
         IF cmd = "NDELETE" THEN GOSUB AddDel                     ' If found, Flag for delete
         IF cmd = "APPEND" THEN GOSUB DoAppend                    ' If found, go append something
         IF cmd = "PREPEND" THEN GOSUB DoPrepend                  ' If found, go prepend something
         IF cmd = "ULINE"  OR cmd = "NULINE" THEN LFOn(i, %User)  ' Add USER
         IF cmd = "REVERT"  OR cmd = "NREVERT" THEN LFOff(i, %User) ' Remove USER
      END IF                                                      '
      DO WHILE i                                                  ' Do till not found
         IF IsCfNF THEN                                           ' Do appropriate test
            i = me.NFSearch(1, lclCursMode)                       ' Do the initial search
         ELSE                                                     '
            i = me.Search(1, lclCursMode)                         ' Do re-Finds
         END IF                                                   '
         IF i THEN                                                ' Found one
            i = cfFLine                                           ' Get found line
            INCR AllCount                                         ' Count successful
            IF cmd = "FLIP" OR cmd = "NFLIP" THEN GOSUB DoFlip    ' If found,Do the Flip
            IF cmd = "NDELETE" THEN GOSUB AddDel                  ' If found, Flag for delete
            IF cmd = "APPEND" THEN GOSUB DoAPPEND                 ' If found, go append something
            IF cmd = "PREPEND" THEN GOSUB DoPrepend               ' If found, go prepend something
            IF cmd = "ULINE"  OR cmd = "NULINE" THEN LFOn(i, %User) ' Add USER
            IF cmd = "REVERT"  OR cmd = "NREVERT" THEN LFOff(i, %User) ' Remove USER
         END IF                                                   '
      LOOP                                                        '
      IF cmd = "NDELETE" THEN GOSUB DelThem                       '
      IF AllCount = 0 THEN                                        ' Issue approp. message
         scError(nMac(%eFail), "No lines found")                  '
      ELSE                                                        '
         scError(0, msgname + " " + FORMAT$(AllCount) + IIF$(AllCount > 1, " lines", " line"))
         sCol = lclCol: sDir = lDir: sLine = lsLine               ' Restore to 1st one found
      END IF                                                      '

   ELSE                                                           ' The non-All version
      IF IsCfNF THEN                                              ' Do appropriate test
         i = me.NFSearch(0, lclCursMode)                          ' Do the initial search
      ELSE                                                        '
         i = me.Search(0, lclCursMode)                            ' Do re-Finds
      END IF                                                      '
      IF i = 0 THEN                                               ' Got it?
         scError(nMac(%eFail), IIF$(IsCfNext, "Bottom of data reached", "Top of data reached"))
         me.cfNotFound = IIF(IsCfNext, 2, 1)                      ' Remember we failed
      ELSE                                                        '
         i = me.sLine                                             ' Get found line
         INCR AllCount                                            '
         IF cmd = "FLIP" OR cmd = "NFLIP" THEN GOSUB DoFlip       ' If found,Do the Flip
         IF cmd = "NDELETE" THEN GOSUB AddDel                     ' If found, Flag for delete
         IF cmd = "DELETE" OR cmd = "NDELETE" THEN GOSUB DelThem  '
         IF cmd = "ULINE"  OR cmd = "NULINE" THEN LFOn(i, %User)  ' Add USER
         IF cmd = "REVERT"  OR cmd = "NREVERT" THEN LFOff(i, %User) ' Remove USER
         scError(0, msgname + " " + FORMAT$(AllCount) + " line")  '
         ErrFlag = %eNone                                         ' Say we 're OK
      END IF                                                      '
   END IF                                                         '
   gfXRebuild = %True                                             ' Force exclude rebuild
   OnUndoFlag                                                     ' Force a snapshot
   MExitMeth

DoAppend:
   IF IsLData(cfFLine) THEN                                       ' Just data lines
      me.LTxtSet(cfFLine, LTxtG(cfFLine) + cfFind)                ' Append the data
      LAttrS(cfFLine) = LAttrG(cfFLine) + REPEAT$(LEN(cfFind), CHR$$(0))  ' Append Attr as well
      me.ModSet(cfFLine)                                          ' Remember we changed something
   END IF                                                         '
   RETURN                                                         '

DoPrepend:
   IF IsLData(cfFLine) THEN                                       ' Just data lines
      me.LTxtSet(cfFLine, cfFind + LTxtG(cfFLine))                ' Prepend the data
      LAttrS(cfFLine) = REPEAT$(LEN(cfFind), CHR$$(0)) + LAttrG(cfFLine)  ' Prepend Attr as well
      me.ModSet(cfFLine)                                          ' Remember we changed something
   END IF                                                         '
   RETURN                                                         '

DoFlip:
   IF IsLData(cfFLine) THEN                                       ' Just data lines
      LFTog(cfFLine, %Invisible)                                  ' Do the FLIP of Invisible status
      IF IsCfPrev THEN                                            ' Set column so no more on this line
         sCol = 1                                                 ' Backward, use col 1
      ELSE                                                        '
         sCol = LEN(L(cfFline).@LTxt)                             ' Forward, use length of line
      END IF                                                      '
   END IF                                                         '
   RETURN                                                         '

AddDel:
   FOR i = 1 TO j                                                 '
      IF dLines(i) = cfFLine THEN RETURN                          ' Don't add line more than once
   NEXT i                                                         '
   INCR j
   IF J > UBOUND(dLines) THEN _                                   ' Expand table if needed
      REDIM PRESERVE dLines(1 TO UBOUND(dLines) * 2)              ' Double it
   dLines(j) = cfFLine                                            ' Save the line number
   RETURN                                                         '

DelThem:
   Allcount = 0
   FOR i = j TO 1 STEP -1                                         ' Now delete the lines (backward)
      IF IsLData(dLines(i)) THEN me.ModSet(dLines(i))             ' Remember we changed something
      me.LTxtFree(dLines(i))                                      ' Go free the dynamic string
      me.LEntDel(dLines(i))                                       ' Remove from the L() array
      me.ModSet(dLines(i))                                        ' Remember we changed something
      me.AdjustPending(dLines(i), -1, 0)                          ' Adjust pending requests
      DECR LastLine: DECR LastReal                                ' Adjust LastLine and LastReal
      INCR AllCount                                               ' Count real deletions
   NEXT i                                                         '
   RETURN
END METHOD

METHOD  pCmdFAV(pCmd AS STRING)
'---------- Add file as a favorite
LOCAL Wrd, wrd2, ln2 AS STRING, i AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF IsMEdit THEN _                                              ' Sorry if MEdit mode
      MErrExit(%eFail, "Can't 'Favorite' in Multi-Edit mode")     '
   wrd2 = IIF$(ISNULL(pCmdOps(1)), "Favorite Files", pCmdOps(1))  ' Use command operand or FAVORITE
   wrd = UUCASE(wrd2)                                             ' UC the working version
   IF wrd = "FAV" THEN wrd = "Favorite Files"                     ' Accept same aliases as command name
   IF wrd = "FAVOURITE" THEN wrd = "Favorite Files"               ' Accept same aliases as command name
   IF wrd = "FAVORITE" THEN wrd = "Favorite Files"                ' Accept same aliases as command name
   ln2 = UUCASE(wrd)                                              ' Uppercase it
   IF ln2 = "RECENT FILES" OR LN2 = "NAMED FAVORITES" OR _        ' Reserved name
      LN2 = "RECENT PATHS" OR LN2 = "FOUND FILES" OR LN2 = "OPEN" OR _  '
      LN2 = "RECENT" OR LN2 = "FOUND" OR LN2 = "PATHS" THEN _     '
      MErrExit(%eFail, "Sorry, name: " + wrd2 + " is reserved")   '
   IF TIP.FilePath = $Empty THEN _                                ' Better be a file name
      MErrExit(%eFail, "Favorite applies only to named files")    '
   CALL scError(%eNone, sFileListAdd(Wrd2, TIP.FilePath))         '
   MExit
END METHOD

METHOD  pCmdFIND(pCmd AS STRING)
'---------- FIND Command
LOCAL lclCmd, lclWord, modemsg, str, fstr, fstr2 AS STRING
LOCAL i, ll AS LONG, AllCount, AllLines, LLastLine AS LONG, lclCursMode AS INTEGER
LOCAL lclCol, ldir, lsline AS LONG                                ' Local saves for FIND ALL
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   lclCursMode = Zone                                             ' Save CursMode at start of command
   me.ResetFunc(%ResetFind)                                       ' Clear previous FIND hilighting

   IF ISTRUE me.RangeSet("", %False, %False) THEN MExitMeth       ' Setup any range command operands
   me.FindReset                                                   ' Reset the parse output fields
   IF sCRTParse("ASDMEC1LPQU") THEN MExitMeth                     ' Scan (All, Subset, Direct, Modifier, Exclude, Cols, Lit1, Left/Right, Pen, QPen, User)
   me.FindSetup                                                   ' Copy CRT parse's answers

   IF pCmdNumOps = 1 AND pCmdOpsType(1) <> %OpQStr AND _          ' Single operand unquoted WORD?
      (IsEQ(pCmdOps(1), "WORD") OR IsEQ(pCmdOps(1), "WORDS")) THEN' or WORDS
      FindWord = %True                                            '
      sDoStatusBar($SBCaseWord)                                   ' re-Do the StatusBar Case box
      MErrExit(%eNone, "FIND mode set to WORD")                   '
   END IF                                                         '

   IF pCmdNumOps = 1 AND pCmdOpsType(1) <> %OpQStr AND _          ' Single operand unquoted CHAR?
      (IsEQ(pCmdOps(1), "CHAR") OR IsEQ(pCmdOps(1), "CHARS")) THEN' or CHARS
      FindWord = %False                                           '
      sDoStatusBar($SBCaseWord)                                   ' re-Do the StatusBar Case box
      MErrExit(%eNone, "FIND mode set to CHARS")                  '
   END IF                                                         '

   IF ISFALSE IsCfNext AND ISFALSE IsCfPrev AND _                 '
      ISFALSE IsCfFirst AND ISFALSE IsCfLast THEN _               '
      CfSet(%CRTNext)                                             ' Set NEXT if no others specified

   IF ISFALSE IsCfLit1 THEN _                                     ' Lit1 present?
      MErrExit(%eFail, "Missing FIND search string")              '
   LocFind = "FIND"                                               ' Remember which LOC / FIND command was last
   IF LEN(CrtL1Raw) < 16 THEN                                     ' Build msg literal
      str = CrtL1Raw                                              '
   ELSE                                                           '
      str = LEFT$(CrtL1Raw, 14) + ".."                            '
   END IF                                                         '
   modemsg = IIF$(IsCfWord, "WORD " + str, "CHARS " + str)        '

   IF IsCfAll THEN                                                ' The ALL version?
      RESET AllCount, AllLines, LLastLine                         ' Reset variables
      i = me.Search(0, lclCursMode)                               ' Do the initial search
      lclCursMode = %CursBad                                      ' Make it look like Bad Cursor
      IF i THEN                                                   '
         i = sLine                                                ' Get found line
         INCR AllCount: lclCol = sCol: lDir = sDir: lsLine = sLine' Count and save 1st located for RFIND stuff
         IF i <> LLastLine THEN INCR AllLines: LLastLine = i      '
      END IF                                                      '
      DO WHILE i                                                  ' Do till not found
         i = me.Search(1, lclCursMode)                            ' Do re-Finds
         IF i THEN                                                ' Found one
            i = sLine                                             ' Get found line
            INCR AllCount                                         ' Count successful
            IF i <> LLastLine THEN INCR AllLines: LLastLine = i   '
         END IF                                                   '
      LOOP                                                        '
      IF AllCount = 0 THEN                                        ' Issue approp. message
         scError(%eNone, modemsg + " not found")                  '
      ELSE                                                        '
         fstr = MID$(LTxtG(cfFLine), cfFCol, cfFLen)              ' Get the found string
         IF IsCfL1Picture OR IsCfL1Hex OR IsCfL1RegEx THEN _      '
            modemsg = IIF$(IsCfWord, "WORD '" + fstr + "'", "CHARS '" + fstr + "'")
         IF VERIFY(fstr, ENV.CharSet + " ") <> 0 THEN             ' Contains non-display chars?
            modemsg = IIF$(IsCfWord, "WORD " + sStr2Hex(fstr), "CHARS " + sStr2Hex(fstr))
         END IF                                                   '
         scError(0, modemsg + " found " + FORMAT$(AllCount) + IIF$(AllCount > 1, " times", " time") + IIF$(AllCount = AllLines, "", " in " + FORMAT$(AllLines) + IIF$(AllLines = 1, " line", " lines")))
         sCol = lclCol: sDir = lDir: sLine = lsLine               ' Restore to 1st one found
      END IF                                                      '
   ELSE                                                           ' The non-All version
      i = me.Search(0, lclCursMode)                               ' Go see if we can find it
      IF i = 0 THEN                                               ' Got it?
         scError(%eNone, IIF$(IsCfNext, "Bottom of data reached", "Top of data reached"))
         cfNotFound = IIF(IsCfNext, 2, 1)                         ' Remember we failed
      ELSE                                                        '
         i = sLine                                                ' Get found line
         lclword = MID$(LTxtG(Sline), sCol, cfFLen)               '
         IF IsCfMX THEN                                           ' X'ing the line?
            scError(%eNone, modemsg + " found and X'd")           '
         ELSE                                                     '
            fstr = MID$(LTxtG(cfFLine), cfFCol, cfFLen)           ' Get the found string
            IF IsCfL1Picture OR IsCfL1Hex OR IsCfL1RegEx THEN _   '
               modemsg = IIF$(IsCfWord, "WORD '" + fstr + "'", "CHARS '" + fstr + "'")
            IF VERIFY(fstr, ENV.CharSet + " ") <> 0 THEN          ' Contains non-display chars?
               modemsg = IIF$(IsCfWord, "WORD " + sStr2Hex(fstr), "CHARS " + sStr2Hex(fstr))
            END IF                                                '
            scError(%eNone, modemsg + " found")                   '
         END IF                                                   '
         ErrFlag = %eNone                                         ' Say we 're OK
      END IF                                                      '
   END IF                                                         '
   MExit
END METHOD

METHOD  pCmdFOLD(pCmd AS STRING)
'---------- Set FOLD ststus
LOCAL nv AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 1 THEN                                         ' Got our operand
      nv = sOnOff(1)                                              ' Go get True/False value for operand 1
      IF nv < 0 THEN MExitMeth                                    ' Error, exit, error msg already issued
      Prf.Fold = nv                                               ' Set the new value
   END IF                                                         '
   scError(0, "FOLD set to " + IIF$(Prf.Fold, "ON", "OFF"))       ' Tell user new value
   sDoStatusBar($SBCaps)                                          ' re-Do the StatusBar CAPS box
   MExit
END METHOD

METHOD  pCmdGLUEWITH(pCmd AS STRING)
'---------- Specify the GLUE joining literal
LOCAL s AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN _                                       ' No ops, just display it
      MErrExit(%eNone, "GLUEWITH set to " + $DQ + ENV.GlueWith + $DQ)
   IF pCmdNumOps > 1 THEN _
      MErrExit(%eFail, "GLUEWITH requires a single literal operand")
   IF pCmdOpsType(1) = %OpQStr AND ISNULL(pCmdOps(1)) THEN        ' Null operand?
      ENV.GlueWith = ""                                           ' Any of the null values
      s = $DQ + $DQ                                               ' Build visible message
   ELSE                                                           '
      IF sCRTParse("1") THEN MExitMeth                            ' Scan (Lit1)
      ENV.GlueWith = CrtL1                                        ' Pick up the literal
      s = CrtL1Raw                                                '
   END IF                                                         '
   scError(%eNone, "GLUEWITH set to " + s)                        '
   sIniSetString("General", "GlueWith", ENV.GlueWith)             ' Save permanently
   MExit
END METHOD

METHOD  pCmdHELP(pCmd AS STRING)
'---------- Invoke HELP file
LOCAL RetC, mapid AS LONG, lclCmd AS STRING
LOCAL rcProcess AS PROCESS_INFORMATION
LOCAL rcStart AS STARTUPINFO

   MEntry
   'HH.EXE -mapid 14 ms-its:D:/Public\Documents\SPFLite6\Docs\SPFLite.chm
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors

   '----- See if a multi-message follow on Help
   IF PCmdNumOps = 0 AND ErrMsgHlpC > 1 THEN                      ' If possible multi-message
      INCR gfDoingMsg                                             '
      DispMessageList()                                           ' Go Display them
      CONTROL SET FOCUS hWnd, %IDC_SPFLiteTAB                     ' Set focus
      DECR gfDoingMsg                                             '
      MExitMeth                                                   '
   END IF

   IF IsEQ(pCmdOps(1), "MACROS") OR IsEQ(pCmdOps(1), "MACRO") THEN'
      lclcmd = ENVIRON$("COMSPEC") + " /C START " + sGetShortName(ENV.EXEPath + "SPFLite Edit Macros.chm")
      retc = SHELL(lclcmd, 0)                                     '
      MExitMeth
   END IF                                                         '
   IF IsEQ(pCmdOps(1), "THINBASIC") OR IsEQ(pCmdOps(1), "BASIC") THEN
      lclcmd = ENVIRON$("COMSPEC") + " /C START " + sGetShortName(ENV.EXEPath + "thinBasic.chm")
      retc = SHELL(lclcmd, 0)                                     '
      MExitMeth
   END IF                                                         '
   mapid = IIF(pCmdNumOps = 1, sHelpIndex(pCmdOps(1)), 0)         ' Get a mapid (or zero)
   lclCmd = "hh.exe -mapid " + FORMAT$(mapid) + " " + sGetShortName(ENV.EXEPath + "SPFLite.chm") + $NUL ' Build command string
   rcStart.cb = LEN(rcStart)                                      ' Initialize the STARTUPINFO structure:
                                                                  ' Start the shelled application:
   RetC = CreateProcess(BYVAL %NULL, BYCOPY lclCmd, BYVAL %NULL, BYVAL %NULL, 1&, _
                        %CREATE_NEW_CONSOLE OR %NORMAL_PRIORITY_CLASS, BYVAL %NULL, _
                        BYVAL %NULL, rcStart, rcProcess)
   CloseHandle(rcProcess.hThread)                                 '
   CloseHandle(rcProcess.hProcess)                                '
   ErrFlag = %eNone                                               '
   MExit
END METHOD

METHOD  pCmdHEX(pCmd AS STRING)
'---------- Turn HEX ON or OFF
LOCAL nv AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   nv = sOnOff(1)                                                 ' Go get True/False value for operand 1
   IF nv < 0 THEN MExitMeth                                       ' Error, exit, error msg already issued
   Prf.HexMode = IIF(ISTRUE nv, 4, 1)                             ' Set the new value
   scError(0, IIF$(ISTRUE nv, "HEX set to ON", "HEX set to OFF")) ' Issue message
   SetCmd                                                         ' Cursor to Cmd line
   MarkedLine = 0                                                 ' Ditto for the Marked line
   MExit
END METHOD

METHOD  pCmdHIDE(pCmd AS STRING)
'---------- Turn Hide ON or OFF
LOCAL nv AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   nv = sOnOff(1)                                                 ' Go get True/False value for operand 1
   IF nv < 0 THEN MExitMeth                                       ' Error, exit, error msg already issued
   IF nv THEN OnHideFlag ELSE OffHideFlag                         ' Set the new value
   scError(%eNone, IIF$(ISTRUE nv, "HIDE set to ON", "HIDE set to OFF"))   ' Issue message
   MExit
END METHOD

METHOD  pCmdHILITE(pCmd AS STRING)
'---------- Turn HILITE ON or OFF
LOCAL Wrd, lclCmd AS STRING, i AS LONG
LOCAL fOn, fOff, fFind, fAuto AS LONG, MSG AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN                                         ' If no operands
      IF ISTRUE Prf.HiFind AND ISTRUE Prf.HiAuto THEN             ' Both ON?
         scError(%eNone, "HILITE FIND & AUTO both set to ON")     ' Tell user
      ELSEIF ISFALSE Prf.HiFind AND ISFALSE Prf.HiAuto THEN       ' Both OFF?
         scError(%eNone, "HILITE FIND & AUTO both set to OFF")    ' Tell user
      ELSE                                                        ' Mixed
         scError(%eNone, "HILITE FIND set to " + IIF$(Prf.HiFind, "ON", "OFF") + _
                              ", AUTO set to " + IIF$(Prf.HiAuto, "ON", "OFF"))
      END IF                                                      '
      MExitMeth                                                   ' Bail out
   END IF                                                         '

   FOR i = 1 TO pCmdNumOps                                        ' Process the operands
      IF INSTR(CHR$(%KWON, %KWOFF, %KWAUTO, %KWFIND),CHR$(pCmdOpsType(i))) = 0 THEN ' Ensure valid keywords?
         scError(%eFail,"Unknown HILITE operand - " + pCmdOps(1)) ' Issue error message
         MExitMeth                                                '
      END IF                                                      '
      SELECT CASE pCmdOpsType(i)                                  ' Handle each one
         CASE %KWON                                               ' ON
            IF fOn THEN GOSUB DupKw                               ' Don't say it twice
            IF fOff THEN GOSUB BadOnOff                           ' Confused!
            fOn = %True                                           '
         CASE %KWOFF                                              ' OFF
            IF fOff THEN GOSUB DupKw                              ' Don't say it twice
            IF fOn THEN GOSUB BadOnOff                            ' Confused!
            fOff = %True                                          '
         CASE %KWFIND                                             ' FIND
            IF fFind THEN GOSUB DupKw                             ' Don't say it twice
            fFind = %True                                         '
         CASE %KWAUTO                                             ' AUTO
            IF fAuto THEN GOSUB DupKw                             ' Don't say it twice
            fAuto = %True                                         '
      END SELECT                                                  '
   NEXT i                                                         '

   IF fOn THEN                                                    ' ON specified
      IF ISFALSE fAuto AND ISFALSE fFind THEN                     ' Without operands
         Prf.HiFind = %True: Prf.HiAuto = %True                   ' Set FIND AUTO ON
         scError(0, "HILITE FIND & AUTO set to ON")               ' Tell user new value
      ELSE                                                        ' Some Ops specified
         IF fFind THEN                                            ' FIND
            Prf.HiFind = %True: MSG = "FIND "                     ' Set FIND ON
         END IF                                                   '
         IF fAuto THEN                                            ' AUTO
            Prf.HiAuto = %True: MSG = IIF$(ISNULL(MSG), "AUTO ", MSG + "& AUTO ") ' Set AUTO ON
         END IF                                                   '
         IF ISNULL(MSG) THEN                                      ' No Operands?
            Prf.HiFind = %True                                    ' Set FIND ON
            Prf.HiAuto = %True                                    ' Set AUTO ON
            MSG = "FIND & AUTO "                                  '
         END IF                                                   '
         scError(0, "HILITE " + MSG + "set to ON")                ' Tell user new value
         MExitMeth                                                '
      END IF                                                      '
   END IF                                                         '
   IF fOff THEN                                                   ' OFF specified
      IF ISFALSE fAuto AND ISFALSE fFind THEN                     ' Without operands
         Prf.HiFind = %False                                      ' Set FIND OFF
         Prf.HiAuto = %False                                      ' Set AUTO OFF
         scError(0, "HILITE FIND & AUTO set to OFF")              ' Tell user new value
      ELSE                                                        ' Some Ops specified
         IF fFind THEN                                            ' FIND
            Prf.HiFind = %False: MSG = "FIND "                    ' Set FIND OFF
         END IF                                                   '
         IF fAuto THEN                                            ' AUTO
            Prf.HiAuto = %False: MSG = IIF$(ISNULL(MSG), "AUTO ", MSG + "& AUTO ") ' Set AUTO OFF
         END IF                                                   '
         IF ISNULL(MSG) THEN                                      ' No Operands?
            Prf.HiFind = %False                                   ' Set FIND OFF
            Prf.HiAuto = %False                                   ' Set AUTO OFF
            MSG = "FIND & AUTO "                                  '
         END IF                                                   '
         scError(0, "HILITE " + MSG + "set to OFF")               ' Tell user new value
         MExitMeth                                                '
      END IF                                                      '
   END IF                                                         '
   IF fFind OR fAuto THEN                                         '
      IF fFind THEN                                               ' Just FIND?
         Prf.HiFind = %True                                       ' Set FIND ON
         MSG = "FIND "                                            ' Setup msg
      END IF                                                      '
      IF fAuto THEN                                               ' Just AUTO?
         Prf.HiAuto = %True                                       ' Set AUTO ON
         MSG = IIF$(ISNULL(MSG), "AUTO ", MSG + " & AUTO ")       ' Setup msg
      END IF                                                      '
      scError(0, "HILITE " + MSG + "set to ON")                   ' Tell user new value
   END IF                                                         '
   MExitMeth

DupKw:
   scError(%eFail, "Duplicate operand - " + Wrd)                  ' Tell user of error
   MExitMeth                                                      '
   RETURN
BadOnOff:
   scError(%eFail, "Both ON and OFF specified ??")                ' Tell user of error
   MExitMeth                                                      '
   RETURN
END METHOD

METHOD  pCmdJOIN(pCmd AS STRING)
'---------- JOIN Command
LOCAL ol, nl, oc, nc AS STRING
LOCAL i AS LONG, AllCount, sp, PrevLastLine AS LONG, lclCursMode AS INTEGER
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   lclCursMode = Zone                                             ' Save CursMode at start of command
   me.ResetFunc(%ResetFind)                                       ' Clear previous FIND hilighting

   IF ISTRUE me.RangeSet("", %False, %False) THEN MExitMeth       ' Setup any range command operands
   me.FindReset                                                   ' Reset the parse output fields

   IF sCRTParse("ASM12PU") THEN MExitMeth                         ' Scan (All, Subset, Modifier, Lit1, Lit2, Clr, User)
   me.FindSetup                                                   ' Copy CRT parse's answers
   CfSet(%CRTNext)                                                ' Set NEXT

   IF ISFALSE IsCfLit1 THEN _                                     ' Lit1 present?
      MErrExit(%eFail, "Missing JOIN search string")              '

   IF ISFALSE IsCfL1Picture AND ISFALSE IsCfL1RegEx THEN _        ' Lit1 must be P type  or R type
      MErrExit(%eFail, "JOIN search string must be P or R type literal")

   IF IsCfL1Picture OR IsCfL1RegEx THEN                           ' P/R type, must be with LM or RM, but not both
      IF (IsCfLM AND ISFALSE IsCfRM) OR _                         '
         (ISFALSE IsCfLM AND IsCfRM) THEN                         '
         '
      ELSE                                                        '
         MErrExit(%eFail, "JOIN search must not have both Line Start and Line End selection")
      END IF                                                      '
   END IF                                                         '

   IF ISFALSE IsCfLit2 THEN                                       ' Lit2 present?
      CfSet(%CrtLit2)                                             ' Provide a default
      cfChange = "!"                                              ' of P'!'
   ELSE
      IF cfSplitPt2 THEN _                                        ' No | chars in Join
         MErrExit(%eFail, "Split character | not allowed in JOIN change string")
   END IF                                                         '
   LocFind = "JOIN"                                               ' Remember which LOC / FIND command was last

   IF IsCfAll THEN                                                ' The ALL version?
      RESET AllCount                                              ' Reset variables
      i = me.Search(0, lclCursMode)                               ' Do the initial search
      lclCursMode = %CursBad                                      ' Make it look like Bad Cursor
      IF i THEN                                                   '
         PrevLastLine = LastLine                                  ' Save LastLine
         IF me.Join() THEN i = %False                             ' Do the Join
         IF LastLine <> PrevLastLine THEN INCR AllCount           ' Linecount altered? count it
      END IF                                                      '
      DO WHILE i                                                  ' Do till not found
         lclCursmode = %CursInternal                              ' Tell StrSearch that sLin/sCol are set
         i = me.Search(1, lclCursMode)                            ' Do re-Finds
         IF i THEN                                                ' Found one
            PrevLastLine = LastLine                               ' Save LastLine
            IF me.Join() THEN                                     ' Do Join, -  signalled LastLine?
               i = %False                                         ' treat like not found
            END IF                                                '
            IF LastLine <> PrevLastLine THEN INCR AllCount        ' Linecount altered? count it
         END IF                                                   '
      LOOP                                                        '
      IF AllCount = 0 THEN                                        ' Issue approp. message
         scError(nMac(%eFail), "Join point not found")            '
      ELSE                                                        '
         scError(0, "Join performed " + FORMAT$(Allcount) + " times") '
      END IF                                                      '

   ELSE                                                           ' The non-All version
      i = me.Search(0, lclCursMode)                               ' Go see if we can find it
      IF i = 0 THEN                                               ' Got it?
         scError(nMac(%eFail), IIF$(IsCfNext, "Bottom of data reached", "Top of data reached"))
         cfNotFound = IIF(IsCfNext, 2, 1)                         ' Remember we failed
      ELSE                                                        '
         PrevLastLine = LastLine                                  ' Save LastLine
         IF me.Join() THEN i = %False                             ' Do the Join
         scError(%eNone, IIF$(LastLine <> PrevLastLine, "Join successful", "Join not possible"))
         ErrFlag = %eNone                                         ' Say we 're OK
      END IF                                                      '
   END IF                                                         '
   MExit
END METHOD

METHOD  pCmdKEYMAP(pCmd AS STRING)
'---------- Display the KeyMap
LOCAL i, j AS LONG, t, tt AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN                                         ' The simple one?
      KbdT.DispKeyMap(0)                                          ' Go Display it
      CONTROL SET FOCUS hWnd, %IDC_SPFLiteTAB                     ' Set focus
      MExitMeth                                                   '
   END IF                                                         '

   IF IsNE(pCmdOps(1), "LIST") THEN _                             ' Valid?
      MErrExit(%eFail, "Unknown KEYMAP operand")                  ' No, kill it

   '----- Do the reverse map
   FOR i = 1 TO 107                                               ' Loop through the Key Table
      IF UUCASE(KbdT.NData(i)) <> "(NULL)" AND UUCASE(KbdT.NData(i)) <> "(PASSTHRU)" THEN     ' Set to something?
         tt = LSET$(KbdT.Labl(i), 24) + KbdT.NData(i) + $CRLF     ' Build title / contents
         t += tt                                                  ' Add line to string
      END IF                                                      '
      IF UUCASE(KbdT.SData(i)) <> "(NULL)" AND UUCASE(KbdT.SData(i)) <> "(PASSTHRU)" THEN     ' Set to something?
         tt = LSET$("Shft-" + KbdT.Labl(i), 24) + KbdT.SData(i) + $CRLF ' Build title / contents
         t += tt                                                  ' Add line to string
      END IF                                                      '
      IF UUCASE(KbdT.CData(i)) <> "(NULL)" AND UUCASE(KbdT.CData(i)) <> "(PASSTHRU)" THEN     ' Set to something?
         tt = LSET$("Ctrl-" + KbdT.Labl(i), 24) + KbdT.CData(i) + $CRLF ' Build title / contents
         t += tt                                                  ' Add line to string
      END IF                                                      '
      IF UUCASE(KbdT.AData(i)) <> "(NULL)" AND UUCASE(KbdT.AData(i)) <> "(PASSTHRU)" THEN     ' Set to something?
         tt = LSET$("Alt-" + KbdT.Labl(i), 24) + KbdT.AData(i) + $CRLF ' Build title / contents
         t += tt                                                  ' Add line to string
      END IF                                                      '
      IF UUCASE(KbdT.SCData(i)) <> "(NULL)" AND UUCASE(KbdT.SCData(i)) <> "(PASSTHRU)" THEN   ' Set to something?
         tt = LSET$("Shft-Ctrl-" + KbdT.Labl(i), 24) + KbdT.SCData(i) + $CRLF ' Build title / contents
         t += tt                                                  ' Add line to string
      END IF                                                      '
      IF UUCASE(KbdT.SAData(i)) <> "(NULL)" AND UUCASE(KbdT.SAData(i)) <> "(PASSTHRU)" THEN   ' Set to something?
         tt = LSET$("Shft-Alt-" + KbdT.Labl(i), 24) + KbdT.SAData(i) + $CRLF ' Build title / contents
         t += tt                                                  ' Add line to string
      END IF                                                      '
      IF UUCASE(KbdT.SCAData(i)) <> "(NULL)" AND UUCASE(KbdT.SCAData(i)) <> "(PASSTHRU)" THEN ' Set to something?
         tt = LSET$("Shft-Ctrl-Alt-" + KbdT.Labl(i), 24) + KbdT.SCAData(i) + $CRLF ' Build title / contents
         t += tt                                                  ' Add line to string
      END IF                                                      '
      IF UUCASE(KbdT.CAData(i)) <> "(NULL)" AND UUCASE(KbdT.CAData(i)) <> "(PASSTHRU)" THEN   ' Set to something?
         tt = LSET$("Ctrl-Alt-" + KbdT.Labl(i), 24) + KbdT.CAData(i) + $CRLF ' Build title / contents
         t += tt                                                  ' Add line to string
      END IF                                                      '
   NEXT i
   gKeyPrimOper = ""                                              ' Null clipboard descriptor
   sWriteClipboard(t)                                             ' Write string to clipboard
   pCmdCLIP("CLIP")                                               '
END METHOD

METHOD  pCmdLCUC(pCmd AS STRING)
'---------- LC / UC / SC / TC Commands
LOCAL lclCmd, lclWord, Txt1, txt2 AS STRING
LOCAL i, j, k, scol, ecol, m, x, y, z, Caps, lclFlag AS LONG, lclCursMode AS INTEGER
LOCAL NumDone AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   lclCursMode = %CursBad                                         ' Pretend a bad cursor location
   me.ResetFunc(%ResetFind)                                       ' Clear previous FIND hilighting
   IF ISTRUE me.RangeSet("", %True, %False) THEN MExitMeth        ' Setup any range command operands
   IF pCmdNumOps = 0 AND ISFALSE DotSet THEN _                    ' Better be something
      MErrExit(%eFail, "Pending line range")                      '
   me.FindReset                                                   ' Reset the parse output fields
   IF sCRTParse("ASEU") THEN MExitMeth                            ' Scan (All, Subset, Exclude, User)
   me.FindSetup                                                   ' Copy CRT parse's answers
   IF DotSet OR _                                                 ' If a line range
      IsCfU OR IsCfNU OR IsCfX OR IsCfNX THEN                     ' Or U / NU / X / NX
      CfSet(%CRTALL)                                              ' Force ALL if no modifiers
   END IF                                                         '

   scol = Prf.BndLeft: ecol = Prf.BndRight                        ' Get local bounds
   IF ecol = 0 THEN ecol = MaxLength                              '
   Caps = %True                                                   ' 1st word overall gets capitalized

   IF IsCfAll THEN                                                ' The ALL version?
      i = me.Search(0, lclCursMode)                               ' Do the initial search
      IF i THEN GOSUB DoChange                                    ' If found, Do the change
      DO WHILE i                                                  ' Do till not found
         i = me.Search(1, lclCursMode)                            ' Do re-Finds
         IF i THEN GOSUB DoChange                                 ' If found, Do the Change
      LOOP                                                        '
   ELSE                                                           ' The non-All version
      i = me.Search(0, lclCursMode)                               ' Go see if we can find it
      IF i THEN GOSUB DoChange                                    ' If found, Do the change
   END IF                                                         '

   IF NumDone = 0 THEN                                            '
      scError(%eFail, "No eligible line(s) altered")              '
   ELSE                                                           '
      scError(%eNone, FORMAT$(NumDone) + " line" + IIF$(NumDone = 1, "", "s") + " altered")
   END IF                                                         '
   MExitMeth                                                      '

DoChange:
   j = cfFLine                                                    ' Get the line index
   me.CurSetReq(%Change, j, 1 + Offset, %True)                    ' Set cursor set attempt
   me.ModSet(j)                                                   ' Remember we changed something
   SELECT CASE UUCASE(pCmdOps(0))                                 ' See which we do
      CASE "UC"                                                   ' UpperCase
         IF IsLData(j) THEN                                       ' Only Data lines
            GOSUB DoPMinus                                        ' Do the +/- stuff
            me.LFlagBitOff(j, %EQChange)                          ' Reset ==CHG>
            me.UpdLControl(j)                                     ' And LLCtl
            IF Prf.BndLeft > 1 OR Prf.BndRight > 0 THEN           ' Honour BNDS setting?
               FOR k = 1 TO LEN(L(j).@LTxt)                       ' Do it the hard way
                  IF k >= scol AND k <= ecol THEN _               '
                     me.LTxtCharRep(j, k, UUCASE(MID$(LTxtG(j), k, 1)))
               NEXT k                                             '
            ELSE                                                  ' No BNDS, the simple way
               me.LTxtSet(j, UUCASE(LTxtG(j)))                    ' Do the Lowercase if a DataLine
            END IF                                                '
            INCR NumDone                                          '
         END IF                                                   '

      CASE "LC"                                                   ' LowerCase
         IF IsLData(j) THEN                                       ' Only Data lines
            GOSUB DoPMinus                                        ' Do the +/- stuff
            me.LFlagBitOff(j, %EQChange)                          ' Reset ==CHG>
            me.UpdLControl(j)                                     ' And LLCtl
            IF Prf.BndLeft > 1 OR Prf.BndRight > 0 THEN           ' Honour BNDS setting?
               FOR k = 1 TO LEN(L(j).@LTxt)                       ' Do it the hard way
                  IF k >= scol AND k <= ecol THEN _               '
                     me.LTxtCharRep(j, k, LLCASE(MID$(LTxtG(j), k, 1)))
               NEXT k                                             '
            ELSE                                                  ' No BNDS, the simple way
               me.LTxtSet(j, LLCASE(LTxtG(j)))                    ' Do the Lowercase if a DataLine
            END IF                                                '
            INCR NumDone                                          '
         END IF                                                   '

      CASE "SC"                                                   ' Sentence Case
         IF IsLData(j) THEN                                       ' Only Data lines
            GOSUB DoPMinus                                        ' Do the +/- stuff
            me.LFlagBitOff(j, %EQChange)                          ' Reset ==CHG>
            me.UpdLControl(j)                                     ' And LLCtl
            IF Prf.BndLeft > 1 OR Prf.BndRight > 0 THEN           ' Honour BNDS setting?
               Txt1 = LLCASE(MID$(LTxtG(j), scol, ecol - scol + 1))' Get copy of substring to process
               GOSUB DoSChunk                                     ' Go process this piece
            ELSE                                                  ' No BNDS, the simple way
               Txt1 = LLCASE(MID$(LTxtG(j), 1, LEN(L(j).@LTxt)))  ' Get copy of substring to process
               GOSUB DoSChunk                                     ' Go process this piece
            END IF                                                '
            txt2 = LTxtG(j)                                       ' Get whole text line
            MID$(txt2, scol, ecol - scol + 1) = Txt1              ' Stuff back in the processed txt
            me.LTxtSet(j, txt2)                                   ' Replace the whole line
            INCR NumDone                                          '
         END IF                                                   '

      CASE "TC"                                                   ' Title Case
         IF IsLData(j) THEN                                       ' Only Data lines
            GOSUB DoPMinus                                        ' Do the +/- stuff
            me.LFlagBitOff(j, %EQChange)                          ' Reset ==CHG>
            me.UpdLControl(j)                                     ' And LLCtl
            IF Prf.BndLeft > 1 OR Prf.BndRight > 0 THEN           ' Honour BNDS setting?
               me.TitleCase(j, scol, ecol)                        ' Go process it
            ELSE                                                  ' No BNDS, the simple way
               me.TitleCase(j, 1, LEN(L(j).@LTxt))                ' Go process it
            END IF                                                '
            INCR NumDone                                          '
         END IF                                                   '
   END SELECT                                                     '
   me.AttrScan(j)                                                 ' Recolorize

   RETURN

DoPMinus:
   lclFlag = DotFlag                                              ' Get a copy of the +/- flag
   IF BIT(lclFlag, %lCmdX) THEN                                   ' Do the +/- processing for the Source
      me.LFlagBitOn(j, %Invisible)                                ' If - make Invisible
      gfXRebuild = %True                                          ' Do exclude stuff
   ELSEIF BIT(lclFlag, %lCmdNX) THEN                              '
      me.LFlagBitOff(j, %Invisible)                               ' If + make visible
      gfXRebuild = %True                                          ' Do exclude stuff
   END IF                                                         '
   RETURN                                                         '

DoSChunk:
   x = PARSECOUNT(Txt1, " ")                                      ' Count words
   REDIM words(1 TO x) AS STRING                                  ' Make table correct size
   PARSE Txt1, words(), " "                                       ' Parse the words into the table
   m = 1                                                          ' Start scan at 1
   FOR y = 1 TO x                                                 ' Loop for each word
      IF ISNOTNULL(words(y)) THEN                                 ' Ignore null 'words'
         z = INSTR(m, Txt1, words(y))                             ' Find location of this word
         IF caps THEN                                             ' Capitalize this word?
            IF z THEN MID$(Txt1, z, 1) = UUCASE(MID$(Txt1, z, 1)) ' Uppercase the 1st letter
            Caps = %False                                         ' Off till next period
         END IF                                                   '
         IF RIGHT$(words(y), 1) = "."  OR _                       ' Look for sentence trigger
            RIGHT$(words(y), 1) = "?"  OR _                       '
            RIGHT$(words(y), 1) = "!"  THEN                       '
            Caps = %True                                          ' Restart sentence if required
         END IF                                                   '
         m = z + 1                                                ' Start scan adjust
      END IF                                                      '
   NEXT j                                                         '
   RETURN                                                         '
END METHOD

METHOD  pCmdLEFT(pCmd AS STRING)
'---------- Scroll LEFT
LOCAL ScrAmt, DataIX AS LONG, ScrCmd AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN                                         ' No Ops, fudge it
      me.CmdParse("LEFT " + Prf.Scroll)                           ' Re-parse it
   END IF
   ScrCmd = pCmdOps(1)                                            ' Pick up the operand
   IF pCmdOpsType(1) <> %OpNum AND INSTR(CHR$(%KWHALF, %KWPAGE, %KWFULL, %KWDATA, %KWMAX, %KWCSR), CHR$(pCmdOpsType(1))) = 0 THEN _
      MErrExit(%eFail, "Invalid scroll amount")                   '
   IF IsCData THEN                                                ' No, In data area?
      me.CurSetReq(%Position, CsrAbsLine, CCol - gLNPadCol + Offset, %False) ' Set cursor set attempt
   ELSEIF IsCLinN THEN                                            ' No, In Line Number area?
      me.CurSetReq(%Position, CsrAbsLine, 0, %False)              ' Set cursor set attempt
   END IF
   IF VAL(ScrCmd) <> 0 THEN                                       ' A Number?
      ScrAmt = VAL(ScrCmd)                                        ' Yes, use it
   ELSEIF pCmdOpsType(1) = %KWHalf THEN                           ' Half?
      ScrAmt = INT(gDataLen / 2)                                  ' Yes, set it
   ELSEIF pCmdOpsType(1) = %KWPage THEN                           ' Page?
      ScrAmt = gDataLen                                           ' Yes, set it
   ELSEIF pCmdOpsType(1) = %KWFull THEN                           ' Full?
      ScrAmt = gDataLen                                           ' Yes, set it
   ELSEIF pCmdOpsType(1) = %KWData THEN                           ' Data?
      ScrAmt = gDataLen - 1                                       ' Yes, set it
   ELSEIF pCmdOpsType(1) = %KWMax THEN                            ' Max?
      ScrAmt = 9999999                                            ' Yes, set a BIG number
   ELSEIF IsCData THEN                                            ' No, In the Data Area?
      ScrAmt = gDataLen - CCol + gLNPadCol                        ' Yes, move cursor to right of screen
      IF ScrAmt = 0 THEN                                          ' If already there go a default amount
         ScrAmt = gDataLen - 1
         me.CurSetReq(%Position, DataIX, CCol - gLNPadCol + Offset - ScrAmt, %False)
      END IF                                                      '
   ELSEIF IsCLinN THEN                                            ' Line number area
      DataIX = sGetIX(CRow)                                       ' Locate data line cursor was on
      ScrAmt = gDataLen - 1                                       ' Scroll width of screen
   ELSE                                                           '
      ScrAmt = gDataLen - 1                                       ' No, Scroll width of screen - 1
   END IF                                                         '
   Offset = MAX(Offset - ScrAmt, 0)                               '
   ErrFlag = %eNone                                               ' Say we 're OK
   MExit
END METHOD

METHOD  pCmdLINE(pCmd AS STRING)
'---------- Do the LINE command in Primary mode
LOCAL lclCursMode AS INTEGER
LOCAL AllLines, fLine AS LONG
LOCAL OneLine, AllCount, i AS LONG, lcmd, errMsg AS STRING, t AS TouchEntry

   MEntry
   '----- Do the basic parsing stuff
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   lclCursMode = %CursBad                                         ' Make it look like Bad Cursor
   me.ResetFunc(%ResetFind)                                       ' Clear previous FIND hilighting

   '----- Extract the line range operands
   IF ISTRUE me.RangeSet("", %False, %False) THEN MExitMeth       ' Setup any range command operands
   IF DotSet AND ISNULL(DotTag) THEN                              ' If a line range and not a :TAG range
      IF DotStart > 0 AND DotEnd = 0 OR _                         ' If a single line range
         DotStart = DotEnd THEN _                                 '
         OneLine = %True                                          ' Set the single line format flag
   END IF                                                         '

   '----- Now scan out the standard search string stuff
   me.FindReset                                                   ' Reset the parse output fields

   IF sCRTParse("SAD1U") THEN MExitMeth                           ' Scan (Subset, ALL, Direction, Lit1, User)
   me.FindSetup                                                   ' Copy CRT parse's answers

   IF ISFALSE IsCfLit1 OR _                                       ' Lit1 present?
      LEN(cfFind) > 8 THEN _                                      ' and not bigger than 8 chars
      MErrExit(%eFail, "Missing/Invalid Line command")            '
   CfClear(%CRTLit1)                                              ' Now remove Lit1 from the search criteria

   IF ISFALSE DotSet AND _                                        ' If no line range
      ISFALSE IsCfU AND ISFALSE IsCfNU AND ISFALSE IsCfX AND ISFALSE IsCfNX THEN  _' Or U / NU / X / NX
      MErrExit(%eFail, "No line reference or line range specified")

   IF ISFALSE IsCfNext AND ISFALSE IsCfPrev AND _                 '
      ISFALSE IsCfFirst AND ISFALSE IsCfLast THEN                 '
      CfSet(%CRTNext)                                             ' Set NEXT if no others specified
      CfSet(%CRTAll)                                              ' then set ALL
   END IF

   lCmd = cfFind                                                  ' Get local copy of the line command
   IF (LEFT$(lCmd, 1) = "." OR LEFT$(lCmd, 1) = ":") AND _        ' Attempt to tweak label or tag?
      ISFALSE OneLine THEN _                                      ' Better be a single line
      MErrExit(%eFail, "Can only modify Label or Tag for a single line")

   '----- Now validate the command
   t.LinNo = DotStart                                             ' Build a Touch entry, complete line number
   t.LinCtl = lCmd                                                ' Add the specified command
   t.LinFlag = 0                                                  ' Clear the flag

   me.LCtlPreProcessEnt(t)                                        ' Go Preprocess (Normalize) the command

   '----- Now validate it's a valid line command
   IF INSTR(".: ", LEFT$(lCmd, 1)) = 0 THEN                       ' if not a Label or tag manipulation?
      me.LCtlValidateEnt(t, errMsg)                               ' Go validate the command
      IF errMsg <> "OK" THEN _                                    ' If not OK an error occurred
         MErrExit(%eFail, errMsg)                                 ' Issue the message
   END IF                                                         '

   '----- If a single line, just stuff it in
   IF OneLine THEN                                                ' Just a single line reference?
      me.TTblAdd(t.LinNo, lCmd)                                   ' Update touched items
      errFlag = 0                                                 ' Reset errFlag to clear any pending etc.
      me.LCtlProcess                                              ' Go start the ugly Line Command Process
      IF errFlag = %eNone THEN me.LCtlCommands                    ' If no errors, go give them a try
      IF ErrFlag = 0 THEN ErrMsg = ""                             '
      me.UpdLControl(t.LinNo)                                     ' Update LLCtl
      MErrExit(%eNone, $SQ + IIF$(TRIM$(lCmd) = "", " ", TRIM$(lCmd)) + $SQ + " applied")
   END IF                                                         '

   '----- OK lets do the search now for line range stuff
   RESET AllCount                                                 ' Reset variables

   '----- The search
   IF ISFALSE IsCfAll THEN                                        ' Not ALL?
      i = me.Search(0, lclCursMode)                               ' Do the initial search
      IF i THEN                                                   '
         i = cfFLine                                              ' Get found line
         fline = i                                                ' Save first line
         me.TTblAdd(i, lCmd)                                      ' Update touched items
         errFlag = 0                                              ' Reset errFlag to clear any pending etc.
         me.LCtlProcess                                           ' Go start the ugly Line Command Process
         IF errFlag = %eNone THEN me.LCtlCommands                 ' If no errors, go give them a try
         MErrExit(%eNone, $SQ + IIF$(TRIM$(lCmd) = "", " ", TRIM$(lCmd)) + $SQ + " applied")
      ELSE                                                        '
         MErrExit(%eFail, "No eligible line found")               '
      END IF                                                      '

   ELSE                                                           ' This is the ALL version
      i = me.Search(0, lclCursMode)                               ' Do the initial search
      IF i THEN                                                   '
         i = cfFLine                                              ' Get found line
         me.TTblAdd(i, lCmd)                                      ' Update touched items
         INCR AllCount                                            ' Count it
      ELSE                                                        '
         MErrExit(%eFail, "No eligible line found")               '
      END IF                                                      '

      lclCursMode = %CursInternal                                 ' Say we're setting resume point
      DO WHILE i                                                  ' Do till not found
         i = me.Search(1, lclCursMode)                            ' Do re-Finds
         IF i THEN                                                ' Found one
            i = cfFLine                                           ' Get found line
            me.TTblAdd(i, lCmd)                                   ' Update touched items
            INCR AllCount                                         ' Count it
         END IF                                                   '
      LOOP                                                        '
   END IF                                                         '

   errFlag = 0                                                    ' Reset errFlag to clear any pending etc.
   me.LCtlProcess                                                 ' Go start the ugly Line Command Process
   IF errFlag = %eNone THEN me.LCtlCommands                       ' If no errors, go give them a try

   '----- Search over, report the result
   IF AllCount = 0 THEN                                           ' Issue approp. message
      scError(nMac(%eFail), "No lines processed")                 '
   ELSE                                                           '
      scError(0, $SQ + IIF$(TRIM$(lCmd) = "", " ", TRIM$(lCmd)) + $SQ + " applied " + FORMAT$(AllCount) + " time" + IIF$(AllCount > 1, "s", ""))
   END IF                                                         '
   MExit

END METHOD

METHOD  pCmdLOCATE(pCmd AS STRING)
'---------- Move screen to a particular line
LOCAL Wrd, lclCmd AS STRING, i, j, k, llastline, lastcol AS LONG, postype AS INTEGER
   MEntry
   LocFind = "LOCATE"                                             ' Remember what Loc / FIND command was last
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN _                                       ' No Ops
      MErrExit(%eFail, "Missing LOCATE operand")                  ' Oops!
   me.ResetFunc(%ResetFind)                                       ' Clear previous FIND hilighting

   IF IsCData THEN                                                ' In the Data Area
      llastline = sGetIX(CRow)                                    ' Save line the cursor is on
      lastcol = CCol - gLNPadCol + Offset + 1                     ' Calc column position in line
   END IF                                                         '
   IF IsCLinN THEN                                                ' In the Line Area
      llastline  = sGetIX(CRow): lastcol = 0                      ' Swap in current cursor location
   END IF                                                         '

   RESET LocFlag, LocLine, LocLabLine1, LocLabLine2, LocLabel1, LocLabel2, LocTag, LocLength, LocNVar, LocColor
   LocDir = 1                                                     '

   FOR i = 1 TO pCmdNumOps                                        ' Process all our operands
      SELECT CASE pCmdOpsType(i)                                  ' Handle each operands

         CASE %KWNEXT:            BIT SET LocFlag, %LocNext       '
         CASE %KWLAST:            BIT SET LocFlag, %LocLast       '
         CASE %KWPREV:            BIT SET LocFlag, %LocPrev       '
         CASE %KWFIRST:           BIT SET LocFlag, %LocFirst      '
         CASE %KWNOT:             BIT SET LocFlag, %LocNot        '
         CASE %KWALL:             BIT SET LocFlag, %LocAll        '
         CASE %KWMX:              BIT SET LocFlag, %LocMX         '
         CASE %KWTOP:             BIT SET LocFlag, %LocTop        '
         CASE %KWCURRENT:         BIT SET LocFlag, %LocCurrent    '
         CASE %KWCHANGE:          BIT SET LocFlag, %LocChange: INCR k
         CASE %KWFIND:            BIT SET LocFlag, %LocFind: INCR k
         CASE %KWFILE:            BIT SET LocFlag, %LocFile: INCR k
         CASE %KWCOMMAND:         BIT SET LocFlag, %LocCommand: INCR k
         CASE %KWERROR:           BIT SET LocFlag, %LocError: INCR k
         CASE %KWEXCLUDE:         BIT SET LocFlag, %LocExcluded: INCR k
         CASE %KWLABEL:           BIT SET LocFlag, %LocLabel: INCR k
         CASE %KWPAGE:            BIT SET LocFlag, %LocPage: INCR k
         CASE %KWSPECIAL:         BIT SET LocFlag, %LocSpecial: INCR k
         CASE %KWKEEP:            BIT SET LocFlag, %LocKeep: INCR k
         CASE %KWTAG:             BIT SET LocFlag, %LocTag: INCR k
         CASE %KWLONG:            BIT SET LocFlag, %LocLong: INCR k
         CASE %KWSIZE:            BIT SET LocFlag, %LocSize: INCR k
         CASE %KWSTD:             BIT SET LocFlag, %LocStd: INCR k
         CASE %KWU:               BIT SET LocFlag, %LocU: INCR k
         CASE %KWNU:              BIT SET LocFlag, %LocNU: INCR k
         CASE %KWNX:              BIT SET LocFlag, %LocNX: INCR k
         CASE %KWNOTE, %KWANOTE, %KWBNOTE, %KWCNOTE, %KWDNOTE, %KWENOTE, %KWFNOTE, _
              %KWGNOTE, %KWHNOTE, %KWINOTE, %KWJNOTE, %KWKNOTE, %KWLNOTE, %KWMNOTE, _
              %KWNNOTE, %KWONOTE, %KWPNOTE, %KWQNOTE, %KWRNOTE, %KWSNOTE, %KWTNOTE, _
              %KWUNOTE, %KWVNOTE, %KWWNOTE, %KWXNOTE, %KWYNOTE, %KWZNOTE
            BIT SET LocFlag, %LocNote: INCR k
            IF pCmdOpsType(i) <> %KWNote THEN                     ' If not the normal NOTE
               LocNVar = INSTR($upper, UUCASE(LEFT$(pCmdOps(i), 1))) ' Save the Note variety
            END IF                                                '

         CASE %OPTag                                              ' :Tag operand
            IF IsEQ(pCmdOps(i), ":ZALL") OR _                     ' Eliminate ZALL
               IsEQ(pCmdOps(i), ":Z") OR _                        ' Eliminate Z
               IsEQ(pCmdOps(i), ":ZF") OR _                       ' Eliminate ZF
               IsEQ(pCmdOps(i), ":ZNF") THEN _                    ' Eliminate ZNF
               MErrExit(%eFail, "Invalid Tag operand")            '
            IF sTagVal(pCmdOps(i), %False) THEN _                 ' See if 1st is a valid Tag
               MErrExit(%eFail, "Invalid Tag operand")            '
            LocTag = LSET$(UUCASE(pCmdOps(i)), 8)                 '
            INCR k                                                '

         CASE %OpDotd, %OpNum                                     ' A .Label or simple number
            IF TRIM$(pCmdOps(i)) = "." THEN _                     ' just a dot?
               MErrExit(%eFail, "Line reference " + pCmdOps(i) + " is invalid") ' Oops!
            j = me.LineNoRef(pCmdOps(i))                          ' Get the line number
            IF j = -1 AND ISFALSE BIT(LocFlag, %LocSize) THEN _   ' If invalid
               MErrExit(%eFail, "Line reference " + pCmdOps(i) + " is undefined") ' Oops!
            IF ISNULL(LocLabel1) THEN                             ' First one?
               LocLabel1 = LSET$(UUCASE(pCmdOps(i)), 8)           '
               LocLabLine1 = j                                    '
            ELSEIF ISNULL(LocLabel2) THEN                         ' Second one?
               LocLabel2 = LSET$(UUCASE(pCmdOps(i)), 8)           '
               LocLabLine2 = j                                    '
            ELSE                                                  '
               MErrExit(%eFail, "Third line reference detected - " + pCmdOps(i)) ' Oops!
            END IF                                                '

         CASE %OpLPtr                                             ' A !LPtr?
            IF TRIM$(pCmdOps(i)) = "!" THEN _                     ' just a !?
               MErrExit(%eFail, "Line reference: " + pCmdOps(i) + " is invalid") ' Oops!
            j = VAL(MID$(pCmdOps(i), 2))                          ' Get the line number
            IF j = 0 OR j > LastLine THEN _                       ' If invalid
               MErrExit(%eFail, "Line reference: " + pCmdOps(i) + " is invalid") ' Oops!
            IF ISNULL(LocLabel1) THEN                             ' First one?
               LocLabel1 = LSET$(UUCASE(pCmdOps(i)), 8)           '
               LocLabLine1 = j                                    '
            ELSEIF ISNULL(LocLabel2) THEN                         ' Second one?
               LocLabel2 = LSET$(UUCASE(pCmdOps(i)), 8)           '
               LocLabLine2 = j                                    '
            ELSE                                                  '
               MErrExit(%eFail, "Third line reference detected - " + pCmdOps(i)) ' Oops!
            END IF                                                '

         CASE ELSE
            '----- Last chance, is it a color?
            ARRAY SCAN nHiLites(), COLLATE UCASE, = pCmdOps(i), TO j ' See if a HiLite Name
            IF j THEN                                             ' A winner
               IF LocColor <> 0 THEN                              ' Already have one?
                  MerrExit(%eFail, "Extra color name operand detected - " + pCmdOps(i)) ' Issue error
               ELSE                                               '
                  LocColor = j                                    ' Save color search
                  BIT SET LocFlag, %LocColor: INCR k
               END IF
            ELSE
               MErrExit(%eFail, "Unknown LOCATE operand - " + pCmdOps(i)) ' Oops!
            END IF
      END SELECT                                                  '
   NEXT i                                                         '

   IF BIT(LocFlag, %LocTop) AND _                                 ' TOP only for simple positions
      (ISNOTNULL(LocLabel2) OR _                                  '
       BIT(LocFlag, %LocMX) OR _                                  '
       k > 1) THEN _                                              '
      MErrExit(%eFail, "TOP only available with simple position requests")

   IF BIT(LocFlag, %LocNot) AND BIT(LocFlag, %LocCurrent) THEN _  ' Filter some obvious errors
      MErrExit(%eFail, "Can't use both NOT and CURRENT")          '
   IF BIT(LocFlag, %LocAll) AND BIT(LocFlag, %LocFile) THEN _     '
      MErrExit(%eFail, "Can't use both ALL and FILE")             '
   IF BIT(LocFlag, %LocLong) AND BIT(LocFlag, %LocSize) THEN _    '
      MErrExit(%eFail, "Can't use both SIZE and LONG")            '
   IF BIT(LocFlag, %LocNext) AND BIT(LocFlag, %LocPrev) THEN _    '
      MErrExit(%eFail, "Can't use both PREV and NEXT")            '
   IF BIT(LocFlag, %LocFirst) AND BIT(LocFlag, %LocLast) THEN _   '
      MErrExit(%eFail, "Can't use both FIRST and LAST")           '
   IF BIT(LocFlag, %LocAll)  AND _                                '
      (BIT(LocFlag, %LocFirst) OR BIT(LocFlag, %LocLast) OR _     '
       BIT(LocFlag, %LocPrev) OR BIT(LocFlag, %LocNext)) THEN _   '
      MErrExit(%eFail, "Can't use ALL with FIRST/LAST/PREV/NEXT") '
   IF k > 1 THEN _                                                ' Only one line type please
      MErrExit(%eFail, "Only 1 (one) line type can be requested") '
   IF k = 0 AND ISNULL(LocLabel1) THEN _                          ' Nothing to do
      MErrExit(%eFail, "No line search parameters present")       '
   IF BIT(LocFlag, %LocCurrent) AND ISFALSE BIT(LocFlag, %LocFind) AND ISFALSE BIT(LocFlag, %LocChange) THEN _
      MErrExit(%eFail, "CURRENT only allowed with FIND or CHANGE")'

   IF (ISNOTNULL(LocLabel1) AND ISNOTNULL(LocLabel2) AND k = 0) OR _  ' Range but nothing else?
      (BIT(LocFlag, %LocAll) AND k = 0) THEN _                    ' All but nothing else?
      MErrExit(%eFail, "No, line type criteria specified")        '

   IF ISNOTNULL(LocLabel1) AND ISNULL(LocLabel2) AND k = 0 THEN   ' Just a simple Loc .A type?
      PosType = IIF(BIT(LocFlag, %LocTop), %FLocate, %Locate)     ' Set position type based on TOP
      IF IsLInvisible(LocLabLine1) THEN                           ' Is this an X'd out line?
         LocLabLine1 = me.LPopOutX(LocLabLine1)                   ' Yes, Uugh! go do lots of work, get adjusted line number
      END IF                                                      '
      IF BIT(LocFlag, %LocMX) THEN                                ' MX requested?
         me.LFlagBitOn(LocLabLine1, %Invisible)                   ' Just make it invisible
         gfXRebuild = %True                                       ' for rebuild exclude state
         me.CurSetReq(%Find, LocLabLine1, 0, %True, %True)        ' Put cursor on the X line, no hilite, XPtr to be set
         CsrLinDX = VAL(LLNumG(LocLabLine1))                      ' Save hidden line number for Status Bar display
      ELSE                                                        '
         IF BIT(LocFlag, %LocTop) THEN                            ' TOP requested?
            me.CurSetReq(postype, LocLabLine1, 0, %True)          ' Put the line at the top
         ELSE                                                     '
            me.CurSetReq(postype, LocLabLine1, 0, %True)          ' Do normal set cursor set attempt
         END IF                                                   '
      END IF                                                      '
      ErrFlag = %eNone                                            ' Say we 're OK
      MExitMeth                                                   '
   END IF                                                         '

   IF BIT(LocFlag, %LocLong) THEN                                 ' LOC LONG type?
      IF ISNOTNULL(LocLabel1) AND ISNULL(LocLabel2) THEN          ' Just LOC LONG nnn type?
         LocLength = VAL(LocLabel1)                               '
      ELSE                                                        '
         IF Prf.LRECL > 0 THEN LocLength = Prf.LRECL              ' Swap LRECL if present
      END IF
      IF LocLength = 0 THEN _                                     '
         MErrExit(%eFail, "Using LONG requires LRECL > 0 or a length operand")
   END IF                                                         '

   IF BIT(LocFlag, %LocSize) THEN                                 ' LOC SIZE type?
      IF ISNOTNULL(LocLabel1) AND ISNULL(LocLabel2) THEN          ' Just LOC SIZE nnn type?
         LocLength = VAL(LocLabel1)                               '
      ELSE                                                        '
         IF Prf.LRECL > 0 THEN LocLength = Prf.LRECL              ' Swap LRECL if present
      END IF
   END IF                                                         '

   IF BIT(LocFlag, %LocPage) THEN                                 ' LOC PAGE type?
      IF ISNOTNULL(LocLabel1) AND ISNULL(LocLabel2) THEN          ' Just LOC Page nnn type?
         LocLength = VAL(LocLabel1)                               ' Save page number in LocLength
         BIT SET LocFlag, %LocFirst                               ' Start at line 1
      ELSE                                                        '
         LocLength = 0                                            ' Zero page number
      END IF
   END IF                                                         '

   IF BIT(LocFlag, %LocMX) THEN                                   ' LOC MX type?
      IF ISFALSE BIT(LocFlag, %LocAll) THEN _                     ' Must have ALL
         MErrExit(%eFail, "MX allowed only with ALL")             '
   END IF                                                         '

   LocLine = TopScrn                                              ' In case no other shows up
   IF ISFALSE BIT(LocFlag, %LocFirst) AND ISFALSE BIT(LocFlag, %LocLast) AND ISFALSE BIT(LocFlag, %LocPrev) THEN
      BIT SET LocFlag, %LocNext                                   ' If nothing else, set NEXT
   END IF                                                         '

   IF ISFALSE me.LocSearch() THEN                                 ' Go do the full type search
      scError(nMac(%eFail), IIF$(locDir = 1, "Bottom of data reached", "Top of data reached"))
      LocNotFound = IIF(BIT(LocFlag, %LocNext), 2, 1)             ' Remember the failure
      LocNFDir = LocDir                                           ' and direction
      IF llastline <> 0 THEN _                                    ' A saved cursor location?
         me.CurSetReq(%Find, llastline, lastcol, %False, %False)  ' Set cursor
   ELSE                                                           '
      IF IsMEdit AND BIT(LocFlag, %LocFile) THEN                  ' A MEdit LOC FILE command?
         scError(%eNone, "File " + FORMAT$(LMIXG(LocLineR)) + " of " + FORMAT$(MEditCount) + " found")
      END IF                                                      '
   END IF                                                         '
   MExit
END METHOD

METHOD   pCmdLRECL(pCmd AS STRING)
'---------- Set LRECL for fixed length records
LOCAL Wrd1, lclCmd AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN _                                       ' If no operands
      MErrExit(%eNone, "LRECL set to " + FORMAT$(Prf.LRECL))      '
   IF pCmdOpsType(1) <> %OpNum THEN _                             '
      MErrExit(%eFail, "Invalid LRECL operand")                   '
   IF VAL(pCmdOps(1)) = 0 AND IsEQ(Prf.EOL, "NONE") THEN _        ' Better not be EOL=NONE
      MErrExit(%eFail, "LRECL = 0 not valid when EOL = NONE")     '
   IF VAL(pCmdOps(1)) = 0 AND Prf.RECFM = "F" THEN _              ' Better not be RECFM F
      MErrExit(%eFail, "LRECL = 0 not valid when RECFM = F")      '
   IF VAL(pCmdOps(1)) > 0 AND Prf.RECFM <> "F" THEN _             ' Better be RECFM F
      MErrExit(%eFail, "LRECL = " + FORMAT$(VAL(pCmdOps(1))) + " not valid when RECFM = " + Prf.RECFM) '

   Prf.LRECL = VAL(pCmdOps(1))                                    ' Set it
   scError(%eNone, "LRECL set to " + FORMAT$(Prf.LRECL))          '
   MExit
END METHOD

METHOD  pCmdMACRO(pCmd AS STRING)
'---------- Add macro commands to the CmdStack
LOCAL macline, t, tt, subname AS STRING
LOCAL i, j, k, m, n, lno, NumLines, mOps AS LONG
LOCAL ONCE AS INTEGER
LOCAL mcIO AS iIO                                                 ' For our I/O stuff

   MEntry

   IF ErrFlag = %ePending THEN MExitMeth                          ' No macros with line commands pending
   IF gMacroMode THEN _                                           ' Can't already be in macro mode
      MErrExit(%eFail, "Already in macro mode, no nesting of macros supported")
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   LET mcIO = CLASS "cIO"                                         '
   gMacroName = pCmdOps(0)                                        ' Get the macro name
   gMacroFile = ENV.MacrosPath + gMacroName + ".MACRO"            ' Make full name of macro file

   mcIO.Setup("IE", "", "", gMacroFile)                           ' Tell IO what we're opening
   IF mcIO.EXEC THEN _                                            ' Go OPEN the file
      MErrExit(%eFail, "Missing " + gMacroName + ".MACRO file")   ' Oops! Should never happen but ...

   '----- Get number of lines in the Macro
   FILESCAN # mcIO.FNum, RECORDS TO NumLines                      ' allow extra line for #INCLUDEDIR
   REDIM gMaclines(1 TO NumLines + 2) AS STRING                   ' Save away the macro in an array

   DO WHILE ISFALSE EOF(# mcIO.FNum)                              ' Read the data
      LINE INPUT# mcIO.FNum, macline                              '
      INCR lno                                                    ' Line number
      gMacLines(lno) = macline
   LOOP
   INCR lno                                                       ' Add a trailing STOP line
   gMacLines(lno) = "STOP"                                        '
   mcIO.Close                                                     ' Close the file

   '----- Check for macro ID line for Basic
   macline = UUCASE(TRIM$(gmaclines(1))) + " "                    ' Get the 1st line
   IF LEFT$(macline, 1) = "'" THEN                                ' A Basic comment line?
      i = INSTR(macline, "MACRO ")                                ' Find MACRO
      IF i THEN                                                   ' Got it
         macline = MID$(macline, i + 6)                           ' Extract prototype operands
         GOSUB BuildOps                                           ' Process the operands
         GOSUB CheckSUB                                           ' Check for 1st line as a SUB
         MB_Call()                                                ' Pass off to the MACRO processor
         MexitMeth                                                ' We're done here
      END IF                                                      '
   END IF                                                         '

   MErrExit(%eFail, "Missing / Invalid Macro header line")        ' Oops

CheckSub:
   '----- See if a SUB type macro
   i = 2                                                          ' Point at line 2
   macline = UUCASE(TRIM$(gmaclines(i)))                          ' Get the 2nd line and it's continuations
   DO WHILE RIGHT$(macline, 1) = "_"                              ' Add continuations if present
      INCR i                                                      ' Point at next line
      macline = LEFT$(macline, LEN(macline) -1) + " " + UUCASE(TRIM$(gmaclines(i)))
   LOOP                                                           '
   IF LEFT$(macline, 4) <> "SUB " THEN RETURN                     ' Nothing to do

   '----- We have a SUB subname(arg [as string], arg [as string], ...) statement
   i = INSTR(macline, "(")                                        ' Find opening (
   IF i = 0 THEN GOTO SubErr                                      ' None?
   i = INSTR(i + 1, macline, ")")                                 ' Find closing )
   IF i = 0 THEN GOTO SubErr                                      ' None?
   k = 0                                                          ' Number of operands
   i = INSTR(macline, "(")                                        ' Find opening (
   subname = TRIM$(MID$(macline, 4 TO i - 1))                     ' Extract Subname

   DO WHILE INSTR(i + 1, macline, ANY ",)")                       ' look for , or )
      j = INSTR(i + 1, macline, ANY ",)")                         ' Look for delimiter
      t = MID$(macline, i + 1 TO j - 1)                           ' Extract the operand
      tt = sGetWord(t, %Strip, %QuoteNotSig)                      ' Get the operand name
      IF tt = "" THEN GOTO SubErr                                 ' Oops!
      tt = sGetWord(t, %Strip, %QuoteNotSig)                      ' Get the possible 'AS'
      IF tt <> "AS" THEN GOTO SubErr                              '
      tt = sGetWord(t, %Strip, %QuoteNotSig)                      ' Get the possible 'AS'
      IF tt <> "STRING" THEN GOTO SubErr                          '
      INCR k                                                      ' Count an operand
      IF MID$(macline, j, 1) = ")" THEN EXIT DO                   ' Terminated by ) ?
      i = j                                                       ' No, Continue scan
   LOOP                                                           ' We now have # arguments

   macline = subname + "("                                        ' Build the MACRO call
   FOR i = 1 TO k                                                 ' Pick up the operands
      IF LEFT$(mOprands(i), 1) <> $DQ THEN                        ' Insert the operands
         macline += $DQ + mOprands(i) + $DQ                       ' Quoted, if not already
      ELSE                                                        '
         macline += $DQ+ $DQ + mOprands(i) + $DQ + $DQ            ' Quoted, add lots more to make it work
      END IF                                                      '
      IF i < k THEN macline += ", "                               ' Add separater
   NEXT i                                                         '
   macline += "): STOP"                                           ' Close it off
   gMacLines(1) = macline                                         ' Replace line 1
   RETURN
SubErr:
   MErrExit(%eFail, "Invalid syntax in the initial SUB statement")' Oops

BuildOps:
  '----- Process the macro operands
   RESET mOprands()                                               ' Reset the Operand table
   DO WHILE ISNOTNULL(TRIM$(macline))                             ' Loop through operands
      INCR mOps                                                   ' Incr count
      mOprands(mOps) = sGetWord(macline, %Strip, %QuoteSig)       ' Get the next operand from string
   LOOP                                                           '

   '----- Overlay prototype operands with real command line operands
   FOR i = 1 TO pCmdNumOps                                        ' Additional operands from the user?
      mOprands(i) = pCmdRaw(i)                                    ' Use it instead
   NEXT i                                                         '

   '----- Do SET substitution on the operands now
   FOR i = 1 TO 50                                                ' Do SET substitution
      IF ISNOTNULL(mOprands(i)) THEN _                            ' If non-null operand
         IF me.MacSubst(mOprands(i)) THEN _                       ' Do substitution
            MErrExit(%eFail, "Malformed macro operand in MACRO prototype: " + mOprands(i)) ' Sorry
   NEXT i                                                         '

   RETURN
END METHOD

METHOD  pCmdMARK(pCmd AS STRING)
'---------- Turn MARK ON or OFF
LOCAL nv AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   nv = sOnOff(1)                                                 ' Go get True/False value for operand 1
   IF nv < 0 THEN MExitMeth                                       ' Error, exit, error msg already issued
   Prf.PMark = nv                                                 ' Set the new value
   scError(0, "MARK set to " + IIF$(nv, "ON", "OFF"))             ' Tell user new value
   MExit
END METHOD

METHOD  pCmdMEDIT(pCmd AS STRING)
'---------- Multi-Edit startup
   sDoMEDIT(pCmd)                                                 ' Must do globally because of Tab additions
END METHOD

METHOD  pCmdMINLEN(pCmd AS STRING)
'---------- Alter the MINLEN setting
LOCAL lCmd AS STRING, OldMINLEN AS LONG
REGISTER i AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps > 0 THEN                                         ' If operands
      IF pCmdOpsType(1) <> %OpNum OR _                            '
         VAL(pCmdOps(1)) < 0 THEN _                               '
         MErrExit(%eFail, "Invalid MINLEN operand")               '
      OldMINLEN = Prf.MINLEN                                      ' Save current MINLEN
      Prf.MINLEN = VAL(pCmdOps(1))                                ' Set it

      '----- Adjust lines
      IF Prf.MINLEN > OldMINLEN THEN                              ' Increasing MINLEN value
         lCmd = "PL 1 " + FORMAT$(LastLine) + " 0 0 "             ' Build internal PL command line
         lCmd = lCmd + lCmd                                       '
         lCmdPL(lCmd)                                             ' Go do a PL/ command for all lines
      END IF                                                      '
   END IF                                                         '

   scError(0, "MINLEN set to " + FORMAT$(Prf.MINLEN))             '
   MExit                                                          '
END METHOD

METHOD  pCmdNFIND(pCmd AS STRING)
'---------- NFIND (Negative Find) Command
LOCAL lclCmd, lclWord, modemsg, str AS STRING
LOCAL i AS LONG, AllCount AS LONG, lclCursMode AS INTEGER
LOCAL lclCol, ldir, lsline AS LONG                                ' Local saves for FIND ALL
   MEntry
   LocFind = "FIND"                                               ' Remember what Loc / FIND command was last
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   lclCursMode = Zone                                             ' Save CursMode at start of command
   me.ResetFunc(%ResetFind)                                       ' Clear previous FIND hilighting
   IF ISTRUE me.RangeSet("", %False, %False) THEN MExitMeth       ' Setup any range command operands
   me.FindReset                                                   ' Reset the parse output fields
   IF sCRTParse("ASDMEC1PU") THEN MExitMeth                       ' Scan (All, Subset, Direct, Modifier, Exclude, Cols, Lit1, Clr, User)
   me.FindSetup                                                   ' Copy CRT parse's answers
   IF ISFALSE IsCfNext AND ISFALSE IsCfPrev AND _                 '
      ISFALSE IsCfFirst AND ISFALSE IsCfLast THEN _               '
      CfSet(%CRTNext)                                             ' Set NEXT if no others specified
      CfSet(%CRTNF)                                               ' Force this to NF
   IF ISFALSE IsCfLit1 THEN _                                     ' Lit1 present?
      MErrExit(%eFail, "Missing NFIND search string")             '

   IF LEN(CrtL1Raw) < 16 THEN                                     ' Build msg literal
      str = CrtL1Raw                                              '
   ELSE                                                           '
      str = LEFT$(CrtL1Raw, 14) + ".."                            '
   END IF                                                         '

   modemsg = IIF$(IsCfWord, "WORD " + str, "CHARS " + str)        '

   IF IsCfAll THEN                                                ' The ALL version?
      AllCount = 0                                                ' Reset variables
      i = me.NFSearch(0, lclCursMode)                             ' Do the initial search
      lclCursMode = %CursBad                                      ' Make it look like Bad Cursor
      IF i THEN                                                   '
         i = sLine                                                ' Get found line
         INCR AllCount: lclCol = sCol: lDir = sDir: lsLine = sLine' Count and save 1st located for RFIND stuff
      END IF                                                      '
      DO WHILE i                                                  ' Do till not found
         i = me.NFSearch(1, lclCursMode)                          ' Do re-Finds
         IF i THEN                                                ' Found one
            i = sLine                                             ' Get found line
            INCR AllCount                                         ' Count successful
         END IF                                                   '
      LOOP                                                        '
      IF AllCount = 0 THEN                                        ' Issue approp. message
         scError(%eFail, modemsg + " found on all lines")         '
      ELSE                                                        '
         scError(0, modemsg + " not found in " + FORMAT$(AllCount) + IIF$(AllCount > 1, " lines", " line"))
         sCol = lclCol: sDir = lDir: sLine = lsLine               ' Restore to 1st one found
      END IF                                                      '
   ELSE                                                           ' The non-All version
      i = me.NFSearch(0, lclCursMode)                             ' Go see if we can find it
      IF i = 0 THEN                                               ' Got it?
         scError(nMac(%eFail), IIF$(IsCfNext, "Bottom of data reached", "Top of data reached"))
         cfNotFound = IIF(IsCfNext, 2, 1)                         ' Remember we failed
      ELSE                                                        '
         IF IsCfMX THEN                                           ' X'ing the line?
            scError(%eNone, modemsg + " not found and X'd")       '
         ELSE                                                     '
            scError(%eNone, modemsg + " not found")               '
         END IF                                                   '
         ErrFlag = %eNone                                         ' Say we 're OK
      END IF                                                      '
   END IF                                                         '
   MExit
END METHOD

METHOD  pCmdNotify(pCmd AS STRING)
'---------- Switch Notify level
LOCAL Wrd, lclCmd AS STRING, i AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps > 0 THEN                                         ' Got an Operand?
      IF INSTR(CHR$(%KWNONE, %KWEDIT, %KWALL, %KWRESET), CHR$(pCmdOpsType(1))) = 0 THEN _
         MErrExit(%eFail, "Invalid operand for NOTIFY")           '
      SELECT CASE pCmdOpsType(1)                                  ' Split off by ON/OFF
         CASE %KWNONE:  ENV.NotifyLevelT = 0                      ' Set Notify to None
         CASE %KWEDIT:  ENV.NotifyLevelT = 1                      ' Set Notify to Edit
         CASE %KWALL:   ENV.NotifyLevelT = 2                      ' Set Notify to All
         CASE %KWRESET: ENV.NotifyLevelT = ENV.NotifyLevel        ' Get the default user's value
      END SELECT                                                  '
   END IF                                                         '
   scError(%eNone, "NOTIFY set to " + SWITCH$(ENV.NotifyLevelT = 0, "None", ENV.NotifyLevelT = 1, "Edit", ENV.NotifyLevelT = 2, "All"))
   MExit
END METHOD

METHOD  pCmdNumber(pCmd AS STRING)
'---------- Do the NUMBER or RENUM command
LOCAL x, xList, NumType, prevseq, t, newnum AS STRING
LOCAL i, j, k, seqCol, seqLen, wSeqCol, trunc, real, seqerror, numfirst, numincr AS LONG
LOCAL LastGoodLineNo, LastGoodSeqNo, BadLineCount AS LONG
   MEntry

   '----- See if a valid NUMTYPE exiats
   NumType = TRIM$(Prf.NumType)                                   ' Get local copy of NumType
   IF ISNULL(NumType) THEN                                        ' Is NUMTYPE null?
      MErrExit(%eFail, "There is no valid NUMTYPE available")     ' Oops!
   END IF                                                         '

   '----- Make sure the file extension is an editable file
   x = "," + UUCASE(MID$(TP.TIPExtn, 2)) + ","                    ' Get extension surrounded by commas
   XList = "," + UUCASE(ENV.FMFFList) + ","                       ' Build XList here since a ProfSkip might have added an entry
   XList = REMOVE$(XList, " ")                                    ' Eliminate any spaces
   IF INSTR(XList, x) THEN                                        ' Is tis file in the exempt list?
      MErrExit(%eFail, "Current filetype is in the non-editable filetype list") ' Oops!
   END IF

   '----- See if ragged right columns
   IF INSTR(NumType, "-") THEN                                    ' Negative columns?
      i = sDoMsgBox("With the current RECFM (" + Prf.RECFM + ")," + $CRLF + _
                 "the specified NUMTYPE (" + NumType + ") could result in" + $CRLF + _
                 "sequence numbers in differing columns.  Do you want to continue?", _
                 %MB_YESNO OR %MB_USERICON OR %MB_DEFBUTTON2, "SPFLite")
      IF i = %IDNO THEN                                           '
         MErrExit(%eNone, "Numbering cancelled")                  ' Exit if NO entered
      END IF                                                      '
   END IF                                                         '

   '----- Setup the sequence columns
   seqCol = VAL(PARSE$(NumType, " ", 1))                          ' Extract start column
   seqLen = VAL(PARSE$(NumType, " ", 2))                          ' Extract end column
   seqLen = ABS(ABS(seqLen) - ABS(seqCol)) + 1                    ' Calc length of sequence field
   IF INSTR(NumType, "TRUNC") THEN trunc = %True                  ' Extract TRUNC
   IF INSTR(NumType, "REAL") THEN real = %True                    ' and REAL flags
   prevseq = REPEAT$(seqLen, " ")                                 ' Initialize previous sequence

   '----- See if the file is already properly numbered
   FOR i = 1 TO LastLine                                          ' Spin through the lines
      IF IsLData(i) THEN                                          ' Only Data lines
         GOSUB GetLineI                                           ' Get Line into t and set WSeqCol

         '----- Prep by making all lines long enough
         IF wSeqCol + seqLen -1 > LEN(t) THEN                     ' If not long enough for a Seq field
            t = LSET$(t, wSeqCol + seqLen -1)                     ' Lengthen line if too short
            me.LTxtSet(i, t)                                      ' Stuff it back
         END IF                                                   '

         '----- Check sequence
         IF VERIFY(MID$(t, wSeqCol, seqLen), $Numeric) <> 0 THEN  ' If not numeric
            seqerror = %True                                      '
         ELSEIF prevseq >= MID$(t, wSeqCol, seqLen) THEN          ' If not > previous sequence field
            seqerror = %True                                      ' Then we have an error
            prevseq = MID$(t, wSeqCol, seqLen)                    ' Save new sequence field
         ELSE                                                     '
            prevseq = MID$(t, wSeqCol, seqLen)                    ' Save new sequence field
         END IF                                                   '
      END IF                                                      '
   NEXT i                                                         '

   IF ISFALSE seqerror THEN _                                     '
      MErrExit(%eNone, "Entire file is already in NUMBER sequence") ' Say all is well

   '----- Try to renumber, keeping what's possible
   LastGoodSeqNo = 0: LastGoodLineNo = 0                          ' Set initial values
   FOR i = 1 TO LastLine                                          ' Spin through the lines
      IF IsLData(i) THEN                                          ' Only Data lines
         GOSUB GetLineI                                           ' Get Line into t and set WSeqCol

         IF VERIFY(MID$(t, wSeqCol, seqLen), $Numeric) <> 0 THEN  ' If not numeric
            INCR BadLineCount                                     ' No good, Count lines needing to be 'fixed'
            ITERATE FOR                                           '
         END IF                                                   '
         IF VAL(MID$(t, wSeqCol, seqLen)) <= LastGoodSeqNo THEN   ' If sequence not higher that previous good sequence
            INCR BadLineCount                                     ' No good, Count lines needing to be 'fixed'
            ITERATE FOR                                           '
         END IF                                                   '
         IF VAL(MID$(t, wSeqCol, seqLen)) <= LastGoodSeqNo + BadLineCount THEN ' If number less than # bad preceeding lines
            INCR BadLineCount                                     ' No good, Count lines needing to be 'fixed'
            ITERATE FOR                                           '
         END IF                                                   '

         '----- We have a new valid sequence number
         IF BadLineCount > 0 THEN                                 ' We have some fixup to perform on badlines
            GOSUB CorrectBadOnes                                  ' Go do so
            GOSUB GetLineI                                        ' Re-get the original i version of the line
         END IF                                                   '
         LastGoodSeqNo = VAL(MID$(t, wSeqCol, seqLen))            ' Save as last good sequence number
         LastGoodLineNo = i                                       ' Save as last good line number
         BadLineCount = 0                                         ' Reset the Bad line count
      END IF                                                      '
   NEXT i                                                         '

   '----- See if we ended up 'clean'
   IF BadLineCount > 0 THEN                                       ' Final stretch a bunch of Bad lines?
      t = SPACE$(wSeqCol + seqLen)                                ' Build a dummy line
      k = LastGoodSeqNo + (100 * (BadLineCount + 1))              ' Calculate a pseudo lastline sequence number                              '
      MID$(t, wSeqCol, seqLen) =  FORMAT$(k, REPEAT$(seqlen, "0"))' Format it into the dummy line
      GOSUB CorrectBadOnes                                        ' Go correct the bad ones
   END IF
   MExitMeth

GetLineI:
   t = LTxtG(i)                                                   ' Get the text line
   wSeqCol = seqCol                                               ' Default sequence column
   IF seqCol < 0 THEN wSeqCol =  LEN(t) - seqCol + 1              ' If negative columns, make it an actual
   RETURN
GetLineJ:
   t = LTxtG(j)                                                   ' Get the text line
   wSeqCol = seqCol                                               ' Default sequence column
   IF seqCol < 0 THEN wSeqCol =  LEN(t) - seqCol + 1              ' If negative columns, make it an actual
   RETURN
CorrectBadOnes:
   k = VAL(MID$(t, wSeqCol, seqLen)) - LastGoodSeqNo              ' Calc the increment to use
   numincr = k / (BadLineCount + 1)                               '
   numfirst = LastGoodSeqNo + numincr                             ' Calc seq number for the first bad line
   FOR j = LastGoodLineNo + 1 TO i - 1                            ' Loop through the bad ones
      IF IsLData(j) THEN                                          ' Only the data lines
         GOSUB GetLineJ                                           ' Go get a bad line
         newnum = FORMAT$(NumFirst, REPEAT$(seqlen, "0"))         ' Format the new number
         MID$(t, wSeqCol, seqLen) = newnum                        ' Stuff it in the line
         IF Trunc THEN t = LEFT$(t, wSeqCol + seqLen - 1)         ' If trunc, do so
         me.LTxtSet(j, t)                                         ' Stuff it back
         me.AttrScan(j)                                           ' Recolorize
         NumFirst += NumIncr                                      ' Incr for next bad line
      END IF                                                      '
   NEXT j                                                         '
   RETURN
END METHOD

METHOD  pCmdNumType(pCmd AS STRING)
'---------- Validate the NUMTYPE command
LOCAL i, j, trunc, real, offReq, fromcol, tocol AS LONG, NumType AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps <> 0 THEN                                        ' No Ops, kill it

      IF pCmdNumOps = 1 THEN                                      ' Just one operand?
         SELECT CASE uucase$(pCmdOps(1))                          ' See which one we've got
            CASE "COB":            pCmd = "NUMTYPE 1 6"           ' Convert the recognised aliases
            CASE "STD":            pCmd = "NUMTYPE 73 80"         '
            CASE "IBM":            pCmd = "NUMTYPE 73 80 TRUNC"   '
            CASE "NONE":           pCmd = "NUMTYPE NONE"          '
            CASE ELSE                                             '
               MErrExit(%eFail, "Unknown NUMTYPE operand - " + pCmdOps(1)) ' Oops!
         END SELECT                                               '
      END IF                                                      '
      IF me.CmdParse(pCmd) THEN MExitMeth                         ' Re-parse the substituted string

      '----- We have a multi-operand NUMTYPE command
      '----- Since negative cols appear as strings, we'll treat them ALL as strings
      FOR i = 1 TO pCmdNumOps                                     ' Loop through them
         IF uucase$(pCmdOps(i)) = "TRUNC" THEN                    ' TRUNC?
            trunc = %True                                         ' Remember that
         ELSEIF uucase$(pCmdOps(i)) = "REAL" THEN                 ' REAL?
            real = %True                                          ' Remember that
         ELSEIF uucase$(pCmdOps(i)) = "NONE" THEN                 ' NONE
            OffReq = %True                                        ' Remember that
         ELSEIF VAL(pCmdOps(i)) <> 0 THEN                         ' Numeric?
            IF fromcol = 0 THEN                                   ' No fromcol yet
               fromcol = VAL(pCmdOps(i))                          ' Save it as fromcol
            ELSEIF tocol = 0 THEN                                 ' No tocol yet?
               tocol = VAL(pCmdOps(i))                            ' Save it as tocol
            ELSE
               MErrExit(%eFail, "NUMTYPE excessive numeric operand - " + pCmdOps(1)) ' Oops!
            END IF                                                '
         ELSE                                                     '
            MErrExit(%eFail, "NUMTYPE unknown operand - " + pCmdOps(1)) ' Oops!
         END IF                                                   '
      NEXT i                                                      '

      '----- Do some other validation
      IF ISFALSE offReq THEN                                      ' Other than OFF
         IF fromcol = 0 OR tocol = 0 THEN _                       ' Got column operands?
            MErrExit(%eFail, "NUMTYPE needs both a from-col and a to-col operand") ' Oops!

         IF SGN(fromcol) <> SGN(tocol) THEN _                     ' Look for illogical values
            MErrExit(%eFail, "NUMTYPE specifies an invalid column range") ' Oops!
         '----- Test +ve values
         IF SGN(fromcol) = 1 THEN                                 ' +ve numbers
            IF fromcol >= tocol OR  _                             ' Left/Right not OK
            tocol - fromcol < 3 OR _                              ' Length not OK
            tocol - fromcol > 7 THEN _                            '
            MErrExit(%eFail, "NUMTYPE specifies an invalid column range") ' Oops!

         ELSE                                                     ' -ve numbers
            j = fromcol - tocol
            IF fromcol >= tocol OR  _                             ' Left/Right not ok
            fromcol - tocol > -3 OR _                             ' Length not OK
            fromcol - tocol < -7 THEN _                           '
            MErrExit(%eFail, "NUMTYPE specifies an invalid column range") ' Oops!
         END IF                                                   '
      END IF

      '----- Now build a valid NUMTYPE string in a standard order
      IF offReq THEN                                              ' If OFF, clear NumType
         NumType = ""                                             '
      ELSE                                                        '
         NumType = FORMAT$(fromcol) + " " + FORMAT$(tocol) + " " + IIF$(trunc, "TRUNC ", "") + IIF$(real, "REAL", "")
      END IF                                                      '
      Prf.NumType = NumType                                       ' Save it in the profile
   END IF                                                         '
   MErrExit(%eNone, "NUMTYPE set to - " + IIF$(Prf.NumType = "", "NONE", Prf.NumType)) ' Echo it
   MExit                                                          '
END METHOD

METHOD  pCmdOPEN(pCmd AS STRING)
'---------- Start a new EDIT session
LOCAL lclCmd, fn AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   fn = IIF$(pCmdNumOps = 0, "", pCmdOps(1))                      ' Get a filename from Command line if there
   IF ISNULL(fn) THEN                                             ' No filename?
      fn = sDoOpenFile("Specify file to Edit")                    ' Go get one
      IF ISNULL(fn) THEN _                                        '
         MErrExit(%eNone, "File selection cancelled")             ' No selection?    Bail out
   END IF                                                         ' We have fn="" for NEW or a filename

   IF ISNOTNULL(fn) THEN                                          ' We have a filename
      IF INSTR(fn,"\") = 0 THEN fn = sGetDefDir + fn              '
      IF ISNULL(PATHSCAN$(FULL, fn)) THEN _                       ' Valid name?
         MErrExit(%eFail, "Invalid Filename or missing file")     ' No,  Bail out
   END IF                                                         '
   sCloneEdit(fn)                                                 ' Go fire up the session
   MExit
END METHOD

METHOD  pCmdOPTIONS(pCmd AS STRING)
'---------- Handle Options Command
LOCAL StartPg AS LONG
LOCAL lclLNoSize, lclENGChars AS LONG, t AS STRING

REGISTER i AS LONG
REGISTER j AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   ENV.SaveScheme                                                 ' Save Scheme array 'as is' in case cancelled
   DIALOG GET LOC hWnd TO i, j                                    ' Get and save location where we are
   ENV.LastScreenX = i: ENV.LastScreenY = j                       '
   sIniSetString("Screen", "LastScrX", FORMAT$(ENV.LastScreenX))  '
   sIniSetString("Screen", "LastScrY", FORMAT$(ENV.LastScreenY))  '
   gfOptCancel = %False                                           ' Reset the cancelled flag
   lclENGChars = ENV.ENGChars                                     ' Save to monitor for changes
   lclLNoSize = ENV.LinNoSize                                     ' Save to monitor for changes
   IF pCmdNumOps = 0 THEN pCmdOps(1) = "GENERAL"                  '
   SELECT CASE UUCASE(pCmdOps(1))                                 ' Convert the operand
      CASE LEFT$("GENERAL", LEN(pCmdOps(1))):   StartPg = 1       '
      CASE LEFT$("FILEMGR", LEN(pCmdOps(1))):   StartPg = 2       '
      CASE LEFT$("FM", LEN(pCmdOps(1))):        StartPg = 2       '
      CASE LEFT$("SUBMIT", LEN(pCmdOps(1))):    StartPg = 3       '
      CASE LEFT$("SCREEN", LEN(pCmdOps(1))):    StartPg = 4       '
      CASE LEFT$("KEYBOARD", LEN(pCmdOps(1))):  StartPg = 5       '
      CASE LEFT$("KB", LEN(pCmdOps(1))):        StartPg = 5       '
      CASE LEFT$("STATUS", LEN(pCmdOps(1))):    StartPg = 6       '
      CASE LEFT$("SCHEME", LEN(pCmdOps(1))):    StartPg = 7       '
      CASE LEFT$("HILITES", LEN(pCmdOps(1))):   StartPg = 8       '
      CASE ELSE: StartPg = 1                                      '
   END SELECT                                                     '
   DispOptions(StartPg)                                           ' Go let user play with them
   IF gfOptCancel THEN                                            ' Cancelled out?
      ENV.LoadScheme                                              ' Restore original Scheme
      ENV.SetGlobalClrs                                           ' Copy all of it to Global
      CONTROL SET FOCUS hWnd, %IDC_SPFLiteTAB                     ' Set focus to regain mousewheel control
      MExitMeth                                                   ' Return
   END IF                                                         '
   ENV.SchemeSaveAll                                              ' Save scheme and HiLites
   ENV.SetGlobalClrs                                              ' Copy all of it to Global

   '----- See if default char string needs modifying due to English vs. non-English ...

   IF ENV.CharSet = "" THEN

      IF ENV.ENGChars THEN                                        ' Default to use English-only chars
         ENV.CharSet = $DefCharSet                                ' Set it without Nat chars
         gUpper = $Upper: gLower = $Lower                         '
      ELSE                                                        ' Use non-English chars
         ENV.CharSet = $DefCharSet + $UpperNat + $LowerNat        ' Set it that way
         gUpper = $Upper + $UpperNat: gLower = $Lower + $LowerNat '
      END IF                                                      '

      sIniSetString("General", "CharSet", ENV.CharSet)            ' Write it out
   END IF

   IF lclENGChars AND ISFALSE ENV.ENGChars THEN                   ' ENGLISH-ONLY is going ON  => OFF

      IF ENV.CharSet = $DefCharSet THEN                           ' If we're still sitting at the default
         ENV.CharSet = $DefCharSet + $UpperNat + $LowerNat        ' Add the national to it
         gUpper = $Upper + $UpperNat: gLower = $Lower + $LowerNat '
         sIniSetString("General", "CharSet", ENV.CharSet)         ' Write it out
      END IF                                                      '

   ELSEIF ISFALSE lclENGChars AND ENV.ENGChars THEN               ' ENGLISH-ONLY is going OFF => ON   gDefCharSet + $UpperNat + $LowerNat

      IF ENV.CharSet = $DefCharSet + $UpperNat + $LowerNat THEN   ' If we're still sitting at the stock Nat Char version
         ENV.CharSet = $DefCharSet                                ' Remove the national
         gUpper = $Upper: gLower = $Lower                         '
         sIniSetString("General", "CharSet", ENV.CharSet)         ' Write it out
      END IF                                                      '

   END IF                                                         '

   IF ENV.LinNoSize <> lclLNoSize THEN                            ' Has line number width changed?
      InitLNText                                                  ' Reset some stuff
   END IF                                                         '
   gDataLen = ENV.ScrWidth - gLNPadCol                            ' Re-Calc derived values
   pCmdLen = ENV.ScrWidth - 24                                    '
   SetCmd                                                         ' Cursor to Cmd line
   TP.PrfWordVal()                                                ' re-do Word in case NatChar changed
   gResizeActive = %True                                          '
   sResizeWindow(0, 0)                                            ' Go resize the window
   gResizeActive = %False                                         '
   sSetupSB                                                       ' Re-do SB for this tab
   LastRulCol = 0: LastRulRow = 0                                 ' In case crosshair cursors active
   IF TabsNum = 0 THEN MExitMeth                                  ' Don't screw up at termination
   sGblOptSet                                                     ' Go do corrections to all tabs
   scError(%eNone, "Options set")                                 ' Issue message
   MExit
END METHOD

METHOD  pCmdORDER(pCmd AS STRING)
'----- Do the simple renum function
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL lclNumb AS STRING
   MEntry
   IF LastLine > ENV.QuickRenum AND pCmd = "F" THEN MExitMeth     ' Exit if only quick renum
   IF ISFALSE IsRenumFlag THEN MExitMeth                          '
   OffRenumFlag                                                   ' Flip the flag
   FOR i = 1 TO LastLine                                          ' Spin through the lines
      IF IsLData(i) THEN                                          ' Only Data lines
         INCR j                                                   ' Bump line number
         lclNumb = FORMAT$(j, "00000000")                         ' Format a number
         IF LLNumG(i) <> lclNumb THEN                             ' Different?
            LLNumS(i) = lclNumb                                   ' Stuff it in
            me.UpdLControl(i)                                     ' Update the LLCtl
         END IF                                                   '
      END IF                                                      '
   NEXT i                                                         '
   LastReal = j                                                   ' Might as well save a known number
   MExit
END METHOD

METHOD  pCmdPAGE(pCmd AS STRING)
'---------- Turn PAGE ON / OFF & set offset
LOCAL TF AS LONG
LOCAL fOn, fOff, fFind, fAuto AS LONG, MSG AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors

   IF pCmdNumOps = 1 THEN                                         ' Process the operands
      IF IsEQ(pCmdOps(1), "ON") THEN                              ' ON
         Prf.PageFlag = 1: Prf.PageOffset = 0                     '
      ELSEIF IsEQ(pCmdOps(1), "SCROLL") THEN                      ' SCROLL
         Prf.PageFlag = 2: Prf.PageOffset = 0                     '
      ELSEIF IsEQ(pCmdOps(1), "OFF") THEN                         ' OFF
         Prf.PageFlag = 0: Prf.PageOffset = 0                     '
      ELSEIF VERIFY(pCmdOps(1), "+-0123456789") = 0 THEN          ' +/- nnn
         Prf.PageFlag = 1: Prf.PageOffset = VAL(pCmdOps(1))       '
      ELSE                                                        '
         MErrExit(%eFail,"Unknown PAGE operand - " + pCmdOps(1))  ' Issue error message
      END IF                                                      '

   ELSEIF pCmdNumOps = 2 THEN                                     ' Two operand format
      IF IsEQ(pCmdOps(1), "ON") THEN                              ' ON
         TF = 1                                                   '
      ELSEIF IsEQ(pCmdOps(1), "SCROLL") THEN                      ' SCROLL
         TF = 2                                                   '
      ELSEIF IsEQ(pCmdOps(1), "OFF") THEN                         ' OFF
         TF = 0                                                   '
      ELSE                                                        '
         MErrExit(%eFail,"Unknown PAGE operand - " + pCmdOps(1))  ' Issue error message
      END IF                                                      '
      IF VERIFY(pCmdOps(2), "+-0123456789") > 0 THEN              ' +/- nnn
         MErrExit(%eFail,"Unknown PAGE operand - " + pCmdOps(2))  ' Issue error message
      END IF
      IF TF = 2 THEN _                                            '
         MErrExit(%eFail, "Page offsets not allowed in PAGE SCROLL mode")
      Prf.PageFlag = TF: Prf.PageOffset = VAL(pCmdOps(2))              '
   END IF                                                         '

   scError(%eNone, "PAGE set to " + IIF$(Prf.PageFlag, IIF$(Prf.PageFlag = 2, "SCROLL", "ON"), "OFF") + IIF$(Prf.PageOffset <> 0, "/" + FORMAT$(Prf.PageOffset, "+#;-#"), "")) ' Tell user
   MExit                                                          '
END METHOD

METHOD  pCmdPASTE(pCmd AS STRING)
'---------- PASTE Clipboard data into the dataset
LOCAL lclTop, lc, lc1, i, j, x, y, z, pl, ABi, NumAdj, NumInDest, SetCur, before, after, cberase, ablabel AS LONG
LOCAL ln, t, cbname, MyDlm AS STRING

   MEntry
   '----- Do basic parsing
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   lclTop = TopScrn                                               ' Save current top
   me.ResetFunc(%ResetFind)                                       ' Clear previous FIND hilighting

   '----- Scan our operands
   FOR i = 1 TO pCmdNumOps                                        '
      IF IsEQ(pCmdOps(i), "BEFORE") THEN                          ' Handle BEFORE
         before = %True                                           ' Remember it
         GOSUB DelOne                                             ' Remove from parse
      ELSEIF IsEQ(pCmdOps(i), "AFTER") THEN                       ' Handle AFTER
         after = %True                                            ' Remember it
         GOSUB DelOne                                             ' Remove from parse
      ELSEIF IsEQ(pCmdOps(i), "ERASE") THEN                       ' Handle ERASE
         cberase = %True                                          ' Remember it
         GOSUB DelOne                                             ' Remove from parse
      ELSEIF pCmdOpsType(i) = %OpDotd THEN                        ' A label?
         IF ablabel > 0 THEN _
            MErrExit(%eFail, "Multiple line reference detected - " + pCmdOps(i)) ' Oops!
         ablabel = me.LineNoRef(pCmdOps(i))                       ' Get the line number
         IF ablabel = -1 THEN _                                   ' If invalid
            MErrExit(%eFail, "Invalid line reference - " + pCmdOps(i)) ' Oops!
         GOSUB DelOne                                             ' Remove from parse
      ELSEIF pCmdOpsType(i) = %OpLPtr THEN                        ' A LPtr reference?
         IF TRIM$(pCmdOps(i)) = "!" THEN _                        ' just a !?
            MErrExit(%eFail, "Line reference: " + pCmdOps(i) + " is invalid") ' Oops!
         IF ablabel > 0 THEN _
            MErrExit(%eFail, "Multiple line reference detected - " + pCmdOps(i)) ' Oops!
         ablabel = VAL(MID$(pCmdOps(i), 2))                       ' Get the line number
         IF ablabel = 0 OR ablabel > LastLine THEN _              ' If invalid
            MErrExit(%eFail, "Line reference: " + pCmdOps(i) + " is invalid") ' Oops!
         GOSUB DelOne                                             ' Remove from parse
      ELSEIF pCmdOpsType(i) = %OpStr THEN                         ' Simple string?
         IF INSTR($Numeric, LEFT$(pCmdOps(i) , 1)) = 0 THEN       ' Not a leading numeric?
            cbname = pCmdOps(i)                                   ' Save clipboard name
            ARRAY DELETE pCmdOps(i): ARRAY DELETE pCmdRaw(i): ARRAY DELETE pCmdOpsType(i)
            DECR i: DECR pCmdNumOps                               ' So we stay on the same index number
         END IF                                                   '
      END IF                                                      '
   NEXT i                                                         '

   '----- Look for stupidities
   IF before AND after THEN _
      MErrExit(%eFail, "Can't specify both BEFORE and AFTER")     ' Issue error
   IF (before OR after) AND ablabel = 0 THEN _                    '
      MErrExit(%eFail, "BEFORE/AFTER require a line label")       ' Issue error

   '----- See if we're empty or not
   IF LastLine <> 2 THEN                                          ' If we're not empty

      '----- Not empty, get our line reference
      IF before > 0 THEN                                          ' Stuff in A/B from cmd line
         lc = ablabel - 1                                         '
         IF ISTRUE LTblRange THEN _                               ' See if an A/B line control waiting
            MErrExit(%eCmdPend, "Cannot specify BEFORE and also an A/B, O/OO, OR/ORR, AA/BB or H/HH line control")

      ELSEIF after > 0 THEN                                       '
         lc = ablabel                                             '
         IF ISTRUE LTblRange THEN _                               ' See if an A/B line control waiting
            MErrExit(%eCmdPend, "Cannot specify AFTER and also an A/B, O/OO, OR/ORR, AA/BB or H/HH line control")

      ELSE                                                        '
         '----- Last chance is a line command range
         IF ISFALSE LTblRange  OR _                               ' A Line range?
            INSTR("A       B       H       HH      O       OO      OR      ORR     OROR    AA      BB      ", LTblDCmd) = 0 THEN ' But not a valid one?
               MErrExit(%eCmdPend, "Specify A/B, O/OO, OR/ORR, AA/BB or H/HH line control for PASTE")
         END IF                                                   '

         IF LTblDCmd = "A       " OR LTblDCmd = "B       " THEN   ' One of the simple ones?
            lc = IIF(LTblDCmd = "A       ", LTblDFrom, LTblDFrom - 1)' Set Insert point
            me.TTblDel(LTblDFrom)                                 ' Remove from Touched lines
            me.UpdLControl(LTblDFrom)                             ' Clear from the line
            IF IsLXClude(lc) THEN _                               ' If this is an Excluded line
               lc += LWrk1G(lc)                                   ' Step over
         ELSE                                                     ' It's one of the complicated ones
            GOSUB PasteBlock                                      ' Go handle specially
            MExitMeth                                             '
         END IF                                                   '
      END IF                                                      '

   '----- File is empty
   ELSE                                                           '
      lc = 1                                                      ' Else insert after top line
      i = me.RangeSet("RESET", %False, %False)                    ' Ensure default range of ALL before reset
      IF ISTRUE LTblRange THEN                                    ' User entered an A/B anyway?
         IF LTblDCmd = "A       " OR LTblDCmd = "B       " THEN   ' One of the simple ones?
            me.TTblDel(LTblDFrom)                                 ' Remove from Touched lines
            me.UpdLControl(LTblDFrom)                             ' Clear from the line
         END IF                                                   '
      END IF                                                      '
   END IF                                                         '

   '----- Insert after the lc pointer
   gKeyPrimOper = cbname: MyDLM = $CRLF                           ' Set clipboardname / delimiter
   IF sReadClipboard(t, MyDLM, cberase) THEN                      ' Go get whatever's there (True = missing CLIP file
      MErrExit(%eFail, "The " + gKeyPrimOper + " CLIP file does not exist")
   END IF                                                         '

   IF LEN(t) < 1 THEN                                             ' Error if nothing there
      MErrExit(%eFail, "The " + IIF$(ISNULL(gKeyPrimOper), "Windows", gKeyPrimOper) + " clipboard was empty")
   END IF                                                         '

   '----- Deblock the clipboard and insert the lines
   me.ModSet(lc)                                                  ' Remember we changed something
   IF lc1 = 0 THEN lc1 = lc                                       ' Save 1st insertion point
   DO WHILE LEN(t) > 1                                            ' Loop extracting text lines
      ln = EXTRACT$(t, MyDLM)                                     ' Get one line
      ln = TAB$(ln, Prf.ImportTabs)                               ' Do any needed tab conversion
      t =  REMAIN$(t, MyDlm)                                      ' Make t = the remainder of the line
      me.LInsertEmpty(lc, 1, %Data)                               ' Request insert 1, %Data line
      INCR lc                                                     ' Bump insertion point
      me.LTxtSet(lc, ln)                                          ' Copy in the extacted text
      me.UpdLControl(lc)                                          ' Setup LLCtl
      me.AttrScan(lc)                                             ' Recolorize
      LPCmdS(lc) = PCmdIX                                         ' Stuff in PCmdIX to make it :ZF
      GOSUB DoPMlc                                                ' Go do PM processing
   LOOP                                                           '

   '----- Clean up and exit
   TopScrn = lclTop                                               ' Restore CRP

   IF ISFALSE IsSetEdit THEN                                      ' If not Set Edit
      scError(0, "Paste complete from " + IIF$(ISNULL(gKeyPrimOper), "Windows", gKeyPrimOper) + " clipboard")
      IF cberase THEN sWriteClipboard("")                         ' Erase if asked for
      RESET gKeyPrimOper                                          ' Blank clipboardname
      me.CurSetReq(%LineCmd, lc1 + 1, 1, %True)                   ' Set cursor set attempt

   ELSE                                                           '
      '----- manage clipboard for special use in SET Edit session

      scError(0, "Current SET values")                            '
      sWinclip_set(gSetClipB)                                     '
   END IF                                                         '

   MExitMeth

'----- Do the paste block mode type
PasteBlock:                                                       ' Do the Paste Overlay stuff

   '----- First get the CBD data into an array
   DIM IStr(1000) AS STRING                                       ' Get a table to hold paste data
   gKeyPrimOper = cbname: MyDLM = $CRLF                           ' Set clipboardname / delimiter
   IF sReadClipboard(t, MyDLM, cberase) THEN                      ' Go get whatever's there (True = missing CLIP file
      scError(%eFail, "The " + gKeyPrimOper + " CLIP file does not exist")
      me.TTblDel(LTblSfrom): me.UpdLControl(LTblSfrom)            ' Remove from Touched lines
      IF LTblSfrom <> LTblSTo THEN me.TTblDel(LTblSTo): me.UpdLControl(LTblSTo)
      MExitMeth                                                   '
   END IF                                                         '
   IF LEN(t) < 1 THEN                                             ' Error if nothing there
      scError(%eFail, "Clipboard was empty")                      '
      MExitMeth                                                   '
   END IF                                                         '

   lclTop = TopScrn                                               ' Save current top

   '----- Extract CBD lines to the IStr array
   pl = 0                                                         ' Reset count
   DO WHILE LEN(t) > 1                                            ' Loop extracting text lines
      INCR pl                                                     '
      IF pl > UBOUND(IStr) THEN                                   ' Expand IStr if needed
         REDIM PRESERVE IStr(UBOUND(IStr) + 1000)                 '
      END IF                                                      '
      IStr(pl) = EXTRACT$(t, MyDLM)                               ' Get one line
      IStr(pl) = TAB$(IStr(pl), Prf.ImportTabs)                   ' Do any needed tab conversion
      t =  REMAIN$(t, MyDlm)                                      ' Make t = the remainder of the line
   LOOP                                                           ' Loop til IStr holds pl paste lines

   '----- Handle the O/OO or OR/ORR style PASTE functions
   IF LTblDCmd = "O       " OR LTblDCmd = "OO      " OR _         ' Is this the O / OO Overlay Type?
      LTblDCmd = "OR      " OR LTblDCmd = "ORR     " OR LTblDCmd = "OROR    " THEN ' or OR/ORR/OROR

      IF LTblDfrom = LTblDTo THEN                                 ' Same From/To  (i.e. O command)
         LTblDTo = LTblDfrom + MAX(1, LTblDRpt) - 1               ' Set To address based on Onnn count (Def 1)
      END IF                                                      '
      me.TTblDel(LTblDfrom)                                       ' Remove from Touched lines
      me.UpdLControl(LTblDfrom)                                   ' Clear from the line
      me.TTblDel(LTblDTo)                                         '
      me.UpdLControl(LTblDTo)                                     '
      me.CurSetReq(%LineCmd, LTblDfrom, 1, %False)                ' Set cursor set attempt
      z = 1                                                       ' Point at 1st Paste line
      FOR x = LTblDfrom TO LTblDTo                                ' Loop through destination lines
         IF IsLData(x) THEN                                       ' Overlay Dest Data lines
            me.ModSet(x)                                          ' Remember we changed something
            IF LTblDCmd = "OR      " OR LTblDCmd = "ORR     " OR LTblDCmd = "OROR    " THEN ' or OR/ORR/OROR
               me.LTxtSet(x, me.OverlayTextRepl(IStr(z), LTxtG(x))) '
            ELSE                                                  '
               me.LTxtSet(x, me.OverlayText(IStr(z), LTxtG(x)))   '
            END IF                                                '
            me.AttrScan(x)                                        ' Recolorize
            me.LFlagBitOff(x, %EQChange)                          ' Reset ==CHG>
            me.UpdLControl(x)                                     ' And LLCtl
            LPCmdS(x) = PCmdIX                                    ' Stuff in PCmdIX to make it :ZF
            GOSUB DoPMx                                           ' Go do PM processing
            IF z < pl THEN INCR z ELSE z = 1                      ' Next paste source line (or repeat from the top)
         END IF                                                   '
      NEXT x                                                      ' Loop Outer

   '----- Handle the H/HH style PASTE functions
   ELSEIF LTblDCmd = "H       " OR LTblDCmd = "HH      " THEN     ' Is this the H / HH Type?

      FOR x = LTblDfrom TO LTblDTo                                ' Loop for each deleted line
         me.ModSet(x)                                             ' Remember we changed something
         me.lTxtFree(x)                                           ' Go free the dynamic string
      NEXT x                                                      '
      FOR x = 1 TO LTblDTo - LTblDfrom + 1                        ' Loop for each deleted line
         me.LEntDel(LTblDfrom)                                    ' Remove from the L() array
      NEXT x                                                      '
      LastLine -= LTblDTo - LTblDfrom + 1: LastReal -= LTblDTo - LTblDfrom + 1 ' Adjust LastLine and LastReal
      lc = LTblDfrom - 1                                          ' Make previous line look like 'A'

      FOR x = 1 TO pl                                             ' Loop through extracted lines
         me.LInsertEmpty(lc, 1, %Data)                            ' Request insert 1, %Data line
         INCR lc                                                  ' Bump insertion point
         me.LTxtSet(lc, IStr(x))                                  ' Copy in the extacted text
         me.AttrScan(lc)                                          ' Recolorize
         me.UpdLControl(lc)                                       ' Setup LLCtl
         LPCmdS(lc) = PCmdIX                                      ' Stuff in PCmdIX to make it :ZF
         GOSUB DoPMlc                                             ' Go do PM processing
      NEXT x                                                      '
      me.CurSetReq(%LineCmd, LTblDfrom, 1, %False)                ' Set cursor set attempt


   '----- Handle the AA style PASTE functions
   ELSEIF LTblDCmd = "AA      " THEN                              ' Is this the AA After type

      me.TTblDel(LTblDfrom)                                       ' Remove from Touched lines
      me.UpdLControl(LTblDfrom)                                   ' Clear from the line
      lc1 = LTblDFrom + 1                                         ' Point where cursor goes
      me.TTblDel(LTblDTo)                                         '
      me.UpdLControl(LTblDTo)                                     '
      LTblDRpt = MAX(1, LTblDRpt)                                 ' Ensure at least count of 1
      FOR x = LTblDfrom TO LTblDTo                                ' Count the data lines
         IF IsLData(x) OR IsLTop(x) OR IsLFile(x) THEN            ' Data only
            INCR NumInDest                                        ' in the Dest. range
            GOSUB DoPMx                                           ' Go do PM processing
         END IF                                                   '
      NEXT x                                                      '
      IF NumInDest <= LTblDRpt THEN _                             ' Dest count reasonable?
         MErrExit(%eFail, "AA number must be < range size")       '
      ABi = LTblDfrom + LTblDRpt - 1                              ' Init for loop
      WHILE ABi <= LTblDTo + NumAdj                               ' Big outer loop
         IF ISFALSE IsLData(ABi) AND ISFALSE IsLTop(ABi) AND ISFALSE IsLFile(ABi) THEN INCR Abi: ITERATE  ' Ignore non Data lines
         me.ModSet(ABi)                                           ' Remember we changed something
         me.LInsertEmpty(ABi, pl, %Data)                          ' Insert enough lines for one copy
         NumAdj += pl                                             ' Track # inserted
         x = ABi + 1                                              ' Calc position to copy the data
         FOR y = 1 TO pl                                          ' For each item in the Clipboard
            me.LTxtSet(x, IStr(y))                                ' Copy in the extacted text
            me.AttrScan(x)                                        ' Recolorize
            me.UpdLControl(x)                                     ' Setup LLCtl
            LPCmdS(x) = PCmdIX                                    ' Stuff in PCmdIX to make it :ZF
            GOSUB DoPMx                                           ' Go do PM processing
            INCR x                                                ' Bump 'to' pointer
         NEXT y                                                   ' Loop inner
         ABI += LTblDRpt + pl                                     ' Big outer loop
      LOOP                                                        ' next ABi
      me.AdjustPending(LTblDfrom + LTblDRpt - 1, NumAdj, 0)       ' Adjust remaining items in Line Cmnd stack
      me.CurSetReq(%LineCmd, lc1, 1, %False)                      '

   '----- Handle the BB style PASTE functions
   ELSEIF LTblDCmd = "BB      " THEN                              ' Is this the BB After type

      me.TTblDel(LTblDfrom)                                       ' Remove from Touched lines
      me.UpdLControl(LTblDfrom)                                   ' Clear from the line
      me.TTblDel(LTblDTo)                                         '
      me.UpdLControl(LTblDTo)                                     '

      LTblDRpt = MAX(1, LTblDRpt)                                 ' Ensure at least count of 1
      FOR x = LTblDfrom TO LTblDTo                                ' Count the data lines
         IF IsLData(x) OR IsLTop(x) OR IsLFile(x) THEN            ' Data only
            INCR NumInDest                                        ' in the Dest. range
            GOSUB DoPMx                                           ' Go do PM processing
         END IF                                                   '
      NEXT x                                                      '
      IF NumInDest <= LTblDRpt THEN _                             ' Dest count reasonable?
         MErrExit(%eFail, "BB number must be < range size")       '
      LTblDRpt = MAX(1, LTblDRpt)                                 ' Ensure at least count of 1
      x = NumInDest \ LTblDRpt                                    ' x = # of insertion points
      ABi = LTblDTo - (x * LTblDRpt)                              ' Initial insertion point
      lc1 = LTblDfrom                                             ' Remember where cursor goes
      WHILE ABi < LTblDTo + NumAdj                                ' Big outer loop
         IF ISFALSE IsLData(ABi) AND ISFALSE IsLTop(ABi) AND ISFALSE IsLFile(ABi) THEN INCR Abi: ITERATE ' Ignore non Data lines
         me.ModSet(ABi)                                           ' Remember we changed something
         me.LInsertEmpty(ABi, pl, %Data)                          ' Insert enough lines for one copy
         NumAdj += pl                                             ' Track # inserted
         x = ABi + 1                                              ' Calc position to copy the data
         FOR y = 1 TO pl                                          ' For each item in the Clipboard
            me.LTxtSet(x, IStr(y))                                ' Copy in the extacted text
            me.AttrScan(x)                                        ' Recolorize
            me.UpdLControl(x)                                     ' Setup LLCtl
            LPCmdS(x) = PCmdIX                                    ' Stuff in PCmdIX to make it :ZF
            GOSUB DoPMx                                           ' Go do PM processing
            INCR x                                                ' Bump 'to' pointer
         NEXT y                                                   ' Loop inner
         ABI += LTblDRpt + pl                                     ' Big outer loop
      LOOP                                                        ' next ABi
      me.AdjustPending(LTblDfrom + LTblDRpt - 1, NumAdj, 0)       ' Adjust pending stuff
      me.CurSetReq(%LineCmd, lc1, 1, %False)                      '

   END IF                                                         '

   '----- Clean up and exit
   i = me.RangeSet("RESET", %False, %False)                       ' Ensure default range of ALL before reset
   TopScrn = lclTop                                               ' Restore CRP
   IF cberase THEN sWriteClipboard("")                            ' Erase if asked for
   scError(0, "Paste complete from " + IIF$(ISNULL(gKeyPrimOper), "Windows", gKeyPrimOper) + " clipboard")
   RESET gKeyPrimOper                                             ' Blank clipboardname
   MExitMeth

'----- Do +/- processing for X lines
DoPMx:
   IF BIT(LTblDFlag, %lCmdX) THEN                                 ' Do the +/- processing for the Source
      me.LFlagBitOn(x, %Invisible)                                ' If - make Invisible
      gfXRebuild = %True                                          ' Ask for Exclude processing
   ELSEIF BIT(LTblDFlag, %lCmdNX) THEN                            '
      me.LFlagBitOff(x, %Invisible)                               ' If + make visible
      gfXRebuild = %True                                          ' Ask for Exclude processing
   END IF                                                         '
   RETURN                                                         '

'----- Do +/- processing for lc lines
DoPMlc:
   IF BIT(LTblDFlag, %lCmdX) THEN                                 ' Do the +/- processing for the Source
      me.LFlagBitOn(lc, %Invisible)                               ' If - make Invisible
      gfXRebuild = %True                                          ' Ask for Exclude processing
   ELSEIF BIT(LTblSFlag, %lCmdNX) THEN                            '
      me.LFlagBitOff(lc, %Invisible)                              ' If + make visible
      gfXRebuild = %True                                          ' Ask for Exclude processing
   END IF                                                         '
   RETURN                                                         '

'----- Remove current parse entry
DelOne:
   ARRAY DELETE pCmdOps(i): ARRAY DELETE pCmdRaw(i): ARRAY DELETE pCmdOpsType(i)
   DECR i: DECR pCmdNumOps                                        ' So we stay on the same index number
   RETURN

END METHOD

METHOD  pCmdPLEASE(pCmd AS STRING)
'---------- PLEASE debugging
LOCAL Wrd, lclCmd AS STRING, nv AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, get an operand
   IF IsEQ(pCmdOps(1), "LOOP") THEN                               ' LOOP?
      DO: LOOP                                                    ' Force a loop
   ELSEIF IsEQ(pCmdOps(1), "CRASH") THEN                          ' CRASH?
      POKE 0, 12345                                               ' Force a crash
   ELSEIF IsEQ(pCmdOps(1), "NV") THEN                             ' No version?
      gNoVersion = %True                                          ' Set the flag
   END IF                                                         '
   MExit
END METHOD

METHOD  pCmdPRESERVE(pCmd AS STRING)
'---------- Turn Preserve ON or OFF or C
LOCAL nv AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps > 0 THEN                                         ' If an operand,
      SELECT CASE AS CONST$ UUCASE(pCmdOps(1))                    ' Which did we get?
         CASE "OFF": nv = 0                                       ' OFF
         CASE "ON": nv = 1                                        ' ON
         CASE "C": nv = 2                                         ' C
         CASE ELSE: MErrExit(%eFail, " Invalid operand")          '
      END SELECT                                                  '
      Prf.PPreserve = nv                                          ' Set the new value
   END IF                                                         '
   scError(0, "PRESERVE set to " + IIF$(Prf.PPreserve = 0 , "OFF", IIF$(Prf.PPreserve = 1, "ON", "C")))  ' Tell user new value
   MExit
END METHOD

METHOD  pCmdPRINT(pCmd AS STRING)
'---------- Handle PRINT Command
LOCAL lclCmd, lclWord, CBD, setup AS STRING, lTxtp AS STRING POINTER
LOCAL i, j AS LONG, lclCursMode AS INTEGER, fl AS QUAD
LOCAL NumDel, lclTop, AppendText, recs, Number, lclmode AS LONG
DIM dlines(1000) AS LONG                                          ' To hold lines to be deleted
   MEntry
   Number = %False                                                ' Default to no-numbered listing
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   lclCursMode = %CursBad                                         ' Pretend a bad cursor location
   me.FindSave                                                    ' Save RFIND and Range stuff
   IF ISTRUE me.RangeSet("", %False, %False) THEN me.FindLoad: MExitMeth   ' Setup any range command operands
   IF pCmdNumOps = 0 AND ISFALSE DotSet THEN                      ' Better be something
      me.FindLoad                                                 ' Reload RFIND stuff
      MErrExit(%eFail, "Pending line range")                      '
   END IF                                                         '
   IF pCmdNumOps > 0 THEN                                         ' Scan any operands
      FOR i = 1 TO pCmdNumOps                                     '
         IF pCmdOpsType(i) = %KWNONUM THEN                        ' Handle NONUM
            Number = %False                                       ' Set NONUM
            ARRAY DELETE pCmdOps(i): ARRAY DELETE pCmdRaw(i): ARRAY DELETE pCmdOpsType(i)
            DECR i: DECR pCmdNumOps                               ' So we stay on the same index number
         ELSEIF pCmdOpsType(i) = %KWNUM THEN                      ' Handle NUM
            Number = %True                                        ' Set NONUM
            ARRAY DELETE pCmdOps(i): ARRAY DELETE pCmdRaw(i): ARRAY DELETE pCmdOpsType(i)
            DECR i: DECR pCmdNumOps                               ' So we stay on the same index number
         ELSEIF pCmdOpsType(i) = %KWSETUP THEN                    ' Handle SETUP
            Setup = UUCASE(pCmdOps(i))                            ' Save it
            IF pCmdNumOps > 1 THEN _                              ' SETUP must be alone
               MErrExit(%eFail, "SETUP must be the only parameter") ' Issue conflict error
            sOpenPrinter(Setup)                                   '
            me.FindLoad                                           ' Reload RFIND stuff
            MExitMeth                                             ' And exit
         END IF                                                   '
      NEXT i                                                      '
   END IF                                                         '
   IF IsFMTab THEN                                                ' Only SETUP allowed for File Manager
      me.FindLoad                                                 ' Reload RFIND stuff
      MErrExit(%eFail,"Only PRINT SETUP allowed in File Manager") '
   END IF                                                         '

   me.FindReset                                                   ' Reset the parse output fields
   IF sCRTParse("ASU") THEN me.FindLoad: MExitMeth                ' Scan (All, Subset, User)
   me.FindSetup                                                   ' Copy CRT parse's answers
   CfSet(%CRTDX)                                                  ' Tell search hands off Invisible status
   IF DotSet OR _                                                 ' If a line range
      IsCfU OR IsCfNU OR IsCfX OR IsCfNX THEN                     ' Or U / NU / X / NX
      CfSet(%CRTALL)                                              ' Force ALL if no modifiers
   END IF                                                         '
   errFlag = %eNone                                               ' Reset any Pending conflict flag
   lclTop = TopScrn                                               '

   IF ISFALSE sOpenPrinter("") THEN                               ' Get printer ready if not already
      me.FindLoad                                                 ' Reload RFIND stuff
      MErrExit(%eFail, "OPEN of printer failed")                  '
   END IF                                                         '

   sPrtPrint(%PRTReset, " ", " ", Number)                         ' Tell sPrtPrint to reset
   IF IsCfAll THEN                                                ' The ALL version?
      i = me.Search(0, lclCursMode)                               ' Do the initial search
      IF i THEN                                                   '
         GOSUB PrintOne                                           ' Go handle one
      END IF                                                      '
      DO WHILE i                                                  ' Do till not found
         i = me.Search(1, lclCursMode)                            ' Do re-Finds
         IF i THEN                                                ' Found one
            GOSUB PrintOne                                        ' Go handle one
         END IF                                                   '
      LOOP                                                        '
   ELSE                                                           ' The non-All version
      i = me.Search(0, lclCursMode)                               ' Go see if we can find it
      IF i THEN                                                   ' Got it?
         GOSUB PrintOne                                           ' Go handle one
      END IF                                                      '
   END IF                                                         '

   sPrtPrint(%PRTFlushClose, " ", " ", Number)                    ' Tell sPrtPrint to flush page and close
   TopScrn = lclTop                                               '
   scError(0, "Print complete to default printer")                '
   me.FindLoad                                                    ' Reload RFIND stuff
   MExitMeth                                                      '

PrintOne:
   i = cfFLine                                                    ' Get the located line number
   IF IsLData(i) AND ISFALSE IsLInsertLine(i) THEN                ' Only data lines of course

      j = DotFlag                                                 ' Copy +/- flag
      IF BIT(j, %lCmdX) THEN                                      ' Do the +/- processing for the Source
         me.LFlagBitOn(i, %Invisible)                             ' If - make Invisible
         gfXRebuild = %True                                       ' Do exclude stuff
      ELSEIF BIT(j, %lCmdNX) THEN                                 '
         me.LFlagBitOff(i, %Invisible)                            ' If + make visible
         gfXRebuild = %True                                       ' Do exclude stuff
      END IF                                                      '

      lclmode = IIF(IsLPage(i), %PRTNewPage, %PRTLine)            ' Set lclmode for forced new page if a =PAGE> line
      IF Number THEN                                              ' Do listing with line numbers?
         sPRTPrint(lclMode, RIGHT$(LLNumG(i), ENV.LinNoSize) + " " + LTxtG(i), _      ' Print line with added linenumber
                            REPEAT$(ENV.LinNoSize + 1, $$TxtLo) + LAttrG(i), Number)
      ELSE                                                        '
         sPRTPrint(lclMode, LTxtG(i), LAttrG(i), Number)          ' Print each line without its line number
      END IF                                                      '
      sPRTPrint(%PRTNewLine, " ", " ", Number)                    ' Do end of line
   END IF                                                         '
   RETURN                                                         '
END METHOD

METHOD  pCmdPROFILE(pCmd AS STRING)
'---------- Handle PROFILE Command
LOCAL SaveProf, t AS STRING, SaveLock, SaveMode AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF IsFMTab AND _                                               ' If FM mode only do certain functions
      pCmdOpsType(1) <> %KWNew AND pCmdOpsType(1) <> %KWEdit THEN _
      MErrExit(%eFail, "Only NEW and EDIT allowed in File Manager")
   IF pCmdNumOps = 0 THEN                                         ' If no operands
      me.ProfDisp                                                 ' Just activate the display
      MExitMeth                                                   '
   END IF                                                         '

   IF (pCmdOpsType(1) = %KWLOCK OR _                              ' These types have only one operand
       pCmdOpsType(1) = %KWUNLOCK OR _                            '
       pCmdOpsType(1) = %KWRESET) AND _                           '
      pCmdNumOps > 1 THEN _                                       '
         MErrExit(%eFail, "Excess operands for the PROFILE command")
   IF LEFT$(pCmdOps(2), 1) = "." THEN pCmdOps(2) = MID$(pCmdOps(2), 2) ' Strip any leading optional period

   IF pCmdOpsType(1) = %KWUNLOCK THEN                             ' UNLOCK / UNLOCKED
      Prf.PLock = %False                                          '
      me.ProfDisp                                                 ' Just activate the display
      MErrExit(%eNone, "Profile is now unlocked")                 '

   ELSEIF pCmdOpsType(1) = %KWLOCK THEN                           ' LOCK / LOCKED
      Prf.PLock = %True                                           '
      me.ProfDisp                                                 ' Just activate the display
      MErrExit(%eNone, "Profile is now locked")                   '

   ELSEIF pCmdOpsType(1) = %KWRESET THEN                          ' Handle RESET
      Prf.Reset()                                                 ' Reset things to standard set
      Prf.WriteAll                                                ' Write it all to the INI file
      me.ProfDisp                                                 ' Just activate the display

   ELSEIF pCmdOpsType(1) = %KWCOPY THEN                           ' Handle COPY
      IF pCmdNumOps < 2 THEN _                                    '
         MErrExit(%eFail, "No Profile name specified with COPY")  '
      IF ISFALSE ISFILE(ENV.PROFPath + pCmdOps(2) + ".INI") THEN _' See if it exists
         MErrExit(%eFail, "Specified Profile name does not exist")'
      SaveProf = Prf.ProfName                                     ' Save current Profile Name
      Prf.SetProfName(pCmdOps(2), %False)                         ' Set to the Profile being copied
      Prf.ReadAll(%True)                                          ' Read everything in
      Prf.SetProfName(SaveProf, %False)                           ' Set back to the original Profile
      SaveLock = Prf.PLock                                        ' Save Lock status
      Prf.PLock = %False                                          ' Unlock so we can write it
      Prf.WriteAll                                                ' Write it all
      Prf.PLock = SaveLock                                        ' Reset the Lock status
      me.ProfDisp                                                 ' Just activate the display
      scError(%eNone, "Specified Profile values loaded")          '

   ELSEIF pCmdOpsType(1) = %KWNEW THEN                            ' Handle NEW
      IF pCmdNumOps < 2 THEN _                                    '
         MErrExit(%eFail, "No Profile name specified with NEW")   '
      IF ISFILE(ENV.PROFPath + pCmdOps(2) + ".INI") THEN _        ' See if it exists
         MErrExit(%eFail, "Specified Profile name already exists")'
      '----- Save current Profile name and build a new one
      SaveProf = Prf.ProfName                                     '
      Prf.SetProfName("DEFAULT", %False)                          ' Load DEFAULT values
      Prf.ReadAll(%True)                                          '
      Prf.SetProfName(pCmdOps(2), %True)                          ' Insert the new name
      scError(%eNone, Prf.EditProf)                               ' Let user set things up, return message
      '----- Remove from ENV.DefaultShr just in case
      ENV.DefaultShr = UUCASE(ENV.DefaultShr)                     ' Uppercase the list
      ENV.DefaultShr = REMOVE$(ENV.DefaultShr, UUCASE(Prf.ProfName) + ",")' Remove this Profname if present
      sIniSetString("General", "DefaultShr", ENV.DefaultShr)      ' Save it back
      '----- Put back the original Profile
      Prf.SetProfName(SaveProf, %False)                           ' Load DEFAULT values
      Prf.ReadAll(%True)                                          '
      CONTROL SET FOCUS hWnd, %IDC_SPFLiteTAB                     ' Set focus

   ELSEIF pCmdOpsType(1) = %KWEDIT THEN                           ' Handle EDIT
      IF pCmdNumOps < 2 OR pCmdOps(2) = Prf.ProfName THEN         ' Doing ourselves?
         IF Prf.UseMode THEN                                      ' If we're un USING mode
            SaveProf = Prf.ProfName                               ' Save our current Profile name
            Prf.SetProfName(Prf.UseName, %False)                  ' Setup Prf to use
            Prf.ReadAll(%False)                                   '
            scError(%eNone, Prf.EditProf)                         ' Go do it and return message
            Prf.SetProfName(SaveProf, %False)                     ' Load DEFAULT values
            Prf.ReadAll(%True)                                    '
         ELSE
            scError(%eNone, Prf.EditProf)                         ' Go do it and return message
         END IF
         sDoStatusBar($AllStatusBarBoxes)                         ' Pick up changes

      ELSE                                                        ' Doing some other Profile
         IF ISFALSE ISFILE(ENV.PROFPath + pCmdOps(2) + ".INI") THEN _' See if it exists
            MErrExit(%eFail, "Specified Profile name does not exist")'
         SaveProf = Prf.ProfName                                  ' Save our current Profile name
         Prf.SetProfName(pCmdOps(2), %False)                      ' Setup Prf to use
         Prf.ReadAll(%False)                                      '
         scError(%eNone, Prf.EditProf)                            ' Go do it and return message
         '----- Put back the original Profile
         Prf.SetProfName(SaveProf, %False)                        ' Load DEFAULT values
         Prf.ReadAll(%True)                                       '
      END IF                                                      '
      CONTROL SET FOCUS hWnd, %IDC_SPFLiteTAB                     ' Set focus

   ELSEIF pCmdOpsType(1) = %KWUSING THEN                          ' USE / USING
      IF IsEQ(Prf.ProfName, "DEFAULT") THEN                       ' Can't do USING on a DEFAULT
         scError(%eFail, "Cannot redirect the DEFAULT Profile with a USING")
         MExitMeth                                                '
      END IF                                                      '
      IF pCmdNumOps < 2 THEN                                      '
         scError(%eNone, "Profile USE has been reset")            '
         Prf.UseName = ""                                         '
         Prf.ReadAll(%True)                                       '

      ELSE                                                        '
         IF IsEQ(pCmdOps(2), Prf.ProfName) THEN _                 ' Recurse?
            MErrExit(%eFail, "Cannot USE the same profile name")  '
         IF ISFALSE ISFILE(ENV.PROFPath + pCmdOps(2) + ".INI") THEN _ ' See if it exists
            MErrExit(%eFail, "Specified Profile name does not exist") '
         SaveProf = Prf.Profname                                  ' Save current Profname

         '----- We need to 'peek' at the UseName of the new profile
         Prf.SetProfName(pCmdOps(2), %False)                      ' Setup to look at the USING Profile
         t = Prf.GetString("ProfUsing", "")                       ' Get any Using value or leave as null
         Prf.SetProfName(SaveProf, %False)                        ' Restore the Profname
         IF ISNOTNULL(t) THEN _                                   ' Is the using one already doing a USING?
            MErrExit(%eFail, "Profile (" + pCmdOps(2) + ") is already redirected via USING to (" + t + ")")
         Prf.UseName = pCmdOps(2)                                 ' Swap as a Using
         Prf.ReadAll(%True)                                       '
         scError(%eNone, "Profile values re-loaded")              '
      END IF                                                      '
      me.ProfDisp                                                 ' Just activate the display

   ELSE                                                           ' Operand is just a name
      IF pCmdNumOps > 1 THEN _                                    ' Better be no more operands
         MErrExit(%eFail, "Excess operands for the PROFILE command")
      IF ISFALSE ISFILE(ENV.PROFPath + pCmdOps(1) + ".INI") THEN _' See if it exists
         MErrExit(%eFail, "Specified Profile name does not exist")'
      Prf.SetProfName(pCmdOps(1), %False)                         ' Swap in place of existing
      Prf.ReadAll(%True)                                          ' Swap in the Profiles values
      me.ProfDisp                                                 ' Just activate the display
      scError(%eNone, "Specified Profile values loaded")          '
   END IF                                                         '
   sDoStatusBar($AllStatusBarBoxes)                               ' re-Do the StatusBar box
   MExit
END METHOD

METHOD  pCmdPTYPE(pCmd AS STRING)
'---------- Power Type command
LOCAL i, j, lclTop, lccol AS LONG, lclCursMode AS INTEGER
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   lclCursMode = %CursBad                                         ' Pretend a bad cursor location
   IF Prf.HexMode <> &1 THEN _                                    '
      MErrExit(%eFail, "PowerType not available in Hex mode")     '
   IF gKbdRecFlag THEN _                                          ' Doing a recording?
      MErrExit(%eFail, "PowerType not available during Keyboard recording")
   me.ResetFunc(%ResetFind)                                       ' Clear previous FIND hilighting


   IF ISTRUE me.RangeSet("", %False, %False) THEN MExitMeth       ' Setup any range command operands
   IF pCmdNumOps = 0 AND ISFALSE DotSet THEN _                    ' Better be something
      MErrExit(%eFail, "Pending line range")                      '
   me.FindReset                                                   ' Reset the parse output fields
   IF sCRTParse("ACSU") THEN MExitMeth                            ' Scan (All, Cols, Subset, User)
   me.FindSetup                                                   ' Copy CRT parse's answers
   CfSet(%CRTDX)                                                  ' Tell search hands off Invisible status
   IF DotSet OR _                                                 ' If a line range
      IsCfU OR IsCfNU OR IsCfX OR IsCfNX THEN                     ' Or U / NU / X / NX
      CfSet(%CRTALL)                                              ' Force ALL if no modifiers
   END IF                                                         '
   IF CrtFCol > 0 THEN                                            ' Was a col number entered?
      lccol = CrtFCol                                             ' Yes, use it
   ELSE
      lccol = IIF(IsCData, CCol - gLNPadCol + Offset, 1 + Offset) ' Calc IX into the Text
   END IF
   errFlag = %eNone                                               ' Reset any Pending conflict flag
   lclTop = TopScrn: gPTblCount = 0                               '

   i = me.Search(0, lclCursMode)                                  ' Do the initial search
   IF i THEN                                                      '
      GOSUB TableOne                                              ' Go handle one
   END IF                                                         '
   DO WHILE i                                                     ' Do till not found
      i = me.Search(1, lclCursMode)                               ' Do re-Finds
      IF i THEN                                                   ' Found one
         GOSUB TableOne                                           ' Go handle one
      END IF                                                      '
   LOOP                                                           '
   IF gPTblCount = 0 THEN _                                       '
      MErrExit(%eFail, "No data lines in range")                  '

   IF IsLInvisible(gPTbl(1).tLin) THEN                            ' Is 'model' line visible?
      me.LFlagBitOff(gPTbl(1).tLin, %Invisible)                   ' No, make it visible

      IF LWrk1G(gPTbl(1).tLin) = 1 AND _                          ' If 1st and last in X'd group (i.e. 1 line)
         LWrk1G(gPTbl(1).tLin) = LWrk1G(gPTbl(1).tLin - LWrk1G(gPTbl(1).tLin)) THEN
         gPTbl(1).tLin = gPTbl(1).tLin - 1                        ' Adjust new location
         FOR j = 2 TO gPTblCount                                  ' Adjust all lower PT entries by -1
            gPTbl(j).tLin = gPTbl(j).tlin - 1                     ' Since the ---- line is disappearing
         NEXT j                                                   '
      ELSEIF LWrk1G(gPTbl(1).tLin) = 1 THEN                       ' Adjust things
         gPTbl(1).tLin = gPTbl(1).tLin - 1                        ' If 1st in X'd range, adjust new location
      ELSEIF LWrk1G(gPTbl(1).tLin) = LWrk1G(gPTbl(1).tLin - LWrk1G(gPTbl(1).tLin)) THEN
         ' Nil                                                    ' If last in X'd range
      ELSE                                                        '
         FOR j = 2 TO gPTblCount                                  ' Adjust all lower PT entries by +1 since we
            gPTbl(j).tLin = gPTbl(j).tlin + 1                     ' are adding a new ---- line
         NEXT j                                                   '
      END IF                                                      '
   END IF
   me.CurSetReq(%Position, gPTbl(1).tLin, lccol, %True)           ' Set cursor set attempt
   OnPTypeMode                                                    ' Set into PType mode now
   MErrExit(%eNone, "Entering PowerType mode, Press Enter to exit")

TableOne:
   j = cfFLine                                                    ' Get the located line number
   IF IsLData(j) THEN                                             ' Only data lines of course
      INCR gPTblCount                                             ' Add to PT Table
      IF gPTblCount > UBOUND(gPTbl()) THEN _                      ' Table full
         REDIM PRESERVE gPTbl(1 TO 2 * UBOUND(gPTbl())) AS PTypeTable ' Enlarge if needed
      gPTbl(gPTblCount).tLin = j                                  ' Save Txt line number
      gPTbl(gPTblCount).sRow = 0                                  ' Say no screen row
   END IF                                                         '
   RETURN                                                         '
END METHOD

METHOD   pCmdQUERY(pCmd AS STRING)
'---------- Do a QUERY Command
LOCAL i, j, k AS LONG, wrd1, t, tt, key AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN _                                       ' If no operands
      MErrExit(%eFail, "Missing QUERY operand")                   '

   SELECT CASE CONST$ UUCASE(pCmdOps(1))                          ' See what we have
      CASE "ACTION"      : scError(%eNone, "ACTION set to " + IIF$(Prf.ActionSave, FORMAT$(Prf.ActionSave) + "," + Prf.ActionVerb, "OFF"))
      CASE "AUTOBKUP"    : scError(%eNone, "AUTOBKUP set to " + IIF$(Prf.AutoBkup, "ON", "OFF"))
      CASE "AUTOCAPS"    : scError(%eNone, "AUTOCAPS set to " + IIF$(Prf.AutoCaps, "ON", "OFF"))
      CASE "AUTONUMS"    : scError(%eNone, "AUTONUM set to " + IIF$(Prf.AutoNum, "ON", "OFF"))
      CASE "AUTOSAVE"    : wrd1 =  IIF$((Prf.AutoSave AND 1) = 1, "ON ", "OFF ")
                           wrd1 += IIF$((Prf.AutoSave AND 2) = 2, "PROMPT", "NOPROMPT")
                           scError(%eNone, "AUTOSAVE set to " + wrd1)
      CASE "BOUNDS"      : scError(%eNone, "BOUNDS set to " + FORMAT$(Prf.BndLeft) + " thru " + IIF$(Prf.BndRight = 0, "MAX", FORMAT$(Prf.BndRight)))
      CASE "CAPS"        : IF Prf.CapsDesired <> 2 THEN
                              scError(0, "CAPS set to " + IIF$(Prf.CapsDesired, "ON", "OFF"))
                           ELSE
                              scError(0, "CAPS set to " + IIF$(Prf.CapsActual, "AUTO:on", "AUTO:off"))
                           END IF
      CASE "CASE"        : scError(%eNone, "CASE set to " + UUCASE(Prf.PCase))
      CASE "CHANGE"      : scError(%eNone, "CHANGE set to " + IIF$(Prf.ChangeMode = "D", "DS", "CS"))

      CASE "COLLATE"     : scError(%eNone, "COLLATE set to " + Prf.PCollate)
      CASE "COLS"        : scError(%eNone, "COLS set to " + IIF$(Prf.Cols, "ON", "OFF"))
      CASE "ENUMWITH"    : scError(%eNone, "ENUMWITH set to " + FORMAT$(gEnumWith))
      CASE "EOL"         : Wrd1 = Prf.EOL
                           IF VERIFY(Wrd1, $Hex) = 0 THEN Wrd1 = "X'" + Wrd1 + "'"
                           scError(%eNone, "EOL set to " + UUCASE(Wrd1))
      CASE "FOLD"        : scError(%eNone, "FOLD set to " + IIF$(Prf.Fold, "ON", "OFF"))
      CASE "GLUEWITH"    : scError(%eNone, "GLUEWITH set to: " + $DQ + ENV.GlueWith + $DQ)
      CASE "HEX"         : scError(%eNone, "HEX set to " + IIF$(Prf.HexMode = 4, "ON", "OFF"))
      CASE "HIDE"        : scError(%eNone, "HIDE set to " + IIF$(ISTRUE IsTPHideFlag, "ON", "OFF"))
      CASE "HILITE"      : IF ISTRUE Prf.HiFind AND ISTRUE Prf.HiAuto THEN
                              scError(%eNone, "HILITE FIND & AUTO both set to ON")
                           ELSEIF ISFALSE Prf.HiFind AND ISFALSE Prf.HiAuto THEN
                              scError(%eNone, "HILITE FIND & AUTO both set to OFF")
                           ELSE
                              scError(%eNone, "HILITE FIND set to " + IIF$(Prf.HiFind, "ON", "OFF") + _
                                                   ", AUTO set to " + IIF$(Prf.HiAuto, "ON", "OFF"))
                           END IF
      CASE "KB"          : GOSUB KBQuery
      CASE "LRECL"       : scError(%eNone, "LRECL set to " + FORMAT$(Prf.LRECL))
      CASE "MARK"        : scError(%eNone, "MARK set to " + IIF$(Prf.PMark, "ON", "OFF"))
      CASE "MINLEN"      : scError(%eNone, "MINLEN set to " + FORMAT$(Prf.MINLEN))
      CASE "NOTIFY"      : scError(%eNone, "NOTIFY set to " + SWITCH$(ENV.NotifyLevelT = 0, "None", ENV.NotifyLevelT = 1, "Edit", ENV.NotifyLevelT = 2, "All"))
      CASE "NUMTYPE"     : scError(%eNone, "NUMTYPE set to " + IIF$(Prf.NumType = "", "NONE", Prf.NumType))
      CASE "PAGE"        : scError(%eNone, "PAGE set to " + IIF$(Prf.PageFlag, IIF$(Prf.PageFlag = 2, "SCROLL", "ON"), "OFF") + IIF$(Prf.PageOffset <> 0, "/" + FORMAT$(Prf.PageOffset, "+#;-#"), ""))
      CASE "PRESERVE"    : scError(%eNone, "PRESERVE set to " + IIF$(Prf.PPreserve = 0 , "OFF", IIF$(Prf.PPreserve = 1, "ON", "C")))
      CASE "RECFM"       : scError(%eNone, "RECFM set to " + Prf.RECFM)
      CASE "SETUNDO"     : scError(%eNone, "SETUNDO set to " + FORMAT$(Prf.UndoNumber))
      CASE "SOURCE"      : scError(%eNone, "SOURCE set to " + Prf.PSource)
      CASE "START"       : scError(%eNone, "START set to " + Prf.Start)
      CASE "STATE"       : scError(%eNone, "STATE set to " + CHOOSE$(Prf.PState + 1, "OFF", "ON", "FEW", "MOST"))
      CASE "SUBARG"      : scError(%eNone, "SUBARG is " + IIF$(ISNULL(Prf.SubArg), "OFF", Prf.SubArg))
      CASE "SUBCMD"      : scError(%eNone, "SUBCMD is " + IIF$(ISNULL(Prf.SubCmd), "OFF", Prf.SubCmd))
      CASE "TABS"        : scError(%eNone, "TABS set to " + IIF$(Prf.Tabs, "ON", "OFF"))
      CASE "XTABS"       : scError(%eNone, "XTABS set to " + FORMAT$(Prf.ImportTabs))
      CASE ELSE          : scError(%eFail, "Unknown QUERY operand: " + pCmdOps(1))
   END SELECT
   MExitMeth

KBQuery:
   IF pCmdNumOps <> 2 THEN _                                      ' If no key operand
      MErrExit(%eFail, "Missing QUERY KB operand")                '
   key = UUCASE(pCmdOps(2)): t = ""                               '
   FOR i = 107 TO 1 STEP -1                                       ' Loop through the Key Table
      t = UUCASE(KbdT.NData(i))                                   ' Get the KB string
      IF t <> "(NULL)" THEN                                       ' Other than (NULL)?
         GOSUB KBSearch                                           ' Do the search
         IF ISNOTNULL(tt) THEN                                    ' If this key contains string
            scError(%eNone, "'" + tt + "' found in " + TRIM$(KbdT.Labl(i)) + "; Map: " + t)
         END IF                                                   '
      END IF                                                      '
      t = UUCASE(KbdT.SData(i))                                   ' Get the KB string
      IF t <> "(NULL)" THEN                                       ' Other than (NULL)?
         GOSUB KBSearch                                           ' Do the search
         IF ISNOTNULL(tt) THEN                                    ' If this key contains string
            scError(%eNone, "'" + tt + "' found in S-" + TRIM$(KbdT.Labl(i)) + "; Map: " + t)
         END IF                                                   '
      END IF                                                      '
      t = UUCASE(KbdT.CData(i))                                   ' Get the KB string
      IF t <> "(NULL)" THEN                                       ' Other than (NULL)?
         GOSUB KBSearch                                           ' Do the search
         IF ISNOTNULL(tt) THEN                                    ' If this key contains string
            scError(%eNone, "'" + tt + "' found in C-" + TRIM$(KbdT.Labl(i)) + "; Map: " + t)
         END IF                                                   '
      END IF                                                      '
      t = UUCASE(KbdT.AData(i))                                   ' Get the KB string
      IF t <> "(NULL)" THEN                                       ' Other than (NULL)?
         GOSUB KBSearch                                           ' Do the search
         IF ISNOTNULL(tt) THEN                                    ' If this key contains string
            scError(%eNone, "'" + tt + "' found in A-" + TRIM$(KbdT.Labl(i)) + "; Map: " + t)
         END IF                                                   '
      END IF                                                      '
      t = UUCASE(KbdT.SCData(i))                                  ' Get the KB string
      IF t <> "(NULL)" THEN                                       ' Other than (NULL)?
         GOSUB KBSearch                                           ' Do the search
         IF ISNOTNULL(tt) THEN                                    ' If this key contains string
            scError(%eNone, "'" + tt + "' found in SC-" + TRIM$(KbdT.Labl(i)) + "; Map: " + t)
         END IF                                                   '
      END IF                                                      '
      t = UUCASE(KbdT.SAData(i))                                  ' Get the KB string
      IF t <> "(NULL)" THEN                                       ' Other than (NULL)?
         GOSUB KBSearch                                           ' Do the search
         IF ISNOTNULL(tt) THEN                                    ' If this key contains string
            scError(%eNone, "'" + tt + "' found in SA-" + TRIM$(KbdT.Labl(i)) + "; Map: " + t)
         END IF                                                   '
      END IF                                                      '
      t = UUCASE(KbdT.SCAData(i))                                 ' Get the KB string
      IF t <> "(NULL)" THEN                                       ' Other than (NULL)?
         GOSUB KBSearch                                           ' Do the search
         IF ISNOTNULL(tt) THEN                                    ' If this key contains string
            scError(%eNone, "'" + tt + "' found in SCA-" + TRIM$(KbdT.Labl(i)) + "; Map: " + t)
         END IF                                                   '
      END IF                                                      '
      t = UUCASE(KbdT.CAData(i))                                  ' Get the KB string
      IF t <> "(NULL)" THEN                                       ' Other than (NULL)?
         GOSUB KBSearch                                           ' Do the search
         IF ISNOTNULL(tt) THEN                                    ' If this key contains string
            scError(%eNone, "'" + tt + "' found in CA-" + TRIM$(KbdT.Labl(i)) + "; Map: " + t)
         END IF                                                   '
      END IF                                                      '
   NEXT i
   IF ErrMsgTblC > 0 THEN RETURN                                  ' We issued a message
   MErrExit(%eFail, "'" + pCmdOps(2) + "'" + " was not found in any assigned key")
   RETURN

KBSearch:
   tt = ""                                                        ' Null answer for not found
   j = INSTR(t, key)                                              ' See if this key contain string
   IF j = 0 THEN RETURN                                           ' No, return not found
   IF j = 1 THEN                                                  ' Found at left end?
      k = j                                                       ' Look for delimiter
      DO WHILE k <= LEN(t) AND INSTR($AlphaNum, MID$(t, k, 1)) <> 0'
         INCR k                                                   '
      LOOP                                                        '
      tt = MID$(t, j TO k - 1)                                    ' Pick up found word
   ELSEIF j + LEN(key) - 1 = LEN(t) THEN                          ' Found at right hand end?
      tt = MID$(t, j)                                             ' Pick up found word
   ELSE                                                           ' Found in the middle
      k = j                                                       ' Look for delimiter
      DO WHILE k >= 1 AND INSTR($AlphaNum, MID$(t, k, 1)) <> 0    '
         DECR k                                                   '
      LOOP                                                        '
      j = IIF(k = 0, 1, k + 1)                                    ' Set left end
      k = j                                                       ' Look for delimiter
      DO WHILE k <= LEN(t) AND INSTR($AlphaNum, MID$(t, k, 1)) <> 0'
         INCR k                                                   '
      LOOP                                                        '
      tt = MID$(t, j TO k - 1)                                    ' Pick up found word
   END IF                                                         '
   RETURN
END METHOD

METHOD  pCmdRCHANGE(pCmd AS STRING)
'---------- RCHANGE command
LOCAL lclCmd, lclWord, modemsg, str AS STRING, lclCursMode AS INTEGER, AllCount, i AS LONG
   MEntry
   lclCmd = pCmd                                                  ' Get local copy
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   lclCursMode = Zone                                             ' Save it
   me.AttrInvClearFind                                            ' Clear any Find HiLite
   IF pCmdNumOps = 0 THEN                                         ' Just a simple RCHANGE with no operands?
      IF LocFind = "SPLIT" THEN pCmdRSplit(pCmd): MExitMeth       ' Peel off the other 'change' types
      IF LocFind = "JOIN" THEN pCmdRJoin(pCmd): MExitMeth         '
      IF LocFind = "DELETE" OR LocFind = "DEL" THEN pCmdRDelete(pCmd): MExitMeth
      LocFind = "FIND"                                            ' Remember what Loc / FIND command was last
      IF ISNULL(cfFind) OR ISNULL(cfOChange) THEN _               '
         MErrExit(%eFail, "No prior CHANGE to use")               '
      IF PrevPCmd = "FIND" OR PrevPCmd = "F" OR PrevPCmd = "RFIND" OR PrevPCmd = "RLOCFIND" THEN ' If the last command was a FIND
         IF cfNotFound = 0 THEN                                   ' And last FIND was successful?
            me.Change()                                           ' Yes, Just do the Change
         ELSE                                                     '
            scError(%eFail, "Last search was not successful")     '
         END IF                                                   '
      ELSE                                                        '
         IF me.Search(1, lclCursMode) THEN                        ' Go see if we can find it
            me.Change()                                           ' Go do the change
         ELSE                                                     ' Not found, issue message
            scError(nMac(%eFail), IIF$(IsCfNext, "Bottom of data reached", "Top of data reached"))
            cfNotFound = IIF(IsCfNext, 2, 1)                      ' Remember we failed
         END IF                                                   '
      END IF                                                      '
   ELSE                                                           ' RCHANGE with operands, treat as CHANGE
      LocFind = "FIND"                                            ' Remember what Loc / FIND command was last
      lclWord = LSET$(UUCASE(pCmdOps(1)), 8)                      ' Get next word to an 8 char field
      IF INSTR("CHANGE  C       CHA     CHG     ", lclWord) <> 0 THEN ' A 'pass-thru' CHANGE command?
         lclWord = sGetWord(lclCmd, %Strip, %QuoteNotSig)         ' Remove RCHANGE from the front
         TP.PCommand = lclCmd                                     '
         me.CmdParse(lclCmd)                                      ' Do the basic parsing
      ELSE                                                        '
         MErrExit(%eFail, "RCHANGE has no operands")              '
      END IF                                                      '

      lclCursMode = Zone                                          ' Save CursMode at start of command
      IF ISTRUE me.RangeSet("", %False, %False) THEN MExitMeth    ' Setup any range command operands
      me.FindReset                                                ' Reset the parse output fields
      IF sCRTParse("ASDMEC12LPQHU") THEN MExitMeth                ' Scan (All, Subset, Direct, Modifier, Cols, Lit1, Lit2, Left/Right, Pen, CPen, Shift, User)
      me.FindSetup                                                ' Copy CRT parse's answers

      IF IsCfALL THEN                                             ' If ALL switch it to FIRST
         CfSet(%CRTFirst)                                         ' Set FIRST
         CfClear(%CRTALL)                                         ' Clear ALL
      END IF

      IF ISFALSE IsCfNext AND ISFALSE IsCfPrev AND _              '
         ISFALSE IsCfFirst AND ISFALSE IsCfLast THEN _            '
         CfSet(%CRTNext)                                          ' Set NEXT if no others specified

      IF ISFALSE IsCfLit1 THEN MErrExit(%eFail, "Missing CHANGE search string")    ' Lit1 present?
      IF ISFALSE IsCfLit2 THEN MErrExit(%eFail, "Missing CHANGE new value string") ' Lit2 present?
      str = IIF$(LEN(CrtL1Raw) < 16, CrtL1Raw, LEFT$(CrtL1Raw, 14) + "..")         ' Build msg literal
      modemsg = IIF$(IsCfWord, "WORD " + str, "CHARS " + str)                      ' Build mode part

      IF IsCfAll THEN                                             ' The ALL version?
         AllCount = 0                                             ' Reset counter
         i = me.Search(0, lclCursMode)                            ' Do the initial search
         lclCursMode = %CursBad                                   ' Make it bad for subsequent searches
         IF i THEN                                                ' See if we have one
            INCR AllCount                                         ' Count successful find
            me.Change()                                           ' Go do the Change
         END IF                                                   '
         DO WHILE i                                               ' Do till not found
            i = me.Search(1, lclCursMode)                         ' Do re-Finds
            IF i THEN                                             '
               INCR AllCount                                      ' Count successful finds
               me.Change()                                        '
            END IF                                                '
         LOOP                                                     '
         IF AllCount = 0 THEN                                     ' Issue approp. message
            scError(nMac(%eFail), modemsg + " not found")               '
         ELSE                                                     '
            scError(0, modemsg + " changed " + FORMAT$(AllCount) + IIF$(AllCount > 1, " times", " time"))
         END IF                                                   '
      ELSE                                                        ' The non-All version
         IF me.Search(0, lclCursMode) THEN                        ' Go see if we can find it
            me.Change()                                           ' Go do the change
         ELSE                                                     ' Not found, issue message
            scError(nMac(%eFail), IIF$(IsCfNext, "Bottom of data reached", "Top of data reached"))                             '
            cfNotFound = IIF(IsCfNext, 2, 1)                      ' Remember we failed
         END IF                                                   '
      END IF                                                      '
   END IF
   MExit
END METHOD

METHOD  pCmdRDelete(pCmd AS STRING)
'---------- RDELETE Command
LOCAL lclCmd, lclWord, modemsg, str AS STRING, lclCursMode AS INTEGER
LOCAL i, ll AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   lclCursMode = Zone                                             ' Save CursMode at start of command
   me.AttrInvClearFind                                            ' Clear any Find HiLite

   IF pCmdNumOps = 0 THEN                                         ' Just a simple RDelete with no operands?
      IF LocFind <> "DELETE" AND LocFind <> "DEL" THEN _          ' Last better be DELETE
         MErrExit(%eFail, "Last search was NOT a DELETE command")
      IF cfNotFound = 0 THEN                                      ' Last FIND was successful?
         i = sLine                                                ' Get the found line
         GOSUB Deli                                               ' Delete it
      ELSE                                                        '
         i = me.Search(0, lclCursMode)                            ' Go see if we can find it
         IF i = 0 THEN                                            ' Got it?
            scError(nMac(%eFail), IIF$(IsCfNext, "Bottom of data reached", "Top of data reached"))
            cfNotFound = IIF(IsCfNext, 2, 1)                      ' Remember we failed
         ELSE                                                     '
            i = sLine                                             ' Get found line
            GOSUB Deli                                            ' Delete it
         END IF                                                   '
      END IF                                                      '
   ELSE                                                           '
      scError(%eFail, "RDELETE has no parameters")                '
   END IF                                                         '
   MExitMeth

Deli:
   me.LTxtFree(i)                                                 ' Go free the dynamic string
   me.LEntDel(i)                                                  ' Remove from the L() array
   me.ModSet(i)                                                   ' Remember we changed something
   me.AdjustPending(i, -1, 0)                                     ' Adjust pending stuff
   me.CurSetReq(%Position, i, 1, %True)                           ' Set cursor set attempt
   DECR LastLine: DECR LastReal                                   ' Adjust LastLine
   scError(%eNone, "Line deleted")                                '
   RETURN
END METHOD

METHOD PCmdRecall(pCmd AS STRING)
'---------- RECALL in Normal Mode
   MEntry
   gTabSwitchCmd = pCmd                                           ' Pass command to the FM tab
   gTabSwitch = 1                                                 ' Switch to FM tab
   MExit                                                          '
END METHOD

METHOD  pCmdRECFM(pCmd AS STRING)
'---------- Set RECFM
LOCAL Wrd1, lclCmd AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN _                                       ' If no operands
      MErrExit(%eNone, "RECFM set to " + Prf.RECFM)               ' Tell user
   SELECT CASE UUCASE(pCmdOps(1))                                 '
      CASE "U", "F", "V", "VBI", "VLI"                            '
         '                                                        '
      CASE ELSE                                                   '
         MErrExit(%eFail, "Invalid RECFM operand")                '
   END SELECT                                                     '

   Prf.RECFM = UUCASE(pCmdOps(1))                                 ' Set it
   SELECT CASE Prf.RECFM                                          ' Check stuff out
      CASE "U"                                                    '
         IF Prf.EOL = "NONE" THEN                                 ' Can't be NONE
            Prf.EOL = "CRLF"                                      ' Set default
            scError(%eNone, "RECFM set to 'U', EOL NONE changed to EOL CRLF" + IIF$(Prf.LRECL > 0, "", ", Warning: LRECL > 0"))
         ELSE                                                     '
            scError(%eNone, "RECFM set to 'U'" + IIF$(Prf.LRECL > 0, "", ", Warning: LRECL > 0"))
         END IF                                                   '
      CASE "F"                                                    '
         scError(%eNone, "RECFM set to 'F'" + IIF$(Prf.LRECL > 0, "", ", Warning: LRECL = 0"))
      CASE "V", "VBI", "VLI"                                      '
         Prf.EOL = "NONE": Prf.LRECL = 0                          '
         scError(%eNone, "RECFM set to 'V', EOL set to NONE, LRECL set to 0")
   END SELECT                                                     '
   MExit                                                          '
END METHOD

METHOD  pCmdRELOAD(pCmd AS STRING)
'---------- Re-Load the current file
LOCAL i, j AS LONG, NewCmd AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF IsClip OR IsSetEdit THEN MErrExit(%eFail,"ReLoad allowed in Edit/Browse/View mode only")  ' Issue error message
   IF TIP.FilePath = $Empty THEN MErrExit(%eFail,"(Empty) files can't be reloaded")              ' Issue error message
   IF ISFALSE IsMEdit THEN                                        ' Simple non MEdit session?
      IF IsModdFlag THEN                                          ' Yes, was data modified?
         i = sDoMsgBox("RELOAD will discard all your current changes." + $CRLF + _
                   "Do you want to continue?", %MB_YESNO + %MB_USERICON, "SPFLite")
         IF i = %IDNO THEN MExitMeth                              '
      END IF                                                      '
      me.UnWatchQueue("")                                         ' Kill Watch and dequeue it
      me.LInitTxtData(TIP.FilePath)                               ' Wipe everything out then
      me.InitaFile(%False)                                        ' Go load it in
      me.UndoInit                                                 ' Init the Undo file names
      me.UndoSave()                                               ' Take an initial one
      TIP.TimeDateRefresh                                         ' Get refreshed Date/Time
      me.WindowTitle                                              ' Alter window/Tab titles
      MExitMeth                                                   '
   END IF                                                         '

   '----- Do things the MEdit way
   IF IsModdFlag THEN                                             ' Yes, was data modified?
      i = sDoMsgBox("RELOAD will lose all your current changes" + $CRLF + _
                "in |KALL|B the loaded files." + $CRLF + _
                "Do you want to continue?", %MB_YESNO + %MB_USERICON, "SPFLite")
      IF i = %IDNO THEN MExitMeth                                 '
   END IF                                                         '
   NewCmd = "MEDIT "                                              ' Start NewCmd
   FOR j = 1 TO MEditCount                                        ' Build a MEdit command to reload things
      NewCmd += $DQ + me.MEditListGet(j) + $DQ + " "              '
      me.UnWatchQueue(me.MEditListGet(j))                         ' Kill Watch and dequeue it
   NEXT j                                                         '
   me.LInitTxtData("")                                            ' Wipe everything out then
   pCmdMEDIT(NewCmd)                                              ' Let MEdit do the work
   MExit
END METHOD

METHOD  pCmdRENAME(pCmd AS STRING)
'---------- Rename the current file
LOCAL i, j, FNum AS LONG
LOCAL lclPath, lclDrive, RenDrive, RenPath, DelFN, RenFn, fn, LErrMsg, lclOldState, lclNewState AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF IsClip OR IsSetEdit OR IsMEdit THEN _                       ' Not CLIP or SET EDit or MEdit mode
      MErrExit(%eFail,"Rename invalid in CLIP/Set Edit/Multi-Edit mode") ' Issue error message
   RenFn = IIF$(pCmdNumOps = 0, "", pCmdOps(1))                   ' Get filename if present

   lclPath = CURDIR$                                              ' Get current path
   IF MID$(lclPath, 2, 1) = ":" THEN _                            ' Extract Drive if present
      lclDrive = LEFT$(lclPath, 2)                                ' and save it

   RenPath = TIP.Path                                             ' Build Path
   IF MID$(RenPath, 2, 1) = ":" THEN                              ' Extract Drive if present
      RenDrive = LEFT$(RenPath, 2)                                ' and save it
      RenPath = MID$(RenPath, 3)                                  '
   END IF                                                         '

   DelFn = TIP.File                                               ' Get file Name

   '----- If no name provided on the command, prompt for one
   IF ISNULL(RenFn) THEN                                          ' Got a new name?
      RenFn = sDoInputBox("Enter the new FileName: ", "Rename", RTRIM$(DelFn))
      IF ISNULL(RenFn) OR IsEQ(RenFn, RTRIM$(DelFn)) THEN MErrExit(%eNone, "Rename cancelled")
   END IF                                                         '

   '----- Ensure it's not in use elsewhere
   IF VAL(sFileQueue("S", " ", RenDrive + RenPath + RenFn)) > 0 THEN _ ' Open in some tab?
     MErrExit(%eFail,"File open elsewhere: " + Renfn)             '

   '----- Do the rename now

   '----- First save STATE names if they exist
   IF IsStateExist THEN                                           ' Does STATE even exist?
      lclOldState = RenDrive + RenPath + DelFN + ".STATE"         ' Build the STATE filename
      REPLACE ANY ":\/" WITH "```" IN lclOldState                 ' Make : / and \ into `
      lclOldState = ENV.StatePath + lclOldState                   ' Add our STATE folder
      IF ISFALSE ISFILE(lclOldState) THEN lclOldState = ""        ' Kill it if no STATE file
      lclNewState = RenDrive + Renpath + RenFn + ".STATE"         ' Build the new STATE filename
      REPLACE ANY ":\/" WITH "```" IN lclNewState                 ' Make : / and \ into `
      lclNewState = ENV.StatePath + lclNewState                   ' Add our STATE folder
      IF ISFILE(lclNewState) THEN lclOldState = ""                ' Kill it if no STATE file
   END IF

   '----- Continue with the file rename
   IF lclDrive <> RenDrive THEN CHDRIVE RenDrive                  '
   CHDIR RenPath                                                  '
   i = me.FileWatch("", %WatchEnd)                                ' Kill any prior Watch
   IF TRIM$(DelFn) <> "" AND DelFN <> $Empty THEN NAME DelFn AS RenFn  ' Do the rename
   LErrMsg = ""                                                   '
   IF ERR = %ERR_FILENOTFOUND THEN                                '
      LErrMsg = "File no longer found??"                          '
   ELSEIF ERR = %ERR_FILEALREADYEXISTS THEN                       '
      LErrMsg = "File already exists."                            '
   ELSEIF ERR = %ERR_PERMISSIONDENIED THEN                        '
      LErrMsg = "File is Read Only."                              '
   ELSEIF ERR = %ERR_PATHFILEACCESSERROR THEN                     '
      LErrMsg = "Path/File Access error."                         '
   ELSEIF ERR <> 0 THEN                                           '
      LErrMsg = "UnKnown Error - Err: " + FORMAT$(ERR)            '
   END IF                                                         '
   IF ISNOTNULL(LErrMsg) THEN _                                   ' Any error?
      MErrExit(%eFail, "Rename failed, " + LErrMsg)               '
   sFileListRename(RenDrive + Renpath + DelFN, RenDrive + Renpath + RenFn) ' Go adjust FILELIST entries

   scError(0, IIF$(ISNULL(DelFn), "File Saved", "Rename complete"))
   fn = RenDrive + RenPath + RenFn                                ' Get a full filename

   IF RenDrive <> lclDrive THEN CHDRIVE lclDrive                  ' Switch drive if needed
   CHDIR lclPath                                                  ' put back the original path

   TIP.Setup("*", "*", "", fn)                                    ' Alter TIP
   me.WindowTitle                                                 ' Alter window/Tab titles

   OnModdFlag                                                     ' Set (M)
   Prf.SetProfName(TIP.Profile, %True)                            ' Set Profile name, create if missing
   Prf.ReadAll(%True)                                             ' Go read this file type's INI values
   me.ClrLoad()                                                   ' Go load Colorize data
   IF lclOldState <> "" THEN                                      ' Doing STATE rename?
      NAME lclOldState AS lclNewState                             ' Yes, do it
   END IF                                                         '

   IF ISNULL(DelFn) OR DelFn = $Empty THEN                        ' If renaming (Empty)
      sMakeNullFile(TIP.FilePath)                                 ' Create null file
      me.FindPushPop("PUSH")                                      ' Save FIND stuff
      me.FindSetAll                                               ' Set parse search output fields for ALL
      me.RangeSetRange(2, LastLine - 1)                           ' Everything for non=MEdit mode
      TOP.Setup("O", TIP.Profile, "", TIP.FilePath)               ' Setup for write
      me.WriteFile(%False)                                        ' Go Write the file
      me.FindPushPop("POP")                                       ' Restore FIND stuff
      OffModdFlag                                                 '
   END IF                                                         '

   sFileQueue("A", " ", TIP.FilePath)                             ' Establish the watch
   IF me.FileWatch(TIP.FilePath, %WatchStart) THEN _              '
      scError(0, "File watch could not be established")           '

   MExit
END METHOD

METHOD  pCmdReNumber(pCmd AS STRING)
'---------- Do the RENUM command
LOCAL x, xList, NumType, t, newnum AS STRING
LOCAL i, seqCol, seqLen, wSeqCol, trunc, real, numfirst, numincr, quiet AS LONG
   MEntry
   '----- Check for QUIET mode
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps > 0 AND pCmdOps(1) = "QUIET" THEN quiet = %True  ' Remember QUIET mode

   '----- See if a valid NUMTYPE exiats
   NumType = TRIM$(Prf.NumType)                                   ' Get local copy of NumType
   IF ISNULL(NumType) THEN                                        ' Is NUMTYPE null?
      IF ISTRUE quiet THEN MExitMeth                              ' If QUIET, simply exit
      MErrExit(%eFail, "There is no valid NUMTYPE available")     ' Oops!
   END IF                                                         '

   '----- Make sure the file extension is an editable file
   x = "," + UUCASE(MID$(TP.TIPExtn, 2)) + ","                    ' Get extension surrounded by commas
   XList = "," + UUCASE(ENV.FMFFList) + ","                       ' Build XList here since a ProfSkip might have added an entry
   XList = REMOVE$(XList, " ")                                    ' Eliminate any spaces
   IF INSTR(XList, x) THEN                                        ' Is tis file in the exempt list?
      IF ISTRUE quiet THEN MExitMeth                              ' If QUIET, simply exit
      MErrExit(%eFail, "Current filetype is in the non-editable filetype list") ' Oops!
   END IF

   '----- Check for ragged right sequence numbers
   IF INSTR(NumType, "-") AND ISFALSE quiet THEN                  ' Negative columns?
      i = sDoMsgBox("With the current RECFM (" + Prf.RECFM + ")," + $CRLF + _
                 "the specified NUMTYPE (" + NumType + ") could result in" + $CRLF + _
                 "sequence numbers in differing columns.  Do you want to continue?", _
                 %MB_YESNO + %MB_USERICON OR %MB_DEFBUTTON2, "SPFLite")
      IF i = %IDNO THEN                                           '
         MErrExit(%eNone, "Numbering cancelled")                  ' Exit if NO entered
      END IF                                                      '
   END IF                                                         '

   '----- Setup the sequence columns
   seqCol = VAL(PARSE$(NumType, " ", 1))                          ' Extract start column
   seqLen = VAL(PARSE$(NumType, " ", 2))                          ' Extract end column
   seqLen = ABS(ABS(seqLen) - ABS(seqCol)) + 1                    ' Calc length of sequence field
   IF INSTR(NumType, "TRUNC") THEN trunc = %True                  ' Extract TRUNC
   IF INSTR(NumType, "REAL") THEN real = %True                    ' and REAL flags
   GOSUB SetupRenum                                               ' Go set Starting and Incr values

   '----- Finally, do the renum
   FOR i = 1 TO LastLine                                          ' Spin through the lines
      IF IsLFile(i) THEN                                          ' An MEdit FILE line?
         GOSUB SetupRenum                                         ' Start Renumbers over again
      END IF                                                      '
      IF IsLData(i) THEN                                          ' Only Data lines
         t = LTxtG(i)                                             ' Get the text line
         wSeqCol = seqCol                                         ' Default sequence column
         IF seqCol < 0 THEN wSeqCol =  LEN(t) + seqCol + 1        ' If negative columns, make it an actual
         IF wSeqCol + seqLen -1 > LEN(t) THEN _                   ' If not long enough for a Seq field, it's an eror
            t = LSET$(t, wSeqCol + seqLen -1)                     ' Lengthen line if too short
         newnum = FORMAT$(NumFirst, REPEAT$(seqlen, "0"))         ' Format the new number
         IF Real THEN  newnum = RIGHT$(LLNumG(i), seqlen)         '
         MID$(t, wSeqCol, seqLen) = newnum                        ' Stuff it in the line
         IF Trunc THEN t = LEFT$(t, wSeqCol + seqLen - 1)         ' If trunc, do so
         me.LTxtSet(i, t)                                         ' Stuff it back
         NumFirst += NumIncr                                      '
      END IF                                                      '
   NEXT i                                                         '
   IF ISFALSE quiet THEN MErrExit(%eNone, "RENUM succesful")      ' If not QUIET, issue a message

   MExitMeth

SetupRenum:
   '----- Figure out renum values
   SELECT CASE AS LONG LastLine                                   ' Base on size of file
      CASE < 1000                                                 ' 3 Digit needed
         SELECT CASE AS LONG seqLen                               ' Now base on field length
            CASE 4:  NumFirst = 10:   NumIncr = 10                '
            CASE 5:  NumFirst = 100:  NumIncr = 100               '
            CASE 6:  NumFirst = 100:  NumIncr = 100               '
            CASE 7:  NumFirst = 100:  NumIncr = 100               '
            CASE 8:  NumFirst = 100:  NumIncr = 100               '
         END SELECT                                               '
      CASE < 10000                                                ' 4 Digit needed
         SELECT CASE AS LONG seqLen                               ' Now base on field length
            CASE 4:  NumFirst = 1:    NumIncr = 1                 '
            CASE 5:  NumFirst = 10:   NumIncr = 10                '
            CASE 6:  NumFirst = 100:  NumIncr = 100               '
            CASE 7:  NumFirst = 100:  NumIncr = 100               '
            CASE 8:  NumFirst = 100:  NumIncr = 100               '
         END SELECT                                               '
      CASE < 100000                                               ' 5 Digit needed
         SELECT CASE AS LONG seqLen                               ' Now base on field length
            CASE 4:  NumFirst = 1:    NumIncr = 1                 '
            CASE 5:  NumFirst = 1:    NumIncr = 1                 '
            CASE 6:  NumFirst = 10:   NumIncr = 10                '
            CASE 7:  NumFirst = 100:  NumIncr = 100               '
            CASE 8:  NumFirst = 100:  NumIncr = 100               '
         END SELECT                                               '
      CASE ELSE                                                   ' 6 Digit needed
         SELECT CASE AS LONG seqLen                               ' Now base on field length
            CASE 4:  NumFirst = 1:    NumIncr = 1                 '
            CASE 5:  NumFirst = 1:    NumIncr = 1                 '
            CASE 6:  NumFirst = 1:    NumIncr = 1                 '
            CASE 7:  NumFirst = 10:   NumIncr = 10                '
            CASE 8:  NumFirst = 100:  NumIncr = 100               '
         END SELECT                                               '
   END SELECT

   RETURN
END METHOD

METHOD  pCmdREPLACE(pCmd AS STRING)
'---------- SAVE data to a specified filename
LOCAL x, lclTop, i AS LONG
   MEntry

   '----- Do the basic parsing stuff
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   lclTop = TopScrn                                               ' Save where we are

   IF IsView AND TIP.FilePath = $Empty THEN _                     ' View mode?
      MErrExit(%eFail, "REPLACE disallowed, you're in VIEW of an Empty file, use CREATE")

   '----- If sGetFileAndRange is successful, we have something to do
   IF me.GetFileAndRange() THEN                                   ' Go get filename etc

      '----- Do a bit of validation
      IF ISFALSE IsCfALL THEN _                                   ' Some kind of line criteria?
         MErrExit(%eFail, "Pending line range")                   '

      '----- Write the file now
      TOP.Setup("OCRE", "", "Select file for REPLACE", "*")       '
      IF me.WriteFile(%False) THEN                                '

         '----- If this is still an (Empty) tab make it a normal one now
         IF TIP.FilePath = $Empty THEN                            ' If still no current file
            TIP.Setup("*", TOP.Profile, "", TOP.FilePath)         ' Swap TOP  to TIP
            me.WindowTitle                                        ' Alter window/Tab titles
            Prf.SetProfName(TIP.Profile, %False)                  ' Setup profile
            Prf.ReadAll(%True)                                    ' Go read this file type's INI values
            me.ClrLoad()                                          ' Get Attr loaded if it exists
            OffModdFlag: OnUndoFlag                               ' Turn off the Modd flag
            i = me.FileWatch("", %WatchEnd)                       ' Kill any prior Watch
            sFileQueue("A", " ", TIP.FilePath)                    ' Add to FQ
            IF me.FileWatch(TIP.FilePath, %WatchStart) THEN _     ' Establish the watch
               scError(0, "File watch could not be established")  '
         END IF                                                   '

         '----- If the line range was a MM (Move) then delete the range
         IF SavFilLCtl = "M" OR SavFilLCtl = "MM" THEN            ' If the line selection was a Take

            '----- But only if we're OK
            IF FileRecs = 0 THEN                                  '
               scError(%eFail, "No records written, File is empty, Record deletion skipped") ' Issue error
            ELSE                                                  '
               me.ModSet(DotStart)                                ' Remember we changed something
               me.LStubDelRange(DotStart, DotEnd)                 ' Delete the range
               me.AdjustPending(DotStart, DotEnd - DotStart + 1, 0) ' Adjust pending stuff
            END IF                                                '
         END IF                                                   '

         '----- Make sure we're in the RECENT list
         sRecentAdd(TOP.FilePath)                                 ' Go add to RECENT list
         DoSet(%LoadReq)                                          ' Request refresh

      END IF                                                      '
   END IF                                                         '
   TopScrn = lclTop                                               '
   MExit
END METHOD

METHOD  pCmdRESET(pCmd AS STRING)
'---------- Do a RESET Command
LOCAL lclTop, i, j, Requests AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   lcltop = TopScrn                                               ' Save where we are
   IF ISTRUE me.RangeSet("", %True, %False) THEN MExitMeth        ' Strip out any line range operands
   ResetClr = 0                                                   ' Clear color to be reset
   IF pCmdNumOps = 0 THEN                                         ' iF no operands
      Requests = %ResetAll - %ResetRetrieve - %ResetLabel - %ResetTag - %ResetCommandK
      IF ISFALSE ENV.ResetU THEN requests -= %ResetUser           ' If not clearing User, then remove it's flag
      me.ResetFunc(Requests)                                      ' Go do the default set
      OffHideFlag                                                 ' Reset Hide as well
      OffSlecSet                                                  ' Clear Slec status
      OffSlecActive                                               '
      me.MarkKill                                                 ' Kill marked areas
      me.SwapKill                                                 '
      IF ISFALSE gMacroMode THEN
         sDoStatusBar($SBSelect)                                  ' re-Do the StatusBar Select box
      END IF
      MExitMeth                                                   '
   ELSE                                                           '
      FOR i = 1 TO pCmdNumOps                                     ' Process each operand
         IF pCmdOpsType(i) = %KWALL AND pCmdNumOps = 1 THEN       ' ALL all by itself
            Requests = %ResetAll - %ResetRetrieve                 ' ALL except Retrieve
            OffHideFlag                                           ' Reset Hide as well
            Prf.WordInput = $WORD                                 ' Also WORD
'            me.ProfWriteAll(%False)                              ' Go update profile if not locked
            OffSlecSet                                            ' Clear Slec status
            OffSlecActive                                         '
            IF ISFALSE gMacroMode THEN
               sDoStatusBar($SBSelect)                            ' re-Do the StatusBar Select box
            END IF
         ELSEIF pCmdOpsType(i) = %KWALL THEN                      ' ALL with something else, ignore it
            '                                                     '
         ELSEIF pCmdOpsType(i) = %KWCHANGE THEN                   ' CHANGE
            Requests = Requests OR %ResetChange                   '
         ELSEIF pCmdOpsType(i) = %KWCOMMAND THEN                  ' COMMAND
            Requests = Requests OR %ResetCommand OR %ResetCommandK'
         ELSEIF pCmdOpsType(i) = %KWEXCLUDE THEN                  ' EXCLUDED / X
            Requests = Requests OR %ResetExcluded                 '
         ELSEIF pCmdOpsType(i) = %KWU THEN                        ' USER / U
            Requests = Requests OR %ResetUser                     '
         ELSEIF pCmdOpsType(i) = %KWTAG THEN                      ' TAG
            Requests = Requests OR %ResetTag                      '
         ELSEIF pCmdOpsType(i) = %KWLABEL THEN                    ' LABEL
            Requests = Requests OR %ResetLabel                    '
         ELSEIF pCmdOpsType(i) = %KWSPECIAL THEN                  ' SPECIAL
            Requests = Requests OR %ResetSpecial                  '
         ELSEIF pCmdOpsType(i) = %KWRETRIEVE THEN                 ' RETRIEVE
            Requests = Requests OR %ResetRetrieve                 '
         ELSEIF pCmdOpsType(i) = %KWFind THEN                     ' FIND
            Requests = Requests OR %ResetFind                     '
         ELSEIF pCmdOpsType(i) = %KWSTD THEN                      ' STD
            Requests = Requests OR %ResetClr                      '
            ResetClr = -1                                         '
         ELSEIF pCmdOpsType(i) = %KWCOLOR THEN                    ' COLOR (alias for STD)
            Requests = Requests OR %ResetClr                      '
            ResetClr = -1                                         '
         ELSEIF pCmdOpsType(i) = %KWSTATE THEN                    ' STATE
            Requests = Requests OR (%ResetLabel OR %ResetTag OR %ResetExcluded) ' Clear all STATE saved info
         ELSEIF pCmdOpsType(i) = %KWHIDE THEN                     ' HIDE
            OffHideFlag                                           ' Set the new value
            scError(0, "HIDE mode OFF")                           ' Issue message
         ELSEIF pCmdOpsType(i) = %KWWORD THEN                     ' WORD
            Prf.WordInput = $WORD                                 ' Reset
'            me.ProfWriteAll(%False)                              ' Go update profile if not locked
            scError(0, "WORD values were reset")                  ' Issue message
         ELSEIF pCmdOpsType(i) = %KWSOURCE THEN                   ' SOURCE
            Prf.PSource = "ANSI"                                  '
            scError(%eNone, "SOURCE set to ANSI")                 ' Issue message
         ELSE                                                     '
         '----- Last chance, is it a Color operand
            ARRAY SCAN nHiLites(), COLLATE UCASE, = pCmdOps(i), TO j ' See if a HiLite Name
            IF j THEN                                             ' A winner
               IF ResetClr <> 0 THEN                              ' Already have a color operand?
                  MErrExit(%eFail, "Extra color name operand detected - " + pCmdOps(i)) ' Issue error
               END IF                                             '
               Requests = Requests OR %ResetClr                   '
               ResetClr = j                                       ' Save the color hilite index
            ELSE                                                  '
               MErrExit(%eFail, "Invalid RESET operand - " + pCmdOps(i))
            END IF                                                '
         END IF                                                   '
      NEXT i                                                      '
      IF Requests <> 0 THEN me.ResetFunc(Requests)                ' Ask sResetFunc to do the specified work
   END IF                                                         '
   i = me.RangeSet("RESET", %False, %False)                       ' Ensure default range of ALL before exit
   OnUndoFlag                                                     ' Remember we changed something
   TopScrn = lclTop                                               ' Go back to where we were
   ErrFlag = %eNone                                               ' Say we 're OK
   scError(0, "")                                                 ' Reset error flag and message
   MExit
END METHOD

METHOD  pCmdRETFWD(pCmd AS STRING)
'---------- Retrieve the oldest command line
LOCAL cmd AS STRING, ix AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   ix = gCmdRtrevIX - 1: IF gCmdRtrevLast = "B" THEN DECR ix      ' Get previous entry
   IF ix < 1 THEN ix = UBOUND(gCmdRtrev())                        ' Rollover if needed
   cmd = TRIM$(gCmdRtrev(ix))                                     ' Copy next command
   IF ISNULL(TRIM$(Cmd)) THEN                                     ' Unused entry
      FOR ix = ix TO 1 STEP -1                                    ' Find the last used entry
         IF ISNOTNULL(TRIM$(gCmdRtrev(ix))) THEN EXIT FOR         ' Any non-null is fine
      NEXT ix                                                     '
      IF ix < 1 THEN ix = UBOUND(gCmdRtrev())                     ' Prevent ix = 0 if all empty
      cmd = TRIM$(gCmdRtrev(ix))                                  ' Copy selected command
   END IF                                                         '
   IF ISNULL(TRIM$(Cmd)) THEN _                                   ' Exit if nothing here
      MErrExit(%eNone, "No saved RETRIEVE commands available")    '
   gCmdRtrevIX = ix                                               ' Swap the used entry
   IF gCmdRtrevIX < 1 THEN gCmdRtrevIX = UBOUND(gCmdRtrev())      ' Setup next entry
   pCommand = LSET$(Cmd, pCmdLen)                                 '
   errFlag = %eRetrieve                                           '
   sCurPrio = 0                                                   ' Reset cursor requests
   gCmdRtrevMsg = "  Cmd: " + FORMAT$(gCmdRtrevIX)                ' Build message
   gCmdRtrevLast = "F"                                            ' Remember last direction
   MExit
END METHOD

METHOD  pCmdRETRIEVE(pCmd AS STRING)
'---------- Retrieve the last command line
LOCAL cmd AS STRING, ix AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF IsEQ(pCmdOps(0), "CRETRIEV") THEN                           ' Is this the CRTRIEV version?
      IF ISFALSE IsCCmnd THEN                                     ' If not in the Command area
         SetCmd                                                   ' Cursor to Cmd line
         MExitMeth                                                ' That's all
      END IF                                                      '
   END IF                                                         ' Else it's a RETRIEVE effectively

   IF gCmdRtrevLast = "F" THEN INCR gCmdRtrevIX                   ' Adjust if direction reverse
   IF gCmdRtrevIX > UBOUND(gCmdRtrev()) THEN gCmdRtrevIX = 1      ' Catch rollover
   cmd = gCmdRtrev(gCmdRtrevIX)                                   ' Get next Retrieve entry
   IF ISNULL(TRIM$(Cmd)) AND ISNULL(TRIM$(gCmdRtrev(1))) THEN _   ' Exit if nothing here
      MErrExit(%eFail, "No saved RETRIEVE commands available")    '
   IF ISNULL(TRIM$(Cmd)) THEN                                     ' Wrap back
      gCmdRtrevIX = 1: cmd = gCmdRtrev(gCmdRtrevIX)               ' To entry 1
   END IF                                                         '
   pCommand = LSET$(Cmd, pCmdLen)                                 '
   errFlag = %eRetrieve                                           '
   sCurPrio = 0                                                   ' Reset cursor requests
   gCmdRtrevMsg = "  Cmd: " + FORMAT$(gCmdRtrevIX)                ' Build message
   INCR gCmdRtrevIX: IF gCmdRtrevIX > UBOUND(gCmdRtrev()) THEN gCmdRtrevIX = 1 ' Bump, rollover if needed
   gCmdRtrevLast = "B"                                            ' Remember last direction
   MExit
END METHOD

METHOD  pCmdRFIND(pCmd AS STRING)
'---------- RFIND Command
LOCAL lclCmd, lclWord, modemsg, str, fstr, fstr2 AS STRING, lclCursMode AS INTEGER, RevOnce AS LONG
LOCAL i, ll AS LONG
   MEntry
   lclCmd = pCommand
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   lclCursMode = Zone                                             ' Save CursMode at start of command
   me.AttrInvClearFind                                            ' Clear any Find HiLite
   IF pCmdNumOps = 1 AND pCmdOpsType(1) = %KWREVERSE THEN         ' Is this the simple REVERSE?
      OnRevOnce: RevOnce = %True                                  ' Reverse direction this time,remember locally
      pCmdNumOps = 0                                              ' Pretend there are no operands now
   END IF                                                         '

   IF pCmdNumOps = 0 THEN                                         ' Just a simple RFIND with no operands?
      IF ISNULL(cfFind) THEN _                                    ' We'd better have something from the previous FIND
         MErrExit(%eFail, "No prior search values to use")        '
      IF IsCfNF THEN                                              ' Was previous an NFIND command
         pCmdRNFIND(pCmd): MExitMeth                              ' Pass it off
      END IF                                                      '

      IF IsCData THEN                                             ' In the Data Area?
         ll = sGetIX(CRow)                                        ' Get line cursor is on
         IF ll AND IsLXclude(ll) THEN                             ' Is this an Xclude line?
            IF CsrLinDX <> 0 THEN ll = sFindLineNum(CsrLinDX)     ' Continuing from within an X'd block?  Use that line number then
         END IF                                                   '
         IF sCol = CCol - gLNPadCol + Offset AND _                ' Same cursor location as last found?
            sLine = ll THEN                                       '
            lclCursMode = %CursBad                                ' Tell StrSearch to ignore it
         END IF                                                   '
      END IF                                                      '
      IF LEN(CrtL1Raw) < 16 THEN                                  ' Build msg literal
         str = CrtL1Raw                                           '
      ELSE                                                        '
         str = LEFT$(CrtL1Raw, 14) + ".."                         '
      END IF                                                      '
      modemsg = IIF$(IsCfWord, "WORD " + str, "CHARS " + str)     '

      IF ISFALSE me.Search(1, lclCursMode) THEN                   ' Go see if we can find it
         IF RevOnce THEN                                          ' Issue correct message
            scError(nmac(%eFail), IIF$(IsCfNext, "Top of data reached", "Bottom of data reached"))
            cfNotFound = IIF(IsCfNext, 1, 2)                      ' Remember we failed
         ELSE                                                     '
            scError(nmac(%eFail), IIF$(IsCfNext, "Bottom of data reached", "Top of data reached"))
            cfNotFound = IIF(IsCfNext, 2, 1)                      ' Remember we failed
         END IF                                                   '
         MExitMeth                                                '
      END IF                                                      '

      i = sLine                                                   ' Get found line number
      lclword = MID$(LTxtG(i), sCol, cfFLen)                      '
      IF IsCfMX THEN                                              ' X'ing the line?
         scError(%eNone, modemsg + " found and X'd")              '
      ELSE                                                        '
         fstr = MID$(LTxtG(cfFLine), cfFCol, cfFLen)              ' Get the found string
         IF IsCfL1Picture OR IsCfL1Hex OR IsCfL1RegEx THEN _      '
            modemsg = IIF$(IsCfWord, "WORD '" + fstr + "'", "CHARS '" + fstr + "'")
         IF VERIFY(fstr, ENV.CharSet + " ") <> 0 THEN             ' Contains non-display chars?
            modemsg = IIF$(IsCfWord, "WORD " + sStr2Hex(fstr), "CHARS " + sStr2Hex(fstr))
         END IF                                                   '
         scError(%eNone, modemsg + " found")                      '
      END IF                                                      '
      MExitMeth                                                   '
   END IF                                                         '

   ' Not the simple kind, lots more work
   lclCmd = PCmd                                                  '
   lclWord = UUCASE(sGetWord(lclCmd, %Strip, %QuoteNotSig))       ' Peel off RFIND
   IF me.CmdParse(lclCmd) THEN MExitMeth                          ' Do basic parsing, exit if errors
   lclWord = LSET$(UUCASE(pCmdOps(0)), 8)                         ' Get next word to an 8 char field

   IF INSTR("FIND    F       FF      ", lclWord) <> 0 THEN        ' A 'pass-thru' FIND command?
      me.FindReset                                                ' Reset the parse output fields
      IF ISTRUE me.RangeSet(lclCmd, %False, %False) THEN MExitMeth' Setup any range command operands
      IF sCRTParse("ASDMEC1LPQU") THEN MExitMeth                  ' Scan FIND (All, Subset, Direct, Modifier, Exlude, Cols, Lit1, LEFT/RIGHT, Pen, QPen, User)
      me.FindSetup                                                ' Copy CRT parse's answers
      IF ISFALSE IsCfLit1 THEN _                                  ' Lit1 present?
         MErrExit(%eFail, "Missing FIND search string")           '
      LocFind = "FIND"                                            ' Remember which LOC / FIND command was last

   ELSEIF INSTR("CHANGE  C       CHA     CHG     ", lclWord) <> 0 THEN ' A 'pass-thru' CHANGE command?
      me.FindReset                                                ' Reset the parse output fields
      IF ISTRUE me.RangeSet(lclCmd, %False, %False) THEN MExitMeth' Setup any range command operands
      IF sCRTParse("ASDMEC12LPQHU") THEN MExitMeth                ' Scan CHANGE (All, Subset, Direct, Modifier, Exclude, Cols, Lit1, Lit2, LEFT/RIGHT, Pen, CPen, Shift, User)
      me.FindSetup                                                ' Copy CRT parse's answers
      IF ISFALSE IsCfLit1 THEN _                                  ' Lit1 present?
         MErrExit(%eFail, "Missing CHANGE search string")         '
      IF ISFALSE IsCfLit2 THEN _                                  ' Lit2 present?
         MErrExit(%eFail, "Missing CHANGE change string")         '
      LocFind = "FIND"                                            ' Remember which LOC / FIND command was last

   ELSEIF INSTR("DELETE  DEL     ", lclWord) <> 0 THEN            ' A 'pass-thru' DELETE command?
      me.FindReset                                                ' Reset the parse output fields
      IF ISTRUE me.RangeSet("", %False, %False) THEN MExitMeth    ' Setup any range command operands
      IF sCRTParse("ASDMC1PU") THEN MExitMeth                     ' Scan (All, Direct, Modifier, Cols, Lit1, User)
      me.FindSetup                                                ' Copy CRT parse's answers
      IF ISFALSE IsCfLit1 THEN _                                  ' Lit1 present?
         MErrExit(%eFail, "Missing DELETE search string")         '
      LocFind = "DELETE"                                          ' Remember which LOC / FIND command was last

   ELSEIF lclWord = "SPLIT   " THEN                               ' A 'pass-thru' SPLIT command?
      me.FindReset                                                ' Reset the parse output fields
      IF ISTRUE me.RangeSet(lclCmd, %False, %False) THEN MExitMeth' Setup any range command operands
      IF sCRTParse("ASCDLM12PU") THEN MExitMeth                   ' Scan (All, Subset, Columns, Modifier, Direction, Lit1, Lit2, Clr, User)
      me.FindSetup                                                ' Copy CRT parse's answers
      CfSet(%CRTNext)                                             ' Set NEXT
      IF ISFALSE IsCfLit1 THEN _                                  ' Lit1 present?
         MErrExit(%eFail, "Missing SPLIT search string")          '
      IF ISFALSE IsCfLit2 THEN _                                  ' Lit2 present?
         MErrExit(%eFail, "Missing SPLIT change string")          '
      IF ISFALSE IsCfL2Picture OR cfSplitPt2 = 0 THEN _           ' Lit1 must be P type
         MErrExit(%eFail, "SPLIT change string must be P'..|..' type literal")
      LocFind = "SPLIT"                                           ' Remember which LOC / FIND command was last

   ELSEIF lclWord = "JOIN    " THEN                               ' A 'pass-thru' JOIN command?
      me.FindReset                                                ' Reset the parse output fields
      IF ISTRUE me.RangeSet(lclCmd, %False, %False) THEN MExitMeth' Setup any range command operands
      IF sCRTParse("ASM12PU") THEN MExitMeth                      ' Scan (All, Subset, Modifier, Lit1, Lit2, Clr, User)
      me.FindSetup                                                ' Copy CRT parse's answers
      CfSet(%CRTNext)                                             ' Set NEXT
      IF ISFALSE IsCfLit1 THEN _                                  ' Lit1 present?
         MErrExit(%eFail, "Missing JOIN search string")           '
      IF ISFALSE IsCfL1Picture AND ISFALSE IsCfL1RegEx THEN _     ' Lit1 must be P type or R type
         MErrExit(%eFail, "JOIN search string must be P or R type literal")
      IF IsCfL1Picture THEN                                       ' P type, must be with LM or RM, but not both
         IF (IsCfLM AND ISFALSE IsCfRM) OR _                      '
            (ISFALSE IsCfLM AND IsCfRM) THEN                      '
            '
         ELSE
            MErrExit(%eFail, "JOIN search string must be P'[....' or P'....]' type literal")
         END IF                                                   '
      ELSE                                                        '
         IF (IsCfLM AND ISFALSE IsCfRM) OR _                      '
            (ISFALSE IsCfLM AND IsCfRM) THEN                      '
            '
         ELSE
            MErrExit(%eFail, "JOIN search string must be R'^....' or P'....$' type literal")
         END IF                                                   '
      END IF
      IF ISFALSE IsCfLit2 THEN                                    ' Lit2 missing?
         CfSet(%CrtLit2)                                          ' Provide a default
         cfChange = "!"                                           ' of P'!'
      END IF                                                      '
      IF cfSplitPt2 THEN _                                        ' No | chars in Join
         MErrExit(%eFail, "Split character | not allowed in JOIN change string")
      LocFind = "JOIN"                                            ' Remember which LOC / FIND command was last

   ELSE                                                           '
      MErrExit(%eFail, "RFIND has no operands")                   '
   END IF                                                         '

   IF IsCfALL THEN                                                ' If ALL switch it to FIRST
      CfSet(%CRTFirst)                                            ' Set FIRST
      CfClear(%CRTALL)                                            ' Clear ALL
   END IF

   IF ISFALSE IsCfNext AND ISFALSE IsCfPrev AND _                 '
      ISFALSE IsCfFirst AND ISFALSE IsCfLast THEN _               '
      CfSet(%CRTNext)                                             ' Set NEXT if no others specified

   IF LEN(CrtL1Raw) < 16 THEN                                     ' Build msg literal
      str = CrtL1Raw                                              '
   ELSE                                                           '
      str = LEFT$(CrtL1Raw, 14) + ".."                            '
   END IF                                                         '
   modemsg = IIF$(IsCfWord, "WORD " + str, "CHARS " + str)        '
   CurrPCmd = "FIND"

   IF ISFALSE me.Search(0, lclCursMode) THEN                      ' Go do a FIND
      scError(%eNone, IIF$(IsCfNext, "Bottom of data reached", "Top of data reached"))
      cfNotFound = IIF(IsCfNext, 2, 1)                            ' Remember we failed
      MExitMeth                                                   '
   END IF                                                         '

   i = sLine                                                      ' Get found line number
   lclword = MID$(LTxtG(i), sCol, cfFLen)                         '
   IF IsCfMX THEN                                                 ' X'ing the line?
      scError(%eNone, modemsg + " found and X'd")                 '
   ELSE                                                           '
      fstr = MID$(LTxtG(cfFLine), cfFCol, cfFLen)                 ' Get the found string
      IF IsCfL1Picture OR IsCfL1Hex OR IsCfL1RegEx THEN _         '
         modemsg = IIF$(IsCfWord, "WORD '" + fstr + "'", "CHARS " + fstr + "'")
      IF VERIFY(fstr, ENV.CharSet + " ") <> 0 THEN                ' Contains non-display chars?
         modemsg = IIF$(IsCfWord, "WORD " + sStr2Hex(fstr), "CHARS " + sStr2Hex(fstr))
      END IF                                                      '
      scError(%eNone, modemsg + " found")                         '
   END IF                                                         '
   MExit
END METHOD

METHOD  pCmdRIGHT(pCmd AS STRING)
'---------- Scroll RIGHT
DIM ScrAmt AS LONG, ScrCmd AS STRING, DataIX AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN                                         ' No Ops, fudge it
      me.CmdParse("RIGHT " + Prf.Scroll)                          ' Re-parse it
   END IF
   ScrCmd = pCmdOps(1)                                            ' Pick up the operand
   IF pCmdOpsType(1) <> %OpNum AND INSTR(CHR$(%KWHALF, %KWPAGE, %KWFULL, %KWDATA, %KWMAX, %KWCSR), CHR$(pCmdOpsType(1))) = 0 THEN _
      MErrExit(%eFail, "Invalid scroll amount")                   '
   IF IsCData THEN                                                ' No, In data area?
      me.CurSetReq(%Position, CsrAbsLine, CCol - gLNPadCol + Offset, %False) ' Set cursor set attempt
   ELSEIF IsCLinN THEN                                            ' No, In Line Number area?
      me.CurSetReq(%Position, CsrAbsLine, 0, %False)              ' Set cursor set attempt
   END IF
   IF VAL(ScrCmd) <> 0 THEN                                       ' A Number?
      ScrAmt = VAL(ScrCmd)                                        ' Yes, use it
   ELSEIF pCmdOpsType(1) = %KWHALF THEN                           ' Half?
      ScrAmt = INT(gDataLen / 2)                                  ' Yes, set it
   ELSEIF pCmdOpsType(1) = %KWPAGE THEN                           ' Page?
      ScrAmt = gDataLen                                           ' Yes, set it
   ELSEIF pCmdOpsType(1) = %KWFULL THEN                           ' Full?
      ScrAmt = gDataLen                                           ' Yes, set it
   ELSEIF pCmdOpsType(1) = %KWDATA THEN                           ' Data?
      ScrAmt = gDataLen - 1                                       ' Yes, set it
   ELSEIF pCmdOpsType(1) = %KWMAX THEN                            ' Max?
      ScrAmt = 9999999                                            ' Yes, set a BIG number
   ELSEIF IsCData THEN                                            ' No, In data area?
      ScrAmt = CCol - gLNData1                                    ' Yes, move cursor to left of screen
      IF ScrAmt = 0 THEN                                          ' If already there go a default amount
         ScrAmt = gDataLen - 1
         me.CurSetReq(%Position, DataIX, CCol - gLNPadCol + Offset + ScrAmt, %False)
      END IF                                                      '
   ELSEIF IsCLinN THEN                                            ' Line number area
      DataIX = sGetIX(CRow)                                       ' Locate data line cursor was on
      ScrAmt = gDataLen - 1                                       ' Scroll width of screen
   ELSE                                                           '
      ScrAmt = gDataLen - 1                                       ' Scroll width of screen
   END IF                                                         '
   me.OffsetAdd(ScrAmt)                                           '
   ErrFlag = %eNone                                               ' Say we 're OK
   MExit
END METHOD

METHOD  pCmdRJOIN(pCmd AS STRING)
'---------- RJOIN Command
LOCAL lclCmd, lclWord, modemsg, str AS STRING, lclCursMode AS INTEGER, RevOnce, PrevLastLine AS LONG
LOCAL i, ll AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   lclCursMode = Zone                                             ' Save CursMode at start of command
   me.ResetFunc(%ResetFind)                                       ' Clear previous FIND hilighting

   IF pCmdNumOps = 0 THEN                                         ' Just a simple RJoin with no operands?
      IF LocFind <> "JOIN" THEN _                                 ' Last better be JOIN
         MErrExit(%eFail, "Last search was NOT a JOIN command")   '
      IF cfNotFound = 0 THEN                                      ' Last FIND was successful?
         PrevLastLine = LastLine                                  ' Save LastLine
         me.Join()                                                ' Do the Join
         scError(%eNone, IIF$(LastLine <> PrevLastLine, "Join successful", "Join not possible"))
         ErrFlag = %eNone                                         ' Say we 're OK
      ELSE                                                        '
         i = me.Search(0, lclCursMode)                            ' Go see if we can find it
         IF i = 0 THEN                                            ' Got it?
            scError(nMac(%eFail), IIF$(IsCfNext, "Bottom of data reached", "Top of data reached"))
            cfNotFound = IIF(IsCfNext, 2, 1)                      ' Remember we failed
         ELSE                                                     '
            PrevLastLine = LastLine                               ' Save LastLine
            me.Join()                                             ' Do the Join
            scError(%eNone, IIF$(LastLine <> PrevLastLine, "Join successful", "Join not possible"))
            ErrFlag = %eNone                                      ' Say we 're OK
         END IF                                                   '
      END IF                                                      '
   ELSE                                                           '
      scError(%eFail, "RJOIN has no parameters")                  '
   END IF                                                         '
   MExit
END METHOD

METHOD  pCmdRLOC(pCmd AS STRING)
'---------- RLOC Command
LOCAL llastline, lastcol, RevOnce AS LONG
   MEntry
   IF LocLine = 0 OR LocFlag = 0 THEN _                           ' A previous LOCATE?
      MErrExit(%eFail, "No previous LOCATE")                      ' No? Too bad
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 1 AND pCmdOpsType(1) = %KWREVERSE THEN         ' Is this the simple REVERSE?
      OnRevOnce: RevOnce = %True                                  ' Reverse just once, remember it locally
      pCmdNumOps = 0                                              ' Pretend there are no operands now
   END IF                                                         '
   IF IsCData THEN                                                ' In the Data Area
      llastline = sGetIX(CRow)                                    ' Save line the cursor is on
      lastcol = CCol - gLNPadCol + Offset + 1                     ' Calc column position in line
   END IF                                                         '
   IF IsCLinn THEN                                                ' In the Line Area
      llastline  = sGetIX(CRow): lastcol = 0                      ' Swap in current cursor location
   END IF                                                         '
   LocFind = "LOCATE"                                             ' Remember what Loc / FIND command was last
   IF ISFALSE me.LocSearch() THEN                                 ' Go do the full type search
      IF RevOnce THEN                                             ' Get the right message
         scError(nMac(%eFail), IIF$(locDir = 1, "Top of data reached", "Bottom of data reached"))
         LocNotFound = IIF(IsCfNext, 1, 2)                        '
      ELSE                                                        '
         scError(nMac(%eFail), IIF$(locDir = 1, "Bottom of data reached", "Top of data reached"))
         LocNotFound = IIF(IsCfNext, 2, 1)                        '
      END IF                                                      '
      IF llastline <> 0 THEN _                                    ' A saved cursor location?
         me.CurSetReq(%Find, llastline, lastcol, %False, %False)  ' Set cursor
   ELSE                                                           '
      IF IsMEdit AND IsLfFile THEN                                ' A MEdit LOC FILE command?
         scError(%eNone, "File " + FORMAT$(LMIXG(LocLineR)) + " of " + FORMAT$(MEditCount) + " found")
      END IF                                                      '
   END IF                                                         '
   MExit
END METHOD

METHOD  pCmdRLOCFIND(pCmd AS STRING)
'---------- RLOCFIND Command
LOCAL lCmd AS STRING
   MEntry
   lCmd = UUCASE(sGetWord(pCmd, %Strip, %QuoteNotSig))            ' Peel off RLOCFIND
   lCmd = UUCASE(sGetWord(pCmd, %NoStrip, %QuoteNotSig))          ' Get next word (if any)
   IF ISNOTNULL(TRIM$(lCmd)) THEN                                 ' Got further operands?
      lCmd = LSET$(lCmd, 8)                                       ' Make it an 8 char field
      IF INSTR("FIND    F       FF      CHANGE  C       CHA     CHG     SPLIT   JOIN    DELETE  DEL     ", lCmd) <> 0 THEN  ' A 'pass-thru' FIND/CHANGE/SPLIT/JOIN command?
         pCmdRFind("RFIND " + pCmd)                               ' Pass to RFIND
         LocFind = TRIM$(lCmd)                                    '
         MExitMeth                                                '
      END IF                                                      '
      IF INSTR("REVERSE REV     ", lCmd) <> 0 THEN                ' A simple REVERSE?
         IF LocFind = "LOCATE" THEN                               '
            pCmdRLoc("RLOC REVERSE" ): MExitMeth                  ' If last was a LOCATE, do an RLOC command
         ELSE                                                     '
            pCmdRFind("RFIND REVERSE"): MExitMeth                 ' Pass to RFIND
         END IF                                                   '
      END IF                                                      '
   END IF                                                         '
   IF LocFind = "LOCATE" THEN pCmdRLoc("RLOC"): MExitMeth         ' If last was a LOCATE, do an RLOC command
   pCmdRFind("RFIND")                                             ' Else treat as a RFIND
   MExit
END METHOD

METHOD  pCmdRNFIND(pCmd AS STRING)
'---------- RNFIND Command
LOCAL lclCmd, lclWord AS STRING, lclCursMode AS INTEGER, RevOnce AS LONG
LOCAL i AS LONG
   MEntry
   LocFind = "FIND"                                               ' Remember what Loc / FIND command was last
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   lclCursMode = Zone                                             ' Save CursMode at start of command
   IF IsRevOnce THEN RevOnce = %True                              ' Save RevOnce locally, TP. version will be reset
   IF ISFALSE me.NFSearch(1, lclCursMode) THEN                    ' Go see if we can find it
      IF RevOnce THEN                                             ' Issue correct message
         scError(nMac(%eFail), IIF$(IsCfNext, "Top of data reached", "Bottom of data reached"))
         cfNotFound = IIF(IsCfNext, 1, 2)                         ' Remember we failed
      ELSE                                                        '
         scError(nMac(%eFail), IIF$(IsCfNext, "Bottom of data reached", "Top of data reached"))
         cfNotFound = IIF(IsCfNext, 2, 1)                         ' Remember we failed
      END IF                                                      '
      MExitMeth                                                   '
   END IF                                                         '
   IF IsCfMX THEN                                                 ' X'ing the line?
      scError(%eNone, "Found line and X'd it")                    '
   ELSE                                                           '
      scError(%eNone, "Found line")                               '
   END IF                                                         '
   MExit                                                          '
END METHOD

METHOD  pCmdRSPLIT(pCmd AS STRING)
'---------- RSPLIT Command
LOCAL lclCmd, lclWord, modemsg, str AS STRING, lclCursMode AS INTEGER, RevOnce AS LONG
LOCAL i, ll AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   lclCursMode = Zone                                             ' Save CursMode at start of command
   me.ResetFunc(%ResetFind)                                       ' Clear previous FIND hilighting

   IF pCmdNumOps = 0 THEN                                         ' Just a simple RSPLIT with no operands?
      IF LocFind <> "SPLIT" THEN _                                ' Last better be SPLIT
         MErrExit(%eFail, "Last search was NOT a SPLIT command")  '
      IF cfNotFound = 0 THEN                                      ' Last FIND was successful?
         me.Split()                                               ' Yes, Just do the Split
         scError(%eNone, "Split successful")                      '
         ErrFlag = %eNone                                         ' Say we 're OK
         cfNotFound = 3                                           ' We're in not found mode again
      ELSE                                                        '
         i = me.Search(0, lclCursMode)                            ' Go see if we can find it
         IF i = 0 THEN                                            ' Got it?
            scError(nMac(%eFail), IIF$(IsCfNext, "Bottom of data reached", "Top of data reached"))
            cfNotFound = IIF(IsCfNext, 2, 1)                      ' Remember we failed
         ELSE                                                     '
            me.Split()                                            ' Do the Split
            scError(%eNone, "Split successful")                   '
            ErrFlag = %eNone                                      ' Say we 're OK
         END IF                                                   '
      END IF                                                      '
   ELSE                                                           '
      scError(%eFail, "RSPLIT has no parameters")                 '
   END IF                                                         '
   MExit
END METHOD

METHOD  pCmdRUN(pCmd AS STRING)
'---------- Run the current file (assumed to be a script)
LOCAL operands, t AS STRING
LOCAL i, j AS LONG, lclCursMode AS INTEGER
LOCAL recs, RC AS LONG
LOCAL lclDrive, lclPath AS STRING
LOCAL runIO  AS iIO                                               ' For our I/O stuff
   MEntry
   IF IsSetEdit OR _                                              ' SetEdit?
      IsClip OR _                                                 ' ClipBoard Mode?
      IsBrowse OR _                                               ' Browsing?
      IsView OR _                                                 ' Viewing?
      IsMedit OR _                                                ' MEdit?
      TIP.FilePath = $Empty OR _                                  ' (Empty) file?
      TIP.Extn = "" THEN                                          ' No extension?
      MErrExit(%eFail, "Not a suitable Edit session for RUN")     ' Say sorry
   END IF                                                         '
   LET runIO = CLASS "cIO"                                        '

   '----- Get remaining command operands
   operands = PCmd                                                ' Get a copy to play with
   t = sGetWord(operands, %Strip, %QuoteSig)                      ' Strip off the RUN command itself


   lclCursMode = %CursBad                                         ' Pretend a bad cursor location
   me.FindPushPop("PUSH")                                         ' Save FIND stuff
   me.FindSetAll                                                  ' Set parse search output fields for ALL
   me.RangeSetRange(2, LastLine - 1)                              ' Everything

   i = me.Search(0, lclCursMode)                                  ' Do the initial search
   IF i THEN                                                      ' We have the 1st RUN line

      '----- Ready the output file
      runIO.Setup("O", "", "", ENV.RunPath + TIP.Base + TIP.Extn) ' Set filename
      IF runIO.EXEC THEN _                                        ' Go OPEN the file
         MErrExit(%eFail, runIO.ResultMsg)                        ' Oops?  Bail out
      GOSUB PutOne                                                ' Go handle one
   END IF
   DO WHILE i                                                     ' Do till not found
      i = me.Search(1, lclCursMode)                               ' Do re-Finds
      IF i THEN GOSUB PutOne                                      ' Found one
   LOOP                                                           '
   runIO.Close                                                    ' Close the file
   me.FindPushPop("POP")                                          ' Restore FIND stuff


'   debug DCmd
   '----- Switch to the EXE path
   lclPath = CURDIR$                                              ' Locate where we are
   IF MID$(lclPath, 2, 1) = ":" THEN _                            ' Extract Drive if present
      lclDrive = LEFT$(lclPath, 2)                                ' and save it
   CHDIR TRIM$(runIO.Drive)                                       ' Switch to RUN path

   '----- Fire it up
   gLoopCtr = -1                                                  ' Suspend loop detection
   t = ENVIRON$("COMSPEC") + " " + ENV.RunFlags + " " + $DQ + runIO.FilePath + $DQ + " " + operands
'   ShellExecute(BYVAL %NULL, "open", $DQ + runIO.FilePath + $DQ + CHR$(0), operands + CHR$(0), BYVAL %NULL, %SW_SHOWNORMAL)
   SHELL t, 1
   RESET gLoopCtr                                                 ' Reset the counter

   '----- Restore PATH
   IF ISNOTNULL(lclDrive) THEN CHDRIVE lclDrive                   ' Switch drive if needed
   CHDIR lclPath                                                  ' put back the original path
   MExitMeth

PutOne:
   PRINT # runIO.FNum, LTxtG(cfFLine)                             ' Write line to SUBMIT file
   INCR recs                                                      ' count it
   RETURN                                                         '

END METHOD

METHOD  pCmdSAVE(pCmd AS STRING)
'---------- SAVE to OrigName
LOCAL i, j, lclTop, fcount, fline, tline, MOnly, MCount, ProfPushed, MEditFLine, Quiet AS LONG
LOCAL BKPName, fn, d, lclDir, lclFile, lclExtn, lclProfName, t, CBD AS STRING
DIM fnlist() AS STRING, FD AS DIRDATA, filesize AS QUAD
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps > 0 AND pCmdOps(1) = "MEditOnly" THEN MOnly = %True ' See if END calling for MEdit only modified
   IF pCmdNumOps > 0 AND (pCmdOps(1) = "Quiet" OR pCmdOps(2) = "Quiet") THEN Quiet = %True     ' Quiet request?
   lclTop = TopScrn                                               ' Save where we are
   IF IsBrowse OR IsView THEN MErrExit(%eFail, "SAVE disallowed, you're in BROWSE/VIEW Mode, use CREATE")

   '----- See if a named CLIP session
   IF IsClip THEN                                                 ' CLIP of a private clipboard?
      gKeyPrimOper = IIF$(ISNOTNULL(ClipName), ClipName, "")      ' Set possible clipboardname
      FOR i = 1 TO LastLine                                       ' Capture all the lines
         IF IsLData(i) AND ISFALSE IsLInsertLine(i) THEN          ' All Data but not Insert lines                                                  ' Get the located line number
            CBD += LTxtG(i) + $CRLF                               ' Build clipboard string
         END IF                                                   '
      NEXT i
      IF ISTRUE sWriteClipboard(CBD) THEN                         ' OK?
         scError(%eNone, "Saved to Clipboard: " + ClipName)       '
         OffModdFlag                                              ' Remove MOD flag
         MexitMeth                                                '
      END IF                                                      '
   END IF

   IF TIP.FilePath = $Empty THEN pCmdSAVEAS("SAVEAS"): MExitMeth  ' Pass odd-ball ones to SAVEAS                                                   '

   fcount = IIF(IsMedit, MAX(1, MEditCount), 1)                   ' Set number of files to save (possibly)
   REDIM fnlist(fcount) AS STRING                                 '

   IF ISFALSE IsMEdit THEN                                        ' Build the list
      fnlist(1) = TIP.FilePath                                    ' Build full name
      IF ISNOTNULL(VSaveFile) THEN sRecycleBin(VSaveFile, "D"): VSaveFile = "" ' Clear VSAVE if present
   ELSE                                                           '
      FOR i = 1 TO fcount                                         '
         fnlist(i) = me.MEditListGet(i)                           ' Build full name from MEdit list
      NEXT i                                                      '
   END IF                                                         '

   '----- Do AUTONUM if needed
   IF Prf.AUTONUM THEN                                            ' AUTONUM requested?
      pCmd = "RENUM QUIET"                                        ' Do renum in quiet mode
      CALL pCmdRenumber(pCmd)                                     ' go do the renum
   END IF

   '----- Now do the SAVE
   FOR j = 1 TO fcount                                            ' OK, now loop doing the SAVE for each one
      IF IsMEdit AND MOnly AND ISFALSE me.MEditFlagGet(j) THEN ITERATE FOR ' If modified only MEdit, then ignore this one
      i = IIF(IsMEdit, me.FileWatch(fnlist(j), %WatchEnd), me.FileWatch("", %WatchEnd)) ' Stop watching
      lclDir = PATHNAME$(PATH, fnlist(j))                         ' Get PATH, up to and incl the \
      lclFile = PATHNAME$(NAME, fnlist(j))                        ' Get File, after \ and before .
      lclExtn = PATHNAME$(EXTN, fnlist(j))                        ' Get Extn, . and all after it

      '----- Create the BKP file if needed
      IF Prf.AutoBkup AND ISFALSE IsBackupDone THEN               ' Need to create BKP file?
         BKPName = lclDir + lclFile + ".BKP" + lclExtn            ' Insert .BKP before the current Extn
         TRY                                                      ' Try the delete of previous
            sRecycleBin(BKPName, "D")                             ' Del any previous BKP file
         CATCH                                                    ' Do nothing, just ignore it
            '                                                     '
         END TRY                                                  '
         TRY                                                      ' Now try the rename
            NAME fnlist(j) AS BKPName                             ' Rename current file to BKP
         CATCH                                                    ' See what happened
            IF ERR = 70 THEN _                                    '
               scError(%eFail, "Can't save, file is in use")      '
            IF ERR <> 0 THEN _                                    ' Already exists?
               scError(%eFail, "Can't create BKP file")           '
         END TRY                                                  '
         IF ERR > 0 THEN MExitMeth                                '
      END IF                                                      '

      '----- Set criteria for the line range and write the file
      me.FindPushPop("PUSH")                                      ' Save FIND stuff
      me.FindSetAll                                               ' Set parse search output fields for ALL
      IF ISFALSE IsMEdit THEN                                     ' Set the correct line range
         me.RangeSetRange(2, LastLine - 1)                        ' Everything for non=MEdit mode
      ELSE                                                        '
         RESET fline, tline                                       ' Zero our answer areas
         FOR i = 1 TO LastLine                                    ' Must search for range of this MEdit file
            IF IsLFile(i) THEN                                    ' Look for =FILE> lines
               IF fnlist(j) = TRIM$(LTxtG(i)) THEN                ' Found start of file
                  MEditFLine = i                                  ' Save the FILE line pointer
                  fline = i + 1                                   ' Save start line number
                  INCR i                                          '
                  DO WHILE ISFALSE IsLFile(i) AND ISFALSE IsLBottom(i)
                     tline = i                                    ' Save as hiwater last line
                     INCR i                                       '
                  LOOP                                            '
                  me.RangeSetRange(fline, tline)                  ' Set fline/tline as the range
                  EXIT FOR                                        ' Done this looping crap
               END IF                                             '
            END IF                                                '
         NEXT i                                                   '
      END IF                                                      '

      '----- If MEdit see if different Profile needed
      IF IsMEdit AND TIP.Profile <> sParseProfile(fnlist(j)) THEN ' Different Profile needed
         Prf.SetProfName(sParseProfile(fnlist(j)), %False)        ' Swap to the oddball one
         Prf.ReadAll(%True)                                       ' Go read the other INI values
      END IF                                                      '

      '----- Range is set, now go write the file
      TOP.Setup("OR", "", "", fnlist(j))                          ' Setup TOP

      IF me.WriteFile(Quiet) THEN                                 ' Let WriteFile have a go
         INCR MCount                                              ' Count saved files
         OffModdFlag: OnUndoFlag                                  ' Turn off the Modd flag
         IF IsMEdit THEN                                          ' Do the reset
            me.MeditFlagSet(j, %False)                            ' Turn off the Modd flag of THIS MEdit file
            me.UpdLControl(MEditFLine)                            ' Update the FILE line
         END IF                                                   '
         IF ISFALSE IsMEdit THEN sRecentAdd(lclDir + lclFile + lclExtn) ' Go add to RECENT list
         t = sAutoMask(TOP.FilePath)                              ' Get AUTOFAV name if available
         IF t <> "" THEN                                          ' Got a name
            sFileListAdd(t, TOP.FilePath)                         ' Add to the list
         END IF                                                   '
      ELSE                                                        '
         scError(%eFail, "File " + lclFile + " not written")      ' Tell User
      END IF                                                      '
      me.FindPushPop("POP")                                       ' Restore FIND stuff

      IF CurrPcmd <> "END" THEN                                   '
         sFileQueue("A", " ", fnlist(j))                          ' Add back to FQ, FileWatch/Watchend removed it
         IF me.FileWatch(fnlist(j), %WatchStart) THEN             ' Re- Establish the watch
            scError(0, "File Watch could not be re-established")
         END IF                                                   '
      END IF                                                      '

      '----- If MEdit see if Profile re-load needed
      IF IsMEdit AND TIP.Profile <> sParseProfile(fnlist(j)) THEN ' Different Profile needed
         Prf.SetProfName(TIP.Profile, %False)                     ' Swap back to the primary
         Prf.ReadAll(%True)                                       ' Go read the other INI values
      END IF                                                      '
   NEXT j                                                         '
   OnBackupDone                                                   ' Remember we did this
   ActionCtr = Prf.ActionSave                                     ' Reinit the action counter

   IF IsMEdit AND MCount = 0 THEN                                 '
      scError(0, "No modified files required saving")             '
   ELSE                                                           '
      IF ISFALSE Quiet THEN                                       ' Issue messages, maybe.
         IF IsMEdit THEN                                          '
            scError(0, FORMAT$(mcount) + " file" + IIF$(mcount = 1, "", "s") + " saved")
         ELSE                                                     '
            scError(0, "File saved")                              '
            sDoStatusBar($SBState)                                ' re-Do the StatusBar boxes
         END IF                                                   '
      END IF                                                      '
      TIP.TimeDateRefresh                                         ' Get refreshed Date/Time
   END IF                                                         '

   TopScrn = lclTop                                               ' Keep our position
   MExitMeth
END METHOD

METHOD  pCmdSAVEALL(pCmd AS STRING)
'---------- Save all modified tabs
LOCAL Cond AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 1 THEN                                         ' Look for COND
      IF IsEQ(pCmdOps(1), "COND") THEN                            ' Flag conditional if COND coded
         Cond = %True                                             '
      ELSE                                                        '
         MErrExit(%eFail, "Unknown SAVEALL parameter")            '
      END IF                                                      '
   END IF                                                         '
   sGblSaveAll(Cond)                                              ' Go do it from global routine
   MExit
END METHOD

METHOD  pCmdSAVEAS(pCmd AS STRING)
'---------- SAVEAS a new filename
LOCAL lclCmd, lclWord, fn, t AS STRING, x, i, lclTop AS LONG
   MEntry
   gLoopCtr = -1
   IF IsSetEdit THEN                                              ' SetEdit?
      lclcmd = pCmd                                               ' Get the command line
      lclWord = sGetWord(lclCmd, %Strip, %QuoteSig)               ' Strip the 1st word
      pCmdCREATE("CREATE " + lclcmd + " .ZF .ZL")                 ' Pass to CREATE
      MErrExit(%eNone, "SAVEAS was converted to a CREATE command")'
   END IF
   IF IsMEdit THEN pCmdSaveAsMedit(pCmd): MExitMeth               ' Pass off MEdit version
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   lclTop = TopScrn                                               ' Save where we are

   IF pCmdNumOps <> 0 THEN                                        ' Got a filename
      IF ISNOTNULL(PATHSCAN$(FULL, pCmdOps(1))) THEN _            '
         MErrExit(%eFail, "File exists, use REPLACE to re-use it")'
   END IF

   TOP.Setup("AECP", "", "Specify SaveAs filename", IIF$(pCmdNumOps = 0, "", pCmdOps(1))) ' Setup for the output
   IF TOP.EXEC THEN                                               ' Go validate
      scError(%eFail, TOP.ResultMsg)                              ' Oops?  Bail out
      SetCmd                                                      ' Cursor to Cmd line
      sPrint ("Command > " + LEFT$(pCommand,pCmdLen), $$TxtLo, 1, 1) '
      IF TOP.RSTabFound THEN                                      ' Found elsewhere already open?
         gTabSwitch = TOP.TabFound                                ' Set found tab number to switch to
         gTabSwitchMsg = TOP.ResultMsg                            ' Set cross-tab message
         MExitMeth                                                '
      ELSE                                                        '
         MErrExit(%eNone, "File selection cancelled")             ' No selection?  Bail out
      END IF                                                      '
   END IF                                                         '

   me.FindPushPop("PUSH")                                         ' Save FIND stuff
   me.FindSetAll                                                  ' Set parse output fields for ALL
   me.RangeSet("RESET", %False, %False)                           ' Ensure default range of ALL before writing
   TOP.Setup("OREC", "*", "*", "*")                               ' Set for Output now
   IF ISFALSE me.WriteFile(%False) THEN _                         '
      MErrExit(%eFail, "File creation failed")                    '
   me.FindPushPop("POP")                                          ' Restore FIND stuff

   TIP.Setup("E", TOP.Profile, "", TOP.FilePath)                  ' Setup TIP now to the file
   TIP.EXEC                                                       ' Fetch the FD data
   TMode = (TMode AND (&HFFFFFFFF - %MClip))                      ' Remove any possible CLIP status

   IF ISFALSE gMacroMode THEN                                     ' If not macro mode
      me.WindowTitle                                              ' Alter window/Tab titles
   END IF                                                         '
   Prf.SetProfName(TIP.Profile, %False)                           ' Setup profile
   Prf.ReadAll(%True)                                             ' Go read this file type's INI values
   me.ClrLoad()                                                   ' Get Attr loaded if it exists
   OffModdFlag: OnUndoFlag                                        ' Turn off the Modd flag
   i = me.FileWatch("", %WatchEnd)                                ' Kill any prior Watch
   sFileQueue("A", " ", TIP.FilePath)                             ' Add to FQ
   IF me.FileWatch(TIP.FilePath, %WatchStart) THEN _              ' Establish the watch
      scError(0, "File watch could not be established")           '
   t = sAutoMask(TIP.FilePath)                                    ' Get AUTOFAV name if available
   IF ISFALSE IsMEdit AND t <> "" THEN                            ' Not Medit sessions
      sFileListAdd(t, TIP.FilePath)                               ' Add to the list
   END IF                                                         '
   sFileListAdd("Recent Files", TIP.FilePath)                     ' Add to Recent
   DoSet(%LoadReq)                                                ' Request refresh
   scError(0, "SaveAs complete")                                  '
   TopScrn = lclTop                                               ' Re-Do position
   MExit
END METHOD

METHOD  pCmdSAVEASMEdit(pCmd AS STRING)
'---------- SAVEAS a MEdit session
LOCAL lclCmd, lclWord, fn, dn, lclDir, lclFile, lclExtn, base1, base2 AS STRING
LOCAL x, i, j, lclTop, fcount, fline, tline AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   lclTop = TopScrn                                               ' Save where we are
   IF pCmdNumOps > 0 THEN _
      MErrExit(%eFail, "SAVEAS for MEDIT sessions takes no operands") ' Say sorry, no operands

   '----- Build table of filenames to be saved
   fcount = MEditCount                                            ' Set number of files to save
   REDIM fnlist(fcount) AS STRING                                 '

   '----- Add names making sure no duplicates
   FOR i = 1 TO fcount                                            ' Build table of filenames
      base1 = me.MEditListGet(i)                                  ' Copy the new name
      IF INSTR(base1, "\") <> 0 THEN base1 = MID$(base1, INSTR(-1, base1, "\") + 1) ' Reduce to base filename
      base1 = UUCASE(base1)                                       ' Uppercase it
      FOR j = 1 TO fcount                                         ' Ensure no base name duplicates
         base2 = fnlist(j)                                        ' Copy one of the previous names
         IF INSTR(base2, "\") <> 0 THEN base2 = MID$(base2, INSTR(-1, base2, "\") + 1) ' Reduce to base filename
         IF ISNOTNULL(base1) AND IsEQ(base1, base2) THEN _        ' Duplicate?
            MErrExit(%eFail, "MEDIT session contains duplicate base file names, SAVEAS cancelled")
      NEXT j                                                      '
      fnlist(i) = me.MEditListGet(i)                              '
   NEXT i                                                         '

   '----- Get the directory where things go
   sPopReady                                                      ' Ready for pop-up
   DISPLAY BROWSE hWnd, , , "Specify directory for SaveAs", sGetDefDir, %BIF_EDITBOX OR %BIF_RETURNONLYFSDIRS _
           TO lclDir                                              '
   sPopReset                                                      ' Reset popup state
   IF ISNULL(lclDir) THEN _                                       '
      MErrExit(%eNone, "Directory selection cancelled")           ' No selection?    Bail out
   lclDir = TRIM$(lclDir) + "\"                                   ' Clean the Dir name, add \

   '----- OK, loop for each file, see if already exists
   FOR j = 1 TO fcount                                            ' OK, now loop doing the SAVE for each one
      lclFile = PATHNAME$(NAME, fnlist(j))                        ' Get File, after \ and before .
      lclExtn = PATHNAME$(EXTN, fnlist(j))                        ' Get Extn, . and all after it
      IF ISNOTNULL(DIR$(lclDir + lclFile + lclExtn)) THEN _       ' See if file exists
         MErrExit(%eFail, "File: " + lclDir + lclFile + lclExtn + " already exists, SaveAs terminated")
      IF VAL(sFileQueue("S", " ", lclDir + lclFile + lclExtn))  > 0 THEN _ ' Open in some tab?
         MErrExit(%eFail, "File: " + lclDir + lclFile + lclExtn + " is open in anoher tab, SaveAs terminated")
   NEXT j                                                         '

   '----- Now loop and actually do something
   me.FindPushPop("PUSH")                                         ' Save FIND stuff
   FOR j = 1 TO fcount                                            ' OK, now loop doing the SAVE for each one
      me.UnWatchQueue(fnlist(j))                                  ' Kill Watch and deque
      lclFile = PATHNAME$(NAME, fnlist(j))                        ' Get File, after \ and before .
      lclExtn = PATHNAME$(EXTN, fnlist(j))                        ' Get Extn, . and all after it
      me.FindSetAll                                               ' Set parse search output fields for ALL
      RESET fline, tline                                          ' Zero our answer areas
      FOR i = 1 TO LastLine                                       ' Must search for range of this MEdit file
         IF IsLFile(i) THEN                                       ' Look for =FILE> lines
            IF fnlist(j) = TRIM$(LTxtG(i)) THEN                   ' Found start of file
               me.LTxtSet(i, lclDir + lclFile + lclExtn)          ' Swap name in =FILE> line
               fline = i + 1                                      ' Save start line number
               INCR i                                             '
               DO WHILE ISFALSE IsLFile(i) AND ISFALSE IsLBottom(i)
                  tline = i                                       ' Save as hiwater last line
                  INCR i                                          '
               LOOP                                               '
               me.RangeSetRange(fline, tline)                     ' Set fline/tline as the range
               EXIT FOR                                           ' Done this looping crap
            END IF                                                '
         END IF                                                   '
      NEXT i                                                      '

      '----- Do the actual write
      TOP.Setup("OCE", "", "", lclDir + lclFile + lclExtn)        ' Setup for output
      IF ISFALSE(me.WriteFile(%False)) THEN _                     ' OK, Write it
         MErrExit(%eFail, "File creation failed")                 '

      lclWord = LEFT$(lclDir, LEN(lclDir) - 1)                    ' Strip off \
      sFileQueue("A", " ", TOP.FilePath)                          ' Add to inuse queue
      IF me.FileWatch(TOP.FilePath, %WatchStart) THEN             '
         scError(0, "File Watch could not be established")        '
      END IF                                                      '

      me.MEditListSet(j, lclDir + lclFile + lclExtn)              ' Modify the MEdit list
      me.MEditFlagSet(j, %False)                                  ' Modify the MEdit list
   NEXT j                                                         '
   me.FindPushPop("POP")                                          ' Restore FIND stuff

   scError(0, "SaveAs complete")                                  '
   TopScrn = lclTop                                               ' Re-Do position
   MExit
END METHOD

METHOD  pCmdSCROLL(pCmd AS STRING)
'---------- Scroll Down (Called internally via ISCROLL command)
LOCAL VisTop, VisBottom, lclScroll, lctop, lcrow, lccol, lcoff, i AS LONG
LOCAL ScrCmd, lclCmd, lclType AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   lclType = pCmdOps(3)                                           ' Get type M=Mouse K=Keyboard
   IF INSTR(CHR$(%KWUP, %KWDOWN, %KWLEFT, %KWRIGHT), CHR$(pCmdOpsType(1))) = 0 THEN _
      MErrExit(%eFail, "Invalid scroll amount")                   '
   IF ENV.AutoScroll = 0 AND lclType = "M" THEN ErrFlag = %eNone: MExitMeth ' Exit if no mouse scrolling
   IF lclType = "K" AND ENV.AutoScroll = 0 THEN                   ' Keyboard but no scroll amount?
      lclScroll = VAL(pCmdOps(2))                                 ' At least treat as if it was one (1)
   ELSE                                                           '
      lclScroll = VAL(pCmdOps(2)) * ENV.AutoScroll                ' Apply Factor
   END IF                                                         '
   lcrow = crow: lccol = ccol                                     ' Get local copies
   IF ISFALSE IsFMTab THEN                                        ' Not the FM screen
      IF ISFALSE (IsCBad OR IsCCmnd OR IsCScrl) THEN              ' In Data or Line area?
         lcrow = sGetIX(CRow)                                     ' Get data line number
         IF lcrow < 0 THEN lcrow = sGetIX(lCRow - ABS(lcrow))     ' Convert to real line index
         IF lCCol > gLNPadCol THEN lccol = Offset + lCCol - gLNPadCol' Get column number, 0 = linenumber area
         IF lcRow = 0 THEN lccol = 0                              ' If no row, then no column
      END IF                                                      '
      IF lcrow > 0 THEN me.CurSetReq(%Low, lcrow, lccol, %False)  ' Set default cursor set attempt to where we are

      ErrFlag = %eNone                                            ' Say we 're OK
      SELECT CASE pCmdOpsType(1)                                  ' OK, see what direction to go
         CASE %KWDOWN                                             ' Scroll DOWN
            Prf.ScrlPageSus = %True                               ' Set suspend PAGE mode
            lctop = TopScrn                                       ' Save current top
            TopScrn = me.CRPFwd(%mVisible, TopScrn, lclScroll)    ' Move TopScrn by scroll amount
            IF lctop = TopScrn OR lcrow = 0 THEN MExitMeth        ' Didn't move, just exit
            IF lclType = "K"  AND lcrow <> 0 THEN lcrow += (TopScrn - lctop) ' Keep cursor from moving
            me.CurSetReq(%Position, lcrow, lccol, %False)         ' Set cursor set attempt

         CASE %KWUP                                               ' Scroll Up
            Prf.ScrlPageSus = %True                               ' Set suspend PAGE mode
            lctop = TopScrn                                       ' Save current top
            TopScrn = me.CRPBack(%mVisible, TopScrn, lclScroll)   ' Move TopScrn by scroll amount
            IF lctop = TopScrn OR lcrow = 0 THEN MExitMeth        ' Didn't move, just exit
            IF lclType = "K"  AND lcrow <> 0 THEN lcrow -= (lctop - TopScrn) ' Keep cursor from moving
            me.CurSetReq(%Position, lcrow, lccol, %False)         ' Set cursor set attempt

         CASE %KWRIGHT                                            ' Scroll Right
            lcoff = OffSet                                        ' Get current offset
            me.OffsetAdd(lclScroll)                               ' Move OffSet over a bit
            IF lclType = "K" THEN lccol += lclScroll              '
            IF lcrow = 0 THEN MExitMeth                           '
            IF lccol - 1 >= Offset AND lccol <= Offset + gDataLen THEN ' Can we keep Cursor on the line?
               me.CurSetReq(%Position, lcrow, lccol, %False)      ' Set cursor set attempt
            ELSE                                                  ' Gone off screen, keep at the left
               me.CurSetReq(%Position, lcrow, Offset + 1, %False) ' Set cursor set attempt
            END IF                                                '

         CASE %KWLEFT                                             ' Scroll Left
            lcoff = OffSet                                        ' Get current offset
            me.OffsetSub(lclScroll)                               ' Move OffSet over a bit
            IF Offset < 0 THEN Offset = 0                         ' Don't go too far
            IF lclType = "K" AND lcoff <> OffSet THEN lccol -= (lcoff - OffSet) ' Keep cursor from moving
            IF lcrow = 0 THEN MExitMeth                           '
            IF lccol >= Offset AND lccol <= Offset + gDataLen THEN' Can we keep Cursor on the line?
               me.CurSetReq(%Position, lcrow, lccol, %False)      ' Set cursor set attempt
            ELSE                                                  ' Gone off screen, keep at the top
               me.CurSetReq(%Position, lcrow, Offset + gDataLen, %False) ' Set cursor set attempt
            END IF                                                '
      END SELECT                                                  '

   ELSE                                                           ' Else it's FM mode

      ErrFlag = %eNone                                            ' Say we 're OK
      SELECT CASE pCmdOpsType(1)                                  ' OK, see what direction to go
         CASE %KWDOWN                                             ' Scroll DOWN
            TopScrn = MIN(TopScrn + lclScroll, MAX(1, AFCount - ENV.ScrHeight + (3 * ENV.FMHelpFlag) + 7)) ' Move TopScrn by scroll amount
            LastTop = 0                                           ' Make sure it moves

         CASE %KWUP                                               ' Scroll Up
            TopScrn = MAX(1, TopScrn - lclScroll)                 ' Move TopScrn by scroll amount
            LastTop = 0                                           ' Make sure it moves

      END SELECT                                                  '
   END IF                                                         '
   MExit
END METHOD

METHOD  pCmdSET(pCmd AS STRING)
'---------- Handle SET Command
LOCAL i, j, k, ctab AS LONG, MSG AS STRING
LOCAL SetVar() AS STRING, SetVarCtr AS LONG, SetKey AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   GOSUB ChkActive                                                ' Make sure SetEdit is not active
   SELECT CASE CONST pCmdNumOps                                   ' Split based on # of operands
      CASE 0                                                      ' No Ops, display all
         sWinclip_get(gSetClipB)                                  ' Go get whatever's there

         IF gSetCount > 0 THEN                                    ' Add them
            FOR i = 1 TO gSetCount                                '

               '----- Each SET item could be a Stack
               k = PARSECOUNT(gSetData(i), BINARY)                ' Get count of number in stack
               REDIM SetVar(1 TO k) AS STRING                     ' Dim variable table
               PARSE gSetData(i), SetVar(), BINARY                ' Extract the table

               '----- 1st item geta Key=value; remainder get just =value style
               FOR j = 1 TO k                                     ' For eack stack item
                  IF j = 1 THEN                                   ' Format primary entry
                     MSG += gSetKey(i) + "=" + SetVar(1) + $CRLF  ' Add to CBD
                  ELSE                                            '
                     MSG += "=" + SetVar(j) + $CRLF               ' Add pushed entry
                  END IF                                          '
               NEXT j                                             '
            NEXT i                                                '
         END IF                                                   '
         IF gSetCount = 0 THEN MSG += "None."                     '

         sWinclip_set(MSG)                                        ' Send it to the Clipboard

         ENV.PMode = %MSetEdit OR %MClipLoad                      ' Set to SetEdit mode + ClipLoad
         ENV.SETETab = PgNumber                                   ' Save current tab number
         me.LoadTheText("", "DEFAULT")                            ' Go load it in

      CASE 1                                                      ' Single Op, display it's value
         MSG = sSetTable("GET", pCmdOps(1))                       ' Go fetch it
         IF VAL(LEFT$(MSG, 1)) = 0 THEN                           '
            MSG = "Set variable " + pCmdOps(1) + " = " + MID$(MSG, 2) ' Add some words
            scError(%eNone, MSG)                                  ' Issue message
         ELSE                                                     '
            scError(VAL(LEFT$(MSG, 1)), MID$(MSG, 2))             ' Issue returned message
         END IF                                                   '

      CASE 2                                                      ' Should only be SET xxx OFF
         IF IsEQ(pCmdOps(2), "POP") THEN                          ' POP command?
            MSG = sSetTable("POP", pCmdOps(1))                    ' Go pop it
            IF VAL(LEFT$(MSG, 1)) = 0 THEN                        '
               MSG = "Popped, variable " + pCmdOps(1) + " is now = " + MID$(MSG, 2) ' Add some words
               scError(%eNone, MSG)                               ' Issue message
            ELSE                                                  '
               scError(VAL(LEFT$(MSG, 1)), MID$(MSG, 2))          ' Issue returned message
            END IF                                                '

         ELSEIF IsEQ(pCmdOps(2), "PUSH") THEN                     ' PUSH with no operands?
            MSG = sSetTable("PUSH", pCmdOps(1))                   ' Go push
            IF VAL(LEFT$(MSG, 1)) = 0 THEN                        '
               MSG = "Push duplicated top value for " + pCmdOps(1) + ", value = " + MID$(MSG, 2) ' Add some words
               scError(%eNone, MSG)                               ' Issue message
            ELSE                                                  '
               scError(VAL(LEFT$(MSG, 1)), MID$(MSG, 2))          ' Issue returned message
            END IF                                                '

         ELSEIF pCmdOpsType(2) = %KWOFF THEN                      ' OFF?
            MSG = sSetTable("DEL", pCmdOps(1))                    ' Go delete it
            scError(VAL(LEFT$(MSG, 1)), MID$(MSG, 2))             ' Issue returned message
         ELSE                                                     '
            MErrExit(%eFail, "Unknown 2nd operand for SET")       '
         END IF                                                   '

      CASE 3                                                      ' Maybe SET xxx = yyyy
         IF pCmdOps(2) = "=" THEN                                 ' =?
            IF IsEQ(pCmdOps(1), "X") THEN MErrExit(%eFail, "Restricted SET variable")
            MSG = sSetTable("SET", pCmdOps(1) + " " + pCmdOps(3)) ' Go set it
            scError(VAL(LEFT$(MSG, 1)), MID$(MSG, 2))             ' Issue returned message

         ELSE
            MErrExit(%eFail, "Unknown 2nd operand for SET")       '
         END IF                                                   '

      CASE 4                                                      ' Better be SET xxx PUSH = value
         IF IsNE(pCmdOps(2), "PUSH") OR _                         ' Validate it
            pCmdOps(3) <> "=" THEN _                              '
            MErrExit(%eFail, "Unknown SET operands")              '
            MSG = sSetTable("PUSH", pCmdOps(1) + " " + pCmdOps(4))' Go push it
            IF VAL(LEFT$(MSG, 1)) = 0 THEN                        '
               MSG = "Push complete for " + pCmdOps(1) + ", value = " + MID$(MSG, 2) ' Add some words
               scError(%eNone, MSG)                               ' Issue message
            ELSE                                                  '
               scError(VAL(LEFT$(MSG, 1)), MID$(MSG, 2))          ' Issue returned message
            END IF                                                '

      CASE ELSE                                                   ' ?? Wonder what's coded
         scError(%eFail, "Unknown SET operands")                  '
   END SELECT                                                     '
   MExitMeth

ChkActive:
   ctab = PgNumber                                                ' Save where we are
   FOR i = 1 TO TabsNum                                           ' Do for each tab
      TP = Tabs(i)                                                ' Pick the Tab
      IF IsSetEdit THEN                                           ' We have a SetEdit tab active
         TP = Tabs(ctab)                                          ' Replace TP
         MErrExit(%eFail, "SET Edit is currently active")         '
      END IF                                                      '
   NEXT i                                                         '
   TP = Tabs(ctab)                                                ' Replace TP
   RETURN
END METHOD

METHOD  pCmdSETUNDO(pCmd AS STRING)
'---------- Handle SETUNDO Command
LOCAL i, j AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN                                         ' No operands, just display it
      IF Prf.UndoNumber = 0 THEN                                  '
         MErrExit(%eNone, "SETUNDO set to 0 (OFF)")               '
      ELSE                                                        '
         MErrExit(%eNone, "SETUNDO set to " + FORMAT$(Prf.UndoNumber))
      END IF                                                      '
   END IF                                                         '

   IF pCmdOpsType(1) = %OpNum THEN                                ' Numeric?
      IF VAL(pCmdOps(1)) = 0 THEN                                 ' If 0 (OFF)0
         IF Prf.UndoNumber > 0 THEN                               ' If UNDO is currently ON
            FOR i = 1 TO Prf.UndoNumber                           ' Delete temporary files
               KILL me.UndoUFnGet(i)                              '
               KILL me.UndoTFnGet(i)                              '
               KILL me.UndoIXFnGet(i)                             '
            NEXT i                                                '
         END IF                                                   '
         Prf.UndoNumber = 0                                       ' Set to zero
         scError(%eNone, "UNDO is now disabled")                  '

      ELSE                                                        '
         IF VAL(pCmdOps(1)) > 25 THEN _                           ' Reasonable?
            MErrExit(%eFail, "A maximum of 25 UNDO levels is supported")
         Prf.UndoNumber = VAL(pCmdOps(1))                         ' then set it
         me.UndoInit()                                            '
         scError(%eNone, "UNDO activated, levels = " + FORMAT$(Prf.UndoNumber))
      END IF                                                      '

   ELSE                                                           '
      scError(%eFail, "Unknown SETUNDO operand")                  '
   END IF                                                         '
   MExit
END METHOD

METHOD  pCmdSORT(pCmd AS STRING)
'---------- SORT the data
LOCAL ans, lclCmd AS STRING, lclTop AS LONG
   MEntry
   lcltop = TopScrn                                               ' Save where we are
   me.FindSave                                                    ' Save RFIND stuff
   IF me.CmdParse(pCmd) THEN me.FindLoad: MExitMeth               ' Do basic parsing, exit if errors
   IF ISTRUE me.RangeSet("", %False, %False) THEN me.FindLoad: MExitMeth ' Setup any range command operands
   me.FindReset                                                   ' Reset the parse output fields
   IF sCRTParse("") THEN me.FindLoad: MExitMeth                   ' Reset Scan
   me.FindSetup                                                   ' Copy CRT parse's answers
   ans = me.LTxtSort()                                            ' Go do the sort
   TopScrn = lclTop                                               ' Go to where we were
   scError(VAL(LEFT$(ans, 1)), MID$(ans, 2))                      ' Issue returned message
   me.FindLoad                                                    ' Reload RFIND stuff
   MExit
END METHOD

METHOD  pCmdSOURCE(pCmd AS STRING)
'---------- Set default Encoding for current profile
LOCAL Wrd1, lclSource, NewSource, A2S, S2A AS STRING, i AS LONG
   MEntry
   lclSource = Prf.PSource                                        ' Save curent SOURCE
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN _                                       ' No Ops?
      MErrExit(%eNone, "SOURCE codepage set to " + Prf.PSource + ", COLLATE codepage set to " + Prf.PCollate)
   NewSource = UUCASE(pCmdOps(1))                                 ' Get an uppercase name
   IF NewSource <> "ANSI" AND _                                   ' Is SOURCE non-ANSI mode?
      NewSource <> "UTF8" AND _                                   '
      NewSource <> "UTF16" AND _                                  '
      NewSource <> "UTF16LE" AND _                                '
      NewSource <> "UTF16BE" THEN                                 '
      IF ISFALSE ISFILE(ENV.INIPath + NewSource + ".SOURCE") THEN _ ' See if the Custom SOURCE file exists
         MErrExit(%eFail, "SOURCE codepage file: " + ENV.INIPath + pCmdOps(1) + ".SOURCE is not found, command rejected")
   END IF                                                         '

   Prf.PSource = NewSource                                        ' Set as SOURCE and
   Prf.PCollate = NewSource                                       ' COLLATE
   IF IsNE(Prf.PSource, lclSource) AND LastReal > 0 THEN          ' Warn the user if live data?
      sDoMsgBox ("SOURCE codepage is changing from |K" + lclSource + "|B to |K" + Prf.PSource + "|B," + $CRLF + _
                 "You should SAVE the current data immediately using this new format" + $CRLF + _
                 "to prevent incorrect loading the next time the file is processed" + $CRLF + $CRLF + _
                 "You may also want to examine your WORD character string setting and" + $CRLF + _
                 "adjust this as well if needed.", %MB_OK OR %MB_USERICON, "Source Warning")
   END IF
   sDoStatusBar($SBSource)                                        ' re-Do the StatusBar Source box
   scError(%eNone, "SOURCE codepage set to " + Prf.PSource + ", COLLATE codepage set to " + Prf.PCollate)
   MExit
END METHOD

METHOD  pCmdSPLIT(pCmd AS STRING)
'---------- SPLIT Command
LOCAL ol, nl, oc, nc AS STRING
LOCAL i AS LONG, AllCount, sp AS LONG, lclCursMode AS INTEGER
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   lclCursMode = Zone                                             ' Save CursMode at start of command
   me.ResetFunc(%ResetFind)                                       ' Clear previous FIND hilighting

   IF ISTRUE me.RangeSet("", %False, %False) THEN MExitMeth       ' Setup any range command operands
   me.FindReset                                                   ' Reset the parse output fields

   IF sCRTParse("ASCDLM12PU") THEN MExitMeth                      ' Scan (All, Subset, Columns, Direction, Left/Right, Modifier, Lit1, Lit2, Clr, User)
   me.FindSetup                                                   ' Copy CRT parse's answers
   CfSet(%CRTNext)                                                ' Set NEXT

   IF ISFALSE IsCfLit1 THEN _                                     ' Lit1 present?
      MErrExit(%eFail, "Missing SPLIT search string")             '
   IF ISFALSE IsCfLit2 THEN _                                     ' Lit2 present?
      MErrExit(%eFail, "Missing SPLIT change string")             '
   IF (ISFALSE IsCfL2Format AND ISFALSE IsCfL2Picture) OR _       ' Lit2 must be F or P type
      cfSplitPt2 = 0 THEN _                                       ' Lit2 must have a | char
      MErrExit(%eFail, "SPLIT change string must be P or F'..|..' type literal")
   LocFind = "SPLIT"                                              ' Remember which LOC / FIND command was last

   IF IsCfAll THEN                                                ' The ALL version?
      RESET AllCount                                              ' Reset variables
      i = me.Search(0, lclCursMode)                               ' Do the initial search
      lclCursMode = %CursBad                                      ' Make it look like Bad Cursor
      IF i THEN                                                   '
         me.Split()                                               ' Do the Split
         INCR AllCount                                            ' Count
      END IF                                                      '
      DO WHILE i                                                  ' Do till not found
         i = me.Search(1, lclCursMode)                            ' Do re-Finds
         IF i THEN                                                ' Found one
            me.Split()                                            ' Do the Split
            INCR AllCount                                         ' Count successful
         END IF                                                   '
      LOOP                                                        '
      IF AllCount = 0 THEN                                        ' Issue approp. message
         scError(nMac(%eFail), "Split point not found")                 '
      ELSE                                                        '
         scError(0, "Split performed " + FORMAT$(Allcount) + " times") '
      END IF                                                      '

   ELSE                                                           ' The non-All version
      i = me.Search(0, lclCursMode)                               ' Go see if we can find it
      IF i = 0 THEN                                               ' Got it?
         scError(nMac(%eFail), IIF$(IsCfNext, "Bottom of data reached", "Top of data reached"))
         cfNotFound = IIF(IsCfNext, 2, 1)                         ' Remember we failed
      ELSE                                                        '
         me.Split()                                               ' Do the Split
         scError(%eNone, "Split successful")                      '
         ErrFlag = %eNone                                         ' Say we 're OK
      END IF                                                      '
   END IF                                                         '
   MExit
END METHOD

METHOD  pCmdSTART(pCmd AS STRING)
'---------- Set START profile value
LOCAL nv AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN _                                       '
      MErrExit(%eNone, "START set to " + Prf.Start)               '
   IF pCmdNumOps <> 1 THEN _                                      ' Better be one operand
      MErrExit(%eFail, "Invalid or missing operand")              '
   SELECT CASE AS CONST$ UUCASE(pCmdOps(1))                       ' Which did we get?
      CASE "PRIOR"  : GOSUB CkState: Prf.Start = "PRIOR"          ' PRIOR
      CASE "LABEL" : GOSUB CkState: Prf.Start = "LABEL"           ' LABEL
      CASE "FIRST": Prf.Start = "FIRST"                           ' FIRST
      CASE "LAST" : Prf.Start = "LAST"                            ' LAST
      CASE "NEW" : GOSUB CkState: Prf.Start = "NEW"               ' NEW
      CASE ELSE: MErrExit(%eFail, "Invalid operand, choose: FIRST, LAST, PRIOR, LABEL or NEW")
   END SELECT
   scError(0, "START set to " + Prf.Start)                        ' Confirm to user
   MExitMeth
CKState:
   IF ISFALSE Prf.PState THEN _                                   ' Only if we're doing STATE
      MErrExit(%eFail, "This START option requires Profile STATE ON")
   RETURN
END METHOD

METHOD  pCmdSTATE(pCmd AS STRING)
'---------- Set State save ststus
LOCAL nv AS LONG, lclFn AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN _                                       ' No Ops?
      MErrExit(0, "STATE set to " + CHOOSE$(Prf.PState + 1, "OFF", "ON", "FEW", "MOST")) ' Tell user the value
   sProfState("RESET")                                            ' If something's changing, clear STATE cache
   SELECT CASE AS CONST$ UUCASE(pCmdOps(1))                       ' What did we get
      CASE "OFF"    : nv = %StateOff                              '
      CASE "ON"     : nv = %StateOn                               '
      CASE "FEW"    : nv = %StateFew                              '
      CASE "MOST"   : nv = %StateMost                             '
      CASE "DELETE", "DEL":                                       ' Delete immediate request?
         IF TIP.FilePath = $Empty OR _                            ' Bad mode?
            IsBrowse OR  _                                        '
            IsView OR  _                                          '
            IsSetEdit OR _                                        ' SetEdit?
            IsClip THEN _                                         ' Clip?
            MErrExit(%eFail, "STATE DELETE allowed only for normal Edit sessions")
         IF Prf.PState <> %StateMost THEN _                       ' Only acceptable if STATE MOST
            MErrExit(%eFail, "STATE DELETE allowed only if STATE MOST is in effect")
         lclFn = TIP.FilePath + ".STATE"                          ' Temp copy with .STATE on the end
         REPLACE ANY ":\/" WITH "```" IN lclFN                    ' Make : / and \ into `
         lclFn = ENV.StatePath + lclFn                            ' Add our STATE folder
         sMakeNullFile(lclFn)                                     ' Create the nullfile
         OffStateExist                                            ' Kill the exists flag
         sDoStatusBar($SBState)                                   ' re-Do the StatusBar boxes
         MErrExit(0, "STATE exemption set for the current file")  ' Tell user

      CASE "CREATE", "CRE":
         IF Prf.PState <> %StateFew THEN _                        ' Only acceptable if STATE FEW
            MErrExit(%eFail, "STATE CREATE allowed only if STATE FEW is in effect")
         OnStateForce                                             ' Force StateSave to override FEW
         MErrExit(0, "The next SAVE/END will create STATE data")  ' Tell user we've done it

      CASE ELSE     : MErrExit(%eFail, "Unknown STATE operand: " + pCmdOps(1))

   END SELECT
   IF nv = %StateOff AND Prf.PState > %StateOff THEN              ' Going from some kind of ON to OFF?
      IF Prf.Start = "PRIOR" OR Prf.Start = "LABEL" OR Prf.Start = "NEW" THEN ' A dependant option?
         Prf.PState = nv                                          ' Set the new value
         Prf.Start = "FIRST"                                      ' Reset the START value
         scError(0, "STATE set to OFF, START also set back to FIRST") ' Tell user what we're doing
         MExitMeth                                                '
      END IF                                                      '
   END IF                                                         '
   Prf.PState = nv                                                ' Set the new value
   scError(0, "STATE set to " + CHOOSE$(Prf.PState + 1, "OFF", "ON", "FEW", "MOST")) ' Tell user the value
   MExit
END METHOD

METHOD  pCmdSUBARG(pCmd AS STRING)
'---------- Set SUBMIT default argument for current profile
LOCAL Wrd1, lclCmd AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 1 THEN                                         ' Got our operand?
      SELECT CASE pCmdOpsType(1)                                  ' What've we got
         CASE %KWOFF                                              ' OFF
            Prf.SubArg = ""                                       '
         CASE ELSE                                                ' Save anything else
            Prf.SubArg = pCmdOps(1)                               '
      END SELECT                                                  '
   END IF                                                         '
   scError(0, "SUBARG is " + IIF$(ISNULL(Prf.SubArg), "OFF", Prf.SubArg))
   MExit
END METHOD

METHOD  pCmdSUBCMD(pCmd AS STRING)
'---------- Set SUBMIT sub command
LOCAL Wrd1, lclCmd AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 1 THEN                                         ' Got our operand?
      SELECT CASE pCmdOpsType(1)                                  ' What've we got
         CASE %KWOFF                                              ' OFF
            Prf.SubCmd = ""                                       '
         CASE ELSE                                                ' Save anything else
            Prf.SubCmd = pCmdOps(1)                               '
      END SELECT                                                  '
   END IF                                                         '
   scError(0, "SUBCMD is " + IIF$(ISNULL(Prf.SubCmd), "OFF", Prf.SubCmd))
   MExit
END METHOD

METHOD  pCmdSUBMIT(pCmd AS STRING)
'---------- Do the SUBMIT process
LOCAL lclCmd, lclWord, CBD, DCmd, t, JobName, W1, W2, W3, w4, BATFile, INCLFile AS STRING, lTxtp AS STRING POINTER
LOCAL i, j AS LONG, lclCursMode AS INTEGER, fl AS QUAD
LOCAL NumDel, lclTop, AppendText, recs, BrkStart, hWatchThread, MSG, RC AS LONG
LOCAL lclDrive, lclPath, CBLines AS STRING
LOCAL fptr AS STRING POINTER
LOCAL subIO AS iIO                                                ' For our I/O stuff
   MEntry
   LET subIO = CLASS "cIO"                                        '
   IF ISNOTNULL(Prf.SubCmd) THEN                                  ' SUBCMD specified?
      IF IsEQ(Prf.SubCmd, "SUB") OR IsEQ(Prf.SubCmd, "SUBMIT") THEN _
         MErrExit(%EFail, "Recursive SUBCMD detected")
      lclcmd = pCmd                                               ' Get the command line
      lclWord = sGetWord(lclCmd, %Strip, %QuoteNotSig)            ' Strip the 1st word
      me.pCmdRun(Prf.SubCmd + " " + lclCmd)                       ' Handle as a different command
      MExitMeth                                                   '
   END IF                                                         '

   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   gDosOperands = "": gResultFile = ""                            ' Null
   FOR i = 1 TO pCmdNumOps                                        ' Extract the DOS operands
      IF pCmdOpsType(i) = %OpSqb AND ISFALSE BrkStart THEN        ' Start of operands?
         BrkStart = %True                                         ' Say we've started
         gDosOperands = MID$(pCmdRaw(i), 2)                       ' Start the string
         GOSUB DelEntry                                           ' Delete from pCmdOps array
         IF RIGHT$(gDosOperands, 1) = "]" THEN                    ' Also have the closing?
            gDosOperands = LEFT$(gDosOperands, LEN(gDosOperands) - 1) ' Remove last ]
            EXIT FOR                                              ' Eliminate the NULL case
         END IF                                                   '
         gDosOperands += " "                                      ' Else just next operand
         ITERATE FOR                                              '
      END IF                                                      '
      IF pCmdOpsType(i) = %OpSqb AND BrkStart THEN                ' End of operands?
         gDosOperands += LEFT$(pCmdRaw(i), LEN(pCmdRaw(i)) - 1)   ' Add the last operand
         GOSUB DelEntry                                           ' Delete from pCmdOps array
         EXIT FOR                                                 '
      END IF                                                      '
      IF BrkStart THEN                                            ' Already building operands?
         gDosOperands += pCmdRaw(i) + " "                         ' Add as simple operand
         GOSUB DelEntry                                           ' Delete from pCmdOps array
         ITERATE FOR                                              ' Onward
      END IF                                                      '
      IF IsEQ(pCmdOps(i), "DEBUG") THEN                           ' Hold DOS Window for messages
         MSG = %True                                              ' Remember that
         GOSUB DelEntry                                           ' Delete from pCmdOps array
         ITERATE FOR                                              ' Onward
      END IF                                                      '
   NEXT i
   pCmdNumOps -= j                                                ' Adjust now for removed operands

   ENV.Jobnumber = ENV.Jobnumber + 1                              ' Bump it
   gJobID = "JOB" + FORMAT$(ENV.Jobnumber, "00000")               ' Build ID
   sIniSetString("General", "JobNumber", FORMAT$(ENV.Jobnumber))  ' Stuff it back

   lclCursMode = %CursBad                                         ' Pretend a bad cursor location
   IF ISTRUE me.RangeSet("", %True, %False) THEN MExitMeth        ' Setup any range command operands
   me.FindReset                                                   ' Reset the parse output fields
   IF sCRTParse("ASU") THEN MExitMeth                             ' Scan (All, Subset, User)
   me.FindSetup                                                   ' Copy CRT parse's answers
   CfSet(%CRTDX)                                                  ' Tell search hands off Invisible status

   IF ISFALSE DotSet AND _                                        ' If no line range
      ISFALSE IsCfU AND ISFALSE IsCfNU AND ISFALSE IsCfX AND ISFALSE IsCfNX THEN ' Or U / NU / X / NX
      CfSet(%CRTALL)                                              ' Force ALL if no modifiers
   END IF

   errFlag = %eNone                                               ' Reset any Pending conflict flag
   lclTop = TopScrn                                               '

   i = me.Search(0, lclCursMode)                                  ' Do the initial search
   IF i THEN                                                      ' We have the 1st SUBMIT line
      t = UUCASE(LTxtG(cfFLine))                                  ' Get it's text
      w1 = sGetWord(t, %Strip, %QuoteNotSig)                      ' Get the 1st 4 'words'
      w2 = sGetWord(t, %Strip, %QuoteNotSig)                      '
      w3 = sGetWord(t, %Strip, %QuoteNotSig)                      '
      w4 = sGetWord(t, %Strip, %QuoteNotSig)                      '
      IF w1 = "//" AND w2 = "JOB" THEN                            ' DOS Job card?
         Jobname = w3                                             ' 3rd operand is JobName
      ELSEIF LEFT$(w1, 2) = "//" AND w2 = "JOB" THEN              ' OS Job card?
         JobName = MID$(w1, 3)                                    ' Then jobname follows the //
      ELSEIF w1 = "*" AND w2 = "$$" AND w3 = "JOB" AND LEFT$(w4, 4) = "JNM=" THEN ' Power JCl?
         IF INSTR(w4, ",") > 0 THEN                               ' Trailing comma?
            JobName = MID$(LEFT$(w4, INSTR(w4, ",") - 1), 5)      ' Extract it
         ELSE                                                     '
            JobName = MID$(w4, 5)                                 '
         END IF                                                   '
      END IF                                                      '

      '----- Ready the output file
      IF ISNOTNULL(JobName) THEN                                  ' If we have a JobName
         gSubmitFile = ENV.SubmitDir + "\SUBMIT_" + JobName + "_" + gJobID + TIP.Extn ' Get SUBMIT temp file allocated
      ELSE                                                        '
         gSubmitFile = ENV.SubmitDir + "\SUBMIT_" + gJobID + TIP.Extn ' Get SUBMIT temp file allocated
      END IF                                                      '
      subIO.Setup("OR", "", "", gSubmitFile)                      ' Setup the IO
      IF subIO.EXEC THEN _                                        ' Go OPEN the file
         MErrExit(%eFail, subIO.ResultMsg)                        ' Oops?  Bail out
      GOSUB PutOne                                                ' Go handle one
   END IF
   DO WHILE i                                                     ' Do till not found
      i = me.Search(1, lclCursMode)                               ' Do re-Finds
      IF i THEN GOSUB PutOne                                      ' Found one
   LOOP                                                           '
   subIO.Close                                                    ' Close the FBO

   IF INSTR(UUCASE(ENV.SubmitCmd), "~R") > 0 OR _                 ' Results file being requested?
      INSTR(UUCASE(ENV.SubmitCmd), "^R") > 0 THEN                 '
      IF ISNOTNULL(JobName) THEN                                  ' Base on Jobname
         gResultFile = ENV.SubmitDir + "\SUBRESULTS_" + JobName + "_" + gJobID + ".Txt"
      ELSE                                                        '
         gResultFile = ENV.SubmitDir + "\SUBRESULTS_" + gJobID + ".txt"
      END IF                                                      '
      sMakeNullFile(gResultFile)                                  ' Create as empty file

      '----- Start the File Watch
      fptr = VARPTR(gResultFile)                                  '
      THREAD CREATE sResultFileWatchThread(BYVAL fptr) 65536, TO hWatchThread ' Fire up the thread
      IF hWatchThread = 0 THEN _                                  ' Failed?
         MErrExit(%eFail, "Watch of Result file could not be started")
      SLEEP 50                                                    ' Wait a bit
      THREAD STATUS hWatchThread TO i                             ' See if running OK
      IF i <> 259 THEN _                                          ' If running OK STATUS returns &H103 (See Help)
         MErrExit(%eFail, "Watch of Result file could not be started")
      THREAD CLOSE hWatchThread TO i                              ' Free up our handle
      SLEEP 1000                                                  '
   END IF                                                         '
   DCmd = ENV.SubmitCmd                                           ' Get the SUBMIT command
   gSubmitType = "S"                                              ' Set Submit type
   IF me.MacSubst(DCmd) THEN _                                    ' Do substitution
      MErrExit(%EFail, "Malformed macro operand in SUBMIT prototype command")

'   debug DCmd
   '----- Create the BAT file
   BATFile = ENV.INIPath + "SPFBatch.BAT"                         '
   subIO.Setup("OR", "", "", BATFile)                             ' Setup for IO
   IF subIO.EXEC THEN _                                           ' Go OPEN the file
      MErrExit(%eFail, subIO.ResultMsg)                           ' Oops?  Bail out
   PRINT # subIO.FNum, DCmd                                       ' Write the command out
   IF ISFALSE MSG THEN PRINT # subIO.FNum, "EXIT"                 ' Close the BAT file if not in DEBUG mode
   subIO.Close                                                    ' Close the File

   '----- Switch to the EXE path
   lclPath = CURDIR$                                              ' Locate where we are
   IF MID$(lclPath, 2, 1) = ":" THEN _                            ' Extract Drive if present
      lclDrive = LEFT$(lclPath, 2)                                ' and save it

   IF MID$(ENV.SubmitDir, 2, 1) = ":" THEN _                      ' See if Submit Path has drive
      CHDRIVE LEFT$(ENV.SubmitDir, 2)                             ' If so, go to it
   CHDIR ENV.SubmitDir                                            ' Switch to EXE path

   '----- Fire it up based on DEBUG / NODEBUG
   gLoopCtr = -1                                                  ' Suspend loop detection
   IF MSG THEN                                                    ' MSG
      t = ENVIRON$("COMSPEC") + " /K " + sGetShortName(ENV.INIPath + "SPFBatch.bat")
      SHELL t, 1
   ELSE                                                           ' MSG omitted
      t = ENVIRON$("COMSPEC") + " /C " + sGetShortName(ENV.INIPath + "SPFBatch.bat")
      SHELL t, 0
   END IF                                                         '
   RESET gLoopCtr                                                 ' Reset the counter

   '----- Check for errors
   IF ERR THEN _                                                  ' Tell user result
      MErrExit(%eFail, "Error issuing SUBMIT " + ERROR$)          '

   '----- Restore PATH
   IF ISNOTNULL(lclDrive) THEN CHDRIVE lclDrive                   ' Switch drive if needed
   CHDIR lclPath                                                  ' put back the original path
   TopScrn = lclTop                                               '

   '----- We're all done!
   scError(0, FORMAT$(recs) + IIF$(recs <> 1, " lines", " line") + " submitted for " + IIF$(ISNULL(JobName), "(JOB", JobName + "(JOB") + FORMAT$(ENV.Jobnumber, "00000") + ")")
   MExitMeth

PutOne:
   IF IsLInsertLine(cfFLine) THEN RETURN                          ' Exempt Insert lines
   t = LTxtG(cfFLine)                                             ' Get the text
   IF uucase(MID$(t, ENV.SubmitInclCol, LEN(ENV.SubmitInclKey))) <> uucase(ENV.SubmitInclKey) THEN  ' If not an #INCLUDE line
      PRINT # subIO.FNum, t                                       ' Just write the line to SUBMIT file
      INCR recs                                                   ' count it
   ELSE                                                           ' Else do the INCLUDE
      INCLFile = TRIM$(MID$(t, ENV.SubmitInclCol + LEN(ENV.SubmitInclKey)))   ' Get the INCLUDE Filename
      sUnQuote(INCLFile)                                          ' Remove any quotes
      IF INSTR(INCLFile, "\") = 0 THEN _                          ' Already include a parh?
         INCLFile = TIP.Path + INCLFile                           ' No, add the current file's path
      IF ISFALSE ISFILE(INCLFile) THEN                            ' If it doesn't exist
         MErrExit(%EFail, "#INCLUDE filename (" + INCLFile + ") not found")
      END IF                                                      '

      '----- Setup to read the Included filename
      TOP.Setup("E","", "", INCLFile)                             ' Setup the TOP area
      IF TOP.EXEC THEN _                                          ' Go validate
         scError(%eFail,TOP.ResultMsg)                            ' Oops?  Bail out
      aPrf.SetProfName(TOP.Profile, %True)                        ' Get aPrf set to the file being copied
      aPrf.ReadAll(%True)                                         ' Get the Profile variables
      me.CopyAFile(0, TOP, APrf, 0, 0, 0, %False, CBLines)        ' Go load the data to CBLines

      '----- Extract CBlines to the output file
      DO WHILE LEN(CBLines) > 1                                   ' Loop extracting text lines
         t = EXTRACT$(CBLines, $CRLF)                             ' Get one line
         t = TAB$(t, Prf.ImportTabs)                              ' Do any needed tab conversion
         PRINT # subIO.FNum, t                                    ' Just write the line to SUBMIT file
         INCR recs                                                ' count it
         CBLines = REMAIN$(CBLines, $CRLF)                        ' Make CBLines = the remainder of the line
      LOOP                                                        ' Loop till all lines extracted

   END IF                                                         '
   RETURN                                                         '

DelEntry:
   ARRAY DELETE pCmdOps(i): ARRAY DELETE pCmdRaw(i): ARRAY DELETE pCmdOpsType(i)
   DECR i: INCR j                                                 ' So we stay on the same index number
   RETURN
END METHOD

METHOD  pCmdSWAP(pCmd AS STRING)
'---------- Switch tab Pages
LOCAL Wrd, lclCmd AS STRING, i AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN                                         ' No Ops? Set a NEXT operand
      pCmdOpsType(1) = %KWNEXT                                    '
   END IF                                                         '

   '----- We have an operand, see if it's OK
   IF INSTR(CHR$(%KWNEXT, %KWHOME, %KWPREV, %KWPRIOR, %KWFIRST, %KWLAST, %KWOPEN, %KWLIST), CHR$(pCmdOpsType(1))) = 0 THEN _
      MErrExit(%eFail, "Invalid operand for SWAP")                '
   SELECT CASE pCmdOpsType(1)                                     ' Split off by ON/OFF
      CASE %KWNEXT                                                '
         i = IIF(PgNumber = TabsNum, 1, PgNumber + 1)             ' Calc tab # we're going to
      CASE %KWPREV                                                '
         i = IIF(PgNumber > 1, PgNumber - 1, TabsNum)             ' Calc tab # we're going to
      CASE %KWPRIOR                                               '
         i = IIF(gTabStackNum > 1, gTabStack(2), 0)               ' Get previous page number
         IF i = 0 OR i > TabsNum THEN i = 1                       ' Don't go to non-existant tab
      CASE %KWFIRST                                               '
         i = IIF(TabsNum > 1, 2, 1)                               ' First real tab (or FM if no Edit tabs)
      CASE %KWHOME                                                '
         i = 1                                                    ' FM tab
      CASE %KWLAST                                                '
         i = TabsNum                                              ' Goto last tab
      CASE %KWOPEN, %KWLIST                                                '
         pCmdRecall("RECALL OPEN")                                '
         MExitMeth                                                '
   END SELECT                                                     '
   gTabSwitch = i                                                 ' Set tab to go to
   CmdStackNum = 0                                                ' If a swap, no more commands allowed
   MExitMeth

END METHOD

METHOD  pCmdTABS(pCmd AS STRING)
'---------- Turn TABS ON or OFF
LOCAL nv AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   nv = sOnOff(1)                                                 ' Go get True/False value for operand 1
   IF nv < 0 THEN MExitMeth                                       ' Error, exit, error msg already issued
   Prf.Tabs = nv                                                  ' Set the new value
   scError(0, "TABS set to " + IIF$(nv, "ON", "OFF"))             ' Tell user new value
   MExit
END METHOD

METHOD  pCmdTAG(pCmd AS STRING)
'---------- TAG command
LOCAL action, lclWord, lclTag, valid AS STRING, lclCursMode AS INTEGER
LOCAL i, AllCount AS LONG, ff AS QUAD
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN _                                       ' Exit if nothing to do
      MErrExit(%eFail, "Missing TAG operands")                    '
   IF pCmdOpsType(1) = %OpTag THEN                                ' Got a tag operand?
      IF sTagVal(pCmdOps(1), %False) THEN _                       ' See if 1st is a valid Tag
         MErrExit(%eFail, "Invalid TAG operand")                  '
      lclTag = LSET$(UUCASE(pCmdOps(1)), 8)                       '
      ARRAY DELETE pCmdOps(1): ARRAY DELETE pCmdRaw(1): ARRAY DELETE pCmdOpsType(1)
      DECR pCmdNumOps                                             ' Strip off the Tag now
   END IF                                                         '
   lclCursMode = Zone                                             ' Save CursMode at start of command
   me.ResetFunc(%ResetFind)                                       ' Clear previous FIND hilighting

   IF ISTRUE me.RangeSet("", %False, %False) THEN MExitMeth       ' Setup any range command operands
   me.FindReset                                                   ' Reset the parse output fields
   IF sCRTParse("ATSDMNEC1PU") THEN MExitMeth                     ' Scan (All, Tag, Subset, Direct, Modifier, Negative, Exclude, Cols, Lit1, Clr, User)
   me.FindSetup                                                   ' Copy CRT parse's answers

   IF IsCfOff THEN INCR i                                         ' Count basic operands
   IF IsCfOn THEN INCR i                                          '
   IF IsCfToggle THEN INCR i                                      '
   IF IsCfAssert THEN INCR i                                      '
   IF IsCfSet THEN INCR i                                         '

   IF i > 1 THEN                                                  ' Too many?
      MErrExit(%eFail, "Only one of ON, OFF, SET, TOGGLE or ASSERT allowed")
   ELSEIF i = 0 THEN                                              ' If no ON/OFF/TOGGLE/ASSERT then its ON
      CFSet(%CrtOn)                                               ' None?  Set ON
   END IF

   IF (IsCfOn OR IsCfSet OR IsCfToggle) AND ISNULL(lcltag) THEN _ '
      MErrExit(%eFail, "Target tag was not provided for " + IIF$(IsCfOn, "ON", IIF$(IsCfSet, "SET", "TOGGLE")))

   IF IsCfALL AND ISFALSE DotSet AND ISFALSE IsCfLit1 AND _
      ISNULL(lcltag) AND ISFALSE IsCfOff AND _
      ISFALSE IsCfX AND ISFALSE IsCfNX THEN _
      MErrExit(%eFail, "ALL specified but no line range or search literal")

   IF ISFALSE IsCfFirst AND ISFALSE IsCfLast AND _                '
      ISFALSE IsCfNext AND ISFALSE IsCfPrev AND _                 '
      (ISTRUE DotSet OR ISTRUE IsCfLit1) THEN _  '                '
         CfSet(%CRTAll)                                           ' Default to All

   IF ISFALSE IsCfFirst AND ISFALSE IsCfLast AND _                '
      ISFALSE IsCfNext AND ISFALSE IsCfPrev AND _                 '
      (ISFALSE DotSet AND ISNOTNULL(lclTag)) THEN _               '
         CfSet(%CRTAll)                                           ' Default to All

   IF ISFALSE IsCfLit1 THEN                                       ' Lit1 present?
      IF ISFALSE IsCfAll AND ISFALSE IsCfFirst AND ISFALSE IsCfLast THEN ' No Lit1, better have ALL or a limit
         MErrExit(%eFail, "No search string, but no ALL/FIRST/LAST specified")
      END IF
      IF IsCfWord OR IsCfPrefix OR IsCfSuffix OR _                ' No extraneous operands
         IsCfFCol OR IsCfTCol THEN  _
         MErrExit(%eFail, "No search string, but string options are present")
      IF IsCfNF THEN _                                            ' No Literal and NF asked for?
         MErrExit(%eFail, "No search string, but NF request is present")
   END IF

   IF IsCfAssert THEN                                             ' Assert?
      IF ISFALSE IsCfLit1 THEN _                                  ' Lit1 present?
         MErrExit(%eFail, "No search string, but ASSERT requested")
      IF ISFALSE DotSet THEN _                                    ' Lit1 present?
         MErrExit(%eFail, "ASSERT, but no line range or tag specified")
      IF IsCfNF THEN                                              ' If NF actually specified
         CfClear(%CRTNF)                                          ' Invert it
      ELSE                                                        '
         CfSet(%CRTNF)                                            ' If not specified, set it
      END IF                                                      '
   END IF                                                         '


   AllCount = 0                                                   ' Reset counter
   IF IsCfOn THEN                                                 ' Build message text
      action = "Tag set ON on "                                   '
   ELSEIF IsCfOff THEN                                            '
      action = "Tagged set OFF on "                               '
   ELSEIF IsCfToggle THEN                                         '
      action = "Tag Toggled on "                                  '
   ELSEIF IsCfAssert THEN                                         '
      action = "Tag set off by ASSERT on "                        '
   ELSEIF IsCfSet THEN                                            '
      action = "Tag SET modified "                                '
   END IF
   IF IsCfAll THEN                                                ' The ALL version?
      IF IsCfNF THEN                                              ' Choose search routine based on NF status
         i = me.NFSearch(0, lclCursMode)                          ' NF version
      ELSE                                                        '
         i = me.Search(0, lclCursMode)                            ' Normal version
      END IF                                                      '
      lclCursMode = %CursBad                                      ' Make it bad for subsequent searches
      IF i THEN                                                   ' See if we have one
         IF IsCfSet THEN _                                        ' If SET clear it so sTagChange will always set it
            LTagS(cfFLine) = $BlankLNo                            '
         AllCount += me.TagChange(lclTag)                         ' Go do the Tag Change
      END IF                                                      '
      DO WHILE i                                                  ' Do till not found
         IF IsCfNF THEN                                           ' Choose search routine based on NF status
            i = me.NFSearch(1, lclCursMode)                       ' NF version
         ELSE                                                     '
            i = me.Search(1, lclCursMode)                         ' Normal version
         END IF                                                   '
         IF i THEN                                                '
            IF IsCfSet THEN _                                     ' If SET clear it so sTagChange will always set it
               LTagS(cfFLine) = $BlankLNo                         '
            AllCount += me.TagChange(lclTag)                      ' Go do the Tag Change
         END IF                                                   '
      LOOP                                                        '
      IF AllCount = 0 THEN                                        ' Issue approp. message
         scError(nMac(%eFail), "No lines (re)tagged")             '
      ELSE                                                        '
         IF IsCfSet THEN GOSUB ClearRest                          ' If SET then clear all the rest
         scError(0, action + FORMAT$(AllCount) + IIF$(AllCount > 1, " lines", " line"))
      END IF                                                      '
   ELSE                                                           ' The non-All version
      IF IsCfNF THEN                                              ' Choose search routine based on NF status
         i = me.NFSearch(0, lclCursMode)                          ' NF version
      ELSE                                                        '
         i = me.Search(0, lclCursMode)                            ' Normal version
      END IF                                                      '
      IF i THEN                                                   ' Did we find it
         IF IsCfSet THEN _                                        ' If SET clear it so sTagChange will always set it
            LTagS(cfFLine) = $BlankLNo                            '
         me.TagChange(lclTag)                                     ' Go do the change
         IF IsCfSet THEN GOSUB ClearRest                          ' If SET then clear all the rest
      ELSE                                                        ' Not found, issue message
         scError(nMac(%eFail), IIF$(IsCfNext, "Top of data reached", "Bottom of data reached"))
         cfNotFound = IIF(IsCfNext, 1, 2)                         ' Remember we failed
      END IF                                                      '
   END IF                                                         '
   gfXRebuild = %True                                             ' Force exclude rebuild
   OnUndoFlag                                                     ' Remember we changed something
   MExitMeth

ClearRest:
   FOR i = 1 TO LastLine                                          ' Reset tags for non-SET winners
      IF ISFALSE IsLData(i) THEN ITERATE FOR                      ' Just do data lines
      IF LTagG(i) = lclTag THEN ITERATE FOR                       ' Also ignore if a SET winner
      IF LTagG(i) = $BlankLNo THEN ITERATE FOR                    ' Also ignore if blank tag
      LTagS(i) = $BlankLNo                                        ' A loser, blank the tag
      me.UpdLControl(i)                                           ' Reset line number
      INCR AllCount                                               '
   NEXT i                                                         '
   RETURN                                                         '
END METHOD

METHOD  pCmdTEST(pCmd AS STRING)
'---------- TEST - Run the SPFTest program
LOCAL Oprand, lclcmd AS STRING
LOCAL RetC, i AS LONG
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 1 THEN                                         '
      SELECT CASE AS CONST$ UUCASE(pCmdOps(1))                    '
         CASE "REGEX"            : Oprand = "1"                   ' RegEx
         CASE "MAP", "MAPPING"   : Oprand = "2"                   ' Map, Mapping
         CASE "CALC"             : Oprand = "3"                   ' Calc
         CASE ELSE                                                '
                  MErrExit(%eFail, "Invalid TEST operand")        '
      END SELECT                                                  '
   END IF
   lclcmd = $DQ + ENV.EXEPath + "SPFTest.EXE" + $DQ + " " + Oprand
   RetC = SHELL(lclcmd, 1)
   IF ERR THEN                                                    ' Tell user result
      scError(%eFail, "Error issuing command " + ERROR$)          '
   END IF                                                         '
   MExit
END METHOD

METHOD  pCmdTOP(pCmd AS STRING)
'---------- Alias for UP MAX
   pCmdUP("UP MAX ")                                              '
END METHOD

METHOD  pCmdUP(pCmd AS STRING)
'---------- Scroll UP
LOCAL ScrAmt, DataIX, ForceLine, Lines2Page AS LONG, ScrCmd, lclCmd AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN                                         ' No Ops, fudge it
      me.CmdParse("UP " + Prf.Scroll)                             ' Re-parse it
   END IF
   ScrCmd = pCmdOps(1)                                            ' Pick up the operand
   IF pCmdOpsType(1) <> %OpNum AND INSTR(CHR$(%KWHALF, %KWPAGE, %KWFULL, %KWDATA, %KWMAX, %KWCSR), CHR$(pCmdOpsType(1))) = 0 THEN _
      MErrExit(%eFail, "Invalid scroll amount")                   '
   IF IsCData THEN                                                ' No, In data area?
      me.CurSetReq(%Position, CsrAbsLine, CCol - gLNPadCol + Offset, %False) ' Set cursor set attempt
   ELSEIF IsCLinN THEN                                            ' No, In Line Number area?
      me.CurSetReq(%Position, CsrAbsLine, 0, %False)              ' Set cursor set attempt
   END IF
   IF VAL(ScrCmd) <> 0 THEN                                       ' A Number?
      ScrAmt = VAL(ScrCmd)                                        ' Yes, use it
   ELSEIF pCmdOpsType(1) = %KWHALF THEN                           ' Half?
      ScrAmt = INT((gwScrHeight - 2) / 2)                         ' Yes, set it
      ScrAmt = ScrAmt \ Prf.HexMode                               ' Adjust for Hex
   ELSEIF pCmdOpsType(1) = %KWPAGE THEN                           ' Page?
      ScrAmt = INT((gwScrHeight - 2 - Prf.Cols))                  ' Yes, set it
      ScrAmt = ScrAmt \ Prf.HexMode                               ' Adjust for Hex
      Prf.ScrlPageSus = %False                                    ' Clear suspend PAGE mode
   ELSEIF pCmdOpsType(1) = %KWFULL THEN                           ' Full?
      ScrAmt = INT((gwScrHeight - 2 - Prf.Cols))                  ' Yes, set it
      ScrAmt = ScrAmt \ Prf.HexMode                               ' Adjust for Hex
   ELSEIF pCmdOpsType(1) = %KWDATA THEN                           ' Data?
      ScrAmt = INT((gwScrHeight - 2) - Prf.Cols) - 1              ' Yes, set it
      ScrAmt = ScrAmt \ Prf.HexMode                               ' Adjust for Hex
   ELSEIF pCmdOpsType(1) = %KWMAX THEN                            ' Max?
      ScrAmt = 9999999                                            ' Yes, set a BIG number
   ELSEIF IsCData THEN                                            ' No, In data area?
      DataIX = CsrAbsLine                                         ' Locate data line cursor was on
      ForceLine = me.CRPBack(%mVisible, DataIX, gwScrHeight - 3 - Prf.Cols) ' Backup by screensize to put line on bottom
      IF ForceLine <> sGetIX(3 + Prf.Cols) THEN                   ' As long as we move
      ELSE                                                        ' Else kill ForceLine
         ForceLine = 0                                            '
         ScrAmt = gwScrHeight - 3 - Prf.Cols                      ' Go back 1 screen
         ScrAmt = ScrAmt \ Prf.HexMode                            ' Adjust for Hex
      END IF                                                      '
   ELSEIF IsCLinN THEN                                            ' No, In Line Number area?
      DataIX = CsrAbsLine                                         ' Locate data line cursor was on
      ForceLine = me.CRPBack(%mVisible, DataIX, gwScrHeight - 3 - Prf.Cols) ' Backup by screensize to put line on bottom
      IF ForceLine <> sGetIX(3 + Prf.Cols) THEN                   ' As long as we move
      ELSE                                                        ' Else kill ForceLine
         ForceLine = 0                                            '
         ScrAmt = gwScrHeight - 3 - Prf.Cols                      ' Go back 1 screen
         ScrAmt = ScrAmt \ Prf.HexMode                            ' Adjust for Hex
      END IF                                                      '
   ELSE                                                           '
      ScrAmt = gwScrHeight - 2 - Prf.Cols                         ' No, Scroll back one screen
      ScrAmt = ScrAmt \ Prf.HexMode                               ' Adjust for Hex
   END IF                                                         '

   '----- Adjust the top of screen location now
   IF LEFT$(Prf.EOL, 4) = "AUTO" AND pCmdOpsType(1) = %KWPAGE THEN' Page in EOL=AUTO mode?
      IF Prf.PageFlag = 2 THEN                                    ' Page SCROLL?
         Lines2Page = me.CRPBack2Page                             ' Get # lines to previous =PAGE> line (or top)
         IF Lines2Page <= (gwScrHeight - Prf.Cols - 2) THEN       ' Will this page all fit on the screen?
            TopScrn = me.CRPBack(%mPage, TopScrn, 1)              ' Move TopScrn by a PAGE
         ELSE                                                     '
            ScrAmt = INT((gwScrHeight - 2 - Prf.Cols))            ' Calc like FULL does
            ScrAmt = ScrAmt \ Prf.HexMode                         '
            TopScrn = me.CRPBack(%mVisible, TopScrn, ScrAmt)      ' Move TopScrn by scroll amount
            Prf.ScrlPageSus = %False                              '
         END IF                                                   '
      ELSEIF Prf.PageFlag = 1 THEN                                ' PAGE ON mode?
         TopScrn = me.CRPBack(%mPage, TopScrn, 1)                 ' Move TopScrn by a PAGE
      ELSE                                                        '
         TopScrn = me.CRPBack(%mVisible, TopScrn, ScrAmt)         ' Move TopScrn by scroll amount
      END IF                                                      '
   ELSE                                                           ' Not AUTO PAGE
      IF ForceLine = 0 THEN                                       ' If we haven't already calculated it
         TopScrn = me.CRPBack(%mVisible, TopScrn, ScrAmt)         ' Move TopScrn by scroll amount
      ELSE                                                        '
         TopScrn = ForceLine                                      ' Use the forced one
      END IF                                                      '
   END IF                                                         '
   ErrFlag = %eNone                                               ' Say we 're ok
   MExit
END METHOD

METHOD  pCmdUNDOREDO(pCmd AS STRING)
'---------- Go back to prev/next checkpoint
LOCAL i, j, k AS LONG, MEditStr AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF Prf.UndoNumber = 0 THEN _                                   ' Are we running UNDO?
      MErrExit(%eFail, "UNDO has been turned off")                ' No? Oops!

   i = UndoCurr                                                   ' Get index of the last save
   IF pCmd = "UNDO" THEN                                          ' UNDO?
      j = IIF(UndoCurr = 1, Prf.UndoNumber, UndoCurr - 1)         ' Point at the previous UNDO entry
   ELSE                                                           ' REDO
      j = IIF(UndoCurr = Prf.UndoNumber, 1, UndoCurr + 1)         ' Point at the next UNDO entry
   END IF                                                         '
   IF me.UndoTimeGet(j) = 0 THEN _                                ' Make sure it's valid
      GOSUB ErrOut                                                ' Oops!
   IF pCmd = "UNDO" THEN                                          ' UNDO?
      IF me.UndoTimeGet(j) > me.UndoTimeGet(i) THEN _             ' Make sure it's valid
         GOSUB ErrOut                                             ' Oops!
   ELSE                                                           '
      IF me.UndoTimeGet(j) < me.UndoTimeGet(i) THEN _             ' Make sure it's valid
         GOSUB ErrOut                                             ' Oops!
   END IF                                                         '
   UndoCurr = j                                                   ' Make the backup the current

   '---------- Copy the TopScr/LastLine
   TopScrn = me.UndoTopScrGet(j)                                  ' Just copy it
   LastLine = me.UndoLastLineGet(j)                               '
   LastReal = me.UndoLastRealGet(j)                               '
   IF me.UndoModifiedGet(j) THEN OnModdFlag ELSE OffModdFlag      '
   IF IsMedit THEN                                                ' If MEdit, must restore individual modified status
      MEditStr = me.UndoMEditGet(j)                               ' Fetch string
      FOR k = 1 TO LEN(MEditStr)                                  ' Restore the MEdit modified values
         IF MID$(MEditStr, k, 1) <> " " THEN _                    ' Just valid T/F entries
            me.MEditFlagSet(k, IIF(MID$(MEditStr, k, 1) = "T", %True, %False))
      NEXT j                                                      '
   END IF
   '---------- Go get the IX, L(), TW() and T() arrays
   k = me.UndoGet(j)                                              ' Load the checkpoint, get PROFLine status
   gfXRebuild = %True                                             ' Force exclude rebuild
   me.ResetFunc(%ResetCommand)                                    ' Clear any pending Line commands
   IF k = &H00000001 THEN me.ProfDisp                             ' Build the Profile display if needed
   OnRenumFlag                                                    ' Let sRenum proceed
   OffUndoFlag                                                    ' So we don't re-save what we just restored
   MExitMeth                                                      '

ErrOut:
   MErrExit(%eNone, "No (more) UNDO entries")                     ' Oops!
   RETURN
END METHOD

METHOD  pCmdUnNumber(pCmd AS STRING)
'---------- Do the UNNUMBER command
LOCAL x, xList, NumType, t AS STRING
LOCAL i, seqCol, seqLen, wSeqCol AS LONG
   MEntry

   '----- Make sure the file extension is an editable file
   NumType = TRIM$(Prf.NumType)                                   ' Get local copy of NumType
   x = "," + UUCASE(MID$(TP.TIPExtn, 2)) + ","                    ' Get extension surrounded by commas
   XList = "," + UUCASE(ENV.FMFFList) + ","                       ' Build XList here since a ProfSkip might have added an entry
   XList = REMOVE$(XList, " ")                                    ' Eliminate any spaces
   IF INSTR(XList, x) THEN                                        ' Is tis file in the exempt list?
      MErrExit(%eFail, "Current filetype is in the non-editable filetype list") ' Oops!
   END IF

   '----- See if a valid NUMTYPE exiats
   IF ISNULL(NumType) THEN                                        ' Is NUMTYPE null?
      MErrExit(%eFail, "There is no valid NUMTYPE available")     ' Oops!
   END IF                                                         '
   IF INSTR(NumType, "-") THEN                                    ' Negative columns?
      i = sDoMsgBox("With the current RECFM (" + Prf.RECFM + ")," + $CRLF + _
                 "the specified NUMTYPE (" + NumType + ") could result in" + $CRLF + _
                 "blanking sequence numbers in differing columns.  Do you want to continue?", _
                 %MB_YESNO OR %MB_USERICON OR %MB_DEFBUTTON2, "SPFLite")
      IF i = %IDNO THEN                                           '
         MErrExit(%eNone, "UnNumbering cancelled")                ' Exit if NO entered
      END IF                                                      '
   END IF                                                         '

   '----- Setup the sequence columns
   seqCol = VAL(PARSE$(NumType, " ", 1))                          ' Extract start column
   seqLen = VAL(PARSE$(NumType, " ", 2))                          ' Extract end column
   seqLen = ABS(ABS(seqLen) - ABS(seqCol)) + 1                    ' Calc length of sequence field

   '----- Do do the UNNUM
   FOR i = 1 TO LastLine                                          ' Spin through the lines
      IF IsLData(i) THEN                                          ' Only Data lines
         t = LTxtG(i)                                             ' Get the text line
         wSeqCol = seqCol                                         ' Default sequence column
         IF seqCol < 0 THEN wSeqCol =  LEN(t) + seqCol + 1        ' If negative columns, make it an actual
         IF wSeqCol + seqLen -1 > LEN(t) THEN _                   ' If not long enough for a Seq field, it's an eror
            t = LSET$(t, wSeqCol + seqLen -1)                     ' Lengthen line if too short
         MID$(t, wSeqCol, seqLen) = SPACE$(seqLen)                ' Stuff it blanks
         me.LTxtSet(i, t)                                         ' Stuff it back
         me.AttrScan(i)                                           ' Recolorize
      END IF                                                      '
   NEXT i                                                         '
   MExit
END METHOD

METHOD  pCmdVIEW(pCmd AS STRING)
'---------- OPEN a new file in VIEW mode
LOCAL SetProf, fn, MSG AS STRING
   MEntry
   '----- Do basic parsing
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   ENV.PMode = %MView                                             ' Set to VIEW mode
   fn = IIF$(pCmdNumOps > 0, pCmdOps(1), "")                      '
   fn = IIF$(fn = "*", sGetFnClipboard, fn)                       ' Swap clipboard if asked for

   '----- See if a specific Profile name provided
   IF pCmdNumOps > 1 AND LEFT$(pCmdOps(2), 1) <> "." THEN _       ' A profile request?
      MErrExit(%eFail, "Unknown VIEW operand: " + pCmdOps(2))     ' Issue error
   SetProf = IIF$(pCmdNumOps > 1, UUCASE(MID$(pCmdOps(2), 2)), "")' Extract the Profile name

   '----- Setup and validate
   TOP.Setup("EAP", SetProf, "Select file for View", fn)          ' Setup IO
   IF TOP.EXEC THEN
      IF TOP.RSTabFound THEN                                      ' Open elsewhere?
         MSG = TOP.ResultMsg                                      ' Save message arounf Tab switch
         TP = Tabs(TOP.TabFound)                                  ' Switch to it
         gTabSwitch = TOP.TabFound                                ' Set global switch
         MErrExit(%eFail, MSG)                                    ' Issue error and exit
      END IF                                                      '
      MErrExit(%eFail, TOP.ResultMsg)                             '
   END IF
   fn = TOP.FilePath: SetProf = TOP.Profile                       ' Copy fn/profile in case from Prompted filename
   me.LoadTheText(fn, SetProf)                                    ' Go load in the data

   MExit
END METHOD

METHOD  pCmdVSAVE(pCmd AS STRING)
'---------- Do a Virtual Save
LOCAL lclTop, Quiet AS LONG, fn, t AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps > 0 AND pCmdOps(1) = "Quiet" THEN Quiet = %True  ' See if Quiet request

   IF IsMEdit THEN _                                              ' MEDIT active?
      MErrExit(%eFail, "VSAVE not supported for MEDIT sessions")  '

   lclTop = TopScrn                                               ' Save where we are

   IF (IsBrowse OR IsView) AND ISFALSE Quiet THEN _               ' BROWSE/VIEW mode?
      MErrExit(%eFail, "VSAVE disallowed, you're in BROWSE/VIEW Mode") '

   IF (TIP.FilePath = $Empty) AND ISFALSE Quiet THEN _            ' (Empty) still?
      MErrExit(%eFail, "VSAVE disallowed, No FileName yet")       '

   IF ISNULL(VSaveFile) THEN                                      ' If we've not done so before
      fn = ENV.VSavePath + "VSAVE."                               ' Build VSave filename
      fn += RIGHT$(DATE$, 4) + LEFT$(DATE$, 2) + MID$(DATE$, 4, 2)' Add the date
      fn += LEFT$(TIME$, 2) + MID$(TIME$, 4, 2) + RIGHT$(TIME$, 2) + RIGHT$(FORMAT$(sOneSecondTimer), 2) ' Add the time
      t = TIP.Path: REPLACE ":\" WITH "." IN t: REPLACE ":" WITH "." IN t: REPLACE "\" WITH "." IN t
      fn += "." + t + "." + TIP.File                              ' Add the tab number and filename
      VSaveFile = fn                                              '
   END IF                                                         '

   '----- Set criteria for the line range and write the file
   me.FindPushPop("PUSH")                                         ' Save FIND stuff
   me.FindSetAll                                                  ' Set parse search output fields for ALL
   me.RangeSetRange(2, LastLine - 1)                              ' Everything gets written

   '----- Range is set, now go write the file
   TOP.Setup("OEC", "", "", VSaveFile)                            '
   IF me.WriteFile(Quiet) THEN                                    ' OK, Write it
      IF ISFALSE Quiet THEN scError(0, "File VSaved")             '
   ELSE                                                           '
      IF ISFALSE Quiet THEN scError(%eFail, "VSAVE of File " + TIP.FilePath + " failed")' Tell User
   END IF                                                         '
   me.FindPushPop("POP")                                          ' Restore FIND stuff

   TopScrn = lclTop                                               ' Keep our position
   MExit
END METHOD

METHOD  pCmdWDIR(pCmd AS STRING)
'---------- Open a Windows Explorer to current file's DIR
LOCAL lclFn AS STRING
   MEntry
   IF IsMEdit THEN _                                              ' For MEdit, make sure all paths are the same
      MErrExit(%eFail, "Not supported in MEdit mode")             '
   IF IsSetEdit OR IsClip THEN _                                  ' Ineligible tab?
      MErrExit(%eFail, "This tab has no implied DIR folder")      '
   lclFn = TIP.FilePath                                           ' Get the filename
   SHELL "EXPLORER.EXE /SELECT, " + $DQ + lclFn + $DQ, 1          '
   MExit                                                          '
END METHOD

METHOD  pCmdXSUBMIT(pCmd AS STRING)
'---------- Do the XSUBMIT process
LOCAL Fn, sline, t, lclPath, lclDrive, lclSubArg AS STRING
LOCAL i, j, lctr, recs AS LONG
LOCAL DCmd, JobName, W1, W2, W3, w4 AS STRING
LOCAL hWatchThread AS LONG
LOCAL fptr AS STRING POINTER
DIM  slines(1 TO 1000) AS STRING
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN _                                       ' Better be a filename operand
      MErrExit(%eFail, "Mising Submit filename")                  '

   TOP.Setup("E", "", "", pCmdOps(1))                             ' Setup TOP
   IF TOP.EXEC THEN _                                             ' Go validate
      MErrExit(%eFail, TOP.ResultMsg)                             ' Oops?  Bail out (File mmissing)
   gSubmitFile = TOP.FilePath                                     ' Set in Global area

   '----- Setup Profile name to get SUBARG
   Prf.SetProfName(TOP.Profile, %False)                           ' Swap profiles
   Prf.ReadAll(%True)                                             '
   lclSubArg = Prf.SubArg                                         ' Save the SubArg
   Prf.SetProfName(TIP.Profile, %False)                           ' Swap profiles back
   Prf.ReadAll(%True)                                             '

   '----- Get next jobnumber
   ENV.Jobnumber = ENV.Jobnumber + 1                              ' Bump jobnumber
   gJobID = "JOB" + FORMAT$(ENV.Jobnumber, "00000")               ' Build ID

   '----- Create the Results file
   IF ISNOTNULL(JobName) THEN                                     ' Base on Jobname
      gResultFile = ENV.SubmitDir + "\SUBRESULTS_" + JobName + "_" + gJobID + ".txt"
   ELSE                                                           '
      gResultFile = ENV.SubmitDir + "\SUBRESULTS_" + gJobID + ".txt"
   END IF                                                         '
   sMakeNullFile(gResultFile)                                     ' Create null file
   gSubmitFile = $DQ + gSubmitFile + $DQ                          ' Add the quotes
   gResultFile = $DQ + gResultFile + $DQ                          '

   '----- Issue the SUBMIT command
   DCmd = ENV.SubmitCmd                                           ' Get the SUBMIT command
   gSubmitType = "J"                                              ' Set Submit type
   IF me.MacSubst(DCmd) THEN                                      ' Do substitution
      sDoMsgBox("Malformed macro operand in SUBMIT prototype command", %MB_OK OR %MB_USERICON, "Submit"): MExitMeth
   END IF                                                         '
'   debug DCmd
   lclPath = CURDIR$                                              ' Locate where we are
   IF MID$(lclPath, 2, 1) = ":" THEN _                            ' Extract Drive if present
      lclDrive = LEFT$(lclPath, 2)                                ' and save it

   IF MID$(ENV.SubmitDir, 2, 1) = ":" THEN _                      ' See if Submit Path has drive
      CHDRIVE LEFT$(ENV.SubmitDir, 2)                             ' If so, go to it
   CHDIR ENV.SubmitDir                                            ' Switch to EXE path
   gLoopCtr = -1                                                  ' Suspend loop detection
   j = SHELL(ENVIRON$("COMSPEC") + " /c " + $DQ + DCmd + $DQ, 0)  '
   RESET gLoopCtr                                                 ' Reset loop detection
   IF ERR THEN                                                    ' Tell user result
      sDoMsgBox("Error issuing XSUBMIT: |K" + ERROR$, %MB_OK OR %MB_USERICON, "XSubmit")
      MExitMeth                                                   '
   END IF                                                         '
   IF ISNOTNULL(lclDrive) THEN CHDRIVE lclDrive                   ' Switch drive if needed
   CHDIR lclPath                                                  ' put back the original path

   '----- Start FileWatch for the Result file
   sUnquote gResultFile                                           ' remove Quotes again
   fptr = VARPTR(gResultFile)                                     '
   THREAD CREATE sResultFileWatchThread(BYVAL fptr) 65536, TO hWatchThread ' Fire up the thread
   IF hWatchThread = 0 THEN _                                     ' Failed?
      sDoMsgBox("Watch of Result file could not be started", %MB_OK OR %MB_USERICON, "XSubmit"): MExitMeth
   SLEEP 50                                                       ' Wait a bit
   THREAD STATUS hWatchThread TO j                                ' See if running OK
   IF j <> 259 THEN _                                             ' If running OK STATUS returns &H103 (See Help)
      sDoMsgBox("Watch of Result file could not be started", %MB_OK OR %MB_USERICON, "XSubmit"): MExitMeth
   THREAD CLOSE hWatchThread TO j                                 ' Free up our handle
   sDoMsgBox("|K" + gSubmitFile + "|B submitted as JobNumber |K" + FORMAT$(ENV.Jobnumber, "00000"), %MB_OK OR %MB_USERICON, "XSubmit")
   MExit
END METHOD

METHOD  pCmdXTABS(pCmd AS STRING)
'---------- Alter the XTabs setting
   MEntry
   IF me.CmdParse(pCmd) THEN MExitMeth                            ' Do basic parsing, exit if errors
   IF pCmdNumOps = 0 THEN _                                       ' If no operands
      MErrExit(%eNone, "XTABS set to " + FORMAT$(Prf.ImportTabs)) '
   IF pCmdOpsType(1) <> %OpNum OR _                               '
      VAL(pCmdOps(1)) < 0 OR VAL(pCmdOps(1)) > 16 THEN _          '
      MErrExit(%eFail, "Invalid XTABS operand")                   '
   Prf.ImportTabs = VAL(pCmdOps(1))                               ' Set it
   scError(%eNone, "XTABS set to " + FORMAT$(Prf.ImportTabs))     '
   MExit
END METHOD
