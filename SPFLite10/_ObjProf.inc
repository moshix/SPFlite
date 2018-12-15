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

CLASS cProf

   '----- File Profile Settings
   INSTANCE ActionSave  AS LONG                                   ' Count between automatic SAVEs
   INSTANCE ActionVerb  AS STRING                                 ' Action verb for the SAVE
   INSTANCE AutoBkup    AS LONG                                   ' Create .BKP files (False = no, True = Yes)
   INSTANCE AUTOCAPS    AS LONG                                   ' AutoCaps On/Off
   INSTANCE AutoNum     AS LONG                                   ' AutoNum On/Off
   INSTANCE AutoSave    AS LONG                                   ' Save on END / =X ? (0 = no, 1 = Yes, 2 = Prompt)
   INSTANCE BndLeft     AS LONG                                   ' Left Bound
   INSTANCE BndRight    AS LONG                                   ' Right Bound (0 = >> no limit)
   INSTANCE BndText     AS STRING                                 ' String representation of BNDS
   INSTANCE CA2S        AS STRING                                 ' COLLATE Translate table ANSI to Source
   INSTANCE CapsActual  AS LONG                                   ' CAPS actual flag
   INSTANCE CapsDesired AS LONG                                   ' CAPS desired flag
   INSTANCE ChangeMode  AS STRING                                 ' Default Change shift mode
   INSTANCE ChrUpper    AS STRING                                 ' Uppercase string
   INSTANCE ChrLower    AS STRING                                 ' Lowercase string
   INSTANCE ChrRegLower AS STRING                                 ' Lowercase string (RegEx)
   INSTANCE ChrRegUpper AS STRING                                 ' Uppercase string (RegEx)
   INSTANCE ColateXlate AS LONG                                   ' COLLATE XLate is active
   INSTANCE Cols        AS LONG                                   ' Ruler Line DesiLAttrG
   INSTANCE CS2A        AS STRING                                 '    "         "      "   SOURCE to ANSI
   INSTANCE EOL         AS STRING                                 ' EOL Default action
   INSTANCE EOLIX       AS LONG                                   ' EOL Select Index
   INSTANCE Fold        AS LONG                                   ' Fold status 0=Off, 1=ON
   INSTANCE GoodBnds    AS STRING                                 ' Last 'good' BNDS string
   INSTANCE HexMode     AS LONG                                   ' Hex Mode 1 = off, 4 = on
   INSTANCE HiAuto      AS LONG                                   ' HiAuto flag
   INSTANCE HiFind      AS LONG                                   ' HiFind flag
   INSTANCE ImportTabs  AS LONG                                   ' Import Tab conversion
   INSTANCE INIFileName AS STRING                                 ' INI file
   INSTANCE INIFileNmZ  AS ASCIIZ * %MAX_PATH                     ' INI file in ASCIIZ format
   INSTANCE LRECL       AS LONG                                   ' LRECL
   INSTANCE MarkLine    AS STRING                                 ' Mark columns line
   INSTANCE MarkWorking AS STRING                                 ' Mark columns line (working)
   INSTANCE MaskLine    AS STRING                                 ' Mask model line
   INSTANCE MINLEN      AS LONG                                   ' MINLEN
   INSTANCE NumType     AS STRING                                 ' NUMTYPE - null or a valid string
   INSTANCE PageFlag    AS LONG                                   ' PAGE ON/OFF
   INSTANCE PageOffset  AS LONG                                   ' PAGE ON offset
   INSTANCE PCase       AS STRING                                 ' Default text case
   INSTANCE PCollate    AS STRING                                 ' Collate option
   INSTANCE PendingUse  AS STRING                                 ' Pending USE name
   INSTANCE PLock       AS LONG                                   ' Profile Lock
   INSTANCE PMark       AS LONG                                   ' MARK lines are active
   INSTANCE PPreserve   AS LONG                                   ' Keep trailing blanks on save? (True = Yes, False = No)
   INSTANCE ProfName    AS STRING                                 ' Profile Name (e.g SPFLite\Profile\xxxxx.INI)
   INSTANCE PSource     AS STRING                                 ' DesiLAttrG encoding ANSI, UTF8 etc.
   INSTANCE PState      AS LONG                                   ' State status 0=Off, 1=ON, 2=FEW, 3=MANY
   INSTANCE PUsing      AS STRING                                 ' Profile USING value
   INSTANCE PWord       AS STRING                                 ' Valid 'word' characters
   INSTANCE RECFM       AS STRING                                 ' RECFM
   INSTANCE SA2S        AS STRING                                 ' SOURCE Translate table ANSI to Source
   INSTANCE ScrlPageSus AS LONG                                   ' Scroll Amount PAGE suspended
   INSTANCE SCROLL      AS STRING * 4                             ' Scroll Amount (Profile)
   INSTANCE SKIP        AS LONG                                   ' ProfRealAll sets if missing Profile and SKIP chosen
   INSTANCE SrceXlate   AS LONG                                   ' SOURCE XLate is active
   INSTANCE SS2A        AS STRING                                 '   "         "      "   SOURCE to ANSI
   INSTANCE Start       AS STRING                                 ' START setting
   INSTANCE SubArg      AS STRING                                 ' SUBARG data
   INSTANCE SubCmd      AS STRING                                 ' SUBCmd data
   INSTANCE Tabs        AS LONG                                   ' Tabs line in use
   INSTANCE TabsLine    AS STRING                                 ' Active Tabs line
   INSTANCE UndoNumber  AS LONG                                   ' No. of UNDO levels
   INSTANCE UseMode     AS LONG                                   ' Profile is in USING mode
   INSTANCE UseName     AS STRING                                 ' USE Name
   INSTANCE WordInput   AS STRING                                 ' Valid 'word' characters (INI version)

   CLASS METHOD CREATE()                                          ' Constructor - Initialize Class stuff
   LOCAL i AS LONG

   '----- initialize the Profile entries
      ActionSave  = 0                                             ' ACTION is OFF
      ActionVerb  = "VSAVE"                                       ' ACTION verb
      AutoBkup    = %False                                        ' AUTOBKUP is OFF
      AUTOCAPS    = %False                                        ' AUTOCAPS is OFF
      AutoNum     = %False                                        ' AUTONum is OFF
      AutoSave    = 2                                             ' AUTOSAVE is OFF,PROMPT
      BndLeft     = 1                                             ' Left Bound
      BndRight    = 0                                             ' Right Bound
      BndText     = ""                                            ' Text representation of BNDS
      CapsActual  = %False                                        ' CAPS actual is OFF
      CapsDesired = %False                                        ' CAPS desired is OFF
      ChangeMode  = "D"                                           ' Default change shift mode
      ChrUpper    = ""                                            ' Uppercase string
      ChrLower    = ""                                            ' Lowercase string
      ChrRegLower = ""                                            ' Lowercase string (RegEx)
      ChrRegUpper = ""                                            ' Uppercase string (RegEx)
      ColateXlate = %False                                        ' COLLATE Xlate is active
      Cols        = %False                                        ' RULER is OFF
      EOL         = "CRLF"                                        ' EOL is CRLF
      EOLIX       = 1                                             ' EOL Index is 1
      Fold        = %False                                        ' FOLD OFF
      GoodBnds    = ""                                            ' Last good BNDS
      HexMode     = 1                                             ' Start in non-Hex mode
      HiAuto      = %True                                         ' HiAuto Flag
      HiFind      = %True                                         ' HiFind Flag
      INIFileName = ENV.PROFPath + "DEFAULT.INI"                  ' Build the File Profile default INI name
      INIFileNmz  = INIFileName                                   ' Build the ASCIIZ version
      ImportTabs  = 4                                             ' XTABS is 4
      LRECL       = 0                                             ' LRECL
      MarkLine    = ""                                            ' Mark line
      MarkWorking = ""                                            ' Mark line (Working)
      MaskLine    = ""                                            ' Mask line
      MINLEN      = 0                                             ' MINLEN
      NumType     = ""                                            ' NUMTYPE - null or a valid string
      PageFlag    = %False                                        ' PAGE ON/OFF
      PageOffset  = 0                                             ' PAGE ON offset
      PCase       = "T"                                           ' Default text case
      PCollate    = "ANSI"                                        ' Collate string
      PendingUse  = ""                                            ' Pending USE name
      PLock       = %False                                        ' Profile is UNLOCKed
      PMark       = %True                                         ' MARK is ON
      PPreserve   = %False                                        ' Preserve
      ProfName    = "Default"                                     ' Profile name
      PPreserve   = %False                                        ' PRESERVE is OFF
      PSource     = "ANSI"                                        ' Default encoding ANSI
      PState      = %StateOff                                     ' State OFF
      PWord       = ""                                            ' Default WORD values
      RECFM       = "U"                                           ' RECFM
      ScrlPageSus = %False                                        ' Scroll Amount PAGE suspended
      SCROLL      = "CSR "                                        ' Scroll is CSR
      Skip        = %False                                        '
      SrceXlate   = %False                                        ' SOURCE Xlate is active
      Start       = "FIRST"                                       ' START value
      SubArg      = ""                                            ' SUBARG data
      SubCmd      = ""                                            ' SUBCmd data
      Tabs        = %True                                         ' TABS is ON
      TabsLine    = ""                                            ' Active Tabs line
      UndoNumber  = 10                                            ' Number of UNDO levels
      UseMode     = %False                                        ' Profile is not in USING mode
      UseName     = ""                                            ' USE Name
      WordInput   = $WORD                                         ' Default INI value
      RESET SA2S, SS2A, CA2S, CS2A                                ' Reset strings before building
      FOR i = 0 TO 255                                            ' SOURCE ANSI to SOURCE table
         SA2S += CHR$(i)                                          '
         SS2A += CHR$(i)                                          '
         CA2S += CHR$(i)                                          '
         CS2A += CHR$(i)                                          '
      NEXT                                                        '

   END METHOD                                                     '

   CLASS METHOD BuildSourceXlate()
   REGISTER i AS LONG
   LOCAL S2A, A2S, t AS STRING

      t = SPACE$(256)                                             ' Make empty string
      FOR i = 0 TO 255                                            ' SOURCE ANSI to SOURCE table
         MID$(t, i + 1, 1) = CHR$(i)                              '
      NEXT                                                        '
      SA2S = t: SS2A = t: SrceXlate = %False                      ' Say no Source Xlate active

      t = SPACE$(256)                                             ' Make empty string
      FOR i = 0 TO 255                                            ' COLLATE ANSI to SOURCE table
         MID$(t, i + 1, 1) = CHR$(i)                              '
      NEXT                                                        '
      CA2S = t: CS2A = t: ColateXlate = %False                    ' Say no Collate Xlate active

      IF PSource <> "ANSI" AND _                                  ' Is SOURCE non-ANSI mode?
         PSource <> "UTF8" AND _                                  '
         PSource <> "UTF16" AND _                                 '
         PSource <> "UTF16LE" AND _                               '
         PSource <> "UTF16BE" THEN                                '
         IF sSourceLoad(PSource, A2S, S2A) THEN                   ' Load table, if successful
            SA2S = A2S: SS2A = S2A: SrceXlate = %True             ' Say source Xlate active
         END IF                                                   '
      END IF

      IF PCollate <> "ANSI" AND _                                 ' Is SOURCE non-ANSI mode?
         PCollate <> "UTF8" AND _                                 '
         PCollate <> "UTF16" AND _                                '
         PCollate <> "UTF16LE" AND _                              '
         PCollate <> "UTF16BE" THEN                               '

         ' if a SOURCE codepage failed validation, and we are asking for the same name as a COLLATE name
         ' we will just get the same error again.  if so, don't bother to load the COLLATE tables
         ' we only make the attempt to load if this is either a different name, or it's the same name
         '/ the SOURCE load above was successful

         IF (ISTRUE SrceXlate) OR (PSource <> PCollate) THEN
            IF sSourceLoad(PCollate, A2S, S2A) THEN               ' Load table, if successful
               CA2S = A2S: CS2A = S2A: ColateXlate = %True        ' Say source Xlate active
            END IF                                                '
         END IF                                                   '
      END IF                                                      '
   END METHOD

   INTERFACE iProf: INHERIT IUNKNOWN                              ' Define the interface

      METHOD BuildBndText()
         IF BndRight > 0 THEN                                     ' We have a right bound
            BndText = RSET$(">", BndRight)                        ' Build string BndRight long
            MID$(BndText, BndLeft, 1) = "<"                       ' Stuff in BndLeft
         ELSE                                                     ' Else build a MAX format line
            BndText = RSET$("+", gDataLen + TP.Offset)            ' Build string to put + at the right side
            MID$(BndText, BndLeft, 1) = "<"                       ' Stuff in BndLeft
         END IF                                                   '
         GoodBnds = BndText                                       ' Save good copy
         TP.BndText = BndText                                     ' Save to the working copy
      END METHOD


      GProp(ActionSave, LONG)
      PROPERTY SET ActionSave(v AS LONG)
         ActionSave = v: IF ISFALSE PLock THEN me.SetString("Action", FORMAT$(ActionSave))
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(ActionVerb, STRING)
      PROPERTY SET ActionVerb(v AS STRING)
         ActionVerb = v: IF ISFALSE PLock THEN me.SetString("ActionVerb", ActionVerb)
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(AutoBkup, LONG)
      PROPERTY SET AutoBkup(v AS LONG)
         AutoBkup = v: IF ISFALSE PLock THEN me.SetString("AutoBkup", FORMAT$(AutoBkup))
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(AUTOCAPS, LONG)
      PROPERTY SET AUTOCAPS(v AS LONG)
         AUTOCAPS = v: IF ISFALSE PLock THEN me.SetString("AutoCaps", FORMAT$(AUTOCAPS))
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(AutoNum, LONG)
      PROPERTY SET AutoNum(v AS LONG)
         AutoNum = v: IF ISFALSE PLock THEN me.SetString("AutoNum", FORMAT$(AutoNum))
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(AutoSave, LONG)
      PROPERTY SET AutoSave(v AS LONG)
         AutoSave = v: IF ISFALSE PLock THEN me.SetString("AutoSave", FORMAT$(AutoSave))
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(BndLeft, LONG)
      PROPERTY SET BndLeft(v AS LONG)
         BndLeft = v: IF ISFALSE PLock THEN me.SetString("BndL", FORMAT$(BndLeft))
         me.BuildBndText
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(BndRight, LONG)
      PROPERTY SET BndRight(v AS LONG)
         BndRight = v: IF ISFALSE PLock THEN me.SetString("BndR", FORMAT$(BndRight))
         me.BuildBndText
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(BndText, STRING)
      GProp(CA2S, STRING)

      GProp(CapsActual, LONG)
      PROPERTY SET CapsActual(v AS LONG)
         CapsActual = v
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(CapsDesired, LONG)
      PROPERTY SET CapsDesired(v AS LONG)
         CapsDesired = v: IF ISFALSE PLock THEN me.SetString("CapsFlag", FORMAT$(CapsDesired))
         CapsActual = CapsDesired
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(ChangeMode, STRING)
      PROPERTY SET ChangeMode(v AS STRING)
         ChangeMode = v: IF ISFALSE PLock THEN me.SetString("Change", ChangeMode)
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(ChrUpper, STRING)
      GProp(ChrLower, STRING)
      GProp(ChrRegLower, STRING)
      GProp(ChrRegUpper, STRING)

      GProp(PCollate, STRING)
      PROPERTY SET PCollate(v AS STRING)
         PCollate = v: IF ISFALSE PLock THEN me.SetString("Collate", PCollate)
         me.BuildSourceXlate()
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(ColateXlate, LONG)

      GProp(Cols, LONG)
      PROPERTY SET Cols(v AS LONG)
         Cols = v: IF ISFALSE PLock THEN me.SetString("ColsFlag", FORMAT$(Cols))
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(CS2A, STRING)

      GProp(EOL, STRING)
      PROPERTY SET EOL(v AS STRING)
         EOL = v: IF ISFALSE PLock THEN me.SetString("EOLFlag", EOL)
         EOLIX = SWITCH(EOL = "CRLF", 1, EOL = "LF", 2, EOL = "CR", 3, EOL = "NL", 4, EOL = "AUTO", 5, EOL = "AUTONL", 6, EOL = "NONE", 7)
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(EOLIX, LONG)

      GProp(Fold, LONG)
      PROPERTY SET Fold(v AS LONG)
         Fold = v: IF ISFALSE PLock THEN me.SetString("FoldFlag", FORMAT$(Fold))
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(GoodBnds, STRING)

      GProp(HexMode, LONG)
      PROPERTY SET HexMode(v AS LONG)
         HexMode = v: IF ISFALSE PLock THEN me.SetString("HexMFlag", FORMAT$(HexMode))
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(HiAuto, LONG)
      PROPERTY SET HiAuto(v AS LONG)
         HiAuto = v: IF ISFALSE PLock THEN me.SetString("HiAuto", FORMAT$(HiAuto))
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(HiFind, LONG)
      PROPERTY SET HiFind(v AS LONG)
         HiFind = v: IF ISFALSE PLock THEN me.SetString("HiFind",   FORMAT$(HiFind))
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(ImportTabs, LONG)
      PROPERTY SET ImportTabs(v AS LONG)
         ImportTabs = v: IF ISFALSE PLock THEN me.SetString("ImportTabs", FORMAT$(ImportTabs))
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(INIFileName, STRING)

      GProp(LRECL, LONG)
      PROPERTY SET LRECL(v AS LONG)
         LRECL = v: IF ISFALSE PLock THEN me.SetString("LRECL", FORMAT$(LRECL))
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(MarkLine, STRING)
      PROPERTY SET MarkLine(v AS STRING)
         MarkLine = v: IF ISFALSE PLock THEN me.SetString("MARK", $DQ + RTRIM$(MarkLine) + $DQ)
         MarkWorking = sMarkSimple(MarkLine)
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(MarkWorking, STRING)

      GProp(MaskLine, STRING)
      PROPERTY SET MaskLine(v AS STRING)
         MaskLine = v: IF ISFALSE PLock THEN me.SetString("MASK", $DQ + RTRIM$(MaskLine) + $DQ)
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(MINLEN, LONG)
      PROPERTY SET MINLEN(v AS LONG)
         MINLEN = v: IF ISFALSE PLock THEN me.SetString("MINLEN", FORMAT$(MINLEN))
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(NumType, STRING)
      PROPERTY SET NumType(v AS STRING)
         NumType = v: IF ISFALSE PLock THEN me.SetString("NUMTYPE", NumType)
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(PageFlag, LONG)
      PROPERTY SET PageFlag(v AS LONG)
         PageFlag = v: IF ISFALSE PLock THEN me.SetString("Page", FORMAT$(PageFlag))
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(PageOffset, LONG)
      PROPERTY SET PageOffset(v AS LONG)
         PageOffset = v: IF ISFALSE PLock THEN me.SetString("PageOffset", FORMAT$(PageOffset))
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(PCase, STRING)
      PROPERTY SET PCase(v AS STRING)
         PCase = v: IF ISFALSE PLock THEN me.SetString("Case", PCase)
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(PLock, LONG)
      PROPERTY SET PLock(v AS LONG)
         PLock = v: me.SetString("ProfLock", FORMAT$(PLock))
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(PMark, LONG)
      PROPERTY SET PMark(v AS LONG)
         PMark = v: IF ISFALSE PLock THEN me.SetString("MarkFlag", FORMAT$(PMark))
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(PPreserve, LONG)
      PROPERTY SET PPreserve(v AS LONG)
         PPreserve = v: IF ISFALSE PLock THEN me.SetString("Preserve", FORMAT$(PPreserve))
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(ProfName, STRING)

      GProp(PSource, STRING)
      PROPERTY SET PSource(v AS STRING)
         PSource = v: IF ISFALSE PLock THEN me.SetString("Source", PSource)
         me.BuildSourceXlate()
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(PState, LONG)
      PROPERTY SET PState(v AS LONG)
         PState = v: IF ISFALSE PLock THEN me.SetString("StateFlag", FORMAT$(PState))
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(PWord, STRING)

      GProp(RECFM, STRING)
      PROPERTY SET RECFM(v AS STRING)
         RECFM = v: IF ISFALSE PLock THEN me.SetString("RECFM", RECFM)
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(SA2S, STRING)

      GSProp(ScrlPageSus, LONG)

      GProp(SCROLL, STRING)
      PROPERTY SET SCROLL(v AS STRING)
         SCROLL = v: IF ISFALSE PLock THEN me.SetString("ScrollAmount", TRIM$(SCROLL))
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GSProp(SKIP, LONG)
      GSProp(SrceXlate, LONG)
      GProp(SS2A, STRING)

      GProp(Start, STRING)
      PROPERTY SET Start(v AS STRING)
         Start = v: IF ISFALSE PLock THEN me.SetString("Start", Start)
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(SubArg, STRING)
      PROPERTY SET SubArg(v AS STRING)
         SubArg = v: IF ISFALSE PLock THEN me.SetString("SubArg", SubArg)
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(SubCmd, STRING)
      PROPERTY SET SubCmd(v AS STRING)
         SubCmd = v: IF ISFALSE PLock THEN me.SetString("SubCmd", SubCmd)
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(Tabs, LONG)
      PROPERTY SET Tabs(v AS LONG)
         Tabs = v: IF ISFALSE PLock THEN me.SetString("TabsFlag", FORMAT$(Tabs))
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(TabsLine, STRING)
      PROPERTY SET TabsLine(v AS STRING)
         TabsLine = v: IF ISFALSE PLock THEN me.SetString("TABS", $DQ + RTRIM$(TabsLine) + $DQ)
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(UndoNumber, LONG)
      PROPERTY SET UndoNumber(v AS LONG)
         UndoNumber = v: IF ISFALSE PLock THEN me.SetString("UndoLevels", FORMAT$(UndoNumber))
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(UseMode, LONG)

      GProp(UseName, STRING)
      PROPERTY SET UseName(v AS STRING)
         IF Usename = "" AND v <> "" THEN                         ' Switching to a USE mode
            UseName = v                                           ' Save it
            IF ISFALSE PLock THEN me.SetString("ProfUsing", UseName) ' Save USE name in the existing Profile
         ELSEIF UseName = v  THEN                                 ' No change?
            '                                                     ' Do nothing
         ELSEIF UseName <> "" AND v = "" THEN                     ' Turning OFF USE?
            UseName = v                                           ' Save it
            INIFileName = ENV.PROFPath + ProfName + ".INI"        ' Setup the original INI filename
            INIFileNmZ = INIFileName                              ' Setup the ASCIIZ format
            IF ISFALSE PLock THEN me.SetString("ProfUsing", UseName) ' Save USE name in the original Profile
         END IF
         me.ReadAll(%True)                                        ' Reload stuff
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      GProp(WordInput, STRING)
      PROPERTY SET WordInput(v AS STRING)
         WordInput = v
         IF me.WordVal THEN EXIT PROPERTY
         IF ISFALSE PLock THEN me.SetString("WORD", WordInput)
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
      END PROPERTY

      METHOD RESET()
      LOCAL i AS LONG
      '---------- Reset all Profile variables to standard set
         ActionSave  = 0                                             ' ACTION is OFF
         ActionVerb  = "VSAVE"                                       ' ACTION verb
         AutoBkup    = %False                                        ' AUTOBKUP is OFF
         AUTOCAPS    = %False                                        ' AUTOCAPS is OFF
         AutoNum     = %False                                        ' AUTONum is OFF
         AutoSave    = 2                                             ' AUTOSAVE is OFF,PROMPT
         BndLeft     = 1                                             ' Left Bound
         BndRight    = 0                                             ' Right Bound
         BndText     = ""                                            ' Text representation of BNDS
         CapsActual  = %False                                        ' CAPS actual is OFF
         CapsDesired = %False                                        ' CAPS desired is OFF
         ChangeMode  = "D"                                           ' Default change shift mode
         ChrUpper    = ""                                            ' Uppercase string
         ChrLower    = ""                                            ' Lowercase string
         ChrRegLower = ""                                            ' Lowercase string (RegEx)
         ChrRegUpper = ""                                            ' Uppercase string (RegEx)
         ColateXlate = %False                                        ' COLLATE Xlate is active
         Cols        = %False                                        ' RULER is OFF
         EOL         = "CRLF"                                        ' EOL is CRLF
         EOLIX       = 1                                             ' EOL Index is 1
         Fold        = %False                                        ' FOLD OFF
         GoodBnds    = ""                                            ' Last good BNDS
         HexMode     = 1                                             ' Start in non-Hex mode
         HiAuto      = %True                                         ' HiAuto Flag
         HiFind      = %True                                         ' HiFind Flag
         ImportTabs  = 4                                             ' XTABS is 4
         LRECL       = 0                                             ' LRECL
         MarkLine    = ""                                            ' Mark line
         MarkWorking = ""                                            ' Mark line (Working)
         MaskLine    = ""                                            ' Mask line
         MINLEN      = 0                                             ' MINLEN
         NumType     = ""                                            ' NumType
         PageFlag    = %False                                        ' PAGE ON/OFF
         PageOffset  = 0                                             ' PAGE ON offset
         PCase       = "T"                                           ' Default text case
         PCollate    = "ANSI"                                        ' Collate string
         PendingUse  = ""                                            ' Pending USE name
         PLock       = %False                                        ' Profile is UNLOCKed
         PMark       = %True                                         ' MARK is ON
         PPreserve   = %False                                        ' Preserve
         PPreserve   = %False                                        ' PRESERVE is OFF
         PSource     = "ANSI"                                        ' Default encoding ANSI
         PState      = %StateOff                                     ' State OFF
         PWord       = ""                                            ' Default WORD values
         RECFM       = "U"                                           ' RECFM
         ScrlPageSus = %False                                        ' Scroll Amount PAGE suspended
         SCROLL      = "CSR "                                        ' Scroll is CSR
         Skip        = %False                                        '
         SrceXlate   = %False                                        ' SOURCE Xlate is active
         Start       = "FIRST"                                       ' START value
         SubArg      = ""                                            ' SUBARG data
         SubCmd      = ""                                            ' SUBCmd data
         Tabs        = %True                                         ' TABS is ON
         TabsLine    = ""                                            ' Active Tabs line
         UndoNumber  = 10                                            ' Number of UNDO levels
         UseMode     = %False                                        ' Profile is not in USING mode
         UseName     = ""                                            ' USE Name
         WordInput   = $WORD                                         ' Default INI value
         RESET SA2S, SS2A, CA2S, CS2A                                ' Reset strings before building
         FOR i = 0 TO 255                                            ' SOURCE ANSI to SOURCE table
            SA2S += CHR$(i)                                          '
            SS2A += CHR$(i)                                          '
            CA2S += CHR$(i)                                          '
            CS2A += CHR$(i)                                          '
         NEXT                                                        '
         me.BuildLines                                               ' Rebuild PROF1/2/3/4/5 lines
      END METHOD

      METHOD SetProfName(pname AS STRING, DefaultNew AS LONG)
      '---------- Get a valid Profile Name set up
      LOCAL t AS STRING
         MEntry
         Skip = %False                                            ' Say we're not skipping the filetype
         IF ISFILE(ENV.PROFPath + pname + ".INI") THEN            ' If it exists, then we use it
            ProfName = UUCASE(pname)                              ' Setup for this profile then
            INIFileName = ENV.PROFPath + ProfName + ".INI"        ' Setup the INI filename
            INIFileNmZ = INIFileName                              ' Setup the ASCIIZ format

         ELSE                                                     '
            IF DefaultNew THEN                                    ' Force create NEW?
               FILECOPY ENV.PROFPath + "DEFAULT.INI", ENV.PROFPath + pname + ".INI" ' Copy the DEFAULT INI file
               ProfName = UUCASE(pname)                           ' Setup for this profile then
               INIFileName = ENV.PROFPath + ProfName + ".INI"     ' Setup the INI filename
               INIFileNmZ = INIFileName                           ' Setup the ASCIIZ format
               ENV.DefaultShr = UUCASE(ENV.DefaultShr)            ' Uppercase the list
               ENV.DefaultShr = REMOVE$(ENV.DefaultShr, UUCASE(pname) + ",")' Remove this Profname if present
               sIniSetString("General", "DefaultShr", ENV.DefaultShr) ' Save it back

            ELSE

               IF INSTR(UUCASE("," + ENV.DefaultShr), UUCASE("," + pname + ",")) THEN  ' Is it already in the Default list?
                  ProfName = "DEFAULT"                            ' Swap in the DEFAULT
                  INIFileName = ENV.PROFPath + "DEFAULT.INI"      ' Setup the INI filename
                  INIFileNmZ = INIFileName                        ' Setup the ASCIIZ format
               ELSE                                               '
                  t = DispDefault(TRIM$(Pname))                   ' Try the display, see what user wants to use
                  IF t = "NEW" THEN                               '
                     FILECOPY ENV.PROFPath + "DEFAULT.INI", ENV.PROFPath + pname + ".INI" ' Copy the DEFAULT INI file
                     ProfName = UUCASE(pname)                     ' Setup for this profile then
                     INIFileName = ENV.PROFPath + ProfName + ".INI"  ' Setup the INI filename
                     INIFileNmZ = INIFileName                     ' Setup the ASCIIZ format
                     MExitMeth                                    '

                  ELSEIF t = "DEF" THEN                           ' Use the DEFAULT
                     ENV.DefaultShr = ENV.DefaultShr + pname + ","   ' Add to the Default string
                     sIniSetString("General", "DefaultShr", ENV.DefaultShr) ' Save it
                     ProfName = "DEFAULT"                         ' Swap in the DEFAULT
                     INIFileName = ENV.PROFPath + "DEFAULT.INI"   ' Build the Using File Profile INI name
                     INIFileNmZ = INIFileName                     ' Setup the ASCIIZ format

                  ELSEIF LEFT$(t, 4) = "USE " THEN                ' Make a USE profile?
                     FILECOPY ENV.PROFPath + "DEFAULT.INI", ENV.PROFPath + pname + ".INI" ' Copy the DEFAULT INI file
                     ProfName = UUCASE(pname)                     ' Setup for this profile then
                     INIFileName = ENV.PROFPath + ProfName + ".INI"  ' Setup the INI filename
                     INIFileNmZ = INIFileName                     ' Setup the ASCIIZ format
                     PendingUse = MID$(t, 5)                      ' Save the Pending USE name
                     MExitMeth                                    ' Done here

                  ELSEIF t = "SKIP" THEN                          ' Skip this file type
                     ENV.FMFFList = ENV.FMFFList + "," + pname    ' Add to the FF exempt list
                     sIniSetString("FManager", "FMFFList", ENV.FMFFList)' Save it
                     Skip = %True                                 ' Say we're skipping this filetype

                  ELSEIF t = "ABORT" THEN                         ' Abort the FF
                     Skip = %True                                 ' Say we're skipping this filetype
                     DIALOG END hIntr                             ' End the Interrupt window
                     gfInterrupt = %True                          ' Flag for mainline

                  END IF                                          '
               END IF                                             '
            END IF                                                '
         END IF                                                   '
         MExit                                                    '
      END METHOD

      METHOD ReadAll(UseOK AS LONG)
      '---------- Get the defaults for the File Profile INI values
      LOCAL t, lclProfName, lclProfINI AS STRING, i AS LONG
         MEntry
         UseMode = %False                                         ' Start in non-USING mode
         UseName = me.GetString("ProfUsing", "")                  ' Get any Using value or leave as null
         IF UseOK AND UseName = "" AND PendingUse <> "" THEN      ' Is this the newly created USE types
            PendingUse = LEFT$(PendingUse, LEN(PendingUse) - 4)   ' Strip off .INI to normalize it
            me.SetString("ProfUsing", PendingUse)                 ' Yes, set in the USE name
            UseName = PendingUse                                  ' And swap it in
            PendingUse = ""                                       ' Clear it now
         END IF                                                   '
         IF UseOK AND ISNOTNULL(UseName) THEN                     ' Are we using another Profile?
            INIFileName = ENV.PROFPath + UseName + ".INI"         ' Re-Build the Using File Profile INI name
            INIFileNmz  = INIFileName                             ' Build the ASCIIZ version
            IF ISFALSE ISFILE(INIFileName) THEN                   ' Better exist
               sDoMsgBox "The |K" + UUCASE(ProfName) + "|B profile specifies USE |K" + UseName + $CRLF + _
               "|Bbut the |K" + UseName + "|B profile is missing. Ignoring the |KUSE|B request.", _
               %MB_OK OR %MB_USERICON, "SPFLite"
               INIFileName = ENV.PROFPath + ProfName + ".INI"     ' Put back the original Profile INI name
               INIFileNmZ = INIFileName                           ' Setup the ASCIIZ format
            ELSE                                                  ' We have the other INI file
               UseMode = %True                                    ' Remember we've loaded a USING version of the Profile
            END IF                                                '
         END IF                                                   ' Remaining INI gets will use the new filename

         me.Touch                                                 ' Go Touch the file's timestamp
         PLock       = VAL(me.GetString("ProfLock", "0"))         '
         SCROLL      = me.GetString("ScrollAmount", "CSR ")       '
         ScrlPageSus = %False                                     '
         ActionSave  = VAL(me.GetString("Action", "0"))           '
         ActionVerb  = me.GetString("ActionVerb", "VSAVE")        '
         AutoBkup    = VAL(me.GetString("AutoBkup", "0"))         '
         AUTOCAPS    = VAL(me.GetString("AutoCaps", "0"))         '
         AutoNum     = VAL(me.GetString("AutoNum", "0"))          '
         AutoSave    = VAL(me.GetString("AutoSave", "2"))         '
         PPreserve   = VAL(me.GetString("Preserve", "0"))         '
         Tabs        = VAL(me.GetString("TabsFlag", "1"))         '
         PMark       = VAL(me.GetString("MarkFlag", "1"))         '
         CapsDesired = VAL(me.GetString("CapsFlag", "0"))         '
         CapsActual  = CapsDesired                                '
         PState      = VAL(me.GetString("StateFlag", "0"))        '
         Fold        = VAL(me.GetString("FoldFlag", "0"))         '
         PCase       = me.GetString("Case", "T")                  '
         HexMode     = VAL(me.GetString("HexMFlag", "1"))         '
         HiFind      = VAL(me.GetString("Hifind", "1"))           '
         HiAuto      = VAL(me.GetString("HiAuto", "1"))           '
         ImportTabs  = VAL(me.GetString("ImportTabs", "4"))       '
         MINLEN      = VAL(me.GetString("MINLEN", "0"))           '
         NumType     = me.GetString("NUMTYPE", "OFF")             '
         LRECL       = VAL(me.GetString("LRECL", "0"))            '
         RECFM       = me.GetString("RECFM", "U")                 '
         UndoNumber  = VAL(me.GetString("UndoLevels", "10"))      '
         ChangeMode  = me.GetString("Change", "D")                '
         EOL         = me.GetString("EOLFlag", "CRLF")            '
         EOLIX       = SWITCH(EOL = "CRLF", 1, EOL = "LF", 2, EOL = "CR", 3, EOL = "NL", 4, EOL = "AUTO", 5, EOL = "AUTONL", 6, EOL = "NONE", 7)
         Cols        = VAL(me.GetString("ColsFlag", "0"))         '
         PageFlag    = VAL(me.GetString("Page", "0"))             '
         PageOffset  = VAL(me.GetString("PageOffSet", "0"))       '
         Cols        = VAL(me.GetString( "ColsFlag", "0"))        '
         SubArg      = me.GetString("SubArg", "")                 '
         SubCmd      = me.GetString("SubCmd", "")                 '
         Start       = UUCASE(me.GetString("Start", "FIRST"))     '

         t = me.GetString("WORD", "MISSING")                      ' Look for WORD string
         WordInput   = IIF$(t = "MISSING", $WORD, t)              ' Set it up
         TP.WordInput = WordInput                                 ' Copy for TP VARPTR usage
         me.WordVal()                                             ' Process it into Prf.PWord

         t = me.GetString("TABS", "MISSING")                      ' Look for TABS string
         TabsLine    = IIF$(t = "MISSING", "", t)                 ' Set it up
         TP.TabsLine = TabsLine                                   ' Copy for TP VARPTR usage

         t = me.GetString("MARK", "MISSING")                      ' Look for MARK string
         MarkLine = IIF$(t = "MISSING", "", t)                    ' Set it up
         TP.MarkLine = MarkLine                                   ' Copy for TP VARPTR usage
         MarkWorking = sMarkSimple(MarkLine)                      ' Setup working version

         t = me.GetString("MASK", "MISSING")                      ' Look for MASK string
         MaskLine    = IIF$(t = "MISSING", "", t)                 ' Set it up
         TP.MaskLine = MaskLine                                   ' Copy for TP VARPTR usage

         '----- Handle old BNDS if new values are not yet present
         BndLeft     = VAL(me.GetString("BndL", "1"))             ' Get Left
         BndRight    = VAL(me.GetString("BndR", "0"))             ' and Right
         me.BuildBndText                                          ' Go build the BndText string
         TP.BndText = BndText                                     ' Copy for TP VARPTR usage

         '----- Get any SOURCE translate tables loaded
         PCollate    = me.GetString("Collate", "ANSI")            '
         PSource     = me.GetString("Source", "ANSI")             '
         me.BuildSourceXlate()                                    '
         TP.ActionCtr = ActionSave                                ' Init action counter
         IF PCase <> "C" AND PCASE <> "T" THEN                    ' Correct an invalid CASE setting.
            PCase = "T"                                           '
            me.SetString("Case", PCase)                           '
         END IF                                                   '
         me.BuildLines                                            ' Rebuild PROF1/2/3/4/5 lines
         MExit
      END METHOD

      METHOD GetCA2SPtr() AS LONG: METHOD = VARPTR(CA2S): END METHOD
      METHOD GetCS2APtr() AS LONG: METHOD = VARPTR(CS2A): END METHOD
      METHOD GetSA2SPtr() AS LONG: METHOD = VARPTR(SA2S): END METHOD
      METHOD GetSS2APtr() AS LONG: METHOD = VARPTR(SS2A): END METHOD


      METHOD GetString(sKey AS STRING, sDefault AS STRING) AS STRING
      '---------- Get string from File Profile ini file
      LOCAL RetVal AS LONG, zResult AS ASCIIZ * 2000
      LOCAL zSection AS ASCIIZ * %MAX_PATH, zKey AS ASCIIZ * %MAX_PATH
      LOCAL zDefault AS ASCIIZ * %MAX_PATH
         zSection = "File" : zKey  = sKey                         '
         zDefault = sDefault                                      '
         RetVal = GetPrivateProfileString(zSection, zKey, zDefault, zResult, SIZEOF(zResult), INIFileNmz)
         IF RetVal THEN METHOD = LEFT$(zResult, RetVal)           '
      END METHOD

      METHOD SetString(sKey AS STRING, sStr AS STRING) AS LONG
      '---------- Set string to ini file                          '
      LOCAL zSection AS ASCIIZ * %MAX_PATH, zKey AS ASCIIZ * %MAX_PATH
      LOCAL zStr AS ASCIIZ * 2000                                 '
         zSection = "File": zKey  = sKey : zStr = sStr            '
         METHOD = WritePrivateProfileString(zSection, zKey, zStr, INIFileNmz)
      END METHOD

      METHOD Touch()
      '---------- Touch the Profile we're loading
      LOCAL hFile AS LONG, hfile2 AS DWORD
      LOCAL CreationTime, LastAccessTime, LastWriteTime AS FILETIME
      LOCAL SystemTimeArea AS SYSTEMTIME
         MEntry
         hFile = FREEFILE                                         ' Get a FileNumber
         OPEN INIFileName FOR APPEND ACCESS READ WRITE LOCK SHARED AS #hFile
         hFile2 = FILEATTR(hFile,2)                               ' Get Windows handle for the file
         GetFileTime hFile2, CreationTime, LastAccessTime, LastWriteTime' Fetch timestamps
         GetSystemTime SystemTimeArea                             ' Get the current time
         SystemtimeToFileTime SystemTimeArea, LastWriteTime       ' Put it in LastWriteTime
         SetFileTime hFile2, CreationTime, LastAccessTime, LastWriteTime'
         CLOSE #hFile                                             ' Close it
         MExit
      END METHOD

      METHOD WordReset()
      '---------- Reset WORD
      LOCAL t AS STRING
         t = me.GetString("WORD", "MISSING")                      ' Look for WORD string
         WordInput = IIF$(t = "MISSING", $WORD, t)                ' Set it up
         me.WordVal()                                             '
      END METHOD

      METHOD WordVal() AS LONG
      '---------- Maniplulate Prf.WordInput to Prf.PWord
      LOCAL tbl, t, op AS STRING
      LOCAL i, j AS LONG
         MEntry
         tbl = SPACE$(256)                                        ' Init the new WORD string
         t = WordInput                                            ' Get a working copy of WORDS line

         DO WHILE LEN(t)                                          ' Break out the operands
            op = sGetWord(t, %Strip, %QuoteNotSig)                ' Get next WORD operand

            IF LEN(op) = 1 THEN                                   ' Just a single character?
               MID$(tbl, ASC(op) + 1, 1) = op                     ' Set it into the final table

            ELSEIF LEN(op) = 2 THEN                               ' Two characters
               IF VERIFY(UUCASE(op), $Hex) = 0 THEN               ' Valid Hex?
                  op = CHR$(VAL("&H" + op))                       ' Convert it to a character
                  MID$(tbl, ASC(op) + 1, 1) = op                  ' Set it into the final table
               ELSE                                               '
                  METHOD = %True                                  ' Flag as error
                  MErrExit(%eFail, "Invalid Hex value in WORD string")
               END IF                                             '

            ELSEIF LEN(op) = 3 THEN                               ' Three, possible a-z type
               IF MID$(op, 2, 1) = "-" AND _                      ' Have a proper - character?
                  MID$(op, 1, 1) <= MID$(op, 3, 1) THEN           ' And in range sequence?
                  FOR i = ASC(MID$(op, 1, 1)) TO ASC(MID$(op, 3, 1))    ' Loop through range
                     MID$(tbl, i + 1, 1) = CHR$(i)                ' Set it into the final table
                  NEXT i                                          '
               ELSE
                  METHOD = %True                                  ' Flag as error
                  MErrExit(%eFail, "Invalid character range in WORD string")
               END IF                                             '

            ELSEIF LEN(op) = 5 THEN                               ' Five, possible hh-hh type
               IF MID$(op, 3, 1) = "-" AND _                      ' Have a proper - character?
                  VERIFY(UUCASE(MID$(op, 1, 2)), $Hex) = 0 AND _  '
                  VERIFY(UUCASE(MID$(op, 4, 2)), $Hex) = 0 AND _  '
                  MID$(op, 1, 2) <= MID$(op, 4, 2) THEN           ' And in range sequence?
                  FOR i = VAL("&H" + MID$(op, 1, 2)) TO VAL("&H" + MID$(op, 4, 2))    ' Loop through range
                     MID$(tbl, i + 1, 1) = CHR$(i)                ' Set it into the final table
                  NEXT i                                          '
               ELSE                                               '
                  METHOD = %True                                  ' Flag as error
                  MErrExit(%eFail, "Invalid hex range in WORD string")
               END IF                                             '

            ELSE                                                  '
               METHOD = %True                                     ' Flag as error
               MErrExit(%eFail, "Invalid string in WORD string")  '
                                                                  '
            END IF                                                '
         LOOP                                                     '
         ChrUpper = $Upper                                        ' Set defaults
         ChrLower = $Lower                                        '
         ChrRegUpper = $RegUpper                                  '
         ChrRegLower = $RegLower                                  '

         IF ISFALSE ENV.ENGChars THEN                             ' Add the non-English ASCII national characters?
            IF INSTR(WordInput, "A-Z") THEN                       ' Uppercase A-Z
               tbl += $UpperNat                                   ' Add them
               ChrUpper = $Upper + $UpperNat                      ' Adjust UC char set
               ChrRegUpper = $RegUpper + $RegUpperNat             '
            END IF                                                '
            IF INSTR(WordInput, "a-z") THEN                       ' Lowercase A-Z
               tbl += $LowerNat                                   ' Add them
               ChrLower = $Lower + $LowerNat                      ' Adjust LC char set
               ChrRegLower = $RegLower + $RegLowerNat             ' Adjust LC char set
            END IF                                                '
         END IF                                                   '

         PWord = REMOVE$(tbl, " ")                                ' Condense the string
         Mexit
      END METHOD

      METHOD WriteAll()
      '---------- Write the File Profile INI values
      LOCAL i, w, h        AS LONG, t AS STRING
         MEntry
         IF PLock THEN MExitMeth                                  ' Don't do if Profile is Locked
         me.SetString("ProfLock", FORMAT$(PLock))                 ' Dump out the values
         me.SetString("ProfUsing", UseName)                       '
         me.SetString("Action", FORMAT$(ActionSave))              '
         me.SetString("AutoBkup", FORMAT$(AutoBkup))              '
         me.SetString("AutoCaps", FORMAT$(AUTOCAPS))              '
         me.SetString("AutoSave", FORMAT$(AutoSave))              '
         me.SetString("StateFlag", FORMAT$(PState))               '
         me.SetString("FoldFlag", FORMAT$(Fold))                  '
         me.SetString("Collate", PCollate)                        '
         me.SetString("Source", PSource)                          '
         me.SetString("Case", PCase)                              '
         me.SetString("Preserve", FORMAT$(PPreserve))             '
         me.SetString("TabsFlag", FORMAT$(Tabs))                  '
         me.SetString("MarkFlag", FORMAT$(PMark))                 '
         me.SetString("CapsFlag", FORMAT$(CapsDesired))           '
         me.SetString("HexMFlag", FORMAT$(HexMode))               '
         me.SetString("HiFind",   FORMAT$(HiFind))                '
         me.SetString("HiAuto", FORMAT$(HiAuto))                  '
         me.SetString("ImportTabs", FORMAT$(ImportTabs))          '
         me.SetString("UndoLevels", FORMAT$(UndoNumber))          '
         me.SetString("MINLEN", FORMAT$(MINLEN))                  '
         me.SetString("NUMTYPE", NumType)                         '
         me.SetString("LRECL", FORMAT$(LRECL))                    '
         me.SetString("RECFM", RECFM)                             '
         me.SetString("EOLFlag", EOL)                             '
         me.SetString("Change", ChangeMode)                       '
         me.SetString("ColsFlag", FORMAT$(Cols))                  '
         me.SetString("ScrollAmount", TRIM$(SCROLL))              '
         me.SetString("Page", FORMAT$(PageFlag))                  '
         me.SetString("PageOffset", FORMAT$(PageOffset))          '
         me.SetString("SubArg", SubArg)                           '
         me.SetString("SubCmd", SubCmd)                           '
         me.SetString("Start", Start)                             '
         me.SetString("WORD", WordInput)                          '
         me.SetString("BndL", FORMAT$(BndLeft))                   '
         me.SetString("BndR", FORMAT$(BndRight))                  '
         me.SetString("TABS", $DQ + RTRIM$(TabsLine) + $DQ)       '
         me.SetString("MARK", $DQ + RTRIM$(MarkLine) + $DQ)       '
         me.SetString("MASK", $DQ + RTRIM$(MaskLine) + $DQ)       '
         MExit
      END METHOD

      METHOD Query(kwd AS STRING) AS STRING
      '---------- Return a query value
         SELECT CASE CONST$ UUCASE(kwd)                           ' See what was asked for
            CASE "NAME":       METHOD = UUCASE(TP.TIPProfile) + IIF$(ISNULL(UseName), "", " USING(" + UUCASE(UseName) + ")") + IIF$(ISTRUE PLock, " LOCKED", " UNLOCKED")
            CASE "ACTION":     METHOD = IIF$(Actionsave, FORMAT$(ActionSave) + " " + ActionVerb, "OFF")
            CASE "AUTOBKUP":   METHOD = IIF$(AutoBkup, "ON", "OFF")
            CASE "AUTOSAVE":   METHOD = IIF$((AutoSave AND 1) = 1, "ON ", "OFF ")  + IIF$((AutoSave AND 2) = 2, "PROMPT", "NOPROMPT")
            CASE "CAPS"                                           '
                               IF CapsDesired <> 2 THEN           '
                                  METHOD = IIF$(CapsDesired, "ON", "OFF")
                               ELSE                               '
                                  METHOD = IIF$(CapsActual, "AUTO:on", "AUTO:off")
                               END IF                             '
            CASE "CASE":       METHOD = UUCASE(PCase)             '
            CASE "CHANGE":     METHOD = IIF$(ChangeMode = "D", "DS", "CS")
            CASE "COLS":       METHOD = IIF$(Cols, "ON", "OFF")   '
            CASE "COLLATE":    METHOD = PCollate                  '
            CASE "EOL":        METHOD = IIF$(VERIFY(EOL, $Hex) = 0, "X'" + EOL + "'", EOL)
            CASE "FOLD":       METHOD = IIF$(Fold, "ON", "OFF")   '
            CASE "HEX":        METHOD = IIF$(HexMode = 4, "ON", "OFF")
            CASE "HIDE":       METHOD = IIF$(IsTPHideFlag, "ON", "OFF")
            CASE "HILITE":                                        '
                               IF ISFALSE HiFind AND ISFALSE HiAuto THEN
                                  METHOD = "OFF"                  '
                               ELSE                               '
                                  METHOD = "FIND " + IIF$(HiFind, "ON", "OFF") + _
                                             " AUTO " + IIF$(HiAuto, "ON", "OFF")
                               END IF                             '
            CASE "XTABS":      METHOD = FORMAT$(ImportTabs)       '
            CASE "LRECL":      METHOD = FORMAT$(LRECL)            '
            CASE "MARK":       METHOD = IIF$(PMark, "ON", "OFF")  '
            CASE "MASK":       METHOD = MaskLine                  '
            CASE "MINLEN":     METHOD = FORMAT$(MINLEN)           '
            CASE "NUMTYPE":    METHOD = NumType                   '
            CASE "NOTIFY":     METHOD = SWITCH$(ENV.NotifyLevelT = 0, "NONE", ENV.NotifyLevelT = 1, "EDIT", ENV.NotifyLevelT = 2, "ALL")
            CASE "PAGE":       METHOD = IIF$(PageFlag, IIF$(PageFlag = 2, "SCROLL", "ON"), "OFF") + IIF$(PageOffset <> 0, "/" + FORMAT$(PageOffset, "+#;-#"), "")
            CASE "PRESERVE":   METHOD = IIF$(PPreserve = 0 , "OFF", IIF$(PPreserve = 1, "ON", "C"))
            CASE "RECFM":      METHOD = RECFM                     '
            CASE "SCROLL":     METHOD = TRIM$(SCROLL)             '
            CASE "SETUNDO":    METHOD = FORMAT$(UndoNumber)       '
            CASE "SOURCE":     METHOD = PSource                   '
            CASE "START":      METHOD = Start                     '
            CASE "STATE":      METHOD = IIF$(PState, "ON", "OFF")
            CASE "SUBARG":     METHOD = IIF$(ISNULL(SubArg), "OFF", SubArg)
            CASE "SUBCMD":     METHOD = IIF$(ISNULL(SubCmd), "OFF", SubCmd)
            CASE "TABS":       METHOD = IIF$(Tabs, "ON", "OFF")   '
            CASE "WORD":       METHOD = PWord                     '

            CASE ELSE                                             ' Oops
               METHOD = ""                                        ' Return null
               TP.ErrFlag = 8: TP.ErrMsg = "Invalid Profile request"
         END SELECT
      END METHOD

      METHOD EditProf() AS STRING
      '---------- Edit a Profile file
      REGISTER j AS LONG
      LOCAL pfn, eProfname AS STRING, InUse, currtab AS LONG
         MEntry                                                   '
         IF ISFALSE ISFILE(INIFileName) THEN _                    ' See if it exists (can't happen, but ...)
            METHOD = "Specified Profile name does not exist": MExitMeth

         me.DispProfile                                           ' Let user set things up
         IF gfDialogDone THEN                                     ' Was dialog cancelled
            METHOD = "Profile EDIT cancelled"                     ' Say so
            MExitMeth                                             '
         END IF                                                   '
         me.WriteAll                                              ' Save it
         sDoStatusBar($AllStatusBarBoxes)                         ' Pick up whatever changed
         METHOD = "Profile Saved"                                 '
         MExit                                                    ' Set focus
      END METHOD

      METHOD DoneCheck() AS LONG
      '---------- Validate dialog variables when DONE clicked
      LOCAL lclScroll, lclText, lclText2, lclkey, lclW1, lclW2, lclRECFM AS STRING, i, j, lclLRECL AS LONG

         '----- Check USING
         CONTROL GET TEXT hPage1, %FILE_PROFUSING TO LclText      ' Validate USING
         lclText = UUCASE(TRIM$(lclText))                         ' Upper case
         IF ISNOTNULL(lclText) THEN                               ' If enteLAttrG, validate it
            IF IsEQ(ProfName, "DEFAULT") THEN                     ' Disallow USING on the DEFAULT Profile
               TAB SELECT hPrf, %TAB_CONTROL, 1                   '
               sDoMsgBox "DEFAULT Profile cannot have a USING value.", %MB_OK OR %MB_USERICON, "SPFLite"
               METHOD = %True: EXIT METHOD                        '
            END IF                                                '
            IF ISFALSE ISFILE(ENV.PROFPath + lclText + ".INI") THEN ' See if it exists
               TAB SELECT hPrf, %TAB_CONTROL, 1                   '
               sDoMsgBox "Specified Using Profile name |K" + lclText + "|B does not exist", %MB_OK OR %MB_USERICON, "SPFLite"
               METHOD = %True: EXIT METHOD                        '
            END IF                                                '
         END IF                                                   '

         '----- Check UNDOLEVELS
         CONTROL GET TEXT hPage1, %FILE_UNDOLEVELS TO LclText     ' Validate UNDOLevels
         IF VAL(lclText) < 0 OR VAL(lclText) > 25 THEN            ' Reasonable range?
            TAB SELECT hPrf, %TAB_CONTROL, 1                      '
            sDoMsgBox "Number of UNDO levels is invalid", %MB_OK OR %MB_USERICON, "SPFLite"
            METHOD = %True: EXIT METHOD                           '
         END IF                                                   '

         '----- Check SCROLL
         CONTROL GET TEXT hPage1, %FILE_SCRLAMTC TO LclScroll     ' Get the Scroll amount
         lclScroll = TRIM$(UUCASE(lclScroll))                     ' Upper case
         IF lclScroll <> "P" AND lclScroll <> "PAGE" AND lclScroll <> "H" AND lclScroll <> "HALF" AND _
            lclScroll <> "D" AND lclScroll <> "DATA" AND lclScroll <> "F" AND lclScroll <> "FULL" AND _
            lclScroll <> "CSR" AND lclScroll <> "C" AND lclScroll <> "CUR" AND VAL(lclScroll) = 0 THEN
            TAB SELECT hPrf, %TAB_CONTROL, 1                      '
            sDoMsgBox "Scroll amount is invalid", %MB_OK OR %MB_USERICON, "SPFLite"       '
            METHOD = %True: EXIT METHOD                           '
         END IF                                                   '

         '----- Check LRECL
         lclLRECL = LRECL                                         ' Save current in case no new one
         CONTROL GET TEXT hPage2, %FILE_LRECL TO LclText          ' Validate LRECL
         IF VAL(lclText) < 0 OR VAL(lclText) > 32768 THEN         ' Reasonable range?
            TAB SELECT hPrf, %TAB_CONTROL, 2                      '
            sDoMsgBox "LRECL is invalid", %MB_OK OR %MB_USERICON, "SPFLite"               '
            METHOD = %True: EXIT METHOD                           '
         END IF                                                   '
         lclLRECL = VAL(lclText)                                  ' Save for next section

         '----- Check RECFM
         lclRECFM = RECFM                                         ' Save current in case no new one
         CONTROL GET TEXT hPage2, %FILE_RECFM TO lclText          '
         lcltext = UUCASE(lcltext)                                ' UC it
         ARRAY SCAN gRECFMList(), = lcltext, TO i                 ' See if one of the 'normal' entries
         IF i THEN                                                ' Yes, must be valid then
            IF IsEQ(lcltext, "F") AND lclLRECL = 0 THEN           '
               TAB SELECT hPrf, %TAB_CONTROL, 2                   '
               sDoMsgBox "Cannot have RECFM=F and LRECL=0 at the same time", %MB_OK OR %MB_USERICON, "SPFLite"
               METHOD = %True: EXIT METHOD                        '
            END IF                                                '
            IF IsEQ(LEFT$(lcltext, 1), "V") AND lclLRECL <> 0 THEN
               TAB SELECT hPrf, %TAB_CONTROL, 2                   '
               sDoMsgBox "RECFM=V will ignore the specified LRECL", %MB_OK OR %MB_USERICON, "SPFLite"
               METHOD = %True: EXIT METHOD                        '
            END IF                                                '
         END IF                                                   '
         lclRECFM = lclText                                       ' Save for next section

         '----- Check EOL
         CONTROL GET TEXT hPage2, %FILE_EOLFlag TO lclText        '
         lcltext = UUCASE(lcltext)                                ' UCASE it
         ARRAY SCAN gEOLFlagList() FOR 7, = lcltext, TO i         ' See if one of the 'normal' entries
         IF i THEN                                                ' Yes, must be valid then
            IF IsEQ(lcltext, "NONE") AND lclLRECL = 0 AND LEFT$(lclRECFM, 1) <> "V" THEN
               TAB SELECT hPrf, %TAB_CONTROL, 2                   '
               sDoMsgBox "Cannot have EOL=NONE, RECFM=" + lclRECFM + ", and LRECL=0 at the same time", %MB_OK OR %MB_USERICON, "SPFLite"
               METHOD = %True: EXIT METHOD                        '
            END IF                                                '
            gEOLFlagList(7) = " "                                 ' Blank the Hex entry now
         ELSE                                                     '
            lcltext = TRIM$(lcltext)                              ' Remove blanks
            IF (LEN(lcltext) <> 2 AND LEN(lcltext) <> 4) OR _     '
            VERIFY(lcltext, $Hex) <> 0 THEN                       ' If not a 2 or 4 byte hex value
               TAB SELECT hPrf, %TAB_CONTROL, 2                   '
               sDoMsgBox "EOL Hex string is not a valid 2 or 4 character hex string", %MB_OK OR %MB_USERICON, "SPFLite"
               METHOD = %True: EXIT METHOD                        '
            END IF                                                '
         END IF                                                   '

         '----- Check XTABS
         CONTROL GET TEXT hPage2, %FILE_XTABS TO LclText          ' Validate XTabs
         IF VAL(lclText) < 0 OR VAL(lclText) > 16 THEN            ' Reasonable range?
            TAB SELECT hPrf, %TAB_CONTROL, 2                      '
            sDoMsgBox "Import tab expand value is invalid", %MB_OK OR %MB_USERICON, "SPFLite"
            METHOD = %True: EXIT METHOD                           '
         END IF                                                   '

         '----- Check CHANGE
         CONTROL GET TEXT hPage2, %FILE_CHANGE TO LclText         ' Validate Change
         IF IsNE(lclText, "DS") AND IsNE(lclText, "CS") THEN      ' Correct choice?
            TAB SELECT hPrf, %TAB_CONTROL, 2                      '
            sDoMsgBox "Change shift type value is invalid", %MB_OK OR %MB_USERICON, "SPFLite"
            METHOD = %True: EXIT METHOD                           '
         END IF                                                   '

         '----- Check MINLEN
         CONTROL GET TEXT hPage2, %FILE_MINLEN TO LclText         ' Validate MinLen
         IF VAL(lclText) < 0 THEN                                 ' Reasonable range?
            TAB SELECT hPrf, %TAB_CONTROL, 2                      '
            sDoMsgBox "Minimum working text length is invalid", %MB_OK OR %MB_USERICON, "SPFLite"
            METHOD = %True: EXIT METHOD                           '
         END IF                                                   '

         '----- Check SOURCE
         CONTROL GET TEXT hPage2, %FILE_SOURCEFLAG TO LclText     ' Validate SOURCE
         LclText = UUCASE(LclText)                                '
         IF LclText <> PSource THEN                               ' If it has been changed
            IF LclText <> "ANSI" AND _                            ' Is SOURCE non-ANSI mode?
               LclText <> "UTF8" AND _                            '
               LclText <> "UTF16" AND _                           '
               LclText <> "UTF16LE" AND _                         '
               LclText <> "UTF16BE" THEN                          '
               IF ISFALSE ISFILE(ENV.INIPath + LclText + ".SOURCE") THEN  ' See if the Custom SOURCE file exists
                  sDoMsgBox "SOURCE translate file: |K" + ENV.INIPath + LclText + ".SOURCE |Bis not found, please correct", _
                            %MB_OK OR %MB_USERICON, "SPFLite"
                  METHOD = %True: EXIT METHOD                     '
               END IF                                             '
            END IF                                                '
         END IF                                                   '

         '----- Check COLLATE
         CONTROL GET TEXT hPage2, %FILE_COLLATE TO LclText        ' Validate SOURCE
         LclText = UUCASE(LclText)                                '
         IF LclText <> PCollate THEN                              ' If it has been changed
            IF LclText <> "ANSI" AND _                            ' Is COLLATE non-ANSI mode?
               LclText <> "UTF8" AND _                            '
               LclText <> "UTF16" AND _                           '
               LclText <> "UTF16LE" AND _                         '
               LclText <> "UTF16BE" THEN                          '
               IF ISFALSE ISFILE(ENV.INIPath + LclText + ".SOURCE") THEN  ' See if the Custom SOURCE file exists
                  sDoMsgBox "COLLATE translate file: |K" + ENV.INIPath + LclText + ".SOURCE |Bis not found, please correct", _
                            %MB_OK OR %MB_USERICON, "SPFLite"
                  METHOD = %True: EXIT METHOD                     '
               END IF                                             '
            END IF                                                '
         END IF                                                   '

         '----- Check ACTION
         CONTROL GET TEXT hPage1, %FILE_ACTION TO LclText         '
         LclText = REMOVE$(LclText, " ")                          ' Remove blanks
         IF INSTR(LclText, ",") = 0 THEN                          ' No comma?
            lclW1 = LclText: lclW2 = ""                           ' Single operand
         ELSE                                                     ' Two operands
            lclW1 = LEFT$(LclText, INSTR(LclText, ",") - 1)       ' Split them
            lclW2 = UUCASE(MID$(LclText, INSTR(LclText, ",") + 1))'
         END IF                                                   '

         IF VERIFY(lclW1, $Numeric) <> 0 THEN                     ' Only numeric
            sDoMsgBox "ACTION 1st operand is invalid", %MB_OK OR %MB_USERICON, "SPFLite"  '
            TAB SELECT hPrf, %TAB_CONTROL, 1                      '
            METHOD = %True: EXIT METHOD                           '
         END IF                                                   '

         IF lclW2 <> "" AND (lclW2 <> "SAVE" AND lclW2 <> "VSAVE") THEN
            sDoMsgBox "ACTION 2nd operand is not |KSAVE|B or |KVSAVE", %MB_OK OR %MB_USERICON, "SPFLite"'
            TAB SELECT hPrf, %TAB_CONTROL, 1                      '
            METHOD = %True: EXIT METHOD                           '
         END IF                                                   '

         '----- Check BOUNDS
         CONTROL GET TEXT hPage1, %FILE_BOUNDS TO LclText         ' Validate BOUNDS
         LclText = REMOVE$(LclText, " ")                          ' Remove blanks
         IF INSTR(LclText, ",") = 0 THEN                          ' No comma?
            lclW1 = LclText: lclW2 = ""                           ' Single operand
         ELSE                                                     ' Two operands
            lclW1 = LEFT$(LclText, INSTR(LclText, ",") - 1)       ' Split them
            lclW2 = MID$(LclText, INSTR(LclText, ",") + 1)        '
         END IF                                                   '
         IF lclW2 = "" AND IsEQ(LclW1, "MAX") THEN                ' Simple MAX
            '                                                     '
         ELSEIF lclW2 = "" AND VAL(lclW1) > 0 THEN                ' Single number?
            '                                                     '
         ELSEIF VAL(lclW1) > 0 AND IsEQ(lclW2, "MAX") THEN        ' nn, MAX
            '                                                     '
         ELSEIF VAL(lclW1) > 0 AND VAL(lclW2) > 0 AND _           '
                VAL(lclW1) < VAL(lclW2) THEN                      ' ll, rr range
            '
         ELSE                                                     '
            TAB SELECT hPrf, %TAB_CONTROL, 1                      '
            sDoMsgBox "BOUNDS specification is invalid", %MB_OK OR %MB_USERICON, "SPFLite"'
            METHOD = %True: EXIT METHOD                           ' ------+
         END IF                                                   '       |
                                                                  '       |
         '----- Now do it again, making the changes               ' NOTE: Keep BOUNDS validation and saving together
                                                                  '       |
         '----- Bounds                                            ' ------+
         IF lclW2 = "" AND IsEQ(LclW1, "MAX") THEN                ' Simple MAX
            BndLeft = 1: BndRight = 0                             '
         ELSEIF lclW2 = "" AND VAL(lclW1) > 0 THEN                ' Single number?
            BndLeft = 1: BndRight = VAL(lclW1)                    '
         ELSEIF VAL(lclW1) > 0 AND IsEQ(lclW2, "MAX") THEN        ' nn, MAX
            BndLeft = VAL(lclW1): BndRight = 0                    '
         ELSEIF VAL(lclW1) > 0 AND VAL(lclW2) > 0 AND _           '
                VAL(lclW1) < VAL(lclW2) THEN                      ' ll, rr range
            BndLeft = VAL(lclW1): BndRight = VAL(lclW2)           '
         END IF                                                   '

         CONTROL GET TEXT hPage1, %FILE_PROFUSING TO lclText      ' Fetch all the values now
         UseName = UUCASE(lclText)                                '
         CONTROL GET TEXT hPage1, %FILE_SUBARG TO lclText         '
         SubArg = UUCASE(lclText)                                 '
         CONTROL GET TEXT hPage1, %FILE_SUBCMD TO lclText         '
         SubCmd = UUCASE(lclText)                                 '
         CONTROL GET TEXT hPage1, %FILE_ACTION TO LclText         '
         LclText = REMOVE$(LclText, " ")                          ' Remove blanks
         IF INSTR(LclText, ",") = 0 THEN                          ' No comma?
            lclW1 = LclText: lclW2 = ""                           ' Single operand
         ELSE                                                     ' Two operands
            lclW1 = LEFT$(LclText, INSTR(LclText, ",") - 1)       ' Split them
            lclW2 = UUCASE(MID$(LclText, INSTR(LclText, ",") + 1))'
         END IF                                                   '
         ActionSave = VAL(lclW1)                                  '
         ActionVerb = UUCASE(lclW2)                               '
         CONTROL GET CHECK hPage1, %FILE_PROFLOCK TO i            '
         PLock = i                                                '
         CONTROL GET CHECK hPage1, %FILE_TABSFLAG TO i            '
         Tabs = i                                                 '
         CONTROL GET CHECK hPage1, %FILE_MARKFLAG TO i            '
         PMark = i                                                '
         CONTROL GET CHECK hPage1, %FILE_COLSFLAG TO i            '
         Cols = i                                                 '
         CONTROL GET CHECK hPage1, %FILE_FOLD TO i                '
         Fold = i                                                 '
         CONTROL GET CHECK hPage1, %FILE_HEXMFLAG TO i            '
         HexMode = IIF(i, 4, 1)                                   '
         CONTROL GET CHECK hPage1, %FILE_HIAUTO TO i              '
         HiAuto = i                                               '
         CONTROL GET CHECK hPage1, %FILE_HIFIND TO i              '
         HiFind = i                                               '
         CONTROL GET CHECK hPage1, %FILE_CASE TO i                '
         PCase = IIF$(i, "C", "T")                                '
         CONTROL GET CHECK hPage1, %FILE_AUTOCAPS TO i            '
         AUTOCAPS = i                                             '
         CONTROL GET CHECK hPage1, %FILE_STATEFLAG TO i           '
         PState = i                                               '
         CONTROL GET TEXT hPage1, %FILE_STATEFLAG TO lclText      '
         IF lclText = "ON" THEN                                   '
            PState = 1                                            '
         ELSEIF lclText = "OFF" THEN                              '
            PState = 0                                            '
         ELSEIF lclText = "FEW" THEN                              '
            PState = 2                                            '
         ELSE                                                     '
            PState = 3                                            '
         END IF                                                   '
         CONTROL GET TEXT hPage1, %FILE_CAPSDFLAG TO lclText      '
         IF lclText = "ON" THEN                                   '
            CapsDesired = 1                                       '
         ELSEIF lclText = "OFF" THEN                              '
            CapsDesired = 0                                       '
         ELSE                                                     '
            CapsDesired = 2                                       '
         END IF                                                   '
         CapsActual = CapsDesired                                 '
         CONTROL GET TEXT hPage1, %FILE_START TO lclText          '
         Start = lclText                                          '
         CONTROL GET TEXT hPage1, %FILE_UNDOLEVELS TO LclText     '
         UndoNumber = VAL(lclText)                                '
         CONTROL GET TEXT hPage1, %FILE_SCRLAMTC TO lclText       '
         IF lclText = "P" OR lclText = "PAGE" THEN                '
            SCROLL = "PAGE"                                       '
         ELSEIF lclText = "F" OR lclText = "FULL" THEN            '
            SCROLL = "FULL"                                       '
         ELSEIF lclText = "H" OR lclText = "HALF" THEN            '
            SCROLL = "HALF"                                       '
         ELSEIF lclText = "D" OR lclText = "DATA" THEN            '
            SCROLL = "DATA"                                       '
         ELSEIF lclText = "CSR" OR lclText = "C" OR lclText = "CUR" THEN
            SCROLL = "CSR "                                       '
         ELSE                                                     '
            SCROLL = lclText                                      '
         END IF                                                   '
         CONTROL GET CHECK hPage2, %FILE_AUTOBKUP TO i            '
         AutoBkup = i                                             '
         CONTROL GET TEXT hPage2, %FILE_PRESERVE TO lclText       '
         IF lclText = "ON" THEN                                   '
            PPreserve = 1                                         '
         ELSEIF lclText = "OFF" THEN                              '
            PPreserve = 0                                         '
         ELSE                                                     '
            PPreserve = 2                                         '
         END IF                                                   '
         CONTROL GET TEXT hPage2, %FILE_SOURCEFLAG TO LclText     '
         IF LclText <> PSource THEN                               ' Is SOURCE changing?
            PSource = UUCASE(LclText)                             ' Set as SOURCE and
            PCollate = UUCASE(LclText)                            ' COLLATE
         ELSE                                                     '
            PSource = UUCASE(lclText)                             ' Just pick it up
         END IF                                                   '
         CONTROL GET TEXT hPage2, %FILE_COLLATE TO LclText        '
         PCollate = UUCASE(LclText)                               ' COLLATE
         CONTROL GET TEXT hPage2, %FILE_AUTOSAVE TO lclText       ' Convert autosave
         IF lclText = "OFF PROMPT" THEN                           ' OFF PROMPT
            AutoSave = 2                                          '
         ELSEIF lclText = "ON NOPROMPT" THEN                      ' ON NOPROMPT
            AutoSave = 1                                          '
         ELSEIF lclText = "ON PROMPT" THEN                        ' ON PROMPT
            AutoSave = 3                                          '
         ELSEIF lclText = "OFF NOPROMPT" THEN                     ' OFF NOPROMPT
            AutoSave = 0                                          '
         END IF                                                   '
         CONTROL GET TEXT hPage2, %FILE_NUMTYPE TO LclText        '
         NumType = uucase(lclText)                                '
         CONTROL GET TEXT hPage2, %FILE_LRECL TO LclText          '
         LRECL = VAL(lclText)                                     '
         CONTROL GET TEXT hPage2, %FILE_RECFM TO LclText          '
         RECFM = lclText                                          '
         CONTROL GET TEXT hPage2, %FILE_EOLFlag TO lclText        '
         EOL = UUCASE(LSET$(lcltext, 4))                          '
         CONTROL GET TEXT hPage2, %FILE_XTABS TO lclText          '
         ImportTabs = VAL(lclText)                                '
         CONTROL GET TEXT hPage2, %FILE_MINLEN TO LclText         '
         MINLEN = VAL(lclText)                                    '
         CONTROL GET TEXT hPage2, %FILE_CHANGE TO lclText         '
         lclText = UUCASE(lclText)                                '
         ChangeMode = IIF$(lclText = "DS", "D", "C")              '
         me.WriteAll                                              ' Save it all
         me.BuildLines                                            ' Rebuild PROF lines
         METHOD = %False                                          ' We made it through
      END METHOD

      METHOD DispProfile()
      '---------- Build and display the Profile Dialog
      LOCAL i AS LONG, t AS STRING
      DIM CAPSFlagList(2) AS STRING
      DIM STATEFlagList(3) AS STRING
      DIM PRESERVEList(2) AS STRING
      DIM STARTList(4) AS STRING
      DIM AutoSaveList(3) AS GLOBAL STRING
         MEntry
         ARRAY ASSIGN CAPSFlagList() = "OFF", "ON", "AUTO"
         ARRAY ASSIGN STATEFlagList() = "OFF", "ON", "FEW", "MANY"
         ARRAY ASSIGN PRESERVEList() = "OFF", "ON", "C Mode"
         ARRAY ASSIGN STARTList() = "FIRST", "LAST", "PRIOR", "LABEL", "NEW"
         ARRAY ASSIGN AutoSaveList() = "OFF NOPROMPT", "ON NOPROMPT", "OFF PROMPT", "ON PROMPT"

         '----- Create the main dialogue
         DIALOG FONT DEFAULT "Tahoma", 13 / gFontScale,0, 0
         DIALOG NEW hWnd, "SPFLite Profile Editor - Profile = " + UUCASE(ProfName), , , 300, 250, %WS_POPUP OR %WS_BORDER OR %WS_DLGFRAME OR %WS_CAPTION OR _
            %WS_CLIPSIBLINGS OR %WS_VISIBLE OR %DS_MODALFRAME OR %DS_3DLOOK OR %DS_NOFAILCREATE OR _
            %DS_SETFONT, %WS_EX_CONTROLPARENT OR %WS_EX_LEFT OR %WS_EX_LTRREADING OR %WS_EX_RIGHTSCROLLBAR, TO hPrf

         '----- Add the Tab Control to main dialog box
         CONTROL ADD TAB,     hPrf, %TAB_CONTROL, "", 1, 1, 299, 225, %TCS_OWNERDRAWFIXED
         CONTROL ADD BUTTON,  hPrf, %DIALOG_CANCEL, "Cancel", 175, 233, 50, 14, %WS_BORDER OR %BS_DEFAULT
         CONTROL SET FONT     hPrf, %DIALOG_CANCEL, hBoldFont
         sToolTipSet (GetDlgItem(hPrf, %DIALOG_CANCEL), " Press Cancel to return without saving the Profile. ")
         CONTROL ADD BUTTON,  hPrf, %DIALOG_DONE, "Done", 235, 233, 50, 14, %WS_BORDER OR %BS_DEFAULT
         sToolTipSet (GetDlgItem(hPrf, %DIALOG_DONE), " Press Done once you have made all desiLAttrG changes. ")
         CONTROL SET FONT     hPrf, %DIALOG_DONE, hBoldFont
         CONTROL ADD IMAGEX,  hPrf, %GENERAL_GPLV3, "GPLV3", 2, 226, 66, 24, %SS_ICON

         '----- Add the Tab 1
         TAB INSERT PAGE hPrf, %TAB_CONTROL, 1, 0, "Editor Options", CALL DlgProfileCallBack TO hPage1

         CONTROL ADD TEXTBOX, hPage1, %FILE_PROFUSING, UseName, 4, 4, 96, 10, %ES_UPPERCASE
         sToolTipSet (GetDlgItem(hPage1, %FILE_PROFUSING), " Enter name of an existing profile which THIS profile should USE," + $CRLF + _
                                                                " No other fields need be completed")
         CONTROL ADD LABEL, hPage1, %FILE_PROFUSING_TEXT, "Profile USE name.", 130, 5, 250, 12
         CONTROL SET COLOR  hPage1, %FILE_PROFUSING_TEXT, %WHITE, -2

         CONTROL ADD TEXTBOX, hPage1, %FILE_SUBCMD, SubCmd, 4,18, 96, 10, %ES_UPPERCASE
         sToolTipSet (GetDlgItem(hPage1, %FILE_SUBCMD), " Enter optional substitute command to use if the SUBMIT command is issued for this file")
         CONTROL ADD LABEL, hPage1, %FILE_SUBCMD_TEXT, "Profile SUBCMD name.", 130, 19, 250, 12
         CONTROL SET COLOR  hPage1, %FILE_SUBCMD_TEXT, %WHITE, -2

         CONTROL ADD TEXTBOX, hPage1, %FILE_SUBARG, SubArg, 4,32, 96, 10, %ES_UPPERCASE
         sToolTipSet (GetDlgItem(hPage1, %FILE_SUBARG), " Enter optional SUBARG string")
         CONTROL ADD LABEL, hPage1, %FILE_SUBARG_TEXT, "Profile SUBARG string.", 130, 33, 250, 12
         CONTROL SET COLOR  hPage1, %FILE_SUBARG_TEXT, %WHITE, -2

         sAddCheck(hPage1,%FILE_PROFLOCK, PLock, 4, 45, 115, "PROFILE LOCK?")
         sToolTipSet (GetDlgItem(hPage1, %FILE_PROFLOCK), "Start the Profile in LOCKed status? ")
         sAddCheck(hPage1, %FILE_TABSFLAG, Tabs, 130, 45, 115, "TABS ON?")
         sToolTipSet (GetDlgItem(hPage1, %FILE_TABSFLAG), " Start the Profile in TABS ON mode? ")
         sAddCheck(hPage1, %FILE_COLSFLAG, Cols, 4, 55, 115, "COLS ON?")
         sToolTipSet (GetDlgItem(hPage1, %FILE_COLSFLAG), "Start the Profile with COLS ON? ")
         sAddCheck(hPage1, %FILE_HEXMFLAG, IIF(HexMode = 1, 0, 1), 130, 55, 115, "HEX ON?")
         sToolTipSet (GetDlgItem(hPage1, %FILE_HEXMFLAG), "Start the Profile with HEX ON? ")
         sAddCheck(hPage1, %FILE_HIFIND, HiFind, 4, 65, 115, "HILITE FIND ON?")
         sToolTipSet (GetDlgItem(hPage1, %FILE_HIFIND), "Start the Profile with HILITE FIND ON? ")
         sAddCheck(hPage1, %FILE_HIAUTO, HiAuto, 130, 65, 115, "HILITE AUTO ON?")
         sToolTipSet (GetDlgItem(hPage1, %FILE_HIAUTO), "Start the Profile with HILITE AUTO ON? ")
         sAddCheck(hPage1, %FILE_CASE, IIF(PCase = "C", 1, 0), 130, 75, 115, "Default CASE=C? (on=C, off=T)")
         sToolTipSet (GetDlgItem(hPage1, %FILE_CASE), "Default CASE for text search strings ")
         sAddCheck(hPage1, %FILE_MARKFLAG, PMark, 4, 75, 115, "MARK ON?")
         sToolTipSet (GetDlgItem(hPage1, %FILE_MARKFLAG), " Start the Profile in MARK ON mode? ")
         sAddCheck(hPage1, %FILE_AUTOCAPS, AUTOCAPS, 130, 85, 115, "AUTOCAPS mode on?")
         sToolTipSet (GetDlgItem(hPage1, %FILE_CASE), "Default for AUTOCAPS mode.  AUTOCAPS converts keywords to CAPS based on the AUTO colorize file settings ")
         sAddCheck(hPage1, %FILE_FOLD, Fold, 130, 95, 115, "FOLD mode on?")
         sToolTipSet (GetDlgItem(hPage1, %FILE_CASE), "Default for FOLD mode.  FOLD displays in UpperCase even if data is actually LowerCase ")

         IF BndLeft = 1 AND BndRight = 0 THEN
            t = "MAX"
         ELSEIF BndRight = 0 THEN
            t = FORMAT$(BndLeft) + ", MAX"
         ELSE
            t = FORMAT$(BndLeft) + ", " + FORMAT$(BndRight)
         END IF
         CONTROL ADD TEXTBOX, hPage1, %FILE_BOUNDS, t, 4, 88, 45, 9, %ES_UPPERCASE
         sToolTipSet (GetDlgItem(hPage1, %FILE_BOUNDS), " Enter 'MAX', 'nn, MAX' or 'll,rr' left/right boundaries,")
         CONTROL ADD LABEL, hPage1, %FILE_BOUNDS_TEXT, "Default BOUNDS?", 60, 88, 250, 11
         CONTROL SET COLOR  hPage1, %FILE_BOUNDS_TEXT, %WHITE, -2

         CONTROL ADD COMBOBOX, hPage1, %FILE_STATEFLAG, STATEFlagList(), 4, 99, 45, 64, %CBS_DROPDOWNLIST OR %WS_VSCROLL
         COMBOBOX SELECT hPage1, %FILE_STATEFLAG, PState + 1
         sToolTipSet (GetDlgItem(hPage1, %FILE_STATEFLAG), " Select desiLAttrG handling of STATE information for this Profile. ")
         CONTROL ADD LABEL, hPage1, %FILE_STATEFLAG_TEXT, "STATE handling?", 60, 100, 150, 11
         CONTROL SET COLOR  hPage1, %FILE_STATEFLAG_TEXT, %WHITE, -2

         CONTROL ADD COMBOBOX, hPage1, %FILE_CAPSDFLAG, CAPSFlagList(), 4, 113, 45, 64, %CBS_DROPDOWNLIST OR %WS_VSCROLL
         COMBOBOX SELECT hPage1, %FILE_CAPSDFLAG, CapsDesired + 1
         sToolTipSet (GetDlgItem(hPage1, %FILE_CAPSDFLAG), " Select ON/OFF to lock the setting, AUTO to base setting on the loaded file. ")
         CONTROL ADD LABEL, hPage1, %FILE_CAPSDFLAG_TEXT, "CAPS mode?", 60, 114, 150, 11
         CONTROL SET COLOR  hPage1, %FILE_CAPSDFLAG_TEXT, %WHITE, -2

         CONTROL ADD TEXTBOX,   hPage1, %FILE_UNDOLEVELS, FORMAT$(UndoNumber), 4, 127, 25, 9
         sToolTipSet (GetDlgItem(hPage1, %FILE_UNDOLEVELS), " DesiLAttrG # of UNDO levels to be maintained, 0 to turn off UNDO.")
         CONTROL ADD LABEL, hPage1, %FILE_UNDOLEVELS_TEXT, "SETUNDO levels, or 0 to turn off UNDO", 60, 127, 250, 11
         CONTROL SET COLOR  hPage1, %FILE_UNDOLEVELS_TEXT, %WHITE, -2

         CONTROL ADD TEXTBOX,   hPage1, %FILE_SCRLAMTC, SCROLL, 4, 139, 45, 9, %ES_CENTER OR %ES_UPPERCASE
         sToolTipSet (GetDlgItem(hPage1, %FILE_SCRLAMTC), " Default scroll amount for PgUp/PgDn commands (CSR, PAGE, FULL, HALF, DATA or ###")
         CONTROL ADD LABEL, hPage1, %FILE_SCRLAMTC_TEXT, "Scroll size when PgUp/PgDn keys are pressed", 60, 139, 250, 11
         CONTROL SET COLOR  hPage1, %FILE_SCRLAMTC_TEXT, %WHITE, -2

         CONTROL ADD COMBOBOX, hPage1, %FILE_START, STARTList(), 4, 150, 35, 80, %CBS_DROPDOWNLIST OR %WS_VSCROLL
         CONTROL ADD LABEL, hPage1, %FILE_START_TEXT, "File positioning on load?", 60, 151, 200, 11
         CONTROL SET COLOR  hPage1, %FILE_START_TEXT, %WHITE, -2
         COMBOBOX SELECT hPage1, %FILE_START, SWITCH(Start = "FIRST", 1, Start = "LAST", 2, Start = "PRIOR", 3, _
                                                     Start = "LABEL", 4, Start = "NEW", 5)
         sToolTipSet (GetDlgItem(hPage1, %FILE_START), " FIRST = Top of file, LAST = Bottom of file, PRIOR = Queue file, LABEL = User's .START label, NEW = Auto .START label. ")

         CONTROL ADD TEXTBOX,   hPage1, %FILE_ACTION, FORMAT$(ActionSave) + "," + ActionVerb, 4, 165, 45, 9
         sToolTipSet (GetDlgItem(hPage1, %FILE_ACTION), " Enter #, SaveVerb. # indicates no. of changes to trigger an automatic SAVE/VSAVE, 0 for NO automatic saves.")
         CONTROL ADD LABEL, hPage1, %FILE_ACTION_TEXT, "SAVE/VSAVE after # changes, or 0 for no saves", 60, 165, 250, 11
         CONTROL SET COLOR  hPage1, %FILE_ACTION_TEXT, %WHITE, -2

         CONTROL ADD LABEL, hPage1, %FILE_MESSAGE1_TEXT, "For NEW - WORD, MARK, MASK and TABS are set to the defaults.", 2, 182, 350, 11
         CONTROL SET COLOR  hPage1, %FILE_MESSAGE1_TEXT, %YELLOW, -2
         CONTROL ADD LABEL, hPage1, %FILE_MESSAGE2_TEXT, "For EDIT - WORD, MARK, MASK and TABS will be unaltered.", 2, 191, 350, 11
         CONTROL SET COLOR  hPage1, %FILE_MESSAGE2_TEXT, %YELLOW, -2
         CONTROL ADD LABEL, hPage1, %FILE_MESSAGE3_TEXT, "Custom values for these can be set the next time the Profile is used.", 2, 200, 350, 11
         CONTROL SET COLOR  hPage1, %FILE_MESSAGE3_TEXT, %YELLOW, -2

         '------ Add the Tab 2
         TAB INSERT PAGE hPrf, %TAB_CONTROL, 2, 0, "File Data Options", CALL DlgProfileCallBack TO hPage2
         sAddCheck(hPage2, %FILE_AUTOBKUP, AutoBkup, 4, 4, 115, "AUTOBKUP ON?")
         sToolTipSet (GetDlgItem(hPage2, %FILE_AUTOBKUP), "Start the Profile with AUTOBKUP ON? ")

         CONTROL ADD COMBOBOX, hPage2, %FILE_PRESERVE, PRESERVEList(), 90, 4, 45, 64, %CBS_DROPDOWNLIST OR %WS_VSCROLL
         CONTROL ADD LABEL, hPage2, %FILE_PRESERVE_TEXT, "PRESERVE spaces on SAVE?", 143, 4, 200, 12
         CONTROL SET COLOR  hPage2, %FILE_PRESERVE_TEXT, %WHITE, -2
         COMBOBOX SELECT hPage2, %FILE_PRESERVE, PPreserve + 1
         sToolTipSet (GetDlgItem(hPage2, %FILE_PRESERVE), " Select OFF to trim spaces, ON to keep them or C Mode to invoke C language mode ON. ")

         CONTROL ADD TEXTBOX, hPage2, %FILE_SOURCEFLAG, PSource, 4, 20, 85, 12
         sToolTipSet (GetDlgItem(hPage2, %FILE_SOURCEFLAG), " Enter the file's SOURCE encoding. ")
         CONTROL ADD LABEL, hPage2, %FILE_SOURCEFLAG_TEXT, "Encoding used (SOURCE)?", 95, 22, 200, 12
         CONTROL SET COLOR  hPage2, %FILE_SOURCEFLAG_TEXT, %WHITE, -2

         CONTROL ADD TEXTBOX, hPage2, %FILE_COLLATE, PCollate, 4, 35, 85, 12
         sToolTipSet (GetDlgItem(hPage2, %FILE_COLLATE), " Choose default Collate handling for HEX and SORT. ")
         CONTROL ADD LABEL, hPage2, %FILE_COLLATE_TEXT, "Collate handling?", 95, 37, 200, 12
         CONTROL SET COLOR  hPage2, %FILE_COLLATE_TEXT, %WHITE, -2

         CONTROL ADD COMBOBOX, hPage2, %FILE_AUTOSAVE, AutoSaveList(), 4, 50, 85, 80, %CBS_DROPDOWNLIST OR %WS_VSCROLL
         COMBOBOX SELECT hPage2, %FILE_AUTOSAVE, AutoSave + 1
         sToolTipSet (GetDlgItem(hPage2, %FILE_AUTOSAVE), " Choose how modified files are to be handled when an END command is issued. ")
         CONTROL ADD LABEL, hPage2, %FILE_AUTOSAVE_TEXT, "Automatic file save at END (AUTOSAVE)?", 95, 52, 200, 12
         CONTROL SET COLOR  hPage2, %FILE_AUTOSAVE_TEXT, %WHITE, -2

         CONTROL ADD LABEL, hPage2, %FILE_NUMBER_TEXT, "NUMBER and NUMTYPE options can only be set using those" + $CRLF + _
                                                       "Primary commands while editing, using this Profile", _
                                    4, 80, 200, 30
         CONTROL SET COLOR  hPage2, %FILE_NUMBER_TEXT, %WHITE, -2

         CONTROL ADD TEXTBOX, hPage2, %FILE_LRECL, FORMAT$(LRECL), 4, 110, 45, 12
         sToolTipSet (GetDlgItem(hPage2, %FILE_LRECL), " If lines are not delimited by special characters, enter the fixed record length. ")
         CONTROL ADD LABEL, hPage2, %FILE_LRECL_TEXT, "Record length if fixed length records (LRECL)?", 60, 112, 200, 12
         CONTROL SET COLOR  hPage2, %FILE_LRECL_TEXT, %WHITE, -2

         i = SWITCH(RECFM = "U", 1, RECFM = "F", 2, RECFM = "V", 3, RECFM = "VBI", 4, RECFM = "VLI", 5)
         CONTROL ADD COMBOBOX, hPage2, %FILE_RECFM, gRECFMList(), 4, 125, 45, 64, %CBS_DROPDOWN OR %WS_VSCROLL
         COMBOBOX SELECT hPage2, %FILE_RECFM, i
         sToolTipSet (GetDlgItem(hPage2, %FILE_RECFM), " Select the file's Record format (U, F or V) ")
         CONTROL ADD LABEL, hPage2, %FILE_RECFM_TEXT, "File Record Format", 60, 127, 200, 12
         CONTROL SET COLOR  hPage2, %FILE_RECFM_TEXT, %WHITE, -2

         i = SWITCH(EOL = "CRLF",     1, _
                    EOL = "LF",       2, _
                    EOL = "CR",       3, _
                    EOL = "NL",       4, _
                    EOL = "AUTO",     5, _
                    EOL = "AUTONL",   6, _
                    EOL = "NONE",     7)
         IF i = 0 THEN                                            ' A hex delimiter?
            i = 8                                                 ' Use the last table item to hold it
            gEOLFlagList(7) = EOL                                 '
         END IF
         CONTROL ADD COMBOBOX, hPage2, %FILE_EOLFLAG, gEOLFlagList(), 4, 140, 45, 64, %CBS_DROPDOWN OR %WS_VSCROLL
         COMBOBOX SELECT hPage2, %FILE_EOLFLAG, i
         sToolTipSet (GetDlgItem(hPage2, %FILE_EOLFLAG), " Select the normal line delimiters or enter a 2 or 4 char. custom delimiter. ")
         CONTROL ADD LABEL, hPage2, %FILE_EOLFLAG_TEXT, "End of line delimiter type (EOL)?", 60, 142, 200, 12
         CONTROL SET COLOR  hPage2, %FILE_EOLFLAG_TEXT, %WHITE, -2

         CONTROL ADD TEXTBOX,   hPage2, %FILE_XTABS, FORMAT$(ImportTabs), 4, 155, 25, 12
         sToolTipSet (GetDlgItem(hPage2, %FILE_XTABS), " The TAB width to be used when expanding tabbed files.")
         CONTROL ADD LABEL, hPage2, %FILE_XTABS_TEXT, "Import Tabs width", 60, 157, 250, 12
         CONTROL SET COLOR  hPage2, %FILE_XTABS_TEXT, %WHITE, -2

         CONTROL ADD TEXTBOX,   hPage2, %FILE_MINLEN, FORMAT$(MINLEN), 4, 170, 25, 12
         sToolTipSet (GetDlgItem(hPage2, %FILE_MINLEN), " Enter the minimum length of text lines DURING the edit.")
         CONTROL ADD LABEL, hPage2, %FILE_MINLEN_TEXT, "Minimum working text line length", 60, 172, 250, 12
         CONTROL SET COLOR  hPage2, %FILE_MINLEN_TEXT, %WHITE, -2

         CONTROL ADD TEXTBOX,   hPage2, %FILE_CHANGE, IIF$(ChangeMode = "D", "DS", "CS"), 4, 185, 25, 12
         sToolTipSet (GetDlgItem(hPage2, %FILE_CHANGE), " Enter your choice for default CHANGE processing")
         CONTROL ADD LABEL, hPage2, %FILE_CHANGE_TEXT, "CHANGE default shift. DS=Data Shift, CS=Column Shift.", 60, 187, 250, 12
         CONTROL SET COLOR  hPage2, %FILE_CHANGE_Text, %WHITE, -2

         '----- DISPLAY the Profile DIALOG BOX
         TAB SELECT hPrf, %TAB_CONTROL, 1                         '
         GRAPHIC REDRAW                                           '
         gfDialogDone = %False                                    ' Clear CANCEL flag
         sPopReady                                                ' Ready for pop-up
         DIALOG SHOW MODAL hPrf CALL DlgProfileCallBack           '
         sPopReset                                                ' Reset popup state
         MExit
      END METHOD

      METHOD BuildLines()
      '---------- Build the Profile Display lines
      LOCAL t1, t2, t3, HoldUsing AS STRING, i, j AS LONG
         MEntry
         t2 =  "PROFILE " + IIF$(ISNULL(UseName), UUCASE(ProfName), ProfName + " " + "USING(" + UUCASE(UseName) + ")") + IIF$(ISTRUE PLock, " LOCKED", " UNLOCKED")
         t2 += ",  ACTION " + IIF$(ActionSave, FORMAT$(ActionSave) + " " + ActionVerb, "OFF")' Build Action status
         t2 += ",  AUTOBKUP " + IIF$(ISTRUE AutoBkup, "ON", "OFF")' Build AutoBkup status
         t2 += ",  AUTOCAPS " + IIF$(ISTRUE AUTOCAPS, "ON", "OFF")' Build AutoCaps status
         t2 += ",  AUTONUM " + IIF$(ISTRUE AutoNum, "ON", "OFF")  ' Build AutoNum status
         TP.Prof1 = t2                                            ' Save it

         t3 =  IIF$((AutoSave AND 1) = 1, "ON ", "OFF ")          ' Build msg string
         t3 += IIF$((AutoSave AND 2) = 2, "PROMPT", "NOPROMPT")   '
         t2 = "AUTOSAVE " + t3                                    ' Add AutoSave
         t2 += ",  CAPS "                                         ' Add Caps
         IF CapsDesired <> 2 THEN                                 ' If simple ON/OFF
            t2 += IIF$(CapsDesired, "ON", "OFF")                  '
         ELSE                                                     ' Else if AUTO
            t2 += IIF$(CapsActual, "AUTO:on", "AUTO:off")         '
         END IF                                                   '
         t2 += ",  CASE " + TRIM$(PCase)                          ' CASE
         t2 += ",  CHANGE " + IIF$(ChangeMode = "D", "DS", "CS")  ' Add Change
         t2 += ",  COLLATE " + PCollate                           ' Add Collate
         t2 += ",  COLS " + IIF$(ISTRUE Cols, "ON", "OFF")        ' Add Cols
         TP.Prof2 = t2                                            ' Save it

         t3 = IIF$(VERIFY(EOL, $Hex) = 0, "X'" + EOL + "'", EOL)  ' If hex operand, frame it for the message
         t2 = "EOL " + TRIM$(t3)                                  ' Add EOL value
         t2 += ",  FOLD " + IIF$(Fold, "ON", "OFF")               ' Add FOLD
         t2 += ",  HEX " + IIF$(HexMode = &1, "OFF", "ON")        ' Add HEX
         t2 += ",  HILITE "                                       ' Build HILITE
         IF ISTRUE HiFind AND ISTRUE HiAuto THEN                  ' All on
            t2 += "FIND AUTO"                                     '
         ELSEIF ISTRUE HiFind THEN                                ' Find
            t2 += "FIND"                                          ' Yes, say FIND is on
         ELSEIF ISTRUE HiAuto THEN                                ' Auto
            t2 += "AUTO"                                          ' Yes, say AUTO is on
         ELSE                                                     ' Neither
            t2 += "OFF"                                           ' Yes, just say OFF
         END IF                                                   '
         t2 += ",  LRECL " + FORMAT$(LRECL)                       ' Add LRECL
         t2 += ",  MARK " + IIF$(ISTRUE PMark, "ON", "OFF")       ' Add MARK
         t2 += ",  MINLEN " + FORMAT$(MINLEN)                     ' Add MINLEN
         TP.Prof3 = t2                                            ' Save it

         t2 = "NUMTYPE " + IIF$(NumType = "", "NONE", NumType)    ' Add NUMTYPE
         t2 += ",  PAGE " + IIF$(ISTRUE PageFlag, IIF$(PageFlag = 2, "SCROLL", "ON"), "OFF")     ' Add PAGE
         IF PageFlag AND PageOffset <> 0 THEN t2 += " " + FORMAT$(PageOffset, "+#;-#") ' Add page offset if needed
         t2 += ",  PRESERVE " + IIF$(PPreserve = 0, "OFF", IIF$(PPreserve = 1, "ON", "C"))   ' Add Preserve
         t2 += ",  RECFM " + RECFM                                ' Add RECFM
         TP.Prof4 = t2                                            ' Save it

         t2 = "SCROLL " + TRIM$(SCROLL)                           ' Add Scroll Default
         t2 += ",  SETUNDO " + FORMAT$(UndoNumber)                ' Add SETUNDO
         t2 += ",  SOURCE " + PSource                             ' SOURCE
         t2 += ",  START " + Start                                ' Start value
         t2 += ",  STATE " + CHOOSE$(PState + 1, "OFF", "ON", "FEW", "MOST") ' State (Labels)
         IF ISNOTNULL(SubArg) THEN                                ' Display Prf.SubArg only if present
            t2 += ",  SUBARG " + SubArg                           ' Prf.SubArg xxxx
         END IF                                                   '
         IF ISNOTNULL(SubCmd) THEN                                ' Display Prf.SubCmd only if present
            t2 += ",  SUBCMD " + SubCmd                           ' Prf.SubCmd xxxx
         END IF                                                   '
         TP.Prof5 = t2                                            ' Save it

         t2 = "TABS " + IIF$(ISTRUE Tabs, "ON", "OFF")            ' Add Tabs
         t2 += ",  XTABS " + FORMAT$(ImportTabs)                  ' Build XTabs
         TP.Prof6 = t2                                            ' Save it
         MExit
      END METHOD

   END INTERFACE

END CLASS
