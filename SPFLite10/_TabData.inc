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

CLASS cObjTabData

   '----- Profile data
   INSTANCE Prf         AS iProf                                  ' Prof data area
   INSTANCE APrf        AS iProf                                  ' Alternate Prof data area
   '----- The following are copies of the Prf.varname, used only so that we can do VARPTR of them for L(i).LTxt
   INSTANCE BndText     AS STRING                                 ' BNDS
   INSTANCE MarkLine    AS STRING                                 ' MARK
   INSTANCE MaskLine    AS STRING                                 ' MASK
   INSTANCE WordInput   AS STRING                                 ' WORD
   INSTANCE TabsLine    AS STRING                                 ' TABS

   '----- Other Tab related data
   INSTANCE CRow        AS LONG                                   ' Cursor Row
   INSTANCE CCol        AS LONG                                   ' Cursor Column
   INSTANCE Zone        AS LONG                                   ' What data field it is in
   INSTANCE ZCol        AS LONG                                   ' Field column (on screen)
   INSTANCE ZLCol       AS LONG                                   ' Field logical column (in field)
   INSTANCE ZLNum       AS LONG                                   ' Line Index if in scrollable line section
   INSTANCE AttnDo      AS LONG                                   ' Flags of ToDo at next attention
   INSTANCE PCmdIX      AS LONG                                   ' Current PCommand Index (since tab initialized)
   INSTANCE ActionCtr   AS LONG                                   ' ACTION counter
   INSTANCE ActionStop  AS LONG                                   ' Stop ACTION (set by End)
   INSTANCE ClipName    AS STRING                                 ' Clipboard name loaded
   INSTANCE CmdClear    AS LONG                                   ' Clear pCommand on next swap
   INSTANCE cDesired    AS INTEGER                                ' Cursor desired
   INSTANCE cCurrent    AS INTEGER                                ' Cursor currently
   INSTANCE cMode       AS LONG                                   ' Cursor Mode
   INSTANCE cModeSave   AS LONG                                   ' Saved Cursor Mode
   INSTANCE cfChange    AS STRING                                 ' Current Change string
   INSTANCE cfChangeT   AS STRING                                 ' Current Change string
   INSTANCE cfCLen      AS LONG                                   ' Current real length of change string
   INSTANCE cfCLenT     AS LONG                                   ' Current real length of change string
   INSTANCE cfFCol      AS LONG                                   ' Current Found column
   INSTANCE cfFColT     AS LONG                                   ' Current Found column
   INSTANCE cfFind      AS STRING                                 ' Current Find string
   INSTANCE cfFindT     AS STRING                                 ' Current Find string
   INSTANCE cfFlag      AS QUAD                                   ' Current Operand flags
   INSTANCE cfFlagT     AS QUAD                                   ' Current Operand flags
   INSTANCE cfFLen      AS LONG                                   ' Current real length of find string
   INSTANCE cfFLenT     AS LONG                                   ' Current real length of find string
   INSTANCE cfFLine     AS LONG                                   ' Current Found line
   INSTANCE cfFLineT    AS LONG                                   ' Current Found line
   INSTANCE cfLMargin   AS LONG                                   ' Current left bound column
   INSTANCE cfLMarginT  AS LONG                                   ' Current left bound column
   INSTANCE cfSplitPt1  AS LONG                                   ' Current SPLIT point (where the | is) Lit1
   INSTANCE cfSplitPt1T AS LONG                                   ' Current SPLIT point (where the | is) Lit1
   INSTANCE cfSplitPt2  AS LONG                                   ' Current SPLIT point (where the | is) Lit2
   INSTANCE cfSplitPt2T AS LONG                                   ' Current SPLIT point (where the | is) Lit2
   INSTANCE cfNotFound  AS LONG                                   ' Current Last search resulted in Not Found 1 = Top, 2 = Bottom
   INSTANCE cfNotFoundT AS LONG                                   ' Current Last search resulted in Not Found 1 = Top, 2 = Bottom
   INSTANCE cfOChange   AS STRING                                 ' Current Original Change string
   INSTANCE cfOChangeT  AS STRING                                 ' Current Original Change string
   INSTANCE cfRMargin   AS LONG                                   ' Current right bound column
   INSTANCE cfRMarginT  AS LONG                                   ' Current right bound column
   INSTANCE cfHiLiteClr AS LONG                                   ' Current HiLite search color
   INSTANCE cfHiLiteClrT AS LONG                                  ' Current HiLite search color
   INSTANCE cfHiLiteOff AS LONG                                   ' Current HiLite -Off color
   INSTANCE cfHiLiteOffT AS LONG                                  ' Current HiLite -Off color
   INSTANCE cfHiLiteOn  AS LONG                                   ' Current HiLite +On  color
   INSTANCE cfHiLiteOnT AS LONG                                   ' Current HiLite +On  color
   INSTANCE ClrComment()AS ClrCommModel                           ' Colorize comment controls
   INSTANCE ClrExclude()AS ClrExcIncModel                         ' Colorize EXCLUDE controls
   INSTANCE ClrInclude()AS ClrExcIncModel                         ' Colorize INCLUDE controls
   INSTANCE ClrDLM      AS STRING                                 ' Delimiters
   INSTANCE ClrESC      AS STRING                                 ' Escape character (or blank)
   INSTANCE ClrQuoted   AS LONG                                   ' Quoted string color
   INSTANCE ClrNumeric  AS LONG                                   ' Numeric string color
   INSTANCE ClrKWrdNum  AS LONG                                   ' Number of KW entries
   INSTANCE ClrKW()     AS STRING                                 ' Table of colorize Kwds
   INSTANCE ClrKWAns()  AS LONG                                   ' Table of colorize Kwd answers (Scheme #)
   INSTANCE CmdStack()  AS STRING                                 ' Command line stack
   INSTANCE CmdStackNum AS INTEGER                                ' Number of entries in CmdStack()
   INSTANCE CsrAbsLine  AS LONG                                   ' Abs line number as of last Attn
   INSTANCE CsrLinDX    AS LONG                                   ' Hidden DX cursor line
   INSTANCE CurrPCmd    AS STRING                                 ' Latest (running) PCmd
   INSTANCE CursLine    AS STRING                                 ' Line under the Cursor at last AttnFlag
   INSTANCE CursWord    AS STRING                                 ' Word under the Cursor at last AttnFlag
   INSTANCE CursFindInv AS rect                                   ' FIND hilite under the cursor
   INSTANCE DotSet      AS LONG                                   ' Range other than .ZFIRST to .ZLAST is set
   INSTANCE DotSetT     AS LONG                                   ' Range other than .ZFIRST to .ZLAST is set
   INSTANCE DotStart    AS LONG                                   ' Start range (.ZFIRST)
   INSTANCE DotStartT   AS LONG                                   ' Start range (.ZFIRST)
   INSTANCE DotSCond    AS STRING                                 ' Condition test for Start
   INSTANCE DotSCondT   AS STRING                                 ' Condition test for Start
   INSTANCE DotAndOr    AS STRING                                 ' Conector for tests
   INSTANCE DotAndOrT   AS STRING                                 ' Conector for tests
   INSTANCE DotEnd      AS LONG                                   ' End Range (.ZLAST)
   INSTANCE DotEndT     AS LONG                                   ' End Range (.ZLAST)
   INSTANCE DotECond    AS STRING                                 ' Condition test for End
   INSTANCE DotECondT   AS STRING                                 ' Condition test for End
   INSTANCE DotFlag     AS LONG                                   ' K/+/- flag from range
   INSTANCE DotFlagT    AS LONG                                   ' K/+/- flag from range
   INSTANCE DotTag      AS STRING                                 ' Range is a Tag range
   INSTANCE DotTagT     AS STRING                                 ' Range is a Tag range
   INSTANCE DotTCond    AS STRING                                 ' Condition test for Tag
   INSTANCE DotTCondT   AS STRING                                 ' Condition test for Tag
   INSTANCE DotMode     AS STRING                                 ' Line control mode (U/T)
   INSTANCE DotModeT    AS STRING                                 ' Line control mode (U/T)
   INSTANCE Encoding    AS STRING                                 ' Character Encoding (ANSI, UTF8, UTF16, UTF16BE)
   INSTANCE ErrFlag     AS LONG                                   ' Error level primary commands
   INSTANCE ErrMsg      AS STRING                                 ' Text of the error message
   INSTANCE ErrMsgTbl() AS STRING                                 ' Table of the error message
   INSTANCE ErrMsgTblC  AS LONG                                   ' # entries in Table of error messages
   INSTANCE ErrMsgHlp() AS STRING                                 ' Copy for HELP 2nd level use
   INSTANCE ErrMsgHlpC  AS LONG                                   ' # entries in Hlp Table
   INSTANCE ErrMsgHigh  AS LONG                                   ' High RC in the list
   INSTANCE FileRecs    AS LONG                                   ' # lines written
   INSTANCE FindWord    AS LONG                                   ' Word/Chars mode
   INSTANCE FIXCtr      AS LONG                                   ' Find Index counter
   INSTANCE Flag        AS LONG                                   ' Misc bit flags
   INSTANCE Flag2       AS LONG                                   ' More misc bit flags
   INSTANCE gHandle     AS DWORD                                  ' Window handle to the graphic area
   INSTANCE hPCRE       AS DWORD                                  ' PCRE area returned by pcre_compile
   INSTANCE pWatchData  AS WatchData POINTER                      ' Pointer to the param list
   INSTANCE pUNDOData   AS UNDOType POINTER                       ' Pointer to the UNDO Data
   INSTANCE UNDOThread  AS LONG                                   ' Thread address
   INSTANCE wResult     AS DWORD                                  '
   INSTANCE InsTbl()    AS LONG                                   ' Table of insert lines
   INSTANCE InsCount    AS LONG                                   ' No. of insert lines
   INSTANCE InsLastClr  AS LONG                                   ' Last line number cleared of Insert mode
   INSTANCE L()         AS Dataline                               ' TEXT control array
   INSTANCE LCtr        AS LONG                                   ' Free items in L()
   INSTANCE LastBnds    AS STRING                                 ' Last Bounds line
   INSTANCE LastLine    AS LONG                                   ' Last line in the L() array
   INSTANCE LastReal    AS LONG                                   ' Last real line (line count)
   INSTANCE LastDrawn   AS LONG                                   ' Last DispScreen drawn line
   INSTANCE LastPTCurs  AS LONG                                   ' Last PT Cursor col
   INSTANCE LastRulCol  AS LONG                                   ' Last Ruler Cursor col
   INSTANCE LastRulRow  AS LONG                                   ' Last Ruler Cursor row
   INSTANCE LineFirst   AS LONG                                   ' First text line (** top of data **)
   INSTANCE LineLast    AS LONG                                   ' Last  text line (** end of data **)
   INSTANCE LocColor    AS LONG                                   ' LOCATE color index
   INSTANCE LocFind     AS STRING                                 ' LOCATE/FIND last command
   INSTANCE LocFlag     AS LONG                                   ' LOCATE Flag
   INSTANCE LocLine     AS LONG                                   ' LOCATE last found line (resume)
   INSTANCE LocLineR    AS LONG                                   ' LOCATE last found line (Real found)
   INSTANCE LocDir      AS LONG                                   ' LOCATE direction
   INSTANCE LocLabel1   AS STRING                                 ' LOCATE label1 string
   INSTANCE LocLabel2   AS STRING                                 ' LOCATE label2 string
   INSTANCE LocLabLine1 AS LONG                                   ' LOCATE label1 string
   INSTANCE LocLabLine2 AS LONG                                   ' LOCATE label2 string
   INSTANCE LocNotFound AS LONG                                   ' LOCATE last was Not Found
   INSTANCE LocNFDir    AS LONG                                   ' LOCATE Not Found Dir
   INSTANCE LocTag      AS STRING                                 ' LOCATE tag string
   INSTANCE LocLength   AS LONG                                   ' LOCATE length for LOCATE LONG nnn
   INSTANCE LocNVar     AS LONG                                   ' LOCATE NOTE variety
   INSTANCE mapstr_sequence_num  AS  LONG                         ' Seq number for Mapstr use
   INSTANCE MarkedLine  AS LONG                                   ' Line No. of marked line number, 0 if none
   INSTANCE MarkSCol    AS LONG                                   ' Mark - Start Column
   INSTANCE MarkECol    AS LONG                                   ' Mark - End Column
   INSTANCE MarkSLin    AS LONG                                   ' Mark - Start Line
   INSTANCE MarkELin    AS LONG                                   ' Mark - End Line
   INSTANCE MarkRect    AS rect                                   ' Mark - Screen coordinates of mark
   INSTANCE MaxLength   AS LONG                                   ' Length of maximum Text line
   INSTANCE MEditCount  AS LONG                                   ' MEDIT count
   INSTANCE MEditList() AS STRING                                 ' MEDIT filename list
   INSTANCE MEditFlag() AS LONG                                   ' MEDIT filename flags
   INSTANCE MiscZone    AS LONG                                   ' MiscMark - Screen Zonen
   INSTANCE MiscLine    AS LONG                                   ' MiscMark - AFList Index for Note
   INSTANCE MiscSCol    AS LONG                                   ' MiscMark - Start Column
   INSTANCE MiscECol    AS LONG                                   ' MiscMark - End Column
   INSTANCE MiscRect    AS rect                                   ' MiscMark - Screen coordinates of mark
   INSTANCE Offset      AS LONG                                   ' Right scroll Offset
   INSTANCE COffset     AS LONG                                   ' Right scroll Offset of pCommand
   INSTANCE pCommand    AS STRING                                 ' Working (current) command
   INSTANCE pCommandRaw AS STRING                                 ' pCommand before PFK prepending
   INSTANCE pCommandPrv AS STRING                                 ' pCommand previous if prefixed by &
   INSTANCE PIC         AS PIC_type                               ' Picture string control area
   INSTANCE PrevPCmd    AS STRING                                 ' Previous PCmd
   INSTANCE pfChange    AS STRING                                 ' Previous Change string
   INSTANCE pfCLen      AS LONG                                   ' Previous real length of change string
   INSTANCE pfFCol      AS LONG                                   ' Previous Found column
   INSTANCE pfFind      AS STRING                                 ' Previous Find string
   INSTANCE pfFlag      AS QUAD                                   ' Previous Operand flags
   INSTANCE pfFLen      AS LONG                                   ' Previous real length of find string
   INSTANCE pfFLine     AS LONG                                   ' Previous Found line
   INSTANCE pfNotFound  AS INTEGER                                ' Previous Last search resulted in Not Found
   INSTANCE pfOChange   AS STRING                                 ' Previous Original Change string
   INSTANCE PgHandle    AS LONG                                   ' Handle of edit window
   INSTANCE PgNumber    AS LONG                                   ' Page number of the Tab
   INSTANCE Prof1       AS STRING                                 ' Profile display 1
   INSTANCE Prof2       AS STRING                                 ' Profile display 2
   INSTANCE Prof3       AS STRING                                 ' Profile display 3
   INSTANCE Prof4       AS STRING                                 ' Profile display 4
   INSTANCE Prof5       AS STRING                                 ' Profile display 5
   INSTANCE Prof6       AS STRING                                 ' Profile display 6
   INSTANCE PTFDisp     AS LONG                                   ' PowerType First line on a screen
   INSTANCE PTLDisp     AS LONG                                   ' PowerType Last  line on a screen
   INSTANCE ResetClr    AS LONG                                   ' Reset color requested color
   INSTANCE S()         AS LONG                                   ' Screen line to L() index crossref
   INSTANCE SavCurCol   AS LONG                                   ' Saved Cursor Column
   INSTANCE SavCurLin   AS LONG                                   ' Saved Cursor Line
   INSTANCE SavFilLCtl  AS STRING                                 ' For range type if used
   INSTANCE SBCount     AS LONG                                   ' Entries in SBTable
   INSTANCE SBXref()    AS LONG                                   ' Xref to SB Left right order
   INSTANCE SBTable()   AS sBarEntry                              ' StatusBar messages
   INSTANCE ScrlAmtLast AS STRING                                 ' Last displayed ScrlAmtC
   INSTANCE sCurPrio    AS INTEGER                                ' Priority of current sCur... entries
   INSTANCE sCurLin     AS LONG                                   ' Current winning line number
   INSTANCE sCurScrl    AS LONG                                   ' Current winning SCRL operand
   INSTANCE sCurHexl    AS LONG                                   ' Current winning Hex line
   INSTANCE sCol        AS LONG                                   ' Search col
   INSTANCE Scr()       AS STRING                                 ' Screen image in text mode
   INSTANCE sDir        AS LONG                                   ' Search Dir
   INSTANCE SlecSCol    AS LONG                                   ' Select - Start Column
   INSTANCE SlecECol    AS LONG                                   ' Select - End Column
   INSTANCE SlecSLin    AS LONG                                   ' Select - Start Line
   INSTANCE SlecELin    AS LONG                                   ' Select - End Line
   INSTANCE sLine       AS LONG                                   ' Search Line
   INSTANCE SwapLines   AS LONG                                   ' Swap - # of data lines
   INSTANCE SwapSCol    AS LONG                                   ' Swap - Start Column
   INSTANCE SwapECol    AS LONG                                   ' Swap - End Column
   INSTANCE SwapSLin    AS LONG                                   ' Swap - Start Line
   INSTANCE SwapELin    AS LONG                                   ' Swap - End Line
   INSTANCE SwapRect    AS rect                                   ' Swap - Saved Line/Column range marked
   INSTANCE T()         AS STRING                                 ' To hold the dynamic TEXT strings
   INSTANCE TW()        AS WSTRING                                ' To hold the dynamic Attr strings
   INSTANCE TTCtr       AS LONG                                   ' Free items in T()
   INSTANCE TabsDflt    AS STRING                                 ' Default tabs
   INSTANCE TCtr        AS LONG                                   ' Touch counter
   INSTANCE TTbl()      AS TouchEntry                             ' Touch table
   INSTANCE TIDX        AS STRING                                 ' 0/1 to indicate whether t(ix) is in use
   INSTANCE TIP         AS iIO                                    ' Tab primary input file IO
   INSTANCE TOP         AS iIO                                    ' Tab primary output file IO
   INSTANCE TopPCmd     AS STRING                                 ' Top Level PCmd
   INSTANCE TMode       AS LONG                                   ' Tab Mode flags
   INSTANCE TopDrawn    AS LONG                                   ' Top DispScreen drawn line
   INSTANCE TopScrn     AS LONG                                   ' Line # at top of screen
   INSTANCE UBoundL     AS LONG                                   ' Current UBound of L()
   INSTANCE UBoundT     AS LONG                                   ' Current UBound of T()
   INSTANCE Undo()      AS UndoType                               ' Array of Undo stuff
   INSTANCE UndoCurr    AS LONG                                   ' Current number saved
   INSTANCE VSaveFile   AS STRING                                 ' VSAVE filename if VSAVE done
   INSTANCE WatchFlag   AS STRING                                 ' WatchFlag "D" Deleted "M" Modified
   INSTANCE WindowID    AS LONG                                   ' Dialog Window ID

   '----- File Manager Data
   INSTANCE NestCtr     AS LONG                                   ' Counter of NestTbl() entries
   INSTANCE NestTbl()   AS STRING                                 ' Save table for FPath, FMask and FileListNm
   INSTANCE NestTopScrn AS LONG                                   ' Desired Topscreen when Nest is popped
   INSTANCE NestCRow    AS LONG                                   ' Desired CRow when Nest is popped
   INSTANCE FMode       AS LONG                                   ' Current FM mode of the tab
   INSTANCE FPath       AS STRING                                 ' Default directory path
   INSTANCE FMask       AS STRING                                 ' Default file mask
   INSTANCE FileListNm  AS STRING                                 ' Active recent file list name
   INSTANCE ScrlAmtC    AS STRING                                 ' Scroll Amount field (Char)
   INSTANCE DefCommand  AS STRING                                 ' Default LMB generated command
   INSTANCE LFPath      AS STRING                                 ' Last File Path
   INSTANCE LFMask      AS STRING                                 ' Last File Mask
   INSTANCE LFileListNm AS STRING                                 ' Last FileList Name
   INSTANCE LScrlAmtC   AS STRING                                 ' Last Scroll Amount
   INSTANCE RQList()    AS STRING                                 ' Current request list (like FILELIST contents)
   INSTANCE RFList()    AS STRING                                 ' Current Forget  list
   INSTANCE RXList()    AS STRING                                 ' Current Exclude list
   INSTANCE RNList()    AS STRING                                 ' Current Note list
   INSTANCE AFList()    AS FMFList                                ' Latest Dir scan results
   INSTANCE RQCount     AS LONG                                   ' Count in RQList
   INSTANCE RFCount     AS LONG                                   ' Count in RFList
   INSTANCE RXCount     AS LONG                                   ' Count in RXList
   INSTANCE RNCount     AS LONG                                   ' Count in RNList
   INSTANCE AFCount     AS LONG                                   ' Count in AFList
   INSTANCE NumFiles    AS LONG                                   ' Number of Files
   INSTANCE NumDirs     AS LONG                                   ' Number of Dirs
   INSTANCE CmdEmpty    AS STRING                                 ' Cmd area for New Empty
   INSTANCE CmdFilePath AS STRING                                 ' Cmd area for FilePath
   INSTANCE CmdFavorite AS STRING                                 ' Cmd area for Favorites
   INSTANCE CmdFileList AS STRING                                 ' Cmd area for FileList
   INSTANCE CmdFound    AS STRING                                 ' Cmd area for Found
   INSTANCE CmdRecent   AS STRING                                 ' Cmd area for Recent Files
   INSTANCE CmdRPaths   AS STRING                                 ' Cmd area for Recent Paths
   INSTANCE CmdOpen     AS STRING                                 ' Cmd area for Open Files
   INSTANCE CmdProfiles AS STRING                                 ' Cmd area for Profiles
   INSTANCE CallMsg     AS STRING                                 ' Msg to issue by FM in a CallTab invocation
   INSTANCE DefSort     AS STRING                                 ' Default FM sort
   INSTANCE RecentSort  AS STRING                                 ' Default FM sort for Recent
   INSTANCE DirSort     AS STRING                                 ' Defalt Dir positioning
   INSTANCE LastTop     AS LONG                                   ' TopScrn as at last display
   INSTANCE FMMarkdSLin AS LONG                                   ' Last marked screen line
   INSTANCE FMMarkdALin AS LONG                                   ' Last marked AFList line number
   INSTANCE SaveCase    AS STRING                                 ' Profile Case Save
   INSTANCE SaveWord    AS LONG                                   ' FindWord Save
   INSTANCE FFSearch    AS STRING                                 ' FM FF search operand
   INSTANCE FSearch     AS STRING                                 ' FM Find search operand
   INSTANCE LastLoc     AS LONG                                   ' FM Last found line number
   INSTANCE LCmdOps()   AS STRING                                 ' FM Line command operands
   INSTANCE LCmdOpsNum  AS LONG                                   ' FM Line command operands number
   INSTANCE FFCmd       AS STRING                                 ' FF command string
   INSTANCE MEditCmd    AS STRING                                 ' The MEDit command string
   INSTANCE OFrmFPath   AS STRING                                 ' Tab opened from FM Path variables
   INSTANCE OFrmFMask   AS STRING                                 '
   INSTANCE OFrmFileL   AS STRING                                 '


   CLASS METHOD CREATE()                                          ' Constructor - Initialize Class stuff
   LOCAL i, j, k AS LONG, t AS STRING
   DIM SBParse(1 TO 15) AS STRING
      LET Prf     = CLASS "cProf"                                 ' Profile Data
      LET APrf    = CLASS "cProf"                                 ' Alternate Profile Data
      '----- The following are copies of the Prf.varname, used only so that we can do VARPTR of them for L(i).LTxt
      BndText     = ""                                            ' BNDS
      MarkLine    = ""                                            ' MARK
      MaskLine    = ""                                            ' MASK
      WordInput   = ""                                            ' WORD
      TabsLine    = ""                                            ' TABS


      PTFDisp     = 0                                             ' PowerType First Displayed
      PTLDisp     = 0                                             ' PowerType Last Displayed
   '----- initialize the Tab defaults
      CRow        = 1                                             ' Cursor Row
      CCol        = 11                                            ' Cursor Column
      Zone        = 0                                             ' What data field it is in
      ZCol        = 0                                             ' Field column (on screen)
      ZLCol       = 0                                             ' Field logical column (in field)
      ZLNum       = 0                                             ' Line Index if in scrollable line section
      AttnDo      = 0                                             ' Flags of ToDos at next attention
      PCmdIX      = 0                                             ' Primary Command Index
      ActionCtr   = 0                                             ' Action counter
      ActionStop  = %False                                        ' Stop Action flag
      ClipName    = ""                                            ' Name of the loaded CLIP data
      CmdClear    = %False                                        ' Clear command next tab switch
      cDesired    = 0                                             ' Cursor desired
      cCurrent    = 0                                             ' Cursor currently
      cMode       = 0                                             ' Cursor Mode (True = Insert mode)
      cModeSave   = 0                                             ' Cursor Mode (Saved)
      cfChange    = ""                                            ' Current Change string
      cfCLen      = 0                                             ' Current real length of change string
      cfFCol      = 0                                             ' Current Found column
      cfFind      = ""                                            ' Current Find string
      cfFlag      = 0                                             ' Current Operand flags
      cfFLen      = 0                                             ' Current real length of find string
      cfFLine     = 0                                             ' Current Found line
      cfLMargin   = 0                                             ' Current left bound column
      cfSplitPt1  = 0                                             ' Current split point (where the | is) Lit1
      cfSplitPt2  = 0                                             ' Current split point (where the | is) Lit2
      cfNotFound  = 0                                             ' Current Last search resulted in Not Found
      cfOChange   = ""                                            ' Current Original Change string
      cfRMargin   = 0                                             ' Current right bound column
      cfHiLiteClr = 0                                             ' Current HiLite search color
      cfHiLiteOff = 0                                             ' Current HiLite -Off color
      cfHiLiteOn  = 0                                             ' Current HiLite +On  color

  DIM ClrComment(1 TO 9) AS INSTANCE ClrCommModel                 ' Colorize comment Control
  DIM ClrExclude(1 TO 9) AS INSTANCE ClrExcIncModel               ' Colorize Exclude Control
  DIM ClrInclude(1 TO 9) AS INSTANCE ClrExcIncModel               ' Colorize Exclude Control
      FOR i = 1 TO 9                                              ' Reset the comment control
         ClrComment(i).Txt1 = ""                                  '
         ClrComment(i).sCol = ""                                  '
         ClrComment(i).clr = 0                                    '
         ClrComment(i).fnd = 0                                    '
         ClrExclude(i).Txt1 = ""                                  '
         ClrExclude(i).sCol = 0
         ClrInclude(i).Txt1 = ""                                  '
         ClrInclude(i).sCol = 0
      NEXT                                                        '
      ClrDLM      = ""                                            ' Delimiters
      ClrESC      = ""                                            ' Escape char
      ClrQuoted   = 0                                             ' Quoted string color
      ClrNumeric  = 0                                             ' Numeric string color
      ClrKWrdNum  = 0                                             ' Number of KW entries
  DIM ClrKW(1000) AS INSTANCE STRING                              ' Table of colorize Kwds
  DIM ClrKWAns(1000) AS INSTANCE LONG                             ' Table of colorize Kwd answers
  DIM CmdStack(20) AS INSTANCE STRING                             ' Stacked Command line(s)
      CmdStackNum = 0                                             ' Number of stacked commands
      CsrAbsLine  = 0                                             ' Abs line number at last Attn
      CsrLinDX    = 0                                             ' Hidden DX line number
      CursLine    = ""                                            ' Line under the Cursor at last AttnFlag
      CurrPCmd    = " "                                           ' Latest (running) PCmd
      CursWord    = ""                                            ' Word under the Cursor at last AttnFlag
      DotSet      = %False                                        ' Range other than .ZFIRST to .ZLAST is set
      DotStart    = 0                                             ' Start range (.ZFIRST)
      DotSCond    = "GE"                                          ' Condition test for Start
      DotAndOr    = "AND"                                         ' Conector for tests
      DotEnd      = 0                                             ' End Range (.ZLAST)
      DotECond    = "LE"                                          ' Condition test for End
      DotFlag     = 0                                             ' K/+/- range flag
      DotTag      = ""                                            ' Range is a Tag range
      DotTCond    = "EQ"                                          ' Condition test for Tag
      DotMode     = " "                                           ' Line control mode (U/T)
      Encoding    = "ANSI"                                        ' Default to ANSI as usual
      ErrFlag     = 0                                             ' Start as no errors
      ErrMsg      = ""                                            ' Text of the error message
  DIM ErrMsgTbl(1 TO 100) AS INSTANCE STRING                      ' Table of error messages
      ErrMsgTblC  = 0                                             ' Count of Table of error messages
  DIM ErrMsgHlp(1 TO 100) AS INSTANCE STRING                      ' Table of error messages
      ErrMsgHlpC  = 0                                             ' Count of Table of error messages
      ErrMsgHigh  = 0                                             ' High RC in the list
      FileRecs    = 0                                             ' # Recs written?
      FindWord    = ENV.FindWord                                  ' WORD / Chars mode from global
      FIXCtr      = 0                                             ' FIND Index counter
      Flag        = 0                                             ' Misc bit flags
      Flag2       = 0                                             ' More misc bit flags
      gHandle     = 0                                             ' Graphic window handle
      hPCRE       = 0                                             ' PCRE compile area
  DIM InsTbl(1 TO 1000) AS INSTANCE LONG                          ' Table of insert lines
      InsCount    = 0                                             ' No. of insert lines
      InsLastClr  = 0                                             ' Last line number cleared of Insert mode
  DIM L(100000)   AS INSTANCE Dataline                            ' TEXT control array
      LCtr        = 100000                                        ' Number free in L()
      LastBnds    = ""                                            ' Last Bnds Line
      LastLine    = 2                                             ' Last Line
      LastReal    = 0                                             ' Last Real Line (Line Count)
      LastDrawn   = 0                                             ' Last Dispscreen drawn line
      LastPTCurs  = 0                                             ' Last PT Cursor col
      LastRulCol  = 0                                             ' Last Ruler Cursor col
      LastRulRow  = 0                                             ' Last Ruler Cursor row
      LineFirst   = 0                                             ' First text line (** top of data **)
      LineLast    = 0                                             ' Last  text line (** end of data **)
      LocColor    = 0                                             ' LOCATE color Index
      LocFind     = ""                                            ' LOCATE/FIND last command
      LocFlag     = 0                                             ' LOCATE Flag
      LocLine     = 0                                             ' LOCATE last found line resume
      LocLineR    = 0                                             ' LOCATE last found line real line
      LocDir      = 1                                             ' LOCATE Direction
      LocLabel1   = ""                                            ' LOCATE label1 string
      LocLabel2   = ""                                            ' LOCATE label2 string
      LocLabLine1 = 0                                             ' LOCATE label1 string
      LocLabLine2 = 0                                             ' LOCATE label2 string
      LocNotFound = 0                                             ' LOCATE last not found
      LocNFDir    = 0                                             ' LOCATE not found dir
      LocTag      = ""                                            ' LOCATE tag string
      LocLength   = 0                                             ' LOCATE length
      LocNVar     = 0                                             ' LOCATE NOTE variety
      mapstr_sequence_num = 0                                     ' Seq number for Mapstr use
      MaxLength   = 0                                             ' Length of longest line
      MarkedLine  = 0                                             ' Line No. of marked line, 0 if none
      MarkSCol    = 0                                             ' Mark - Start Column
      MarkECol    = 0                                             ' Mark - End Column
      MarkSLin    = 0                                             ' Mark - Start Line
      MarkELin    = 0                                             ' Mark - End Line
      MEditCount  = 0                                             ' MEDIT count
  DIM MEditList(1 TO 20) AS INSTANCE STRING                       ' MEDIT filename list
  DIM MEditFlag(1 TO 20) AS INSTANCE LONG                         ' MEDIT filename flags
      MiscZone    = 0                                             ' MiscMark - Screen Zone
      MiscLine    = 0                                             ' MiscMark - AFList Index for Note
      MiscSCol    = 0                                             ' MiscMark - Start Column
      MiscECol    = 0                                             ' MiscMark - End Column
      Flag2       = Flag2 OR IIF(ENV.InsMode, %NsrtFlag, 0)       ' Set default Insert mode
      Flag2       = Flag2 OR IIF(ENV.InsMode, %NsrtSave, 0)       ' Set default Insert mode
      Offset      = 0                                             ' Right scroll Offset
      COffset     = 0                                             ' Right scroll Offset of pCommand
      pCommand    = ""                                            ' Working (current) command
      pCommandRaw = ""                                            ' pCommand prior to PFK prepending
      pCommandPrv = ""                                            ' pCommand previous if & prefixed
      PrevPCmd    = " "                                           ' Previous PCmd
      pfChange    = ""                                            ' Previous Change string
      pfCLen      = 0                                             ' Previous real length of change string
      pfFCol      = 0                                             ' Previous Found column
      pfFind      = ""                                            ' Previous Find string
      pfFlag      = 0                                             ' Previous Operand flags
      pfFLen      = 0                                             ' Previous real length of find string
      pfFLine     = 0                                             ' Previous Found line
      pfNotFound  = %False                                        ' Previous Last search resulted in Not Found
      pfOChange   = ""                                            ' Previous Original Change string
      PgHandle    = 0                                             ' No handle to start
      PgNumber    = 0                                             ' No Page Number to start
      Prof1       = ""                                            ' Profile display 1
      Prof2       = ""                                            ' Profile display 2
      Prof3       = ""                                            ' Profile display 3
      Prof4       = ""                                            ' Profile display 4
      Prof5       = ""                                            ' Profile display 5
      Prof6       = ""                                            ' Profile display 6
      ResetClr    = 0                                             ' Reset color requested color
  DIM S(300) AS INSTANCE LONG                                     ' Screen line to L() index crossref
      SavCurCol   = 11                                            ' Save Cursor Column
      SavCurLin   = 0                                             ' Save Cursor Line
      SBCount     = 0                                             ' Count in SBTable
  DIM SBXref(1 TO 15) AS INSTANCE LONG                            ' Xref to SB Left to right
  DIM SBTable(15) AS INSTANCE SBarEntry                           ' StatusBar messages
      ScrlAmtLast = Prf.Scroll                                    ' Last displayed scroll amount
      sCurPrio    = 0                                             ' Priority of current sCur... entries
      sCurLin     = 0                                             ' Current winning line
      sCurScrl    = 0                                             ' Current winning SCRL operand
      sCurHexl    = 0                                             ' Current winning Hex Line
      sCol        = 0                                             ' Search col
  DIM Scr(20) AS INSTANCE STRING                                  ' Screen image in text mode
      sDir        = 0                                             ' Search Dir
      SlecSCol    = 0                                             ' Select - Start Column
      SlecECol    = 0                                             ' Select - End Column
      SlecSLin    = 0                                             ' Select - Start Line
      SlecELin    = 0                                             ' Select - End Line
      sLine       = 0                                             ' Search Line
      SwapLines   = 0                                             ' Swap - # of data lines
      SwapSCol    = 0                                             ' Swap - Start Column
      SwapECol    = 0                                             ' Swap - End Column
      SwapSLin    = 0                                             ' Swap - Start Line
      SwapELin    = 0                                             ' Swap - End Line
  DIM T(100000)    AS INSTANCE STRING                             ' To hold the dynamic TEXT strings
  DIM TW(100000)   AS INSTANCE WSTRING                            ' To hold the dynamic Attr strings
      TTCtr       = 100000                                        ' Number free in T()
      TabsDflt    = ""                                            ' Default tabs
      TCtr        = 0                                             ' Touch counter
  DIM TTbl(1 TO 300) AS INSTANCE TouchEntry                       ' Touch table
  DIM TTbl(1 TO 300) AS INSTANCE TouchEntry                       ' Touch table
      LET TIP     = CLASS "cIO"                                   ' Primary Input IO
      LET TOP     = CLASS "cIO"                                   ' Primary Output IO
      TopPCmd     = ""                                            ' Top PCmd
      TMode       = Env.PMode                                     ' Tab Mode flags
      TopScrn     = 1                                             ' Start at line 1
      TopDrawn    = 0                                             ' Top Dispscreen drawn line
      UBoundL     = 100000                                        ' Current table ubound
      UBoundT     = 100000                                        ' Current table ubound
  DIM Undo(Prf.UndoNumber) AS INSTANCE UndoType                   ' Undo slot data
      UndoCurr    = 0                                             ' Current undo number
      VSaveFile   = ""                                            ' VSAVE filename if VSAVE done
      WatchFlag = " "                                             ' WatchFlag
      WindowID    = 0                                             ' No Dialog ID to start

   '----- initialize the FM Data
      NestCtr     = 0                                             ' Counter of NestTbl() entries
      NestTopScrn = 0                                             ' Desired TopScrn
      NestCRow    = 0                                             ' Desired CRow
  DIM NestTbl(1 TO 10) AS INSTANCE STRING                         ' Save table for FPath, FMask and FileListNm
      FMode       = 0                                             ' Current FM mode of the tab
      FPath       = ""                                            ' Default directory path
      FMask       = ""                                            ' Default file mask
      ScrlAmtC    = "PAGE"                                        ' Scroll Amount field (Char)
      DefCommand  = "Select"                                      ' Default LMB Command
      LFPath      = ""                                            ' Last File Path
      LFMask      = ""                                            ' Last File Mask
      LFileListNm = ""                                            ' Last FileList Name
      LScrlAmtC   = ""                                            ' Last Scroll Amount
  DIM RQList(1 TO 100) AS INSTANCE STRING                         ' Current request list
  DIM RFList(1 TO 100) AS INSTANCE STRING                         ' Current Forget List
  DIM RXList(1 TO 100) AS INSTANCE STRING                         ' Current Exclude list
  DIM RNList(1 TO 100) AS INSTANCE STRING                         ' Current Note list
  DIM AFList(1 TO 100) AS INSTANCE FMFList                        ' Latest Dir scan results
      RQCount     = 0                                             ' Count in RQList
      RFCount     = 0                                             ' Count in RFList
      RXCount     = 0                                             ' Count in RXList
      RNCount     = 0                                             ' Count in RNList
      AFCount     = 0                                             ' Count in AFList
      NumFiles    = 0                                             ' Number of Files
      NumDirs     = 0                                             ' Number of Dirs
      CmdEmpty    = " "                                           ' Cmd area for New Empty
      CmdFilePath = " "                                           ' Cmd area for Norm Path
      CmdFavorite = " "                                           ' Cmd area for Favorites
      CmdFileList = " "                                           ' Cmd area for FileList
      CmdFound    = " "                                           ' Cmd area for Found
      CmdRecent   = " "                                           ' Cmd area for Recent Files
      CmdRPaths   = " "                                           ' Cmd area for Recent Paths
      CmdOpen     = " "                                           ' Cmd area for Open Files
      CmdProfiles = " "                                           ' Cmd area for Profiles
      CallMsg     = ""                                            ' Msg to issue by FM in a CallTab invocation
      DefSort     = "Name+"                                       ' Default FM sort
      DirSort     = "Dir+"                                        ' Default Fm Dir positioning
      LastTop     = 0                                             ' TopScrn as at last display
      FMMarkdSLin = 0                                             ' Last marked screen line number
      FMMarkdALin = 0                                             ' Last marked AFList line number
      FilelistNm  = ""                                            ' Active recent file list name
      SaveCase    = ""                                            ' Profile Case Save
      SaveWord    = 0                                             ' FindWord Save
      FFSearch    = ""                                            ' FM FF search operand
      FSearch     = ""                                            ' FM Find search operand
      LastLoc     = 0                                             ' FM Last found line number
      MEditCmd    = ""                                            ' MEdit command string
  DIM LCmdOps(1 TO 100) AS INSTANCE STRING                        ' FM Line command operands
      LCmdOpsNum  = 0                                             ' FM Line command operands number
      FFCmd       = ""                                            ' FF command string
      OFrmFPath   = ""                                            ' Tab opened from FM Path variables
      OFrmFMask   = ""                                            '
      OFrmFileL   = ""                                            '

      FOR i = 1 TO 15                                             '
         SBTable(i) = gSBTable(i)                                 ' Start by Copying the global table
      NEXT i                                                      '
      t = ENV.SBLayout + ",PAD"                                   ' Ensure PAD at the end
      k = PARSECOUNT(t)                                           ' Get number of entries
      PARSE t, SBParse()                                          ' Break it into pieces
      FOR i = 1 TO k                                              ' Process each entry
         t = LSET$(SBParse(i), 6)                                 ' Make name 6 chars
         FOR j = 1 TO 15                                          ' Find name in table
            IF t = SBTable(j).SBName THEN                         ' Found it
               SBTable(j).SBActive = "Y"                          ' Mark this one active
               INCR SBCount                                       ' Bump our TP count
               SBXref(SBCount) = j                                ' Put the entry index into the Xref table
               EXIT FOR                                           ' We're done
            END IF                                                '
         NEXT j                                                   '
      NEXT i                                                      ' Onward to next ENV KW
   END METHOD                                                     '

   INTERFACE iObjTabData: INHERIT IUNKNOWN                        ' Define the interface

      '----- Setup all the Properties and Methods for the tab related data
      GProp(CRow, LONG)
      GSProp(BndText, STRING)
      GSProp(MarkLine, STRING)
      GSProp(MaskLine, STRING)
      GSProp(WordInput, STRING)
      GSProp(TabsLine, STRING)

      GSProp(AttnDo, LONG)
      GSProp(ActionCtr, LONG)
      GSProp(ActionStop, LONG)
      GSProp(cCurrent, INTEGER)
      GSProp(cfFCol, LONG)
      GSProp(cfFLen, LONG)
      GSProp(cfFLine, LONG)
      GSProp(cfNotFound, INTEGER)
      GSProp(cfSplitPt2, LONG)
      GSProp(cfHiLiteClr, LONG)
      GSProp(cfHiLiteOff, LONG)
      GSProp(cfHiLiteOn, LONG)

      GSProp(CsrLinDX, LONG)
      GSProp(CmdClear, LONG)
      GSProp(CurrPcmd, STRING)
      GSProp(CursLine, STRING)
      GSProp(CursWord, STRING)
      GSProp(errFlag, LONG)
      GSProp(errMsg, STRING)
      GSProp(ErrMsgTblC, LONG)
      GSProp(ErrMsgHlpC, LONG)
      GSProp(ErrMsgHigh, LONG)
      GSProp(FindWord, LONG)
      GSProp(FIXCtr, LONG)
      GSProp(Flag, LONG)
      GSProp(Flag2, LONG)
      GSProp(gHandle, DWORD)
      GSProp(InsLastClr, LONG)
      GSProp(hPCRE, DWORD)
      GSProp(LastLine, LONG)
      GSProp(LastPTCurs, LONG)
      GSProp(LastReal, LONG)
      GSProp(LastRulCol, LONG)
      GSProp(LineFirst, LONG)
      GSProp(LineLast,  LONG)
      GSProp(LastRulRow, LONG)
      GSProp(LocLineR, LONG)
      GSProp(mapstr_sequence_num, LONG)
      GSProp(MarkECol, LONG)
      GSProp(MarkedLine, LONG)
      GSProp(MarkELin, LONG)
      GSProp(MarkSCol, LONG)
      GSProp(MarkSLin, LONG)
      GSProp(MaxLength, LONG)
      GSProp(MEditCount, LONG)
      GSProp(Offset, LONG)
      GSProp(pCommand, STRING)
      GSProp(pCommandRaw, STRING)
      GSProp(PgHandle, LONG)
      GSProp(PgNumber, LONG)
      GSProp(Prof1, STRING)
      GSProp(Prof2, STRING)
      GSProp(Prof3, STRING)
      GSProp(Prof4, STRING)
      GSProp(Prof5, STRING)
      GSProp(Prof6, STRING)
      GSProp(PTFDisp, LONG)
      GSProp(PTLDisp, LONG)
      GSProp(SBCount, LONG)
      GSProp(SlecECol, LONG)
      GSProp(SlecELin, LONG)
      GSProp(SlecSCol, LONG)
      GSProp(SlecSLin, LONG)
      GSProp(sLine, LONG)
      GSProp(SwapSLin, LONG)
      GSProp(TopPcmd, STRING)
      GSProp(TopScrn, LONG)
      GSProp(UBoundL, LONG)
      GSProp(UBoundT, LONG)
      GSProp(WatchFlag, STRING)
      GSProp(WindowID, LONG)

      '----- Setup all the FM Properties

      GSProp(CmdEmpty, STRING)                                    ' Cmd area for New Empty
      GSProp(CmdFilePath, STRING)                                 ' Cmd area for FilePath
      GSProp(CmdFavorite, STRING)                                 ' Cmd area for Favoritte files
      GSProp(CmdFileList, STRING)                                 ' Cmd area for FileLists
      GSProp(CmdFound, STRING)                                    ' Cmd area for Found Files
      GSProp(CmdOpen, STRING)                                     ' Cmd area for Open Files
      GSProp(CmdProfiles, STRING)                                 ' Cmd area for Profiles
      GSProp(CmdRecent, STRING)                                   ' Cmd area for Recent Files
      GSProp(CmdRPaths, STRING)                                   ' Cmd area for Recent Paths
      GSProp(DefCommand, STRING)                                  ' Defalt LMB Command
      GSProp(FileListNm, STRING)                                  ' Active recent file list name
      GSProp(FMask, STRING)                                       ' Default file mask
      GSProp(FPath, STRING)                                       ' Default directory path
      GSProp(DefSort, STRING)                                     ' Default FM sort
      GSProp(RecentSort, STRING)                                  ' Default FM sort Recent
      GSProp(DirSort, STRING)                                     ' Default FM directory positioning
      GSProp(LastTop, LONG)                                       ' TopScrn as at last display
      GSProp(LFileListNm, STRING)                                 ' Last FileList Name
      GSProp(LFMask, STRING)                                      ' Last File Mask
      GSProp(LFPath, STRING)                                      ' Last File Path
      GSProp(NestCtr, LONG)                                       ' Nesting Depth
      GSProp(NestTopScrn, LONG)                                   ' Desired Topscreen when Nest is popped
      GSProp(NestCRow, LONG)                                      ' Desired CRow when Nest is popped
      GSProp(OFrmFPath, STRING)                                   ' Tab opened from FM Path variables
      GSProp(OFrmFMask, STRING)                                   '
      GSProp(OFrmFileL, STRING)                                   '
      GSProp(ScrlAmtC, STRING)                                    ' Scroll Amount field (Char)
      GSProp(TMode, LONG)                                         ' Tab mode


      METHOD Change()
      '---------- Do the Change processing
      LOCAL i, j, k, l, m, x, y, si, sj, sk, RetCode, GotColumns AS LONG, PosType AS INTEGER
      LOCAL lh, lh2, rh, lclChg, LastCase, t, lclPic AS STRING
      LOCAL cl, cl2, cr, CTxt AS WSTRING
      LOCAL AttrAsc, tClr AS WORD

         MEntry
         PosType = IIF(IsCfTop, %FChange, %Change)                ' Set positioning for TOP or not

         '----- Get common variables into local data
         i = cfFLine                                              ' Line we're working on
         j = cfFCol                                               ' Column where it was found
         k = cfFLen                                               ' Length of found string (Use orig. length)
         l = LTxtGLen(i)                                          ' Length of text line

         '----- Adjust k (Found length) if a picture type; for \ escapes and for { } boundary
         IF IsCfL1Picture THEN                                    ' If doing a picture type search
            lclPic = cfFind                                       ' Get a local copy
            IF LEFT$(lclPic, 1) = "{" THEN lclPic = MID$(lclPic, 2) ' Remove the { and } if present
            IF RIGHT$(lclPic, 1) = "}" THEN lclPic = LEFT$(lclPic, LEN(lclPic) - 1)

            m = 1: k = 0                                          '
            t = LTxtG(i): RESET si, sj                            ' Get a local text copy
            DO WHILE m <= LEN(cfFind)                             ' Get Pic true length
               IF MID$(cfFind, m, 1) = "{" AND m = 1 THEN         ' A Left boundary spaces request
                  DECR k                                          ' Alow for the {
                  k += (VERIFY(t, " ") - 1 - (VERIFY(lclPic, " ") - 1)) ' Adjust k (found length
               END IF                                             '

               IF MID$(cfFind, m, 1) = "}" AND m = LEN(cfFind) THEN ' A Right boundary spaces
                  DECR k                                          ' Allow for the }

                  IF RIGHT$(lclPic, 1) = " " THEN                 ' Trailing spaces in the picture?
                     FOR si = LEN(lclPic) TO 1 STEP -1            ' See how many
                        IF MID$(lclPic, si, 1) = " " THEN INCR sk ' Count trailing blanks
                        IF MID$(lclPic, si, 1) <> " " THEN EXIT FOR ' Exit when a non-blank
                     NEXT si                                      '
                  END IF                                          ' sk has trailing picture spaces

                  FOR si = LEN(t) TO 1 STEP -1                    ' See how many
                     IF MID$(t, si, 1) = " " THEN INCR sj         ' Find last non-blank
                     IF MID$(t, si, 1) <> " " THEN EXIT FOR       '
                  NEXT si                                         ' sj has trailing text line blank count
                  k += sj - sk                                    ' Adjust k (found length
               END IF                                             '
               INCR k                                             '
               m = m + IIF(m < LEN(cfFind) AND MID$(cfFind, m, 1) = "\", 2, 1)
            LOOP                                                  '
         END IF                                                   '

         me.ModSet(i)                                             ' Remember we changed something
         me.LFlagBitOn(i, %EQChange)                              ' Mark the line number changed
         me.UpdLControl(i)                                        ' Update LLCtl

         '----- Set color for any highlighting request

         '----- First see if any existing highlighting on the found string
         AttrAsc = ASC(MID$(LAttrG(i), j))                        ' Get Attr of 1st change byte
         IF REPEAT$(k, CHR$$(AttrAsc)) <> MID$(LAttrG(i), j, k) THEN _  ' If not the same attr for whole string
            AttrAsc = AttrAsc AND (%AttrAll - %AttrHiLite)        ' Just make it normal text

         '----- Override if specific Clr asked for
         IF cfHiLiteOn <> 0 THEN                                  ' +color asked for?
            tClr = cfHiLiteOn: SHIFT LEFT tClr, 8                 ' Add it in
            AttrAsc = (AttrAsc AND (%AttrAll - %AttrHiLite)) OR tClr '
         END IF                                                   '
         IF IsCfPStd THEN AttrAsc = (AttrAsc AND (%AttrAll - %AttrHiLite)) ' Std? Remove color

         lh = LEFT$(LTxtG(i), j - 1)                              ' Get the LH portion of the text
         rh = MID$(LTxtG(i), j + k)                               ' Get the RH portion of the text
         cl = LEFT$(LAttrG(i), j - 1)                             ' Get the LH portion of the Attributes
         cr = MID$(LAttrG(i), j + k)                              ' Get the RH portion of the Attributes
         lclChg = cfChange                                        ' Get local copy of change string

         '----- At this point we have the LH/RH parts of the Text and Attribute lines
         '-----                       and the Scheme and color defaults for the change string
         '----- Now let various options fiddle the change string

         '----- Handle the EXEC Macro type
         IF IsCfL2Exec THEN                                       ' An EXEC type?
            gMacFString = MID$(L(i).@LTxt, j, k)                  ' Setup FIND string
            gMacCString = gMacFString                             ' Setup default answer to do nothing
            pCmdMacro(lclChg)                                     ' Let MACRO command handle it
            lclChg = gMacCString:  cfChange = gMacCString         ' Copy the answer
         END IF

         '----- Handle the Mapping type change
         IF IsCfL2Map THEN                                        ' A Mapping type?
            gMacFString = MID$(L(i).@LTxt, j, k)                  ' Setup FIND string
            gMacCString = gMacFString                             ' Setup default answer to do nothing
            gMacErrString = ""                                    ' Clear error
            RetCode = mapstr_process(gMacFString, cfChange, gMacCString, gMacErrString)
            IF RetCode <> 0 THEN                                  ' Better say OK
               scError(%eFail, "CHANGE Map string error: " + gMacErrString) ' Issue error returned by MapStr
            END IF                                                '
            lclChg = gMacCString: cfChange = gMacCString          ' Copy the answer
         END IF

         '----- Handle T' type changes
         IF IsCfL2CaseInComp THEN                                 ' If replace is a T'string' type
            IF LEN(lclChg) > 0 THEN                               ' If change string is not Null
               LastCase = IIF$(ISTRUE INSTR(gUpper, LEFT$(lclChg, 1)), "U", "L")  ' Set a default LastCase
               lh = cfFind                                        '
               FOR x = 1 TO LEN(lclChg)                           ' Adjust case of the change string
                  IF LEN(cfFind) >= x THEN                        ' If Search string long enough
                     IF INSTR(gUpper, MID$(LTxtG(i), x + j - 1, 1)) OR INSTR(gLower, MID$(LTxtG(i), x + j - 1, 1)) THEN _
                        LastCase = IIF$(ISTRUE INSTR(gUpper, MID$(LTxtG(i), x + j - 1, 1)), "U", "L")  ' Set a LastCase for this char
                  END IF                                          '
                  IF LastCase = "U" THEN MID$(lclChg, x, 1) = UUCASE(MID$(lclChg, x, 1))
                  IF LastCase = "L" THEN MID$(lclChg, x, 1) = LLCASE(MID$(lclChg, x, 1))
               NEXT x                                             '
            END IF                                                '

         '----- UC it all if CAPS ON
         ELSEIF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) THEN  ' CAPS ON?
            lclChg = UUCASE(lclChg)                               ' UC if CAPS is ON
         END IF                                                   '

         '----- If this is a F' or P' type change
         IF IsCfL2Format OR IsCfL2Picture THEN                    ' If doing a format change
            lclChg = ""                                           ' Start result as null
            y = cfFCol                                            ' Found data

            '----- Handle each of the special character requests
            FOR x = 1 TO LEN(cfChange)                            ' Loop doing the special change
               t = LTxtG(i)                                       ' Get a local text copy
               SELECT CASE AS CONST$ MID$(cfChange, x, 1)         ' Do it based on the char inside P'xxx'
                  CASE "="                                        ' Copy the character
                     lclChg += MID$(t, y, 1)                      '
                     INCR y                                       '
                  CASE "<"                                        ' Lowercase it
                     lclChg += LLCASE(MID$(t, y, 1))              '
                     INCR y                                       '
                  CASE ">"                                        ' Uppercase it
                     lclChg += UUCASE(MID$(t, y, 1))              '
                     INCR y                                       '
                  CASE "!"                                        ' Entire found string
                     lclChg += MID$(t, j, k)                      '
                  CASE "«", "®"                                   ' Entire found string lowercased  "®" is "«" in OEM
                     lclChg += LLCASE(MID$(t, j, k))              '
                  CASE "»", "¯"                                   ' Entire found string uppercase   "¯" is "»" in OEM
                     lclChg += UUCASE(MID$(t, j, k))              '
                  CASE "~"                                        ' Ignore it
                     INCR y                                       '
                  CASE "\"                                        ' Escape?
                     lclChg += MID$(cfChange, x + 1, 1)           ' Copy next char
                     INCR x                                       ' extra step
                     IF IsCfL2Picture THEN INCR y                 ' If doing a picture
                  CASE "{"                                        ' Leading spaces
                     lclChg += SPACE$(VERIFY(t, " ") - 1)         '
                  CASE "}"                                        ' Trailing spaces
                     RESET si, sj                                 '
                     FOR si = LEN(t) TO 1 STEP -1                 ' See how many
                        IF MID$(t, si, 1) <> " " THEN sj = si: EXIT FOR ' Find last non-blank
                     NEXT si                                      '
                     IF sj > 0 THEN                               ' Found it
                        lclChg += SPACE$(LEN(t) - sj)             ' Add # spaces
                     END IF                                       '

                  '----- Not special character, just copy it
                  CASE ELSE                                       ' Copy L12 character
                     lclChg += MID$(cfChange, x, 1)               '
                     IF IsCfL2Picture THEN INCR y                 ' If doing a picture
               END SELECT                                         '
            NEXT x                                                '
         END IF                                                   '

         '----- OK, everyone has fiddled lclChg their own way

         '----- See if there are possible columns to try and retain
         GotColumns = %False                                      ' Reset flag
         IF IsCfCS THEN                                           ' Explicit Column Shift?
            '                                                     ' Do nothing
         ELSEIF Prf.ChangeMode = "C" AND ISFALSE IsCfDS THEN      ' Default CS and not overriden?
            '                                                     ' Do nothing
         ELSEIF Prf.ChangeMode = "D" AND IsCfCS THEN              ' Default DS but overriden by CS?
            '                                                     ' Do nothing
         ELSE                                                     ' Else check for columns
            IF LEN(TRIM$(cfFind)) <> 0 AND _                      ' Only do columns if search and replace was non-blank
               LEN(TRIM$(cfChange)) <> 0 AND _                    '
               ISFALSE IsColSupp THEN                             ' And nobody wants this suppressed
               x = INSTR(j + k, LTxtG(i), "  ")                   ' Look for at least 2 blanks following found string
               IF x THEN                                          ' If we 've got spare blanks
                  FOR y = x TO l                                  ' Look through remainder of line for a non- blank
                     IF MID$(LTxtG(i), y, 1) <> " " THEN          ' Got one?
                        GotColumns = %True                        ' Remember we have columns
                        EXIT FOR                                  ' Exit, y - 1 is split point for lh/rh
                    END IF                                        '
                  NEXT y                                          '
               END IF                                             '
            END IF                                                '
         END IF                                                   '

         '----- If there are columns do the change so we try and retain them
         IF GotColumns AND ISFALSE IsCfL2Trunc THEN               ' We've columns and no TRUNC, do things the fancy way
            lh = LEFT$(LTxtG(i), y - 2)                           ' Create full lh portion
            rh = MID$(LTxtG(i), y - 1)                            ' Create full rh portion
            lh2 = RTRIM$(MID$(lh, j + k))                         ' Create rh part of the full lh portion
            lh = LEFT$(lh, j - 1)                                 ' Create lh part of the full lf portion
            lh = lh + lclChg + lh2                                ' Re-Build the full lh portion with the change
            IF LEN(lh) < y - 1 THEN                               ' If lh is less than orig length of lh portion
               lh = lh + STRING$((y - 2) - LEN(lh), " ")          ' make it back up to original length
            END IF                                                '
            me.LTxtSet(i, lh + rh)                                ' Re-build the altered line now

            '----- Ditto for the Attr line
            cl = LEFT$(LAttrG(i), y - 2)                          ' Create full lh portion
            cr = MID$(LAttrG(i), y - 1)                           ' Create full rh portion
            cl2 = RTRIM$(MID$(cl, j + k))                         ' Create rh part of the full lh portion
            cl = LEFT$(cl, j - 1)                                 ' Create lh part of the full lf portion
            cl = cl + REPEAT$(LEN(lclChg), CHR$$(AttrAsc)) + cl2  ' Re-Build the full lh portion with the change
            IF LEN(cl) < y - 1 THEN                               ' If lh is less than orig length of lh portion
               cl = cl + REPEAT$((y - 2) - LEN(cl), CHR$$(AttrAsc)) ' make it back up to original length
            END IF                                                '
            LAttrS(i) = cl + cr                                   ' Re-build the Attr line
            me.AttrScan(i)                                        ' Recolorize

         '----- Just a simple non-column change
         ELSE                                                     ' We've a plain simple change to do
            IF IsCfL2Trunc THEN                                   ' TRUNC requested?
               me.LTxtSet(i, lh + lclChg)                         ' Re-build the line minus the rh portion
               CTxt = cl + REPEAT$(LEN(lclChg), CHR$$(AttrAsc))   ' Re-build the Attr line
            ELSE                                                  ' Else no TRUNC requested?
               me.LTxtSet(i, lh + lclChg + rh)                    ' Re-build the altered line
               CTxt = cl + REPEAT$(LEN(lclChg), CHR$$(AttrAsc)) + cr ' Re-build the Attr line
            END IF                                                '
            LAttrS(i) = Ctxt                                      ' Store the Attr line
            me.AttrScan(i)                                        ' Recolorize
         END IF                                                   '

         IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) THEN      ' CAPS ON?
            me.LTxtSet(i, UUCASE(LTxtG(i)))                       ' UC if CAPS is ON
         END IF                                                   '

         '----- Figure out if and what to HiLite
         IF ISTRUE Prf.HiFind THEN                                ' See whether to hilite the string
            me.AttrInvSet(i, j, j + LEN(lclChg) - 1)              ' Set the Attr flags
         END IF                                                   '

         '----- Do cursor positioning
         IF ISTRUE Prf.HiFind THEN                                ' See whether to hilite the string
            IF IsLInvisible(i) THEN                               ' Still invisible (DX is active), else it'd be popped out
               me.CurSetReq(PosType, i, j + LEN(lclChg) - 1, %True, %True)  ' Put cursor on the X line
               CsrLinDX = VAL(LLNumG(i))                          ' Save hidden line number
            ELSE                                                  ' Not invisible
               IF sDir = 1 THEN                                   ' Forward?
                  me.CurSetReq(PosType, i, j + LEN(lclChg) - 1, %True) ' Set cursor set attempt
               ELSE                                               '
                  me.CurSetReq(PosType, i, j - 1, %True)          ' Set cursor set attempt
               END IF                                             '
            END IF                                                '

         '----- No Hi-Lite required
         ELSE                                                     ' Else no hilite
            IF sDir = 1 THEN                                      ' Forward?
               me.CurSetReq(PosType, i, j + LEN(lclChg) - 1, %True) ' Set cursor set attempt
            ELSE                                                  '
               me.CurSetReq(PosType, i, j - 1, %True)             ' Set cursor set attempt
            END IF                                                '
         END IF                                                   '

         '----- Set resume search parameters
         IF sDir = 1 THEN                                         ' Forward search?
            sCol += LEN(lclChg) - 1                               ' Adjust continue search col for picture types
            IF IsCfLeft THEN sCol = LEN(L(i).@LTxt)               ' LEFT? Go to end of line
         ELSE                                                     '
            sCol = (j - 1) - LEN(L(sLine).@LTxt)                  ' No, calc new scan start
            IF IsCfLeft THEN                                      ' LEFT going backward?
               sLine -= 1                                         ' 'next' line
               sCol = -1                                          '
            END IF                                                '
         END IF                                                   '
         MExit                                                    '

      END METHOD

      METHOD CleanInsLines()
      '---------- Remove unused Insert Lines
      REGISTER i AS LONG

         MEntry
         '----- Skip if someone wants a one time pass on this
         IF IsInsClnSupp THEN OffInsClnSupp: MExitMeth            ' Suppress once

         '----- Don't waste time if nobody asked for it at all
         IF InsCount = 0 OR LastLine < 3 THEN MExitMeth           ' Don't scan if not needed

         '----- OK, grunt work to delete the unused '''''' lines
         ARRAY SORT InsTbl() FOR InsCount                         ' Sort into line number order
         FOR i = InsCount TO 1 STEP -1                            ' Loop through Insert table
            IF IsLInsertLine(InsTbl(i)) THEN                      ' An unused Insert Line?
               me.LTxtFree(InsTbl(i))                             ' Yes, Go free the dynamic string
               me.LEntDel(InsTbl(i))                              ' Remove from the L() array
               DECR LastLine: DECR LastReal                       ' Adjust Last Line
            END IF                                                '
         NEXT i                                                   '
         InsCount = 0                                             ' Say we have none
         MExit
      END METHOD

      METHOD CleanTabData()
      '---------- Clean up Edit tab data areas
      LOCAL i, j AS LONG, fn AS STRING
         IF ISNOTNULL(VSaveFile) THEN                             ' A VSAVE file to delete?
            sRecycleBin(VSaveFile, "D")                           ' Then go do it
         END IF                                                   '

         IF IsMedit THEN                                          ' If a MEdit clear each filename
            FOR j = MEditCount TO 1 STEP -1                       ' Loop in reverse (Medit table shrinks as deleted)
               fn = me.MEditListGet(j)                            ' Get a filename
               me.UnWatchQueue(fn)                                ' Kill Watch and dequeue
               me.MeditTbl("D", fn)                               ' Remove from Medit table
            NEXT j                                                '

         ELSE                                                     ' Else the simple one
            me.UnWatchQueue("")                                   ' Kill the current file
         END IF                                                   '

         me.LInitTxtData("")                                      ' Wipe everything out then

         IF UBOUND(Undo) > 0 THEN                                 ' If any UNDO table, get rid of it
            FOR i = 1 TO UBOUND(Undo())                           ' Delete temporary files
               TRY
                  IF ISNOTNULL(TRIM$(Undo(i).UFn)) THEN KILL TRIM$(Undo(i).UFn)
                  IF ISNOTNULL(TRIM$(Undo(i).TFn)) THEN KILL TRIM$(Undo(i).TFn)
                  IF ISNOTNULL(TRIM$(Undo(i).IXFn)) THEN KILL TRIM$(Undo(i).IXFn)
               CATCH                                              '
               END TRY                                            '
            NEXT i                                                '
         END IF                                                   '

      END METHOD

      METHOD ClrKWSrch (sWord AS STRING) AS LONG
      LOCAL Upper, Lower, iPos AS LONG, srch AS STRING
         srch = IIF$(IsClrCase, TRIM$(sWord), UUCASE(TRIM$(sWord))) ' Set case correctly
         METHOD = %SCTxtLo                                        ' Default fail answer
         Lower = 1 : Upper = ClrKwrdNum                           ' Set starting

         IF srch = ClrKW(Lower) THEN METHOD = ClrKWAns(Lower) : EXIT METHOD
         IF srch = ClrKW(Upper) THEN METHOD = ClrKWAns(Upper) : EXIT METHOD
         IF srch < ClrKW(Lower) THEN EXIT METHOD                  ' Exit with the %SCTxtHi default
         IF srch > ClrKW(Upper) THEN EXIT METHOD                  '

         DO UNTIL (Upper <= (Lower + 1))                          ' OK, zero in on it
            iPos = (Lower + Upper) / 2                            '
            IF srch > ClrKW(iPos) THEN                            '
               Lower = iPos                                       '
            ELSEIF srch < ClrKW(iPos) THEN                        '
               Upper = iPos                                       '
            ELSE                                                  '
               METHOD = ClrKWAns(iPos) : EXIT METHOD              '
            END IF                                                '
         LOOP                                                     '
      END METHOD

      METHOD CmdStackParse (CmdString AS STRING)                  ' Break command string down and save it
      LOCAL tCmd, lcmd AS STRING, c, i, j, lastchr AS LONG
         MEntry
         CmdStackNum = 0                                          ' Start with none stacked
         tCmd = TRIM$(cmdstring)                                  ' Working copy
         i = TALLY(tCmd, ENV.CmdChr)                              ' See how many Command delimiters
         IF i > UBOUND(CmdStack) THEN _                           ' Expand CmdStack if needed
            REDIM PRESERVE CmdStack(i + 2)                        ' Get room for our count plus some slack

         IF i = 0 THEN MExitMeth                                  ' If no stacking, we're done

         lastchr = 0: c = 0                                       '
         FOR i = 1 TO LEN(tCmd)                                   ' Process stacked commands
            IF MID$(tCmd, i, 1) = CHR$(34) THEN                   ' Skip quoted
               WHILE i <= LEN(tCmd)                               '
                  INCR i                                          '
                  IF MID$(tCmd, i, 1) = CHR$(34) THEN EXIT LOOP
               WEND                                               '
               INCR i                                             ' Look for trailing
            END IF                                                '
            IF MID$(tCmd, i, 1) = "'" THEN                        ' Skip quoted
               WHILE i <= LEN(tCmd)                               '
                  INCR i                                          '
                  IF MID$(tCmd, i, 1) = "'" THEN EXIT LOOP
               WEND                                               '
               INCR i                                             ' Look for trailing
            END IF                                                '
            IF MID$(tCmd, i, 1) = "`" THEN                        ' Skip ` quoted
               WHILE i <= LEN(tCmd)                               '
                  INCR i                                          '
                  IF MID$(tCmd, i, 1) = "`" THEN EXIT LOOP
               WEND                                               '
               INCR i                                             ' Look for trailing
            END IF                                                '

            IF MID$(tCmd, i, 1) = ENV.CmdChr THEN                 ' Got a ENV.CmdChr
               lcmd = MID$(tCmd, lastchr + 1, i - lastchr - 1)    ' Extract command
               INCR c                                             '
               IF c = 1 THEN                                      ' 1st one?
                  pCommand = lcmd                                 ' Set it in pCommand
               ELSE                                               ' Else
                  CmdStack(c-2) = lcmd                            ' in the CmdStack array
                  INCR CmdStackNum                                ' Count it
               END IF                                             '
               lastchr = i                                        ' remember where last ENV.CmdChr was
            END IF                                                '
         NEXT i                                                   '
         IF i <> lastchr THEN                                     ' If we didn't end with a ;
            IF c > 0 THEN                                         ' If we're stacking, add another, else leave pCommand alone
               lcmd = MID$(tCmd, lastchr + 1)                     ' Extract last part
               CmdStack(CmdStackNum) = lcmd                       ' in the CmdStack array
               INCR CmdStackNum                                   ' Count it
            END IF                                                '
         END IF                                                   '
         MExit
      END METHOD

      METHOD ClrLoad() AS LONG
      '---------- Load Colourize values for a named file extension
      LOCAL i, j, k, ClrOpVal, ClrOpTyp, b, g, r, Upper, LineNumber, lclScheme, DeprFlag AS LONG
      LOCAL Txt1, lclext, ClrOpStr, t AS STRING
      LOCAL AliasOnce AS INTEGER
      DIM parsed(100) AS STRING
      DIM Numbers(1 TO 10) AS STRING
      LOCAL ClrIO AS iIO                                          ' For our I/O stuff

         MEntry
         IF ISFALSE Prf.HiAuto THEN MExitMeth                     ' Don't bother if no HiAuto
         LET ClrIO = CLASS "cIO"                                  '
         AliasOnce = %True                                        ' Set 1 time flag

         '----- First ensure we do the correct profile
         lclext = IIF$(ISNULL(Prf.UseName), TIP.Profile, Prf.UseName) ' Pick up correct Profile name
         GOSUB DoAClrFile                                         ' Go process a Attr file
'         IF DeprFlag THEN                                         ' Remind User
'            sDoMsgBox "The |KAUTO|B file used for this Profile contains" + $CRLF + _
'                   "Obsoleted KW values, which were ignored." + $CRLF + $CRLF + _
'                   "The |KAUTO|B file should be Edited to conform to the" + $CRLF + _
'                   "latest SPFLite release." + $CRLF + $CRLF + _
'                   "The provided |KAUTOCNVT|B Macro will perform the needed changes", _
'                   %MB_OK OR %MB_USERICON, "AUTO file conversion"
'         END IF
         MExitMeth                                                '

      DoACLRFile:
         '----- Set up a bunch of dummy values
         OffClrFlag                                               ' Colourize inactive
         OffClrCase                                               ' No MIXEDCASE
         ClrQuoted = 0                                            ' Setup dummy values
         ClrNumeric = 0                                           '
         FOR i = 1 TO 9                                           ' Reset the comment control
            ClrComment(i).Txt1 = ""                               '
            ClrComment(i).sCol = ""                               '
            ClrComment(i).clr = 0                                 '
            ClrComment(i).fnd = 0                                 '
            ClrExclude(i).Txt1 = ""                               '
            ClrExclude(i).sCol = 0                                '
            ClrInclude(i).Txt1 = ""                               '
            ClrInclude(i).sCol = 0                                '
         NEXT                                                     '
         ClrDLM = " "    :    ClrESC = ""     :   ClrKwrdNum = 0  '

         '----- If .AUTO file doesn't exist, just exit
         IF ISFALSE ISFILE(ENV.AutoPath + lclext + ".AUTO") THEN MExitMeth  ' See if Colorize file exists

         '----- File present, process it
         ClrIO.Setup("IE", "", "", ENV.AutoPath + lclext + ".AUTO") ' Tell IO what we're opening
         ClrIO.EXEC                                               '
         DO WHILE ISFALSE EOF(ClrIO.FNum)                         ' Read the data
            LINE INPUT # ClrIO.FNum, Txt1                         ' Get a line
            INCR LineNumber                                       ' Incr linenumber for messages
            OnClrFlag                                             ' Somethings here
            IF LEN(TRIM$(Txt1)) > 0 AND LEFT$(Txt1, 1) <> ";" THEN ' Process the statement

               '----- If line has data and not a comment line, parse it
               PARSE Txt1, parsed(), " "                          '

               '----- Since we use ' ' delimiters, eliminate multi-blank spans
               i = 0: j = 0                                       ' Shrink unused entries
               DO WHILE j <= 100                                  '
                  IF ISNULL(parsed(i)) THEN ARRAY DELETE parsed(i) ELSE INCR i
                  INCR j                                          '
               LOOP                                               '

               '----- Handle each command type
               SELECT CASE AS CONST$ UUCASE(parsed(0))            ' Do different types

                  '----- SCHEME command
                  CASE "SCHEME"                                   ' SCHEME
                     DeprFlag = %True                             ' Deprecated, remember that

                  '----- NORMHI command
                  CASE "NORMHI"                                   ' NORMHI
                     DeprFlag = %True                             ' Deprecated, remember that

                  '----- NORMLO command
                  CASE "NORMLO"                                   ' NORMLO
                     DeprFlag = %True                             ' Deprecated, remember that

                  '----- LINEHI command
                  CASE "LINEHI"                                   ' LINEHI
                     DeprFlag = %True                             ' Deprecated, remember that

                  '----- LINELO command
                  CASE "LINELO"                                   ' LINELO
                     DeprFlag = %True                             ' Deprecated, remember that

                  '----- NUMERIC command
                  CASE "NUMERIC"                                  ' NUMERIC, added, default set to 0
                     IF parsed(1) = "" THEN GOSUB KillClr         ' Enough params?
                     GOSUB ValidateScheme                         ' Make sure Scheme number is valid
                     ClrNumeric = lclScheme                       ' Save Scheme number

                  '----- EXCLUDE command
                  CASE "EXCLUDE"                                  ' EXCLUDE
                     IF parsed(2) = "" THEN GOSUB KillClr         ' Enough params?
                     j = 0                                        ' Reset found index
                     FOR i = 1 TO 9                               ' Look for an unused slot
                        IF TRIM$(ClrExclude(i).Txt1) = "" THEN j = i: EXIT FOR
                     NEXT i                                       '
                     IF j = 0 THEN GOSUB KillClr                  ' All full?  Kill it
                     ClrExclude(j).Txt1 = UUCASE(parsed(2))       ' Save the text
                     ClrExclude(j).sCol = VAL(parsed(1))          ' Column No

                  '----- INCLUDE command
                  CASE "INCLUDE"                                  ' INCLUDE
                     IF parsed(2) = "" THEN GOSUB KillClr         ' Enough params?
                     j = 0                                        ' Reset found index
                     FOR i = 1 TO 9                               ' Look for an unused slot
                        IF TRIM$(ClrInclude(i).Txt1) = "" THEN j = i: EXIT FOR
                     NEXT i                                       '
                     IF j = 0 THEN GOSUB KillClr                  ' All full?  Kill it
                     ClrInclude(j).Txt1 = UUCASE(parsed(2))       ' Save the text
                     ClrInclude(j).sCol = VAL(parsed(1))          ' Column No

                  '----- DELIMS command
                  CASE "DELIMS"                                   ' DELIMS
                     IF parsed(1) = "" THEN GOSUB KillClr         ' Enough params?
                     ClrDLM = parsed(1) + " "                     ' Dlm text + a blank

                  '----- ESCAPECHR command
                  CASE "ESCAPECHR"                                ' ESCAPECHR
                     IF parsed(1) = "" THEN GOSUB KillClr         ' Enough params?
                     ClrESC = parsed(1)                           ' ESC text

                  '----- MIXEDCASE command
                  CASE "MIXEDCASE"                                ' MIXEDCASE
                     IF parsed(1) = "" THEN GOSUB KillClr         ' Enough params?
                     IF IsEQ(parsed(1), "YES") THEN ONClrCase: EXIT SELECT ' Remember to honour case
                     IF IsEQ(parsed(1), "NO") THEN EXIT SELECT    ' False is the default

                  '----- WORD / AUTOCAPS command
                  CASE "WORD", "AUTOCAPS"                         ' WORD or AUTOCAPS?
                     IF parsed(2) = "" THEN GOSUB KillClr         ' Enough params?
                     GOSUB ValidateScheme                         ' Make sure Scheme number is valid
                     IF UUCASE(parsed(0)) = "AUTOCAPS" THEN lclScheme += 100 ' Indicate uppercase request
                     ClrKWrdNum += 1                              ' Count word
                     IF ClrKWrdNum > UBOUND(ClrKW) THEN           ' Expand tables if needed
                        REDIM PRESERVE ClrKW(0 TO ClrKWrdNum + 500) AS INSTANCE STRING
                        REDIM PRESERVE ClrKWAns(0 TO ClrKWrdNum + 500) AS INSTANCE LONG
                     END IF                                       '
                     ClrKw(ClrKWrdNum) = IIF$(IsClrCase, parsed(2), UUCASE(parsed(2))) ' Add it, uppercase if needed
                     ClrKwAns(ClrKWrdNum) = lclScheme             '

                  '----- QUOTED command
                  CASE "QUOTED"                                   ' Quoted strings
                     IF parsed(1) = "" THEN GOSUB KillClr         ' Enough params?
                     GOSUB ValidateScheme                         ' Make sure Scheme number is valid
                     ClrQuoted = lclScheme                        ' Scheme number

                  '----- Error for anything else if not COMMENTx
                  CASE ELSE                                       ' Better never get here
                     '----- Try a COMMENT? command
                     IF LEFT$(UUCASE(parsed(0)), 7) <> "COMMENT" THEN GOSUB KillClr ' If not a COMMENTx statement
                     IF parsed(3) = "" THEN GOSUB KillClr         ' Enough params?
                     IF INSTR("123456789", MID$(parsed(0), 8, 1)) = 0 THEN GOSUB KillClr' Make sure 1-9
                     GOSUB ValidateScheme                         ' Make sure Scheme number is valid
                     i = VAL(MID$(parsed(0), 8, 1))               ' Extract the COMMENT number
                     ClrComment(i).Clr = lclScheme                ' Scheme index
                     ClrComment(i).Txt1 = UUCASE(parsed(2))       ' Comment1 text
                     ClrComment(i).sCol = parsed(3)               ' Column No / Dlm
               END SELECT                                         '
            END IF                                                '
         LOOP                                                     '
         ClrIO.Close                                              ' Close file

         '----- Sort now for Binary Search to work
         ARRAY SORT ClrKW(1) FOR ClrKWrdNum, TAGARRAY ClrKWAns()  '
         RETURN                                                   '

      ValidateScheme:
         lclScheme = VAL(parsed(1))                               ' Get the scheme number
         IF lclScheme < 1 OR lclScheme > 14 THEN GOSUB KillClr    ' Invalid scheme number kill it
         lclScheme += 1                                           ' Adjust scheme up past TxtLo and TxtHi
         RETURN                                                   '

      KillClr:
         OffClrFlag                                               ' Turn off the flag
         sDoMsgBox ClrIO.FilePath + $CRLF + "File contains an invalid statement at Line: " + FORMAT$(LineNumber) + $CRLF + _
                "|K" + Txt1 + $CRLF + "|BAuto Colorize has been turned off", %MB_OK OR %MB_USERICON, "SPFLite"
         ClrIO.Close                                              ' Close the file
         MExitMeth
         RETURN
      END METHOD


      METHOD CmdAssign()
      '---------- Process the pCommand line
      LOCAL i AS LONG, s, t, u AS STRING
      LOCAL CallAddr AS DWORD
         MEntry
         IF ISNOTNULL(TRIM$(pCommand)) THEN                       ' Something in the line?

            '----- Handle any alias commands
            s = pCommand                                          ' Get working copy of command line
            t = UUCASE(sGetWord(s, %Strip, %QuoteNotSig))         ' Get the basic command name

            '----- Check FM SELECT command fudge
            IF t = "S" OR t = "SEL" OR t = "SELECT" THEN          ' The magic word?
               t = DefCommand                                     ' Swap the command name
               pCommand = DefCommand + " " + s                    ' Rebuild command line

            ELSE                                                  '
               u = sSetTable("GET", "ALIAS." + t)                 ' See if an alias SET symbol
               IF VAL(LEFT$(u, 1)) = 0 THEN                       ' Found?
                  pCommand = MID$(u, 2) + " " + s                 ' Re-build a new command string
               END IF                                             '
            END IF                                                '

            '----- Now process the command
            PrevPCmd = CurrPcmd                                   ' Push back commands
            CurrPcmd = UUCASE(sGetWord(pCommand, %NoStrip, %QuoteNotSig)) ' Get the basic command name
            TopPCmd = CurrPCmd                                    ' Save top level command name
            IF CurrPCmd <> "H" AND CurrPCmd <> "HELP" THEN RESET ErrMsgHlpC, ErrMsgHlp()
            i = PCmdT.GetCmdIX(CurrPcmd)                          ' Get the command table index?

            IF i THEN                                             ' We found it
               gCrashLastPCmd = PCommand                          ' Last Primary command
               IF gMacroMode THEN                                 ' In macro mode?
                  IF ISFALSE PCmdT.IsMacOK(i) THEN                ' If not allowable in macro mode
                     scError(%eFail, CurrPcmd + " command not supported in Macro mode") ' Set error message
                     MExitmeth                                    ' and leave
                  END IF                                          '
               END IF                                             '

               IF IsFMTab THEN                                    ' If this is the FM tab
                  IF ISFALSE PCmdT.IsFMAllowed(i) THEN            ' If not allowable in FM Mode
                     MErrExit(%eFail, "Command not supported in FM tab") ' Set error message
                  ELSE                                            '
                     CallAddr = PCmdT.GetFMCode(i)                ' Get Routine address
                     CALL DWORD CallAddr USING pCmdModel(pCommand)' Call the routine
                  END IF                                          '

               ELSE                                               ' Not FM mode
                  IF ISFALSE PCmdT.IsEditAllowed(i) THEN          ' If not allowable in Edit Mode
                     MErrExit(%eFail, "Command only supported in FM tab")' Set error message
                  ELSE                                            '
                     IF IsBrowse THEN                             ' If Browse, see if allowable
                        IF ISFALSE PCmdT.IsBrowseOK(i) THEN _     ' If not allowable in BROWSE mode
                           MErrExit(%eFail, "Command not allowed in Browse mode")' Set error message
                     END IF                                       '

                     IF ErrFlag > %eRetrieve THEN                 ' Probable pending?
                        IF PCmdT.IsScrollOK(i) THEN               ' If scrolling type command?
                           me.MarkKill                            ' Do command anyway
                           IF PCmdT.IsChainOK(i) THEN INCR PCmdIX ' Bump Command index if Chainable
                           sCurPrio = 0                           ' Allow cursor requests for the new command
                           CallAddr = PCmdT.GetEditCode(i)        ' Get Routine address
                           CALL DWORD CallAddr USING pCmdModel(pCommand) ' Call the routine
                        END IF                                    '
                     ELSE                                         '
                        me.MarkKill                               ' Kill any active block select
                        IF PCmdT.IsChainOK(i) THEN INCR PCmdIX    ' Bump Command index if Chainable
                        sCurPrio = 0                              ' Allow cursor requests for the new command
                        CallAddr = PCmdT.GetEditCode(i)           ' Get Routine address
                        CALL DWORD CallAddr USING pCmdModel(pCommand) ' Call the routine
                     END IF                                       '
                  END IF                                          '
               END IF                                             '

            ELSE                                                  '
               IF ISFALSE IsFMTab AND IsMacro(CurrPCmd) THEN      ' Not FM and a macro file?
                  me.pCmdMacro(pCommand)                          ' Go do it
               ELSE                                               ' No
                  scError(%eFail, "Unknown command: " + CurrPcmd) ' Set error message
                  MExitMeth                                       ' and leave
               END IF                                             '
            END IF                                                '

            IF TabsNum = 0 THEN MExitMeth                         ' Prevent Memory Access violations at termination
         END IF                                                   '
         MExit                                                    '
      END METHOD

      METHOD CmdParse(pCmd AS STRING) AS LONG
      '---------- Do basic Command parsing
      LOCAL c AS STRING, i, j AS LONG
         MEntry
         RESET pCmdOps(), pCmdRaw(), pCmdOpsType()                ' Clear the scan arrays

         '----- Separate the command line into operands
         c = pCmd                                                 ' Get a copy to play with
         DO WHILE LEN(c)                                          ' Break out the operands
            IF i = UBOUND(pCmdOps()) THEN                         ' Time to expand?
               REDIM PRESERVE pCmdOps(0 TO 2 * UBOUND(pCmdOps())) AS STRING
               REDIM PRESERVE pCmdRaw(0 TO 2 * UBOUND(pCmdRaw())) AS STRING
               REDIM PRESERVE pCmdOpsType(0 TO 2 * UBOUND(pCmdOpsType())) AS LONG
            END IF                                                '
            pCmdOps(i) = sGetWord(c, %Strip, %QuoteSig)           '
            INCR i                                                '
         LOOP                                                     '

         '----- 1st operand is the Cmd, pre-process all the other operands
         pCmdNumOps = i - 1                                       ' Operands are in array(1) thru ... 0 = No Operands

         IF pCmdNumOps > 0 THEN                                   ' Got operands?
            j = 1                                                 '
            DO WHILE j <= pCmdNumOps                              ' Examine for @ | # shortcuts
               IF pCmdOps(j) = "@" THEN                           ' Column range?
                  IF IsSlecActive THEN                            ' Active Slec?
                     pCmdOps(j) = FORMAT$(SlecSCol)               ' Replace with start col
                     ARRAY INSERT pCmdOps(j + 1), FORMAT$(SlecECol) ' Insert the end column
                     INCR pCmdNumOps                              ' Bump operand count
                  ELSE                                            '
                     ARRAY DELETE pCmdOps(j)                      ' Delete the operand
                     DECR pCmdNumOps                              '
                  END IF                                          '
               ELSEIF pCmdOps(j) = "#" THEN                       ' Line range?
                  IF IsSlecActive THEN                            ' Active Slec?
                     pCmdOps(j) = "." + FORMAT$(SlecSLIN)         ' Replace with start line
                     ARRAY INSERT pCmdOps(j + 1), "." + FORMAT$(SlecELin) ' Insert the end column
                     INCR pCmdNumOps                              ' Bump operand count
                  ELSE                                            '
                     ARRAY DELETE pCmdOps(j)                      ' Delete the operand
                     DECR pCmdNumOps                              '
                  END IF                                          '
               END IF                                             '
               INCR j                                             '
            LOOP                                                  '
         END IF                                                   '

         IF sCmdOpType THEN METHOD = %True: MExitMeth             ' Go preprocess the operands

         '----- Find the command (operand 0) in the command table
         i = PCmdT.GetCmdIX(pCmdOps(0))                           ' Get the command table index?

         '----- Validate the number of operands
         IF i THEN                                                ' We found it
            IF pCmdNumOps > PCmdT.GetNumOps(i) THEN               ' Do 'too many ops' test
               scError(%eFail, "Excessive operands for the " + pCmdOps(0) + " command") ' Set error message
               METHOD = %True: MExitMeth                          ' and leave with error posted
            END IF                                                '
         END IF                                                   '
         CurrPCmd = UUCASE(pCmdOps(0))                            ' Save as CurrPCmd in case entered from FM or such
         MExit
      END METHOD

      METHOD CmdAddStack() AS INTEGER
      '---------- Break apart pCommand to multiple commands       '
      LOCAL i, j AS LONG, tCmd, t, tx1, tx2, raw AS STRING        '
         MEntry
         METHOD = %False                                          ' Indicate it's not a RESET command pending
         IF me.SetSubst THEN MExitMeth                            ' Do SET substitution

         IF ISFALSE IsPFKInsert THEN                              ' Save for RETRIEVE if not built by PFK
            IF LEN(TRIM$(pCommand)) >= ENV.MinRetrieve THEN       ' Long enough to save?
               t = UUCASE(sGetWord(pCommand, %NoStrip, %QuoteSig))' Get command name
               IF t <> "RETRIEVE"  AND t <> "CRETRIEV" AND t <> "RETF" AND t <> "ISCROLL" THEN
                  IF gCmdRtrev(1) <> TRIM$(pCommand) THEN         ' If not the same as the current top of stack
                     ARRAY INSERT gCmdRtrev(1), TRIM$(pCommand)   ' Save new entry for Retrieve command
                  END IF                                          '
                  gCmdRtrevIX = 1                                 ' Even if Dup. of the top, Reset retrieve back to start of stack
               END IF                                             '
            END IF                                                '
         ELSE                                                     ' PFKIFlag is on
            IF LEN(TRIM$(pCommandRaw)) >= ENV.MinRetrieve THEN    ' Long enough to save?
               t = UUCASE(sGetWord(pCommand, %NoStrip, %QuoteSig))' Get command name
               IF t <> "RETRIEVE"  AND t <> "CRETRIEV" AND t <> "RETF" AND t <> "ISCROLL" THEN
                  IF gCmdRtrev(1) <> TRIM$(pCommandRaw) THEN      ' If not the same as the current top of stack
                     ARRAY INSERT gCmdRtrev(1), TRIM$(pCommandRaw)' Save new entry for Retrieve command
                  END IF                                          '
                  gCmdRtrevIX = 1                                 ' Even if Dup. of the top, Reset retrieve back to start of stack
               END IF                                             '
            END IF                                                '
         END IF                                                   '

'         '----- Do symbol Substitution
'         IF ISNOTNULL(pCommand) THEN                              '
'            tx2 = pCommand                                        '
'            tx1 = sSymbolSub(tx2)                                 ' Do symbol substitution
'            IF VAL(LEFT$(tx1, 1)) <> 0 THEN                       ' Subst. error?
'               scError(%eFail, MID$(tx1, 2))                      '
''               DoSet(%Refresh)                                     '
'               MExitMeth                                          ' Kill it
'            END IF                                                '
'            pCommand = tx2                                        ' Swap in substituted command
'         END IF                                                   '

         me.CmdStackParse(pCommand)                               ' Go parse out to the command stack / pCommand
         tCmd = UUCASE(sGetWord(pCommand, %NoStrip, %QuotenotSig))' Get command name

         IF tCmd = "RESET" OR tCmd = "RES" OR tCmd = "CAN" OR tCmd = "CANCEL" THEN ' A RESET or CANCEL command?
            METHOD = %True                                        ' Say we 've got a RESET command pending
         END IF                                                   '
         pCommandRaw = ""                                         ' Null the raw command
         MExit                                                    '
      END METHOD

      METHOD CmdStackNext () AS STRING                            ' Fetch next stacked command
      LOCAL j AS LONG, s AS STRING
         MEntry
         IF CmdStackNum < 1 THEN                                  ' If none (just in case)
            METHOD  = ""                                          ' Stuff blanks
            CmdStackNum = 0                                       '
            MExitMeth                                             ' And exit
         END IF                                                   '
         s = CmdStack(0)                                          ' Stuff into Command
         FOR j = 0 TO CmdStackNum                                 ' Shift the stack
            SWAP CmdStack(j), CmdStack(j+1)                       '
         NEXT j                                                   ' loop
         DECR CmdStackNum                                         ' Reduce remaining count
         METHOD = s                                               ' Pass back the result
         MExit
      END METHOD

      METHOD CopyAFile(lin AS LONG, ZIP AS iIO, ZPrf AS iProf, lclPM AS LONG, frmlin AS LONG, tolin AS LONG, quick AS LONG, OPTIONAL CBLines AS STRING)
      '---------- Read file and insert after passed line pointer
      LOCAL ln, i, j, k, fnum, lclLower, lclUpper, effLRECL, lctr, NoteLines, UTFError1, UTFError2 AS LONG
      LOCAL IsCopy, OverConf, StateFrom, StateTo, PCtr AS LONG
      LOCAL lclSource,lfn, chunk, leftover, dlm, t, tt, c AS STRING
      LOCAL chunklen, bytesleft, bytescopied, bytesinfile AS QUAD
      LOCAL OneTime, PageFlag, LastPage, Overprint AS INTEGER
      LOCAL lTxtP AS STRING POINTER
      LOCAL C1, C2 AS BYTE POINTER
      LOCAL tmpEOL AS STRING
      LOCAL FD AS DIRDATA
      LOCAL cpIO AS iIO
         MEntry
         t = ZPrf.Profname
         leftover = ""                                            ' Init leftover
         LET cpIO = CLASS "cIO"                                   '

         cpIO.Setup("B", ZIP.Profile, "", ZIP.FilePath)           ' Setup for open
         IF cpIO.EXEC THEN _                                      ' Go do OPEN, True if error
            MErrExit(%eFail, cpIO.ResultMsg)                      ' Tell user of error

         '----- If a full OPEN (not a COPY) extract and setup some file data
         IF CurrPcmd <> "COPY" THEN                               ' If not the COPY command
            CHDIR cpIO.Path                                       ' Make this the 'current' directory
            sIniSetString("General", "LastPath", cpIO.Path)       ' Save data in the INI file
            me.WindowTitle                                        ' Alter window title
         END IF                                                   '

         '----- Get the Profile loaded
         ZPrf.SetProfName(ZIP.Profile, %False)                    ' Set the Profile name needed
         ZPrf.ReadAll(%True)                                      ' Go read this file type's Profile values
         IF ZPrf.Skip THEN GOTO VDeblockDone                      ' Told to skip the filetype (FF mode)
         IF CurrPcmd <> "COPY" AND _                              ' If not the COPY command
            ISFALSE quick THEN                                    ' Not quick mode
            me.ClrLoad                                            ' Go load Colorize data
         END IF                                                   '

         '----- Init for the deblocking loop
         lclSource = ZPrf.PSource                                 ' Default to Profile's SOURCE value
         bytesleft = LOF(cpIO.FNum): bytesinfile = bytesleft      ' Get # bytes in file
         IF bytesleft > 3000000 THEN gLoopCtr = - 1               ' Prevent loop detection for big files
         ln = lin: StateFrom = Lin + 1                            ' Copy requested insertion point, remember for STATE
         tmpEOL = "CRLF": dlm = $CRLF                             ' Set default
         IF LEFT$(ZPrf.EOL, 4) = "AUTO" THEN PageFlag = %True     ' Force a =PAGE> at the top for AUTO files

         IF LEFT$(ZPrf.RECFM, 1) = "V" THEN                       ' If a RECFM=V type do it all differently
            GOSUB SetSource                                       ' Go set lclSource

            '----- Variable RDW type deblocking
            DO WHILE bytesleft > 0                                ' While we have data
               IF gfInterrupt THEN EXIT DO                        ' Break during FF search

               GET$ # cpIO.FNum, 4, Chunk                         ' Get in the RDW
               bytesleft = bytesleft - 4                          ' Reduce bytes left
               SELECT CASE AS CONST$ ZPrf.RECFM                   ' See how to handle it
                  CASE "V"                                        ' Old style V
                     i = CVL(MID$(Chunk, 2, 1) + LEFT$(Chunk, 1) + CHR$(0,0)) - 4 ' Get actual data length
                  CASE "VBI"                                      ' Big Endian
                     i = CVL(Chunk)                               '
                  CASE "VLI"                                      ' Little EndiAN
                     i = CVL(RIGHT$(Chunk, 1) + MID$(Chunk, 3, 1) + MID$(Chunk, 2, 1) + LEFT$(Chunk, 1))
               END SELECT                                         '
               IF i <= bytesleft THEN                             ' Better be this much left
                  GET$ # cpIO.FNum, i, t                          ' Get the record data now
                  bytesleft -= i                                  ' reduce bytesleft
               ELSE                                               '
                  j = sDoMsgBox("File: |K" & cpIO.FilePath + $CRLF + "|B does not appear to be a valid RECFM=" + ZPrf.RECFM + " file." + $CRLF + _
                           "The data loaded should not be trusted", %MB_OK + %MB_USERICON, _
                           "SPFLite FileLoad")     '
                  GOTO VDeblockDone                               ' Bail out
               END IF                                             '

               '----- Do whatever translate is needed
               SELECT CASE AS CONST$ lclSource                    ' See what, if any, Unicode Xlate is needed
                  CASE "ANSI"                                     ' Top choice - do nothing
                     EXIT SELECT                                  '
                  CASE "UTF16"                                    ' Standard UTF16?
                     t = ACODE$(t)                                '
                  CASE "UTF8"                                     ' UTF8?
                     t = sUtf8ToAnsi(t, UTFError1)                '
                     UTFError2 = (UTFError1 OR UTFError2)         ' Merge error flags
                  CASE "UTF16BE"                                  ' UTF16 BigEndian?
                     FOR i = 0 TO LEN(t) - 1 STEP 2               ' We have to reverse Char. Pairs 1st
                        C1 = STRPTR(t) + i: C2 = C1 + 1           ' Point at 1st two characters
                        SWAP @C1, @C2                             ' to get BE into LE format
                     NEXT                                         '
                     t = ACODE$(t)                                ' Now we can Xlate it
                  CASE ELSE                                       ' Some other SOURCE value
                     me.Translate(t, ZPrf.GetSS2APtr)             ' Translate SOURCE to ANSI
               END SELECT                                         '

               GOSUB InsertT                                      ' Insert t as a line

            LOOP                                                  '
            GOTO VDeblockDone                                     '
         END IF                                                   '

         '----- Process the file a chunk (64K) at a time
         DO WHILE bytesleft > 0                                   ' Loop through file
            IF gfInterrupt THEN EXIT DO                           ' Break during FF search
            chunklen = MIN(65536, bytesleft)                      ' How big a chunk to read
            GET$ # cpIO.FNum, chunklen, Chunk                     ' Get a good size chunk
            bytesleft = bytesleft - LEN(Chunk)                    ' Reduce bytes left

            '----- If the 1st block, do a bunch of extra stuff
            IF OneTime = 0 THEN                                   ' 1st data block?
               OneTime = 1                                        ' Flip it
               GOSUB SetSource                                    ' Go set lclSource

               IF ISFALSE ZPrf.RECFM = "V" AND ISFALSE LEFT$(ZPrf.EOL, 4) = "AUTO" THEN ' Skip next if RECFM=V or EOL=AUTO
                  IF LEFT$(Chunk, 4) = CHR$(&h00, &h00, &hFE, &hFF) OR _ ' See if a Unicode file UTF32BE
                     LEFT$(Chunk, 4) = CHR$(&hFF, &hFE, &h00, &h00) THEN ' See if a Unicode file UTF32LE
                     scError(%eFail, "Unsupported Unicode file")  ' Tell user we don't do this one
                     Chunk = MID$(Chunk, 5)                       ' Strip off the leader
                     lclSource = "ANSI"                           ' Default to ANSI
                  ELSEIF LEFT$(Chunk, 3) = CHR$(&hEF, &hBB, &hBF) THEN  ' UTF8?
                     lclSource = "UTF8"                           ' Set it
                     Chunk = MID$(Chunk, 4)                       ' Strip off the leader
                  ELSEIF LEFT$(Chunk, 2) = CHR$(&hFF, &hFE) THEN  ' UTF16LE?
                     lclSource = "UTF16"                          ' Set it
                     Chunk = MID$(Chunk, 3)                       ' Strip off the leader
                  ELSEIF LEFT$(Chunk, 2) = CHR$(&hFE, &hFF) THEN  ' UTF16BE?
                     lclSource = "UTF16BE"                        ' Set it
                     Chunk = MID$(Chunk, 3)                       ' Strip off the leader
                  END IF                                          '
               END IF                                             '

               '----- Determine EOL type present for ANSI files
               IF ZPrf.RECFM = "V" THEN                           ' Skip if RECFM=V
                  tmpEOL = "NONE": dlm = ""                       '
               ELSE                                               '
                  ARRAY SCAN gEOLFlagList() FOR 7, = ZPrf.EOL, TO i ' Profile using one of the standard list items?
                  IF ISFALSE i THEN                               ' No, must be a Hex value
                     dlm = sHex2Str(ZPrf.EOL)                     ' Calc the delimiter
                     tmpEOL = UUCASE(ZPrf.EOL)                    ' Swap it
                  ELSEIF ZPrf.EOL = "NONE" THEN                   ' The NONE case?
                     tmpEOL = "NONE": dlm = ""                    '
                  ELSE                                            '
                     IF INSTR(Chunk, $CRLF) THEN                  ' Normal?
                        '                                         '
                     ELSEIF INSTR(Chunk, $CR) THEN                ' CR
                        tmpEOL = "CR": dlm = $CR                  '
                     ELSEIF INSTR(Chunk, $LF) THEN                ' LF
                        tmpEOL = "LF": dlm = $LF                  '
                     ELSEIF INSTR(Chunk, $NL) THEN                ' NL
                        tmpEOL = "NL": dlm = $NL                  '
                     END IF                                       '
                  END IF                                          '
               END IF                                             '

               IF lclSource = "UTF16"  OR lclSource = "UTF16LE" OR lclSource = "UTF16BE" THEN  ' For UTF 16 bit check other way
                  IF INSTR(Chunk, CHR$(0, 13, 0, 10)) THEN        ' Normal?
                     tmpEOL = "CRLF": dlm = $CRLF                 '
                  ELSEIF INSTR(Chunk, CHR$(13, 0, 10, 0)) THEN    ' Normal?
                     tmpEOL = "CRLF": dlm = $CRLF                 '
                  ELSEIF INSTR(Chunk, CHR$(0, 13)) THEN           ' CR
                     tmpEOL = "CR": dlm = $CR                     '
                  ELSEIF INSTR(Chunk, CHR$(13, 0)) THEN           ' CR
                     tmpEOL = "CR": dlm = $CR                     '
                  ELSEIF INSTR(Chunk, CHR$(10, 0)) THEN           ' LF
                     tmpEOL = "LF": dlm = $LF                     '
                  ELSEIF INSTR(Chunk, CHR$(0, 10)) THEN           ' LF
                     tmpEOL = "LF": dlm = $LF                     '
                  ELSEIF INSTR(Chunk, CHR$(0, 21)) THEN           ' NL
                     tmpEOL = "NL": dlm = $NL                     '
                  ELSEIF INSTR(Chunk, CHR$(21, 0)) THEN           ' NL
                     tmpEOL = "NL": dlm = $NL                     '
                  END IF                                          '
               END IF                                             '

               '----- Ensure Non-ANSI EOL is handled OK
               IF ZPrf.PSource <> "ANSI" AND _                    ' Treat all these SOURCE values as ANSI
                  ZPrf.PSource <> "UTF8" AND _                    '
                  ZPrf.PSource <> "UTF16" AND _                   '
                  ZPrf.PSource <> "UTF16LE" AND _                 '
                  ZPrf.PSource <> "UTF16BE" THEN                  '
                  IF ZPrf.EOL <> "NONE" THEN _                    ' If EOL NONE, skip setting the dlm
                     me.Translate(dlm, ZPrf.GetSS2APtr)           ' For other SOURCE values convert dlm to pre-translate version
               END IF                                             '
               IF ZPrf.EOL = "NONE" THEN                          ' If EOL NONE
                  DLM = "": tmpEOL = "NONE"                       ' Null the DLM so it doesn't mess up LRECL
               END IF                                             '

               '----- Check possible LRECL problem
               IF ZPrf.LRECL > 0 THEN                             ' If Fixed records, see if user should be warned
                  IF LOF(cpIO.FNum) MOD (ZPrf.LRECL + LEN(dlm)) <> 0 THEN  ' Evenly divisible?
                     j = sDoMsgBox("File: |K" & cpIO.FilePath + $CRLF + "|Bis not a multiple of LRECL " + FORMAT$(ZPrf.LRECL) + ", EOL=" + tmpEOL + $CRLF + _
                              "the last record will be blank padded", %MB_OK + %MB_USERICON, _
                              "SPFLite FileLoad")  '
                  END IF                                          '
               END IF                                             '
            END IF                                                '

            '----- Do whatever translate is needed
            SELECT CASE AS CONST$ lclSource                       ' See what, if any, Unicode Xlate is needed
               CASE "ANSI"                                        ' Top choice - do nothing
                  EXIT SELECT                                     '
               CASE "UTF16"                                       ' Standard UTF16?
                  Chunk = ACODE$(Chunk)                           '
               CASE "UTF8"                                        ' UTF8?
                  Chunk = sUtf8ToAnsi(Chunk, UTFError1)           '
                  UTFError2 = (UTFError1 OR UTFError2)            ' Merge error flags
               CASE "UTF16BE"                                     ' UTF16 BigEndian?
                  FOR i = 0 TO LEN(Chunk) - 1 STEP 2              ' We have to reverse Char. Pairs 1st
                     C1 = STRPTR(Chunk) + i: C2 = C1 + 1          ' Point at 1st two characters
                     SWAP @C1, @C2                                ' to get BE into LE format
                  NEXT                                            '
                  Chunk = ACODE$(Chunk)                           ' Now we can Xlate it
               CASE ELSE                                          ' Some other SOURCE value
                  me.Translate(Chunk, ZPrf.GetSS2APtr)            ' Translate SOURCE to ANSI
            END SELECT                                            '

            '----- Handle pass over from one chunk to the next
            Chunk = leftover + Chunk                              ' Pick up any leftover
            leftover = ""                                         ' Clear leftover

            '----- Finally, try deblocking a line out of the block
            DO WHILE LEN(Chunk) > 0                               ' Parse out the chunk
               IF gfInterrupt THEN EXIT DO                        ' Break during FF search
               RESET gLoopCtr                                     ' Reset LoopCtr to avoid treating as a loop

               '----- If EOL AUTO we do lots of fiddles to try and get things neat
               IF LEFT$(ZPrf.EOL, 4) = "AUTO" THEN                ' The ugly AUTO/NL type?
                  IF LEN(Chunk) < 250 AND bytesleft > 0 THEN      ' Nearing the end of the chunk?
                     leftover = Chunk                             ' Read another bit to avoid splitting delimiter sets up
                     EXIT DO                                      '
                  END IF                                          '
                  i = INSTR(Chunk, ANY $EOLDLM)                   ' See if any interesting delimiters
                  IF i > 0 OR bytesleft = 0 THEN                  ' If there's a line present, or we're out of data
                     IF MID$(Chunk, i, 3) = $CR + $FF + $LF THEN  ' The weird CR/FF/LF type?
                        t = IIF$(i = 1, "", LEFT$(Chunk, i - 1))  ' Set t = the line's data
                        Chunk = MID$(Chunk, i + 3)                ' Remove it from Chunk
                        IF OverPrint THEN GOSUB DoOverlay         ' Go combine t and tt into t
                        GOSUB InsertT                             ' Insert t as a line
                        PageFlag = %True                          ' Make next line a PAGE
                     ELSEIF MID$(Chunk, i, 3) = $CR + $CR + $LF THEN ' The extra CR type?
                        t = IIF$(i = 1, "", LEFT$(Chunk, i - 1))  ' Set t = the line's data
                        Chunk = MID$(Chunk, i + 3)                ' Remove it from Chunk
                        IF OverPrint THEN GOSUB DoOverlay         ' Go combine t and tt into t
                        GOSUB InsertT                             ' Insert t as a line
                     ELSEIF MID$(Chunk, i, 2) = $CRLF THEN        ' The normal line?
                        t = IIF$(i = 1, "", LEFT$(Chunk, i - 1))  ' Set t = the line's data
                        Chunk = MID$(Chunk, i + 2)                ' Remove it from Chunk
                        IF OverPrint THEN GOSUB DoOverlay         ' Go combine t and tt into t
                        GOSUB InsertT                             ' Insert t as a line
                     ELSEIF MID$(Chunk, i, 2) = $CR + $FF THEN    ' The normal line?
                        PageFlag = %True                          ' Mark PAGE
                        t = IIF$(i = 1, "", LEFT$(Chunk, i - 1))  ' Set t = the line's data
                        Chunk = MID$(Chunk, i + 2)                ' Remove it from Chunk
                        IF OverPrint THEN GOSUB DoOverlay         ' Go combine t and tt into t
                        GOSUB InsertT                             ' Insert t as a line
                     ELSEIF MID$(Chunk, i, 2) = $LF + $NUL THEN   ' Another weird LF line?
                        t = IIF$(i = 1, "", LEFT$(Chunk, i - 1))  ' Set t = the line's data
                        Chunk = MID$(Chunk, i + 2)                ' Remove it from Chunk
                        IF OverPrint THEN GOSUB DoOverlay         ' Go combine t and tt into t
                        GOSUB InsertT                             ' Insert t as a line
                     ELSEIF MID$(Chunk, i, 1) = $LF THEN          ' The LF line?
                        t = IIF$(i = 1, "", LEFT$(Chunk, i - 1))  ' Set t = the line's data
                        Chunk = MID$(Chunk, i + 1)                ' Remove it from Chunk
                        IF OverPrint THEN GOSUB DoOverlay         ' Go combine t and tt into t
                        GOSUB InsertT                             ' Insert t as a line
                     ELSEIF MID$(Chunk, i, 1) = $CR THEN          ' The CR line?
                        t = IIF$(i = 1, "", LEFT$(Chunk, i - 1))  ' Set t = the line's data
                        Chunk = MID$(Chunk, i + 1)                ' Remove it from Chunk
                        IF OverPrint THEN GOSUB DoOverlay         ' Go combine t and tt into t
                        GOSUB InsertT                             ' Insert t as a line
                     ELSEIF MID$(Chunk, i, 1) = $EOF THEN         ' The EOF?
                        Chunk = MID$(Chunk, i + 1)                ' Remove it from Chunk
                     ELSEIF bytescopied = 0 AND MID$(Chunk, i, 1) = $FF THEN ' If file starts with a FF then ignore it
                        Chunk = MID$(Chunk, 2)                    '
                     ELSEIF MID$(Chunk, i, 1) = $FF THEN          ' If an FF then set flag for next line
                        Chunk = MID$(Chunk, 2): PageFlag = %True  '
                     ELSEIF MID$(Chunk, i, 3) = $CR + $CR + $LF THEN ' Extra $CR, ignore it
                        Chunk = MID$(Chunk, 2)                    '
                     ELSEIF MID$(Chunk, i, 1) = $CR AND INSTR($CR + $LF + $FF, MID$(Chunk, i + 1, 1)) = 0 THEN ' Lonely $CR (overlay)
                        IF ZPrf.EOL = "AUTO" THEN                 ' Normal AUTO type?
                           OverPrint = %True                      ' Switch to OverPrint mode
                           tt = IIF$(i = 1, "", LEFT$(Chunk, i - 1)) ' Set tt = the line's data
                           Chunk = MID$(Chunk, i + 1)             ' Remove it from Chunk
                        ELSE                                      ' Must be AUTONL
                           t = IIF$(i = 1, "", LEFT$(Chunk, i - 1)) ' Set t = the line's data
                           Chunk = MID$(Chunk, i + 1)             ' Remove it from Chunk
                           IF OverPrint THEN GOSUB DoOverlay      ' Go combine t and tt into t
                           GOSUB InsertT                          ' Insert t as a line
                        END IF                                    '
                     ELSEIF bytesleft = 0 THEN                    ' Last piece with no delimiter
                        t = Chunk                                 ' Set t = the remaining data
                        Chunk = ""                                ' Remove it from Chunk
                        IF OverPrint THEN GOSUB DoOverlay         ' Go combine t and tt into t
                        GOSUB InsertT                             ' Insert t as a line
                     END IF                                       '
                  ELSE                                            ' We need continued data
                     leftover = Chunk                             ' Save whats left
                     EXIT DO                                      ' Exit to read more data
                  END IF                                          '

               '----- A simple delimited line
               ELSEIF ZPrf.RECFM = "U" THEN                       ' If a delimited type file
                  IF INSTR(Chunk, dlm) OR _                       ' If there's a line present
                     bytesleft = 0 THEN                           ' Or we're simply out of data
                     t = EXTRACT$(Chunk, dlm)                     ' Read the data
                     Chunk = REMAIN$(Chunk, dlm)                  '
                     GOSUB InsertT                                ' Go insert the line
                  ELSE                                            ' We need continued data
                     leftover = Chunk                             ' Save whats left
                     EXIT DO                                      ' Exit to read more data
                  END IF                                          '

               '----- Fixed LRECL type deblocking
               ELSEIF ZPrf.RECFM = "F" THEN                       ' Is this the fixed LRECL?
                  i = LEN(Chunk): j = ZPrf.LRECL + LEN(dlm): k = bytesleft
                  IF LEN(Chunk) >= ZPrf.LRECL + LEN(dlm) THEN     ' If there's enough data for a record
                     t = LEFT$(Chunk, ZPrf.LRECL + LEN(dlm))      ' Extract a line
                     Chunk = MID$(Chunk, ZPrf.LRECL + LEN(dlm) + 1)'
                     GOSUB InsertT                                ' Go insert the line
                  ELSE                                            ' We need continued data
                     leftover = Chunk                             ' Save whats left
                     EXIT DO                                      ' Exit to read more data
                  END IF                                          '

               '----- Variable RDW type deblocking
               ELSEIF ZPrf.RECFM = "V" THEN                       ' Is this the variable?
                  i = MAK(LONG, PEEK(STRPTR(Chunk)+1), PEEK(STRPTR(Chunk)))  ' Get length from RDW
                  IF i < LEN(Chunk) OR bytesleft = 0 THEN         ' Whole record within this chunk, or EOF?
                     IF i > LEN(Chunk) THEN                       ' We have a leftover partial record
                        i = LEN(Chunk)                            ' just do the last bit
                        j = sDoMsgBox("File: |K" & TRIM$(cpIO.FilePath) + $CRLF + "|Bis not a valid RECFM=V file," + $CRLF + _
                                 "the last record is probably invalid", %MB_OK + %MB_USERICON, _
                                 "SPFLite FileLoad") '
                     END IF                                       '
                     t = MID$(Chunk, 5, i - 4)                    ' Extract a line
                     GOSUB InsertT                                ' Go insert the line
                     Chunk = MID$(Chunk, i + 1)                   '
                  ELSE                                            ' We need continued data
                     leftover = Chunk                             ' Save whats left
                     EXIT DO                                      ' Exit to read more data
                  END IF                                          '
               END IF                                             '
            LOOP                                                  '
         LOOP                                                     '

         '----- End of data, handle the remainder
         IF LEN(Chunk) AND ZPrf.LRECL > 0 THEN                    ' We have a partial RECFM LRECL leftover
            t = Chunk                                             ' Pick up the remaining text
            me.LInsertEmpty(ln, 1, %Data)                         ' Request insert 1, %Data line
            GOSUB InsertT                                         ' Go insert the line
         END IF                                                   '

         VDeblockDone:
         '----- Close the file
         cpIO.Close                                               ' Close it
         IF lin = 0 THEN MExitMeth                                ' If a simple read, just exit right away

         IF ISFALSE quick AND ISFALSE IsMedit THEN sRecentAdd(cpIO.FilePath) ' Go add to RECENT list

         '----- Figure out what STATE processing to do
         IF CurrPcmd = "COPY" AND _                               ' If the COPY command
            ZPrf.PState > %StateOff THEN                          ' State somehow active?
            IsCopy = %True                                        ' Tell STATE it's COPY calling
            GOSUB DoState                                         ' Then go add any NOTEs
         END IF                                                   '

         IF CurrPcmd <> "COPY" AND _                              ' If not the COPY command
            ISFALSE quick THEN                                    ' Not quick mode
            IF ZPrf.PState > %StateOff THEN                       ' State somehow active?
               IsCopy = %False                                    ' Tell STATE it's NOT COPY calling
               GOSUB DoState                                      ' Then go add the STATE data
            END IF                                                '
         END IF                                                   '

         '----- Check Overlay Conflict
         IF OverConf THEN _                                       ' Did a conflict occur?
            scError(%eNone, "EOL AUTO overprinting resulted in data loss")

         '----- Check for UTF8 conversion error
         IF UTFError2 <> 0 THEN                                   ' Was there an error?
            IF (UTFError2 AND 1) = 1 THEN                         ' Value outside ANSI range?
               scError(%eNone, "File contains UTF8 data outside ANSI range, UTF8 conversion terminated")
            ELSEIF (UTFError2 AND 2) = 2 THEN                     ' Malformed UTF8 value, X'A4' substituted
               scError(%eNone, "File contains malformed UTF8 data, X'A4' chars substituted")
            END IF
         END IF                                                   '

         '----- Do the CAPS AUTO setting
         IF ZPrf.CapsDesired = 2 AND CurrPcmd <> "COPY" THEN      ' If not COPY and CAPS=AUTO
            IF lclUpper AND ISFALSE lclLower THEN                 ' If only Upper, then set temp CAPS ON
               ZPrf.CapsActual = 1                                '
               scError(%eNone, "CAPS set to AUTO:on")             '
            ELSE                                                  '
               ZPrf.CapsActual = 0                                '
               scError(%eNone, "CAPS set to AUTO:off")            '
            END IF                                                '
         END IF                                                   '

      '----- Check EOL conflict possibility
         IF CurrPcmd <> "COPY" THEN                               ' If not the COPY command
            ZPrf.PSource = lclSource                              ' Save the Encoding
            IF LEFT$(ZPrf.EOL, 4) <> "AUTO" AND _                 '
               LEFT$(ZPrf.EOL, 4) <> "NONE" AND _                 '
               tmpEOL <> ZPrf.EOL THEN                            ' Tell user of any EOL difference
               t = TRIM$(tmpEOL)                                  ' Format tmpEOL
               IF VERIFY(t, $Hex) = 0 THEN t = "X'" + t + "'"     '
               tt = TRIM$(ZPrf.EOL)                               ' Format ZPrf.EOL
               IF VERIFY(tt, $Hex) = 0 THEN tt = "X'" + tt + "'"  ' If hex operand, frame it for the message
               scError(0, "Warning, Profile EOL is set to " & tt & ", file loaded was " & t)
            END IF                                                '
         ELSE                                                     '
            IF ZPrf.PSource <> lclSource THEN                     ' Warn if encoding is different
               scError(0, "Warning, current SOURCE is " & ZPrf.PSource & ", file loaded was " & TRIM$(lclSource))
            END IF                                                '
         END IF                                                   '
         IF ISFALSE quick THEN OffUndoFlag                        ' Make sure we get an Undo recorded
         gfXRebuild  = %True                                      '

         MExitMeth                                                '

      '----- Insert a single line from the deblocking routines
      InsertT:
         IF LEN(t) = 1 AND t = CHR$(26) AND ISNULL(Chunk) AND bytesleft = 0 THEN RETURN ' Ignore old EOF markers.
         IF ZPrf.LRECL > 0 AND LEN(t) = ZPrf.LRECL + LEN(dlm) THEN' LONG because of useless dlms?
            t = LEFT$(t, LEN(t) - LEN(dlm))                       ' Drop the dlm
         END IF                                                   '

         '----- If CopyaFile called as a simple line reader, stuff lines in a clipboard
         IF lin = 0 THEN                                          ' Called to insert after line zero?
            CBLines += t + $CRLF                                  ' Add to CBLines
            RETURN                                                ' and return
         END IF                                                   '

         '----- Not line reader mode, insert the line
         INCR lctr                                                ' Count the line
         IF lctr = 1 THEN LastPage = %False                       '
         IF frmlin <> 0 OR tolin <> 0 THEN                        ' Doing a line range
            IF lctr < frmlin OR lctr > tolin THEN RETURN          ' Only do lines in range
         END IF                                                   '
         me.LInsertEmpty(ln, 1, %Data)                            ' Request insert 1, %Data line
         INCR ln                                                  ' Bump insertion point
         me.LTxtSet(ln, t)                                        ' Set it in
         IF PageFlag THEN                                         ' Set as a PAGE line?
            IF ISFALSE LastPage THEN                              ' Try to prevent two in a row
               me.LFlagBitOn(ln, %Page)                           ' Yes, set the flag
               INCR PCtr                                          ' Count pages
               LWrk2S(ln) = PCtr                                  ' save it
               LastPage = %True                                   ' Remember for next
            END IF                                                '
            PageFlag = %False                                     ' Flip flag
         ELSE                                                     '
            LastPage = %False                                     ' No longer LastPage
         END IF                                                   '

         IF ZPrf.CapsDesired = 2 AND CurrPcmd <> "COPY" THEN      ' If not COPY and CAPS=AUTO
            IF INSTR(t, ANY gLower) THEN lclLower = %True         ' Do tests for CAPS AUTO
            IF INSTR(t, ANY gUpper) THEN lclUpper = %True         '
         END IF                                                   '

         IF ZPrf.LRECL > 0 AND LEN(t) > Prf.LRECL THEN            ' A LONG LRECL record?
            scError(%eNone, "Records exceeding LRECL " + FORMAT$(ZPrf.LRECL) + " were detected")
         END IF                                                   '

         me.UpdLControl(ln)                                       ' Setup LLCtl
         GOSUB DoPM                                               ' Go do PM processing

         IF ZPrf.ImportTabs > 0 THEN                              ' Are we expanding tabs?
            lTxtP = LTxtGP(ln)                                    ' Point at the Txt
            me.LTxtSet(ln, TAB$(@lTxtP, ZPrf.ImportTabs))         ' Save it and do tab conversion
         END IF                                                   '
         bytescopied += LEN(t)                                    ' Accum output
         me.AttrScan(ln)                                          ' Setup attributes
         StateTo = ln                                             ' Last line for STATE processing
         RETURN                                                   '

      '----- Do COPY Overlay
      DoOverlay:
         i = MAX(LEN(t), LEN(tt))                                 ' Loop for longest line
         t = LSET$(t, i): tt = LSET$(tt, i)                       ' Make both lines same length
         FOR k = 1 TO i                                           '
            IF MID$(tt, k, 1) = " " THEN                          ' If a blank, overlay with impunity
               MID$(tt, k, 1) = MID$(t, k, 1)                     ' Just copy backward
            ELSEIF MID$(tt, k, 1) = "_" AND MID$(t, k, 1) <> " " THEN   ' If first line is _ and 2nd is not blank
               MID$(tt, k, 1) = MID$(t, k, 1)                           ' Replace the _ with the character
            ELSEIF MID$(tt, k, 1) <> " " AND MID$(t, k, 1) <> " " THEN  ' If still a conflict, remember that
               OverConf = %True                                   ' Set the flag
            END IF                                                '
         NEXT k                                                   '
         t = tt: OverPrint = %False                               ' Swap the merged line.  Turn flag off.
         RETURN                                                   '

      '----- Handle the STATE loading
      DoState:
         me.StateLoad(cpIO.FilePath, StateFrom, StateTo, IsCopy)  ' Go load them if valid (%True=COPY)
         RETURN

      '----- Do the Post exclude processing
      DoPM:
         IF BIT(lclPM, %lCmdX) THEN                               ' Do the +/- processing for the Source
            me.LFlagBitOn(ln, %Invisible)                         ' If - make Invisible
            gfXRebuild = %True                                    ' Ask for Exclude processing
         ELSEIF BIT(lclPM, %lCmdNX) THEN                          '
            me.LFlagBitOff(ln, %Invisible)                        ' If + make visible
            gfXRebuild = %True                                    ' Ask for Exclude processing
         END IF                                                   '
         RETURN                                                   '

      SetSource:
         '----- Determine translation type
         IF ZPrf.PSource = "ANSI" OR _                            ' Call all these SOURCE values ANSI
            ZPrf.PSource = "UTF8" OR _                            '
            ZPrf.PSource = "UTF16" OR _                           '
            ZPrf.PSource = "UTF16LE" OR _                         '
            ZPrf.PSource = "UTF16BE" THEN                         '
            lclSource = "ANSI"                                    '
         ELSE                                                     '
            lclSource = ZPrf.PSource                              ' Else use what's coded
         END IF                                                   '
         RETURN                                                   '

      END METHOD

      METHOD CRPBack(MoveMode AS INTEGER, FromCRP AS LONG, NumRows AS LONG) AS LONG
      '---------- Return an altered Row pointer
      REGISTER i AS LONG
      REGISTER j AS LONG
      LOCAL k, l AS LONG
         MEntry
         '----- Take a line pointer and move it backward accounting for X'd and special lines
         i = NumRows: j = FromCRP                                 ' Copy params

         DO WHILE i                                               ' While more lines to go back
            '----- Visible lines
            IF MoveMode = %mVisible THEN                          ' Doing Visible lines?
               DO                                                 ' Yes, loop backward till a visible line is seen
                  j = MAX(j - 1, 1)                               '
               LOOP UNTIL (ISFALSE IsLInvisible(j) AND ISFALSE IsLXclude(j)) OR _
                          (IsLXclude(j) AND ISFALSE IsHideFlag) OR _
                          J = 1                                   '

            '----- Data lines
            ELSEIF MoveMode = %mData THEN                         ' Doing Data Lines?
               DO                                                 ' Yes, loop backward till a Data line is seen
                  j = MAX(j - 1, 1)                               '
               LOOP UNTIL IsLData(j) OR IsLNote(j) OR j = 1       '

            '----- Only Data lines
            ELSEIF MoveMode = %mOData THEN                        ' Doing Data Lines?
               DO                                                 ' Yes, loop backward till a Data line is seen
                  j = MAX(j - 1, 1)                               '
               LOOP UNTIL IsLData(j) OR j = 1                     '

            '----- Visible Data lines
            ELSEIF MoveMode = %mVisData THEN                      ' Doing visible Data Lines?
               DO                                                 ' Yes, loop backward till a Data line is seen
                  j = MAX(j - 1, 1)                               '
               LOOP UNTIL ((IsLData(j) OR IsLNote(j)) AND ISFALSE IsLInvisible(j)) OR j = 1

            '----- Page mode backward
            ELSEIF MoveMode = %mPage THEN                         ' Doing PAGE mode Lines?
               DO                                                 ' Yes, loop backward till a Page line is seen
                  j = MAX(j - 1, 1)                               '
               LOOP UNTIL (IsLData(j) AND ISFALSE IsLInvisible(j) AND IsLPage(j)) OR j = 1

            END IF                                                '
            IF j = 1 THEN EXIT DO                                 ' OK, if Row is zero, we'd better stop
            DECR i                                                ' Else DECR the number we're supposed to do
         LOOP                                                     ' and loop till done
         METHOD = j                                               ' Pass back the Row we ended up with
         MExit
      END METHOD

      METHOD CRPBack2Page() AS LONG
      '---------- return # of lines to next page
      REGISTER i AS LONG
      REGISTER j AS LONG
         j = 0                                                    ' Count lines to next page
         i = TopScrn                                              ' i = current top of screen
         DO                                                       ' Find next page line or lastline
            INCR j: i = MAX(i - 1, 1)                             ' Count
         LOOP UNTIL IsLPage(i) OR i = 1                           '
         METHOD = j                                               ' Return count
      END METHOD

      METHOD CRPFwd(MoveMode AS INTEGER, FromCRP AS LONG, NumRows AS LONG) AS LONG
      '---------- Return an altered Row pointer
      REGISTER i AS LONG
      REGISTER j AS LONG
      LOCAL k, l, indent AS LONG, dlm AS INTEGER
         MEntry
         '----- Take a line pointer and move it foreward accounting for X'd and special lines
         i = NumRows: j = FromCRP                                 ' Copy params

         '----- Fudge text values
         IF MoveMode = %MText THEN                                '
            i = 99999999                                          ' If Text mode set a LARGE limit
            k = j                                                 ' Look for Text line 2
            DO                                                    '
               INCR k                                             '
            LOOP WHILE ISFALSE IsLData(k) AND k < LastLine        ' k ends up at next data line
            indent = LEN(L(k).@LTxt) - LEN(LTRIM$(LTxtG(k))) + 1  ' Set indent of 2nd line
         END IF                                                   '

         DO WHILE i                                               ' While more lines to go forward

            '----- Visible lines
            IF MoveMode = %mVisible THEN                          ' Doing Visible lines?
               DO                                                 ' Yes, loop foreward till a Visible line is seen
                  j = MIN(j + 1, LastLine)                        '
               LOOP UNTIL (ISFALSE IsLInvisible(j) AND ISFALSE IsLXclude(j)) OR _
                          (IsLXclude(j) AND ISFALSE IsHideFlag) OR _'
                          J = LastLine                            '

            '----- Data lines
            ELSEIF MoveMode = %mData THEN                         ' Doing Data lines?
               DO                                                 ' Yes, loop foreward till a Data line is seen
                  j = MIN(j + 1, LastLine)                        '
               LOOP UNTIL IsLData(j) OR IsLNote(j) OR j = LastLine'

            '----- Only Data lines
            ELSEIF MoveMode = %mOData THEN                        ' Doing Data lines?
               DO                                                 ' Yes, loop foreward till a Data line is seen
                  j = MIN(j + 1, LastLine)                        '
               LOOP UNTIL IsLData(j) OR j = LastLine              '

            '----- Visible data lines
            ELSEIF MoveMode = %mVisData THEN                      ' Doing Visible Data lines?
               DO                                                 ' Yes, loop foreward till a Data line is seen
                  j = MIN(j + 1, LastLine)                        '
               LOOP UNTIL ((IsLData(j) OR IsLNote(j)) AND ISFALSE IsLInvisible(j)) OR j = LastLine

            '----- Page mode
            ELSEIF MoveMode = %mPage THEN                         ' Doing PAGE mode Lines?
               IF IsLTop(j) THEN INCR j                           ' Fudge top of file PAGE
               DO                                                 ' Yes, loop forward till a Page line is seen
                  j = MIN(j + 1, LastLine)                        '
               LOOP UNTIL (IsLData(j) AND ISFALSE IsLInvisible(j) AND IsLPage(j)) OR j = LastLine

            '----- Text mode
            ELSEIF MoveMode = %mText THEN                         ' Doing Text lines?
               dlm = 0                                            ' Reset DLM flag
               DO                                                 ' Yes, loop foreward till end of paragraph
                  j = MIN(j + 1, LastLine)                        '
                  IF IsLData(j) THEN                              ' Look only at Data lines
                     IF LEFT$(LTxtG(j), 1) = "." OR _             ' Look for leading delimiters
                        LEFT$(LTxtG(j), 1) = ":" OR _             '
                        LEFT$(LTxtG(j), 1) = "&" OR _             '
                        LEFT$(LTxtG(j), 1) = "<" THEN             '
                        dlm = 1                                   ' Break out of loop
                        EXIT DO                                   '
                     END IF                                       '
                     IF ISNULL(TRIM$(LTxtG(j))) THEN              ' Look for blank line
                        dlm = 1                                   ' Break out of loop
                        EXIT DO                                   '
                     END IF                                       '
                     IF indent <> LEN(L(j).@LTxt) - LEN(LTRIM$(LTxtG(j))) + 1 THEN ' Different indent than 2nd line
                        dlm = 1                                   ' Break out of loop
                        EXIT DO                                   '
                     END IF                                       '
                  ELSEIF j = LastLine THEN                        '
                     EXIT DO                                      '
                  END IF                                          '
               LOOP UNTIL dlm = 1                                 '
               DECR j                                             ' Point at the last one
               EXIT DO                                            '
            END IF                                                '

            IF j = LastLine THEN EXIT DO                          ' OK if Row is at the end, stop
            DECR i                                                ' Else DECR the number we're supposed to do
         LOOP                                                     ' and loop till done

         METHOD = j                                               ' Pass back the resulting Row
         MExit
      END METHOD

      METHOD CRPFwd2Page() AS LONG
      '---------- return # of lines to next page
      REGISTER i AS LONG
      REGISTER j AS LONG
         j = 0                                                    ' Count lines to next page
         i = TopScrn                                              ' i = current top of screen
         DO                                                       ' Find next page line or lastline
            INCR j: INCR i                                        ' Count
         LOOP UNTIL IsLPage(i) OR i = LastLine                    '
         METHOD = j                                               ' Return count
      END METHOD

      PROPERTY GET CsrCol() AS LONG: PROPERTY = CCol: END PROPERTY
      PROPERTY SET CsrCol(v AS LONG): CCol = v: me.CsrMode: END PROPERTY

      METHOD CsrColAdd (amt AS LONG): CCol += amt: me.CsrMode: END METHOD
      METHOD CsrColSub (amt AS LONG): CCol -= amt: me.CsrMode: END METHOD

      PROPERTY GET CsrRow() AS LONG: PROPERTY = CRow: END PROPERTY
      PROPERTY SET CsrRow(v AS LONG): CRow = v: me.CsrMode: END PROPERTY

      METHOD CsrRowAdd (amt AS LONG): CRow += amt: me.CsrMode: END METHOD
      METHOD CsrRowSub (amt AS LONG): CRow -= amt: me.CsrMode: END METHOD

      METHOD CsrMode()
      LOCAL DataIX, i AS LONG
      '---------- Set Csr Zone values
      IF ISFALSE IsFMTab THEN                                     ' Edit screen?
         SELECT CASE AS LONG CRow                                 ' Select first by row
            CASE 1                                                ' Command line
               SELECT CASE AS LONG CCol                           ' Select by column
                  CASE < 11: Zone = %CursBad                      ' We're in Bad area
                  CASE <= 10 + pCmdLen                            ' In Command area?
                     Zone = %CursCmnd                             '
                     ZLCol = CCol - 10 + COffset                  ' Calc logical position in field
                     ZCol = CCol - 10                             ' Calc position in screen field
                  CASE > 20 + pCmdLen                             ' In Scroll Amt area?
                     Zone = %CursScrl                             '
                     ZLCol = CCol - 20 - pCmdLen                  ' Calc logical offset in field
                     ZCol = CCol - 20 - pCmdLen                   ' Calc offset in screen field
               END SELECT                                         ' End of row 1

            CASE 2                                                ' Line 2 (dashed)
               Zone = %CursBad                                    ' Bad

            CASE > 2 + Prf.Cols                                   ' In text area
               CursFindInv.nTop = 0                               ' Null Find Hilite pointer
               SELECT CASE AS LONG CCol                           ' Select by column
                  CASE < gLNPadCol                                ' In line command area
                     IF S(CRow) = 0 THEN _                        ' If not a valid line
                        Zone = %CursBad: EXIT SELECT              ' flag it Bad instead
                     Zone = %CursLinN                             ' Yep, remember that
                     ZLCol = CCol                                 ' Calc logical offset in field
                     ZCol = CCol                                  ' Calc offset in screen field
                     ZLNum = S(CRow)                              ' Pass back the line index
                  CASE > gLNPadCol                                ' In text data area?
                     Zone = %CursData                             ' Start as good Data
                     DataIX = sGetIX(CRow)                        ' Get real data IX pointer
                     IF DataIX = 0 OR DataIX = -3 THEN            ' Eliminate bad lines
                        Zone = %CursBad: EXIT SELECT              ' If it is, flag it Bad instead
                     ELSEIF DataIX > 0 AND _                      '
                        (ISTRUE (L(DataIX).LFlag AND (%Top OR %Bottom))) THEN
                        Zone = %CursBad: EXIT SELECT              ' If it is, flag it Bad instead
                     ELSEIF DataIx < 0 THEN                       ' A Hex Line?
                        DataIX = sGetIX(CRow - ABS(DataIX))       ' Convert to real line index
                     END IF                                       '
                     ZLCol = CCol - gLNPadCol + Offset            ' Calc logical offset in field
                     ZCol  = CCol - gLNPadCol                     ' Calc offset in screen field
                     ZLNum = DataIX                               ' Pass back the line index
                     me.AttrInvGrab(ZLNum, ZLCol)                 ' See if a find hilite to grab
               END SELECT                                         ' End of row 1

            CASE ELSE
               Zone = %CursBad                                    ' Make it Bad
         END SELECT                                               '

      ELSE                                                        ' FM Mode

         SELECT CASE AS LONG CRow                                 ' OK, figure it out

            CASE 1                                                ' Possible pCommand line?
               SELECT CASE AS LONG CCol                           ' Now look at column
                  CASE < 11: Zone = %CursBad                      ' We're in Bad area
                  CASE <= 10 + pCmdLen                            ' In Command area?
                     Zone = %CursCmnd                             '
                     ZLCol = CCol - 10 + COffset                  ' Calc logical offset in field
                     ZCol = CCol - 10                             ' Calc offset in screen field
                  CASE > 20 + pCmdLen                             ' In Scroll Amt area?
                     Zone = %CursScrl                             '
                     ZLCol = CCol - 20 - pCmdLen                  ' Calc logical offset in field
                     ZCol = CCol - 20 - pCmdLen                   ' Calc offset in screen field
               END SELECT

            CASE 2, FM_Head_Line, > ENV.ScrHeight + 1 - (2 * ENV.FMHelpFlag) ' Static rows
               Zone = %CursBad                                    ' So remember we're in Bad area

            CASE FM_Quick_Line_1                                  ' First quick line?
               i = CCol
               SELECT CASE AS LONG CCol                           ' Now look at column
                  CASE FM_Quick_Pos_1                             ' We're in New Empty Cmd area
                     Zone = %CursEmpty                            '
                  CASE FM_Quick_Pos_2                             ' We're in FilePath
                     Zone = %CursFilePath                         '
                  CASE FM_Quick_Pos_3                             ' We're in Recent Cmd area
                     Zone = %CursRecent                           '
                  CASE FM_Quick_Pos_4                             ' We're in Found Cmd area
                     Zone = %CursFound                            '
                  CASE FM_Quick_Pos_5                             ' We're in Open Files Cmd area
                     Zone = %CursOpen                             '
                  CASE FM_Quick_Pos_6                             ' We're in Favorite Cmd area
                     Zone = %CursFavorite                         '
                  CASE FM_Quick_Pos_7                             ' We're in FileList Cmd area
                     Zone = %CursFileList                         '
                  CASE FM_Quick_Pos_8                             ' We're in RPaths Cmd area
                     Zone = %CursRPaths                           '
                  CASE FM_Quick_Pos_9                             ' We're in Profiles Cmd area
                     Zone = %CursProfiles                         '
                  CASE ELSE: Zone = %CursBad                      ' We're in Bad area
               END SELECT                                         '

            CASE FM_Path_Line                                     ' Possible Default Path?
               IF ISNOTNULL(FileListNm) THEN Zone = %CursBad: EXIT SELECT  ' Bad in RECALL mode
               IF CCol < FM_Path_Left THEN Zone = %CursBad: EXIT SELECT  ' Left end?  Bad
               Zone = %CursPath                                   ' We're in Path field
               ZLCol = CCol - FM_Path_Left + 1                    ' Calc logical offset in field
               ZCol = CCol - FM_Path_Left + 1                     ' Calc offset in screen field

            CASE FM_Mask_Line                                     ' Not legal for FILELIST mode
               IF ISNOTNULL(FileListNm) THEN Zone = %CursBad: EXIT SELECT  ' Bad in RECALL mode
               IF CCol < FM_Mask_Left THEN Zone = %CursBad: EXIT SELECT  ' Done
               Zone = %CursMask                                   ' We're in Types
               ZLCol = CCol - FM_Mask_Left + 1                    ' Calc logical offset in field
               ZCol = CCol - FM_Mask_Left + 1                     ' Calc offset in screen field

            CASE >= FM_Top_File_Line                              ' File List area
               IF CRow >= FM_Top_File_Line + AFCount - TopScrn THEN Zone = %CursBad: EXIT SELECT
               SELECT CASE AS LONG CCol                           ' Now look at column
                  CASE 1 TO ENV.FMLCmdWidth                       ' We're in the line command area
                     ZLNum = CRow - FM_Top_File_Line + TopScrn    ' Calc  AFList index
                     Zone = %CursLCmd                             '
                     ZLCol = CCol + AFList(ZLNum).CmdOff          ' Calc logical offset in field
                     ZCol = CCol                                  ' Calc offset in screen field

                  CASE FM_Head_Note_Left TO FM_Head_Note_Left + FM_Note_Size - 1
                     IF FM_Head_Note_Left > 0 THEN                ' We're in Note (if Note active)
                        IF FMode = %FMFileList THEN               ' Have an associated FILELIST?
                           ZLNum = CRow - FM_Top_File_Line + TopScrn     ' Calc  AFList index
                           Zone = %CursLNote                      '
                           ZLCol = CCol - FM_Head_Note_Left + 1 + AFList(ZLNum).NoteOff ' Calc logical offset in field
                           ZCol = CCol - FM_Head_Note_Left + 1    ' Calc offset in screen field
                        ELSE                                      '
                           Zone = %CursLLin                       '
                        END IF                                    '
                     ELSE                                         '
                        Zone = %CursLLin                          '
                     END IF                                       '

                  CASE ELSE: Zone = %CursLLin                     ' We're elsewhere in the line
               END SELECT                                         '
         END SELECT                                               '
      END IF                                                      '
      END METHOD                                                  '

      METHOD CursCmnd() AS INTEGER: METHOD = ISTRUE Zone = %CursCmnd: END METHOD
      METHOD CursData() AS INTEGER: METHOD = ISTRUE Zone = %CursData: END METHOD
      METHOD CursLinN() AS INTEGER: METHOD = ISTRUE Zone = %CursLinN: END METHOD

      METHOD CurSetReq(prio AS INTEGER, Lin AS LONG, pCol AS LONG, fscrl AS INTEGER, OPT Back AS LONG, OPT HexRow AS LONG)
      '---------- Add an entry to the set cursor request queue
      LOCAL i, j AS LONG
         i = sCurLin: j = sCurPrio
         IF prio > sCurPrio THEN                                  ' A higher priority request?
            IF sCurPrio > 0 THEN                                  ' A previous item set?
               L(sCurLin).LCol =  0                               ' Clear it then
               me.LFlagBitOff(sCurLin, %Cursor)                   '
               me.LFlagBitOff(sCurLin, %Scroll)                   '
               me.LFlagBitOff(sCurLin, %XPtr)                     '
            END IF                                                '
            sCurPrio = prio                                       ' Save request reason (caller)
            sCurLin = lin                                         ' Save line
            sCurScrl = fscrl                                      ' Save SCRL flag
            sCurHexl = IIF(ISMISSING(HexRow), 0, HexRow)          ' Setup optional Hex Row
            L(lin).LCol =  pCol                                   ' Store in the actual text line area
            me.LFlagBitOn(lin, %Cursor)                           '
            IF fscrl THEN me.LFlagBitOn(lin, %Scroll)             '
            IF ISFALSE ISMISSING(Back) THEN                       '
               IF Back THEN me.LFlagBitOn(lin, %XPtr)             '
            END IF                                                '
         END IF                                                   '
      END METHOD

      METHOD CursLCmd()  AS INTEGER: METHOD = ISTRUE Zone = %CursLCmd: END METHOD
      METHOD CursLLin()  AS INTEGER: METHOD = ISTRUE Zone = %CursLLin: END METHOD

      METHOD DoMarkLines()
      '---------- Draw mark lines on the screen
      REGISTER i AS LONG
      LOCAL u AS STRING
         MEntry
         IF Prf.PMark AND ISFALSE IsPTypeMode AND ISNOTNULL(TRIM$(Prf.MarkLine)) THEN ' Column marking
            TopDrawn = 3 + Prf.Cols                               '
            i = sGetIX(TopDrawn)                                  '
            IF IsLTop(sGetIX(TopDrawn)) THEN TopDrawn = TopDrawn + 1 ' Adjust top line
            IF IsLBottom(sGetIX(LastDrawn)) THEN LastDrawn = LastDrawn - 1 ' Adjust bottom line
            i = 1: u = Prf.MarkWorking                            ' Scan for Mark columns
            DO                                                    '
               i = INSTR(i, u, "*")                               ' Look for next marker
               IF i = 0 THEN EXIT DO                              '
               IF i >= Offset + 1 AND i <= Offset + gDataLen THEN ' Within screen boundary
                  GRAPHIC SET MIX %MIX_COPYSRC                    '
                  GRAPHIC LINE ((i - 1 + gLNPadCol - Offset) * gFontWidth + %GLM - 1, (TopDrawn - 1) * gFontHeight) - ((i - 1 + gLNPadCol - Offset) * gFontWidth + %GLM - 1, LastDrawn * gFontHeight), ENV.cMarkLine
               END IF                                             '
               INCR i                                             '
            LOOP                                                  '
         END IF                                                   '
         MExit                                                    '
      END METHOD

      METHOD DoPFKShow
      LOCAL c, i, j, k, pline, lclHiLite AS LONG
      LOCAL Txt1, Txt2, Txt3, Txt4, Txt5, TLine AS STRING
         MEntry
         IF ENV.PFKShow = 0 OR IsFMTab THEN MExitMeth             ' Skip if not needed
         c = INT(ENV.ScrWidth / (KbdT.KbdPFShowMax + 2))          '
         j = 1: txt1 = "": txt2 = "": txt3 = "": txt4 = "": txt5 = ""   ' Starting variables
         FOR i = 1 TO c                                           ' Build a line
            IF j <= KbdT.KbdPFShowNum THEN                        ' Still PFSHOW entries left?
               txt1 += LSET$(MID$(KbdT.KbdPFShow(j), 2), KbdT.KbdPFShowMax + 2) '
               INCR j                                             '
            END IF                                                '
         NEXT i                                                   '
         FOR i = 1 TO c                                           ' Build a line
            IF j <= KbdT.KbdPFShowNum THEN                        ' Still PFSHOW entries left?
               txt2 += LSET$(MID$(KbdT.KbdPFShow(j), 2), KbdT.KbdPFShowMax + 2) '
               INCR j                                             '
            END IF                                                '
         NEXT i                                                   '
         FOR i = 1 TO c                                           ' Build a line
            IF j <= KbdT.KbdPFShowNum THEN                        ' Still PFSHOW entries left?
               txt3 += LSET$(MID$(KbdT.KbdPFShow(j), 2), KbdT.KbdPFShowMax + 2) '
               INCR j                                             '
            END IF                                                '
         NEXT i                                                   '
         FOR i = 1 TO c                                           ' Build a line
            IF j <= KbdT.KbdPFShowNum THEN                        ' Still PFSHOW entries left?
               txt4 += LSET$(MID$(KbdT.KbdPFShow(j), 2), KbdT.KbdPFShowMax + 2) '
               INCR j                                             '
            END IF                                                '
         NEXT i                                                   '
         FOR i = 1 TO c                                           ' Build a line
            IF j <= KbdT.KbdPFShowNum THEN                        ' Still PFSHOW entries left?
               txt5 += LSET$(MID$(KbdT.KbdPFShow(j), 2), KbdT.KbdPFShowMax + 2) '
               INCR j                                             '
            END IF                                                '
         NEXT i                                                   '
         IF ENV.PFKShow > 0 THEN _
            sPrtHelp(LSET$(txt1, ENV.ScrWidth), gwScrHeight + 1)
         IF ENV.PFKShow > 1 THEN _
            sPrtHelp(LSET$(txt2, ENV.ScrWidth), gwScrHeight + 2)
         IF ENV.PFKShow > 2 THEN _
            sPrtHelp(LSET$(txt3, ENV.ScrWidth), gwScrHeight + 3)
         IF ENV.PFKShow > 3 THEN _
            sPrtHelp(LSET$(txt4, ENV.ScrWidth), gwScrHeight + 4)
         IF ENV.PFKShow > 4 THEN _
            sPrtHelp(LSET$(txt5, ENV.ScrWidth), gwScrHeight + 5)
         MExit
      END METHOD

      METHOD ErrMAdd(MSG AS STRING, sev AS LONG) AS LONG                       ' Go Queue the message text
         IF ErrMsgTblC + 1 > UBOUND(ErrMsgTbl()) THEN             ' Exceeding current table size?
            REDIM PRESERVE ErrMsgTbl(1 TO 2 * UBOUND(ErrMsgTbl())) AS INSTANCE STRING' Make it bigger
         END IF
         ARRAY INSERT ErrMsgTbl(), MSG                            ' Insert it and push the rest down
         INCR ErrMsgTblC                                          ' Count it
         IF ErrMsgTblC = 1 THEN                                   ' If this is the 1st message of a block
            RESET ErrMsgHigh, ErrMsgHlpC, ErrMsgHlp()             ' Clear the old HELP copy
         END IF                                                   '
         ErrMsgHigh = IIF(sev > ErrMsgHigh, sev, ErrMsgHigh)      ' Save new high RC
      END METHOD

      METHOD ErrMHGet(i AS LONG) AS STRING: METHOD = ErrMsgHlp(i): END METHOD

      METHOD FileWatch (WatchFile AS STRING, WatchFunc AS LONG) AS LONG
      '---------- Setup a directory watch
      LOCAL hSearch    AS DWORD                                   '
      LOCAL i, j AS LONG, strbuffer AS ASCIIZ * 255               '
      LOCAL FD         AS DIRDATA                                 '
      LOCAL SleepCtr AS LONG

         MEntry
         SELECT CASE WatchFunc                                    ' Split by what function requested
            CASE %WatchStart                                      ' %WatchStart
               j = VAL(sFileQueue("I", " ", WatchFile))           ' Already in FQ table?
               IF j = 0 THEN                                      ' If not, Whoops!
                  METHOD = %True                                  '
                  MExitMeth                                       ' Oops!  Exit
               END IF                                             '
               IF gFQ(j).gWatchThread <> 0 THEN                   ' Already got thread address
                  METHOD = %True                                  ' Say we failed
                  MExitMeth                                       '
               END IF                                             '
               hSearch = GetFileAttributesEx(gFQ(j).gWatchFile, &H0, BYREF FD) ' Search for the filename
               IF hSearch = 0 THEN                                ' No file, error
                  METHOD = %True                                  '
                  MExitMeth                                       ' Oops!  Exit
               END IF                                             '
               gFQ(j).gFileAttrib = FD.FileAttributes             ' Save Status data
               gFQ(j).gFileLWTime = FD.LastWriteTime              '
               gFQ(j).gFileCRTime = FD.CreationTime               '
               gFQ(j).gFileSizeHigh = FD.FileSizeHigh             '
               gFQ(j).gFileSizeLow = FD.FileSizeLow               '
               gFQ(j).gActive = %False                            '
               gFQ(j).gEvent = CreateEvent(BYVAL %NULL, %TRUE, %FALSE, BYVAL %NULL)
               IF gFQ(j).gEvent = 0 THEN                          '
                  i = GetLastError                                '
                  FormatMessage %FORMAT_MESSAGE_FROM_SYSTEM, BYVAL %NULL, i, %NULL, strBuffer, SIZEOF(strBuffer), BYVAL %NULL
                  METHOD = %True                                  '
                  MExitMeth                                       ' Oops!  Exit
               END IF                                             '
               gFQ(j).gChanged = %False                           ' The change flag
               pWatchData = VARPTR(gFQ(j))                        ' Setup the pointer
               THREAD CREATE sFileWatchThread(pWatchData) 65536, TO gFQ(j).gWatchThread ' Fire up the thread
               IF gFQ(j).gWatchThread = 0 THEN                    ' Failed?
                  METHOD = %True                                  ' Say we failed
                  MExitMeth                                       '
               END IF                                             '
               DO WHILE gFQ(j).gActive = %False                   ' Wait till thread starts
                  SLEEP 50                                        ' Wait a bit
               LOOP                                               '
               THREAD STATUS gFQ(j).gWatchThread TO wResult       ' See if running OK
               IF wResult <> 259 THEN                             ' If running OK STATUS returns &H103 (See Help)
                  METHOD = %True                                  ' Flag an error
               ELSE                                               '
                  METHOD = %False                                 ' Else, Say we started it
               END IF                                             '
               THREAD CLOSE gFQ(j).gWatchThread TO wResult        ' Free up our handle
               MExitMeth

            CASE %WatchEnd                                        ' %WatchEnd
               i = %False
               IF ISNULL(WatchFile) THEN                          ' If a normal End
                  FOR j = 1 TO UBOUND(gFQ())                      ' Just search for Tab number
                     IF gFQ(j).gPgNumber = PgNumber AND _         ' An entry for this tab?
                        gFQ(j).gInUse = %True THEN                '
                        i = %True                                 '
                        EXIT FOR                                  '
                     END IF                                       '
                  NEXT j                                          '
                  IF ISFALSE i THEN                               ' Didn't set j
                     METHOD = %True                               '
                     MExitMeth                                    ' Oops!  Exit
                  END IF                                          '
               ELSE                                               '
                  j = VAL(sFileQueue("I", " ", WatchFile))        ' Fetch gFQ() index
                  IF j = 0 THEN                                   ' If not, Whoops!
                     METHOD = %True                               '
                     MExitMeth                                    ' Oops!  Exit
                  END IF                                          '
               END IF                                             '

               IF gFQ(j).gWatchThread = 0 THEN                    ' Never started?
                  METHOD = %False                                 ' Just return False
               ELSE                                               '
                  i = SetEvent(gFQ(j).gEvent)                     ' Tell watch thread to quit now
                  IF i <> 0 THEN                                  ' Successful SetEvent
                     DO WHILE gFQ(j).gActive = %True              ' Wait till thread goes inactive
                        SLEEP 10                                  ' Wait a bit
                     LOOP
                  END IF
                  METHOD = gFQ(j).gChanged                        ' Copy the Changed flag result
                  gFQ(j).gWatchThread = 0                         '
               END IF                                             '
         END SELECT
         MExit
      END METHOD                                                  '

      METHOD FindPushPop(which AS STRING)
      STATIC lcfFind, lcfChange, lcfOChange AS STRING
      STATIC lcfFLen, lcfCLen, lcfLMargin, lcfRMargin, lcfFLine AS LONG
      STATIC lcfFCol, lcfSplitPt1, lcfSplitpt2, lcfNotFound, lcfHiLiteClr, lcfHiLiteOff, lcfHiLiteOn AS LONG
      STATIC lcfFlag AS QUAD
         '---------- Save / Restore the saved search data
         IF which = "PUSH" THEN
            lcfFind      = cfFind:      lcfChange    = cfChange:     lcfOChange   = cfOChange
            lcfFLen      = cfFLen:      lcfCLen      = cfCLen:       lcfLMargin   = cfLMargin
            lcfRMargin   = cfRMargin:   lcfFLine     = cfFLine:      lcfFCol      = cfFCol
            lcfSplitPt1  = cfSplitPt1:  lcfSplitPt2  = cfSplitPt2:   lcfNotFound  = cfNotFound
            lcfFlag      = cfFlag:      lcfHiLiteClr = cfHiLiteClr:  lcfHiLiteOff = cfHiLiteOff
            lcfHiLiteOn  = cfHiLiteOn
         ELSE
            cfFind       = lcfFind:     cfChange     = lcfChange:    cfOChange    = lcfOChange
            cfFLen       = lcfFLen:     cfCLen       = lcfCLen:      cfLMargin    = lcfLMargin
            cfRMargin    = lcfRMargin:  cfFLine      = lcfFLine:     cfFCol       = lcfFCol
            cfSplitPt1   = lcfSplitPt1: cfSplitPt2   = lcfSplitPt2:  cfNotFound   = lcfNotFound
            cfFlag       = lcfFlag:     cfHiLiteClr  = lcfHiLiteClr: cfHiLiteOff  = lcfHiLiteOff
            cfHiLiteOn   = lcfHiLiteOn
         END IF
      END METHOD

      METHOD FindReset()
      '---------- Reset the FIND control areas
         cfFind = ""                                              ' Reset the parse output fields
         cfChange = ""                                            '
         cfOChange = ""                                           '
         cfFLen = 0                                               '
         cfCLen = 0                                               '
         cfLMargin = 1                                            '
         cfRMargin= 0                                             '
         cfFlag = 0                                               '
         cfFLine = 0                                              '
         cfFCol = 0                                               '
         cfSplitPt1 = 0                                           '
         cfSplitPt2 = 0                                           '
         cfNotFound = 0                                           '
         cfHiLiteClr = 0                                          '
         cfHiLiteOff = 0                                          '
         cfHiLiteOn = 0                                           '
         OffColSupp                                               '

         cfFindT = ""                                             '
         cfChangeT = ""                                           '
         cfOChangeT = ""                                          '
         cfFLenT = 0                                              '
         cfCLenT = 0                                              '
         cfLMarginT = 1                                          '
         cfRMarginT = 0                                           '
         cfFlagT = 0                                              '
         cfFLineT = 0                                             '
         cfFColT = 0                                              '
         cfSplitPt1T = 0                                          '
         cfSplitPt2T = 0                                          '
         cfNotFoundT = 0                                          '
         cfHiLiteClrT = 0                                         '
         cfHiLiteOffT = 0                                         '
         cfHiLiteOnT = 0                                          '
      END METHOD

      METHOD FindSave()
      '---------- Save the FIND control areas
         cfFindT = cfFind                                         ' Save the RFIND stuff
         cfChangeT = cfChange                                     '
         cfOChangeT = cfOChange                                   '
         cfFLenT = cfFLen                                         '
         cfCLenT = cfCLen                                         '
         cfLMarginT = cfLMargin                                   '
         cfRMarginT = cFRMargin                                   '
         cfFlagT = cfFlag                                         '
         cfFLineT = cfFLine                                       '
         cfFColT = cfFCol                                         '
         cfSplitPt1T = cfSplitPt1                                 '
         cfSplitPt2T = cfSplitPt2                                 '
         cfNotFoundT = cfNotFound                                 '
         cfHiLiteClrT = cfHiLiteClr                               '
         cfHiLiteOffT = cfHiLiteOff                               '
         cfHiLiteOnT = cfHiLiteOn                                 '
         DotSetT =   DotSet                                       '
         DotStartT = DotStart                                     '
         DotSCondT = DotSCond                                     '
         DotAndOrT = DotAndOr                                     '
         DotEndT =   DotEnd                                       '
         DotECondT = DotECond                                     '
         DotFlagT =  DotFlag                                      '
         DotTagT =   DotTag                                       '
         DotTCondT = DotTCond                                     '
         DotModeT =  DotMode                                      '
      END METHOD

      METHOD FindLOAD()
      '---------- ReLoad the FIND control areas
         cfFind = cfFindT                                         ' ReLoad the RFIND stuff
         cfChange = cfChangeT                                     '
         cfOChange = cfOChangeT                                   '
         cfFLen = cfFLenT                                         '
         cfCLen = cfCLenT                                         '
         cfLMargin = cfLMarginT                                   '
         cfRMargin= cfRMarginT                                    '
         cfFlag = cfFlagT                                         '
         cfFLine = cfFLineT                                       '
         cfFCol = cfFColT                                         '
         cfSplitPt1 = cfSplitPt1T                                 '
         cfSplitPt2 = cfSplitPt2T                                 '
         cfNotFound = cfNotFoundT                                 '
         cfHiLiteClr = cfHiLiteClrT                               '
         cfHiLiteOff = cfHiLiteOffT                               '
         cfHiLiteOn = cfHiLiteOnT                                 '
         DotSet =   DotSetT                                       '
         DotStart = DotStartT                                     '
         DotSCond = DotSCondT                                     '
         DotAndOr = DotAndOrT                                     '
         DotEnd =   DotEndT                                       '
         DotECond = DotECondT                                     '
         DotFlag =  DotFlagT                                      '
         DotTag =   DotTagT                                       '
         DotTCond = DotTCondT                                     '
         DotMode =  DotModeT                                      '
      END METHOD

      METHOD FindSetAll()
      '---------- Set the FIND control areas for ALL records
         cfFind = ""                                              ' Reset the parse output fields
         cfChange = ""                                            '
         cfOChange = ""                                           '
         cfFLen = 0                                               '
         cfCLen = 0                                               '
         cfLMargin = 1                                            '
         cfRMargin= 0                                             '
         cfFlag = 0                                               '
         BIT SET cfFlag, %CRTAll                                  '
         BIT SET cfFlag, %CRTDX                                   '
         cfFLine = 0                                              '
         cfFCol = 0                                               '
         cfSplitPt1 = 0                                           '
         cfNotFound = 0                                           '
         cfHiLiteClr = 0                                          '
      END METHOD

      METHOD  FindSetup()
      '---------- Setup the FIND control areas from the CRT scan data
      LOCAL f AS QUAD

         MEntry
         cfLMargin = Prf.BndLeft                                  ' Set defaults
         cfRMargin = Prf.BndRight                                 '
         cfFlag = CrtFlag                                         ' Copy the CRT flag area
         IF BIT(CrtFlag, %CrtFCol) THEN                           ' A From override?
            cfLMargin = CrtFCol                                   ' then set it
            BIT SET cfFlag, %CRTFCol                              ' Remember we set it
         END IF                                                   '
         IF BIT(CrtFlag, %CrtTCol) THEN                           ' A To override?
            cfRMargin = CrtTCol                                   '
            BIT SET cfFlag, %CrtTCol                              '
         END IF                                                   '
         IF cfRMargin = 0 THEN cfRMargin= MaxLength               '

         IF BIT(CrtFlag, %CRTLit1) THEN                           ' Got a Lit1
            cfFind = CrtL1                                        ' Copy the basic find string
            IF BIT(CrtFlag, %CrtL1Picture) OR BIT(CrtFlag, %CrtL1Hex) OR BIT(CrtFlag, %CrtL1RegEx) THEN _ ' Lit1 present?
               cfFind = CrtL1RData                                ' Save the massaged
            cfFLen = LEN(cfFind)                                  ' Save it's length
         END IF                                                   '

         cfHiLiteClr = CRTHiLiteClr                               ' Setup color operands
         cfHiLiteOff = CRTHiLiteOff                               '
         cfHiLiteOn  = CRTHiLiteOn                                '

         IF cfFind    = "*" AND CRTL1 = CRTL1Raw AND ISNOTNULL(pfFind) THEN  ' Request for previous search?
            cfFind    = pfFind                                    ' Copy previous find stuff
            cfFLen    = pfFLen                                    '
            BIT RESET cfFlag, %CrtL1Picture                       ' Clear L1 related flags
            BIT RESET cfFlag, %CrtL1Hex                           '
            BIT RESET cfFlag, %CrtL1RegEx                         '
            BIT RESET cfFlag, %CrtL1CaseComp                      '
            BIT RESET cfFlag, %CrtL1CaseInComp                    '
            IF BIT(pfFlag, %CrtL1Picture) THEN BIT SET cfFlag, %CrtL1Picture ' Transfer the old flags
            IF BIT(pfFlag, %CrtL1Hex) THEN BIT SET cfFlag, %CrtL1Hex '
            IF BIT(pfFlag, %CrtL1RegEx) THEN BIT SET cfFlag, %CrtL1RegEx
            IF BIT(pfFlag, %CrtL1CaseComp) THEN BIT SET cfFlag, %CrtL1CaseComp
            IF BIT(pfFlag, %CrtL1CaseInComp) THEN BIT SET cfFlag, %CrtL1CaseInComp
         END IF                                                   '

         IF BIT(CrtFlag, %CrtLit2) THEN                           ' Got a Lit2
            cfChange = CrtL2                                      ' Save it
            IF BIT(CrtFlag, %CrtL2Format) OR BIT(CrtFlag, %CrtL2Picture) OR BIT(CrtFlag, %CrtL2Hex) THEN _  ' Lit2 present?
               cfChange = CrtL2RData                              ' Save the massaged
            cfCLen = LEN(cfChange)                                ' Save it's length
            cfOChange = CrtL2Raw                                  ' Save Raw form
         END IF                                                   '

         IF cfChange  = "*" AND CrtL2 = CrtL2Raw AND ISNOTNULL(pfChange) THEN ' Request for previous change?
            cfChange  = pfChange                                  ' Copy previous change stuff
            cfOChange = pfOChange                                 '
            cfCLen    = pfCLen                                    '
            BIT RESET cfFlag, %CrtL2Picture                       ' Clear L2 related flags
            BIT RESET cfFlag, %CrtL2Hex                           '
            BIT RESET cfFlag, %CrtL2CaseComp                      '
            BIT RESET cfFlag, %CrtL2CaseInComp                    '
            BIT RESET cfFlag, %CrtL2Map                           '
            BIT RESET cfFlag, %CrtL2EXEC                          '
            IF BIT(pfFlag, %CrtL2Picture) THEN BIT SET cfFlag, %CrtL2Picture ' Transfer the old flags
            IF BIT(pfFlag, %CrtL2Hex) THEN BIT SET cfFlag, %CrtL2Hex
            IF BIT(pfFlag, %CrtL2CaseComp) THEN BIT SET cfFlag, %CrtL2CaseComp
            IF BIT(pfFlag, %CrtL2CaseInComp) THEN BIT SET cfFlag, %CrtL2CaseInComp
            IF BIT(pfFlag, %CrtL2Map) THEN BIT SET cfFlag, %CrtL2Map
            IF BIT(pfFlag, %CrtL2Exec) THEN BIT SET cfFlag, %CrtL2Exec
         END IF
         MExit
      END METHOD

      METHOD FMPrtFileName(stp AS LONG, sp AS LONG, ForceNameHi AS LONG)
      '---------- Format and print one FM filename
      LOCAL tText1, tText2 AS STRING, i, j AS LONG, lclScheme AS WSTRING
         tText1 = IIF$(FMode = %FMPath OR FMode = %FMNamed OR FMode = %FMProfiles OR FileListNm <> "", _ THEN ' FLIST just gets filename too
                      TRIM$(AFList(stp).FD.FileName), _           ' Build just filename
                      TRIM$(AFList(stp).Path) + TRIM$(AFList(stp).FD.FileName)) ' Build full path\filename
         tText1 = me.FileAbbrev(tText1, FM_File_Size)             ' Abbreviate it as needed

         tText2 = SPACE$(10)                                      ' Default to no extension
         i = INSTR(-1, tText1, "."): j = INSTR(-1, tText1, "\")   ' Get locations of last . and \
         IF i AND  i > j THEN tText2 = MID$(tText1, i)            ' If . is right of \ then we have an extension
         tText2 = LEFT$(tText2, 8) + "  "                         ' Corect length and add pad
         IF AFList(stp).Flag = %FDirUp THEN tText1 = LSET$("..\", FM_File_Size): tText2 = SPACE$(10)
         tText1 += " "                                           ' Add pad

         '----- Choose color for the filename printing
         IF LEFT$(DefSort, 4) <> "Name" AND LEFT$(DefSort, 4) <> "Extn" THEN ' Sorted on other than Filename?
            lclScheme = IIF$(ForceNameHi, $$TxtHi, $$TxtLo)       ' Set scheme
            GOSUB PrtFName                                        ' Print filename
            lclScheme = $$TxtLo                                   ' Set Lo
            GOSUB PrtExtn                                         ' Print Extension

         ELSEIF LEFT$(DefSort, 4) = "Name" THEN                   ' Sorted on Filename?
            lclScheme = IIF$(ForceNameHi, $$TxtLo, $$TxtHi)       ' Set scheme
            GOSUB PrtFName                                        ' Print filename
            lclScheme = $$TxtLo                                   ' Set Lo
            GOSUB PrtExtn                                         ' Print Extension

         ELSEIF LEFT$(DefSort, 4) = "Extn" THEN                   ' Sorted on Extension?
            lclScheme = IIF$(ForceNameHi, $$TxtHi, $$TxtLo)       ' Set scheme
            GOSUB PrtFName                                        ' Print filename
            lclScheme = $$TxtHi                                   ' Set Hi
            GOSUB PrtExtn                                         ' Print Extension
         END IF                                                   '

         EXIT METHOD                                              '

         PrtFName:
            sPrint (tText1, lclScheme, sp, FM_Head_Name_Left)     '
            Rightmost = MAX(Rightmost, FM_Head_Name_Left + LEN(tText1) - 1)
            RETURN
         PrtExtn:
            IF FM_Head_Ext_Left > 0 THEN                          ' Splitting off Ext?
               sPrint (tText2, lclScheme, sp, FM_Head_Ext_Left)   '
               Rightmost = MAX(Rightmost, FM_Head_Ext_Left + 9)   '
            END IF
            RETURN

      END METHOD

      METHOD IndentAsPrev(l AS LONG) AS LONG
      '---------- Return an indent value = to indent of previous line
      LOCAL i AS LONG
         MEntry
         i = me.CRPBack(%mData, l, 1)                             ' Backup to previous Data line
         IF ISNULL(TRIM$(LTxtG(i))) THEN                          ' If line empty, return 1
            METHOD = 1                                            '
         ELSE                                                     '
            METHOD = VERIFY(LTxtG(i), " ")                        ' Get the indent
         END IF                                                   '
         MExit                                                    '
      END METHOD

      METHOD InitaFile(quick AS LONG)
      '---------- Initialize file control stuff
      LOCAL i AS LONG, t AS STRING
         MEntry
         '----- Fudge CLIP and SET EDIT modes
         IF ISClipLoad THEN                                       ' Are we to load the clipboard?
            Prf.SetProfName("DEFAULT", %False)                    ' Use DEFAULT
            Prf.ReadAll(%True)                                    ' Go read the default values
            ClipName = ENV.CLIPCmd                                ' Save the clipboard name
            ENV.ClipCmd = ""                                      ' Reset the clipboard name
            pCmdPASTE("PASTE " + ClipName)                        '
            OffModdFlag                                           '

         ELSE                                                     '
            '----- We have a file to OPEN
            IF TIP.FilePath <> $Empty THEN                        ' Passed a filename?
               Prf.SetProfName(TIP.Profile, %False)               ' Setup the Profile

               '----- Get file data loaded
               TopScrn = 1                                        ' N.B. May be overridden by StateLoad

               me.CopyAFile(1, TIP, Prf, 0, 0, 0, quick)          ' Go load the data, MIX is always 1

               '----- Add to OPEN queue
               i = me.FileWatch("", %WatchEnd)                    ' Kill any prior Watch
               sFileQueue("A", " ", TIP.FilePath)                 ' Add to Open queue

               '----- If not quick (i.e. FF command mode) start the File Watch
               IF ISFALSE quick THEN                              '
                  IF me.FileWatch(TIP.FilePath, %WatchStart) THEN ' Establish the watch
                     scError(0, "File watch could not be established")  '
                  END IF                                          '
               END IF                                             '

               OffModdFlag                                        ' Start off as not modified

               '----- If a MEDIT session we need a bit more work
               IF IsMEdit THEN                                    ' Starting a MEdit session?
                  me.MeditTbl("A", TIP.FilePath)                  ' Add to Medit table
                  me.LInsertEmpty(1, 1, %File)                    ' Insert for the =FILE> line
                  LMixS(2) = 1                                    ' Mark as File 1
                  me.LTxtSet(2, TIP.FilePath)                     ' Stuff filename in as the text
                  me.UpdLControl(2)                               ' Setup LLCtl
               END IF                                             '

            END IF                                                '
         END IF                                                   '

         SetCmd                                                   ' Cursor to Cmd line
         IF ISFALSE quick THEN                                    '
            OnUndoFlag                                            ' Take a 1st save checkpoint
            me.WindowTitle                                        ' Alter window/Tab titles
         END IF                                                   '
         sDoStatusBar($AllStatusBarBoxes)                         ' re-Do the StatusBar box
         MExit
      END METHOD

      METHOD InsBounds(lno AS LONG)
         me.LInsertEmpty(lno - 1, 1, %Bounds)                     ' Insert the requested line(s)
         me.AdjustPending(lno - 1, 1, 0)                          ' Adjust pending stuff
         me.LTxtFree(lno)                                         ' Go free the normal dynamic string
         me.LTxt2Bnds(lno)                                        ' Point L.Txt to the Bnds data
         me.UpdLControl(lno)                                      ' Update the lineno area
      END METHOD

      METHOD InsCols(lno AS LONG)
         me.LInsertEmpty(lno - 1, 1, %Cols)                       ' Insert a line
         me.AdjustPending(lno - 1, 1, 0)                          ' Adjust pending stuff
         me.UpdLControl(lno)                                      ' Setup LLCtl
      END METHOD

      METHOD InsMark(lno AS LONG)
         me.LInsertEmpty(lno - 1, 1, %Mark)                       ' Insert a line
         me.AdjustPending(lno - 1, 1, 0)                          ' Adjust pending stuff
         me.LTxtFree(lno)                                         ' Go free the normal dynamic string
         me.LTxt2Mark(lno)                                        ' Point L.Txt to the Mark Data
         me.UpdLControl(lno)                                      ' Setup LLCtl
      END METHOD

      METHOD InsMask(lno AS LONG)
         me.LInsertEmpty(lno - 1, 1, %Mask)                       ' Insert a line
         me.AdjustPending(lno - 1, 1, 0)                          ' Adjust pending stuff
         me.LTxtFree(lno)                                         ' Go free the normal dynamic string
         me.LTxt2Mask(lno)                                        ' Point L.Txt to the Mark Data
         me.UpdLControl(lno)                                      ' Setup LLCtl
      END METHOD

      METHOD InsProf(lno AS LONG, which AS LONG)
         me.LInsertEmpty(lno - 1, 1, %Prof)                       ' Insert a line
         me.AdjustPending(lno - 1, 1, 0)                          ' Adjust pending stuff
         me.LTxtFree(lno)                                         ' Go free the normal dynamic string
         SELECT CASE AS CONST which                               ' Point correct PROF line
            CASE 1: me.LTxt2Prof1(lno)                            '
            CASE 2: me.LTxt2Prof2(lno)                            '
            CASE 3: me.LTxt2Prof3(lno)                            '
            CASE 4: me.LTxt2Prof4(lno)                            '
            CASE 5: me.LTxt2Prof5(lno)                            '
            CASE 6: me.LTxt2Prof6(lno)                            '
         END SELECT
         me.UpdLControl(lno)                                      ' Setup LLCtl
      END METHOD

      METHOD InsTabs(lno AS LONG)
         me.LInsertEmpty(lno - 1, 1, %Tabs)                       ' Insert a line
         me.AdjustPending(lno - 1, 1, 0)                          ' Adjust pending stuff
         me.LTxtFree(lno)                                         ' Go free the normal dynamic string
         me.LTxt2Tabs(lno)                                        ' Point L.Txt to the Tabs Data
         me.UpdLControl(lno)                                      ' Setup LLCtl
      END METHOD

      METHOD InsWord(lno AS LONG)
         me.LInsertEmpty(lno - 1, 1, %Word)                       ' Insert a line
         me.AdjustPending(lno - 1, 1, 0)                          ' Adjust pending stuff
         me.LTxtFree(lno)                                         ' Go free the normal dynamic string
         me.LTxt2Word(lno)                                        ' Point L.Txt to the WORD Data
         me.UpdLControl(lno)                                      ' Setup LLCtl
      END METHOD


      METHOD InsTblAdd(lno AS LONG)
      REGISTER i AS LONG
         MEntry
         INCR InsCount                                            ' Incr count
         IF InsCount > UBOUND(InsTbl()) THEN _                    ' Keep table big enough
            REDIM PRESERVE InsTbl(1 TO InsCount + 500) AS GLOBAL LONG
         InsTbl(InsCount) = lno                                   ' Save Insert Line
         FOR i = 1 TO InsCount                                    ' Adjust if a new insert lower than existing
            IF InsTbl(i) > lno THEN INCR InsTbl(i)                '
         NEXT i                                                   '
         MExit
      END METHOD


      METHOD LAttrGet (ix AS LONG) AS WSTRING: METHOD = L(ix).@LAttr: END METHOD
      METHOD LAttrLen (ix AS LONG) AS LONG: METHOD = LEN(L(ix).@LAttr): END METHOD
      METHOD LAttrPtr (ix AS LONG) AS LONG: METHOD = L(ix).LAttr: END METHOD

      METHOD LAttrAdjust(ix AS LONG, icol AS LONG, iamt AS LONG)
      LOCAL ichr AS WSTRING
         IF iamt > 0 THEN                                         ' Doing an insert
            ichr = REPEAT$(iamt, MID$(L(ix).@LAttr, icol, 1))     ' Get char at insert position (repeated)
            IF icol = 1 THEN ichr = REPEAT$(iamt, CHR$$(0))       '
            L(ix).@LAttr = STRINSERT$(L(ix).@LAttr, ichr, icol)   '
         ELSE                                                     ' Doing a delete
            L(ix).@LAttr = STRDELETE$(L(ix).@LAttr, icol, ABS(iamt)) ' Remove characters
         END IF                                                   '
      END METHOD

      METHOD LFlagGet (ix AS LONG) AS LONG: METHOD = L(ix).LFlag: END METHOD
      METHOD LWrk2Get (ix AS LONG) AS LONG: METHOD = L(ix).LWrk2: END METHOD
      METHOD LLblGet (ix AS LONG) AS STRING: METHOD = L(ix).LLbl: END METHOD
      METHOD LLblSet (ix AS LONG, vl AS STRING): L(ix).LLbl = vl: END METHOD
      METHOD LLCtlGet (ix AS LONG) AS STRING: METHOD = L(ix).LLCtl: END METHOD
      METHOD LLCtlSet (ix AS LONG, vl AS STRING): L(ix).LLCtl = vl: END METHOD
      METHOD LLCtlTagSet(ix AS LONG)
       IF L(ix).LTag <> $BlankLNo THEN L(ix).LLCtl = L(ix).LTag
        END METHOD
      METHOD LLCtlLblSet(ix AS LONG)
       IF L(ix).LLbl <> $BlankLNo THEN L(ix).LLCtl = L(ix).LLbl
        END METHOD
      METHOD LLNumGet (ix AS LONG) AS STRING: METHOD = IIF$(ISTRUE (L(ix).LFlag AND %Data), L(ix).LLNum, "00000000"): END METHOD
      METHOD LMIXGet (ix AS LONG) AS LONG: METHOD = L(ix).LMix: END METHOD
      METHOD LMIXSet (ix AS LONG, vl AS LONG): L(ix).LMix = vl: END METHOD
      METHOD LWrk1Get (ix AS LONG) AS LONG: METHOD = L(ix).LWrk1: END METHOD

      METHOD  LStubDelRange(lfrom AS LONG, lto AS LONG)
      '---------- Delete a range of lines
      REGISTER i AS LONG
         MEntry
         IF IsSlecSet THEN                                        ' Need we check Slec line range?
            IF (lfrom >= me.LineNoRef(FORMAT$(SlecSLin)) AND lfrom <= me.LineNoRef(FORMAT$(SlecELin))) OR _
               (lto   >= me.LineNoRef(FORMAT$(SlecSLin)) AND lto   <= me.LineNoRef(FORMAT$(SlecELin))) THEN
               OffSlecSet                                         ' Clear Slec status
               OffSlecActive                                      '
               sDoStatusBar($SBSelect)                            ' re-Do the StatusBar Select box
               me.MarkKill                                        ' Reset the select frame
            END IF                                                '
         END IF                                                   '
         FOR i = lfrom TO lto                                     ' Loop for each deleted line
            DECR LastLine                                         ' Adjust Last Line
            IF IsLData(i) THEN DECR LastReal                      ' If Data, Adjust LastReal
            me.lTxtFree(i)                                        ' Go free the dynamic string
         NEXT i                                                   '
         me.LEntDelRange(lfrom, lto - lfrom +1)                   ' Remove from the L() array
         MExit
      END METHOD

      METHOD  lStubDoPM(flag AS LONG, lno AS LONG)
      '---------- Do +/- processing for a line
         MEntry
         IF BIT(flag, %lCmdX) THEN                                ' Do the +/- processing for the line
            IF ISFALSE IsLFile(lno) THEN me.LFlagBitOn(lno, %Invisible)' If - make Invisible (except =FILE> lines)
            gfXRebuild = %True                                    ' Force exclude rebuild

         ELSEIF BIT(flag, %lCmdNX) THEN                           '
            me.LFlagBitOff(lno, %Invisible)                       ' If + make visible
            gfXRebuild = %True                                    ' Force exclude rebuild
         END IF                                                   '
         MExit
      END METHOD

      METHOD LTagGet (BYVAL ix AS LONG) AS STRING: METHOD = L(ix).LTag: END METHOD

      METHOD LTblSize(needed AS LONG)
      '----- Expand T() and TW() if running out of room
      REGISTER x AS LONG
         IF needed  > TTCtr  THEN                                 ' If running out of room in T() make it bigger
            REDIM PRESERVE T(UBOUNDT + 25000) AS INSTANCE STRING  ' Get room for another bunch of lines in T() array
            REDIM PRESERVE TW(UBOUNDT + 25000) AS INSTANCE WSTRING' Ditto for TW()
            TIDX = TIDX + STRING$(25000, "1")                     ' Make all new T() lines available
            TTCtr += 25000                                        ' Add to Free count
            FOR x = 1 TO UBOUND(L)                                ' Correct the @LTxt pointers in the L() array
               IF L(x).LTxt <> 0 THEN                             '
                  L(x).LTxt = VARPTR(T(L(x).LTxtIX))              ' Correct using the saved TxtIX value
               END IF                                             '
               IF L(x).LAttr <> 0 THEN                            '
                  L(x).LAttr = VARPTR(TW(L(x).LAttrIX))           ' Correct using the saved AttrIX value
               END IF                                             '
            NEXT x                                                '
            UBoundT += 25000                                      ' Save new UBound value
         END IF                                                   '
      END METHOD

      METHOD LTxtGet (BYVAL ix AS LONG) AS STRING: METHOD = L(ix).@LTxt: END METHOD
      METHOD LTxtSet (BYVAL ix AS LONG, vl AS STRING)
      LOCAL t AS STRING, i AS LONG
         IF Prf.MINLEN > 0 AND LEN(vl) < Prf.MINLEN THEN
            t = LSET$(vl, Prf.MINLEN)
            L(ix).@LTxt = t
         ELSE
            L(ix).@LTxt = vl
         END IF
         MaxLength = MAX(MaxLength, LEN(L(ix).@LTxt))             ' Track MaxLength
      END METHOD
      METHOD LFlagData (BYVAL ix AS LONG) AS LONG: METHOD = ISTRUE (L(ix).LFlag AND %Data): END METHOD
      METHOD LFlagTop (BYVAL ix AS LONG) AS LONG: METHOD = ISTRUE (L(ix).LFlag AND %Top): END METHOD
      METHOD LFlagBottom (BYVAL ix AS LONG) AS LONG: METHOD = ISTRUE (L(ix).LFlag AND %Bottom): END METHOD
      METHOD LFlagWord (BYVAL ix AS LONG) AS LONG: METHOD = ISTRUE (L(ix).LFlag AND %Word): END METHOD
      METHOD LFlagTabs (BYVAL ix AS LONG) AS LONG: METHOD = ISTRUE (L(ix).LFlag AND %Tabs): END METHOD
      METHOD LFlagBounds (BYVAL ix AS LONG) AS LONG: METHOD = ISTRUE (L(ix).LFlag AND %Bounds): END METHOD
      METHOD LFlagCols (BYVAL ix AS LONG) AS LONG: METHOD = ISTRUE (L(ix).LFlag AND %Cols): END METHOD
      METHOD LFlagXclude (BYVAL ix AS LONG) AS LONG: METHOD = ISTRUE (L(ix).LFlag AND %Xclude): END METHOD
      METHOD LFlagInvisible (BYVAL ix AS LONG) AS LONG: METHOD = ISTRUE (L(ix).LFlag AND %Invisible): END METHOD
      METHOD LFlagMark (BYVAL ix AS LONG) AS LONG: METHOD = ISTRUE (L(ix).LFlag AND %Mark): END METHOD
      METHOD LFlagMask (BYVAL ix AS LONG) AS LONG: METHOD = ISTRUE (L(ix).LFlag AND %Mask): END METHOD
      METHOD LFlagNote (BYVAL ix AS LONG) AS LONG: METHOD = ISTRUE (L(ix).LFlag AND %Note): END METHOD
      METHOD LFlagProf (BYVAL ix AS LONG) AS LONG: METHOD = ISTRUE (L(ix).LFlag AND %Prof): END METHOD
      METHOD LFlagUser (BYVAL ix AS LONG) AS LONG: METHOD = ISTRUE (L(ix).LFlag AND %User): END METHOD
      METHOD LFlagPage (BYVAL ix AS LONG) AS LONG: METHOD = ISTRUE (L(ix).LFlag AND %Page): END METHOD
      METHOD LFlagFile (BYVAL ix AS LONG) AS LONG: METHOD = ISTRUE (L(ix).LFlag AND %File): END METHOD
      METHOD LFlagBitOn (BYVAL ix AS LONG, vl AS LONG): L(ix).LFlag = (L(ix).LFlag OR VL): END METHOD
      METHOD LFlagBitOff (BYVAL ix AS LONG, vl AS LONG): L(ix).LFlag = (L(ix).LFlag AND (&HFFFFFFFF - vl)): END METHOD
      METHOD LFlagBitTog (BYVAL ix AS LONG, vl AS LONG): L(ix).LFlag = (L(ix).LFlag XOR vl): END METHOD

      METHOD LTxtCharDel (BYVAL ix AS LONG, posit AS LONG)
         LOCAL t AS STRING POINTER: t = LTxtGP(ix)
         L(ix).@LTxt = STRDELETE$(@t, posit, 1)
         IF Prf.MINLEN > 0 AND Prf.MINLEN > LEN(@t) THEN L(ix).@LTxt = LSET$(@t, Prf.MINLEN)
      END METHOD
      METHOD LTxtCharRep (BYVAL ix AS LONG, posit AS LONG, char AS STRING): MID$(L(ix).@LTxt, posit, 1) = char: END METHOD
      METHOD LTxtLen (BYVAL ix AS LONG) AS LONG: METHOD = LEN(L(ix).@LTxt): END METHOD
      METHOD LTxt2Bnds (BYVAL ix AS LONG): L(ix).LTxt = VARPTR(BndText): END METHOD
      METHOD LTxt2Mark (BYVAL ix AS LONG): L(ix).LTxt = VARPTR(MarkLine): END METHOD
      METHOD LTxt2Mask (BYVAL ix AS LONG): L(ix).LTxt = VARPTR(MaskLine): END METHOD
      METHOD LTxt2Prof1 (BYVAL ix AS LONG): L(ix).LTxt = VARPTR(Prof1): END METHOD
      METHOD LTxt2Prof2 (BYVAL ix AS LONG): L(ix).LTxt = VARPTR(Prof2): END METHOD
      METHOD LTxt2Prof3 (BYVAL ix AS LONG): L(ix).LTxt = VARPTR(Prof3): END METHOD
      METHOD LTxt2Prof4 (BYVAL ix AS LONG): L(ix).LTxt = VARPTR(Prof4): END METHOD
      METHOD LTxt2Prof5 (BYVAL ix AS LONG): L(ix).LTxt = VARPTR(Prof5): END METHOD
      METHOD LTxt2Prof6 (BYVAL ix AS LONG): L(ix).LTxt = VARPTR(Prof6): END METHOD
      METHOD LTxt2Word (BYVAL ix AS LONG): L(ix).LTxt = VARPTR(WordInput): END METHOD
      METHOD LTxt2Tabs (BYVAL ix AS LONG): L(ix).LTxt = VARPTR(TabsLine): END METHOD
      METHOD LEntDel (BYVAL ix AS LONG)
      REGISTER i AS LONG
         i = UBOUND(L())                                          ' Point at last entry
         ARRAY DELETE L(ix)                                       ' Delete the entry
         OnRenumFlag                                              ' Call for rename
         L(i).LTxt = 0                                            ' Reset the last entry
         L(i).LTxtIX = 0                                          '
         L(i).LAttr = 0                                           '
         L(i).LAttrIX = 0                                         '
         L(i).LFlag = 0                                           '
         L(i).LCol = 0                                            '
         L(i).LLCtl = $BlankLNo                                   '
         L(i).LLNum = $BlankLNo                                   '
         L(i).LLbl = $BlankLNo                                    '
      END METHOD

      METHOD LEntDelRange (BYVAL ix AS LONG, cnt AS LONG)
      REGISTER i AS LONG
      REGISTER j AS LONG

         MEMORY COPY VARPTR(L(ix + cnt)), VARPTR(L(ix)), SIZEOF(DataLine) * (UBoundL - ix - cnt + 1)

         FOR i = UBoundL - cnt + 1 TO UBoundL                     ' Pad the entries at the end
            L(i).LTxt = 0                                         '
            L(i).LTxtIX = 0                                       '
            L(i).LAttr = 0                                        '
            L(i).LAttrIX = 0                                      '
            L(i).LFlag = 0                                        '
            L(i).LCol = 0                                         '
            L(i).LLCtl = $BlankLNo                                '
            L(i).LLNum = $BlankLNo                                '
            L(i).LLbl = $BlankLNo                                 '
         NEXT i                                                   '
         LCtr += cnt                                              ' Add back to available L() count
         OnRenumFlag                                              '
      END METHOD

      METHOD LLCtlCharRep (BYVAL ix AS LONG, posit AS LONG, char AS STRING)
         MID$(L(ix).LLCtl, posit, 1) = char
      END METHOD

      METHOD LLCtlScan (BYVAL key AS STRING) AS LONG
         LOCAL k AS LONG
         ARRAY SCAN L(), FROM 1 TO 8, COLLATE UCASE, = LSET$(key, 8), TO k
         METHOD = k
      END METHOD

      METHOD LLNumScan (BYVAL key AS STRING) AS LONG
      LOCAL x, y, z AS STRING
         LOCAL k AS LONG
         LOCAL i AS LONG
         ARRAY SCAN L(), FROM 9 TO 16, = LSET$(key, 8), TO k
         METHOD = k
         x = L(2).LLNum: y = L(2).@LTxt
      END METHOD

      METHOD LEntXPop (BYVAL ix AS LONG, cnt AS LONG)
         LOCAL i, j AS LONG, tmp AS DataLine
         MEntry
         L(ix).LWrk1 = L(ix).LWrk1 - i                            ' Adjust # in reduced range
         tmp = L(ix)                                              ' Save the XX marker line
         ARRAY DELETE L(ix)                                       ' Delete the XX marker line from its old location
         FOR j = ix TO ix + cnt - 1                               ' Un-'hide' the lines in the range
            L(j).LFlag = (L(j).LFlag AND (&HFFFFFFFF - %Invisible)) '
            L(j).LWrk1 = 0                                        ' Reset work field
         NEXT j                                                   '
         ARRAY INSERT L(ix + cnt), tmp                            ' Stuff saved XX marker back in
         FOR j = 1 TO tmp.LWrk1                                   ' Redo Work numbers in the new range
            L(ix + j + cnt).LWrk1 = j                             '
         NEXT j                                                   '
         MExit
      END METHOD                                                  '

      METHOD LLTagScan (BYVAL key AS STRING) AS LONG
         LOCAL k AS LONG, lkey AS STRING * 8
         MEntry
         lkey = key
         IF lkey = ":Z      " OR lkey = ":ZF     " OR lkey = ":ZNF    " THEN
            METHOD = 1
            MexitMeth
         END IF
         ARRAY SCAN L(), FROM 25 TO 32, COLLATE UCASE, = lkey, TO k
         METHOD = k
         MExit
      END METHOD

      METHOD LTxtFree(IX AS LONG)
         ' ---------- Free up T() line from the main L() entry
         MEntry
         IF ISTRUE (L(IX).LFlag AND (%Tabs OR %Mark OR %Mask OR %Prof OR %Word OR %Bounds)) THEN MExitMeth ' Don't do this for 'special' lines
         MID$(TIDX, L(IX).LTxtIX, 1) = "1"                        ' Mark slot available
         L(IX).@LTxt = ""                                         ' Null the text string
         L(IX).LTxt = 0                                           ' Kill the pointer
         L(IX).LTxtIX = 0                                         ' Kill the IX value
         L(IX).@LAttr = ""                                        ' Null the Attr string
         L(IX).LAttr = 0                                          ' Kill the pointer
         L(IX).LAttrIX = 0                                        ' Kill the IX value
         MExit
      END METHOD

      METHOD LTxtAlloc(DataIX AS LONG)
      ' ---------- Allocate a T() and TW() line to the main L() entry
      REGISTER i AS LONG
      REGISTER x AS LONG
         MEntry
         me.LTblSize(5)                                           ' Ensure enough room in text table

         i = INSTR(TIDX, "1")                                     ' Scan for next a free entry
         MID$(TIDX, i, 1) = "0"                                   ' Mark it no longer available
         T(i) = "": TW(i) = ""                                    ' Null the strings
         L(DataIX).LTxtIX = i                                     ' Remember the index in the L() entry
         L(DataIX).LTxt = VARPTR(T(i))                            ' Point L.LTxt to the dynamic String
         L(DataIX).LAttrIX = i                                    ' Remember the index in the L() entry
         L(DataIX).LAttr = VARPTR(TW(i))                          ' Point L.LAttr to the dynamic String
         MExit
      END METHOD

      METHOD LTxtSort () AS STRING
      '---------- SORT the data
      DIM SCmd(1 TO 5) AS SortCtl
      REGISTER i AS LONG
      REGISTER j AS LONG
      LOCAL k, x, y, z, NumKey AS LONG
      LOCAL SNum AS LONG, sk, lclCmd, t, lclSave, cstring, lkey AS STRING, DidSort, Uniq, MarkU, delcount, basekey AS LONG
         MEntry
         FOR i = 1 TO 5                                           ' Init the sort key
            sCmd(i).AorD = " ": sCmd(i).CASE = Prf.PCase          '
            sCmd(i).sFrom = 0: sCmd(i).sTo = 0                    '
         NEXT i                                                   '
         IF pCmdNumOps = 0 THEN SCmd(1).AorD = "A": SCmd(1).sFrom = 1: SCmd(1).sTo = MaxLength ' Setup a default key
         FOR i = 0 TO 255                                         ' Build collate string
            cstring += CHR$(i)                                    '
         NEXT                                                     '
         IF Prf.SrceXlate OR Prf.ColateXlate THEN                 ' In non-ANSI
            me.Translate(cstring, Prf.GetCA2SPtr)                 ' If needed
         END IF                                                   '
         i = 1                                                    ' Start Cmd table index at 1
         CfSet(%CRTFirst)                                         ' Ensure we start at the top row

         FOR j = 1 TO pCmdNumOps                                  ' Process the remaining operands
            pCmdOps(j) = UUCASE(pCmdOps(j))                       ' Uppercase it
            IF VAL(pCmdOps(j)) > 0 THEN                           ' We got a #, see where to put it
               IF SCmd(i).sFrom = 0 THEN                          ' Haven't got From yet
                  SCmd(i).sFrom = VAL(pCmdOps(j))                 ' Now we do
               ELSEIF SCmd(i).sTo = 0 THEN                        ' If no To
                  SCmd(i).sTo = VAL(pCmdOps(j))                   ' Now we have a To
               ELSEIF i = 5 THEN                                  ' At our max?
                  METHOD = "8Excessive SORT operands"             ' Pass back error message
                  MExitMeth                                       '
               ELSE                                               '
                  IF SCmd(i).AorD = " " THEN SCmd(i).AorD = "A"   ' Assign direction if not already done
                  INCR i                                          ' Bump
                  SCmd(i).sFrom = VAL(pCmdOps(j))                 ' assign this as the From value
               END IF                                             '
               ITERATE FOR                                        '
            END IF                                                '

            IF pCmdOps(j) = "UNIQ" THEN                           ' UNIQ?
               uniq = %True                                       ' Add the flag
               ITERATE FOR                                        '
            END IF                                                '

            IF pCmdOps(j) = "MARKUNIQ" THEN                       ' MARKUNIQ?
               MarkU = %True                                      ' Add the flag
               ITERATE FOR                                        '
            END IF                                                '

            IF pCmdOps(j) = "X" THEN                              ' X?
               CfSet(%CRTX)                                       ' Add the flag
               ITERATE FOR                                        '
            END IF                                                '

            IF pCmdOps(j) = "NX" THEN                             ' NX?
               CfSet(%CRTNX)                                      ' Add the flag
               ITERATE FOR                                        '
            END IF                                                '

            IF pCmdOps(j) = "U" THEN                              ' U?
               CfSet(%CRTU)                                       ' Add the flag
               ITERATE FOR                                        '
            END IF                                                '

            IF pCmdOps(j) = "NU" THEN                             ' NU?
               CfSet(%CRTNU)                                      ' Add the flag
               ITERATE FOR                                        '
            END IF                                                '

            IF pCmdOps(j) = "DX" THEN                             ' DX?
               CfSet(%CRTDX)                                      ' Add the flag
               ITERATE FOR                                        '
            END IF                                                '

            IF pCmdOps(j) = "MX" THEN                             ' X?
               CfSet(%CRTMX)                                      ' Add the flag
               ITERATE FOR                                        '
            END IF                                                '

            SELECT CASE LEFT$(pCmdOps(j), 1)                      ' See if a direction/case operand
               CASE "A", "D"                                      ' A direction?
                  IF SCmd(i).AorD = " " THEN                      ' Save direction if not yet set
                     SCmd(i).AorD = LEFT$(pCmdOps(j), 1)          '
                     GOSUB Do2ndOper                              ' Go handle 2nd oper
                  ELSEIF SCmd(i).sFrom <> 0 AND SCmd(i).sTo <> 0 THEN
                     IF i = 5 THEN                                ' At max?
                        METHOD = "8Excessive SORT operands"       ' Pass back error
                        MExitMeth                                 '
                     ELSE                                         '
                        INCR i                                    ' Bump to next slot
                        SCmd(i).AorD = LEFT$(pCmdOps(j), 1)       '
                        GOSUB Do2ndOper                           ' Go handle 2nd oper
                     END IF                                       '
                  ELSE                                            '
                     METHOD =  "8Illogical SORT operands"         ' Pass back message
                     MExitMeth                                    '
                  END IF                                          '
               CASE "T", "C"                                      '
                  IF LEN(pCmdOps(j)) > 1 THEN                     ' Alone?
                     METHOD =  "8Invalid SORT operands"           ' Pass back message
                     MExitMeth                                    '
                  END IF
                  SCmd(i).CASE = LEFT$(pCmdOps(j), 1)             ' Save Case
               CASE ELSE                                          '
                  METHOD = "8Invalid SORT operands"               ' Pass back error
                  MExitMeth                                       '
            END SELECT                                            '
         NEXT j                                                   '
         IF ISTRUE MarkU AND ISTRUE Uniq THEN                     ' Can't do both MarkU and Uniq
            METHOD = "8Cannot use MARKU and UNIQ together"        ' Pass back error
            MExitMeth                                             '
         END IF
         NumKey = i                                               '

         IF IsCfX AND IsCfNX THEN _                               '
            METHOD = "8Both X and NX not allowed on the SORT command.": MExitMeth
         IF IsCfMX AND IsCfDX THEN _                              '
            METHOD = "8Both MX and DX not allowed on the SORT command.": MExitMeth

         FOR i = 1 TO NumKey                                      ' Validate entries

            IF SCmd(i).AorD = " " THEN SCmd(i).AorD = "A"         ' Set default direction
            IF SCmd(i).sFrom = 0 AND SCmd(i).sTo = 0 THEN         ' JUST the direction?
               sCmd(i).sFrom = 1: sCmd(i).sTo = MaxLength         ' then use defaults
            END IF                                                '
            IF SCmd(i).sFrom <> 0 AND SCmd(i).sTo = 0 THEN SCmd(i).sTo = MaxLength
            IF SCmd(i).sFrom = 0 OR SCmd(i).sTo = 0 THEN          ' Better be no zeros
               METHOD =  "8Illogical SORT operands"               ' Pass back error msg
               MExitMeth                                          '
            END IF                                                '
            IF (SCmd(i).sFrom > SCmd(i).sTo AND SCmd(i).sTo <> 0) THEN ' From GT To ?
               METHOD = "8Illogical SORT operands"                ' Pass back error
               MExitMeth                                          '
            END IF                                                '
         NEXT i

         REDIM sKey(LastLine) AS STRING                           ' Get the SORT arrays
         REDIM sKey2(LastLine) AS STRING                          '
         REDIM STxt(LastLine) AS STRING                           '
         REDIM CTxt(LastLine) AS WSTRING                          '
         REDIM SLins(LastLine) AS LONG                            '
         REDIM DLins(LastLine) AS LONG                            '

         i = me.Search(0, %CursBad)                               ' Do the initial search
         IF i THEN GOSUB DoTblAdd                                 ' If found, add to tables
         DO WHILE i                                               ' Do till not found
            i = me.Search(1, %CursBad)                            ' Do re-Finds
            IF i THEN GOSUB DoTblAdd                              ' If found, add to tables
         LOOP                                                     '

         IF (Prf.SrceXlate OR Prf.ColateXlate) AND ISFALSE uniq THEN  ' If Non-ANSI, or UNIQ don't do check
            FOR i = 0 TO y - 2                                    ' See if already in order
               IF LEFT$(sKey(i), basekey) > LEFT$(sKey(i + 1), basekey) THEN DidSort = %True: EXIT FOR
            NEXT i                                                '
         ELSE                                                     '
            DidSort = %True                                       ' Say to do SORT
         END IF                                                   '

         IF DidSort THEN                                          ' Out of order, do the sort
            ARRAY SORT sKey() FOR y, COLLATE cstring, TAGARRAY STxt() ' Do the sort
            ARRAY SORT sKey2() FOR y, COLLATE cstring, TAGARRAY CTxt() ' Do the sort of the Clr records

            lkey = ""                                             ' Reset last key
            FOR i = 0 TO y - 1                                    ' Put the lines back now
               L(SLins(i)).@LTxt = STxt(i)                        '
               L(SLins(i)).@LAttr = CTxt(i)                       ' Put back Attr
               me.ModSet(SLins(i))                                ' Remember we changed something
               LPCmdS(SLins(i)) = PCmdIX                          ' Mark all lines as :ZF
               IF MarkU THEN                                      ' If MarkU mode
                  me.LFlagBitOn(SLins(i), %User)                  ' Start by setting USER on all lines
                  IF LEFT$(sKey(i), basekey) <> lkey THEN         ' Different key as last one?
                     lkey = LEFT$(sKey(i), basekey)               ' Yes? save the new key
                     INCR delcount                                ' Count as a uniq item
                  ELSE                                            ' Same key
                     DECR delcount                                ' Previous line is no longer unique
                     me.LFlagBitOff(SLins(i - 1), %User)          ' Remove User from it
                     me.LFlagBitOff(SLins(i), %User)              ' And also from this line
                  END IF

               ELSEIF uniq THEN                                   ' Do some extra stuff for UNIQ
                  IF LEFT$(sKey(i), basekey) <> lkey THEN         ' Different key as last one?
                     lkey = LEFT$(sKey(i), basekey)               ' No? save the new key
                  ELSE                                            ' Same key!
                     dLins(delcount) = sLins(i)                   ' Mark line to be deleted
                     INCR delcount                                ' Count it
                  END IF                                          '
               END IF                                             '
            NEXT i                                                '
         END IF                                                   '

         IF ISTRUE uniq AND delcount > 0 THEN                     ' Do we have lines to delete?
            ARRAY SORT dLins() FOR delcount, DESCEND              ' Sort in descending order
            FOR i = 0 TO delcount - 1                             ' Delete the lines now
               me.lTxtFree(dLins(i))                              ' Go free the dynamic string
               me.LEntDel(dLins(i))                               ' Remove from the L() array
               DECR LastLine: DECR LastReal                       ' Adjust Last Line
            NEXT i                                                '
         END IF                                                   '
         REDIM sKey()                                             ' Free the Temp arrays
         REDIM sKey2()                                            '
         REDIM sTxt()                                             '
         REDIM cTxt()                                             '
         REDIM sLins()                                            '
         REDIM dLins()                                            '
         IF uniq THEN                                             ' UNIQ request?
            IF delcount > 0 THEN                                  '
               METHOD = "0SORT Complete, " + FORMAT$(delcount) + " non-Uniq lines deleted"   ' Say all is well
            ELSE                                                  '
               METHOD = "0SORT Complete, all lines are unique"    ' Say all is well
            END IF                                                '
         ELSEIF MarkU THEN                                        ' If MarkU mode
            IF delcount > 0 THEN                                  '
               METHOD = "0SORT Complete, " + FORMAT$(delcount) + " Uniq lines marked as USER lines"
            ELSE                                                  '
               METHOD = "0SORT Complete, all lines are unique and marked as USER lines"
            END IF                                                '
         ELSE                                                     '
            METHOD = "0SORT Complete" + IIF$(DidSort, "", ", already sorted")   ' Say all is well
         END IF                                                   '
         MExitMeth

         DoTblAdd:                                                ' Add a sort item
            x = cfFLine                                           ' Get the line number
            sKey(y) = ""                                          ' Init key to null
            FOR k = 1 TO NumKey                                   ' Build the sort key

               j = SCmd(k).sTo - SCmd(k).sFrom + 1                ' Length of this Key
               t = LSET$(MID$(L(x).@LTxt, SCmd(k).sFrom, j), j)   '
               IF SCmd(k).CASE = "T" THEN t = UUCASE(t)           ' If CASE=T, uppercase it
               IF SCmd(k).AorD = "D" THEN t = sInvert(t)          ' Invert key if descending sort
               sKey(y) += t                                       ' Add to key
            NEXT k                                                '
            basekey = LEN(sKey(y))                                ' Save length of base key for UNIQ test
            sKey(y) = skey(y) + FORMAT$(x, "00000000")            ' Add record number for a tie breaker
            STxt(y) = L(x).@LTxt: SLins(y) = x                    ' Save Text and original source line #
            sKey2(y) = sKey(y)                                    ' Setup second dup key
            CTxt(y) = L(x).@LAttr                                 ' Save it
            INCR y                                                '
            RETURN

         Do2ndOper:
            IF LEN(pCmdOps(j)) = 2 THEN                           ' A 2nd character?
               IF INSTR("CT", RIGHT$(pCmdOps(j), 1)) = 0 THEN     ' Valid
                  METHOD = "8Invalid SORT operands"               ' Pass back error
                  MExitMeth                                       '
               ELSE                                               '
                  SCmd(i).CASE = RIGHT$(pCmdOps(j), 1)            ' Else save it
               END IF                                             '
            ELSEIF LEN(pCmdOps(j)) = 1 THEN                       ' Ignore these
               '
            ELSE                                                  '
               METHOD = "8Invalid SORT operands"                  ' Pass back error
               MExitMeth                                          '
            END IF                                                '
            RETURN                                                '
      END METHOD

      METHOD LInsertEmpty(ln AS LONG, n AS LONG, IFlag AS LONG)
      '---------- Add empty line(s) following passed line pointer
      LOCAL x, p, ol, nl, ot, nt,i, fMix, slno, inc, rerr AS LONG
      LOCAL a, b, c AS LONG, w1 AS WSTRING
         MEntry
         UBoundT = UBOUND(T): UBoundL = UBOUND(l)                 ' Save current sizes
         IF IsSlecSet THEN                                        ' Need we check Slec line range?
            IF (ln >= me.LineNoRef(FORMAT$(SlecSLin)) AND ln <= me.LineNoRef(FORMAT$(SlecELin))) THEN
               OffSlecSet                                         ' Clear Slec status
               OffSlecActive                                      '
               sDoStatusBar($SBSelect)                            ' re-Do the StatusBar Select box
            END IF                                                '
         END IF                                                   '
         OnRenumFlag                                              ' Say Renum needed
         slno = ln                                                ' Setup param for GetIncrement
         me.GetIncrement(slno, n, inc)                            ' Pass lineno/#lines, get back StartNo, Incr
         slno += inc                                              ' Create 1st inserted line number
         IF IsMEdit THEN fMIX = me.MEditTbl("I", FORMAT$(ln))     ' In MEdit mode? Yes, get the MIX value of previous line

         me.LTblSize(n + 5)                                       ' Ensure enough room in text table

         '----- Expand L() if running out of room
         IF n + 5  > LCtr  THEN                                   ' If running out of room in L() make it bigger
            ol = UBOUND(L()): nl = ol + n + (ol * 2)              ' Save old array size, calc new size
            REDIM PRESERVE L(nl) AS INSTANCE DataLine             ' Get room for another bunch of lines in L() array
            rerr = ERR
            UBoundL = UBOUND(L())                                 ' Save new UBound value
            LCtr += (nl - ol)                                     ' Bump LCtr
            FOR x = 1 TO ol                                       ' Correct the @LTxt pointers
               IF L(x).LTxt <> 0 THEN                             '
                  L(x).LTxt = VARPTR(T(L(x).LTxtIX))              ' Correct using the saved TxtIX value
               END IF                                             '
               IF L(x).LAttr <> 0 THEN                            '
                  L(x).LAttr = VARPTR(TW(L(x).LAttrIX))           ' Correct using the saved AttrIX value
               END IF                                             '
            NEXT x                                                '
            FOR x = ol + 1 TO UBOUND(L())                         ' Reset the new entries
               L(x).LTxt = 0                                      '
               L(x).LTxtIX = 0                                    '
               L(x).LAttr = 0                                     '
               L(x).LAttrIX = 0                                   '
               L(x).LFlag = 0                                     '
               L(x).LCol = 0                                      '
               L(x).LWrk2 = 0                                     '
               L(x).LLCtl = $BlankLNo                             '
               L(x).LLNum = $BlankLNo                             '
               L(x).LLbl = $BlankLNo                              '
               L(x).LTag = $BlankLNo                              '
            NEXT x                                                '
         END IF                                                   '
w1 = LAttrG(2)

         p = ln                                                   ' Copy passed parameter

         '----- Do fast alternative to ARRAY INSERT 1 line at a time

         IF n > 1 THEN                                            ' If n > 1 then shift it ourselves
            MEMORY COPY VARPTR(L(p + 1)), VARPTR(L(p + n + 1)), SIZEOF(DataLine) * (UBOUND(L) - p - n)
         ELSE                                                     ' Else just let ARRAY INSERT do it
            ARRAY INSERT L(ln + 1) FOR (LastLine - p + 2)         ' Insert an entry in L() array
         END IF                                                   '

         '----- Fill in the newly inserted lines
         FOR x = 1 TO n                                           ' Loop for the number to be inserted
            INCR p                                                ' Incr local line pointer value
            INCR LastLine                                         ' Incr total lines in array
            IF (Iflag AND %Data) THEN INCR LastReal               ' Incr LastReal if a data line
            DECR LCtr                                             ' Reduce L() count remaining

            i = INSTR(TIDX, "1")                                  ' Scan for next available entry
            MID$(TIDX, i, 1) = "0"                                ' Mark it no longer available
            DECR TTCtr                                            ' Reduce T() count remaining
            T(i) = ""                                             ' Null the string
            L(p).LTxtIX = i                                       ' Remember the index in the L() entry
            L(p).LTxt = VARPTR(T(i))                              ' Point L.LTxt to the dynamic String

            TW(i) = CHR$$(0)                                      ' Null the string
            L(p).LAttrIX = i                                      ' Remember the index in the L() entry
            L(p).LAttr = VARPTR(TW(i))                            ' Point L.LAttr to the dynamic String

            L(p).LFlag = IFlag                                    ' Setup flag as requested
            IF (IFlag AND %SpecialLine) = 0 THEN _                ' If not a special line
               L(p).LLNum = FORMAT$(slno, "00000000")             ' then assign it a line number
            L(p).LLCtl = FORMAT$(slno, "00000000")                '
            IF ISTRUE (L(p).LFlag AND %InsertLine) THEN L(p).LLCtl = "''''''''"
            IF ISTRUE (L(p).LFlag AND %Note)       THEN L(p).LLCtl = gLnoTextTxt(14) ' =  NOTE>
            slno += inc                                           ' Bump inserted line number
            L(p).LLbl = $BlankLNo                                 '
            L(p).LTag = $BlankLNo                                 '
            L(p).LWrk1 = 0                                        '
            L(p).LWrk2 = 0                                        '
            IF IsMEdit THEN L(p).LMix = fMIX                      ' In MEdit mode? Yes, Update the MIX
         NEXT x                                                   ' Loop for number to be inserted
         MExit
      END METHOD

      METHOD LineNoRef(operand AS STRING) AS LONG
      '---------- Convert operand to line number
      LOCAL oper AS STRING, Wrd8 AS STRING * 8, i AS LONG
         MEntry
         IF LEN(operand) > 8 THEN METHOD = -1: MExitMeth          ' If too big, kill it
         oper = UUCASE(operand)                                   ' Uppercase it
         IF oper = ".0" THEN                                      ' Special .0 case?
            METHOD = 1: MExitMeth                                 '
         END IF                                                   '
         IF LEFT$(oper, 1) = "." THEN                             ' Is this a label?
            IF LastLine = 2 THEN                                  ' Empty data area?
               METHOD = -1                                        ' Everything else is invalid
            ELSE                                                  '
               IF oper = ".ZFIRST" OR oper = ".ZF" THEN           ' Look for built in labels
                  METHOD = 2                                      '
               ELSEIF oper = ".ZLAST" OR oper = ".ZL" THEN        '
                  METHOD = LastLine -1                            '
               ELSEIF oper = ".ZFIND" THEN                        '
                  METHOD = IIF(cfFLine, cfFLine, -1)              ' If we have one, pass it back
               ELSEIF oper = ".ZLOC" THEN                         '
                  METHOD = IIF(LocLineR, LocLineR, -1)            ' If we have one, pass it back
               ELSEIF oper = ".ZCSR" THEN                         '
                  METHOD = S(CRow)                                '
               ELSEIF VAL(MID$(oper, 2)) > 0 THEN                 ' Numeric dotted label?
                  i = VAL(MID$(oper, 2))                          ' Get the value
                  i = sFindLineNum(i)                             ' Convert to proper IX value
                  IF i < 0 THEN i = ABS(i)                        ' If -ve it was a 999999 type operand
                  METHOD = i                                      ' Pass it back
               ELSE                                               '
                  Wrd8 = oper                                     ' Convert to fixed string for search
                  i = me.LLCtlScan(Wrd8)                          ' See if we can find it
                  METHOD = IIF(i = 0, -1, i - 1)                  ' Pass zero if NFG, else LLNum of the line
               END IF                                             '
            END IF                                                '

         ELSEIF LEFT$(oper, 1) = "!" THEN                         ' Is this an LPtr type?
            IF VAL(MID$(oper,2)) = 0 OR _                         ' !0 is not allowed
               VAL(MID$(oper, 2)) > Lastline THEN                 ' Or > Lastline
               METHOD = -1                                        ' Kill it
            ELSE                                                  '
               METHOD = VAL(MID$(oper, 2))                        ' Pass it back
            END IF                                                '

         ELSE                                                     '
            i = VAL(oper)                                         ' See if numeric
            IF i > 0 THEN i = sFindLineNum(i)                     ' Convert to proper IX value
            IF i < 0 THEN i = ABS(i)                              ' If -ve it was a 999999 type operand
            METHOD = i                                            ' Pass it back
         END IF
         MExit
      END METHOD

      METHOD LInitTxtData(fn AS STRING)
      '---------- Initialize the text work areas
      LOCAL ct, t AS STRING, i, tsize AS LONG, fsize AS QUAD, FD AS DIRDATA
         MEntry
         IF TRIM$(fn) = "" THEN                                   ' Null filename, a reset to empty
            tsize = 50000                                         ' Use 50000
         ELSE                                                     '
            t = DIR$(fn, 6, TO FD)                                ' Search for filename  (6 = Normal, hidden, system)
            IF TRIM$(t) = "" THEN                                 ' Can't find it?
               tsize = 100000                                     ' Use 100000
            ELSE                                                  '
               fsize = MAK(QUAD, FD.FileSizeLow , FD.FileSizeHigh)' Get the filesize
               SELECT CASE fsize                                  '
                  CASE < 500000:   tsize = 100000                 ' Set a 'best guess'
                  CASE < 4000000:  tsize = 250000                 '
                  CASE < 10000000: tsize = 500000                 '
                  CASE ELSE: tsize = 500000                       '
               END SELECT                                         '
            END IF                                                '
         END IF                                                   '

         REDIM T(tsize) AS INSTANCE STRING                        ' To hold the dynamic TEXT strings
         REDIM TW(tsize) AS INSTANCE WSTRING                      ' To hold the dynamic Attr strings
         REDIM L(tsize) AS INSTANCE Dataline                      ' TEXT control array
         UBoundL = tsize: UBoundT = tsize                         '
         tTCtr = tsize                                            ' Set number available in table
         LCtr = tsize                                             '

         FOR i = 0 TO UBOUND(T)                                   ' Reset the Text array
            T(i) = "": TW(i) = ""                                 ' and the Attr array
         NEXT i                                                   '
         FOR i = 0 TO UBOUND(L)                                   ' Reset the DataLine array
            L(i).LTxt = 0                                         '
            L(i).LTxtIX = 0                                       '
            L(i).LAttr = 0                                        '
            L(i).LAttrIX = 0                                      '
            L(i).LFlag = 0                                        '
            L(i).LLCtl = $BlankLNo                                '
            L(i).LLNum = $BlankLNo                                '
            L(i).LLbl = $BlankLNo                                 '
            L(i).LTag = $BlankLNo                                 '
         NEXT i                                                   '
         TIDX = STRING$(UBOUND(L), "1")                           ' Make all T() and TW() available
         REDIM MEditList(1 TO 20) AS INSTANCE STRING              ' MEDIT filename list
         REDIM MEditFlag(1 TO 20) AS INSTANCE LONG                ' MEDIT filename flags
         MEditCount = 0                                           '
         me.LTxtAlloc(1)                                          ' Initialize Top of Dataset Line
         L(1).LFlag = L(1).LFlag OR %Top                          '
         L(1).LLNum = "00000000"                                  '
         L(1).LLCtl = "********"                                  '
         me.LTxtAlloc(2)                                          ' Initialize Bottom of Dataset Line
         L(2).LFlag = L(2).LFlag OR %Bottom                       '
         L(2).LLNum = "00000000"                                  '
         L(2).LLCtl = "********"                                  '
         RESET s()                                                '
         me.CsrRow = 1: Offset = 0: COffset = 0                   ' Set Startup values
         me.CsrCol = 11: MaxLength = 0: sCurPrio = 0              '
         MarkedLine = 0: Flag = 0: LastLine = 2: TopScrn = 1      '
         OffBackupDone: Encoding = "ANSI"                         '
         LastReal = 0: TMode = (TMode AND %MFMTab)                ' Kill all but FMTab flag
         MExit
      END METHOD

      METHOD LocSearch() AS LONG
      '---------- Do the full LOCATE Search procedure
      LOCAL i, j, k, x, y, z, CurrWinner, WinnerLine, RevOnce, AllCount AS LONG
      LOCAL PosType AS INTEGER, t1, t2 AS STRING, aTxt AS WSTRING POINTER, AttrAsc AS WORD
         MEntry
         METHOD = %False                                          ' Start off saying we've failed
         PosType = IIF(IsLfTop, %FLocate, %Locate)                ' Set position type based on TOP
         IF IsRevOnce THEN                                        ' Reverse one time only?
            OffRevOnce: RevOnce = %True                           ' Remove globally, remember locally
         END IF                                                   '
         LocDir = IIF(IsLfPrev, -1, 1)                            ' If PREV then reverse search direction
         IF IsCData OR IsCLinn THEN                               ' In the Data Area or Line number area
            LocLine = sGetIX(CRow) + IIF(RevOnce, -1 * LocDir, LocDir) ' Swap in current cursor location
         END IF                                                   '
         IF IsLfFirst OR IsLfAll THEN                             ' If FIRST/ALL then start as the top
            LocLine = 1: LfClear(%LocFirst): LfSet(%LocNext)      ' Remove so we do it just once
         END IF                                                   '
         IF IsLfLast THEN                                         ' If LAST then start as the bottom
            LocLine = LastLine: LocDir = -1                       '
            LfClear(%LocLast): LfSet(%LocPrev)                    ' Remove LAST and insert PREV
         END IF                                                   '

         IF LocNotFound  <> 0 THEN                                ' Was last search a not found?
            IF LocNotFound = 1 THEN                               ' Hit the top?
               LocLine = LastLine                                 ' Backward, restart at the bottom
            ELSE                                                  ' Else the bottom
               LocLine = 1                                        ' Forward, restart at the top
            END IF                                                '
            LocNotFound = 0                                       ' Clear, we wrap just once
         END IF                                                   '

         i = LocLine: LocNFDir = IIF(RevOnce, -1 * LocDir, LocDir)' Init for DO loop
         DO WHILE 1 = 1                                           ' Do forever
            IF (IIF(RevOnce, -1 * LocDir, LocDir) = 1  AND IsLBottom(i)) OR _  ' Do till top or bottom
               (IIF(RevOnce, -1 * LocDir, LocDir) = -1 AND IsLTop(i)) THEN EXIT DO  '

            IF IsLBottom(i) OR IsLTop(i) THEN i += IIF(RevOnce, -1 * LocDir, LocDir): ITERATE DO  '

            IF LocLabLine1 > 0 AND LocLabLine2 > 0 THEN           ' Line range request
               IF i < LocLabLine1 OR i > LocLabLine2 THEN         ' See if in range
                  i += IIF(RevOnce, -1 * LocDir, LocDir)          ' Next entry
                  ITERATE DO                                      '
               END IF                                             '
            END IF                                                '

            '----- Mark as :ZNF line to say we looked at it
            LPCmdS(i) = (0 - PCmdIX)                              ' Stuff in negated PCmdIX to make it :ZNF

            IF ISNOTNULL(LocTag) THEN                             ' Looking for a specific Tag?
               IF LTagG(i) = LocTag THEN                          ' Have a Match?
                  IF ISFALSE IsLfNot THEN j = %True               ' If it isn't NOT, set True
               ELSE                                               ' Not a Match
                  IF IsLfNot THEN j = %True                       ' If it is NOT, set True
               END IF                                             '

            ELSEIF IsLfChange THEN                                ' Looking for a ==CHG> line
               IF ISFALSE IsLfCurrent THEN                        ' Looking for the current one?
                  IF IsLEQChange(i) THEN                          ' No, normal CHANGE, Is this  one
                     IF ISFALSE IsLfNot THEN j = %True            ' If it isn't NOT, set True
                  ELSE                                            ' Not a Match
                     IF IsLfNot THEN j = %True                    ' If it is NOT, set True
                  END IF                                          '
               ELSE                                               '
                  IF LFIXG(i) > CurrWinner THEN CurrWinner = LFIXG(i): WinnerLine = i ' If more current, save the winner
               END IF                                             '

            ELSEIF IsLfFind THEN                                  ' Looking for a FIND line
               IF ISFALSE IsLfCurrent THEN                        ' Looking for the current one?
                  IF LFIXG(i) > 0 THEN                            ' No, is this one
                     IF ISFALSE IsLfNot THEN j = %True            ' If it isn't NOT, set True
                  ELSE                                            ' Not a Match
                     IF IsLfNot THEN j = %True                    ' If it is NOT, set True
                  END IF                                          '
               ELSE                                               '
                  IF LFIXG(i) > CurrWinner THEN CurrWinner = LFIXG(i): WinnerLine = i ' If more current, save the winner
               END IF                                             '

            ELSEIF IsLfCommand THEN                               ' Looking for a pending command line
               IF me.LLCtlGet(i) <> LLNumG(i) AND _               ' If not a line number
                  me.LLCtlGet(i) <> "''''''''" AND _              ' and not an Insert line
                  me.LLCtlGet(i) <> "--------" AND _              ' and not a -------- line
                  LEFT$(me.LLCtlGet(i), 1) <> "." AND _           ' and not a label
                  LEFT$(me.LLCtlGet(i), 1) <> ":" AND _           ' and not a tag
                  LEFT$(me.LLCtlGet(i), 1) <> "=" THEN            ' and not a special, then its a command
                  IF ISFALSE IsLfNot THEN j = %True               ' If it isn't NOT, set True
               ELSE                                               ' Not a Match
                  IF IsLfNot THEN j = %True                       ' If it is NOT, set True
               END IF                                             '

            ELSEIF IsLfError THEN                                 ' Looking for an Error
               '                                                  ' Not yet

            ELSEIF IsLfExcluded THEN                              ' Looking for an Excluded line?
               IF ISFALSE IsLfNot THEN                            ' If it isn't NOT, look for Xclude lines
                  IF IsLXclude(i) THEN j = %True                  ' This is one
               ELSE                                               ' We're doing the NOT X search
                  IF IsLData(i) AND ISFALSE IsLInvisible(i) THEN j = %True ' A visible Data line
               END IF                                             '

            ELSEIF IsLfNX THEN                                    ' Looking for an NX line?
               IF ISFALSE IsLfNot THEN                            ' If it isn't NOT, look for Normal lines
                  IF IsLData(i) AND ISFALSE IsLInvisible(i) THEN j = %True ' This is one
               ELSE                                               ' We're doing the NOT X search
                  IF IsLXclude(i) THEN j = %True                  ' An Exclude line line
               END IF                                             '

            ELSEIF IsLfU THEN                                     ' Looking for a User line?
               IF ISFALSE IsLfNot THEN                            ' If it isn't NOT, look for User lines
                  IF IsLUser(i) THEN j = %True                    ' This is one
               ELSE                                               ' We're doing the NOT U search
                  IF ISFALSE IsLUser(i) THEN j = %True            ' A visible Data line
               END IF                                             '

            ELSEIF IsLfNU THEN                                    ' Looking for a non-User line?
               IF ISFALSE IsLfNot THEN                            ' If it isn't NOT, look for non-User lines
                  IF ISFALSE IsLUser(i) THEN j = %True            ' This is one
               ELSE                                               ' We're doing the NOT NU search
                  IF IsLUser(i) THEN j = %True                    ' A User line
               END IF                                             '

            ELSEIF IsLfLabel THEN                                 ' Looking for an Excluded line?
               IF L(i).LLbl <> $BlankLNo THEN                     ' This is one
                  IF ISFALSE IsLfNot THEN j = %True               ' If it isn't NOT, set True
               ELSE                                               ' Not a Match
                  IF IsLfNot THEN j = %True                       ' If it is NOT, set True
               END IF

            ELSEIF IsLfNote THEN                                  ' Looking for a Note line?
               IF IsLNote(i) THEN                                 ' This is one
                  IF LocNVar = 0 THEN                             ' An old NOTE type?
                     IF L(i).LWrk2 = 0 THEN                       ' And this is an old Note
                        IF ISFALSE IsLfNot THEN j = %True         ' If it isn't NOT, set True
                     ELSE                                         '
                        IF IsLfNot THEN j = %True                 ' If it is NOT, set True
                     END IF                                       '
                  ELSEIF LocNVar < 26 THEN                        ' A specific xNOTE type?
                     IF L(i).LWrk2 = LocNVar THEN                 ' And this is that NOTE type
                        IF ISFALSE IsLfNot THEN j = %True         ' If it isn't NOT, set True
                     ELSE                                         '
                        IF IsLfNot THEN j = %True                 ' If it is NOT, set True
                     END IF                                       '
                  ELSE                                            ' Must be ZNOTE
                     IF L(i).LWrk2 <> 0 THEN                      ' And this is a special NOTE type
                        IF ISFALSE IsLfNot THEN j = %True         ' If it isn't NOT, set True
                     ELSE                                         '
                        IF IsLfNot THEN j = %True                 ' If it is NOT, set True
                     END IF                                       '
                  END IF                                          '
               ELSE                                               ' Not a Match
                  IF IsLfNot THEN j = %True                       ' If it is NOT, set True
               END IF                                             '

            ELSEIF IsLfSpecial THEN                               ' Looking for a Special line?
               IF IsLWord(i) OR _                                 ' One of them?
                  IsLTabs(i) OR _                                 '
                  IsLBounds(i) OR _                               '
                  IsLCols(i) OR _                                 '
                  IsLMark(i) OR _                                 '
                  IsLMask(i) OR _                                 '
                  IsLProf(i) THEN                                 ' Yes, set true
                  IF ISFALSE IsLfNot THEN j = %True               ' If it isn't NOT, set True
               ELSE                                               ' Not a Match
                  IF IsLfNot THEN j = %True                       ' If it is NOT, set True
               END IF                                             '

            ELSEIF IsLfFile THEN                                  ' Looking for a File line?
               IF IsLFile(i) THEN                                 ' One of them?
                  IF ISFALSE IsLfNot THEN j = %True               ' If it isn't NOT, set True
               ELSE                                               ' Not a Match
                  IF IsLfNot THEN j = %True                       ' If it is NOT, set True
               END IF                                             '

            ELSEIF IsLfPage THEN                                  ' Looking for a Page line?
               IF IsLPage(i) THEN                                 ' One of them?
                  IF LocLength > 0 THEN                           ' Specific page number?
                     IF LocLength - Prf.PageOffset = 1 AND i = 2 THEN' Page 1 fudge?
                        IF ISFALSE IsLfNot THEN j = %True         ' If it isn't NOT, set True
                     ELSEIF LocLength - Prf.PageOffset = L(i).LWrk2 THEN ' Correct page?
                        IF ISFALSE IsLfNot THEN j = %True         ' If it isn't NOT, set True
                     ELSE                                         ' Not right page
                        IF IsLfNot THEN j = %True                 ' If it is NOT a page, set True
                     END IF                                       '
                  ELSE                                            ' Is PAGE, and is any old page
                     IF ISFALSE IsLfNot THEN j = %True            ' If it isn't NOT, set True
                  END IF                                          '
               ELSE                                               ' Not a Match
                  IF IsLfNot THEN j = %True                       ' If it is NOT a page, set True
               END IF                                             '

            ELSEIF IsLfKeep THEN                                  ' Looking for a Kept command?
               IF INSTR(me.LLCtlGet(i), "&") <> 0 THEN            ' This is one
                  IF ISFALSE IsLfNot THEN j = %True               ' If it isn't NOT, set True
               ELSE                                               ' Not a Match
                  IF IsLfNot THEN j = %True                       ' If it is NOT, set True
               END IF                                             '

            ELSEIF IsLfTag THEN                                   ' Looking for any Tag?
               IF LTagG(i) <> $BlankLNo THEN                      ' This is one
                  IF ISFALSE IsLfNot THEN j = %True               ' If it isn't NOT, set True
               ELSE                                               ' Not a Match
                  IF IsLfNot THEN j = %True                       ' If it is NOT, set True
               END IF                                             '

            ELSEIF IsLfLong THEN                                  ' Looking for any Long?
               IF IsLData(i) THEN                                 ' Only data lines
                  IF LEN(L(i).@LTxt) > LocLength THEN             ' This is one
                     IF ISFALSE IsLfNot THEN j = %True            ' If it isn't NOT, set True
                  ELSE                                            ' Not a Match
                     IF IsLfNot THEN j = %True                    ' If it is NOT, set True
                  END IF                                          '
               END IF                                             '

            ELSEIF IsLfSize THEN                                  ' Looking for Size?
               IF IsLData(i) THEN                                 ' Only data lines
                  IF LEN(L(i).@LTxt) = LocLength THEN             ' This is one
                     IF ISFALSE IsLfNot THEN j = %True            ' If it isn't NOT, set True
                  ELSE                                            ' Not a Match
                     IF IsLfNot THEN j = %True                    ' If it is NOT, set True
                  END IF                                          '
               END IF                                             '

            ELSEIF IsLfStd THEN                                   ' Looking for Std color
               IF L(i).LAttr <> 0 THEN                            ' Line have a Attr line?
                  aTxt = LAttrGP(i)                               ' Get pointer to Attr text string
                  k = %False                                      ' Say no HiLites
                  FOR x = 1 TO LEN(@aTxt)                         ' Scan through it
                     AttrAsc = ASC(MID$(@aTxt, x)) AND %AttrHiLite' Get the HiLite bits
                     IF AttrAsc <> 0 THEN k = %True: EXIT FOR     ' Found one
                  NEXT x                                          '

                  IF ISFALSE k THEN                               ' But nothing there?
                     IF ISFALSE IsLfNot THEN j = %True            ' If it isn't NOT, set True
                  ELSE                                            ' Not a Match
                     IF IsLfNot THEN j = %True                    ' If it is NOT, set True
                  END IF                                          '
               ELSE                                               ' No Attr values
                  IF ISFALSE IsLfNot THEN j = %True               ' If it isn't NOT, set True
               END IF                                             '

            ELSEIF IsLfColor THEN                                 ' Looking for a specific color
               IF L(i).LAttr <> 0 THEN                            ' Line have a Attr line?
                  aTxt = LAttrGP(i)                               ' Get pointer to Attr text string
                  k = %False                                      ' Say not found
                  FOR x = 1 TO LEN(@aTxt)                         ' Scan through it
                     AttrAsc = ASC(MID$(@aTxt, x)) AND %AttrHiLite' Get the HiLite bits
                     SHIFT RIGHT AttrAsc, 8                       ' Right align
                     IF AttrAsc = LocColor THEN k = %True: EXIT FOR ' Found one
                  NEXT x                                          '

                  IF k THEN                                       ' Found?
                     IF ISFALSE IsLfNot THEN j = %True            ' If it isn't NOT, set True
                  ELSE                                            ' Not a Match
                     IF IsLfNot THEN j = %True                    ' If it is NOT, set True
                  END IF                                          '
               ELSE                                               ' No Attr values
                  IF IsLfNot THEN j = %True                       ' If it isn't NOT, set True
               END IF                                             '

            END IF                                                '

            IF j THEN                                             ' If we have a 'found' condition
               METHOD = %True                                     '
               LPCmdS(i) = PCmdIX                                 ' Stuff in PCmdIX to make it :ZF
               INCR AllCount                                      ' Count winners
               IF ISFALSE IsLInvisible(i) THEN                    ' A normal non-X'd line?
                  me.CurSetReq(PosType, i, 0, %True)              ' Set cursor set attempt
                  IF IsLfMX THEN                                  ' MX requested
                     me.LFlagBitOn(i, %Invisible)                 ' Then make it invisible
                     gfXRebuild = %True                           ' Rebuild exclude state
                     OnUndoFlag                                   ' Take a snapshot
                  END IF                                          '
               ELSE                                               ' We have a line already X'd out
                  IF ISFALSE IsLfMX THEN                          ' No MX requested
                     me.LFlagBitOff(i, %Invisible)                '
                     gfXRebuild = %True                           ' Rebuild exclude state
                     OnUndoFlag                                   ' Take a snapshot
                     me.CurSetReq(PosType, i, 0, %True)           ' Set cursor set attempt
                  END IF                                          '
               END IF                                             '
               LocLineR = i                                       ' Save found line
               LocLine = i + IIF(RevOnce, -1 * LocDir, LocDir)    ' Save continue line
               IF ISFALSE IsLfAll THEN                            '
                  EXIT DO                                         '
               ELSE                                               '
                  j = 0                                           '
               END IF                                             '
            END IF                                                '
            i += IIF(RevOnce, -1 * LocDir, LocDir)                '
         LOOP                                                     '

         IF IsLfAll THEN                                          ' Issue message for ALL
            scError(0, "LOCATE matched " + FORMAT$(AllCount) + IIF$(AllCount = 1, " line", " lines"))
         END IF                                                   '
         IF ISFALSE IsLfCurrent THEN MExitMeth                    ' If not CURRENT
         IF ISFALSE IsLfChange AND ISFALSE IsLfFind THEN MExitMeth'
         IF CurrWinner = 0 THEN MExitMeth                         ' If no winner
         METHOD = %True                                           '
         IF ISFALSE IsLInvisible(WinnerLine) THEN                 ' A normal non-X'd line?
            me.CurSetReq(PosType, WinnerLine, 0, %True)           ' Set cursor set attempt
         ELSE                                                     ' We have a line already X'd out
            me.LFlagBitOff(WinnerLine, %Invisible)                '
            gfXRebuild = %True                                    ' Rebuild exclude state
            me.CurSetReq(PosType, WinnerLine, 0, %True)           ' Set cursor set attempt
         END IF                                                   '
         LocLineR = WinnerLine                                    ' Save winner line
         LocLine = WinnerLine + IIF(RevOnce, -1 * LocDir, LocDir) ' Save continue line
         MExit                                                    '
      END METHOD

      METHOD LPopOutX(ln AS LONG) AS LONG
      '---------- Pop out an X'd line
      REGISTER i AS LONG
      REGISTER j AS LONG
      LOCAL k AS LONG, tmp AS DataLine
         MEntry
         i = ln - L(ln).LWrk1                                     ' Point i back at the XX control line
         j = i + l(i).LWrk1                                       ' Point j at last XX'd out line in the range

         IF L(i).LWrk1 = 1 THEN                                   ' Doing a whole 1 line range
            me.LTxtFree(i)                                        ' Go free the dynamic string of the XX line
            ARRAY DELETE L(i)                                     ' Remove from the L() array
            DECR LastLine                                         ' Adjust for removed XX line
            L(i).LFlag = L(i).LFlag AND (&HFFFFFFFF - %Invisible) ' Remove Invisible (i now = real line #)
            L(i).LFlag = L(i).LFlag OR %Popped                    ' Mark as popped during this interaction
            L(i).LWrk1 = 0                                        ' Clear Work
            IF ln <= DotEnd THEN DotEnd = DotEnd - 1              ' Adjust the end range if needed
            METHOD = i                                            ' Pass Back the line number
            MExitMeth                                             ' We're done, a simple one
         END IF                                                   '

         IF j = ln THEN                                           ' Doing last line of a range
            L(ln).LFlag = L(ln).LFlag AND (&HFFFFFFFF - %Invisible) ' Remove Invisible, that's all that's needed
            L(ln).LFlag = L(ln).LFlag OR %Popped                  ' Mark as popped during this interaction
            L(ln).LWrk1 = 0                                       ' Clear Work in the Pop'ed out line
            DECR L(i).LWrk1                                       ' and reduce count in the XX marker line
            METHOD = ln                                           ' Pass Back the line number
            MExitMeth                                             ' We're done, another simple one
         END IF                                                   '

         IF L(ln).LWrk1 = 1 THEN                                  ' Doing the 1st line of a range
            DECR L(i).LWrk1                                       ' Adjust # in reduced range
            tmp = L(i)                                            ' Save the XX marker line for a bit
            ARRAY DELETE L(i)                                     ' Delete the XX marker line from its old location
            l(i).LFlag = L(i).LFlag AND (&HFFFFFFFF - %Invisible) ' Make the line visible (i now = real line #)
            L(i).LFlag = L(i).LFlag OR %Popped                    ' Mark as popped during this interaction
            L(i).LWrk1 = 0                                        ' Clear Work in the Pop'ed out line
            ARRAY INSERT L(i + 1), tmp                            ' Stuff saved XX marker back in after Popped line
            FOR j = 1 TO L(i + 1).LWrk1                           ' Redo Work numbers in the new range
               L(i + 1 + j).LWrk1 = j                             '
            NEXT j                                                '
            METHOD = i                                            ' Pass Back the line number
            MExitMeth                                             '
         END IF
                                                                  ' Doing a line in middle of a range
         k = L(i).LWrk1 - L(ln).LWrk1                             ' Calc count for the new 'lower' range
         L(i).LWrk1 = L(ln).LWrk1 - 1                             ' Set a new count in current marker
         L(ln).LFlag = L(ln).LFlag AND (&HFFFFFFFF - %Invisible)  ' Pop out the requested line
         L(ln).LFlag = (L(ln).LFlag OR %Popped)                   ' Mark as popped during this interaction
         me.LInsertEmpty(ln, 1, %Xclude)                          ' Insert a line to hold the Exclude marker
         L(ln + 1).LWrk1 = k                                      ' Stuff in the calculated # of X'd lines in new area
         L(ln + 1).LLNum = "00000000"                             ' Clear line number

         FOR j = 1 TO k                                           ' Redo Work numbers in the new range
            L(ln + 1 + j).LWrk1 = j                               '
         NEXT j                                                   '
         METHOD = ln                                              ' Pass Back the line number
         me.AdjustPending(ln, 1, 0)                               ' Adjust pending stuff
         IF ln <= DotEnd THEN DotEnd = DotEnd + 1                 ' Adjust the end range if needed
         MExit
      END METHOD

      METHOD LPopOutY(ln AS LONG) AS LONG
      '---------- Pop out an X'd line, return needed adjustment for subsequent lines
      REGISTER i AS LONG
      REGISTER j AS LONG
      LOCAL k AS LONG, tmp AS DataLine
         MEntry
         i = ln - L(ln).LWrk1                                     ' Point i back at the XX control line
         j = i + l(i).LWrk1                                       ' Point j at last XX'd out line in the range

         IF L(i).LWrk1 = 1 THEN                                   ' Doing a whole 1 line range
            me.LTxtFree(i)                                        ' Go free the dynamic string of the XX line
            ARRAY DELETE L(i)                                     ' Remove from the L() array
            DECR LastLine                                         ' Adjust for removed XX line
            L(i).LFlag = L(i).LFlag AND (&HFFFFFFFF - %Invisible) ' Remove Invisible (i now = real line #)
            L(i).LFlag = L(i).LFlag OR %Popped                    ' Mark as popped during this interaction
            L(i).LWrk1 = 0                                        ' Clear Work
            METHOD = -1                                           ' We ended up deleting one line
            MExitMeth                                             ' We're done, a simple one
         END IF                                                   '

         IF j = ln THEN                                           ' Doing last line of a range
            L(ln).LFlag = L(ln).LFlag AND (&HFFFFFFFF - %Invisible) ' Remove Invisible, that's all that's needed
            L(ln).LFlag = L(ln).LFlag OR %Popped                  ' Mark as popped during this interaction
            L(ln).LWrk1 = 0                                       ' Clear Work in the Pop'ed out line
            DECR L(i).LWrk1                                       ' and reduce count in the XX marker line
            METHOD = 0                                            ' Pass Back zero, we made no inserts/deletes
            MExitMeth                                             ' We're done, another simple one
         END IF                                                   '

         IF L(ln).LWrk1 = 1 THEN                                  ' Doing the 1st line of a range
            DECR L(i).LWrk1                                       ' Adjust # in reduced range
            tmp = L(i)                                            ' Save the XX marker line for a bit
            ARRAY DELETE L(i)                                     ' Delete the XX marker line from its old location
            L(i).LFlag = L(i).LFlag AND (&HFFFFFFFF - %Invisible) ' Make the line visible (i now = real line #)
            L(i).LFlag = L(i).LFlag OR %Popped                    ' Mark as popped during this interaction
            L(i).LWrk1 = 0                                        ' Clear Work in the Pop'ed out line
            ARRAY INSERT L(i + 1), tmp                            ' Stuff saved XX marker back in after Popped line
            FOR j = 1 TO L(i + 1).LWrk1                           ' Redo Work numbers in the new range
               L(i + 1 + j).LWrk1 = j                             '
            NEXT j                                                '
            METHOD = 0                                            ' Pass Back zero, we made no real inserts/deletes
            MExitMeth                                             '
         END IF
                                                                  ' Doing a line in middle of a range
         k = L(i).LWrk1 - L(ln).LWrk1                             ' Calc count for the new 'lower' range
         L(i).LWrk1 = L(ln).LWrk1 - 1                             ' Set a new count in current marker
         L(ln).LFlag = L(ln).LFlag AND (&HFFFFFFFF - %Invisible)  ' Pop out the requested line
         L(ln).LFlag = L(ln).LFlag OR %Popped                     ' Mark as popped during this interaction
         me.LInsertEmpty(ln, 1, %Xclude)                          ' Insert a line to hold the Exclude marker
         L(ln + 1).LWrk1 = k                                      ' Stuff in the calculated # of X'd lines in new area
         L(ln + 1).LLNum = "00000000"                             ' Clear line number

         FOR j = 1 TO k                                           ' Redo Work numbers in the new range
            L(ln + 1 + j).LWrk1 = j                               '
         NEXT j                                                   '
         METHOD = 1                                               ' Pass Back 1 for the inserted line
         me.AdjustPending(ln, 1, 0)                               ' Adjust pending stuff
         MExit
      END METHOD

      METHOD UndoPut(BYVAL pData AS UNDOType POINTER)
      REGISTER j AS LONG
      LOCAL tResult AS LONG
         MEntry
         @pData.pTIDX = VARPTR(TIDX)                              ' Build parameter list
         @pData.pL  = VARPTR(L(0))                                '
         @pData.pT  = VARPTR(T(0))                                '
         @pData.pTW = VARPTR(TW(0))                               '
         THREAD CREATE sUNDOSaveThread(BYVAL pData) 65536, TO UNDOThread ' Fire up the thread
         IF UNDOThread = 0 THEN                                   ' Failed?
            '?????                                                ' Say we failed
            MExitMeth                                             '
         END IF                                                   '
         THREAD STATUS UndoThread TO tResult                      ' See if running OK
         IF tResult <> 259 THEN                                   ' If running OK STATUS returns &H103 (See Help)
            THREAD CLOSE UNDOThread TO tResult                    ' Free up our handle
         END IF                                                   '
         j = 0                                                    '
         DO WHILE @PData.mCpyBusy = %True                         ' Wait till thread has done copies
            SLEEP 5                                               ' Wait a bit
            INCR j                                                ' Count for escape hatch
            IF j > 400 THEN EXIT DO                               ' Break out after 2 seconds
         LOOP                                                     '
         MExit
      END METHOD

      METHOD UndoGet(i AS LONG) AS LONG
      LOCAL j, ProfLine AS LONG, buf AS STRING, bufw AS WSTRING
      REGISTER ii AS LONG
      LOCAL uIO AS iIO                                            ' For our I/O stuff
         MEntry
         LET uIO = CLASS "cIO"                                    '

         uIO.Setup("BE", "", "", me.UndoIXFnGet(i))               ' Get the IX file
         uIO.EXEC                                                 '
         j = LOF(# uIO.FNum)                                      ' Get its length
         GET$ # uIO.FNum, j, TIDX                                 ' Read it back
         uIO.Close                                                ' Close it

         uIO.Setup("BE", "", "", me.UndoUFnGet(i))                ' Get the U file
         uIO.EXEC                                                 '
         j = LOF(# uIO.FNum) / SIZEOF(DataLine)                   ' Get size of an L entry
         REDIM L(j)                                               ' Redim L() to match save data
         GET # UIO.FNum, ,L()                                     ' Read the Data (Can't pass L() to a METHOD)
         uIO.Close                                                '

         uIO.Setup("BE", "", "", me.UndoTFnGet(i))                ' Get the T file
         uIO.EXEC                                                 '
         GET$ # uIO.FNum, LOF(# uIO.FNum), buf                    ' Read it back
         PARSE buf, T(), $CRLF                                    ' Assign back to the array
         uIO.Close                                                '

         FOR ii = 1 TO UBoundL                                    ' Loop re-assigning pointers
            IF IsLProf(ii) THEN ProfLine = %True                  ' Remember if PRof lines exist
            L(ii).LTxt = VARPTR(T(L(ii).LTxtIX))                  ' Correct the Txt pointer
         NEXT ii                                                  '

         uIO.Setup("BE", "", "", me.UndoTWFnGet(i))               ' Get the TW file
         uIO.EXEC                                                 '
         GET$$ # uIO.FNum, LOF(# uIO.FNum), bufw                  ' Read it back
         PARSE bufw, TW(), $CRLF                                  ' Assign back to the array
         uIO.Close                                                '

         FOR ii = 1 TO UBoundL                                    ' Loop re-assigning pointers
            L(ii).LAttr = VARPTR(TW(L(ii).LAttrIX))               ' Correct the Txt pointer
         NEXT ii                                                  '

         METHOD = IIF(ProfLine, &H00000001, 0)                    ' Say whether Prof lines exist
         MExit
      END METHOD

      METHOD IsTouchable(lin AS LONG, sCol AS LONG) AS LONG
      '---------- See if line area is 'touchable'
      LOCAL ldata AS STRING
         MEntry
         METHOD = %True                                           ' Default return
         IF TCtr = 0 THEN MExitMeth                               ' No touch entries, it's touchable
         ldata = me.TTblScan(lin)                                 ' See if a pending for this line
         IF ISNOTNULL(ldata) THEN METHOD = %False                 ' Swap it in
         MExit
      END METHOD

      METHOD Join() AS LONG
      '---------- Join  a line for JOIN/RJOIN
      LOCAL i, j, k, l, l1, l2, x, y AS LONG
      LOCAL TTxt, L1Txt, L2Txt, lclChg AS STRING
      LOCAL CTxt, CWrk, L1Clr, L2Clr AS WSTRING, AttrAsc AS WORD

         MEntry
         '----- Get common variables into local data
         i = cfFLine                                              ' Find line
         j = cfFCol                                               ' Column where it was found
         k = cfFLen                                               ' Length of found string (Use orig. length)
         l = LEN(L(i).@LTxt)                                      ' Length of text line
         GOSUB GetPrevAttr                                        ' Get current Attr of the found string

         '----- Prepare the Change string
         y = j                                                    ' Found data
         TTxt = LTxtG(i)                                          ' Copy of the line's text
         FOR x = 1 TO LEN(cfChange)                               ' Loop building the change string
            SELECT CASE AS CONST$ MID$(cfChange, x, 1)            ' Do it based on the char inside P'xxx'
               CASE "=": lclChg += MID$(TTxt, y, 1): INCR y       ' Copy the character
               CASE "<": lclChg += LLCASE(MID$(TTxt, y, 1)): INCR y ' Lowercase it
               CASE ">": lclChg += UUCASE(MID$(TTxt, y, 1)): INCR y ' Uppercase it
               CASE "!": lclChg += MID$(TTxt, j, k)               ' Copy found string
               CASE "~": INCR y                                   ' Ignore it
               CASE "\"                                           ' Escape?
                  lclChg += MID$(cfChange, x + 1, 1)              ' Copy next char
                  INCR x                                          ' extra step
                  IF IsCfL2Picture THEN INCR y                    ' If doing a picture
               '----- Not special character, just copy it
               CASE ELSE                                          ' Copy L12 character
                  lclChg += MID$(cfChange, x, 1)                  '
                  IF IsCfL2Picture THEN INCR y                    ' If doing a picture
            END SELECT                                            '
         NEXT x                                                   '

         '----- Setup the two line numbers
         IF IsCfLM THEN                                           ' LM join
            l1 = me.CRPBack(%mOData, i, 1)                        ' L1 point at previous line
            l2 = i                                                ' L2 at this line
            IF IsLTop(l1) THEN                                    ' Ignore Joins to the Top line
               sLine = l2 + 1: sCol = 0                           ' Skip onward
               cfNotFound = 3                                     ' New search needed
               MExitMeth                                          '
            END IF                                                '
            L1Txt = LTxtG(l1)                                     ' Get L1 text
            L1Clr = LAttrG(l1)                                    ' Get L1 Attr

            L2Txt = lclChg + MID$(LTxtG(l2), k + 1)               ' Get text
            L2Clr = REPEAT$(LEN(lclChg), CHR$$(AttrAsc)) + MID$(LAttrG(l2), k + 1)  ' And Attr
            sLine = l1: sCol = 2                                  ' Set resume point

         ELSE                                                     ' Else, RM join
            l1 = i                                                ' L1 points at this line
            l2 = me.CRPFwd(%MOData, i, 1)                         ' L2 points at next line
            IF l2 = LastLine THEN                                 ' Ignore Joins to the Bottom line
               cfNotFound = 0                                     '
               METHOD = %True                                     ' Tell caller to stop
               MExitMeth                                          '
            END IF                                                '
            cfNotFound = 3                                        ' New search needed

            L1Txt = LEFT$(LTxtG(l1), j - 1) + lclChg              ' Get L1 text
            L1Clr = LAttrG(l1)                                    ' Get L1 Attr
            L1Clr = LEFT$(L1Clr, j - 1) + REPEAT$(LEN(lclChg), CHR$$(AttrAsc))    '

            L2Txt = LTxtG(l2)                                     ' Get text
            L2Clr = LAttrG(l2)                                    ' And Attr
            sLine = l1: sCol = 0                                  ' Set resume point
         END IF                                                   '

         me.LFlagBitOff(l1, %EQChange)                            ' Reset ==CHG>
         me.LFlagBitOff(l2, %EQChange)                            '
         me.ModSet(l1)                                            ' Remember we changed something
         IF IsLPopped(l2) THEN                                    ' Did L2 just Pop?
            me.LFlagBitOff(l1, %Invisible)                        ' Ditto it to L1
            me.LFlagBitOn(l1, %Popped)                            '
         END IF                                                   '

         me.LTxtSet(l1, L1Txt + L2Txt)                            ' Swap in combined lines
         IF L(L1).LLbl = $BlankLNo THEN LLblS(l1) = TRIM$(L(l2).LLbl) ' Copy the .LLbl if it's blank
         IF LTagG(l1) = $BlankLNo THEN LTagS(l1) = TRIM$(LTagG(l2))  ' Copy the :tag if it's blank
         me.UpdLControl(l1)                                       '
         LAttrS(l1) = L1Clr + L2Clr                               ' Use it
         me.AttrScan(l1)                                          ' Recolorize

         me.lTxtFree(l2)                                          ' Remove the 2nd line
         me.LEntDel(l2)                                           '
         DECR LastLine: DECR LastReal                             ' Adjust LastLine and LastReal
         me.AdjustPending(l1, -1, 0)                              ' Adjust pending stuff
         DotEnd -= 1                                              ' Adjust range
         me.CurSetReq(%Position, l1, 1, %True)                    ' Set cursor set attempt

         MExitMeth                                                '

      GetPrevAttr:
         '----- First see if any existing highlighting on the found string
         AttrAsc = 0                                              '
         cTxt = LAttrG(i)                                         ' Get Attr line
         CWrk = MID$(ctxt, j, k)                                  ' Get Attr char for the found field
         IF CWrk = REPEAT$(k, LEFT$(CWrk, 1)) THEN                ' All the same color?
            AttrAsc = ASC(LEFT$(Cwrk, 1))                         ' Set AttrAsc to the current state
         END IF                                                   '
         RETURN                                                   '
      END METHOD

      METHOD LAllocTxt(ln AS LONG)
      '---------- Alloc a txt entry to this line
      LOCAL x, p, ol, nl, i AS LONG
         MEntry
         i = INSTR(TIDX, "1")                                     ' Scan for next available entry
         MID$(TIDX, i, 1) = "0"                                   ' Mark it no longer available
         T(i) = ""                                                ' Null the string
         L(ln).LTxtIX = i                                         ' Remember the index in the L() entry
         L(ln).LTxt = VARPTR(T(i))                                ' Point L.Txt to the dynamic String
         TW(i) = ""                                               ' Null the string
         L(ln).LAttrIX = i                                        ' Remember the index in the L() entry
         L(ln).LAttr = VARPTR(TW(i))                              ' Point L.Txt to the dynamic String
         MExit
      END METHOD

      METHOD MarkKill()
      '---------- UnMark and kill any text selection
         IF IsMarkDrawn THEN                                      ' Blank any hi-lite
            sCaretDestroy                                         '
            GRAPHIC SET MIX %MIX_NOT                              '
            GRAPHIC BOX (MarkRect.nLeft, MarkRect.nTop) - (MarkRect.nRight, MarkRect.nBottom), 0, %WHITE, -1
            OffMarkDrawn                                          '
            GRAPHIC REDRAW                                        '
            sCaretCreate                                          '
            sDoCursor                                             ' Get cursor back
            sCaretShow                                            '
         END IF                                                   '
         OffMarkActive                                            '

      '---------- UnMark and kill any misc selection
         IF IsMiscDrawn THEN                                      ' Blank any hi-lite
            sCaretDestroy                                         '
            GRAPHIC SET MIX %MIX_NOT                              '
            GRAPHIC BOX (MiscRect.nLeft, MiscRect.nTop) - (MiscRect.nRight, MiscRect.nBottom), 0, %WHITE, -1
            OffMiscDrawn                                          '
            GRAPHIC REDRAW                                        '
            sCaretCreate                                          '
            sDoCursor                                             ' Get cursor back
            sCaretShow                                            '
         END IF                                                   '
         OffMiscActive: MiscZone = 0                              '
      END METHOD

      METHOD MarkScr()
      '---------- Mark the selected area on the screen
      LOCAL tl, tt, tr, tb, vb, i, j AS LONG, cLoc, char AS STRING
         MEntry
         IF ISFALSE IsMarkActive OR IsFMTab THEN MExitMeth        ' Just in case
         SlecSCol = MarkSCol                                      ' Save for permanent select area
         SlecECol = MarkECol                                      '
         SlecSLin = VAL(LLNumG(MarkSLin))                         ' Save lines as external references
         SlecELin = VAL(LLNumG(MarkELin))                         '
         OnSlecSet                                                ' Say we've got an active set
         OnSlecActive                                             '

         FOR i = 3 + Prf.Cols TO 300                              ' Locate line on visible bottom of screen
            IF S(i) <> 0 THEN vb = i ELSE EXIT FOR                '
         NEXT i                                                   '
         tl = MAX(MarkSCol - Offset + gLNPadCol, gLNData1)        ' Calc char screen loc of left side
         tr = MIN(MarkECol - Offset + gLNPadCol, ENV.ScrWidth)    ' Calc char screen loc of right side
         tt = me.sSearch(MarkSLin)                                ' Calc char screen loc of top
         tb = me.sSearch(MarkELin)                                ' Calc char screen loc of bottom
         IF tt = -1 AND tb = -1 THEN MExitMeth                    ' Above the screen totally, just exit
         IF tt = -2 THEN MExitMeth                                ' Below the screen totally, just exit
         IF tt = -1 AND tb = -2 THEN                              ' Extends above and below the screen?
            tt = 3 + Prf.Cols: tb = vb                            ' Do the entire height
         ELSEIF tt = -1 AND tb > 0 THEN                           ' Extends from above into the screen?
            tt = 3 + Prf.Cols                                     ' Do top down to end line
         ELSEIF tt > 0 AND TB = -2 THEN                           ' Extends from middle of screen past the bottom
            tb = vb                                               ' Do from start line to bottom
         ELSEIF tt > 0 AND tb > 0 THEN                            ' Extends for part of screen
            '                                                     ' tt and tb should be OK
         END IF                                                   '
         IF tl > ENV.ScrWidth THEN MExitMeth                      ' Right of the screen totally
         IF tr < tl THEN MExitMeth                                '
         sCaretDestroy                                            ' Turn cursor off

         GRAPHIC SET MIX %MIX_NOT                                 '
         IF IsMarkDrawn THEN                                      ' Do we need to get rid of an old one?
            GRAPHIC BOX (MarkRect.nLeft, MarkRect.nTop) - (MarkRect.nRight, MarkRect.nBottom), 0, %WHITE, -1
            OffMarkDrawn                                          '
         END IF                                                   '
         MarkRect.nLeft = (tl - 1) * gFontWidth + %GLM            ' Build new location
         MarkRect.nRight = tr * gFontWidth + %GLM                 '
         MarkRect.nTop = (tt - 1) * gFontHeight                   '
         MarkRect.nBottom = tb * gFontHeight                      '
         GRAPHIC BOX (MarkRect.nLeft, MarkRect.nTop) - (MarkRect.nRight, MarkRect.nBottom), 0, %WHITE, -1
         GRAPHIC REDRAW                                           '
         OnMarkDrawn                                              '
         sCaretCreate                                             '
         sDoCursor                                                ' Get cursor back
         sCaretShow                                               '
         sDoStatusBar($SBSelect)                                  '
         MExit
      END METHOD

      METHOD MarkReDraw(lno AS LONG)
      '---------- Redraw Mark lines for a single screen line
      LOCAL i AS LONG, u AS STRING
         MEntry
         IF ISNULL(TRIM$(Prf.MarkLine)) OR ISFALSE Prf.PMark OR IsPTypeMode THEN MExitMeth ' Column marking OFF, just exit
         i = 1: u = Prf.MarkWorking                               ' Scan for Mark columns
         DO                                                       '
            i = INSTR(i, u, "*")                                  ' Look for next marker
            IF i = 0 THEN EXIT DO                                 '
            IF i >= Offset + 1 AND i <= Offset + gDataLen THEN    ' Within screen boundary
               GRAPHIC SET MIX %MIX_COPYSRC                       '
               GRAPHIC LINE ((i - 1 + gLNPadCol - Offset) * gFontWidth + %GLM - 1, (lno - 1) * gFontHeight) - ((i - 1 + gLNPadCol - Offset) * gFontWidth + %GLM - 1, lno * gFontHeight), ENV.cMarkLine
            END IF                                                '
            INCR i                                                '
         LOOP                                                     '
         MExit                                                    '
      END METHOD

      METHOD MarkSave()
      '---------- Save the Mark line by updating the INI file
         OffMarkAFlag                                             ' Clear flag that got us here
         Prf.MarkLine = TP.MarkLine                               ' Set it into the Profile
      END METHOD

      METHOD MarkStart(trow AS LONG, tcol AS LONG, StartRow AS LONG, StartCol AS LONG)
      '---------- See if starting a mouse mark area
      LOCAL j AS LONG
         IF tCol > gLNPadCol AND tRow > 2 + Prf.Cols AND _        ' Make sure its safe
            tRow < gwScrHeight + 1 AND Prf.HexMode = &1 AND ISFALSE IsFMTab THEN ' In true text area?
            IF IsLXClude(S(tRow)) OR _                            ' Can't start on an X'd line or Top/Bottom
               IsLTop(S(tRow)) OR _                               '
               IsLBottom(S(tRow)) THEN sDoBeep: EXIT METHOD       '
            me.MarkKill                                           ' Kill any active text block select
            me.MarkKill                                           ' Kill any active cmd  select
            MarkSCol = tCol + Offset - gLNPadCol                  ' Remember starting x/y of Select frame
            MarkSLin = S(tRow)                                    '
            MarkECol = MarkSCol                                   '
            MarkELin = MarkSLin                                   '
            IF IsSwapActive THEN                                  ' If a SWAP is active
               IF MarkSLin >= SwapSLin AND MarkSLin <= SwapELin THEN ' Test start line
                  IF MarkSCol >= SwapSCol AND MarkSCol <= SwapECol THEN sDoBeep: EXIT METHOD ' And horizontally?
               END IF                                             '
               j = SwapLines - 1                                  ' Get number of data lines in Swap group
               DO WHILE j                                         ' See if enough lines here
                  IF IsLData(MarkELin) THEN                       ' If a data line
                     DECR j                                       ' Reduce count
                  END IF                                          '
                  INCR MarkELin                                   ' Next line
                  IF IsLBottom(MarkELin) THEN sDoBeep: EXIT METHOD'
               LOOP                                               '
               IF MarkELin >= SwapSLin AND MarkELin <= SwapELin THEN ' Test start line
                  IF MarkECol >= SwapSCol AND MarkECol <= SwapECol THEN sDoBeep: EXIT METHOD ' And horizontally?
               END IF                                             '
            END IF                                                '
            StartRow = MarkSLin                                   '
            StartCol = MarkSCol                                   '
            OnMarkActive                                          '

         '----- In the Cmd area, start a Cmd highlight
         ELSEIF tRow = 1 AND tCol < 10 + pCmdLen THEN             ' Make sure its safe
            me.MarkKill                                           ' Kill any active text block select
            me.MarkKill                                           ' Kill any active cmd  select
            MiscSCol = tCol + COffset - 10                        ' Remember starting x/y of Select frame
            MiscECol = MiscSCol                                   '
            StartRow = 1                                          '
            StartCol = MiscSCol                                   '
            MiscZone = %CursCmnd                                  ' Which area are we marking
            OnMiscActive                                          '
         END IF
         gfLeftDown = %True                                       '
      END METHOD

      METHOD MaskSave()
      '---------- Save the Mask line by updating the INI file
         OffMaskAFlag                                             ' Clear flag that got us here
         Prf.MaskLine = TP.MaskLine                               ' Set it into the Profile
      END METHOD

      METHOD MEditListSet(BYVAL ix AS LONG, fn AS STRING)
         IF ix > UBOUND(MEditList()) THEN _
            REDIM PRESERVE MEditList(1 TO 2 * UBOUND(MEditList())) AS INSTANCE STRING
         MEditList(ix) = fn
      END METHOD

      METHOD MEditListGet(BYVAL ix AS LONG) AS STRING: METHOD = MEditList(ix): END METHOD

      METHOD MEditFlagSet(BYVAL ix AS LONG, vl AS LONG)
         IF ix > UBOUND(MEditFlag()) THEN _
            REDIM PRESERVE MEditFlag(1 TO 2 * UBOUND(MEditFlag())) AS INSTANCE LONG
         MEditFlag(ix) = vl
      END METHOD

      METHOD MEditFlagGet(BYVAL ix AS LONG) AS LONG: METHOD = MEditFlag(ix): END METHOD

      METHOD MeditTbl(func AS STRING, fn AS STRING) AS LONG
      '---------- Manipulate the MEdit table
      LOCAL i, j, k AS LONG
         MEntry
         METHOD = 0                                               ' Default RC
         SELECT CASE AS CONST$ func                               ' What function?
            CASE "S"                                              ' Search
               IF MEditCount = 0 THEN MExitMeth                   ' No table, then not found
               FOR i = 1 TO MEditCount                            ' Search table then
                  IF IsEQ(fn, me.MEditListGet(i)) THEN METHOD = i: MExitMeth
               NEXT i                                             '

            CASE "A"                                              ' Add, Search should have failed already
               MEditCount += 1                                    ' Bump
               me.MEditListSet(MEditCount, fn)                    ' Add the new filename
               me.MEditFlagSet(MEditCount, %False)                ' Clear it's flag
               METHOD = MEditCount                                ' Pass back count (index)

            CASE "D"                                              ' Delete
               IF MEditCount = 0 THEN MExitMeth                   ' No table, then not found
               k = 0                                              ' Found index
               FOR i = 1 TO MEditCount                            ' Search table then
                  IF IsEQ(fn, me.MEditListGet(i)) THEN _          ' Got it?
                     k = i: EXIT FOR                              ' Save found entry
               NEXT i                                             '
               FOR i = k TO MEditCount - 1                        ' Ripple table down
                  me.MEditListSet(i, me.MEditListGet(i + 1))      '
                  me.MEditFlagSet(i, me.MEditFlagGet(i + 1))      '
               NEXT i                                             '
               MEditCount -= 1                                    '
               METHOD = k                                         ' Pass back deleted entry number

            CASE "I"                                              ' Get Index from line number
               FOR i = VAL(fn) TO 1 STEP -1                       ' BackScan
                  j = LMIXG(i)                                    ' Get line's MIX value
                  IF j <> 0 THEN METHOD = j: MExitMeth            ' Found, we're done, pass it back
               NEXT i                                             '
               METHOD = 1: MExitMeth                              ' In case not MEdit mode

         END SELECT
         MExit
      END METHOD

      METHOD MiscMark()
      '---------- Mark a misc area
      LOCAL tl, tt, tr, tb, vb, i, j AS LONG, cLoc, char AS STRING
         MEntry
         IF ISFALSE IsMiscActive THEN MExitMeth                   ' Just in case
         sCaretDestroy                                            '
         GRAPHIC SET MIX %MIX_NOT                                 '
         IF IsMiscDrawn THEN                                      ' Do we need to get rid of an old one?
            GRAPHIC BOX (MiscRect.nLeft, MiscRect.nTop) - (MiscRect.nRight, MiscRect.nBottom), 0, %WHITE, -1
            OffMiscDrawn                                          '
         END IF                                                   '

         SELECT CASE AS LONG MiscZone                             ' Which area are we marking
            CASE %CursCmnd                                        ' Command line
               tl = MAX(MiscSCol - COffset + 10, 11)              ' Calc char screen loc of left side
               tr = MIN(MiscECol - COffset + 10, 10 + pCmdLen)    ' Calc char screen loc of right side
               tt = 1                                             ' Calc char screen loc of top
               tb = 1                                             ' Calc char screen loc of bottom
               IF tr < tl THEN MExitMeth                          '
               MiscRect.nLeft = (tl - 1) * gFontWidth + %GLM      ' Build new location
               MiscRect.nRight = tr * gFontWidth + 1              '
               MiscRect.nTop = (tt - 1) * gFontHeight             '
               MiscRect.nBottom = tb * gFontHeight                '

            CASE %CursPath                                        ' FM Path line
               tl = FM_Path_Left + MiscSCol - 1                   ' Calc char screen loc of left side
               tr = FM_Path_Left + MiscECol - 1                   ' Calc char screen loc of right side
               tt = FM_Path_Line                                  ' Calc char screen loc of top
               tb = FM_Path_Line                                  ' Calc char screen loc of bottom
               IF tr < tl THEN MExitMeth                          '
               MiscRect.nLeft = (tl - 1) * gFontWidth + %GLM      ' Build new location
               MiscRect.nRight = tr * gFontWidth + 1              '
               MiscRect.nTop = (tt - 1) * gFontHeight             '
               MiscRect.nBottom = tb * gFontHeight                '

            CASE %CursMask                                        ' FM Mask line
               tl = FM_Mask_Left + MiscSCol - 1                   ' Calc char screen loc of left side
               tr = FM_Mask_Left + MiscECol - 1                   ' Calc char screen loc of right side
               tt = FM_Mask_Line                                  ' Calc char screen loc of top
               tb = FM_Mask_Line                                  ' Calc char screen loc of bottom
               IF tr < tl THEN MExitMeth                          '
               MiscRect.nLeft = (tl - 1) * gFontWidth + %GLM      ' Build new location
               MiscRect.nRight = tr * gFontWidth + 1              '
               MiscRect.nTop = (tt - 1) * gFontHeight             '
               MiscRect.nBottom = tb * gFontHeight                '

            CASE %CursLNote                                       ' FM Note line
               tl = FM_Head_Note_Left + MiscSCol - 1 - AFList(MiscLine).NoteOff ' Calc char screen loc of left side
               tr = FM_Head_Note_Left + MiscECol - 1 - AFList(MiscLine).NoteOff ' Calc char screen loc of right side
               tt = MiscLine - TopScrn + FM_Top_File_Line         ' Calc char screen loc of top
               tb = tt                                            ' Calc char screen loc of bottom
               IF tr < tl THEN MExitMeth                          '
               MiscRect.nLeft = (tl - 1) * gFontWidth + %GLM      ' Build new location
               MiscRect.nRight = tr * gFontWidth + 1              '
               MiscRect.nTop = (tt - 1) * gFontHeight             '
               MiscRect.nBottom = tb * gFontHeight                '

         END SELECT
         GRAPHIC BOX (MiscRect.nLeft, MiscRect.nTop) - (MiscRect.nRight, MiscRect.nBottom), 0, %WHITE, -1
         GRAPHIC REDRAW                                           '
         OnMiscDrawn                                              '
         sCaretCreate                                             '
         sDoCursor                                                ' Get cursor back
         sCaretShow                                               '
         MExit
      END METHOD

      METHOD OffsetAdd (amt AS LONG): Offset += amt: END METHOD
      METHOD OffsetSub (amt AS LONG): Offset = MAX(0, Offset - amt): END METHOD
      METHOD COffsetAdd (amt AS LONG): COffset += amt: END METHOD
      METHOD COffsetSub (amt AS LONG): COffset = MAX(0, COffset - amt): END METHOD

      METHOD OverlayText(SOURCE AS STRING, Dest AS STRING) AS STRING
      '---------- Overlay text from one string to another
      REGISTER i AS LONG
      LOCAL TX AS STRING, lclRBnd AS LONG
         MEntry
         TX = Dest                                                ' Copy Dest as a work string
         lclRBnd = Prf.BndRight: IF lclRBnd = 0 THEN lclRBnd = MaxLength    ' Get correct right bounds
         IF LEN(SOURCE) > LEN(TX) THEN _                          ' The Source string is longer
            TX = LSET$(TX, LEN(SOURCE))                           ' Make the Dest as long as the Source

         FOR i = 1 TO LEN(SOURCE)                                 ' Loop through length of Source string
            IF Prf.BndLeft > 1 OR Prf.BndRight > 0 THEN           ' BNDS active?
               IF i >= Prf.BndLeft AND i <= lclRBnd THEN          ' Yes, current column within BNDS?
                  IF MID$(TX, i, 1) = " " THEN MID$(TX, i, 1) = MID$(SOURCE, i, 1) ' Overlay if Dest is blank
               END IF                                             '
            ELSE                                                  ' No BNDS
               IF MID$(TX, i, 1) = " " THEN MID$(TX, i, 1) = MID$(SOURCE, i, 1) ' Overlay if Dest is blank
            END IF                                                '
         NEXT I                                                   '

         METHOD = TX                                              ' Pass back the result
         MExit
      END METHOD

      METHOD OverlayTextDel(SOURCE AS STRING, Dest AS STRING) AS STRING
      '---------- Overlay text from one string to another, return whether all data merged or not
      REGISTER i AS LONG
      REGISTER j AS LONG
      LOCAL STxt, DTxt AS STRING, lclRBnd AS LONG
         MEntry
         OffOvrTextDel                                            ' Start as saying no error occurred
         lclRBnd = Prf.BndRight: IF lclRBnd = 0 THEN lclRBnd = MaxLength ' Get correct right bounds
         DTxt = Dest: STxt = SOURCE                               ' Copy to work strings
         IF LEN(STxt) > LEN(DTxt) THEN _                          ' The Source string is longer
            DTxt = LSET$(DTxt, LEN(STxt))                         ' Make the Dest as long as the Source

         FOR i = 1 TO LEN(STxt)                                   ' Loop through length of Source string
            IF Prf.BndLeft > 1 OR Prf.BndRight > 0 THEN           ' BNDS active?
               IF i >= Prf.BndLeft AND i <= lclRBnd THEN          ' Yes, current column within BNDS?
                  IF MID$(DTxt, i, 1) = " " OR MID$(DTxt, i, 1) = MID$(STxt, i, 1) THEN '
                     MID$(DTxt, i, 1) = MID$(STxt, i, 1)          ' Overlay if Dest is blank or = char
                     MID$(STxt, i, 1) = " "                       ' Blank the char in the source string
                  END IF                                          '
               END IF                                             '
            ELSE                                                  ' No bounds, do it the simple way
               IF MID$(DTxt, i, 1) = " " OR MID$(DTxt, i, 1) = MID$(STxt, i, 1) THEN '
                  MID$(DTxt, i, 1) = MID$(STxt, i, 1)             ' Overlay if Dest is blank
                  MID$(STxt, i, 1) = " "                          ' Blank the char in the source string
               END IF                                             '
            END IF                                                '
         NEXT I                                                   '

         METHOD = DTxt                                            ' Pass back the result
         IF ISNOTNULL(TRIM$(Stxt)) THEN OnOvrTextDel              ' If Source not all merged, then flag an error
         MExit
      END METHOD

      METHOD OverlayTextRepl(SOURCE AS STRING, Dest AS STRING) AS STRING
      '---------- Overlay text replace from one string to another
      REGISTER i AS LONG
      LOCAL TX AS STRING, lclRBnd AS LONG
         MEntry
         TX = Dest                                                ' Copy Dest as a work string
         lclRBnd = Prf.BndRight: IF lclRBnd = 0 THEN lclRBnd = MaxLength  ' Get correct right bounds
         IF LEN(SOURCE) > LEN(TX) THEN _                          ' The Source string is longer
            TX = LSET$(TX, LEN(SOURCE))                           ' Make the Dest as long as the Source

         FOR i = 1 TO LEN(SOURCE)                                 ' Loop through length of Source string
            IF Prf.BndLeft > 1 OR Prf.BndRight > 0 THEN           ' BNDS active?
               IF i >= Prf.BndLeft AND i <= lclRBnd THEN          ' Yes, current column within BNDS?
                  IF MID$(SOURCE, i, 1) <> " " THEN _             ' If source is non-blank
                     MID$(TX, i, 1) = MID$(SOURCE, i, 1)          ' Copy it
               END IF                                             '
            ELSE                                                  ' No BNDS
               IF MID$(SOURCE, i, 1) <> " " THEN _                ' If source is non-blank
                  MID$(TX, i, 1) = MID$(SOURCE, i, 1)             ' Copy it
            END IF                                                '
         NEXT I                                                   '

         METHOD = TX                                              ' Pass back the result
         MExit
      END METHOD

      METHOD pCmdCharRep (posit AS LONG, char AS STRING)
         IF posit > LEN(pCommand) THEN                            ' If needed, lengthen the string
            pCommand += SPACE$(posit - LEN(pCommand))             '
         END IF                                                   '
         MID$(pCommand, posit, 1) = char                          ' Stuff in the character
      END METHOD

      METHOD PicSetWord():    PIC.WORD_WORD_changed = 1: END METHOD
      METHOD PicSetCharSet(): PIC.CharSet_changed = 1:   END METHOD
      METHOD PicSetAll(): PIC.defaults_initialized = 0:   END METHOD

      METHOD PicInit()
      '---------- initialize class-specific data areas for Picture processing
      LOCAL i, ASCII AS LONG, lclCharSet AS STRING
         MEntry
         IF PIC.defaults_initialized = 0 THEN                     ' Do all?
            FOR i = 0 TO 255                                      ' Reset tables
               PIC.data_mask(i) = 0: PIC.pic_mask(i) = 0          '
            NEXT i                                                '
            PIC.WORD_WORD_changed = 1                             ' Force rebuild of WORD mapping
            PIC.CharSet_changed = 1                               ' Force rebuild of P'.' mapping
            PIC.mapping_initialized = 0                           ' Force basic initialize
            PIC.defaults_initialized = 1                          ' Do basic class initialization just once
         END IF
         IF PIC.mapping_initialized = 0 THEN

            FOR ASCII = 0 TO 255                                  ' Create mapping between data char and characterizing flag bits
               PIC.data_mask(ASCII) = (%PicMap_Any OR %PicMap_NonBlank) ' Map P'=' and also initialize other bits to 0
               IF (ASCII >= ASC("0")) AND (ASCII <= ASC("9")) THEN' Map P'#' and P'-'
                  PIC.data_mask(ASCII) OR= %PicMap_Numeric        '
               ELSE                                               '
                  PIC.data_mask(ASCII) OR= %PicMap_NonNumeric     '
               END IF                                             '
            NEXT                                                  '
            PIC.data_mask(ASC(" ")) XOR= %PicMap_NonBlank         ' Map P'^' and P'¬' (unmap only position which isn't nonblank)
            FOR I = 1 TO LEN(Prf.ChrUpper)                        ' Map P'>' and P'@'
               ASCII = ASC(Prf.ChrUpper, I)                       '
               PIC.data_mask(ASCII) OR= %PicMap_Upper             '
               PIC.data_mask(ASCII) OR= %PicMap_Alpha             '
            NEXT                                                  '
            FOR I = 1 TO LEN(Prf.ChrLower)                        ' Map P'<' and P'@'
               ASCII = ASC(Prf.ChrLower, I)                       '
               PIC.data_mask(ASCII) OR= %PicMap_Lower             '
               PIC.data_mask(ASCII) OR= %PicMap_Alpha             '
            NEXT                                                  '
            FOR ASCII = 0 TO 255                                  ' Map P'$'
               IF (PIC.data_mask(ASCII) AND (%PicMap_Alpha OR %PicMap_Numeric)) = 0 THEN
                  IF ASCII <> ASC(" ") THEN                       ' (undocumented) IBM ISPF P'$' does not match blanks
                     PIC.data_mask(ASCII) OR= %PicMap_Special     ' characters that are not P'@' or P'#' are P'$'
                  END IF                                          '
               END IF                                             '
            NEXT                                                  '

            FOR ASCII = 0 TO 255                                  ' Create picture map mask array to categorize the picture codes
               PIC.pic_mask(ASCII) = %PicMap_Literal              '
            NEXT                                                  '

            PIC.pic_mask(ASC("=")) = %PicMap_Any                  '
            PIC.pic_mask(ASC("^")) = %PicMap_NonBlank             '
            PIC.pic_mask(170)      = %PicMap_NonBlank             '  hex AA not sign (OEM)
            PIC.pic_mask(172)      = %PicMap_NonBlank             '  hex AC not sign (Ansi)
            PIC.pic_mask(ASC(".")) = %PicMap_NonDisplay           '
            PIC.pic_mask(ASC("#")) = %PicMap_Numeric              '
            PIC.pic_mask(ASC("-")) = %PicMap_NonNumeric           '
            PIC.pic_mask(ASC("@")) = %PicMap_Alpha                '
            PIC.pic_mask(ASC("<")) = %PicMap_Lower                '
            PIC.pic_mask(ASC(">")) = %PicMap_Upper                '
            PIC.pic_mask(ASC("$")) = %PicMap_Special              '
            PIC.pic_mask(ASC("&")) = %PicMap_WORD_Char            '
            PIC.pic_mask(ASC("%")) = %PicMap_Non_Word_Char        '
            PIC.pic_mask(ASC("{")) = %PicMap_LMargin              '
            PIC.pic_mask(ASC("}")) = %PicMap_RMargin              '
            FOR ASCII = 0 TO 255                                  ' Note literal alphabetic chars subject to CASE C/T handling
               IF (PIC.data_mask(ASCII) AND %PicMap_Alpha) THEN   '
                  PIC.pic_mask(ASCII) = %PicMap_LiteralAlpha      '
               END IF                                             '
            NEXT                                                  '
            PIC.pic_mask(ASC("\")) = %PicMap_Escape               '
            PIC.mapping_initialized = 1                           '
         END IF                                                   '

         IF PIC.WORD_WORD_changed THEN                            ' If a change to the WORD string, redo maps for P'%' and P'&'
            FOR ASCII = 0 TO 255                                  ' Clear all WORD/Non-WORD bits first
               PIC.data_mask(ASCII) AND= (NOT (%PicMap_WORD_Char OR %PicMap_Non_WORD_Char))
            NEXT                                                  '
            FOR I = 1 TO LEN(Prf.PWord)                           ' Mark the WORD chars
               ASCII = ASC(Prf.PWord, I)                          '
               PIC.data_mask(ASCII) OR= %PicMap_WORD_Char         '
            NEXT
            FOR ASCII = 0 TO 255                                  ' Mark the non-WORD chars as WORD chars
               IF (PIC.data_mask(ASCII) AND %PicMap_WORD_Char) = 0 THEN
                   PIC.data_mask(ASCII) OR= %PicMap_Non_WORD_Char '
               END IF                                             '
            NEXT                                                  '
            PIC.WORD_WORD_changed = 0
         END IF

         IF PIC.CharSet_changed THEN                              ' If a change to the Charset string, redo maps for P'.'
            lclCharset = ENV.CharSet + " "                        ' Add a blank to normal CharSet
            FOR ASCII = 0 TO 255                                  ' Set all %PicMap_NonDisplay first
               PIC.data_mask(ASCII) OR= %PicMap_NonDisplay        '
            NEXT                                                  '
            FOR I = 1 TO LEN(lclCharSet)                          ' Clear non-display from displayable chars
               ASCII = ASC(lclCharSet, I)                         '
               PIC.data_mask(ASCII) XOR= %PicMap_NonDisplay       '
            NEXT                                                  '
            PIC.CharSet_changed = 0
         END IF
         MExit
      END METHOD

      METHOD PicSearch(BYREF arg_data AS STRING, BYREF arg_pic AS STRING, BYREF arg_start_col AS LONG, _
                          BYREF ret_col  AS LONG, BYREF ret_len  AS LONG)
      '---------- Perform Picture type search
      LOCAL len_true_pic, scan_first_col, scan_last_col, scan_incr, scan_col, curr_data_mask, j, k AS LONG
      LOCAL lclPic, lclData AS STRING
      REGISTER curr_pic_mask AS LONG
      REGISTER i AS LONG
      LOCAL dp, pp AS BYTE POINTER
         MEntry
         ret_col = 0                                              ' Default return values
         ret_len = 0                                              '
         lclPic = arg_pic                                         ' Copy the picture

         '----- Fudge mask for { usage
         IF INSTR(lclPic, "{") > 0 THEN                           ' Got a leading { picture?
            lclPic = MID$(lclPic, 2)                              ' Remove it
            IF LEFT$(arg_data, 1) = " " THEN                      ' And some leading spaces?
               lclPic = SPACE$(VERIFY(arg_data, " ") - 1 - (VERIFY(lclPic, " ") - 1) ) + lclPic' Rebuild picture string
            END IF                                                '
         END IF                                                   '

         '----- Fudge mask for } usage
         IF INSTR(arg_pic, "}") > 0 THEN                          ' Got a trailing } picture?
            lclPic = LEFT$(LclPic, LEN(lclPic) - 1)               ' Remove it
            RESET j, k                                            ' Clear some work vars
            IF RIGHT$(lclPic, 1) = " " THEN                       ' Trailing spaces in the picture?
               FOR i = LEN(lclPic) TO 1 STEP -1                   ' See how many
                  IF MID$(lclPic, i, 1) = " " THEN INCR k         ' Count trailing blanks
                  IF MID$(lclPic, i, 1) <> " " THEN EXIT FOR      ' Exit when a non-blank
               NEXT i                                             '
            END IF                                                '
            IF RIGHT$(arg_data, 1) = " " THEN                     ' And some trailing spaces?
               FOR i = LEN(arg_data) TO 1 STEP -1                 ' See how many
                  IF MID$(arg_data, i, 1) <> " " THEN j = i: EXIT FOR ' Find last non-blank
               NEXT i                                             '
               IF j > 0 THEN                                      ' Found it
                  lclPic = lclPic + SPACE$(LEN(arg_data) - j - k) ' Rebuild picture string
               END IF                                             '
            END IF                                                '
         END IF                                                   '

         IF ISFALSE arg_start_col THEN MExitMeth                  ' Exit if invalid arg
         IF ISFALSE LEN(arg_data) OR ISFALSE LEN(arg_pic) OR ABS(arg_start_col) > LEN(arg_data) THEN MExitMeth ' Sanity check on line and picture string lengths

         i = 1
         DO WHILE i <= LEN(lclPic)                                ' Get Pic true length
            INCR len_true_pic                                     '
            i = i + IIF(i < LEN(lclPic) AND MID$(lclPic, i, 1) = "\", 2, 1)
         LOOP                                                     '
         '+--------------------------------------------------------------------------
         '| Compute a scan window to search for
         '| we create a left-justified window so only complete pic strings are
         '| checked for.
         '|
         '| example: line is 30 long, pic is 5 long
         '| forward, start = 1, scan column range is 1 to 26 (30-5+1)
         '| backward, start = 30, scan is from 26 (30-5+1) to 1
         '|
         '| ----+----1----+----2----+----3
         '| 12345                    12345
         '+--------------------------------------------------------------------------
         IF arg_start_col > 0 THEN                                ' Validate arg_start_col
            IF len_true_pic > LEN(arg_data) - arg_start_col + 1 THEN MExitMeth ' Pic is larger than remaining data
            scan_first_col = arg_start_col                        ' Forward scan: right boundary is fixed, left is moving
            scan_last_col = LEN(arg_data) - len_true_pic + 1      '
            scan_incr = 1                                         '
         ELSEIF arg_start_col < 0 THEN                            ' Backward scan: left boundary is fixed at 1, right is moving
            scan_last_col = 1                                     ' Set left boundary
            scan_first_col = LEN(arg_data) + arg_start_col + 1    ' Get start scan
            IF len_true_pic > scan_first_col THEN MExitMeth       ' Pic is larger than remaining data
            scan_incr = -1                                        '
         END IF                                                   '

         FOR scan_col = scan_first_col TO scan_last_col STEP scan_incr  ' Do the search
            GOSUB ScanPicture                                     '
            IF ret_col <> 0 THEN EXIT FOR                         ' We're done loop if found
         NEXT                                                     '
         MExitMeth

      ScanPicture:
         pp = STRPTR(lclPic): dp = STRPTR(arg_data) + scan_col - 1' Point at picture and data strings
         FOR i = 1 TO len_true_pic                                ' Loop for picture length
            curr_pic_mask  = PIC.pic_mask(@pp)                    ' Get mask which maps meaning of pic char
            curr_data_mask = PIC.data_mask(@dp)                   '

            IF curr_pic_mask = %PicMap_Escape THEN                '
               IF i < LEN(lclPic) THEN INCR pp                    ' Next char is taken literally; a final \ is used as-is
               curr_pic_mask = %PicMap_Literal                    ' It must be exact match
            END IF                                                '

            IF curr_pic_mask = %PicMap_Literal THEN               '  Begin comparison process.  RETURN on first failed match
               IF @dp <> @pp THEN RETURN                          ' Literal pic char must match exactly to data char

            ELSEIF (curr_pic_mask AND %PicMap_LiteralAlpha) THEN  ' Literal alpha is exact match if CASE C, or insensitive if CASE T
               IF Prf.PCase = "C" THEN                            '
                  IF @dp <> @pp THEN RETURN                       '
               ELSE                                               ' Assuming "T" if not "C"
                  IF IsNE(CHR$(@dp),CHR$(@pp)) THEN RETURN        '
               END IF                                             '
            ELSE                                                  ' all remaining pic codes are class tests
               IF (curr_pic_mask AND curr_data_mask) = 0 THEN RETURN '
            END IF                                                '
            INCR pp: INCR dp                                      ' Step to next picture column
         NEXT                                                     '
         ret_col = scan_col: ret_len = len_true_pic               ' if no failure, set success values
         RETURN                                                   '
      END METHOD                                                  '

      METHOD Search(MODE AS INTEGER, FCursMode AS INTEGER) AS INTEGER
      '---------- This is the single does-it-all string search routine
      LOCAL i, j, k, ii, jj, kk, l, x, lclRBnd, lclLBnd, EventCount, PClr AS LONG, t1, t2, c1, t, lclWord AS STRING, ONCE, Srch, RevOnce AS LONG, PosType AS INTEGER
      STATIC PassCount AS LONG
         MEntry
         i = sLine                                                ' Init values for debugging
         j = sCol                                                 '

         '----- Get set up, set Failure as the default exit
         lclWord = Prf.PWord: Srch = ISTRUE IsCfLit1              ' Fetch local copies
         ONCE = %True                                             '
         METHOD = %False                                          ' Start off saying we've failed
         IF LastLine > 50000 THEN gLoopCtr = - 1                  ' Prevent loop detection for big files

         '----- Handle the REVERSE option
         IF IsRevOnce THEN                                        ' Reverse one time only?
            OffRevOnce: RevOnce = %True                           ' Remove globally, remember locally
         END IF                                                   '

         '----- If we're actually doing a search string, CASE process it
         IF Srch THEN                                             ' If a literal type search
            t2 = cfFind                                           ' Setup data from Search arg
            IF ISFALSE IsCfL1CaseComp AND _                       '
               ISFALSE IsCfL1RegEx AND _                          '
               ISFALSE IsCfL1Picture THEN t2 = UUCASE(t2)         ' UC if not case sensitive search or RegEx
         END IF                                                   '

         '----- This is an initial FIND/CHANGE vs an RFIND/RCHANGE
         IF MODE = 0 THEN                                         ' Mode = 0 = Full Find

            '----- Figure out the line and column we're to start on
            sLine = TopScrn: sCol = 1: sDir = 1                   ' Set start defaults
            IF IsCfAll THEN sLine = 1                             ' If ALL, start at line 1
            IF FCursMode = %CursData AND ISFALSE IsCfAll THEN     ' In the Data Area and not an ALL request?
               sLine = sGetIX(CRow)                               ' Swap in current cursor location
               sCol = CCol - gLNPadCol + Offset                   ' Calc column position in line
               IF IIF(RevOnce, -1 * sDir, sDir) = 1 THEN          ' Forward?
                  sCol += 1                                       ' Adjust right by 1
               ELSE                                               '
                  sCol = MAX(1, sCol - 1)                         '
               END IF                                             '
            END IF                                                '
            IF FCursMode = %CursLinn AND ISFALSE IsCfAll THEN     ' In the Line # Area and not an ALL request?
               sLine = sGetIX(CRow): sCol = 1                     ' Swap in current cursor location and Col 1
            END IF                                                '
            IF IsCfFirst THEN                                     ' If FIRST then start as the top
               sLine = 1: sCol = 1                                '
               CfClear(%CRTFirst)                                 ' Remove so we do it just once
               CfSet(%CRTNext)                                    ' Make it NEXT
               CfClear(%CRTPrev)                                  ' Kill Prev
            END IF                                                '
            IF IsCfLast THEN                                      ' If LAST then start as the bottom
               sLine = LastLine: sDir = -1                        '
               CfClear(%CRTLast): CfSet(%CRTPrev)                 ' Remove LAST and insert PREV
               CfClear(%CRTNext)                                  '
            END IF                                                '
            IF IsCfPrev THEN                                      ' If Prev set reverse
               CfSet(%CRTPrev): sDir = -1                         '
               CfClear(%CRTNext)                                  '
            ELSE                                                  '
               sDir = 1                                           '
            END IF                                                '

            '----- Set the search direction up
            IF IsCfPrev THEN                                      ' If PREV then reverse search direction
               sDir = -1: sCol = -1                               '
               sLine = MAX(sLine - 1, 1)                          '
            END IF                                                '

         '----- This is the RFIND/RCHANGE setup
         ELSEIF MODE = 1 THEN                                     ' Mode = 1 = ReFind

            '----- Watch out for rollover at Top/Bottom
            IF cfNotFound = 0 OR cfNotFound = 3 THEN              ' Have we just hit Top/Bottom not found?
               cfNotFound = 0                                     ' Reset it

               '----- Not a Top/Bottom rollover, set our starting line and column
               IF FCursMode = %CursInternal THEN                  ' sCol/sLin internally set?
                  '
               ELSEIF FCursMode = %CursData THEN                  ' No, In the Data Area?
                  sCol = CCol - gLNPadCol + Offset                ' Calc column position in line
                  l = sGetIX(CRow)                                ' Get line cursor is on
                  IF l AND IsLXclude(l) THEN                      ' Is this an Xclude line?
                     i = CsrLinDX                                 '
                     IF CsrLinDX <> 0 THEN sLine = sFindLineNum(CsrLinDX) ' Continuing from within an X'd block?  Use that line number then
                  ELSE                                            '
                     sLine = sGetIX(CRow)                         ' Swap in current cursor location
                     IF IIF(RevOnce, -1 * sDir, sDir) = 1 THEN    ' Forward?
                        sCol += 1                                 ' Adjust right by 1
                     ELSE                                         '
                        sCol = MAX(1, sCol - 1)                   '
                     END IF                                       '
                  END IF                                          '
                  IF IsCfLEFT THEN                                ' If previous a LEFT
                     IF sDir = 1 THEN                             ' Going forward?
                        sLine += 1: sCol = 1                      ' Next line, col 1
                     ELSE                                         '
                        sLine -= 1: sCol = -1                     ' Next line, col 1
                     END IF                                       '
                  END IF                                          '
               ELSEIF FCursMode = %CursCmnd THEN                  ' In the Command area?
                  sLine = TopScrn                                 ' Swap in top line of screen
                  sCol = 1                                        ' and Col 1
               ELSEIF FCursMode = %CursLinN THEN                  ' In the Line # area?
                  sLine = sGetIX(CRow)                            ' Swap in the line #
                  sCol = 1                                        ' and Col 1
               ELSEIF FCursMode = %CursBad THEN                   ' Bad cursor
                  IF IIF(RevOnce, -1 * sDir, sDir) = 1 THEN       ' Forward?
                     sCol += 1                                    ' Adjust right by 1
                     IF IsCfLEFT THEN                             ' If previous a LEFT
                        IF sDir = 1 THEN                          ' Going forward?
                           sLine += 1: sCol = 1                   ' Next line, col 1
                        ELSE                                      '
                           sLine -= 1: sCol = -1                  ' Next line, col 1
                        END IF                                    '
                     END IF                                       '
                  ELSE                                            '
                     IF IsCfL1Picture THEN                        ' If P'xxx'
                        sCol = sCol - LEN(L(sLine).@LTxt) - 2     ' Yes, calc new scan start
                     ELSEIF IsCfL1RegEx THEN                      ' If Regex
                        sCol = sCol - LEN(L(sLine).@LTxt) - 2     ' Yes, calc new scan start
                        sLine -= 1                                '
                     ELSE                                         '
                        sCol = sCol + LEN(t2) - 3 - LEN(L(sLine).@LTxt) - IIF(RIGHT$(t2, 1) <> " ", 0, LEN(t2)) ' No, calc new scan start
                     END IF                                       '
                  END IF                                          '
               END IF                                             '
            '----- This is the Top/Bottom rollover setup
            ELSE                                                  ' Last one hit top/bottom
               sLine = IIF(cfNotFound = 1, LastLine, 1)           ' Set restart point
               cfNotFound = 0                                     ' Clear, we wrap just once
            END IF                                                '
         END IF                                                   '

         '----- Save for * operands etc.
         pfFind     = cfFind                                      ' Save now as a Previous set of parameters
         pfChange   = cfChange                                    '
         pfOChange  = cfOChange                                   '
         pfFLen     = cfFLen                                      '
         pfClen     = cfClen                                      '
         pfFlag     = cfFlag                                      '
         pfFLine    = cfFLine                                     '
         pfFCol     = cfFCol                                      '
         pfNotFound = cfNotFound                                  '

         '----- OK, finally we start the search
         IF sCol = 0 THEN sCol = 1                                '
         i = sLine                                                ' Init for DO loop
         j = sCol                                                 '
         DO WHILE 1 = 1                                           ' Do forever
            IF ISTRUE gfInterrupt THEN
               EXIT DO                                            ' Break for an FF search
            END IF

            '----- Watch for hitting Top/Bottom i.e. Not Found
            IF (IIF(RevOnce, -1 * sDir, sDir) = 1 AND IsLBottom(i)) OR _  ' Do till top or bottom
               (IIF(RevOnce, -1 * sDir, sDir) = -1 AND IsLTop(i)) THEN EXIT DO

            '----- Do the line range filtering
            IF IsCfRangePass THEN                                 ' A one time pass?
               CfClear(%CrtRangePass)                             ' Clear it
               Passcount = 2                                      ' Allow one more
            END IF                                                '
            IF PassCount > 0 THEN                                 ' Still a pass remaining?
               DECR PassCount                                     ' Yes, use it up

            ELSE                                                  ' else do the test
               '----- Do the line range filtering
               IF ISFALSE Me.RangeVal(i) OR _                     ' Do basic filtering
                  ISFALSE IsLData(i) THEN                         '
                  i += IIF(RevOnce, -1 * sDir, sDir)              ' Next entry
                  ONCE = %False                                   ' Ignore restart column
                  sCol = IIF(sDir = 1, 1, -1)                     ' Reset starting col number
                  ITERATE DO                                      ' Onward to next line
               END IF                                             '
            END IF                                                '

            '----- Do the NX line filtering
            IF IsCfNX THEN                                        ' Do the X test a little differently
               IF IsLInvisible(i) THEN                            ' Not Visible?
                  ONCE = %False                                   ' Skip this guy
                  i += IIF(RevOnce, -1 * sDir, sDir)              '
                  sCol = IIF(sDir = 1, 1, -1)                     ' Reset starting col number
                  ITERATE DO                                      '
               END IF                                             '
            END IF                                                '

            '----- Do the X line filtering
            IF IsCfX THEN                                         ' Do the X test a little differently
               IF ISFALSE IsLInvisible(i) AND ISFALSE IsLPopped(i) THEN ' Visible and not just Popped?
                  ONCE = %False                                   ' Skip this guy
                  i += IIF(RevOnce, -1 * sDir, sDir)              '
                  sCol = IIF(sDir = 1, 1, -1)                     ' Reset starting col number
                  ITERATE DO                                      '
               END IF                                             '
            END IF                                                '

            '----- Do the U line filtering
            IF IsCfU THEN                                         ' Do the U test
               IF ISFALSE IsLUser(i) THEN                         ' Not a User line?
                  ONCE = %False                                   ' Skip this guy
                  i += IIF(RevOnce, -1 * sDir, sDir)              '
                  sCol = IIF(sDir = 1, 1, -1)                     ' Reset starting col number
                  ITERATE DO                                      '
               END IF                                             '
            END IF                                                '

            '----- Do the NU line filtering
            IF IsCfNU THEN                                        ' Do the NU test
               IF IsLUser(i) THEN                                 ' Is this a User line?
                  ONCE = %False                                   ' Skip this guy
                  i += IIF(RevOnce, -1 * sDir, sDir)              '
                  sCol = IIF(sDir = 1, 1, -1)                     ' Reset starting col number
                  ITERATE DO                                      '
               END IF                                             '
            END IF                                                '

            '----- Do +/- processing for the source range
            IF DotSet AND DotFlag <> 0 THEN                       ' Possible + / - processing
               x = DotFlag                                        ' Get local copy
               IF BIT(x, %lCmdX) THEN                             ' Do the +/- processing for the Source
                  me.LFlagBitOn(i, %Invisible)                    ' If - make Invisible
                  gfXRebuild = %True                              ' Ask for Exclude processing
               ELSEIF BIT(x, %lCmdNX) THEN                        '
                  me.LFlagBitOff(i, %Invisible)                   ' If + make visible
                  gfXRebuild = %True                              ' Ask for Exclude processing
               END IF                                             '
            END IF                                                '

            '----- Stuff in command index as a -ve value (Looked at, not selected)
            IF ABS(LPCmdG(i)) <> PCmdIX THEN _                    ' If this command hasn't been here before (e.g. FIND ALL xxx)
               LPCmdS(i) = 0 - PCmdIX                             ' Stuff in PCmdIX (negated) to make it :ZNF

            '----- OK, we actually have to compare data to do this
            IF Srch THEN                                          ' Actually have to do string compare?

               '----- Prep the text line for proper CASE handling
               t1 = LTxtG(i)                                      ' Setup data from Txt line
               IF ISFALSE IsCfL1CaseComp AND _                    '
                  ISFALSE IsCfL1RegEx AND _                       '
                  ISFALSE IsCfL1Picture THEN t1 = UUCASE(t1)      ' Uppercase if not case sensitive search or picture type

               '----- Only for the first line we honor the starting column
               IF ONCE THEN                                       ' For 1st line, honour start col
                  ONCE = %False                                   ' Just once please

                  '----- Do Normal, Picture or RegEx search as appropriate using start column
                  IF IsCfL1Picture THEN                           ' Picture type search?
                     me.PicSearch(t1, t2, sCol, j, k)             ' Call search routine
                  ELSEIF IsCfL1RegEx THEN                         ' RegEx type search?
                     PCRE_Regex_Test(t1, sCol, j, k)              ' See if we can find it
                  ELSE                                            '
                     k = sCol                                     '
                     j = INSTR(sCol, t1, t2)                      ' OK, see if its there
                     k = LEN(t2)                                  '
                  END IF                                          '

               '----- Don't honor the start column for 2nd and subsequent lines
               ELSE                                               '
                  IF IIF(RevOnce, -1 * sDir, sDir) = -1 THEN      ' If LAST or PREV then start as the end of the line
                     IF IsCfL1Picture THEN                        ' Picture type search?
                        me.PicSearch(t1 , t2 , sCol , j, k)       ' Call search routine
                     ELSEIF IsCfL1RegEx THEN                      ' RegEx type search?
                        PCRE_Regex_Test(t1, 1, j, k)              ' See if we can find it
                     ELSE                                         '
                        j = INSTR(-1, t1, t2)                     '
                        k = LEN(t2)                               '
                     END IF                                       '
                  ELSE                                            '
                     IF IsCfL1Picture THEN                        ' Picture type search?
                        me.PicSearch(t1 , t2 , sCol , j, k)       ' Call search routine
                     ELSEIF IsCfL1RegEx THEN                      ' RegEx type search?
                        PCRE_Regex_Test(t1, 1, j, k)              ' See if we can find it
                     ELSE                                         '
                        j = INSTR(t1, t2)                         ' OK, see if its there
                        k = LEN(t2)                               '
                     END IF                                       '
                  END IF                                          '
               END IF                                             '

            '----- If no real search, make it look like we did one and found it
            ELSE                                                  ' If no Srch, fudge a 'found' answer
               j = 1: k = 1                                       ' Say found 1 byte in column 1
            END IF                                                '

            '----- OK if we found it, we have still a lot of work to do
            IF j THEN                                             ' If we have a 'found' condition

               IF Srch THEN                                       ' Doing a string search

                  '----- Do column bounds if asked for
                  IF IsCfFCol AND ISFALSE IsCfTCol THEN           ' Request for locate in single column?
                     IF j <> cfLMargin THEN                       ' then is this in the right column?
                        GOSUB KeepSearching                       ' No, keep looking
                        ITERATE DO                                '
                     END IF                                       '
                  ELSEIF IsCfFCol AND IsCfTCol THEN               ' Request within Left/Right columns?
                     IF j < cfLMargin OR j + k - 1 > cfRMargin THEN ' Is it within bounds?
                        GOSUB KeepSearching                       ' No, keep looking
                        ITERATE DO                                '
                     END IF                                       '
                  END IF                                          '

                  '----- Do the WORD test if asked for
                  IF IsCfWord THEN                                ' WORD requested?
                     IF j = 1 AND j + k - 1 = LEN(t1) THEN        ' If in left margin and only thing on line
                        GOSUB MarkTrue: MExitMeth                 ' We found it
                     ELSEIF j = 1 AND j + k - 1 <> LEN(t1) THEN   ' In left margin, more on line
                        IF INSTR(lclWord, MID$(t1, j + k, 1)) = 0 THEN  ' Trailing char is a delimiter
                           GOSUB MarkTrue: MExitMeth              ' We found it
                        ELSE                                      '
                           GOSUB KeepSearching                    '
                           ITERATE DO                             '
                        END IF                                    '
                     ELSEIF INSTR(lclWord, MID$(t1, j - 1, 1)) = 0 THEN ' Left end is after a DLM
                        IF j + k - 1 = LEN(t1) THEN               ' and right is at the end
                           GOSUB MarkTrue: MExitMeth              ' Another winner
                        ELSE                                      '
                           IF INSTR(lclWord, MID$(t1, j + k, 1)) = 0 THEN ' Trailing char is a delimiter
                              GOSUB MarkTrue: MExitMeth           ' Yet another winner
                           ELSE                                   '
                              GOSUB KeepSearching                 '
                              ITERATE DO                          '
                           END IF                                 '
                        END IF                                    '
                     ELSE                                         '
                        GOSUB KeepSearching                       '
                        ITERATE DO                                '
                     END IF                                       '
                  END IF                                          '

                  '----- Do the PREFIX test if asked for
                  IF IsCfPrefix THEN                              ' PREFIX requested?
                     IF j = 1 OR INSTR(lclWord, MID$(t1, j - 1, 1)) = 0 THEN ' If in left margin or left after a DLM
                        IF INSTR(lclWord, MID$(t1, j + k, 1)) <> 0 AND j + k - 1 <> LEN(t1) THEN ' Trailing char is a non-delimiter
                           GOSUB MarkTrue: MExitMeth              ' Another found it
                        ELSE                                      '
                           GOSUB KeepSearching                    '
                           ITERATE DO                             '
                        END IF                                    '
                     ELSE                                         '
                        GOSUB KeepSearching                       '
                        ITERATE DO                                '
                     END IF                                       '
                  END IF                                          '

                  '----- Do the SUFFIX test if asked for
                  IF IsCfSuffix THEN                              ' SUFFIX requested?
                     IF j = 1 OR INSTR(lclWord, MID$(t1, j - 1, 1)) = 0 THEN ' Can't be found in col 1 or preceeded by a DLM
                        GOSUB KeepSearching                       '
                        ITERATE DO                                '
                     ELSEIF INSTR(lclWord, MID$(t1, j + k, 1)) = 0 OR J + k -1 = LEN(t1) THEN ' Trailing char is delimiter or EOL
                        GOSUB MarkTrue: MExitMeth                 ' We found it
                     ELSE                                         '
                        GOSUB KeepSearching                       '
                        ITERATE DO                                '
                     END IF                                       '
                  END IF                                          '

                  '----- Do the LM test if asked for
                  IF IsCfLM THEN                                  ' LM requested?
                     IF j <> Prf.BndLeft THEN                     ' OK if in LM column
                        GOSUB KeepSearching                       '
                        ITERATE DO                                '
                     END IF                                       '
                  END IF                                          '

                  '----- Do the RM test if asked for
                  IF IsCfRM THEN                                  ' RM requested?
                     lclRBnd = LEN(L(i).@LTxt)                    ' Set where RM is to be found
                     IF j + k - 1 = lclRBnd THEN                  ' OK if in RM column
                        IF Prf.BndLeft > 1 OR Prf.BndRight > 0 THEN ' BOUNDS set?
                           IF lclRBnd > Prf.BndRight OR lclRBnd < Prf.BndLeft THEN ' Outside the LM / RM BNDS columns?
                              GOSUB KeepSearching                 ' Then fail it
                              ITERATE DO                          '
                           END IF                                 '
                        END IF                                    '
                     ELSE                                         ' Not at RM
                        GOSUB KeepSearching                       ' Then fail it
                        ITERATE DO                                '
                     END IF                                       '
                  END IF                                          '

                  '----- Do the Color tests if asked for
                  IF IsCfSTD OR cfHiLiteClr <> 0 OR _             ' Any color type requests?
                     IsCfSOLID OR IsCfMSOLID THEN                 '

                     c1 = me.AttrHiLiteGet(i, j, j + k - 1)       ' Get the string

                     IF IsCfStd AND c1 = REPEAT$(k, CHR$(0)) THEN GOTO ClrTestDone  ' Lets try STD first
                     IF c1 = REPEAT$(k, CHR$(cfHiLiteClr)) THEN GOTO ClrTestDone    ' Lets try a simple color test
                     IF IsCfSOLID THEN                            ' Lets try SOLID next
                        IF c1 = REPEAT$(k, LEFT$(c1, 1)) AND _    ' Attr line field all the same color?
                           c1 <> REPEAT$(k, CHR$(0)) THEN GOTO ClrTestDone  ' and not STD
                     END IF                                       '
                     IF IsCfMSOLID THEN                           ' Lets try -SOLID next
                        IF c1 <> REPEAT$(k, LEFT$(c1, 1)) THEN GOTO ClrTestDone     ' Attr line field not all the same color?
                     END IF                                       '

                     GOSUB KeepSearching                          '
                     ITERATE DO                                   '
                  END IF                                          '

                  ClrTestDone:
               END IF                                             '

               '----- Whoopee! an actual successful search, do the MarkTrue stuff and exit
               GOSUB MarkTrue: MExitMeth                          ' No Specials (Word, Prefix, Suffix), we're done
            END IF                                                '

            '----- Failed the search, onward to the next line
            i += IIF(RevOnce, -1 * sDir, sDir)                    '
            sCol = IIF(IIF(RevOnce, -1 * sDir, sDir) = -1, -1, 1) ' Adjust start column
         LOOP                                                     '
         MExitMeth

         '----- This does all the stuff needed for a True search condition
         MarkTrue:
            PosType = IIF(IsCfTop, %FLocate, %Find)               ' Set positioning for TOP or not

            '----- Handle the weird LEFT/RIGHT considerations
            IF IsCfLeft AND sDir = -1 THEN                        ' LEFT and doing a backward scan
               IF IsCfL1Picture THEN                              ' Picture type search?
                  me.PicSearch(t1 , t2 , 1 , j, k)                ' Search foreward from column 1
               ELSE                                               '
                  j = INSTR(1, t1, t2)                            ' Search foreward from column 1
                  k = LEN(t2)                                     '
               END IF                                             '
            END IF                                                '

            IF IsCfRight AND sDir = 1 THEN                        ' RIGHT and doing a foreward scan
               IF IsCfL1Picture THEN                              ' Picture type search?
                  me.PicSearch(t1 , t2 , -1 , j, k)               ' Search backward from the end
               ELSE                                               '
                  j = INSTR(-1, t1, t2)                           ' Search backward from the end
                  k = LEN(t2)                                     '
               END IF                                             '
            END IF

            ' OK, finally we can consider it a Found condition
            METHOD = %True                                        '
            INCR FixCtr                                           ' Bump FIXCtr
            LFIXS(i) = FixCtr                                     ' Save in line control area
            LPCmdS(i) = PCmdIX                                    ' Stuff in PCmdIX to make it :ZF

            '----- If a normal non-X'd line see if that should be altered
            IF ISFALSE IsLInvisible(i) THEN                       ' A normal non-X'd line?
               IF IsCfMX THEN                                     ' If caller is someone who wants X out the line
                  me.LFlagBitOn(i, %Invisible)                    ' Just make it invisible
                  me.LFlagBitOn(i, %Popped)                       ' Say we've just popped it
                  gfXRebuild = %True                              ' Call for rebuild exclude state
                  me.CurSetReq(PosType, i, j, %True, %True)       ' Put cursor on the X line, no hilite, XPtr to be set
                  GOSUB SeeSetClr                                 ' Go see if Pen colors to be set
                  CsrLinDX = VAL(LLNumG(i))                       ' Save hidden line number for Status Bar display
               ELSE                                               ' This is someone who is just looking like FIND / CHANGE'
                  t1 = LSET$(CurrPCmd, 8)                         ' Make a temp fixed length field
                  IF INSTR("FIND    F       FF      RFIND   ", t1) <> 0 THEN ' A FIND command?
                     GOSUB SeeSetClr                              ' Go see if Pen colors to be set
                     IF Srch THEN                                 ' Was this a string search?
                        me.CurSetReq(PosType, i, j, %True)        ' Set cursor
                        IF ISTRUE Prf.HiFind THEN                 ' See whether to hilite the find string
                           me.AttrInvSet(i, j, j + k - 1)         ' Mark for Hi-Lite
                        END IF                                    '
                     END IF                                       '
                  END IF                                          '
               END IF                                             '

            '----- If an X'd line see if that should be altered
            ELSE                                                  ' We have a line already X'd out
               IF IsCfMX THEN                                     ' If caller is someone who wants X out the line
                  IF Srch THEN                                    ' A real string search?
                     me.CurSetReq(PosType, i, j, %True, %True)    ' Just set the cursor
                     GOSUB SeeSetClr                              ' Go see if Pen colors to be set
                     CsrLinDX = VAL(LLNumG(i))                    ' Save hidden line number for Status Bar display
                  END IF                                          '
               ELSE                                               '
                  IF IsCfDX THEN                                  ' If told to leave X status alone
                     t1 = LSET$(CurrPCmd, 8)                      ' Make a temp fixed length field
                     IF INSTR("CHANGE  C       CHA     CHG     RCHANGE ", t1) <> 0 THEN ' A CHANGE command?
                        ' Do nothing, CHANGE should set lilighting
                     ELSE                                         '
                        IF Srch THEN                              ' If a real string search
                           me.CurSetReq(PosType, i, j, %True, %True) ' Put cursor on the X line, no hilite, XPtr to be set
                           GOSUB SeeSetClr                        ' Go see if Pen colors to be set
                           CsrLinDX = VAL(LLNumG(i))              ' Save hidden line number for Status Bar display
                           ii = CsrLinDX                          '
                        END IF                                    '
                     END IF                                       '
                  ELSE                                            '
                     me.LFlagBitOff(i, %Invisible)                ' Remove the invisibility flag
                     me.LFlagBitOn(i, %Popped)                    ' Say we've just popped it
                     LWrk1S(i) = 0                                ' Clear ptr back to the Xclude line
                     gfXRebuild = %True                           ' Rebuild exclude state
                     IF Srch THEN                                 ' If a real string search
                        GOSUB SeeSetClr                           ' Go see if Pen colors to be set
                        me.CurSetReq(PosType, i, j, %True)        ' Set Cursor
                        IF ISTRUE Prf.HiFind THEN                 ' See whether to hilite the find string
                           me.AttrInvSet(i, j, j + k - 1)         ' Set for Hi-Lite
                        END IF                                    '
                     END IF                                       '
                  END IF                                          '
               END IF                                             '
            END IF                                                '

            '----- Save the 'found' location for others
            sLine = i: sCol = j                                   ' Save line and column where found
            IF ISFALSE Srch THEN sLine += 1                       ' If not really searching, continue on next line
            cfFLine = i: cfFCol = j                               ' Save in global areas
            IF IsCfL1RegEx THEN                                   '
               cfFLen = k                                         ' If Regex use found length
            END IF                                                '
         RETURN                                                   '

         '----- We failed, set things up to continue the search
         KeepSearching:                                           '
            IF IIF(RevOnce, -1 * sDir, sDir) = 1 THEN             ' Forward
               sCol = j + 1                                       ' This isn't the one, set to continue at next column
               IF sCol > LEN(t1) THEN                             ' End of record?
                  sCol = 1                                        ' Yes, start back at 1
                  ONCE = %False                                   '
                  INCR i                                          '
               ELSE                                               '
                  ONCE = %True                                    ' Say to honour column, stay on same line
               END IF                                             '
            ELSE                                                  '
               IF IsCfL1RegEx THEN                                ' RegEx type search?
                  ONCE = %True                                    ' Stay on same line
                  sCol += 1                                       ' Step over this found one
                  RETURN                                          '
               END IF                                             '
               sCol = j - LEN(t1) - 3                             ' Backward, adjust the other way
               l = sCol                                           '
               IF ABS(sCol) > LEN(t1) THEN                        ' Beginning of record?
                  sCol = -1                                       ' Yes, start back at end of record
                  ONCE = %False                                   '
                  DECR i                                          ' and next(prev) line
               ELSE                                               '
                  ONCE = %True                                    ' No, Honour column, stay on same line
               END IF                                             '
            END IF                                                '
         RETURN                                                   ' We'll never get here

      SeeSetClr:                                                  ' See if Pen colors to be set
         IF cfHiLiteOn <> 0 THEN
            PClr = cfHiLiteOn                                     ' Override if requested
         ELSEIF IsCfPStd THEN
            PClr = %AttrHiSTD                                     ' +Std asked for?
         ELSE
            RETURN
         END IF
         me.AttrHiLiteSet(i, j, j + k - 1, PClr)                  ' Set the color
         RETURN                                                   '

      END METHOD

      METHOD SetNewStart()
      '---------- Set a new .START label at the last line
      LOCAL k AS LONG
         MEntry
         IF Prf.Start <> "NEW" THEN MExitMeth                     ' Do only if START PRIOR is set
         IF LastLine < 3 THEN MExitMeth                           ' Skip also if no lines
         k = me.LLCtlScan(".START")                               ' See if label exists already
         IF k THEN LLblS(k - 1) = $BlankLNo: LLCtlS(k - 1) = $BlankLNo ' Blank label elsewhere if used
         LLblS(LastLine - 1) = ".START"                           ' Add the Label
         me.UpdLControl(LastLine - 1)                             ' Put back the line number
         MExit
      END METHOD

      METHOD SetPrfNumType(v AS STRING)
      '---------- Set the Prf NUMTYPE string
         MEntry
         Prf.NumType = v                                          ' Stuff in the value
         MExit
      END METHOD

      METHOD SetStart()
      '---------- Do initial positioning based on the Profile START operand
      LOCAL i, j AS LONG
         MEntry
         SELECT CASE AS CONST$ Prf.Start                          ' What is our default
            CASE "FIRST"                                          ' FIRST
               Topscrn = IIF(Prf.Pageflag, 2, 1)                  ' Set to 1 or 2 based on PAGE value

            CASE "LAST"                                           ' LAST
               TopScrn = LastLine                                 ' Set to bottom - 1 page
               TopScrn = me.CRPBack(%mVisible, TopScrn, gwScrHeight - 3 - Prf.Cols)
            CASE "LABEL", "NEW"                                   ' Both these position to .START
               i = me.LineNoRef(".START")                         ' Get the line number
               IF i = -1 THEN                                     ' Found?
                  i = 1                                           ' No, go to top
               ELSE                                               '
                  me.LFlagBitOff(i, %Invisible)                   ' Remove the invisibility flag
                  gfXRebuild = %True                              ' Rebuild exclude state
               END IF                                             '
               TopScrn = i                                        '
         END SELECT                                               '
         MExit
      END METHOD

      METHOD SetSubst() AS LONG
      '---------- See if SET substitution is needed
      LOCAL i, j AS LONG, t, u, v, cmd, ncmd, pcmd AS STRING
         MEntry
         METHOD = %False
         pcmd = pCommand + " "                                    ' Get a local copy
         i = INSTR(pcmd, "=")                                     ' Look for '='
         IF i = 0 THEN MExitMeth                                  ' No '=', then nothing to do
         DO WHILE ISNOTNULL(TRIM$(pcmd))                          ' Process it
            cmd = sGetWord(pcmd, %Strip, %QuoteSig) + " "         ' Get an Operand
            i = INSTR(cmd, "=")                                   ' Look for '='
            IF INSTR("'`" + $DQ, LEFT$(cmd, 1)) THEN i = 0        ' If a quoted string, kill any = location
            IF i = 1 THEN                                         ' Got '=', and not quoted
               j = INSTR(cmd, " ")                                ' Get trailing blank
               IF j <> 2 THEN                                     ' If not a Solitary = sign?
                  t = LEFT$(cmd, j - i)                           ' extract =xxx key
                  IF IsNE(t, "=X") THEN                           ' If not =X
                     u = MID$(t, 2)                               ' Extract xxx
                     v = sSetTable("GET", u)                      ' Retrieve substitution value
                     IF VAL(LEFT$(v, 1)) > 0 THEN                 ' Not found?
                        i = INSTR(j, cmd, "=")                    ' Ignore it, look onward
                        METHOD = %True                            '
                        scError(%eFail, MID$(v, 2) + ": " + t): MExitMeth ' Else tell user
                     ELSE                                         ' We have a value
                        cmd = MID$(v, 2) + " "                    ' Substitute it
                     END IF                                       '
                  END IF                                          '
               END IF                                             '
            END IF                                                '
            ncmd += cmd                                           ' this operand done
         LOOP                                                     '
         pCommand = ncmd                                          ' Save replacement back
         MExit
      END METHOD

      METHOD SetupFMSBXref()
         '----- Re-Build the SBTable using ENV.SBLayout
      LOCAL i, j, k AS LONG, t AS STRING
      DIM SBParse(1 TO 15) AS STRING
         FOR i = 1 TO 15                                          '
            SBTable(i) = gSBTable(i)                              ' Start by Copying the global table
         NEXT i                                                   '
         SBCount = 0                                              ' Reset since we're starting over
         t = ENV.SBLayout + ",PAD"                                ' Ensure PAD at the end
         IF IsFMTab THEN t = "MODE,LINNO,INSOVR,CASE,MISC,PAD"    ' FM gets a fixed set of fields
         k = PARSECOUNT(t)                                        ' Get number of entries
         PARSE t, SBParse()                                       ' Break it into pieces
         FOR i = 1 TO k                                           ' Process each entry
            t = LSET$(SBParse(i), 6)                              ' Make name 6 chars
            FOR j = 1 TO 15                                       ' Find name in table
               IF t = SBTable(j).SBName THEN                      ' Found it
                  SBTable(j).SBActive = "Y"                       ' Mark this one active
                  INCR SBCount                                    ' Bump our TP count
                  SBXref(SBCount) = j                             ' Put the entry index into the Xref table
                  EXIT FOR                                        ' We're done
               END IF                                             '
            NEXT j                                                '
         NEXT i                                                   ' Onward to next ENV KW
      END METHOD                                                  '

      METHOD sGet (BYVAL ix AS LONG) AS LONG: METHOD = s(ix): END METHOD

      METHOD SPLIT()
      '---------- Split a line for SPLIT/RSPLIT
      LOCAL sp1, sp2, i, j, k, l, x, y, bias AS LONG
      LOCAL ol, nl, TTxt, lclChg, t AS STRING
      LOCAL oc, nc, cTxt, CWrk AS WSTRING, AttrAsc AS WORD

         MEntry
         '----- Get common variables into local data
         i = cfFLine                                              ' Find line
         j = cfFCol                                               ' Column where it was found
         k = cfFLen                                               ' Length of found string (Use orig. length)
         l = LEN(L(i).@LTxt)                                      ' Length of text line
         GOSUB GetPrevAttr                                         ' Get any Attr of the found string

         '----- Prepare the Change string
         sp1 = cfSplitPt2 - 1                                     ' Calc the Split point in the source literal (Use Lit-2 instead)
         sp2 = cfSplitPt2 - 1                                     ' Calc the Split point in the replace literal
         y = j                                                    ' Found data
         TTxt = LTxtG(i)                                          ' Copy of the line's text
         FOR x = 1 TO LEN(cfChange)                               ' Loop building the change string
            SELECT CASE AS CONST$ MID$(cfChange, x, 1)            ' Do it based on the char inside P'xxx'
               CASE "=": lclChg += MID$(TTxt, y, 1): INCR y       ' Copy the character
               CASE "<": lclChg += LLCASE(MID$(TTxt, y, 1)): INCR y ' Lowercase it
               CASE ">": lclChg += UUCASE(MID$(TTxt, y, 1)): INCR y ' Uppercase it
               CASE "!"                                           ' Copy found string
                  lclChg += MID$(TTxt, j, k)                      '
                  IF x + bias <= sp2 THEN                         '
                     sp2 += (k - 1)                               ' Shift split point by insert
                     bias += (k - 1)                              '
                  END IF                                          '
               CASE "~": INCR y                                   ' Ignore it
               CASE "\"                                           ' Escape?
                  lclChg += MID$(cfChange, x + 1, 1)              ' Copy next char
                  INCR x                                          ' extra step
                  IF IsCfL2Picture THEN INCR y                    ' If doing a picture
               '----- Not special character, just copy it
               CASE ELSE                                          ' Copy L12 character
                  lclChg += MID$(cfChange, x, 1)                  '
                  IF IsCfL2Picture THEN INCR y                    ' If doing a picture
            END SELECT                                            '
         NEXT x                                                   '

         me.LFlagBitOff(i, %EQChange)                             ' Reset ==CHG>
         me.ModSet(i)                                             ' Remember we changed something

         ol = LEFT$(LTxtG(i), j - 1) + LEFT$(lclChg, sp2)         ' Split line into left + lh of change literal
         nl = MID$(lclChg, sp2 + 1) + MID$(TTxt, j + k)           ' and right = rh of chg literal + line remainder

         oc = LAttrG(i)                                           ' Get Attr line
         nc = oc                                                  ' Copy to nc before we alter it
         oc = LEFT$(oc, j - 1) + REPEAT$(LEN(lclChg) - sp2, CHR$$(AttrAsc))  ' Split color into left + lh of change literal
         nc = REPEAT$(LEN(lclChg) - sp2, CHR$$(AttrAsc)) + MID$(nc, j + k)  ' and right = rh of change literal + line remainder
         me.LInsertEmpty(i, 1, %Data)                             ' Insert the split data line

         me.LTxtSet(i, ol)                                        ' Swap in left hand data
         me.LTxtSet(i + 1, nl)                                    ' Swap in right hand data
         me.LFlagBitOff(i, %EQChange)                             ' Reset ==CHG>
         LTagS(i + 1) = TRIM$(LTagG(i))                           ' Copy the :tag
         me.UpdLControl(i + 1)                                    '

         me.AdjustPending(i, 1,0)                                 ' Adjust pending requests
         DotEnd += 1                                              ' Adjust range

         LAttrS(i) = oc                                           '
         LAttrS(i + 1) = nc                                       '
         me.AttrScan(i)                                           ' Recolorize
         me.AttrScan(i + 1)                                       ' Recolorize

         IF IsCfNext THEN                                         ' If FWD
            sLine += 1: sCol = 1 + LEN(MID$(lclChg, sp2 + 2))     ' Then start of new line + L'changestring
         ELSE                                                     '
            sLine += 1: sCol = 1                                  ' Else start of new line
         END IF                                                   '
         me.CurSetReq(%Position, sLine, sCol, %False)             ' Set cursor set attempt
         cfNotFound = 3                                           ' Set to srch needed

         IF IsCfAll THEN                                          ' An ALL request?
            CfSet(%CrtRangePass)                                  ' Issue a one time pass for the range test
         END IF                                                   '
         MExitMeth

      GetPrevAttr:
         '----- First see if any existing highlighting on the found string
         AttrAsc = 0                                              '
         cTxt = LAttrG(i)                                         ' Get Attr line
         CWrk = MID$(ctxt, j, k)                                  ' Get data for the found field
         IF CWrk = REPEAT$(k, LEFT$(CWrk, 1)) THEN                ' All the same?
            AttrAsc = ASC(LEFT$(Cwrk, 1))                         ' Set AttrAsc to the 1st char
         END IF                                                   '
         RETURN                                                   '
      END METHOD

      METHOD sSearch (BYVAL lno AS LONG) AS LONG
      LOCAL i AS LONG
         MEntry
         IF lno < s(3 + Prf.Cols) THEN METHOD = -1: MExitMeth     ' If above screen, return -1
         FOR i = 3 + Prf.Cols TO 300                              ' See if ON the screen
            IF s(i) = lno THEN METHOD = i: MExitMeth              ' If so, return screen line number
            IF s(i) = 0 THEN EXIT FOR                             ' Not above or ON the screen, must be below
         NEXT i                                                   '
         METHOD = -2                                              ' Return -2
         MExit
      END METHOD

      METHOD ScreenDim (height AS LONG, lwidth AS LONG)
         MEntry
         LOCAL i AS LONG
         IF height > UBOUND(Scr) THEN                             '
            REDIM Scr(height) AS INSTANCE STRING                  ' Resize bigger if needed
         END IF                                                   '
         FOR i = 1 TO height                                      ' Fill it
            Scr(i) = REPEAT$(lwidth, " ")                         '
         NEXT i
         MExit
      END METHOD

      METHOD ScreenChrHit(slin AS LONG, scol AS LONG) AS LONG     ' non-blank or adjacent to one?
         LOCAL lclCol, rcol AS LONG
         lclCol = MAX(1, scol - 1): rcol = MIN(scol + 1, ENV.ScrWidth) ' Set boundaries
         IF MID$(Scr(slin), lclCol TO rcol) = SPACE$(rcol - lclCol + 1) THEN EXIT METHOD
         METHOD = %True                                           ' Something non-blank, return true
      END METHOD                                                  '

      METHOD ScreenGet (i AS LONG) AS STRING: METHOD = Scr(i): END METHOD
      METHOD ScreenRep (i AS LONG, CCol AS LONG, str AS STRING): MID$(Scr(i), CCol, LEN(str)) = str: END METHOD

      METHOD ProfScrlRep (posit AS LONG, char AS STRING)
      LOCAL t AS STRING
         t = Prf.Scroll
         IF LEN(t) < posit THEN t = LEFT$(t + "    ", posit)
         MID$(t, posit, 1) = char
         Prf.Scroll = t
      END METHOD

      METHOD StateLoad(fn AS STRING, fromline AS LONG, ToLine AS LONG, IsCopy AS LONG)
      '---------- If valid, re-load the STATE information
      LOCAL RealLines, LblCount, TagCount, UsrCount, ClrCount, XCount, NoteCount, UseNewHash, AttrClr AS LONG
      LOCAL OHash, Hash AS DWORD
      LOCAL H1, HL, HT, HU, HC, HX, HN, t, lclFn AS STRING
      REGISTER i AS LONG
      REGISTER j AS LONG
      LOCAL k, l, a, lno, bias, missing, duplabel, tcol, tlen AS LONG, tt, tc, llabel, lcltag, tclr, lclr, lhash AS STRING
      LOCAL SIO AS iIO                                            ' For our I/O stuff
         MEntry
         OffStateExist                                            ' Remove STATE data exists
         LET SIO = CLASS "cIO"                                    '

         '----- Make the fn into a 'safe' fn
         lclFn = fn + ".STATE"                                    ' Temp copy with .STATE on the end
         REPLACE ANY ":\/" WITH "```" IN lclFN                    ' Make : / and \ into `
         lclFn = ENV.StatePath + lclFn                            ' Add our STATE folder

         IF ISFALSE ISFILE(lclFn) THEN MExitMeth                  ' Exit if no STATE file
         IF sIsNullFile(lclFn) THEN MExitMeth                     ' Also exit if the nullfile version

         '----- See if passed reasonable data
         SIO.Setup("IE", "", "", lclFn)                           ' Tell IO what we're opening
         IF SIO.EXEC THEN _                                       ' Open it
            MErrExit(%eFail, "Warning, STATE file could not be opened, STATE ignored")
         LINE INPUT # SIO.FNum, H1                                ' Get a line
         IF LEFT$(H1, 2) <> "#1" THEN                             ' Better be the header line
            SIO.Close                                             '
            MErrExit(%eFail, "STATE file is missing it's header line, STATE ignored")
         ELSE                                                     '
            H1 = MID$(H1, 3)                                      ' Strip off the #1
         END IF                                                   '

         '----- Check Hash and line counts
         sFCS32Init(Hash)                                         ' Initialize the Hash count
         FOR i = Fromline TO ToLine                               ' Loop through the line range

            '----- Do Data lines
            IF IsLData(i) THEN                                    ' Data lines
               INCR RealLines                                     ' Count real text lines
               GOSUB GetHashText                                  ' Get hashable text
               Hash = sFCS32Update(Hash, BYVAL STRPTR(t), LEN(t)) ' Accum the hash
               oHash = sHash(t, oHash)                            ' Do it the old way as well
            END IF                                                '
         NEXT i                                                   '
         sFCS32Final(Hash)                                        ' Finalize Hash

         IF RealLines <> VAL(PARSE$(H1, ",", 2)) THEN _           ' Correct # lines?
            SIO.Close: MErrExit(%eFail, "STATE file line count does not match the file, STATE ignored")

         '----- See which Hash count was used
         IF HEX$(Hash, 8) = PARSE$(H1, ",", 1) THEN               ' Match the New hash?
            UseNewHash = %True                                    ' Remember that
         ELSEIF HEX$(OHash, 8) <> PARSE$(H1, ",", 1) THEN         ' Not the Old hash either?
            scError(%eFail, "STATE file hash count does not match the file, Attempting STATE regardless") ' Else tell user
         END IF                                                   '

         TopScrn = VAL(PARSE$(H1, ",", 3))                        ' Set top of screen line

         '----- Do Labels if present
         LblCount = VAL(PARSE$(H1, ",", 4))                       ' Get label count
         IF LblCount > 0 THEN                                     ' We have some to do
            LINE INPUT # SIO.FNum, HL                             ' Get the label line
            IF LEFT$(HL, 2) <> "#L" THEN                          ' Better be the label line
               SIO.Close                                          '
               MErrExit(%eFail, "STATE file label line is misplaced, STATE terminating.")
            ELSE                                                  '
               HL = MID$(HL, 3)                                   ' Strip off the #L
            END IF                                                '
            j = PARSECOUNT(HL) - 1                                ' How many on the line
            IF j <> LblCount THEN _                               ' If not a match
               scError(%eFail, "STATE file has " + FORMAT$(j) + " labels, Header says " + FORMAT$(LblCount) + " labels, Attempting STATE regardless") ' Else tell user
            FOR k = 1 TO PARSECOUNT(HL) - 1                       ' Loop through the entries
               t = PARSE$(HL, ",", k)                             ' Get next string
               llabel = PARSE$(t, "/", 1)                         ' Get line label
               IF ISFALSE IsCopy THEN                             ' and NOT COPY?
                  i = VAL(PARSE$(t, "/", 2)) + FromLine - 1       ' Get line no
                  lhash = PARSE$(t, "/", 3)                       ' Get line hash
                  IF i > LastLine - 1 THEN                        ' Avoid out of range errors
                     missing = %True                              ' A label failed to be reset
                     ITERATE                                      ' Next For loop item
                  END IF                                          '
                  GOSUB GetHashText                               ' Get hashable text
                  GOSUB GetLineHash                               ' Get hash for this line
                  tt = HEX$(Hash, 8)                              ' Get hash for the pointed to line
                  IF tt = lhash THEN                              ' The simple case, the text matches
                     llabel = UUCASE(llabel)                      ' Uppercase it
                     FOR l = 1 TO LastLine                        ' See if label already exists
                        IF TRIM$(L(l).LLbl) = llabel THEN Duplabel = %True: RETURN ' A duplicate? Remember that
                     NEXT k                                       '
                     LLblS(I) = llabel                            ' Then Re-insert the label
                     me.UpdLControl(i)                            ' Update LLCtl
                     ITERATE FOR                                  ' Done this one
                  ELSE                                            ' no match, lets search a bit
                     scError(%eFail, "STATE failed label line hash verification, Attempting to continue")
                     missing = %True                              ' A label failed to be reset
                  END IF                                          '
               END IF                                             '
            NEXT K                                                '
         END IF

         '----- Do Tags if present
         TagCount = VAL(PARSE$(H1, ",", 5))                       ' Get Tag count
         IF TagCount > 0 THEN                                     ' We have some to do
            LINE INPUT # SIO.FNum, HT                             ' Get the tag line
            IF LEFT$(HT, 2) <> "#T" THEN                          ' Better be the tag line
               SIO.Close                                          '
               MErrExit(%eFail, "STATE file tag line is misplaced, STATE terminating.")
            ELSE                                                  '
               HT = MID$(HT, 3)                                   ' Strip off the #T
            END IF                                                '
            j = PARSECOUNT(HT) - 1                                ' How many on the line
            IF j <> TagCount THEN _                               ' If not a match
               scError(%eFail, "STATE file has " + FORMAT$(j) + " tags, Header says " + FORMAT$(TagCount) + " tags, Attempting STATE regardless") ' Else tell user
            FOR k = 1 TO PARSECOUNT(HT) - 1                       ' Loop through the entries
               t = PARSE$(HT, ",", k)                             ' Get next string
               lcltag = PARSE$(t, "/", 1)                         ' Get line tag
               IF ISFALSE IsCopy THEN                             ' and NOT COPY?
                  i = VAL(PARSE$(t, "/", 2)) + FromLine - 1       ' Get line no
                  lhash = PARSE$(t, "/", 3)                       ' Get line hash
                  IF i > LastLine - 1 THEN                        ' Avoid out of range errors
                     missing = %True                              ' A label failed to be reset
                     ITERATE                                      ' Next For loop item
                  END IF                                          '
                  GOSUB GetHashText                               ' Get hashable text
                  GOSUB GetLineHash                               ' Get hash for this line
                  tt = HEX$(Hash, 8)                              ' Get hash for the pointed to line
                  IF tt = lhash THEN                              ' The simple case, the text matches
                     lcltag = UUCASE(lcltag)                      ' Uppercase it
                     LTagS(I) = lcltag                            ' Then Re-insert the tag
                     me.UpdLControl(i)                            ' Update LLCtl
                     ITERATE FOR                                  ' Done this one
                  ELSE                                            ' no match, lets search a bit
                     scError(%eFail, "STATE failed tag line hash verification, Attempting to continue")
                     missing = %True                              ' A label failed to be reset
                  END IF                                          '
               END IF                                             '
            NEXT K                                                '
         END IF                                                   '

         '----- Do User flags if present
         UsrCount = VAL(PARSE$(H1, ",", 6))                       ' Get Usr count
         IF UsrCount > 0 THEN                                     ' We have some to do
            LINE INPUT # SIO.FNum, HU                             ' Get the user line
            IF LEFT$(HU, 2) <> "#U" THEN                          ' Better be the Usr line
               SIO.Close                                          '
               MErrExit(%eFail, "STATE file User line is misplaced, STATE terminating.")
            ELSE                                                  '
               HU = MID$(HU, 3)                                   ' Strip off the #U
            END IF                                                '
            j = PARSECOUNT(HU) - 1                                ' How many on the line
            IF j <> UsrCount THEN _                               ' If not a match
               scError(%eFail, "STATE file has " + FORMAT$(j) + " User flags, Header says " + FORMAT$(UsrCount) + " flags, Attempting STATE regardless") ' Else tell user
            FOR k = 1 TO PARSECOUNT(HU) - 1                       ' Loop through the entries
               t = PARSE$(HU, ",", k)                             ' Get next string
               i = VAL(PARSE$(t, "/", 1)) + FromLine - 1          ' Get line number
               lhash = PARSE$(t, "/", 2)                          ' Get line hash
               IF i > LastLine - 1 THEN                           ' Avoid out of range errors
                  missing = %True                                 ' A flag failed to be reset
                  ITERATE                                         ' Next For loop item
               END IF                                             '
               GOSUB GetHashText                                  ' Get hashable text
               GOSUB GetLineHash                                  ' Get hash for this line
               tt = HEX$(Hash, 8)                                 ' Get hash for the pointed to line
               IF tt = lhash THEN                                 ' The simple case, the text matches
                  me.LFlagBitOn(i, %User)                         ' Then Re-insert the flag
                  ITERATE FOR                                     ' Done this one
               ELSE                                               ' no match
                  scError(%eFail, "STATE failed User line hash verification, Attempting to continue")
                  missing = %True                                 ' A flag failed to be reset
               END IF                                             '
            NEXT K                                                '
         END IF                                                   '

         '----- Do Attr if present
         ClrCount = VAL(PARSE$(H1, ",", 7))                       ' Get Attr count
         IF ClrCount > 0 THEN                                     ' We have some to do
            LINE INPUT # SIO.FNum, HC                             ' Get the Attr line
            IF LEFT$(HC, 2) <> "#C" THEN                          ' Better be the Attr line
               SIO.Close                                          '
               MErrExit(%eFail, "STATE file color line is misplaced, STATE terminating.")
            ELSE                                                  '
               HC = MID$(HC, 3)                                   ' Strip off the #C
            END IF                                                '
            j = PARSECOUNT(HC) - 1                                ' How many on the line
            IF j <> ClrCount THEN _                               ' If not a match
               scError(%eFail, "STATE file has " + FORMAT$(j) + " color blocks, Header says " + FORMAT$(ClrCount) + " blocks, Attempting STATE regardless") ' Else tell user
            FOR k = 1 TO PARSECOUNT(HC) - 1                       ' Loop through the entries
               t = PARSE$(HC, ",", k)                             ' Get next string
               i = VAL(PARSE$(t, "/", 1)) + FromLine - 1          ' Get line no
               tclr = PARSE$(t, "/", 2)                           ' Get color value
               tcol = VAL(PARSE$(t, "/", 3))                      ' Get column value
               tlen = VAL(PARSE$(t, "/", 4))                      ' Get length value
               lhash = PARSE$(t, "/", 5)                          ' Get line hash
               GOSUB GetHashText                                  ' Get hashable text
               GOSUB GetLineHash                                  ' Get hash for this line
               tt = HEX$(Hash, 8)                                 ' Get hash for the pointed to line
               IF tt = lhash THEN                                 ' The simple case, the text matches
                  a = INSTR(nHiLitesChrs, tclr)                   ' Get index of the color char
                  me.AttrHiLiteSet(i, tcol, tcol + tlen - 1, a)   '
                  ITERATE FOR                                     ' Done this one
               ELSE                                               ' no match
                  scError(%eFail, "STATE failed color line hash verification, Attempting to continue")
                  missing = %True                                 ' A label failed to be reset
               END IF                                             '
            NEXT K                                                '
         END IF                                                   '

         '----- Do Exclude if present
         XCount = VAL(PARSE$(H1, ",", 8))                         ' Get X count
         IF XCount > 0 THEN                                       ' We have some to do
            LINE INPUT # SIO.FNum, HX                             ' Get the X line
            IF LEFT$(HX, 2) <> "#X" THEN                          ' Better be the X line
               SIO.Close                                          '
               MErrExit(%eFail, "STATE file Exclude line is misplaced, STATE terminating.")
            ELSE                                                  '
               HX = MID$(HX, 3)                                   ' Strip off the #X
            END IF                                                '
            j = PARSECOUNT(HX) - 1                                ' How many on the line
            IF j <> XCount THEN _                                 ' If not a match
               scError(%eFail, "STATE file has " + FORMAT$(j) + " exclude blocks, Header says " + FORMAT$(XCount) + " blocks, Attempting STATE regardless") ' Else tell user
            FOR k = 1 TO PARSECOUNT(HX) - 1                       ' Loop through the entries
               t = PARSE$(HX, ",", k)                             ' Get next string
               i = VAL(PARSE$(t, "/", 1)) + FromLine - 1          ' Get line no
               l = VAL(PARSE$(t, "/", 2))                         ' Get count of lines
               FOR a = i TO i + l - 1                             ' Init the loop
                  IF ISFALSE IsLBottom(a) THEN me.LFlagBitOn(a, %Invisible) ' Mark the line Invisible now
               NEXT a                                             '
            NEXT k                                                '
         END IF                                                   '

         '----- Do NOTE lines if present
         NoteCount = VAL(PARSE$(H1, ",", 9))                      ' Get Note count
         IF NoteCount > 0 THEN                                    ' We have some to do
            DO WHILE NoteCount                                    ' Loop doing them
               IF ISFALSE EOF(SIO.FNum) THEN                      ' Better be some lines here
                  LINE INPUT # SIO.FNum, HN                       ' Get a NOTE line
                  IF LEFT$(HN, 2) <> "#N" THEN                    ' Better be a NOTE line
                     SIO.Close                                    '
                     MErrExit(%eFail, "STATE file NOTE line is misplaced, STATE terminating.")
                  ELSE                                            '
                     HN = MID$(HN, 3)                             ' Strip off the #N
                  END IF                                          '
                  DECR NoteCount                                  ' Count them
                  j = VAL(PARSE$(HN, "/", 1))                     ' Get Note type
                  i = VAL(PARSE$(HN, "/", 2)) + FromLine - 1 + bias ' Get line no
                  t = PARSE$(HN, "/", 3)                          ' Get the text
                  me.LInsertEmpty(i, 1, %Note)                    ' Insert a line to hold the Note
                  me.LTxtSet(i + 1, t)                            ' Add the text
                  LWrk2S(i + 1) = j                               ' Set the note type
                  IF IsLInvisible(i) AND IsLInvisible(i + 2) THEN me.LFlagBitOn(i + 1, %Invisible)
                  me.UpdLControl(i + 1)                           ' Setup LLCtl
                  INCR Bias                                       '
               ELSE                                               ' We've hit EOF
                  SIO.Close                                       '
                  MErrExit(%eFail, "STATE file is missing NOTE lines, STATE terminating.")
               END IF                                             '
            LOOP                                                  '
         END IF                                                   '

         SIO.Close                                                '

         '----- Finish off
         IF missing THEN                                          ' Fail to reset them all?
            scError(%eFail, "Warning, STATE could not reload 1 or more Labels, could not locate line(s)")
         END IF                                                   '
         IF duplabel THEN                                         ' Duplicate?
            scError(%eFail, "Warning, MEDIT detected duplicate labels, only the 1st reloaded")
         END IF                                                   '
         OnStateExist                                             ' Remember STATE data existed
         sDoStatusBar($SBState)                                   ' re-Do the StatusBar box
         MExitMeth                                                '

      GetHashText:
         IF Prf.PPreserve = 0 THEN                                ' PRESERVE OFF?
            t = RTRIM$(LTxtG(i))                                  '
         ELSE                                                     '
            IF Prf.PPreserve = 1 THEN                             ' PRESERVE ON?
               t = LTxtG(i)                                       '
            ELSE                                                  ' Must be PRESERVE C
               j = INSTR(-1, LTxtG(i), "\")                       ' Get location of last \ in the line
               IF j = 0 THEN                                      ' No \ character
                  t = LTxtG(i)                                    ' Then just like PRESERVE ON
               ELSE                                               ' We have a \
                  IF j = LEN(L(i).@LTxt) THEN                     ' Right at the end?
                     t = LTxtG(i)                                 ' Then just like PRESERVE ON
                  ELSE                                            ' Else we truncate it at the \
                     t = LEFT$(LTxtG(i), j)                       '
                  END IF                                          '
               END IF                                             '
            END IF                                                '
         END IF                                                   '

         '----- If AUTOxx file, preserve FF page breaks
         IF LEFT$(Prf.EOL, 4) = "AUTO" THEN                       '
            IF IsLPage(i) AND i<> 1 THEN _                        ' =PAGE> line?
               t = $FF + t                                        ' Re-Insert the FF page breaks
         END IF                                                   '

         IF Prf.LRECL > 0 THEN                                    '
            t = LSET$(t, Prf.LRECL)                               '
         END IF                                                   '
         RETURN                                                   '

      GetLineHash:
         Hash = 0                                                 ' Init hash
         IF UseNewHash THEN                                       ' New hash?
            sFCS32Init(Hash)                                      ' Initialize the Hash count
            Hash = sFCS32Update(Hash, BYVAL STRPTR(t), LEN(t))    ' Accum the hash
            sFCS32Final(Hash)                                     ' Finalize Hash
         ELSE                                                     ' Still the old hash
            Hash = sHash(t, Hash)                                 ' Do hash the old way
         END IF                                                   '
         RETURN                                                   '
      END METHOD

      METHOD StateSave(fn AS STRING, fline AS LONG, tline AS LONG, filter AS LONG)
      '---------- Save the STATE information
      REGISTER i AS LONG
      REGISTER j AS LONG
      LOCAL RealLines, TopLine, LblLines, TagLines, PenLines, ClrLines, XLines, UsrLines, NoteLines AS LONG
      LOCAL Hash, LHash AS DWORD
      LOCAL H1, HL, HT, HP, HC, HX, HN, HU, lclFn, AttrClr  AS STRING
      LOCAL k, x, lno, bail, attr AS LONG, t, tt AS STRING,  aTxt AS WSTRING POINTER
      LOCAL CreationTime, LastAccessTime, LastWriteTime AS FILETIME
      LOCAL AttrAsc1, AttrAsc2 AS WORD
      LOCAL hFile1            AS LONG
      LOCAL hFile2            AS DWORD
      LOCAL Notes() AS STRING
      DIM Notes(1 TO 100) AS STRING
      LOCAL ssIO AS iIO                                           ' For our I/O stuff

         MEntry

         '----- Make the fn into a 'safe' fn
         lclFn = fn + ".STATE"                                    ' Temp copy with .STATE on the end
         REPLACE ANY ":\/" WITH "```" IN lclFN                    ' Make : / and \ into `
         lclFn = ENV.StatePath + lclFn                            ' Add our STATE folder

         '----- See if we even do anything at all
         IF fn = "\" THEN MExitMeth                               ' If this is an empty tab bail out

         '----- Check STATE is on
         IF Prf.PState = %StateOff THEN                           ' If we're not even doing STATE

            '----- No STATE data, delete any existing STATE file
            sRecycleBin(lclFN, "D")                               ' Delete STATE file to the recycle bin
         OffStateExist                                            ' Remember no STATE data existed
         sDoStatusBar($SBState)                                   ' re-Do the StatusBar box

         '----- OK, it appears STATE needs to be taken under some kind of circumstances
         ELSE

            '----- See if FEW/MANY conditions to be met
            IF Prf.PState = %StateFew  AND ISFALSE ISFILE(lclFn) AND ISFALSE IsStateForce THEN MExitMeth
            IF Prf.PState = %StateMost AND sIsNullFile(lclFn) THEN MExitMeth

            '----- Spin through data lines collecting STATE data
            sFCS32Init(Hash)                                      ' Initialize the Hash count
            FOR i = fline TO tline                                ' Loop through the line range

               '----- Do filtering if needed to match the callers
               IF filter THEN                                     ' Filtering asked for?
                  IF (IsCfNX AND IsLInvisible(i)) OR _            ' NX? If Invisible skip it
                     (IsCfX AND ISFALSE IsLInvisible(i)) OR _     ' X? If not Invisible skip it
                     (IsCfU AND ISFALSE IsLUser(i)) OR _          ' U?  If Not a User line, skip it?
                     (IsCfNU AND IsLUser(i)) THEN ITERATE FOR     ' NU?  Skip if a User line?
               END IF                                             '

               '----- Do Data lines
               IF IsLData(i) THEN                                 ' Data lines
                  INCR RealLines                                  ' Count real text lines
                  IF Prf.PPreserve = 0 THEN                       ' PRESERVE OFF?
                     t = RTRIM$(LTxtG(i))                         '
                  ELSE                                            '
                     IF Prf.PPreserve = 1 THEN                    ' PRESERVE ON?
                        t = LTxtG(i)                              '
                     ELSE                                         ' Must be PRESERVE C
                        j = INSTR(-1, LTxtG(i), "\")              ' Get location of last \ in the line
                        IF j = 0 THEN                             ' No \ character
                           t = LTxtG(i)                           ' Then just like PRESERVE ON
                        ELSE                                      ' We have a \
                           IF j = LEN(L(i).@LTxt) THEN            ' Right at the end?
                              t = LTxtG(i)                        ' Then just like PRESERVE ON
                           ELSE                                   ' Else we truncate it at the \
                              t = LEFT$(LTxtG(i), j)              '
                           END IF                                 '
                        END IF                                    '
                     END IF                                       '
                  END IF                                          '

                  '----- If AUTOxx file, preserve FF page breaks
                  IF LEFT$(Prf.EOL, 4) = "AUTO" THEN              '
                     IF IsLPage(i) AND RealLines <> 1 THEN _      ' =PAGE> line?
                        t = $FF + t                               ' Re-Insert the FF page breaks
                  END IF                                          '

                  IF Prf.LRECL > 0 THEN                           '
                     t = LSET$(t, Prf.LRECL)                      '
                  END IF                                          '
                  Hash = sFCS32Update(Hash, BYVAL STRPTR(t), LEN(t)) ' Accum the hash
                  sFCS32Init(LHash)                               ' Initialize the Hash count
                  LHash = sFCS32Update(LHash, BYVAL STRPTR(t), LEN(t)) ' Accum the hash for line by itself
                  sFCS32Final(LHash)                              ' Finalize Hash
                  tt = HEX$(LHash, 8)                             ' Get hash for this line by itself

                  '----- Handle Labels
                  IF L(i).LLbl <> $BlankLNo THEN                  ' Look for any Labels
                     INCR LblLines                                ' Count them
                     IF LblLines = 1 THEN HL = "#L"               ' Initialize Label line
                     HL += TRIM$(L(i).LLbl) + "/" + FORMAT$(RealLines) + "/" + tt + "," ' Add to HL string
                  END IF                                          '

                  '----- Handle Tags
                  IF LTagG(i) <> $BlankLNo THEN                   ' Look for any tags
                     INCR TagLines                                ' Count them
                     IF TagLines = 1 THEN HT = "#T"               ' Initialize Tag line
                     HT += TRIM$(LTagG(i)) + "/" + FORMAT$(RealLines) + "/" + tt + ","  ' Add to TL string
                  END IF                                          '

                  '----- Handle User Lines
                  IF IsLUser(i) THEN                              ' Look for %User flag
                     INCR UsrLines                                ' Count them
                     IF UsrLines = 1 THEN HU = "#U"               ' Initialize User Line
                     HU += FORMAT$(RealLines) + "/" + tt + ","    ' Add to HU string
                  END IF                                          '

                  '----- PEN Clr entries
                  aTxt = LAttrGP(i)                               ' Get Attr string
                  FOR j = 1 TO LEN(@Atxt)                         ' Loop through it
                     AttrAsc1 = ASC(MID$(@aTxt, j)) AND %AttrHiLite' Get any HiLite
                     IF AttrAsc1 <> 0 THEN                        ' Anything
                        k = AttrAsc1                              ' Temp copy
                        SHIFT RIGHT k, 8                          ' Right align
                        AttrClr = MID$(nHiLitesChrs, k, 1)        ' Convert to a single Alpha character
                        INCR ClrLines                             ' Count
                        IF ClrLines = 1 THEN HC = "#C"            '
                        HC += FORMAT$(RealLines) + "/" + AttrClr + "/"  + FORMAT$(j) + "/" ' Start building the entry
                        FOR k = j TO LEN(@aTxt)                   ' Look for end of hilight string
                           AttrAsc2 = ASC(MID$(@aTxt, k)) AND %AttrHiLite' Get next Attr byte
                           IF AttrAsc2 = AttrAsc1 AND k <> LEN(@aTxt) THEN ITERATE FOR ' Still the same color? Keep looking for the end
                           IF k = LEN(@aTxt) THEN INCR k          ' Adjust for end of line
                           HC += FORMAT$(k - j)   '               '
                           j = k: EXIT FOR                        '
                        NEXT j                                    '
                        HC += "/" + tt + ","                      '
                     END IF                                       '
                  NEXT i                                          '

               '----- Exclude line(s) marker
               ELSEIF IsLXclude(i) AND ISFALSE IsCfNX AND ISFALSE IsCfX THEN  ' Look for X Marker
                  j = LWrk1G(i)                                   ' Get lines in the block
                  FOR k = i + 1 TO i + j                          ' Scan block for NOTE/U/NU lines
                     IF IsLNote(k) THEN DECR j                    ' If a NOTE reduce X range
                     IF (IsCfU AND ISFALSE IsLUser(k)) OR _       ' If U or NU filtered,
                        (IsCfNU AND IsLUser(k)) THEN DECR j       ' then adjust count
                  NEXT                                            '
                  IF j > 0 THEN                                   ' Any left?
                     INCR XLines                                  ' Count it
                     IF XLines = 1 THEN HX = "#X"                 ' Init string
                     HX += FORMAT$(RealLines + 1) + "/" + FORMAT$(j) + "," ' Add to XL string
                  END IF

               '----- NOTE line(s)
               ELSEIF IsLNote(i) THEN                             ' Look for NOTE lines
                  INCR NoteLines                                  ' Start counting
                  IF NoteLines > UBOUND(Notes()) THEN _           ' Keep array big enough
                     REDIM PRESERVE Notes(1 TO UBOUND(Notes()) * 2) AS STRING
                  Notes(NoteLines) = "#N" + FORMAT$(L(i).LWrk2) + "/" + FORMAT$(RealLines) + "/" + LTxtG(i) ' Save data
               END IF                                             '

            NEXT i                                                '
            sFCS32Final(Hash)                                     ' Finalize Hash

            '----- Write the STATE file
            LET ssIO = CLASS "cIO"                                '
            ssIO.Setup("O", "", "", lclFn)                        ' Tell IO what we're opening
            IF ssIO.EXEC THEN                                     ' See if OK
               ssIO.Close                                         '
               sDoMsgBox("Error in writing file", %MB_OK OR %MB_USERICON, "STATESave")    '
            ELSE
               t = "#1" + HEX$(Hash, 8) + "," + FORMAT$(RealLines) + "," + IIF$(IsMEdit, "1", FORMAT$(TopScrn)) + "," + _
                   FORMAT$(LblLines) + "," + FORMAT$(TagLines) + "," + FORMAT$(UsrLines) + "," + FORMAT$(ClrLines) + "," + _
                   FORMAT$(XLines) + "," + FORMAT$(NoteLines) + ","
               PRINT # ssIO.FNum, t                               '
                                                                  '
               '----- Dump Label line if present                  '
               IF LblLines > 0 THEN _                             '
                  PRINT # ssIO.FNum, HL                           '
               '----- Dump Tag line if present                    '
               IF TagLines > 0 THEN _                             '
                  PRINT # ssIO.FNum, HT                           '
               '----- Dump Usr line if present                    '
               IF UsrLines > 0 THEN _                             '
                  PRINT # ssIO.FNum, HU                           '
               '----- Dump Clr line if present                    '
               IF ClrLines > 0 THEN _                             '
                  PRINT # ssIO.FNum, HC                           '
               '----- Dump Exclude line if present                '
               IF XLines > 0 THEN _                               '
                  PRINT # ssIO.FNum, HX                           '

               '----- Dump out NOTE lines if present
               IF NoteLines > 0 THEN                              ' NOTE lines saved?
                  FOR j = 1 TO NoteLines                          ' Yes, dump them out
                     PRINT # ssIO.FNum, Notes(j)                  '
                  NEXT j                                          '
               END IF                                             '
               ssIO.Close                                         ' Then close it
               OffStateForce                                      ' Clear Force just in case
            END IF                                                '
            OnStateExist                                          ' Remember STATE data existed
            sDoStatusBar($SBState)                                ' re-Do the StatusBar box
         END IF
         MExitmeth
      END METHOD

      METHOD SwapKill()
      '---------- UnMark and kill any text selection
      LOCAL ul AS rect, i AS LONG
         IF IsSwapDrawn THEN                                      ' A previous one drawn?
            FOR i = SwapRect.nTop TO SwapRect.nBottom             ' Clear previous UL Attr for the lines
               me.AttrULClear(i, SwapRect.nLeft, SwapRect.nRight)
            NEXT i                                                '
            OffSwapDrawn                                          '
         END IF                                                   '
         OffSwapActive                                            '
      END METHOD

      METHOD SwapScr()
      '---------- Underline the selected text on the screen
      REGISTER i AS LONG
         IF ISFALSE IsSwapActive THEN EXIT METHOD                 ' Just in case
         IF IsSwapDrawn THEN                                      ' A previous one drawn?
            FOR i = SwapRect.nTop TO SwapRect.nBottom             ' Clear previous UL Attr for the lines
               me.AttrULSet(i, SwapRect.nLeft, SwapRect.nRight)   '
            NEXT i                                                '
            OffSwapDrawn                                          '
         END IF                                                   '

         FOR i = SwapSLin TO SwapELin                             ' Set UL Attr for the lines
            me.AttrULSet(i, SwapSCol, SwapECol)                   '
         NEXT i                                                   '

         SwapRect.nLeft = SwapSCol: SwapRect.nRight  = SwapECol   ' Save where we drew it
         SwapRect.nTop  = SwapSLin: SwapRect.nBottom = SwapELin   ' Save where we drew it
         DoSet(%Refresh)                                          ' Have it looked at
         OnSwapDrawn                                              '
         sCaretCreate                                             '
         sDoCursor                                                ' Get cursor back
         sCaretShow                                               '
      END METHOD

      '----- Following stube allow access to the TIP Object from outside the TP Object
      METHOD TIPBase() AS STRING: METHOD = TIP.Base: END METHOD   ' Pass on requests to the TIP object
      METHOD TIPDate() AS STRING: METHOD = TIP.Date: END METHOD
      METHOD TIPEXEC() AS LONG: METHOD = TIP.EXEC: END METHOD
      METHOD TIPExtn() AS STRING: METHOD = TIP.Extn: END METHOD
      METHOD TIPFile() AS STRING: METHOD = TIP.File: END METHOD
      METHOD TIPFilePath() AS STRING: METHOD = TIP.FilePath: END METHOD
      METHOD TIPPath() AS STRING: METHOD = TIP.Path: END METHOD
      METHOD TIPProfile() AS STRING: METHOD = TIP.Profile: END METHOD
      METHOD TIPResultMsg() AS STRING: METHOD = TIP.ResultMsg: END METHOD
      METHOD TIPROStat() AS LONG: METHOD = TIP.ROStat: END METHOD
      METHOD TIPSetup(flags AS STRING, prof AS STRING, prompt AS STRING, fn AS STRING): TIP.Setup(flags, prof, prompt, fn): END METHOD
      METHOD TIPTime() AS STRING: METHOD = TIP.Time: END METHOD
      METHOD TIPTimeDateRefresh(): TIP.TimeDateRefresh: END METHOD

      METHOD TitleCase(lno AS LONG, scol AS LONG, ecol AS LONG)
      '---------- Do title case on one line
      DIM words() AS STRING
      LOCAL txt1, txt2 AS STRING
      LOCAL i, j, k, m AS LONG
         MEntry
         txt1 = LLCASE(MID$(LTxtG(lno), scol, ecol - scol + 1))   ' Get copy of substring to process

         '----- Break the line into words
         i = PARSECOUNT(txt1, " ")                                ' Count words
         REDIM words(1 TO i) AS STRING                            ' Make table correct size
         PARSE txt1, words(), " "                                 ' Parse the words into the table
         m = 1                                                    ' Start scan at 1

         '----- Find each word now and uppercase it's 1st letter
         FOR j = 1 TO i                                           ' Loop for each word
            IF ISNOTNULL(words(j)) THEN                           ' Ignore null 'words'
               k = INSTR(m, txt1, words(j))                       ' Find location of this word
               IF k THEN MID$(txt1, k, 1) = UUCASE(MID$(txt1, k, 1))' Uppercase the 1st letter
               m = k + LEN(words(j))                              ' Start scan adjust
            END IF                                                '
         NEXT j                                                   '

         '----- Stuff back the converted string
         txt2 = LTxtG(lno)                                        ' Get whole text line
         MID$(txt2, scol, ecol - scol + 1) = txt1                 ' Stuff back in the processed txt
         me.LTxtSet(lno, txt2)                                    ' Replace the whole line
         me.AttrScan(lno)                                         ' Recolorize
         MExit                                                    '
      END METHOD

      METHOD TPPrfGetCols() AS LONG: METHOD = Prf.Cols: END METHOD
      METHOD TPPrfGetSS2APtr() AS LONG: METHOD = Prf.GetSS2APtr: END METHOD
      METHOD TPPrfGetSA2SPtr() AS LONG: METHOD = Prf.GetSA2SPtr: END METHOD
      METHOD TPPrfGetCA2SPtr() AS LONG: METHOD = Prf.GetCA2SPtr: END METHOD
      METHOD TPPrfGetCS2APtr() AS LONG: METHOD = Prf.GetCS2APtr: END METHOD

      METHOD Translate(BYREF InText AS STRING, TTbl AS LONG)
      '---------- Do a character translate
      REGISTER i AS LONG
      LOCAL pSrc AS BYTE PTR
      LOCAL Rep() AS BYTE
      LOCAL tt AS STRING PTR
         tt = TTbl
         REDIM rep(255) AT STRPTR(@tt)                            ' Get the table
         pSrc = STRPTR(InText)                                    ' Create pointer to string
         FOR i = 1 TO LEN(InText)                                 ' Do each character
            @pSrc = rep(@pSrc)                                    '
            INCR pSrc                                             '
         NEXT i                                                   '
      END METHOD

      METHOD TTblLinNoGet(x AS LONG) AS LONG: METHOD = TTbl(x).LinNo: END METHOD

      METHOD TTblLinNoSet(x AS LONG, lno AS LONG) AS LONG: TTbl(x).LinNo = lno: END METHOD

      METHOD TTblGet(x AS LONG) AS TouchEntry
         METHOD = TTbl(x)
      END METHOD
      METHOD TTblSet(x AS LONG, t AS TouchEntry)
         TTbl(x) = t
      END METHOD
      METHOD TTblReset(): RESET TCtr, TTbl(): END METHOD

      METHOD TTblScan(lno AS LONG) AS STRING
      REGISTER i AS LONG
      LOCAL cv AS LAsString, tv AS STRING
         MEntry
         cv.LVar = lno: tv = cv.CVar
         ARRAY SCAN TTbl(), FROM 1 TO 4, = tv, TO i               ' See if it's here
         IF i AND TTbl(i).LinNo <> 0 THEN                         ' and not been nulled
            METHOD = TTbl(i).LinCtl                               ' Pass String
         ELSE                                                     '
            METHOD = ""                                           ' Pass back Null
         END IF                                                   '
         MExit
      END METHOD

      METHOD TTblAdd(lno AS LONG, ldata AS STRING)
      REGISTER i AS LONG
      REGISTER j AS LONG
         MEntry
         IF TCtr > 0 THEN                                         ' Some already, see if already here
            FOR i = 1 TO TCtr                                     '
               IF TTbl(i).LinNo = lno THEN                        ' Already here?
                  TTbl(i).LinCtl = LSET$(ldata, 8)                ' Then replace the data
                  MExitMeth                                       ' We're done
               END IF                                             '
            NEXT i                                                '
         END IF                                                   '
         INCR TCtr                                                ' Must add it, Incr count
         IF TCtr > UBOUND(TTbl()) THEN _                          ' Keep table big enough
            REDIM PRESERVE TTbl(1 TO TCtr + 500) AS GLOBAL TouchEntry
         TTbl(TCtr).LinNo = lno                                   ' Save touched line
         TTbl(TCtr).LinCtl = LSET$(ldata, 8)                      ' And touched data
         TTbl(TCtr).LinCmd = $BlankLNo                            ' Clear all other fields
         TTbl(TCtr).LinRpt = 0                                    '
         TTbl(TCtr).LinFlag = 0                                   '
         TTbl(TCtr).LinCmdIX = 0                                  '
         TTbl(TCtr).LinCType = " "                                '
         TTbl(TCtr).LinCVar = " "                                 '
         MExit
      END METHOD

      METHOD TTblDel(lno AS LONG)
      REGISTER i AS LONG
         MEntry
         IF TCtr = 0 THEN MExitMeth
         FOR i = TCtr TO 1 STEP -1                                ' Loop through the Touched table
            IF BIT(TTbl(i).LinFlag, %LCmdR) THEN ITERATE FOR      ' If a R (retain) line command, ignore this
            IF TTbl(i).LinNo = lno THEN                           ' The one we're looking for?
               ARRAY DELETE TTbl(i)                               ' Delete the entry
               DECR TCtr                                          '
               EXIT FOR                                           '
            END IF                                                '
         NEXT i                                                   '
         MExit
      END METHOD

      METHOD TTblSort()
      LOCAL i AS LONG
         ARRAY SORT TTbl() FOR TCtr, FROM 1 TO 4, CALL sBinSort()
      END METHOD

      METHOD TTblPrune()
      REGISTER i AS LONG
      REGISTER j AS LONG
         MEntry
         FOR i = TCtr TO 1 STEP -1                                ' Loop through the Touched table
            IF TTbl(i).LinNo = 0 THEN                             ' A 'nulled' entry
               ARRAY DELETE TTbl(i)                               ' Delete the entry
               INCR j                                             ' Count deleted, stay at tis index
            END IF                                                '
         NEXT i                                                   '
         TCtr -= j                                                '
         MExit
      END METHOD

      METHOD UndoInit()                                           '
      REGISTER i AS LONG                                          '
         MEntry
         IF UBOUND(Undo) > 0 THEN                                 ' If any prior table, get rid of it
            FOR i = 1 TO UBOUND(Undo())                           ' Delete temporary files
               ON ERROR GOTO UndoError                            ' In case there are none to Kill
               IF ISNOTNULL(TRIM$(Undo(i).UFn)) THEN KILL TRIM$(Undo(i).UFn)
               IF ISNOTNULL(TRIM$(Undo(i).TFn)) THEN KILL TRIM$(Undo(i).TFn)
               IF ISNOTNULL(TRIM$(Undo(i).TWFn)) THEN KILL TRIM$(Undo(i).TWFn)
               IF ISNOTNULL(TRIM$(Undo(i).IXFn)) THEN KILL TRIM$(Undo(i).IXFn)
               UndoResume:
               ON ERROR GOTO 0                                    ' Kill error trap
            NEXT i                                                '
         END IF                                                   '

         IF Prf.UndoNumber > 0 THEN                               ' A new table needed?
            REDIM Undo(1 TO Prf.UndoNumber) AS INSTANCE UndoType  ' Allocate Undo slot data
            FOR i = 1 TO Prf.UndoNumber                           ' Init Undotable
               Undo(i).UFn = sGetNewTempFile("UNDO")              ' Get temp files allocated
               Undo(i).TFn = sGetNewTempFile("UNDO")              '
               Undo(i).TWFn = sGetNewTempFile("UNDO")             '
               Undo(i).IXFn = sGetNewTempFile("UNDO")             '
               Undo(i).Time = 0                                   '
            NEXT i                                                '
         END IF                                                   '

         MExitMeth

         UndoError:
            RESUME UndoResume                                     ' Ignore error
      END METHOD                                                  '


      METHOD UndoModifiedGet (BYVAL ix AS LONG) AS LONG: METHOD = Undo(ix).Modified: END METHOD
      METHOD UndoModifiedSet (BYVAL ix AS LONG, BYVAL num AS LONG): Undo(ix).Modified = num: END METHOD
      METHOD UndoTopScrGet (BYVAL ix AS LONG) AS LONG: METHOD = Undo(ix).TopScr: END METHOD
      METHOD UndoTopScrSet (BYVAL ix AS LONG, BYVAL num AS LONG): Undo(ix).TopScr = num: END METHOD
      METHOD UndoLastLineGet (BYVAL ix AS LONG) AS LONG: METHOD = Undo(ix).LastLine: END METHOD
      METHOD UndoLastLineSet (BYVAL ix AS LONG, BYVAL num AS LONG): Undo(ix).LastLine = num: END METHOD
      METHOD UndoLastRealGet (BYVAL ix AS LONG) AS LONG: METHOD = Undo(ix).LastReal: END METHOD
      METHOD UndoLastRealSet (BYVAL ix AS LONG, BYVAL num AS LONG): Undo(ix).LastReal = num: END METHOD
      METHOD UndoTimeGet (BYVAL ix AS INTEGER) AS DOUBLE: METHOD = Undo(ix).Time: END METHOD
      METHOD UndoTimeSet (BYVAL ix AS INTEGER, BYVAL num AS DOUBLE): Undo(ix).Time = num: END METHOD
      METHOD UndoUFnGet (BYVAL ix AS LONG) AS STRING: METHOD = Undo(ix).UFn: END METHOD
      METHOD UndoBusyGet (BYVAL ix AS LONG) AS LONG: METHOD = ISTRUE Undo(ix).Busy: END METHOD
      METHOD UndoBusySet (BYVAL ix AS LONG): Undo(ix).Busy = %True: END METHOD

      METHOD UndoSave()
      '---------- Take an Undo save if needed
      LOCAL i, j AS LONG, MEditStr AS STRING
      LOCAL LUndo AS UNDOType POINTER                             ' Pointer to the UNDO Data
         MEntry                                                   '
         IF ISFALSE IsUndoFlag OR Prf.UndoNumber = 0 THEN MExitMeth  ' Only if needed
         OffUndoFlag                                              ' Reset it
         IF UndoCurr < Prf.UndoNumber THEN                        ' Select next save slot
            UndoCurr += 1                                         '
         ELSE                                                     '
            UndoCurr = 1                                          ' Wrap if at the end
         END IF                                                   '
         LUndo = VARPTR(Undo(UndoCurr))                           ' Point at the entry used
         j = 0                                                    ' Set escape counter
         DO WHILE @LUndo.Busy                                     ' Must wait if last write of this slot still active
            SLEEP 5                                               '
            INCR j                                                '
            IF j > 600 THEN EXIT DO                               ' 3 seconds?  Escape
         LOOP                                                     '

         '---------- Copy the TopScr/LastLine
         @LUndo.TopScr = TopScrn                                  ' Just copy it
         @LUndo.LastLine = LastLine                               '
         @LUndo.LastReal = LastReal                               '
         @LUndo.Modified = IIF(IsModdFlag, %True, %False)         '
         @LUndo.UBoundL = UBoundL                                 '
         @LUndo.UBoundT = UBoundT                                 '
         @LUndo.Time = TIMER                                      ' Save the time
         IF IsMedit THEN                                          ' If MEdit, must save individual modified status
            MEditStr = SPACE$(255)                                '
            FOR j = 1 TO MEditCount                               ' Pick up the MEdit modified values
               MID$(MEditStr, j, 1) = IIF$(me.MEditFlagGet(j), "T", "F")' Build string with T/F for each file
            NEXT j                                                '
            @LUndo.MEditMod = MEditStr                            ' Store it
         END IF                                                   '

         '---------- Do the IX string, L(), T() and TW() arrays
         @LUndo.MCpyBusy = %True                                  ' Mark this as memory copy busy
         @LUndo.Busy = %True                                      ' Mark this slot Busy
         me.UndoPut(LUndo)                                        ' Go Write UNDO checkpoint
         MExit
      END METHOD

      METHOD UndoTFnGet (BYVAL ix AS LONG) AS STRING: METHOD = Undo(ix).TFn: END METHOD
      METHOD UndoTWFnGet (BYVAL ix AS LONG) AS STRING: METHOD = Undo(ix).TWFn: END METHOD
      METHOD UndoIXFnGet (BYVAL ix AS LONG) AS STRING: METHOD = Undo(ix).IXFn: END METHOD
      METHOD UndoMEditGet (BYVAL ix AS LONG) AS STRING: METHOD = Undo(ix).MEditMod: END METHOD
      METHOD UndoMEditSet (BYVAL ix AS LONG, BYVAL st AS STRING): Undo(ix).MEditMod = st: END METHOD

      METHOD AdjustPending(ln AS LONG, Number AS LONG, CsrAdj AS LONG)
      '---------- Adjust all Pending stuff
      REGISTER x AS LONG
         MEntry

         '----- Adjust Cursor positioning
         IF sCurPrio > 0 THEN                                     ' If Cursor request pending
            IF sCurLin > (ln + CsrAdj) THEN                       ' Only higher line numbers should need doing
               sCurLin = MAX(1, sCurLin + Number)                 '
            END IF                                                '
            IF (ln + CsrAdj) >= CsrAbsLine OR (ln + CsrAdj) + Number >= CsrAbsLine THEN '
               ' Null                                             '
            ELSE                                                  '
               IF number > 0 THEN CsrAbsLine += number            ' Adjust CsrAbsLine if action above it
            END IF                                                '
         END IF                                                   '

         '----- Adjust LTblB entries
         IF LTblBIX <> 0 THEN                                     ' If LTblB entries left
            FOR x = 1 TO LTblBIX                                  ' Adjust remaining pending commands
               IF LTblB(x).SrcFrom > ln THEN LTblB(x).SrcFrom += Number ' Only higher line numbers should need doing
               IF LTblB(x).SrcTo > ln THEN LTblB(x).SrcTo += Number     '
               IF LTblB(x).DstFrom > ln THEN LTblB(x).DstFrom += Number '
               IF LTblB(x).DstTo > ln THEN LTblB(x).DstTo += Number     '
            NEXT x                                                ' Next pending command
         END IF                                                   '

         '----- Adjust LTblRange entries
         IF LTblRange THEN                                        ' Waiting line-range operands?
            IF LTblSFrom > ln THEN LTblSFrom += Number            ' If so, adjust source
            IF LTblSTo   > ln THEN LTblSTo   += Number            '
            IF LTblDFrom > ln THEN LTblDFrom += Number            ' If so, adjust destination
            IF LTblDTo   > ln THEN LTblDTo   += Number            '
         END IF                                                   '

         '----- Adjust Touch Table
         IF TCtr > 0 THEN                                         ' TTbl entries?
            FOR x = 1 TO TCtr                                     ' Loop through the Touched table
               IF me.TTblLinNoGet(x) > ln THEN                    ' A higher line number
                  me.TTblLinNoSet(x, me.TTblLinNoGet(x) + Number) ' Adjust it
               END IF                                             '
            NEXT x                                                '
         END IF                                                   '
         MExit
      END METHOD

      METHOD AFLCmdSet(i AS LONG, vl AS STRING)
         AFList(i).Cmd = vl
      END METHOD

      METHOD AttrInvClear(lno AS LONG, fcolumn AS LONG, tcolumn AS LONG)
      REGISTER i AS LONG
      REGISTER j AS LONG
         j = LAttrSPtr(lno)                                       ' Get pointer to actual characters
         FOR i = j + ((fcolumn - 1) * 2) TO j + ((tcolumn - 1) * 2) STEP 2  ' Attr by Attr
            POKE WORD, i, (PEEK(WORD, i) AND (%AttrAll - %AttrInv)) ' Remove the INV flag
         NEXT i                                                   '
      END METHOD

      METHOD AttrInvClearFind()
      REGISTER i AS LONG
      REGISTER j AS LONG
         IF CursFindInv.nTop <> 0 THEN                            ' If something marked
            j = LAttrSPtr(CursFindInv.nTop)                        ' Get pointer to the line
            FOR i = j + (CursFindInv.nLeft - 1) * 2 TO j + (CursFindInv.nRight - 1) * 2       ' Attr by Attr
               POKE WORD, i, (PEEK(WORD, i) AND (%AttrAll - %AttrInv)) ' Remove the INV flag
            NEXT i                                                '
         END IF                                                   '
      END METHOD

      METHOD AttrInvGet(lno AS LONG, pcolumn AS LONG) AS LONG
      REGISTER i AS LONG
      REGISTER j AS LONG
      LOCAL AttrLine AS WSTRING POINTER
         METHOD = %False                                          ' Default answer
         j = LAttrSPtr(lno)                                       ' Get pointer to actual characters
         i = j + ((pcolumn - 1) * 2)                              ' Point at Attr char
         IF (PEEK(WORD, i) AND %AttrInv) THEN METHOD = %True      ' Pass back INV status
      END METHOD

      METHOD AttrInvGrab( lno AS LONG, ccolumn AS LONG)
      REGISTER i AS LONG
      REGISTER j AS LONG
      LOCAL k AS LONG
         CursFindInv.ntop = 0                                     ' Default to saying no hilite
         j = LAttrSPtr(lno)                                       ' Get pointer to actual characters
         k = j + (LAttrGLen(lno) - 1) * 2                         ' Get pointer to last Attr character
         IF LAttrGLen(lno) >= ccolumn THEN                        ' Only if a string is present and long enough
            i = j + (ccolumn - 1) * 2                             ' Point at cursor's location in line
            IF (PEEK(WORD, i) AND %AttrInv) = 0 THEN EXIT METHOD  ' If not sitting on a INV string, exit
            CursFindInv.nTop = lno                                ' Save the line number
            DO WHILE (i - 2) > j AND (PEEK(WORD, i - 2) AND %AttrInv) <> 0  ' Find start of INV string
               i -= 2                                             '
            LOOP                                                  '
            CursFindInv.nLeft = ((i - j) / 2) + 1                 ' Save Left end of INV string
            i = j + (ccolumn - 1) * 2                             ' Point at cursor's location in line
            DO WHILE (i + 2) <= k AND (PEEK(WORD, i + 2) AND %AttrInv) <> 0 ' Find end of INV string
               i += 2                                             '
            LOOP                                                  '
            CursFindInv.nRight = ((i - j) / 2) + 1                ' Save Right end of INV string
         END IF                                                   '
      END METHOD

      METHOD AttrInvSet(lno AS LONG, fcolumn AS LONG, tcolumn AS LONG)
      REGISTER i AS LONG
      REGISTER j AS LONG
         j = LAttrSPtr(lno)                                       ' Get pointer to actual characters
         FOR i = j + ((fcolumn - 1) * 2) TO j + ((tcolumn - 1) * 2) STEP 2  ' Attr by Attr
            POKE WORD, i, (PEEK(WORD, i) OR %AttrInv)             ' Add the INV flag
         NEXT i                                                   '
      END METHOD

      METHOD AttrHiLiteClear(lno AS LONG, fcolumn AS LONG, tcolumn AS LONG, pcolor AS LONG)
      REGISTER i AS LONG
      REGISTER j AS LONG
      LOCAL scolor AS LONG, AttrAsc AS WORD
         scolor = pcolor                                          ' Get the color
         SHIFT LEFT scolor, 8                                     ' Shift up to hi-byte
         j = LAttrSPtr(lno)                                       ' Get pointer to actual characters
         FOR i = j + ((fcolumn - 1) * 2) TO j + ((tcolumn - 1) * 2) STEP 2  ' Attr by Attr
            AttrAsc = PEEK(WORD, i) AND %AttrHiLite               ' Get current hilite color
            IF scolor = AttrAsc THEN                              ' The requested color?
               POKE WORD, i, (PEEK(WORD, i) AND (%AttrAll - %AttrHiLite)) ' Remove the color
            END IF                                                '
         NEXT i                                                   '
      END METHOD

      METHOD AttrHiLiteGet(lno AS LONG, fcolumn AS LONG, tcolumn AS LONG) AS STRING
      REGISTER i AS LONG
      REGISTER j AS LONG
      LOCAL t AS STRING, AttrLine AS WSTRING POINTER, AttrAsc AS WORD
         METHOD = ""                                              ' Default answer
         j = LAttrSPtr(lno)                                       ' Get pointer to actual characters
         FOR i = j + ((fcolumn - 1) * 2) TO j + ((tcolumn - 1) * 2) STEP 2  ' Attr by Attr
            AttrAsc = PEEK(WORD, i) AND %AttrHiLite               ' Get current hilite color
            SHIFT RIGHT AttrAsc, 8                                ' Shift to low order
            t += CHR$(AttrAsc)                                    ' Add in the CHR$ of the color
         NEXT i                                                   '
         METHOD = t                                               '
      END METHOD

      METHOD AttrHiLiteSet(lno AS LONG, fcolumn AS LONG, tcolumn AS LONG, pcolor AS LONG)
      REGISTER i AS LONG
      REGISTER j AS LONG
      LOCAL scolor AS LONG, AttrAsc AS WORD
         scolor = pcolor                                          ' Get the color
         SHIFT LEFT scolor, 8                                     ' Shift up to hi-byte
         j = LAttrSPtr(lno)                                       ' Get pointer to actual characters
         FOR i = j + ((fcolumn - 1) * 2) TO j + ((tcolumn - 1) * 2) STEP 2  ' Attr by Attr
            POKE WORD, i, (PEEK(WORD, i) AND (%AttrAll - %AttrHiLite)) OR scolor ' Add in the color
         NEXT i                                                   '
      END METHOD

      METHOD AttrScan(lno AS LONG) AS LONG
      '---------- Scan a line for Colorize
      LOCAL tClr, StartPos, ci, wFlag, TxtLen AS LONG
      LOCAL letter, nxtletter, prvbyte, QSkip AS BYTE POINTER
      LOCAL ESCChar AS BYTE
      LOCAL sTxt AS STRING POINTER
      LOCAL aTxt AS WSTRING POINTER
      REGISTER i AS LONG
      REGISTER j AS LONG
         MEntry

         '----- Get set to go
         ESCChar = IIF(ISNOTNULL(ClrESC), ASC(ClrESC), 0)         ' Get local copy of ESC
         sTxt = LTxtGP(lno): aTxt = LAttrGP(lno)                  ' Setup txt pointer variables
         TxtLen = LTXTGLen(lno)                                   ' Get value once
         IF LEN(@atxt) < TxtLen THEN @aTxt = LSET$(@aTxt, TxtLen USING CHR$$(%SCTxtLo)) ' Init Attr line if needed
         i = LEN(@aTxt)
         IF ISFALSE Prf.HiAuto OR ISFALSE IsClrFlag THEN MExitMeth' If colorization not possible, exit

         '----- If ludicrous, just make it one long string and exit
         IF TxtLen > 2000 THEN                                    ' If ridiculously long, it ain't a source line
            me.AttrSchemeSet(lno, 1, TxtLen, %SCTxtLo)            ' Build the Attr line
            MExitMeth                                             ' We're done
         END IF                                                   '

         '----- Start off doing EXCLUDE / INCLUDE stuff
         IF TRIM$(ClrExclude(1).Txt1) <> "" THEN GOSUB DoExclude  ' We have some excludes to do
         IF TRIM$(ClrInclude(1).Txt1) <> "" THEN GOSUB DoInclude  ' Else test for INCLUDEs


         '----- First look for posible comments
         FOR i = 1 TO 9                                           ' Look at all 9 possible comment types
            IF LEN(TRIM$(ClrComment(i).Txt1)) THEN                ' A comment string present?
               ClrComment(i).Fnd = INSTR(UUCASE(@sTxt), TRIM$(ClrComment(i).Txt1))        '
               IF VAL(TRIM$(ClrComment(i).sCol)) > 0 AND ISTRUE ClrComment(i).Fnd <> VAL(TRIM$(ClrComment(i).sCol)) THEN ClrComment(i).Fnd = 0
            END IF                                                   '
         NEXT i                                                   '

         '----- Set all the scan starting values
         Letter = STRPTR(@sTxt) : wFlag = %False: StartPos = 1: i = 1 ' Startup values

         '----- Now the hard work, byte by byte scanning
         DO                                                       ' Start real scanning
            '----- Past the end?   We're done
            IF i > TxtLen THEN                                    ' No more string? We're done
               IF ISTRUE wFlag THEN GOSUB CloseWord               ' If doing a word, go close it off
               MExitMeth                                          '
            END IF                                                '

            '----- See if we're at the point where any of the comments start
            FOR ci = 1 TO 9                                       ' Loop through the comment controls
               '----- Are we're at the point where a comment starts?
               IF (LEN(TRIM$(ClrComment(ci).Txt1)) AND i = ClrComment(ci).Fnd) THEN ' Finish line if rest is a Comment
                  IF ISTRUE wFlag THEN GOSUB CloseWord            ' If doing a word, go close it off

                  '----- Simple comment type
                  IF INSTR($Numeric, LEFT$(ClrComment(ci).sCol, 1)) THEN '
                     me.AttrSchemeSet(lno, i, TxtLen, ClrComment(ci).Clr) ' Build the Attr line
                     MExitMeth                                    ' We're done

                  '----- Start/End comment type like /*   */
                  ELSE                                            '
                     j = INSTR(I + 1, @sTxt, TRIM$(ClrComment(ci).sCol))' Look for end comment delimiter
                     IF j = 0 THEN                                ' None, treat rest of line as comment
                        me.AttrSchemeSet(lno, i, TxtLen, ClrComment(ci).Clr) ' Build the Attr line
                        MExitMeth                                 ' We're done
                     ELSE                                         '
                        me.AttrSchemeSet(lno, i, j + LEN(TRIM$(ClrComment(ci).sCol)), ClrComment(ci).Clr) ' Build the Attr line
                        i = j + LEN(TRIM$(ClrComment(ci).sCol)): letter = letter + LEN(TRIM$(ClrComment(ci).sCol)) - i ' Adjust to continue scan
                        ClrComment(ci).Fnd = INSTR(i, UUCASE(@sTxt), TRIM$(ClrComment(ci).Txt1)) ' See if a 2nd comment on the line
                        ITERATE DO                                ' Continue
                     END IF                                       '
                  END IF                                          '
               END IF                                             '
            NEXT ci

            '----- White space detected
            IF @letter = 32 THEN                                  ' WhiteSpace?
               IF ISFALSE wFlag THEN                              ' Working on a word?
                  StartPos = i                                    ' No, Save start of whitespace area
                  DO WHILE @letter = 32 AND i <= TxtLen           ' find start of word
                     INCR letter: INCR i                          ' Move right
                  LOOP                                            '
                  me.AttrSchemesET(lno, StartPos, i, %SCTxtLo)    ' Build the Attr line
                  IF i > TxtLen THEN                              ' Stopped by the end?
                     MExitMeth                                    ' We're done
                  ELSE                                            ' No, stopped by next string
                     ITERATE DO                                   ' Just loop back
                  END IF                                          '
               ELSE                                               ' We have a space delimited word
                  GOSUB CloseWord                                 ' Create an entry for it
                  ITERATE DO                                      ' Loop back
               END IF                                             '
            END IF                                                '

            '----- Quote detected
            IF @letter = 34 OR @letter = 39 OR @letter = 96 THEN  ' Double, single or back quote?
               QSkip = letter                                     ' Save which one
               IF ISFALSE wflag THEN                              ' Not in a word?
                  wFlag = %True: StartPos = i                     ' Now we are
               ELSE                                               '
                  GOSUB Closeword                                 ' If in a word, close it
                  wFlag = %True: StartPos = i                     ' and start another
               END IF                                             '
               INCR letter: INCR i                                ' Move right
               DO WHILE (@letter <> @QSkip) AND (i <= TxtLen)     ' find next quote
                  INCR letter: INCR i                             ' Move right
                  IF (@letter = ESCChar) AND (EscChar <> 0) THEN  ' Is this as ESC char?
                     IF i = TxtLen THEN                           ' At very last position of TXT?  Then ESC is 'orphaned'
                        Letter += 1: i += 1                       ' Step over 'orphaned' escape char ONLY
                     ELSE                                         ' they are not doing anything strange with ESC
                        Letter += 2: i += 2                       ' Step over escape AND escaped char
                     END IF                                       '
                  END IF                                          '
               LOOP                                               '
               IF i > TxtLen THEN                                 ' Stopped by the end?
                  me.AttrSchemeSet(lno, StartPos, i, ClrQuoted)   ' Build the Attr line
                  wFlag = %False                                  ' Turn off the word flag
                  MExitMeth                                       ' and bail out
               ELSE                                               ' No, stopped by next Quote
                  me.AttrSchemeSet(lno, StartPos, i + 1, ClrQuoted)  ' Build the Attr line
                  wFlag = %False                                  ' Turn off the word flag
                  INCR letter: INCR i                             ' Move right
                  ITERATE DO                                      ' Just loop back
               END IF                                             '
            END IF                                                '

            '----- User specified delimiter detected
            IF INSTR(ClrDLM, CHR$(@letter)) THEN                  ' Look for a delimiter
               GOSUB GotDLM                                       ' Go handle DLM
               INCR letter: INCR i                                ' Step over
               ITERATE DO                                         ' and loop back
            END IF

            '----- Just any old character detected
            IF ISFALSE wFlag THEN                                 ' Some other old character
               wFlag = %True: StartPos = i                        ' Just say we're in a word
            ELSE                                                  ' Maybe letter after a DLM?
               prvByte = letter - 1                               '
               IF INSTR(ClrDLM, CHR$(@prvbyte)) THEN              ' Look for a previous delimiter
                  GOSUB CloseWord                                 ' Yes, go output it
                  wFlag = %True: StartPos = i                     ' Start over
               END IF                                             '
            END IF

            '----- the ESC character detected
            IF @letter = ESCChar THEN                             ' Is this as ESC char?
               IF i < TxtLen THEN                                 ' Still remaining text?
                  Letter += 2: i += 2                             ' Step over 2 characters
               ELSE                                               '
                  INCR Letter: INCR i                             ' Do normal step over
               END IF                                             '
            ELSE                                                  '
               INCR Letter: INCR i                                ' Do normal step over
            END IF                                                '
         LOOP                                                     '
         MExitMeth

      '----- End of word found, do KW lookup and build O/P control entry
      CloseWord:
         tClr = me.ClrKwSrch(MID$(@sTxt, StartPos, i - StartPos)) ' Save colour from word lookup (zero default)
         me.AttrSchemeSet(lno, StartPos, i, tClr)                 ' Build the Attr line
         '----- See if a Numeric 'word'
         IF tClr = %SCTxtLo THEN                                  ' If just a default scheme assigned by KwSrch
            IF VERIFY(MID$(@sTxt, Startpos, i - StartPos), "0123456789.,") = 0 THEN ' We have a numeric 'word'
               me.AttrSchemeSet(lno, StartPos, i, ClrNumeric)     ' Build the Attr line with the ClrNumeric value
            END IF                                                '
         END IF                                                   '
         wFlag = %False                                           ' Turn off the word flag
         RETURN                                                   '

      '----- Delim found, maybe a word end or a word start
      GotDLM:
         IF ISTRUE wFlag THEN                                     ' Working on a word?
            GOSUB CloseWord                                       ' Yes, go handle it
         END IF                                                   '
         StartPos = i: wFlag = %True                              ' Start over, call DLM a 'word'
         RETURN                                                   '

      '----- See if this line to be excluded
      DoExclude:
         FOR i = 1 TO 9                                           ' Loop through Exclude list
            IF TRIM$(ClrExclude(i).Txt1) = "" THEN RETURN         ' Done? return as unexcluded
            IF ClrExclude(i).sCol = 0 THEN                        ' A scan type exclude?
               IF INSTR(UUCASE(@sTxt), TRIM$(ClrExclude(i).Txt1)) <> 0 THEN GOSUB XcludeIt
            ELSE                                                  ' Fixed column test
               IF UUCASE(MID$(@sTxt, ClrExclude(i).sCol, LEN(TRIM$(ClrExclude(i).Txt1)))) = TRIM$(ClrExclude(i).Txt1) THEN GOSUB XcludeIt
            END IF                                                '
         NEXT i                                                   '
         RETURN

      '----- See if this line to be included
      DoInclude:
         FOR i = 1 TO 9                                           ' Loop through Include list
            IF TRIM$(ClrInclude(i).Txt1) = "" THEN GOSUB XcludeIt ' Didn't find it, exclude the line
            IF ClrInclude(i).sCol = 0 THEN                        ' A scan type include?
               IF INSTR(UUCASE(@sTxt), TRIM$(ClrInclude(i).Txt1)) <> 0 THEN RETURN
            ELSE                                                  ' Fixed column test
               IF IsEQ(MID$(@sTxt, ClrInclude(i).sCol, LEN(TRIM$(ClrInclude(i).Txt1))), TRIM$(ClrInclude(i).Txt1)) THEN RETURN
            END IF                                                '
         NEXT i                                                   '
         GOSUB Xcludeit                                           ' Didn't match, exclude it
         RETURN                                                   '

      XcludeIt:
         me.AttrSchemeSet(lno, 1, TxtLen, %SCTxtLo)               ' Build the Attr line
         MExitmeth                                                ' We're done

      END METHOD

      METHOD AttrSchemeGet(lno AS LONG, pcolumn AS LONG) AS LONG
      REGISTER i AS LONG
      REGISTER j AS LONG
      LOCAL AttrLine AS WSTRING POINTER
         METHOD = %SCTxtLo                                        ' Default answer
         j = LAttrSPtr(lno)                                       ' Get pointer to actual characters
         i = j + ((pcolumn - 1) * 2)                              ' Point at Attr char
         METHOD = PEEK(WORD, i) AND %AttrScheme                   ' Pass back Scheme number
      END METHOD

      METHOD AttrSchemeSet(lno AS LONG, fcolumn AS LONG, tcolumn AS LONG, scheme AS LONG)
      REGISTER i AS LONG
      REGISTER j AS LONG
      LOCAL lclUC, lclScheme AS WORD
         j = LAttrSPtr(lno)                                       ' Get pointer to actual characters
         IF scheme > 100 THEN                                     ' An UC theme?
            lclUC = %AttrUC: lclScheme = scheme - 100             ' Set that up
         ELSE                                                     '
            lclUC = &H0000: lclScheme = scheme                    '
         END IF                                                   '
         FOR i = j + ((fcolumn - 1) * 2) TO j + ((tcolumn - 1) * 2) STEP 2  ' Attr by Attr
            POKE WORD, i, (PEEK(WORD, i) AND (%AttrAll - %AttrScheme - %AttrUC)) OR (lclscheme OR lclUC) ' Add in the scheme numberr
         NEXT i                                                   '
      END METHOD

      METHOD AttrULClear(lno AS LONG, fcolumn AS LONG, tcolumn AS LONG)
      REGISTER i AS LONG
      REGISTER j AS LONG
         j = LAttrSPtr(lno)                                       ' Get pointer to actual characters
         FOR i = j + ((fcolumn - 1) * 2) TO j + ((tcolumn - 1) * 2) STEP 2  ' Attr by Attr
            POKE WORD, i, (PEEK(WORD, i) AND (%AttrAll - %AttrUL)) ' Remove the UL flag
         NEXT i                                                   '
      END METHOD

      METHOD AttrULSet(lno AS LONG, fcolumn AS LONG, tcolumn AS LONG)
      REGISTER i AS LONG
      REGISTER j AS LONG
         j = LAttrSPtr(lno)                                       ' Get pointer to actual characters
         FOR i = j + ((fcolumn - 1) * 2) TO j + ((tcolumn - 1) * 2) STEP 2  ' Attr by Attr
            POKE WORD, i, (PEEK(WORD, i) OR %AttrUL)              ' Add the INV flag
         NEXT i                                                   '
      END METHOD

      METHOD AFIsFileDir(i AS LONG) AS LONG
         METHOD = ISTRUE (AFList(i).Flag = %FDirDown OR AFList(i).Flag = %FEntry OR AFList(i).Flag = %FFLEntry)
      END METHOD

      METHOD AFIsReadOnly(i AS LONG) AS LONG
         METHOD = ISTRUE ((AFList(i).FD.FileAttributes AND %FILE_ATTRIBUTE_READONLY) = %FILE_ATTRIBUTE_READONLY)
      END METHOD

      METHOD   LoadData()
      '---------- Load the Actual File list from the Request List
      REGISTER i AS LONG
      REGISTER j AS LONG
      LOCAL tfn, lclPath, t, RQPath, RQMask, RQFlags, RQNote AS STRING, FD AS DIRDATA
      LOCAL NumExcl, NumForgotten, CurrTab AS LONG
         MEntry
         RESET AFCount, NumDirs, NumFiles, FMMarkdSLin            ' Reset stuff

         '----- Setup custom sorts
         SELECT CASE LONG FMode
            CASE %FMPath                                          ' FilePath
               DirSort = ENV.GetFMSort("FILEPATH", 1)             ' Fetch the Dir sort
               DefSort = ENV.GetFMSort("FILEPATH", 2)             ' Fetch the Def sort
            CASE %FMOpen                                          ' Open files
               DirSort = ENV.GetFMSort("OPEN", 1)                 ' Fetch the Dir sort
               DefSort = ENV.GetFMSort("OPEN", 2)                 ' Fetch the Def sort
            CASE %FMNamed                                         ' FLISTS
               DirSort = ENV.GetFMSort("FLISTS", 1)               ' Fetch the Dir sort
               DefSort = ENV.GetFMSort("FLISTS", 2)               ' Fetch the Def sort
            CASE %FMRecPaths                                      ' Paths
               DirSort = ENV.GetFMSort("PATHS", 1)                ' Fetch the Dir sort
               DefSort = ENV.GetFMSort("PATHS", 2)                ' Fetch the Def sort
            CASE %FMProfiles                                      ' Profiles
               DirSort = ENV.GetFMSort("PROFILES", 1)             ' Fetch the Dir sort
               DefSort = ENV.GetFMSort("PROFILES", 2)             ' Fetch the Def sort
            CASE %FMFileList                                      ' A Filelist type
               SELECT CASE AS CONST$ FileListNm                   ' Which type
                  CASE "Recent Files"                             ' Recent
                     DirSort = ENV.GetFMSort("RECENT", 1)         ' Fetch the Dir sort
                     DefSort = ENV.GetFMSort("RECENT", 2)         ' Fetch the Def sort
                  CASE "Found Files"                              ' Found
                     DirSort = ENV.GetFMSort("FOUND", 1)          ' Fetch the Dir sort
                     DefSort = ENV.GetFMSort("FOUND", 2)          ' Fetch the Def sort
                  CASE "Favorite Files"                           ' Favorites
                     DirSort = ENV.GetFMSort("FAVORITES", 1)      ' Fetch the Dir sort
                     DefSort = ENV.GetFMSort("FAVORITES", 2)      ' Fetch the Def sort
                  CASE ELSE                                       ' Any others
                     DirSort = ENV.GetFMSort("FILEPATH", 1)       ' Fetch the Dir sort
                     DefSort = ENV.GetFMSort("FILEPATH", 2)       ' Fetch the Def sort
               END SELECT                                         '
            CASE ELSE                                             '
               DirSort = ENV.GetFMSort("FILEPATH", 1)             ' Fetch the Dir sort
               DefSort = ENV.GetFMSort("FILEPATH", 2)             ' Fetch the Def sort
         END SELECT                                               '

         IF FMode = %FMFilelist OR FMode = %FMPath THEN           ' One of the DIR list types?

            '----- Extract data from the file directories
            FOR i = 1 TO RQCount                                  ' Loop through all RQList items
               me.RQSplit(RQList(i), RQPath, RQMask, RQFlags, RQNote) ' Get request parameter operands
               IF INSTR(RQFlags, "G") THEN ITERATE FOR            ' Ignore generic entries
               IF RIGHT$(RQPath, 1) = "\" AND FileListNm = "" THEN' A PathName while doing a normal path?
                  IF ISFALSE me.IsAFDup(RQPath + "..\") THEN      ' If not already there
                     me.IncAFSize                                 ' Allocate a new entry
                     INCR NumDirs                                 ' Count as dirs
                     GOSUB BlankBasic                             ' Reset everything
                     AFList(AFCount).RQIX = i                     ' Save RQ Index
                     AFList(AFCount).Path = RQPath                ' Save Path name
                     lclPath = RQPath                             ' Save the path
                     AFList(AFCount).FD.FileName = "..\"          ' Stuff in ..\
                     AFList(AFCount).Flag = %FDirUp               ' Set DirUp entry type
                     RQPath += "*"                                ' Add * if a pathname
                  END IF                                          '
                  RQPath += "*"                                   ' Add * if a pathname
               ELSE                                               '
                  lclPath = LEFT$(RQPath, INSTR(-1, RQPath, "\")) ' Extract path from the full name
               END IF                                             '

               RQPath = DIR$(RQPath, 22 TO FD)                    ' Look for the first entry

               '----- Loop through the returned entries
               DO WHILE ISNOTNULL(RQPath)                         ' While we're getting entries
                  IF (FD.FileAttributes AND %FILE_ATTRIBUTE_DIRECTORY) = %FILE_ATTRIBUTE_DIRECTORY THEN
                     IF ISFALSE me.IsAFDup(lclPath + TRIM$(FD.FileName) + "\") THEN  ' If not already there
                        me.IncAFSize                              ' Allocate a new entry
                        INCR NumDirs                              ' Count as dirs
                        GOSUB BlankBasic                          ' Reset everything
                        AFList(AFCount).RQIX = i                  ' Save RQ Index
                        AFList(AFCount).FD = FD                   ' Copy the FD data into it
                        AFList(AFCount).FD.FileName = TRIM$(AFList(AFCount).FD.FileName) + "\" 'Add \ to indicate directory
                        AFList(AFCount).Flag = %FDirDown          ' Set DirDown entry type
                        AFList(AFCount).LWTime = sMakePrettyTime(FD.LastWriteTime) ' Format LWTime
                        AFList(AFCount).Path = lclPath            ' Add the path
                     END IF

                  ELSE                                            ' It's a File Entry
                     IF ISTRUE me.TestMask(TRIM$(FD.FileName), RQMask) AND _' If we pass the Mask test
                        ISFALSE me.TestForget(TRIM$(lclPath) + TRIM$(FD.FileName)) AND _        ' And the Forget test
                        ISFALSE me.TestExclude(TRIM$(lclPath) + TRIM$(FD.FileName)) THEN        ' And the Exclude test
                        IF ISFALSE me.IsAFDup(lclPath + TRIM$(FD.FileName)) THEN  ' If not already there
                           me.IncAFSize                           ' Allocate a new entry
                           INCR NumFiles                          ' Count as dirs
                           GOSUB BlankBasic                       ' Reset everything
                           AFList(AFCount).RQIX = i               ' Save RQ Index
                           AFList(AFCount).Path = lclPath         ' Add the path
                           AFList(AFCount).FD = FD                ' Copy the FD data into it
                           AFList(AFCount).Flag = IIF(FMode = %FMPath, %FEntry, %FFLEntry)  ' Set the file entry type
                           AFList(AFCount).Note = me.NoteSearch(lclPath + TRIM$(FD.FileName)) ' Fetch Mote data (if any)
                           AFList(AFCount).LWTime = sMakePrettyTime(FD.LastWriteTime) ' Format LWTime
                           AFList(AFCount).Ext = UUCASE(MID$(TRIM$(FD.FileName), INSTR(-1,TRIM$(FD.FileName), "."))) ' Extract extension
                           AFList(AFCount).LinesInt = sGetLines(lclPath + TRIM$(FD.FileName)) ' Get Line count from STATE

                           IF INSTR(AFList(AFCount).Ext, "\") THEN AFList(AFCount).Ext = SPACE$(32)
                        END IF                                    '
                     END IF                                       '

                  END IF                                          '
                  RQPath = DIR$(NEXT, TO FD)                      ' Get next FILELIST entry
               LOOP                                               '

            NEXT i ' FM.RQCount                                   '

         ELSEIF FMode = %FMRecPaths THEN                          ' Recent paths?

            '----- Extract data from the file directories
            FOR i = 1 TO RQCount                                  ' Loop through all RQList items
               me.RQSplit(RQList(i), RQPath, RQMask, RQFlags, RQNote) ' Get request parameter operands

               '----- Handle old or new format FILELIST entry
               IF RIGHT$(RQPath, 4) = "\*.*" THEN                 ' Old format path FILELIST?
                  RQPath = LEFT$(RQPath, LEN(RQPath) - 3)         ' Remove the *.*, leave the trailing \
               ELSEIF RIGHT$(RQPath, 1) <> "\" THEN               ' If no trailing \ ignore it
                  ITERATE FOR                                     '
               END IF

               IF ISFALSE ISFOLDER(RQPath) THEN ITERATE FOR       ' Ignore if not a real folder

               me.IncAFSize                                       ' Allocate a new entry
               INCR NumDirs                                       ' Count as dirs
               GOSUB BlankBasic                                   ' Reset everything
               AFList(AFCount).RQIX = i                           ' Save RQ Index
               AFList(AFCount).FD.FileName = RQPath               ' Build an entry
               AFList(AFCount).Flag = %FPath                      ' Set RPath entry type
               AFList(AFCount).Ext = UUCASE(MID$(TRIM$(AFList(AFCount).FD.FileName), INSTR(-1,TRIM$(AFList(AFCount).FD.FileName), "."))) ' Extract extension
               IF INSTR(AFList(AFCount).Ext, "\") THEN AFList(AFCount).Ext = SPACE$(32)

            NEXT i ' FM.RQCount                                   '

         ELSEIF FMode = %FMNamed THEN                             ' Favourite FILELISTS?

            '----- Extract data from the file directories
            tfn = DIR$(ENV.FileListPath + "*.FLIST" TO FD)        ' Look for other FILELIST files
            DO WHILE ISNOTNULL(tfn)                               ' While we're getting entries
               IF (FD.FileAttributes AND %FILE_ATTRIBUTE_DIRECTORY) <> %FILE_ATTRIBUTE_DIRECTORY THEN
                  IF me.TestExclude(ENV.FileListPath + TRIM$(FD.FileName)) THEN GOTO NextFList ' And the Exclude test
                  IF UUCASE(RIGHT$(TRIM$(FD.FileName), 6)) = ".FLIST" THEN  ' Any other filelists?
                     me.IncAFSize                                 ' Allocate a new entry
                     INCR NumFiles                                ' Count as dirs
                     GOSUB BlankBasic                             ' Reset everything
                     AFList(AFCount).FD = FD                      ' Copy the FD data into it
                     AFList(AFCount).Flag = %FFileList            ' Set as FILELIST entry type
                     AFList(AFCount).LWTime = sMakePrettyTime(FD.LastWriteTime) ' Format LWTime
                     AFList(AFCount).Path = ENV.FileListPath      ' Add the path
                     AFList(AFCount).Ext = UUCASE(MID$(TRIM$(AFList(AFCount).FD.FileName), INSTR(-1,TRIM$(AFList(AFCount).FD.FileName), "."))) ' Extract extension
                     IF INSTR(AFList(AFCount).Ext, "\") THEN AFList(AFCount).Ext = SPACE$(32)
                  END IF                                          '
               END IF                                             '
               NextFList:
               tfn = DIR$(NEXT, TO FD)                            ' Get next entry
            LOOP                                                  '

         ELSEIF FMode = %FMProfiles THEN                          ' Profiles?

            '----- Extract data from the file directories
            tfn = DIR$(ENV.ProfPath + "*.INI" TO FD)              ' Look for INI files
            DO WHILE ISNOTNULL(tfn)                               ' While we're getting entries
               IF (FD.FileAttributes AND %FILE_ATTRIBUTE_DIRECTORY) <> %FILE_ATTRIBUTE_DIRECTORY THEN
                  me.IncAFSize                                    ' Allocate a new entry
                  INCR NumFiles                                   ' Count as dirs
                  GOSUB BlankBasic                                ' Reset everything
                  AFList(AFCount).FD = FD                         ' Copy the FD data into it
                  AFList(AFCount).Flag = %FProfile                ' Set as Profile entry type
                  AFList(AFCount).LWTime = sMakePrettyTime(FD.LastWriteTime) ' Format LWTime
                  AFList(AFCount).Path = ENV.ProfPath             ' Add the path
                  AFList(AFCount).Ext = UUCASE(MID$(TRIM$(AFList(AFCount).FD.FileName), INSTR(-1,TRIM$(AFList(AFCount).FD.FileName), "."))) ' Extract extension
                  IF INSTR(AFList(AFCount).Ext, "\") THEN AFList(AFCount).Ext = SPACE$(32)
               END IF                                             '
               tfn = DIR$(NEXT, TO FD)                            ' Get next entry
            LOOP                                                  '

         ELSEIF FMode = %FMOpen THEN                              ' OPEN files?

            '----- Fetch Open Files
            FOR i = 1 TO UBOUND(gFQ())                            ' Process the File Queue
               IF gFQ(i).gInUse = %False THEN ITERATE FOR         ' Skip if it no longer in use
               me.IncAFSize                                       ' Allocate a new entry
               INCR NumFiles                                      ' Count as a file
               GOSUB BlankBasic                                   ' Reset everything
               AFList(AFCount).Uniq = gFQ(i).gPgNumber            ' Uniq to hide the tab number
               AFList(AFCount).Path = gFQ(i).gWatchDir + "\"      '
               AFList(AFCount).FD.FileName = MID$(gFQ(i).gWatchFile, INSTR(-1, gFQ(i).gWatchFile, "\") + 1)    '
               AFList(AFCount).Flag = %FOpen                      ' Set open file entry type
               AFList(AFCount).Ext = UUCASE(MID$(TRIM$(AFList(AFCount).FD.FileName), INSTR(-1,TRIM$(AFList(AFCount).FD.FileName), "."))) ' Extract extension
               IF INSTR(AFList(AFCount).Ext, "\") THEN AFList(AFCount).Ext = SPACE$(32)
            NEXT i                                                '

            '----- Now pick up the Special tabs
            CurrTab = PgNumber                                    ' Save current tab number
            FOR i = TabsNum TO 1 STEP - 1                         ' Now look for Special Tabs
               TP = Tabs(i)                                       ' Pick the Tab
               IF IsClip THEN                                     ' Look for specials
                  me.IncAFSize                                    ' Allocate a new entry
                  INCR NumFiles                                   ' Count as a file
                  GOSUB FillSpecial                               ' Complete special entry
                  AFList(AFCount).FD.FileName = "(Clip)"          '
               ELSEIF IsSetEdit THEN                              '
                  me.IncAFSize                                    ' Allocate a new entry
                  INCR NumFiles                                   ' Count as a file
                  GOSUB FillSpecial                               ' Complete special entry
                  AFList(AFCount).FD.FileName = "(SET Edit)"      '
               ELSEIF ISNULL(TP.TIPFilePath) AND ISFALSE IsFMTab THEN '
                  me.IncAFSize                                    ' Allocate a new entry
                  INCR NumFiles                                   ' Count as a file
                  GOSUB FillSpecial                               ' Complete special entry
                  AFList(AFCount).FD.FileName = $Empty            '
               ELSEIF LEFT$(TIP.File, 1) = "*" THEN               ' CLONEd tab?
                  me.IncAFSize                                    ' Allocate a new entry
                  INCR NumFiles                                   ' Count as a file
                  GOSUB FillSpecial                               ' Complete special entry
                  AFList(AFCount).FD.FileName = LSET$(TIP.FilePath, FM_File_Size)
               END IF                                             '
            NEXT i                                                '
            TP = Tabs(CurrTab)                                    ' Restore current tab
         END IF                                                   '

         '----- Add total line
         me.IncAFSize                                             ' Allocate a new entry
         GOSUB BlankBasic                                         '
         AFList(AFCount).Flag = %FTotals                          ' Set as Total Line
         t = "-- End of List"                                     ' Build a last line marker
         IF NumDirs <> 0 THEN                                     '
            t +=  ", " + FORMAT$(NumDirs) + IIF$(NumDirs = 1, " Dir", " Dirs")
         END IF                                                   '
         IF NumFiles <> 0 THEN                                    ' Adding Files?
            t += ", " + FORMAT$(NumFiles) + IIF$(NumFiles = 1, " File", " Files")
            IF (FileListNm = "Recent Files" OR FileListNm = "Recent Paths") AND NumFiles = ENV.RecentCtr THEN ' Warn of RECENT list full
               t += " (Max)"                                      '
            END IF                                                '
         END IF                                                   '
         IF RXCount <> 0 THEN                                     ' Adding Excluded?
            t += ", " + FORMAT$(RXCount) + " Excluded"            '
         END IF                                                   '
         IF RFCount <> 0 THEN                                     ' Adding Excluded?
            t += ", " + FORMAT$(RFCount) + " Forgotten"           '
         END IF                                                   '
         t += " --"                                               '
         AFList(AFCount).FD.FileName = t                          ' Save in FileName

         SELECT CASE AS CONST$ TP.DefSort                         ' See how to sort it
            CASE "Name+"                                          ' Name Ascending
               ARRAY SORT AFList() FOR AFCount, CALL sFMSortFileUp() '
            CASE "Name-"                                          ' Name Descending
               ARRAY SORT AFList() FOR AFCount, CALL sFMSortFileDown()
            CASE "Extn+"                                          ' Ext Ascending
               ARRAY SORT AFList() FOR AFCount, CALL sFMSortExtUp()  '
            CASE "Extn-"                                          ' Ext Descending
               ARRAY SORT AFList() FOR AFCount, CALL sFMSortExtDown()'
            CASE "Date+"                                          ' Date Ascending
               ARRAY SORT AFList() FOR AFCount, CALL sFMSortDateUp() '
            CASE "Date-"                                          ' Date Descending
               ARRAY SORT AFList() FOR AFCount, CALL sFMSortDateDown()
            CASE "Size+"                                          ' Size Ascending
               ARRAY SORT AFList() FOR AFCount, CALL sFMSortSizeUp() '
            CASE "Size-"                                          ' Size Descending
               ARRAY SORT AFList() FOR AFCount, CALL sFMSortSizeDown()
            CASE "Lines+"                                         ' Lines Ascending
               ARRAY SORT AFList() FOR AFCount, CALL sFMSortLinesUp() '
            CASE "Lines-"                                         ' Lines Descending
               ARRAY SORT AFList() FOR AFCount, CALL sFMSortLinesDown()
            CASE "Note+"                                          ' Note Ascending
               ARRAY SORT AFList() FOR AFCount, CALL sFMSortNoteUp() '
            CASE "Note-"                                          ' Note Descending
               ARRAY SORT AFList() FOR AFCount, CALL sFMSortNoteDown()
         END SELECT                                               '

         DoClear(%LoadData)                                       ' Say we've done it
         DoSet(%Refresh)                                          ' New data always means a Refresh

         TopScrn = IIF(NestTopScrn > 0, NestTopScrn, TopScrn)     ' Set top to whatever Nest Pop wants
         IF NestCRow = 0 THEN                                     ' If no Nest Pop request
            me.CsrRow = 1: me.CsrCol = 11: COffset = 0            ' Reset things
         ELSE                                                     ' Else set Nest value
            COffset = 0: me.CsrRow = NestCRow: me.CsrCol = IIF(NestCRow = 1, 11, 1) ' Reset things
         END IF                                                   '
         RESET NestTopScrn, NestCRow                              '

         MExitMeth                                                '

         FillSpecial:
            GOSUB BlankBasic                                      ' Reset everything
            AFList(AFCount).Uniq =  i                             ' Uniq to hide the tab number
            AFList(AFCount).Flag = %FOpen                         ' Set Open file entry type
         RETURN

         BlankBasic:
            AFList(AFCount).FD.FileSizeLow = 0                    ' Clear fields
            AFList(AFCount).FD.FileSizeHigh = 0                   '
            AFList(AFCount).SizeInt = 0                           '
            AFList(AFCount).Flag = 0                              '
            AFList(AFCount).CmdOff = 0                            '
            AFList(AFCount).NoteOff = 0                           '
            AFList(AFCount).Uniq = 0                              '
            AFList(AFCount).Ext = SPACE$(32)                      '
            AFList(AFCount).Msg = SPACE$(64)                      '
            AFList(AFCount).Size = SPACE$(15)                     '
            AFList(AFCount).LWTime = SPACE$(17)                   '
            AFList(AFCount).Note = " "                            '
            AFList(AFCount).Path = " "                            '
            AFList(AFCount).Cmd = " "                             '
            AFList(AFCount).RQIX = 0                              '
            AFList(AFCount).LinesInt = -1                         '
         RETURN

      END METHOD

      METHOD   LoadFileList(flist AS STRING) AS LONG
      '---------- Load the Actual File list from the Request List
      REGISTER i AS LONG
      REGISTER j AS LONG
      LOCAL table(), lclMask, t, RQPath, RQMask, RQFlags, RQNote AS STRING
      LOCAL TableCount AS LONG
      LOCAL LIO AS iIO                                            ' For our I/O stuff
         MEntry
         LET LIO = CLASS "cIO"                                    '
         GOSUB LoadData                                           ' Load raw FILELIST data
         lclmask = "*"                                            ' Default to a mask of *
         IF TableCount = 0 THEN _                                 ' If empty, we had an error
            scError(%eFail, "FILELIST (" + flist + ") was empty"): METHOD = %True: MexitMeth ' Issue msg, just exit

         '----- Now process what's left in the table
         FOR i = 1 TO TableCount                                  ' Each line of the FILELIST file
            IF LEFT$(Table(i), 10) = "FileSrch: " THEN            ' The magic word?
               FFSearch = MID$(Table(i), 11)                      ' Save the FF command
            END IF                                                '
            me.RQSplit(Table(i), RQPath, RQMask, RQFlags, RQNote) ' Split up the operands

            me.IncRQSize                                          ' Allocate an RQ slot
            IF ISNULL(RQMask) THEN                                ' If no RQMask, use lclMask
               RQList(RQCount) = BUILD$(RQPath, "|", lclMask, "|", RQFlags, "|", RQNote)  ' Create an entry
            ELSE                                                  '
               RQList(RQCount) = BUILD$(RQPath, "|", RQMask, "|", RQFlags, "|", RQNote)  ' Create an entry
            END IF                                                '

            IF INSTR(RQFlags, "F") THEN                           ' This is a FORGET entry
               me.IncRFSize                                       ' Allocate an RF slot
               RFList(RFCount) = RQPath                           ' Create a Forget entry
            END IF                                                '

            IF ISNOTNULL(TRIM$(RQNote)) THEN                      ' This have a Note item?
               me.IncRNSize                                       ' Allocate an RN slot
               RNList(RNCount) = RQPath + "|" + RQNote            ' Create a Note entry
            END IF                                                '

         NEXT i ' TableCount                                      '

         METHOD = %False                                          ' Say we were successful
         MExitMeth

         LoadData:                                                ' Load the FILELIST data
            LIO.Setup("IE", "", "", ENV.FileListPath + flist + ".FLIST") ' Tell IO what we're opening
            IF LIO.EXEC THEN                                      '
               scError(%eFail, "FILELIST (" + flist + ".FLIST - " + LIO.ResultMsg)  ' Oops
               METHOD = %True: MExitMeth                          ' Say we've failed
            END IF                                                '
            TableCount = 0                                        ' Say no records
            FILESCAN # LIO.FNum, RECORDS TO i                     ' Get the number of records
            IF i > 0 THEN                                         ' Some records?
               REDIM table(1 TO i) AS STRING                      ' Redim array to match save data
               LINE INPUT # LIO.FNum, table() TO j                ' Read it all
            END IF                                                '
            LIO.Close                                             ' Close the FBO
            TableCount = j                                        ' Pass back record count
            FOR i = 1 TO j                                        ' Convert old comma delimmited stuff
               IF INSTR(table(i), "|") = 0 THEN                   ' If not yet in the new | format
                  REPLACE ANY "," WITH "|" IN table(i)            ' Swap commas to |
               END IF                                             '
            NEXT i                                                '
         RETURN

      END METHOD

      METHOD   LoadReq(PATH AS STRING, mask AS STRING, flist AS STRING)
      '---------- Load the Request list
      REGISTER i AS LONG
      REGISTER j AS LONG
         MEntry
         RESET RQCount, RFCount, RXCount, RNCount
         REDIM RQList(1 TO 100) AS INSTANCE STRING
         REDIM RFList(1 TO 100) AS INSTANCE STRING
         REDIM RXList(1 TO 100) AS INSTANCE STRING
         REDIM RNList(1 TO 100) AS INSTANCE STRING
         REDIM AFList(1 TO 100) AS INSTANCE FMFList
         RESET FFSearch
         DoSet(%LoadData)                                         ' LoadReq always requires LoadData
         SELECT CASE AS CONST$ TRIM$(flist)                       ' Set our initial mode
            CASE ""                                               ' No FileListNm
               FMode = %FMPath                                    ' Path/Mask mode
               RQList(1) =  TRIM$(PATH) + "|" + TRIM$(mask)       ' Create a '1 line' FILELIST
               RQCount = 1                                        '
            CASE "Recent Paths"                                   '
               FMode = %FMRecPaths                                ' Recent Paths
               me.LoadFileList(flist)                             ' Go Load the FileList
            CASE "Open Files"                                     '
               FMode = %FMOpen                                    ' Open files
            CASE "Named Favorites"                                '
               FMode = %FMNamed                                   ' Favorites
            CASE "Profiles"                                       '
               FMode = %FMProfiles                                ' Profiles
            CASE ELSE                                             '
               FMode = %FMFileList                                ' FILELIST format
               me.LoadFileList(flist)                             ' Go Load the FileList
         END SELECT                                               '
         DoClear(%LoadReq)                                        ' Say we've done it
         TopScrn = 1                                              ' We're going to be at the top
         me.CsrRow = 1: me.CsrCol = 11: COffset = 0               ' Reset things
         MExit
      END METHOD

      METHOD LoadTheText(fn AS STRING, Prof AS STRING)
      '--------- Load the text either in this tab or a new one
         MEntry
         '----- Decide whether a new tab is needed or not
         IF LastLine > 2  THEN                                    ' We need a new tab, there are lines present
            '---- Open this in a new tab
            sTabAdd(fn, Prof)                                     ' Yes, let sTabAdd do the work
            gTabSwitch = TP.PgNumber                              ' Set to 'switch' there

         ELSE                                                     '
            '----- Open this in the current tab
            TIP.Setup("B", Prof, "", fn)                          ' Setup for Open
            me.InitaFile(%False)                                  ' Go load it in
            me.UndoInit                                           ' Init the Undo file names
            me.UndoSave()                                         ' Take an initial one
            me.WindowTitle                                        ' Alter window/Tab titles
            me.SetStart()                                         ' Do initial positioning
            gTabSwitch = PgNumber                                 ' Set to 'switch' here
         END IF                                                   '
         MExit
      END METHOD

      METHOD BndsSave()
      '---------- Save the BNDS line by updating the INI file
      LOCAL Wrd1, Wrd2 AS STRING, i, j, k, l, LeftCtr, RightCtr, PlusCtr AS LONG
         MEntry
         OffBndsAFlag                                             ' Clear flag that got us here
         '----- Validate the BNDS line contents
         Wrd1 = Prf.GoodBnds                                      ' Wrd1 = last good bounds
         Wrd2 = TP.BndText                                        ' Lets test it
         IF ISNULL(TRIM$(Wrd2)) OR TRIM$(Wrd2) = "<" THEN Wrd2 = "<+"  ' If null'ed or just a <, stuff in a default

         LeftCtr = TALLY(Wrd2, "<")                               ' Count important chars
         RightCtr = TALLY(Wrd2, ">")                              '
         PlusCtr = TALLY(Wrd2, "+")                               '

         Wrd2 = REMOVE$(Wrd2, ANY " <>+")                         ' Clear all valid characters
         IF Wrd2 <> "" THEN _                                     ' Test for garbage characters
            MErrExit(%eFail, "Invalid or extra characters in BNDS line")
         Wrd2 = TP.BndText                                        ' Start over
         IF ISNULL(TRIM$(Wrd2)) OR TRIM$(Wrd2) = "<" THEN Wrd2 = "<+"  ' If null'ed or just a <, stuff in a default

         IF (LeftCtr = 1 AND RightCtr = 1 AND PlusCtr = 0) OR _   ' Look for valid combination   < >
            (LeftCtr = 1 AND RightCtr = 0 AND PlusCtr = 1) THEN   '      or                      < +
            '                                                     '
         ELSE                                                     '
            MErrExit(%eFail, "Invalid combination of characters in BNDS line")
         END IF                                                   '

         Wrd2 = REMOVE$(Wrd2, ANY "+")                            ' Clear any MAX characters

         '----- We seem to have a valid BNDS line in Wrd2
         IF LeftCtr = 1 AND RightCtr = 1 THEN                     ' Normal L/R bounds
            Prf.BndLeft = INSTR(Wrd2,"<"): Prf.BndRight = INSTR(Wrd2, ">") ' Save the values
         ELSEIF LeftCtr = 1 AND PlusCtr = 1 THEN                  ' Normal L/Max
            Prf.BndLeft = INSTR(Wrd2,"<"): Prf.BndRight = 0       ' Save the values
         END IF                                                   '
         TP.BndText = Wrd2                                        ' Replace working copy
         IF ISFALSE gMacroMode THEN
            sDoStatusBar($SBBnds)                                 ' re-Do the StatusBar BNDS box
         END IF
         MExit                                                    ' We're done
      END METHOD

      METHOD  CallTab(cmd AS STRING, oprnd AS STRING)
      '---------- Call another Tab and return
      LOCAL OurTab AS LONG, lclFPath, lclFMask, lclFListNm AS STRING
         MEntry
         OurTab = TP.PgNumber                                     ' Save our page number
         lclFPath = ENV.FMPath: lclFMask = ENV.FMMask: lclFListNm = ENV.FMFileList ' Save around calls
         ENV.FMPath = FPath                                       ' Pass our FM variables to the called command
         ENV.FMMask = FMask                                       '
         ENV.FMFileList = FileListNm                              '
         SELECT CASE AS CONST$ cmd                                ' Call the correct guy
            CASE "BROWSE":   pCmdBROWSE(cmd + " " + oprnd)        '
            CASE "CANCEL":   pCmdCANCEL(cmd + " " + oprnd)        '
            CASE "CLONE":    pCmdCLONE(cmd + " " + oprnd)         '
            CASE "EDIT":     pCmdEDIT(cmd + " " + oprnd)          '
            CASE "END":      pCmdEND(cmd + " " + oprnd)           '
            CASE "MEDIT":    pCmdMEDIT(cmd + " " + oprnd)         '
            CASE "OPEN":     pCmdOPEN(cmd + " " + oprnd)          '
            CASE "OPENV":    pCmdOPEN(cmd + " " + oprnd)          '
            CASE "OPENB":    pCmdOPEN(cmd + " " + oprnd)          '
            CASE "SAVE":     pCmdSAVE(cmd + " " + oprnd)          '
            CASE "VIEW":     pCmdVIEW(cmd + " " + oprnd)          '
            CASE "XSUBMIT":  pCmdXSUBMIT(cmd + " " + oprnd)       '
         END SELECT                                               '
         ENV.FMPath = lclFPath: ENV.FMMask = lclFMask: ENV.FMFileList = lclFListNm ' Restore things
         IF gTabSwitch <> OurTab THEN                             '
            gTabSwitchMsg = CallMsg                               ' If action occurred in another tab
         ELSE                                                     '
            IF ISNOTNULL(CallMsg) THEN scError(4, CallMsg)        ' Issue message in the other tab
         END IF                                                   '
         TP = Tabs(OurTab)                                        ' Switch back to original tab
         CallMsg = ""                                             ' Clear CallMsg
         MExit                                                    ' We're done
      END METHOD

      METHOD CursorWord()
      '---------- Extract 'word' under the cursor if possible
      REGISTER i AS LONG
      REGISTER j AS LONG
      LOCAL DataIX AS LONG
      LOCAL lclWord, t AS STRING

         MEntry
         '----- Exit if not even in the data area
         IF ISFALSE IsCData THEN GOTO CursBailOut                 ' If not in data area then null things

         lclWord = Prf.PWord                                      ' Fetch a local copy
         DataIX = sGetIX(CRow)                                    ' Get the cursor line

         '----- Exit if not a data line
         IF DataIX < 1 THEN GOTO CursBailOut                      ' None, bail out
         IF ISFALSE IsLData(DataIX) THEN GOTO CursBailOut         ' Bail if not on a data line

         CursLine = LTxtG(DataIX)                                 ' Save the line the cursor's on
         t = CursLine                                             ' Temp copy
         i = CCol - gLNPadCol + Offset                            ' Calc index into string where cursor is located
         '----- Exit if cursor sitting on a delimiter
         IF INSTR(lclWord, MID$(t, i, 1)) = 0 THEN                ' If cursor is not sitting on a word character
            CursWord = "": MExitMeth                              ' Also bail out
         END IF                                                   '

         '----- Extract the word now, cursor in Col 1
         IF i = 1 THEN                                            ' Simple case (in column 1)
            j = 1                                                 ' Look for a delimiter
            WHILE INSTR(lclWord, MID$(t, j, 1)) <> 0 AND j <= LEN(t) '
               INCR j                                             '
            WEND                                                  ' j = delim or end + 1
            CursWord = LEFT$(t, j - 1)                            ' Then Cursword is 1 thru DLM (or end) -1

         '----- Extract the word now, cursor somewhere else
         ELSE                                                     ' We're not in column 1
            DO WHILE i                                            ' Backup to 1st DLM char
               IF INSTR(lclWord, MID$(t, i, 1)) <> 0 THEN         '
                  DECR i                                          ' Backup 1 char
               ELSE                                               '
                  INCR i: EXIT DO                                 ' i now equals 1st char of word
               END IF                                             '
            LOOP                                                  '
            i = MAX(1, i)                                         ' Minimum of 1
            j = i                                                 ' Look for a delimiter
            WHILE INSTR(lclWord, MID$(t, j, 1)) <> 0 AND j <= LEN(t) '
               INCR j                                             '
            WEND                                                  ' j = delim or end + 1
            CursWord = MID$(t, i TO j - 1)                        ' Then Cursword is i thru DLM (or end) -1
         END IF                                                   '
         MExitMeth                                                '

      CursBailOut:
            CursWord = "": CursLine = ""                          ' Null CursWord and CursLine
            MExitMeth                                             ' Exit
         RETURN                                                   '
      END METHOD

      METHOD CursWordLoc()
      '---------- Mark 'word' under the cursor if possible
      REGISTER i AS LONG
      REGISTER j AS LONG
      LOCAL DataIX AS LONG
      LOCAL lclWord, t AS STRING
         MEntry
         OffMarkActive                                            '

         '----- See if cursor is at least somewhere reasonable
         lclWord = Prf.PWord                                      ' Fetch a local copy
         DataIX = sGetIX(CRow)                                    ' Get the cursor line
         IF ISFALSE IsLData(DataIX) THEN MExitMeth                ' Bail if not on a data line
         t = LTxtG(DataIX)                                        ' Get the text
         i = CCol - gLNPadCol + Offset                            ' Calc index into string where cursor is located
         IF INSTR(lclWord, MID$(t, i, 1)) = 0 THEN MExitMeth      ' If cursor in a deimiter type area, also bail out

         '----- Set the MARK values
         MarkSLin = S(CRow)                                       '
         MarkELin = MarkELin                                      '
         OnMarkActive                                             '

         IF i <> 1 THEN                                           ' If not the Simple case (in column 1)
            DO WHILE i                                            ' Backup to 1st DLM char
               IF INSTR(lclWord, MID$(t, i, 1)) <> 0 THEN         ' On a word char?
                  DECR i                                          ' Backup 1 char
               ELSE                                               '
                  INCR i: EXIT DO                                 ' i now equals 1st char of word
               END IF                                             '
            LOOP                                                  '
            i = MAX(i, 1)                                         ' i must be at least 1
         END IF                                                   '
         j = i                                                    ' Look for a delimiter
         WHILE INSTR(lclWord, MID$(t, j, 1)) <> 0 AND j <= LEN(t) '
            INCR j                                                '
         WEND                                                     ' j = delim or end + 1

         MarkSCol = i                                             ' Remember starting x/y of Select frame
         MarkECol = j - 1                                         '

         '----- Don't collide with SWAP highliting
         IF IsSwapActive THEN                                     ' Prevent Swap overlap
            IF MarkSLin = SwapSLin THEN                           ' Same line?
               IF MarkSCol >= SwapSCol AND MarkSCol <= SwapECol THEN GOSUB ClrMark: MExitMeth
               IF MarkECol >= SwapSCol AND MarkECol <= SwapECol THEN GOSUB ClrMark: MExitMeth
            END IF                                                '
         END IF                                                   '
         MExitMeth

      ClrMark:
         MarkSLin = 0: MarkELin = 0                               ' Reset things
         MarkSCol = 0: MarkECol = 0                               '
         OffMarkActive                                            '
         RETURN                                                   '

      END METHOD

      METHOD DataInsert(iText AS STRING, iInsert AS STRING, i AS LONG) AS STRING
      '---------- Handle inserting a string into a string
      LOCAL x, y, GotColumns AS LONG, lclTxt, lh, lh2, rh AS STRING

         MEntry
         lclTxt = iText
         IF ISFALSE IsNsrtData THEN                               ' If not the complex DataInsert, do it quick and exir
            IF i > LEN(lclTxt) THEN lclTxt = LSET$(lclTxt, i)     ' Past the right end? Then lengthen string
            METHOD = STRINSERT$(lclTxt, iInsert, i)               ' Simple string insert
            MExitMeth                                             ' We're done
         END IF                                                   '

         '----- See if any 'columns' in the data
         x = INSTR(i, lclTxt, " ")                                ' Look for at least 2 blanks following found string
         IF x THEN                                                ' If we 've got spare blanks
            FOR y = x TO LEN(lclTxt)                              ' Look through remainder of line for a non- blank
               IF MID$(lclTxt, y, 1) <> " " THEN                  ' Got one?
                  GotColumns = %True                              ' Remember we have columns
                  EXIT FOR                                        ' Exit, y - 1 is split point for lh/rh
               END IF                                             '
            NEXT y                                                '
         END IF                                                   '

         '----- There are columns and DataInsert requested
         IF GotColumns THEN                                       ' Data Insert AND columns?
               lh = LEFT$(lclTxt, y - 2)                          ' Create full lh portion
               rh = MID$(lclTxt, y - 1)                           ' Create full rh portion
               lh2 = RTRIM$(MID$(lh, i))                          ' Create rh part of the full lh portion
               lh = LEFT$(lh, i - 1)                              ' Create lh part of the full lf portion
               lh = lh + iInsert + lh2                            ' Re-Build the full lh portion with the change
               IF LEN(lh) < y - 1 THEN                            ' If lh is less than orig length of lh portion
                  lh = lh + STRING$((y - 2) - LEN(lh), " ")       ' make it back up to original length
               END IF                                             '
               METHOD = lh + rh                                   ' Re-build the altered text

         '----- Normal non- DataInsert mode
         ELSE
            IF i > LEN(lclTxt) THEN lclTxt = LSET$(lclTxt, i)     ' Past the right end? Then lengthen string
            METHOD = STRINSERT$(lclTxt, iInsert, i)               ' Simple string insert

         END IF                                                   '
         MExit
      END METHOD


      METHOD DispANSI(BYVAL colMode AS LONG)
      '---------- Fire up ANSI window in a new thread
      ' (CharSet)    passes colMode = %False
      ' (CharSetCol) passes colMode = %True

      LOCAL hANSIThread, i AS LONG                                ' fptr no longer used

         THREAD CREATE DispANSI2(BYVAL colMode) 65536, TO hANSIThread

         IF hANSIThread = 0 THEN _                                ' Failed?
            MErrExit(%eFail, "Could not create ANSI Window")
         SLEEP 200                                                ' Wait a bit
         THREAD STATUS hANSIThread TO i                           ' See if running OK
         IF i <> 259 THEN _                                       ' If running OK STATUS returns &H103 (See Help)
            MErrExit(%eFail, "Could not run ANSI Window")
         THREAD CLOSE hANSIThread TO i                            ' Free up our handle
      END METHOD


      METHOD  DispFM()
      '---------- Display the File Manager Screen
      LOCAL d, e, HMask AS STRING, sp, stp AS LONG, CurSet AS INTEGER
      LOCAL i, j, k, pcol, plen, sClr1, sClr2, sClr3, sClr4, LastScrLine AS LONG
      LOCAL tText, tText1, tText2, u AS STRING, tfn AS WSTRINGZ * %MAX_PATH
      LOCAL lScrl AS INTEGER, lColumn, lLin, lLength, cOffset AS LONG
      LOCAL HDDir, HDName, HDExt, HDSize, HDDate, HDNote, HDLines, Layout, lField AS STRING
      LOCAL lclAttr AS WSTRING
         MEntry

         '----- Always get refreshed data if needed
         IF IsDoLoadReq THEN me.LoadReq(FPath, FMask, FileListNm) ' Maybe Load Requests if needed
         IF IsDoLoadData AND ISFALSE IsDoMsg THEN me.LoadData     ' Load Data if needed and no messages waiting
         DoClear(%Refresh)

         '----- Display the FM window
         GRAPHIC ATTACH PgHandle, WindowID, REDRAW                ' Set as the default graphic area
         cBandBG = %False                                         ' Clear any prior banding state

         sPrint ("Command > ", $$TxtHi, 1, 1)                     ' Re-do the command line
         IF ISNOTNULL(pCommandPrv) THEN                           ' Got an & command line?
            pCommand = pCommandPrv                                ' Swap it in
            pCommandPrv = ""                                      ' null it
         END IF                                                   '

         me.WindowCmd                                             '
         sPrint ("Scroll > ", $$TxtHi,  1, 11 + pCmdLen + 1)      '
         sPrint (LSET$(Prf.Scroll, 4), $$TxtLo, 1, 11 + pCmdLen + 9 + 1)  '
         ScrlAmtLast = Prf.Scroll                                 ' Save as last valid Scroll value

         sPrint (STRING$(ENV.ScrWidth, " "), $$TxtLo, 2, 1)       ' Print normal dash line
            IF ISFALSE gResizeActive  THEN                        ' Don't do messages during a RESIZE
               IF ErrMsgTblC > 0 THEN                             ' If there's an Error message?
                  REDIM ErrMsgHlp(1 TO ErrMsgTblC)                ' Set Hlp copy table correct size
                  FOR i = 1 TO ErrMsgTblC                         ' Copy it
                     ErrMsgHlp(i) = ErrMsgTbl(i)                  '
                  NEXT i                                          '
                  ErrMsgHlpC = ErrMsgTblC                         ' Copy count
                  IF ErrMsgTblC > 1 THEN                          ' More than 1 message waiting?
                     sPrint (" +" + LEFT$(ErrMsg, 70), $$Error, 2, ENV.ScrWidth - LEN(LEFT$(ErrMsg, 70)) - 1)  ' Print error message
                  ELSE                                            '
                     sPrint (" " + LEFT$(ErrMsg, 70), $$Error, 2, ENV.ScrWidth - LEN(LEFT$(ErrMsg, 70)))       ' Print error message
                  END IF                                          '
                  IF errFlag = %eFail THEN sDoBeep                ' Beep on errors
               ELSEIF ISNOTNULL(gGlblMessage) THEN                ' Got a global one?
                  sPrint (" " + gGlblMessage, $$Error, 2, ENV.ScrWidth - LEN(gGlblMessage))' Print error message
                  RESET gGlblMessage                              '
               END IF                                             '
            END IF                                                '
            RESET ErrFlag, ErrMsg, ErrMsgTbl(), ErrMsgTblC        ' Clear error status

         '----- Do the Quick entries
         '_ New  _ Norm Path  _ Recent  _ Found  _ Opened  _ Favorites  _ Lists  _ Paths  _ Profiles
         sPrint ("_ ", $$FMTool, FM_Quick_Line_1, FM_Quick_Pos_1) ' Print the quick entries
         sPrint ("New  ", $$FMTool, FM_Quick_Line_1, FM_Quick_Pos_1 + 2)
         sPrint ("_ ", $$FMTool, FM_Quick_Line_1, FM_Quick_Pos_2)
         sPrint ("FilePath  ", $$FMTool, FM_Quick_Line_1, FM_Quick_Pos_2 + 2)
         sPrint ("_ ", $$FMTool, FM_Quick_Line_1, FM_Quick_Pos_3)
         sPrint ("Recent  ", $$FMTool, FM_Quick_Line_1, FM_Quick_Pos_3 + 2)
         sPrint ("_ ", $$FMTool, FM_Quick_Line_1, FM_Quick_Pos_4)
         sPrint ("Found  ", $$FMTool, FM_Quick_Line_1, FM_Quick_Pos_4 + 2)
         sPrint ("_ ", $$FMTool, FM_Quick_Line_1, FM_Quick_Pos_5)
         sPrint ("Opened  ", $$FMTool, FM_Quick_Line_1, FM_Quick_Pos_5 + 2)
         sPrint ("_ ", $$FMTool, FM_Quick_Line_1, FM_Quick_Pos_6)
         sPrint ("Favorites  ", $$FMTool, FM_Quick_Line_1, FM_Quick_Pos_6 + 2)
         sPrint ("_ ", $$FMTool, FM_Quick_Line_1, FM_Quick_Pos_7)
         sPrint ("Flists  ", $$FMTool, FM_Quick_Line_1, FM_Quick_Pos_7 + 2)
         sPrint ("_ ", $$FMTool, FM_Quick_Line_1, FM_Quick_Pos_8)
         sPrint ("Paths  ", $$FMTool, FM_Quick_Line_1, FM_Quick_Pos_8 + 2)
         sPrint ("_ ", $$FMTool, FM_Quick_Line_1, FM_Quick_Pos_9)
         sPrint ("Profiles  ", $$FMTool, FM_Quick_Line_1, FM_Quick_Pos_9 + 2)
         IF ENV.ScrWidth > FM_Quick_Pos_9 + 12 THEN
            sPrint (SPACE$(ENV.ScrWidth - FM_Quick_Pos_9 + 12), $$FMTool, FM_Quick_Line_1, FM_Quick_Pos_9 + 12)
         END IF

         IF ISNULL(FileListNm) THEN                               ' Normal FM Screen
            sPrint ("File Path/Name> ", $$TxtHi, FM_Path_Line, 1) '
            tText = LSET$(me.FileAbbrev(FPath, FM_Crit_Size), FM_Crit_Size)
            sPrint (tText, $$TxtLo, FM_Path_Line, FM_Path_Left)   '
            sPrint ("File Patterns > ", $$TxtHi, FM_Mask_Line, 1) '
            tText = LSET$(FMask, FM_Crit_Size)                    '
            sPrint (tText, $$TxtLo, FM_Mask_Line, FM_Mask_Left)   '

         '----- Do a FM FILELIST type heading
         ELSE                                                     '
            sPrint ("FLIST         > ", $$TxtHi, FM_Path_Line, 1) '
            IF FFSearch = "" THEN                                 ' Handle Found list differently
               sPrint (LSET$(FileListNm, FM_Crit_Size), $$TxtLo, FM_Path_Line, FM_Path_Left)
            ELSE                                                  '
               sPrint (LSET$(FileListNm + ": " + FFSearch, FM_Crit_Size), $$TxtLo, FM_Path_Line, FM_Path_Left)
            END IF                                                '
            sPrint ("FSPEC         > ", $$TxtHi, FM_Mask_Line, 1) ' Display the header
            sPrint (LSET$("*",FM_Crit_Size), $$TxtLo, FM_Mask_Line, FM_Mask_Left)
            IF RQCount = 1 THEN                                   ' If just 1 request
               tText = RQList(1)                                  ' Get the RQList item
               i = INSTR(tText, "|"): i = INSTR(i+1, tText, "|")  ' Find 2nd comma
               tText = LEFT$(tText, i - 1)                        ' Strip of junk user doesn't need to see
               i = INSTR(tText, "|")                              ' Find 1st comma now
               IF MID$(tText, i + 1) = "*" OR MID$(tText, i + 1) = "*.*" THEN tText =LEFT$(tText, i - 1)
               tText = LSET$(tText, FM_Crit_Size)                 '
               sPrint (tText, $$TxtLo, FM_Mask_Line, FM_Mask_Left)'
            END IF                                                '

         END IF

         '----- Setup the column header text
         HDDir = DirSort
         HDName = "Name": HDExt = "Ext":  HDSize = "Size":  HDDate = "Date":  HDNote = "Note": HDLines = "Lines" ' Set all to defaults
         SELECT CASE AS CONST$ DefSORT                            '
            CASE "Name-"    : HDName  = "Name-"
            CASE "Name+"    : HDName  = "Name+"
            CASE "UNSORTED" : '
            CASE "Extn-"    : HDExt   = "Ext-"
            CASE "Extn+"    : HDExt   = "Ext+"
            CASE "Size-"    : HDSize  = "Size-"
            CASE "Size+"    : HDSize  = "Size+"
            CASE "Date-"    : HDDate  = "Date-"
            CASE "Date+"    : HDDate  = "Date+"
            CASE "Note-"    : HDNote  = "Note-"
            CASE "Note+"    : HDNote  = "Note+"
            CASE "Lines+"   : HDLines = "Lines+"
            CASE "Lines-"   : HDLines = "Lines-"
         END SELECT

         '----- Do the column headings
         sPrint (REPEAT$(ENV.ScrWidth, " "), $$FMTool,  FM_Head_Line, 1) ' Blank the line on the screen
         sPrint (IIF$(ENV.FMLCmdWidth > 6, "Command", "Cmd"), $$FMTool, FM_Head_Line, 1)
         sPrint (HDDir + "  " + HDName, $$FMTool, FM_Head_Line, FM_Head_Name_Left)

         '----- Now add the Misc fields as requested
         Layout = ENV.FMLayout                                    ' Get working copy of layout
         cOffset = FM_Head_Name_Left + FM_File_Size + 1           ' Next available position in line
         FM_Head_Ext_Left = 0: FM_Head_Size_Left = 0: FM_Head_Date_Left = 0: FM_Head_Note_Left = 0: FM_Head_Lines_Left = 0
         DO WHILE ISNOTNULL(TRIM$(Layout))                        ' Step through the layout
            lField = EXTRACT$(Layout, ",")                        ' Get next operand
            Layout = REMAIN$(Layout, ",")                         ' Remove it from string
            SELECT CASE AS CONST$ lField                          ' See what field
               CASE "EXT"                                         '
                  FM_Head_Ext_Left = cOffset: cOffset += %FMLExt + 2 ' EXT + pad
                  sPrint (LSET$(HDExt, %FMLExt), $$FMTool, FM_Head_Line, FM_Head_Ext_Left)
               CASE "SIZESHORT"                                   '
                  FM_Head_Size_Left = cOffset: cOffset += %FMLSizeS + 2 ' Short size + pad
                  sPrint (RSET$(HDSize, %FMLSizeS), $$FMTool, FM_Head_Line, FM_Head_Size_Left)
               CASE "SIZELONG"                                    '
                  FM_Head_Size_Left = cOffset: cOffset += %FMLSizeL + 2 ' Long size + pad
                  sPrint (RSET$(HDSize, %FMLSizeL), $$FMTool, FM_Head_Line, FM_Head_Size_Left)
               CASE "DATE"                                        '
                  FM_Head_Date_Left = cOffset: cOffset += %FMLDate + 2  ' Date + pad
                  sPrint (LSET$(HDDate, %FMLDate), $$FMTool, FM_Head_Line, FM_Head_Date_Left)
               CASE "DATETIME"                                    '
                  FM_Head_Date_Left = cOffset: cOffset += %FMLDateT + 2 ' Date/Time size + pad
                  sPrint (LSET$(HDDate, %FMLDateT), $$FMTool, FM_Head_Line, FM_Head_Date_Left)
               CASE "LINES"                                       '
                  FM_Head_Lines_Left = cOffset: cOffset += %FMLLines + 2 ' Lines size + pad
                  sPrint (RSET$(HDLines, %FMLLines), $$FMTool, FM_Head_Line, FM_Head_Lines_Left)
               CASE "NOTE"                                        '
                  FM_Head_Note_Left = cOffset: cOffset += FM_Note_Size + 2  ' Note + pad
                  IF FMode = %FMFileList THEN                     ' Have an associated FILELIST?
                     sPrint (LSET$(HDNote, FM_Note_Size), $$FMTool, FM_Head_Line, FM_Head_Note_Left)
                  END IF                                          '
            END SELECT                                            '
         LOOP                                                     '

         '----- Loop through data lines on the screen
         stp = TopScrn                                            ' Set Screen Text Pointer to current TopScrn setting
         FOR sp = FM_Top_File_Line TO FM_Top_File_Line + FM_List_Height - 1 ' Loop for logical Screen data lines
            sCalcFMBG(sp)                                         ' Get Banding setup

            Rightmost = 0                                         ' Start over

            '----- Below last data line
            IF stp > AFCount THEN                                 ' Past end of dataset table?
               sPrint (REPEAT$(ENV.ScrWidth, " "), $$TxtLo, sp, 1)' Blank the line on the screen
               INCR stp                                           '
               ITERATE FOR                                        '

            ELSEIF AFList(stp).Flag = %FTotals THEN
               sPrint REPEAT$(ENV.FMLCmdWidth + 1, " "), $$TxtLo, sp, 1   '
               sPrint (LSET$(TRIM$(AFList(stp).FD.FileName), ENV.ScrWidth), $$TxtHi, sp, FM_Head_Name_Left) ' Dump total line
               LastScrLine = sp                                   ' Save last printed screen line
               INCR stp                                           '
               ITERATE FOR                                        '

            '----- A Normal data line
            ELSE                                                  '

               LastScrLine = sp                                   ' Save last printed screen line
               '----- Do the line command field
               IF AFList(stp).Flag <> %FTotals THEN               '
                  sPrint(me.WindowLCmd(stp), $$TxtLo, sp, 1)      '
               ELSE                                               '
                  sPrint (REPEAT$(ENV.FMLCmdWidth + 1, " "), $$TxtLo, sp, 1)
               END IF                                             '

               '----- Do the Filename / Extension or Message field
               IF ISNULL (TRIM$(AFList(stp).Msg)) THEN            ' Is there a message waiting?
                  me.FMPrtFileName(stp, sp, %False)               ' Go fiddle and print it (No forced HiLight)

               ELSE                                               ' We have a message instead of the filename
                  sPrint (LSET$(AFList(stp).msg, FM_File_Size+2), $$Error, sp, FM_Head_Name_Left)
                  Rightmost = FM_Head_Name_Left + FM_File_Size - 1'
                  DoSet(%LoadData)                                ' Trigger clear message
                  DOClear(%Msg)                                   ' Clear message flag
               END IF                                             '

               '----- Do the Size field
               IF FM_Head_Size_Left > 0 THEN                      ' Doing Size at all?
                  IF AFList(stp).Flag <> %FDirDown AND _          '
                     AFList(stp).Flag <> %FDirUp THEN
                     lclAttr = IIF$(LEFT$(DefSort, 4) = "Size", $$TxtHi, $$TxtLo)
                     IF INSTR(ENV.FMLayout, "SIZESHORT") THEN     ' Short size?
                        tText = sMakePrettySizeSmall(MAK(QUAD, AFList(stp).FD.FileSizeLow, AFList(stp).FD.FileSizeHigh)) + "  "
                        sPrint (tText, lclAttr, sp, FM_Head_Size_Left)
                        Rightmost = MAX(Rightmost, FM_Head_Size_Left + LEN(tText) - 1)
                     ELSE                                         ' Else Large size
                        tText = sMakePrettySizeLarge(MAK(QUAD, AFList(stp).FD.FileSizeLow, AFList(stp).FD.FileSizeHigh)) + "  "
                        sPrint (tText, lclAttr, sp, FM_Head_Size_Left)
                        Rightmost = MAX(Rightmost, FM_Head_Size_Left + LEN(tText) - 1)
                     END IF                                       '
                  ELSE                                            '
                     sPrint (IIF$(INSTR(ENV.FMLayout, "SIZESHORT"), SPACE$(%FMLSizeS + 2), SPACE$(%FMLSizeL + 2)), $$TxtLo, sp, FM_Head_Size_Left)
                  END IF                                          '
               END IF                                             '

               '----- Do the Lines field
               IF FM_Head_Lines_Left > 0 THEN                     ' Doing Lines at all?
                  IF AFList(stp).Flag <> %FDirDown AND _          '
                     AFList(stp).Flag <> %FDirUp THEN             '
                     lclAttr = IIF$(LEFT$(DefSort, 5) = "Lines", $$TxtHi, $$TxtLo)
                     tText = RIGHT$(FORMAT$(AFList(stp).LinesInt, "* #,###,###"), %FMLLines) + "  "
                     IF AFList(stp).LinesInt = -1 THEN tText = SPACE$(11)
                     sPrint (tText, lclAttr, sp, FM_Head_Lines_Left)
                     Rightmost = MAX(Rightmost, FM_Head_Lines_Left + LEN(tText) - 1)
                  ELSE                                            '
                     sPrint (SPACE$(%FMLLines + 2), $$TxtLo, sp, FM_Head_Lines_Left)
                     Rightmost = MAX(Rightmost, FM_Head_Lines_Left + %FMLLines + 2 - 1)
                  END IF                                          '
               END IF                                             '

               '----- Do the Date / Time
               IF FM_Head_Date_Left > 0 THEN                      ' Doing Date at all?
                  lclAttr = IIF$(LEFT$(DefSort, 4) = "Date", $$TxtHi, $$TxtLo)
                  IF INSTR(ENV.FMLayout, "DATETIME") THEN         ' Date / Time?
                     tText = LSET$(AFList(stp).LWTime, 19)        ' Get correct length text item
                     IF ENV.FMDateHilite THEN                     ' Do the Date hilite version?
                        IF TRIM$(AFList(stp).LWTime) >= gDateActive THEN
                           sPrint (tText, $$Red, sp, FM_Head_Date_Left)
                           Rightmost = MAX(Rightmost, FM_Head_Date_Left + 18)
                        ELSEIF TRIM$(AFList(stp).LWTime) > gDateActive1 THEN
                           sPrint (tText, $$Yellow, sp, FM_Head_Date_Left)
                           Rightmost = MAX(Rightmost, FM_Head_Date_Left + 18)
                        ELSEIF TRIM$(AFList(stp).LWTime) > gDateActive8 THEN
                           sPrint (tText, $$Green, sp, FM_Head_Date_Left)
                           Rightmost = MAX(Rightmost, FM_Head_Date_Left + 18)
                        ELSEIF TRIM$(AFList(stp).LWTime) > gDateActive24 AND TRIM$(AFList(stp).LWTime) > gDateActive48 THEN
                           sPrint (tText, $$Blue, sp, FM_Head_Date_Left)
                           Rightmost = MAX(Rightmost, FM_Head_Date_Left + 18)
                        ELSE
                           sPrint (tText, lclAttr, sp, FM_Head_Date_Left)
                           Rightmost = MAX(Rightmost, FM_Head_Date_Left + 18)
                        END IF                                    '
                     ELSE
                        sPrint (tText, lclAttr, sp, FM_Head_Date_Left)
                        Rightmost = MAX(Rightmost, FM_Head_Date_Left + 18)
                     END IF
                  ELSE                                            ' Else just the date
                     tText = LSET$(AFList(stp).LWTime, 12)        ' Get correct length text item
                     IF ENV.FMDateHilite THEN                     ' Do the Date hilite version?
                        IF TRIM$(AFList(stp).LWTime) >= gDateActive THEN
                           sPrint (tText, $$Red, sp, FM_Head_Date_Left)
                           Rightmost = MAX(Rightmost, FM_Head_Date_Left + 11)
                        ELSEIF TRIM$(AFList(stp).LWTime) > gDateActive1 THEN
                           sPrint (tText, $$Yellow, sp, FM_Head_Date_Left)
                           Rightmost = MAX(Rightmost, FM_Head_Date_Left + 11)
                        ELSEIF TRIM$(AFList(stp).LWTime) > gDateActive8 THEN
                           sPrint (tText, $$Green, sp, FM_Head_Date_Left)
                           Rightmost = MAX(Rightmost, FM_Head_Date_Left + 11)
                        ELSEIF TRIM$(AFList(stp).LWTime) > gDateActive24 AND TRIM$(AFList(stp).LWTime) > gDateActive48 THEN
                           sPrint (tText, $$Blue, sp, FM_Head_Date_Left)
                           Rightmost = MAX(Rightmost, FM_Head_Date_Left + 11)
                        ELSE
                           sPrint (tText, lclAttr, sp, FM_Head_Date_Left)
                           Rightmost = MAX(Rightmost, FM_Head_Date_Left + 11)
                        END IF
                     ELSE
                        sPrint (tText, lclAttr, sp, FM_Head_Date_Left)
                        Rightmost = MAX(Rightmost, FM_Head_Date_Left + 11)
                     END IF                                       '
                  END IF
               END IF

               '----- Do the Note
               IF FM_Head_Note_Left > 0 THEN                      ' Doing Note at all?
                  tText = me.WindowNote(stp) + "  "               ' Format Text
                  IF FMode = %FMFileList THEN                     ' Have an associated FILELIST?
                     lclAttr = IIF$(LEFT$(DefSort, 4) = "Note", $$TxtHi, $$TxtLo)
                     sPrint (tText, lclAttr, sp, FM_Head_Note_Left)
                     Rightmost = MAX(Rightmost, FM_Head_Note_Left + LEN(tText) - 1)
                  ELSE
                     sPrint (SPACE$(FM_Note_Size + 2), $$TxtLo, sp, FM_Head_Note_Left)
                     Rightmost = MAX(Rightmost, FM_Head_Note_Left + FM_Note_Size + 1)
                  END IF
               END IF

               IF Rightmost < ENV.ScrWidth THEN                   ' Need to blank remainder?
                  sPrint (SPACE$(ENV.ScrWidth - Rightmost), $$TxtLo, sp, Rightmost + 1) ' Print pad after the size
               END IF                                             '

            END IF                                                '
            INCR stp                                              '
         NEXT sp                                                  '
         Me.DispFMHelp                                            ' Add help lines if needed
         IF IsCBad OR CRow > LastScrLine THEN                     ' If we ended up somewhere bad
            me.CsrRow = 1: me.CsrCol = 11                         ' Reset it to Cmd line
         ELSEIF CRow >= FM_Top_File_Line THEN                     ' If in line area
            IF TRIM$(AFList(CRow - FM_Top_File_Line + TopScrn).Cmd) = "" AND _  ' If nothing in command
               TRIM$(AFList(CRow - FM_Top_File_Line + TopScrn).Note) = "" THEN  ' and nothing in Note
               me.CsrCol = 1                                      ' Reset it to Cmd line
            END IF                                                '
         END IF                                                   '

         me.MarkLineNumb                                          ' Mark the line number

         LFPath = UUCASE(TRIM$(FPath))                            ' Save last Path
         LFMask = UUCASE(TRIM$(FMask))                            ' Save last Mask
         LFileListNm = UUCASE(TRIM$(FileListNm))                  ' Save last FileListNm
         LScrlAmtC = UUCASE(Prf.Scroll)                           ' Save last Scroll amount
         LastTop = TopScrn                                        ' Save where we are
         GRAPHIC REDRAW                                           '
         GRAPHIC ATTACH PgHandle, WindowID                        ' Set as the default graphic area
         MExitMeth

      END METHOD

      METHOD  DispFMHelp()
      '----- Display FM HELP if asked for
      LOCAL HMask, tText AS STRING
         IF ISFALSE ENV.FMHelpFlag THEN EXIT METHOD               ' Skip if not needed
         SELECT CASE CONST$ FileListNm                            ' Setup the Command mask
            CASE ""                : HMask = "ABCDEIJLMOPRSTUWX"  ' With valid commands
            CASE "Found Files"     : HMask = "ABCDEFIJLMOPRSTUWX"
            CASE "Recent Files"    : HMask = "ABCDEFIJLMOPRSTUWX"
            CASE "Recent Paths"    : HMask = "BEFISX"
            CASE "Favorite Files"  : HMask = "BCDEFIJLMOPRSTUWX"
            CASE "Named Favorites" : HMask = "BDEGHIRSU"
            CASE "Open Files"      : HMask = "ACJOPVZNWX"
            CASE "Profiles"        : HMask = "BCDEIRSTX"
            CASE ELSE:             : HMask = "ABCDEFIJLMOPRSTU"
         END SELECT

         tText = " Line Commands: "                               ' Build help line 1
         IF INSTR(HMask, "A") THEN tText += $FMHAdd               '
         IF INSTR(HMask, "G") THEN tText += $FMHAll               '
         IF INSTR(HMask, "B") THEN tText += $FMHBrowse            '
         IF INSTR(HMask, "N") THEN tText += $FMHCancel            '
         IF INSTR(HMask, "C") THEN tText += $FMHClone             '
         IF INSTR(HMask, "D") THEN tText += $FMHDelete            '
         IF INSTR(HMask, "W") THEN tText += $FMHDIR               '
         IF INSTR(HMask, "E") THEN tText += $FMHEdit              '
         IF INSTR(HMask, "Z") THEN tText += $FMHEND               '
         IF INSTR(HMask, "I") THEN tText += $FMHExclude           '
         sPrtHelp (LSET$(tText, ENV.ScrWidth), ENV.ScrHeight - 2) ' Go print it
         tText = "                "                               ' Build help line 2
         IF INSTR(HMask, "F") THEN tText += $FMHForget            '
         IF INSTR(HMask, "J") THEN tText += $FMHJob               '
         IF INSTR(HMask, "L") THEN tText += $FMHLines             '
         IF INSTR(HMask, "M") THEN tText += $FMHMEdit             '
         IF INSTR(HMask, "H") THEN tText += $FMHNorm              '
         IF INSTR(HMask, "O") THEN tText += $FMHOpen              '
         IF INSTR(HMask, "P") THEN tText += $FMHPrint             '
         IF INSTR(HMask, "U") THEN tText += $FMHPurge             '
         IF INSTR(HMask, "R") THEN tText += $FMHRename            '
         IF INSTR(HMask, "V") THEN tText += $FMHSave              '
         IF INSTR(HMask, "S") THEN tText += $FMHSelect            '
         IF INSTR(HMask, "T") THEN tText += $FMHTouch             '
         IF INSTR(HMask, "B") THEN tText += $FMHView              '
         IF INSTR(HMask, "X") THEN tText += $FMHWDIR              '
         sPrtHelp(LSET$(tText, ENV.ScrWidth), ENV.ScrHeight - 1)  ' Go print it
         tText = " File Patterns: Masks separated by ;  such as *.BAS;*.TXT;*.INC;*.H  " ' Hint line
         sPrtHelp(LSET$(tText, ENV.ScrWidth), ENV.ScrHeight)      ' Go print it
      END METHOD                                                  '

      METHOD DispLine(stp AS LONG, sp AS LONG)
      '---------- Display one single text data line
      REGISTER i AS LONG
      IF ISFALSE IsLData(stp) THEN EXIT METHOD                    ' If not Data line, exit

         '----- Extra setup needed for PowerType mode
         IF IsPTypeMode THEN                                      ' If in PType mode
            FOR i = 1 TO gPTblCount                               ' Update the PT table with screen row
               IF stp = gPTbl(i).tLin THEN                        ' A PT line being displayed
                  gPTbl(i).sRow = sp                              ' Tell PT code what screen line it is
                  IF PTFDisp = 0 THEN PTFDisp = sp                ' Set the top/bottom PT lines
                  PTLDisp = sp                                    '
                  EXIT FOR                                        '
               END IF                                             '
            NEXT i                                                '
         END IF                                                   '

         sPrint (LTxtG(stp), LAttrG(stp), sp, gLNData1, Offset, gDatalen)   ' Go print the line, honor Offset/Datalen
         me.MarkReDraw(sp)                                        ' Re-draw the MARK lines
      END METHOD

      METHOD DispScreen()
      '---------- Display new screen from TopScrn down
      LOCAL VisTop, VisBot, Scrl, PTPrinted, cBG, cBG1, cBG2, ci, PageSeen AS LONG
      LOCAL d, e, title, tText, u, v AS STRING                    '
      LOCAL i, j, k, x, y, pcol, plen, sclr, eclr, TargetCol, lColumn, lLin, sp, stp, HideFlag AS LONG
      LOCAL lScrl, CurSet AS INTEGER         '

         MEntry
         DoClear(%Refresh)
         '----- If FM, pass off
         IF IsFMTab THEN me.DispFM: MexitMeth                     ' If this is the FM tab, pass it off
         '----- See if FileWatch has posted us
         IF WatchFlag <> " " THEN                                 ' Posted by File Watch?
            IF ISFALSE gfEndAll OR IsModdFlag THEN sFileChanged   ' If not doing shutdown or file modified
         END IF                                                   '

         '----- Start building the screen
         IF ISFALSE gMacroMode THEN GRAPHIC ATTACH PgHandle, WindowID, REDRAW ' Set as the default graphic area

         '----- Do the Xcluded line rebuild if needed
         me.LFlagBitOn(TopScrn, %TopScreen)                       ' Ensure TopScrn gets saved
         me.XRebuild(i, j, k)                                     ' Go rebuild the Excluded stuff
         IF i THEN sCurLin = i                                    ' Save the answers
         IF j THEN TopScrn = j                                    '
         IF k THEN sCurScrl = k                                   '

         '----- If a cursor request pending, get screen positioned
         IF sCurPrio > 0 THEN                                     ' Cursor request provided?
            IF sCurScrl THEN                                      ' If forced scrolling permitted
               VisTop = TopScrn                                   ' Get Top/Bottom visible screen lines
               IF sCurPrio <> %FLocate AND sCurPrio <> %FChange THEN ' If not a forced locate/change
                  IF IsLInvisible(sCurLin) THEN sCurLin -= LWrk1G(sCurLin) ' Backup to exclude line if needed

                  FOR i = 3 + Prf.Cols TO gwScrHeight             ' Get last line on screen
                     j = S(i)                                     ' Get the L() index from S()
                     IF j THEN                                    ' Got a line number?
                        VisBot = j                                ' Save as hi-water line number
                     ELSE                                         ' No line?
                        IF VisBot > 0 THEN EXIT FOR               ' We have a winner
                     END IF                                       '
                  NEXT i                                          '
                  IF sCurLin < VisTop OR sCurLin > VisBot THEN    ' Move if needed and allowed
                     TopScrn = IIF(sCurPrio <> %Locate, MAX(1, sCurLin - 1), sCurLin) ' Move TopScrn accordingly
                  END IF                                          '
               ELSE                                               '
                  TopScrn = MAX(1, sCurLin)                       ' Move TopScrn accordingly
               END IF                                             '
               TargetCol = L(sCurLin).LCol                        ' Copy column request
               IF TargetCol <> 0 AND (TargetCol < Offset + 1 OR TargetCol > Offset + gDataLen - 2) THEN
                  SELECT CASE AS CONST$ Prf.Scroll                ' Get the default Scrl amount
                     CASE "HALF": Scrl = INT(gDataLen / 2)        '
                     CASE "PAGE": Scrl = gDataLen                 '
                     CASE "FULL": Scrl = gDataLen                 '
                     CASE "DATA": Scrl = gDataLen - 1             '
                     CASE ELSE:   Scrl = gDataLen                 '
                  END SELECT                                      '
                  OffSet = 0                                      ' Set temp back to zero
                  DO WHILE OffSet + gDataLen < TargetCol          ' While target column is offscreen
                     me.OffSetAdd(Scrl)                           '
                  LOOP                                            '
               END IF                                             '
            END IF                                                '
         END IF                                                   '

         '----- Setup the Tab title
         me.TabTitleSet(%False)                                   '

         '----- Building the screen, do line 1
         cBandBG = %False                                         ' Reset any prior Banding value
         CurSet = %False: PTFDisp = 0: PTLDisp = 0                ' Start as no cursor set
         LastRulCol = 0: LastRulRow = 0                           ' No line cursors are present

         sPrint ("Command > ", $$TxtHi, 1, 1)                     ' Re-do the command line
         IF ISNOTNULL(pCommandPrv) THEN                           ' Got an & command line?
            pCommand = pCommandPrv                                ' Swap it in
            pCommandPrv = ""                                      ' null it
         END IF                                                   '
         me.WindowCmd                                             '
         sPrint ("Scroll > ", $$TxtHi, 1, 11 + pCmdLen + 1)               '
         sPrint (LSET$(Prf.Scroll, 4), $$TxtLo, 1, 11 + pCmdLen + 9 + 1)  '
         ScrlAmtLast = Prf.Scroll                                 ' Save as last valid Scroll value

         '----- Do line 2 and the error message
         sPrint (STRING$(ENV.ScrWidth, " "), $$TxtLo, 2, 1)       ' Print normal dash line
         IF ISFALSE gResizeActive  THEN                           ' Don't do messages during a RESIZE
            IF ErrMsgTblC > 1 THEN                                ' If there are multiple messages
               IF ErrMsgHigh > 0 THEN                             ' Some errors?
                  scError(%eMultiple, "Multiple messages (including errors) were triggered, press F1 (Help) to see full list")
               END IF                                             '
            END IF                                                '
            IF ErrMsgTblC > 0 THEN                                ' If there's an Error message?
               REDIM ErrMsgHlp(1 TO ErrMsgTblC)                   ' Set Hlp copy table correct size
               FOR i = 1 TO ErrMsgTblC                            ' Copy it
                  ErrMsgHlp(i) = ErrMsgTbl(i)                     '
               NEXT i                                             '
               ErrMsgHlpC = ErrMsgTblC                            ' Copy count
               sPrint (IIF$(ErrMsgTblC > 1, " +", "  ") + LEFT$(ErrMsg, 70), $$Error, 2, ENV.ScrWidth - LEN(LEFT$(ErrMsg, 70)) - 1) ' Print error message
               IF errFlag = %eFail THEN sDoBeep                   ' Beep on errors
            ELSEIF ISNOTNULL(gGlblMessage) THEN                   ' Got a global one?
               sPrint (" " + gGlblMessage, $$Error, 2, ENV.ScrWidth - LEN(gGlblMessage)) ' Print error message
               RESET gGlblMessage                                 '
            END IF                                                '
         END IF                                                   '
         RESET ErrFlag, ErrMsg, ErrMsgTbl(), ErrMsgTblC           ' Clear error status

         '----- Display fixed COLS line if called for
         IF Prf.Cols THEN                                         ' Constant Cols line?
            s(3) = 0                                              ' Null data line pointer
            sp = 3                                                ' Set line for print
            GOSUB DoColsLine                                      ' Go print it
         END IF                                                   '

         '----- Prep the text data
         me.pCmdOrder("F")                                        ' Re-Do the line numbers (perhaps)

         '----- Loop through text screen lines
         stp = TopScrn                                            ' Set Screen Text Pointer to current TopScrn setting
         FOR sp = 3 + Prf.Cols TO gwScrHeight                     ' Loop for logical Screen data lines
            sCalcEditBG(sp)                                       ' Calc banding value

            IF ISTRUE IsLPage(stp) THEN                           ' A =PAGE> line ?
               i = sp: j = Prf.Cols                               '
               IF LEFT$(Prf.EOL, 4) = "AUTO" AND Prf.Scroll = "PAGE" AND ISTRUE Prf.PageFlag AND _  ' PAGE mode AUTO processing?
                  sp > Prf.Cols + 4 AND _                         ' Not top line
                  ISFALSE Prf.ScrlPageSus THEN                    ' In correct mode?
                  PageSeen = %True                                ' =PAGE> below the top line
               END IF                                             '
            END IF                                                '

            '----- If below last line
            IF stp > LastLine  OR ISTRUE PageSeen THEN            ' Past end of dataset or end of Page in PAGE mode?
               s(sp) = 0                                          ' Null Txt index pointer
               sPrint (SPACE$(ENV.ScrWidth), $$TxtLo, sp, 1)      ' Blank the line on the screen

            '----- Do a file line
            ELSE                                                  '
               LastDrawn = sp                                     ' Remember last drawn line
               HideFlag = IIF(IsLXclude(stp + 1) AND IsHideFlag, %True, %False) ' Set HIDE mode underline

               '----- Base on line type
               IF IsLInvisible(stp) OR _                          ' Show this line at all?
                  (IsLXclude(stp) AND IsHideFlag) THEN            '
                  sp -= 1                                         ' No, keep sp at this screen slot
                  INCR stp                                        ' Move on to the next line
                  ITERATE FOR

               ELSE                                               '

                  '----- First do the line number area
                  s(sp) = stp                                     ' Save which L() entry we're showing
                  IF ISFALSE gMacroMode THEN                      ' If not Macro Mode
                     IF IsPTypeMode THEN                          ' If in PType mode all line numbers go Hi
                        FOR i = 1 TO gPTblCount                   ' Update the PT table with screen row
                           IF stp = gPTbl(i).tLin THEN            ' A PT line being displayed
                              IF HideFlag THEN                    ' Underlined or not
                                 sPrint (RIGHT$(LLNumG(stp), ENV.LinNoSize), $$LnoHiUL, sp, 1) ' Print the Line number
                              ELSE                                '
                                 sPrint (RIGHT$(LLNumG(stp), ENV.LinNoSize), $$LnoHi, sp, 1) ' Print the Line number
                              END IF                              '
                              EXIT FOR                            '
                           END IF                                 '
                        NEXT i                                    '
                     ELSE
                        IF HideFlag THEN                          ' Underlined or not                                     '
                           sPrint (RIGHT$(LLNumG(stp), ENV.LinNoSize), $$LnoLoUL, sp, 1) ' Print the Line number
                        ELSE                                      '
                           sPrint (RIGHT$(LLNumG(stp), ENV.LinNoSize), $$LnoLo, sp, 1) ' Print the Line number
                        END IF                                    '
                     END IF                                       '
                     sPrint(" ", $$TxtLo, sp, gLNPadCol)          ' Clear col gLNPadCol with normal blank
                  END IF                                          '

                  '----- Now overlay if needed with special data
                  IF stp = 1 THEN
                     u = me.LLCtlGet(stp): v = LLNumG(stp)
                     '
                  END IF
                  IF me.LLCtlGet(stp) <> LLNumG(stp) THEN         ' A pending line control?
                      IF ISFALSE gMacroMode THEN                  ' If not Macro Mode
                         sPrint (LEFT$(me.LLCtlGet(stp), ENV.LinNoSize), $$LnoHi, sp, 1) ' Print the Pending
                         IF LTagG(stp) <> $BlankLNo AND L(stp).LLbl <> $BlankLNo THEN _ ' A Label/Tag collision?
                            sPrint (":", $$TxtLo, sp, gLNPadCol)  ' Flag column gLNPadCol
                      END IF                                      '
                  END IF                                          '

                  '----- Now add the User status
                  IF IsLUser(stp) THEN                            ' A User Line?
                      sPrint ("|", $$TxtHi, sp, gLNPadCol)        ' Flag column gLNPadCol
                  END IF                                          '

'                  sPrint (format$(L(stp).LMix), $$TxtHi, sp, gLNPadCol)  ' Uncomment for MIX display during debugging
'                  sPrint (IIF$(L(stp).PCmd<0, "-",RIGHT$(FORMAT$(L(stp).PCmd, "0000"), 1)), $$TxtHi, sp, gLNPadCol) ' Uncomment for PCmdIX display during debugging

                  IF IsLTop(stp) THEN                             ' Handle types differently
                     tText = STRING$(gDataLen, "*")               '
                     MID$(tText, (gDataLen - 16) \ 2 + 1, 16) = "* Top of Data **"
                     sPrint (tText, $$TxtHi, sp, gLNData1)        '

                  ELSEIF IsLBottom(stp) THEN                      '
                     tText = STRING$(gDataLen, "*")               '
                     MID$(tText, (gDataLen - 16) \ 2 + 1, 16) = " Bottom of Data "
                     sPrint (tText, $$TxtHi, sp, gLNData1)        '

                  ELSEIF IsLInsertLine(stp) THEN                  '
                     sPrint (me.Window(LTxtG(stp)), $$TxtLo, sp, gLNData1)

                  ELSEIF IsLFile(stp) THEN                        '
                     sPrint (me.Window(LTxtG(stp)), $$TxtHi, sp, gLNData1)

                  ELSEIF IsLData(stp) THEN                        '
                     me.DispLine(stp, sp)                         '

                  ELSEIF IsLCols(stp) THEN                        '
                     GOSUB DoColsLine                             '

                  ELSEIF IsLWord(stp) THEN                        '
                     sPrint (LSET$(Prf.WordInput, gDataLen), $$TxtHi, sp, gLNData1)

                  ELSEIF IsLMark(stp) THEN                        '
                     sPrint (me.Window(LTxtG(stp)), $$TxtHi, sp, gLNData1)

                  ELSEIF IsLMask(stp) THEN                        '
                     sPrint (me.Window(LTxtG(stp)), $$TxtHi, sp, gLNData1)

                  ELSEIF IsLProf(stp) THEN                        '
                     sPrint (me.Window(LTxtG(stp)), $$TxtHi, sp, gLNData1)

                  ELSEIF IsLTabs(stp) THEN                        '
                     sPrint (me.Window(LTxtG(stp)), $$TxtHi, sp, gLNData1)

                  ELSEIF IsLNote(stp) THEN                        '
                     sPrint (me.Window(LTxtG(stp)), $$TxtHi, sp, gLNData1)
                     me.MarkReDraw(sp)                            ' Re-draw the MARK lines

                  ELSEIF IsLBounds(stp) THEN                      '
                     sPrint (me.WindowBnds, $$TxtHi, sp, gLNData1)'

                  ELSEIF IsLXclude(stp) THEN                      '
                     tText = STRING$(gDataLen, " ")               ' Build the marker line
                     LSET tText = REPEAT$((gDataLen), "-")        '
                     IF LWrk1G(stp) > 999999 THEN                 ' Choose a format
                        u = " <" + FORMAT$(LWrk1G(stp), "00000000") + "> " ' Build the sentence
                     ELSE                                         '
                        u = " < " + FORMAT$(LWrk1G(stp), "000000") + " > " ' Build the sentence
                     END IF                                       '
                     IF LEN(u) + 5 < LEN(tText) THEN              '
                        MID$(tText, (gDataLen - LEN(u)) - 5, LEN(u)) = u' Insert into the ---- line
                     ELSE                                         '
                        MID$(tText, (gDataLen - LEN(u)), LEN(u)) = u    ' Insert into the ---- line
                     END IF                                       '
                     sPrint (tText, $$TxtHi, sp, gLNData1)        '

                  END IF                                          '
               END IF                                             '

               '----- See if this is the cursor line
               IF sCurPrio > 0 THEN                               ' Is there a cursor request?
                  IF stp = sCurLin THEN                           ' Is this the line?
                     lScrl     = sCurScrl                         ' Fetch the winners data locally
                     lColumn   = L(stp).LCol                      '
                     me.CsrRow = sp + sCurHexl                    ' Yes set cursor on this row
                     IF lColumn = 0 THEN                          ' a LineNum request?
                        me.CsrCol = 1                             ' Yes, stuff in column 1
                     ELSE                                         '
                        me.CsrCol = lColumn - Offset + gLNPadCol  ' Calc column position
                        IF CCol < 1 OR CCol > ENV.ScrWidth THEN me.CsrCol = gLNData1
                     END IF                                       '
                     CurSet = %True                               '
                  END IF                                          '
               END IF                                             '

               '----- If HEX mode, do all the extra stuff
               IF Prf.HexMode = &4 THEN                           ' If HEX mode, do some more
                  INCR sp                                         ' Next Screen Line
                  IF sp <= gwScrHeight THEN                       ' If still on screen
                     s(sp) = -1                                   ' Mark as Top Hex line
                     IF IsLData(stp) AND ISFALSE IsLInvisible(stp) THEN '
                        sPrint ($BlankLNo, $$TxtLo, sp, 1)        ' Blank Line Number
                        sPrint (me.WindowHexTop(LTxtG(stp)), $$TxtHi, sp, gLNData1) ' Just print upper hex
                     ELSE                                         '
                        sPrint (SPACE$(ENV.ScrWidth), $$TxtLo, sp, 1)
                     END IF                                       '
                  END IF                                          '
                  INCR sp                                         ' Next Screen line
                  IF sp <= gwScrHeight THEN                       ' If still on screen
                     s(sp) = -2                                   ' Mark as Bottom Hex line
                     IF IsLData(stp) AND ISFALSE IsLInvisible(stp) THEN '
                        sPrint ($BlankLNo, $$TxtLo, sp, 1)        ' Blank Line Number
                        sPrint (me.WindowHexBottom(LTxtG(stp)), $$TxtHi, sp, gLNData1) ' Just print lower hex
                     ELSE                                         '
                        sPrint (SPACE$(ENV.ScrWidth), $$TxtLo, sp, 1)
                     END IF                                       '
                  END IF                                          '
                  INCR sp                                         ' Next screen line
                  IF sp <= gwScrHeight THEN                       ' If still on screen
                     s(sp) = -3                                   ' Mark as Dash line
                     IF stp <= LastLine THEN sPrint (STRING$(ENV.ScrWidth,"-"), $$TxtLo, sp, 1) ' Dash line on the screen
                  END IF                                          '
               END IF                                             '
            END IF                                                '
            INCR stp                                              ' Bump text line
         NEXT sp                                                  ' Bump Screen line

         IF sCurPrio > 0 THEN                                     ' Was this a specified cursor location?
            sCurPrio = 0                                          ' Reset for next time
            sCurLin  = 0                                          '
            me.LFlagBitOff(sCurLin, %Cursor)                      '
            L(sCurLin).LCol = 0                                   '
         END IF                                                   '

         IF ISFALSE gMacroMode THEN                               ' If not Macro Mode
            '----- Mark selected text
            IF IsMarkActive THEN                                  ' Do we have a marked area?
               OffMarkDrawn                                       ' Previous is gone, we just re-drew the screen
               me.MarkScr                                         ' Go re-draw the marked area
            END IF                                                '
            '----- Mark selected Cmd
            IF IsMiscActive THEN                                  ' Do we have a marked Cmd?
               OffMiscDrawn                                       ' Previous is gone, we just re-drew the screen
               me.MiscMark                                        ' Go re-draw the marked area
            END IF                                                '

            '---- Do MARK lines if called for
            me.MarkLineNumb                                       ' Mark the current line number area on the screen
            me.DoPFKShow                                          ' Re-do PFK show

            '----- Tell system to actually draw it
            GRAPHIC REDRAW                                        '
            GRAPHIC ATTACH PgHandle, WindowID                     ' Set as the default graphic area
         END IF
      BailExit:
         ON ERROR GOTO 0                                          '
         MExitMeth                                                '

      '----- Do special handling of COLS line
      DoColsLine:
         sPrint (gLnoTextTxt(5) + " ", $$TxtHi, sp, 1)            '
         tText = me.WindowCol()                                   ' Build Cols line
         i = 1: j = gLNData1                                      ' Print it the hard way

         DO WHILE i <= LEN(tText)                                 '
            k = INSTR(i, tText, ANY $UpperSpec)                   ' Look for any > 128 chars
            IF k = 0 THEN sPrint (MID$(tText, i), $$TxtHi, sp, j): EXIT DO    ' No more, remainder is normal
            IF k = i THEN                                         ' Found at start column?
               sPrint (CHR$(ASC(MID$(tText, i, 1)) - 128), $$TxtHiUL, sp, j)
               INCR i: INCR j                                     ' Step over
            ELSE                                                  '
               sPrint (MID$(tText, i, k - i), $$TxtHi, sp, j)     '
               j += (k - i): i = k                                ' Adjust continue values
            END IF                                                '
         LOOP                                                     '
         RETURN                                                   '
      BailOut:
         RESUME BailExit                                          '
      END METHOD

      METHOD DoAutoCaps(lno AS LONG) AS STRING
      '---------- Return an AUTOCAP version of a  line
      REGISTER i AS LONG
      REGISTER j AS LONG
      LOCAL sTxt AS STRING, k AS LONG
         sTxt = lTxtG(lno)                                        ' Get the text line
         IF ISFALSE Prf.HiAuto OR ISFALSE IsClrFlag THEN _        ' If colorize not active
            METHOD = sTxt: EXIT METHOD                            ' Return the normal line

         k = LAttrSPtr(lno): j = 1                                ' Get pointer to actual characters
         FOR i = k TO k + ((LEN(sTxt) - 1) * 2) STEP 2            ' Attr by Attr
            IF (PEEK(WORD, i) AND %AttrUC) THEN                   ' An UC position
               MID$(sTxt, j, 1) = uucase(MID$(sTxt, j, 1))        ' Then do it
            END IF                                                '
            INCR j                                                ' Index through sTxt
         NEXT i                                                   '
         METHOD = sTxt                                            ' Pass back the modified line
      END METHOD

      METHOD DoBROWSE(i AS LONG)
      '---------- Pass file to BROWSE
      LOCAL lcmd, cmd, SetProf AS STRING
         lcmd = AFlist(i).Cmd                                     ' Save command string
         WHILE lcmd <> ""                                         ' Process lcmd for .profile name
            cmd = sGetWord(lcmd, %Strip, %QuoteNotSig)            ' Get a word from the command
            IF LEFT$(cmd, 1) = "." THEN SetProf = UUCASE(cmd)     '
         WEND                                                     '
         me.CallTab("BROWSE", $DQ + RTRIM$(AFList(i).Path) + RTRIM$(AFList(i).FD.FileName) + $DQ + " " + SetProf)                                     '
      END METHOD

      METHOD  DoDirDown(i AS LONG)
      '---------- Do a Dir Down function
         FPath = RTRIM$(AFList(i).Path) + RTRIM$(AFList(i).FD.FileName)
         FileListNm = ""                                          ' Kill FileList name
         DoSet(%LoadReq)                                          ' Full refresh
      END METHOD

      METHOD  DoDirUp(i AS LONG)
      '---------- Do a Dir Up function
      LOCAL cmd AS STRING, j AS LONG
         cmd = TRIM$(AFList(i).Path)                              '
         j = INSTR(-2, cmd, "\")                                  ' Find 2nd last \
         IF j <> 0 THEN                                           ' Got one
            FPath = LEFT$(cmd, j)                                 ' Cut off the last level
            FileListNm = ""                                       ' Kill FileList name
            DoSet(%LoadReq)                                       ' Full refresh
         END IF                                                   ' Else do nothing
      END METHOD

      METHOD DoEDIT(i AS LONG)
      '---------- Pass file to EDIT
      LOCAL lcmd, cmd, SetProf AS STRING
         lcmd = AFlist(i).Cmd                                     ' Save command string
         WHILE lcmd <> ""                                         ' Process lcmd for .profile name
            cmd = sGetWord(lcmd, %Strip, %QuoteNotSig)            ' Get a word from the command
            IF LEFT$(cmd, 1) = "." THEN SetProf = UUCASE(cmd)     '
         WEND                                                     '

         IF (AFList(i).FD.FileAttributes AND %FILE_ATTRIBUTE_READONLY) = %FILE_ATTRIBUTE_READONLY THEN
            CallMsg = "File was ReadOnly, it has been opened in View mode"
            me.CallTab("VIEW", $DQ + RTRIM$(AFList(i).Path) + RTRIM$(AFList(i).FD.FileName) + $DQ + " " + SetProf)
         ELSE                                                     '
            me.CallTab("EDIT", $DQ + RTRIM$(AFList(i).Path) + RTRIM$(AFList(i).FD.FileName) + $DQ + " " + SetProf)
         END IF                                                   '
      END METHOD

      METHOD DoOpen(i AS LONG, oper AS STRING)
      '---------- Pass file to Open
      LOCAL tFn, lcmd, cmd, SetProf AS STRING
         lcmd = AFlist(i).Cmd                                     ' Save command string
         WHILE lcmd <> ""                                         ' Process lcmd for .profile name
            cmd = sGetWord(lcmd, %Strip, %QuoteNotSig)            ' Get a word from the command
            IF LEFT$(cmd, 1) = "." THEN SetProf = UUCASE(cmd)     '
         WEND                                                     '
         me.CallTab(oper, $DQ + RTRIM$(AFList(i).Path) + RTRIM$(AFList(i).FD.FileName) + $DQ + " " + SetProf)
      END METHOD

      METHOD DoVIEW(i AS LONG)
      '---------- Pass file to VIEW
      LOCAL tFn, lcmd, cmd, SetProf AS STRING
         lcmd = AFlist(i).Cmd                                     ' Save command string
         WHILE lcmd <> ""                                         ' Process lcmd for .profile name
            cmd = sGetWord(lcmd, %Strip, %QuoteNotSig)            ' Get a word from the command
            IF LEFT$(cmd, 1) = "." THEN SetProf = UUCASE(cmd)     '
         WEND                                                     '
         me.CallTab("VIEW", $DQ + RTRIM$(AFList(i).Path) + RTRIM$(AFList(i).FD.FileName) + $DQ + " " + SetProf)
      END METHOD

      METHOD  FileAbbrev(fnn AS STRING, dl AS LONG) AS STRING
      '---------- Abbreviate a filename to specified length
         METHOD = IIF$(LEN(Fnn) - Dl < 1, LSET$(fnn, Dl), LEFT$(fnn, (Dl - 3) \ 3) + "..." & RIGHT$(fnn, ((Dl - 3) * 2) \ 3))
      END METHOD

      METHOD GetFileAndRange() AS INTEGER
      '---------- Get a filename and line range
      LOCAL fn AS STRING
      LOCAL f, t, fCtl AS STRING, oldOpCount, HadRange, i, fnum, tnum AS LONG
         MEntry
         METHOD = %FALSE                                          ' Start off saying we've failed (easy EXIT's)

         '----- See if Cmd line has a range
         fCtl = "CMD"                                             ' To indicate line range came from the Cmd
         oldOpCount = pCmdNumOps                                  ' Save # operands
         IF ISTRUE me.RangeSet("", %True, %False) THEN MExitMeth  ' Strip out any line range operands (True = numerics allowed)
         IF oldOpCount <> pCmdNumOps THEN                         ' If # changed, we got something
            HadRange = %True                                      ' Range operands were processed, we had some operands
            SavFilLCtl = "CMD"                                    ' Save it
         END IF

         '----- If no range from Cmd, try line command range
         IF ISFALSE HadRange THEN                                 ' Was a Line control range provided
            IF ISFALSE LTblRange THEN _                           ' A line control range available?
               MErrExit(%ePending, "Pending line command Range")  '
            SavFilLCtl = TRIM$(LTblSCmd)                          ' Save it in TP area
            DotStart = LTblSFrom: DotEnd = LTblSTo: DotSet = %True'
            me.TTblDel(LTblSFrom)                                 ' Remove from Touched lines
            me.UpdLControl(LTblSFrom)                             ' Clear from the line
            me.TTblDel(LTblSTo)                                   ' Remove from Touched lines
            me.UpdLControl(LTblSTo)                               ' Clear from the line
         ELSE                                                     ' Else we got range from the command
            SavFilLCtl = "CMD"                                    ' Save it in TP area
         END IF                                                   '

         '----- Now try for X/NX and filename
         me.FindReset                                             ' Reset the parse output fields
         IF sCRTParse("S1U") THEN MExitMeth                       ' Scan (Exclude, Lit1, User)
         me.FindSetup                                             ' Copy CRT parse's answers
         CfSet(%CrtDX)                                            ' Don't alter exclude status
         IF DotSet OR _                                           ' If a line range
            IsCfU OR IsCfNU OR IsCfX OR IsCfNX THEN _             ' Or U / NU / X / NX
            CfSet(%CrtALL)                                        ' Force ALL

         '----- Get TOP setup for the output
         TOP.Setup("AP", "", "Specify output filename", cfFind)   ' Setup TOP for the file
         IF TOP.EXEC THEN                                         ' Go validate
            scError(%eFail, TOP.ResultMsg)                        ' Oops?  Bail out
            i = me.RangeSet("RESET", %False, %False)              ' Ensure default range of ALL before reset
            me.ResetFunc(%ResetCommand)                           '
            IF TOP.RSTabFound THEN                                ' Found elsewhere already open?
               gTabSwitch = TOP.TabFound                          ' Set tab to go to
               MExitMeth                                          '
            END IF                                                '
         END IF                                                   '
         CfClear(%CRTLit1)                                        ' Remove it as a search literal

         METHOD = %True                                           '
         MExit
      END METHOD

      METHOD GetIncrement(BYREF slno AS LONG, BYREF insno AS LONG, BYREF inc AS LONG)
      '---------- Get line number increment for an insert block
      REGISTER i AS LONG
      REGISTER j AS LONG
      LOCAL sp, ep, x, y, z AS LONG
         MEntry

         sp = slno                                                ' Set starting point to passed line number
         ep = slno + 1                                            ' Ending = starting + 1

         '----- Adjust if sp/ep not sitting on data lines
         IF ISFALSE (L(sp).Lflag AND %Data) THEN sp = me.CRPBack(%mOData, sp, 1) ' Backup to previous Data line

         IF ISFALSE (L(ep).LFlag AND %Data) THEN ep = me.CRPFwd(%mOData, ep, 1)  ' Fwd to next Data line

         '----- Calc increment, see if a simple 'fit'
         x = VAL(l(sp).LLNum): y = VAL(l(ep).LLNum)               ' Get line numbers of prev/next lines
         slno = x                                                 ' Pass back the previous line number
         IF ep = LastLine THEN inc = 10: MExitMeth                ' If inserting before LastLine, pass back 10
         z = FIX((y - x) / (insno + 1))                           ' See if room for insert line numbers
         IF z > 0 THEN inc = z: MExitMeth                         ' Simple 'fit', pass back the answer

         '----- Won't fit, must renumber some subsequent lines
         inc = 2                                                  ' Set 2 as a default incr
         z = (inc * insno) + slno + inc                           ' Calc line # of 1st line after last inserted line (using 2 as an incr)

         DO WHILE VAL(l(ep).LLNum) <= z AND ep <> LastLine        ' Shift line number of subseq lines upward
            L(ep).LLNum = FORMAT$(z, "00000000")                  ' Update the LLNum field
            me.UpdLControl(ep)                                    ' And the LLCtl
            z += inc                                              ' Bump new line number
            ep = me.CRPFwd(%mOData, ep, 1)                        ' Move to next data line
         LOOP                                                     '
         MExit
      END METHOD

      METHOD IncAFSize()
         INCR AFCount
         IF AFCount > UBOUND(AFList()) THEN                       ' Keep RQList big enough
            REDIM PRESERVE AFList(1 TO 2 * AFCount) AS INSTANCE FMFList ' Double it
         END IF                                                   '
      END METHOD

      METHOD IncRFSize()
         INCR RFCount
         IF RFCount > UBOUND(RFList()) THEN                       ' Keep RFList big enough
            REDIM PRESERVE RFList(1 TO 2 * RFCount) AS INSTANCE STRING ' Double it
         END IF                                                   '
      END METHOD

      METHOD IncRQSize()
         INCR RQCount
         IF RQCount > UBOUND(RQList()) THEN                       ' Keep RQList big enough
            REDIM PRESERVE RQList(1 TO 2 * RQCount) AS INSTANCE STRING ' Double it
         END IF                                                   '
      END METHOD

      METHOD IncRXSize()
         INCR RXCount
         IF RXCount > UBOUND(RXList()) THEN                       ' Keep RXList big enough
            REDIM PRESERVE RXList(1 TO 2 * RXCount) AS INSTANCE STRING ' Double it
         END IF                                                   '
      END METHOD

      METHOD IncRNSize()
         INCR RNCount
         IF RNCount > UBOUND(RNList()) THEN                       ' Keep RNList big enough
            REDIM PRESERVE RNList(1 TO 2 * RXCount) AS INSTANCE STRING ' Double it
         END IF                                                   '
      END METHOD

      METHOD IsAFDup(fn AS STRING) AS LONG
      '---------- See if a filename is already in the AFList array
      REGISTER i AS LONG
         FOR i = 1 TO AFCount                                     ' Spin through the table
            IF IsEQ(fn, TRIM$(AFList(i).Path) + TRIM$(AFList(i).FD.FileName)) THEN METHOD = %True: EXIT METHOD
         NEXT i                                                   '
      END METHOD

      METHOD IsOnScrn(ln AS LONG) AS LONG
      '---------- Is a line on screen?
      REGISTER i AS LONG
         ARRAY SCAN s() FOR ENV.ScrHeight, =ln, TO i              ' Is line number in the table
         METHOD = IIF(i, %True, %False)                           ' Return a True/False
      END METHOD

      METHOD MacSubst(mline AS STRING) AS LONG
      '---------- Do macro substitution
      LOCAL i, j AS LONG, fn, lsubarg, lclSet, TX, MOp AS STRING
      DIM Dops(1 TO 9) AS STRING
         '----- If gDosOperands present, break into sub-parts
         IF ISNOTNULL(gDosOperands) THEN                          ' Break out gDosOperands?
            TX = gDosOperands                                     ' Get working copy
            DOps(1) = sGetWord(TX, %Strip, %QuoteSig)             ' Get an Operand
            DOps(2) = sGetWord(TX, %Strip, %QuoteSig)             '
            DOps(3) = sGetWord(TX, %Strip, %QuoteSig)             '
            DOps(4) = sGetWord(TX, %Strip, %QuoteSig)             '
            DOps(5) = sGetWord(TX, %Strip, %QuoteSig)             '
            DOps(6) = sGetWord(TX, %Strip, %QuoteSig)             '
            DOps(7) = sGetWord(TX, %Strip, %QuoteSig)             '
            DOps(8) = sGetWord(TX, %Strip, %QuoteSig)             '
            DOps(9) = TX                                          '
         END IF                                                   '

         '----- Setup Subarg and full filename
         lsubarg = IIF$(ISNULL(Prf.SubArg), "OFF", Prf.SubArg)    '
         fn = TIP.FilePath                                        '

         '----- If any ^M arguments
         i = INSTR(UUCASE(mline), "~M"): IF i = 0 THEN i = INSTR(UUCASE(mline), "^M") ' See if it's present
         IF i THEN                                                ' Lots of extra stuff here
           MOp =  "-SUB:" + gSubmitType                           ' Start building the ^M string
           MOp += " -EOL:" + TRIM$(Prf.EOL)                       '
           MOp += " -LEN:" + FORMAT$(Prf.LRECL)                   '
           MOp += " -FMT:" + Prf.RECFM                            '
           MOp += " -SRC:" + Prf.PSource                          '
           MOp += " -TAB:" + FORMAT$(Prf.ImportTabs)              '
         END IF                                                   '

         '----- If any Key arguments
         i = INSTR(UUCASE(mline), "~K"): IF i = 0 THEN i = INSTR(UUCASE(mline), "^K") ' See if the hard one is present
         IF i THEN                                                ' Lots of extra stuff here
           IF MID$(mline, i + 2, 1) <> "(" OR INSTR(i + 2, mline, ")") = 0 THEN
              METHOD = %True: EXIT METHOD                         ' Sorry
           END IF                                                 '
           IF sKMacro(mline) THEN METHOD = %True: EXIT METHOD     ' Pass off the hard work (sKMacro does ALL ~K in line)
         END IF                                                   '

         '----- If any Set arguments
         i = INSTR(UUCASE(mline), "~S("): IF i = 0 THEN i = INSTR(UUCASE(mline), "^S(") ' See if the hard one is present
         IF i THEN                                                ' Lots of extra stuff here
            IF MID$(mline, i + 2, 1) <> "(" OR INSTR(i + 2, mline, ")") = 0 THEN
               METHOD = %True: EXIT METHOD                        ' Sorry
            END IF                                                '
            IF sSSet(mline) THEN METHOD = %True: EXIT METHOD      ' Pass off the hard work (sSSet does all ~S in line)
         END IF                                                   '

         '----- If any Environment arguments
         i = INSTR(UUCASE(mline), "~V("): IF i = 0 THEN i = INSTR(UUCASE(mline), "^V(") ' See if the hard one is present
         IF i THEN                                                ' Lots of extra stuff here
            IF MID$(mline, i + 2, 1) <> "(" OR INSTR(i + 2, mline, ")") = 0 THEN
               METHOD = %True: EXIT METHOD                        ' Sorry
            END IF                                                '
            IF sVSet(mline) THEN METHOD = %True: EXIT METHOD      ' Pass off the hard work (sVSet does all ~V in line)
         END IF                                                   '

         '----- If any SET values, get lclSet built

         REPLACE "^#" WITH gPageNumber IN mline                   ' Do Uppercase variation
         REPLACE "^1" WITH UUCASE(Dops(1)) IN mline               '
         REPLACE "^2" WITH UUCASE(Dops(2)) IN mline               '
         REPLACE "^3" WITH UUCASE(Dops(3)) IN mline               '
         REPLACE "^4" WITH UUCASE(Dops(4)) IN mline               '
         REPLACE "^5" WITH UUCASE(Dops(5)) IN mline               '
         REPLACE "^6" WITH UUCASE(Dops(6)) IN mline               '
         REPLACE "^7" WITH UUCASE(Dops(7)) IN mline               '
         REPLACE "^8" WITH UUCASE(Dops(8)) IN mline               '
         REPLACE "^9" WITH UUCASE(Dops(9)) IN mline               '
         REPLACE "^a" WITH UUCASE(gDosOperands) IN mline          '
         REPLACE "^A" WITH UUCASE(gDosOperands) IN mline          '
         REPLACE "^c" WITH UUCASE(ENV.INIPath) IN mline           '
         REPLACE "^C" WITH UUCASE(ENV.INIPath) IN mline           '
         REPLACE "^d" WITH UUCASE(TIP.Date) IN mline              '
         REPLACE "^D" WITH UUCASE(TIP.Date) IN mline              '
         REPLACE "^e" WITH UUCASE(ENV.EXEPath) IN mline           '
         REPLACE "^E" WITH UUCASE(ENV.EXEPath) IN mline           '
         REPLACE "^f" WITH UUCASE(TIP.FilePath) IN mline          '
         REPLACE "^F" WITH UUCASE(TIP.Base) IN mline              '
         REPLACE "^i" WITH UUCASE(gSubmitFile) IN mline           '
         REPLACE "^I" WITH UUCASE(gSubmitFile) IN mline           '
         REPLACE "^j" WITH UUCASE(gJobID) IN mline                '
         REPLACE "^J" WITH UUCASE(gJobID) IN mline                '
         REPLACE "^l" WITH UUCASE(CursLine) IN mline              '
         REPLACE "^L" WITH UUCASE(CursLine) IN mline              '
         REPLACE "^m" WITH UUCASE(MOp) IN mline                   '
         REPLACE "^M" WITH UUCASE(MOp) IN mline                   '
         REPLACE "^n" WITH UUCASE(fn) IN mline                    '
         REPLACE "^N" WITH UUCASE(fn) IN mline                    '
         REPLACE "^p" WITH UUCASE(TIP.Path) IN mline              '
         REPLACE "^P" WITH UUCASE(TIP.Path) IN mline              '
         REPLACE "^r" WITH UUCASE(gResultFile) IN mline           '
         REPLACE "^R" WITH UUCASE(gResultFile) IN mline           '
         REPLACE "^t" WITH UUCASE(TIP.Time) IN mline              '
         REPLACE "^T" WITH UUCASE(TIP.Time) IN mline              '
         REPLACE "^w" WITH UUCASE(CursWord) IN mline              '
         REPLACE "^W" WITH UUCASE(CursWord) IN mline              '
         REPLACE "^x" WITH UUCASE(TIP.Extn) IN mline              '
         REPLACE "^X" WITH UUCASE(TIP.Extn) IN mline              '
         REPLACE "^z" WITH UUCASE(lsubarg) IN mline               '
         REPLACE "^Z" WITH UUCASE(lsubarg) IN mline               '
         REPLACE "~#" WITH gPageNumber IN mline                   ' Do Lowercase variation
         REPLACE "~1" WITH Dops(1) IN mline                       '
         REPLACE "~2" WITH Dops(2) IN mline                       '
         REPLACE "~3" WITH Dops(3) IN mline                       '
         REPLACE "~4" WITH Dops(4) IN mline                       '
         REPLACE "~5" WITH Dops(5) IN mline                       '
         REPLACE "~6" WITH Dops(6) IN mline                       '
         REPLACE "~7" WITH Dops(7) IN mline                       '
         REPLACE "~8" WITH Dops(8) IN mline                       '
         REPLACE "~9" WITH Dops(9) IN mline                       '
         REPLACE "~a" WITH gDosOperands IN mline                  '
         REPLACE "~A" WITH gDosOperands IN mline                  '
         REPLACE "~c" WITH ENV.INIPath IN mline                   '
         REPLACE "~C" WITH ENV.INIPath IN mline                   '
         REPLACE "~d" WITH TIP.Date IN mline                      '
         REPLACE "~D" WITH TIP.Date IN mline                      '
         REPLACE "~e" WITH ENV.EXEPath IN mline                   '
         REPLACE "~E" WITH ENV.EXEPath IN mline                   '
         REPLACE "~f" WITH TIP.Base IN mline                      '
         REPLACE "~F" WITH TIP.Base IN mline                      '
         REPLACE "~i" WITH gSubmitFile IN mline                   '
         REPLACE "~I" WITH gSubmitFile IN mline                   '
         REPLACE "~j" WITH gJobID IN mline                        '
         REPLACE "~J" WITH gJobID IN mline                        '
         REPLACE "~l" WITH CursLine IN mline                      '
         REPLACE "~L" WITH CursLine IN mline                      '
         REPLACE "~m" WITH MOp IN mline                           '
         REPLACE "~M" WITH MOp IN mline                           '
         REPLACE "~n" WITH fn IN mline                            '
         REPLACE "~N" WITH fn IN mline                            '
         REPLACE "~p" WITH TIP.Path IN mline                      '
         REPLACE "~P" WITH TIP.Path IN mline                      '
         REPLACE "~r" WITH gResultFile IN mline                   '
         REPLACE "~R" WITH gResultFile IN mline                   '
         REPLACE "~t" WITH TIP.Time IN mline                      '
         REPLACE "~T" WITH TIP.Time IN mline                      '
         REPLACE "~w" WITH CursWord IN mline                      '
         REPLACE "~W" WITH CursWord IN mline                      '
         REPLACE "~x" WITH TIP.Extn IN mline                      '
         REPLACE "~X" WITH TIP.Extn IN mline                      '
         REPLACE "~z" WITH lsubarg IN mline                       '
         REPLACE "~Z" WITH lsubarg IN mline                       '
         METHOD = %False                                          ' Say we're OK
      END METHOD

      METHOD MarkExtend(trow AS LONG, tcol AS LONG, StartRow AS LONG, StartCol AS LONG)
      '---------- See if a mouse mark area is to be extended
      DIM KeyStateArray(0 TO 255) AS BYTE                         ' Keyboard stats

      IF IsMarkActive THEN                                        '
         IF ISFALSE (tCol > gLNPadCol AND tRow > 2 + Prf.Cols AND tRow < gwScrHeight + 1) THEN ' Outside true text area?
            me.MarkKill                                           ' Kill it
            EXIT METHOD                                           '
         END IF                                                   '
         tCol = tCol + Offset - gLNPadCol                         ' Convert to text column
         GetKeyboardState(BYVAL VARPTR(KeyStateArray(0)))         ' Get all key states etc
         IF (ISTRUE ENV.Allow2DMouse OR (KeyStateArray(%vk_shift) AND &H80) OR (KeyStateArray(%vk_control) AND &H80) OR (KeyStateArray(%vk_menu) AND &H80)) AND ISFALSE IsSwapActive THEN
            tRow = S(tRow)                                        ' Get text line
         ELSE                                                     '
            tRow = MarkSLin                                       ' Use previous text line
         END IF                                                   '
         IF (LFlagG(tRow) AND (%Tabs OR %Word OR %Mark OR %Mask OR %Prof)) <> 0 THEN sDoBeep: EXIT METHOD

         '----- Don't let MARK overlap with SWAP area
         IF IsSwapActive THEN                                     ' Prevent Swap overlap
            IF trow = SwapSLin THEN                               '
               IF tcol >= SwapSCol AND tCol <= SwapECol THEN EXIT METHOD
               IF MarkSCol <= SwapSCol AND tCol >= SwapSCol THEN EXIT METHOD
               IF MarkSCol >= SwapECol AND tCol <= SwapECol THEN EXIT METHOD
            END IF                                                '
         END IF                                                   '

         '----- Set the MARK boundaries
         IF tCol > StartCol THEN                                  ' Build SlectRC RECT structure appropriately
            MarkSCol = StartCol                                   '
            MarkECol = tCol                                       '
         ELSE                                                     '
            MarkSCol = tCol                                       '
            MarkECol = StartCol                                   '
         END IF                                                   '
         IF tRow <= StartRow THEN                                 '
            MarkSLin = tRow                                       '
            IF ISFALSE IsSwapActive THEN MarkELin = StartRow      '
         ELSE                                                     '
            MarkSLin = StartRow                                   '
            IF ISFALSE IsSwapActive THEN MarkELin = tRow          '
         END IF                                                   '

         '----- Get the marked area drawn
         me.MarkScr                                               ' Try drawing a hi-lite

      '----- Check if the Cmd line hi-light version
      ELSEIF IsMiscActive THEN                                    ' Misc highlight?
         IF ISFALSE (tCOL > 10 AND tCOL < 10 + pCmdLen) OR ISNULL(TRIM$(pCommand)) THEN   ' Outside the cmd area?
            me.MarkKill                                           ' Kill it
            EXIT METHOD                                           '
         END IF                                                   '
         tCol = tCol + COffset - 10                               ' Convert to cmd column
         tRow = 1                                                 ' Always row 1

         '----- Set the CmdM boundaries
         MiscSCol = IIF(tCol > StartCol, StartCol, tCol)          ' Set appropriately
         MiscECol = IIF(tCol > StartCol, tCol, StartCol)          '

         '----- Get the marked area drawn
         me.MiscMark                                              ' Try drawing a hi-lite
      END IF                                                      '
      END METHOD

      METHOD   MarkFMLineNumb()
      '---------- Mark the line number where the cursor is
      REGISTER i AS LONG
      REGISTER lr AS LONG
      LOCAL j AS LONG, tText, tText1, tText2 AS STRING
         MEntry
         lr = CRow                                                ' Copy Cursor row
         IF lr < FM_Top_File_Line THEN lr = 0                     ' Only do if below headers
         IF lr <> FMMarkdSLin AND ISTRUE FMMarkdSLin THEN         ' Clear any previous marked line
            cBandBG = %False                                      ' remove banding
            IF ISNOTNULL(FileListNm) THEN sPrint (LSET$("*", FM_Crit_Size), $$TxtLo, FM_Mask_Line, FM_Mask_Left)
            sCalcFMBG(FMMarkdSLin)                                ' Calc BG value for banding

            IF ISNOTNULL (TRIM$(AFList(FMMarkdALin).Msg)) THEN    ' Is there a message waiting?
               sPrint (LSET$(AFList(FMMarkdALin).msg, FM_File_Size), $$Error, FMMarkdSLin, FM_Head_Name_Left)
            ELSE                                                  ' Else print the real filename
               '----- Split filename and extension
               me.FMPrtFileName(FMMarkdALin, FMMarkdSLin, %False) ' Print filename normally
            END IF                                                '
            FMMarkdSLin = 0                                       '
         END IF

         ' Now see if a new line to mark
         IF lr = 0 OR lr >= FM_Top_File_Line + AFCount - TopScrn OR lr < FM_Top_File_Line OR lr > FM_Top_File_Line + FM_List_Height - 1 THEN
            MExitMeth                                             '
         END IF                                                   '

         i = lr - FM_Top_File_Line + TopScrn                      ' Set i to AFList index

         IF ISNOTNULL (TRIM$(AFList(i).Msg)) THEN                 ' Is there a message waiting?
            sPrint (LSET$(AFList(i).msg, FM_File_Size), $$Error, lr, FM_Head_Name_Left)

         ELSE                                                     ' Else print the real filename
            IF ISNOTNULL(FileListNm) THEN
               cBandBG = %False                                   ' remove banding
               sPrint (LSET$("*",FM_Crit_Size), $$TxtLo, FM_Mask_Line, FM_Mask_Left)
               IF AFList(i).RQIX <> 0 THEN                        ' Get the RQList item
                  tText = RQList(AFList(i).RQIX)                  '
                  j = INSTR(tText, "|")                           ' Find 1st comma
                  tText1 = LEFT$(tText, j - 1)                    ' Extract FilePath
                  j = INSTR(j+1, tText, "|")                      ' Find 2nd comma
                  tText = LEFT$(tText, j - 1)                     ' Strip of junk user doesn't need to see
                  j = INSTR(tText, "|")                           ' Find 1st comma now
                  IF MID$(tText, j + 1) = "*" OR MID$(tText, j + 1) = "*.*" THEN tText = LEFT$(tText, j - 1)
                  tText = LSET$(tText, FM_Crit_Size)              '
                  sPrint (tText, $$TxtLo, FM_Mask_Line, FM_Mask_Left)
               END IF                                             '
            END IF                                                '
            sCalcFMBG(lr)                                         ' Calc BG value for banding

            me.FMPrtFileName(i, lr, %True)                        ' Print filename in HiLight
            FMMarkdSLin = lr                                      ' Save cursor line
            FMMarkdALin = i                                       ' Save FMFiles index being highlighted
         END IF                                                   '
         MExit
      END METHOD

      METHOD MarkLineNumb()
      '---------- Mark the line number where the cursor is
      REGISTER i AS LONG
      REGISTER logicalrow AS LONG
      LOCAL lclScheme AS WSTRING
         IF IsFMTab THEN me.MarkFMLineNumb: EXIT METHOD           ' Pass off the FM style quickly
         MEntry
         logicalrow = CRow                                        ' Copy Cursor row
         IF logicalrow > 2 + Prf.Cols THEN                        ' Only do if below headers
            IF S(logicalrow) < 0 THEN logicalrow = logicalrow - ABS(S(logicalrow))
         ELSE                                                     '
            logicalrow = 0                                        ' Force to zero
         END IF                                                   '

         IF S(logicalrow) <> 0 AND s(logicalrow) = InsLastClr THEN MExitMeth ' Skip if typing in a '''''' row
         sCalcEditBG(MarkedLine)                                  ' Calc banding value

         IF logicalrow <> MarkedLine AND ISTRUE MarkedLine THEN   ' Clear any previous marked line
            IF S(MarkedLine) <> 0 THEN i = sGetIX(MarkedLine)     ' Get IX of real data line if still on screen
            lclScheme = IIF$(IsLXclude(i + 1) AND IsHideFlag, $$LnoHiUL, $$LnoHi) ' Set HIDE mode underline
            IF me.LLCtlGet(i) <> LLNumG(i) OR _                   ' A pending line control?
               L(i).LLbl <> $BlankLNo OR LTagG(i) <> $BlankLNo OR _' A label ot Tag?
               ISTRUE (LFlagG(i) AND (%Top OR %Bottom OR %Word OR %Mark OR %Mask OR %Prof OR %Tabs OR %File OR %Cols OR %EQChange OR %Bounds )) THEN ' or special line type
               IF LTagG(i) <> $BlankLNo AND L(i).LLbl <> $BlankLNo THEN ' A Label/Tag collision?
                  sPrint (LEFT$(me.LLCtlGet(i), ENV.LinNoSize), lclScheme, MarkedLine, 1) ' Print the Pending
                  sPrint (":", $$TxtHi, MarkedLine, gLNPadCol)    ' Inverse
               ELSE                                               ' No Label/Tag collision
                  lclScheme = IIF$(IsLXclude(i + 1) AND IsHideFlag, $$LnoLoUL, $$LnoLo) ' Set HIDE mode underline
                  sPrint (LEFT$(me.LLCtlGet(i), ENV.LinNoSize), lclScheme, MarkedLine, 1) ' Print the Pending
               END IF                                             '

            ELSE                                                  '
               lclScheme = IIF$(IsLXclude(i + 1) AND IsHideFlag, $$LnoLOUL, $$LnoLO) ' Set HIDE mode underline
               sPrint (RIGHT$(LLNumG(i), ENV.LinNoSize), lclScheme, MarkedLine, 1) ' No, just make it normal
            END IF                                                '
            IF IsLUser(i) THEN sPrint ("|", $$TxtHi, MarkedLine, gLNPadCol) ' Flag column gLNPadCol id %User
            me.MarkReDraw(MarkedLine)                             ' Re-draw the MARK lines
            MarkedLine = 0                                        ' Forget this marked location
         END IF                                                   '

         IF logicalrow > 2 + Prf.Cols THEN                        ' Set any new marked line
            IF S(logicalrow) > 0 THEN                             ' On an active line?
               sCalcEditBG(logicalrow)                            ' Calc banding value
               i = sGetIX(logicalrow)                             ' Get IX of real data line if still on screen
               IF me.LLCtlGet(i) = LLNumG(i) THEN                 ' A Pending line command?
                  lclScheme = IIF$(IsLXclude(i + 1) AND IsHideFlag, $$LnoHiUL, $$LnoHi) ' Set HIDE mode underline
                  sPrint (RIGHT$(LLNumG(i), ENV.LinNoSize), lclScheme, logicalrow, 1) ' No, just make the line number
               ELSE                                               '
                  lclScheme = IIF$(IsLXclude(i + 1) AND IsHideFlag, $$LnoHiUL, $$LnoHi) ' Set HIDE mode underline
                  IF LTagG(i) <> $BlankLNo AND L(i).LLbl <> $BlankLNo THEN ' A Label/Tag collision?
                     sPrint (LEFT$(me.LLCtlGet(i), ENV.LinNoSize), lclScheme, logicalrow, 1) ' Print the Pending
                     sPrint (":", $$LnoHi, logicalrow, gLNPadCol) ' Inverse
                  ELSE                                            ' No Label/Tag collision
                     sPrint (LEFT$(me.LLCtlGet(i), ENV.LinNoSize), lclScheme, logicalrow, 1) ' Print the Pending
                  END IF                                          '
               END IF                                             '
               IF IsLUser(i) THEN sPrint ("|", $$TxtHi, logicalrow, gLNPadCol) ' Flag column gLNPadCol as %User
               MarkedLine = logicalrow                            ' Remember where we did this
            END IF                                                '
         END IF                                                   '
         MExit
      END METHOD

      METHOD ModSet(ln AS LONG)
      '---------- Set Modified flags for the edit
      REGISTER i AS LONG
      REGISTER j AS LONG
         MEntry
         IF IsLTabs(ln) THEN OnTabsAFlag: MExitMeth               ' Watch for special lines
         IF IsLWord(ln) THEN OnWordAFlag: MExitMeth               '
         IF IsLMark(ln) THEN OnMarkAFlag: MExitMeth               '
         IF IsLMask(ln) THEN OnMaskAFlag: MExitMeth               '
         IF IsLBounds(ln) THEN OnBndsAFlag: MExitMeth             '
         IF (ISFALSE IsLData(ln) AND ISFALSE IsLTop(ln) AND ISFALSE IsLFile(ln)) THEN MExitMeth ' Only set MODD for Data lines
         OnModdFlag: OnUndoFlag: OnModdEver                       ' Remember we changed something
         me.LFlagBitOff(ln, %EQChange)                            ' Reset ==CHG>
         IF IsLInsertLine(ln) THEN                                ' Is this an Insert line?
            me.LFlagBitOff(ln, %InsertLine)                       '
            InsLastClr = ln                                       ' Save Last line number cleared
         END IF                                                   '
         IF IsMEdit THEN                                          ' In a multi edit session, more to do
            me.MEditFlagSet(LMixG(ln), %True)                     ' Flag modified in the MEdit table
            FOR i = 1 TO LastLine                                 ' Find =FILE line
               IF IsLFile(i) THEN                                 ' A =FILE line?
                  INCR j                                          ' Count it
                  IF j = LMIXG(ln) THEN                           ' The one we want?
                     me.UpdLControl(i)                            ' Update it
                     MExitMeth                                    ' We're all done
                  END IF                                          '
               END IF                                             '
            NEXT i                                                '
         END IF                                                   '
         me.UpdLControl(ln)                                       ' Update LLCtl
         MExit                                                    '
      END METHOD

      METHOD  NestAdd(NewPath AS STRING, NewMask AS STRING, NewFileListNm AS STRING)
      '---------- Add a nested FM structure
         MEntry                                                   '
         INCR NestCtr                                             ' Incr Nest count
         IF NestCtr > UBOUND(NestTbl()) THEN _                    '
            REDIM PRESERVE NestTbl(1 TO NestCtr + 50) AS INSTANCE STRING
         NestTbl(NestCtr) = FPath + "|" + FMask + "|" + FileListNm + "|" + FORMAT$(TopScrn) + "|" + FORMAT$(CRow)  ' Save the key values
         FPath       = NewPath                                    ' Insert the passed values
         FMask       = NewMask                                    '
         FileListNm  = NewFileListNm                              '
         LFPath      = UUCASE(FPath)                              ' Save as 'previous' values
         LFMask      = UUCASE(FMask)                              '
         LFileListNm = UUCASE(FileListNm)                         '
         DoSet(%LoadReq)                                          ' Trigger initial Req
         MExit                                                    '
      END METHOD

      METHOD  NestDel()
      '---------- Remove a nested FM structure
         MEntry                                                   '
         IF NestCtr = 0 THEN MExitMeth                            ' Oops! Do nothing
         FPath       = PARSE$(NestTbl(NestCtr), "|", 1)           ' Restore the key values
         FMask       = PARSE$(NestTbl(NestCtr), "|", 2)           '
         FileListNm  = PARSE$(NestTbl(NestCtr), "|", 3)           '
         NestTopScrn = VAL(PARSE$(NestTbl(NestCtr), "|", 4))      '
         NestCRow    = VAL(PARSE$(NestTbl(NestCtr), "|", 5))      '
         LFPath      = UUCASE(FPath)                              ' Save as 'previous' values
         LFMask      = UUCASE(FMask)                              '
         LFileListNm = UUCASE(FileListNm)                         '
         DECR NestCtr                                             '
         DoSet(%LoadReq)                                          ' Trigger initial Req and Data load
         MExit                                                    '
      END METHOD

      METHOD NFSearch(MODE AS INTEGER, FCursMode AS INTEGER) AS INTEGER
      '---------- Do the NF (Not Found) version of full string Search procedure
      LOCAL i, j, k, l, lclLBnd, lclRBnd AS LONG, t1, t2, lclWord, c1 AS STRING, ONCE, RevOnce AS LONG
         MEntry
         lclWord = Prf.PWord                                      ' Fetch local copies
         ONCE = %True                                             '
         METHOD = %False                                          ' Start off saying we've failed
         IF LastLine > 50000 THEN gLoopCtr = - 1                  ' Prevent loop detection for big files
         IF IsRevOnce THEN                                        ' Reverse one time only?
            OffRevOnce: RevOnce = %True                           ' Remove globally, remember locally
         END IF                                                   '
         t2 = cfFind                                              ' Setup data from Search arg
         IF ISFALSE IsCfL1CaseComp AND _                          '
            ISFALSE IsCfL1RegEx AND _                             '
            ISFALSE IsCfL1Picture THEN t2 = UUCASE(t2)            ' UC if not case sensitive search or RegEx

         IF MODE = 0 THEN                                         ' Mode = 0 = Full Find
            sLine = TopScrn: sCol = 1: sDir = 1                   ' Set start defaults

            IF ISFALSE IsCfAll AND _                              ' If not ALL
               (FCursMode = %CursData OR FCursMode = %CursLinn) THEN    ' And in the Data/Line Area
               sLine = sGetIX(CRow)                               ' Use the cursor line
            END IF                                                '
            IF IsCfAll THEN                                       ' ALL
               sLine = 1                                          ' Start at line 1

            ELSEIF IsCfFirst THEN                                 ' FIRST
               sLine = 1: CfClear(%CRTFirst)                      '

            ELSEIF IsCfLast THEN                                  ' LAST
               sLine = LastLine: sDir = -1                        '
               CfClear(%CRTLast)                                  '
               CfClear(%CRTNext)                                  '
               CfSet(%CRTPrev)                                    '

            ELSEIF IsCfPrev THEN                                  ' PREV
               sDir = -1: sCol = -1                               '
               sLine = MAX(sLine - 1, 1)                          '
            END IF                                                '

         ELSEIF MODE = 1 THEN                                     ' Mode = 1 = ReFind
            IF cfNotFound = 0  OR cfNotFound = 3 THEN             ' If not a Top/Bottom not found situation
               cfNotFound = 0                                     ' Clear in case 3
               IF FCursMode = %CursCmnd THEN                      ' In the Command area?
                  sLine = TopScrn                                 ' Swap in top line of screen
               ELSEIF (FCursMode = %CursData OR FCursMode = %CursLinN) THEN ' No, In the Data/Line Area?
                  l = sGetIX(CRow)                                ' Get line cursor is on
                  IF l AND IsLXclude(l) THEN                      ' Is this an Xclude line?
                     i = CsrLinDX                                 '
                     IF CsrLinDX <> 0 THEN sLine = sFindLineNum(CsrLinDX) ' Continuing from within an X'd block?  Use that line number then
                  ELSE                                            '
                     sLine = sGetIX(CRow)                         ' Swap in current cursor location
                  END IF                                          '
                  sLine = sLine + IIF(RevOnce, -1 * sDir, sDir)   ' NF continues at next line
               ELSE                                               '
                  sLine = sLine + IIF(RevOnce, -1 * sDir, sDir)   ' NF continues at next line
               END IF                                             '
               sCol = 1                                           '
               IF IIF(RevOnce, -1 * sDir, sDir) = -1 THEN sCol = -1 ' Switch if backward
            ELSE                                                  ' Last one hit top/bottom
               sLine = IIF(cfNotFound = 1, LastLine, 1)           ' Set restart point
               cfNotFound = 0                                     ' Clear, we wrap just once
            END IF                                                '
         END IF                                                   '

         pfFind     = cfFind                                      ' Save now as a Previous set of parameters
         pfChange   = cfChange                                    '
         pfOChange  = cfOChange                                   '
         pfFLen     = cfFLen                                      '
         pfClen     = cfClen                                      '
         pfFlag     = cfFlag                                      '
         pfFLine    = cfFLine                                     '
         pfFCol     = cfFCol                                      '
         pfNotFound = cfNotFound                                  '

         i = sLine                                                ' Init for DO loop
         DO WHILE 1 = 1                                           ' Do forever
            IF gfInterrupt THEN EXIT DO                           ' Break for an FF search
            IF (IIF(RevOnce, -1 * sDir, sDir) = 1 AND IsLBottom(i)) OR _  ' Do till top or bottom
               (IIF(RevOnce, -1 * sDir, sDir) = -1 AND IsLTop(i)) THEN EXIT DO '

            IF ISFALSE me.RangeVal(i) OR _                        ' Do basic filtering
               ISFALSE IsLData(i) THEN                            '
               i += IIF(RevOnce, -1 * sDir, sDir)                 ' Next entry
               ONCE = %False                                      ' Kill column search start
               ITERATE DO                                         '
            END IF                                                '

            IF IsCfNX THEN                                        ' Do the NX test a little differently
               IF IsLInvisible(i) THEN                            ' Not Visible?
                  ONCE = %False                                   ' Skip this guy
                  i += IIF(RevOnce, -1 * sDir, sDir)              '
                  ITERATE DO                                      '
               END IF                                             '
            END IF                                                '

            IF IsCfX THEN                                         ' Do the X test a little differently
               IF ISFALSE IsLInvisible(i) AND ISFALSE IsLPopped(i) THEN ' Visible and not just Popped?
                  ONCE = %False                                   ' Skip this guy
                  i += IIF(RevOnce, -1 * sDir, sDir)              '
                  ITERATE DO                                      '
               END IF                                             '
            END IF                                                '

            IF IsCfU THEN                                         ' Do the U test
               IF ISFALSE IsLUser(i) THEN                         ' Not a User line?
                  ONCE = %False                                   ' Skip this guy
                  i += IIF(RevOnce, -1 * sDir, sDir)              '
                  ITERATE DO                                      '
               END IF                                             '
            END IF                                                '

            '----- Do the NU line filtering
            IF IsCfNU THEN                                        ' Do the NU test
               IF IsLUser(i) THEN                                 ' Is this a User line?
                  ONCE = %False                                   ' Skip this guy
                  i += IIF(RevOnce, -1 * sDir, sDir)              '
                  ITERATE DO                                      '
               END IF                                             '
            END IF                                                '

            '----- Stuff in command index as a -ve value (Looked at, not selected)
            IF ABS(LPCmdG(i)) <> PCmdIX THEN _                    ' If this command hasn't been here before (e.g. FIND ALL xxx)
               LPCmdS(i) = PCmdIX                                 ' Stuff in PCmdIX to make it :ZNF   (Neg searches have :ZF/:ZNF reversed)

            t1 = LTxtG(i)                                         ' Setup data from Txt line
            IF ISFALSE IsCfL1CaseComp AND _                       '
               ISFALSE IsCfL1RegEx AND _                          '
               ISFALSE IsCfL1Picture THEN t1 = UUCASE(t1)         ' Uppercase if not case sensitive search or picture type
            IF ONCE THEN                                          ' For 1st line, honour start col
               ONCE = %False                                      ' Just once please
               IF IsCfL1Picture THEN                              ' Picture or RegEx type search?
                  me.PicSearch(t1 , t2 , sCol, j, k)              ' Call search routine
               ELSEIF IsCfL1RegEx THEN                            ' RegEx type search?
                  PCRE_Regex_Test(t1, sCol, j, k)                 ' See if we can find it
               ELSE                                               '
                  j = INSTR(sCol, t1, t2)                         ' OK, see if its there
                  k = LEN(t2)                                     '
               END IF                                             '
            ELSE                                                  '
               IF IIF(RevOnce, -1 * sDir, sDir) = -1 THEN         ' If LAST or PREV then start as the end of the line
                     IF IsCfL1Picture THEN                        ' Picture type search?
                        me.PicSearch(t1 , t2 , sCol , j, k)       ' Call search routine
                     ELSEIF IsCfL1RegEx THEN                      ' RegEx type search?
                        PCRE_Regex_Test(t1, 1, j, k)              ' See if we can find it
                     ELSE                                         '
                        j = INSTR(-1, t1, t2)                     '
                        k = LEN(t2)                               '
                     END IF                                       '
               ELSE                                               '
                  IF IsCfL1Picture THEN                           ' Picture type search?
                     me.PicSearch(t1 , t2 , sCol , j, k)          ' Call search routine
                  ELSEIF IsCfL1RegEx THEN                         ' RegEx type search?
                     PCRE_Regex_Test(t1, 1, j, k)                 ' See if we can find it
                  ELSE                                            '
                     j = INSTR(t1, t2)                            ' OK, see if its there
                     k = LEN(t2)                                  '
                  END IF                                          '
               END IF                                             '
            END IF                                                '

            IF j THEN                                             ' If we have a 'found' string condition

               IF IsCfFCol AND ISFALSE IsCfTCol THEN              ' Request for locate in single column?
                  IF j <> cfLMargin THEN                          ' then is this in the right column?
                     GOSUB IsLineDone                             ' No, keep looking
                     ITERATE DO                                   '
                  END IF                                          '
               ELSEIF IsCfFCol AND IsCfTCol THEN                  ' Request within Left/Right columns?
                  IF j < cfLMargin OR j + k - 1 > cfRMargin THEN  ' Is it within bounds?
                     GOSUB IsLineDone                             ' No, keep looking
                     ITERATE DO                                   '
                  END IF                                          '
               END IF                                             '

               IF IsCfWord THEN                                   ' WORD requested?
                  IF j = 1 AND j + k - 1 = LEN(t1) THEN           ' If in left margin and only thing on line
                     i += IIF(RevOnce, -1 * sDir, sDir): ITERATE DO ' We found it, keep looking
                  ELSEIF j = 1 AND j + k - 1 <> LEN(t1) THEN      ' In left margin, more on line
                     IF INSTR(lclWord, MID$(t1, j + k, 1)) = 0 THEN ' Trailing char is a delimiter
                        i += IIF(RevOnce, -1 * sDir, sDir): ITERATE DO ' We found it, keep looking
                     ELSE                                         '
                        GOSUB IsLineDone: ITERATE DO              '
                     END IF                                       '
                  ELSEIF INSTR(lclWord, MID$(t1, j - 1, 1)) = 0 THEN  ' Left end is after a DLM
                     IF j + k - 1 = LEN(t1) THEN                  ' and right is at the end
                        i += IIF(RevOnce, -1 * sDir, sDir): ITERATE DO ' We found it, keep looking
                     ELSE                                         '
                        IF INSTR(lclWord, MID$(t1, j + k, 1)) = 0 THEN  ' Trailing char is a delimiter
                           i += IIF(RevOnce, -1 * sDir, sDir): ITERATE DO ' We found it, keep looking
                        ELSE                                      '
                           GOSUB IsLineDone: ITERATE DO           '
                        END IF                                    '
                     END IF                                       '
                  ELSE                                            '
                     GOSUB IsLineDone: ITERATE DO                 ' Keep looking
                  END IF                                          '
               END IF                                             '

               IF IsCfPrefix THEN                                 ' PREFIX requested?
                  IF j = 1 THEN                                   ' If in left margin
                     i += IIF(RevOnce, -1 * sDir, sDir): ITERATE DO ' We found it, keep looking
                  ELSEIF INSTR(lclWord, MID$(t1, j - 1, 1)) = 0 THEN' Left end after a delimiter
                     i += IIF(RevOnce, -1 * sDir, sDir): ITERATE DO ' We found it, keep looking
                  ELSE                                            '
                     GOSUB IsLineDone: ITERATE DO                 '
                  END IF                                          '
               END IF                                             '

               IF IsCfSuffix THEN                                 ' SUFFIX requested?
                  IF j + k - 1 = LEN(t1) THEN                     ' If only thing on line
                     i += IIF(RevOnce, -1 * sDir, sDir): ITERATE DO ' We found it, keep looking
                  ELSEIF INSTR(lclWord, MID$(t1, j + k, 1)) = 0 THEN' Trailing char is a delimiter
                     i += IIF(RevOnce, -1 * sDir, sDir): ITERATE DO ' We found it, keep looking
                  ELSE                                            '
                     GOSUB IsLineDone: ITERATE DO                 '
                  END IF                                          '
               END IF                                             '

               '----- Do the LM test if asked for
               IF IsCfLM THEN                                     ' LM requested?
                  IF j = Prf.BndLeft THEN                            ' OK if in LM column
                     i += IIF(RevOnce, -1 * sDir, sDir): ITERATE DO ' We found it, keep looking
                  ELSE                                            ' Else fail it
                     GOSUB IsLineDone: ITERATE DO                 '
                  END IF                                          '
               END IF                                             '

               '----- Do the RM test if asked for
               IF IsCfRM THEN                                     ' RM requested?
                  lclRBnd = IIF(Prf.BndRight > 0, Prf.BndRight, LEN(L(i).@LTxt)) ' Set where RM is to be found
                  IF j + k - 1 = lclRBnd THEN                     ' OK if in RM column
                     i += IIF(RevOnce, -1 * sDir, sDir): ITERATE DO ' We found it, keep looking
                  ELSE                                            ' Else fail it
                     GOSUB IsLineDone: ITERATE DO                 '
                  END IF                                          '
               END IF                                             '

               '----- Do the Color tests if asked for
               IF IsCfSTD OR _                                    ' Any color type requests?
                  cfHiLiteClr <> 0 OR IsCfSOLID OR IsCfMSOLID THEN
                  c1 = me.AttrHiLiteGet(i, j, j + k - 1)          ' Get the string

                  IF IsCfStd AND c1 = REPEAT$(k, CHR$(0)) THEN    ' Lets try STD first
                     i += IIF(RevOnce, -1 * sDir, sDir): ITERATE DO ' We found it, keep looking
                  END IF                                          '

                  IF c1 = REPEAT$(k, CHR$(cfHiLiteClr)) THEN       ' Matches requested color?
                     i += IIF(RevOnce, -1 * sDir, sDir): ITERATE DO ' We found it, keep looking
                  END IF                                          '


                  IF IsCfSOLID THEN                               ' Lets try SOLID next
                     IF c1 = REPEAT$(k, LEFT$(c1, 1)) AND _       ' Attr line field all the same color?
                        c1 <> REPEAT$(k, CHR$(0)) THEN _          ' and not STD
                        i += IIF(RevOnce, -1 * sDir, sDir): ITERATE DO ' We found it, keep looking
                  END IF                                          '

                  IF IsCfMSOLID THEN                              ' Lets try -SOLID next
                     IF c1 <> REPEAT$(k, LEFT$(c1, 1)) THEN   _   ' Attr line field not all the same color?
                        i += IIF(RevOnce, -1 * sDir, sDir): ITERATE DO ' We found it, keep looking
                  END IF                                          '

                  GOSUB IsLineDone: ITERATE DO                    '
               END IF                                             '

               i += IIF(RevOnce, -1 * sDir, sDir): ITERATE DO     ' A valid string compare winner, keep looking
            END IF                                                '

            GOTO MarkANotTrue                                     ' Not found, what we're looking for

         LOOP                                                     '
         MExitMeth                                                '

         NextLine:                                                ' Onward to next line
            IF IIF(RevOnce, -1 * sDir, sDir) = 1 THEN             ' Forward
               sCol = j + 1                                       ' This isn't the one, set to continue at next column
               IF sCol > LEN(t1) THEN                             ' End of record?
                  GOTO MarkANotTrue                               ' Yes, we have a 'not found' line
               ELSE                                               '
                  ONCE = %True                                    ' Say to honour column, stay on same line
               END IF                                             '
            ELSE                                                  '
               IF ABS(sCol) = LEN(t1) THEN                        ' Beginning of record?
                  GOTO MarkANotTrue                               ' Yes, we have a 'not found' line
               ELSE                                               '
                  ONCE = %True                                    ' No, Honour column, stay on same line
               END IF                                             '
               sCol = j - LEN(t1) -2                              ' Backward, adjust the other way
            END IF                                                '
         RETURN                                                   ' We'll never get here

         IsLineDone:                                              ' See if this line is done or not
            IF IIF(RevOnce, -1 * sDir, sDir) = 1 THEN             ' Forward
               sCol = j + 1                                       ' This isn't the one, set to continue at next column
               IF sCol > LEN(t1) THEN                             ' End of record?
                  GOTO MarkANotTrue                               ' Yes, we have a 'not found' line
               ELSE                                               '
                  ONCE = %True                                    ' Say to honour column, stay on same line
               END IF                                             '
            ELSE                                                  '
               IF ABS(sCol) = LEN(t1) THEN                        ' Beginning of record?
                  GOTO MarkANotTrue                               ' Yes, we have a 'not found' line
               ELSE                                               '
                  ONCE = %True                                    ' No, Honour column, stay on same line
               END IF                                             '
               sCol = j - LEN(t1) -2                              ' Backward, adjust the other way
            END IF                                                '
         RETURN                                                   ' We'll never get here

         MarkANotTrue:
            METHOD = %True                                        '
            INCR FixCtr                                           ' Bump FIXCtr
            LFIXS(i) = FixCtr                                     ' Save in line control area

            '----- Stuff in command index as a -ve value (selected in an NF kind of way)
            LPCmdS(i) = 0 - PCmdIX                                ' Stuff in negated PCmdIX to make it :ZNF

            IF ISFALSE IsLInvisible(i) THEN                       ' A normal non-X'd line?
               IF IsCfMX THEN                                     ' If caller is someone who wants X out the line
                  me.LFlagBitOn(i, %Invisible)                    ' Just make it invisible
                  me.LFlagBitOn(i, %Popped)                       ' Say we've just popped it
                  gfXRebuild = %True                              ' Rebuild exclude state
                  me.CurSetReq(%Find, i, 1, %True, %True)         ' Put cursor on the X line, Col 1, no hilite, XPtr to be set
                  CsrLinDX = VAL(LLNumG(i))                       ' Save hidden line number for Status Bar display
               ELSE                                               ' This is someone who is just looking like FIND / CHANGE'
                  IF ISTRUE Prf.HiFind THEN                       ' See whether to hilite the find string
                     me.CurSetReq(%Find, i, 1, %True)             ' Set cursor set attempt
                  ELSE                                            '
                     me.CurSetReq(%Find, i, 1, %True)             ' Do it without hilite
                  END IF                                          '
               END IF                                             '
            ELSE                                                  ' We have a line already X'd out
               IF IsCfMX THEN                                     ' If caller is someone who wants X out the line
                  me.CurSetReq(%Find, i, 1, %True, %True)         ' Just set the cursor
                  CsrLinDX = VAL(LLNumG(i))                       ' Save hidden line number for Status Bar display
               ELSE                                               '
                  IF IsCfDX THEN                                  ' If told to leave X status alone
                     me.CurSetReq(%Find, i, 1, %True, %True)      ' Put cursor on the X line, Col 1, no hilite, XPtr to be set
                     CsrLinDX = VAL(LLNumG(i))                    ' Save hidden line number for Status Bar display
                  ELSE                                            '
                     me.LFlagBitOff(i, %Invisible)                ' Remove the invisibility flag
                     me.LFlagBitOn(i, %Popped)                    ' Say we've just popped it
                     gfXRebuild = %True                           ' Rebuild exclude state
                     IF ISTRUE Prf.HiFind THEN                    ' See whether to hilite the find string
                        me.CurSetReq(%Find, i, 1, %True)          ' Set cursor set attempt
                     ELSE                                         '
                        me.CurSetReq(%Find, i, 1, %True)          ' Do it without hilite
                     END IF                                       '
                  END IF                                          '
               END IF                                             '
            END IF                                                '
            sLine = i: sCol = 1                                   ' Save line and column where found
            cfFLine = i: cfFCol = 1: cfFLen = 1                   ' Save in global areas
            IF IIF(RevOnce, -1 * sDir, sDir) = -1 THEN sCol = j - LEN(t1) -2  ' Backward, adjust the other way
            MExitMeth                                             ' Bail out
        RETURN                                                    ' We'll never do the return
      END METHOD

      METHOD  NoteSearch(fname AS STRING) AS STRING
      '---------- See if Note data for a new entry
      LOCAL i AS LONG
         ARRAY SCAN RNList() FOR RNCount, FROM 1 TO LEN(fname), COLLATE UCASE, = fname, TO i
         METHOD = IIF$(i = 0, " ", MID$(RNList(i), INSTR(RNList(i), "|") + 1))
      END METHOD

      METHOD  NoteUpdate(fname AS STRING, fNote AS STRING) AS STRING
      '---------- Update Note data for an entry
      LOCAL i AS LONG, RQPath, RQMask, RQFlags, RQNote AS STRING
         ARRAY SCAN RQList() FOR RQCount, FROM 1 TO LEN(fname), COLLATE UCASE, = fname, TO i
         IF i THEN                                                ' If in the RQ List, update it
            me.RQSplit(RQList(i), RQPath, RQMask, RQFlags, RQNote)' Split up operands
            RQNote = fNote                                        ' Copy new Note value
            RQList(i) = BUILD$(RQPath, "|", RQMask, "|", RQFlags, "|", RQNote) ' Write it back
         ELSE                                                     ' Must create a new Generic entry
            me.IncRQSize                                          ' Add entry to RQ list
            RQList(RQCount) = BUILD$(fname, "||G|", fNote)        '
         END IF
         DoSet(%SaveReq)                                          ' Request a REQ list save
      END METHOD

      METHOD PostKeyBoard()
      '---------- Do post keyboard activities if requested
         LOCAL i, j, currtab, InUse, ActionSaveNeeded AS LONG, fn, fn2, Cmd, ErrMsg, lclDrive, lclPath, DelDrive, DelPath, Delfn, RenFn, t, tt AS STRING
         LOCAL SaveCmd AS STRING
         MEntry
         IF IsDoAttention THEN                                    ' Attention requested?

            '----- Reset the Attention flag, set a default low priority cursor location request
            DoClear(%Attention)                                   ' Clear it right away
            DoSet(%Refresh)                                       ' Attention always does a refresh as well
            ErrFlag = 0: ErrMsg = ""                              ' Start clean

            IF LEFT$(TRIM$(pCommand), 1) = "&" THEN               ' Handle any & prefixed command line?
               pCommandPrv = pCommand                             ' Save & prefixed commands
               pCommand = MID$(pCommand, INSTR(pCommand, "&") + 1)' Strip the &
            END IF

            '----- Setup the ~w and ~l macro variables if not FM mode
            IF ISFALSE IsFMTab THEN                               ' Only do for non FM screens
               IF CRow > 2 + Prf.Cols AND CRow <= gwScrHeight THEN' Save data line if cursor in data area
                  j = sGetIX(CRow)                                ' Locate and save data line cursor was on
                  IF j < 0 THEN                                   '
                     me.CsrRowSub(ABS(j))                         ' Backup to 1st line
                     j = sGetIX(CRow)                             ' Backward to real data, not hex
                  END IF                                          '
                  CsrAbsLine = j                                  '
               ELSE                                               '
                  CsrAbsLine = 0                                  ' else null
               END IF                                             '
               me.CursorWord()                                    ' Go see if extractable data under the cursor

               '----- Also do special processing if any of WORD, MARK, TABS or BNDS lines were modified.
               IF IsWordAFlag THEN me.WordSave                    ' Validate WORD if the user has modified them
               IF IsMarkAFlag THEN me.MarkSave                    ' Validate Mark if the user has modified them
               IF IsMaskAFlag THEN me.MaskSave                    ' Validate Mask if the user has modified them
               IF IsTabsAFlag THEN me.TabsSave                    ' Validate TABS if the user has modified them
               IF IsBndsAFlag THEN me.BndsSave                    ' Validate BNDS if the user has modified them
            END IF                                                '

            '----- Must validate Scrl amount on each Attention (Ugh)
            me.ValScroll()                                        ' Go see if Scroll Amount is reasonable

            '----- Handle CmdSepChr stuff, stack the individual commands
            TP.MacSubst(pCommand)                                 ' Do substitution
            IF me.CmdAddStack() = %False THEN                     ' Process I/P line to multiple commands and see if RESET

               IF IsFMTab THEN                                    ' If not RESET, do either File Manager version
                  IF me.FMLCmdValidate THEN GOTO RefreshOut       ' Go do so, Skip processing if an error
                  me.FMLCmdSpread                                 ' If OK, spread out the repeaters
                  me.FMLCmdProcess                                ' Process them now
                  IF ErrFlag > 0 THEN GOTO RefreshOut             ' Errors, skip out

               ELSE                                               ' Or Normal Edit version of line commands
                  '----- Scan the Edit Line commands
                  me.LCtlProcess                                  ' Let LCtlProcess have a go
                  IF LTblBIX > 0 THEN                             ' We end up with items in the B table?
                     me.LCtlCommands                              ' Go do them
                  END IF
               END IF                                             '
            END IF                                                '

            '---------- Set ACTION for typing and Line command changes
            IF IsUndoFlag AND ActionCtr THEN                      ' Set SAVE needed
               ActionSaveNeeded = %True                           ' Say SAVE might be needed
            END IF

            '----- Lets handle the command line if we get here
            DO WHILE (ISNOTNULL(TRIM$(pCommand)) OR CmdStackNum > 0) AND errFlag <= %eCmdPend ' Do some commands

               '----- Get next stacked command if needed
               IF ISNULL(TRIM$(pCommand)) THEN                    ' If no command
                  pCommand = me.CmdStackNext                      ' Get one from the stack
               END IF                                             '
               IF ISNULL(TRIM$(pCommand)) THEN ITERATE DO         ' If still null, ignore it

               '----- Go route and execute the command line
               me.CmdAssign                                       '
               IF TabsNum = 0 THEN EXIT METHOD                    ' If terminating , exit quickly

               '----- If we just did a RETRIEVE, we're done
               IF errFlag = %eRetrieve THEN SetCmd: EXIT DO       '

               '---------- Use this location to handle the ACTION option for chained commands
               IF (IsUndoFlag AND ActionCtr) OR ActionSaveNeeded THEN  ' If UndoFlag and ACTION running
                  GOSUB DoActionSave                              ' Yes, go do it
               END IF                                             '

               '----- If no errors, clear the command
               IF errFlag <> %eNone AND errFlag <> %ePending THEN EXIT DO  ' If some kind of error, exit Cmd loop
               pCommand = ""                                      '

               ' Blank the Command now if not needed
               me.WindowCmd                                       ' Re-do the command line
               IF NestTopScrn = 0 THEN SetCmd                     ' Put Cursor on the line if no pending Nest request

               '----- Reset Retrieve stuff, watch out for Pending
               gCmdRtrevIX = 1: gCmdRtrevMsg = ""                 '
               IF errFlag = %eCmdPend THEN EXIT DO                ' If CmdPend let scrn display etc. happen
            LOOP                                                  '

            '----- Do some stuff now based on FM / Non-FM tab
            IF ISFALSE IsFMTab THEN                               ' If not FM

               '---------- Do actual SAVE for Typing and Line Command Changes
               IF ActionSaveNeeded THEN                           ' If save WAITING
                  GOSUB DoActionSave                              ' Yes, go do it
               END IF                                             '

               '----- Reset flags; Clean '''''' lines; take an UNDO Chkpt
               OffPFKInsert                                       ' Clear flags now
               me.CleanInsLines()                                 ' Go remove unused Insert Lines
               me.UndoSave                                        ' Go take an Undo save if needed

            ELSE                                                  ' It IS an FM Tab
               IF IsDoSaveReq THEN me.SaveReq                     ' Save Req list if any action above modified something

               '----- See if Path/Mask/etc have changed
               FPath = TRIM$(FPath)                               ' Clean up
               IF RIGHT$(FPath, 1) <> "\" THEN FPath += "\"       ' Ensure FPath has trailing \
               FMask = TRIM$(FMask)                               '
               IF IsNE(LScrlAmtC, TRIM$(Prf.Scroll)) THEN _       ' Do Scrl separately
                  sIniSetString("FManager", "ScrlAmtC", Prf.Scroll) ' So refresh isn't triggered

               IF IsNE(LFPath, TRIM$(FPath)) OR _                 ' User changed criteria?
                  IsNE(lFMask, TRIM$(FMask)) OR _                 '
                  IsNE(lFileListNm, TRIM$(FileListNm)) THEN       '
                  LastTop = 0: DoSet(%LoadReq)                    ' Kill position

                  IF ISNULL(FMask) THEN FMask = "*"               '
                  IF ISNULL(FileListNm) THEN                      ' If this is a normal Directory
                     sIniSetString("FManager", "DefDir1", FPath)  '
                     sIniSetString("FManager", "DefTypes", FMask) '
                  ELSE                                            ' Else save FileListNm
                     sIniSetString("FManager", "Recall", FileListNm) '
                  END IF                                          '
               END IF                                             '
            END IF                                                '
            IF ISFALSE gMacroMode THEN                            ' If not macro mode
               IF ENV.InsReset THEN                               ' If ENV.InsReset requested
                  IF ENV.InsMode THEN OnNsrtFlag ELSE OffNsrtFlag ' then pick up the default
                  IF IsNsrtFlag THEN OffNsrtFlag ELSE OnNsrtFlag  ' Flip/Flop since krInsert also does this
                  me.krInsert                                     ' and go set it
               END IF                                             '
            END IF                                                '
            IF IsDOMarkKill THEN                                  ' Kill marked area?
               IF IsDoMarkKillSkip THEN                           ' Skip once?
                  DoClear(%MarkKillSkip)                          ' Yes, just clear the flag
               ELSE
                  me.MarkKill                                     ' Do so
                  DoClear(%MarkKill)                              ' Clear flag
               END IF                                             '
            END IF                                                '
         END IF                                                   '

      RefreshOut:
         IF IsDoRefresh OR IsDoMsg OR ErrFlag > 0 THEN            ' Refresh requested or an Error?
            me.DispScreen                                         ' Then re-display
         END IF                                                   '
         MExitMeth

      DoActionSave:
         DECR ActionCtr                                           ' OK, decrement it
         IF ActionCtr > 0 OR ActionStop THEN RETURN               ' Not time to do a SAVE or Stopped by END?
         ActionCtr = Prf.ActionSave                               ' Yes, reset the counter for next time
         IF IsMEdit AND Prf.ActionVerb = "VSAVE" THEN             ' Unsupported combo?
            '                                                     ' Do nothing
         ELSE                                                     '
            IF Prf.ActionVerb = "SAVE" THEN                       ' Call the right guy
               pCmdSave("SAVE" + IIF$(IsMEdit, " MeditOnly Quiet", " Quiet"))
            ELSE                                                  '
               pCmdVSave("VSAVE Quiet")                           '
            END IF                                                '
         END IF                                                   '
         ActionSaveNeeded = %False                                '
         RETURN                                                   '
      END METHOD

      '----- Following stubs allow access to the Prf Object from outside the TP Object
      METHOD PrfAutoBkup() AS LONG: METHOD = Prf.AutoBkup: END METHOD        ' Pass on request to the Prf Object
      METHOD PrfAutoSave() AS LONG: METHOD = Prf.AutoSave: END METHOD
      METHOD PrfBndLeft() AS LONG: METHOD = Prf.BndLeft: END METHOD
      METHOD PrfBndRight() AS LONG: METHOD = Prf.BndRight: END METHOD
      METHOD PrfCapsActual() AS LONG: METHOD = Prf.CapsActual: END METHOD
      METHOD PrfCapsDesired() AS LONG: METHOD = Prf.CapsDesired: END METHOD
      METHOD PrfChangeMode() AS STRING: METHOD = Prf.ChangeMode: END METHOD
      METHOD PrfCols() AS LONG: METHOD = Prf.Cols: END METHOD
      METHOD PrfEOL() AS STRING: METHOD = Prf.EOL: END METHOD
      METHOD PrfFold() AS LONG: METHOD = Prf.Fold: END METHOD
      METHOD PrfHexMode() AS LONG: METHOD = Prf.HexMode: END METHOD
      METHOD PrfLock() AS LONG: METHOD = Prf.PLock: END METHOD
      METHOD PrfMarkWorking() AS STRING: METHOD = Prf.MarkWorking: END METHOD
      METHOD PrfPageFlag() AS LONG: METHOD = Prf.PageFlag: END METHOD
      METHOD PrfPageOffset() AS LONG: METHOD = Prf.PageOffset: END METHOD
      METHOD PrfPCase() AS STRING: METHOD = Prf.PCase: END METHOD
      METHOD PrfPCollate() AS STRING: METHOD = Prf.PCollate: END METHOD
      METHOD PrfDoneCheck() AS LONG: METHOD = Prf.DoneCheck: END METHOD
      METHOD PrfPMark() AS LONG: METHOD = Prf.PMark: END METHOD
      METHOD PrfProfName() AS STRING: METHOD = Prf.ProfName: END METHOD
      METHOD PrfPSource() AS STRING: METHOD = Prf.PSource: END METHOD
      METHOD PrfPWord() AS STRING: METHOD = Prf.PWord: END METHOD
      METHOD PrfQuery(lclTxt AS STRING) AS STRING: METHOD = Prf.Query(lclTxt): END METHOD
      METHOD PrfReadAll(UseOK AS LONG): Prf.ReadAll(UseOK): END METHOD
      METHOD PrfReset(): Prf.Reset(): END METHOD
      METHOD PrfSetProfName(nm AS STRING): Prf.SetProfName(nm, %False): END METHOD
      METHOD PrfSrceXlate() AS LONG: METHOD = Prf.SrceXlate: END METHOD
      METHOD PrfCollateXlate() AS LONG: METHOD = Prf.ColateXlate: END METHOD
      METHOD PrfStart() AS STRING: METHOD = Prf.Start: END METHOD
      METHOD PrfUndoNumber() AS LONG: METHOD = Prf.UndoNumber: END METHOD
      METHOD PrfUseName() AS STRING: METHOD = Prf.UseName: END METHOD
      METHOD PrfWordVal(): Prf.WordVal: END METHOD

      METHOD PrfGetSS2APtr() AS LONG: METHOD = Prf.GetSS2APtr: END METHOD
      METHOD PrfGetSA2SPtr() AS LONG: METHOD = Prf.GetSA2SPtr: END METHOD
      METHOD PrfGetCA2SPtr() AS LONG: METHOD = Prf.GetCA2SPtr: END METHOD
      METHOD PrfGetCS2APtr() AS LONG: METHOD = Prf.GetCS2APtr: END METHOD


      METHOD ProfDisp()
      '---------- Build and display the PROF lines
      LOCAL lcltop AS LONG
         MEntry
         me.ResetFunc(%ResetSpecial)                              ' Clear any prior PROF display
         Prf.BuildLines                                           ' Build the Profile display lines (5)
         lclTop = MIN(S(3 + Prf.Cols), LastLine - 1)              ' Get line at top of scrn
         IF lclTop = 1 THEN lclTop = 2                            ' So we don't insert before the ** Top of File **
         IF lclTop = LastLine - 1 THEN                            ' Is this the last data line
            DO WHILE lclTop > 1 AND ISFALSE TP.LFlagData(lclTop)  ' If last line isn't data line, find it
               DECR lclTop                                        ' Backup
            LOOP                                                  '
         END IF                                                   '
         me.InsBounds(lclTop)                                     ' Insert the BNDS line
         me.InsCols(lclTop)                                       ' Insert the COLS line
         me.InsTabs(lclTop)                                       ' Insert the TABS line
         me.InsMask(lclTop)                                       ' Insert the MASK lin
         me.InsMark(lclTop)                                       ' Insert the MARK lin
         me.InsWord(lclTop)                                       ' Insert the WORD lin
         me.InsProf(lclTop, 6)                                    ' Insert the PROF lin
         me.InsProf(lclTop, 5)                                    ' Insert the PROF lin
         me.InsProf(lclTop, 4)                                    ' Insert the PROF lin
         me.InsProf(lclTop, 3)                                    ' Insert the PROF lin
         me.InsProf(lclTop, 2)                                    ' Insert the PROF lin
         me.InsProf(lclTop, 1)                                    ' Insert the PROF lin
         MExit
      END METHOD


      METHOD RangeSet(cmd AS STRING, Numerics AS LONG, DelNeed AS LONG) AS LONG
      '---------- Scan out range operands
      LOCAL i, j, k, l, SOpSet, sHadOp AS LONG
      LOCAL StrStart, StrSCond, StrEnd, StrECond, StrTag, StrTCond, WCond, t AS STRING

         MEntry
         DotStart = 2: DotEnd = LastLine - 1                      ' Set defaults in case no operands
         DotSCond = "GE": DotECond = "LE": DotTCond = "EQ": DotFlag = 0: DotMode = " "
         DotSet = %False: DotTag = "": DotAndOr = "AND": METHOD = %False
         IF cmd = "RESET" THEN MExitMeth                          ' Exit if just called to RESET

         FOR i = 1 TO pCmdNumOps                                  ' Lets loop through operands

            IF (pCmdOpsType(i) = %OpNum AND ISTRUE Numerics) THEN ' Numeric and allowable?
               IF ISNULL(StrStart) THEN                           ' Save as 1st if not set yet
                  StrStart = pCmdOps(i): StrSCond = "GE"          '
                  ARRAY DELETE pCmdOps(i): ARRAY DELETE pCmdRaw(i): ARRAY DELETE pCmdOpsType(i)
                  DECR i: INCR j: sHadOp = %True                  ' So we stay on the same index number
               ELSEIF ISNULL(StrEnd) THEN                         ' Or save as 2nd
                  StrEnd = pCmdOps(i): StrECond = "LE"            '
                  ARRAY DELETE pCmdOps(i): ARRAY DELETE pCmdRaw(i): ARRAY DELETE pCmdOpsType(i)
                  DECR i: INCR j: sHadOp = %True                  ' So we stay on the same index number
               END IF                                             '
            END IF                                                '

            IF pCmdOpsType(i) = %OpDotd THEN                      ' A dotted operand?
               IF ISNULL(StrStart) THEN                           ' Save as 1st if not set yet
                  WCond = "GE": GOSUB ExtCond                     ' Go Extract condition
                  StrStart = UUCASE(pCmdOps(i)): StrSCond = WCond ' Save answer
                  IF VERIFY(2, StrStart, "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_") <> 0 THEN _ ' Valid Chars?
                     METHOD = %True: MErrExit(%eFail, "Invalid line range operand")
                  ARRAY DELETE pCmdOps(i): ARRAY DELETE pCmdRaw(i): ARRAY DELETE pCmdOpsType(i)
                  DECR i: INCR j: sHadOp = %True                  ' So we stay on the same index number
               ELSEIF ISNULL(StrEnd) THEN                         ' Or save as 2nd
                  WCond = "LE": GOSUB ExtCond                     ' Go Extract condition
                  StrEnd = UUCASE(pCmdOps(i)): StrECond = WCond   '
                  IF VERIFY(2, StrEnd, "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_") <> 0 THEN _ ' Valid characters?
                     METHOD = %True: MErrExit(%eFail, "Invalid line range operand")
                  ARRAY DELETE pCmdOps(i): ARRAY DELETE pCmdRaw(i): ARRAY DELETE pCmdOpsType(i)
                  DECR i: INCR j: sHadOp = %True                  ' So we stay on the same index number
               END IF                                             '
            END IF                                                '

            IF pCmdOpsType(i) = %OpLPtr THEN                      ' An LPtr operand?
               IF ISNULL(StrStart) THEN                           ' Save as 1st if not set yet
                  WCond = "GE": GOSUB ExtCond                     ' Go Extract condition
                  StrStart = UUCASE(pCmdOps(i)): StrSCond = WCond ' Save answer
                  ARRAY DELETE pCmdOps(i): ARRAY DELETE pCmdRaw(i): ARRAY DELETE pCmdOpsType(i)
                  DECR i: INCR j: sHadOp = %True                  ' So we stay on the same index number
               ELSEIF ISNULL(StrEnd) THEN                         ' Or save as 2nd
                  WCond = "LE": GOSUB ExtCond                     ' Go Extract condition
                  StrEnd = UUCASE(pCmdOps(i)): StrECond = WCond   '
                  ARRAY DELETE pCmdOps(i): ARRAY DELETE pCmdRaw(i): ARRAY DELETE pCmdOpsType(i)
                  DECR i: INCR j: sHadOp = %True                  ' So we stay on the same index number
               END IF                                             '
            END IF                                                '

            IF pCmdOpsType(i) = %OpTag THEN                       ' A tag operand?
               IF ISNULL(StrTag) THEN                             ' Got a tag already?
                  WCond = "EQ": GOSUB ExtCond                     ' Go Extract condition
                  StrTag = UUCASE(pCmdOps(i)): StrTCond = WCond   ' Save it
                  IF sTagVal(StrTag, %True) THEN _                ' Valid?
                     METHOD = %True: MErrExit(%eFail, "Invalid tag operand")
                  ARRAY DELETE pCmdOps(i): ARRAY DELETE pCmdRaw(i): ARRAY DELETE pCmdOpsType(i)
                  DECR i: INCR j: sHadOp = %True                  ' So we stay on the same index number
               ELSE                                               '
                  METHOD = %True: MErrExit(%eFail, "More than 1 :tag operand present")
               END IF                                             '
            END IF                                                '

            IF pCmdOpsType(i) = %KWAnd THEN                       ' AND?
               DotAndOr = "AND"                                   ' Save it
               ARRAY DELETE pCmdOps(i): ARRAY DELETE pCmdRaw(i): ARRAY DELETE pCmdOpsType(i)
               DECR i: INCR j                                     ' So we stay on the same index number
            END IF                                                '

            IF pCmdOpsType(i) = %KWOr THEN                        ' OR?
               DotAndOr = "OR"                                    ' Save it
               ARRAY DELETE pCmdOps(i): ARRAY DELETE pCmdRaw(i): ARRAY DELETE pCmdOpsType(i)
               DECR i: INCR j                                     ' So we stay on the same index number
            END IF                                                '
         NEXT i                                                   '

         IF ISNULL(DotAndOr) THEN DotAndOr = "AND"                ' Default AND if not provided
         pCmdNumOps -= j                                          ' Adjust now for removed operands

         IF ISNOTNULL(StrTag) THEN                                ' If we have a tag, better be no others
            IF StrTCond <> "EQ" AND StrTCond <> "NE" THEN _       '
               METHOD = %True: MErrExit(%eFail, "Only =, <>, ¬=, and \= allowed with a :tag operand")
            DotSet = %True: DotTag = StrTag: DotTCond = StrTCond  ' Save our answer
         END IF                                                   '

         IF ISNULL(StrStart) AND ISNULL(StrEnd) THEN              ' Anything scanned out?
            IF ISFALSE LTblRange THEN MExitMeth                   ' No line range, exit
            DotMode = TRIM$(LTblSCmd)                             ' Save C / M line mode
            IF LEFT$(LTblSCmd, 1) <> "M" AND ISTRUE DelNeed THEN _'
               METHOD = %True: MErrExit(%eFail, "Line range is destructive, use M/MM line command")
            DotStart = LTblSFrom                                  ' Something there set it up
            me.TTblDel(LTblSFrom)                                 ' Remove from Touched lines
            me.UpdLControl(LTblSFrom)                             ' Clear from the line
            DotSCond = "GE"                                       '
            DotEnd = LTblSTo                                      '
            me.TTblDel(LTblSTo)                                   ' Remove from Touched lines
            me.UpdLControl(LTblSTo)                               ' Clear from the line
            DotECond = "LE"                                       '
            DotSet = %True                                        '
            DotFlag = LTblSFlag                                   '
            MExitMeth                                             '
         END IF                                                   '

         IF LTblRange THEN _                                      ' If a line range also
            METHOD = %True: MErrExit(%eFail, "Can't have both line range operands and marked line range")

         IF ISNOTNULL(StrStart) AND ISNULL(StrEnd) THEN           ' Just one operand
            DotStart = me.LineNoRef(StrStart)                     ' Yes, set 1st as start
            IF DotStart = -1 THEN _                               '
               METHOD = %True: MErrExit(%eFail, "Line reference " + StrStart + " is undefined")
            DotSCond = StrSCond                                   ' Set the conditional
            IF ISFALSE sOpSet THEN DotSCond = "EQ"                ' If it wasn't specifically set, Force EQ condition
            DotEnd = 0                                            ' End = 0 to say not there
            DotSet = %True: MExitMeth                             ' Say we're OK and leave
         END IF                                                   '

         DotStart = me.LineNoRef(StrStart)                        ' Yes, set 1st as start
         IF DotStart = -1 THEN _                                  '
            METHOD = %True: MErrExit(%eFail, "Line reference " + StrStart + " is undefined")
         DotSCond = StrSCond                                      '
         DotEnd = me.LineNoRef(StrEnd)                            ' 2nd as end
         IF DotEnd = -1 THEN _                                    '
            METHOD = %True: MErrExit(%eFail, "Line reference " + StrEnd + " is undefined")
         DotECond = StrECond                                      '

         IF DotStart > DotEnd AND _                               ' Look for the simple out of order swap
            DotSCond = "GE" AND DotECond = "LE" AND _             '
            DotAndOr = "AND" THEN                                 ' All looks safe, we can swap them
            DotStart = me.LineNoRef(StrEnd): DotEnd = me.LineNoRef(StrStart)
         END IF                                                   '

         IF (DotSCond = "LE" OR DotSCond = "LT") AND (DotECond = "GE" OR DotECond = "GT") AND _
            DotAndOr = "AND" AND _                                '
            DotStart < DotEnd THEN _                              '
               METHOD = %True: MErrExit(%eFail, "Line range illogical, represents 0 lines")

         IF DotStart < DotEnd THEN                                ' Look for another illogical set
            IF DotsCond = "LT" OR DotsCond = "LE" OR DotECond = "GE" OR DotECond = "GT" THEN _
               METHOD = %True: MErrExit(%eFail, "Line range is illogical")
         END IF                                                   '

         IF DotSCond = "NE" OR DotECond = "NE" THEN _
               METHOD = %True: MErrExit(%eFail, "Cannot use a not equal condition with two operands")

         IF DotStart = 0 OR DotEnd = 0 OR _                       ' Valid?
            DotStart > DotEnd THEN _                              '
            METHOD = %True: MErrExit(%eFail, "Start or end operand is invalid")
         IF DotStart <> 2 OR DotEnd <> LastLine - 1 OR _          ' Are we still at default .ZFIRST to .ZLAST ?
            sHadOp = %True THEN _                                 '
            DotSet = %True                                        ' No
         MExitMeth

      ExtCond:                                                    ' Extract condition from Operand
         k = VERIFY(2, pCmdOps(i), "<>=\" + CHR$(172))            ' Find 1st Label/Tagname character
         IF k = 2 THEN RETURN                                     ' No condition, return 1st op default
         IF ISNULL(StrStart) THEN sOpSet = %True                  ' Remember is specific 1st operand
         SELECT CASE AS CONST$ MID$(pCmdOps(i), 2, k - 2)         ' Examine the condition
            CASE "<=": WCond = "LE": EXIT SELECT                  '
            CASE ">=": WCond = "GE": EXIT SELECT                  '
            CASE "<>": WCond = "NE": EXIT SELECT                  '
            CASE "<":  WCond = "LT": EXIT SELECT                  '
            CASE ">":  WCond = "GT": EXIT SELECT                  '
            CASE "\":  WCond = "NE": EXIT SELECT                  '
            CASE CHR$(172):  WCond = "NE": EXIT SELECT            '
         END SELECT                                               '
         pCmdOps(i) = LEFT$(pCmdOps(i), 1) + MID$(pCmdOps(i), k)  ' Rebuild Op without the condition
         RETURN                                                   '
      END METHOD

      METHOD RangeSetRange(fline AS LONG, tline AS LONG)
      '---------- Set a specific line range
         DotStart = fline: DotEnd = tline                         ' Set the passed line range
         DotSCond = "GE": DotECond = "LE": DotTCond = "EQ": DotFlag = 0: DotMode = " "
         DotSet = %True: DotTag = "": DotAndOr = "AND"            '
      END METHOD

      METHOD RangeVal(lno AS LONG) AS LONG
      '---------- See if Line number is within specified range
      LOCAL s, e, Tst1, ans AS LONG, tag1, tag2 AS STRING * 8

         MEntry
         IF ISNOTNULL(TRIM$(DotTag)) THEN                         ' Is this a Tag 'range'?
            tag1 = LTagG(lno): tag2 = LSET$(DotTag, 8)            ' Get tag from the line and DotTag

            '----- Handle the magic :Z :ZF and :ZNF tags
            IF tag2 = ":Z      " THEN                             ' The :Z version
               IF PCmdIX = LPCmdG(lno) THEN                       ' Already found by this command, then we know it's true
                  ans = %True                                     ' So mark it True again
               ELSEIF DotTCond = "EQ" THEN                        ' The EQ version
                  ans = IIF((PCmdIX - 1) = ABS(LPCmdG(lno)), %True, %False)
               ELSE                                               ' Else do the NE test
                  ans = IIF((PCmdIX - 1) = ABS(LPCmdG(lno)), %False, %True)
               END IF
            ELSEIF tag2 = ":ZF     " THEN                         ' The :ZF version
               IF PCmdIX = LPCmdG(lno) THEN                       ' Already found by this command, then we know it's true
                  ans = %True                                     ' So mark it True again
               ELSEIF DotTCond = "EQ" THEN                        ' The EQ version
                  s = LPCmdG(lno)
                  ans = IIF((PCmdIX - 1) = LPCmdG(lno), %True, %False)
               ELSE                                               ' Else do the NE test
                  ans = IIF((PCmdIX - 1) = LPCmdG(lno), %False, %True)
               END IF
            ELSEIF tag2 = ":ZNF    " THEN                         ' The :ZNF version
               IF PCmdIX = LPCmdG(lno) THEN                       ' Already found by this command, then we know it's true
                  ans = %True                                     ' So mark it True again
               ELSEIF DotTCond = "EQ" THEN                        ' The EQ version
                  ans = IIF(((0 - (PCmdIX - 1))) = LPCmdG(lno), %True, %False)
               ELSE                                               ' Else do the NE test
                  ans = IIF(((0 - (PCmdIX - 1))) = LPCmdG(lno), %False, %True)
               END IF

            '----- Handle the normal tags
            ELSEIF DotTCond = "EQ" THEN                           ' Select test type
               IF Tag2 = ":ZALL   " THEN                          ' The ZALL guy?
                  ans = IIF(ISNOTNULL(TRIM$(tag1)), %True, %False)' Set true if any tag at all
               ELSE                                               '
                  ans = IIF(tag1 = tag2, %True, %False)           ' Set answer
               END IF                                             '
            ELSE                                                  ' NE version
               IF Tag2 = ":ZALL   " THEN                          ' The ZALL guy?
                  ans = IIF(ISNULL(TRIM$(tag1)), %True, %False)   ' Set true if no tag at all
               ELSE                                               '
                  ans = IIF(tag1 = tag2, %False, %True)           '
               END IF                                             '
            END IF                                                '
            IF ISTRUE ans THEN EXIT IF                            ' Something True, continue
            METHOD = %False                                       ' Exit failure, no more tests needed
            MExitMeth                                             '
         END IF                                                   '

         s = DotStart: e = DotEnd
         IF DotEnd = 0 THEN                                       ' Just a single operand?
            SELECT CASE AS CONST$ DotSCond                        ' Split by type of test
               CASE "LT": METHOD = IIF(lno <  DotStart, %True, %False)
               CASE "LE": METHOD = IIF(lno <= DotStart, %True, %False)
               CASE "EQ": METHOD = IIF(lno =  DotStart, %True, %False)
               CASE "NE": METHOD = IIF(lno <> DotStart, %True, %False)
               CASE "GE": METHOD = IIF(lno >= DotStart, %True, %False)
               CASE "GT": METHOD = IIF(lno >  DotStart, %True, %False)
            END SELECT                                            '
            MExitMeth                                             '
         END IF                                                   '

         SELECT CASE AS CONST$ DotSCond                           ' Two tests, do the 1st
            CASE "LT": Tst1 = IIF(lno <  DotStart, %True, %False) '
            CASE "LE": Tst1 = IIF(lno <= DotStart, %True, %False) '
            CASE "EQ": Tst1 = IIF(lno =  DotStart, %True, %False) '
            CASE "NE": Tst1 = IIF(lno <> DotStart, %True, %False) '
            CASE "GE": Tst1 = IIF(lno >= DotStart, %True, %False) '
            CASE "GT": Tst1 = IIF(lno >  DotStart, %True, %False) '
         END SELECT                                               '

         IF ISTRUE Tst1 AND DotAndOr = "OR" THEN METHOD = %True: MExitMeth      ' 1st test OK and OR, no need for the 2nd set
         IF ISFALSE Tst1 AND DotAndOr = "AND" THEN METHOD = %False: MExitMeth   ' Failed 1st test and AND, no need for the 2nd set

         SELECT CASE AS CONST$ DotECond                           ' Do the 2nd
            CASE "LT": METHOD = IIF(lno <  DotEnd, %True, %False) '
            CASE "LE": METHOD = IIF(lno <= DotEnd, %True, %False) '
            CASE "EQ": METHOD = IIF(lno =  DotEnd, %True, %False) '
            CASE "NE": METHOD = IIF(lno <> DotEnd, %True, %False) '
            CASE "GE": METHOD = IIF(lno >= DotEnd, %True, %False) '
            CASE "GT": METHOD = IIF(lno >  DotEnd, %True, %False) '
         END SELECT                                               '
         MExit
      END METHOD

      METHOD ResetFunc(Request AS LONG)
      '---------- Do a RESET function
      LOCAL i, j, lcltop, NumDel AS LONG, lCmd, t AS STRING, tPtr AS STRING POINTER, pLCtl AS STRING * 8, AttrByte AS BYTE
         MEntry
         lcltop = TopScrn                                         ' Save where we were
         IF IsLXclude(lcltop) THEN INCR lcltop                    '

         '----- Fudge line range if needed
         IF DotStart = 2 AND DotEnd = LastLine - 1 AND Request = %ResetCommand THEN
            DotStart = 1: DotEnd = LastLine                       ' Add Top/Bottom if an internal resetcommand
         END IF                                                   '

         FindWord = ENV.FindWord                                  '

         '----- Do RETRIEVE separately
         IF (Request AND %ResetRetrieve) <> 0 THEN                ' RETRIEVE
            RESET gCmdRtrev(): gCmdRtrevIX = 1                    '
         END IF                                                   '

         '----- Clear line command working tables
         IF (Request AND %ResetCommand) <> 0 THEN                 ' If %ResetCommand
            RESET LTblAIX, LTblBIX                                ' Reset LLCtl table indexes
            me.TTblReset                                          ' Reset the touched table
         END IF                                                   '

         '----- Loop thru line range
         FOR i = 1 TO LastLine                                    ' Loop through them all

            '----- Just those in the range
            IF ISFALSE me.RangeVal(i) THEN ITERATE                ' Honour range

            '----- Data, Top, Bottom and Excluded lines
            IF ISTRUE (LFlagG(i) AND (%Data OR %Top OR %Bottom OR %Xclude OR %File OR %NOTE)) THEN ' Do Data lines
               LFIXS(i) = 0                                       ' Reset the FIND index
               IF ISTRUE (LFlagG(i) AND (%Data)) THEN _           ' Do Data lines
                  LPCmdS(i) = (0 - PCmdIX)                        ' Stuff in negated PCmdIX to make it :ZNF

               '----- X indicator lines
               IF IsLXclude(i) THEN                               ' Reset X lines
                  IF (Request AND %ResetExcluded) <> 0 THEN       ' If %ResetExcluded
                     me.LTxtFree(i)                               ' Go free the dynamic string
                     me.LEntDel(i)                                ' Remove from the L() array
                     DECR i                                       ' Make sure we don't miss one
                     DECR LastLine: DECR LastReal                 ' Adjust LastLine and LastReal
                     INCR NumDel                                  ' reduce count of total lines
                     IF i <= lcltop THEN lcltop = MAX(1, lcltop - 1) ' Adjust final screen positioning
                     gfXRebuild = %True                           ' Do an X rebuild
                     ITERATE FOR                                  '
                  END IF                                          '
               END IF                                             '

               '----- Clear Invisible status
               IF (Request AND %ResetExcluded) <> 0 THEN          ' If %ResetExcluded
                  IF IsLInvisible(i) THEN LPCmdS(i) = PCmdIX      ' Make it :ZF if we changed it
                  me.LFlagBitOff(i, %Invisible)                   ' Reset Invisible
                  LWrk1S(i) = 0                                   ' Reset any Xclude back index
               END IF                                             '

               '----- Clear User status
               IF (Request AND %ResetUser) <> 0 THEN              ' If %ResetExcluded
                  IF IsLUser(i) THEN LPCmdS(i) = PCmdIX           ' Make it :ZF if we changed it
                  me.LFlagBitOff(i, %User)                        ' Reset User
               END IF                                             '

               '----- Clear ==CHG> status
               IF (Request AND %ResetChange) <> 0 THEN            ' If %ResetChange
                  IF IsLChange(i) THEN                            ' And Change is on
                     me.LFlagBitOff(i, %EQChange)                 ' Reset ==CHG>
                     me.UpdLControl(i)                            ' Redo the LLCtl
                     LPCmdS(i) = PCmdIX                           ' Stuff in PCmdIX to make it :ZF
                  END IF                                          '
               END IF                                             '

               '----- Clear line Tags
               IF (Request AND %ResetTag) <> 0 THEN               ' If %ResetTag
                  IF LTagG(i) <> $BlankLNo THEN                   ' If one exists
                     LTagS(i) = $BlankLNo                         ' Reset the Tag
                     me.UpdLControl(i)                            ' Redo the LLCtl
                     LPCmdS(i) = PCmdIX                           ' Stuff in PCmdIX to make it :ZF
                  END IF                                          '
               END IF                                             '

               '----- Clear Color status
               IF ResetClr <> 0 THEN                              ' If a RESET color
                  IF ResetClr = -1 THEN                           '
                     me.AttrHiLiteSet(i, 1, LAttrGLen(i), %AttrHiStd)' Clear all colors
                  ELSE                                            '
                     me.AttrHiLiteClear(i, 1, LAttrGLen(i), ResetClr) ' Clear the requested color only
                  END IF                                          '
               END IF

               '----- Clear line labels
               IF (Request AND %ResetLabel) <> 0 THEN             ' If %ResetLabel
                  IF L(i).LLbl <> $BlankLNo THEN                  ' If one exists
                     LLblS(i) = $BlankLNo                         ' Reset Line label
                     me.UpdLControl(i)                            ' Redo the LLCtl
                     LPCmdS(i) = PCmdIX                           ' Stuff in PCmdIX to make it :ZF
                  END IF                                          '
               END IF                                             '

               '----- Clear line commands
               IF (Request AND %ResetCommand) <> 0 THEN           ' If %ResetCommand
                  pLCtl = L(i).LLCtl                              ' Save current LLCtl
                  me.UpdLControl(i)                               ' Redo the LLCtl
                  IF L(i).LLCtl <> pLCtl THEN LPCmdS(i) = PCmdIX  ' If we're doing something, make it :ZF
               END IF                                             '

               '----- Clear FIND hiliting
               IF (Request AND %ResetFind) <> 0 THEN              ' If %ResetFind
                  me.AttrInvClear(i, 1, LAttrGLen(i))             ' Clear Inv from all lines
               END IF                                             '

            '----- Misc. line types
            ELSEIF ISTRUE (LFlagG(i) AND (%Tabs OR %Word OR %Mark OR %Mask OR %Prof OR %Bounds OR %Cols)) THEN    ' Delete Misc. Lines
               IF (Request AND %ResetSpecial) <> 0 THEN           ' If %ResetSpecial
                  IF LFlagG(i) AND %Prof = 0 THEN me.LTxtFree(i)  ' Go free the dynamic string
                  me.LEntDel(i)                                   ' Remove from the L() array
                  DECR i                                          ' Make sure we don't miss one
                  DECR LastLine: DECR LastReal                    ' Adjust LastLine and LastReal
                  INCR NumDel                                     ' reduce count of total lines
               END IF                                             '
               IF (Request AND %ResetCommand) <> 0 THEN           ' If %ResetCommand
                  me.UpdLControl(i)                               ' Reset Line number
               END IF                                             '

            END IF                                                '
         NEXT i                                                   '
         ResetClr = 0                                             ' Clear type of Clr reset request

         '----- Adjust for deleted lines
         me.AdjustPending(1, 0 - NumDel, 0)                       ' Adjust pending requests

         '----- Find/Change previous values
         IF (Request AND %ResetAll) = %ResetAll THEN              '
            sCol = 1: sDir = 1: sLine = 1                         ' Restore RFind stuff
            cfFLine = 0: cfFCol = 0: pfFLine = 0: pfFCol = 0      '
            errFlag = %eNone: ErrMsg = ""                         ' Reset flags
         END IF                                                   '
         TopScrn = lclTop                                         ' Go back to where we were
         OnUndoFlag                                               ' Call for an Undo Save
         MExit
      END METHOD

      METHOD  RQSplit(RQuest AS STRING, RQPath AS STRING, RQMask AS STRING, RQFlags AS STRING, RQNote AS STRING)
      '---------- Split out the RQ entries
      REGISTER j AS LONG
      REGISTER k AS LONG
      LOCAL t AS STRING
         MEntry                                                   '
         RESET RQPath, RQMask, RQFlags, RQNote                    ' Start all answers as null
         t = RQuest                                               ' Get RQ string

         j = INSTR(t, "|")                                        ' Look for 1st |
         RQPath = IIF$(j, LEFT$(t, j - 1), t)                     ' Return 1st operand
         IF j = 0 THEN MExitMeth                                  ' No |, just a path

         k = INSTR(j + 1, t, "|")                                 ' Look for 2nd |
         RQMask = IIF$(k, MID$(t, j + 1 TO k - 1), MID$(t, j + 1))'
         IF k = 0 THEN MExitMeth                                  ' No |, just a mask
         j = k                                                    ' Step over |

         k = INSTR(j + 1, t, "|")                                 ' Look for 3rd |
         RQFlags = IIF$(k, MID$(t, j + 1 TO k - 1), MID$(t, j + 1))
         IF k = 0 THEN MExitMeth                                  ' No |, just a flag

         RQNote = MID$(t, k + 1)                                  ' Remainder is the Note
         MExit                                                    '
      END METHOD

      METHOD  SaveReq()
      '---------- Save the request list, it has been modified
      REGISTER i AS LONG
      LOCAL srIO AS iIO                                           ' For our I/O stuff
         MEntry
         LET srIO = CLASS "cIO"                                   '
         DoClear(%SaveReq)                                        ' Say we did it
         IF FileListNm = "" THEN MExitMeth                        ' No FILELIST name, bail out

         srIO.Setup("OR", "", "", ENV.FileListPath + FileListNm + ".FLIST")  ' Tell IO what we're opening
         IF srIO.EXEC THEN _                                      '
            scError(%eFail, "FILELIST Save failed: " + srIO.ResultMsg): MExitMeth  ' Oops?  Bail out

         FOR i = 1 TO RQCount                                     ' Write the file
            PRINT # srIO.FNum, RQList(i)                          ' Print a line
         NEXT i                                                   '
         srIO.Close                                               ' Close the File
         MExit                                                    '
      END METHOD

      METHOD SBGetText(i AS LONG) AS STRING:  METHOD =  TRIM$(SBTable(i).SBText): END METHOD
      METHOD SBGetMySelfP(i AS LONG) AS LONG:   METHOD =  VARPTR(SBTable(i).SBMySelf): END METHOD
      METHOD SBGetDfScheme(i AS LONG) AS LONG:  METHOD = SBTable(i).SBDfScheme: END METHOD
      METHOD SBGetOVScheme(i AS LONG) AS LONG:  METHOD = SBTable(i).SBOvScheme: END METHOD
      METHOD SBGetID(i AS LONG) AS STRING:  METHOD = SBTable(i).SBID: END METHOD
      METHOD SBGetXrWidth(i AS LONG) AS LONG: METHOD = SBTable(SBXref(i)).SBWidth: END METHOD
      METHOD SBGetAlign(i AS LONG) AS STRING:  METHOD = SBTable(i).SBAlign: END METHOD
      METHOD SBGetPosBar(i AS LONG) AS STRING:  METHOD = SBTable(i).SBPosBar: END METHOD
      METHOD SBGetActive(i AS LONG) AS STRING:  METHOD = SBTable(i).SBActive: END METHOD

      METHOD SBSetText(i AS LONG, str AS STRING):  SBTable(i).SBText = str: END METHOD
      METHOD SBSetOvScheme(i AS LONG, clr AS LONG):  SBTable(i).SBOvScheme = clr: END METHOD
      METHOD SBXrefGet(i AS LONG) AS LONG: METHOD = SBXref(i): END METHOD

      METHOD TestForget(FName AS STRING) AS LONG
      '---------- See if filename is in the Forget list
      LOCAL i AS LONG
         IF RFCount = 0 THEN METHOD = 0                           ' Nothing in table, then not found
         ARRAY SCAN RFList() FOR RFCount, COLLATE UCASE, =FName, TO i ' Can we find the filename?
         METHOD = i                                               ' Pass back result
      END METHOD

      METHOD TabsSimple(ln AS LONG) AS STRING
      '---------- Return a TABS line with + expanded at least to ln
      LOCAL i, j, k AS LONG, nTabs AS STRING
         MEntry
         k = MAX(ln, ENV.ScrWidth)                                ' Get min length of output
         i = INSTR(-1, Prf.TabsLine, "*")                         ' Get location of last *
         j = INSTR(-1, Prf.TabsLine, "+")                         ' Get location of last +
         IF j = 0 THEN METHOD = Prf.TabsLine: MExitMeth           ' No +, nothing to do
         nTabs = LEFT$(Prf.TabsLine, i)                           ' Start with LH end of Tabs line
         DO WHILE LEN(nTabs) < k                                  ' Extend until > length
            nTabs += SPACE$(j - i -1) + "*"                       ' Add ...* to create another tab stop
         LOOP                                                     '
         METHOD = nTabs                                           '
         MExit                                                    '
      END METHOD

      METHOD TabsSave()
      '---------- Save the TABS line and CAPS setting by updating the INI file
      LOCAL Wrd1 AS STRING, i, j AS LONG
         MEntry
         OffTabsAFlag                                             ' Clear flag that got us here
         Wrd1 = TP.TabsLine                                       ' Get a working copy of Tabs line
         i = INSTR(-1, Wrd1, "+")                                 ' Get location of last + sign
         j = INSTR(-1, Wrd1, "*")                                 ' Get location of last * sign
         IF i > 0 AND j > 0 AND j > i THEN                        '
            Prf.TabsLine = ""                                     '
            MErrExit(%eFail, "Repeating tab (+) is left of last fixed tab (*), TABS line nulled.") ' Else tell user
         END IF                                                   '
         Prf.TabsLine = TP.TabsLine                               ' Set it into the Profile
         MExit                                                    '
      END METHOD

      METHOD TabTitleSet(force AS LONG)
      '---------- Setup the tab title
      LOCAL title AS STRING, lcltcitem AS TC_ITEMHEADER, newtxt AS ASCIIZ * 255, tWnd AS LONG
         MEntry
         IF gMacroMode THEN MExitMeth                             ' If macro mode, exit
         IF IsFMTab THEN                                          ' The simple case?
            title = "File Manager"                                '
            GOSUB DoSwitch                                        ' Switch it
            MExitMeth                                             ' We're done
         END IF                                                   '
         IF ((IsModdFlag AND IsModdLast) OR (ISFALSE IsModdFlag AND ISFALSE IsModdLast)) AND _
            force = %False THEN MExitMeth                         ' Don't do if unchanged, causes flickering
         IF IsModdFlag THEN OnModdLast ELSE OffModdLast           ' Save as last status
         IF IsClip THEN                                           '
            title += "(CLIP)"                                     '
            IF ClipName <> "" THEN title += " - " + ClipName      '
         ELSEIF IsSetEdit THEN                                    '
            title += "(SET Edit)"                                 '
         ELSEIF IsMEdit THEN                                      '
            title += "(M-Edit)"                                   '
         ELSEIF ISNULL(TP.TIPFilePath) THEN                       '
            title += $Empty                                       ' So we don't display nulls
         ELSE                                                     '
            title += TP.TIPFile                                   ' Start with the file name
         END IF                                                   '
         IF IsModdFlag THEN _                                     ' If modified, Add 128 to first char as a signal
            title = CHR$(ASC(LEFT$(title, 1)) + 128) + MID$(title, 2)'
         GOSUB DoSwitch                                           '
         MExitMeth                                                '

      DoSwitch:
         CONTROL HANDLE hWnd, %IDC_SPFLiteTAB TO tWnd             ' Get the handle
         newtxt = title                                           ' Copy title to ASCIIZ string
         lcltcitem.mask = %TCIF_TEXT                              ' Say which item we're changing
         lcltcitem.pszText = VARPTR(NewTxt)                       ' Point at the new text
         SendMessage(tWnd, %TCM_SETITEM, TP.PgNumber - 1, VARPTR(lcltcitem)) ' Make the change
         RETURN                                                   '
      END METHOD

      METHOD TagChange(NewTag AS STRING) AS LONG
      '---------- Do the TAG status change, return 1 if a change occurred
      LOCAL i, result AS LONG, t AS STRING
         MEntry
         i = cfFLine: t = UUCASE(NewTag)                          ' Line we're working on
         IF IsCfOn OR IsCfSet THEN                                ' TAG :xxx ON/SET?
            IF LTagG(i) <> t THEN                                 ' Already set to this tag?
               LTagS(i) = t                                       ' No? Copy in the new Tag
               me.UpdLControl(i)                                  ' Reset line control
               result = 1                                         ' Pass back 1 to say we did something
            END IF                                                '
         ELSEIF IsCfOff OR IsCfAssert THEN                        ' TAG :xxx OFF or ASSERT?
            IF ISNULL(t) THEN                                     ' If no tag specified
               IF LTagG(i) <> $BlankLNo THEN                      ' If not already a blank tag?
                  LTagS(i) = $BlankLNo                            ' Clear the Tag
                  me.UpdLControl(i)                               ' Reset line control
                  result = 1                                      ' Pass back 1 to say we did something
               END IF                                             '
            ELSE                                                  ' Tag specified, OFF only if we match
               IF t = LTagG(i) THEN                               ' So, do we match?
                  LTagS(i) = $BlankLNo                            ' Clear the Tag
                  me.UpdLControl(i)                               ' Reset line control
                  result = 1                                      ' Pass back 1 to say we did something
               END IF                                             '
            END IF                                                '
         ELSEIF IsCfToggle THEN                                   ' TAG :xxx TOGGLE?
            IF ISNULL(TRIM$(LTagG(i))) THEN                       ' If currently null
               LTagS(i) = t                                       ' Copy in the new Tag
               me.UpdLControl(i)                                  ' Reset line control
               result = 1                                         ' Pass back 1 to say we did something
            ELSE                                                  ' Already a tagname
               IF t = LTagG(i) THEN                               ' If we match
                  LTagS(i) = $BlankLNo                            ' Clear the Tag
                  me.UpdLControl(i)                               ' Reset line control
                  result = 1                                      ' Pass back 1 to say we did something
               END IF                                             '
            END IF                                                '
         END IF                                                   '
         IF IsCfMX THEN                                           ' Was this an MX operation?
            IF result = 0 THEN me.LFlagBitOff(i, %Invisible)      ' Then UnDo the flip the search made
         END IF                                                   '
         METHOD = result                                          ' Pass back 0/1 for counting
         MExit
      END METHOD

      METHOD TestExclude(FName AS STRING) AS LONG
      '---------- See if filename is in the Exclude list
      LOCAL i AS LONG
         IF RXCount = 0 THEN METHOD = 0                           ' Nothing in table, then not found
         ARRAY SCAN RXList() FOR RXCount, COLLATE UCASE, =FName, TO i ' Can we find the filename?
         METHOD = i                                               ' Pass back result
      END METHOD

      METHOD TestMask(FName AS STRING, FTMask AS STRING) AS LONG
      '---------- See if filename matches a mask
      LOCAL FType, Fn, tmask AS STRING, TName, TMask2 AS ASCIIZ * %MAX_PATH, MPtr, FPtr, TPtr, LPtr AS BYTE POINTER
      LOCAL i, j, chosen, request, DblA AS LONG
      LOCAL Masks() AS STRING
         MEntry
         tmask = IIF$(LEFT$(FTMask, 1) = "-", "*;" + FTMask, FTMask) ' Add a leading *; if mask starts off negative
         REDIM Masks(1 TO PARSECOUNT(tmask, ";")) AS STRING       ' Size the array
         PARSE UUCASE(TRIM$(tmask)), Masks(), ";"                 ' Split into the table
         TName = UUCASE(FName)                                    ' Make filename ASCIIZ

         FOR j = 1 TO UBOUND(Masks())                             ' Now for each mask
            request = %True                                       ' Start as True
            tmask = masks(j)                                      ' Copy it
            IF LEFT$(tmask, 1) = "-" THEN                         ' Unwanted?
               request = %False: tmask = MID$(tmask, 2)           ' Make Request false, drop the -
            ELSEIF LEFT$(tmask, 1) = "+" THEN                     ' Wanted?
               tmask = MID$(tmask, 2)                             ' Request true, drop the +
            END IF                                                '

            IF LEFT$(tmask, 1) = "=" THEN                         ' Set symbol?
               tmask = sSetTable("GET", MID$(tmask, 2))           ' Retrieve substitution value
               IF VAL(LEFT$(tmask, 1)) > 0 THEN                   ' Not found?
                  tmask = masks(j)                                ' Put back untouched operand
               ELSE                                               ' We have a value
                  tmask = UUCASE(MID$(tmask, 2))                  ' Substitute it
               END IF                                             '
            END IF                                                '
            tmask2 = tmask                                        '

            IF PathMatchSpec(TName, TMask2) THEN                  ' Do we have a match?
               Chosen = Request: ITERATE FOR                      ' Set match
            ELSE                                                  '
               ITERATE FOR                                        '
            END IF                                                '
         NEXT j                                                   '

         IF chosen THEN METHOD = %True ELSE METHOD = %False       ' Result is final Chosen value
         MExit                                                    '
      END METHOD

      METHOD UnWatchQueue(fn AS STRING)
      '---------- Kill watch and dequeue the file
         MEntry
         IF ISNOTNULL(TRIM$(fn)) THEN                             ' If a filename
            me.FileWatch(fn, %WatchEnd)                           ' Kill specific filename Watch
            sFileQueue("D", " ", fn)                              ' Remove from the Open queue
         ELSE                                                     '
            me.FileWatch("", %WatchEnd)                           ' Kill any prior Watch for the tab
            sFileQueue("D", " ", TIP.FilePath)                    ' Remove default file from the Open queue
         END IF                                                   '
         Mexit                                                    '
      END METHOD

      METHOD UpdLControl(lno AS LONG)
      '---------- Update the Line Ctl field
      REGISTER i AS LONG
      REGISTER j AS LONG
      LOCAL f, t AS LONG, KCmd AS STRING
         MEntry
         i = lno                                                  ' Swap line number to local
         f = (LFlagG(i) AND %SpecialLine)                         ' Get a copy of the line flag, keep just the Type bits
         IF f THEN                                                ' It's a special one
            FOR j = 1 TO UBOUND(gLnoTextType)                     ' Lets search
               IF (f AND gLnoTextType(j)) THEN                    ' Match of type?
                  LLCtlS(i) = gLnoTextTxt(j)                      ' Set the text
                  EXIT FOR                                        ' We're done
               END IF                                             '
            NEXT j                                                '

            '----- Fiddle =FILE line if needed
            IF (f AND %File) THEN                                 ' Is this a =FILE> line?
               IF me.MeditFlagGet(LMIXG(i)) THEN                  ' Modified
                  KCmd = me.LLCtlGet(i)                           ' Get the current LLCtl
                  MID$(KCmd, CHOOSE(ENV.LinNoSize, 1, 1, 1, 1, 5, 6, 7, 8 ELSE 1), 1) = "*" ' Add modified *
                  LLCtlS(i) = KCmd                                ' Stuff it back
               END IF                                             '
            END IF                                                '

            '----- Fiddle =NOTE line if needed
            IF (f AND %Note) THEN                                 ' Is this a =NOTE> line?
               j = L(i).LWrk2                                     ' Get the LWrk2 value
               IF j THEN                                          ' Other than default note?
                  KCmd = me.LLCtlGet(i)                           ' Get the current LLCtl
                  MID$(KCmd, CHOOSE(ENV.LinNoSize, 1, 1, 1, 1, 2, 1, 2, 3 ELSE 1), 1) = MID$($Upper, j, 1) ' Add type of xNOTE
                  LLCtlS(i) = KCmd                                ' Stuff it back
               END IF                                             '
            END IF                                                '

         ELSE                                                     ' Nothing special, just a line number
            LLCtlS(i) = LLNumG(i)                                 ' Set the line number in
         END IF                                                   '

         '----- Add Labels and tags
         me.LLctlTagSet(i)                                        ' Overlay with Tag
         me.LLctlLblSet(i)                                        ' Overlay with LLbl

         IF TCtr > 0 THEN                                         ' Pending line commands?
            KCmd = me.TTblScan(i)                                 ' See if a pending line command
            IF ISNOTNULL(TRIM$(KCmd)) THEN LLCtlS(i) = UUCASE(KCmd) ' Swap it in
         END IF                                                   '
         MExit                                                    '
      END METHOD

      METHOD ValScroll()
      '---------- Validate Scroll Amount
      LOCAL Scr, t AS STRING, i AS LONG
         MEntry
         IF IsEQ(Prf.Scroll, ScrlAmtLast) THEN MexitMeth          ' Nothing changed?
         scr = UUCASE(Prf.Scroll)                                 ' Get local copy

         '----- Use default if field nulled
         IF ISNULL(TRIM$(scr)) THEN                               ' Null'd?
            Prf.Scroll = ScrlAmtLast                              ' set in default
            IF IsFMTab THEN sIniSetString("FManager", "ScrlAmtc", Prf.Scroll) ' Save as the FM default if an FM Tab
            MExitMeth                                             '
         END IF                                                   '

         '----- If numeric, handle it first
         scr = LSET$(scr,5)                                       ' Make it 5 chars long
         IF VERIFY(LEFT$(scr, 1), "0123456789") = 0 THEN          ' If leading numeric
            i = VERIFY(scr, "0123456789")                         ' Find first non-numeric
            scr = LEFT$(scr, i - 1)                               ' Throw away non-numerics
            IF VAL(scr) > 0 THEN                                  ' Must not be zero or neg
               Prf.Scroll = FORMAT$(VAL(scr), "####")             ' Stuff back pretty version
               IF IsFMTab THEN sIniSetString("FManager", "ScrlAmtc", Prf.Scroll) ' Save as the FM default if an FM Tab
               MexitMeth                                          ' Done
            ELSE                                                  '
               Prf.Scroll = ScrlAmtLast                           ' set in default
               MErrExit(%eFail, "Invalid scroll Amt. Re-establishing the previous value")
            END IF                                                '
         END IF                                                   '

         '----- See which alpha type it is
         scr = TRIM$(scr)                                         ' Trim it
         SELECT CASE CONST$ LEFT$(scr, 1)                         ' Which was it?
            CASE "P": Prf.Scroll = "PAGE"                         '
            CASE "F": Prf.Scroll = "FULL"                         '
            CASE "H": Prf.Scroll = "HALF"                         '
            CASE "D": Prf.Scroll = "DATA"                         '
            CASE "C": Prf.Scroll = "CSR "                         '
            CASE ELSE                                             '
               Prf.Scroll = UUCASE(ScrlAmtLast)                   ' Just in case
               MErrExit(%eFail, "Invalid scroll Amt. Re-establishing the previous value")
         END SELECT                                               '

         IF IsFMTab THEN sIniSetString("FManager", "ScrlAmtc", Prf.Scroll) ' Save as the FM default if an FM Tab
         MExit                                                    '

      END METHOD

      METHOD WINDOW(BYREF Tx AS STRING) AS STRING
      '---------- Return a portion of a Text string to match window offset
         METHOD = LSET$(MID$(TX, Offset + 1, gDataLen), gDataLen) ' Pass back string (Padded if needed)
      END METHOD

      METHOD WindowBnds() AS STRING
      '---------- Return a BNDS line to match current Window offset
      LOCAL t AS STRING
         t = MID$(BndText, Offset + 1, gDataLen)                  ' Pickup gDataLen characters from the Offset location
         IF LEN(t) < gDataLen THEN _                              ' Make sure at least screen width
            t = LSET$(t, gDataLen)                                ' Lengthen to gDataLen if needed
         METHOD = t                                               ' Pass back the result
      END METHOD

      METHOD WindowCmd()
      '---------- Reprint the command line
      LOCAL lCommand1, lCommand2, Wd1, Wd2, KWs, t1 AS STRING, i AS LONG, AttrLine, t2 AS WSTRING
         MEntry
         cBandBG = %False                                         ' Clear any prior banding state
         AttrLine = REPEAT$(MAX(1, LEN(pCommand)), CHR$$(%SCTxtLo)) ' Make whole line Lo intensity
         Wd1 = UUCASE(sGetWord(pCommand, %NoStrip, %QuoteNotSig)) ' Get the basic command name
         i = PCmdT.GetCmdIX(Wd1)                                  ' Get the command table index?
         IF i = 0 THEN                                            ' If we can't identify a command
            sPrint (LSET$(MID$(pCommand, COffset+1, pCmdLen), pCmdLen + 1), AttrLine, 1, 11)
            GOTO DoMark                                           '
         END IF                                                   '

         '----- OK we have a recognized command that has KW type operands
         KWs = "," + PCmdT.GetOprands(i) + ","                    ' Copy valid operands
         lCommand1 = UUCASE(pCommand) + " "                       ' Working copy of pCommand
         lCommand2 = lCommand1                                    ' Another copy
         Wd1 = sGetWord(lCommand1, %Strip, %QuoteSig)             ' Strip off the command name
         i = INSTR(lCommand2, Wd1)                                ' Locate it in line in case not at the left
         MID$(AttrLine, i, LEN(Wd1)) = REPEAT$(LEN(Wd1), CHR$$(%ScTxtHi))

         IF ISNULL(PCmdT.GetOprands(i)) THEN GOTO DoMark          ' Also if a no operand command, we're done

         DO WHILE TRIM$(lCommand1) <> ""                          ' While we have operands
            Wd1 = sGetWord(lCommand1, %Strip, %QuoteSig)          ' Get next operand, add commas
            Wd2 = "," + Wd1 + ","                                 ' Add commas
            i = INSTR(KWs, Wd2)                                   ' See if we've found a KW
            IF i THEN                                             ' We've got one
               i = INSTR(lCommand2, " " + Wd1 + " ")              ' Locate it in line
               MID$(AttrLine, i + 1, LEN(Wd1)) = REPEAT$(LEN(Wd1), CHR$$(%ScTxtHi))' Mark it Hi
            END IF                                                '
         LOOP

         DoMark:
         sPrint (LSET$(MID$(pCommand, COffset+1, pCmdLen), pCmdLen + 1), AttrLine, 1, 11)

         '----- Mark selected Cmd
         IF IsMiscActive THEN                                     ' Do we have a marked Cmd?
            OffMiscDrawn                                          ' Previous is gone, we just re-drew the screen
            me.MiscMark                                           ' Go re-draw the marked area
         END IF                                                   '
         MExit                                                    '
      END METHOD

      METHOD WindowCol() AS STRING
      '---------- Return a COLS line to match current Window offset
      LOCAL t, u, lcltabs AS STRING, i, j AS LONG
         lcltabs = me.TabsSimple(Offset + gDataLen)               ' Get a working TABS line
         t = STRING$(gDataLen + 5, "-")                           ' Fill line with dashes (plus a cushion)

         FOR i = 1 TO gDataLen                                    ' Loop now for each character
            j = i + Offset                                        ' j = real data column
            IF (j MOD 10) = 0 THEN                                ' At a column interval=10 ?
               u = FORMAT$(j/10, "#####")                         ' Format the column number
               MID$(t,i,LEN(u)) = u                               ' Overlay the number
            ELSEIF (j MOD 5) = 0 THEN                             ' At a column interval=5
               MID$(t,i,1) = "+"                                  ' Stuff in a +
            END IF                                                '
            IF MID$(lcltabs, j, 1) = "*" THEN                     ' Also a tab position
               MID$(t, i, 1) = CHR$(ASC(MID$(t, i, 1)) + 128)     ' Flag for reverse video
            END IF                                                '
         NEXT i                                                   '

         METHOD = LEFT$(t, gDataLen)                              ' Pass back the result
      END METHOD

      METHOD WindowHexBottom(BYREF Txt1 AS STRING) AS STRING
      '---------- Return a portion of a Text string to match window offset
      LOCAL t, c, HexTable2 AS STRING
      REGISTER i AS LONG
         HexTable2 = REPEAT$(16, $Hex)                            '
         t = LSET$(MID$(Txt1, Offset+1, gDataLen), gDataLen)      ' Get windowed substring

         FOR i = 1 TO LEN(t)                                      ' Loop through characters
            IF Prf.SrceXlate OR Prf.ColateXlate THEN              ' Non-ANSI SOURCE or COLLATE?
               c = MID$(t, i, 1)                                  ' Do the translate
               me.Translate(c, Prf.GetCA2SPtr)                    '
               MID$(t, i, 1) = MID$(HexTable2, 1 + ASC(c), 1)     '
            ELSE                                                  '
               MID$(t, i, 1) = MID$(HexTable2, 1 + ASC(MID$(t, i, 1)), 1) '
            END IF                                                '
         NEXT i                                                   '

         METHOD = t                                               ' Pass back result
      END METHOD

      METHOD WindowHexTop(BYREF Txt1 AS STRING) AS STRING
      '---------- Return a portion of a Text string as Upper Hex
      LOCAL t, c, HexTable1 AS STRING
      REGISTER i AS LONG
         HexTable1 = BUILD$("0000000000000000111111111111111122222222222222223333333333333333", _
                            "4444444444444444555555555555555566666666666666667777777777777777", _
                            "88888888888888889999999999999999AAAAAAAAAAAAAAAABBBBBBBBBBBBBBBB", _
                            "CCCCCCCCCCCCCCCCDDDDDDDDDDDDDDDDEEEEEEEEEEEEEEEEFFFFFFFFFFFFFFFF")
         t = LSET$(MID$(Txt1, Offset+1, gDataLen), gDataLen)      ' Get windowed substring

         FOR i = 1 TO LEN(t)                                      ' Loop through characters
            IF Prf.SrceXlate OR Prf.ColateXlate THEN              ' Non-ANSI SOURCE or COLLATE?
               c = MID$(t, i, 1)                                  ' Do the translate
               me.Translate(c, Prf.GetCA2SPtr)                    '
               MID$(t, i, 1) = MID$(HexTable1, 1 + ASC(c), 1)     '
            ELSE                                                  '
               MID$(t, i, 1) = MID$(HexTable1, 1 + ASC(MID$(t, i, 1)), 1) '
            END IF                                                '
         NEXT i                                                   '

         METHOD = t                                               ' Pass back result
      END METHOD

      METHOD WindowLCmd(idx AS LONG) AS STRING
      '---------- Extract the visible FM Line command portion
         IF TRIM$(AFList(idx).Cmd) <> "" THEN                     ' Something in command?
            METHOD = LSET$(MID$(AFList(idx).Cmd, AFList(idx).CmdOff + 1), ENV.FMLCmdWidth) + " "
         ELSE                                                     '
            METHOD = REPEAT$(ENV.FMLCmdWidth, IIF$((AFList(idx).FD.FileAttributes AND %FILE_ATTRIBUTE_READONLY) = %FILE_ATTRIBUTE_READONLY, ".","_")) + " "
         END IF                                                   '
      END METHOD

      METHOD WindowNote(idx AS LONG) AS STRING
      '---------- Extract the visible FM Note portion
         IF TRIM$(AFList(idx).Note) <> "" THEN                    ' Something in command?
            METHOD = LSET$(MID$(AFList(idx).Note, AFList(idx).NoteOff + 1), FM_Note_Size)
         ELSE                                                     '
            METHOD = REPEAT$(FM_Note_Size, "_")                   '
         END IF                                                   '
      END METHOD

      METHOD WindowTitle()
      '---------- Return the formatted Window Title Bar text
      LOCAL TX AS ASCIIZ * 200, lText, lVer AS STRING
         IF ISFALSE gNoVersion THEN _                             ' Build version string
            lVer = "(v" + ENV.PgmVers + ")"                       ' if needed
         IF IsClip THEN                                           ' CLIP mode?
            TX = "CLIP - SPFLite" + lVer                          '

         ELSEIF IsSetEdit THEN                                    ' SETEdit mode?
            TX = "SET-Edit - SPFLite" + lVer                      '

         ELSEIF IsMEdit THEN                                      ' Multi-Edit mode?
            TX = "Multi-Edit - SPFLite" + lVer                    '

         ELSEIF IsFMTab THEN                                      ' File Manager?
            IF ISNOTNULL(FileListNm) THEN                         ' Got a filelist name?
               TX = FileListNm + " - FILELIST - File Manager - SPFLite" + lVer
            ELSE                                                  '
               TX = "File Manager - SPFLite" + lVer               '
            END IF                                                '

         ELSEIF TIP.FilePath = $Empty THEN                        ' FileTitle chosen?
            TX = $Empty + " - SPFLite" + lVer                     '

         ELSE                                                     '
            lText = SWITCH$(IsBrowse, "BROWSE - ", IsView, "VIEW - ", %True, "EDIT - ")
            TX = lText + TP.TIPFile + " - SPFLite" + lVer + " - " + TP.TIPFilePath
         END IF                                                   '
         SetWindowText(hWnd, TX)                                  ' Alter window title
         me.TabTitleSet(%True)                                    ' If setting Window Title, always do Tab Title
      END METHOD

      METHOD WordSave()
      '---------- Save the WORD line by updating the INI file
      LOCAL t AS STRING
         MEntry
         OffWordAFlag                                             ' Clear flag that got us here
         t = TP.WordInput                                         ' Get what was entered
         IF ISNULL(TRIM$(t)) THEN t = $WORD                       ' Reset if field nulled.
         Prf.WordInput = t                                        '
         IF prf.WordVal() THEN Prf.WordReset                      ' If invalid request, reset it
         me.PicSetWord                                            ' Say WORD has changed
         me.PicInit                                               ' Re-Initialize Picture control area
         MExit
      END METHOD

      METHOD WriteFile(Quiet AS LONG) AS INTEGER
      LOCAL x, y, i, j, recs AS LONG, t, tt AS STRING, C1, C2 AS BYTE POINTER
      LOCAL pt, DLM, RDW  AS STRING, T1W AS WSTRING
         MEntry

         '----- If RECFM=F see if any lines > LRECL exist, tell user if so
         IF Prf.RECFM = "F" THEN                                  ' Any need to warn user at all?
            MaxLength = 0                                         ' Reset biggest size record
            FOR i = 2 TO LastLine - 1                             ' Ugh! must do a scan
               IF IsLData(i) THEN _                               ' Look at data lines only
                  MaxLength = MAX(MaxLength, LEN(L(i).@LTxt))     ' Track MaxLength
            NEXT i                                                '
            IF MaxLength > Prf.LRECL AND ISFALSE Quiet THEN       ' Lines longer than LRECL?
               j =sDoMsgBox("File contains lines longer then the LRECL of " + FORMAT$(Prf.LRECL) + "," + $CRLF + _
                         "Click |KYES|B to continue and truncate lines, or" + $CRLF + _
                         "Click |KNO|B  to abandon the file write and return to Edit", %MB_YESNO + %MB_USERICON, "SPFLite File Write")
               IF j = %IDYES THEN                                 ' OK?
                  '                                               '
               ELSE                                               '
                  METHOD = %False                                 ' Set return code
                  MExitMeth                                       '
               END IF                                             '
            END IF                                                '
         END IF                                                   '

         '----- Continue with writing the file
         METHOD = %False                                          ' Set default error exit
         IF ISNULL(TOP.FilePath) THEN MExitMeth                   ' Better be here, else exit

         '----- Choose the correct EOL delimiter
         dlm = Prf.EOL                                            '
         IF Prf.EOL = "CRLF" THEN                                 ' Setup the EOL delimiter
            DLM = $CRLF                                           '
         ELSEIF Prf.EOL = "LF" THEN                               '
            DLM = $LF                                             '
         ELSEIF Prf.EOL = "CR" THEN                               '
            DLM = $CR                                             '
         ELSEIF Prf.EOL = "NL" THEN                               '
            DLM = $NL                                             '
         ELSEIF Prf.EOL = "NONE" THEN                             '
            DLM = ""                                              '
         ELSEIF LEFT$(Prf.EOL, 4) = "AUTO" THEN                   '
            DLM = $CRLF                                           '
         ELSE                                                     ' Must be hex
            DLM = sHex2Str(TRIM$(Prf.EOL))                        ' Set it up
         END IF                                                   '

         '---------- OK, now write the file                       '
         IF LastLine * MaxLength > 4000000 THEN gLoopCtr = - 1    ' Prevent loop detection for big files

         IF TOP.EXEC THEN _                                       ' Go do OPEN, True if error
            MErrExit(%eFail, TOP.ResultMsg)                       ' Tell user of error
         LOCK # TOP.FNum                                          '
         IF ERR <> 0 AND ISFALSE Quiet THEN                       ' OPEN OK?
            scError(%eFail,"Failed to LOCK output file, ERR=" + STR$(ERRCLEAR))
            TOP.Close                                             '
            MExitMeth                                             '
         END IF                                                   '
T1W = LAttrG(2)

         '----- Write UTF header if appropriate
         IF Prf.PSource = "UTF8" THEN                             ' UTF8
            PRINT # TOP.FNum, CHR$(&hEF, &hBB, &hBF);             ' ....... write EFBBBF leader
         ELSEIF Prf.PSource = "UTF16" THEN                        ' UTF16
            PRINT # TOP.FNum, CHR$(&hFF, &hFE);                   ' ....... write FFFE   leader
         ELSEIF Prf.PSource = "UTF16BE" THEN                      ' UTF16BE
            PRINT # TOP.FNum, CHR$(&hFE, &hFF);                   ' ....... write FEFF   leader
         END IF                                                   '

         '----- Loop through the lines writing them
         x = me.Search(0, INT(Zone))                              ' Do the initial search
         IF X THEN                                                '
            x = cfFLine                                           ' Get found line
            GOSUB WriteIt                                         ' Go Write it
         END IF                                                   '
         DO WHILE X                                               ' Do till not found
T1W = LAttrG(2)
            X = me.Search(1, %CursBad)                            ' Do re-Finds
T1W = LAttrG(2)
            IF X THEN                                             ' Found one
               X = cfFLine                                        ' Get found line
               GOSUB WriteIt                                      ' Go Write it
            END IF                                                '
         LOOP                                                     '
         UNLOCK # TOP.FNum                                        ' Unlock it
         TOP.Close                                                ' All done, CLOSE it
T1W = LAttrG(2)

         '----- Issue messages and do the STATE handliong
         IF ISFALSE IsMEdit AND ISFALSE Quiet THEN scError(%eNone, FORMAT$(recs) + " records written to file") ' Issue message
         FileRecs = recs                                          ' Save if caller wants to know
         me.SetNewStart()                                         ' Set .START if needed
         me.StateSave(TOP.FilePath, DotStart, DotEnd, %True)      ' Save the STATE data, X/NX and U/NU filtering
T1W = LAttrG(2)
         METHOD = %True: MExitMeth                                '

      '----- Do output for a single line
      WriteIt:
         RESET gLoopCtr                                           ' Reset LoopCtr to avoid treating as a loop
         IF IsLData(x) AND ISFALSE IsLInsertLine(x) THEN          ' Only completed data lines of course

            '----- If AUTOCAPS, get the proper case'd version
            IF Prf.AutoCaps THEN                                  ' Do we need to do this?
               t = me.DoAutoCaps(x)                               ' Go do it
            ELSE                                                  '
               t = LTxtG(x)                                       ' Else get the normal text
            END IF                                                '

            '----- Do the PRESERVE trailing blanks stuff
            IF Prf.PPreserve = 0 THEN                             ' PRESERVE OFF?
               t = RTRIM$(t)                                      '
            ELSE                                                  '
               IF Prf.PPreserve = 1 THEN                          ' PRESERVE ON?
                  ' Do nothing                                    '
               ELSE                                               ' Must be PRESERVE C
                  i = INSTR(-1, t, "\")                           ' Get location of last \ in the line
                  IF ISNOTNULL(TRIM$(MID$(t, i + 1))) THEN i = 0 ' Remainder better be blank, else pretend we didn't see \
                  IF i = 0 THEN                                   ' No \ character
                     ' Do nothing                                 ' Then just like PRESERVE ON
                  ELSE                                            ' We have a \
                     IF i = LEN(t) THEN                           ' Right at the end?
                        ' Do nothing                              ' Then just like PRESERVE ON
                     ELSE                                         ' Else we truncate it at the \
                        t = LEFT$(t, i)                           '
                        me.LTxtSet(x, t)                          ' Replace the actual text too
                     END IF                                       '
                  END IF                                          '
               END IF                                             '
            END IF                                                '

            '----- If AUTOxx file, preserve FF page breaks
            IF LEFT$(Prf.EOL, 4) = "AUTO" THEN                    '
               IF IsLPage(x) AND recs <> 0 THEN _                 ' =PAGE> line? (and not the first)
                  t = $FF + t                                     ' Re-Insert the FF page break
            END IF                                                '

            '----- Handle FIXED, Non-ANSI and add the EOL delimiter
            IF Prf.LRECL > 0 THEN                                 ' If Fixed records, make it so
               t = LSET$(t, Prf.LRECL)                            '
            END IF                                                '

            IF Prf.SrceXlate THEN                                 ' If Non-ANSI do before EOL
               me.Translate(t, Prf.GetSA2SPtr)                    ' Do the translate
            END IF                                                '

            t += DLM                                              ' Add the appropriate delimiter

            '----- Do UTF translation if required
            IF Prf.PSource = "UTF8" THEN                          ' See what encoding to use
               t = sUtf8FromAnsi(t)                               ' UTF8
            ELSEIF Prf.PSource = "UTF16" THEN                     '
               t = UCODE$(t)                                      ' UTF16
            ELSEIF Prf.PSource = "UTF16BE" THEN                   '
               t = UCODE$(t)                                      ' UTF16BE
               FOR i = 0 TO LEN(t) - 1 STEP 2                     ' We have to reverse Char. Pairs now
                  C1 = STRPTR(t) + i: C2 = C1 + 1                 ' Point at 1st two characters
                  SWAP @C1, @C2                                   ' to get BE into LE format
               NEXT                                               '
            END IF                                                '

            '----- Handle RECFM=V if needed
            IF LEFT$(Prf.RECFM, 1) <> "V" THEN                    ' If not RECFM=V
               PRINT # TOP.FNum, t;                               ' Finally, actually write it
            ELSE                                                  ' RECFM=V, add an RDW
               SELECT CASE AS CONST$ Prf.RECFM                    ' See how to handle it
                  CASE "V"                                        ' Old style V
                     i = LEN(t) + 4                               ' Calc length including RDW
                     RDW = CHR$(PEEK(VARPTR(i) + 1), PEEK(VARPTR(i)), 0, 0) ' Build the RDW
                  CASE "VBI"                                      ' Big Endian
                     i = LEN(t)                                   ' Calc length not including RDW
                     RDW = CHR$(PEEK(VARPTR(i)), PEEK(VARPTR(i) + 1), PEEK(VARPTR(i) + 2), PEEK(VARPTR(i) + 3)) ' Build the RDW
                  CASE "VLI"                                      ' Little EndiAN
                     i = LEN(t)                                   ' Calc length not including RDW
                     RDW = CHR$(PEEK(VARPTR(i) + 3), PEEK(VARPTR(i) + 2), PEEK(VARPTR(i) + 1), PEEK(VARPTR(i))) ' Build the RDW
               END SELECT                                         '
               PRINT # TOP.FNum, RDW;                             ' Write RDW
               PRINT # TOP.FNum, t;                               ' Finally, write the data
            END IF                                                '

            '----- Count and do the +/- post line processing
            INCR recs                                             ' Count it
            IF DotSet AND DotFlag <> 0 THEN                       ' Possible +/- processing?
               y = DotFlag                                        ' Fetch the flag locally
               IF BIT(y, %lCmdX) THEN                             ' Do the +/- processing for the Source
                  me.LFlagBitOn(x, %Invisible)                    ' If - make Invisible
                  gfXRebuild = %True                              ' Ask for exclude processing
               ELSEIF BIT(y, %lCmdNX) THEN                        '
                  me.LFlagBitOff(x, %Invisible)                   ' If + make visible
                  gfXRebuild = %True                              ' Ask for exclude processing
               END IF                                             '
            END IF                                                '
         END IF                                                   '
      RETURN
      END METHOD

      METHOD XRebuild(BYREF csr AS LONG, BYREF top AS LONG, BYREF scrl AS LONG)
      '---------- Rebuild the Excluded lines
      REGISTER i AS LONG
      REGISTER j AS LONG
      LOCAL k, l, fMIX AS LONG, FLineNumber AS STRING
      LOCAL InitLastLine AS LONG
         MEntry
         gfXRebuild = %False                                      ' Turn of the flag that got us here
         MaxLength = 0: FLineNumber = "00000000"                  ' Reset MaxLength and FLine
         InitLastLine = LastLine                                  ' Save LastLine when we start
         IF gMacroMode AND cfFLine > 0 THEN                       ' Is an FLine set?
            FLineNumber = IIF$(ISTRUE (L(cfFLine).LFlag AND %Data), L(cfFLine).LLNum, "00000000") ' Save cfFLine in line number format
         END IF                                                   '
         i = 1                                                    ' Now rebuild them
         DO WHILE ISFALSE IsLBottom(i)                            ' Do till we reach the end

            IF IsLXclude(i) THEN                                  ' An Exclude marker line?

               '----- See if this Excluded marker is even needed
               IF ISFALSE IsLInvisible(i + 1) THEN                ' If next line is not Invisible, the marker is un-needed
                  me.LTxtFree(i)                                  ' Go free the dynamic string
                  me.LEntDel(i)                                   ' Remove from the L() array
                  DECR LastLine                                   ' Adjust LastLine
                  ITERATE DO                                      ' We're done here
               END IF                                             '

               '----- Re-Verify this Excluded block
               GOSUB TestSpec                                     ' Look for our special cursor marking lines
               LWrk1S(i) = 0                                      ' Clear count
               j = i + 1                                          ' Set j to the 1st Invisible line
               DO WHILE IsLInvisible(j) OR IsLXclude(j)           ' While still a possible excluded block candidate
                  IF IsLXclude(j) THEN                            ' A 'butted up' 2nd Excluded range?
                     me.LTxtFree(j)                               ' Delete the un-needed marker line
                     me.LEntDel(j)                                ' Remove from the L() array
                     DECR LastLine                                ' Adjust LastLine
                     ITERATE DO                                   ' We're done here
                  END IF                                          '

                  '----- Must now be a normal Invisible line
                  LWrk1S(i) += 1                                  ' Bump count in marker line
                  LWrk1S(j) = j - i                               ' Stuff how far back the marker line is
                  GOSUB TestSpec                                  ' Look for our special cursor marking lines

                  '----- Fudge the XPrt stuff
                  IF IsLXPtr(j) THEN                              ' Is this a Backward pointer?
                     me.LFlagBitOff(j, %XPtr)                     ' Clear it
                     L(i).LCol = L(j).LCol                        ' Copy data backward
                     me.LFlagBitOn(i, %Cursor)                    ' Set this as Cursor line
                     csr = i                                      ' Pass back this line
                     me.LFlagBitOff(j, %Cursor)                   ' Kill Cursor on the invisible line
                     IF IsLScroll(j) THEN                         ' Allowed to scroll?
                        me.LFlagBitOn(i, %Scroll)                 ' Transfer the Scroll flag
                        scrl = i                                  ' Pass back since we're past it
                     END IF                                       '
                     me.LFlagBitOff(j, %Scroll)                   ' Clear it
                  END IF                                          '
                  INCR j                                          ' Onward in the Invisible block
               LOOP                                               '
               i = j                                              ' Swap j to be continue line

            ELSEIF IsLInvisible(i) THEN                           ' A start of Invisible line block?

               '----- We have an Invisible line, but no preceeding Exclude marker
               j = i                                              ' Save insertion point
               me.LInsertEmpty(j - 1, 1, %Xclude)                 ' Insert a line to hold the Exclude marker
               LWrk1S(j) = 0                                      ' Set Work to zero
               LLNumS(j) = "00000000"                             ' Set KNum to zero
               me.UpdLControl(j)                                  ' Get LLCtl set
               IF IsMEdit THEN                                    ' In MEdit mode?
                  fMIX = me.MEditTbl("I", FORMAT$(j - 1))         ' Go get the MIX value of previous line
                  LMixS(j) = fMIX                                 ' Set into the new line
               END IF                                             '
               i = j                                              ' Adjust i to the inserted Exclude marker
               ITERATE DO                                         ' Back to outer loop

            ELSE                                                  ' Just a line we don't care about right now

               '----- Just a normal line, continue onward
               GOSUB TestSpec                                     ' Look for our special cursor marking lines
               IF ISFALSE IsLBottom(i) THEN INCR i                ' Bump and continue
            END IF                                                '
         LOOP                                                     '
         IF gMacroMode AND FLineNumber <> "00000000" THEN         ' If macro and an FLine then re-create the cfFLine
            cfFline = me.LLNumScan(FLineNumber) - 1               ' Set to the current LPtr
         END IF

         '----- Adjust Dotend if sitting at LastLine
         IF InitLastLine = DotEnd + 1 AND LastLine <> InitLastLine THEN ' If LastLine moved and DotEnd was pointing there
            DotEnd += (LastLine - InitLastLine)                   ' Adjust DotEnd to match
            DotEndT = DotEnd                                      '
         END IF
         MExitmeth

      TestSpec:
         me.LFlagBitOff(i, %Popped)                               ' Remove as popped during this interaction
         IF IsLData(i) THEN MaxLength = MAX(MaxLength, LEN(L(i).@LTxt)) ' Track MaxLength
         IF IsLCursor(i) THEN me.LFlagBitOff(i, %Cursor): csr = i '
         IF IsLScroll(i) THEN me.LFlagBitOff(i, %Scroll): scrl = i'
         IF IsLTopScreen(i) THEN me.LFlagBitOff(i, %TopScreen): top = i
         RETURN                                                   '
      END METHOD

      #INCLUDE ONCE "_KB.inc"                                     ' Keyboard Primitives routines
      #INCLUDE ONCE "_FMPCmd.inc"                                 ' FM Primary command routines
      #INCLUDE ONCE "_FMLCmd.inc"                                 ' FM Line commands command routines
      #INCLUDE ONCE "_PCmd.inc"                                   ' Edit Primary command routines
      #INCLUDE ONCE "_LCmd.inc"                                   ' Edit Line command routines

   END INTERFACE

END CLASS

'----- High use mini code stubs
FASTPROC QErrorNull: TP.ErrFlag = 0: TP.ErrMsg = "": END FASTPROC ' Set to 'No Error'


'----- Stubs since you can't make a CODEPTR to an Object METHOD
FASTPROC FMkbBackSpace:     TP.FMkbBackSpace:           END FASTPROC
FASTPROC FMkbBackTab:       TP.FMkbBackTab:             END FASTPROC
FASTPROC FMkbChar:          TP.FMkbChar:                END FASTPROC
FASTPROC FMkbClipName:      TP.FMkbClipName:            END FASTPROC
FASTPROC FMkbClipPath:      TP.FMkbClipPath:            END FASTPROC
FASTPROC FMkbCopyLCmd:      TP.FMkbCopyLCmd:            END FASTPROC
FASTPROC FMkbDelete:        TP.FMkbDelete:              END FASTPROC
FASTPROC FMkbDown:          TP.FMkbDown:                END FASTPROC
FASTPROC FMkbEndOfLine:     TP.FMkbEndOfLine("L"):      END FASTPROC
FASTPROC FMkbEndOfText:     TP.FMkbEndOfLine("T"):      END FASTPROC
FASTPROC FMkbEnter:         TP.FMkbEnter:               END FASTPROC
FASTPROC FMkbEraseEOL:      TP.FMkbEraseEOL:            END FASTPROC
FASTPROC FMkbFirstLCmd:     TP.FMkbFirstLCmd:           END FASTPROC
FASTPROC FMkbHome:          TP.FMkbHome:                END FASTPROC
FASTPROC FMkbLeft:          TP.FMkbLeft:                END FASTPROC
FASTPROC FMkbLineNo:        TP.FMkbLineNo:              END FASTPROC
FASTPROC FMkbNewLine:       TP.FMkbNewLine:             END FASTPROC
FASTPROC FMkbRight:         TP.FMkbRight:               END FASTPROC
FASTPROC FMkbPaste:         TP.FMkbPaste:               END FASTPROC
FASTPROC FMkbTab:           TP.FMkbTab:                 END FASTPROC
FASTPROC FMkbUp:            TP.FMkbUp:                  END FASTPROC
SUB FMPCmdAll(pCmd AS STRING):         TP.FMPCmdAll(pCmd):         END SUB
SUB FMPCmdBottom(pCmd AS STRING):      TP.FMPCmdBottom(pCmd):      END SUB
SUB FMPCmdBrowse(pCmd AS STRING):      TP.FMPCmdBrowse(pCmd):      END SUB
SUB FMPCmdCase(pCmd AS STRING):        TP.FMPCmdCase(pCmd):        END SUB
SUB FMPCmdDefault(pCmd AS STRING):     TP.FMPCmdDefault(pCmd):     END SUB
SUB FMPCmdDown(pCmd AS STRING):        TP.FMPCmdDown(pCmd):        END SUB
SUB FMPCmdEdit(pCmd AS STRING):        TP.FMPCmdEdit(pCmd):        END SUB
SUB FMPCmdEND(pCmd AS STRING):         TP.FMPCmdEND(pCmd):         END SUB
SUB FMPCmdExclude(pCmd AS STRING):     TP.FMPCmdExclude(pCmd):     END SUB
SUB FMPCmdFF(pCmd AS STRING):          TP.FMPCmdFF(pCmd):          END SUB
SUB FMPCmdFind(pCmd AS STRING):        TP.FMPCmdFind(pCmd):        END SUB
SUB FMPCmdLocate(pCmd AS STRING):      TP.FMPCmdLocate(pCmd):      END SUB
SUB FMPCmdMakeList(pCmd AS STRING):    TP.FMPCmdMakeList(pCmd):    END SUB
SUB FMPCmdRecall(pCmd AS STRING):      TP.FMPCmdRecall(pCmd):      END SUB
SUB FMPCmdReset(pCmd AS STRING):       TP.FMPCmdReset(pCmd):       END SUB
SUB FMPCmdRFind(pCmd AS STRING):       TP.FMPCmdRFind(pCmd):       END SUB
SUB FMPCmdTop(pCmd AS STRING):         TP.FMPCmdTop(pCmd):         END SUB
SUB FMPCmdView(pCmd AS STRING):        TP.FMPCmdView(pCmd):        END SUB
SUB FMPCmdUp(pCmd AS STRING):          TP.FMPCmdUp(pCmd):          END SUB
FASTPROC krANSI:            TP.krANSI:                  END FASTPROC
FASTPROC krANSIcol:         TP.krANSIcol:               END FASTPROC
FASTPROC krBackTab:         TP.krBackTab:               END FASTPROC
FASTPROC krBPaste:          TP.krBPaste:                END FASTPROC
FASTPROC krBrowse:          TP.krEditBrowse("BROWSE"):  END FASTPROC
FASTPROC krBS:              TP.krBS:                    END FASTPROC
FASTPROC krChar:            TP.krChar:                  END FASTPROC
FASTPROC krClipClear:       TP.krClipClear:             END FASTPROC
FASTPROC krClipDate:        TP.krClipDate:              END FASTPROC
FASTPROC krClipISODate:     TP.krClipISODate:           END FASTPROC
FASTPROC krClipISOTime:     TP.krClipISOTime:           END FASTPROC
FASTPROC krClipName:        TP.krClipName:              END FASTPROC
FASTPROC krClipPath:        TP.krClipPath:              END FASTPROC
FASTPROC krClipTime:        TP.krClipTime:              END FASTPROC
FASTPROC krColumn:          TP.krColumn:                END FASTPROC
FASTPROC krCondLineNo:      TP.krCondLineNo:            END FASTPROC
FASTPROC krCopyLCmd:        TP.krCopyLCmd:              END FASTPROC
FASTPROC krCopyPaste:       TP.krCopyPaste:             END FASTPROC
FASTPROC krCopyPasteAdd:    TP.krCopyPasteAdd:          END FASTPROC
FASTPROC krCopyPasteRaw:    TP.krCopyPasteRaw:          END FASTPROC
FASTPROC krCtlLeft:         TP.krCtlLeft:               END FASTPROC
FASTPROC krCtlRight:        TP.krCtlRight:              END FASTPROC
FASTPROC krDataBackspace:   TP.krDataDelBS("B"):        END FASTPROC
FASTPROC krDataDelete:      TP.krDataDelBS("D"):        END FASTPROC
FASTPROC krDataDeleteMark:  TP.krDataDeleteMark:        END FASTPROC
FASTPROC krDataInsert:      TP.krDataInsert:            END FASTPROC
FASTPROC krDate:            TP.krDate:                  END FASTPROC
FASTPROC krDel:             TP.krDel:                   END FASTPROC
FASTPROC krDeleteMark:      TP.krDeleteMark:            END FASTPROC
FASTPROC krDown:            TP.krDown:                  END FASTPROC
FASTPROC krDup:             TP.krDup:                   END FASTPROC
FASTPROC krEdit:            TP.krEditBrowse("EDIT"):    END FASTPROC
FASTPROC krEndLine:         TP.krEndLine:               END FASTPROC
FASTPROC krEndText:         TP.krEndText:               END FASTPROC
FASTPROC krEnter:           TP.krEnter:                 END FASTPROC
FASTPROC krEnumD:           TP.krEnum("D"):             END FASTPROC
FASTPROC krEnumL:           TP.krEnum("L"):             END FASTPROC
FASTPROC krEnumU:           TP.krEnum("U"):             END FASTPROC
FASTPROC krEOL:             TP.krEOL:                   END FASTPROC
FASTPROC krErase:           TP.krErase:                 END FASTPROC
FASTPROC krFindNext:        TP.krFind("NEXT"):          END FASTPROC
FASTPROC krFindPrev:        TP.krFind("PREV"):          END FASTPROC
FASTPROC krFirstLCmd:       TP.krFirstLCmd:             END FASTPROC
FASTPROC krHome:            TP.krHome:                  END FASTPROC
FASTPROC krInsert:          TP.krinsert:                END FASTPROC
FASTPROC krISODate:         TP.krISODate:               END FASTPROC
FASTPROC krISOTime:         TP.krISOTime:               END FASTPROC
FASTPROC krJustify:         TP.krJustify:               END FASTPROC
FASTPROC krKCpy:            TP.krKCpy:                  END FASTPROC
FASTPROC krKCpyAdd:         TP.krKCpyAdd:               END FASTPROC
FASTPROC krKCpyRaw:         TP.krKCpyRaw:               END FASTPROC
FASTPROC krKCut:            TP.krKCut:                  END FASTPROC
FASTPROC krLastTab:         TP.krLastTab:               END FASTPROC
FASTPROC krLeft:            TP.krLeft:                  END FASTPROC
FASTPROC krLift:            TP.krLift:                  END FASTPROC
FASTPROC krLineNo:          TP.krLineNo:                END FASTPROC
FASTPROC krLowerCase:       TP.krLowerCase:             END FASTPROC
FASTPROC krMarkD:           TP.krMark("D", IIF(IsFMTab, CODEPTR(FMkbDown), CODEPTR(krDown))): END FASTPROC
FASTPROC krMarkE:           TP.krMark("E", IIF(IsFMTab, CODEPTR(FMkbEndofText), CODEPTR(krEndText))): END FASTPROC
FASTPROC krMarkL:           TP.krMark("L", IIF(IsFMTab, CODEPTR(FMkbLeft), CODEPTR(krLeft))): END FASTPROC
FASTPROC krMarkR:           TP.krMark("R", IIF(IsFMTab, CODEPTR(FMkbRight), CODEPTR(krRight))): END FASTPROC
FASTPROC krMarkU:           TP.krMark("U", IIF(IsFMTab, CODEPTR(FMkbUp), CODEPTR(krUp))): END FASTPROC
FASTPROC krNewLine:         TP.krNewLine:               END FASTPROC
FASTPROC krNewLineNS:       TP.krNewLine(%True):        END FASTPROC
FASTPROC krPaste:           TP.krPaste:                 END FASTPROC
FASTPROC krPen:             TP.krPenCommon:             END FASTPROC
FASTPROC krPenBlue:         TP.krPenBlue:               END FASTPROC
FASTPROC krPenGreen:        TP.krPenGreen:              END FASTPROC
FASTPROC krPenRed:          TP.krPenRed:                END FASTPROC
FASTPROC krPenYellow:       TP.krPenYellow:             END FASTPROC
FASTPROC krPenStd:          TP.krPenStd:                END FASTPROC
FASTPROC krRecord:          TP.krRecord:                END FASTPROC
FASTPROC krResetInsert:     TP.krResetInsert:           END FASTPROC
FASTPROC krRestoreCursor:   TP.krRestoreCursor:         END FASTPROC
FASTPROC krRestoreInsert:   TP.krRestoreInsert:         END FASTPROC
FASTPROC krRight:           TP.krRight:                 END FASTPROC
FASTPROC krSaveCursor:      TP.krSaveCursor:            END FASTPROC
FASTPROC krScrollD:         TP.krScrollD:               END FASTPROC
FASTPROC krScrollL:         TP.krScrollL:               END FASTPROC
FASTPROC krScrollR:         TP.krScrollR:               END FASTPROC
FASTPROC krScrollU:         TP.krScrollU:               END FASTPROC
FASTPROC krSentenceCase:    TP.krSentenceCase:          END FASTPROC
FASTPROC krSetInsert:       TP.krSetInsert:             END FASTPROC
FASTPROC krSwap:            TP.krSwap:                  END FASTPROC
FASTPROC krTab:             TP.krTab:                   END FASTPROC
FASTPROC krTime:            TP.krTime:                  END FASTPROC
FASTPROC krTitleCase:       TP.krTitleCase:             END FASTPROC
FASTPROC krToggleHome:      TP.krToggleHome:            END FASTPROC
FASTPROC krToggleSelect:    TP.krToggleSelect:          END FASTPROC
FASTPROC krTxtHome:         TP.krTxtHome:               END FASTPROC
FASTPROC krTxtNewLine:      TP.krTxtNewLine:            END FASTPROC
FASTPROC krTxtNewLineNS:    TP.krTxtNewLine(%True):     END FASTPROC
FASTPROC krUp:              TP.krUp:                    END FASTPROC
FASTPROC krUpperCase:       TP.krUpperCase:             END FASTPROC
FASTPROC krView:            TP.krEditBrowse("VIEW"):    END FASTPROC

SUB pCmdAction(pCmd AS STRING)     : TP.pCmdAction(pCmd):        END SUB
SUB pCmdAdd(pCmd AS STRING)        : TP.pCmdAdd(pCmd):           END SUB
SUB pCmdAppend(pCmd AS STRING)     : TP.pCmdEXCLFLIPSHOW(pCmd, "APPEND", "Appended"): END SUB
SUB pCmdAutoBkup(pCmd AS STRING)   : TP.pCmdAutoBkup(pCmd):      END SUB
SUB pCmdAutoCaps(pCmd AS STRING)   : TP.pCmdAutoCaps(pCmd):      END SUB
SUB pCmdAutoNum(pCmd AS STRING)    : TP.pCmdAutoNum(pCmd):       END SUB
SUB pCmdAutoSave(pCmd AS STRING)   : TP.pCmdAutoSave(pCmd):      END SUB
SUB pCmdBottom(pCmd AS STRING)     : TP.pCmdBottom(pCmd):        END SUB
SUB pCmdBounds(pCmd AS STRING)     : TP.pCmdBounds(pCmd):        END SUB
SUB pCmdBrowse(pCmd AS STRING)     : TP.pCmdBrowse(pCmd):        END SUB
SUB pCmdCancel(pCmd AS STRING)     : TP.pCmdCancel(pCmd):        END SUB
SUB pCmdCaps(pCmd AS STRING)       : TP.pCmdCaps(pCmd):          END SUB
SUB pCmdCase(pCmd AS STRING)       : TP.pCmdCase(pCmd):          END SUB
SUB pCmdChange(pCmd AS STRING)     : TP.pCmdChange(pCmd):        END SUB
SUB pCmdClip(pCmd AS STRING)       : TP.pCmdClip(pCmd):          END SUB
SUB pCmdClone(pCmd AS STRING)      : TP.pCmdClone(pCmd):         END SUB
SUB pCmdCmd(pCmd AS STRING)        : TP.pCmdCmd(pCmd):           END SUB
SUB pCmdCollate(pCmd AS STRING)    : TP.pCmdCollate(pCmd):       END SUB
SUB pCmdCols(pCmd AS STRING)       : TP.pCmdCols(pCmd):          END SUB
SUB pCmdCompress(pCmd AS STRING)   : TP.pCmdCompress(pCmd):      END SUB
SUB pCmdCopy(pCmd AS STRING)       : TP.pCmdCopy(pCmd):          END SUB
SUB pCmdCreate(pCmd AS STRING)     : TP.pCmdCreate(pCmd):        END SUB
SUB pCmdCut(pCmd AS STRING)        : TP.pCmdCut(pCmd):           END SUB
SUB pCmdDelete(pCmd AS STRING)     : TP.pCmdDelete(pCmd):        END SUB
SUB pCmdDo(pCmd AS STRING)         : TP.pCmdDo(pCmd):            END SUB
SUB pCmdDown(pCmd AS STRING)       : TP.pCmdDown(pCmd):          END SUB
SUB pCmdDir(pCmd AS STRING)        : TP.pCmdDir(pCmd):           END SUB
SUB pCmdDropKeep(pCmd AS STRING)   : TP.pCmdDropKeep(pCmd):      END SUB
SUB pCmdEdit(pCmd AS STRING)       : TP.pCmdEdit(pCmd):          END SUB
SUB pCmdEnd(pCmd AS STRING)        : TP.pCmdEnd(pCmd):           END SUB
SUB pCmdEnumWith(pCmd AS STRING)   : TP.pCmdEnumWith(pCmd):      END SUB
SUB pCmdEOL(pCmd AS STRING)        : TP.pCmdEOL(pCmd):           END SUB
SUB pCmdExclude(pCmd AS STRING)    : TP.pCmdEXCLFLIPSHOW(pCmd, "EXCLUDE", "Excluded"): END SUB
SUB pCmdFav(pCmd AS STRING)        : TP.pCmdFav(pCmd):           END SUB
SUB pCmdFind(pCmd AS STRING)       : TP.pCmdFind(pCmd):          END SUB
SUB pCmdFlip(pCmd AS STRING)       : TP.pCmdEXCLFLIPSHOW(pCmd, "FLIP", "Flipped"): END SUB
SUB pCmdFold(pCmd AS STRING)       : TP.pCmdFold(pCmd):          END SUB
SUB pCmdGlueWith(pCmd AS STRING)   : TP.pCmdGlueWith(pCmd):      END SUB
SUB pCmdHelp(pCmd AS STRING)       : TP.pCmdHelp(pCmd):          END SUB
SUB pCmdHex(pCmd AS STRING)        : TP.pCmdHex(pCmd):           END SUB
SUB pCmdHide(pCmd AS STRING)       : TP.pCmdHide(pCmd):          END SUB
SUB pCmdHilite(pCmd AS STRING)     : TP.pCmdHilite(pCmd):        END SUB
SUB pCmdJoin(pCmd AS STRING)       : TP.pCmdJoin(pCmd):          END SUB
SUB pCmdKeyMap(pCmd AS STRING)     : TP.pCmdKeyMap(pCmd):        END SUB
SUB pCmdLCUC(pCmd AS STRING)       : TP.pCmdLCUC(pCmd):          END SUB
SUB pCmdLeft(pCmd AS STRING)       : TP.pCmdLeft(pCmd):          END SUB
SUB pCmdLine(pCmd AS STRING)       : TP.pCmdLine(pCmd):          END SUB
SUB pCmdLocate(pCmd AS STRING)     : TP.pCmdLocate(pCmd):        END SUB
SUB pCmdLRECL(pCmd AS STRING)      : TP.pCmdLRECL(pCmd):         END SUB
SUB pCmdMacro(pCmd AS STRING)      : TP.pCmdMacro(pCmd):         END SUB
SUB pCmdMark(pCmd AS STRING)       : TP.pCmdMark(pCmd):          END SUB
SUB pCmdMEdit(pCmd AS STRING)      : TP.pCmdMEdit(pCmd):         END SUB
SUB pCmdMinLen(pCmd AS STRING)     : TP.pCmdMinLen(pCmd):        END SUB
SUB pCmdNDelete(pCmd AS STRING)    : TP.pCmdEXCLFLIPSHOW(pCmd, "NDELETE", "Deleted"): END SUB
SUB pCmdNExclude(pCmd AS STRING)   : TP.pCmdEXCLFLIPSHOW(pCmd, "NEXCLUDE", "Excluded"): END SUB
SUB pCmdNFind(pCmd AS STRING)      : TP.pCmdNFind(pCmd):         END SUB
SUB pCmdNFlip(pCmd AS STRING)      : TP.pCmdEXCLFLIPSHOW(pCmd, "NFLIP", "Flipped"): END SUB
SUB pCmdNotify(pCmd AS STRING)     : TP.pCmdNotify(pCmd):        END SUB
SUB pCmdNRevert(pCmd AS STRING)    : TP.pCmdEXCLFLIPSHOW(pCmd, "NREVERT", "UnMarked"): END SUB
SUB pCmdNUline(pCmd AS STRING)     : TP.pCmdEXCLFLIPSHOW(pCmd, "NULINE", "Marked"): END SUB
SUB pCmdNumber(pCmd AS STRING)     : TP.pCmdNumber(pCmd):        END SUB
SUB pCmdNumType(pCmd AS STRING)    : TP.pCmdNumType(pCmd):       END SUB
SUB pCmdNShow(pCmd AS STRING)      : TP.pCmdEXCLFLIPSHOW(pCmd, "NSHOW", "Showed"): END SUB
SUB pCmdOpen(pCmd AS STRING)       : TP.pCmdOpen(pCmd):          END SUB
SUB pCmdOptions(pCmd AS STRING)    : TP.pCmdOptions(pCmd):       END SUB
SUB pCmdOrder(pCmd AS STRING)      : TP.pCmdOrder(pCmd):         END SUB
SUB pCmdPage(pCmd AS STRING)       : TP.pCmdPage(pCmd):          END SUB
SUB pCmdPaste(pCmd AS STRING)      : TP.pCmdPaste(pCmd):         END SUB
SUB pCmdPlease(pCmd AS STRING)     : TP.pCmdPlease(pCmd):        END SUB
SUB pCmdPrepend(pCmd AS STRING)    : TP.pCmdEXCLFLIPSHOW(pCmd, "PREPEND", "Prepended"): END SUB
SUB pCmdPreserve(pCmd AS STRING)   : TP.pCmdPreserve(pCmd):      END SUB
SUB pCmdPrint(pCmd AS STRING)      : TP.pCmdPrint(pCmd):         END SUB
SUB pCmdProfile(pCmd AS STRING)    : TP.pCmdProfile(pCmd):       END SUB
SUB pCmdPType(pCmd AS STRING)      : TP.pCmdPType(pCmd):         END SUB
SUB pCmdQuery(pCmd AS STRING)      : TP.pCmdQuery(pCmd):         END SUB
SUB pCmdRChange(pCmd AS STRING)    : TP.pCmdRChange(pCmd):       END SUB
SUB pCmdRDelete(pCmd AS STRING)    : TP.pCmdRDelete(pCmd):       END SUB
SUB pCmdRecall(pCmd AS STRING)     : TP.pCmdRecall(pCmd):        END SUB
SUB pCmdRECFM(pCmd AS STRING)      : TP.pCmdRECFM(pCmd):         END SUB
SUB pCmdREDO(pCmd AS STRING)       : TP.pCmdUndoRedo("REDO"):    END SUB
SUB pCmdReload(pCmd AS STRING)     : TP.pCmdReload(pCmd):        END SUB
SUB pCmdRename(pCmd AS STRING)     : TP.pCmdRename(pCmd):        END SUB
SUB pCmdReNumber(pCmd AS STRING)   : TP.pCmdReNumber(pCmd):      END SUB
SUB pCmdReplace(pCmd AS STRING)    : TP.pCmdReplace(pCmd):       END SUB
SUB pCmdReset(pCmd AS STRING)      : TP.pCmdReset(pCmd):         END SUB
SUB pCmdRetFwd(pCmd AS STRING)     : TP.pCmdRetFwd(pCmd):        END SUB
SUB pCmdRetrieve(pCmd AS STRING)   : TP.pCmdRetrieve(pCmd):      END SUB
SUB pCmdRevert(pCmd AS STRING)     : TP.pCmdEXCLFLIPSHOW(pCmd, "REVERT", "UnMarked"): END SUB
SUB pCmdRFind(pCmd AS STRING)      : TP.pCmdRFind(pCmd):         END SUB
SUB pCmdRight(pCmd AS STRING)      : TP.pCmdRight(pCmd):         END SUB
SUB pCmdRJoin(pCmd AS STRING)      : TP.pCmdRJoin(pCmd):         END SUB
SUB pCmdRLoc(pCmd AS STRING)       : TP.pCmdRLoc(pCmd):          END SUB
SUB pCmdRLocFind(pCmd AS STRING)   : TP.pCmdRLocFind(pCmd):      END SUB
SUB pCmdRNFind(pCmd AS STRING)     : TP.pCmdRNFind(pCmd):        END SUB
SUB pCmdRSplit(pCmd AS STRING)     : TP.pCmdRSplit(pCmd):        END SUB
SUB pCmdRun(pCmd AS STRING)        : TP.pCmdRun(pCmd):           END SUB
SUB pCmdSave(pCmd AS STRING)       : TP.pCmdSave(pCmd):          END SUB
SUB pCmdSaveAll(pCmd AS STRING)    : TP.pCmdSaveAll(pCmd):       END SUB
SUB pCmdSaveAs(pCmd AS STRING)     : TP.pCmdSaveAs(pCmd):        END SUB
SUB pCmdSaveAsMEdit(pCmd AS STRING): TP.pCmdSaveAsMEdit(pCmd):   END SUB
SUB pCmdScroll(pCmd AS STRING)     : TP.pCmdScroll(pCmd):        END SUB
SUB pCmdSet(pCmd AS STRING)        : TP.pCmdSet(pCmd):           END SUB
SUB pCmdSetUndo(pCmd AS STRING)    : TP.pCmdSetUndo(pCmd):       END SUB
SUB pCmdShow(pCmd AS STRING)       : TP.pCmdEXCLFLIPSHOW(pCmd, "SHOW", "Showed"): END SUB
SUB pCmdSort(pCmd AS STRING)       : TP.pCmdSort(pCmd):          END SUB
SUB pCmdSource(pCmd AS STRING)     : TP.pCmdSource(pCmd):        END SUB
SUB pCmdSplit(pCmd AS STRING)      : TP.pCmdSplit(pCmd):         END SUB
SUB pCmdStart(pCmd AS STRING)      : TP.pCmdStart(pCmd):         END SUB
SUB pCmdState(pCmd AS STRING)      : TP.pCmdState(pCmd):         END SUB
SUB pCmdSubArg(pCmd AS STRING)     : TP.pCmdSubArg(pCmd):        END SUB
SUB pCmdSubCmd(pCmd AS STRING)     : TP.pCmdSubCmd(pCmd):        END SUB
SUB pCmdSubmit(pCmd AS STRING)     : TP.pCmdSubmit(pCmd):        END SUB
SUB pCmdSwap(pCmd AS STRING)       : TP.pCmdSwap(pCmd):          END SUB
SUB pCmdTabs(pCmd AS STRING)       : TP.pCmdTabs(pCmd):          END SUB
SUB pCmdTag(pCmd AS STRING)        : TP.pCmdTag(pCmd):           END SUB
SUB pCmdTest(pCmd AS STRING)       : TP.pCmdTest(pCmd):          END SUB
SUB pCmdTop(pCmd AS STRING)        : TP.pCmdTop(pCmd):           END SUB
SUB pCmdUp(pCmd AS STRING)         : TP.pCmdUp(pCmd):            END SUB
SUB pCmdULine(pCmd AS STRING)      : TP.pCmdEXCLFLIPSHOW(pCmd, "ULINE", "Marked"): END SUB
SUB pCmdUNDO(pCmd AS STRING)       : TP.pCmdUndoRedo("UNDO"):    END SUB
SUB pCmdUnNumber(pCmd AS STRING)   : TP.pCmdUnNumber(pCmd):      END SUB
SUB pCmdView(pCmd AS STRING)       : TP.pCmdView(pCmd):          END SUB
SUB pCmdVSave(pCmd AS STRING)      : TP.pCmdVSave(pCmd):         END SUB
SUB pCmdWDir(pCmd AS STRING)       : TP.pCmdWDir(pCmd):          END SUB
SUB pCmdXSubmit(pCmd AS STRING)    : TP.pCmdXSubmit(pCmd):       END SUB
SUB pCmdXTabs(pCmd AS STRING)      : TP.pCmdXTabs(pCmd):         END SUB

FASTPROC PTBackTab:         TP.PTBackTab:               END FASTPROC
FASTPROC PTBS:              TP.PTBS:                    END FASTPROC
FASTPROC PTChar:            TP.PTChar:                  END FASTPROC
FASTPROC PTCopyPaste:       TP.PTCopyPaste:             END FASTPROC
FASTPROC PTCopyPasteAdd:    TP.PTCopyPasteAdd:          END FASTPROC
FASTPROC PTCopyPasteRaw:    TP.PTCopyPasteRaw:          END FASTPROC
FASTPROC PTCtlLeft:         TP.PTCtlLeft:               END FASTPROC
FASTPROC PTCtlRight:        TP.PTCtlRight:              END FASTPROC
FASTPROC PTCut:             TP.PTCut:                   END FASTPROC
FASTPROC PTDataBackSpace:   TP.PTDataDelBS("B"):        END FASTPROC
FASTPROC PTDataDelete:      TP.PTDataDelBS("D"):        END FASTPROC
FASTPROC PTDate:            TP.PTDate:                  END FASTPROC
FASTPROC PTDel:             TP.PTDel:                   END FASTPROC
FASTPROC PTDup:             TP.PTDup:                   END FASTPROC
FASTPROC PTEnter:           TP.PTEnter:                 END FASTPROC
FASTPROC PTEnumD:           TP.PTEnum("D"):             END FASTPROC
FASTPROC PTEnumL:           TP.PTEnum("L"):             END FASTPROC
FASTPROC PTEnumU:           TP.PTEnum("U"):             END FASTPROC
FASTPROC PTEOL:             TP.PTEOL:                   END FASTPROC
FASTPROC PTErase:           TP.PTErase:                 END FASTPROC
FASTPROC PTISODate:         TP.PTISODate:               END FASTPROC
FASTPROC PTISOTime:         TP.PTISOTime:               END FASTPROC
FASTPROC PTJustify:         TP.PTJustify:               END FASTPROC
FASTPROC PTLastTab:         TP.PTLastTab:               END FASTPROC
FASTPROC PTLeft:            TP.PTLeft:                  END FASTPROC
FASTPROC PTLift:            TP.PTLift:                  END FASTPROC
FASTPROC PTLowerCase:       TP.PTLowerCase:             END FASTPROC
FASTPROC PTMarkE:           TP.PTMark("E", CODEPTR(krEndText)): END FASTPROC
FASTPROC PTMarkL:           TP.PTMark("L", CODEPTR(PTLeft)): END FASTPROC
FASTPROC PTMarkR:           TP.PTMark("R", CODEPTR(PTRight)): END FASTPROC
FASTPROC PTPaste:           TP.PTPaste:                 END FASTPROC
FASTPROC PTPowerCopy:       TP.PTPowerCopy:             END FASTPROC
FASTPROC PTPowerCut:        TP.PTPowerCut:              END FASTPROC
FASTPROC PTPowerPaste:      TP.PTPowerPaste:            END FASTPROC
FASTPROC PTRight:           TP.PTRight:                 END FASTPROC
FASTPROC PTSentenceCase:    TP.PTSentenceCase:          END FASTPROC
FASTPROC PTSeq:             TP.PTSeq:                   END FASTPROC
FASTPROC PTSwap:            TP.PTSwap:                  END FASTPROC
FASTPROC PTTab:             TP.PTTab:                   END FASTPROC
FASTPROC PTTime:            TP.PTTime:                  END FASTPROC
FASTPROC PTTitleCase:       TP.PTTitleCase:             END FASTPROC
FASTPROC PTUpperCase:       TP.PTUpperCase:             END FASTPROC

SUB      lCmdBNDS(x AS STRING):        TP.LCmdBNDS(x):            END SUB
SUB      lCmdCC(x AS STRING):          TP.LCmdCC(x):              END SUB
SUB      lCmdCOLS(x AS STRING):        TP.LCmdCOLS(x):            END SUB
SUB      lCmdDD(x AS STRING):          TP.LCmdDD(x):              END SUB
SUB      lCmdF(x AS STRING):           TP.LCmdF(x):               END SUB
SUB      lCmdI(x AS STRING):           TP.LCmdI(x):               END SUB
SUB      lCmdJJ(x AS STRING):          TP.LCmdJJ(x):              END SUB
SUB      lCmdL(x AS STRING):           TP.LCmdL(x):               END SUB
SUB      lCmdLCLC(x AS STRING):        TP.LCmdBasic(x):           END SUB
SUB      lCmdLDShift(x AS STRING):     TP.LCmdBasic(x):           END SUB
SUB      lCmdLIShift(x AS STRING):     TP.LCmdBasic(x):           END SUB
SUB      lCmdLShift(x AS STRING):      TP.LCmdBasic(x):           END SUB
SUB      lCmdMakeData(x AS STRING):    TP.lCmdMakeDN(x, %Data):   END SUB
SUB      lCmdMakeNote(x AS STRING):    TP.lCmdMakeDN(x, %Note):   END SUB
SUB      lCmdMark(x AS STRING):        TP.lCmdMark(x):            END SUB
SUB      lCmdMask(x AS STRING):        TP.lCmdMask(x):            END SUB
SUB      lCmdMM(x AS STRING):          TP.lCmdMM(x):              END SUB
SUB      lCmdNote(x AS STRING):        TP.lCmdNote(x):            END SUB
SUB      lCmdPL(x AS STRING):          TP.lCmdPL(x):              END SUB
SUB      lCmdRDShift(x AS STRING):     TP.LCmdBasic(x):           END SUB
SUB      lCmdRIShift(x AS STRING):     TP.LCmdBasic(x):           END SUB
SUB      lCmdRR(x AS STRING):          TP.LCmdRR(x):              END SUB
SUB      lCmdRShift(x AS STRING):      TP.LCmdBasic(x):           END SUB
SUB      lCmdSC(x AS STRING):          TP.LCmdBasic(x):           END SUB
SUB      lCmdSI(x AS STRING):          TP.LCmdSI(x):              END SUB
SUB      lCmdSS(x AS STRING):          TP.LCmdSS(x):              END SUB
SUB      lCmdTABS(x AS STRING):        TP.LCmdTABS(x):            END SUB
SUB      lCmdTC(x AS STRING):          TP.LCmdBasic(x):           END SUB
SUB      lCmdTF(x AS STRING):          TP.LCmdTF(x):              END SUB
SUB      lCmdTJJ(x AS STRING):         TP.LCmdTJJ(x):             END SUB
SUB      lCmdTL(x AS STRING):          TP.LCmdTL(x):              END SUB
SUB      lCmdTR(x AS STRING):          TP.LCmdTR(x):              END SUB
SUB      lCmdTS(x AS STRING):          TP.LCmdTS(x):              END SUB
SUB      lCmdTT(x AS STRING):          TP.LCmdTT(x):              END SUB
SUB      lCmdTU(x AS STRING):          TP.LCmdTU(x):              END SUB
SUB      lCmdTX(x AS STRING):          TP.LCmdTX(x):              END SUB
SUB      lCmdUCUC(x AS STRING):        TP.LCmdBasic(x):           END SUB
SUB      lCmdUU(x AS STRING):          TP.LCmdBasic(x):           END SUB
SUB      lCmdVV(x AS STRING):          TP.LCmdBasic(x):           END SUB
SUB      lCmdWord(x AS STRING):        TP.LCmdWord(x):            END SUB
SUB      lCmdXX(x AS STRING):          TP.LCmdXX(x):              END SUB

SUB      pCmdModel(pCmd AS STRING)
'---------- Model method to use in CALL DWORD USING
END SUB
