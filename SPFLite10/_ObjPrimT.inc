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

CLASS cPrimTable

   '----- Primitive Data
   INSTANCE PrimCount         AS LONG                             ' Number of Primitives
   INSTANCE PrimTable()       AS iPrimitiveEntry                  ' Table
   INSTANCE PName()           AS STRING                           ' Primitive name list
   INSTANCE PrimitiveList()   AS STRING                           ' List for KEYMAP

   CLASS METHOD CREATE()                                          ' Constructor - Initialize Class stuff
   LOCAL i AS LONG
      DIM PName(1 TO 127) AS INSTANCE STRING                      ' Dim the Name table
      DIM PrimTable (1 TO 127) AS INSTANCE iPrimitiveEntry        ' Dim the Primitive data table

      ' DIM size for PName and PrimTable above must be as large as number of Me.InitPrim calls below
      ' The current number of Me.InitPrim calls below is 101

      Me.InitPrim("(BackSpace)", %False, %True,       CODEPTR(krBS),           CODEPTR(FMkbBackSpace),    CODEPTR(PTBS))
      Me.InitPrim("(BackTab)", %False, %True,         CODEPTR(krBackTab),      CODEPTR(FMkbBackTab),      CODEPTR(PTBackTab))
      Me.InitPrim("(BlockPaste)", %False, %False,     CODEPTR(krBPaste),       0,                         0)
      Me.InitPrim("(Browse)", %True, %True,           CODEPTR(krBrowse),       0,                         0)
      Me.InitPrim("(CharSet)", %False, %True,         CODEPTR(krANSI),         CODEPTR(krANSI),           CODEPTR(krANSI))
      Me.InitPrim("(CharSetCol)", %False, %True,      CODEPTR(krANSIcol),      CODEPTR(krANSIcol),        CODEPTR(krANSIcol))
      Me.InitPrim("(ClipClear)", %False, %True,       CODEPTR(krClipClear),    CODEPTR(krClipClear),      CODEPTR(krClipClear))
      Me.InitPrim("(ClipDate)", %False, %True,        CODEPTR(krClipDate),     CODEPTR(krClipDate),       CODEPTR(krClipDate))
      Me.InitPrim("(ClipISODate)", %False, %True,     CODEPTR(krClipISODate),  CODEPTR(krClipISODate),    CODEPTR(krClipISODate))
      Me.InitPrim("(ClipISOTime)", %False, %True,     CODEPTR(krClipISOTime),  CODEPTR(krClipISOTime),    CODEPTR(krClipISOTime))
      Me.InitPrim("(ClipName)", %False, %True,        CODEPTR(krClipName),     CODEPTR(FMkbClipName),     CODEPTR(krClipName))
      Me.InitPrim("(ClipPath)", %False, %True,        CODEPTR(krClipPath),     CODEPTR(FMkbClipPath),     CODEPTR(krClipPath))
      Me.InitPrim("(ClipTime)", %False, %True,        CODEPTR(krClipTime),     CODEPTR(krClipTime),       CODEPTR(krClipTime))
      Me.InitPrim("(Column)", %False, %True,          CODEPTR(krColumn),       0,                         0)
      Me.InitPrim("(CondLineNo)", %False, %True,      CODEPTR(krCondLineNo),   CODEPTR(FMkbLineNo),       0)
      Me.InitPrim("(Copy)", %True, %True,             CODEPTR(krKCpy),         CODEPTR(krKCpy),           CODEPTR(krKCpy))
      Me.InitPrim("(CopyAdd)", %True, %True,          CODEPTR(krKCpyAdd),      0,                         CODEPTR(krKCpyAdd))
      Me.InitPrim("(CopyLCmd)", %False, %True,        CODEPTR(krCopyLCmd),     CODEPTR(FMkbCopyLCmd),     CODEPTR(krCopyLCmd))
      Me.InitPrim("(CopyPaste)", %False, %True,       CODEPTR(krCopyPaste),    CODEPTR(krCopyPaste),      CODEPTR(PTCopyPaste))
      Me.InitPrim("(CopyPasteAdd)", %False, %True,    CODEPTR(krCopyPasteAdd), 0,                         CODEPTR(PTCopyPasteAdd))
      Me.InitPrim("(CopyPasteRaw)", %False, %True,    CODEPTR(krCopyPasteRaw), 0,                         CODEPTR(PTCopyPasteRaw))
      Me.InitPrim("(CopyRaw)", %True, %True,          CODEPTR(krKCpyRaw),      0,                         CODEPTR(krKCpyRaw))
      Me.InitPrim("(Cut)", %True, %False,             CODEPTR(krKCut),         CODEPTR(krKCut),           CODEPTR(PTCut))
      Me.InitPrim("(DataBackspace)", %False, %False,  CODEPTR(krDataBackSpace),0,                         CODEPTR(PTDataBackspace))
      Me.InitPrim("(DataDelete)", %False, %False,     CODEPTR(krDataDelete),   0,                         CODEPTR(PTDataDelete))
      Me.InitPrim("(DataDeleteMark)", %False, %False, CODEPTR(krDataDeleteMark),0,                        0)
      Me.InitPrim("(DataInsert)", %False, %False,     CODEPTR(krDataInsert),   CODEPTR(krDataInsert),     CODEPTR(krDataInsert))
      Me.InitPrim("(Date)", %False, %False,           CODEPTR(krDate),         0,                         CODEPTR(PTDate))
      Me.InitPrim("(Delete)", %False, %True,          CODEPTR(krDel),          CODEPTR(FMkbDelete),       CODEPTR(PTDel))
      Me.InitPrim("(DeleteMark)", %False, %True,      CODEPTR(krDeleteMark),   0,                         0)
      Me.InitPrim("(Down)", %False, %True,            CODEPTR(krDown),         CODEPTR(FMkbDown),         0)
      Me.InitPrim("(Dup)", %True, %False,             CODEPTR(krDup),          0,                         CODEPTR(PTDup))
      Me.InitPrim("(Edit)", %True, %True,             CODEPTR(krEdit),         0,                         0)
      Me.InitPrim("(EndOfLine)", %False, %True,       CODEPTR(krEndLine),      CODEPTR(FMkbEndOfLine),    CODEPTR(krEndLine))
      Me.InitPrim("(EndOfText)", %False, %True,       CODEPTR(krEndText),      CODEPTR(FMkbEndOfText),    CODEPTR(krEndText))
      Me.InitPrim("(Enter)", %False, %True,           CODEPTR(krEnter),        CODEPTR(FMkbEnter),        CODEPTR(PTEnter))
      Me.InitPrim("(Enum)", %True, %False,            CODEPTR(krEnumD),        0,                         CODEPTR(PTEnumD))
      Me.InitPrim("(EnumHexLC)", %True, %False,       CODEPTR(krEnumL),        0,                         CODEPTR(PTEnumL))
      Me.InitPrim("(EnumHexUC)", %True, %False,       CODEPTR(krEnumU),        0,                         CODEPTR(PTEnumU))
      Me.InitPrim("(Erase)", %True, %False,           CODEPTR(krErase),        0,                         CODEPTR(PTErase))
      Me.InitPrim("(EraseEOL)", %False, %True,        CODEPTR(krEOL),          CODEPTR(FMkbEraseEOL),     CODEPTR(PTEOL))
      Me.InitPrim("(FindNext)", %False, %True,        CODEPTR(krFindNext),     0,                         0)
      Me.InitPrim("(FindPrev)", %False, %True,        CODEPTR(krFindPrev),     0,                         0)
      Me.InitPrim("(FirstLineCmd)", %False, %True,    CODEPTR(krFirstLCmd),    CODEPTR(FMkbFirstLCmd),    0)
      Me.InitPrim("(Home)", %False, %True,            CODEPTR(krHome),         CODEPTR(FMkbHome),         0)
      Me.InitPrim("(Insert)", %False, %True,          CODEPTR(krInsert),       CODEPTR(krInsert),         CODEPTR(krInsert))
      Me.InitPrim("(ISODate)", %False, %False,        CODEPTR(krISODate),      0,                         CODEPTR(PTISODate))
      Me.InitPrim("(ISOTime)", %False, %False,        CODEPTR(krISOTime),      0,                         CODEPTR(PTISOTime))
      Me.InitPrim("(JustifyC)", %True, %False,        CODEPTR(krJustify),      0,                         CODEPTR(PTJustify))
      Me.InitPrim("(JustifyL)", %True, %False,        CODEPTR(krJustify),      0,                         CODEPTR(PTJustify))
      Me.InitPrim("(JustifyR)", %True, %False,        CODEPTR(krJustify),      0,                         CODEPTR(PTJustify))
      Me.InitPrim("(LastTab)", %False, %True,         CODEPTR(krLastTab),      0,                         CODEPTR(PTLastTab))
      Me.InitPrim("(Left)", %False, %True,            CODEPTR(krLeft),         CODEPTR(FMkbLeft),         CODEPTR(PTLeft))
      Me.InitPrim("(Lift)", %True, %True,             CODEPTR(krLift),         CODEPTR(krLift),           CODEPTR(PTLift))
      Me.InitPrim("(LineNo)", %False, %True,          CODEPTR(krLineNo),       CODEPTR(FMkbLineNo),       0)
      Me.InitPrim("(LowerCase)", %True, %False,       CODEPTR(krLowerCase),    0,                         CODEPTR(PTLowerCase))
      Me.InitPrim("(MarkDown)", %False, %True,        CODEPTR(krMarkD),        CODEPTR(krMarkD),          0)
      Me.InitPrim("(MarkEnd)", %False, %True,         CODEPTR(krMarkE),        CODEPTR(krMarkE),          CODEPTR(PTMarkE))
      Me.InitPrim("(MarkLeft)", %False, %True,        CODEPTR(krMarkL),        CODEPTR(krMarkL),          CODEPTR(PTMarkL))
      Me.InitPrim("(MarkRight)", %False, %True,       CODEPTR(krMarkR),        CODEPTR(krMarkR),          CODEPTR(PTMarkR))
      Me.InitPrim("(MarkUp)", %False, %True,          CODEPTR(krMarkU),        CODEPTR(krMarkU),          0)
      Me.InitPrim("(NewLine)", %False, %True,         CODEPTR(krNewLine),      CODEPTR(FMkbNewLine),      0)
      Me.InitPrim("(NewLineNS)", %False, %True,       CODEPTR(krNewLineNS),    CODEPTR(FMkbNewLine),      0)
      Me.InitPrim("(Null)", %False, %True,            0,                       0,                         0)
      Me.InitPrim("(PassThru)", %False, %True,        0,                       0,                         0)
      Me.InitPrim("(Paste)", %False, %True,           CODEPTR(krPaste),        CODEPTR(FMkbPaste),        CODEPTR(PTPaste))
      Me.InitPrim("(Pen)", %True, %True,              CODEPTR(krPen),          0,                         0)
      Me.InitPrim("(PenBlue)", %True, %True,          CODEPTR(krPenBlue),      0,                         0)
      Me.InitPrim("(PenGreen)", %True, %True,         CODEPTR(krPenGreen),     0,                         0)
      Me.InitPrim("(PenRed)", %True, %True,           CODEPTR(krPenRed),       0,                         0)
      Me.InitPrim("(PenYellow)", %True, %True,        CODEPTR(krPenYellow),    0,                         0)
      Me.InitPrim("(PenStd)", %True, %True,           CODEPTR(krPenStd),       0,                         0)
      Me.InitPrim("(PowerCopy)", %True, %False,       0,                       0,                         CODEPTR(PTPowerCopy))
      Me.InitPrim("(PowerCut)", %True, %False,        0,                       0,                         CODEPTR(PTPowerCut))
      Me.InitPrim("(PowerPaste)", %False, %False,     0,                       0,                         CODEPTR(PTPowerPaste))
      Me.InitPrim("(PrtScrnClipboard)", %False, %True,CODEPTR(sPrtScreen),     CODEPTR(sPrtScreen),       0)
      Me.InitPrim("(PrtScrnLog)", %False, %True,      CODEPTR(sPrtScreen),     CODEPTR(sPrtScreen),       0)
      Me.InitPrim("(PrtScrnPrinter)", %False, %True,  CODEPTR(sPrtScreen),     CODEPTR(sPrtScreen),       0)
      Me.InitPrim("(PrtTextClipboard)", %False, %True,CODEPTR(sPrtScreen),     CODEPTR(sPrtScreen),       0)
      Me.InitPrim("(Record)", %False, %True,          CODEPTR(krRecord),       CODEPTR(krRecord),         CODEPTR(krRecord))
      Me.InitPrim("(ResetInsert)", %False, %True,     CODEPTR(krResetInsert),  CODEPTR(krResetInsert),    CODEPTR(krResetInsert))
      Me.InitPrim("(RestoreCursor)", %False, %True,   CODEPTR(krRestoreCursor),0,                         CODEPTR(krRestoreCursor))
      Me.InitPrim("(RestoreInsert)", %False, %True,   CODEPTR(krRestoreInsert),CODEPTR(krRestoreInsert),  CODEPTR(krRestoreInsert))
      Me.InitPrim("(Right)", %False, %True,           CODEPTR(krRight),        CODEPTR(FMkbRight),        CODEPTR(PTRight))
      Me.InitPrim("(SaveCursor)", %False, %True,      CODEPTR(krSaveCursor),   0,                         CODEPTR(krSaveCursor))
      Me.InitPrim("(ScrollDown)", %False, %True,      CODEPTR(krScrollD),      CODEPTR(krScrollD),        0)
      Me.InitPrim("(ScrollLeft)", %False, %True,      CODEPTR(krScrollL),      0,                         CODEPTR(krScrollL))
      Me.InitPrim("(ScrollRight)", %False, %True,     CODEPTR(krScrollR),      0,                         CODEPTR(krScrollR))
      Me.InitPrim("(ScrollUp)", %False, %True,        CODEPTR(krScrollU),      CODEPTR(krScrollU),        0)
      Me.InitPrim("(SentenceCase)", %True, %False,    CODEPTR(krSentenceCase), 0,                         CODEPTR(PTSentenceCase))
      Me.InitPrim("(Seq)", %True, %False,             0,                       0,                         CODEPTR(PTSeq))
      Me.InitPrim("(SetInsert)", %False, %True,       CODEPTR(krSetInsert),    CODEPTR(krSetInsert),      CODEPTR(krSetInsert))
      Me.InitPrim("(Swap)", %False, %False,           CODEPTR(krSwap),         0,                         CODEPTR(PTSwap))
      Me.InitPrim("(Tab)", %False, %True,             CODEPTR(krTab),          CODEPTR(FMkbTab),          CODEPTR(PTTab))
      Me.InitPrim("(Time)", %False, %False,           CODEPTR(krTime),         0,                         CODEPTR(PTTime))
      Me.InitPrim("(TitleCase)", %True, %False,       CODEPTR(krTitleCase),    0,                         CODEPTR(PTTitleCase))
      Me.InitPrim("(ToggleHome)", %False, %True,      CODEPTR(krToggleHome),   0,                         CODEPTR(krToggleHome))
      Me.InitPrim("(ToggleSelect)", %False, %True,    CODEPTR(krToggleSelect), 0,                         0)
      Me.InitPrim("(TxtHome)", %False, %True,         CODEPTR(krTxtHome),      0,                         CODEPTR(krTxtHome))
      Me.InitPrim("(TxtNewLine)", %False, %True,      CODEPTR(krTxtNewLine),   0,                         0)
      Me.InitPrim("(TxtNewLineNS)", %False, %True,    CODEPTR(krTxtNewLineNS), 0,                         0)
      Me.InitPrim("(Up)", %False, %True,              CODEPTR(krUp),           CODEPTR(FMkbUp),           0)
      Me.InitPrim("(UpperCase)", %True, %False,       CODEPTR(krUpperCase),    0,                         CODEPTR(PTUpperCase))
      Me.InitPrim("(View)", %True, %True,             CODEPTR(krView),         0,                         0)
      Me.InitPrim("(WordLeft)", %False, %True,        CODEPTR(krCtlLeft),      0,                         CODEPTR(PTCtlLeft))
      Me.InitPrim("(WordRight)", %False, %True,       CODEPTR(krCtlRight),     0,                         CODEPTR(PTCtlRight))
      DIM gPrimTable(PrimCount) AS STRING                         ' Hint list for KeyMap
      gPrimTable(0) = " Valid keyboard Primitives"                ' Insert heading
      FOR i = 1 TO PrimCount                                      ' Build the KEYMAP list
         gPrimTable(i) = "(" + PName(i) + ")"                     '
      NEXT i
      ARRAY SORT gPrimTable(0) FOR PrimCount + 1                  ' Sort the primitive table

   END METHOD                                                     '

   INTERFACE iPrimTable: INHERIT IUNKNOWN                         ' Define the interface


      METHOD GetPrimIX(Cmd AS STRING) AS LONG
      '----- Validate Primitive name and return Index (0 if invalid)
      REGISTER i AS LONG
         ARRAY SCAN PName() FOR PrimCount, COLLATE UCASE, =Cmd, TO i
         METHOD = i                                               ' Return ARRAY SCAN result
      END METHOD

      METHOD BrowseOK(IX AS LONG) AS LONG
      '----- Return BrowseOK
         METHOD = PrimTable(ix).BrowseOK                          ' Pass back flag
      END METHOD

      METHOD FPath(IX AS LONG) AS LONG
      '----- Return FPath
         METHOD = PrimTable(ix).FPath                             ' Pass back address
      END METHOD

      METHOD NDMarkd(IX AS LONG) AS LONG
      '----- Return NDMarked
         METHOD = PrimTable(ix).NDMarkd                           ' Pass back flag
      END METHOD

      METHOD NPath(IX AS LONG) AS LONG
      '----- Return NPath
         METHOD = PrimTable(ix).NPath                             ' Pass back address
      END METHOD

      METHOD PPath(IX AS LONG) AS LONG
      '----- Return PPath
         METHOD = PrimTable(ix).PPath                             ' Pass back address
      END METHOD

   END INTERFACE

   CLASS METHOD InitPrim(kw AS STRING, NMark AS LONG, OKBrowse AS LONG, norm AS LONG, fm AS LONG, pt AS LONG)
   '---------- Add a single Primitive to the table
   LOCAL Cmd AS STRING
      INCR PrimCount                                                 ' Bump count
      LET PrimTable(PrimCount) = CLASS "cPrimitiveEntry"             ' Allocate an entry
      Cmd = MID$(kw, 2)                                              ' Strip out the name
      Cmd = LEFT$(Cmd, INSTR(Cmd, ")") - 1)                          '
      PName(PrimCount) = Cmd                                         ' Add the primitive name
      PrimTable(PrimCount).NdMarkd = NMark                           ' Add the data items
      PrimTable(PrimCount).BrowseOK = OKBrowse                       '
      PrimTable(PrimCount).NPath = norm                              '
      PrimTable(PrimCount).FPath = fm                                '
      PrimTable(PrimCount).PPath = pt                                '
   END METHOD

END CLASS

CLASS cPrimitiveEntry
   INSTANCE NdMarkd  AS LONG
   INSTANCE NPath    AS LONG
   INSTANCE FPath    AS LONG
   INSTANCE PPath    AS LONG
   INSTANCE BrowseOK AS LONG

   INTERFACE iPrimitiveEntry: INHERIT IUNKNOWN
      gsProp(NdMarkd, LONG)
      gsProp(NPath, LONG)
      gsProp(FPath, LONG)
      gsProp(PPath, LONG)
      gsProp(BrowseOK, LONG)
   END INTERFACE
END CLASS
