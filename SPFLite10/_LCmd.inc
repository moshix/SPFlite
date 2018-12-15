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

'---------- Edit Line commands

METHOD  lCmdBasic(lCmd AS STRING)
'---------- Basic type line command handler
REGISTER x AS LONG
REGISTER i AS LONG
LOCAL j, k, lclLBnd, lclRBnd, Caps, m, NumShift, lclRepeat AS LONG
DIM words() AS STRING
LOCAL Txt1, Txt2, lh, mh, rh AS STRING, w AS lCtlCmd
LOCAL shift_data AS STRING, shift_Attr AS WSTRING
LOCAL shift_error, shift_change, shift_error_total AS LONG

   MEntry
   me.lCmdParse(w, lCmd)                                          ' Parse the command
   me.CurSetReq(%LineCmd, w.SrcFrom, 1 + Offset, %False)          ' Set cursor set attempt
   lclRBnd = Prf.BndRight: IF lclRBnd = 0 THEN lclRBnd = MaxLength' Get working RBnds
   lclRepeat = w.SrcRepeat                                        ' Save outside of line loop
   Caps = %True                                                   ' 1st word overall gets capitalized

   '----- Loop through the line range
   FOR x = w.SrcFrom TO w.SrcTo                                   ' Loop for each line
      IF IsLData(x) OR IsLNote(x) THEN                            ' Only Data/Note lines
         me.lStubDoPM(w.SrcFlag, x)                               ' Do + / - processing for the line

         SELECT CASE AS CONST$ w.SrcCmd                           ' Split up

            '----- LCLC Lowercase command ---------------------------------------------------------
            CASE "LC      ", "LCLC    ", "LCC     "               ' LowerCase command
               me.ModSet(x)                                       ' Remember we changed something
               IF Prf.BndLeft > 1 OR Prf.BndRight > 0 THEN        ' Honour BNDS setting?
                  FOR i = 1 TO LEN(L(x).@LTxt)                    ' Do it the hard way
                     IF i >= Prf.BndLeft AND i <= lclRBnd THEN _  '
                        me.LTxtCharRep(x, i, LLCASE(MID$(LTxtG(x), i, 1)))
                  NEXT i                                          '

               ELSE                                               ' No BNDS, the simple way
                  me.LTxtSet(x, LLCASE(LTxtG(x)))                 ' Do the Lowercase
               END IF                                             '
               me.AttrScan(x)                                     ' Recolorize

            '----- UCUC Uppercase command ---------------------------------------------------------
            CASE "UC      ", "UCUC    ", "UCC     "               ' Uppercase command
               me.ModSet(x)                                       ' Remember we changed something
               IF Prf.BndLeft > 1 OR Prf.BndRight > 0 THEN        ' Honour BNDS setting?
                  FOR i = 1 TO LEN(L(x).@LTxt)                    ' Do it the hard way
                     IF i >= Prf.BndLeft AND i <= lclRBnd THEN _  '
                        me.LTxtCharRep(x, i, UUCASE(MID$(LTxtG(x), i, 1)))
                  NEXT i                                          '

               ELSE                                               ' No BNDS, the simple way
                  me.LTxtSet(x, UUCASE(LTxtG(x)))                 ' Do the Lowercase if a DataLine
               END IF                                             '
               me.AttrScan(x)                                     ' Recolorize

            '----- SCSC Sentence case command -----------------------------------------------------
            CASE "SC      ", "SCSC    ", "SCC     "               ' Sentence Case
               me.ModSet(x)                                       ' Remember we changed something
               IF Prf.BndLeft > 1 OR Prf.BndRight > 0 THEN        ' Honour BNDS setting?
                  Txt1 = LLCASE(MID$(LTxtG(x), Prf.BndLeft, lclRBnd - Prf.BndLeft + 1)) ' Get copy of substring to process
                  GOSUB DoSCase                                   ' Go process this piece
               ELSE                                               ' No BNDS, the simple way
                  Txt1 = LLCASE(MID$(LTxtG(x), 1, LEN(L(x).@LTxt)))' Get copy of substring to process
                  GOSUB DoSCase                                   ' Go process this piece
               END IF                                             '

               txt2 = LTxtG(x)                                    ' Get whole text line
               MID$(txt2, Prf.BndLeft, lclRBnd - Prf.BndLeft + 1) = Txt1' Stuff back in the processed txt
               me.LTxtSet(x, txt2)                                ' Replace the whole line
               me.AttrScan(x)                                     ' Recolorize

            '----- TCTC Title case command --------------------------------------------------------
            CASE "TC      ", "TCTC    ", "TCC     "
               me.ModSet(x)                                       ' Remember we changed something
               IF Prf.BndLeft > 1 OR Prf.BndRight > 0 THEN        ' Honour BNDS setting?
                  me.TitleCase(x, Prf.BndLeft, lclRbnd)           ' Go process it
               ELSE                                               ' No BNDS, the simple way
                  me.TitleCase(x, 1, LEN(L(x).@LTxt))             ' Go process it
               END IF                                             '
               me.AttrScan(x)                                     ' Recolorize

            '----- << Left  Data shift ------------------------------------------------------------
            '----- >> Right Data shift ------------------------------------------------------------

            CASE "<       ", "<<      " , _
                 ">       ", ">>      "
               IF w.SrcRepeat = 0 THEN                            '
                  w.SrcRepeat = ENV.DefDataShift                  ' Default shift if not provided
               END IF                                             '
               NumShift = w.SrcRepeat                             '
               lclLBnd = Prf.BndLeft                              ' Get default lBnd
               lclRBnd = Prf.BndRight                             '
               IF lclRBnd = 0 THEN                                '
                  lclRBnd = MaxLength                             ' Get working RBnds
               END IF                                             '
               shift_data = LTxtG(x)                              '
               shift_attr = LAttrG(x)                             '
               IF LEFT$(w.SrcCmd, 1) = "<" THEN                   ' do left data shift
                  sLCmdDataShiftLeft    _                         '
                      (  shift_data    _                          '
                      ,  shift_attr    _                          '
                      ,  NumShift      _                          '
                      ,  lclLBnd       _                          '
                      ,  lclRBnd       _                          '
                      ,  shift_error   _                          '
                      ,  shift_change  )                          '
               ELSE                                               ' do right data shift
                  sLCmdDataShiftRight   _                         '
                      (  shift_data    _                          '
                      ,  shift_attr    _                          '
                      ,  NumShift      _                          '
                      ,  lclLBnd       _                          '
                      ,  lclRBnd       _                          '
                      ,  shift_error   _                          '
                      ,  shift_change  )                          '
               END IF                                             '
               IF shift_change > 0 THEN                           '
                  me.ModSet(x)                                    ' flag current line as changed
                  me.LTxtSet(x, shift_data)                       ' store new data line
                  LAttrS(x) = shift_attr                          ' update attributes
                  me.AttrScan(x)                                  ' Recolorize
               END IF                                             '

               IF shift_error > 0 THEN                            ' could not shift this line
                  INCR shift_error_total                          ' record num of lines with shift error
               END IF

            '----- (( Shift Left command ----------------------------------------------------------
            CASE "(       ", "((      ", "[       ", "[[      "   '
               me.ModSet(x)                                       ' Remember we changed something
               W.SrcRepeat = lclRepeat                            ' Reset to the default
               IF LEFT$(w.SrcCmd, 1) = "(" THEN                   ' The normal (( shift?
                  W.SrcRepeat = IIF(w.SrcRepeat = 0, ENV.DefDataShift, lclRepeat)
               ELSE                                               ' Must be the [[ shift
                  W.SrcRepeat = IIF(W.SrcRepeat = 0, ENV.DefDataShift, w.SrcRepeat * ENV.DefDataShift) ' If provided, use it, else default
               END IF                                             '

               IF Prf.BndLeft > 1 OR Prf.BndRight > 0 THEN        ' Have to honour BNDS?

                  '----- Short
                  IF LEN(L(x).@LTxt) < Prf.BndLeft THEN           ' Length < Left Bound
                     ' Null                                       ' Do nothing

                  '----- Medium, less than Right Bounds
                  ELSEIF LEN(L(x).@LTxt) <= lclRBnd THEN          ' Length > Left and not longer than Right
                     lh = LEFT$(LTxtG(x), Prf.BndLeft - 1)        ' Save lh untouched portion
                     rh = MID$(LTxtG(x), Prf.BndLeft)             ' Save the rh shiftable portion
                     rh = IIF$(w.SrcRepeat + 1 <= LEN(rh), MID$(rh, w.SrcRepeat + 1), "") ' If shift < line length, pick up rest of line
                     me.LTxtSet(x, lh + rh)                       ' put pieces back together
                     me.LAttrAdjust(x, Prf.BndLeft, -(w.SrcRepeat))' Adjust any Attr hilite
                     me.AttrScan(x)                               ' Recolorize

                  '----- Long, past Right Bounds
                  ELSEIF LEN(L(x).@LTxt) >= lclRBnd THEN          ' Length > Right
                     lh = LEFT$(LTxtG(x), Prf.BndLeft - 1)        ' Save lh untouched portion
                     mh = MID$(LTxtG(x), Prf.BndLeft, lclRBnd - Prf.BndLeft +1) ' Save middle shiftable portion
                     rh = MID$(LTxtG(x), lclRBnd + 1)             ' Save the rh untouched portion
                     mh = IIF$(w.SrcRepeat + 1 <= LEN(mh), MID$(mh, w.SrcRepeat + 1), "") ' If shift < line length, pick up rest of line
                     mh = mh + STRING$(lclRBnd - Prf.BndLeft + 1 - LEN(mh), " ") ' Lengthen to original length
                     me.LTxtSet(x, lh + mh + rh)                  ' put pieces back together
                     me.AttrScan(x)                               ' Recolorize
                  END IF                                          '

               '----- Unbounded version
               ELSE                                               ' Not BNDS, do the simple version
                  IF w.SrcRepeat + 1 <= LEN(L(x).@LTxt) THEN      ' If shift less than line length
                     me.LTxtSet(x, MID$(LTxtG(x), w.SrcRepeat + 1))' Pick up remaining part of line
                     me.LAttrAdjust(x, Prf.BndLeft, -(w.SrcRepeat))' Adjust any Attr hilite
                  ELSE                                            '
                     me.LTxtSet(x, "")                            ' Else just Null it
                     LAttrS(x) = ""                               ' Adjust any Attr hilite
                  END IF                                          '
                  me.AttrScan(x)                                  ' Recolorize
               END IF                                             '

            '----- )) Right Shift -----------------------------------------------------------------
            CASE ")       ", "))      ", "]       ", "]]      "
               me.ModSet(x)                                       ' Remember we changed something
               w.SrcRepeat = lclRepeat                            '
               IF LEFT$(w.SrcCmd, 1) = ")" THEN                   ' The normal )) shift?
                  w.SrcRepeat = IIF(w.SrcRepeat = 0, ENV.DefDataShift, lclRepeat) ' Default shift if not provided
               ELSE                                               ' Must be the ]] shift
                  w.SrcRepeat = IIF(w.SrcRepeat = 0, ENV.DefDataShift, w.SrcRepeat * ENV.DefDataShift) ' Default shift if not provided
               END IF                                             '
               IF Prf.BndLeft > 1 OR Prf.BndRight > 0 THEN        ' Have to honour BNDS?
                  IF LEN(L(x).@LTxt) < Prf.BndLeft THEN           ' Length < Left Bound
                     ' Null                                       ' Do nothing

                  ELSEIF LEN(L(x).@LTxt) <= lclRBnd THEN          ' Length > Left and not longer than Right
                     lh = LEFT$(LTxtG(x), Prf.BndLeft - 1)        ' Save lh untouched portion
                     rh = MID$(LTxtG(x), Prf.BndLeft)             ' Save the rh shiftable portion
                     rh = STRING$(w.SrcRepeat, " ") + rh          ' Shift it
                     IF Prf.BndRight > 0 THEN rh = LEFT$(rh, lclRBnd - Prf.BndLeft + 1)  ' if not + keep it within the BNDS
                     me.LTxtSet(x, lh + rh)                       ' put pieces back together
                     me.LAttrAdjust(x, Prf.BndLeft, w.SrcRepeat)  ' Adjust any Attr hilite
                     me.AttrScan(x)                               ' Recolorize

                  ELSEIF LEN(L(x).@LTxt) >= lclRBnd THEN          ' Length > Right
                     lh = LEFT$(LTxtG(x), Prf.BndLeft - 1)        ' Save lh untouched portion
                     mh = MID$(LTxtG(x), Prf.BndLeft, lclRBnd - Prf.BndLeft +1) ' Save middle shiftable portion
                     rh = MID$(LTxtG(x), lclRBnd + 1)             ' Save the rh untouched portion
                     mh = STRING$(w.SrcRepeat, " ") + mh          ' Shift it
                     IF Prf.BndRight > 0 THEN mh = LEFT$(mh, lclRBnd - Prf.BndLeft + 1) ' If not + keep it within the BNDS
                     me.LTxtSet(x, lh + mh + rh)                  ' put pieces back together
                     me.LAttrAdjust(x, Prf.BndLeft, w.SrcRepeat)  ' Adjust any Attr hilite
                     me.AttrScan(x)                               ' Recolorize
                  END IF                                          '

               ELSE                                               ' Normal, non-BNDS version
                  me.LTxtSet(x, STRING$(w.SrcRepeat, " ") + LTxtG(x)) ' Shift it
                  me.LAttrAdjust(x, Prf.BndLeft, w.SrcRepeat)     ' Adjust any Attr hilite
                  me.AttrScan(x)                                  ' Recolorize
               END IF                                             '

            '----- U/UU Mark User command ---------------------------------------------------------
            CASE "U       ", "UU      "                           ' Mark User command
               IF IsLData(x) THEN                                 ' Just Data lines
                  me.LFlagBitOn(x, %User)                         ' Mark the line as a USER line
                  OnUndoFlag                                      ' Force a snapshot
               END IF                                             '

            '----- V/VV Revert Mark User command ---------------------------------------------------------
            CASE "V       ", "VV      "                           ' Revert Mark User command
               IF IsLData(x) THEN                                 ' Just Data lines
                  me.LFlagBitOff(x, %User)                        ' Clear USER line
                  OnUndoFlag                                      ' Force a snapshot
               END IF                                             '

         END SELECT                                               '
      END IF                                                      '
   NEXT x                                                         '
   IF shift_error_total > 0 THEN                                  ' Data shift errors?
      scError(%eNone, "Data shifting incomplete on " + FORMAT$(shift_error_total) + " line" + IIF$(shift_error_total = 1, "", "s"))
   END IF
   MExitMeth

DoSCase:
   i = PARSECOUNT(Txt1, " ")                                      ' Count words
   REDIM words(1 TO i) AS STRING                                  ' Make table correct size
   PARSE Txt1, words(), " "                                       ' Parse the words into the table

   m = 1                                                          ' Start scan at 1
   FOR j = 1 TO i                                                 ' Loop for each word
      IF ISNOTNULL(words(j)) THEN                                 ' Ignore null 'words'
         k = INSTR(m, Txt1, words(j))                             ' Find location of this word
         IF caps THEN                                             ' Capitalize this word?
            IF k THEN MID$(Txt1, k, 1) = UUCASE(MID$(Txt1, k, 1)) ' Uppercase the 1st letter
            Caps = %False                                         ' Off till next period
         END IF                                                   '
         IF RIGHT$(words(j), 1) = "."  OR _                       ' Look for sentence trigger
            RIGHT$(words(j), 1) = "?"  OR _                       '
            RIGHT$(words(j), 1) = "!"  THEN                       '
            Caps = %True                                          ' Restart sentence if required
         END IF                                                   '
         m = k + 1                                                ' Start scan adjust
      END IF                                                      '
   NEXT j                                                         '
   RETURN                                                         '

AdjustBnds:
   IF LEN(L(x).@LTxt) > lclLbnd THEN                              ' Any data within the bounds?
      IF MID$(LTxtG(x), lclLBnd, 1) <> " " THEN                   ' Data AT the LBound?
         FOR i = lclLBnd TO lclRBnd - 1                           ' Look for a space (end of 'word')
            IF MID$(LTxtG(x), i, 1) = " " THEN                    ' Found it
               lclLBnd = i + 1: EXIT FOR                          ' Shift Lbnds
            END IF                                                '
         NEXT i                                                   '
      END IF                                                      '
   END IF                                                         '

   IF LEN(L(x).@LTxt) >= lclRBnd THEN                             ' Any data at right boundary?
      IF MID$(LTxtG(x), lclRBnd, 1) <> " " THEN                   ' Data AT the RBound?
         FOR i = lclRBnd TO lclLBnd STEP - 1                      ' Look for a space (start of 'word')
            IF MID$(LTxtG(x), i, 1) = " " THEN                    ' Found it
               lclRBnd = i - 1: EXIT FOR                          ' Shift wRbnds
            END IF                                                '
         NEXT i                                                   '
      END IF                                                      '
   END IF
   RETURN

END METHOD

METHOD  lCmdBNDS(lCmd AS STRING)
'---------- Insert the BNDS line
LOCAL w AS LctlCmd
   MEntry
   me.lCmdParse(w, lCmd)                                          ' Parse the command
   me.InsBounds(w.SrcFrom)                                        ' Go insert it
   me.CurSetReq(%Insert, w.SrcFrom + 1, 1 + Offset, %False)       ' Set cursor set attempt
   MExit
END METHOD

METHOD  lCmdCC(lCmd AS STRING)
'---------- COPY lines in the dataset
LOCAL i, j, x, y, z, n, r, ABi, ABj, NumAdj, SetCur, NumInDest, fMIX AS LONG, w AS lCtlCmd

   MEntry
   me.lCmdParse(w, lCmd)                                          ' Parse the command
   i = w.SrcTo - w.SrcFrom + 1                                    ' Calc # in range
   IF (w.DstCmd = "AA      " OR w.DstCmd = "BB      ") AND _      ' Flag the combo with no repeat support
      w.SrcRepeat > 0 THEN                                        '
      me.CurSetReq(%Position, w.SrcFrom, 0, %True)                ' Set cursor set attempt
      MErrExit(%eFail, "CC / AA or CC / BB cannot use repeat value")
   END IF                                                         '

   '----- Do the simple COPY A/B style
   IF w.DstCmd = "A       " OR w.DstCmd = "B       " THEN         ' The simple Copy command?
      IF w.DstCmd = "B       " THEN w.DstFrom = w.DstFrom - 1     ' Make a'B' into 'A'
      r = MAX(1, w.SrcRepeat, w.DstRepeat)                        ' Get Repeat factor
      me.LInsertEmpty(w.DstFrom, i * r, 0)                        ' Insert enough lines for it all
      me.AdjustPending(w.DstFrom, i * r, i * r)                   ' Adjust pending stuff
      FOR x = 1 TO r                                              ' Do loop'Repeat' times
         z = w.SrcFrom                                            ' Point at 1st line in From Range
         IF w.SrcFrom > w.DstFrom THEN z += (i * r)               ' Adjust if moving upward in the file
         FOR y = 1 TO i                                           ' For each item in the From Range
            n = w.DstFrom + y - 1 + ((x - 1) * i) + 1             ' Calc position to copy the data
               GOSUB CopyZ2N                                      ' Go copy data for one line
            IF w.SrcFrom > w.DstFrom THEN                         ' If moving upward in the file
               IF z < w.SrcTo + (i * r) THEN INCR z               ' Next source line if doing a group replicate
            ELSE
               IF z < w.SrcTo THEN INCR z                         ' Next source line if doing a group replicate
            END IF
            IF y = 1 THEN me.CurSetReq(%LineCmd, n, 0, %False)    ' Set cursor set attempt
         NEXT y                                                   ' Loop inner
      NEXT x                                                      ' Loop Outer

   '----- Do the COPY overlay style
   ELSEIF w.DstCmd = "O       " OR w.DstCmd = "OO      " OR _     ' The Overlay type?
          w.DstCmd = "OR      " OR w.DstCmd = "ORR     " OR w.DstCmd = "OROR     " THEN ' The Overlay type?

      j = w.DstTo - w.DstFrom + 1                                 ' Calc # in range

      me.CurSetReq(%LineCmd, w.DstFrom, 0, %False)                ' Set cursor set attempt

      z = w.SrcFrom                                               ' Point at 1st line in From Range
      FOR x = w.DstFrom TO w.DstTo                                ' Loop through dest lines
         IF IsLData(x) THEN                                       ' Overlay Dest Data lines
            me.lStubDoPM(w.Dstflag, x)                            ' Do + / - processing for the line

            IF IsLData(z) THEN                                    ' and only Source Data lines
               me.lStubDoPM(w.Srcflag, z)                         ' Do + / - processing for the line
               IF w.DstCmd = "OR      " OR w.DstCmd = "ORR     " OR w.DstCmd = "OROR     " THEN ' The Overlay replace type?
                  me.LTxtSet(x, me.OverlayTextRepl(LTxtG(z), LTxtG(x)))  ' Both Data lines, overlay replace them
               ELSE                                               '
                  me.LTxtSet(x, me.OverlayText(LTxtG(z), LTxtG(x))) ' Both Data lines, overlay them
               END IF                                             '
               me.ModSet(x)                                       ' Remember we changed something
               me.LFlagBitOff(x, %EQChange)                       ' Reset ==CHG>
               me.UpdLControl(x)                                  ' And LLCtl
               me.AttrScan(x)                                     ' Recolorize
               IF z < w.SrcTo THEN INCR z ELSE z = w.SrcFrom      ' Next source line (or repeat from the top)
            ELSE                                                  '
               INCR z: DECR x                                     ' Don't get 'stuck' on the non-data line
            END IF                                                '
         END IF                                                   '
      NEXT x                                                      ' Loop Outer

   '----- Do the COPY HERE style
   ELSEIF w.DstCmd = "H       " OR w.DstCmd = "HH      " THEN     ' The Target type?

      me.LStubDelRange(w.DstFrom, w.DstTo)                        ' Delete the HH lines

      IF w.SrcFrom > w.DstFrom THEN                               ' Adjust Source if copying upward in the file
         w.SrcFrom -= (w.DstTo - w.DstFrom + 1)                   '
         w.SrcTo -= (w.DstTo - w.DstFrom + 1)                     '
      END IF                                                      '

      me.AdjustPending(w.DstFrom, 0 - (w.DstTo - w.DstFrom + 1), 0) ' Adjust pending stuff
                                                                  ' Target lines are now deleted
      r = MAX(1, w.SrcRepeat, w.DstRepeat)                        ' Get Repeat factor
      w.DstFrom = w.DstFrom - 1                                   ' Make previous line the 'A' line
      me.LInsertEmpty(w.DstFrom, i * r, 0)                        ' Insert enough lines for it all
      FOR x = 1 TO r                                              ' Do loop 'Repeat' times
         z = w.SrcFrom                                            ' Point at 1st line in From Range
         IF w.SrcFrom > w.DstFrom THEN z += (i * r)               ' Adjust if moving upward in the file
         FOR y = 1 TO i                                           ' For each item in the From Range
            n = w.DstFrom + y - 1 + ((x - 1) * i) + 1             ' Calc position to copy the data
               GOSUB CopyZ2N                                      ' Go copy data for one line
            IF w.SrcFrom > w.DstFrom THEN                         ' If moving upward in the file
               IF z < w.SrcTo + (i * r) THEN INCR z               ' Next source line if doing a group replicate
            ELSE                                                  '
               IF z < w.SrcTo THEN INCR z                         ' Next source line if doing a group replicate
            END IF                                                '
         NEXT y                                                   ' Loop inner
      NEXT x                                                      ' Loop Outer
      me.AdjustPending(w.DstFrom, i * r, i * r)                   ' Adjust pending stuff
      me.CurSetReq(%LineCmd, w.DstFrom + 1, 0, %False)            ' Set cursor set attempt

   '----- Do the COPY AA repeat block style
   ELSEIF w.DstCmd = "AA      " THEN                              ' The multiple Copy After command?

      r = MAX(1, w.SrcRepeat)                                     ' Get Repeat factor
      w.DstRepeat = MAX(1, w.DstRepeat)                           ' Ensure at least count of 1
      FOR x = w.DstFrom TO w.DstTo                                ' Count the data lines
         IF IsLData(x) OR IsLTop(x) THEN                          ' A data line or top line?
            INCR NumInDest                                        ' + 1
            me.lStubDoPM(w.Dstflag, x)                            ' Do + / - processing for the line
         END IF                                                   '
      NEXT x                                                      '

      IF NumInDest <= w.DstRepeat THEN _                          ' Dest count reasonable?
         MErrExit(%eFail, "AA number must be < range size")       '

      ABi = w.DstFrom + w.DstRepeat - 1                           ' Init for loop
      SetCur = w.DstFrom                                          ' Remember where cursor goes

      WHILE ABi <= w.Dstto + NumAdj                               ' Big outer loop
         IF ISFALSE IsLData(ABi) AND ISFALSE IsLTop(ABi) THEN INCR Abi: ITERATE  ' Ignore non Data lines
         IF SetCur > 0 THEN                                       ' Set cursor just once
            me.CurSetReq(%LineCmd, SetCur, 1, %False)             '
            SetCur = 0                                            '
         END IF                                                   '
         me.LInsertEmpty(ABi, i * r, 0)                           ' Insert enough lines for one set copy
         IF ISMEdit THEN                                          ' In MEdit mode?
            fMIX = me.MEditTbl("I", FORMAT$(ABi))                 ' Go get the MIX value of previous line
            FOR x = ABi + 1 TO ABi + (i * r)                      '
               LMixS(x) = fMIX                                    ' Set into the new line
            NEXT k                                                '
         END IF                                                   '
         NumAdj += i * r                                          ' Track # inserted
         IF SetCur > 0 THEN                                       ' Set cursor just once
            me.CurSetReq(%LineCmd, SetCur, 1, %False)             '
            SetCur = 0                                            '
         END IF                                                   '
         n = ABi + 1                                              ' Calc position to copy the data
         FOR x = 1 TO r                                           ' Do loop'Repeat' times
            z = w.SrcFrom                                         ' Point at 1st line in From Range
            IF w.SrcFrom > w.DstFrom THEN z += NumAdj             ' Adjust if moving upward in the file
            FOR y = 1 TO i                                        ' For each item in the From Range
               GOSUB CopyZ2N                                      ' Go copy data for one line
               INCR n                                             ' Bump 'to' pointer
               IF w.SrcFrom > w.DstFrom THEN                      ' If moving upward in the file
                  IF z < w.SrcTo + NumAdj THEN INCR z             ' Next source line if doing a group replicate
               ELSE                                               '
                  IF z < w.SrcTo THEN INCR z                      ' Next source line if doing a group replicate
               END IF                                             '
            NEXT y                                                ' Loop inner
         NEXT x                                                   ' Loop Outer
         ABI += w.DstRepeat + (i * r)                             ' Big outer loop
      LOOP                                                        ' next ABi
      me.AdjustPending(w.DstFrom + w.DstRepeat - 1, NumAdj, 0)    ' Adjust pending stuff

   '----- Do the COPY BB repeat block style
   ELSEIF w.DstCmd = "BB      " THEN                              ' The multiple Copy Before command?

      r = MAX(1, w.SrcRepeat)                                     ' Get Repeat factor
      w.DstRepeat = MAX(1, w.DstRepeat)                           ' Ensure at least count of 1
      FOR x = w.DstFrom TO w.DstTo                                ' Count the data lines
         IF IsLData(x) OR IsLBottom(x) THEN                       ' A data line?
            INCR NumInDest                                        ' + 1
            me.lStubDoPM(w.Dstflag, x)                            ' Do + / - processing for the line
         END IF                                                   '
      NEXT x                                                      '
      IF NumInDest <= w.DstRepeat THEN _                          ' Dest count reasonable?
         MErrExit(%eFail, "BB number must be < range size")       '
      x = NumInDest \ w.DstRepeat                                 ' x = # of insertion points
      ABi = w.DstTo - (x * w.DstRepeat)                           ' Initial insertion point
      SetCur = w.DstFrom                                          ' Remember where cursor goes
      WHILE ABi < w.Dstto + NumAdj                                ' Big outer loop
         IF ISFALSE IsLData(ABi) AND ISFALSE IsLTop(ABi) THEN INCR Abi: ITERATE ' Ignore non Data lines
         IF SetCur > 0 THEN                                       ' Set cursor just once
            me.CurSetReq(%LineCmd, SetCur, 1, %False)             '
            SetCur = 0                                            '
         END IF                                                   '
         me.LInsertEmpty(ABi, i * r, 0)                           ' Insert enough lines for one set copy
         IF IsMEdit THEN                                          ' In MEdit mode?
            fMIX = me.MEditTbl("I", FORMAT$(ABi))                 ' Go get the MIX value of previous line
            FOR x = ABi + 1 TO ABi + (i * r)                      '
               LMixS(x) = fMIX                                    ' Set into the new line
            NEXT k                                                '
         END IF                                                   '
         NumAdj += i * r                                          ' Track # inserted
         IF SetCur > 0 THEN                                       ' Set cursor just once
            me.CurSetReq(%LineCmd, SetCur, 1, %False)             '
            SetCur = 0                                            '
         END IF                                                   '
         n = ABi + 1                                              ' Calc position to copy the data
         FOR x = 1 TO r                                           ' Do loop'Repeat' times
            z = w.SrcFrom                                         ' Point at 1st line in From Range
            IF w.SrcFrom > w.DstFrom THEN z += NumAdj             ' Adjust if moving upward in the file
            FOR y = 1 TO i                                        ' For each item in the From Range
               GOSUB CopyZ2N                                      ' Go copy data for one line
               INCR n                                             ' Bump 'to' pointer
               IF w.SrcFrom > w.DstFrom THEN                      ' If moving upward in the file
                  IF z < w.SrcTo + NumAdj THEN INCR z             ' Next source line if doing a group replicate
               ELSE                                               '
                  IF z < w.SrcTo THEN INCR z                      ' Next source line if doing a group replicate
               END IF                                             '
            NEXT y                                                ' Loop inner
         NEXT x                                                   ' Loop Outer
         ABI += w.DstRepeat + (i * r)                             ' Big outer loop
      LOOP                                                        ' next ABi
      me.AdjustPending(w.DstFrom + w.DstRepeat - 1, NumAdj, 0)    ' Adjust pending stuff
   END IF                                                         '
   gfXRebuild = %True                                             ' Force exclude rebuild
   MExitMeth

'----- Copy a line indexed by Z to line indexed by N
CopyZ2N:
   me.LFlagBitOff(n, %EQChange)                                   ' Reset ==CHG>
   LLblS(n) = $BlankLNo                                           ' Copy the entry control data
   LTagS(n) = TRIM$(LTagG(z))                                     '
   LLCtlS(n) = $BlankLNo                                          '
   LWrk1S(n) = LWrk1G(z)                                          '
   me.LTxtSet(n, LTxtG(z))                                        '
   IF ISTRUE (LFlagG(z) AND (%Tabs OR %Word OR %Mark OR %Mask OR %Bounds)) THEN ' Free the dynamic string for special lines
      me.LTxtFree(n)                                              '
      IF IsLWord(z) THEN me.LTxt2Word(n)                          '
      IF IsLMark(z) THEN me.LTxt2Mark(n)                          '
      IF IsLMask(z) THEN me.LTxt2Mask(n)                          '
      IF IsLTabs(z) THEN me.LTxt2Tabs(n)                          '
      IF IsLBounds(z) THEN me.LTxt2Bnds(n)                        '
   END IF                                                         '
   LAttrS(n) = LAttrG(z)                                          ' Copy it over
   LFlagS(n) = LFlagG(z)                                          ' Copy Flag last to sTxtFree doesn't see it
   me.LFlagBitOff(n, %Cursor)                                     ' And never copy the %Cursor flag
   me.LFlagBitOff(n, %Scroll)                                     ' or the scroll
   me.LFlagBitOff(n, %InsertLine)                                 ' or the Insert status
   me.UpdLControl(n)                                              '
   me.lStubDoPM(w.Srcflag, z)                                     ' Do + / - processing for the source
   me.lStubDoPM(w.Dstflag, n)                                     ' Do + / - processing for the destination
   me.ModSet(n)                                                   ' Remember we changed something
   RETURN                                                         '
END METHOD

METHOD  lCmdCOLS(lCmd AS STRING)
'---------- Insert a COLS line(s) into the dataset
LOCAL w AS LCtlCmd
   MEntry
   me.lCmdParse(w, lCmd)                                          ' Parse the command
   me.InsCols(w.SrcFrom)                                          ' Insert the COLS line
   me.CurSetReq(%Insert, w.SrcFrom + 1, 0, %False)                ' Set cursor set attempt
   MExit
END METHOD

METHOD  lCmdDD(lCmd AS STRING)
'---------- Delete line(s) in the Dataset
LOCAL i, j, x AS LONG, w AS lCtlCmd

   MEntry
   me.lCmdParse(w, lCmd)                                          ' Parse the command
   '----- If a DELETE =FILE> line, make all the adjustments
   IF w.SrcTo = w.SrcFrom AND IsLFile(w.SrcFrom) THEN             ' Is this the whole =FILE> delete?
      FOR x = w.SrcFrom + 1 TO LastLine                           ' Adjust w.SrcTo to the end of the 'file'
         IF LMIXG(x) = LMIXG(w.SrcFrom) THEN                      ' Still the same 'file'
            w.SrcTo = x                                           ' Set new 'to' line
         ELSE                                                     ' Oops, end of 'file' or Bottom
            EXIT FOR                                              ' Scan no more
         END IF                                                   '
      NEXT x                                                      '
      me.MeditTbl("D", TRIM$(LTxtG(w.SrcFrom)))                   ' Remove from Medit table
      me.UnWatchQueue(TRIM$(LTxtG(w.SrcFrom)))                    ' Kill Watch and dequeue
      FOR x = w.SrcTo TO LastLine                                 ' Ripple down the Mix entries above ours
         IF IsLData(x) OR IsLFile(x) OR IsLNote(x) THEN           '
            IF LMIXG(x) > LMIXG(w.SrcFrom) THEN _                 '
               LMixS(x) = LMIXG(x) - 1                            '
         END IF                                                   '
      NEXT x                                                      '
   END IF                                                         '

   '----- Now do the actual deletes
   i = w.SrcTo - w.SrcFrom + 1                                    ' Calc # in range
   me.ModSet(w.SrcFrom)                                           ' Remember we changed something
   me.LStubDelRange(w.SrcFrom, w.SrcTo)                           ' Delete the lines

   '----- Adjust everything for the deleted lines
   me.AdjustPending(w.SrcFrom, 0 - i, 0)                          ' Adjust pending stuff
   me.CurSetReq(%LineCmd, w.SrcFrom, 0, %True)                    ' Set cursor set attempt
   gfXRebuild = %True                                             ' Force exclude rebuild
   MExit
END METHOD

METHOD  lCmdF(lCmd AS STRING)
'---------- Re-Display First 'n' lines of an eXcluded range
LOCAL i, j AS LONG, w AS lCtlCmd
   MEntry
   me.lCmdParse(w, lCmd)                                          ' Parse the command
   IF ISFALSE IsLXclude(w.SrcFrom) THEN _                         ' Better be on an eXcluded line
      MErrExit(%eFail, "FIRST allowed on EXCLUDED lines only")    '

   '----- Get set up
   i = MAX(1, w.SrcRepeat)                                        ' Get number to be 'exploded'
   i = MIN(i, LastLine - w.SrcFrom)                               ' Prevent going past LastLine

   '----- If entire X's range, do it
   IF i >= LWrk1G(w.SrcFrom) THEN                                 ' Is this the whole range?
      i = LWrk1G(w.SrcFrom)                                       ' Yes, set i to range size
      me.LTxtFree(w.SrcFrom)                                      ' Go free the dynamic string of the XX line
      me.LEntDel(w.SrcFrom)                                       ' Remove from the L() array
      DECR LastLine: DECR LastReal                                ' Adjust LastLine
      FOR j = w.SrcFrom TO w.SrcFrom + i - 1                      ' Un-'hide' the lines in the range
         me.LFlagBitOff(j, %Invisible)                            '
         LWrk1S(j) = 0                                            ' Reset work count
      NEXT j                                                      '
      me.AdjustPending(j, -1, 0)                                  ' Adjust pending stuff

   '----- Show part of an X'd range
   ELSE
      LWrk1S(w.SrcFrom) = LWrk1G(w.SrcFrom) - i                   ' Adjust # in reduced range
      me.LEntXPop(w.SrcFrom, i)                                   ' Go un-POP the lines
   END IF                                                         '
   me.CurSetReq(%LineCmd, w.SrcFrom, 1 + Offset, %False)          ' Set cursor set attempt
   gfXRebuild = %True                                             ' Force exclude rebuild
   MExit
END METHOD

METHOD  lCmdI(lCmd AS STRING)
'---------- Insert line(s) into the dataset
LOCAL i, j, k, fMIX AS LONG, w AS lCtlCmd

   MEntry
   me.lCmdParse(w, lCmd)                                          '
   j = w.SrcTo                                                    ' Get insert point
   me.ModSet(j)                                                   ' Remember we changed something
   i = MAX(1, w.SrcRepeat)                                        ' Get number to be inserted

   '----- Insert either I type ('''''''') or N type (nnnnnnnn) lines
   IF w.SrcCmd = "I       " THEN                                  ' If a normal Insert
      me.LInsertEmpty(j, i, %InsertLine OR %Data)                 ' Insert the requested line(s)
      FOR k = j + 1 TO J + i                                      ' Add to Ins Table
         me.InsTblAdd(k)                                          '
         me.LTxtSet(k, Prf.MaskLine)                              ' Stuff in the mask
      NEXT k                                                      '
   ELSE                                                           ' Else insert as normal data lines
      me.LInsertEmpty(j, i, %Data)                                '
      FOR k = j + 1 TO J + i                                      '
         me.LTxtSet(k, "")                                        ' Stuff in a null
      NEXT k                                                      '
   END IF                                                         '

   '----- If MEdit mode ensure the MIX value is correct
   IF IsMEdit THEN                                                ' In MEdit mode?
      fMIX = me.MEditTbl("I", FORMAT$(j))                         ' Go get the MIX value of previous line
      FOR k = j + 1 TO j + i                                      '
         LMixS(k) = fMIX                                          ' Set into the new line
      NEXT k                                                      '
   END IF                                                         '

   '----- Adjust now for the inserted lines
   me.AdjustPending(j, i, 0)                                      ' Adjust pending stuff
   me.CurSetReq(%Insert, j + 1, me.IndentAsPrev(j + 1), %True)    ' Set cursor set attempt
   OnInsClnSupp                                                   ' Suppress Insert cleanup for 1 Attn pass
   MExit
END METHOD

METHOD  lCmdJJ(lCmd AS STRING)
'---------- Join tagged line with the next line
LOCAL i, j, k, x AS LONG, sclr AS STRING, w AS lCtlCmd
   MEntry
   me.lCmdParse(w, lCmd)                                          ' Parse the command
   i = w.SrcTo - w.SrcFrom                                        ' Calc # in range
   IF w.SrcFrom = w.SrcTo THEN                                    ' A single x(nn) command?
      i = MAX(1, w.SrcRepeat - 1)                                 ' Get requested # lines
   END IF                                                         '

   '----- Validate stuff
   FOR x = w.SrcFrom TO w.SrcFrom + i                             ' Check for X'd stuff
      IF ISFALSE IsLData(x) OR IsLInvisible(x) OR IsLFile(x) THEN _ ' All normal data and visible?
         MErrExit(%eFail, "Join/Glue range must be only normal Data lines")
   NEXT x                                                         '
   IF w.SrcFrom => LastLine - 1 THEN _                            ' Better be a next line
      MErrExit(%eFail, "No next line to join with")               '

   '----- OK, let's do it
   j = w.SrcFrom: k = j + 1: x = 0                                ' Point at 1st line and next line
   me.ModSet(w.SrcFrom)                                           ' Remember we changed something
   DO WHILE i > 0                                                 ' Loop for requested count
      IF w.SrcCmd = "J       " OR w.SrcCmd = "JJ      " THEN      ' Join?
         me.LTxtSet(j, LTxtG(j) + " " + LTxtG(k))                 ' Do the Join

         LAttrS(j) = LAttrG(j) + CHR$$(0) + LAttrG(k)             ' Set up
      ELSE                                                        ' Else must be Glue
         me.LTxtSet(j, LTxtG(j) + ENV.GlueWith + lTxtG(k))        ' Do the Glue
         LAttrS(j) = LAttrG(j) + REPEAT$(LEN(ENV.GlueWith), CHR$$(0)) + LAttrG(k) ' Both present
      END IF                                                      '
      me.AttrScan(j)                                              ' Recolorize
      me.LTxtFree(k)                                              ' Go free the dynamic string
      me.LEntDel(k)                                               ' and remove the Joined/Glued line
      DECR i: INCR x                                              ' Loop control
   LOOP                                                           '

   '----- Adjust now for the joined lines
   me.AdjustPending(w.SrcFrom, 0 - x, 0)                          ' Adjust pending stuff
   me.CurSetReq(%LineCmd, w.SrcFrom, 0, %True)                    ' Set cursor set attempt
   LastLine -= x: LastReal -= x                                   ' Adjust LastLine
   IF (w.SrcCmd = "G       " OR w.SrcCmd = "GG      ") AND ISNOTNULL(ENV.GlueWith) THEN _ ' Glue with non-null ENV.GlueWith?
      scError(%eNone, "Glued with " + $DQ + ENV.GlueWith + $DQ)
   MExit
END METHOD

METHOD  lCmdL(lCmd AS STRING)
'---------- Re-Display Last 'n' lines of an eXcluded range
LOCAL i, j AS LONG, w AS lCtlCmd

   MEntry
   me.lCmdParse(w, lCmd)                                          ' Parse the command
   IF ISFALSE IsLXclude(w.SrcFrom) THEN _                         ' Better be on an eXcluded line
      MErrExit(%eFail, "LAST allowed on EXCLUDED lines only")     '
   i = MAX(1, w.SrcRepeat)                                        ' Get number to be 'exploded'
   i = MIN(i, LastLine - w.SrcFrom)                               ' Prevent going past LastLine

   '----- An entire range uncovered
   IF i >= LWrk1G(w.SrcFrom) THEN                                 ' Is this the whole range?
      i = LWrk1G(w.SrcFrom)                                       ' Yes, set i to range size
      me.LTxtFree(w.SrcFrom)                                      ' Go free the dynamic string of the XX line
      me.LEntDel(w.SrcFrom)                                       ' Remove from the L() array
      DECR LastLine: DECR LastReal                                ' Adjust LastLine
      me.CurSetReq(%LineCmd, w.SrcFrom, 1 + Offset, %False)       ' Set cursor set attempt
      FOR j = w.SrcFrom TO w.SrcFrom + i - 1                      ' Un-'hide' the lines in the range
         me.LFlagBitOff(j, %Invisible)                            '
         LWrk1S(j) = 0                                            ' Reset work counter
      NEXT j                                                      '
      me.AdjustPending(j, -1, 0)                                  ' Adjust pending stuff

   '----- Just part of a range to expose
   ELSE
      me.CurSetReq(%LineCmd, w.SrcFrom + LWrk1G(w.SrcFrom) - i + 1, 1 + Offset, %False) ' Set cursor set attempt
      FOR j = w.SrcFrom + LWrk1G(w.SrcFrom) - i + 1 TO w.SrcFrom + LWrk1G(w.SrcFrom) ' Un-'hide' the lines in the range
         me.LFlagBitOff(j, %Invisible)                            '
         LWrk1S(j) = 0                                            '
      NEXT j                                                      '
      LWrk1S(w.SrcFrom) = LWrk1G(w.SrcFrom) - i                   ' Adjust # in reduced range
   END IF                                                         '
   gfXRebuild = %True                                             ' Force exclude rebuild
   MExit
END METHOD

METHOD  lCmdMakeDN(lCmd AS STRING, lType AS LONG)
'---------- Change lines to NOTEs
LOCAL i, j, x AS LONG, t AS STRING, w AS lCtlCmd

   MEntry
   me.lCmdParse(w, lCmd)                                          ' Parse the command
   me.CurSetReq(%LineCmd, w.SrcFrom, 1 + Offset, %False)          ' Set cursor set attempt

   FOR x = w.SrcFrom TO w.SrcTo                                   ' Loop for each line

      '----- Leaving NOTE status
      IF IsLNote(x) THEN                                          ' Note line?
         me.ModSet(x)                                             ' Remember we changed something
         me.LFlagBitOff(x, %Note)                                 ' Remove %Note
         me.LFlagBitOn(x, lType)                                  ' Add desired type
         GOSUB MarkData                                           ' See if moving to data

      '----- Leaving Normal Text type
      ELSEIF IsLData(x) THEN                                      ' Data line?
         me.LFlagBitOff(x, %Data)                                 ' Remove %Data
         me.LFlagBitOn(x, lType)                                  ' Add desired type
         IF lType <> %Data THEN                                   ' If losing Data
            me.ModSet(x)                                          ' Flag modified
            LLblS(x) = $BlankLNo                                  ' Clear any Labels/Tags
            LTagS(x) = $BlankLNo                                  '
            me.UpdLControl(x)                                     ' Refresh LIne number area
            LastReal -= 1                                         ' Update Real count
         END IF                                                   '

      '----- Leaving BNDS status
      ELSEIF IsLBounds(x) THEN                                    ' BNDS line?
         me.LFlagBitOff(x, %Bounds)                               ' Remove %Bounds
         me.LFlagBitOn(x, lType)                                  ' Add desired type
         me.LAllocTxt(x)                                          ' Allocate a text string
         me.LTxtSet(x, Prf.BndText)                               ' Copy the bounds string
         GOSUB MarkData                                           ' See if moving to data

      '----- Leaving TABS ststus
      ELSEIF IsLTabs(x) THEN                                      ' TABS line?
         me.LFlagBitOff(x, %Tabs)                                 ' Remove %Tabs
         me.LFlagBitOn(x, lType)                                  ' Add desired type
         me.LAllocTxt(x)                                          ' Allocate a text string
         me.LTxtSet(x, Prf.TabsLine)                                 ' Copy the Tabs string
         GOSUB MarkData                                           ' See if moving to data

      '----- Leaving MARK ststus
      ELSEIF IsLMark(x) THEN                                      ' MARK line?
         me.LFlagBitOff(x, %Mark)                                 ' Remove %Mark
         me.LFlagBitOn(x, lType)                                  ' Add desired type
         me.LAllocTxt(x)                                          ' Allocate a text string
         me.LTxtSet(x, Prf.MarkLine)                              ' Copy the Mark string
         GOSUB MarkData                                           ' See if moving to data

      '----- Leaving MASK ststus
      ELSEIF IsLMask(x) THEN                                      ' MASK line?
         me.LFlagBitOff(x, %Mask)                                 ' Remove %Mask
         me.LFlagBitOn(x, lType)                                  ' Add desired type
         me.LAllocTxt(x)                                          ' Allocate a text string
         me.LTxtSet(x, Prf.MaskLine)                              ' Copy the Mask string
         GOSUB MarkData                                           ' See if moving to data

      '----- Leaving WORD status
      ELSEIF IsLWord(x) THEN                                      ' WORD line?
         me.LFlagBitOff(x, %Word)                                 ' Remove %WORD
         me.LFlagBitOn(x, lType)                                  ' Add desired type
         me.LAllocTxt(x)                                          ' Allocate a text string
         me.LTxtSet(x, Prf.WordInput)                             ' Copy the WORD string
         GOSUB MarkData                                           ' See if moving to data

      '----- Leaving COLS ststus
      ELSEIF IsLCols(x) THEN                                      ' COLS line?
         me.LFlagBitOff(x, %Cols)                                 ' Remove %Cols
         me.LFlagBitOn(x, lType)                                  ' Add desired type
         me.LAllocTxt(x)                                          ' Allocate a text string
         t = me.WindowCol                                         ' Get COLS line
         FOR i = 1 TO LEN(t)                                      ' Purify it
            IF ASC(MID$(t, i, 1)) > 127 THEN _                    ' Reduce all > 128 to under 128
               MID$(t, i, 1) = CHR$(ASC(MID$(t, i, 1)) - 128)     '
         NEXT i
         me.LTxtSet(x, t)                                         ' Copy the COLS data
         GOSUB MarkData                                           ' See if moving to data

      '----- Leaving PROF ststus
      ELSEIF IsLProf(x) THEN                                      ' PROF line?
         t = LTxtG(x)                                             ' Get the text
         me.LFlagBitOff(x, %Prof)                                 ' Remove %Prof
         me.LFlagBitOn(x, lType)                                  ' Add desired type
         me.LAllocTxt(x)                                          ' Allocate a text string
         me.LTxtSet(x, t)                                         ' Copy the Prof string
         GOSUB MarkData                                           ' See if moving to data
      END IF                                                      '
   NEXT x
   MExitMeth

MarkData:
   IF lType = %Data THEN                                          ' If making Data
      me.ModSet(x)                                                ' Flag modified
      me.UpdLControl(x)                                           ' Refresh LIne number area
      me.AttrScan(x)                                              ' Recolorize
      LastReal += 1                                               ' Update Real count
      OnRenumFlag                                                 '
   END IF                                                         '
   RETURN

END METHOD

METHOD  lCmdMark(lCmd AS STRING)
'---------- Insert the Mark line
LOCAL w AS lCtlCmd
   MEntry
   me.lCmdParse(w, lCmd)                                          ' Parse the command
   me.InsMark(w.SrcFrom)                                          ' Insert the MARK line
   me.CurSetReq(%Insert, w.SrcFrom + 1, 1 + Offset, %False)       ' Set cursor set attempt
   MExit
END METHOD

METHOD  lCmdMask(lCmd AS STRING)
'---------- Insert the Mask line
LOCAL w AS lCtlCmd
   MEntry
   me.lCmdParse(w, lCmd)                                          ' Parse the command
   me.InsMask(w.SrcFrom)                                          ' Insert the MASK line
   me.CurSetReq(%Insert, w.SrcFrom + 1, 1 + Offset, %False)       ' Set cursor set attempt
   MExit
END METHOD

METHOD  lCmdMM(lCmd AS STRING)
'---------- Move lines within the dataset
LOCAL i, j, x, y, z, n, r, ABi, SetCur, NumAdj, NumInDest, fMIX AS LONG
LOCAL S1Start, S1End, S2Start, S2End AS LONG
LOCAL OverlayError, LblOnce AS INTEGER, w AS lCtlCmd

   MEntry
   me.lCmdParse(w, lCmd)                                          ' Parse the command
   OverlayError = %False                                          ' Start as 'No error'
   IF (w.DstCmd = "W       " OR w.DstCmd = "WW      ") AND _      ' Flag the combo with no repeat support
      w.SrcRepeat > 0 THEN                                        '
      me.CurSetReq(%Position, w.SrcFrom, 0, %True)                ' Set cursor set attempt
      MErrExit(%eFail, "MM / WW cannot use repeat value")         '
   END IF                                                         '

   IF (w.DstCmd = "AA      " OR w.DstCmd = "BB      ") AND _      ' Flag the combo with no repeat support
      w.SrcRepeat > 0 THEN                                        '
      me.CurSetReq(%Position, w.SrcFrom, 0, %True)                ' Set cursor set attempt
      MerrExit(%eFail, "MM / AA or MM / BB cannot use repeat value")
   END IF                                                         '

   i = w.SrcTo - w.SrcFrom + 1                                    ' Calc # in range

   '----- Simple move after type
   IF w.DstCmd = "A       " OR w.DstCmd = "B       " THEN         ' Is this the simple Move command?

      r = MAX(1, w.SrcRepeat, w.DstRepeat)                        ' Get Repeat factor
      IF w.DstCmd = "B       " THEN w.DstFrom = w.DstFrom - 1     ' Make a'B' into 'A'
      IF IsLXclude(w.DstFrom) THEN _                              ' If XX range, add # XX'd to reach end of range
         w.DstFrom += LWrk1G(w.DstFrom)                           '
      me.LInsertEmpty(w.DstFrom, i * r, 0)                        ' Insert enough lines for it all

      IF IsMEdit THEN                                             ' In MEdit mode?
         fMIX = me.MEditTbl("I", FORMAT$(w.DstFrom))              ' Go get the MIX value of previous line
         FOR x = w.DstFrom + 1 TO w.DstFrom + (i * r)             '
            LMixS(x) = fMIX                                       ' Set into the new line
         NEXT k                                                   '
      END IF                                                      '

      LblOnce = %True                                             ' Only move labels once
      FOR x = 1 TO r                                              ' Do loop'Repeat' times
         z = w.SrcFrom                                            ' Point at 1st line in From Range
         IF w.SrcFrom > w.DstFrom THEN z += (i * r)               ' Adjust if moving upward in the file
         FOR y = 1 TO i                                           ' For each item in the From Range
            n = w.DstFrom + y - 1 + ((x - 1) * i) + 1             ' Calc position to copy the data
            GOSUB CopyZ2N                                         ' Go copy data for one line
            IF w.SrcFrom > w.DstFrom THEN                         ' If moving upward in the file
               IF z < w.SrcTo + (i * r) THEN INCR z               ' Next source line if doing a group replicate
            ELSE                                                  '
               IF z < w.SrcTo THEN INCR z                         ' Next source line if doing a group replicate
            END IF                                                '
         NEXT y                                                   ' Loop inner
         LblOnce = %False                                         ' Flip labels once switch
      NEXT x                                                      ' Loop Outer

      me.AdjustPending(w.DstFrom, i * r, i * r)                   ' Adjust pending stuff
      me.CurSetReq(%LineCmd, w.DstFrom + 1, 0, %False)            ' Set cursor set attempt

      IF w.SrcFrom > w.DstFrom THEN w.SrcFrom = w.SrcFrom + (i * r) ' Adjust if we moved upward in the file
      me.LStubDelRange(w.SrcFrom, w.SrcFrom + i - 1)              ' Delete the lines
      me.AdjustPending(w.SrcFrom, 0 - i, i * r)                   ' Adjust pending stuff

   '----- Move Overlay type
   ELSEIF w.DstCmd = "O       " OR w.DstCmd = "OO      " OR _     ' The Overlay style?
          w.DstCmd = "OR      " OR w.DstCmd = "ORR     " OR w.DstCmd = "OROR    " THEN

      IF w.DstFrom = w.DstTo THEN                                 ' A single O(nn) command?
         i = MAX(1, w.DstRepeat)                                  ' Get requested # lines
         w.DstRepeat = 0                                          ' No longer a repeat, used as a count
         i = MIN(i, LastLine - w.DstFrom)                         ' Prevent going past LastLine
         j = w.DstFrom                                            ' Point at 1st line
         DO WHILE i > 0                                           ' Loop for requested count
            IF IsLXclude(j) THEN                                  ' If an XX 'd range, add its size
               j += LWrk1G(j)                                     '
            END IF                                                '
            DECR i: INCR j                                        ' Loop control
         LOOP                                                     '
         w.DstTo = j - 1                                          ' Stuff back adjusted 'To' line number
         IF w.DstTo > LastLine - 1 THEN w.DstTo = LastLine - 1    '
      END IF                                                      '
      me.CurSetReq(%LineCmd, w.DstFrom, 1 + Offset, %False)       ' Set cursor set attempt
      j = w.DstTo - w.DstFrom + 1                                 ' Calc # in range

      me.ModSet(w.SrcFrom)                                        ' Remember we changed something
      me.ModSet(w.DstFrom)                                        '
      z = w.SrcFrom                                               ' Point at 1st line in From Range
      FOR x = w.DstFrom TO w.DstTo                                ' Loop through dest lines

         IF IsLData(x) THEN                                       ' Overlay Dest Data lines

            IF IsLData(z) THEN                                    ' and only Source Data lines

               me.lStubDoPM(w.Dstflag, x)                         ' Do + / - processing for the line

               IF w.DstCmd = "OR      " OR w.DstCmd = "ORR     " OR w.DstCmd = "OROR    " THEN
                  me.LTxtSet(x, me.OverlayTextRepl(LTxtG(z), LTxtG(x)))  ' Both Data lines, overlay replace them
               ELSE                                               '
                  me.LTxtSet(x, me.OverlayTextDel(LTxtG(z), LTxtG(x)))  ' Both Data lines, overlay them
                  IF IsOvrTextDel THEN OverlayError = %True       ' Remember if not all chars 'swallowed up'
               END IF                                             '
               me.AttrScan(x)                                     ' Recolorize

               IF z < w.SrcTo THEN INCR z ELSE z = w.SrcFrom      ' Next source line (or repeat from the top)

            ELSE                                                  '
               INCR z: DECR x                                     ' Don't get 'stuck' on the non-data line
            END IF                                                '
         END IF                                                   '
      NEXT x                                                      ' Loop Outer

      IF ISFALSE OverlayError THEN                                ' If all data 'swallowed'
         i = w.SrcTo - w.SrcFrom + 1                              ' Re-Calc # in source range
         r = w.Dstto - w.DstFrom + 1                              ' And in Dest range

         IF i > r THEN                                            ' Source > Dest?
            i = r                                                 ' Pick smaller of them
            scError(%eFail, "Not all data moved")
         END IF                                                   '

         me.LStubDelRange(w.SrcFrom, w.SrcFrom + i - 1)           ' Delete the lines

         me.AdjustPending(w.SrcFrom, 0 - i, 0)                    ' Adjust pending stuff
      ELSE                                                        '
         scError(%eFail, "MOVE data not deleted - all data not overlaid")
      END IF                                                      '

   '----- Simple move 'HERE' type
   ELSEIF w.DstCmd = "H       " OR w.DstCmd = "HH      " THEN     ' The Target style?

      IF IsLXclude(w.DstTo) THEN _                                ' If ends in an XX range, add # XX'd to reach end of range
         w.DstTo += LWrk1G(w.DstTo)                               '

      me.LStubDelRange(w.DstFrom, w.DstTo)                        ' Delete the lines

      IF w.SrcFrom > w.DstFrom THEN                               ' Adjust Source if copying upward in the file
         w.SrcFrom -= (w.DstTo - w.DstFrom + 1)                   '
         w.SrcTo -= (w.DstTo - w.DstFrom + 1)                     '
      END IF                                                      '

      me.AdjustPending(w.DstFrom, 0 - (w.DstTo - w.DstFrom + 1), 0) ' Adjust pending stuff

      r = MAX(1, w.SrcRepeat, w.DstRepeat)                        ' Get Repeat factor
      w.DstFrom = w.DstFrom - 1                                   ' Set Dest as if previous line was an 'A'

      IF IsLXclude(w.DstFrom) THEN _                              ' If XX range, add # XX'd to reach end of range
         w.DstFrom += LWrk1G(w.DstFrom)                           '
      me.LInsertEmpty(w.DstFrom, i * r, 0)                        ' Insert enough lines for it all

      IF IsMEdit THEN                                             ' In MEdit mode?
         fMIX = me.MEditTbl("I", FORMAT$(w.DstFrom))              ' Go get the MIX value of previous line
         FOR x = w.DstFrom + 1 TO w.DstFrom + (i * r)             '
            LMixS(x) = fMIX                                       ' Set into the new line
         NEXT k                                                   '
      END IF                                                      '

      LblOnce = %True                                             ' Only move labels once
      FOR x = 1 TO r                                              ' Do loop'Repeat' times
         z = w.SrcFrom                                            ' Point at 1st line in From Range
         IF w.SrcFrom > w.DstFrom THEN z += (i * r)               ' Adjust if moving upward in the file
         FOR y = 1 TO i                                           ' For each item in the From Range
            n = w.DstFrom + y - 1 + ((x - 1) * i) + 1             ' Calc position to copy the data
            GOSUB CopyZ2N                                         ' Go copy data for one line
            IF w.SrcFrom > w.DstFrom THEN                         ' If moving upward in the file
               IF z < w.SrcTo + (i * r) THEN INCR z               ' Next source line if doing a group replicate
            ELSE                                                  '
               IF z < w.SrcTo THEN INCR z                         ' Next source line if doing a group replicate
            END IF                                                '
         NEXT y                                                   ' Loop inner
         LblOnce = %False                                         ' Flip labels once switch
      NEXT x                                                      ' Loop Outer

      me.AdjustPending(w.DstFrom, i * r, i * r)                   ' Adjust pending stuff
      me.CurSetReq(%LineCmd, w.DstFrom + 1, 0, %False)            ' Set cursor set attempt

      IF w.SrcFrom > w.DstFrom THEN w.SrcFrom = w.SrcFrom + (i * r) ' Adjust if we moved upward in the file

      me.LStubDelRange(w.SrcFrom, w.SrcFrom + i - 1)              ' Delete the lines

      me.AdjustPending(w.SrcFrom, 0 - i, i * r)                   ' Adjust pending stuff

   '----- Move after AA (interspersed)
   ELSEIF w.DstCmd = "AA      " THEN                              ' The multiple Move After command?

      r = MAX(1, w.SrcRepeat)                                     ' Get Repeat factor
      IF IsLXclude(w.DstTo) THEN _                                ' If Dest ends in XX range, add # XX'd to reach real end
         w.DstTo += LWrk1G(w.DstTo)                               '
      w.DstRepeat = MAX(1, w.DstRepeat)                           ' Ensure at least count of 1

      FOR x = w.DstFrom TO w.DstTo                                ' Count the data lines
         IF IsLData(x) OR IsLTop(x) THEN                          '
            INCR NumInDest                                        ' + 1
            me.lStubDoPM(w.Dstflag, x)                            ' Do + / - processing for the line
         END IF                                                   '
      NEXT x                                                      '
      IF NumInDest <= w.DstRepeat THEN _                          ' Dest count reasonable?
         MErrExit(%eFail, "AA number must be < range size")       '

      ABi = w.DstFrom + w.DstRepeat - 1                           ' Init for loop
      SetCur = w.DstFrom                                          ' Remember where cursor goes
      WHILE ABi <= w.Dstto + NumAdj                               ' Big outer loop
         IF ISFALSE IsLData(ABi) AND ISFALSE IsLTop(ABi) THEN INCR Abi: ITERATE ' Ignore non Data lines
         IF SetCur > 0 THEN                                       ' Set cursor just once
            me.CurSetReq(%LineCmd, SetCur, 1, %False)             '
            SetCur = 0                                            '
         END IF                                                   '

         me.LInsertEmpty(ABi, i * r, 0)                           ' Insert enough lines for one set copy
         IF IsMEdit THEN                                          ' In MEdit mode?
            fMIX = me.MEditTbl("I", FORMAT$(ABi))                 ' Go get the MIX value of previous line
            FOR x = ABi + 1 TO ABi + (i * r)                      '
               LMixS(x) = fMIX                                    ' Set into the new line
            NEXT k                                                '
         END IF                                                   '

         NumAdj += i * r                                          ' Track # inserted

         n = ABi + 1                                              ' Calc position to copy the data
         FOR x = 1 TO r                                           ' Do loop 'Repeat' times
            z = w.SrcFrom                                         ' Point at 1st line in From Range
            IF w.SrcFrom > w.DstFrom THEN z += NumAdj             ' Adjust if moving upward in the file

            FOR y = 1 TO i                                        ' For each item in the From Range
               GOSUB CopyZ2N                                      ' Go copy data for one line
               INCR n                                             ' Bump 'to' pointer
               IF w.SrcFrom > w.DstFrom THEN                      ' If moving upward in the file
                  IF z < w.SrcTo + NumAdj THEN INCR z             ' Next source line if doing a group replicate
               ELSE                                               '
                  IF z < w.SrcTo THEN INCR z                      ' Next source line if doing a group replicate
               END IF                                             '
            NEXT y                                                ' Loop inner
         NEXT x                                                   ' Loop Outer
         ABI += w.DstRepeat + (i * r)                             ' Big outer loop
      LOOP                                                        ' next ABi


      me.AdjustPending(w.DstFrom + w.DstRepeat - 1, NumAdj, 0)    ' Adjust pending stuff

      IF w.SrcFrom > w.DstFrom THEN w.SrcFrom = w.SrcFrom + NumAdj' Adjust if we moved upward in the file

      me.LStubDelRange(w.SrcFrom, w.SrcFrom + i - 1)              ' Delete the lines

      me.AdjustPending(w.SrcFrom, 0 - i, i * r)                   ' Adjust pending stuff

   '----- Move after BB (interspersed)
   ELSEIF w.DstCmd = "BB      " THEN                              ' The multiple Move Before command?

      r = MAX(1, w.SrcRepeat)                                     ' Get Repeat factor
      IF IsLXclude(w.DstTo) THEN _                                ' If Dest ends in XX range, add # XX'd to reach real end
         w.DstTo += LWrk1G(w.DstTo)                               '
      w.DstRepeat = MAX(1, w.DstRepeat)                           ' Ensure at least count of 1

      FOR x = w.DstFrom TO w.DstTo                                ' Count the data lines
         IF IsLData(x) OR IsLBottom(x) THEN                       '
            INCR NumInDest                                        ' + 1
            me.lStubDoPM(w.Dstflag, x)                            ' Do + / - processing for the line
         END IF                                                   '
      NEXT x                                                      '
      IF NumInDest <= w.DstRepeat THEN _                          ' Dest count reasonable?
         MErrExit(%eFail, "BB number must be < range size")       '

      x = NumInDest \ w.DstRepeat                                 ' x = # of insertion points
      ABi = w.DstTo - (x * w.DstRepeat)                           ' Initial insertion point
      SetCur = w.DstFrom                                          ' Remember where cursor goes

      WHILE ABi < w.Dstto + NumAdj                                ' Big outer loop
         IF ISFALSE IsLData(ABi) AND ISFALSE IsLTop(ABi) THEN INCR Abi: ITERATE ' Ignore non Data lines
         IF SetCur > 0 THEN                                       ' Set cursor just once
            me.CurSetReq(%LineCmd, SetCur, 1, %False)             '
            SetCur = 0                                            '
         END IF                                                   '

         me.LInsertEmpty(ABi, i * r, 0)                           ' Insert enough lines for one set copy
         IF IsMEdit THEN                                          ' In MEdit mode?
            fMIX = me.MEditTbl("I", FORMAT$(ABi))                 ' Go get the MIX value of previous line
            FOR x = ABi + 1 TO ABi + (i * r)                      '
               LMixS(x) = fMIX                                    ' Set into the new line
            NEXT k                                                '
         END IF                                                   '

         NumAdj += i * r                                          ' Track # inserted

         n = ABi + 1                                              ' Calc position to copy the data
         FOR x = 1 TO r                                           ' Do loop'Repeat' times
            z = w.SrcFrom                                         ' Point at 1st line in From Range
            IF w.SrcFrom > w.DstFrom THEN z += NumAdj             ' Adjust if moving upward in the file
            FOR y = 1 TO i                                        ' For each item in the From Range
               GOSUB CopyZ2N                                      ' Go copy data for one line
               INCR n                                             ' Bump 'to' pointer
               IF w.SrcFrom > w.DstFrom THEN                      ' If moving upward in the file
                  IF z < w.SrcTo + NumAdj THEN INCR z             ' Next source line if doing a group replicate
               ELSE                                               '
                  IF z < w.SrcTo THEN INCR z                      ' Next source line if doing a group replicate
               END IF                                             '
            NEXT y                                                ' Loop inner
         NEXT x                                                   ' Loop Outer
         ABI += w.DstRepeat + (i * r)                             ' Big outer loop
      LOOP                                                        ' next ABi

      me.AdjustPending(w.DstFrom + w.DstRepeat - 1, NumAdj, 0)    ' Adjust pending stuff

      IF w.SrcFrom > w.DstFrom THEN w.SrcFrom = w.SrcFrom + NumAdj' Adjust if we moved upward in the file

      me.LStubDelRange(w.SrcFrom, w.SrcFrom + i - 1)              ' Delete the lines

      me.AdjustPending(w.SrcFrom, 0 - i, i * r)                   ' Adjust pending stuff

   '----- Move - SWAP type
   ELSEIF w.DstCmd = "W       " OR w.DstCmd = "WW      " THEN     ' The Swap style?

      IF w.DstFrom = w.DstTo THEN                                 ' A single W(nn) command?
         i = MAX(1, w.DstRepeat)                                  ' Get requested # lines
         w.DstRepeat = 0                                          ' No longer a repeat, used as a count
         i = MIN(i, LastLine - w.DstFrom)                         ' Prevent going past LastLine
         j = w.DstFrom                                            ' Point at 1st line
         DO WHILE i > 0                                           ' Loop for requested count
            IF IsLXclude(j) THEN                                  ' If an XX 'd range, add its size
               j += LWrk1G(j)                                     '
            END IF                                                '
            DECR i: INCR j                                        ' Loop control
         LOOP                                                     '
         w.DstTo = j - 1                                          ' Stuff back adjusted 'To' line number
         IF w.DstTo > LastLine - 1 THEN w.DstTo = LastLine - 1    '
      END IF                                                      '

      IF w.SrcFrom < w.DstFrom THEN                               ' Put MM and WW ranges into Top/Bottom order
         S1Start = w.SrcFrom: S1End = w.SrcTo                     '
         S2Start = w.DstFrom: S2End = w.DstTo                     '
      ELSE                                                        '
         S1Start = w.DstFrom: S1End = w.DstTo                     '
         S2Start = w.SrcFrom: S2End = w.SrcTo                     '
      END IF                                                      '

      IF IsLXclude(S1End) THEN _                                  ' If ends in an XX range, add # XX'd to reach end of range
         S1End += LWrk1G(S1End)                                   '
      IF IsLXclude(S2End) THEN _                                  ' If ends in an XX range, add # XX'd to reach end of range
         S2End += LWrk1G(S2End)                                   '

      '----- First copy the S1 lines AFTER the S2 lines
      i = S1End - S1Start + 1: r = 1                              ' Fudge i and r variables
      me.LInsertEmpty(S2End, i, 0)                                ' Insert lines equal to size of S1

      IF IsMEdit THEN                                             ' In MEdit mode?
         fMIX = me.MEditTbl("I", FORMAT$(S2End))                  ' Go get the MIX value of previous line
         FOR x = S2End + 1 TO S2End + i                           '
            LMixS(x) = fMIX                                       ' Set into the new line
         NEXT k                                                   '
      END IF                                                      '

      me.AdjustPending(S2End, i, i)                               ' Adjust pending stuff

      z = S1Start                                                 ' Point at 1st line in S1 Range
      LblOnce = %True                                             ' Only move labels once
      FOR y = 1 TO i                                              ' For each item in the From Range
         n = S2END + 1 + y - 1                                    ' Calc position to copy the data
         GOSUB CopyZ2N                                            ' Go copy data for one line
         INCR z                                                   ' Next source line
      NEXT y                                                      ' Loop

      '----- Now MM -> HH from S2 to S1
      i = S2End - S2Start + 1: r = 1                              ' Fudge i and r variables

      me.LStubDelRange(S1Start, S1End)                            ' Delete the lines

      S2Start -= (S1End - S1Start + 1)                            '
      S2End -= (S1End - S1Start + 1)                              '

      me.AdjustPending(S1Start, 0 - (S1End - S1Start + 1), 0)     ' Adjust pending stuff

      r = 1                                                       ' Get Repeat factor
      S1Start = S1Start - 1                                       ' Set Dest as if previous line was an 'A'
      IF IsLXclude(S1Start) THEN _                                ' If XX range, add # XX'd to reach end of range
         S1Start += LWrk1G(S1Start)                               '
      me.LInsertEmpty(S1Start, i, 0)                              ' Insert enough lines for it all

      IF IsMEdit THEN                                             ' In MEdit mode?
         fMIX = me.MEditTbl("I", FORMAT$(S1Start))                ' Go get the MIX value of previous line
         FOR x = S1Start + 1 TO S1Start + i                       '
            LMixS(x) = fMIX                                       ' Set into the new line
         NEXT k                                                   '
      END IF                                                      '

      LblOnce = %True                                             ' Only move labels once
      z = S2Start                                                 ' Point at 1st line in S2 Range
      z += i                                                      ' Adjust since moving upward in the file
      FOR y = 1 TO i                                              ' For each item in the S2 Range
         n = S1Start + y                                          ' Calc position to copy the data
         GOSUB CopyZ2N                                            ' Go copy data for one line
         INCR z                                                   ' Next S2 line
      NEXT y                                                      ' Loop

      me.AdjustPending(S1Start, i, i)                             ' Adjust pending stuff
      me.CurSetReq(%LineCmd, S1Start + 1, 0, %False)              ' Set cursor set attempt

      S2Start = S2Start + i                                       ' Adjust if we moved upward in the file

      me.LStubDelRange(S2Start, S2Start + i - 1)                  ' Delete the lines

      me.AdjustPending(S2Start, 0 - i, i)                         ' Adjust pending stuff

   END IF                                                         '
   gfXRebuild = %True                                             ' Force exclude rebuild
   MExitmeth

CopyZ2N:
   me.LFlagBitOff(n, %EQChange)                                   ' Reset ==CHG>
   LLblS(n) = $BlankLNo                                           ' Copy the entry control data
   IF ISTRUE LblOnce THEN LLblS(n) = TRIM$(L(z).LLbl)             ' Copy the label
   LTagS(n) = TRIM$(LTagG(z))                                     ' And the Tag
   LLnumS(n) = LLNumG(z)                                          ' LNumc too
   LLCtlS(n) = $BlankLNo                                          '
   LWrk1S(n) = LWrk1G(z)                                          '
   me.LTxtSet(n, LTxtG(z))                                        '

   IF ISTRUE (LFlagG(z) AND ( %Tabs OR %Word OR %Mark OR %Mask OR %Bounds)) THEN ' Free the dynamic string for special lines
      me.lTxtFree(n)                                              '
      IF IsLWord(z) THEN me.LTxt2Word(n)                          '
      IF IsLMark(z) THEN me.LTxt2Mark(n)                          '
      IF IsLMask(z) THEN me.LTxt2Mask(n)                          '
      IF IsLTabs(z) THEN me.LTxt2Tabs(n)                          '
      IF IsLBounds(z) THEN me.LTxt2Bnds(n)                        '
   END IF                                                         '
   LAttrS(n) = LAttrG(z)                                          ' Copy it over

   LFlagS(n) = LFlagG(z)                                          ' Copy Flag last to sTxtFree doesn't see it
   me.LFlagBitOff(n, %Cursor)                                     ' And never copy the %Cursor flag
   me.LFlagBitOff(n, %Scroll)                                     ' or the scroll
   me.LFlagBitOff(n, %InsertLine)                                 ' or the Insert status
   me.UpdLControl(n)                                              ' Setup LLCtl
   me.lStubDoPM(w.Dstflag, n)                                     ' Do + / - processing for the line
   me.ModSet(n)                                                   ' Remember we changed something
   me.ModSet(z)                                                   ' Remember we changed something
   RETURN
END METHOD

METHOD  lCmdNote(lCmd AS STRING)
'---------- Insert Note(s) into the dataset
LOCAL i, j, k, fMIX AS LONG, w AS lCtlCmd
   MEntry
   me.lCmdParse(w, lCmd)                                          ' Parse the command
   j = w.SrcFrom                                                  ' Get insert point
   i = MAX(1, w.SrcRepeat)                                        ' Get number to be inserted

   me.LInsertEmpty(j, i, %Note)                                   ' Insert the requested line(s)
   FOR k = j + 1 TO j + i                                         '
      IF w.SrcVar <> "." THEN                                     ' A NOTE variation?
         LWrk2S(k) = INSTR($Upper, TRIM$(w.SrcVar))               ' Save index to $Upper of the xNOTE letter
      END IF                                                      '
      me.UpdLControl(k)                                           ' Refresh the line number area
   NEXT k                                                         '
   IF IsMEdit THEN                                                ' In MEdit mode?
      fMIX = me.MEditTbl("I", FORMAT$(j))                         ' Go get the MIX value of previous line
      FOR k = j + 1 TO j + i                                      '
         LMixS(k) = fMIX                                          ' Set into the new line
      NEXT k                                                      '
   END IF                                                         '

   me.AdjustPending(j, i, 0)                                      ' Adjust pending stuff
   me.CurSetReq(%Insert, j + 1, me.IndentAsPrev(j + 1), %True)    ' Set cursor set attempt
   MExit
END METHOD

METHOD  LCmdParse(w AS LCtlCmd, lcmd AS STRING) AS LONG
'---------- Parse the line command into the w. entry format
DIM pEnt(1 TO 12) AS STRING
LOCAL lstr AS STRING
   lstr = SHRINK$(lCmd)
   IF PARSECOUNT(lStr, " ") <> 12 THEN METHOD = %True: EXIT METHOD
   PARSE lStr, pEnt(), " "                                        ' Break it apart
   w.SrcCmd = pEnt(1):  w.SrcFrom = VAL(pEnt(2)): w.SrcTo = VAL(pEnt(3)): w.SrcRepeat = VAL(pEnt(4)): w.SrcFlag = VAL(pEnt(5)): w.SrcVar = pEnt(6)
   w.DstCmd = pEnt(7):  w.DstFrom = VAL(pEnt(8)): w.DstTo = VAL(pEnt(9)): w.DstRepeat = VAL(pEnt(10)): w.DstFlag = VAL(pEnt(11)): w.DstVar = pEnt(12)
END METHOD

METHOD  lCmdPL(lCmd AS STRING)
'---------- Pad lines
LOCAL i, j, x AS LONG, Txt1 AS STRING, w AS lCtlCmd
   MEntry
   me.lCmdParse(w, lCmd)                                          ' Parse the command
   w.SrcRepeat = MAX(MAX(1, Prf.MINLEN), w.SrcRepeat)             ' Minimum of Prf.MINLEN or 1
   IF w.SrcFrom = w.SrcTo THEN                                    ' A single PL(nn) command?
      IF IsLXclude(w.SrcFrom) THEN _                              ' If an XX 'd range, add its size
            w.SrcTo += LWrk1G(w.SrcFrom)                          '
   END IF

   FOR x = w.SrcFrom TO w.SrcTo                                   ' Loop for each line

      IF IsLData(x) THEN                                          ' Only data lines
         IF LEN(L(x).@LTxt) < w.SrcRepeat THEN                    ' If shorter
            me.LTxtSet(x, LSET$(LTxtG(x), w.SrcRepeat))           ' Lengthen it
            me.ModSet(x)                                          ' Remember we changed something
            me.AttrScan(x)                                        ' Recolorize
            me.CurSetReq(%LineCmd, x, 1 + Offset, %False)         ' Set cursor set attempt
         END IF                                                   '
         me.lStubDoPM(w.SrcFlag, x)                               ' Do + / - processing for the line
      END IF                                                      '
   NEXT x                                                         '
   MExit
END METHOD

METHOD  lCmdRR(lCmd AS STRING)
'---------- REPLICATE lines in the dataset
LOCAL i, j, x, y, z, n, fMIX AS LONG, w AS lCtlCmd

   MEntry
   me.lCmdParse(w, lCmd)                                          ' Parse the command
   i = IIF(w.SrcFrom = w.SrcTo, 1, MAX(1, w.SrcTo - w.SrcFrom + 1)) ' Set # lines in the range
   w.SrcRepeat = MAX(w.SrcRepeat, 1)                              ' Default to 1 repeat
   me.LInsertEmpty(w.SrcTo, i * w.SrcRepeat, 0)                   ' Insert enough lines for it all

   FOR x = 1 TO w.SrcRepeat                                       ' Do loop w.SrcRepeat times

      z = w.SrcFrom                                               ' Point at 1st line in From Range
      FOR y = 1 TO i                                              ' For each item in the From Range
         n = w.SrcTo + y + ((x - 1) * i)                          ' Calc position to copy the data
         me.LFlagBitOff(n, %EQChange)                             ' Reset ==CHG>
         LLblS(n) = $BlankLNo                                     ' Copy the entry control data
         LTagS(n) = TRIM$(LTagG(z))                               '
         LWrk1S(n) = LWrk1G(z)                                    '
         me.LTxtSet(n, LTxtG(z))                                  '

         IF ISTRUE (LFlagG(z) AND (%Tabs OR %Word OR %Mark OR %Mask OR %Bounds)) THEN ' Free the dynamic string for special lines
            me.LTxtFree(n)                                        '
            IF IsLWord(z)   THEN me.LTxt2Word(n)                  '
            IF IsLMark(z)   THEN me.LTxt2Mark(n)                  '
            IF IsLMask(z)   THEN me.LTxt2Mask(n)                  '
            IF IsLTabs(z)   THEN me.LTxt2Tabs(n)                  '
            IF IsLBounds(z) THEN me.LTxt2Bnds(n)                  '
         END IF                                                   '
         LAttrS(n) = LAttrG(z)                                    ' Copy it over
         LFlagS(n) = LFlagG(z)                                    ' Copy Flag last so sTxtFree doesn't see it
         me.LFlagBitOff(n, %Cursor)                               ' And never copy the %Cursor flag
         IF IsLData(n) THEN LastReal += 1                         ' Incr LastReal if a data line
         me.UpdLControl(n)                                        ' Setup LLCtl
         me.lStubDoPM(w.SrcFlag, z)                               ' Do + / - processing for the source line
         me.lStubDoPM(w.SrcFlag, n)                               ' Do + / - processing for the new line

         me.ModSet(n)                                             ' Remember we changed something
         IF z < w.SrcTo THEN INCR z                               ' Next source line if doing a group replicate
      NEXT y                                                      ' Loop inner
   NEXT x                                                         ' Loop Outer

   me.AdjustPending(w.SrcTo, (w.SrcTo - w.SrcFrom + 1) * w.SrcRepeat, 1) ' Adjust pending stuff
   me.CurSetReq(%LineCmd, w.SrcTo + ((w.SrcTo - w.SrcFrom + 1) * w.SrcRepeat), 0, %False) ' Set new cursor set attempt
   gfXRebuild = %True                                             ' Force exclude rebuild
   MExit
END METHOD

METHOD  lCmdSI(lCmd AS STRING)
'---------- Show next indent level of an eXcluded range
LOCAL i, j, k, l, popped, lownum, justonce AS LONG, w AS lCtlCmd
   MEntry
   me.lCmdParse(w, lCmd)                                          ' Parse the command
   IF ISFALSE IsLXclude(w.SrcFrom) THEN _                         ' Better be on an eXcluded line
      MErrExit(%eFail, "SI allowed on EXCLUDED lines only")       '

   i = MAX(1, w.SrcRepeat)                                        ' Get number to be 'exploded'
   REDIM Indent(LWrk1G(w.SrcFrom)) AS STRING                      ' Create array as big as the XX block

   k = w.SrcFrom + 1                                              ' k = start line of XX range
   FOR j = 1 TO UBOUND(Indent)                                    ' For each line of XX range
      l = VERIFY(LTxtG(k), " ")                                   ' Get the indent of this line
      IF l = 0 THEN l = MaxLength                                 ' if 0 (empty line) treat as max indent
      Indent(j) = FORMAT$(l, "000000") + " " + FORMAT$(k, "000000")
      INCR k                                                      ' Bump XX line number
   NEXT j                                                         '

   ARRAY SORT Indent(1)                                           ' Sort lowest indent first
   lownum = VAL(LEFT$(Indent(1), 6))                              ' Save the lowest Indent value

   IF i >= LWrk1G(w.SrcFrom) - 1 THEN                             ' Is this the whole range (-1)?
      i = LWrk1G(w.SrcFrom)                                       ' Yes, set i to range size
      me.LTxtFree(w.SrcFrom)                                      ' Go free the dynamic string of the XX line
      me.LEntDel(w.SrcFrom)                                       ' Remove from the L() array
      DECR LastLine: DECR LastReal                                ' Adjust LastLine
      me.CurSetReq(%LineCmd, w.SrcFrom, 1 + Offset, %False)       ' Set cursor set attempt
      FOR j = w.SrcFrom TO w.SrcFrom + i - 1                      ' Un-'hide' the lines in the range
         me.LFlagBitOff(j, %Invisible)                            '
         LWrk1S(j) = 0                                            ' Reset work counter
      NEXT j                                                      '

   ELSE
      REDIM Indent2(i) AS STRING                                  ' Get a 2nd work array
      FOR j = 1 TO i                                              ' Copy the line numbers we're popping out
         Indent2(j) = RIGHT$(Indent(j), 6)                        ' Copy just the line number
      NEXT j
      ARRAY SORT Indent2(1)                                       ' Now sort by line number

      popped = 0                                                  ' Clear popped count
      FOR j = 1 TO i                                              ' Loop for at most the number of Show lines requested
         k = VAL(Indent2(j))                                      ' Get the line number of this line
         k += popped                                              ' Adjust for any previous popped adjustment
         popped = popped + me.LPopOutY(k)                         ' Go popout the line and accum adjustment factor
         IF justonce = 0 THEN                                     ' Try setting the cursor
            justonce = 1                                          ' flip it
            IF IsLXclude(k) THEN                                  ' Back on Xclude line?
               me.CurSetReq(%LineCmd, k - 1, 1 + Offset, %False)  ' Set cursor set attempt
            ELSE                                                  '
               me.CurSetReq(%LineCmd, k, 1 + Offset, %False)      ' Set cursor set attempt
            END IF                                                '
         END IF                                                   '
      NEXT j                                                      '
   END IF                                                         '

   ERASE Indent()                                                 ' Free arrays
   ERASE Indent2()                                                '
   gfXRebuild = %True                                             ' Force exclude rebuild
   MExit
END METHOD

METHOD  lCmdSS(lCmd AS STRING)
'---------- Show excluded range
LOCAL i, j, x AS LONG, w AS lCtlCmd
   MEntry
   me.lCmdParse(w, lCmd)                                          ' Parse the command
   IF w.SrcTo = w.SrcFrom AND IsLFile(w.SrcFrom) THEN             ' Is this the whole =FILE> select?
      FOR x = w.SrcFrom + 1 TO LastLine                           ' Adjust w.SrcTo to the end of the 'file'
         IF LMIXG(x) = LMIXG(w.SrcFrom) THEN                      ' Still the same 'file'
            w.SrcTo = x                                           ' Set new 'to' line
         ELSE                                                     ' Oops, end of 'file' or Bottom
            EXIT FOR                                              ' Scan no more
         END IF                                                   '
      NEXT x                                                      '
   END IF                                                         '

   IF w.SrcFrom = w.SrcTo THEN                                    ' A single x(nn) command?
      i = MAX(1, w.SrcRepeat)                                     ' Get requested # lines
      i = MIN(i, LastLine - w.SrcFrom)                            ' Prevent deleting past LastLine
      j = w.SrcFrom                                               ' Point at 1st line
      DO WHILE i > 0                                              ' Loop for requested count

         IF IsLXclude(j) THEN                                     ' If an XX 'd range, add its size
            j += LWrk1G(j)                                        '
         END IF                                                   '
         DECR i: INCR j                                           ' Loop control
      LOOP                                                        '
      w.SrcTo = j - 1                                             ' Stuff back adjusted 'To' line number
      IF w.SrcTo > LastLine - 1 THEN w.SrcTo = LastLine - 1       ' Prevent deleting past LastLine
   END IF
   w.SrcRepeat = 0                                                ' Nullify Repeat (used as source count)

   me.CurSetReq(%LineCmd, w.SrcFrom, 1 + Offset, %False)          ' Set cursor set attempt
   FOR x = w.SrcFrom TO w.SrcTo                                   ' Loop for each line
      me.LFlagBitOff(x, %Invisible)                               ' Remove invisible
   NEXT x
   gfXRebuild = %True                                             ' Force exclude rebuild
   MExit
END METHOD

METHOD  lCmdTABS(lCmd AS STRING)
'---------- Insert the TABS line
LOCAL w AS lCtlCmd
   MEntry
   me.lCmdParse(w, lCmd)                                          ' Parse the command
   me.InsTabs(w.SrcFrom)                                          ' Insert the Tabs line
   me.CurSetReq(%Insert, w.SrcFrom + 1, 1 + Offset, %False)       ' Set cursor set attempt
   MExit
END METHOD

METHOD  lCmdTF(lCmd AS STRING)
'---------- Text flow
LOCAL WordIX, i, j, k, l, x, y, z, Got1, Got2, L1Done,lclLBnd, lclRBnd, lclLen, fMIX, TM AS LONG
DIM Words(1 TO 1000) AS STRING
DIM WAttr(1 TO 1000) AS WSTRING
LOCAL Txt1, Txt2, nl AS STRING, L1Indent, L2Indent, lclFrom, lclTo, CIP AS LONG, w AS lCtlCmd
LOCAL aTxt AS WSTRING
   MEntry
   me.lCmdParse(w, lCmd)                                          ' Parse the command
   IF LEFT$(w.SrcCmd, 2) = "TM" THEN TM = %True                   ' Set TM mode
   IF ISFALSE TM AND _                                            ' Test this only for TF/TFF/TFTF commands
      (Prf.BndLeft > 1 OR Prf.BndRight > 0) AND w.SrcRepeat > 0 THEN _   ' If bounds active and an entered boundary
      MErrExit(%eFail, "TF right bound specified, but global BNDS are set. TF abandoned.")

   '----- Position past any leading blank lines or non data lines
   me.CurSetReq(%LineCmd, w.SrcFrom, 1 + Offset, %False)          ' Set cursor set attempt
   DO WHILE ISFALSE IsLData(w.SrcFrom) OR (IsLData(w.SrcFrom) AND TRIM$(LTxtG(w.SrcFrom)) = "")
      IF IsLBottom(w.SrcFrom) OR w.SrcFrom > w.SrcTo THEN MExitMeth
      INCR w.SrcFrom                                              ' Skip till next actual data line
   LOOP                                                           '

   lclFrom = w.SrcFrom                                            ' Save start of range (if TFF or TFTF or TMM / TMTM)
   lclTo = w.SrcTo                                                ' Save end of range
   lclLBnd = Prf.BndLeft                                          ' Copy defaults
   lclRBnd = Prf.BndRight: IF lclRBnd = 0 THEN lclRBnd = MaxLength' Get working RBnds
   me.ModSet(lclFrom)                                             ' Remember we changed something

   '----- Outer loop of ALL paragraphs
   DO WHILE lclFrom <= lclTo                                      ' Till we're done
      DO WHILE ISFALSE IsLData(lclFrom) OR (IsLData(lclFrom) AND TRIM$(LTxtG(lclFrom)) = "")
         IF IsLBottom(lclFrom) OR lclFrom > lclTo THEN EXIT DO: EXIT DO
         INCR lclFrom                                             ' Skip till next actual data line
      LOOP                                                        '
      RESET WordIX, L1Indent, L2Indent, Got1, Got2                ' Reset # words stacked and indent values
      fMIX = LMIXG(i)                                             ' Get starting MIX
      w.SrcFrom = lclFrom: w.SrcTo = lclTo                        '
      i = w.SrcFrom                                               ' Start of a single paragraph loop

      '----- Scan lines to determine end of paragraph line
      DO WHILE 1 = 1                                              ' Do till we break out
         IF w.SrcCmd <> "TF      " AND w.SrcCmd <> "TM      " AND i > LclTo THEN EXIT DO ' Break if TFF and at the end

         '----- Just Data lines
         IF IsLData(i) THEN                                       ' Look at text lines

            '----- Exit if =FILE> changes
            IF fMIX <> LMIXG(i) THEN  EXIT DO                     ' Don't flow through =FILE> change

            '----- Get working text from line within the bounds
            IF TM THEN                                            ' If TM/TMTM
               Txt1 = LTxtG(i)                                    ' Then get all the text
            ELSEIF Prf.BndLeft <> 1 OR Prf.BndRight > 0 THEN      ' If bounds active
               Txt1 = MID$(LEFT$(LTxtG(i), lclRBnd), lclLBnd)     ' Get the bounded text
            ELSE                                                  '
               Txt1 = LTxtG(i)                                    ' Get all the text
            END IF                                                '

            '----- Exit if null or a Script languiage delimiter
            IF INSTR(LEFT$(Txt1, 1), ANY ".:&<") AND Got1 THEN EXIT DO  ' Leading char a script delimiter? Yes, exit
            IF ISNULL(TRIM$(Txt1)) THEN EXIT DO                   ' Look for blank line, if so, exit

            '----- Extract the Line 1 Indent
            IF ISFALSE Got1 THEN                                  ' Get L1 indent if needed
               L1Indent = LEN(Txt1) - LEN(LTRIM$(Txt1))           ' Save the indent
               Got1 = %True                                       ' Just once

            '----- Extract the Line 2 Indent
            ELSEIF ISFALSE Got2 THEN                              ' Get L2 indent if needed
               L2Indent = LEN(Txt1) - LEN(LTRIM$(Txt1))           ' Save the indent
               Got2 = %True                                       ' Just once
            END IF                                                '

            '----- If we have Line 2 Indent and Indent changes, we have end of paragraph
            IF Got2 THEN                                          ' If we've got L2 indent
               IF L2indent <> LEN(Txt1) - LEN(LTRIM$(Txt1)) THEN EXIT DO  ' Different indent than 2nd line, exit
            END IF                                                '

            '----- Else keep going
            INCR i                                                ' Incr line number then

         '----- Also exit at end of file
         ELSEIF IsLBottom(i) THEN                                 ' Don't go past the end
            EXIT DO                                               '
         '----- Else, onward
         ELSE                                                     '
            INCR i                                                '
         END IF                                                   '
      LOOP                                                        ' Exit here, i = last line of 'paragraph'

      '----- Save last line of paragraph
      w.SrcTo = i - 1                                             ' Set SrcTo to end of paragraph

      '----- Now extract all the 'words' from the paragraph
      FOR x = w.SrcFrom TO w.SrcTo                                ' Loop for all lines in this paragraph

         '----- Again, only data lines
         IF IsLData(x) THEN                                       '

            '----- Get working text from line within the bounds
            IF TM THEN                                            ' If TM / TMTM
               Txt1 = LTxtG(x)                                    ' Get all the text
            ELSEIF Prf.BndLeft <> 1 OR Prf.BndRight > 0 THEN      ' If bounds active
               Txt1 = MID$(LEFT$(LTxtG(x), lclRBnd), lclLBnd)     ' Get the bounded text
            ELSE                                                  '
               Txt1 = LTxtG(x)                                    ' Get all the text
            END IF                                                '
            aTxt = LAttrG(x)                                      ' Get any associated highlite line
            '----- Scan a text line for words
            i = 1                                                 ' Start in col 1
            DO WHILE ISNOTNULL(Txt1)                              ' Extract all words from each line

               '----- Keep Word table big enough
               INCR WordIX                                        ' Bump Index
               IF WordIX > UBOUND(Words()) THEN                   ' Exceeding table?
                  REDIM PRESERVE Words(1 TO UBOUND(Words()) + 500)' Get another 500 entries
                  REDIM PRESERVE WAttr(1 TO UBOUND(Words()) + 500) '
               END IF                                             '

               '----- Find start of a word
               j = INSTR(i, Txt1, ANY $NonBlank)                  ' Find start of next 'word'
               IF j = 0 THEN EXIT DO                              ' If none, we're done

               '----- Now find end of word
               k = INSTR(j, Txt1, " ")                            ' Look for blank @ end of word

               '----- We have the last word on the line
               IF k = 0 THEN                                      ' No trailing blank?  (i.e. end of line)
                  Words(WordIX) = MID$(Txt1, j) + " "             ' Save as next word
                  WAttr(WordIX) = MID$(aTxt, j + lclLBnd - 1)     ' Save any assoc attribute
                  WAttr(WordIX) = LSET$(WAttr(WordIX), LEN(Words(WordIX)) USING CHR$$(0)) ' Keep lengths the same
                  Txt1 = ""                                       ' Use up the Txt1 string

               '----- We have a 'word'
               ELSE                                               ' We have a trailing blank

                  '----- Look for beginning of next word
                  l = INSTR(k, Txt1, ANY $NonBlank)               ' Look for start of next word

                  '----- No next word, save only one blank
                  IF l = 0 THEN                                   ' No next word, just trailing blanks
                     Words(WordIX) = TRIM$(MID$(Txt1, j)) + " "   ' Add with just one trailing blank
                     WAttr(WordIX) = TRIM$(MID$(aTxt, j + lclLBnd - 1)) + CHR$$(0) ' Save any assoc highlight
                     WAttr(WordIX) = LSET$(WAttr(WordIX), LEN(Words(WordIX)) USING CHR$$(0)) ' Keep lengths the same
                     Txt1 = ""                                    ' Use up the Txt1 string

                  '----- Multiple embedded blanks, save them all
                  ELSE                                            ' We have a word followed by another word
                     Words(WordIX) = MID$(Txt1, j, l - j)         ' Save word plus it's trailing blanks
                     Txt1 = MID$(Txt1, l)                         ' Chop off what we've saved
                     WAttr(WordIX) = MID$(aTxt, j + lclLBnd - 1, l - j)' Save any assoc highlight
                     WAttr(WordIX) = LSET$(WAttr(WordIX), LEN(Words(WordIX)) USING CHR$$(0)) ' Keep lengths the same
                     aTxt = MID$(aTxt, l)                         '
                  END IF                                          '
               END IF                                             '

               '----- Don't accidentally save 'blank words'
               IF ISNULL(TRIM$(Words(WordIX))) THEN _             ' Dummy null word?
                  DECR WordIX                                     ' Pretend we never saw it

               '----- If end of sentence word, add an extra space
               IF RIGHT$(Words(WordIX), 2) = ". " OR _            ' If end of sentence word, add an extra blank
                  RIGHT$(Words(WordIX), 2) = "? " OR _            '
                  RIGHT$(Words(WordIX), 2) = "! " THEN            '
                  Words(WordIX) = Words(WordIX) + " "             '
                  WAttr(WordIX) = WAttr(WordIX) + CHR$$(0)        '
               END IF                                             '
            LOOP                                                  ' Loop till this line all done
         END IF                                                   '
      NEXT x                                                      ' Loop till all lines in Paragraph done

      '----- Select the right bounds
      IF lclRbnd = MaxLength THEN                                 ' If no active right bounds or TM / TMTM
         lclRBnd = IIF(w.SrcRepeat = 0, gDataLen, w.SrcRepeat)    ' If not specified, set Right Margin = to Screen width
      ELSEIF TM THEN                                              '
         lclRBnd = IIF(w.SrcRepeat = 0, lclRBnd, w.SrcRepeat)     '
      END IF

      '----- Make sure we're not asked to do the impossible
      IF L1Indent > 0 OR L2Indent > 0 THEN                        ' Any indents?
         IF L1Indent + LEN(Words(1)) >= lclRbnd OR L2Indent + LEN(Words(1)) >= lclRbnd THEN _  ' Impossible request?
            MErrExit(%eFail, "Line indents exceed the right bound, Text Flow abandoned.")
      END IF                                                      '

      '----- If we're doing entire lines, do things the non-bounds way
      IF (Prf.BndLeft = 1 AND Prf.BndRight = 0) OR TM THEN        ' If no active bounds, delete and insert lines

         '----- First, delete the current paragraph
         i = 0                                                    '
         FOR x = w.SrcFrom TO w.SrcTo                             ' Loop for each deleted line
            IF IsLData(x) THEN                                    ' If Data line
               me.LTxtFree(x)                                     ' Go free the dynamic string
               INCR i                                             ' Count deleted lines
            END IF                                                '
         NEXT x                                                   '
         z = w.SrcFrom                                            ' Set start delete point
         FOR x = w.SrcFrom TO w.SrcTo                             ' Loop for each deleted line
            IF IsLData(z) THEN                                    ' Just data lines
               me.LEntDel(z)                                      ' Remove from the L() array
            ELSE                                                  ' Otherwise
               INCR z                                             ' Step over non data line
            END IF                                                '
         NEXT x                                                   '

         me.AdjustPending(w.SrcFrom, 0 - i, 0)                    ' Adjust pending stuff
         LastLine -= i: LastReal -= i                             ' Adjust LastLine
         lclTo -= i                                               ' Adjust overall 'last' line

         '----- Setup for paragraph rebuild
         x = w.SrcFrom - 1: y = 1                                 ' Set x as the insert point, y as Word index
         nl = STRING$(L1Indent, " ")                              ' Start with initial indent
         aTxt = REPEAT$(L1Indent, CHR$$(0))                       ' Start Attr line

         '----- Loop till all 'words' are re-inserted
         DO WHILE y <= WordIX                                     ' OK, now dump the stuff back out

            '----- Done yet?
            IF y <= WordIX THEN                                   ' Still word(s) to go?

               '----- See if the word will fit on the current line
               IF LEN(nl) + LEN(Words(y)) - 1 <= lclRBnd THEN     ' Still room in line?
                  nl += Words(y)                                  ' Add to the current line
                  aTxt += IIF$(LEN(Words(y)) = LEN(WAttr(y)), WAttr(y), LSET$(WAttr(y), LEN(Words(y)) USING CHR$$(0)))
                  INCR y                                          ' Bump our index

               '----- If word is too big for ANY line
               ELSEIF LEN(Words(y)) > lclRBnd THEN                ' Excessive length word

                  '----- If data waiting in nl, save it now
                  IF ISNOTNULL(TRIM$(nl)) THEN                    ' If something waiting in nl
                     GOSUB InsertNlAtX                            ' Save the nl string
                  END IF                                          '

                  '----- Now save the long guy by himself
                  nl = IIF$(L1Done, STRING$(L2Indent, " "), STRING$(L1Indent, " ")) + Words(y)
                  aTxt = IIF$(L1Done, REPEAT$(L2Indent, CHR$$(0)), REPEAT$(L1Indent, CHR$$(0))) + WAttr(y)
                  GOSUB InsertNlAtX                               ' Save the nl string

                  '----- Get a new line to keep going with
                  nl = STRING$(L2Indent, " ")                     ' Set Line 2 indent
                  aTxt = REPEAT$(L2Indent, CHR$$(0))              '
                  scError(%eFail, "Excessive text length for flow, please check result")
                  INCR y                                          ' Bump our index

               '----- Line is full, save and get another
               ELSE                                               ' No room in line, get another

                  '----- If data waiting in nl, save it now
                  IF ISNOTNULL(TRIM$(nl)) THEN                    ' If something waiting in nl
                     GOSUB InsertNlAtX                            ' Save the nl string
                  END IF                                          '
                  nl = STRING$(L2Indent, " ")                     ' Set Line 2 indent
                  aTxt = REPEAT$(L2Indent, CHR$$(0))              '
               END IF                                             '
            END IF                                                '
         LOOP                                                     '
         IF ISNOTNULL(TRIM$(nl)) THEN                             ' Handle remaining data in line
            GOSUB InsertNlAtX                                     ' Save the nl string
         END IF                                                   '
         lclFrom = x + 1                                          ' Set start of next paragraph

      '----- Now do it all the BNDS way
      ELSE                                                        ' Now do it the BOUNDS active way
         x = w.SrcFrom: y = 1: lclLen = lclRBnd - lclLBnd + 1     ' Set x as this line, y as Word index
         aTxt = LAttrG(x)                                         ' Get Attr line
         IF LEN(aTxt) < lclRBnd THEN aTxt = LSET$(aTxt, lclRBnd USING CHR$$(0)) ' Make it at least lclRBnd long
         MID$(atxt, lclLBnd TO lclRBnd) = REPEAT$(lclRBnd - lclLBnd + 1, CHR$$(0)) ' Blank bounded color settings
         nl = STRING$(L1Indent, " ")                              ' Restore initial indent

         '----- Loop through words
         DO WHILE y <= WordIX                                     ' OK, now dump the stuff back out

            '----- If it still fits, just add to the line
            IF LEN(nl) + LEN(Words(y)) - 1 <= lclLen THEN         ' Still room in line?
               CIP = LEN(nl) + lclLBnd                            ' Point where Attr string goes
               nl += Words(y)                                     ' Yes, add this word
               MID$(aTxt, CIP, LEN(WAttr(y))) = WAttr(y)          ' Stuff Attr back
               INCR y                                             ' And, point at the next one

            '----- No more room in line
            ELSE                                                  ' No room in line, stuff it back
               nl = LSET$(nl, lclLen)                             ' Pad length to size of bounded area
               Txt2 = LTxtG(x)                                    ' Get the current string
               IF LEN(Txt2) < lclrBnd THEN Txt2 = LSET$(Txt2, lclRBnd) ' Pad length to size of bounded area if needed
               MID$(Txt2, lclLBnd, lclLen) = nl                   ' Insert the revised string into the text line
               me.LTxtSet(x, Txt2)                                '
               LAttrS(x) = aTxt                                   ' replece it
               me.AttrScan(x)                                     ' Recolorize
               nl = STRING$(L2Indent, " ")                        ' Set Line 2 indent for the next line
               INCR x                                             ' Bump x
               aTxt = LAttrG(x)                                   ' Get Attr line
               IF LEN(aTxt) < lclRBnd THEN aTxt = LSET$(aTxt, lclRBnd USING CHR$$(0)) ' Make it at least lclRBnd long
               MID$(atxt, lclLBnd TO lclRBnd) = REPEAT$(lclRBnd - lclLBnd + 1, CHR$$(0)) ' Blank bounded color settings
            END IF                                                '
         LOOP                                                     '

         '----- Clear out last data if needed
         IF ISNOTNULL(TRIM$(nl)) THEN                             ' Data in last line?
            nl = LSET$(nl, lclLen)                                ' Pad length to size of bounded area
            Txt2 = LTxtG(x)                                       ' Get the current string
            IF LEN(Txt2) < lclrBnd THEN Txt2 = LSET$(Txt2, lclRBnd) ' Pad length to size of bounded area if needed
            MID$(Txt2, lclLBnd, lclLen) = nl                      ' Insert the revised string into the text line
            me.LTxtSet(x, Txt2)                                   '
            LAttrS(x) = aTxt                                      ' replace it
            me.AttrScan(x)                                        ' Recolorize
            nl = STRING$(L2Indent, " ")                           ' Set Line 2 indent for the next line
            INCR x                                                ' Bump x
            aTxt = LAttrG(x)                                      ' Get Attr line
            IF LEN(aTxt) < lclRBnd THEN aTxt = LSET$(aTxt, lclRBnd USING CHR$$(0)) ' Make it at least lclRBnd long
            MID$(atxt, lclLBnd TO lclRBnd) = REPEAT$(lclRBnd - lclLBnd + 1, CHR$$(0)) ' Blank bounded color settings
         END IF                                                   '

         '----- Pad blanks in lines below last filled line
         DO WHILE x <= w.SrcTo                                    ' Pad remainder of lines
            Txt2 = LTxtG(x)                                       ' Get the current string
            IF LEN(Txt2) < lclrBnd THEN Txt2 = LSET$(Txt2, lclRBnd) ' Pad length to size of bounded area if needed
            MID$(Txt2, lclLBnd, lclLen) = SPACE$(lclLen)          ' Insert the revised string into the text line
            me.LTxtSet(x, Txt2)                                   '
            aTxt = LAttrG(x)                                      ' Get Attr line
            IF LEN(aTxt) < lclRBnd THEN aTxt = LSET$(aTxt, lclRBnd USING CHR$$(0)) ' Make it at least lclRBnd long
            MID$(atxt, lclLBnd TO lclRBnd) = REPEAT$(lclLen, CHR$$(0))  ' Blank bounded color settings
            LAttrS(x) = aTxt                                      ' Put it back
            me.AttrScan(x)                                        ' Recolorize
            INCR x                                                ' Next line
         LOOP
         lclFrom = x                                              ' Set start of next paragraph
      END IF                                                      '
   LOOP                                                           ' End outer paragraph loop
   MExitmeth

InsertNlAtX:
   L1Done = %True                                                 ' At least one line was done
   me.LInsertEmpty(x, 1, %Data)                                   ' Insert a line
   IF IsMEdit THEN LMIXS(x + 1) = fMIX                            ' If MEdit, keep same MIX index
   me.LTxtSet(x + 1 , RTRIM$(nl))                                 ' Stuff in the nl string
   me.AdjustPending(x, 1, 0)                                      ' Adjust pending stuff
   me.UpdLControl(x + 1)                                          ' Tidy LLCtl
   LAttrS(x + 1) = aTxt                                           ' Store it
   me.AttrScan(x + 1)                                             ' Recolorize
   INCR x                                                         ' Bump line insertion point
   INCR lclTo                                                     ' Bump last range line
   RETURN
END METHOD

METHOD  lCmdTJJ(lCmd AS STRING)
'---------- Text Join tagged line with the next line
LOCAL i, j, k, x AS LONG, t1, t2 AS STRING, w AS lCtlCmd
LOCAL a1, a2, aTxt AS WSTRING
   MEntry
   me.lCmdParse(w, lCmd)                                          ' Parse the command
   i = w.SrcTo - w.SrcFrom                                        ' Calc # in range
   IF w.SrcFrom = w.SrcTo THEN                                    ' A single J(nn) command?
      i = MAX(1, w.SrcRepeat - 1)                                 ' Get requested # lines
   END IF

   FOR x = w.SrcFrom TO w.SrcFrom + i                             ' Check for X'd stuff
      IF ISFALSE IsLData(x) OR IsLInvisible(x) OR IsLFile(x)THEN _' All data and visible?
         MErrExit(%eFail, "Text Join/Glue range must be only normal Data lines")
   NEXT x                                                         '

   IF w.SrcFrom => LastLine - 1 THEN _                            ' Better be a next line
      MErrExit(%eFail, "No next line to join with")               '

   w.SrcTo = w.SrcFrom + i                                        '
   j = w.SrcFrom: k = j + 1: x = 0                                ' Point at 1st line and next line
   me.ModSet(w.SrcFrom)                                           ' Remember we changed something

   DO WHILE i > 0                                                 ' Loop for requested count
      t1 = LTxtG(j): t2 = LTxtG(k)                                ' Get the two pieces
      a1 = LAttrG(j): a2 = LAttrG(k)                              ' Get any Attr strings
      a1 = LSET$(a1, LEN(t1)): a2 = LSET$(a2, LEN(t2))            ' Make same lengths as T1/T2

      IF j <> w.SrcFrom THEN                                      ' Ltrim all but the first line
         t1 = LTRIM$(t1)                                          '
         a1 = RIGHT$(a1, LEN(t1))                                 ' Trim the Attr the same
      END IF                                                      '

      IF j <> w.SrcTo THEN                                        ' Rtrim all but the last line
         t1 = RTRIM$(t1)                                          '
         a1 = LEFT$(a1, LEN(t1))                                  ' Trim the Attr the same
      END IF                                                      '

      IF k <> w.SrcFrom THEN                                      ' Ltrim all but the last line
         t2 = LTRIM$(t2)                                          '
         a2 = RIGHT$(a2, LEN(t2))                                 ' Trim the Attr the same
      END IF                                                      '

      IF k <> w.SrcTo THEN                                        ' Rtrim all but the last line
         t2 = RTRIM$(t2)                                          '
         a2 = LEFT$(a2, LEN(t2))                                  ' Trim the Attr the same
      END IF                                                      '

      IF w.SrcCmd = "TJ      " OR w.SrcCmd = "TJJ     " THEN      ' Join?
         me.LTxtSet(j, t1 + " " + t2)                             ' Do the Join
         aTxt = a1 + CHR$$(0) + a2                                '
         LAttrS(j) = aTxt                                         ' Save it

      ELSE                                                        ' Else must be Glue
         me.LTxtSet(j, t1 + ENV.GlueWith + t2)                    ' Do the Glue style
         aTxt = a1 + REPEAT$(LEN(ENV.GlueWith), CHR$$(0)) + a2    '
         LAttrS(j) = aTxt                                         ' Save it
      END IF                                                      '
      me.AttrScan(j)                                              ' Recolorize
      me.LTxtFree(k)                                              ' Go free the dynamic string
      me.LEntDel(k)                                               ' and remove the Joined/Glued line
      DECR i: INCR x                                              ' Loop control
   LOOP                                                           '

   me.AdjustPending(w.SrcFrom, 0 - x, 0)                          ' Adjust pending stuff
   me.CurSetReq(%LineCmd, w.SrcFrom, 0, %True)                    ' Set cursor set attempt

   LastLine -= x: LastReal -= x                                   ' Adjust LastLine
   IF (w.SrcCmd = "TG      " OR w.SrcCmd = "TGG     ") AND ISNOTNULL(ENV.GlueWith) THEN _ ' Glue with non-null ENV.GlueWith?
      scError(%eNone, "Glued with " + $DQ + ENV.GlueWith + $DQ)
   MExit
END METHOD

METHOD  lCmdTL(lCmd AS STRING)
'---------- Truncate lines
LOCAL i, j, x, nl AS LONG, aTXT AS WSTRING, w AS lCtlCmd
   MEntry
   me.lCmdParse(w, lCmd)                                          ' Parse the command
   IF w.SrcFrom = w.SrcTo THEN                                    ' A single x(nn) command?
      IF IsLXclude(w.SrcFrom) THEN                                ' If an XX 'd range, add its size
         w.SrcTo = w.SrcFrom + LWrk1G(w.SrcTo)                    '
      END IF                                                      '
   END IF                                                         '
   IF w.SrcFrom = w.SrcTo AND w.SrcRepeat = 0 THEN MExitMeth      ' Do nothing for TL with no number

   nl = w.SrcRepeat                                               ' Get requested line length

   IF nl = 0 THEN                                                 ' If no value specified
      FOR x = w.SrcFrom TO w.SrcTo                                ' Loop for each line in range
         IF IsLData(x) THEN                                       ' Only data lines
            nl = MAX(LEN(L(x).@LTxt), nl)                         ' Find the longest line in the range
         END IF                                                   '
      NEXT x                                                      '
   END IF                                                         '

   FOR x = w.SrcFrom TO w.SrcTo                                   ' Loop for each line
      IF IsLData(x) THEN                                          ' Only data lines
         IF LEN(L(x).@LTxt) > nl THEN                             ' If longer, truncate it
            me.CurSetReq(%LineCmd, x, 1 + Offset, %False)         ' Set cursor set attempt
            me.LTxtSet(x, LEFT$(LTxtG(x), nl))                    ' Truncate the line
            aTxt = LAttrG(x)                                      ' Get Attr if present
            aTxt = LEFT$(atxt, nl)                                ' Truncate it as well
            LAttrS(x) = aTxt                                      ' Put back
            me.AttrScan(x)                                        ' Recolorize
            me.ModSet(x)                                          ' Remember we changed something

         ELSEIF LEN(L(x).@LTxt) = nl THEN                         ' If equal, do nothing
            '
         ELSE                                                     ' If shorter, extend it
            me.CurSetReq(%LineCmd, x, 1 + Offset, %False)         ' Set cursor set attempt
            me.LTxtSet(x, LSET$(LTxtG(x), nl))                    ' Extend the line
            me.AttrScan(x)                                        ' Recolorize
            me.ModSet(x)                                          ' Remember we changed something
         END IF                                                   '

         me.lStubDoPM(w.SrcFlag, x)                               ' Do + / - processing for the line
      END IF                                                      '
   NEXT x                                                         '
   MExit
END METHOD

METHOD  lCmdTR(lCmd AS STRING)
'---------- Trim lines
LOCAL x AS LONG, aTxt AS WSTRING, Txt1 AS STRING, w AS lCtlCmd
   MEntry
   me.lCmdParse(w, lCmd)                                          ' Parse the command
   FOR x = w.SrcFrom TO w.SrcTo                                   ' Loop for each line
      IF IsLData(x) THEN                                          ' Only data lines
         Txt1 = RTRIM$(LTxtG(x))                                  ' Trim the line

         IF LEN(L(x).@LTxt) > w.SrcRepeat AND _                   ' Only if longer than specified value
            LEN(Txt1) <> LEN(L(x).@LTxt) THEN                     ' AND actually some trailing blanks
            me.ModSet(x)                                          ' Remember we changed something
            me.CurSetReq(%LineCmd, x, 1 + Offset, %False)         ' Set cursor set attempt
            IF LEN(Txt1) < w.SrcRepeat THEN Txt1 = LSET$(Txt1, w.SrcRepeat) ' Add back if now less than requested
            me.LTxtSet(x, Txt1)                                   ' Save text back
            aTxt = LAttrG(x)                                      ' Get Attr if present
            aTxt = LEFT$(atxt, LEN(Txt1))                         ' Truncate it as well
            LAttrS(x) = aTxt                                      ' Put back
            me.AttrScan(x)                                        ' Recolorize
         END IF                                                   '

         me.lStubDoPM(w.SrcFlag, x)                               ' Do + / - processing for the line
      END IF                                                      '
   NEXT x                                                         '
   MExit
END METHOD

METHOD  lCmdTS(lCmd AS STRING)
'---------- Text Split Commands TS/TB/TBB
LOCAL i, j, k, indent, wfrom, wto AS LONG, ol, nl, ob, nb, obc, nbc AS STRING, w AS lCtlCmd
LOCAL oc, nc AS WSTRING
   MEntry
   me.lCmdParse(w, lCmd)                                          ' Parse the command
   IF w.SrcRepeat > 100 THEN _                                    ' Reasonable # ?
      MErrExit(%eFail, "TS/TB [n] seems unreasonable")            '

   IF ISFALSE IsLData(w.SrcFrom) AND ISFALSE IsLNote(w.SrcFrom) THEN _ ' Is Cursor on an eligible line
      MErrExit(%eFail, "Cursor is not on an eligible line")       '

   IF sGetIX(CRow) < w.SrcFrom OR sGetIX(CRow) > w.SrcTo THEN _   ' Is Cursor on the TS/TB/TBB range of lines?
      MErrExit(%eFail, "Cursor was NOT on the TS/TB/TBB lines")   '

   '----- Loop as needed
   wFrom = w.SrcFrom: wTo = w.SrcTo                               ' Get loop values
   DO WHILE wfrom <= wTo                                          ' Do all this for each line

      me.LFlagBitOff(wFrom, %EQChange)                            ' Reset ==CHG>
      me.LFlagBitOff(wFrom, %Invisible)                           ' Reset any invisible
      LWrk1S(wFrom) = 0                                           ' Reset work count
      me.UpdLControl(wFrom)                                       ' And LLCtl
      me.ModSet(wFrom)                                            ' Remember we changed something
      IF IsLXclude(wFrom) THEN                                    ' Is this an Exclude line?
         me.LTxtFree(wFrom)                                       ' Go free the dynamic string of the XX line
         me.LEntDel(wFrom)                                        ' Remove from the L() array
         DECR LastLine: DECR LastReal                             ' Adjust LastLine
         wTo -= 1                                                 ' Adjust wTo line
         ITERATE DO                                               ' Loop back
      END IF                                                      '

      IF Prf.BndLeft = 1 AND Prf.BndRight = 0 THEN                ' If no bounds set

         i = CCol - gLNPadCol + Offset                            ' Calc IX into data line where cursor is
         IF i < 1 THEN MExitMeth                                  '
         IF ISNULL(TRIM$(LTxtG(wFrom))) THEN                      ' If line empty, indent = 1
            indent = 1                                            '
         ELSE                                                     '
            indent = VERIFY(LTxtG(wFrom), " ")                    ' Get the indent
         END IF                                                   '

         ol = LEFT$(LTxtG(wFrom), i - 1)                          ' Split into left
         nl = LTRIM$(MID$(LTxtG(wFrom), i))                       ' and right
         j = MAX(1, w.SrcRepeat)                                  ' Get insert blank line number
         IF BIT(w.SrcFlag, %lCmdForce0) THEN j = 0                ' Specific zero?  Set zero then
         oc = LEFT$(LAttrG(wFrom), i - 1)                         ' Split into left
         nc = MID$(LAttrG(wFrom), i)                              ' and right
         nc = RIGHT$(nc, LEN(nl))                                 '

         IF j > 0 THEN                                            ' Any lines to insert?
            me.LInsertEmpty(wFrom, j, IIF(LEFT$(w.SrcCmd, 2) = "TS", %InsertLine OR %Data, %Data)) ' Insert the requested insert line(s)
            FOR k = wFrom + 1 TO wFrom + j                        ' Add to Ins Table
               IF LEFT$(w.SrcCmd, 2) = "TS" THEN                  ' TS Insert lines?
                  me.InsTblAdd(k)                                 ' Setup for '''''' clear
                  me.LTxtSet(k, Prf.MaskLine)                     ' Stuff in the mask
               END IF                                             '
               me.UpdLControl(k)                                  '
            NEXT k                                                '
         END IF                                                   '
         me.LInsertEmpty(wFrom + j, 1, %Data)                     ' Insert the split data line

         me.AdjustPending(wFrom, j + 1, 0)                        ' Adjust pending stuff
         me.CurSetReq(%LineCmd, wFrom, i, %False)                 ' Set cursor set attempt

         me.LTxtSet(wFrom, ol)                                    ' Swap in left hand data
         me.LTxtSet(wFrom + j + 1, SPACE$(indent - 1) + nl)       ' Swap in right hand data

         LAttrS(wFrom) = oc                                       ' Swap the Attr values
         LAttrS(wFrom + j + 1) = nc                               '
         me.AttrScan(wFrom)                                       ' Recolorize
         me.AttrScan(wFrom + j + 1)                               ' Recolorize

         IF IsLNote(wFrom) THEN                                   ' Fixup NOTEs if needed
            FOR i = WFrom TO wFrom + j + 1                        ' Setup loop
               me.LFlagBitOff(i, %Data)                           ' Fixup bits
               me.LFlagBitOff(i, %Insert)                         '
               me.LFlagBitOn(i, %Note)                            '
            NEXT i                                                '

         ELSE                                                     '
            OnInsClnSupp                                          ' Suppress Insert cleanup for 1 Attn pass
         END IF                                                   '
         wFrom += j + 2: wTo += j + 1                             '
      ELSE                                                        ' This is one with Bounds active

         '----- Do it the difficult way with BOUNDS active

         i = CCol - gLNPadCol + Offset                            ' Calc IX into data line where cursor is
         IF i < Prf.BndLeft OR i > Prf.BndRight THEN MExitMeth    ' Split point not within bounds

         indent = Prf.BndLeft                                     ' Indent = LBnds

         ol = LTxtG(wFrom)                                        ' Get the old text line
         MID$(ol, i TO Prf.BndRight) = SPACE$(Prf.BndRight - i + 1) ' Split into left
         ol = RTRIM$(ol)                                          '
         nl = LTRIM$(MID$(LTxtG(wFrom), i TO Prf.BndRight))       ' and right
         j = MAX(1, w.SrcRepeat)                                  ' Get insert blank line number
         IF BIT(w.SrcFlag, %lCmdForce0) THEN
             j = 0                                                ' Specific zero?  Set zero then
         END IF
         oc = LAttrG(wFrom)                                       ' Get the old text line
         oc = LSET$(oc, Prf.BndRight)                             '
         MID$(oc, i TO Prf.BndRight) = REPEAT$(Prf.BndRight - i + 1, CHR$$(0)) ' Split into left
         nc = LTRIM$(MID$(LAttrG(wFrom), i TO Prf.BndRight))      ' and right
         nc = RIGHT$(nc, LEN(nl))                                 '

         IF j > 0 THEN                                            ' Any lines to insert?
            me.LInsertEmpty(wFrom, j, IIF(LEFT$(w.SrcCmd, 2) = "TS", %InsertLine OR %Data, %Data)) ' Insert the requested insert line(s)
            FOR k = wFrom + 1 TO wFrom + j                        ' Add to Ins Table
               IF LEFT$(w.SrcCmd, 2) = "TS" THEN                  ' TS Insert lines?
                  me.InsTblAdd(k)                                 ' Setup for '''''' clear
                  me.LTxtSet(k, Prf.MaskLine)                     ' Stuff in the mask
               END IF                                             '
               me.UpdLControl(k)                                  '
            NEXT k                                                '
         END IF                                                   '
         me.LInsertEmpty(wFrom + j, 1, %Data)                     ' Insert the split data line

         me.AdjustPending(wFrom, j + 1, 0)                        ' Adjust pending stuff
         me.CurSetReq(%LineCmd, wFrom, i, %False)                 ' Set cursor set attempt

         me.LTxtSet(wFrom, ol)                                    ' Swap in left hand data
         me.LTxtSet(wFrom + j + 1, SPACE$(indent - 1) + nl)       ' Swap in right hand data

         LAttrS(wFrom) = oc                                       '
         LAttrS(wFrom + j + 1) = REPEAT$(indent - 1, CHR$$(0)) + nc  '
         me.AttrScan(wFrom)                                       ' Recolorize
         me.AttrScan(wFrom + j + 1)                               ' Recolorize

         IF IsLNote(wFrom) THEN                                   ' Fixup NOTEs if needed
            FOR i = WFrom TO wFrom + j + 1                        ' Setup loop
               me.LFlagBitOff(i, %Data)                           ' Fixup bits
               me.LFlagBitOff(i, %Insert)                         '
               me.LFlagBitOn(i, %Note)                            '
            NEXT i                                                '

         ELSE                                                     '
            OnInsClnSupp                                          ' Suppress Insert cleanup for 1 Attn pass
         END IF                                                   '
         wFrom += j + 2: wTo += j + 1                             '
      END IF
   LOOP                                                           ' Again
      DoSet(%Refresh)                                             ' Have it looked at
   MExit
END METHOD

METHOD  lCmdTT(lCmd AS STRING)
'---------- Text Select command
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL w AS lCtlCmd
   MEntry
   me.lCmdParse(w, lCmd)                                          ' Parse the command

   IF Prf.HexMode = &4 THEN sDoBeep: MErrExit(%eFail, "Text select not supported in HEX mode")

   FOR i = w.SrcFrom TO w.SrcTo                                   ' Lets calc a maximum line length
      IF ISFALSE IsLData(i) THEN ITERATE FOR                      ' Skip non-Data lines
      j = MAX(j, LEN(L(i).@LTxt))                                 ' Save max length
   NEXT i                                                         '

   MarkSLin = w.SrcFrom: MarkELin = w.SrcTo                       ' Set mark boundaries
   IF Prf.BndLeft = 1 AND Prf.BndRight = 0 THEN                   ' Entire line bounds?
      MarkSCol = 1:         MarkECol = j                          '
   ELSEIF Prf.BndLeft > 1 AND Prf.BndRight = 0 THEN               ' LB to Max?
      MarkSCol = Prf.BndLeft:  MarkECol = j                       '
   ELSEIF Prf.BndLeft > 1 AND Prf.BndRight > 0 THEN               ' LB to RB?
      MarkSCol = Prf.BndLeft:  MarkECol = Prf.BndRight            '
   END IF                                                         '
   OnMarkActive                                                   '
   me.MarkScr                                                     ' Redraw the hilight

   me.CurSetReq(%LineCmd, w.SrcFrom, 0, %True)                    ' Set cursor set attempt
   DoSet(%MarkKillSkip)                                           ' Prevent Enter from clearing this.
   MExit
END METHOD

METHOD  lCmdTU(lCmd AS STRING)
'---------- Toggle User line(s) in the Display
LOCAL i, j, x, y, fMIX AS LONG, w AS lCtlCmd
   MEntry
   me.lCmdParse(w, lCmd)                                          ' Parse the command

   '----- Do the TU of a =FILE> line
   IF w.SrcTo = w.SrcFrom AND IsLFile(w.SrcFrom) THEN             ' Do  a TU on a FILE line
      FOR x = w.SrcFrom + 1 TO LastLine                           ' Adjust w.SrcTo to the end of the 'file'
         IF LMIXG(x) = LMIXG(w.SrcFrom) THEN                      ' Still the same 'file'
            w.SrcTo = x                                           ' Set new 'to' line
         ELSE                                                     ' Oops, end of 'file' or Bottom
            EXIT FOR                                              ' Scan no more
         END IF                                                   '
      NEXT x                                                      '
   END IF                                                         '

   FOR x = w.SrcFrom TO w.SrcTo                                   ' Loop for each XX'd line
      IF IsLFile(x) THEN ITERATE FOR                              ' Ignore FILE lines
      IF IsLData(x) THEN                                          ' Just Data lines
         me.LFlagBitTog(x, %User)                                 ' Toggle User
      END IF                                                      '
   NEXT x                                                         '

   me.CurSetReq(%LineCmd, w.SrcFrom, 1 + Offset, %False)          ' Set cursor set attempt
   OnUndoFlag                                                     ' Force a snapshot
   gfXRebuild = %True                                             ' Force exclude rebuild
   MExit
END METHOD

METHOD  lCmdTX(lCmd AS STRING)
'---------- Toggle Exclude line(s) in the Display
LOCAL i, j, x, y, fMIX AS LONG, w AS lCtlCmd
   MEntry
   me.lCmdParse(w, lCmd)                                          ' Parse the command

   '----- DO the TX of a =FILE> line
   IF w.SrcTo = w.SrcFrom AND IsLFile(w.SrcFrom) THEN             ' Do  a TX on a FILE line
      FOR x = w.SrcFrom + 1 TO LastLine                           ' Adjust w.SrcTo to the end of the 'file'
         IF LMIXG(x) = LMIXG(w.SrcFrom) THEN                      ' Still the same 'file'
            w.SrcTo = x                                           ' Set new 'to' line
         ELSE                                                     ' Oops, end of 'file' or Bottom
            EXIT FOR                                              ' Scan no more
         END IF                                                   '
      NEXT x                                                      '
   END IF                                                         '

   FOR x = w.SrcFrom TO w.SrcTo                                   ' Loop for each XX'd line
      IF IsLFile(x) THEN ITERATE FOR                              ' Ignore FILE lines
      IF IsLData(x) OR IsLNote(x) THEN                            ' Just Data and Note lines
         me.LFlagBitTog(x, %Invisible)                            ' Toggle Invisible
      END IF                                                      '
   NEXT x                                                         '

   me.CurSetReq(%LineCmd, w.SrcFrom + 1, 1 + Offset, %False)      ' Set cursor set attempt
   OnUndoFlag                                                     ' Force a snapshot
   gfXRebuild = %True                                             ' Force exclude rebuild
   MExit
END METHOD

METHOD  lCmdWord(lCmd AS STRING)
'---------- Insert the WORD line
LOCAL w AS lCtlCmd
   MEntry
   me.lCmdParse(w, lCmd)                                          ' Parse the command
   me.InsWord(w.SrcFrom)                                          ' Insert the WORD line
   me.CurSetReq(%Insert, w.SrcFrom + 1, 1 + Offset, %False)       ' Set cursor set attempt
   MExit
END METHOD

METHOD  lCmdXX(lCmd AS STRING)
'---------- Exclude line(s) from the Display
LOCAL i, j, x, y, fMIX AS LONG, w AS lCtlCmd
   MEntry
   me.lCmdParse(w, lCmd)                                          ' Parse the command

   '----- Handle the X of a =FILE> line
   IF w.SrcTo = w.SrcFrom AND IsLFile(w.SrcFrom) THEN             ' Is this the whole =FILE> exclude?
      FOR x = w.SrcFrom + 1 TO LastLine                           ' Adjust w.SrcTo to the end of the 'file'
         IF LMIXG(x) = LMIXG(w.SrcFrom) THEN                      ' Still the same 'file'
            w.SrcTo = x                                           ' Set new 'to' line
         ELSE                                                     ' Oops, end of 'file' or Bottom
            EXIT FOR                                              ' Scan no more
         END IF                                                   '
      NEXT x                                                      '
   END IF                                                         '

   FOR x = w.SrcFrom TO w.SrcTo                                   ' Loop for each XX'd line
      IF ISFALSE IsLFile(x) AND _                                 ' =FILE> line never exclude
         ISFALSE IsLXClude(x) THEN _                              ' Xclude lines themselves
         me.LFlagBitOn(x, %Invisible)                             ' Mark the line Invisible now
   NEXT x                                                         '

   me.CurSetReq(%LineCmd, w.SrcTo + 1, 0, %False)                 ' Set cursor set attempt
   OnUndoFlag                                                     ' Force a snapshot
   gfXRebuild = %True                                             ' Force exclude rebuild
   MExit
END METHOD

METHOD  LCtlCommands()
'---------- Process the Line commands we've scanned out
LOCAL w AS LCtlCmd
LOCAL c, i, j, k, f, t AS LONG, lCmd AS STRING, CallAddr AS DWORD
   MEntry
   IF LTblRange THEN MExitMeth                                    ' If a primary range waiting, do nothing

   DO WHILE LTblBIX > 0                                           ' OK, loop through them
      w = LTblB(1)                                                ' Copy 1st entry to work data
      ARRAY DELETE LTblB(1)                                       ' Delete this entry
      DECR LTblBIX                                                ' Decr count of entries remaining

      GOSUB MEditCheck                                            ' Look for invalid MEdit ranges

      '----- Locate command in table, pass control if executable
      c = LCmdT.GetCmdIX(TRIM$(w.SrcCmd))                         '
      gCrashLastLCmd = w.SrcCmd                                   ' Last Line command

      GOSUB BuildLCmd                                             ' Build the w. entry into a command

      CallAddr = LCmdT.GetEditCode(c)                             ' Get the address of the routine
      CALL DWORD CallAddr USING lCmdI(lCmd)                       ' Call the routine
   LOOP                                                           ' Loop back till done
   MExitMeth

BuildLCmd:
   lCmd  = w.SrcCmd + STR$(w.SrcFrom) + STR$(w.SrcTo) + STR$(w.SrcRepeat) + STR$(w.SrcFlag) + " " + w.SrcVar + " "
   lCmd += w.DstCmd + STR$(w.DstFrom) + STR$(w.DstTo) + STR$(w.DstRepeat) + STR$(w.DstFlag) + " " + w.DstVar + " "
   RETURN

MEditCheck:
   IF ISFALSE IsMEdit THEN RETURN                                 ' Do nothing if not MEdit mode
   IF INSTR("R       RR      C       CC      M       MM      D       DD      ", UUCASE(w.SrcCmd)) = 0 THEN RETURN ' If not an R/C/M/D command, do nothing
   f = w.SrcFrom: t = w.SrcTo                                     ' Get working copy

   IF f = t THEN                                                  ' A single x(nn) command?
      IF IsEQ(LEFT$(w.SrcCmd, 1), "D") AND w.SrcRepeat < 2 AND IsLFile(w.SrcFrom) THEN RETURN ' Allow single explicit D of a =FILE> line
      i = MAX(1, w.SrcRepeat)                                     ' Get requested # lines
      i = MIN(i, LastLine - f)                                    ' Prevent going past LastLine
      j = f                                                       ' Point at 1st line
      DO WHILE i > 0                                              ' Loop for requested count
         IF IsLXclude(j) THEN                                     ' If an XX 'd range, add its size
            j += LWrk1G(j)                                        '
         END IF                                                   '
         DECR i: INCR j                                           ' Loop control
      LOOP                                                        '
      t = j - 1                                                   ' Stuff back adjusted 'To' line number
   END IF                                                         '

   FOR i = f TO t                                                 ' See if any =FILE> lines in range
      IF IsLFile(i) THEN                                          ' Yes, kill it
         scError(%eFail, "Can't have = FILE > line within data range")
         me.CurSetReq(%LError, i, 0, %True)                       '
         MExitMeth                                                ' Bail out
      END IF                                                      '
   NEXT i                                                         '
   RETURN                                                         '
END METHOD

METHOD  LCtlImmediate()
'---------- Build TblB entries from Immediate TblA entries
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL d, x, y AS LONG, t AS TouchEntry

   MEntry
   '----- Process the "I" entries in TblA
   FOR i = 1 TO LTblAIX                                           ' See if any still in Pending status
      IF LTblA(i).CmdType <> "I" THEN ITERATE FOR                 ' Ignore non-immediate commands

      INCR LTblBIX                                                ' Bump TblB Index
      LTblB(LTblBIX).SrcCmd = LTblA(i).Cmd                        ' Copy the data over
      LTblB(LTblBIX).SrcFrom = LTblA(i).FromIX                    '
      LTblB(LTblBIX).SrcTo = LTblA(i).ToIX                        '
      LTblB(LTblBIX).SrcRepeat = LTblA(i).Repeat                  '
      LTblB(LTblBIX).SrcFlag = LTblA(i).CmdFlag                   '
      LTblB(LTblBIX).SrcVar = LTblA(i).CmdVar                     '
      LTblB(LTblBIX).DstCmd = LTblA(i).Cmd                        '
      LTblB(LTblBIX).DstFrom = LTblA(i).FromIX                    '
      LTblB(LTblBIX).DstTo = LTblA(i).ToIX                        '
      LTblB(LTblBIX).DstRepeat = LTblA(i).Repeat                  '
      LTblB(LTblBIX).DstFlag = LTblA(i).CmdFlag                   '
      LTblB(LTblBIX).DstVar = LTblA(i).CmdVar                     '

      IF UUCASE(LEFT$(LTblB(LTblBIX).SrcCmd, 1))  = "D" THEN      ' A Delete command?
         IF LTblAIX > (i + d) THEN                                ' More line commands to go?
            IF LTblA(i + 1).FromIX <= LTblB(LTblBIX).SrcTo THEN   ' Within range of the delete?
               me.CurSetReq(%LError, LTblA(i).FromIX, 0, %True)   '
               MErrExit(%eFail, "Line command at line - " + RIGHT$(LLNumG(LTblA(i).FromIX), ENV.LinNoSize) + " overlaps subsequent line commands")
            END IF                                                '
         END IF                                                   '
      END IF                                                      '

      IF BIT(LTblB(LTblBIX).SrcFlag, %lCmdBSlash) AND LTblBIX > 1 THEN ' A \ modified command and a prior one built?
         me.CurSetReq(%LError, LTblA(i).ToIX, 0, %True)           '
         MErrExit(%eFail, "Other line commands preceed a \ modified line Cmd. at line - " + RIGHT$(LLNumG(LTblA(i).ToIX), ENV.LinNoSize))
      END IF                                                      '

      IF BIT(LTblB(LTblBIX).SrcFlag, %lCmdFSlash) AND LTblAIX > (i + d) THEN ' A / modified command and more to go?
         me.CurSetReq(%LError, LTblA(i).FromIX, 0, %True)         '
         MErrExit(%eFail, "Other line commands follow a / modified line Cmd. at line - " + RIGHT$(LLNumG(LTblA(i).FromIX), ENV.LinNoSize))
      END IF                                                      '

      FOR j = 1 TO TCtr                                           ' Remove now from Touched table
         t = me.TTblGet(j)                                        ' Get a touch entry
         IF t.LinNo = LTblA(i).FromIX OR t.LinNo = LTblA(i).ToIXOrig THEN ' Is this our entry?
            IF ISFALSE BIT(t.LinFlag, %LCmdR) THEN _              ' If not an R (retain) line command, kill the entry
               t.LinNo = 0: t.LinCmd = $BlankLno: t.LinCtl = $BlankLno
            me.TTblSet(j, t)                                      ' Stuff back the updated entry
         END IF                                                   '
      NEXT j                                                      '

      me.UpdLControl(LTblA(i).FromIX)                             ' Remove from the LLCtl area
      me.UpdLControl(LTblA(i).ToIXOrig)                           '
      ARRAY DELETE LTblA(i)                                       ' Delete this entry
      INCR d: DECR i                                              ' Count deleted items, keep index on this entry
   NEXT i                                                         ' Loop
   LTblAIX -= d                                                   ' Adjust LTblAIX by number processed
   MExit
END METHOD

METHOD LCtlLookUp(Ent AS STRING) AS LONG
'---------- Lookup Line Control Command
LOCAL i, x, scrollOK AS LONG, mCmd, aCmd,zCmd, IsMac, lcmd, MMode AS STRING
   MEntry
   i = LCmdT.GetCmdIX(Ent)                                        ' Can we find the command?
   IF i THEN                                                      ' Found it
      METHOD = i                                                  ' Return the Index
   ELSE                                                           '
      mCmd = Ent                                                  ' Let's see if a possible macro
      MMode = sSetTable("GET", "MACROMODE." + mCmd)               ' See if a forced Macro Mode
      IF LEFT$(MMode, 1) = "0" THEN                               ' Yes, something's there
         MMode = UUCASE(MID$(MMode, 2))                           ' Extract the answer
         IsMac = IIF$(MMode = "BLOCK", "B", "S")                  ' Set based on BLOCK or not

      ELSEIF LEN(mCmd) >= 2 THEN                                  ' At least 2 chars?
         IF RIGHT$(mCmd, 1) = MID$(mCmd, LEN(mCmd) - 1, 1) THEN   ' Last two characters are equal?  (XYY)
            mCmd = LEFT$(mCmd, LEN(mCmd) - 1): IsMac = "B"        ' This is also a block mode XY macro
         ELSE                                                     ' Just a SINGLE type
            IsMac = "S"                                           '
         END IF                                                   '

      ELSE                                                        ' Not a block mode command
         IsMac = "S"                                              ' Set SINGLE
      END IF

      IF IsMacro(mCmd) THEN                                       ' Now see if it even exists?
         '----- See if pCommand is a scrollable one
         me.SetSubst                                              ' Do SET substitution
         lcmd = UUCASE(pCommand)                                  ' Get command line
         lcmd = sGetWord(lcmd, %NoStrip, %QuoteNotSig)            ' Get the 1st word
         IF ISNOTNULL(TRIM$(lCmd)) THEN                           ' Got something?
            x = PCmdT.GetCmdIX(lCmd)                              ' Get the command table index?
            IF x > 0 THEN                                         '
               IF PCmdT.IsScrollOK(x) THEN _                      ' Command allows scrolling
                  ScrollOK = %True                                ' Remember that
            END IF                                                '
         ELSE                                                     ' Check the ALIAS way
            aCmd = UUCASE(sSetTable("GET", "ALIAS." + lcmd))      ' See if an alias SET symbol
            IF VAL(LEFT$(aCmd, 1)) = 0 THEN                       ' Found?
               aCmd = MID$(aCmd, 2)                               ' Strip off the return code
               zCmd = sGetWord(aCmd, %NoStrip, %QuoteNotSig)      ' Get the 1st word of the Alias string
               x = PCmdT.GetCmdIX(zCmd)                           ' Get the command table index?
               IF x > 0 THEN                                      '
               IF PCmdT.IsScrollOK(x) THEN _                      ' Command allows scrolling
                  ScrollOK = %True                                ' Remember that
               END IF                                             '
            END IF                                                '
         END IF                                                   '

         IF TRIM$(pCommand) = "" OR TRIM$(pCommand) = mCmd  OR ISTRUE ScrollOK THEN ' But only if a blank command line or other half of pair
            IF IsMac = "S" THEN METHOD = 1                        ' Pass back prototype to a SINGLE entry
            IF IsMac = "B" THEN METHOD = 2                        '   "    "      "     "  " BLOCK  entry
            IF ISFALSE ScrollOK THEN                              '
               pCommand = mCmd                                    ' Stuff macro name in command line
               Ent = mCmd                                         ' Replace caller's name
            END IF                                                '
         ELSE                                                     ' A Prim Command is present
            scError(8, "Line macro not allowed with a primary command") ' Issue error
            METHOD = 0                                            '
         END IF                                                   '
      ELSE                                                        ' Not a .MACRO file
         METHOD = 0                                               ' Return 0
      END IF                                                      '
   END IF                                                         '
   MExit
END METHOD

METHOD  LCtlPairs()
'---------- Build LTblB from LTblA pairs
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL t AS TouchEntry

   MEntry
   FOR i = 1 TO LTblAIX                                           ' Process what's left in the TblA

      IF (LTblA(i).CmdType = "S" AND LTblA(i + 1).CmdType = "D") THEN ' Build a Source / Dest entry
         INCR LTblBIX                                             '
         LTblB(LTblBIX).SrcCmd = LTblA(i).Cmd                     '
         LTblB(LTblBIX).SrcFrom = LTblA(i).FromIX                 '
         LTblB(LTblBIX).SrcTo = LTblA(i).ToIX                     '
         LTblB(LTblBIX).SrcRepeat = LTblA(i).Repeat               '
         LTblB(LTblBIX).SrcFlag = LTblA(i).CmdFlag                '
         LTblB(LTblBIX).SrcVar = LTblA(i).CmdVar                  '
         LTblB(LTblBIX).DstCmd = LTblA(i + 1).Cmd                 '
         LTblB(LTblBIX).DstFrom = LTblA(i + 1).FromIX             '
         LTblB(LTblBIX).DstTo = LTblA(i + 1).ToIX                 '
         LTblB(LTblBIX).DstRepeat = LTblA(i + 1).Repeat           '
         LTblB(LTblBIX).DstFlag = LTblA(i + 1).CmdFlag            '
         LTblB(LTblBIX).DstVar = LTblA(i + 1).CmdVar              '


      ELSEIF (LTblA(i).CmdType = "D" AND LTblA(i + 1).CmdType = "S") THEN  ' Build a Dest / Source entry
         INCR LTblBIX                                             '
         LTblB(LTblBIX).SrcCmd = LTblA(i + 1).Cmd                 '
         LTblB(LTblBIX).SrcFrom = LTblA(i + 1).FromIX             '
         LTblB(LTblBIX).SrcTo = LTblA(i + 1).ToIX                 '
         LTblB(LTblBIX).SrcRepeat = LTblA(i + 1).Repeat           '
         LTblB(LTblBIX).SrcFlag = LTblA(i + 1).CmdFlag            '
         LTblB(LTblBIX).SrcVar = LTblA(i + 1).CmdVar              '
         LTblB(LTblBIX).DstCmd = LTblA(i).Cmd                     '
         LTblB(LTblBIX).DstFrom = LTblA(i).FromIX                 '
         LTblB(LTblBIX).DstTo = LTblA(i).ToIX                     '
         LTblB(LTblBIX).DstRepeat = LTblA(i).Repeat               '
         LTblB(LTblBIX).DstFlag = LTblA(i).CmdFlag                '
         LTblB(LTblBIX).DstVar = LTblA(i).CmdVar                  '

      END IF                                                      '

      GOSUB CheckOverlap                                          ' Source overlaps Destination?

      FOR j = 1 TO TCtr                                           ' Remove now from Touched table
         t = me.TTblGet(j)                                        ' Get a touch entry
         IF t.LinNo = LTblA(i).FromIX OR t.LinNo = LTblA(i).ToIXOrig OR _ ' Is this our entry?
            t.LinNo = LTblA(i + 1).FromIX OR t.LinNo = LTblA(i + 1).ToIXOrig THEN
            IF ISFALSE BIT(t.LinFlag, %LCmdR) THEN _              ' If not an R (retain) line command, kill the entry
               t.LinNo = 0: t.LinCmd = $BlankLno: t.LinCtl = $BlankLno
            me.TTblSet(j, t)                                      ' Stuff back the updated entry
         END IF                                                   '
      NEXT j                                                      '

      me.UpdLControl(LTblB(LTblBIX).SrcFrom)                      ' Clear from the LLCtl area
      me.UpdLControl(LTblB(LTblBIX).SrcTo)                        '
      me.UpdLControl(LTblB(LTblBIX).DstFrom)                      '
      me.UpdLControl(LTblB(LTblBIX).DstTo)                        '

      INCR i                                                      ' Incr i, we've combined two entries

   NEXT i                                                         ' Next 'A' table item
   MExitmeth

CheckOverlap:
   IF LTblB(LTblBIX).SrcFrom < LTblB(LTblBIX).DstFrom THEN        ' Moving downward in the file
      IF LTblB(LTblBIX).SrcTo > LTblB(LTblBIX).DstFrom THEN       ' Last line of source is > 1st line in dest
         GOSUB OverlapError                                       ' Go kill it
      END IF                                                      '

   ELSE                                                           ' Moving upward in the file
      IF LTblB(LTblBIX).SrcFrom < LTblB(LTblBIX).DstTo THEN       ' First line of source is < last line in dest
         GOSUB OverlapError                                       ' Go kill it
      END IF                                                      '
   END IF                                                         '
   RETURN                                                         '

OverlapError:
   me.CurSetReq(%LError, LTblB(LTblBIX).SrcFrom, 0, %True)        '
   MErrExit(%eFail, "Overlapping line ranges at line - " + RIGHT$(LLNumG(LTblB(LTblBIX).SrcFrom), ENV.LinNoSize))
   RETURN                                                         ' Never get here
END METHOD

METHOD  LCtlPreProcess()
'---------- Prelim massage of line number field
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL c AS STRING
LOCAL t AS TouchEntry

   MEntry
   FOR i = 1 TO TCtr                                              ' Loop through the Touched table
      t = me.TTblGet(i)                                           ' Get local copy of the entry
      me.LCtlPreProcessEnt(t)                                     ' Let common routine process it
      c = t.LinCtl                                                ' Stuff it in normal string
      IF INSTR(LEFT$(c, 1), ANY $Numeric) THEN                    ' Just a number?
         c = $BlankLNo: t.LinCtl = $BlankLNo                      ' blank it
      END IF                                                      '
      LLCtlS(t.LinNo) = c                                         ' Stuff it back in real LLCtl area
      me.TTblSet(i, t)                                            ' Return local copy to real table
   NEXT i
   MExit
END METHOD

METHOD  LCtlPreProcessEnt(t AS TouchEntry)
'---------- Massage a single line number field
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL numv AS LONG, Cmd, Num, c AS STRING

   MEntry
   c = TRIM$(t.LinCtl)                                            ' Get local copy of the line control data
   IF LEFT$(c, 1) = "." OR LEFT$(c, 1) = ":" THEN MExitMeth       ' Probable label or Tag? Do nothing
   c = UUCASE(c)                                                  '
   j = INSTR(c, ANY $Numeric)                                     ' Get loc of any trailing numeric
   IF j THEN                                                      ' We have a trailing number
      Cmd = LEFT$(c, j - 1)                                       ' Cmd = left part
      Num = MID$(c, j)                                            ' Num = number portion
      Numv = VAL(num)                                             ' Numv = remaining number
   ELSE                                                           ' No trailing number
      Cmd = c: Numv = 0                                           '
   END IF                                                         '

   '----- Fudge xNOTE through
   t.LinCVar = "."                                                ' Set "." to mean no entry
   IF IsEQ(MID$(Cmd, 2), "NOTE") AND VERIFY(UUCASE(LEFT$(Cmd, 1)), $Upper) = 0 THEN             ' xNOTE and x = A -> Z
      IF IsNE(LEFT$(Cmd, 1), "Z") THEN t.LinCVar = UUCASE(LEFT$(Cmd, 1)): Cmd = MID$(Cmd, 2)    ' Set the x into LinCVar if not Z
   END IF                                                         '

   IF INSTR("+-&.", RIGHT$(Num, 1)) > 0 THEN Cmd += RIGHT$(Num, 1)
   IF ISNOTNULL(Num) AND Numv = 0 THEN                            ' Specific zero?
      BIT SET  t.LinFlag, %LCmdForce0                             ' Remember the specific zero was here
   END IF
   Num = FORMAT$(Numv)                                            ' Back to a stripped string
   IF Num = "0" THEN Num = " "                                    ' Now eliminate the zero
   t.LinCtl = LSET$(Cmd + Num, 8)                                 ' Put back tidied up operand
   t.LinCmd = LSET$(Cmd, 8)                                       '
   t.LinRpt = Numv                                                '
   MExit
END METHOD

METHOD  LCtlProcess()
'---------- Start the Line Control process
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL x, nextSD, numS, numD, locS, locD AS LONG, lCmd, mCmd, aCmd, zCmd, PendType, LTblTest AS STRING, t AS TouchEntry

   MEntry
   '---------- PreProcess the Operands
   RESET LTblAIX, LTblBIX, LTblRange
   RESET LTblSRpt, LTblSFlag, LTblSFrom, LTblSTo
   RESET LTblDRpt, LTblDFlag, LTblDFrom, LTblDTo
   LTblSCmd = "": LTblDCmd = ""

   '----- Exit quickly if nothing to do
   IF TCtr = 0 THEN MExitMeth                                     ' Nothing to do

   '----- Pre Massage the Line control data
   me.LCtlPreProcess                                              ' Do basic preprocessing
   me.TTblPrune                                                   ' Prune handled entries

   '----- Handle any Tags and/or Labels
   me.LCtlTagLbl                                                  ' Go do Tag and Label processing
   IF ErrFlag > 0 THEN MExitMeth                                  ' Exit if an error was posted
   me.TTblPrune                                                   ' Prune handled entries
   IF TCtr = 0 THEN MExitMeth                                     ' Leave if nothing left to do

   '----- Validate the command as a valid command
   me.LCtlValidate                                                ' Go validate the entries
   IF ErrFlag > 0 THEN MExitMeth                                  ' Exit if an error was posted
   me.TTblPrune                                                   ' Prune handled entries
   IF TCtr = 0 THEN MExitMeth                                     ' Leave if nothing left to do

   '----- Do the TblA build
   me.LCtlTblA                                                    ' Go build TblA
   me.TTblPrune                                                   ' Prune handled entries
   IF ErrFlag > 0 OR LTblAIX = 0 THEN MExitMeth                   ' Exit if an error or no TblA built

   '----- See if a Primary Command Source range specified
   LTblTest = " ": RESET NumS, NumD, locS, locD                   ' Clear eligibility

   '----- First see if any eligible Primary command
   me.SetSubst                                                    ' Do SET substitution
   lcmd = UUCASE(pCommand)                                        ' Get command line
   lcmd = sGetWord(lcmd, %NoStrip, %QuoteNotSig)                  ' Get the 1st word
   IF ISNOTNULL(TRIM$(lCmd)) THEN                                 ' Got something?
      x = PCmdT.GetCmdIX(lcmd)                                    ' Get the command table index?
      IF x > 0 THEN                                               '
         IF PCmdT.GetName(x) = "FIND" OR _                        ' Temporarily cripple FIND/NFIND from
            PCmdT.GetName(x) = "F" OR _                           ' using line ranges marked via line commands
            PCmdT.GetName(x) = "NFIND" OR _                       '
            PCmdT.GetName(x) = "NF" THEN                          '
               '
         ELSE
            IF PCmdT.GetLinOps(x) = "S" OR PCmdT.GetLinOps(x) = "D" THEN
               LTblTest = PCmdT.GetLinOps(x)                      ' Save S or D
               IF PCmdT.IsDualOK(x) THEN LTblTest = "Y"           ' Override with Dual if allowed
            END IF                                                '
         END IF                                                   '
      ELSE
         aCmd = UUCASE(sSetTable("GET", "ALIAS." + lcmd))         ' See if an alias SET symbol
         IF VAL(LEFT$(aCmd, 1)) = 0 THEN                          ' Found?
            aCmd = MID$(aCmd, 2)                                  ' Strip off the return code
            zCmd = sGetWord(aCmd, %NoStrip, %QuoteNotSig)         ' Get the 1st word of the Alias string
            x = PCmdT.GetCmdIX(zcmd)                              ' Get the command table index?
            IF x > 0 THEN                                         '
               IF PCmdT.GetLinOps(x) = "S" OR PCmdT.GetLinOps(x) = "D" THEN
                  LTblTest = PCmdT.GetLinOps(x)                   ' Save S or D
                  IF PCmdT.IsDualOK(x) THEN LTblTest = "Y"        ' Override with Dual if allowed
               END IF                                             '
            ELSE                                                  '
               IF IsMacro(zcmd) THEN LTblTest = "Y"               ' Flag Macro candidates
            END IF                                                '

         ELSE                                                     '
            IF IsMacro(lcmd) THEN LTblTest = "Y"                  ' Flag Macro candidates
         END IF                                                   '
      END IF                                                      '

      '----- If LTblTest then possibility of line operands
      IF LTblTest <> " " THEN                                     ' Further checking now needed
         FOR i = 1 TO LTblAIX                                     ' Search LTBlA
            IF LTblA(i).CmdType = "S" THEN INCR NumS: locS = i    ' Count and remember Source items
            IF LTblA(i).CmdType = "D" THEN INCR NumD: locD = i    ' Count and remember Dest items
         NEXT i                                                   '
         IF NumS < 2 AND NumD < 2 THEN                            ' No more than one Source or Dest?

            '----- Just a Source entry?
            IF NumS = 1 AND NumD = 0 THEN                         ' Just a Source then?
               IF LTblTest = "S" OR LTblTest = "Y" THEN           ' Match the Primary command?
                  LTblSCmd = LTblA(locS).Cmd                      ' Save the line command
                  LTblSRpt = LTblA(locS).Repeat                   ' and repeat
                  LTblSFlag = LTblA(locS).CmdFlag                 ' and the flag
                  LTblSFrom = LTblA(locS).FromIX                  ' Line range
                  LTblSTo = LTblA(locS).ToIX                      '
                  IF ISFALSE BIT(LTblA(locS).CmdFlag, %lCmdR) THEN me.TTblDel(LTblA(locS).FromIX)
                  IF ISFALSE BIT(LTblA(locS).CmdFlag, %lCmdR) THEN me.TTblDel(LTblA(locS).ToIX)
                  me.UpdLControl(LTblA(locS).FromIX)              ' Do the line control
                  me.UpdLControl(LTblA(locS).ToIX)                '
                  LTblDCmd = "        "                           ' Reset the destination fields
                  LTblDRpt = 0                                    '
                  LTblDFlag = 0                                   '
                  LTblDFrom = 0                                   '
                  LTblDTo = 0                                     '
                  ARRAY DELETE LTblA(locS)                        ' Delete this entry
                  LTblAIX -= 1                                    ' Adjust LTblAIX
                  LTblRange = %True                               ' Set the 'we have a line operand' flag
               END IF                                             '

            '----- Just a Destination entry?
            ELSEIF NumS = 0 AND NumD = 1 THEN                     ' Just a Destination then?
               IF LTblTest = "D" OR LTblTest = "Y" THEN           ' Match the Primary command?
                  LTblDCmd = LTblA(locD).Cmd                      ' Save the line command
                  LTblDRpt = LTblA(locD).Repeat                   ' and repeat
                  LTblDFlag = LTblA(locD).CmdFlag                 ' and the flag
                  IF LTblA(locD).Cmd = "A   " THEN                ' Simple A ?
                     LTblDFrom = LTblA(locD).FromIX               ' Line range
                     LTblDTo = LTblA(locD).FromIX                 '
                  ELSEIF LTblA(locD).Cmd = "B   " THEN            ' Simple B ?
                     LTblDFrom = LTblA(locD).FromIX - 1           ' Line range
                     LTblDTo = LTblA(locD).FromIX - 1             '
                  ELSE                                            ' Must be a block command
                     LTblDFrom = LTblA(locD).FromIX               ' Line range
                     LTblDTo = LTblA(locD).ToIX                   '
                  END IF                                          '
                  IF ISFALSE BIT(LTblA(locD).CmdFlag, %lCmdR) THEN me.TTblDel(LTblA(locD).FromIX)
                  IF ISFALSE BIT(LTblA(locD).CmdFlag, %lCmdR) THEN me.TTblDel(LTblA(locD).ToIX)
                  me.UpdLControl(LTblA(locD).FromIX)              ' Do the line control
                  me.UpdLControl(LTblA(locD).ToIX)                '
                  LTblSCmd = "        "                           ' Reset the destination fields
                  LTblSRpt = 0                                    '
                  LTblSFlag = 0                                   '
                  LTblSFrom = 0                                   '
                  LTblSTo = 0                                     '
                  ARRAY DELETE LTblA(locD)                        ' Delete this entry
                  LTblAIX -= 1                                    ' Adjust LTblAIX
                  LTblRange = %True                               ' Set the 'we have a line operand' flag
               END IF                                             '

            '----- Both Source and Destination entry?
            ELSEIF NumS = 1 AND NumD = 1 THEN                     ' Both source and Destination then?
               IF LTblTest = "Y" THEN                             ' Match the Primary command?
                  LTblSCmd = LTblA(locS).Cmd                      ' Save the line command
                  LTblSRpt = LTblA(locS).Repeat                   ' and repeat
                  LTblSFlag = LTblA(locS).CmdFlag                 ' and the flag
                  LTblSFrom = LTblA(locS).FromIX                  ' Line range
                  LTblSTo = LTblA(locS).ToIX                      '
                  IF ISFALSE BIT(LTblA(locS).CmdFlag, %lCmdR) THEN me.TTblDel(LTblA(locS).FromIX)
                  IF ISFALSE BIT(LTblA(locS).CmdFlag, %lCmdR) THEN me.TTblDel(LTblA(locS).ToIX)
                  me.UpdLControl(LTblA(locS).FromIX)              ' Do the line control
                  me.UpdLControl(LTblA(locS).ToIX)                '
                  IF ISFALSE BIT(LTblA(locD).CmdFlag, %lCmdR) THEN me.TTblDel(LTblA(locD).FromIX)
                  IF ISFALSE BIT(LTblA(locD).CmdFlag, %lCmdR) THEN me.TTblDel(LTblA(locD).ToIX)
                  me.UpdLControl(LTblA(locD).FromIX)              ' Do the line control
                  me.UpdLControl(LTblA(locD).ToIX)                '
                  LTblDCmd = LTblA(locD).Cmd                      ' Save the line command
                  LTblDRpt = LTblA(locD).Repeat                   ' and repeat
                  LTblDFlag = LTblA(locD).CmdFlag                 ' and the flag
                  LTblDFrom = LTblA(locD).FromIX                  ' Line range
                  LTblDTo = LTblA(locD).ToIX                      '
                  IF locS < locD THEN                             ' Delete TblA items
                     ARRAY DELETE LTblA(locS)                     '
                     ARRAY DELETE LTblA(locD - 1)                 '
                  ELSE                                            '
                     ARRAY DELETE LTblA(locD)                     '
                     ARRAY DELETE LTblA(locS - 1)                 '
                  END IF                                          '
                  LTblAIX -= 2                                    ' Adjust LTblAIX
                  LTblRange = %True                               ' Set the 'we have a line operand' flag
               END IF                                             '
            END IF                                                '
         END IF                                                   '
      END IF                                                      '
   END IF                                                         '

   me.LCtlImmediate                                               ' Go Extract the Immediate commands
   me.TTblPrune                                                   ' Prune handled entries

   IF ErrFlag > 0 OR LTblAIX = 0 THEN MExitMeth                   ' Exit if an error was posted or no more A table entries

   FOR i = 1 TO LTblAIX                                           ' See if any still in Pending status
      nextSD = i + 1                                              ' Set NextSD to point at 'next' line command

      IF LTblA(i).CmdType = "S" AND LTblA(NextSD).CmdType = "D" THEN  ' Pair a Source/Dest pair
         i = nextSD                                               ' Continue after the matching pair

      ELSEIF LTblA(i).CmdType = "D" AND LTblA(NextSD).CmdType = "S" THEN  ' Pair a Dest/Source pair
         i = nextSD                                               ' Continue after the matching pair

      ELSEIF LTblA(i).CmdType = "S" AND LTblA(NextSD).CmdType = " " THEN  ' Source waiting for Dest?
         PendType = "S"                                           ' Remember that

      ELSEIF LTblA(i).CmdType = "D" AND LTblA(NextSD).CmdType = " " THEN  ' Dest waiting for Source?
         PendType = "D"                                           ' Remember that

      ELSE                                                        '
         IF LTblA(i).CmdIX < 3 OR LTblA(NextSD).CmdIX < 3 THEN                 ' A macro involved somehow?
            IF (LTblA(i).CmdIX = 1 AND LTblA(NextSD).CmdIX = 1) OR _           ' Both the same type (single/block)?
               (LTblA(i).CmdIX = 2 AND LTblA(NextSD).CmdIX = 2) THEN           '
               scError(%ePending, "Multiple line-command macros are not supported") ' We don't do that
            ELSEIF (LTblA(i).CmdIX < 3 AND LTblA(NextSD).CmdIX < 3) AND _      ' Both macros?
                   (LTblA(i).CmdIX <> LTblA(NextSD).CmdIX) THEN                ' but different types
               scError(%ePending, "Single line and Block mode line macros cannot be used together") ' We don't do that
            ELSE                                                               ' Macro plus some other source
               scError(%ePending, "Macro line commands cannot mix with built-in source line commands") ' We don't do that
            END IF                                                '
         ELSE                                                     '
            scError(%ePending, "Illogical line command grouping") ' We're really confused
         END IF                                                   '
         me.CurSetReq(%LError, LTblA(i).FromIX, 0, %False)        '
         MExitMeth                                                '
      END IF                                                      '
   NEXT i                                                         '

   IF ISNOTNULL(PendType) THEN                                    ' A pending condition?
      IF PendType = "D" THEN _                                    ' A Dest waiting
         MErrExit(%ePending, "Pending Source line range command") ' No, we're also Pending then

      IF PendType = "S" THEN _                                    ' A pending line command source
         MErrExit(%ePending, "Pending Destination line range command")  ' No, we're also Pending then
   END IF                                                         '

   '----- We have valid pairs, go build them
   me.LCtlPairs                                                   '
   me.TTblPrune                                                   ' Prune handled entries
   MExit
END METHOD

METHOD  LCtlTagLbl()
'---------- Process the Tags and Labels
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL c, errMsg AS STRING, k AS LONG
LOCAL t AS TouchEntry

   MEntry
   FOR i = 1 TO TCtr                                              ' Loop through the Touched table
      t = me.TTblGet(i)                                           ' Get local copy of Touch entry

      j = t.LinNo                                                 ' Get the text line number
      c = t.LinCtl                                                ' Get a local copy of the line area data

      me.LCtlTagLblEnt(t, errMsg)                                 ' Process the individual entry

      IF errMsg = "OK" THEN                                       ' Simply ignored?
         '                                                        ' Then we will too
      ELSEIF errMsg <> "NULLED" THEN                              ' Nulled successfully?
         GOSUB DoErrMsg                                           ' No, go issue message and EXIT the SUB
      END IF                                                      '

      me.TTblSet(i, t)                                            ' Return local copy to real table
      me.UpdLControl(j)                                           ' Format the LinCtl area properly

   NEXT i                                                         '
   MExitMeth

DoErrMsg:
   me.CurSetReq(%LError, j, 0, %True)                             ' Set the cursor
   MErrExit(%eFail, errMsg)                                       ' Issue errMsg
   RETURN                                                         ' Never reach here
END METHOD

METHOD  LCtlTagLblEnt(t AS TouchEntry, errMsg AS STRING)
'---------- Process a single line Tags or Label entry
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL c AS STRING, k AS LONG

   MEntry

   j = t.LinNo                                                    ' Get the text line number
   c = t.LinCtl                                                   ' Get a local copy of the line area data

   IF LEFT$(c, 1) = "." THEN                                      ' A Label was entered

      IF ISTRUE (LFlagG(j) AND %NonLineNum) THEN                  ' Watch for special lines
         errMsg = "Illegal line for a 'Dot' label"                ' Set text msg
         MExitMeth                                                ' Exit with the error
      END IF                                                      '

      IF TRIM$(c) <> REMOVE$(c, " ") OR _                         ' Ensure no crappy chars
         INSTR($Numeric, MID$(c, 2, 1)) <> 0 OR _                 '
         INSTR(c, ANY "'," + $DQ) THEN                            '
         errMsg = "Illegal characters in 'Dot' label"             ' Set text msg
         MExitMeth                                                ' Exit with the error
      END IF                                                      '
      c = LSET$(TRIM$(UUCASE(c)), 8)                              ' Uppercase it, make full size
      t.LinCtl = c                                                ' Stuff it back

      IF TRIM$(c) = "." THEN                                      ' A clear request?
         LLblS(j) = $BlankLNo                                     ' Blank the Label
         GOSUB DoUnTouch                                          ' Say we've handled this one

      ELSEIF LEFT$(c, 2) = ".." THEN                              ' A label toggle request

         IF MID$(c, 3) = "" THEN                                  ' No label included?
            GOSUB DoUnTouch                                       ' Say we've handled this one

         ELSEIF L(j).LLbl = $BlankLNo THEN                        ' If none presently, set this one
            k = me.LLCtlScan(MID$(c, 2) + " ")                    ' See if label exists already
            IF k THEN                                             ' It exists elsewhere
               LLblS(k - 1) = $BlankLNo                           ' Blank label elsewhere if used
               me.UpdLControl(k - 1)                              ' Redo LLCtl
            END IF                                                '
            LLblS(j) = MID$(c, 2) + " "                           ' Add the Label

         ELSE                                                     ' Got a current label
            IF L(j).LLbl <> MID$(c, 2) + " " THEN                 ' And different?
               LLblS(j) = MID$(c, 2) + " "                        ' Swap to this label
            ELSE                                                  ' Same label, toggle it off
               LLblS(j) = $BlankLNo                               ' Blank it
            END IF                                                '
         END IF                                                   '
         GOSUB DoUnTouch                                          ' Say we've handled this one

      ELSEIF LEFT$(L(j).LLbl, 1) <> "." THEN                      ' No existing Label present
         k = me.LLCtlScan(c)                                      ' See if label exists already
         IF k THEN                                                '
            LLblS(k - 1) = $BlankLNo                              ' Blank label elsewhere if used
            me.UpdLControl(k - 1)                                 ' Redo LLCtl
         END IF                                                   '
         LLblS(j) = c                                             ' Add the Label
         GOSUB DoUnTouch                                          ' Say we've handled this one

      ELSEIF LEFT$(L(j).LLbl, 1) = "." AND _                      ' Are both dotted?
         c <> L(j).LLbl THEN                                      ' And different?
         k = me.LLCtlScan(c)                                      ' Scan to see if there
         IF k THEN                                                '
            LLblS(k - 1) = $BlankLNo                              ' Blank label elsewhere if used
            me.UpdLControl(k - 1)                                 ' Redo LLCtl
         END IF                                                   '
         LLblS(j) = c                                             ' Add the Label
         GOSUB DoUnTouch                                          ' Say we've handled this one

      ELSEIF LEFT$(L(j).LLbl, 1) = "." AND _                      ' Are both dotted?
         c = L(j).LLbl THEN                                       ' And the same?
         GOSUB DoUnTouch                                          ' Say we've handled this one
      END IF                                                      '

   ELSEIF LEFT$(c, 1) = ":" THEN                                  ' A Tag has been entered

      IF ISTRUE (LFlagG(j) AND %NonLineNum) THEN                  ' Watch for special lines
         errMsg = "Illegal line for a 'Tag' label"                ' Set text msg
         MExitMeth                                                ' Exit with the error
      END IF                                                      '

      IF TRIM$(c) <> REMOVE$(c, " ") OR _                         ' Ensure no crappy chars
         INSTR($Numeric, MID$(c, 2, 1)) <> 0 OR _                 '
         INSTR(c, ANY "'," + $DQ) THEN                            '
         errMsg = "Illegal characters in 'Tag' label"             ' Set text msg
         MExitMeth                                                ' Exit with the error
      END IF                                                      '

      IF IsEQ(TRIM$(c), ":ZALL") THEN                             ' Reserved label?
         errMsg = "Illegal reserved 'Tag' label"                  ' Set text msg
         MExitMeth                                                ' Exit with the error
      END IF                                                      '
      c = LSET$(TRIM$(UUCASE(c)), 8)                              ' Uppercase it and fill it
      t.LinCtl = c                                                ' Stuff it back

      IF TRIM$(c) = ":" THEN                                      ' A tag removal?
         LTagS(j) = $BlankLNo                                     ' Blank the Tag
         GOSUB DoUnTouch                                          ' Say we've handled this one

      ELSEIF LEFT$(c, 2) = "::" THEN                              ' A tag toggle request

         me.CurSetReq(%LError, j, 0, %True)                       '
         IF MID$(c, 3) = "" THEN                                  ' No tagname?
            GOSUB DoUnTouch                                       ' Say we've handled this one

         ELSEIF LTagG(j) = $BlankLNo THEN                         ' If none presently, set this one
            LTagS(j) = MID$(c, 2) + " "                           ' Add the Tag
            GOSUB DoUnTouch                                       ' Say we've handled this one

         ELSE                                                     ' Got a current label
            IF LTagG(j) <> MID$(c, 2) + " " THEN                  ' And different?
               LTagS(j) = MID$(c, 2) + " "                        ' Swap to this label
            ELSE                                                  ' Same label, toggle it off
               LTagS(j) = $BlankLNo                               ' Blank it
            END IF                                                '
            GOSUB DoUnTouch                                       ' Say we've handled this one
         END IF                                                   '

      ELSE                                                        '
         LTagS(j) = c                                             ' Add the Tag
         GOSUB DoUnTouch                                          ' Say we've handled this one
      END IF                                                      '

   ELSEIF LEFT$(c, 1)  = " " THEN                                 ' A Label / Tag being removed
      IF LEFT$(L(j).LLbl, 1) = "." THEN                           ' A Label is present, we'll remove it
         LLblS(j) = $BlankLNo                                     ' Blank the Label

      ELSEIF LEFT$(LTagG(j), 1) = ":" THEN                        ' A Tag is present, we'll remove it
         LTagS(j) = $BlankLNo                                     ' Blank the Tag
      END IF                                                      '
      GOSUB DoUnTouch                                             ' Say we've handled this one
   ELSE                                                           ' Not touched by this routine
      errMsg = "OK"                                               ' Just say OK
   END IF                                                         '
   MExitmeth

DoUnTouch:
   t.LinNo = 0                                                    ' Null in the Touched table
   errMsg = "NULLED"                                              ' Set error message to NULLED
   OnUndoFlag                                                     ' Call for an Undo Save
   RETURN
END METHOD

METHOD  LCtlTblA()
'---------- Build TblA now
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL c, cmd, errMsg AS STRING, k, m, n AS LONG, Lkup AS STRING * 8, w AS LCtlScan
LOCAL t, pt AS TouchEntry

   MEntry
   LTblAIX = 0                                                    ' Reset # in table and loop control
   LTblA(0).Cmd = "        "                                      ' Setup dummy (0) entry
   LTblA(0).FromIX = -1                                           '
   LTblA(0).ToIX = -1                                             '
   LTblA(0).Repeat = -1                                           '
   LTblA(0).CmdType = " "                                         '
   LTblA(0).CmdFlag = 0                                           '
   LTblA(0).CmdIX = 0                                             '
   LTblA(0).CmdVar = " "                                          '
   LTblA(1).Cmd = "        "                                      ' Setup dummy (1) entry
   LTblA(1).FromIX = -1                                           '
   LTblA(1).ToIX = -1                                             '
   LTblA(1).Repeat = -1                                           '
   LTblA(1).CmdType = " "                                         '
   LTblA(1).CmdFlag = 0                                           '
   LTblA(1).CmdIX = 0                                             '
   LTblA(1).CmdVar = " "                                          '

   me.TTblSort                                                    ' Get array in line number order

   FOR i = 1 TO TCtr                                              ' Loop through the Touched table

      t = me.TTblGet(i)                                           ' Get working copy of Touch entry
      j = t.LinNo                                                 ' Get the text line number
      c = t.LinCtl                                                ' Get a local copy of the line area data

      k = INSTR(c, ANY " 0123456789")                             ' Parse out and validate the command
      cmd = UUCASE(LEFT$(c, k - 1))                               ' Extract the command

      w.cmd = t.LinCmd                                            ' Build the w. entry
      IF Env.LinRepeatMax > 0 AND t.LinRpt > ENV.LinRepeatMax THEN' Ensure not over max repeat value
         errMsg = "Repeat value exceeds maximum allowable - line " + RIGHT$(LLNumG(j), ENV.LinNoSize)
         GOSUB DoErrMsg                                           ' Issue message and EXIT the SUB
      END IF
      w.Repeat = t.LinRpt                                         ' Save numeric operand
      w.CmdType = t.LinCType                                      ' Save cmd type
      w.CmdFlag = t.LinFlag                                       ' Save cmd flag
      w.CmdIX = t.LinCmdIX                                        ' Save cmd table index
      w.CmdVar = t.LinCVar                                        ' Save cmd variation

      '----- A Non-Range command
      IF ISFALSE LCmdT.IsRangeOK(t.LinCmdIX) THEN                 ' If a non-Range Command
         IF LTblA(LTblAIX).ToIX = 0 THEN                          ' Is there a prev cmd waiting for a 'To'?
            errMsg = "Previous range command incomplete - line " + RIGHT$(LLNumG(j), ENV.LinNoSize)
            GOSUB DoErrMsg                                        ' Issue message and EXIT the SUB
         END IF                                                   '

         w.FromIX = j                                             ' Set start range
         w.ToIX = j                                               ' ToIX ditto
         w.ToIXOrig = j                                           ' ToIXOrig for safekeeping
         IF ISFALSE LCmdT.IsRFactorOK(t.LinCmdIX) AND w.Repeat > 1 THEN  ' If Repeat # is actually a line count, process it
            m = MAX(1, w.Repeat)                                  ' Get requested # lines
            w.Repeat = 0                                          ' No longer a repeat, used as a count
            m = MIN(m, LastLine - j)                              ' Prevent going past LastLine
            n = j                                                 ' Point at 1st line

            DO WHILE m > 0                                        ' Loop for requested count
               IF IsLXclude(n) THEN                               ' If an XX 'd range, add its size
                  m -= (LWrk1G(n) + 1)                            '
                  n += (LWrk1G(n) + 1)                            '
               ELSE                                               '
                  INCR n                                          ' Else just add 1
               END IF                                             '
               DECR m                                             ' Loop control
            LOOP                                                  '

            w.ToIX = n - 1                                        ' Stuff back adjusted 'To' line number
            IF w.ToIX > LastLine - 1 THEN w.ToIX = LastLine - 1   '
         END IF                                                   '

         IF IsLXclude(W.ToIX) THEN                                ' If an XX 'd range, add its size
            w.ToIX += LWrk1G(w.ToIX)                              '
         END IF                                                   '

         IF ISTRUE BIT(w.CmdFlag, %lCmdBSlash) THEN               ' A \ modified command?
            w.FromIX = 2                                          ' Set From to 1st data line
            w.ToIX = j                                            ' Set To = this line
         END IF                                                   '

         IF ISTRUE BIT(w.CmdFlag, %lCmdFSlash) THEN               ' A / modified command?
            w.FromIX = j                                          ' Set From to this line
            w.ToIX = LastLine - 1                                 ' Set To = the last line
         END IF                                                   '

         INCR LTblAIX                                             ' Add to the A table
         LTblA(LTblAIX) = w                                       '
         LTblA(LTblAIX + 1).CmdType = " "                         ' Set a blank type in next entry
         LTblA(LTblAIX + 1).ToIX = -1                             '
         me.TTblSet(i, t)                                         ' Save it back

         ITERATE FOR                                              ' Bail out, we're done for now
      END IF                                                      '

      '----- A Range command
      IF LTblA(LTblAIX).ToIX = 0 THEN                             ' It's a Range Cmnd, is prev cmd waiting for a 'To'?

         IF LCmdT.GetEditCode(LTblA(LTblAIX).CmdIX) <> LCmdT.GetEditCode(w.CmdIX) THEN ' Then it had better be the same command
            errMsg = "Pending command(s)": j = 0                  ' No?  It's an error, null j for no cursor set
            GOSUB DoErrMsg                                        ' Issue message and EXIT the SUB
         END IF                                                   '

         IF LTblA(LTblAIX).Repeat <> 0 AND w.Repeat <> 0 THEN     ' If they both have repeat #s
            IF LTblA(LTblAIX).Repeat <> w.Repeat  THEN            ' They'd better be the same
               errMsg = "Different numeric operands on 1 range - line " + RIGHT$(LLNumG(j), ENV.LinNoSize)
               GOSUB DoErrMsg                                     ' Issue message and EXIT the SUB
            END IF                                                '
         END IF                                                   '

         LTblA(LTblAIX).ToIX = j                                  ' Validated, complete the range
         LTblA(LTblAIX).ToIXOrig = j                              ' Safe keep it
         IF IsLXclude(j) THEN _                                   ' If XX range, add # XX'd to reach end of range
            LTblA(LTblAIX).ToIX = j + LWrk1G(j)                   '

         IF LTblA(LTblAIX).Repeat = 0 AND w.Repeat > 0 THEN LTblA(LTblAIX).Repeat = w.Repeat

         IF LTblA(LTblAIX).CmdFlag <> 0 AND w.CmdFlag <> 0 AND LTblA(LTblAIX).CmdFlag <> w.CmdFlag THEN
            errMsg = "Different K/+/- line modifiers on 1 range - line " + RIGHT$(LLNumG(j), ENV.LinNoSize)
            GOSUB DoErrMsg                                        ' Issue message and EXIT the SUB
         END IF                                                   '

         LTblA(LTblAIX).CmdFlag = LTblA(LTblAIX).CmdFlag OR w.CmdFlag  ' Merge flags  ????
         LTblA(LTblAIX + 1).CmdType = " "                         ' Set a blank type in next entry
         pt = me.TTblGet(i - 1)                                   ' Now re-get the previous waiting entry
         IF BIT(pt.LinFlag, %LCmdR) OR BIT(t.LinFlag, %LCmdR) THEN' If either is an R (retain) line command
            BIT SET pt.LinFlag, %LCmdR                            ' Make them both R
            BIT SET  t.LinFlag, %LCmdR                            '
         END IF
         me.TTblSet(i, t)                                         ' Save them back
         me.TTblSet(i - 1, pt)                                    '

         ITERATE FOR                                              '

      ELSE                                                        ' Range with no previous waiting
         w.FromIX = j: w.ToIX = 0: w.ToIXOrig = 0                 ' Add this as a new one, waiting for "To"
         INCR LTblAIX                                             ' Add to the A table
         LTblA(LTblAIX) = w                                       '
         LTblA(LTblAIX + 1).CmdType = " "                         ' Set a blank type in next entry
      END IF                                                      '

   NEXT i                                                         ' Keep Looping

   IF LTblA(LTblAIX).ToIX = 0 THEN                                ' Last item still waiting?
      errMsg = "Pending command(s)": j = 0                        ' No, we're also Pending then. Null j for no cursor set
      GOSUB DoErrMsg                                              ' Issue message and EXIT the SUB                                                    '
   END IF                                                         '
   MExitmeth

DoErrMsg:
   IF j THEN me.CurSetReq(%LError, j, 0, %True)                   ' Set the cursor if j valid
   MErrExit(%ePending, errMsg)                                    ' Issue errMsg
   MExitMeth                                                      ' Bail out
   RETURN                                                         ' Never reach here

END METHOD

METHOD  LCtlValidate()
'---------- Validate the line commands
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL c, cmd, errMsg AS STRING, k, m, n, NumK, NumSlsh, NumPM, NumDotted AS LONG, Lkup AS STRING * 8
LOCAL t AS TouchEntry

   MEntry
   FOR i = 1 TO TCtr                                              ' Loop through the Touched table

      t = me.TTblGet(i)                                           ' Get a local copy of Touch entry
      j = t.LinNo                                                 ' Get the text line number
      me.LCtlValidateEnt(t, errMsg)                               ' Go validate

      IF errMsg = "NULLED" THEN                                   ' Nulled Cmd
         me.UpdLControl(i)                                        ' Reset the line number field
      ELSEIF errMsg <> "OK" THEN                                  ' Error flagged?
         GOSUB DoErrMsg                                           ' Issue message and EXIT the SUB
      END IF                                                      '

      me.TTblSet(i, t)                                            ' Save work entry back
   NEXT i                                                         '
   MExitMeth

DoErrMsg:
   IF j THEN me.CurSetReq(%LError, j, 0, %True)                   ' Set the cursor if j valid
   MErrExit(%eFail, errMsg)                                       ' Issue errMsg
   RETURN                                                         ' Never reach here
END METHOD

METHOD  LCtlValidateEnt(t AS TouchEntry, errMsg AS STRING)
'---------- Validate the line commands
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL c, cmd, modlist AS STRING, k, m, n, NumR, NumSlsh, NumPM, NumDotted AS LONG, Lkup AS STRING

   MEntry
   j = t.LinNo                                                    ' Get the text line number
   cmd = UUCASE(t.LinCmd)                                         ' Extract the command
   RESET NumR, NumPM, NumSlsh, NumDotted                          '

   IF ISNULL(TRIM$(cmd)) THEN                                     ' Nulled Cmd
      t.LinNo = 0                                                 ' Kill the entry
      me.UpdLControl(i)                                           ' Reset the line number field
      errMsg = "NULLED"                                           ' Say we NULLED it
      MexitMeth                                                   ' We're done
   END IF                                                         '

   '----- Check out the suffix modifiers
   c = TRIM$(cmd)                                                 ' Trim it
   IF RIGHT$(c, 2) = ".." THEN REPLACE ".." WITH "\" IN c         ' Support .. alias of \
   IF RIGHT$(c, 1) = "."  THEN REPLACE "."  WITH "/" IN c         ' Support .  alias of /
   BIT RESET t.LinFlag, %lCmdNX                                   ' Reset the flags we'll fiddle with
   BIT RESET t.LinFlag, %lCmdX                                    '
   BIT RESET t.LinFlag, %lCmdBSlash                               '
   BIT RESET t.LinFlag, %lCmdFSlash                               '
   BIT RESET t.LinFlag, %lCmdR                                    '
   modlist = "+-\/&"                                              ' Set modifier list
   WHILE ISNOTNULL(c) AND INSTR(modlist, RIGHT$(c, 1)) <> 0       ' While suffixes are present
      SELECT CASE AS CONST$ RIGHT$(c, 1)                          ' Split them
         CASE "+": BIT SET t.LinFlag, %lCmdNX: INCR NumPM         ' Set Flag
         CASE "-": BIT SET t.LinFlag, %lCmdX: INCR NumPM          '
         CASE "\": BIT SET t.LinFlag, %lCmdBSlash: INCR NumSlsh   '
         CASE "/": BIT SET t.LinFlag, %lCmdFSlash: INCR NumSlsh   '
         CASE "&": BIT SET t.LinFlag, %lCmdR: INCR NumR           '
      END SELECT                                                  '
      c = LEFT$(c, LEN(c) -1)                                     ' Remove the char
   LOOP                                                           '
   cmd = LSET$(c, 8)                                              ' Swap it

   t.LinCmd = cmd                                                 ' Save stripped command
   IF (ISTRUE NumR OR ISTRUE NumPM) AND ISTRUE NumSlsh THEN       ' Both /\ and K+- ?
      errMsg = "Cannot combine \ / and & + - line modifiers on same line"
      MExitMeth                                                   ' Exit
   END IF                                                         '

   IF NumSlsh > 1 OR NumR > 1 OR NumPM > 1 OR NumDotted > 1 THEN  ' Stuttering?
      errMsg = "Multiple line command modifiers detected on the same line"
      MExitMeth                                                   ' Exit
   END IF                                                         '

   IF (NumR + NumPM + NumSlsh) > 0 AND ISNULL(TRIM$(cmd)) THEN    ' Was that just a \/ or &/+/- on the line?
      errMsg = "Line modifiers (& + - \ /) detected alone on a line"
      MExitMeth                                                   ' Exit
   END IF                                                         '

   '----- Now validate the actual command
   Lkup = TRIM$(cmd)                                              ' Copy
   IF Lkup = ""  OR Lkup = "''''''''" THEN                        ' If just trash
      t.LinNo = 0                                                 ' Kill this entry
      me.UpdLControl(i)                                           ' Reset the line number area
      errMsg = "NULLED"                                           ' Say we NULLED it
      MExitMeth                                                   ' We're done here
   END IF                                                         '

   k = me.LCtlLookup(Lkup)                                        ' Lookup what we found
   IF k = 0 THEN                                                  ' If zero, then not found
      IF IsEQ(Lkup, "ZNOTE   ") THEN                              ' Fudge message for ZNOTE
         errMsg = "ZNOTE reserved as a Primary Command keyword on line " + RIGHT$(LLNumG(j), ENV.LinNoSize)
      ELSE
         errMsg = "Unknown line command - '" + Lkup + "' on line " + RIGHT$(LLNumG(j), ENV.LinNoSize)
      END IF                                                      '
      MExitMeth                                                   ' Exit
   END IF                                                         '

   IF j = 1 AND (Lkup <> "A" AND Lkup <> "AA" AND Lkup <> "I" AND Lkup <> "N" AND Lkup <> "NOTE") THEN ' Handle Top of Data line
      errMsg = "Illegal line command for top line": j = 1         '
      MExitMeth                                                   ' Exit
   END IF                                                         '

   IF j = LastLine AND Lkup <> "B" AND Lkup <> "BB" THEN  ' Handle Bottom of Data line
      errMsg = "Illegal line command on bottom line": j = LastLine
      MExitMeth                                                   ' Exit
   END IF                                                         '

   IF IsBrowse AND ISFALSE LCmdT.IsBrowseOK(k) THEN               ' Browse mode and not valid?
      errMsg = "Illegal line command in Browse mode, line " + RIGHT$(LLNumG(j), ENV.LinNoSize)
      MExitMeth                                                   ' Exit
   END IF                                                         '

   IF t.LinRpt > 0 AND ISFALSE LCmdT.IsNumVOK(k) THEN             ' A valid number argument?
      errMsg = "Illegal numeric line operand on line " + RIGHT$(LLNumG(j), ENV.LinNoSize)
      MExitMeth                                                   ' Exit
   END IF                                                         '

   IF NumR > 0 AND ISFALSE LCmdT.IsKeepOK(k) THEN                 ' Valid K ?
      errMsg = "Cannot use & modifier on line command - line " + RIGHT$(LLNumG(j), ENV.LinNoSize)
      MExitMeth                                                   ' Exit
   END IF                                                         '

   IF NumPM > 0 AND ISFALSE LCmdT.IsPMOK(k) THEN                  ' Valid + - ?
      errMsg = "Cannot use + - modifier on line command - line " + RIGHT$(LLNumG(j), ENV.LinNoSize)
      MExitMeth                                                   ' Exit
   END IF                                                         '

   IF NumSlsh > 0 AND ISFALSE LCmdT.IsSlashOK(k) THEN             ' Valid slash ?
      errMsg = "Cannot use / \ modifier on line command - line " + RIGHT$(LLNumG(j), ENV.LinNoSize)
      MExitMeth                                                   ' Exit
   END IF                                                         '

   '----- Add command Index and Type to Touch table
   t.LinCmdIX = k                                                 ' Save Cmd Table index
   t.LinCType = LCmdT.GetType(k)                                  ' Save command type (I/S/D)

   ErrMsg = "OK"                                                  ' Say we're OK
   MExit                                                          ' Done
END METHOD
