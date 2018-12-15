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

'----- Keyboard related code included within the cObjTabData Object
'
   METHOD   FMkbBackSpace()
   '---------- FM Mode BackSpace key
      MEntry
      SELECT CASE AS LONG Zone
         CASE %CursCmnd                                           ' In pCommand?
            IF CCol > 11 THEN                                     ' Room to BKSP?
               pCommand = STRDELETE$(pCommand, ZLCol - 1, 1) + " "' Delete it
               me.WindowCmd                                       ' Re-do the command line
               me.CsrColSub(1)                                    ' Adjust cursor location
            END IF

         CASE %CursScrl                                           ' In Scroll field
            IF CCol > 21 + pCmdLen THEN                           ' Room to BKSP?
               Prf.Scroll = STRDELETE$(Prf.Scroll, ZLCol - 1, 1) + " " ' Delete it
               sPrint (LSET$(Prf.Scroll, 4), $$TxtLo, 1, 21 + pCmdLen)    ' Redisplay the entire string
               me.CsrColSub(1)                                    ' Adjust cursor location
            END IF                                                '

         CASE %CursEmpty
            CmdEmpty = " "                                        ' Just blank it
            sPrint (" ", $$TxtLo,CRow, CCol)                      '

         CASE %CursRecent
            CmdRecent = " "                                       ' Just blank it
            sPrint (" ", $$TxtLo, CRow, CCol)                     '

         CASE %CursFound
            CmdFound = " "                                        ' Just blank it
            sPrint (" ", $$TxtLo, CRow, CCol)                     '

         CASE %CursOpen
            CmdOpen = " "                                         ' Just blank it
            sPrint (" ", $$TxtLo, CRow, CCol)                     '

         CASE %CursFavorite
            CmdFavorite = " "                                     ' Just blank it
            sPrint (" ", $$TxtLo, CRow, CCol)                     '

         CASE %CursFileList
            CmdFileList = " "                                     ' Just blank it
            sPrint (" ", $$TxtLo, CRow, CCol)                     '

         CASE %CursRPaths
            CmdRPaths = " "                                       ' Just blank it
            sPrint (" ", $$TxtLo, CRow, CCol)                     '

         CASE %CursProfiles
            CmdProfiles = " "                                     ' Just blank it
            sPrint (" ", $$TxtLo, CRow, CCol)                     '

         CASE %CursPath                                           ' In Path
            IF CCol > FM_Path_Left THEN                           ' Room to BKSP?
               FPath = STRDELETE$(FPath, ZLCol - 1, 1) + " "      ' Delete it
               sPrint (FPath, $$TxtLo, CRow, FM_Path_Left)        ' Redisplay the entire string
               me.CsrColSub(1)                                    ' Adjust cursor location
            END IF                                                '

         CASE %CursMask                                           ' In Mask
            IF CCol > FM_Mask_Left THEN                           ' Room to BKSP?
               FMask = STRDELETE$(FMask, ZLCol - 1, 1) + " "      ' Delete it
               sPrint (FMask, $$TxtLo, CRow, FM_Mask_Left)        ' Redisplay the entire string
               me.CsrColSub(1)                                    ' Adjust cursor location
            END IF                                                '

         CASE %CursLCmd                                           ' In Line Cmd Area
            IF ZLCol > 1 THEN                                     ' Room to BKSP?
               AFList(ZLNum).Cmd  = STRDELETE$(AFList(ZLNum).Cmd, ZLCol - 1, 1) ' Delete it
               IF AFList(ZLNum).CmdOff > 0 THEN                   ' Room to BKSP within offset?
                  AFList(ZLNum).CmdOff = AFList(ZLNum).CmdOff - 1 '
                  me.CsrColAdd(0)                                 ' To trigger CsrMode processing
               ELSE                                               ' Else no offset available
                  me.CsrColSub(1)                                 ' Adjust cursor location
               END IF                                             '
               sPrint (me.WindowLCmd(ZLNum), $$TxtHi, CRow, 1)    ' Redisplay the entire string
            END IF                                                '

         CASE %CursLNote                                          ' In Line Note Area
            IF ZLCol > 1 THEN                                     ' Room to BKSP?
               AFList(ZLNum).Note = STRDELETE$(AFList(ZLNum).Note, ZLCol - 1, 1) ' Delete it
               IF AFList(ZLNum).NoteOff > 0 THEN                  ' Room to BKSP within offset?
                  AFList(ZLNum).NoteOff = AFList(ZLNum).NoteOff - 1    '
                  me.CsrColAdd(0)                                 ' To trigger CsrMode processing
               ELSE                                               ' Else no offset available
                  me.CsrColSub(1)                                 ' Adjust cursor location
               END IF                                             '
               sPrint (me.WindowNote(ZLNum), $$TxtHi, CRow, FM_Head_Note_Left) ' Redisplay the entire string
               me.NoteUpdate(TRIM$(AFList(ZLNum).Path) + TRIM$(AFList(ZLNum).FD.FileName), TRIM$(AFList(ZLNum).Note))

            END IF                                                '
      END SELECT                                                  '
      MExit
   END METHOD

   METHOD   FMkbBackTab()
   '---------- FM Mode BackTab key
      MEntry
      SELECT CASE AS LONG CRow                                    ' First base things on the row
         CASE 1
            SELECT CASE AS LONG CCol                              ' Now base on the column
               CASE > 22 + pCmdLen:       SetScrl                 ' Far right, go to Scroll
               CASE ELSE:                 me.CsrCol = 11          ' All else goes to Home
            END SELECT                                            '

         CASE 2                                                   ' Dash Line
            SetScrl                                               ' Go to Scroll field

         CASE FM_Quick_Line_1                                     ' Quick Line?
            SELECT CASE AS LONG CCol                              ' Now base on the column
               CASE > FM_Quick_Pos_9: me.CsrCol = FM_Quick_Pos_9  ' Past Ninth quick Entry
               CASE > FM_Quick_Pos_8: me.CsrCol = FM_Quick_Pos_8  ' Past Eighth quick Entry
               CASE > FM_Quick_Pos_7: me.CsrCol = FM_Quick_Pos_7  ' Past Seventh quick Entry
               CASE > FM_Quick_Pos_6: me.CsrCol = FM_Quick_Pos_6  ' Past Sixth quick Entry
               CASE > FM_Quick_Pos_5: me.CsrCol = FM_Quick_Pos_5  ' Past Fifth quick Entry
               CASE > FM_Quick_Pos_4: me.CsrCol = FM_Quick_Pos_4  ' Past Fourth quick Entry
               CASE > FM_Quick_Pos_3: me.CsrCol = FM_Quick_Pos_3  ' Past Third quick Entry
               CASE > FM_Quick_Pos_2: me.CsrCol = FM_Quick_Pos_2  ' Past Second quick Entry
               CASE > FM_Quick_Pos_1: me.CsrCol = FM_Quick_Pos_1  ' Past First quick Entry
               CASE ELSE: SetScrl                                 ' Back to Scroll field
            END SELECT                                            '

         CASE FM_Path_Line                                        ' Def Path Line
            IF CCol < FM_Path_Left THEN                           ' In left header area
               me.CsrRow = FM_Quick_Line_1: me.CsrCol = FM_Quick_Pos_9 ' Go to last Quick entry
            ELSE                                                  ' Else
               me.CsrRow = FM_Quick_Line_1: me.CsrCol = FM_Quick_Pos_9 ' Go to last Quick entry
            END IF                                                '

         CASE FM_Mask_Line                                        ' Mask Line
            IF ISNOTNULL(FileListNm) THEN                         ' FILELIST
               me.CsrRow = FM_Quick_Line_1: me.CsrCol = FM_Quick_Pos_9 ' Go to last Quick entry
            ELSE                                                  '
               me.CsrRow = FM_Path_Line: me.CsrCol = FM_Path_Left ' Top File line for FileLists
            END IF                                                '

         CASE FM_Head_Line                                        '
            IF ISNOTNULL(FileListNm) THEN                         ' FILELIST
               me.CsrRow = FM_Quick_Line_1: me.CsrCol = FM_Quick_Pos_9 ' Go to last Quick entry
            ELSE                                                  '
               me.CsrRow = FM_Mask_Line: me.CsrCol = FM_Mask_Left ' Top File line for FileLists
            END IF                                                '

         CASE > FM_Top_File_Line                                  ' Lower than 1st file line?
            IF CCol > FM_Head_Note_Left AND FM_Head_Note_Left > 0 THEN ' Right of an actual Note column
               IF FMode = %FMFileList THEN                        ' If a valid Note field
                  me.CsrCol = FM_Head_Note_Left                   ' Backup to the Note field
               ELSE                                               ' Else
                  me.CsrCol = 1                                   ' Ignore Note field
               END IF                                             '
            ELSEIF CCol > 1 THEN                                  ' Within the line
               me.CsrCol = 1                                      ' Just col 1
            ELSE                                                  '
               me.CsrRowSub(1)                                    ' Simple DECR
               IF FM_Head_Note_Left > 0 THEN                      ' Is there an actual Note column
                  me.CsrCol = FM_Head_Note_Left                   ' Go to it
               ELSE                                               '
                  me.CsrCol = 1                                   '
               END IF                                             '
            END IF                                                '

         CASE = FM_Top_File_Line                                  '
            IF CCol > FM_Head_Note_Left AND FM_Head_Note_Left > 0 THEN ' Right of an actual Note column
               IF FMode = %FMFileList THEN                        ' If a valid Note field
                  me.CsrCol = FM_Head_Note_Left                   ' Backup to the Note field
               ELSE                                               ' Else
                  me.CsrCol = 1                                   ' Ignore Note field
               END IF                                             '
            ELSEIF CCol > 1 THEN                                  ' Within the line
               me.CsrCol = 1                                      ' Just col 1
            ELSE                                                  '
               IF ISNOTNULL(FileListNm) THEN                      ' FILELIST
                  me.CsrRow = FM_Quick_Line_1: me.CsrCol = FM_Quick_Pos_9 ' Go to last Quick entry
               ELSE                                               '
                  me.CsrRow = FM_Mask_Line: me.CsrCol = FM_Mask_Left ' Top File line for FileLists
               END IF                                             '
            END IF                                                '

      END SELECT                                                  '
      me.MarkLineNumb                                             ' Refresh the line number
      MExit
   END METHOD

   METHOD   FMkbChar()
   '---------- FM Mode normal character
      MEntry

      SELECT CASE AS LONG Zone

         CASE %CursCmnd
            IF IsNotNsrtFlag THEN                                 ' Normal non-Insert?
               me.pCmdCharRep(ZLCol, gKeyChr)                     ' Just overlay the Key into the string
            ELSE                                                  ' Insert mode
               pCommand = STRINSERT$(pCommand, gKeyChr, ZLCol)    ' Insert the Key
            END IF                                                '
            IF ZCol = pCmdLen THEN                                ' In the last position?
               me.COffsetAdd(1)                                   ' Shift a bit
               me.CsrColAdd(0)                                    ' To trigger CsrMode update
            ELSE                                                  '
               me.CsrColAdd(1)                                    ' Adjust cursor
            END IF                                                '
            me.WindowCmd                                          ' Re-do the command line

         CASE %CursScrl
            IF IsNotNsrtFlag THEN                                 ' Normal non-Insert?
               me.ProfScrlRep(ZLCol, gKeyChr)                     ' Just overlay the Key into the string
            ELSE                                                  '
               Prf.Scroll = LEFT$(STRINSERT$(Prf.Scroll, gKeyChr, ZLCol), 4) ' Insert the Key
            END IF                                                '
            sPrint (LSET$(Prf.Scroll, 4), $$TxtLo, 1, 21 + pCmdLen) ' Print the whole string again
            me.CsrColAdd(1)                                       ' Adjust cursor

         CASE %CursEmpty
            CmdEmpty = gKeyChr                                    ' Just stuff it in
            sPrint (gKeyChr, $$TxtLo, CRow, CCol)                 '
            me.FMkbTab                                            ' Tab to next field

         CASE %CursFilePath
            CmdFilePath = gKeyChr                                 ' Just stuff it in
            sPrint (gKeyChr, $$TxtLo, CRow, CCol)                 '
            me.FMkbTab                                            ' Tab to next field

         CASE %CursRecent
            CmdRecent = gKeyChr                                   ' Just stuff it in
            sPrint (gKeyChr, $$TxtLo, CRow, CCol)                 '
            me.FMkbTab                                            ' Tab to next field

         CASE %CursFound
            CmdFound = gKeyChr                                    ' Just stuff it in
            sPrint (gKeyChr, $$TxtLo, CRow, CCol)                 '
            me.FMkbTab                                            ' Tab to next field

         CASE %CursOpen
            CmdOpen = gKeyChr                                     ' Just stuff it in
            sPrint (gKeyChr, $$TxtLo, CRow, CCol)                 '
            me.FMkbTab                                            ' Tab to next field

         CASE %CursFavorite
            CmdFavorite = gKeyChr                                 ' Just stuff it in
            sPrint (gKeyChr, $$TxtLo, CRow, CCol)                 '
            me.FMkbTab                                            ' Tab to next field

         CASE %CursFileList
            CmdFileList = gKeyChr                                 ' Just stuff it in
            sPrint (gKeyChr, $$TxtLo, CRow, CCol)                 '
            me.FMkbTab                                            ' Tab to next field

         CASE %CursRPaths
            CmdRPaths = gKeyChr                                   ' Just stuff it in
            sPrint (gKeyChr, $$TxtLo, CRow, CCol)                 '
            me.FMkbTab                                            ' Tab to next field

         CASE %CursProfiles
            CmdProfiles = gKeyChr                                 ' Just stuff it in
            sPrint (gKeyChr, $$TxtLo, CRow, CCol)                 '
            me.FMkbTab                                            ' Tab to next field

         CASE %CursPath
            IF LEN(FPath) < ZLCol THEN FPath = LSET$(FPath, ZLCol)' Lengthen field if needed
            IF IsNotNsrtFlag THEN                                 ' Normal non-Insert?
               MID$(FPath, ZLCol, 1) = gKeyChr                    ' Just overlay the Key into the string
            ELSE                                                  '
               FPath = STRINSERT$(FPath, gKeyChr, ZLCol)          ' Insert the key
            END IF                                                '
            sPrint (FPath, $$TxtLo, CRow, FM_Path_Left)           ' Print the whole string again
            me.CsrColAdd(1)                                       ' Adjust cursor

         CASE %CursMask
            IF LEN(FMask) < ZLCol THEN FMask = LSET$(FMask, ZLCol)' Lengthen field if needed
            IF IsNotNsrtFlag THEN                                 ' Normal non-Insert?
               MID$(FMask, ZLCol, 1) = gKeyChr                    ' Just overlay the Key into the string
            ELSE                                                  '
               FMask = STRINSERT$(FMask, gKeyChr, ZLCol)          ' Insert the key
            END IF                                                '
            sPrint (FMask, $$TxtLo, CRow, FM_Mask_Left)           ' Print the whole string again
            me.CsrColAdd(1)                                       ' Adjust cursor

         CASE %CursLCmd
            IF LEN(TRIM$(AFList(ZLNum).Cmd)) < ZLCol THEN AFList(ZLNum).Cmd = LSET$(AFList(ZLNum).Cmd, ZLCol)  ' Lengthen field if needed
            IF IsNotNsrtFlag THEN                                 ' Normal non-Insert?
               MID$(AFList(ZLNum).Cmd, ZLCol, 1) = gKeyChr        ' Stuff in the character
            ELSE                                                  '
               AFList(ZLNum).Cmd = STRINSERT$(AFList(ZLNum).Cmd, gKeyChr, ZLCol)  ' Insert the key
            END IF                                                '
            IF ZCol = ENV.FMLCmdWidth THEN                        ' In the last position?
               AFList(ZLNum).CmdOff = AFList(ZLNum).CmdOff + 1    ' Bump offset
               me.CsrColAdd(0)                                    ' To trigger CsrMode
            ELSE                                                  '
               me.CsrColAdd(1)                                    ' Adjust cursor
            END IF                                                '
            sPrint(me.WindowLCmd(ZLNum), $$TxtHi, CRow, 1)        '

         CASE %CursLNote
            IF LEN(TRIM$(AFList(ZLNum).Note)) < ZLCol THEN AFList(ZLNum).Note = LSET$(AFList(ZLNum).Note, ZLCol)  ' Lengthen field if needed
            IF IsNotNsrtFlag THEN                                 ' Normal non-Insert?
               MID$(AFList(ZLNum).Note, ZLCol, 1) = gKeyChr       ' Stuff in the character
            ELSE                                                  '
               AFList(ZLNum).Note = STRINSERT$(AFList(ZLNum).Note, gKeyChr, ZLCol)  ' Insert the key
            END IF                                                '
            IF ZCol = FM_Note_Size THEN                           ' In the last position?
               AFList(ZLNum).NoteOff = AFList(ZLNum).NoteOff + 1  ' Bump offset
               me.CsrColAdd(0)                                    ' To trigger CsrMode
            ELSE                                                  '
               me.CsrColAdd(1)                                    ' Adjust cursor
            END IF                                                '
            sPrint(me.WindowNote(ZLNum), $$TxtHi, CRow, FM_Head_Note_Left)'
            me.NoteUpdate(TRIM$(AFList(ZLNum).Path) + TRIM$(AFList(ZLNum).FD.FileName), TRIM$(AFList(ZLNum).Note))

         CASE ELSE                                                '
            sDoBeep                                               ' Beep on errors
      END SELECT
      MExit
   END METHOD

   METHOD   FMkbClipName()
   '---------- Stuff file name into the clipboard
   LOCAL ln AS STRING, j AS LONG
      MEntry
      IF ISFALSE IsCLLin AND ISFALSE IsCLCmd AND ISFALSE IsCLNote THEN sDoBeep: MExitMeth ' Exit if not in line area
      IF CRow >= FM_Top_File_Line + AFCount - TopScrn THEN sDoBeep: MExitMeth ' Exit if below last line
      ln = TRIM$(AFList(ZLNum).FD.FileName)                       '
      j = INSTR(-1, ln, "\")                                      ' Find last \
      IF j THEN ln = MID$(ln, j + 1)                              ' Reduce to just filename
      sWinclip_set(ln)                                            ' Send it to the Clipboard
      MExit
   END METHOD

   METHOD   FMkbClipPath()
   '---------- Stuff full file path into the clipboard
   LOCAL ln AS STRING, j AS LONG
      MEntry
      IF ISFALSE IsCLLin AND ISFALSE IsCLCmd AND ISFALSE IsCLNote THEN sDoBeep: MExitMeth ' Exit if not in line area
      IF CRow >= FM_Top_File_Line + AFCount - TopScrn THEN sDoBeep: MExitMeth ' Exit if below last line
      ln = RTRIM$(AFList(ZLNum).Path) + RTRIM$(AFList(ZLNum).FD.FileName)
      sWinclip_set(ln)                                            ' Send it to the Clipboard
      MExit
   END METHOD

   METHOD   FMkbCopyLCmd()
   '---------- FM mode Delete key
   LOCAL CBD AS STRING                                            ' new name of ClipBoard Data
      MEntry
      IF CRow < FM_Top_File_Line OR CRow >= FM_Top_File_Line + AFCount - TopScrn THEN     ' In valid area?
         CBD = " "                                                ' Clear the clipboard
         sWriteClipboard(CBD)                                     '
      ELSE                                                        '
         CBD = TRIM$(AFList(ZLNum).Cmd)                           ' Copy line command
         IF ISNULL(CBD) THEN CBD = " "                            '
         sWriteClipboard(CBD)                                     '
      END IF                                                      '
      MExit
   END METHOD

   METHOD   FMkbDelete()
   '---------- FM mode Delete key
      MEntry
      SELECT CASE AS LONG Zone
         CASE %CursCmnd                                           ' In pCommand?
            pCommand = STRDELETE$(pCommand, ZLCol, 1) + " "       ' Delete it
            me.WindowCmd                                          ' Re-do the command line

         CASE %CursScrl                                           ' In Scroll field
            Prf.Scroll = STRDELETE$(Prf.Scroll, ZLCol, 1) + " "   ' Delete it
            sPrint (LSET$(Prf.Scroll, 4), $$TxtLo, 1, 21 + pCmdLen) ' Redisplay the entire string

         CASE %CursEmpty
            CmdEmpty = " "                                        ' Just blank it
            sPrint (" ", $$TxtLo, CRow, CCol)                     '

         CASE %CursRecent
            CmdRecent = " "                                       ' Just blank it
            sPrint (" ", $$TxtLo, CRow, CCol)                     '

         CASE %CursFound
            CmdFound = " "                                        ' Just blank it
            sPrint (" ", $$TxtLo, CRow, CCol)                     '

         CASE %CursOpen
            CmdOpen = " "                                         ' Just blank it
            sPrint (" ", $$TxtLo, CRow, CCol)                     '

         CASE %CursFavorite
            CmdFavorite = " "                                     ' Just blank it
            sPrint (" ", $$TxtLo, CRow, CCol)                     '

         CASE %CursFileList
            CmdFileList = " "                                     ' Just blank it
            sPrint (" ", $$TxtLo, CRow, CCol)                     '

         CASE %CursRPaths
            CmdRPaths = " "                                       ' Just blank it
            sPrint (" ", $$TxtLo, CRow, CCol)                     '

         CASE %CursProfiles
            CmdProfiles = " "                                     ' Just blank it
            sPrint (" ", $$TxtLo, CRow, CCol)                     '

         CASE %CursPath                                           ' In Path
            FPath = STRDELETE$(FPath, ZLCol, 1) + " "             ' Delete it
            sPrint (FPath, $$TxtLo, CRow, FM_Path_Left)           ' Redisplay the entire string

         CASE %CursMask                                           ' In Mask
            FMask = STRDELETE$(FMask, ZLCol, 1) + " "             ' Delete it
            sPrint (FMask, $$TxtLo, CRow, FM_Mask_Left)           ' Redisplay the entire string

         CASE %CursLCmd                                           ' In Line Cmd Area
            AFList(ZLNum).Cmd  = STRDELETE$(AFList(ZLNum).Cmd, ZLCol, 1) ' Delete it
            sPrint (me.WindowLCmd(ZLNum), $$TxtHi, CRow, 1)       ' Redisplay the entire string

         CASE %CursLNote                                          ' In Line Note Area
            AFList(ZLNum).Note = STRDELETE$(AFList(ZLNum).Note, ZLCol, 1) ' Delete it
            sPrint (me.WindowNote(ZLNum), $$TxtHi, CRow, FM_Head_Note_Left) ' Redisplay the entire string
            me.NoteUpdate(TRIM$(AFList(ZLNum).Path) + TRIM$(AFList(ZLNum).FD.FileName), TRIM$(AFList(ZLNum).Note))

      END SELECT                                                  '
      MExit
   END METHOD

   METHOD   FMkbDown()
   '---------- FM Mode Down key
      MEntry
      me.CsrRowAdd(1)                                             ' Simple INCR
      IF CRow >= FM_Top_File_Line + FM_List_Height THEN           ' Scroll the screen if needed
         TopScrn = MIN(TopScrn + 1, MAX(1, AFCount - FM_List_Height + 7)) ' Move TopScrn by scroll amount
         me.CsrRowSub(1)                                          ' Undo INCR
         LastTop = 0                                              ' Make sure it moves
         me.DispScreen                                            ' and display screen
      END IF                                                      '
      me.MarkLineNumb                                             ' Mark the line number
      MExit
   END METHOD

   METHOD   FMkbEndOfLine(LorT AS STRING)
   '---------- FM Mode EndOfLine key
   LOCAL j AS LONG
      MEntry
      SELECT CASE AS LONG Zone
         CASE %CursCmnd                                           ' In the Cmd area
            j = IIF(LorT = "L", LEN(pCommand), LEN(RTRIM$(pCommand)))  ' How long is Cmd
            IF j > pCmdLen + COffset THEN                         ' Is it to the right of the Window?
               COffset = j - (pCmdLen / 2)                        ' Adjust offset
            ELSEIF j < COffset THEN                               ' Is it to the left of the Window?
               COffset = MAX(0, j - (pCmdLen / 2))                ' Adjust offset
            END IF                                                '
            me.WindowCmd                                          ' Redisplay the line
            me.CsrCol = j - COffset + 11                          '

         CASE %CursScrl                                           ' In Scroll area
            j = IIF(LorT = "L", LEN(Prf.Scroll), LEN(RTRIM$(Prf.Scroll)))  ' How long is fieldCmd
            j = MIN(3, j)                                         ' So we don't point cursor off the right edge
            me.CsrCol = j + 21 + pCmdLen                          '

         CASE %CursEmpty, %CursRecent, %CursFound, %CursOpen, _
              %CursFavorite, %CursFileList, %CursRPaths, %CursProfiles

         CASE %CursPath                                           ' Path line
            me.CsrCol = IIF(LorT = "L", LEN(FPath), LEN(RTRIM$(FPath))) + FM_Path_Left

         CASE %CursMask                                           ' Mask Line
            me.CsrCol = IIF(LorT = "L", LEN(FMask), LEN(RTRIM$(FMask))) + FM_Mask_Left

         CASE %CursLCmd                                           ' Line command
            j = IIF(LorT = "L", LEN(AFList(ZLNum).Cmd), LEN(RTRIM$(AFList(ZLNum).Cmd))) + 1   ' Get length of LCmd
            IF j > ENV.FMLCmdWidth + AFList(ZLNum).CmdOff THEN    ' Is it to the right of the Window?
               AFList(ZLNum).CmdOff = j - ENV.FMLCmdWidth         ' Adjust offset
            ELSEIF j < AFlist(ZLNum).CmdOff THEN                  ' Is it to the left of the Window?
               AFList(ZLNum).CmdOff = MAX(0, j - ENV.FMLCmdWidth) ' Adjust offset
            END IF                                                '
            sPrint(me.WindowLCmd(ZLNum), $$LnoHi, CRow, 1)        '
            me.CsrCol = j - AFList(ZLNum).CmdOff                  '

         CASE %CursLNote                                          ' Note areaLine command
            j = IIF(LorT = "L", LEN(AFList(ZLNum).Note), LEN(RTRIM$(AFList(ZLNum).Note))) + 1 ' Get length of Note
            IF j > FM_Note_Size + AFList(ZLNum).NoteOff THEN      ' Is it to the right of the Window?
               AFList(ZLNum).NoteOff = j - FM_Note_Size           ' Adjust offset
            ELSEIF j < AFlist(ZLNum).NoteOff THEN                 ' Is it to the left of the Window?
               AFList(ZLNum).NoteOff = MAX(0, j - FM_Note_Size)   ' Adjust offset
            END IF                                                '
            sPrint(me.WindowNote(ZLNum), $$LnoHi, CRow, FM_Head_Note_Left)
            me.CsrCol = j - AFList(ZLNum).NoteOff + 1 + FM_Head_Note_Left

         CASE ELSE                                                '
            sDoBeep                                               '
      END SELECT                                                  '
      MExit
   END METHOD

   METHOD  FMkbEnter()
   '---------- Enter pressed on File Manager Tab
      DoSet(%Attention)                                           ' Just call Attention
   END METHOD

   METHOD   FMkbEraseEOL()
   '---------- FM mode EraseEOL key
   LOCAL j AS LONG
      MEntry
      SELECT CASE AS LONG Zone                                    ' Base things on the field
         CASE %CursCmnd                                           ' pCommand line
            pCommand = LEFT$(pCommand, ZLCol - 1)                 '
            me.WindowCmd                                          ' Re-do the command line

         CASE %CursScrl                                           ' Scroll field
            Prf.Scroll = LEFT$(LEFT$(Prf.Scroll, ZLCol) + "    ", 4)  ' Truncate the line then
            sPrint (LSET$(Prf.Scroll, 4), $$TxtLo, 1, 21 + pCmdLen)' Print the whole string again

         CASE %CursEmpty
            CmdEmpty = " "                                        ' Just blank it
            sPrint (" ", $$TxtLo, CRow, CCol)                     '

         CASE %CursRecent
            CmdRecent = " "                                       ' Just blank it
            sPrint (" ", $$TxtLo, CRow, CCol)                     '

         CASE %CursFound
            CmdFound = " "                                        ' Just blank it
            sPrint (" ", $$TxtLo, CRow, CCol)                     '

         CASE %CursOpen
            CmdOpen = " "                                         ' Just blank it
            sPrint (" ", $$TxtLo, CRow, CCol)                     '

         CASE %CursFavorite
            CmdFavorite = " "                                     ' Just blank it
            sPrint (" ", $$TxtLo, CRow, CCol)                     '

         CASE %CursFileList
            CmdFileList = " "                                     ' Just blank it
            sPrint (" ", $$TxtLo, CRow, CCol)                     '

         CASE %CursRPaths
            CmdRPaths = " "                                       ' Just blank it
            sPrint (" ", $$TxtLo, CRow, CCol)                     '

         CASE %CursProfiles
            CmdProfiles = " "                                     ' Just blank it
            sPrint (" ", $$TxtLo, CRow, CCol)                     '

         CASE %CursPath
            FPath = LEFT$(FPath, ZLCol - 1)                       ' Truncate the line then
            sPrint (LSET$(FPath, ENV.ScrWidth - FM_Path_Left), $$TxtLo, CRow, FM_Path_Left) ' Print the whole string again

         CASE %CursMask
            FMask = LEFT$(FMask, ZLCol - 1)                       ' Truncate the line then
            sPrint (LSET$(FMask, ENV.ScrWidth - FM_Mask_Left), $$TxtLo, CRow, FM_Mask_Left) ' Print the whole string again

         CASE %CursLCmd
            AFList(ZLNum).Cmd = LEFT$(AFList(ZLNum).Cmd, ZLCol - 1)
            sPrint (me.WindowLCmd(ZLNum), $$TxtHi, CRow, 1)        ' Redisplay the entire string

         CASE %CursLNote
            AFList(ZLNum).Note = LEFT$(AFList(ZLNum).Note, ZLCol - 1)
            sPrint (me.WindowNote(ZLNum), $$TxtHi, CRow, FM_Head_Note_Left)  ' Redisplay the entire string
            me.NoteUpdate(TRIM$(AFList(ZLNum).Path) + TRIM$(AFList(ZLNum).FD.FileName), TRIM$(AFList(ZLNum).Note))

         CASE ELSE                                                '
            sDoBeep                                               '

      END SELECT                                                  '
      MExit
   END METHOD

   METHOD  FMkbFirstLCmd()
   '---------- First Line Cmd area
      me.CsrRow = FM_Top_File_Line: me.CsrCol = 1                 ' Go to 1st File Line
   END METHOD

   METHOD   FMkbHome()
   '---------- FM Mode Home key
      MEntry
      me.CsrRow = 1: me.CsrCol = 11: COffset = 0                  ' Reset things
      me.WindowCmd                                                ' Re-do the command line in case we shifted
      me.MarkLineNumb                                             ' Refresh the line number
      MExit
   END METHOD

   METHOD   FMkbLeft()
   '---------- FM Mode Left key
      MEntry
      SELECT CASE AS LONG Zone                                    ' Basee things on the field
         CASE %CursCmnd                                           ' pCommand line
            IF ZCol = 1 AND COffset > 0 THEN                      ' Left end and offset available?
               me.COffsetAdd(-1)                                  ' Decr the offset
               me.WindowCmd                                       ' Re-do the command line
            ELSE                                                  '
               GOSUB Back1                                        ' Simple DECR
            END IF                                                '

         CASE %CursLCmd                                           ' Line command
            IF ZCol = 1 AND AFList(ZLNum).CmdOff > 0 THEN         ' If left and and offset > 0
               AFList(ZLNum).CmdOff = AFList(ZLNum).CmdOff - 1    ' Reduce the offset
               sPrint(me.WindowLCmd(ZLNum), $$LnoHi, CRow, 1)     ' Re-Display it
            ELSE                                                  '
               GOSUB Back1                                        ' Simple DECR
            END IF                                                '

         CASE %CursLNote                                          ' Line note
            IF ZCol = 1 AND AFList(ZLNum).NoteOff > 0 THEN        ' If at left end and offset > 0
               AFList(ZLNum).NoteOff = AFList(ZLNum).NoteOff - 1  ' Decr the offset
               sPrint(me.WindowNote(ZLNum), $$LnoHi, CRow, FM_Head_Note_Left) ' Re-Display it
            ELSE                                                  '
               GOSUB Back1                                        ' Simple DECR
            END IF                                                '

         CASE ELSE                                                ' Everything else
            GOSUB Back1                                           ' Simple Decr

      END SELECT                                                  '
      me.MarkLineNumb                                             ' In case line number moved
      MExitMeth
      Back1:
         me.CsrColAdd(-1)                                         ' Simple Decr
         IF CCol < 1 THEN                                         ' Column wrap?
            me.CsrCol = ENV.ScrWidth: me.CsrRowAdd(-1)            ' Yes, adjust for it
            IF CRow < 1 THEN me.CsrRow = ENV.ScrHeight - (3 * ENV.FMHelpFlag) ' also for CRow wrap if needed
         END IF                                                         '
         RETURN
   END METHOD

   METHOD   FMkbLineNo()
   '---------- LineNo key
      IF ISFALSE IsCLLin AND ISFALSE IsCLCmd AND ISFALSE IsCLNote THEN sDoBeep: MExitMeth ' Exit if not in line area
      me.CsrCol = 1                                               ' Go back to LCmd area
   END METHOD

   METHOD   FMkbNewLine()
   '---------- FM Mode NewLine key
      MEntry
      SELECT CASE AS LONG CRow                                    ' First base things on the row
         CASE 1, 2                                                ' pCommand line, dash line
            me.CsrRow = FM_Quick_Line_1: me.CsrCol = FM_Quick_Pos_1

         CASE FM_Quick_Line_1                                     ' Quick Line?
            IF ISNOTNULL(FileListNm) THEN                         ' FileList mode?
               me.CsrRow = FM_Top_File_Line: me.CsrCol = 1        ' Top File line in Named Favorite mode
            ELSE                                                  '
               me.CsrRow = FM_Path_Line: me.CsrCol = FM_Path_Left ' Go to mask line
            END IF                                                '

         CASE FM_Path_Line                                        ' Path Line
            IF ISNOTNULL(FileListNm) THEN                         ' FileList mode?
               me.CsrRow = FM_Top_File_Line: me.CsrCol = 1        ' Top File line in Named Favorite mode
            ELSE                                                  '
               me.CsrRow = FM_Mask_Line: me.CsrCol = FM_Mask_Left ' Go to mask line
            END IF                                                '

         CASE FM_Mask_Line                                        ' Def Types Line
            me.CsrRow = FM_Top_File_Line: me.CsrCol = 1           ' Go to 1st File Line

         CASE FM_Head_Line                                        '
            me.CsrRow = FM_Top_File_Line: me.CsrCol = 1           ' Go to 1st File Line

         CASE >= FM_Top_File_Line                                 ' In File list
            IF CRow >= ENV.ScrHeight - (2 * ENV.FMHelpFlag) OR _  '
               CRow >= FM_Top_File_Line - 1 + AFCount - TopScrn THEN
               me.CsrRow = 1: me.CsrCol = 11                      ' Go to top line
            ELSE                                                  '
               me.CsrRowAdd(1)                                    '
               me.CsrCol = 1                                      '
            END IF                                                '
      END SELECT                                                  '
      me.MarkLineNumb                                             ' Refresh the line number
      MExit
   END METHOD

   METHOD   FMkbPaste
   '---------- FM Mode Paste key
   LOCAL t, ln, MyDlm, Tx AS STRING, i, j AS LONG
      MEntry
      sReadClipboard(t, MyDLM, %False)                            ' Go get Clipboard, keep the file
      IF LEN(t) < 1 THEN sDoBeep: MExitMeth                       ' Error if nothing there
      ln = EXTRACT$(t, MyDLM)                                     ' Get 1st line from Clipboard
      me.MarkKill                                                 ' Kill any active block select

      SELECT CASE AS LONG Zone                                    ' Split now by what field we're on

         CASE %CursCmnd                                           ' pCommand
            Tx = pCommand                                         ' Get working copy
            GOSUB Pasteit                                         ' Paste the string
            PCommand = Tx                                         ' Copy back
            i = ZLCol + LEN(ln)                                   ' Move cursor logically
            IF i > pCmdLen + COffset THEN                         ' Is it to the right of the Window?
               COffset = i - (pCmdLen / 2)                        ' Adjust offset
            ELSEIF i < COffset THEN                               ' Is it to the left of the Window?
               COffset = MAX(0, i - (pCmdLen / 2))                ' Adjust offset
            END IF                                                '
            me.WindowCmd                                          ' Print the whole string again
            me.CsrCol = i - COffset + 10                          ' Move cursor

         CASE %CursPath
            Tx = FPath                                            ' Get working copy
            GOSUB Pasteit                                         ' Paste the string
            FPath = Tx                                            ' Copy it back
            sPrint (FPath, $$TxtLo, FM_Path_Line, FM_Path_Left)   '
            me.CsrCol = ZLCol + LEN(ln) + FM_Path_Left - 1        ' Move cursor

         CASE %CursMask
            Tx = FMask                                            ' Get working copy
            GOSUB Pasteit                                         ' Paste the string
            FMask = Tx                                            ' Copy it back
            sPrint (FMask, $$TxtLo, FM_Mask_Line, FM_Mask_Left)   '
            me.CsrCol = ZLCol + LEN(ln) + FM_Mask_Left - 1        ' Move cursor

         CASE %CursLCmd                                           ' In the Line Command area?
            Tx = AFList(ZLNum).Cmd                                ' Get working copy
            GOSUB Pasteit                                         ' Paste the string
            AFList(ZLNum).Cmd = Tx                                ' Copy back
            i = CCol + AFList(ZLNum).CmdOff + LEN(ln)             ' Move cursor logically
            IF i > ENV.FMLCmdWidth + AFList(ZLNum).CmdOff THEN    ' Is it to the right of the Window?
               AFList(ZLNum).CmdOff = i - (ENV.FMLCmdWidth / 2)   ' Adjust offset
            ELSEIF i < AFList(ZLNum).CmdOff THEN                  ' Is it to the left of the Window?
               AFList(ZLNum).CmdOff = MAX(0, i - (ENV.FMLCmdWidth / 2)) ' Adjust offset
            END IF                                                '
            sPrint(me.WindowLCmd(ZLNum), $$TxtHi, CRow, 1)        ' Re-Display it
            me.CsrCol = i - AFList(ZLNum).CmdOff                  ' Move cursor

         CASE %CursLNote                                          ' In the Note Command area?
            Tx = AFList(ZLNum).Note                               ' Get working copy
            GOSUB Pasteit                                         ' Paste the string
            AFList(ZLNum).Note = Tx                               ' Copy back
            i = ZLCol + LEN(ln)                                   ' Move cursor logically
            IF i > FM_Note_Size + AFList(ZLNum).NoteOff THEN      ' Is it to the right of the Window?
               AFList(ZLNum).NoteOff = i - (FM_Note_Size / 2)     ' Adjust offset
            ELSEIF i < AFList(ZLNum).NoteOff THEN                 ' Is it to the left of the Window?
               AFList(ZLNum).NoteOff = MAX(0, i - (FM_Note_Size / 2)) ' Adjust offset
            END IF                                                '
            sPrint(me.WindowNote(ZLNum), $$TxtHi, CRow, FM_Head_Note_Left)' Re-Display it
            me.CsrCol = i - AFList(ZLNum).NoteOff + FM_Head_Note_Left - 1 ' Move cursor
            me.NoteUpdate(TRIM$(AFList(ZLNum).Path) + TRIM$(AFList(ZLNum).FD.FileName), TRIM$(AFList(ZLNum).Note))

      END SELECT                                                  '
      MExitMeth

      PasteIt:
         IF IsNotNsrtFlag THEN                                    ' Do the non-Insert mode handling
            IF LEN(Tx) < ZLCol + LEN(ln) - 1 THEN Tx = LSET$(Tx, ZLCol + LEN(ln) - 1) ' Lengthen if needed
            MID$(Tx, ZLCol, LEN(ln)) = ln                         ' Overlay the CBD text
         ELSE                                                     ' Do the Insert method
            Tx = LEFT$(Tx, ZLCol - 1) + ln + MID$(Tx, ZLCol)      '
         END IF                                                   '
      RETURN
   END METHOD

   METHOD   FMkbRight
   '---------- FM Mode Right key
      MEntry
      SELECT CASE AS LONG Zone                                    ' Basee things on the field
         CASE %CursCmnd                                           ' pCommand line
            IF ZCol = pCmdLen THEN                                ' End of line?
               me.COffsetAdd(1)                                   ' Bump the offset
               me.WindowCmd                                       ' Re-do the command line
            ELSE                                                  '
               GOSUB Fwd1                                         ' Simple Incr
            END IF                                                '

         CASE %CursLCmd                                           ' Line command
            IF ZCol = ENV.FMLCmdWidth AND ISNOTNULL(TRIM$(AFList(ZLNum).Cmd)) THEN ' If at right bound and data present
               AFList(ZLNum).CmdOff = AFList(ZLNum).CmdOff + 1    ' Bump the offset
               sPrint(me.WindowLCmd(ZLNum), $$LnoHi, CRow, 1)     ' Re-Display it
            ELSE                                                  '
               GOSUB Fwd1                                         ' Simple Incr
            END IF                                                '

         CASE %CursLNote                                          ' Line note
            IF ZCol = FM_Note_Size AND ISNOTNULL(TRIM$(AFList(ZLNum).Note)) THEN ' If at right bound and data present
               AFList(ZLNum).NoteOff = AFList(ZLNum).NoteOff + 1  ' Bump the offset
               sPrint(me.WindowNote(ZLNum), $$LnoHi, CRow, FM_Head_Note_Left)   ' Re-Display it
            ELSE                                                  '
               GOSUB Fwd1                                         ' Simple Incr
            END IF                                                '

         CASE ELSE                                                ' Everything else
            GOSUB Fwd1                                            ' Simple INCR

      END SELECT                                                  '
      me.MarkLineNumb                                             ' In case line number moved
      MExitMeth                                                   '
      Fwd1:
         me.CsrColAdd(1)                                          ' Simple INCR
         IF CCol > ENV.ScrWidth THEN                              ' Column wrap?
            me.CsrCol = 1: me.CsrRowAdd(1)                        ' Yes, adjust for it
            IF CRow > ENV.ScrHeight + 1 - (3 * ENV.FMHelpFlag) THEN me.CsrRow = 1 ' also for Row wrap if needed
         END IF                                                   '
      RETURN
   END METHOD

   METHOD   FMkbTab()
   '---------- FM Mode Tab key
   LOCAL i, j AS LONG

      MEntry
      SELECT CASE AS LONG CRow                                    ' First base things on the row
         CASE 1                                                   ' pCommand line
            SELECT CASE AS LONG CCol                              ' Now base on the column
               CASE < 11:               me.CsrCol = 11            ' Left of Cmd, go to Cmd
               CASE < 10 + pCmdLen:     SetScrl                   ' In Cmd, go to Scroll
               CASE ELSE:                                         ' All Else
                  me.CsrRow = FM_Quick_Line_1: me.CsrCol = FM_Quick_Pos_1  ' First Quick Entry
            END SELECT                                            '

         CASE 2                                                   ' Dash Line
            me.CsrRow = FM_Quick_Line_1: me.CsrCol = FM_Quick_Pos_1  ' First Quick Entry

         CASE FM_Quick_Line_1                                     ' Quick Line?
            SELECT CASE AS LONG CCol                              ' Now base on the column
               CASE < FM_Quick_Pos_2: me.CsrCol = FM_Quick_Pos_2  ' Second quick Entry
               CASE < FM_Quick_Pos_3: me.CsrCol = FM_Quick_Pos_3  ' Third quick Entry
               CASE < FM_Quick_Pos_4: me.CsrCol = FM_Quick_Pos_4  ' Fourth quick Entry
               CASE < FM_Quick_Pos_5: me.CsrCol = FM_Quick_Pos_5  ' Fifth quick Entry
               CASE < FM_Quick_Pos_6: me.CsrCol = FM_Quick_Pos_6  ' Sixth quick Entry
               CASE < FM_Quick_Pos_7: me.CsrCol = FM_Quick_Pos_7  ' Seventh quick Entry
               CASE < FM_Quick_Pos_8: me.CsrCol = FM_Quick_Pos_8  ' Eighth quick Entry
               CASE < FM_Quick_Pos_9: me.CsrCol = FM_Quick_Pos_9  ' Ninth quick Entry
               CASE ELSE
                  IF ISNOTNULL(FileListNm) THEN                   ' FILELIST
                     me.CsrRow = FM_Top_File_Line: me.CsrCol = 1  ' Top File line for FileLists
                  ELSE                                            '
                     me.CsrRow = FM_Path_Line: me.CsrCol = FM_Path_Left ' Path Line for all else
                  END IF                                          '
            END SELECT                                            '

         CASE FM_Path_Line                                        ' Def Path Line
            IF CCol < FM_Path_Left THEN                           ' In left header area
               me.CsrCol = FM_Path_Left                           ' Go to Path
            ELSEIF ISNOTNULL(FileListNm) THEN                     ' FILELIST
               me.CsrRow = FM_Top_File_Line: me.CsrCol = 1        ' Top File line for FileLists
            ELSE                                                  '
               me.CsrRow = FM_Mask_Line: me.CsrCol = FM_Mask_Left ' Path Line for all else
            END IF                                                '

         CASE FM_Mask_Line                                        ' Mask Line
            IF CCol < FM_Mask_Left THEN                           ' In left header area
               me.CsrCol = FM_Mask_Left                           ' Go to Mask
            ELSE                                                  '
               me.CsrRow = FM_Top_File_Line: me.CsrCol = 1        ' Top File line for FileLists
            END IF                                                '

         CASE FM_Head_Line                                        '
            me.CsrRow = FM_Top_File_Line                          ' Go to 1st File Line
            me.CsrCol = 1                                         '

         CASE >= FM_Top_File_Line                                 '
            IF CRow >= ENV.ScrHeight - (2 * ENV.FMHelpFlag) OR _  '
               CRow >= FM_Top_File_Line + AFCount - TopScrn THEN  '
               me.CsrRow = 1                                      ' Go to top line
               me.CsrCol = 11                                     '
            ELSEIF CCol < FM_Head_Note_Left THEN                  ' Left of a Note column
               IF FMode = %FMFileList THEN                        ' If a valid Note field
                  me.CsrCol = FM_Head_Note_Left                   ' Skip to the Note field
               ELSE                                               '
                  me.CsrRowAdd(1)                                 '
                  me.CsrCol = 1                                   '
               END IF
            ELSE                                                  '
               me.CsrRowAdd(1)                                    '
               me.CsrCol = 1                                      '
            END IF                                                '
      END SELECT                                                  '
      me.MarkLineNumb                                             ' Refresh the line number
      MExit
   END METHOD

   METHOD   FMkbUp()
   '---------- FM Mode Up key
   LOCAL i, j AS LONG
      MEntry
      me.CsrRowSub(1)                                             ' Simple DECR
      IF CRow = 0 THEN me.CsrRow = FM_Top_File_Line + FM_List_Height - 1 ' Handle wrap if needed
      IF CRow + 1 < FM_Top_File_Line THEN MExitMeth               ' In header area
      IF CRow + 1 = FM_Top_File_Line THEN                         ' Top line see if scrolling
         IF TopScrn = 1 THEN MexitMeth                            ' No place to go, move cursor up
         TopScrn = MAX(1, TopScrn - 1)                            ' Move TopScrn
         me.CsrRowAdd(1)                                          ' Keep the cursor here
         LastTop = 0                                              ' Let scrolling proceed
         me.FMkbEnter                                             '
         MexitMeth                                                '
      END IF                                                      '
      me.MarkLineNumb                                             ' Refresh the line number
      MExit
   END METHOD


   METHOD  krANSI
   '---------- ANSI pop-up in row mode (original layout)
      me.DispANSI (%False)                                        ' Pop it open, flag = not column mode
      GRAPHIC ATTACH PgHandle, WindowID                           ' Set graphics back
   END METHOD

   METHOD  krANSIcol
   '---------- ANSI pop-up in column mode
      me.DispANSI (%True)                                         ' Pop it open, flag = column mode
      GRAPHIC ATTACH PgHandle, WindowID                           ' Set graphics back
   END METHOD


   METHOD  krBackTab()
   '---------- Tab backward To previous field
   LOCAL i, j, k, CsrLin AS LONG, lclTabs AS STRING
      OnInsClnSupp                                                ' Suppress Insert cleanup for 1 Attn pass

      SELECT CASE CRow                                            ' First base things on the row

         '----- Cursor on top line
         CASE 1                                                   ' Top row
            IF CCol < 12 THEN GOSUB BotLine: GOSUB LastTab: EXIT METHOD ' Simple one
            IF CCol > 11 AND CCol < 22 + pCmdLen THEN SetCmd: EXIT METHOD ' Another
            SetScrl                                               '

         '----- Cursor on Line 2
         CASE 2                                                   ' Line 2
            SetScrl                                               ' Back to Scroll Amt.

         '----- Cursor in the data area
         CASE ELSE                                                ' Lower data area

            '----- If on fixed COLS line, get out
            IF CRow = 3 AND Prf.Cols THEN                         ' Cols line?
               SetScrl: EXIT METHOD                               '
            END IF                                                '

            '----- In data, now look at column
            CsrLinDX = 0                                          ' Reset possible hidden DX line location
            SELECT CASE CCol                                      ' Now base it on the column

               '----- Data - Column 1
               CASE 1                                             ' In Col 1?
                  IF (CRow - 1) < 3 + Prf.Cols THEN               ' Too Far?
                     SetScrl: EXIT METHOD                         ' Yes, go to Scroll Amt.
                  ELSE                                            '
                     CsrLin = sGetIX(CRow)                        ' Get real line number

                     '----- Below the Bottom line?
                     IF CsrLin = 0 THEN                           ' Below last line?
                        GOSUB BotLine: GOSUB LastTab: EXIT METHOD ' Go to last tab on last line
                     END IF                                       '

                     IF CsrLin < 0 THEN                           ' If Hex line
                        CsrLin = sGetIX(CRow + CsrLin)            ' Get real line number of actual line
                        DECR CRow                                 ' Backup on screen
                     ELSE                                         ' Else
                        DECR CsrLin                               ' Decr real line
                        DECR CRow                                 ' Decr screen line
                     END IF
                     GOSUB LastTab                                ' go to Last Tab
                     EXIT METHOD                                  ' and leave
                  END IF                                          '

               '----- Data - LineNum area
               CASE < gLNData1                                    ' In LinNum area, just goto col 1
                  me.CsrCol = 1: EXIT METHOD                      '

               '----- Data - Actual text columns
               CASE ELSE                                          ' In real data area

                  '----- Get the text line we're on
                  CsrLin = sGetIX(CRow)                           ' Get data line index

                  '----- Below the Bottom line?
                  IF CsrLin = 0 THEN                              ' Below last line?
                     GOSUB BotLine: GOSUB LastTab: EXIT METHOD    ' Go to last tab on last line
                  END IF                                          '

                     IF CsrLin < 0 THEN                           ' If Hex line
                        CsrLin = sGetIX(CRow + CsrLin)            ' Get real line number of actual line
                     ELSE                                         ' Else
                        DECR CsrLin                               ' Decr real line
                     END IF

                  '----- Top / Bottom line?
                  IF IsLTop(CsrLin) OR IsLBottom(CsrLin) THEN     ' On the Top or bottom of Data Line
                     me.CsrCol = 1                                ' Just go to Col 1
                     EXIT METHOD                                  '
                  END IF                                          '

                  '----- If TABS OFF, we can't do much
                  IF ISFALSE Prf.Tabs THEN                        ' No TABS, just go to left margin
                     IF CCol = gLNData1 THEN me.CsrCol = 1 ELSE me.CsrCol = gLNData1 '
                     EXIT METHOD                                    '
                  END IF                                          '

                  '----- TABS ON, figure out where to go
                  i = CCol - gLNPadCol + Offset                   ' Calc index into data string
                  lclTabs = me.TabsSimple(LEN(L(CsrLin).@LTxt))   ' Get a working tabs line
                  IF i > LEN(lclTabs) THEN                        ' If right of last tab
                     GOSUB LastTab                                ' Go set to the last tab
                  ELSE                                            '
                     i = INSTR(i - LEN(lclTabs) - 2, lclTabs, "*")' look for prev tab marker
                     IF i THEN                                    ' If we have one
                        IF i < Offset + 1 OR i > Offset + gDataLen - 1 THEN ' Still on current screen
                           me.CurSetReq(%Position, j, i, %True)   ' No, Set cursor set attempt
                           DoSet(%Refresh)                        ' Have it looked at
                           EXIT METHOD                            '
                        ELSE                                      ' Still on screen, just move cursor
                           me.CsrCol = i + gLNPadCol - Offset     '
                        END IF                                    '
                     ELSE                                         ' No more to the left
                        IF CCol = gLNData1 THEN me.CsrCol = 1 ELSE me.CsrCol = gLNData1 ' Left margin if not already there
                     END IF                                       '
                  END IF                                          '
            END SELECT                                            '
      END SELECT                                                  '

      EXIT METHOD                                                   '

      '----- Find Bottom line of screen
      BotLine:                                                    ' Find Last Screen Line
         FOR i = gwScrHeight TO 3 + Prf.Cols STEP -1              '
            IF sGetIX(i) > 0 THEN                                 '
               me.CsrRow = i                                      ' Got it, exit
               CsrLin = sGetIX(CRow)                              ' Get data line index
               EXIT FOR                                           '
            END IF                                                '
         NEXT i                                                   '
      RETURN                                                      '

      '----- Find the last tab on a line
      LastTab:                                                    ' Find Last Tab
         lclTabs = me.TabsSimple(LEN(L(CsrLin).@LTxt))            ' Get a working tabs line
         IF Prf.Tabs THEN                                         ' If TABS active
            i = INSTR(-1, lclTabs, "*")                           '
         ELSE                                                     '
            i = 1                                                 '
         END IF                                                   '
         IF IsLTop(CsrLin) OR IsLBottom(CsrLin) THEN              ' On the Top or bottom of Data Line
            me.CsrCol = 1                                         ' Just go to Col 1
            RETURN                                                '
         END IF                                                   '
         IF i < Offset + 1 OR i > Offset + gDataLen - 1 THEN      ' Still on current screen
            me.CurSetReq(%Position, CsrLin, i, %True)             ' Set cursor set attempt
            DoSet(%Refresh)                                       ' Have it looked at
         ELSE                                                     '
            me.CsrCol = i + gLNPadCol - Offset                    '
         END IF                                                   '
      RETURN                                                      '
   END METHOD

   METHOD  krBPaste()
   '---------- Block PASTE Clipboard data into the dataset
   LOCAL lclTop, i, j, tcol, curline, curcol, fMIX AS LONG
   LOCAL ln, t, tt AS STRING
   LOCAL MyDLM AS STRING

      MEntry
      '----- Reset any marked area
      me.MarkKill                                                 ' Kill any active block select
      me.SwapKill                                                 '
      me.AttrInvClearFind                                         ' Also any Find HiLite
      lclTop = TopScrn                                            ' Save current top

      '----- Ensure we're in the text area, locate the text line number
      IF CCol < gLNData1 OR CRow < 3 + Prf.Cols THEN sDoBeep: MExitMeth ' Better be in text area
      tcol = CCol - gLNPadCol + Offset                            ' Calc column insert position
      i = sGetIX(CRow)                                            ' Get 1st line number
      IF i = 0 OR NOT IsLData(i) THEN sDoBeep: MExitMeth          ' Better be valid
      CsrLinDX = 0                                                ' Reset possible hidden DX line location

      '----- Fetch the Clipboard data
      sReadClipboard(t, MyDLM, %False)                            ' Go get Clipboard, keep the file

      '----- Extract Clipboard a line at a time
      DO WHILE LEN(t) > 1                                         ' Loop extracting text lines
         ln = EXTRACT$(t, MyDLM)                                  ' Get one line
         t =  REMAIN$(t, MyDlm)                                   ' Make t = the remainder of the line

         '----- Skip till eligible line, or insert one if at the end
         WHILE NOT IsLData(i)                                     ' Only paste into data lines
            IF IsLBottom(i) OR IsLFile(i) THEN                    ' Bottom line?  (or bottom of a %File group)
               me.LInsertEmpty(i - 1, 1, %Data)                   ' Request insert 1, %Data line
               IF IsMEdit THEN                                    ' In MEdit mode?
                  fMIX = me.MEditTbl("I", FORMAT$(i - 1))         ' Go get the MIX value of previous line
                  LMixS(i) = fMIX                                 ' Set into the new line
               END IF                                             '
            ELSE                                                  '
               INCR i                                             '
            END IF                                                '
         WEND                                                     '

         j = CCol - gLNPadCol + Offset                            ' Calc index into string
         tt = LTxtG(i)                                            ' Get working copy of text

         '----- Do it either Insert or non-Insert way
         IF IsNotNsrtFlag THEN                                    ' Do the non-Insert mode handling
            IF j + LEN(ln) - 1 > LEN(tt) THEN tt = LSET$(tt , j + LEN(ln) - 1)  ' Lengthen if needed
            MID$(tt, j, LEN(ln)) = ln                             ' Overlay the CBD text
         ELSE                                                     ' Do the Insert method
            tt = me.DataInsert(tt, ln, j)                         ' Insert the data
            me.LAttrAdjust(i, j, LEN(ln))                         ' Adjust any Attr hilite
         END IF                                                   '

         IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) THEN _    ' CAPS ON?
            tt = UUCASE(tt)                                       ' UC if CAPS is ON
         me.LTxtSet(i, tt)                                        ' Save it back
         me.ModSet(i)                                             ' Remember we changed something
         curline = i: curcol = tcol + LEN(ln)                     '
         INCR i                                                   ' Next line
      LOOP                                                        '

      '----- We're done, clean up and exit
      TopScrn = lclTop                                            ' Restore CRP
      me.CurSetReq(%LineCmd, curline, curcol, %True)              ' Set cursor set attempt
      DoSet(%Refresh)                                             ' Have it looked at
      MExit                                                       '
   END METHOD

   METHOD  krBS()
   '---------- Do a PC type Bksp function
   LOCAL ix, i, j, x AS LONG, lcltxt AS STRING

      MEntry
         '----- If a highlighted area, see if within its boundaries
         IF IsBrowse THEN sDoBeep: MExitMeth                      ' Not allowed in Browse mode
         IF IsMarkActive AND IsCData THEN                         ' Hi-lighted text and cursor in Data area?
            IF ZLNum >= MarkSLin AND _                            ' Is cursor within the marked area?
               ZLNum <= MarkELin AND _                            '
               ZLCol >= MarkSCol AND _                            '
               ZLCol <= MarkECol THEN                             '
               me.AttrInvClearFind                                ' Clear any Find HiLite

               IF IsNsrtFlag THEN                                 ' Do the non-Insert mode handling
                  '----- Do a char mode delete on each line
                  FOR i = MarkSLin TO MarkELin                    ' Process each line in the marked area
                     IF ISTRUE IsLData(i) OR ISTRUE IsLNote(i) THEN  ' Just data/note lines
                        IF ZLCol > MarkSCol AND ZLCol <= MarkECol + 1 THEN ' Room to BKSP?
                           me.ModSet(i)                           ' Remember we changed something
                           me.LTxtCharDel(i, ZLCol - 1)           ' Delete 1 char
                           me.LAttrAdjust(i, ZLCol - 1, -1)       ' Adjust any Attr hilite
                           IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) AND IsLData(ix) THEN  ' CAPS ON and data line?
                              me.LTxtSet(i, UUCASE(LTxtG(i)))     ' UC it
                           END IF                                 '
                           me.AttrScan(i)                         ' Recolorize
                        ELSE                                      ' Not in the block
                           GOTO BSNonInsert                       ' Do the Non-Insert version
                        END IF                                    '
                     END IF                                       '
                  NEXT i                                          '
                  DECR MarkECol                                   ' Pull in the Mark End column
                  IF MarkSCol > MarkECol THEN                     ' Nothing left?
                     me.MarkKill                                  ' Reset the select frame
                     me.SwapKill                                  '
                  END IF                                          '
                  me.CurSetReq(%Position, ZLNum, ZLCol - 1, %True)' Set cursor set attempt
                  me.CsrCol = ZLCol - 1 + Offset + gLNData1       '
                  DoSet(%Refresh)                                 ' Have it looked at
                  MExitMeth                                       '

               ELSE                                               ' Do the non-Insert mode handling
                  BSNonInsert:
                  me.MarkKill                                     ' Reset the select frame
                  me.SwapKill                                     '

                  '----- OK, now do the Delete of the marked range
                  FOR x = MarkSLin TO MarkELin                    ' Loop through the requested range
                     IF ISTRUE IsLData(x) OR ISTRUE IsLNote(x) THEN  ' Just data/note lines
                        me.ModSet(x)                              ' Remember we changed something
                        lclTxt = LTxtG(x)                         ' Get the text
                        lclTxt = STRDELETE$(lclTxt, MarkSCol, MarkECol - MarkSCol + 1) ' Delete the characters
                        me.LAttrAdjust(x, MarkSCol, -(MarkECol - MarkSCol + 1)) ' Adjust any Attr hilite
                        IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) THEN     ' CAPS ON?
                           me.LTxtSet(x, UUCASE(LTxtG(x)))        ' UC if CAPS is ON
                        END IF                                    '
                        me.LTxtSet(x, lclTxt)                     ' Put work string back
                        me.CurSetReq(%Position, MarkELin, MarkSCol, %True) ' Set cursor set attempt
                        me.CsrCol = MarkSCol + Offset + gLNPadCol '
                     END IF                                       '
                  NEXT x                                          '
                  DoSet(%Refresh)                                 ' Have it looked at
                  MExitMeth                                       '

               END IF                                             '
            END IF                                                '
            me.MarkKill                                           ' Reset the highlight
            me.SwapKill                                           '
         END IF                                                   '
                                                                  '
      '----- Bksp in the command line
      IF IsCCmnd THEN                                             ' In Command line
         IF COffset > 0 THEN                                      ' Cmd shifted?
            i = CCol - 11 + COffset                               ' Save old cursor location
            pCommand = STRDELETE$(pCommand, i, 1)                 ' Delete char to it's left
            IF CCol > 11 AND CCol < 15 THEN                       ' Left hand side?
               COffset = MAX(0, COffset - (pCmdLen / 2))          '
            END IF                                                '
            me.WindowCmd                                          ' Print the whole string again
            me.CsrCol = i - COffset + 11 - 1                      ' Move cursor left now
         ELSEIF CCol > 11 THEN                                    '
            pCommand = STRDELETE$(pCommand, CCol - 11 + COffset, 1) ' Delete it
            me.WindowCmd                                          ' Print the whole string again
            me.CsrColSub(1)                                       '
         END IF                                                   '

      '----- Bksp in the Scroll amount field
      ELSEIF IsCScrl THEN                                         ' In Scroll Amt.
         IF CCol > 21 + pCmdLen THEN                              ' Room to BKSP?
            Prf.Scroll = STRDELETE$(Prf.Scroll, CCol - 21 - pCmdLen, 1) + " " ' Delete it
            sPrint (LSET$(Prf.Scroll, 4), $$TxtLo, 1, 21 + pCmdLen) ' Redisplay the entire string
            me.CsrColSub(1)                                       ' Adjust cursor location
         END IF                                                   '

      '----- Bksp in the text data area
      ELSEIF IsCData THEN                                         ' In Data area?
         me.AttrInvClearFind                                      ' Clear any Find HiLite
         IF IsBrowse THEN sDoBeep: MExitMeth                      ' Not allowed in Browse mode
         me.MarkKill                                              ' Reset the select frame
         me.SwapKill                                              '
         IF CCol > gLNData1 THEN                                  ' Room to BKSP?
            ix = sGetIX(CRow)                                     ' Get IX of real DataLine
            IF ix = 0 THEN MExitMeth                              ' Bail out if below last line
            IF Prf.HexMode = &4 AND ix < 0 THEN sDoBeep: MExitMeth' If on a non-Char line
            CsrLinDX = 0                                          ' Reset possible hidden DX line location

            IF ISTRUE (LFlagG(ix) AND %NonTypable) THEN           ' Ignore non-typable lines
               ' Null

            ELSE                                                  '
               me.ModSet(ix)                                      ' Remember we changed something
               i = CCol - gLNData1 + Offset                       ' Calc IX into the Text
               me.LTxtCharDel(ix, i)                              ' Delete 1 char
               me.LAttrAdjust(ix, i, -1)                          ' Adjust any Attr hilite
               me.CsrColSub(1)                                    ' Adjust cursor
               IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) AND IsLData(ix) THEN  ' CAPS ON and data line?
                  me.LTxtSet(ix, UUCASE(LTxtG(ix)))               ' UC it
               END IF                                             '
               me.AttrScan(ix)                                    ' Recolorize
               me.DispLine(ix, CRow)                              ' Re-display the line

               '----- Extra stuff if in HEX mode
               IF Prf.HexMode = &4 THEN                           ' If Hex mode, do extra
                  IF CRow + 1 <= gwScrHeight THEN                 ' Make sure we don't go off the bottom
                     sPrint (me.WindowHexTop(LTxtG(ix)), $$TxtHi, CRow + 1, gLNData1) ' Print upper hex
                     me.MarkReDraw(CRow + 1)                      ' Re-draw the MARK lines
                  END IF                                          '
                  IF CRow + 2 <= gwScrHeight THEN                 ' Make sure we don't go off the bottom
                     sPrint (me.WindowHexBottom(LTxtG(ix)), $$TxtHi, CRow + 2, gLNData1) ' Print lower hex
                     me.MarkReDraw(CRow + 2)                      ' Re-draw the MARK lines
                  END IF                                          '
               END IF                                             '
            END IF                                                '
         END IF                                                   '

      '----- Bksp in the Line number area
      ELSEIF IsCLinN THEN                                         ' In the Line control area
         ix = sGetIX(CRow)                                        ' Get IX of real DataLine
         IF ix = 0 THEN MExitMeth                                 ' Bail out if below last line
         IF me.IsTouchable(ix, CCol) THEN                         ' If 1st interaction in the line number area
            LLCtlS(ix) = $BlankLNo: me.CsrCol = 1                 ' Blank the Line number
            sPrint (BYCOPY LEFT$(LLCtlG(ix), ENV.LinNoSize), $$LnoHi, CRow, 1)' ReDisplay it
            me.TTblAdd(ix, TRIM$(LLCtlG(ix)))                     ' Update touched items
         END IF                                                   '
         IF CCol > 1 THEN                                         ' Room to BKSP?
            LLCtlS(ix) = STRDELETE$(LLCtlG(ix), CCol - 1, 1) + " "' Delete a char
            me.TTblAdd(ix, TRIM$(LLCtlG(ix)))                     ' Update touched items
            sPrint (BYCOPY LEFT$(LLCtlG(ix), ENV.LinNoSize), $$LnoHi, CRow, 1)' ReDisplay it
            me.CsrColSub(1)                                       '
         END IF                                                   '
      END IF                                                      '
      MExit                                                       '
   END METHOD

   METHOD  krChar()
   '---------- Process normal character
   LOCAL i, j, k, GotColumns AS LONG, lclTxt, t AS STRING
      MEntry
      CsrLinDX = 0                                                ' Reset possible hidden DX line location

      '----- If text is hilited and this key is within it, replace the marked text
      IF IsMarkActive AND IsCData THEN                            ' Hi-lighted text and cursor in Data area?
         IF ZLNum >= MarkSLin AND _                               ' Is cursor within the marked area?
            ZLNum <= MarkELin AND _                               '
            ZLCol >= MarkSCol AND _                               '
            ZLCol <= MarkECol THEN                                '
            me.AttrInvClearFind                                   ' Clear any Find HiLite
            IF IsBrowse THEN sDoBeep: MExitMeth                   ' Not allowed in Browse mode
            IF IsNotNsrtFlag THEN                                 ' Do the non-Insert mode handling
               me.MarkKill                                        ' Reset the select frame
               me.SwapKill                                        '
               FOR i = MarkSLin TO MarkELin                       ' Process each line in the marked area
                  IF ISTRUE IsLData(i) OR ISTRUE IsLNote(i) THEN  ' Just data/note lines
                     me.ModSet(i)                                 ' Remember we changed something
                     lclTxt = LTxtG(i)                            ' Get the text
                     IF LEN(lclTxt) < MarkECol THEN _             ' Need to lengthen
                        lclTxt = LSET$(lclTxt, MarkECol)          ' Do so
                     lclTxt = STRDELETE$(lclTxt, MarkSCol, MarkECol - MarkSCol + 1) ' Delete the characters
                     me.LAttrAdjust(i, MarkSCol, -(MarkECol - MarkSCol + 1))    ' Adjust any Attr hilite
                     lclTxt = STRINSERT$(lclTxt, gKeyChr, MarkSCol) ' Insert the replacement
                     me.LAttrAdjust(i, MarkSCol, 1)               ' Adjust any Attr hilite
                     IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) THEN ' CAPS ON?
                        lclTxt = UUCASE(lclTxt)                   ' UC if CAPS is ON
                     END IF                                       '
                     me.LTxtSet(i, lclTxt)                        ' Put work string back
                     me.AttrScan(i)                               ' Recolorize
                     me.CurSetReq(%Position, MarkELin, MarkSCol + 1, %True) ' Set cursor set attempt
                     me.CsrCol = MarkSCol + Offset + gLNData1     '
                  END IF                                          '
               NEXT i                                             '
               DoSet(%Refresh)                                    ' Have it looked at
               MExitMeth                                          '

            ELSE                                                  ' This is the Insert mode version
               FOR i = MarkSLin TO MarkELin                       ' Process each line in the marked area
                  IF ISTRUE IsLData(i) OR ISTRUE IsLNote(i) THEN  ' Just data/note lines
                     me.ModSet(i)                                 ' Remember we changed something
                     lclTxt = LTxtG(i)                            ' Get the text
                     IF LEN(lclTxt) < MarkECol THEN _             ' Need to lengthen
                        lclTxt = LSET$(lclTxt, MarkECol)          ' Do so
                     lclTxt = STRINSERT$(lclTxt, gKeyChr, ZLCol)  ' Insert the character
                     me.LAttrAdjust(i, ZLcol, 1)                  ' Adjust any Attr hilite
                     IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) THEN ' CAPS ON?
                        lclTxt = UUCASE(lclTxt)                   ' UC if CAPS is ON
                     END IF                                       '
                     me.LTxtSet(i, lclTxt)                        ' Put work string back
                     me.AttrScan(i)                               ' Recolorize
                  END IF                                          '
               NEXT i                                             '
               INCR MarkECol                                      ' Push over the Mark End column
               me.CurSetReq(%Position, ZLNum, ZLCol + 1, %True)   ' Set cursor set attempt
               me.CsrCol = ZLCol + 1 + Offset + gLNData1          '
               DoSet(%Refresh)                                    ' Have it looked at
               MExitMeth                                          '
            END IF
         END IF                                                   '
      END IF                                                      '

      '----- Major split -This is the non-INSERT mode
      IF IsNotNsrtFlag THEN                                       ' Do the non-Insert mode handling

         '----- Type in the Command line
         IF IsCCmnd THEN                                          ' In Command line
            me.pCmdCharRep(ZLCol, gKeyChr)                        ' Just overlay the Key into the string
            me.WindowCmd                                          ' Just overlay the Key into the string
'           sPrint (gKeyChr, $$TxtLo, CRow, CCol)                  '
            IF ZCol = pCmdLen THEN                                ' In the last position?
               me.COffsetAdd(1)                                   ' Shift a bit
               me.CsrColAdd(0)                                    ' To trigger CsrMode update
               me.WindowCmd                                       ' Re-do the command line
            ELSE                                                  '
               me.CsrColAdd(1)                                    ' Adjust cursor
            END IF                                                '
            MExitMeth                                             ' Exit to avoid kbRight on SUB exit

         '----- Type in the Scroll amount field
         ELSEIF IsCScrl THEN                                      ' In the Scroll Amt area?
            me.ProfScrlRep(ZLCol, gKeyChr)                        ' Just overlay the Key into the string
            sPrint (gKeyChr, $$TxtLo, CRow, CCol)                 '

         '----- Type in the text data area
         ELSEIF IsCData THEN                                      ' In the Data Area?
            me.AttrInvClearFind                                   ' Clear any Find HiLite
            IF IsBrowse THEN sDoBeep: MExitMeth                   ' Not allowed in Browse mode
            me.MarkKill                                           ' Reset the select frame
            me.SwapKill                                           '
            j = S(CRow)                                           ' Get raw line index type
                                                                  '
            IF ISTRUE (LFlagG(ZLnum) AND %NonTypable) THEN        ' Ignore non-typable lines
               ' Null

            ELSE                                                  '
               t = lTxtG(ZLnum)                                   ' Get current text
               me.ModSet(ZLnum)                                   ' Remember we changed something
               IF ZLCol > LEN(L(ZLNum).@LTxt) THEN                ' If right of the end-of-text, lengthen it to reach
                  k = ZLCol - LEN(L(ZLnum).@LTxt)
                  me.LTxtSet(ZLNum, lTxtG(ZLnum) + STRING$(k, " "))
               END IF                                             '
               IF Prf.HexMode = &1 THEN                           ' If not hex mode
                  me.LTxtCharRep(ZLnum, ZLCol, gKeyChr)           ' Now overlay the Key into the string
                  IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) THEN _ ' CAPS ON?
                     me.LTxtSet(ZLnum, UUCASE(LTxtG(ZLNum)))      ' UC if CAPS is ON
                  me.AttrScan(ZLnum)                              ' Recolorize
                  me.DispLine(ZLnum, CRow)                        ' Re-display the line

               '----- Lots of work in HEX mode
               ELSE                                               '
                  SELECT CASE j                                   ' Handle differently for Char/Upper/Lower

                     '----- Upper Hex line
                     CASE -1                                      ' Upper hex
                        GOSUB HexVal                              ' Validate it
                        gKeyChr = sHexUpper(gKeyChr, MID$(LTxtG(ZLNum), ZLCol, 1)) ' Go insert the upper half
                        me.LTxtCharRep(ZLnum, ZLCol, gKeyChr)     ' Overlay the Key into the string
                        sPrint (gKeyChr, $$TxtLo, CRow - 1, CCol) '
                        me.MarkReDraw(CRow - 1)                   ' Re-draw the MARK lines
                        sPrint (me.WindowHexTop(LTxtG(ZLNum)), $$TxtHi, CRow, gLNData1) ' Print upper hex
                        me.MarkReDraw(CRow)                       ' Re-draw the MARK lines

                     '----- Lower HEX line
                     CASE -2                                      ' Lower hex
                        GOSUB HexVal                              ' Validate it
                        gKeyChr = sHexLower(gKeyChr, MID$(LTxtG(ZLNum), ZLCol, 1)) ' Go insert the lower half
                        me.LTxtCharRep(ZLnum, ZLCol, gKeyChr)     ' Overlay the Key into the string
                        sPrint (gKeyChr, $$TxtLo, CRow - 2, CCol) '
                        me.MarkReDraw(CRow - 2)                   ' Re-draw the MARK lines
                        sPrint (me.WindowHexBottom(LTxtG(ZLNum)), $$TxtHi, CRow, gLNData1) ' Print lower hex
                        me.MarkReDraw(CRow)                       ' Re-draw the MARK lines

                     '----- Dash line
                     CASE -3                                      ' Dash line
                        sDoBeep: MExitMeth                        ' do nothing

                     '----- Normal character line
                     CASE ELSE                                    ' Normal Char
                        me.LTxtCharRep(ZLnum, ZLCol, gKeyChr)     ' Overlay the Key into the string
                        me.AttrScan(ZLnum)                        ' Recolorize
                        me.DispLine(ZLnum, CRow)                  ' Re-display the line
                        IF CRow + 1 <= gwScrHeight THEN           ' Make sure we don't go off the bottom
                           sPrint (me.WindowHexTop(LTxtG(ZLNum)), $$TxtHi, CRow + 1, gLNData1) ' Print upper hex
                           me.MarkReDraw(CRow + 1)                ' Re-draw the MARK lines
                        END IF                                    '
                        IF CRow + 2 <= gwScrHeight THEN           ' Make sure we don't go off the bottom
                           sPrint (me.WindowHexBottom(LTxtG(ZLNum)), $$TxtHi, CRow + 2, gLNData1) ' Print lower hex
                           me.MarkReDraw(CRow + 2)                ' Re-draw the MARK lines
                        END IF                                    '
                  END SELECT                                      '
               END IF                                             '
            END IF                                                '

         '----- Line number area
         ELSEIF IsCLinN THEN                                      ' In the Line Number area?
            IF me.IsTouchable(ZLnum, CCol) THEN                   ' If 1st interaction in the line number area
               me.LFlagBitOff(ZLnum, %EQChange)                   ' Reset ==CHG>
               LLCtlS(ZLnum) = $BlankLNo: me.CsrCol = 1           ' Blank the Line number
               sPrint (BYCOPY LEFT$(LLCtlG(ZLnum), ENV.LinNoSize), $$LnoHi, CRow, 1) ' ReDisplay it
            END IF                                                '
            me.LLCtlCharRep(ZLnum, CCol, gKeyChr)                 ' Overlay the Key into the LineNum string
            me.TTblAdd(ZLnum, TRIM$(LLCtlG(ZLnum)))               ' Update touched items
            sPrint (gKeyChr, $$TxtLo, CRow, CCol)                 '
         ELSE                                                     '
            sDoBeep: MExitMeth                                    '
         END IF                                                   '

      '----- INSERT mode versions
      ELSE                                                        ' Do all the same stuff in Insert mode

         '----- Type Insert mode command line
         IF IsCCmnd THEN                                          ' In Command line
            IF ZLCol > LEN(pCommand) THEN                         ' If too short, lengthen it
               pCommand += SPACE$(ZLCol - LEN(pCommand))          '
            END IF                                                '
            pCommand = STRINSERT$(pCommand, gKeyChr, ZLCol)       ' Insert the Key
            IF ZCol = pCmdLen THEN                                ' In the last position?
               me.COffsetAdd(1)                                   ' Shift a bit
               me.CsrColAdd(0)                                    ' To trigger CsrMode update
            ELSE                                                  '
               me.CsrColAdd(1)                                    ' Adjust cursor
            END IF                                                '
            me.WindowCmd                                          ' Re-do the command line
            MExitMeth                                             ' Exit to avoid kbRight on SUB exit

         '----- Type Insert mode Scroll amount field
         ELSEIF IsCScrl THEN                                      ' In the Scroll Amt  area?
            Prf.Scroll = LEFT$(STRINSERT$(Prf.Scroll, gKeyChr, ZLCol), 4) ' Insert the Key
            sPrint (LSET$(Prf.Scroll, 4), $$TxtLo, 1, 21 + pCmdLen) ' Print the whole string again

         '----- Type Insert mode text data area
         ELSEIF IsCData THEN                                      ' In the Data area?
            me.AttrInvClearFind                                   ' Clear any Find HiLite
            IF IsBrowse THEN sDoBeep: MExitMeth                   ' Not allowed in Browse mode
            me.MarkKill                                           ' Reset the select frame
            me.SwapKill                                           '
            j = S(CRow)                                           ' Get raw line index type
            IF Prf.HexMode = &4 AND j < 0 THEN sDoBeep: MExitMeth ' Insert and in the Hex lines
            i = CCol - gLNPadCol + Offset                         ' Calc index into string

            IF ISTRUE (LFlagG(ZLnum) AND %NonTypable) THEN        ' Ignore non-typable lines
               ' Null

            ELSE                                                  '
               t = lTxtG(ZLnum)                                   ' Get current text
               me.ModSet(ZLnum)                                   ' Remember we changed something
               me.LTxtSet(ZLnum, me.DataInsert(LTxtG(ZLNum), gKeyChr, ZLCol))
               IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) THEN _ ' CAPS ON?
                  me.LTxtSet(ZLnum, UUCASE(LTxtG(ZLNum)))         ' UC if CAPS is ON
               me.LAttrAdjust(ZLnum, ZLCol, 1)                    ' Adjust any Attr hilite

               me.AttrScan(ZLNum)                                 ' Recolorize
               me.DispLine(ZLnum, CRow)                           ' Re-display the line

               '----- Hex mode version
               IF Prf.HexMode = &4 THEN                           ' If Hex mode, do extra
                  IF CRow + 1 <= gwScrHeight THEN                 ' Make sure we don't go off the bottom
                     sPrint (me.WindowHexTop(LTxtG(ZLNum)), $$TxtHi, CRow + 1, gLNData1) ' Print upper hex
                     me.MarkReDraw(CRow + 1)                      ' Re-draw the MARK lines
                  END IF                                          '
                  IF CRow + 2 <= gwScrHeight THEN                 ' Make sure we don't go off the bottom
                     sPrint (me.WindowHexBottom(LTxtG(ZLNum)), $$TxtHi, CRow + 2, gLNData1) ' Print lower hex
                     me.MarkReDraw(CRow + 2)                      ' Re-draw the MARK lines
                  END IF                                          '
               END IF                                             '
            END IF                                                '

         '----- Type Insert mode Line number area
         ELSEIF IsCLinN THEN                                      ' In Line Number area?
            gKeyChr = UUCASE(gKeyChr)                             ' Make all Uppercase

            IF me.IsTouchable(ZLnum, CCol) THEN                   ' If 1st interaction in the line number area
               me.LFlagBitOff(ZLnum, %EQChange)                   ' Reset ==CHG>
               LLCtlS(ZLnum) = $BlankLNo: me.CsrCol = 1           ' Blank the Line number
               sPrint (BYCOPY LEFT$(LLCtlG(ZLnum), ENV.LinNoSize), $$LnoHi, CRow, 1) ' ReDisplay it
            END IF                                                '
            LLCtlS(ZLnum) = LEFT$(STRINSERT$(LLCtlG(ZLnum), gKeyChr, CCol), 8)

            sPrint (BYCOPY LEFT$(LLCtlG(ZLnum), ENV.LinNoSize), $$LnoHi, CRow, 1)
            me.TTblAdd(ZLnum, TRIM$(LLCtlG(ZLnum)))               ' Update touched items
         ELSE                                                     '
            sDoBeep: MExitMeth                                    '
         END IF                                                   '

      END IF                                                      '
      krRight                                                     ' We can finally move over to the next column
      MExitMeth                                                   ' We're done

   HexVal:
      IF INSTR("0123456789ABCDEF", UUCASE(gKeyChr)) = 0 THEN      '
         sDoBeep: MExitMeth                                       ' Better be valid Hex char
      END IF                                                      '
      gKeyChr = UUCASE(gKeyChr)                                   ' Uppercase it
      RETURN                                                      ' OK, Back
   END METHOD

   METHOD  krClipClear()
   '---------- Clear the clipboard
    sWriteClipboard("")                                           ' Write null to the clipboard
   END METHOD

   METHOD  krClipDate()
   '---------- Stuff Date into the clipboard
      sWriteClipboard(sDate())                                    ' Write it
   END METHOD

   METHOD  krClipISODate()
   '---------- Put ISO Date into the Clipboard
   LOCAL ln AS STRING
      ln = DATE$                                                  ' Go get the date
      ln = MID$(ln, 7) + "-" + LEFT$(ln, 2) + "-" + MID$(ln, 4, 2)' Reformat it to ISO standard
      sWriteClipboard(ln)                                         ' Write it
   END METHOD

   METHOD  krClipISOTime()
   '---------- Put ISO Time into the Clipboard
      sWriteClipboard(TIME$)                                      ' Write it
   END METHOD

   METHOD  krClipName()
   '---------- Stuff file name into the clipboard
   LOCAL ln AS STRING, ix, i, j AS LONG
      ln = TIP.File                                               ' Get the filename
      IF IsMEdit THEN                                             ' In a multi edit session, more to do
         IF ISFALSE IsCData AND ISFALSE IsCLinN THEN sDoBeep: EXIT METHOD ' Ignore totally if not in Data area
         ix = sGetIX(CRow)                                        ' Locate data line cursor was on
         IF ix = 0 THEN sDoBeep: EXIT METHOD                      ' Exit if below bottom line
         i = me.MEditTbl("I", FORMAT$(ix))                        ' Get the Medit index
         ln = me.MEditListGet(i)                                  ' Get the MEdit filename
         j = INSTR(-1, ln, "\")                                   ' Find last \
         IF j THEN ln = MID$(ln, j + 1)                           ' Reduce to just filename
      END IF                                                      '
      sWriteClipboard(ln)                                         ' Write it
   END METHOD

   METHOD  krClipPath()
   '---------- Stuff full file path into the clipboard
   LOCAL ln AS STRING, ix, i, j AS LONG
      ln = TIP.FilePath                                           ' Get the full path
      IF IsMEdit THEN                                             ' In a multi edit session, more to do
         IF ISFALSE IsCData AND ISFALSE IsCLinN THEN sDoBeep: EXIT METHOD ' Ignore totally if not in Data area
         ix = sGetIX(CRow)                                        ' Locate data line cursor was on
         IF ix = 0 THEN sDoBeep: EXIT METHOD                      ' Exit if below bottom line
         i = me.MEditTbl("I", FORMAT$(ix))                        ' Get the Medit index
         ln = me.MEditListGet(i)                                  ' Get the MEdit filename
      END IF                                                      '
      sWriteClipboard(ln)                                         ' Write it
   END METHOD

   METHOD  krClipTime()
   '---------- Stuff Time into the clipboard
      sWriteClipboard(sTime)                                      ' Write it
   END METHOD

   METHOD  krColumn()
   '---------- Jump to specified column
   LOCAL ix, ccCol, OldOffset, HexRow AS LONG
   LOCAL t AS STRING
      ccCol = IIF(VAL(gKeyPrimOper) > 0, VAL(gKeyPrimOper), 1)    ' Set the column number
      ix = sGetIX(CRow)                                           ' Locate data line cursor was on
      IF ix = 0 THEN EXIT METHOD                                  ' Exit if below bottom line
      IF ix < 0 THEN                                              ' A Hex line?
         HexRow = ABS(ix)                                         ' Save the hex line
         ix = sGetIX(CRow - ABS(ix))                              ' Get real data line
      END IF                                                      '
      me.CurSetReq(%Position, ix, ccCol, %True, 0, HexRow)        ' Set cursor set attempt
      OnInsClnSupp                                                ' Suppress Insert cleanup for 1 Attn pass
      DoSet(%Refresh)                                             ' Have it looked at
   END METHOD

   METHOD  krCondLineNo()
   '---------- Conditional LineNo key
   LOCAL i, j, k, l AS LONG
      CsrLinDX = 0                                                ' Reset possible hidden DX line location
      l = CRow                                                    ' Copy row

      '----- Only if in an eligible line
      IF CRow <= 2 + Prf.Cols THEN                                ' In header?
         l = 3 + Prf.Cols                                         ' Set l to top screen line
         GOSUB FindLine                                           ' Go find top Data line
      ELSE                                                        '
         GOSUB FindLine                                           ' Go find next Data line
      END IF                                                      '
      EXIT METHOD                                                 '

   FindLine:
      FOR i = l TO 300                                            ' Search screen lines
         j = S(i)                                                 ' Get the L() index from S()
         IF j AND IsLData(j) THEN                                 ' Got a line number of a data line?
            k = j: EXIT FOR                                       ' Save as found line
         ELSEIF j AND IsLTop(j) AND _                             ' On the Top?
            IsLBottom(j+1) THEN                                   ' And next line is the bottom
            k = j: EXIT FOR                                       ' Use the Top line
         ELSE                                                     ' No line?
            IF j = 0 THEN EXIT FOR                                ' No more screen lines
         END IF                                                   '
      NEXT i                                                      '
      IF k THEN                                                   ' Did we find a line?
         me.CsrRow = i: me.CsrCol = 1                             '
      ELSE                                                        '
         sDoBeep                                                  '
      END IF                                                      '
      RETURN                                                      '
   END METHOD

   METHOD  krCopyLCmd()
   '---------- Copy the line command (if present)
   LOCAL i, j AS LONG, CBD AS STRING
      IF ISFALSE IsCData AND ISFALSE IsCLinN THEN                 ' Not a valid location?
         CBD = " "                                                ' Clear the clipboard
         sWriteClipboard(CBD)                                     '
         EXIT METHOD                                              ' All done
      END IF

      j = sGetIX(CRow)                                            ' Get Line number we're on
      IF LLCtlG(j) <> LLNumG(j) THEN CBD = LLCtlG(j)              ' Get the LLCtl data
      ARRAY SCAN gLnoTextTxt(), = LSET$(CBD, 8), TO i             ' One of the special lines?
      IF i OR LEFT$(CBD, 1) = "." OR LEFT$(CBD, 1) = ":" THEN CBD = " " ' If special or a label or a tag, then null the CBD

      '----- Stuff it in the clipboard
      sWriteClipboard(CBD)                                        ' Write it
   END METHOD

   METHOD   krCopyPaste()
   '---------- Do Copy or Paste depending on whether selected
      IF (IsMarkActive AND IsMarkDrawn) OR _                      ' Is something selected
         (IsMiscActive AND IsMiscDrawn) THEN                      '
         krKCpy                                                   ' yes, COPY it
      ELSE                                                        '
         IF IsBrowse THEN sDoBeep: MExitMeth                      ' Not allowed in Browse mode
         IF IsFMTab THEN                                          ' FM screen?
            FMkbPaste                                             ' Do it that way
         ELSE                                                     '
            krPaste                                               ' No? Normal Paste
         END IF                                                   '
      END IF                                                      '
   END METHOD

   METHOD  krCopyPasteAdd()
   '---------- Block copypaste in ADD mode
      gKeyPrimOper = "$RAA" + gKeyPrimOper                        ' Tell krCopyPaste it's ADD mode
      CALL krCopyPaste()                                          ' Let IT do the work
   END METHOD

   METHOD  krCopyPasteRaw()
   '---------- Block copypaste in RAW mode
      gKeyPrimOper = "$RAW" + gKeyPrimOper                        ' Tell krCopyPaste it's RAW mode
      CALL krCopyPaste()                                          ' Let IT do the work
   END METHOD

   METHOD  krCtlLeft()
   '---------- Tab Left one 'word'
   LOCAL ix, i, j AS LONG
   LOCAL t AS STRING
      MEntry
      '----- Make sure we're in a valid location
      IF ISFALSE IsCData THEN MExitMeth                           ' Ignore totally if not in Data area
      ix = sGetIX(CRow)                                           ' Locate data line cursor was on
      IF ix = 0 THEN MExitMeth                                    ' Exit if below bottom line
      IF ISFALSE IsLData(ix) AND ISFALSE IsLNote(ix) THEN MExitMeth ' Also must be on a data/note line

      '----- Get local copy of the text and the column # we're in
      t = RTRIM$(LTxtG(ix))                                       ' Get a copy of the text
      i = CCol + Offset - gLNPadCol                               ' Where is cursor now

      '----- Col 1, simple, go to prev line
      IF i = 1 THEN                                               ' Column 1? (The simple one)
         GOSUB PrevLine                                           ' Then Prev line
         MExitMeth                                                ' We're done
      END IF                                                      '

      '----- Past the End, go to the last word
      IF i > LEN(RTRIM$(t)) THEN                                  ' At or past last char?
         GOSUB CurrLastWord                                       ' Go to Last Word

      '----- Within the actual text, work to do
      ELSE                                                        ' Within line, find next location

         '----- On space between words?
         IF MID$(t, i, 1) = " " THEN                              ' Currently sitting on a blank?
            j = INSTR(-(LEN(t) - i) - 2, t, ANY ENV.CharSet)      ' Look for prev visible character
            j = INSTR(-(LEN(t) - j)  - 2, t, " ") + 1             ' Look for prev blank
            IF j = 0 THEN                                         ' None?
               GOSUB PrevLine                                     ' Go to prev line then
            ELSE                                                  ' We've got the next 'word'
               IF j < Offset + 1 OR j > Offset + gDataLen - 1 THEN' Still on current screen
                  me.CurSetReq(%Position, ix, j, %True)           ' No, Set cursor set attempt
                  DoSet(%Refresh)                                 ' Have it looked at
               ELSE                                               '
                  me.CsrCol = j + gLNPadCol - Offset              ' Just move cursor on the screen
               END IF                                             '
            END IF                                                '

         ELSE                                                     ' We're sitting on a character

            '----- Sitting on 1st character of a word
            IF i > 1 AND MID$(t, i - 1, 1) = " " THEN             ' Are we sitting on the 1st char of a word?
               j = INSTR(-(LEN(t) - i) - 3, t, ANY ENV.CharSet)   ' Look for prev visible character
               j = INSTR(-(LEN(t) - j) - 2, t, " ") + 1           ' Look for prev blank
               IF j = 0 THEN                                      ' None?
                  GOSUB PrevLine                                  ' Go to prev line then
               ELSE                                               ' We've got the next 'word'
                  IF j < Offset + 1 OR j > Offset + gDataLen - 1 THEN ' Still on current screen
                     me.CurSetReq(%Position, ix, j, %True)        ' No, Set cursor set attempt
                     DoSet(%Refresh)                              ' Have it looked at
                  ELSE                                            '
                     me.CsrCol = j + gLNPadCol - Offset           ' Just move cursor on the screen
                  END IF                                          '
               END IF                                             '

            ELSE                                                  ' Sitting on other than 1st char of word
               '----- Sitting on some other character in the word
               j = INSTR(-(LEN(t) - i) - 2, t, " ") + 1           ' Look for prev blank + 1
               IF j = 0 THEN                                      ' None?
                  GOSUB PrevLine                                  ' Go to prev line then
               ELSE                                               ' We've got the next 'word'
                  IF j < Offset + 1 OR j > Offset + gDataLen - 1 THEN ' Still on current screen
                     me.CurSetReq(%Position, ix, j, %True)        ' No, Set cursor set attempt
                     DoSet(%Refresh)                              ' Have it looked at
                  ELSE                                            '
                     me.CsrCol = j + gLNPadCol - Offset           ' Just move cursor on the screen
                  END IF                                          '
               END IF                                             '
            END IF                                                '
        END IF                                                    '
     END IF                                                       '
     MExitMeth                                                    ' We're done

   '----- Go to last word on the previous line
   PrevLine:
      ix = me.CrpBack(%mData, ix, 1)                              ' Get address of prev data line
      IF ix = 0 THEN MExitMeth                                    ' Do nothing
      IF IsLInvisible(ix) THEN                                    ' Invisible?
         ix = me.LPopOutX(ix)                                     ' Yes, Uugh! go do lots of work, get adjusted line number
      END IF                                                      '
      t = RTRIM$(LTxtG(ix))                                       ' Get text of prev line
      j = INSTR(-1, t, " ")                                       ' Look for last blank
      IF j <> 0 THEN INCR j                                       ' 1st char of last word
      j = MAX(1, j)                                               ' If not found, set to 1
      me.CurSetReq(%Position, ix, j, %True)                       ' Set cursor set attempt
      DoSet(%Refresh)                                             ' Have it looked at
      RETURN                                                      '

   '----- Go to last word on the current line
   CurrLastWord:
      j = INSTR(-1, t, " ")                                       ' Look for last blank
      IF j <> 0 THEN INCR j                                       ' 1st char of last word
      j = MAX(1, j)                                               ' If not found, set to 1
      IF j < Offset + 1 OR j > Offset + gDataLen - 1 THEN         ' Still on current screen
         me.CurSetReq(%Position, ix, j, %True)                    ' No, Set cursor set attempt
         DoSet(%Refresh)                                          ' Have it looked at
      ELSE                                                        '
         me.CsrCol = j + gLNPadCol - Offset                       ' Just move cursor on the screen
      END IF                                                      '
      RETURN                                                      '
   END METHOD

   METHOD  krCtlRight()
   '---------- Tab right one 'word'
   LOCAL ix, i, j AS LONG
   LOCAL t AS STRING
      MEntry
      '----- Make sure we're somewhere legal
      IF ISFALSE IsCData THEN MExitMeth                           ' Ignore totally if not in Data area
      ix = sGetIX(CRow)                                           ' Locate data line cursor was on
      IF ix = 0 THEN MExitMeth                                    ' Exit if below bottom line
      IF ISFALSE IsLData(ix) AND ISFALSE IsLNote(ix) THEN MExitMeth ' Also must be on a data/note line

      '----- Get current line and column #
      t = RTRIM$(LTxtG(ix)) + " "                                 ' Get a copy of the text (+ 1 blank)
      i = CCol + Offset - gLNPadCol                               ' Where is cursor now

      '----- Past last character, go to next line
      IF i > LEN(RTRIM$(t)) THEN                                  ' At or past last char?
         GOSUB NextLine                                           ' Go to Col 1 next line

      ELSE                                                        ' Within line, find next location

         '----- Sitting ON a character
         IF MID$(t, i, 1) <> " " THEN                             ' Currently sitting on a char?
            j = INSTR(i + 1, t, " ")                              ' Look for next blank
            j = INSTR(j + 1, t, ANY ENV.CharSet)                  ' Look for next visible character
            IF j = 0 THEN                                         ' None?
               GOSUB NextLine                                     ' Go to next line then
            ELSE                                                  ' We've got the next 'word'
               IF j < Offset + 1 OR j > Offset + gDataLen - 1 THEN' Still on current screen
                  me.CurSetReq(%Position, ix, j, %True)           ' No, Set cursor set attempt
                  DoSet(%Refresh)                                 ' Have it looked at
               ELSE                                               '
                  me.CsrCol = j + gLNPadCol - Offset              ' Just move cursor on the screen
               END IF                                             '
            END IF                                                '
         ELSE                                                     ' We're sitting on a blank

            '----- Sitting on a blank
            j = INSTR(i + 1, t, ANY ENV.CharSet)                  ' Look for next visible character
            IF j = 0 THEN                                         ' None?
               GOSUB NextLine                                     ' Go to next line then
            ELSE                                                  ' We've got the next 'word'
               IF j < Offset + 1 OR j > Offset + gDataLen - 1 THEN' Still on current screen
                  me.CurSetReq(%Position, ix, j, %True)           ' No, Set cursor set attempt
                  DoSet(%Refresh)                                 ' Have it looked at
               ELSE                                               '
                  me.CsrCol = j + gLNPadCol - Offset              ' Just move cursor on the screen
               END IF                                             '
            END IF                                                '
        END IF                                                    '
     END IF                                                       '
     MExitMeth                                                    ' We're done

   '----- Go to the next line
   NextLine:
      ix = me.CrpFwd(%mData, ix, 1)                               ' Move to next line
      IF ix = 0 THEN MExitMeth                                    ' No more lines, exit
      IF IsLInvisible(ix) THEN                                    ' Invisible?
         ix = me.LPopOutX(ix)                                     ' Yes, Uugh! go do lots of work, get adjusted line number
      END IF                                                      '
      me.CurSetReq(%Position, ix, 1, %True)                       ' No, Set cursor set attempt
      DoSet(%Refresh)                                             ' Have it looked at
      RETURN                                                      '
   END METHOD

   METHOD  krDataDeleteMark()
   '---------- Data elete a marked area (if there is one)
      IF ISFALSE IsMarkActive THEN EXIT METHOD                    ' Is something selected?  No?  Exit
      me.krDataDelBS("D")                                         ' Just do a Data Delete
   END METHOD

   METHOD  krDataDelBS(which AS STRING)
   '---------- DataDeleteBackspace
   LOCAL i, j, x, y, z, OldLen, ix, GotColumns AS LONG, lcltxt, lh AS STRING, lclWTxt, lhW AS WSTRING

      MEntry
      IF ISFALSE IsCData THEN me.krDel(): MExitMeth               ' If not in Data area, pass off to krDel
      IF ISFALSE IsMarkActive THEN                                ' If normal (nothing selected)
         me.AttrInvClearFind                                      ' Clear any Find HiLite
         ix = sGetIX(CRow)                                        ' Get IX of real DataLine
         IF ix = 0 THEN MExitMeth                                 ' Bail out if below last line
         IF Prf.HexMode = &4 AND ix < 0 THEN sDoBeep: MExitMeth   ' If on a non-Char line alert user
         IF ISTRUE (LFlagG(ix) AND %NonTypable) THEN MExitMeth    ' Ignore non-typable lines

         '----- Delete a text character now
         me.ModSet(ix)                                            ' Remember we changed something
         i = CCol - gLNPadCol + Offset                            ' Calc IX into the Text
         IF which = "B" THEN i = IIF(i = 1, 1, i - 1)             ' Point at prev char if this is a backspace

         '----- See if any 'columns' in the data
         lcltxt = LTxtG(ix)                                       ' Get the text
         x = INSTR(i, lcltxt, "  ")                               ' Look for at least 2 blanks following found string
         IF x THEN                                                ' If we 've got spare blanks
            FOR y = x TO LEN(lcltxt)                              ' Look through remainder of line for a non- blank
               IF MID$(lcltxt, y, 1) <> " " THEN                  ' Got one?
                  GotColumns = %True                              ' Remember we have columns
                  EXIT FOR                                        ' Exit, y - 1 is split point for lh/rh
               END IF                                             '
            NEXT y                                                '
         END IF                                                   '

         '----- There are columns do it the hard way
         IF GotColumns THEN                                       ' Data Delete AND columns?
            lh = LEFT$(lcltxt, y - 2)                             ' Create full lh portion
            lh = STRDELETE$(lh, i, 1) + " "                       ' Delete char and add a blank
            lcltxt = lh + MID$(lcltxt, y - 1)                     ' Re-build the altered text
            me.LTxtSet(ix, lcltxt)                                ' Save it back
         ELSE
            me.LTxtCharDel(ix, i)                                 ' Delete 1 char
            me.LAttrAdjust(ix, i, 1)                              ' Adjust any Attr hilite
         END IF                                                   '

         IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) AND IsLData(ix) THEN ' CAPS ON?
            me.LTxtSet(ix, UUCASE(LTxtG(ix)))                     ' UC if CAPS is ON
         END IF                                                   '
         me.AttrScan(ix)                                          ' Recolorize
         me.DispLine(ix, CRow)                                    ' Re-display the line
         IF which = "B" THEN me.CsrCol = CCol - 1                 ' Move cursor back if this is a backspace

         '----- If HEX mode, we have to do the other lines
         IF Prf.HexMode = &4 THEN                                 ' If Hex mode, do extra
            IF CRow + 1 <= gwScrHeight THEN                       ' Make sure we don't go off the bottom
               sPrint (me.WindowHexTop(LTxtG(ix)), $$TxtHi, CRow + 1, gLNData1) ' Print upper hex
               me.MarkReDraw(CRow + 1)                            ' Re-draw the MARK lines
            END IF                                                '
            IF CRow + 2 <= gwScrHeight THEN                       ' Make sure we don't go off the bottom
               sPrint (me.WindowHexBottom(LTxtG(ix)), $$TxtHi, CRow + 2, gLNData1) ' Print lower hex
               me.MarkReDraw(CRow + 2)                            ' Re-draw the MARK lines
            END IF                                                '
         END IF                                                   '
         MExitMeth                                                '
      END IF

      IF which = "B" THEN MexitMeth                               ' No DataBackspace of a marked block
      '----- DataDelete of a highlighted area
      me.MarkKill                                                 ' Reset the select frame
      me.SwapKill                                                 '
      ix = sGetIX(CRow)                                           ' Get IX of real DataLine
      IF ix = 0 THEN ix = MarkSLin                                ' Use FromLin if not on screen

      '----- OK, now do the Delete of the marked range
      me.CurSetReq(%Position, ix, MarkSCol, %True)                ' Set cursor set attempt
      FOR z = MarkSLin TO MarkELin                                ' Loop through the requested range
         IF ISTRUE IsLData(z) OR ISTRUE IsLNote(z) THEN           ' Just data/note lines
            me.ModSet(z)                                          ' Remember we changed something
            OldLen = LEN(L(z).@LTxt)                              ' Save old length of line

            '----- See if any 'columns' in the data
            GotColumns = %False                                   ' Start as 'no columns'
            lcltxt = LTxtG(z)                                     ' Get the text
            x = INSTR(MarkECol + 1, lcltxt, "  ")                 ' Look for at least 2 blanks following deleted string
            IF x THEN                                             ' If we've got spare blanks
               FOR y = x TO LEN(lcltxt)                           ' Look through remainder of line for a non- blank
                  IF MID$(lcltxt, y, 1) <> " " THEN               ' Got one?
                     GotColumns = %True                           ' Remember we have columns
                     EXIT FOR                                     ' Exit, y - 1 is split point for lh/rh
                  END IF                                          '
               NEXT y                                             '
            END IF                                                '

            '----- There are columns do it the hard way
            IF GotColumns THEN                                    ' Data Delete AND columns?
               lh = LEFT$(lcltxt, y - 1)                          ' Pick up LH portion
               lh = STRDELETE$(lh, MarkSCol, MarkECol - MarkSCol + 1)  ' Delete chars
               lh = LSET$(lh, y - 1)                              ' Pad back the length
               lcltxt = lh + MID$(lcltxt, y)                      ' Re-build the altered text
            ELSE                                                  '
               lh = LEFT$(lcltxt, MarkSCol - 1)                   ' Create full lh portion
               lcltxt = lh + MID$(lcltxt, MarkECol + 1)           ' Re-build the altered text
            END IF                                                '
            me.LTxtSet(z, lcltxt)                                 ' Save it back

            '----- Now do highlights
            lclWTxt = LAttrG(z)                                   ' Get the Attr text
            IF GotColumns THEN                                    ' Data Delete AND columns?
               lhw = LEFT$(lclWTxt, y - 1)                        ' Pick up LH portion
               lhW = STRDELETE$(lhW, MarkSCol, MarkECol - MarkSCol + 1)  ' Delete chars
               lhW = LSET$(lhW, y - 1)                            ' Pad back the length
               lclWTxt = lhW + MID$(lclWTxt, y)                   ' Re-build the altered text
            ELSE                                                  '
               lhW = LEFT$(lclWTxt, MarkSCol - 1)                 ' Create full lh portion
               lclWTxt = lhW + MID$(lclWTxt, MarkECol + 1)        ' Re-build the altered text
            END IF                                                '
            LAttrS(z) = lclWTxt                                   ' Save it back

            IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) AND IsLData(x) THEN ' CAPS ON?
               me.LTxtSet(z, UUCASE(LTxtG(z)))                    ' UC if CAPS is ON
            END IF                                                '
            me.AttrScan(z)                                        ' Recolorize

         END IF                                                   '
      NEXT x                                                      '
      DoSet(%Refresh)                                             ' Have it looked at
      MExitMeth                                                   '
   END METHOD

   METHOD  krDataInsert()
   '---------- DataInsert ON
      IF IsNsrtFlag THEN onNsrtSave ELSE OffNsrtSave              ' Save it
      OnNsrtFlag                                                  ' Basic Insert goes ON
      OnNsrtData                                                  ' DataInsert on
      sDoStatusBar($SBInsOvr)                                     ' re-Do the StatusBar INS/OVR box
   END METHOD

   METHOD  krDate()
   '---------- Paste Date at cursor location
   LOCAL ln AS STRING
      ln = sDate()                                                ' Go get the date
      me.krTDPaste(ln)
   END METHOD

   METHOD  krDel()
   '---------- Delete character at current position
   LOCAL ix, x, i AS LONG, lclTxt AS STRING
      MEntry
      CsrLinDX = 0                                                ' Reset possible hidden DX line location

      '----- Located on the Command line
      IF IsCCmnd THEN                                             ' In Command line
         IF IsMiscActive THEN                                     ' Marked area
            me.MarkKill                                           ' Reset the select text
            pCommand = STRDELETE$(pCommand, MiscSCol, MiscECol - MiscSCol + 1) + REPEAT$(MiscECol - MiscSCol + 1, " ")
            me.CsrCol = MiscSCol - COffset + 11 - 1               ' Set cursor to start of deleted area

         ELSE                                                     '
            pCommand = STRDELETE$(pCommand, CCol - 10 + COffset, 1) + " " ' Delete it
         END IF                                                   '
         me.WindowCmd                                             ' Re-do the command line

      '----- Located in the Scroll Amt. area
      ELSEIF IsCScrl THEN                                         ' In Scrl Amt area?
         Prf.Scroll = STRDELETE$(Prf.Scroll, CCol - 20 - pCmdLen, 1) + " " ' Delete it
         sPrint (LSET$(Prf.Scroll, 4), $$TxtLo, 1, 21 + pCmdLen)  ' Redisplay the entire string

      '----- Located in the normal text area
      ELSEIF IsCData THEN                                         ' In Data area
         me.AttrInvClearFind                                      ' Clear any Find HiLite
         IF IsBrowse THEN sDoBeep: MExitMeth                      ' Not allowed in Browse mode

         '----- If a highlighted area, get it's boundaries, then kill it
         IF IsMarkActive THEN                                     ' Hi-lighted text and cursor in Data area?
            IF ZLNum >= MarkSLin AND _                            ' Is cursor within the marked area?
               ZLNum <= MarkELin AND _                            '
               ZLCol >= MarkSCol AND _                            '
               ZLCol <= MarkECol THEN                             '
               IF IsNsrtFlag THEN                                 ' Do the Insert mode handling
                  '----- Do a char mode delete on each line
                  FOR i = MarkSLin TO MarkELin                    ' Process each line in the marked area
                     IF ISTRUE IsLData(i) OR ISTRUE IsLNote(i) THEN ' Just data/note lines
                        IF ZLCol >= MarkSCol AND ZLCol <= MarkECol THEN ' Room to Delete?
                           me.ModSet(i)                           ' Remember we changed something
                           me.LTxtCharDel(i, ZLCol)               ' Delete 1 char
                           me.LAttrAdjust(i, ZLCol, -1)           ' Adjust any Attr hilite
                           IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) AND IsLData(ix) THEN  ' CAPS ON and data line?
                              me.LTxtSet(i, UUCASE(LTxtG(i)))     ' UC it
                           END IF                                 '
                           me.AttrScan(i)                         ' Recolorize
                        ELSE                                      ' Not in the block
                           GOTO DelNonInsert                      ' Do nothing
                        END IF                                    '
                     END IF                                       '
                  NEXT i                                          '
                  DECR MarkECol                                   ' Pull in the Mark End column
                  IF MarkSCol > MarkECol THEN                     ' Nothing left?
                     me.MarkKill                                  ' Reset the select frame
                     me.SwapKill                                  '
                  END IF                                          '
                  me.CurSetReq(%Position, ZLNum, ZLCol, %True)    ' Set cursor set attempt
                  me.CsrCol = ZLCol + Offset + gLNData1           '
                  DoSet(%Refresh)                                 ' Have it looked at
                  MExitMeth                                       '

               ELSE                                               ' Do the non-Insert mode handling
                  DelNonInsert:
                  me.MarkKill                                     ' Reset the select frame
                  me.SwapKill                                     '

                  '----- OK, now do the Delete of the marked range
                  FOR x = MarkSLin TO MarkELin                    ' Loop through the requested range
                     IF ISTRUE IsLData(x) OR ISTRUE IsLNote(x) THEN  ' Just data/note lines
                        me.ModSet(x)                              ' Remember we changed something
                        lclTxt = LTxtG(x)                         ' Get the text
                        lclTxt = STRDELETE$(lclTxt, MarkSCol, MarkECol - MarkSCol + 1) ' Delete the characters
                        me.LAttrAdjust(x, MarkSCol, -(MarkECol - MarkSCol + 1)) ' Adjust any Attr hilite
                        IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) THEN        ' CAPS ON?
                           me.LTxtSet(x, UUCASE(LTxtG(x)))        ' UC if CAPS is ON
                        END IF                                    '
                        me.LTxtSet(x, lclTxt)                     ' Put work string back
                        me.AttrScan(x)                            ' Recolorize
                        me.CurSetReq(%Position, MarkELin, MarkSCol, %True) ' Set cursor set attempt
                        me.CsrCol = MarkSCol + Offset + gLNPadCol '
                     END IF                                       '
                  NEXT x                                          '
                  DoSet(%Refresh)                                 ' Have it looked at
                  MExitMeth                                       '
               END IF                                             '
            END IF                                                '
            me.MarkKill                                           ' Reset the select frame
            me.SwapKill                                           '

         END IF                                                   '
                                                                  '
         '----- Get detail location and eliminate invalid ones
         ix = sGetIX(CRow)                                        ' Get IX of real DataLine
         IF ix = 0 THEN MExitMeth                                 ' Bail out if below last line
         IF Prf.HexMode = &4 AND ix < 0 THEN sDoBeep: MExitMeth      ' If on a non-Char line alert user
         IF ISTRUE (LFlagG(ix) AND %NonTypable) THEN              ' Ignore non-typable lines
            ' Null
         ELSE                                                     '

            '----- Delete a text character now
            me.ModSet(ix)                                         ' Remember we changed something
            i = CCol - gLNPadCol + Offset                         ' Calc IX into the Text
            me.LTxtCharDel(ix, i)                                 ' Delete 1 char
            me.LAttrAdjust(ix, i, -1)                             ' Adjust any Attr hilite
            IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) AND IsLData(ix) THEN ' CAPS ON?
               me.LTxtSet(ix, UUCASE(LTxtG(ix)))                  ' UC if CAPS is ON
            END IF                                                '
            me.AttrScan(ix)                                       ' Recolorize
            me.DispLine(ix, CRow)                                 ' Re-display the line

            '----- If HEX mode, we have to do the other lines
            IF Prf.HexMode = &4 THEN                                 ' If Hex mode, do extra
               IF CRow + 1 <= gwScrHeight THEN                    ' Make sure we don't go off the bottom
                  sPrint (me.WindowHexTop(LTxtG(ix)), $$TxtHi, CRow + 1, gLNData1) ' Print upper hex
                  me.MarkReDraw(CRow + 1)                         ' Re-draw the MARK lines
               END IF                                             '
               IF CRow + 2 <= gwScrHeight THEN                    ' Make sure we don't go off the bottom
                  sPrint (me.WindowHexBottom(LTxtG(ix)), $$TxtHi, CRow + 2, gLNData1) ' Print lower hex
                  me.MarkReDraw(CRow + 2)                         ' Re-draw the MARK lines
               END IF                                             '
            END IF                                                '
         END IF                                                   '

      '----- Located in the Line Number area
      ELSEIF IsCLinN THEN                                         ' In the Line control area
         ix = sGetIX(CRow)                                        ' Get IX of real DataLine
         IF ix = 0 THEN MExitMeth                                 ' Bail out if below last line
         IF me.IsTouchable(ix, CCol) THEN                         ' If 1st interaction in the line number area
            LLCtlS(ix) = $BlankLNo: me.CsrCol = 1                 ' Blank the Line number
            sPrint (BYCOPY LEFT$(LLCtlG(ix), ENV.LinNoSize), $$LnoHi, CRow, 1)' ReDisplay it
         END IF                                                   '
         LLCtlS(ix) = STRDELETE$(LLCtlG(ix), CCol, 1) + " "       ' Delete a char
         sPrint (BYCOPY LEFT$(LLCtlG(ix), ENV.LinNoSize), $$LnoHi, CRow, 1)   ' ReDisplay it
         me.TTblAdd(ix, TRIM$(LLCtlG(ix)))                        ' Update touched items
      END IF                                                      '
      MExit                                                       '
   END METHOD

   METHOD  krDeleteMark()
   '---------- Delete a marked area (if there is one
   LOCAL ix, i AS LONG, lclTxt AS STRING
      '----- If a marked area, then remove it
      me.AttrInvClearFind                                         ' Clear any Find HiLite
      IF IsMarkActive THEN                                        ' Is something selected?
         IF IsBrowse THEN sDoBeep: EXIT METHOD                    ' Not allowed in Browse mode
         IF MarkSLin = MarkELin AND _                             ' and a single line mark?
            CCol - gLNPadCol + Offset >= MarkSCol AND _           ' and cursor within the marked area?
            CCol - gLNPadCol + Offset <= MarkECol THEN            '
            me.MarkKill                                           ' Reset the select frame
            me.SwapKill                                           '

            '----- OK, now do the Delete
            me.ModSet(MarkSLin)                                   ' Remember we changed something
            lclTxt = LTxtG(MarkSLin)                              ' Get the text
            lclTxt = STRDELETE$(lclTxt, MarkSCol, MarkECol - MarkSCol + 1) ' Delete the characters
            IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) THEN              ' CAPS ON?
               lclTxt = UUCASE(lclTxt)                            ' UC if CAPS is ON
            END IF                                                '
            me.LTxtSet(me.MarkSLin, lclTxt)                       ' Put work string back
            me.LAttrAdjust(MarkSLin, MarkSCol, -(MarkECol - MarkSCol + 1)) ' Adjust any Attr hilite
            me.CurSetReq(%Position, MarkSLin, MarkSCol, %True)    ' Set cursor set attempt
            me.CsrCol = MarkSCol + Offset + gLNPadCol             '
            DoSet(%Refresh)                                       ' Have it looked at
         END IF                                                   '
      END IF                                                      '
   END METHOD

   METHOD  krDown()
   '---------- Increment CsrRow
   LOCAL t AS STRING
   LOCAL ix, CurData, cIncr AS LONG

      MEntry
      cIncr = IIF(VAL(gKeyPrimOper) > 0, VAL(gKeyPrimOper), 1)    ' Set the INCR amount?
      '----- Move the cursor down
      CurData = sGetIX(CRow)                                      ' Get line cursor is on
      IF CurData <> 0 AND CRow + cIncr <= gwScrHeight AND sGetIX(CRow + 1) = 0 THEN ' End of Page display?
         Prf.ScrlPageSus = %True                                  ' Set suspend PAGE mode
         OnInsClnSupp                                             ' Suppress Insert cleanup for 1 Attn pass
         DoSet(%Refresh)                                          ' Have it looked at
         MExitMeth                                                '
      END IF                                                      '
      CsrLinDX = 0                                                ' Reset possible hidden DX line location
      me.CsrRowAdd(cIncr)                                         ' Simple INCR

      '----- Handle edge wrap or scroll
      IF ISFALSE ENV.KBVerScroll OR IsCCmnd THEN                  ' AutoScroll or not
         IF CRow > gwScrHeight THEN me.CsrRow = 1                 ' NO - Handle Wrap if needed

      ELSE                                                        '

         '----- Scroll it at the bottom
         IF CRow > gwScrHeight THEN                               ' AutoScroll in Data area - At bottom line yet?
            me.CsrRowSub(cIncr)                                   ' Yes, Undo the INCR for now

            '----- Get the actual line the cursor was on
            ix = sGetIX(CRow)                                     ' Locate data line cursor was on
            IF ix < 0 THEN                                        ' A Hex extra line?
               ix = sGetIX(CRow - ABS(ix))                        ' Point at real character line
            END IF                                                '

            '----- Can't scroll if already below last line
            IF ix = 0 THEN MExitMeth                              ' Exit if not data line

            '----- If not yet at the bottom line, find the next visible line
            IF ISFALSE IsLBottom(ix) THEN                         ' See if cursor is at the bottom
               ix = me.CRPFwd(%mVisible, ix, cIncr)               ' Move to next visible line
               IF CCol > gLNPadCol THEN                           ' If in data area
                  me.CurSetReq(%Position, ix, CCol + Offset - gLNPadCol, %False) ' Set cursor set attempt
               ELSE                                               ' Must be in Line Number area
                  me.CurSetReq(%Position, ix, 0, %False)          ' Set cursor set attempt
               END IF                                             '
               TopScrn = me.CRPFwd(%mVisible, TopScrn, cIncr)     ' Move TopScrn by scroll amount
               OnInsClnSupp                                       ' Suppress Insert cleanup for 1 Attn pass
               DoSet(%Refresh)                                    ' Have it looked at
            END IF                                                '
         END IF                                                   '
      END IF                                                      '
      MExit                                                       '
   END METHOD

   METHOD  krDup()
   '---------- Duplicate prior line into the marked block
   LOCAL x, i AS LONG, lclTxt, t AS STRING

      me.AttrInvClearFind                                         ' Clear any Find HiLite
      '----- Get the data from the previous line
      i = me.CrpBack(%mData, MarkSLin, 1)                         ' Get address of prev data line
      IF ISFALSE IsLData(i) THEN sDoBeep: EXIT METHOD             ' Beep and exit if no prior
      t = LTxtG(i)                                                ' Get text from the previous line
      IF LEN(t) < MarkECol THEN t = LSET$(t, MarkECol)            ' If line shorter than select columns, lengthen it
      t = MID$(t, MarkSCol TO MarkECol)                           ' Get the string to be Dup'd

      '---- OK, now stuff it into the lines
      me.ModSet(MarkSLin)                                         ' Remember we changed something
      FOR x = MarkSLin TO MarkELin                                ' Loop through the requested range
         IF ISFALSE (LFlagG(x) AND %NonTypable) THEN              ' Just typable lines
            lclTxt = LTxtG(x)                                     ' Get the text
            IF LEN(lclTxt) < MarkECol THEN lclTxt = LSET$(lclTxt, MarkECol) ' Padd if needed
            MID$(lclTxt, MarkSCol TO MarkECol) = t                ' Extract the columns
            me.LTxtSet(x, lclTxt)                                 ' Save it back
            me.AttrScan(x)                                        ' Recolorize
         END IF                                                   '
      NEXT x                                                      '

      '----- We're done, clean up and exit
      me.MarkKill                                                 ' Reset the select frame
      me.CurSetReq(%LineCmd, MarkSLin, MarkSCol, %True)           ' Set cursor set attempt
      DoSet(%Refresh)                                             ' Have it looked at
   END METHOD

   METHOD  krEditBrowse(which AS STRING)
   '---------- Edit/Browse the highlighted filename
   LOCAL x, i, AppendText AS LONG, lclTxt, a, f, t, CBD AS STRING
      '----- Ensure a marked area and get its boundaries
      IF MarkSLin <> MarkELin OR _                                ' If not one line
         ISFALSE (IsLData(MarkELin)) AND ISFALSE (IsLNote(MarkELin)) THEN _ ' or not Data/Note line
         sDoBeep: EXIT METHOD

      '----- OK, now build the string
      lclTxt = LTxtG(MarkELin)                                    ' Get the text
      IF LEN(lclTxt) < MarkECol THEN lclTxt = lclTxt + SPACE$(MarkECol - LEN(lclTxt)) ' Padd if needed
      lclTxt = MID$(lclTxt, MarkSCol, MarkECol - MarkSCol + 1)    ' Extract the columns
      IF TRIM$(lclTxt) = "" THEN sDoBeep: EXIT METHOD             '
      DO WHILE INSTR($DQ + "'<([ ", LEFT$(lclTxt, 1)) <> 0        ' Left trim it
         lclTxt = MID$(lclTxt, 2)                                 '
      LOOP                                                        '
      DO WHILE INSTR($DQ + "'>)];,. ", RIGHT$(lclTxt, 1)) <> 0    ' Right trim it
         lclTxt = LEFT$(lclTxt, LEN(lclTxt) - 1)                  ' Remove quotes if present
      LOOP                                                        '
      REPLACE "/" WITH "\" IN lclTxt                              ' Normalize path delimiters
      IF INSTR(lclTxt, "\") = 0 THEN lclTxt = TIP.Path + lclTxt   ' Add path is none there

      '----- Pass off to EDIT / BROWSE as needed
      IF which = "EDIT" THEN
         PCmdEDIT("EDIT " + $DQ + lclTxt + $DQ)                   ' Call EDIT
      ELSEIF which = "BROWSE" THEN                                '
         PCmdBrowse("BROWSE " + $DQ + lclTxt + $DQ)               ' Call BROWSE
      ELSEIF which = "VIEW" THEN                                  '
         PCmdView("VIEW " + $DQ + lclTxt + $DQ)                   ' Call VIEW
      END IF                                                      '
      DoSet(%Refresh)                                             ' Have it looked at
   END METHOD

   METHOD  krEndLine()
   '---------- EndLine key
   LOCAL ix, i, HexRow AS LONG
      '----- On text data line
      IF IsCData OR IsCLinN THEN                                  '
         ix = sGetIX(CRow)                                        ' Get IX of line the cursor is on
         IF ix = 0 THEN EXIT METHOD                               ' Bail out if below last line
         IF ix = -3 THEN                                          ' Dash line?
            sDoBeep: EXIT METHOD                                  ' Error
         ELSEIF ix < 0 THEN                                       ' A Hex extra line?
            IF ix < 0 THEN                                        ' A Hex line?
               HexRow = ABS(ix)                                   ' Save the hex line
               ix = sGetIX(CRow - ABS(ix))                        ' Get real data line
            END IF
         END IF                                                   '
         IF ISFALSE (LFlagG(ix) AND %NonTypable) THEN             ' See if cursor on a typable line
            i = LEN(LTxtG(ix))                                    ' Calc where cursor should go
            IF i >= Offset AND i <= Offset + gDataLen THEN        ' Still on displayed screen?
               me.CsrCol = i + gLNData1 - Offset                  ' Just set the new Col value
            ELSE                                                  '
               me.CurSetReq(%Posend, ix, i + 1, %True, 0, Hexrow) ' Set cursor set attempt
               DoSet(%Refresh)                                    ' Have it looked at
            END IF                                                '
         END IF                                                   '

      '----- On the Command line
      ELSEIF IsCCmnd THEN                                         '
         i = LEN(pCommand)                                        '
         IF i > pCmdLen + COffset THEN                            ' Is it to the right of the Window?
            COffset = i - (pCmdLen / 2)                           ' Adjust offset
         ELSEIF i < COffset THEN                                  ' Is it to the left of the Window?
            COffset = MAX(0, i - (pCmdLen / 2))                   ' Adjust offset
         END IF                                                   '
         me.WindowCmd                                             ' Redisplay the line
         me.CsrCol = i - COffset + 11                             '

      '----- In the Scroll Command field
      ELSEIF IsCScrl THEN                                         '
         i = LEN(RTRIM$(Prf.Scroll))                              '
         i = MIN(3, i)                                            ' So we don't point cursor off the right edge
         me.CsrCol = i + 21 + pCmdLen                             '
      END IF                                                      '
   END METHOD

   METHOD  krEndText()
   '---------- EndText key
   LOCAL ix, i, HexRow AS LONG
      '----- On a text data line
      IF IsCData OR IsCLinN THEN                                  '
         ix = sGetIX(CRow)                                        ' Get IX of line the cursor is on
         IF ix = 0 THEN EXIT METHOD                               ' Bail out if below last line
         IF ix = -3 THEN                                          ' Dash line?
            sDoBeep: EXIT METHOD                                  ' Error
         ELSEIF ix < 0 THEN                                       ' A Hex extra line?
            IF ix < 0 THEN                                        ' A Hex line?
               HexRow = ABS(ix)                                   ' Save the hex line
               ix = sGetIX(CRow - ABS(ix))                        ' Get real data line
            END IF
         END IF                                                   '
         IF ISFALSE (LFlagG(ix) AND %NonTypable) THEN             ' See if cursor on a typable line
            i = LEN(RTRIM$(LTxtG(ix)))                            ' Calc where cursor should go
            IF i >= Offset AND i <= Offset + gDataLen THEN        ' Still on displayed screen?
               me.CsrCol = i + gLNData1 - Offset                  ' Just set the new Col value
            ELSE                                                  '
               me.CurSetReq(%Posend, ix, i + 1, %True, 0, HexRow) ' Set cursor set attempt
               DoSet(%Refresh)                                    ' Have it looked at
            END IF                                                '
         END IF                                                   '

      '----- On the Command line
      ELSEIF IsCCmnd THEN                                         '
         i = LEN(RTRIM$(pCommand))                                '
         IF i > pCmdLen + COffset THEN                            ' Is it to the right of the Window?
            COffset = i - (pCmdLen / 2)                           ' Adjust offset
         ELSEIF i < COffset THEN                                  ' Is it to the left of the Window?
            COffset = MAX(0, i - (pCmdLen / 2))                   ' Adjust offset
         END IF                                                   '
         me.WindowCmd                                             ' Redisplay the line
         me.CsrCol = i - COffset + 11                             '

      '----- On the Scroll Amount field
      ELSEIF IsCScrl THEN                                         '
         i = LEN(RTRIM$(Prf.Scroll))                              '
         i = MIN(3, i)                                            ' So we don't point cursor off the right edge
         me.CsrCol = i + 21 + pCmdLen                             '
      END IF                                                      '
   END METHOD

   METHOD  krEnter()
   '---------- Enter pressed
   LOCAL ix, lbefore, lafter, i, j AS LONG, w AS STRING
      MEntry
      me.AttrInvClearFind                                         ' Clear any Find HiLite
      IF gmacroMode THEN MexitMeth
      IF IsFMTab THEN me.FMkbEnter: MExitMeth                     ' In case internal Enter passed
      DOSet(%Attention)                                           ' Enter always causes Attention
      DOSet(%MarkKill)                                            ' Clear any maeked area

      '----- Exit quickly if a CRETREIV
      w = UUCASE(sGetWord(pCommand, %NoStrip, %QuoteNotSig))      ' Get any pending command
      IF IsPFKInsert AND w = "CRETRIEV" THEN                      ' If a CRTRIEV?  (Avoid doing a sCurSetReq below)
         MExitMeth                                                '
      END IF                                                      '

      '----- If in text data area, add a new insert line if appropriate
      IF IsCData THEN                                             ' In the data area?
         ix = sGetIX(CRow)                                        ' Get IX of line the cursor is on
         IF ix < 0 THEN ix = sGetIX(CRow - ABS(ix))               ' Get real data line if in Hex mode
         IF ix <> 0 THEN                                          ' Make sure not below bottom line
            i = IsLInsertLine(ix)
            j = IsLInsertLine(ix + 1)
            IF ix = InsLastClr AND _                              ' See if cursor last used Insert Line
               ISFALSE IsLInsertLine(ix + 1) THEN                 ' and the next line is not an Insert
               LLCtlS(ix) = "I1      "                            ' Default an extra insert line
               me.TTblAdd(ix, TRIM$(LLCtlG(ix)))                  ' Update touched items
               me.InsLastClr = 0                                  ' Clear last line
            END IF                                                '
         END IF                                                   '
      END IF                                                      '

      '----- If in line # area, get real line number
      IF IsCLinN THEN                                             ' If cursor in the line number field
         ix = sGetIX(CRow)                                        ' Get IX of line the cursor is on
         IF ix < 0 THEN ix = sGetIX(CRow - ABS(ix))               ' Get real data line if in Hex mode
      END IF                                                      '

      '----- If in the data area, do the complex 'where does the cursor go' routine
      IF CRow > 2 + Prf.Cols AND CRow <= gwScrHeight THEN         ' If cursor in the data lines
         IF ISNULL(TRIM$(pCommand)) THEN                          ' and the command is blank
            ix = sGetIX(CRow)                                     ' Get IX of line the cursor is on
            IF ix < 0 THEN ix = sGetIX(CRow - ABS(ix))            ' Get real data line if in Hex mode
            IF IsLXclude(ix) THEN                                 ' Skip over excluded lines
               ix = me.CRPFwd(%MVisible, ix, 1)                   '
            END IF                                                '
            IF ix <> 0 THEN                                       ' Ensure a data line
               IF CCol > ENV.LinNoSize THEN _                     ' If not in line number area
                  ix = me.CRPFwd(%MVisible, ix, 1)                ' Get next visible line

               IF LEN(L(ix).@LTxt) > 0 THEN                       ' If something in the line
                  me.CurSetReq(%Low, ix, MAX(VERIFY(MAX(Offset + 1, 1), LTxtG(ix), " "), Offset + 1), %False)
               ELSE                                               '
                  lbefore = me.CRPBack(%MVisible, ix, 1)          ' Get previous line
                  lafter = me.CRPFwd(%MVisible, ix, 1)            ' Get next line
                  IF lbefore > 1 AND LEN(L(lbefore).@LTxt) = 0 THEN ' If it's zero as well
                     IF lafter < LastLine AND LEN(L(lafter).@LTxt) = 0 THEN
                        me.CurSetReq(%Low, ix, 1 + Offset, %False)' Set cursor to the data column 1
                     ELSE                                         ' lafter is present
                        me.CurSetReq(%Low, ix, MAX(VERIFY(MAX(Offset + 1, 1), LTxtG(lafter), " "), Offset + 1), %False) ' Set cursor as per next line
                     END IF                                       '
                  ELSE                                            '
                     me.CurSetReq(%Low, ix, MAX(VERIFY(MAX(Offset + 1, 1), LTxtG(lbefore), " "), Offset + 1), %False)   ' Set cursor as per previous line
                  END IF                                          '
               END IF                                             '
            END IF                                                '
         END IF                                                   '
      END IF                                                      '

      '----- In the Command line
      IF IsCCmnd THEN                                             ' In the Command area?

         '----- Reset retrieve stuff if no actual command present
         IF ISNULL(TRIM$(pCommand)) THEN                          ' and the command is blank
            gCmdRtrevIX = 1: gCmdRtrevMsg = ""                    ' Reset Retrieve Index
            IF CRow > 1 THEN                                      ' If not actually ON the Cmd line
               ix = sGetIX(3 + Prf.Cols)                          ' Get IX of 1st data line on screen
               me.CurSetReq(%Low, ix, 0, %False)                  ' Set cursor to top line linenum field
            END IF                                                '
         END IF                                                   '
      END IF                                                      '

      MExit
   END METHOD

   METHOD  krEnum(which AS STRING)
   '---------- Do the Enum function
   LOCAL i, j, k, enumz, ln AS LONG
   LOCAL menum AS DWORD
   LOCAL t, tt, mprefix, msuffix, digits, nv AS STRING
      MEntry
      '----- Do some basic setup
      IF ISFALSE IsMarkActive THEN MExitMeth                      ' Hi-lighted text? No? Bail out
      me.AttrInvClearFind                                         ' Clear any Find HiLite

      CsrLinDX = 0                                                ' Reset possible hidden DX line location
      digits = $Numeric: IF which <> "D" THEN digits = $Hex + "abcdef" ' Set the valid 'digits' string
      t = LTxtG(MarkSLin)                                         ' Get the first text line
      IF LEN(t) < MarkECol THEN t += SPACE$(MarkECol - LEN(t))    ' Padd if needed
      t = MID$(t, MarkSCol, MarkECol - MarkSCol + 1)              ' Extract the columns containg the model

      '----- See if there's a suffix
      IF INSTR(digits, RIGHT$(t, 1)) > 0 THEN                     ' If it ends in a digit,
         msuffix = ""                                             ' then there is NO suffix
      ELSE                                                        ' Non-digit, we must extract the suffix
         i = INSTR(-1, t, ANY digits)                             ' Look for the last digit
         IF i = 0 THEN                                            ' None? whoops
            DoSet(%Refresh)                                       ' Have it looked at
            MErrExit(%eFail, "Enum model has no starting value")  ' Error
         END IF                                                   '
         msuffix = MID$(t, i + 1)                                 ' Found digit, extract suffix
         t = LEFT$(t, i)                                          ' Strip it off the model
      END IF                                                      '
      me.MarkKill                                                 ' Reset the select frame
      me.SwapKill                                                 '

      '----- Now extract the prefix
      j = 0                                                       ' Reset j
      FOR i = LEN(t) TO 1 STEP -1                                 ' Look backward for 1st non-digit
         IF INSTR(digits, MID$(t, i, 1)) > 0 THEN ITERATE FOR     ' If a digit, keep looking
         j = i: EXIT FOR                                          ' Found a non digit, save its loc and exit for
      NEXT i                                                      '
      IF j = 0 THEN                                               ' All were digits
         mprefix = ""                                             ' Therefor no prefix
      ELSE                                                        ' Non-digit, we have a prefix
         mprefix = LEFT$(t, j)                                    ' Save prefix
         t = MID$(t, j + 1)                                       ' Extract it leaving t as the eval string.
      END IF                                                      '

      '----- Extract the enum value
      ln = LEN(t)                                                 ' Save length for rebuilding
      IF which = "D" THEN                                         ' Decimal Enum?
         menum = VAL(t)                                           ' Save it
         IF LEN(t) <> LEN(FORMAT$(menum)) THEN enumz = %True      ' Entered as leading zeros?, remember that
      ELSE                                                        ' One of the Hex
         menum = VAL("&H" + t)                                    ' Save it that way
         IF LEN(t) <> LEN(HEX$(menum)) THEN enumz = %True         ' Entered as leading zeros?, remember that
      END IF                                                      '

      '----- Now modify the PT lines
      FOR k = MarkSLin TO MarkELin                                ' Loop for each line
         IF ISFALSE IsLData(k) THEN ITERATE FOR                   ' Only Data lines
         me.ModSet(k)                                             ' Remember we changed something
         tt = LTxtG(k)                                            ' Get working copy of text
         IF LEN(tt) < MarkECol THEN tt = tt + SPACE$(MarkECol - LEN(tt)) ' Padd if needed

         '----- Select which Enum type to do
         SELECT CASE AS CONST$ which                              ' Build output string
            CASE "D"                                              ' Decimal
               t = FORMAT$(menum)                                 '
               GOSUB AdjustPad                                    ' Do padding etc.
            CASE "L"                                              ' Lowercase Hex
               t = LLCASE(HEX$(menum))                            '
               GOSUB AdjustPad                                    ' Do padding etc.
            CASE "U"                                              ' Lowercase Hex
               t = HEX$(menum)                                    '
               GOSUB AdjustPad                                    ' Do padding etc.
         END SELECT                                               '

         '----- Build the value and stuff it in
         t = mprefix + t + msuffix                                ' Add prefix and suffix
         t = RIGHT$(RSET$(t, MarkECol - MarkSCol + 1), MarkECol - MarkSCol + 1) ' Right align and trim it
         MID$(tt, MarkSCol, LEN(t)) = t                           ' Overlay the enum text
         me.LTxtSet(k, tt)                                        ' Save it back
         me.AttrScan(k)                                           ' Recolorize

         menum += gEnumWith                                       ' Incr enum for next line
      NEXT k                                                      '
      DoSet(%Refresh)                                             ' Have it looked at
      MExitMeth                                                   '

   '----- Fiddle with the prefix padding
   AdjustPad:
      IF LEN(t) > ln AND RIGHT$(mprefix, 1) = " " THEN            ' If enum needs more room, and Prefix has some
         mprefix = LEFT$(mprefix, LEN(mprefix) - 1)               ' Steal right hand blank from prefix
         INCR ln                                                  ' Make enum longer
      ELSEIF LEN(t) > ln AND LEFT$(mprefix, 1) = " " THEN         ' If not right end, try left end
         mprefix = MID$(mprefix, 2)                               ' Steal left hand blank from prefix
         INCR ln                                                  ' Make enum longer
      END IF                                                      '
      IF LEN(t) > ln AND RIGHT$(msuffix, 1) = " " THEN            ' If enum needs more room, and Suffix has some
         msuffix = LEFT$(msuffix, LEN(msuffix) - 1)               ' Steal right hand blank from suffix
         INCR ln                                                  ' Make enum longer
      ELSEIF LEN(t) > ln AND LEFT$(msuffix, 1) = " " THEN         ' If not right end, try left end
         msuffix = MID$(msuffix, 2)                               ' Steal left hand blank from suffix
         INCR ln                                                  ' Make enum longer
      END IF                                                      '
      IF enumz THEN                                               ' Pad appropriately
         t = RIGHT$(REPEAT$(ln, "0") + t, ln)                     '
      ELSE                                                        '
         t = RIGHT$(REPEAT$(ln, " ") + t, ln)                     '
      END IF                                                      '
      RETURN                                                      '

   END METHOD

   METHOD  krEOL()
   '---------- EOL key
   LOCAL ix, i, j AS LONG

      MEntry
      '----- Cursor in the Command area
      IF IsCCmnd THEN                                             ' In the Command area?
         pCommand = LEFT$(pCommand, CCol - 11 + COffset)          ' Truncate the line then
         me.WindowCmd                                             ' Re-do the command line

      '----- Cursor in the Scroll amount field
      ELSEIF IsCScrl THEN                                         ' In the Scroll Amt area?
         j = CCol - 21 - pCmdLen                                  ' Calc length of truncated part
         Prf.Scroll = LEFT$(LEFT$(Prf.Scroll, j) + "    ", 4)     ' Truncate the line then
         sPrint (LSET$(Prf.Scroll, 4), $$TxtLo, 1, 21 + pCmdLen)  ' Print the whole string again

      '----- Cursor in the text data
      ELSEIF IsCData  THEN                                        ' In the Data area?
         me.AttrInvClearFind                                      ' Clear any Find HiLite
         IF IsBrowse THEN sDoBeep: MExitMeth                      ' Not allowed in Browse mode

         '----- Ensure we're on a valid line
         ix = sGetIX(CRow)                                        ' Get IX of line the cursor is on
         IF ix = 0 THEN MExitMeth                                 ' Bail out if below last line
         IF ix < 0 THEN sDoBeep: MExitMeth                        ' Hex extra line, error
         IF ISFALSE (LFlagG(ix) AND %NonTypable) THEN             ' See if cursor on a typable line

            '----- Finally, truncate the line
            i = CCol - gLNPadCol + Offset                         ' Calc index into string
            me.LTxtSet(ix, LEFT$(LTxtG(ix), i - 1))               ' Truncate the line then
            LAttrS(ix) = LEFT$(LAttrG(ix), i - 1)                 ' And the Attr line
            me.CurSetReq(%EOL, ix, i, %True)                      ' Set cursor set attempt
            me.ModSet(ix)                                         ' Remember we changed something
            IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) AND IsLData(ix) THEN ' CAPS ON?
               me.LTxtSet(ix, UUCASE(LTxtG(ix)))                  ' UC if CAPS is ON
            END IF                                                '
            me.AttrScan(ix)                                       ' Recolorize
            me.DispLine(ix, CRow)                                 ' Re-display the line

            '----- Do a lot more if in HEX mode
            IF Prf.HexMode = &4 THEN                                 ' If Hex mode, do extra
               IF CRow + 1 <= gwScrHeight THEN                    ' Make sure we don't go off the bottom
                  sPrint (me.WindowHexTop(LTxtG(ix)), $$TxtHi, CRow + 1, gLNData1) ' Print upper hex
                  me.MarkReDraw(CRow + 1)                         ' Re-draw the MARK lines
               END IF                                             '
               IF CRow + 2 <= gwScrHeight THEN                    ' Make sure we don't go off the bottom
                  sPrint (me.WindowHexBottom(LTxtG(ix)), $$TxtHi, CRow + 2, gLNData1) ' Print lower hex
                  me.MarkReDraw(CRow + 2)                         ' Re-draw the MARK lines
               END IF                                             '
            END IF                                                '
         END IF                                                   '

      '----- Cursor in the Line Number area
      ELSEIF IsCLinN  THEN                                        ' In the Line Number area
         ix = sGetIX(CRow)                                        ' Get IX of line the cursor is on
         IF ix = 0 THEN MExitMeth                                 ' Bail out if below last line
         IF me.IsTouchable(ix, CCol) THEN                         ' If 1st interaction in the line number area
            LLCtlS(ix) = $BlankLNo: me.CsrCol = 1                 ' Blank the Line number
            sPrint (BYCOPY LEFT$(LLCtlG(ix), ENV.LinNoSize), $$LnoHi, CRow, 1)' ReDisplay it
         END IF                                                   '
         LLCtlS(ix) = LEFT$(LEFT$(LLCtlG(ix), CCol - 1) + $BlankLNo, 8)
         me.TTblAdd(ix, TRIM$(LLCtlG(ix)))                        ' Update touched items
         sPrint (BYCOPY LEFT$(LLCtlG(ix), ENV.LinNoSize), $$LnoHi, CRow, 1)   ' ReDisplay it
      END IF                                                      '
      MExit                                                       '
   END METHOD

   METHOD  krErase()
   '---------- Erase (blank) a selected block
   LOCAL x, i, j, AppendText AS LONG, lclTxt, a, f, t, CBD AS STRING
      '---- OK, do the erase
      me.AttrInvClearFind                                         ' Clear any Find HiLite
      FOR x = MarkSLin TO MarkELin                                ' Loop through the requested range
         IF ISTRUE IsLData(x) THEN                                ' Just data lines
            lclTxt = LTxtG(x)                                     ' Get the text
            IF LEN(lclTxt) < MarkECol THEN lclTxt = lclTxt + SPACE$(MarkECol - LEN(lclTxt)) ' Padd if needed
            MID$(lclTxt, MarkSCol, MarkECol - MarkSCol + 1) = SPACE$(MarkECol - MarkSCol + 1) ' Blank the columns
            me.LTxtSet(x, lclTxt)                                 ' Stuff it back
            me.AttrScan(x)                                        ' Recolorize
            me.ModSet(x)                                          ' Remember we changed something
         END IF                                                   '
      NEXT x                                                      '
      me.MarkKill                                                 ' Reset the select frame
      DoSet(%Refresh)                                             ' Have it looked at
   END METHOD

   METHOD  krFind(which AS STRING)
   '---------- Find the highlighted filename
   LOCAL x, i, AppendText AS LONG, lclTxt, a, f, t, CBD, cmd AS STRING
      '----- See if anything's highlighted, if not treat as RFIND
      me.AttrInvClearFind                                         ' Clear any Find HiLite
      IF ISFALSE IsMarkActive THEN                                ' If nothing selected, treat as RFIND
         PCmdRFIND(IIF$(which = "NEXT", "RFIND ", "RFIND REV "))  ' Call RFIND
         me.DispScreen                                            ' Redisplay things

      '----- Stuff highlighted, grab the boundaries
      ELSE                                                        ' Else do the full marked thing
         IF MarkSLin <> MarkELin THEN sDoBeep: EXIT METHOD        ' Better be just one line
         IF ISFALSE IsLData(MarkELin) AND ISFALSE IsLNote(MarkELin) THEN sDoBeep: EXIT METHOD ' Better be a data/note line

         '----- OK, now build the string
         lclTxt = LTxtG(MarkELin)                                 ' Get the text
         IF LEN(lclTxt) < MarkECol THEN lclTxt = lclTxt + SPACE$(MarkECol - LEN(lclTxt)) ' Padd if needed
         lclTxt = MID$(lclTxt, MarkSCol, MarkECol - MarkSCol + 1) ' Extract the columns
         lclTxt = sDoQuoteString(lclTxt)                          ' Quote it to protect spaces, quotes etc.
         me.MarkKill                                              ' Reset the select frame
         me.SwapKill                                              '
         PCmdFIND("FIND " + which + " " + lclTxt)                 ' Call FIND
         me.DispScreen                                            ' Redisplay things
      END IF                                                      '
   END METHOD

   METHOD  krFirstLCmd()
   '---------- First Line Cmd area
      me.CsrRow = 3 + Prf.Cols: me.CsrCol = 1
   END METHOD

   METHOD  krHome()
   '---------- Home the cursor
      IF IsCCmnd THEN                                             ' If IN the command line,
         COffset = 0: CCol = 11                                   ' Remove Offset, home cursor
         me.WindowCmd                                             ' Redisplay the line
      END IF                                                      '
      SetCmd                                                      ' Cursor to the command line
   END METHOD

   METHOD  krInsert()
   '---------- Insert key ON or OFF
      IF IsNsrtFlag THEN OffNsrtFlag ELSE OnNsrtFlag              ' Flip/Flop the mode
      IF IsNsrtFlag THEN OnNsrtSave ELSE OffNsrtSave              ' Save it
      OffNsrtData                                                 ' Always set DataInsert off
      sCaretDestroy                                               '
      sCaretCreate                                                '
      sDoCursor                                                   '
      sCaretShow                                                  '
      sDoStatusBar($SBInsOvr)                                     ' re-Do the StatusBar Ins/Ovr box
   END METHOD

   METHOD  krISODate()
   '---------- Paste ISO Date at cursor location
   LOCAL ln AS STRING
      ln = DATE$                                                  ' Go get the date
      ln = MID$(ln, 7) + "-" + LEFT$(ln, 2) + "-" + MID$(ln, 4, 2)' Reformat it to ISO standard
      me.krTDPaste(ln)
   END METHOD

   METHOD  krISOTime()
   '---------- Paste ISO Time at cursor location
   LOCAL ln AS STRING
      ln = TIME$                                                  ' Go get the time
      me.krTDPaste(ln)
   END METHOD

   METHOD  krJustify()
   '---------- Justify a selected block
   LOCAL x, i AS LONG, lclTxt, jtxt, ttxt AS STRING
      '----- OK, now do the Justify
      me.AttrInvClearFind                                         ' Clear any Find HiLite
      FOR x = MarkSLin TO MarkELin                                ' Loop through the requested range
         IF ISTRUE IsLData(x) OR ISTRUE IsLNote(x) THEN           ' Just data lines
            lclTxt = LTxtG(x)                                     ' Get the text
            me.ModSet(x)                                          ' Remember we changed something
            IF LEN(lclTxt) < MarkECol THEN _                      ' Lengthen text if needed
               lclTxt = LSET$(lclTxt, MarkECol)                   '
            jtxt = MID$(lclTxt, MarkSCol, MarkECol - MarkSCol +1) ' Get the string to be justified
            SELECT CASE AS CONST$ gKeyChr                         ' Do the right type of justify
               CASE "JUSTIFYL": jtxt = LSET$(LTRIM$(jtxt), LEN(jtxt))'
               CASE "JUSTIFYR": jtxt = RSET$(RTRIM$(jtxt), LEN(jtxt))'
               CASE "JUSTIFYC": ttxt = SPACE$(LEN(jtxt)): jtxt = TRIM$(jtxt)
                                MID$(ttxt, (LEN(ttxt) - LEN(jtxt)) \ 2 + 1, LEN(jtxt)) = jtxt
                                jtxt = ttxt                       '
            END SELECT                                            '
            MID$(lclTxt, MarkSCol, MarkECol - MarkSCol + 1) = jtxt' Re-insert the string
            me.LTxtSet(x, lclTxt)                                 ' Put work string back
            me.AttrScan(x)                                        ' Recolorize
            me.CurSetReq(%Position, MarkELin, MarkSCol, %True)    ' Set cursor set attempt
         END IF                                                   '
      NEXT x                                                      '
      DoSet(%Refresh)                                             ' Have it looked at
      me.MarkKill                                                 ' Reset the select frame
   END METHOD

   METHOD  krKCpy()
   '---------- Block COPY based on selected window text
   LOCAL x, i, j, AppendText, FNum AS LONG, lclTxt, a, f, t, CBD, fn, MyDLM AS STRING

      MEntry
      me.AttrInvClearFind                                         ' Clear any Find HiLite
      IF IsMarkActive THEN                                        ' Normal Text mark?
         '---- OK, now build the Clipboard string
         IF LEFT$(gKeyPrimOper, 4) <> "$RAA" THEN                 ' If not APPEND mode
            CBD = ""                                              ' Start as ""
         ELSE                                                     '
            sReadClipboard(CBD, MyDLM, %False)                    ' Go get whatever's there, keep the file
         END IF                                                   '

         FOR x = MarkSLin TO MarkELin                             ' Loop through the requested range
            IF ISTRUE IsLData(x) OR ISTRUE IsLNote(x) THEN        ' Just Data / Note lines
               lclTxt = LTxtG(x)                                  ' Get the text
               IF LEN(lclTxt) < MarkECol THEN lclTxt = LSET$(lclTxt, MarkECol) ' Padd if needed
               lclTxt = MID$(lclTxt, MarkSCol TO MarkECol)        ' Extract the columns
               IF LEFT$(gKeyPrimOper, 4) <> "$RAW"  AND LEFT$(gKeyPrimOper, 4) <> "$RAA" THEN  ' If not RAW or RAA, do it the normal way
                  CBD += lclTxt + $CR + $LF                       ' Add each line with CR/LF
               ELSE                                               ' Else
                  CBD += lclTxt                                   ' Do it without CRLF
               END IF                                             '
            END IF                                                '
         NEXT x                                                   '
         me.MarkKill                                              ' Reset the select frame
         IF LEFT$(gKeyPrimOper, 4) = "$RAW" OR LEFT$(gKeyPrimOper, 4) = "$RAA" THEN gKeyPrimOper = MID$(gKeyPrimOper, 5) ' If RAW, strip it off now

      '----- Must be Misc line mark
      ELSE                                                        '
         SELECT CASE AS LONG MiscZone                             ' Which one?
            CASE %CursCmnd                                        ' Command line
               CBD = MID$(pCommand, MiscSCol TO MiscECol)         ' Get string from command line
               me.MarkKill                                        ' Reset the select text
            CASE %CursPath                                        ' FM Path line
               CBD = MID$(FPath, MiscSCol TO MiscECol)            ' Get string from the line
               me.MarkKill                                        ' Reset the select text
            CASE %CursMask                                        ' FM Mask line
               CBD = MID$(FMask, MiscSCol TO MiscECol)            ' Get string from the line
               me.MarkKill                                        ' Reset the select text
            CASE %CursLNote                                       ' FM Note line
               CBD = MID$(AFList(MiscLine).Note, MiscSCol TO MiscECol) ' Get string from the line
               me.MarkKill                                        ' Reset the select text
         END SELECT                                               '
      END IF                                                      '

      '----- Stuff it in the clipboard
      sWriteClipboard(CBD)                                        ' Write it
      MExit
   END METHOD

   METHOD  krKCpyAdd()
   '---------- Block copy in RAW APPEND mode
      gKeyPrimOper = "$RAA" + gKeyPrimOper                        ' Tell krKCpy it's RAW mode
      me.krKCpy()                                                 ' Let IT do the work
   END METHOD

   METHOD  krKCpyRaw()
   '---------- Block copy in RAW mode
      gKeyPrimOper = "$RAW" + gKeyPrimOper                        ' Tell krKCpy it's RAW mode
      me.krKCpy()                                                 ' Let IT do the work
   END METHOD

   METHOD  krKCut()
   '---------- Block CUT based on selected window text
   LOCAL x, i, j, AppendText, FNum AS LONG, lclTxt, a, f, t, CBD, fn AS STRING

      MEntry
      me.AttrInvClearFind                                         ' Clear any Find HiLite
      IF IsMarkActive THEN                                        ' Normal Text mark?
         '----- OK, now build the Clipboard string
         CBD = ""                                                 ' Start as ""
         FOR x = MarkSLin TO MarkELin                             ' Loop through the requested range
            IF ISTRUE IsLData(x) OR ISTRUE IsLNote(x) THEN        ' Just Data.Note lines
               lclTxt = LTxtG(x)                                  ' Get the text
               IF LEN(lclTxt) < MarkECol THEN lclTxt = LSET$(lclTxt, MarkECol) ' Padd if needed
               lclTxt = MID$(lclTxt, MarkSCol TO MarkECol)        ' Extract the columns
               CBD += lclTxt + $CR + $LF                          ' Add each line with CR/LF
            END IF                                                '
         NEXT x                                                   '

         '----- Stuff it in the clipboard
         IF sWriteClipboard(CBD) THEN _                           ' Write it
            me.MarkKill                                           ' Reset the select frame

         '----- OK, now do the Delete
         FOR x = MarkSLin TO MarkELin                             ' Loop through the requested range
            IF ISTRUE IsLData(x) OR ISTRUE IsLNote(x) THEN        ' Just data/note lines
               me.ModSet(x)                                       ' Remember we changed something
               lclTxt = LTxtG(x)                                  ' Get the text
               lclTxt = STRDELETE$(lclTxt, MarkSCol, MarkECol - MarkSCol + 1) ' Delete the characters
               me.LAttrAdjust(x, MarkSCol, -(MarkECol - MarkSCol + 1))    ' Adjust any Attr hilite
               IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) THEN           ' CAPS ON?
                  me.LTxtSet(x, UUCASE(LTxtG(x)))                 ' UC if CAPS is ON
               END IF                                             '
               me.LTxtSet(x, lclTxt)                              ' Put work string back
               me.AttrScan(x)                                     ' Recolorize
               me.CurSetReq(%Position, MarkELin, MarkECol - (MarkECol - MarkSCol), %True) ' Set cursor set attempt
            END IF                                                '
         NEXT x                                                   '
      ELSE                                                        '
         SELECT CASE AS LONG MiscZone                             ' Which one?
            CASE %CursCmnd                                        ' Command line
               CBD = MID$(pCommand, MiscSCol TO MiscECol)         ' Get string from command line
               pCommand = STRDELETE$(pCommand, MiscSCol, MiscECol - MiscSCol + 1) ' Delete the characters
               me.MarkKill                                        ' Re-Display it
            CASE %CursPath                                        ' FM Path line
               CBD = MID$(FPath, MiscSCol TO MiscECol)            ' Get string from the line
               me.MarkKill                                        '
               FPath = STRDELETE$(FPath, MiscSCol, MiscECol - MiscSCol + 1) ' Delete the characters
            CASE %CursMask                                        ' FM Mask line
               CBD = MID$(FMask, MiscSCol TO MiscECol)            ' Get string from the line
               me.MarkKill                                        '
               FMask = STRDELETE$(FMask, MiscSCol, MiscECol - MiscSCol + 1) ' Delete the characters
            CASE %CursLNote                                       ' FM Note line
               CBD = MID$(AFList(MiscLine).Note, MiscSCol TO MiscECol) ' Get string from the line
               me.MarkKill                                        '
               AFList(MiscLine).Note = STRDELETE$(AFList(MiscLine).Note, MiscSCol, MiscECol - MiscSCol + 1) ' Delete the characters
         END SELECT                                               '
         sWriteClipboard(CBD)                                     ' Write it
      END IF                                                      '
      DoSet(%Refresh)                                             ' Have it looked at
      MExit
   END METHOD

   METHOD  krLastTab()
   '---------- Tab To the Last Tab marker
   LOCAL i, j, k AS LONG, lclTabs AS STRING
      MEntry
      OnInsClnSupp                                                ' Suppress Insert cleanup for 1 Attn pass

      SELECT CASE CRow                                            ' Do tabbing based on where we are

         '----- Top line
         CASE 1                                                   ' Top Row
            IF CCol < 11 THEN SetCmd: MExitMeth                   '
            IF CCol < 20 + pCmdLen THEN me.CsrCol = 20 + pCmdLen: MExitMeth
            me.CsrRow = 3 + Prf.Cols: me.CsrCol = 1: MExitMeth    '

         '----- Dash line
         CASE 2, 2 + Prf.Cols                                     ' Line 2 or COLS line
            me.CsrRow = 3 + Prf.Cols: me.CsrCol = 1: MExitMeth    '

         '----- Not in the headers
         CASE ELSE                                                ' In data area

            '----- Get text line number and working TABS line
            j = sGetIX(CRow)                                      ' Get Line number of cursor line
            lclTabs = me.TabsSimple(LEN(L(j).@LTxt))              ' Get a working tabs line

            i = INSTR(Prf.TabsLine, "+")                             ' + type tab line

            '----- Normal, non + type tab line
            IF ISFALSE i THEN                                     ' Can't really do LastTab with extensible tabs
               IF j < 0 THEN j = sGetIX(CRow - ABS(j))            ' Get real data line
               k = 0                                              ' Say no next ine
               IF CRow < gwScrHeight THEN k = sGetIX(CRow + 1)    ' See if another screen line
               IF ISFALSE Prf.Tabs THEN                           ' If not tabbing
                  me.CsrRowAdd(1): me.CsrCol = 1                  ' Do basic bump
                  IF CRow > gwScrHeight OR k = 0 THEN SetCmd      ' Down too far?  Cursor to the command line
                  MExitMeth                                       ' Leave now
               END IF                                             '
               IF IsLTop(j) THEN                                  ' On the Top of Data Line
                  me.CsrCol = 1: me.CsrRowAdd(1)                  ' Col 1 of next line
               END IF                                             '
                                                                  ' Tabbing IN a data line
               i = INSTR(-1, lclTabs, "*")                        ' look for last tab marker
            END IF                                                '

            '----- Tab line has + repeating tab
            IF i THEN                                             ' we have one
               IF i < Offset + 1 OR i > Offset + gDataLen - 1 THEN' Still on current screen
                  me.CurSetReq(%Position, j, i, %True)            ' No, Set cursor set attempt
                  DoSet(%Refresh)                                 ' Have it looked at
                  MExitMeth                                       '
               ELSE                                               '
                  me.CsrCol = i + gLNPadCol - Offset              ' Just move cursor on the screen
                  MExitMeth                                       '
               END IF                                             '
            ELSE                                                  ' ?? No tabs?
               me.CsrCol = gLNData1 - Offset                      ' Just move cursor on the screen
            END IF                                                '
         END SELECT                                               '
      MExit                                                       '
   END METHOD

   METHOD  krLeft()
   '---------- Decrement CsrCol
   LOCAL ix, cIncr, OldOffset, Hexrow AS LONG
   LOCAL t AS STRING
      MEntry
      cIncr = IIF(VAL(gKeyPrimOper) > 0, VAL(gKeyPrimOper), 1)    ' Set the INCR amount?
      me.CsrColSub(cIncr)                                         ' Simple DECR
      IF CCol = gLNPadCol AND Offset = 0 THEN me.CsrColSub(1)     ' Make Col gLNPadCol an Auto-Skip field

      '----- In the command line
      IF cRow = 1 THEN                                            ' In the command area
         IF CCol < 11 AND COffset > 0 THEN                        ' If left and and offset > 0
            OldOffset = COffset                                   ' Save our original Offset
            me.CsrColAdd(cIncr)                                   ' Undo the DECR we did
            cIncr = MIN(cIncr, COffSet)                           ' Make sure we don't go too far left
            me.COffsetSub(cIncr)                                  ' Reduce the offset
            me.CsrCol = MIN(11, CCol - (COffSet - OldOffset) + cIncr)  ' Position the cursor
            me.WindowCmd                                          ' Re-do the command line
            MExitMeth                                             '
         END IF                                                   '
      END IF                                                      '

      '----- Not in text data or no horizontal scrolling
      IF ISFALSE ENV.KBHorScroll OR ISFALSE IsCData THEN          ' No KeybScroll or not in data area
         IF CCol = 0 THEN                                         ' Line Wrap?
            IF ENV.KBHorScroll THEN                               ' If KeybScroll, don't move
               me.CsrColAdd(1)                                    ' Undo the DECR we did
            ELSE                                                  '
               me.CsrCol = ENV.ScrWidth: me.CsrRowSub(1)          ' Yes, make adjustment
               IF CRow = 0 THEN me.CsrRow = gwScrHeight           ' Also Row Wrap if needed
            END IF                                                '
         END IF                                                   '

      ELSE                                                        '
         '----- Scrolling is possible, do it
         IF CCol < gLNData1 AND Offset > 0 THEN                   ' YES - At left side of screen and room to scroll?
            me.CsrColAdd(cIncr)                                   ' Yes, Undo the DECR for now
            cIncr = MIN(cIncr, OffSet)                            ' Make sure we don't go too far left
            ix = sGetIX(CRow)                                     ' Locate data line cursor was on
            IF ix = 0 THEN MExitMeth                              ' Exit if below bottom line
            IF ix < 0 THEN                                        ' A Hex line?
               HexRow = ABS(ix)                                   ' Save the hex line
               ix = sGetIX(CRow - ABS(ix))                        ' Get real data line
            END IF
            me.CurSetReq(%Position, ix, CCol + Offset - gLNPadCol - cIncr, %False, 0, Hexrow) ' Set cursor set attempt
            me.OffsetSub(cIncr)                                   ' Drop the offset
            OnInsClnSupp                                          ' Suppress Insert cleanup for 1 Attn pass
            DoSet(%Refresh)                                       ' Have it looked at
         END IF                                                   '
      END IF                                                      '
      MExit                                                       '
   END METHOD

   METHOD  krLift()
   '---------- Block COPY and Erase based on selected window text
   LOCAL x, i, j, AppendText, FNum AS LONG, lclTxt, ltxt2, a, f, t, CBD, fn AS STRING

      MEntry
      me.AttrInvClearFind                                         ' Clear any Find HiLite
      IF IsMarkActive THEN                                        ' Normal text hilighted
         IF IsBrowse THEN sDoBeep: MExitMeth                      ' Not allowed in Browse mode

         '---- OK, now build the Clipboard string
         CBD = ""                                                 ' Start as ""
         FOR x = MarkSLin TO MarkELin                             ' Loop through the requested range
            IF ISTRUE IsLData(x) THEN                             ' Just data lines
               lclTxt = LTxtG(x)                                  ' Get the text
               IF LEN(lclTxt) < MarkECol THEN lclTxt = LSET$(lclTxt, MarkECol) ' Padd if needed
               lTxt2 = MID$(lclTxt, MarkSCol TO MarkECol)         ' Extract the columns
               CBD += lTxt2 + $CR + $LF                           ' Add each line with CR/LF
               MID$(lclTxt, MarkSCol TO MArkECol) = SPACE$(MarkECol - MarkSCol + 1) ' Blank the columns
               me.LTxtSet(x, lclTxt)                              ' Stuff it back
               me.AttrScan(x)                                     ' Recolorize
               me.ModSet(x)                                       ' Remember we changed something
            END IF                                                '
         NEXT x                                                   '
         '----- Stuff it in the clipboard
         IF sWriteClipboard(CBD) THEN _                           ' Write it
            me.MarkKill                                           ' Reset the select frame

      ELSE
         SELECT CASE AS LONG MiscZone                             ' Which one?
            CASE %CursCmnd                                        ' Command line
               CBD = MID$(pCommand, MiscSCol TO MiscECol)         ' Get string from command line
               lclTxt = pCommand                                  ' Get a copy of the line
               MID$(lclTxt, MiscSCol TO MiscECol) = SPACE$(LEN(CBD)) ' Blank the columnss
               pCommand = lclTxt                                  ' Store it back
               me.MarkKill                                        ' Reset the select frame
            CASE %CursPath                                        ' FM Path line
               CBD = MID$(FPath, MiscSCol TO MiscECol)            ' Get string from the line
               lclTxt = FPath                                     ' Get a copy of the line
               MID$(lclTxt, MiscSCol TO MiscECol) = SPACE$(LEN(CBD)) ' Blank the columnss
               FPath = lclTxt                                     ' Store it back
               me.MarkKill                                        ' Reset the select frame
            CASE %CursMask                                        ' FM Mask line
               CBD = MID$(FMask, MiscSCol TO MiscECol)            ' Get string from the line
               lclTxt = FMask                                     ' Get a copy of the line
               MID$(lclTxt, MiscSCol TO MiscECol) = SPACE$(LEN(CBD)) ' Blank the columnss
               FMask = lclTxt                                     ' Store it back
               me.MarkKill                                        ' Reset the select frame
            CASE %CursLNote                                       ' FM Note line
               CBD = MID$(AFList(MiscLine).Note, MiscSCol TO MiscECol) ' Get string from the line
               lclTxt = AFList(MiscLine).Note                     ' Get a copy of the line
               MID$(lclTxt, MiscSCol TO MiscECol) = SPACE$(LEN(CBD)) ' Blank the columnss
               AFList(MiscLine).Note = lclTxt                     ' Store it back
               me.MarkKill                                        '
         END SELECT                                               '
            me.MarkKill                                           ' Reset the select frame
      END IF                                                      '
      DoSet(%Refresh)                                             ' Have it looked at
      MExit                                                       '
   END METHOD

   METHOD  krLineNo()
   '---------- LineNo key
   LOCAL ix, i AS LONG
      CsrLinDX = 0                                                ' Reset possible hidden DX line location

      '----- Only if in an eligible line
      IF IsCData OR IsCLinN THEN                                  '
         ix = sGetIX(CRow)                                        ' Get IX of line the cursor is on
         IF ix = 0 OR ix = -3 THEN sDoBeep: EXIT METHOD           ' Below last line or Hex dash line? Bail out
         me.CsrCol = 1                                            ' Move cursor back to col 1
      ELSE                                                        '
         sDoBeep                                                  ' Beep
      END IF                                                      '
   END METHOD

   METHOD  krLowerCase()
   '---------- LowerCase a selected block
   LOCAL x, i AS LONG, lclTxt AS STRING
      me.AttrInvClearFind                                         ' Clear any Find HiLite
      '----- OK, now do the Lowercase
      FOR x = MarkSLin TO MarkELin                                ' Loop through the requested range
         IF ISTRUE IsLData(x) OR ISTRUE IsLNote(x) THEN           ' Just data/note lines
            lclTxt = LTxtG(x)                                     ' Get the text
            me.ModSet(x)                                          ' Remember we changed something
            IF MarkSCol <= LEN(lclTxt) THEN                       ' Do only if long enough
               FOR i = MarkSCol TO MIN(MarkECol, LEN(lclTxt))     ' Just ones that are there
                  MID$(lclTxt, i, 1) = LLCASE(MID$(lclTxt, i, 1)) '
               NEXT i                                             '
               me.LTxtSet(x, lclTxt)                              ' Put work string back
               me.AttrScan(x)                                     ' Recolorize
               me.CurSetReq(%Position, MarkELin, MarkECol + 1, %True)  ' Set cursor set attempt
            END IF                                                '
         END IF                                                   '
      NEXT x                                                      '
      DoSet(%Refresh)                                             ' Have it looked at
      me.MarkKill                                                 ' Reset the select frame
   END METHOD

   METHOD  krMark(Dir AS STRING, Pass AS DWORD)
   '---------- Mark Down Arrow
   LOCAL i, j, TLin, TCol, OldLin, NewLin, ColAdj AS LONG
   LOCAL t AS STRING
   LOCAL sl, el, ll AS LONG
      MEntry
      IF IsMiscActive AND MiscZone <> Zone THEN _                 ' Moved outside current marked zone?
         me.MarkKill: MExitMeth                                   ' Kill it and leave cursor alone

      SELECT CASE AS LONG Zone                                    ' Split by where we are

         '----- Mark in the command line
         CASE %CursCmnd                                           ' Command line

            '----- If highlight is active
            IF ISFALSE IsMiscActive THEN                          ' If no active area right now, start one
               '----- No active bounds right now, Start one up
               MiscZone = Zone                                    ' Save the zone we're marking in
               IF dir = "E" THEN                                  ' If END case
                  MiscSCol = ZLCol: MiscECol = LEN(RTRIM$(pCommand))
                  CALL DWORD Pass                                 ' Then just move the cursor
               ELSE                                               '
                  MiscSCol = ZLCol: MiscECol = ZLCol              '
               END IF                                             '
               OnMiscActive                                       ' We're starting a marked area
               me.MiscMark                                        ' Go Draw it

            ELSE                                                  '
               ' We have an active area already
               GOSUB KillOrMove                                   ' Handle the simple conditions

               '----- We're at a boundary, see how to adjust the bounds
               SELECT CASE dir                                    ' Adjust boundaries
                  CASE "L"                                        ' Left
                     GOSUB MiscLeft                               ' Go adjust for Left

                  CASE "R"                                        ' Right
                     GOSUB MiscRight                              ' Go adjust for right

                  CASE "E"                                        ' End
                     MiscECol = LEN(RTRIM$(pCommand))             ' Move right to text length column
                     me.MiscMark: CALL DWORD Pass: MExitMeth      ' Redraw the hilight and we're done

                  CASE "U", "D"                                   ' Up/Down?
                     me.MarkKill: CALL DWORD Pass: MExitMeth      ' Just kill and move the cursor
               END SELECT                                         '
            END IF                                                '

         '----- Mark in the Path line
         CASE %CursPath                                           ' FM Path?

            '----- If highlight is active
            IF ISFALSE IsMiscActive THEN                          ' If no active area right now, start one
               '----- No active bounds right now, Start one up
               MiscZone = Zone                                    ' Save the zone we're marking in
               IF dir = "E" THEN                                  ' If END case
                  MiscSCol = ZLCol: MiscECol = LEN(RTRIM$(FPath)) '
                  CALL DWORD Pass                                 ' Then just move the cursor
               ELSE                                               '
                  MiscSCol = ZLCol: MiscECol = ZLCol              '
               END IF                                             '
               OnMiscActive                                       ' We're starting a marked area
               me.MiscMark                                        ' Go Draw it

            ELSE                                                  '
               ' We have an active area already
               GOSUB KillOrMove                                   ' Handle the simple conditions

               '----- We're at a boundary, see how to adjust the bounds
               SELECT CASE dir                                    ' Adjust boundaries
                  CASE "L"                                        ' Left
                     GOSUB MiscLeft                               ' Go adjust for Left

                  CASE "R"                                        ' Right
                     GOSUB MiscRight                              ' Go adjust for right

                  CASE "E"                                        ' End
                     MiscECol = LEN(RTRIM$(FPath))                ' Move right to text length column
                     me.MiscMark: CALL DWORD Pass: MExitMeth      ' Redraw the hilight and we're done

                  CASE "U", "D"                                   ' Up/Down?
                     me.MarkKill: CALL DWORD Pass: MExitMeth      ' Just kill and move the cursor

               END SELECT                                         '
            END IF                                                '

         CASE %CursMask                                           ' FM Mask?

            '----- If highlight is active
            IF ISFALSE IsMiscActive THEN                          ' If no active area right now, start one
               '----- No active bounds right now, Start one up
               MiscZone = Zone                                    ' Save the zone we're marking in
               IF dir = "E" THEN                                  ' If END case
                  MiscSCol = ZLCol: MiscECol = LEN(RTRIM$(FMask)) '
                  CALL DWORD Pass                                 ' Then just move the cursor
               ELSE                                               '
                  MiscSCol = ZLCol: MiscECol = ZLCol              '
               END IF                                             '
               OnMiscActive                                       ' We're starting a marked area
               me.MiscMark                                        ' Go Draw it

            ELSE                                                  '
               ' We have an active area already
               GOSUB KillOrMove                                   ' Handle the simple conditions

               '----- We're at a boundary, see how to adjust the bounds
               SELECT CASE dir                                    ' Adjust boundaries
                  CASE "L"                                        ' Left
                     GOSUB MiscLeft                               ' Go adjust for Left

                  CASE "R"                                        ' Right
                     GOSUB MiscRight                              ' Go adjust for right

                  CASE "E"                                        ' End
                     MiscECol = LEN(RTRIM$(FMask))                ' Move right to text length column
                     me.MiscMark: CALL DWORD Pass: MExitMeth      ' Redraw the hilight and we're done

                  CASE "U", "D"                                   ' Up/Down?
                     me.MarkKill: CALL DWORD Pass: MExitMeth      ' Just kill and move the cursor

               END SELECT                                         '
            END IF                                                '

         CASE %CursLNote                                          ' FM Note?

            '----- If highlight is active
            IF ISFALSE IsMiscActive THEN                          ' If no active area right now, start one
               '----- No active bounds right now, Start one up
               MiscZone = Zone                                    ' Save the zone we're marking in
               MiscLine = CRow - FM_Top_File_Line + TopScrn       ' Save which line it is
               IF dir = "E" THEN                                  ' If END case
                  MiscSCol = ZLCol: MiscECol = LEN(RTRIM$(AFList(MiscLine).Note))
                  CALL DWORD Pass                                 ' Then just move the cursor
               ELSE                                               '
                  MiscSCol = ZLCol: MiscECol = ZLCol              '
               END IF                                             '
               OnMiscActive                                       ' We're starting a marked area
               me.MiscMark                                        ' Go Draw it

            ELSE                                                  '
               ' We have an active area already
               GOSUB KillOrMove                                   ' Handle the simple conditions

               '----- We're at a boundary, see how to adjust the bounds
               SELECT CASE dir                                    ' Adjust boundaries
                  CASE "L"                                        ' Left
                     GOSUB MiscLeft                               ' Go adjust for Left

                  CASE "R"                                        ' Right
                     GOSUB MiscRight                              ' Go adjust for right

                  CASE "E"                                        ' End
                     MiscECol = LEN(RTRIM$(AFList(MiscLine).Note))' Move right to text length column
                     me.MiscMark: CALL DWORD Pass: MExitMeth      ' Redraw the hilight and we're done

                  CASE "U", "D"                                   ' Up/Down?
                     me.MarkKill: CALL DWORD Pass: MExitMeth      ' Just kill and move the cursor

               END SELECT                                         '
            END IF                                                '

         CASE %CursData                                           ' In text data area
            me.AttrInvClearFind                                   ' Clear any Find HiLite
            '----- Eliminate the can't handle type
            CsrLinDX = 0                                          ' Reset possible hidden DX line location
            IF Prf.HexMode = &4  THEN CALL DWORD Pass: sDoBeep: MExitMeth' In HEX? Skip

            '----- Watch out for Swap overlap
            IF IsSwapActive THEN                                  ' Prevent Swap overlap
               TCol = ZLCol                                       '
               TCol += IIF(dir = "L", -1, 0)                      ' Create the 'next col'
               TCol += IIF(dir = "R", 1, 0)                       '
               TLin = ZLNum                                       ' Check the top line
               IF TLin >= SwapSLin AND TLin <= SwapELin THEN      ' Within bounds vertically?
                  IF TCol >= SwapSCol AND TCol <= SwapECol THEN sDoBeep: MExitMeth ' And horizontally?
               END IF                                             '
               j = SwapLines - 1                                  ' Get number of data lines in Swap group
               DO WHILE j                                         ' See if enough lines here
                  IF IsLData(TLin) THEN                           ' If a data line
                     DECR j                                       ' Reduce count
                  END IF                                          '
                  INCR TLin                                       ' Next line
                  IF IsLBottom(TLin) THEN sDoBeep: MExitMeth      '
               LOOP                                               '
               IF TLin >= SwapSLin AND TLin <= SwapELin THEN      ' Check the bottom lines
                  IF TCol >= SwapSCol AND TCol <= SwapECol THEN sDoBeep: MExitMeth ' And horizontally?
               END IF                                             '
            END IF                                                '

            '----- If highlight is active
            IF ISFALSE IsMarkActive THEN                          ' If no current
               IF IsLXClude(ZLNum) THEN MExitMeth                 ' Can't start on an X'd line
               IF dir <> "E" THEN                                 ' If not the END case
                  MarkSLin = ZLNum: MarkELin = ZLNum              ' Save 1 character area
                  MarkSCol = ZLCol: MarkECol = ZLCol              '
               ELSE                                               ' END
                  MarkSLin = ZLNum: MarkELin = ZLNum              ' Save area to end of text
                  MarkSCol = ZLCol: MarkECol = LEN(RTRIM$(LTxtG(ZLNum)))
                  CALL DWORD Pass                                 ' Then just move the cursor
               END IF                                             '
               IF IsSwapActive THEN                               ' If a SWAP is active
                  j = SwapLines - 1                               ' Get number of data lines in Swap group
                  DO WHILE j                                      ' See if enough lines here
                     IF IsLData(MarkELin) THEN                    ' If a data line
                        DECR j                                    ' Reduce count
                     END IF                                       '
                     MarkELin = MarkELin + 1                      ' Next line
                     IF IsLBottom(MarkELin) THEN sDoBeep: MExitMeth
                  LOOP                                            '
               END IF                                             '
               OnMarkActive                                       ' Start a marked area
               me.MarkScr                                         ' Draw it

            ELSE                                                  ' We already have a Marked area

               '----- Kill mark if outside current boundaries
               IF ZLNum < MarkSLin OR ZLNum > MarkELin OR _       ' Somehow outside it's boundaries?
                  ZLCol < MarkSCol OR ZLCol > MarkECol THEN       '
                  me.MarkKill: MExitMeth                          ' Kill it and leave cursor alone
               END IF                                             '

               '----- If within current bounds, just move the cursor
               IF ZLNum > MarkSLin AND ZLNum < MarkELin AND _     ' Somehow within the bounds?
                  ZLCol > MarkSCol AND ZLCol < MarkECol THEN _    '
                  CALL DWORD Pass: MExitMeth                      ' Just move the cursor

               '----- We're at a boundary, see how to adjust the bounds
               SELECT CASE dir                                    ' Adjust boundaries

                  '----- Move boundary left
                  CASE "L"                                        ' Left
                     IF ZLCol = MarkSCol AND MarkSCol > 1 THEN    ' At left and got room?
                        MarkSCol = MarkSCol - 1                   ' Move left boundary
                        me.MarkScr: CALL DWORD Pass: MExitMeth    ' Redraw the hilight
                     ELSEIF ZLCol = MarkECol AND MarkECol > MarkSCol THEN ' At right and have room
                        MarkECol = MarkECol - 1                   ' Move right boundary
                        me.MarkScr: CALL DWORD Pass: MExitMeth    ' Redraw the hilight
                     ELSE                                         '
                        CALL DWORD Pass: MExitMeth                ' Just move the cursor
                     END IF                                       '
                     IF IsMarkActive THEN me.MarkScr              ' Possibly redraw the highlight

                  '----- Move boundary right
                  CASE "R"                                        ' Right
                     IF ZLCol = MarkSCol AND MarkSCol < MarkECol THEN ' At left end and have room?
                        MarkSCol = MarkSCol + 1                   ' Move left boundary
                        me.MarkScr: CALL DWORD Pass: MExitMeth    ' Redraw the hilight
                     ELSEIF ZLCol = MarkECol THEN                 ' At right
                        MarkECol = MarkECol + 1                   ' Move right boundary
                        me.MarkScr: CALL DWORD Pass: MExitMeth    ' Redraw the hilight
                     ELSE                                         '
                        CALL DWORD Pass: MExitMeth                ' Just move the cursor
                     END IF                                       '

                  '----- Move boundary up
                  CASE "U"                                        ' Up
                     IF IsSwapActive THEN MExitMeth               ' Do nothing if SWAP active
                     IF ZLNum = MarkSLin AND MarkSLin > 2 THEN    ' At the top or mark area and have room?
                        OldLin = me.sSearch(MarkSLin)             ' Get where current top is
                        MarkSLin = me.CRPBack(%MVisData, MarkSLin, 1)' Move top boundary
                        NewLin = me.sSearch(MarkSLin)             ' Get where new top is
                        OldLin = OldLin - NewLin                  ' OldLin = # screen lines we went up
                        GOSUB DoDraw                              ' Draw, move cursor and exit
                     ELSEIF ZLNum = MarkELin AND MarkELin > MarkSLin THEN  ' At bottom and got room?
                        OldLin = me.sSearch(MarkELin)             ' Get where current bottom is
                        MarkELin = me.CRPBack(%MVisData, MarkELin, 1)' Move bottom boundary
                        NewLin = me.sSearch(MarkELin)             ' Get where new bottom is
                        OldLin = OldLin - NewLin                  ' OldLin = # screen lines we went up
                        GOSUB DoDraw                              ' Draw, move cursor and exit
                     ELSE                                         '
                        CALL DWORD Pass: MExitMeth                ' Just move the cursor
                     END IF                                       '

                  '----- Move boundary down
                  CASE "D"                                        ' Down
                     IF IsSwapActive THEN MExitMeth               ' Do nothing if SWAP active
                     IF ZLNum = MarkSLin AND MarkSLin < MarkELin THEN ' At the top and have room?
                        OldLin = me.sSearch(MarkSLin)             ' Get where current top is
                        MarkSLin = me.CRPFwd(%MVisData, MarkSLin, 1)' Move top boundary
                        NewLin = me.sSearch(MarkSLin)             ' Get where new top is
                        OldLin = NewLin - OldLin                  ' OldLin = # screen lines we went down
                        GOSUB DoDraw                              ' Draw, move cursor and exit
                     ELSEIF ZLNum = MarkELin AND MarkELin < LastLine - 1 THEN ' At bottom and got room?
                        OldLin = me.sSearch(MarkELin)             ' Get where current bottom is
                        MarkELin = me.CRPFwd(%MVisData, MarkELin, 1)' Move bottom boundary
                        NewLin = me.sSearch(MarkELin)             ' Get where new bottom is
                        OldLin = NewLin - OldLin                  ' OldLin = # screen lines we went down
                        IF OldLin < 0 THEN OldLin = 1             ' Adjust if off the bottom and scrolling
                        GOSUB DoDraw                              ' Draw, move cursor and exit
                     ELSE                                         '
                        CALL DWORD Pass: MExitMeth                ' Just move the cursor
                     END IF                                       '

                  '----- Move boundary to End of text
                  CASE "E"                                        ' End
                     MarkECol = LEN(RTRIM$(LTxtG(ZLNum)))         ' Move right to text length column
                     me.MarkScr                                   ' Redraw the hilight
                     CALL DWORD Pass: MExitMeth                   ' Then just move the cursor

               END SELECT                                         '
            END IF                                                '

         CASE ELSE
            CALL DWORD Pass: sDoBeep: MExitMeth                   ' Not in Data. Skip
      END SELECT                                                  '
      MExitMeth                                                   '

   KillOrMove:
      '----- Kill mark if outside current boundaries
      IF ZLCol < MiscSCol OR ZLCol > MiscECol THEN _              ' Outside it's boundaries?
         me.MarkKill: MExitMeth                                   ' Kill it and leave cursor alone

      '----- Kill mark if Note zone and different line
      IF MiscZone = %CursLNote AND MiscLine <> CRow - FM_Top_File_Line + TopScrn THEN _     ' Outside it's boundaries?
         me.MarkKill: MExitMeth                                   ' Kill it and leave cursor alone

      '----- If within current bounds, just move the cursor
      IF ZLCol > MiscSCol AND ZLCol < MiscECol THEN _             ' Within the bounds?
         CALL DWORD Pass: MExitMeth                               ' Just move the cursor

      RETURN

   AdjMiscCol:
      IF ColAdj < 0 AND MiscSCol > 1 THEN                         ' At left and got room
         MiscSCol = MiscSCol + ColAdj                             ' Move left boundary
      ELSEIF ColAdj > 0 AND ZLCol = MiscSCol AND MiscSCol <> MiscECol THEN              ' At left edge and going right
         MiscSCol = MiscSCol + ColAdj                             ' Move left boundary
      ELSEIF ColAdj > 0 THEN                                      ' Going right
         MiscECol = MiscECol + ColAdj                             ' Move right boundary
      ELSE                                                        '
         sDoBeep                                                  '
      END IF                                                      '
      me.MiscMark: CALL DWORD Pass: MExitMeth                     ' Redraw the hilight and we're done
      RETURN


   DoDraw:
      me.MarkScr                                                  ' Redraw the hilight
      DO WHILE OldLin > 0                                         ' Adjust the cursor correctly
         CALL DWORD Pass                                          ' Then just move the cursor
         DECR OldLin                                              '
      LOOP                                                        '
      MExitMeth                                                   '
      RETURN

   MiscLeft:
      IF ZLCol = MiscSCol AND MiscSCol = MiscECol THEN            ' Sitting on one char field?
         MiscSCol = MAX(1, MiscSCol - 1)                          ' Adjust left end
      ELSEIF ZLCol = MiscSCol THEN                                ' Sitting on left end of string
         MiscSCol = MAX(1, MiscSCol - 1)                          ' Adjust left end
      ELSE                                                        ' Must be on Right Boundary
         MiscECol = MAX(MiscSCol, MiscECol - 1)                   ' Adjust left end
      END IF                                                      '
      me.MiscMark: CALL DWORD Pass: MExitMeth                     ' Redraw the hilight and we're done
      RETURN

   MiscRight:
      IF ZLCol = MiscSCol AND MiscSCol = MiscECol THEN            ' Sitting on a one char field?
         MiscECol = MiscECol + 1                                  ' Adjust right end
      ELSEIF ZLCol = MiscSCol THEN                                ' Sitting on Left boundary
         MiscSCol = MIN(MiscECol, MiscSCol + 1)                   ' Adjust left end
      ELSE                                                        ' Must be on Right Boundary
         MiscECol = MiscECol + 1                                  ' Adjust right end
      END IF                                                      '
      me.MiscMark: CALL DWORD Pass: MExitMeth                     ' Redraw the hilight and we're done
      RETURN                                                      '
   END METHOD

   METHOD  krNewLine(OPTIONAL ns AS LONG)
   '---------- Next line on screen
   LOCAL i, j, noscroll, ix AS LONG
      MEntry
      noscroll = IIF(ISMISSING(ns), %False, ns)                   ' Pick up No Scroll flag
      CsrLinDX = 0                                                ' Reset possible hidden DX line location

      '----- Top of screen area
      IF CRow < 3 + Prf.Cols THEN                                 ' Command line, line 2 or COLS line
         me.CsrRow = 3 + Prf.Cols: me.CsrCol = 1: MExitMeth       ' Go to 1st data line
      END IF                                                      ' Else in data area

      '----- If not in text data area, go to Command line
      ix = sGetIX(CRow)                                           ' Locate data line cursor was on
      IF ix = 0 THEN                                              ' Not in data area?
         SetCmd                                                   ' Cursor to the command line
         MExitMeth                                                '
      END IF                                                      '

      '----- If in Text data area go to next line
      IF ix < 0 THEN ix = sGetIX(CRow - ABS(ix))                  ' If Hex extra line, get real line
      me.CsrRowAdd(1)                                             ' Simple INCR
      IF CRow > gwScrHeight OR IsLBottom(ix) THEN                 ' At bottom line yet?
         IF ix = 0 OR IsLBottom(ix) OR ISTRUE noscroll THEN       ' Not in data area or noscroll requested?
            SetCmd                                                ' Cursor to the command line
            MExitMeth                                             '
         END IF                                                   '
         ix = me.CRPFwd(%mVisible, ix, 1)                         ' Move to next visible line
         me.CurSetReq(%Position, ix, 0, %True)                    ' Set cursor set attempt
         TopScrn = me.CRPFwd(%mVisible, TopScrn, 1)               ' Move TopScrn by scroll amount
         OnInsClnSupp                                             ' Suppress Insert cleanup for 1 Attn pass
         DoSet(%Refresh)                                          ' Have it looked at
      ELSE                                                        '
         me.CsrCol = 1                                            '
      END IF                                                      '
      MExit
   END METHOD

   METHOD  krPaste()
   '---------- Paste text at cursor location
   LOCAL ix, i, j AS LONG
   LOCAL ln, t, tt, MyDLM AS STRING

      MEntry
      '----- Get the Clipboard text
      CsrLinDX = 0                                                ' Reset possible hidden DX line location
      sReadClipboard(t, MyDLM, %False)                            ' Go get Clipboard, keep the file
      IF LEN(t) < 1 THEN MExitMeth                                ' Error if nothing there
      me.AttrInvClearFind                                         ' Clear any Find HiLite

      IF RIGHT$(t, LEN(MyDLM)) = MyDlm THEN t = LEFT$(t, LEN(t) - LEN(MyDlm)) ' Remove trailing MyDlm to get PARSECOUNT accurate
      IF PARSECOUNT(t, MyDlm) > 1 AND ISTRUE IsCData THEN         ' Potential BlockPaste?
         me.krBPaste                                              ' Yes, pass it off
         MExitMeth                                                ' Done here
      END IF                                                      '

      CsrLinDX = 0                                                ' Reset possible hidden DX line location

      '----- Get first line
      ln = EXTRACT$(t, MyDLM)                                     ' Get 1st line from Clipboard
      me.MarkKill                                                 ' Kill any active block select
      me.SwapKill                                                 '

      '----- Paste into Command line
      IF IsCCmnd THEN                                             ' In the Command area?
         tt = pCommand                                            ' Get working copy
         IF IsNotNsrtFlag THEN                                    ' Do the non-Insert mode handling
            IF LEN(tt) - (CCol - 10 + COffset) < LEN(ln) THEN     '
               tt = tt + SPACE$(LEN(ln))                          ' Lengthen if needed
            END IF                                                '
            MID$(tt, CCol - 10 + COffset, LEN(ln)) = ln           ' Overlay the CBD text
         ELSE                                                     ' Do the Insert method
            tt = LEFT$(tt, CCol - 11 + COffset) + ln + MID$(tt, CCol - 10 + COffset)
         END IF                                                   '
         PCommand = tt                                            ' Copy back
         i = CCol + COffset + LEN(ln)                             ' Move cursor
         IF i > pCmdLen + COffset THEN                            ' Is it to the right of the Window?
            COffset = i - (pCmdLen / 2)                           ' Adjust offset
         ELSEIF i < COffset THEN                                  ' Is it to the left of the Window?
            COffset = MAX(0, i - (pCmdLen / 2))                   ' Adjust offset
         END IF                                                   '
         me.WindowCmd                                             ' Print the whole string again
         me.CsrCol = i - COffset                                  ' Move cursor

      '----- Paste into Scroll Amt (Weird but ...)
      ELSEIF IsCScrl THEN                                         ' Scroll area?
         tt = Prf.Scroll                                          ' Get working copy
         IF IsNotNsrtFlag THEN                                    ' Do the non-Insert mode handling
            IF LEN(tt) - (CCol - 20 - pCmdLen - 1) < LEN(ln) THEN tt = tt + SPACE$(LEN(ln)) ' Lengthen if needed
            MID$(tt, CCol - 20 - pCmdLen, LEN(ln)) = ln           ' Overlay the CBD text
         ELSE                                                     ' Do the Insert method
            tt = LEFT$(tt, CCol - 20 - pCmdLen - 1) + ln + MID$(tt, CCol - 20 - pCmdLen)
         END IF                                                   '
         Prf.Scroll = LEFT$(tt, 4)                                ' Copy back
         sPrint (LSET$(Prf.Scroll, 4), $$TxtLo, 1, 19 + pCmdLen)  ' Print the whole string again

      '----- Paste into Line Number area
      ELSEIF IsCLinN THEN                                         ' Line Number area?
         ix = sGetIX(CRow)                                        ' Get IX of real data line
         IF ix = 0 THEN sDoBeep: MExitMeth                        ' Below last line? Bail out
         IF me.IsTouchable(ix, CCol) THEN                         ' If 1st interaction in the line number area
            me.LFlagBitOff(ix, %EQChange)                         ' Reset ==CHG>
            LLCtlS(ix) = $BlankLNo: CCol = 1                      ' Blank the Line number
            sPrint (BYCOPY LEFT$(LLCtlG(ix), ENV.LinNoSize), $$LnoHi, CRow, 1)' ReDisplay it
         END IF                                                   '

         LLCtlS(ix) = LEFT$(ln, 8)                                ' Overlay the data into the LineNum string
         me.TTblAdd(ix, TRIM$(LLCtlG(ix)))                        ' Update touched items
         sPrint (LEFT$(LLCtlG(ix), ENV.LinNoSize), $$LnoHi, CRow, 1) ' Print the whole string again
         me.CsrCol = 1                                            ' Move cursor

      '----- Paste into the Text data area
      ELSEIF IsCData THEN                                         ' In the Data Area?
         IF IsBrowse THEN sDoBeep: MExitMeth                      ' Not allowed in Browse mode
         j = S(CRow)                                              ' Get raw line index type
         ix = sGetIX(cRow)                                        ' Get IX of real data line
         IF ix < 0 THEN ix = sGetIX(CRow - ABS(ix))               ' Convert to real line index
         IF ix = 0 OR j < 0 THEN sDoBeep: MExitMeth               ' Bail out if below last line or a Hex Line
         IF ISTRUE (LFlagG(ix) AND %NonTypable) THEN              ' Ignore non-typable lines
            sDoBeep: MExitMeth                                    '
         END IF                                                   '
         me.ModSet(ix)                                            ' Remember we changed something
         i = CCol - gLNPadCol + Offset                            ' Calc index into string
         tt = LTxtG(ix)                                           ' Get working copy of text

         '----- Do it either Insert or non-Insert way
         IF IsNotNsrtFlag THEN                                    ' Do the non-Insert mode handling
            IF i + LEN(ln) - 1 > LEN(tt) THEN tt = LSET$(tt , i + LEN(ln) - 1)  ' Lengthen if needed
            MID$(tt, i, LEN(ln)) = ln                             ' Overlay the CBD text
         ELSE                                                     ' Do the Insert method
            tt = me.DataInsert(tt, ln, i)                         ' Insert the data
            me.LAttrAdjust(ix, i, LEN(ln))                        ' Adjust any Attr hilite
         END IF                                                   '

         IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) AND IsLData(ix) THEN ' CAPS ON?
            tt = UUCASE(tt)                                       ' UC if CAPS is ON
         END IF                                                   '
         me.LTxtSet(ix, tt)                                       ' Save it back
         me.AttrScan(ix)                                          ' Recolorize

         '----- Print it, do extra of course if Hex mode
         me.DispLine(ix, CRow)                                    ' Re-display the line
         IF Prf.HexMode <> &1 THEN                                   ' If hex mode
            IF CRow + 1 <= gwScrHeight THEN                       ' Make sure we don't go off the bottom
               sPrint (me.WindowHexTop(tt), $$TxtLo, CRow + 1, gLNData1) ' Print upper hex
            me.MarkReDraw(CRow + 1)                               ' Re-draw the MARK lines
            END IF                                                '
            IF CRow + 2 <= gwScrHeight THEN                       ' Make sure we don't go off the bottom
               sPrint (me.WindowHexBottom(tt), $$TxtLo, CRow + 2, gLNData1) ' Print lower hex
               me.MarkReDraw(CRow + 2)                            ' Re-draw the MARK lines
            END IF                                                '
         END IF                                                   '

         i += LEN(ln)                                             ' Move cursor
         IF i > Offset + gDataLen - 1 THEN                        ' Off current screen?
            me.CurSetReq(%Position, ix, i, %True)                 ' No, Set cursor set attempt
            DoSet(%Refresh)                                       ' Have it looked at
         ELSE                                                     '
            me.CsrCol = i - Offset + gLNPadCol                    '
         END IF                                                   '

      ELSE                                                        ' Anything else (CursBad)
         sDoBeep                                                  ' Beep it
      END IF
      MExit
   END METHOD

   METHOD  krPenBlue()
      gkeyPrimOper = "BLUE"
      me.krPenCommon
   END METHOD

   METHOD  krPenGreen()
      gkeyPrimOper = "GREEN"
      me.krPenCommon
   END METHOD

   METHOD  krPenRed()
      gkeyPrimOper = "RED"
      me.krPenCommon
   END METHOD

   METHOD  krPenYellow()
      gkeyPrimOper = "YELLOW"
      me.krPenCommon
   END METHOD

   METHOD  krPenStd()
      gkeyPrimOper = "STD"
      me.krPenCommon
   END METHOD

   METHOD  krPenCommon()
   '---------- Set Pen color for a hi-lite block
   LOCAL x, i, j, FNum, clr AS LONG, lclTxt, a, f, t, fn, clrname AS STRING
      clrname = uucase(IIF$(gKeyPrimOper = "", "STD", gKeyPrimOper)) ' Get the requested color
      IF clrname = "STD" THEN                                     ' If STD, set scheme 0
         clr = 0                                                  '
      ELSE                                                        ' else look it up
         ARRAY SCAN nHiLites(), = clrname, TO clr                 ' See if a HiLite Name
      END IF                                                      '
      me.MarkKill                                                 ' Reset the select frame

      '---- OK, process the lines
      FOR x = MarkSLin TO MarkELin                                ' Loop through the requested range
         IF ISFALSE (LFlagG(x) AND %NonTypable) THEN              ' Just typable lines
            IF MarkScol > LAttrGLen(x) THEN ITERATE FOR           ' If past end of line, do nothing
            i = MIN(MarkECol, LEN(L(x).@LTxt))                    ' Setup end of mark section
            me.AttrHiLiteSet(x, MarkScol, i, clr)                 ' Set in the requested color
         END IF                                                   '
      NEXT x                                                      '
      DoSet(%Refresh)                                             ' Have it looked at
   END METHOD

   METHOD  krRecord()
   '---------- Start/Stop Keyboard recording
      IF ISTRUE gKbdRecFlag THEN                                  ' If ON, then turn off
         gKbdRecFlag = %False                                     ' Turn on the flag
         REPLACE "(Record)" WITH "" IN gKbdRecording              ' Remove the OFF command
         IF ISNOTNULL(TRIM$(gKbdRecording)) THEN                  ' See if we recorded anything
            sWinclip_set(gKbdRecording)                           ' Stuff it into Clipboard
            scError(%eNone, "KB Recording saved to the Clipboard")' Say we've done so
            DoSet(%Attention)                                     ' Get it issued
         END IF                                                   '
      ELSE                                                        ' else turn it on
         gKbdRecording = ""                                       ' Start string as null
         gKbdRecFlag = %True                                      ' Turn on the flag
      END IF                                                      '
      sDoStatusBar($SBMisc)                                       ' re-Do the StatusBar Misc box
   END METHOD

   METHOD  krResetInsert()
   '---------- Insert key OFF
      IF IsNsrtFlag THEN OnNsrtSave ELSE OffNsrtSave              ' Save it
      OffNsrtFlag                                                 ' Set it OFF
   END METHOD

   METHOD  krRestoreCursor()
   '---------- Restore Cursor location
   LOCAL i, j AS LONG
      IF SavCurLin = 0 THEN                                       ' No line number?
         SetCmd                                                   ' Cursor to the command line
      ELSE                                                        '
         i = me.sSearch(SavCurLin)                                ' See if location is on screen
         IF i < 0 THEN                                            ' Not on screen
            SetCmd                                                ' Cursor to the command line
         ELSE                                                     '
            me.CsrRow = i                                         ' Put cursor back
            me.CsrCol = IIF(SavCurCol = 0, 1, SavCurCol + gLNPadCol - OffSet) '
         END IF                                                   '
      END IF                                                      '
   END METHOD

   METHOD  krRestoreInsert()
   '---------- Restore Insert mode
      IF IsNsrtSave THEN OnNsrtFlag ELSE OffNsrtFlag              ' Restore it
      OffNsrtData                                                 ' Always set DataInsert off
   END METHOD

   METHOD  krRight()
   '---------- Increment Column
   LOCAL ix, cIncr, OldOffset, Hexrow AS LONG
   LOCAL t AS STRING
      MEntry
      cIncr = IIF(VAL(gKeyPrimOper) > 0, VAL(gKeyPrimOper), 1)    ' Set the INCR amount?
      me.CsrColAdd(cIncr)                                         ' Move cursor
      IF CCol = gLNPadCol THEN me.CsrColAdd(1)                    ' Make Col gLNPadCol an Auto-Skip field

      '----- Cursor in the Command line
      IF IsCCmnd THEN                                             ' In the command area
         IF CCol > 10 + pCmdLen THEN                              ' Past end of line?
            OldOffset = COffset                                   ' Save our original Offset
            COffset = CCol - (10 + pCmdLen) + COffSet             ' Adjust the offset
            me.CsrCol = CCol - (COffSet - OldOffset)              ' Position the cursor
            me.WindowCmd                                          ' Re-do the command line
            MExitMeth                                             '
         END IF                                                   '
      END IF                                                      '

      '----- Cursor not in the Text Data area
      IF ISFALSE ENV.KBHorScroll OR ISFALSE IsCData THEN          ' No KeybScroll or not in data area
         IF CCol > ENV.ScrWidth THEN                              ' Column wrap?
            me.CsrCol = 1: me.CsrRowAdd(1)                        ' Yes, adjust for it
            IF CRow > gwScrHeight THEN me.CsrRow = 1              ' also for Row wrap if needed
         END IF                                                   '

      '----- Cursor in the Text area
      ELSE                                                        '
         IF CCol > ENV.ScrWidth THEN                              ' YES - At right side of screen?
            me.CsrColSub(cIncr)                                   ' Yes, Undo the INCR for now
            ix = sGetIX(CRow)                                     ' Locate data line cursor was on
            IF ix = 0 THEN MExitMeth                              ' Exit if below bottom line
            IF ix < 0 THEN                                        ' A Hex line?
               HexRow = ABS(ix)                                   ' Save the hex line
               ix = sGetIX(CRow - ABS(ix))                        ' Get real data line
            END IF
            me.CurSetReq(%Position, ix, CCol + Offset - gLNPadCol + cIncr, %False, 0, HexRow) ' Set cursor set attempt
            me.OffsetAdd(cIncr)                                   ' Bump the offset
            OnInsClnSupp                                          ' Suppress Insert cleanup for 1 Attn pass
            DoSet(%Refresh)                                       ' Have it looked at
         END IF                                                   '
      END IF                                                      '
      MExit                                                       '
   END METHOD

   METHOD  krSaveCursor()
   '---------- Save Cursor location
   LOCAL i, j AS LONG
      SELECT CASE CRow                                            ' Do tabbing based on where we are

         '----- Sitting in header area
         CASE 1, 2, 2 + Prf.Cols                                  ' Top Rows?
            SavCurCol = 11                                        ' Set for Command area
            SavCurLin = 0                                         '

         '----- Sitting in text area
         CASE ELSE                                                ' In data area
            j = sGetIX(CRow)                                      ' Get Line number of cursor line
            IF j < 0 THEN j = sGetIX(CRow - ABS(j))               ' Get real data line
            IF j = 0 THEN                                         ' Not a real text line?
               SavCurCol = 11                                     ' Set for Command area
               SavCurLin = 0                                      '
            ELSE                                                  '
               SavCurLin = j                                      ' Save line number
               SavCurCol = IIF(CCol < gLNData1, 0, CCol - gLNPadCol + Offset) ' Save column
            END IF                                                '
      END SELECT                                                  '
   END METHOD

   METHOD  krScrollD
   '---------- Scroll Down
      pCommand = "ISCROLL DOWN 1 K"                               ' Create internal command
      DoSet(%Attention)                                           ' Have it looked at
   END METHOD

   METHOD  krScrollL
   '---------- Scroll Left
      pCommand = "ISCROLL LEFT 1 K"                               ' Create internal command
      DoSet(%Attention)                                           ' Have it looked at
   END METHOD

   METHOD  krScrollR
   '---------- Scroll Right
      pCommand = "ISCROLL RIGHT 1 K"                              ' Create internal command
      DoSet(%Attention)                                           ' Have it looked at
   END METHOD

   METHOD  krScrollU
   '---------- Scroll Up
      pCommand = "ISCROLL UP 1 K"                                 ' Create internal command
      DoSet(%Attention)                                           ' Have it looked at
   END METHOD

   METHOD  krSentenceCase()
   '---------- Sentence case a block
   DIM words() AS STRING
   LOCAL Txt1, txt2 AS STRING
   LOCAL i, j, k, m, x, scol, ecol, Caps AS LONG

      MEntry
      Caps = %True                                                ' 1st word overall gets capitalized
      me.AttrInvClearFind                                         ' Clear any Find HiLite

      '----- Set the cursor and do the actual casing
      me.CurSetReq(%Low, MarkELin, 1 + MarkECol + 1, %True)       ' Set cursor set attempt
      FOR x = MarkSLin TO MarkELin                                ' Loop for each line
         IF IsLData(x) OR ISTRUE IsLNote(x) THEN                  ' Only Data/note lines
            me.ModSet(x)                                          ' Remember we changed something
            Txt1 = LLCASE(MID$(LTxtG(x), MarkSCol TO MarkECol))   ' Get copy of substring to process
            GOSUB DoAChunk                                        ' Go process this piece
            txt2 = LTxtG(x)                                       ' Get whole text line
            MID$(txt2, MarkSCol TO MarkECol) = Txt1               ' Stuff back in the processed txt
            me.LTxtSet(x, txt2)                                   ' Replace the whole line
            me.AttrScan(x)                                        ' Recolorize
         END IF                                                   '
      NEXT x                                                      '
      DoSet(%Refresh)                                             ' Have it looked at
      me.MarkKill                                                 ' Reset the select frame
      MExitMeth

   '----- Do the Sentence case stuff
   DoAChunk:
      i = PARSECOUNT(Txt1, " ")                                   ' Count words
      REDIM words(1 TO i) AS STRING                               ' Make table correct size
      PARSE Txt1, words(), " "                                    ' Parse the words into the table
      m = 1                                                       ' Start scan at 1
      FOR j = 1 TO i                                              ' Loop for each word
         IF ISNOTNULL(words(j)) THEN                              ' Ignore null 'words'
            k = INSTR(m, Txt1, words(j))                          ' Find location of this word
            IF caps THEN                                          ' Capitalize this word?
               IF k THEN MID$(Txt1, k, 1) = UUCASE(MID$(Txt1, k, 1))' Uppercase the 1st letter
               Caps = %False                                      ' Off till next period
            END IF                                                '
            IF RIGHT$(words(j), 1) = "."  OR _                    ' Look for sentence trigger
               RIGHT$(words(j), 1) = "?"  OR _                    '
               RIGHT$(words(j), 1) = "!"  THEN                    '
               Caps = %True                                       ' Restart sentence if required
            END IF                                                '
            m = k + 1                                             ' Start scan adjust
         END IF                                                   '
      NEXT j                                                      '
      RETURN                                                      '
   END METHOD

   METHOD  krSetInsert()
   '---------- Set Insert key ON
   IF IsNsrtFlag THEN OnNsrtSave ELSE OffNsrtSave                 ' Save it
   OnNsrtFlag                                                     ' Set the mode
   OffNsrtData                                                    ' Always set DataInsert off
   sDoStatusBar($SBInsOvr)                                        ' re-Do the StatusBar Ins/Ovr box
   END METHOD

   METHOD  krSwap()
   '---------- Swap command
   LOCAL i, j, swapLnum, markLnum, SNum AS LONG
   LOCAL dSwapLine, dMarkLine, dSwapText, dMarkText, dPart1, dPart2, dPart3, dPart4 AS STRING
   LOCAL cSwapLine, cMarkLine, cSwapText, cMarkText, cPart1, cPart2, cPart3, cPart4 AS WSTRING
   LOCAL rightECol AS LONG
      MEntry
      me.AttrInvClearFind                                         ' Clear any Find HiLite

      rightECol = MAX(SwapECol, MarkECol)

      '----- If both areas marked, we can do the Swap
      IF IsMarkActive AND IsSwapActive THEN                       ' We have two marked areas ready

         '----- Init for line loop
         me.SwapKill                                              ' To remove UL so they don't get copied
         swapLnum = SwapSLin: markLnum = MarkSLin                 ' Point at the 1st two lines

         FOR SNum = 1 TO SwapLines                                ' Loop for the number of lines
            '----- Extract the two strings

            dSwapLine = LSET$(LTxtG(swapLnum), MAX(LEN(L(swapLnum).@LTxt), rightECol)) ' Get the data text from the Swap line
            cSwapLine = LSET$(LAttrG(swapLnum), MAX(LEN(L(swapLnum).@LTxt), rightECol)) ' Get the attr text from the Swap line

            dSwapText = MID$(dSwapLine, SwapSCol TO SwapECol)     ' Extract the data columns
            cSwapText = MID$(cSwapLine, SwapSCol TO SwapECol)     ' Extract the attr columns

            me.ModSet(swapLnum)                                   ' Remember we changed something

            dMarkLine = LSET$(LTxtG(markLNum), MAX(LEN(L(markLnum).@LTxt), rightECol)) ' Get the data text from the Mark line
            cMarkLine = LSET$(LAttrG(markLnum), MAX(LEN(L(markLnum).@LTxt), rightECol)) ' Get the attr text from the Mark line

            dMarkText = MID$(dMarkLine, MarkSCol TO MarkECol)     ' Extract the data columns
            cMarkText = MID$(cMarkLine, MarkSCol TO MarkECol)     ' Extract the attr columns

            me.ModSet(markLnum)                                   ' Remember we changed something

            IF swapLnum <> markLnum THEN                          ' We're dealing with two different lines

               '----- Handle the two different line combination

               IF SwapSCol = 1 THEN                               ' If old Swap text is at LH end
                  dSwapLine = dMarkText + MID$(dSwapLine, SwapECol + 1) ' Just replace with the Mark text
                  cSwapLine = cMarkText + MID$(cSwapLine, SwapECol + 1) ' Just replace with the Mark text

               ELSE                                               '
                  '----- Insert dMarkText to replace dSwapText text
                  dSwapLine = LEFT$(dSwapLine, SwapSCol - 1) + dMarkText + MID$(dSwapLine, SwapECol + 1)
                  cSwapLine = LEFT$(cSwapLine, SwapSCol - 1) + cMarkText + MID$(cSwapLine, SwapECol + 1)

               END IF                                             '

               me.LTxtSet (swapLnum, dSwapLine)                   ' Save the Swap line back
               LAttrS(swapLnum) = cSwapLine                       ' Save the Swap line back
               me.AttrScan(swapLnum)                              ' Recolorize

               IF MarkSCol = 1 THEN                               ' If old Mark text is at LH end
                  dMarkLine = dSwapText + MID$(dMarkLine, MarkECol + 1) ' Just replace with the Swap text
                  cMarkLine = cSwapText + MID$(cMarkLine, MarkECol + 1) ' Just replace with the Swap text

               ELSE                                               '
                  '----- Insert dMarkText to replace dSwapText text
                  dMarkLine = LEFT$(dMarkLine, MarkSCol - 1) + dSwapText + MID$(dMarkLine, MarkECol + 1)
                  cMarkLine = LEFT$(cMarkLine, MarkSCol - 1) + cSwapText + MID$(cMarkLine, MarkECol + 1)

               END IF                                             '

               me.LTxtSet (markLnum, dMarkLine)                   ' Save the Swap line back
               LAttrS(markLnum) = cMarkLine                       ' Save the Swap line back
               me.AttrScan(markLnum)                              ' Recolorize

            ELSE                                                  ' The Mark and Swap are on the same line

               '----- Swap on a single line
               IF SwapSCol < MarkSCol THEN                        ' Swap is left of Mark

                  '----- Swap area is Left of the Mark area

                  dPart1 = LEFT$(dSwapLine, SwapSCol - 1)         ' Get part 1, the left of Swap string
                  cPart1 = LEFT$(cSwapLine, SwapSCol - 1)         ' Get part 1, the left of Swap string

                  '----- Get part 2, between the strings

                  dPart2 = IIF$(SwapEcol + 1 = MarkSCol, "", MID$(dSwapLine, SwapECol + 1, MarkSCol - SwapECol - 1))
                  cPart2 = IIF$(SwapEcol + 1 = MarkSCol, "", MID$(cSwapLine, SwapECol + 1, MarkSCol - SwapECol - 1))

                  dPart3 = MID$(dSwapLine, MarkECol + 1)          ' Get part 3, After the strings
                  cPart3 = MID$(cSwapLine, MarkECol + 1)          ' Get part 3, After the strings

                  dSwapLine = BUILD$(dPart1, dMarkText, dPart2, dSwapText, dPart3) ' Rebuild the line
                  cSwapLine = BUILD$(cPart1, cMarkText, cPart2, cSwapText, cPart3) ' Rebuild the line

               ELSE                                               ' Mark is left of swap

                  '----- Swap area is Right of the Mark area

                  dPart1 = LEFT$(dSwapLine, MarkSCol - 1)         ' Get part 1, the left of Mark string
                  cPart1 = LEFT$(cSwapLine, MarkSCol - 1)         ' Get part 1, the left of Mark string

                  '----- Get part 2, between the strings

                  dPart2 = IIF$(MarkEcol + 1 = SwapSCol, "", MID$(dSwapLine, MarkECol + 1, SwapSCol - MarkECol - 1))
                  cPart2 = IIF$(MarkEcol + 1 = SwapSCol, "", MID$(cSwapLine, MarkECol + 1, SwapSCol - MarkECol - 1))

                  dPart3 = MID$(dSwapLine, SwapECol + 1)          ' Get part 3, After the strings
                  cPart3 = MID$(cSwapLine, SwapECol + 1)          ' Get part 3, After the strings

                  dSwapLine = BUILD$(dPart1, dSwapText, dPart2, dMarkText, dPart3) ' Rebuild the line
                  cSwapLine = BUILD$(cPart1, cSwapText, cPart2, cMarkText, cPart3) ' Rebuild the line

               END IF                                             '

               me.LTxtSet (swapLnum, dSwapLine)                   ' Save the Swap line back
               LAttrS(swapLnum) = cSwapLine                       ' Save the Swap line back
               me.AttrScan(swapLnum)                              ' Recolorize

            END IF                                                '

            swapLnum = me.CRPFwd(%mOData, swapLnum, 1)            ' Move to next line
            markLnum = me.CRPFwd(%mOData, markLnum, 1)            ' Move to next line

         NEXT SNum                                                ' Loop till all lines in block done

         me.MarkKill                                              ' Kill the mark area

         IF SwapSLin < MarkSLin THEN                              ' Set the cursor to lowest line number
            me.CurSetReq(%Position, SwapSLin, 1, %True)           ' Set cursor set attempt
         ELSE                                                     '
            me.CurSetReq(%Position, MarkSLin, 1, %True)           ' Set cursor set attempt
         END IF                                                   '

         DoSet(%Refresh)                                          ' Have it looked at

      '----- Only Mark area active, start a Swap activity
      ELSEIF IsMarkActive AND ISFALSE IsSwapActive THEN           ' We have a MARK but no SWAP, start a Swap
         SwapSCol = MarkSCol                                      ' Copy marked area as a swap area
         SwapECol = MarkECol                                      '
         SwapSLin = MarkSLin                                      '
         SwapELin = MarkELin                                      '
         j = 0                                                    ' Clear counter

         FOR i = SwapSLin TO SwapELin                             ' Count data lines
            IF IsLData(i) THEN INCR j                             '
         NEXT i                                                   '

         SwapLines = j                                            ' Save # of data lines
         me.MarkKill                                              ' Kill the mark area
         OnSwapActive                                             ' Make Swap active
         me.SwapScr                                               ' Draw the underline

      '----- No Mark area, See if Move mode, or Kill time
      ELSEIF ISFALSE IsMarkActive AND IsSwapActive THEN           ' We have swap and no Mark, See if maybe Move mode

         IF IsCData THEN                                          ' Cursor In the Data Area?
            j = sGetIX(CRow)                                      ' Get line number of the cursor
            IF j < 0 THEN j = sGetIX(CRow - ABS(j))               ' Convert to real line index
            IF j = 0 THEN GOSUB KillSwap: MExitMeth               ' Bail out if below last line
            i = CCol - gLNPadCol + Offset                         ' Calc column index into line

            IF j >= SwapSLin AND j <= SwapELin AND _              ' Cursor within the SWAP area?
               i >= SwapSCol AND i <= SwapECol THEN GOSUB KillSwap: MExitMeth ' Kill if

            '----- This looks like a SWAP Move mode request
            '----- Init for line loop
            me.SwapKill                                           ' Remove UL so it doesn't get copied
            swapLnum = SwapSLin: markLnum = j                     ' Point at the 1st two lines

            FOR SNum = 1 TO SwapLines                             ' Loop for the number of lines in Swap area

               '----- Extract the two lines

               dSwapLine = LTxtG(swapLnum)                        ' Get the text from the Swap line
               cSwapLine = LAttrG(swapLnum)                       ' Get the text from the Swap line

               dMarkLine = LTxtG(markLnum)                        ' Get the text from the Mark line
               cMarkLine = LAttrG(markLnum)                       ' Get the text from the Mark line

               IF LEN(dSwapLine) < j THEN dSwapLine = LSET$(dSwapLine, j) ' Make sure lines are as long as the mark point
               IF LEN(cSwapLine) < j THEN cSwapLine = LSET$(cSwapLine, j) '

               IF LEN(dMarkLine) < j THEN dMarkLine = LSET$(dMarkLine, j) '
               IF LEN(cMarkLine) < j THEN cMarkLine = LSET$(cMarkLine, j) '

               IF swapLnum <> markLnum THEN                       ' We're dealing with two different lines

                  dSwapText = MID$(dSwapLine, SwapSCol TO SwapECol) ' Extract the swap data
                  cSwapText = MID$(cSwapLine, SwapSCol TO SwapECol) ' Extract the swap data

                  dSwapLine = LEFT$(dSwapLine, SwapSCol - 1) + MID$(dSwapLine, SwapECol + 1) ' Remove if from the line
                  cSwapLine = LEFT$(cSwapLine, SwapSCol - 1) + MID$(cSwapLine, SwapECol + 1) ' Remove if from the line

                  me.LTxtSet (swapLnum, dSwapLine)                ' Save the Swap line back
                  LAttrS(swapLnum) = cSwapLine                    ' Save the Swap line back
                  me.AttrScan(swapLnum)                           ' Recolorize

                  me.ModSet(swapLnum)                             ' Remember we changed something

                  '----- Handle the two different line combination

                  IF i = 1 THEN                                   ' If insert point is at LH end
                     dMarkLine = dSwapText + dMarkLine            ' Just insert it
                     cMarkLine = cSwapText + cMarkLine            ' Just insert it

                  ELSE                                            '
                     dMarkLine = LEFT$(dMarkLine, i - 1) + dSwapText + MID$(dMarkLine, i) ' Insert dMarkText at the cursor column
                     cMarkLine = LEFT$(cMarkLine, i - 1) + cSwapText + MID$(cMarkLine, i) ' Insert dMarkText at the cursor column

                  END IF                                          '

                  me.LTxtSet (markLnum, dMarkLine)                ' Save the Swap line back
                  LAttrS(markLnum) = cMarkLine                    ' Save the Swap line back
                  me.AttrScan(markLnum)                           ' Recolorize

                  me.ModSet(markLnum)                             ' Remember we changed something

               ELSE                                               ' The Mark and Swap are on the same line

                  '----- Swap on a single line
                  IF SwapSCol < i THEN                            ' Swap is left of Mark

                     '----- Swap area is Left of the Mark area

                     dPart1 = LEFT$(dSwapLine, SwapSCol - 1)      ' Get part 1, the left of Swap string
                     cPart1 = LEFT$(cSwapLine, SwapSCol - 1)      ' Get part 1, the left of Swap string

                     dPart2 = MID$(dSwapLine, SwapECol + 1 TO i - 1) ' Get part 2, between the strings
                     cPart2 = MID$(cSwapLine, SwapECol + 1 TO i - 1) ' Get part 2, between the strings

                     dPart3 = MID$(dSwapLine, SwapSCol TO SwapECol) ' Get part 3, The swap string
                     cPart3 = MID$(cSwapLine, SwapSCol TO SwapECol) ' Get part 3, The swap string

                     dPart4 = MID$(dSwapLine, i)                  ' Get part 4, remainder of line
                     cPart4 = MID$(cSwapLine, i)                  ' Get part 4, remainder of line

                     dSwapLine = BUILD$(dPart1, dPart2, dPart3, dPart4) ' Rebuild the line
                     cSwapLine = BUILD$(cPart1, cPart2, cPart3, cPart4) ' Rebuild the line

                  ELSE                                            ' Mark is left of swap

                     '----- Swap area is Right of the Mark point

                     dPart1 = LEFT$(dSwapLine, i - 1)             ' Get part 1, the left of Mark string
                     cPart1 = LEFT$(cSwapLine, i - 1)             ' Get part 1, the left of Mark string

                     dPart2 = MID$(dSwapLine, SwapSCol TO SwapECol) ' Get part 2, the swap string
                     cPart2 = MID$(cSwapLine, SwapSCol TO SwapECol) ' Get part 2, the swap string

                     dPart3 = MID$(dSwapLine, i TO SwapSCol - 1)  ' Get part 3, Between mark and swap string
                     cPart3 = MID$(cSwapLine, i TO SwapSCol - 1)  ' Get part 3, Between mark and swap string

                     dPart4 = MID$(dSwapLine, SwapECol + 1)       ' Get part 4, After the Swap string
                     cPart4 = MID$(cSwapLine, SwapECol + 1)       ' Get part 4, After the Swap string

                     dSwapLine = BUILD$(dPart1, dPart2, dPart3, dPart4) ' Rebuild the line
                     cSwapLine = BUILD$(cPart1, cPart2, cPart3, cPart4) ' Rebuild the line

                  END IF                                          '

                  me.LTxtSet (swapLnum, dSwapLine)                ' Save the Swap line back
                  LAttrS(swapLnum) = cSwapLine                    ' Save the Swap line back
                  me.AttrScan(swapLnum)                           ' Recolorize

                  me.ModSet(swapLnum)                             ' Remember we changed something

               END IF                                             '

               swapLnum = me.CRPFwd(%mOData, swapLnum, 1)         ' Move to next line
               markLnum = me.CRPFwd(%mOData, markLnum, 1)         ' Move to next line

            NEXT SNum                                             ' Loop till all lines in block done

            me.MarkKill                                           ' Kill the mark area
            me.SwapKill                                           ' Kill the Swap area

            me.CurSetReq(%Position, j, i, %True)                  ' Set cursor set attempt
            DoSet(%Refresh)                                       ' Have it looked at

         ELSE                                                     ' Not even in data area
            GOSUB KillSwap                                        ' Kill things
         END IF                                                   '

      '----- Anything else, ignore
      ELSE                                                        ' Nothing going on, just exit
         MExitMeth                                                '

      END IF                                                      '
      MExitMeth

   KillSwap:
      IF IsSwapDrawn THEN me.SwapKill                             ' Kill underline if drawn
      OffSwapActive                                               ' Now kill Swap
      RETURN
   END METHOD

   METHOD  krTab()
   '---------- Tab forward To Next field
   LOCAL i, j, k AS LONG, lclTabs AS STRING
      MEntry
      CsrLinDX = 0                                                ' Reset possible hidden DX line location
      OnInsClnSupp                                                ' Suppress Insert cleanup for 1 Attn pass

      SELECT CASE CRow                                            ' Do tabbing based on where we are

         '----- Sitting on Line 1
         CASE 1                                                   ' Top Row
            IF CCol < 11 THEN SetCmd: MExitMeth                   '
            IF CCol < 21 + pCmdLen THEN SetScrl: MExitMeth        '
            me.CsrRow = 3 + Prf.Cols: me.CsrCol = 1: MExitMeth    '

         '----- Sitting on Line 2
         CASE 2, 2 + Prf.Cols                                     ' Line 2 or Ruler line
            me.CsrRow = 3 + Prf.Cols: me.CsrCol = 1: MExitMeth    '

         '----- Sitting in text area
         CASE ELSE                                                ' In data area
            j = sGetIX(CRow)                                      ' Get Line number of cursor line
            IF j = 0 THEN me.krHome: MexitMeth                    '
            IF j < 0 THEN j = sGetIX(CRow - ABS(j))               ' Get real data line
            lclTabs = me.TabsSimple(LEN(L(j).@LTxt))              ' Get a working tabs line
            k = 0                                                 ' Say no next line
            IF CRow < gwScrHeight THEN k = sGetIX(CRow + 1)       ' See if another screen line
            SELECT CASE CCol                                      ' Now base it on column
               CASE < gLNData1                                    ' In line number area
                  IF IsLTop(j) OR IsLBottom(j) THEN               ' Don't stop on TOP/BOTTOM lines
                     IF CRow < gwScrHeight THEN j = sGetIX(CRow + 1) ' See if another screen line
                     IF j THEN me.CsrRow = CRow + 1               '
                     IF j = 0 THEN SetCmd: MExitMeth              ' If not a real line, go to command line
                     me.CsrCol = 1: MExitMeth                     '
                  END IF                                          '
                  IF j = 0 THEN SetCmd: MExitMeth                 ' If not a real line, go to command line
                  IF ISNULL(TRIM$(RETAIN$(Prf.TabsLine, "*"))) OR ISFALSE Prf.Tabs THEN me.CsrCol = gLNData1: MExitMeth ' If no tabs, just tab over to data area
                  me.CsrCol = gLNPadCol                           ' Pretend we're in col gLNPadCol
                  GOSUB DataTab                                   ' Else try the dataline type
               CASE ELSE                                          ' In the real data area
                  GOSUB DataTab                                   '
               END SELECT                                         '
         END SELECT                                               '
      MExitMeth                                                   '

   '----- Handle the actual tabbing logic
   DataTab:
      IF ISFALSE Prf.Tabs THEN                                    ' If not tabbing
         me.CsrRowAdd(1): me.CsrCol = 1                           ' Do basic bump
         IF CRow > gwScrHeight OR k = 0 THEN SetCmd               ' Down too far? Cursor to the command line
         MExitMeth                                                ' Leave now
      END IF                                                      '

      '----- On Top of Data
      IF IsLTop(j) THEN                                           ' On the Top of Data Line
         me.CurSetReq(%Position, j + 1, 0, %True)                 ' Yes, Set cursor set attempt
         me.CsrCol = 1: me.CsrRowAdd(1)                           ' Col 1 of next line
         DoSet(%Refresh)                                          ' Have it looked at
         MExitMeth                                                '
      END IF                                                      '

      '----- On Bottom of Data
      IF IsLBottom(j) THEN                                        ' On the Bottom
         SetCmd                                                   ' Cursor to Cmd line
         MExitMeth                                                ' Leave now
      END IF                                                      '
                                                                  ' Tabbing IN a data line
      '----- Within a data line
      i = CCol - gLNPadCol + Offset                               ' Calc index into data string
      i = INSTR(i + 1, lclTabs, "*")                              ' Look for next tab marker

      '----- There's another marker on the line
      IF i THEN                                                   ' We have one
         IF i < Offset + 1 OR i > Offset + gDataLen - 1 THEN      ' Still on current screen
            me.CurSetReq(%Position, j, i, %True)                  ' No, Set cursor set attempt
            DoSet(%Refresh)                                       ' Have it looked at
            MExitMeth                                             '
         ELSE                                                     '
            me.CsrCol = i + gLNPadCol - Offset                    ' Just move cursor on the screen
            MExitMeth                                             '
         END IF                                                   '

      '----- No more tabs, we need to get to the next line
      ELSE                                                        ' We're gone past last tab, go to next line
         IF INSTR(Prf.TabsLine, "+") = 0 THEN                        ' Extensible tabs?
            j = me.CRPFwd(%MVisible, j, 1)                        ' No
            Offset = 0                                            '
            me.CurSetReq(%Position, j, 0, %True)                  ' Set cursor set attempt
            DoSet(%Refresh)                                       ' Have it looked at
         ELSE                                                     ' Yes, extensible tabs
            k = LEN(L(j).@LTxt)                                   ' Get current length of line
            DO UNTIL i > 0                                        ' OK, extend till we get a tab marker
               k += k                                             ' Double the length
               lclTabs = me.TabsSimple(k)                         ' Get a longer working tabs line
               i = CCol - gLNPadCol + Offset                      ' Calc index into data string
               i = INSTR(i + 1, lclTabs, "*")                     ' Look for next tab marker
            LOOP                                                  '
            IF i < Offset + 1 OR i > Offset + gDataLen - 1 THEN   ' Still on current screen
               me.CurSetReq(%Position, j, i, %True)               ' No, Set cursor set attempt
               DoSet(%Refresh)                                    ' Have it looked at
               MExitMeth                                          '
            ELSE                                                  '
               me.CsrCol = i + gLNPadCol - Offset                 ' Just move cursor on the screen
               MExitMeth                                          '
            END IF                                                '
         END IF                                                   '
      END IF                                                      '
      RETURN                                                      '
   END METHOD

   METHOD  krTDPaste(ln AS STRING)
   '---------- Paste in the passed string
   LOCAL ix, i, j AS LONG
   LOCAL t, tt AS STRING
      MEntry
      CsrLinDX = 0                                                ' Reset possible hidden DX line location
      me.MarkKill                                                 ' Kill any active block select
      me.SwapKill                                                 '
      me.AttrInvClearFind                                         ' Clear any Find HiLite

      '----- Paste into the Command line
      IF IsCCmnd THEN                                             ' In the Command area?
         tt = pCommand                                            ' Get working copy
         IF IsNotNsrtFlag THEN                                    ' Do the non-Insert mode handling
            IF LEN(tt) - (CCol - 10 + COffset) < LEN(ln) THEN tt = tt + SPACE$(LEN(ln)) ' Lengthen if needed
            MID$(tt, CCol - 10 + COffset, LEN(ln)) = ln           ' Overlay the CBD text
         ELSE                                                     ' Do the Insert method
            tt = LEFT$(tt, CCol - 11 + COffset) + ln + MID$(tt, CCol - 10 + COffset)
         END IF                                                   '
         PCommand = tt                                            ' Copy back
         i = CCol + COffset + LEN(ln)                             ' Move cursor
         IF i > pCmdLen + COffset THEN                            ' Is it to the right of the Window?
            COffset = i - (pCmdLen / 2)                           ' Adjust offset
         ELSEIF i < COffset THEN                                  ' Is it to the left of the Window?
            COffset = MAX(0, i - (pCmdLen / 2))                   ' Adjust offset
         END IF                                                   '
         me.WindowCmd                                             ' Print the whole string again
         me.CsrCol = i - COffset                                  ' Move cursor

      '----- Paste into the text data
      ELSEIF IsCData THEN                                         ' In the Data Area?
         j = S(CRow)                                              ' Get raw line index type
         ix = sGetIX(cRow)                                        ' Get IX of real data line

         '----- Make sure it's a valid data line
         IF ix < 0 THEN ix = sGetIX(CRow - ABS(ix))               ' Convert to real line index
         IF ix = 0 OR j < 0 THEN sDoBeep: MExitMeth               ' Bail out if below last line or a Hex Line
         IF ISTRUE (LFlagG(ix) AND %NonTypable) THEN              ' Ignore non-typable lines
            sDoBeep: MExitMeth                                    '
         END IF                                                   '

         '----- Do the paste
         me.ModSet(ix)                                            ' Remember we changed something
         i = CCol - gLNPadCol + Offset                            ' Calc index into string
         tt = LTxtG(ix)                                           ' Get working copy of text

         '----- Do it either Insert or non-Insert mode
         IF IsNotNsrtFlag THEN                                    ' Do the non-Insert mode handling
            IF LEN(tt) - i < LEN(ln) THEN tt = tt + SPACE$(i + LEN(ln) - LEN(tt))  ' Lengthen if needed
            MID$(tt, i, LEN(ln)) = ln                             ' Overlay the CBD text
         ELSE                                                     ' Do the Insert method
            tt = me.DataInsert(tt, ln, i)                         ' Insert the data
            me.LAttrAdjust(ix, i, LEN(ln))                        ' Adjust any Attr hilite
         END IF                                                   '

         IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) AND IsLData(ix) THEN ' CAPS ON?
            tt = UUCASE(tt)                                       ' UC if CAPS is ON
         END IF                                                   '
         me.LTxtSet(ix, tt)                                       ' Save it back
         me.AttrScan(ix)                                          ' Recolorize

         '----- Print, do extra stuff if Hex mode
         me.DispLine(ix, CRow)                                    ' Re-display the line
         IF Prf.HexMode <> &1 THEN                                   ' If hex mode
            IF CRow + 1 <= gwScrHeight THEN                       ' Make sure we don't go off the bottom
               sPrint (me.WindowHexTop(tt), $$TxtLo, CRow + 1, gLNData1)  ' Print upper hex
               me.MarkReDraw(CRow + 1)                            ' Re-draw the MARK lines
            END IF                                                '
            IF CRow + 2 <= gwScrHeight THEN                       ' Make sure we don't go off the bottom
               sPrint (me.WindowHexBottom(tt), $$TxtLo, CRow + 2, gLNData1) ' Print lower hex
            me.MarkReDraw(CRow + 2)                               ' Re-draw the MARK lines
            END IF                                                '
         END IF                                                   '

         me.CsrCol = CCol + LEN(ln)                               ' Move cursor
         IF CCol < Offset + 1 OR CCol > Offset + gDataLen - 1 THEN' Still on current screen
            me.CurSetReq(%Position, ix, CCol, %True)              ' No, Set cursor set attempt
         DoSet(%Refresh)                                          ' Have it looked at
         END IF                                                   '

      ELSE                                                        ' Anything else (CursBad)
         sDoBeep                                                  ' Beep it
      END IF                                                      '
      MExit                                                       '
   END METHOD

   METHOD  krTime()
   '---------- Paste Date at cursor location
   LOCAL ln AS STRING
      ln = sTime()                                                ' Go get the date
      me.krTDPaste(ln)                                            ' Go do common code
   END METHOD

   METHOD  krTitleCase()
   '---------- TitleCase a selected block
   LOCAL x, i AS LONG, lclTxt AS STRING
      me.AttrInvClearFind                                         ' Clear any Find HiLite
      '----- OK, now do the Titlecase
      FOR x = MarkSLin TO MarkELin                                ' Loop through the requested range
         IF ISTRUE IsLData(x) OR ISTRUE IsLNote(x) THEN           ' Just data lines
            me.ModSet(x)                                          ' Remember we changed something
            me.TitleCase(x, MarkSCol, MarkECol)                   ' Go process it
         END IF                                                   '
      NEXT x                                                      '
      DoSet(%Refresh)                                             ' Have it looked at
      me.MarkKill                                                 ' Reset the select frame
   END METHOD

   METHOD  krToggleHome()
   '---------- ToggleHome key
   LOCAL ix, i AS LONG
      '----- Make sure it's a valid line
      IF CRow < 3 + Prf.Cols THEN sDoBeep: EXIT METHOD            ' Command line, line 2 or COLS line
      ix = sGetIX(CRow)                                           ' Get IX of line the cursor is on
      IF ix = 0 OR ix = -3 THEN sDoBeep: EXIT METHOD              ' Below last line or Hex dash line? Bail out
      IF ISFALSE (LFlagG(ix) AND %NonTypable) THEN                ' See if cursor on a typable line

         '----- Toggle location back and forth
         IF CCol <> gLNData1 THEN                                 ' If not in col 1
            me.CsrCol = gLNData1                                  ' Move cursor back to 1st data col
         ELSE                                                     '
            IF ISNULL(TRIM$(LTxtG(ix))) THEN                      ' No text?
               me.CsrCol = gLNData1                               ' Stay in col 1
            ELSE                                                  '
               me.CsrCol = gLNPadCol + VERIFY(LTxtG(ix), " ")     '
            END IF                                                '
         END IF                                                   '
      ELSE                                                        '
         sDoBeep                                                  ' Beep
      END IF                                                      '
   END METHOD

   METHOD  krToggleSelect()
   '---------- Toggle Select status
      IF IsSlecSet THEN                                           ' Have a selection?
         IF IsSlecActive THEN                                     ' Toggle the Active status
            OffSlecActive                                         '
            IF IsMarkActive THEN me.MarkKill                      ' Something active, erase it
         ELSE                                                     '
            OnSlecActive                                          '
            IF ISFALSE IsMarkActive THEN                          ' Nothing active, flip it
               MarkSCol = SlecSCol                                ' Restore Marked area
               MarkECol = SlecECol                                '
               MarkSLin = me.LineNoRef(FORMAT$(SlecSLin))         '
               MarkELin = me.LineNoRef(FORMAT$(SlecELin))         '
               OnMarkActive                                       '
               me.MarkScr                                         ' Redraw screen
            END IF                                                '
         END IF                                                   '
         sDoStatusBar($SBSelect)                                  ' re-Do the StatusBar Select box
      END IF                                                      '
   END METHOD

   METHOD  krTxtHome()
   '---------- TxtHome key
   LOCAL ix, i AS LONG
      '----- Make sure the line's a valid one
      IF CRow < 3 + Prf.Cols THEN sDoBeep: EXIT METHOD            ' Command line, line 2 or COLS line
      ix = sGetIX(CRow)                                           ' Get IX of line the cursor is on
      IF ix = 0 OR ix = -3 THEN sDoBeep: EXIT METHOD              ' Below last line or Hex dash line? Bail out
      IF ISFALSE (LFlagG(ix) AND %NonTypable) THEN                ' See if cursor on a typable line
         me.CsrCol = gLNData1                                     ' Move cursor back to 1st data col
      ELSE                                                        '
         sDoBeep                                                  ' Beep
      END IF                                                      '
   END METHOD

   METHOD  krTxtNewLine(OPTIONAL ns AS LONG)
   '---------- TxtNewLine key
   LOCAL i, j, noscroll, ix AS LONG
      noscroll = IIF(ISMISSING(ns), %False, ns)                   ' Pick up No Scroll flag
      CsrLinDX = 0                                                ' Reset possible hidden DX line location

      '----- Top of screen area
      IF CRow < 3 + Prf.Cols THEN                                 ' Command line, line 2 or COLS line
         me.CsrRow = 3 + Prf.Cols: me.CsrCol = gLNData1: EXIT METHOD' Go to 1st data line
      END IF                                                      ' Else in data area

      '----- Get real line #
      ix = sGetIX(CRow)                                           ' Locate data line cursor was on
      IF ix = 0 THEN                                              ' Not in data area?
         SetCmd                                                   ' Cursor to Cmd line
         EXIT METHOD                                              '
      END IF                                                      '
      IF ix < 0 THEN ix = sGetIX(CRow - ABS(ix))                  ' If Hex extra line, get real line

      '----- Go to next line
      me.CsrRowAdd(1)                                             ' Simple INCR
      IF CRow > gwScrHeight OR IsLBottom(ix) THEN                 ' At bottom line yet?
         IF ix = 0 OR IsLBottom(ix) OR ISTRUE noscroll THEN       ' Not in data area or noscroll requested?
            SetCmd                                                ' Cursor to Cmd line
            EXIT METHOD                                           '
         END IF                                                   '
         ix = me.CRPFwd(%mVisible, ix, 1)                         ' Move to next visible line
         me.CurSetReq(%Position, ix, Offset + 1, %True)           ' Set cursor set attempt
         TopScrn = me.CRPFwd(%mVisible, TopScrn, 1)               ' Move TopScrn by scroll amount
         OnInsClnSupp                                             ' Suppress Insert cleanup for 1 Attn pass
         DoSet(%Refresh)                                          ' Have it looked at
      ELSE                                                        '
         me.CsrCol = gLNData1                                     '
      END IF                                                      '
   END METHOD

   METHOD  krUp()
   '---------- Decrement CsrRow
   LOCAL t AS STRING
   LOCAL ix, cIncr AS LONG
   LOCAL i, j AS LONG
      MEntry
      cIncr = IIF(VAL(gKeyPrimOper) > 0, VAL(gKeyPrimOper), 1)    ' Set the INCR amount?
      CsrLinDX = 0                                                ' Reset possible hidden DX line location
      IF TopScrn = 1 THEN                                         ' No scrolling possible
         me.CsrRow = MAX(1, CRow - cIncr)                         ' Simple DECR, go no lower than 1
         MExitMeth                                                '
      END IF                                                      '

      '----- Possibility of scrolling upward
      SELECT CASE CRow                                            ' OK, lets see what we got

         '----- Command line, always wrap
         CASE <= 2                                                ' Currently on line 1 or 2, go to Cmnd line
            me.CsrRow = 1: MExitMeth                              ' Stay on the command line

         CASE 2 + Prf.Cols                                        ' On the fixed COLS> line?
            me.CsrRow = MAX(1, CRow - cIncr)                      ' Simple DECR, go no lower than 1
            MExitMeth                                             '

         '----- Top text data line then do scrolling if allowed
         CASE (2 + Prf.Cols + cIncr)                              ' Time for a scroll?
            IF ISFALSE ENV.KBVerScroll THEN                       ' No KeybScroll then we're done
               me.CsrRow = MAX(1, CRow - cIncr)                   ' Simple DECR, go no lower than 1
               MExitMeth                                          '
            END IF                                                '

            ix = sGetIX(CRow)                                     ' Locate data line cursor was on
            Prf.ScrlPageSus = %True                               ' Set suspend PAGE mode
            IF ISFALSE IsLTop(ix) THEN                            ' See if cursor is at the top
               ix = me.CRPBack(%mVisible, ix, cIncr)              ' Move back to next visible line
               IF IsCData THEN                                    ' In data area?
                  me.CurSetReq(%Position, ix, CCol + Offset - gLNPadCol, %False) ' Set cursor set attempt
               ELSE                                               ' Line Number area then
                  me.CurSetReq(%Position, ix, 0, %False)          ' Set cursor set attempt
               END IF                                             '
               TopScrn = me.CRPBack(%mVisible, TopScrn, cIncr)    ' Move TopScrn by scroll amount
               OnInsClnSupp                                       ' Suppress Insert cleanup for 1 Attn pass
               DoSet(%Refresh)                                    ' Have it looked at

            '----- Anything else, just a simple cursor up
            ELSE                                                  '
               me.CsrRow = MAX(1, CRow - cIncr)                   ' Simple DECR, go no lower than 1
            END IF                                                '

         CASE ELSE
            me.CsrRow = MAX(1, CRow - cIncr)                      ' Simple DECR, go no lower than 1
      END SELECT                                                  '
      MExit
   END METHOD

   METHOD  krUpperCase()
   '---------- UpperCase a selected block
   LOCAL x, i AS LONG, lclTxt AS STRING
      me.AttrInvClearFind                                         ' Clear any Find HiLite
      '----- Ensure a marked area and get its bounds
      IF ISFALSE IsMarkActive THEN sDoBeep: EXIT METHOD           ' Better be something selected, else exit

      '----- OK, now do the Uppercase
      FOR x = MarkSLin TO MarkELin                                ' Loop through the requested range
         IF ISTRUE IsLData(x) OR ISTRUE IsLNote(x) THEN           ' Just data/note lines
            me.ModSet(x)                                          ' Remember we changed something
            lclTxt = LTxtG(x)                                     ' Get the text
            IF MarkSCol <= LEN(lclTxt) THEN                       ' Do only if long enough
               FOR i = MarkSCol TO MIN(MarkECol, LEN(lclTxt))     ' Just ones that are there
                  MID$(lclTxt, i, 1) = UUCASE(MID$(lclTxt, i, 1)) '
               NEXT i                                             '
               me.LTxtSet(x, lclTxt)                              ' Put work string back
               me.AttrScan(x)                                     ' Recolorize
               me.CurSetReq(%Position, MarkELin, MarkECol + 1, %True)  ' Set cursor set attempt
            END IF                                                '
         END IF                                                   '
      NEXT x                                                      '
      DoSet(%Refresh)                                             ' Have it looked at
      me.MarkKill                                                 ' Reset the select frame
   END METHOD

   METHOD  PTBackTab()
   '---------- Tab backward To previous field in PowerType mode
   LOCAL i, j AS LONG, lclTabs AS STRING
      MEntry
      OnInsClnSupp                                                ' Suppress Insert cleanup for 1 Attn pass

      '----- Do it based on where we are
      j = sGetIX(CRow)                                            ' Get data line index
      IF ISFALSE Prf.Tabs THEN _                                  ' No TABS, just go to left margin
         me.CsrCol = gLNData1: MExitMeth                          '
      i = CCol - gLNPadCol + Offset                               ' Calc index into data string
      lclTabs = me.TabsSimple(i + 80)                             ' Get a working tabs line

      '----- If far right, go to last tab
      IF i > LEN(lclTabs) THEN                                    ' If right of last tab
         GOSUB LastTab                                            ' Go set to the last tab

      '----- Else backup one tab
      ELSE                                                        '
         i = INSTR(i - LEN(lclTabs) - 2, lclTabs, "*")            ' look for prev tab marker
         IF i THEN                                                ' If we have one
            IF i < Offset + 1 OR i > Offset + gDataLen - 1 THEN   ' Still on current screen
               me.CurSetReq(%Position, j, i, %True)               ' No, Set cursor set attempt
               DoSet(%Refresh)                                    ' Have it looked at
               MExitMeth                                          '
            ELSE                                                  ' Still on screen, just move cursor
               me.CsrCol = i + gLNPadCol - Offset                 '
            END IF                                                '
         ELSE                                                     ' No more to the left
            me.CsrCol = gLNData1                                  ' Left margin if not already there
         END IF                                                   '
      END IF                                                      '

   MExitMeth                                                      '

   '----- Find the last tab
   LastTab:                                                       ' Find Last Tab
      lclTabs = me.TabsSimple(LEN(L(j).@LTxt))                    ' Get a working tabs line
      IF Prf.Tabs THEN                                            ' If TABS active
         i = INSTR(-1, lclTabs, "*")                              '
      ELSE                                                        '
         i = 1                                                    '
      END IF                                                      '
      IF i < Offset + 1 OR i > Offset + gDataLen - 1 THEN         ' Still on current screen
         me.CurSetReq(%Position, j, i, %True)                     ' Set cursor set attempt
         DoSet(%Refresh)                                          ' Have it looked at
      ELSE                                                        '
         me.CsrCol = i + gLNPadCol - Offset                       '
      END IF                                                      '
      RETURN                                                      '
   END METHOD

   METHOD  PTBS()
   '---------- Do a PC type Bksp function in PowerType mode
   LOCAL ix, i, j, k AS LONG
      '----- Do it if there's room
      IF CCol > gLNData1 THEN                                     ' Room to BKSP?

         '----- Repeat for each line in the PT group
         FOR k = 1 TO gPTblCount                                  ' Loop for each PT line
            j = gPTbl(k).sRow                                     ' Get screen row
            ix = gPTbl(k).tLin                                    ' Get IX of real data line
            IF CCol > gLNData1 AND IsLData(ix) THEN               ' Room to BKSP?
               me.ModSet(ix)                                      ' Remember we changed something
               i = CCol - gLNData1 + Offset                       ' Calc IX into the Text
               me.LTxtCharDel(ix, i)                              ' Delete 1 char
               me.LAttrAdjust(ix, i, -1)                          ' Adjust any Attr hilite
               IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) THEN ' CAPS ON and data line?
                  me.LTxtSet(ix, UUCASE(LTxtG(ix)))               ' UC it
               END IF                                             '
               me.AttrScan(ix)                                    ' Recolorize

               '----- If the line is visible on the screen, re-display it
               IF j > 0 THEN me.DispLine(ix, j)                   ' If visible Re-display the line
            END IF                                                '
         NEXT k                                                   '
         me.CsrColSub(1)                                          ' Adjust cursor
      ELSE                                                        '
         sDoBeep                                                  ' Beep if not possible
      END IF                                                      '
      TP.LastPTCurs = 0                                           ' Force cursor to draw
      sDoCursor                                                   ' To get back | | lines
   END METHOD

   METHOD  PTChar()
   '---------- Process normal character in PowerType Mode
   LOCAL i, j, k, ix AS LONG
      MEntry
      CsrLinDX = 0                                                ' Reset possible hidden DX line location

      '----- Do for each line in the PT Group
      FOR k = 1 TO gPTblCount                                     ' Loop for each PT line

         '----- Do the non-Insert mode style
         IF IsNotNsrtFlag THEN                                    ' Do the non-Insert mode handling
            j = gPTbl(k).sRow                                     ' Get screen row
            ix = gPTbl(k).tLin                                    ' Get IX of real data line
            IF ISFALSE IsLData(ix) THEN ITERATE FOR               ' Only Data lines
            me.ModSet(ix)                                         ' Remember we changed something
            i = CCol - gLNPadCol + Offset                         ' Calc index into string
            IF i > LEN(L(ix).@LTxt) THEN me.LTxtSet(ix, LSET$(LTxtG(ix), i)) ' Past the right end? Then lengthen string
            me.LTxtCharRep(ix, i, gKeyChr)                        ' Insert the character

            '----- Do the CAPS thing and re-display if line is on the screen
            IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) THEN              ' CAPS ON?
               me.LTxtSet(ix, UUCASE(LTxtG(ix)))                  ' UC if CAPS is ON
            END IF                                                '
            me.AttrScan(ix)                                       ' Recolorize
            IF j > 0 THEN me.DispLine(ix, j)                      ' If visible on the screen Re-display the line

         '----- Do the Insert mode style
         ELSE                                                     ' Do all the same stuff in Insert mode
            j = gPTbl(k).sRow                                     ' Get screen row
            ix = gPTbl(k).tLin                                    ' Get IX of real data line
            IF ISFALSE IsLData(ix) THEN ITERATE FOR               ' Only Data lines
            me.ModSet(ix)                                         ' Remember we changed something
            i = CCol - gLNPadCol + Offset                         ' Calc index into string
            me.LTxtSet(ix, me.DataInsert(LTxtG(ix), gKeyChr, i))  '
            me.LAttrAdjust(ix, i, 1)                              ' Adjust any Attr hilite

            '----- Do the CAPS thing and re-display if line is on the screen
            IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) THEN              ' CAPS ON?
               me.LTxtSet(ix, UUCASE(LTxtG(ix)))                  ' UC if CAPS is ON
            END IF                                                '
            me.AttrScan(ix)                                       ' Recolorize
            IF j > 0 THEN me.DispLine(ix, j)                      ' If visible on the screen Re-display the line

         END IF                                                   '
      NEXT k                                                      ' Next PT line
      me.krRight                                                  ' We can finally move over to the next column
      TP.LastPTCurs = 0                                           ' Force cursor to draw
      sDoCursor                                                   ' To get back | | lines
      MExit
   END METHOD

   METHOD  PTCopyPaste()
   '---------- Do Copy or Paste depending on whether selected in PT mode
      IF IsMarkActive AND IsMarkDrawn THEN                        ' Is something selected
         me.krKCpy                                                ' yes, COPY it
      ELSE                                                        '
         me.PTPaste                                               ' No? Then do a Paste
         TP.LastPTCurs = 0                                        ' Force cursor to draw
         sDoCursor                                                ' To get back | | lines
      END IF                                                      '
   END METHOD

   METHOD  PTCopyPasteAdd()
   '---------- Block copypaste in RAW APPEND mode
      gKeyPrimOper = "$RAA" + gKeyPrimOper                        ' Tell PTCopyPaste it's RAW mode
      me.PTCopyPaste()                                            ' Let IT do the work
   END METHOD

   METHOD  PTCopyPasteRaw()
   '---------- Block copypaste in RAW mode
      gKeyPrimOper = "$RAW" + gKeyPrimOper                        ' Tell PTCopyPaste it's RAW mode
      me.PTCopyPaste()                                            ' Let IT do the work
   END METHOD

   METHOD  PTCtlLeft()
   '---------- Tab Left one 'word' in PowerType mode
   LOCAL ix, i, j AS LONG
   LOCAL t AS STRING

      MEntry
      '----- Get the line we're on
      ix = gPTbl(1).tLin                                          ' Locate data line cursor was on
      t = RTRIM$(LTxtG(ix))                                       ' Get a copy of the text
      i = CCol + Offset - gLNPadCol                               ' Where is cursor now

      '----- Handle col 1, the simple one
      IF i = 1 THEN MExitMeth                                     ' Column 1? Do nothing

      '----- Past the end, go to last word
      IF i > LEN(RTRIM$(t)) THEN                                  ' At or past last char?
         GOSUB CurrLastWord                                       ' Go to Last Word

      '----- Backup one word
      ELSE                                                        ' Within line, find next location

         '----- Sitting on a blank
         IF MID$(t, i, 1) = " " THEN                              ' Currently sitting on a blank?
            j = INSTR(-(LEN(t) - i) - 2, t, ANY ENV.CharSet)      ' Look for prev visible character
            j = INSTR(-(LEN(t) - j)  - 2, t, " ") + 1             ' Look for prev blank
            IF j = 0 THEN                                         ' None?
               MExitMeth                                          ' Do nothing
            ELSE                                                  ' We've got the next 'word'
               IF j < Offset + 1 OR j > Offset + gDataLen - 1 THEN' Still on current screen
                  me.CurSetReq(%Position, ix, j, %True)           ' No, Set cursor set attempt
                  DoSet(%Refresh)                                 ' Have it looked at
               ELSE                                               '
                  me.CsrCol = j + gLNPadCol - Offset              ' Just move cursor on the screen
               END IF                                             '
            END IF                                                '

         '----- Sitting on an actual character
         ELSE                                                     ' We're sitting on a character

            '----- 1st character of a word
            IF i > 1 AND MID$(t, i - 1, 1) = " " THEN             ' Are we sitting on the 1st char of a word?
               j = INSTR(-(LEN(t) - i) - 3, t, ANY ENV.CharSet)   ' Look for prev visible character
               j = INSTR(-(LEN(t) - j) - 2, t, " ") + 1           ' Look for prev blank
               IF j = 0 THEN                                      ' None?
                  MExitMeth                                       ' Do nothing
               ELSE                                               ' We've got the next 'word'
                  IF j < Offset + 1 OR j > Offset + gDataLen - 1 THEN ' Still on current screen
                     me.CurSetReq(%Position, ix, j, %True)        ' No, Set cursor set attempt
                     DoSet(%Refresh)                              ' Have it looked at
                  ELSE                                            '
                     me.CsrCol = j + gLNPadCol - Offset           ' Just move cursor on the screen
                  END IF                                          '
               END IF                                             '

            '----- Non 1st char of word
            ELSE                                                  ' Sitting on other than 1st char of word
               j = INSTR(-(LEN(t) - i) - 2, t, " ") + 1           ' Look for prev blank + 1
               IF j = 0 THEN                                      ' None?
                  MExitMeth                                       ' Do Nothing
               ELSE                                               ' We've got the next 'word'
                  IF j < Offset + 1 OR j > Offset + gDataLen - 1 THEN ' Still on current screen
                     me.CurSetReq(%Position, ix, j, %True)        ' No, Set cursor set attempt
                     DoSet(%Refresh)                              ' Have it looked at
                  ELSE                                            '
                     me.CsrCol = j + gLNPadCol - Offset           ' Just move cursor on the screen
                  END IF                                          '
               END IF                                             '
            END IF                                                '
        END IF                                                    '
     END IF                                                       '
     MExitMeth                                                    ' We're done

   '-----Go to last word
   CurrLastWord:
      j = INSTR(-1, t, " ")                                       ' Look for last blank
      IF j <> 0 THEN INCR j                                       ' 1st char of last word
      j = MAX(1, j)                                               ' If not found, set to 1
      IF j < Offset + 1 OR j > Offset + gDataLen - 1 THEN         ' Still on current screen
         me.CurSetReq(%Position, ix, j, %True)                    ' No, Set cursor set attempt
         DoSet(%Refresh)                                          ' Have it looked at
      ELSE                                                        '
         me.CsrCol = j + gLNPadCol - Offset                       ' Just move cursor on the screen
      END IF                                                      '
      RETURN                                                      '
   END METHOD

   METHOD  PTCtlRight()
   '---------- Tab right one 'word' in PowerType mode
   LOCAL ix, i, j AS LONG
   LOCAL t AS STRING

      MEntry
      '----- Get line we're on
      ix = gPTbl(1).tLin                                          ' Locate data line cursor was on
      t = RTRIM$(LTxtG(ix)) + " "                                 ' Get a copy of the text (+ 1 blank)
      i = CCol + Offset - gLNPadCol                               ' Where is cursor now

      '----- If past end of line, do nothing
      IF i > LEN(RTRIM$(t)) THEN                                  ' At or past last char?
         MExitMeth                                                ' Do nothing

      '----- Somewhere within the line
      ELSE                                                        ' Within line, find next location

         '----- Sitting ON a character
         IF MID$(t, i, 1) <> " " THEN                             ' Currently sitting on a char?
            j = INSTR(i + 1, t, " ")                              ' Look for next blank
            j = INSTR(j + 1, t, ANY ENV.CharSet)                  ' Look for next visible character
            IF j = 0 THEN                                         ' None?
               MExitMeth                                          ' Do Nothing
            ELSE                                                  ' We've got the next 'word'
               IF j < Offset + 1 OR j > Offset + gDataLen - 1 THEN' Still on current screen
                  me.CurSetReq(%Position, ix, j, %True)           ' No, Set cursor set attempt
                  DoSet(%Refresh)                                 ' Have it looked at
               ELSE                                               '
                  me.CsrCol = j + gLNPadCol - Offset              ' Just move cursor on the screen
               END IF                                             '
            END IF                                                '

         '----- Sitting ON a blank
         ELSE                                                     ' We're sitting on a blank
            j = INSTR(i + 1, t, ANY ENV.CharSet)                  ' Look for next visible character
            IF j = 0 THEN                                         ' None?
               MExitMeth                                          ' Do nothing
            ELSE                                                  ' We've got the next 'word'
               IF j < Offset + 1 OR j > Offset + gDataLen - 1 THEN' Still on current screen
                  me.CurSetReq(%Position, ix, j, %True)           ' No, Set cursor set attempt
                  DoSet(%Refresh)                                 ' Have it looked at
               ELSE                                               '
                  me.CsrCol = j + gLNPadCol - Offset              ' Just move cursor on the screen
               END IF                                             '
            END IF                                                '
        END IF                                                    '
     END IF                                                       '
     MExit                                                        '
   END METHOD

   METHOD  PTCut()
   '---------- Cut selected columns in PT mode
   LOCAL x, i, j, k, ix AS LONG, lclTxt, CBD AS STRING
      '---------- OK, build the Clipboard string                  '
      CBD = ""                                                    ' Start as ""
      me.MarkKill                                                 ' Reset the select frame
      FOR k = 1 TO gPTblCount                                     ' Loop for each PT line
         j = gPTbl(k).sRow                                        ' Get screen row
         ix = gPTbl(k).tLin                                       ' Get IX of real data line
         IF ISFALSE IsLData(ix) THEN ITERATE FOR                  ' Only Data lines
         me.ModSet(ix)                                            ' Remember we changed something
         lclTxt = LTxtG(ix)                                       ' Get the text
         IF LEN(lclTxt) < MarkECol THEN lclTxt = LSET$(lclTxt, MarkECol)' Padd if needed
         lclTxt = MID$(lclTxt, MarkSCol TO MarkECol)              ' Extract the columns
         IF k = 1 THEN CBD += lclTxt + $CRLF                      ' Only the 1st line goes to the CBD
         lclTxt = LTxtG(ix)                                       ' Get the text again
         lclTxt = STRDELETE$(lclTxt, MarkSCol, MarkECol - MarkSCol + 1) ' Delete the characters
         me.LAttrAdjust(ix, MarkSCol, -(MarkECol - MarkSCol + 1)) ' Adjust any Attr hilite
         IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) THEN      ' CAPS ON?
            lclTxt = UUCASE(lclTxt)                               ' UC if CAPS is ON
         END IF                                                   '
         me.LTxtSet(ix, lclTxt)                                   ' Put work string back
         me.AttrScan(ix)                                          ' Recolorize
         IF j > 0 THEN me.DispLine(ix, j)                         ' If visible on the screen Re-display the line
      NEXT k                                                      '
      sWriteClipboard(CBD)                                        ' Write it
      TP.LastPTCurs = 0                                           ' Force cursor to draw
      sDoCursor                                                   ' To get back | | lines
   END METHOD

   METHOD  PTDataDelBS(which AS STRING)
   LOCAL ix, x, y, GotColumns, i, j, k AS LONG, lclTxt, lh AS STRING
      MEntry
      CsrLinDX = 0                                                ' Reset possible hidden DX line location
      IF which = "B" THEN me.CsrCol = CCol - 1                    ' Move cursor back if this is a backspace

      FOR k = 1 TO gPTblCount                                     ' Do for each PT line
         j = gPTbl(k).sRow                                        ' Get screen row
         ix = gPTbl(k).tLin                                       ' Get IX of real data line
         IF ISFALSE IsLData(ix) THEN ITERATE FOR                  ' Only Data lines
         me.ModSet(ix)                                            ' Remember we changed something
         i = CCol - gLNPadCol + Offset                            ' Calc IX into the Text
         IF which = "B" THEN i = IIF(i = 1, 1, i - 1)             ' Point at prev char if this is a backspace

         '----- See if any 'columns' in the data
         lclTxt = LTxtG(ix)                                       ' Get the text
         x = INSTR(i, lclTxt, " ")                                ' Look for at least 2 blanks following found string
         IF x THEN                                                ' If we 've got spare blanks
            FOR y = x TO LEN(lclTxt)                              ' Look through remainder of line for a non- blank
               IF MID$(lclTxt, y, 1) <> " " THEN                  ' Got one?
                  GotColumns = %True                              ' Remember we have columns
                  EXIT FOR                                        ' Exit, y - 1 is split point for lh/rh
               END IF                                             '
            NEXT y                                                '
         END IF                                                   '

         '----- There are columns do it the hard way
         IF GotColumns THEN                                       ' Data Insert AND columns?
            lh = LEFT$(lclTxt, y - 2)                             ' Create full lh portion
            lh = STRDELETE$(lh, i, 1) + " "                       ' Delete char and add a blank
            lclTxt = lh + MID$(lclTxt, y - 1)                     ' Re-build the altered text
            me.LTxtSet(ix, lclTxt)                                ' Save it back
         ELSE                                                     '
            me.LTxtCharDel(ix, i)                                 ' Delete 1 char
            me.LAttrAdjust(ix, i, -1)                             ' Adjust any Attr hilite
         END IF                                                   '

         IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) THEN      ' CAPS ON?
            me.LTxtSet(ix, UUCASE(LTxtG(ix)))                     ' UC if CAPS is ON
         END IF                                                   '
         me.AttrScan(ix)                                          ' Recolorize
         IF j > 0 THEN me.DispLine(ix, j)                         ' If visible on the screen Re-display the line
      NEXT k                                                      '
      TP.LastPTCurs = 0                                           ' Force cursor to draw
      sDoCursor                                                   ' To get back | | lines
      MExit                                                       '
   END METHOD

   METHOD  PTDate()
   '---------- Paste Date at cursor location in PowerType mode
   LOCAL ln AS STRING
      ln = sDate()                                                ' Go get the date
      me.PTTDPaste(ln)                                            ' Go to common routine
   END METHOD

   METHOD  PTDel()
   '---------- Delete character at current position in PowerType Mode
   LOCAL ix, x, i, j, k AS LONG, lclTxt AS STRING
      MEntry
      CsrLinDX = 0                                                ' Reset possible hidden DX line location
      IF IsMarkActive THEN                                        ' Is something selected?
         me.MarkKill                                              ' Reset the select frame

         '----- OK, now do the Delete based on the marked column range
         FOR k = 1 TO gPTblCount                                  ' Do for each PT line
            j = gPTbl(k).sRow                                     ' Get screen row
            ix = gPTbl(k).tLin                                    ' Get IX of real data line
            IF ISFALSE IsLData(ix) THEN ITERATE FOR               ' Only Data lines
            me.ModSet(ix)                                         ' Remember we changed something
            lclTxt = LTxtG(ix)                                    ' Get the text
            lclTxt = STRDELETE$(lclTxt, MarkSCol, MarkECol - MarkSCol + 1) ' Delete the characters
            me.LAttrAdjust(ix, MarkSCol, -(MarkECol - MarkSCol + 1))' Adjust any Attr hilite
            IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) THEN   ' CAPS ON?
               lclTxt = UUCASE(lclTxt)                            ' UC if CAPS is ON
            END IF                                                '
            me.LTxtSet(ix, lclTxt)                                ' Put work string back
            me.AttrScan(ix)                                       ' Recolorize
            IF j > 0 THEN me.DispLine(ix, j)                      ' If visible on the screen Re-display the line
         NEXT k                                                   '
         me.CsrCol = MarkSCol - Offset + gLNPadCol                ' Position cursor
         TP.LastPTCurs = 0                                        ' Force cursor to draw
         sDoCursor                                                ' To get back | | lines
         MExitMeth                                                '
      END IF                                                      '

      FOR k = 1 TO gPTblCount                                     ' Do for each PT line
         j = gPTbl(k).sRow                                        ' Get screen row
         ix = gPTbl(k).tLin                                       ' Get IX of real data line
         IF ISFALSE IsLData(ix) THEN ITERATE FOR                  ' Only Data lines
         me.ModSet(ix)                                            ' Remember we changed something
         i = CCol - gLNPadCol + Offset                            ' Calc IX into the Text
         me.LTxtCharDel(ix, i)                                    ' Delete 1 char
         me.LAttrAdjust(ix, i, -1)                                ' Adjust any Attr hilite
         IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) THEN                 ' CAPS ON?
            me.LTxtSet(ix, UUCASE(LTxtG(ix)))                     ' UC if CAPS is ON
         END IF                                                   '
         me.AttrScan(ix)                                          ' Recolorize
         IF j > 0 THEN me.DispLine(ix, j)                         ' If visible on the screen Re-display the line
      NEXT k                                                      '
      TP.LastPTCurs = 0                                           ' Force cursor to draw
      sDoCursor                                                   ' To get back | | lines
      MExit                                                       '
   END METHOD

   METHOD  PTDup()
   '---------- Duplicate prior line into the marked block
   LOCAL i, j, k, ix AS LONG, lclTxt, t AS STRING
      '----- Get the data from the previous line
      i = me.CrpBack(%mData, MarkSLin, 1)                         ' Get address of prev data line
      IF ISFALSE IsLData(i) THEN sDoBeep: EXIT METHOD             ' Beep and exit if no prior
      t = LTxtG(i)                                                ' Get text from the previous line
      IF LEN(t) < MarkECol THEN t = LSET$(t, MarkECol)            ' If line shorter than select columns, lengthen it
      t = MID$(t, MarkSCol TO MarkECol)                           ' Get the string to be Dup'd

      '----- Loop now through the PT lines
      FOR k = 1 TO gPTblCount                                     ' Loop for each PT line
         j = gPTbl(k).sRow                                        ' Get screen row
         ix = gPTbl(k).tLin                                       ' Get IX of real data line
         IF ISFALSE IsLData(ix) THEN ITERATE FOR                  ' Only Data lines
         me.ModSet(ix)                                            ' Remember we changed something
         lclTxt = LTxtG(ix)                                       ' Get the text
         IF LEN(lclTxt) < MarkECol THEN lclTxt = LSET$(lclTxt, MarkECol)' Padd if needed
         MID$(lclTxt, MarkSCol TO MarkECol) = t                   ' Extract the columns
         me.LTxtSet(ix, lclTxt)                                   ' Save it back
         me.AttrScan(ix)                                          ' Recolorize
         IF j > 0 THEN me.DispLine(ix, j)                         ' If visible on the screen Re-display the line
      NEXT k                                                      '

      '----- We're done, clean up and exit
      me.MarkKill                                                 ' Reset the select frame
      me.CurSetReq(%LineCmd, MarkSLin, MarkSCol, %True)           ' Set cursor set attempt
      DoSet(%Refresh)                                             ' Have it looked at
   END METHOD

   METHOD  PTEnter
   '---------- Enter in PT mode
      '----- Kludge to get LEFT / RIGHT through in PType Mode
      IF IsEQ(TRIM$(pCommand), "LEFT") THEN                       ' Sneak LEFT through
         pCommand = ""                                            ' Blow the command away
         me.PTLeftScroll                                          ' Simulate a left Scroll
      ELSEIF IsEQ(TRIM$(pCommand), "RIGHT") THEN                  ' Sneak RIGHT through
         pCommand = ""                                            ' Blow the command away
         me.PTRightScroll                                         ' Simulate a right Scroll
      ELSE                                                        ' Else normal ENTER
         OffPTypeMode                                             ' Turn off PowerType
         gPTblCount = 0                                           '
         ErrMsg = ""                                              '
         me.krEnter                                               '
         DoSet(%Refresh)                                          '
      END IF                                                      '
   END METHOD

   METHOD  PTEnum(which AS STRING)
   '---------- Do the Enum function
   LOCAL ix, i, j, k, enumz, ln AS LONG
   LOCAL menum AS DWORD
   LOCAL t, tt, mprefix, msuffix, digits, nv AS STRING
      MEntry
      '----- Do some basic setup
      CsrLinDX = 0                                                ' Reset possible hidden DX line location
      digits = $Numeric: IF which <> "D" THEN digits = $Hex + "abcdef" ' Set the valid 'digits' string
      ix = gPTbl(1).tLin                                          ' Get IX of 1st PT line
      me.MarkKill                                                 ' Reset the select frame
      t = LTxtG(ix)                                               ' Get the text line
      IF LEN(t) < MarkECol THEN t = LSET$(t, MarkECol)            ' Padd if needed
      t = MID$(t, MarkSCol TO MarkECol)                           ' Extract the columns containg the model

      '----- See if there's a suffix
      IF INSTR(digits, RIGHT$(t, 1)) > 0 THEN                     ' If it ends in a digit,
         msuffix = ""                                             ' then there is NO suffix
      ELSE                                                        ' Non-digit, we must extract the suffix
         i = INSTR(-1, t, ANY digits)                             ' Look for the last digit
         IF i = 0 THEN                                            ' None? whoops
            tt = "Enum model has no starting value"               '
            GOSUB DispMsg                                         '
            MExitMeth                                             '
         END IF                                                   '
         msuffix = MID$(t, i + 1)                                 ' Found digit, extract suffix
         t = LEFT$(t, i)                                          ' Strip it off the model
      END IF                                                      '

      '----- Now extract the prefix
      j = 0                                                       ' Reset j
      FOR i = LEN(t) TO 1 STEP -1                                 ' Look backward for 1st non-digit
         IF INSTR(digits, MID$(t, i, 1)) > 0 THEN ITERATE FOR     ' If a digit, keep looking
         j = i: EXIT FOR                                          ' Found a non digit, save its loc and exit for
      NEXT i                                                      '
      IF j = 0 THEN                                               ' All were digits
         mprefix = ""                                             ' Therefor no prefix
      ELSE                                                        ' Non-digit, we have a prefix
         mprefix = LEFT$(t, j)                                    ' Save prefix
         t = MID$(t, j + 1)                                       ' Extract it leaving t as the eval string.
      END IF                                                      '

      '----- Extract the enum value
      ln = LEN(t)                                                 ' Save length for rebuilding
      IF which = "D" THEN                                         ' Decimal Enum?
         menum = VAL(t)                                           ' Save it
         IF LEN(t) <> LEN(FORMAT$(menum)) THEN enumz = %True      ' Entered as leading zeros?, remember that
      ELSE                                                        ' One of the Hex
         menum = VAL("&H" + t)                                    ' Save it that way
         IF LEN(t) <> LEN(HEX$(menum)) THEN enumz = %True         ' Entered as leading zeros?, remember that
      END IF                                                      '

      '----- Now modify the PT lines
      FOR k = 1 TO gPTblCount                                     ' Loop for each PT line
         j = gPTbl(k).sRow                                        ' Get screen row
         ix = gPTbl(k).tLin                                       ' Get IX of real data line
         IF ISFALSE IsLData(ix) THEN ITERATE FOR                  ' Only Data lines
         me.ModSet(ix)                                            ' Remember we changed something
         tt = LTxtG(ix)                                           ' Get working copy of text
         IF LEN(tt) < MarkECol THEN tt = LSET$(tt, MarkECol)      ' Padd if needed

         '----- Select which Enum type to do
         SELECT CASE AS CONST$ which                              ' Build output string
            CASE "D"                                              ' Decimal
               t = FORMAT$(menum)                                 '
               GOSUB AdjustPad                                    ' Do padding etc.
            CASE "L"                                              ' Lowercase Hex
               t = LLCASE(HEX$(menum))                            '
               GOSUB AdjustPad                                    ' Do padding etc.
            CASE "U"                                              ' Lowercase Hex
               t = HEX$(menum)                                    '
               GOSUB AdjustPad                                    ' Do padding etc.
         END SELECT                                               '

         '----- Build the value and stuff it in
         t = mprefix + t + msuffix                                ' Add prefix and suffix
         t = RIGHT$(RSET$(t, MarkECol - MarkSCol + 1), MarkECol - MarkScol + 1) ' Right align and trim it
         MID$(tt, MarkSCol, LEN(t)) = t                           ' Overlay the enum text
         me.LTxtSet(ix, tt)                                       ' Save it back
         me.AttrScan(ix)                                          ' Recolorize

         '----- If on the screen, display it
         IF j > 0 THEN me.DispLine(ix, j)                         ' If visible on the screen Re-display the line
         menum += gEnumWith                                       ' Incr enum for next line
      NEXT k                                                      '
      TP.LastPTCurs = 0                                           ' Force cursor to draw
      sDoCursor                                                   ' To get back | | lines
      MExitMeth

   '----- Fiddle with the prefix padding
   AdjustPad:
      IF LEN(t) > ln AND RIGHT$(mprefix, 1) = " " THEN            ' If enum needs more room, and Prefix has some
         mprefix = LEFT$(mprefix, LEN(mprefix) - 1)               ' Steal right hand blank from prefix
         INCR ln                                                  ' Make enum longer
      ELSEIF LEN(t) > ln AND LEFT$(mprefix, 1) = " " THEN         ' If not right end, try left end
         mprefix = MID$(mprefix, 2)                               ' Steal left hand blank from prefix
         INCR ln                                                  ' Make enum longer
      END IF                                                      '
      IF LEN(t) > ln AND RIGHT$(msuffix, 1) = " " THEN            ' If enum needs more room, and Suffix has some
         msuffix = LEFT$(msuffix, LEN(msuffix) - 1)               ' Steal right hand blank from suffix
         INCR ln                                                  ' Make enum longer
      ELSEIF LEN(t) > ln AND LEFT$(msuffix, 1) = " " THEN         ' If not right end, try left end
         msuffix = MID$(msuffix, 2)                               ' Steal left hand blank from suffix
         INCR ln                                                  ' Make enum longer
      END IF                                                      '
      IF enumz THEN                                               ' Pad appropriately
         t = RIGHT$(REPEAT$(ln, "0") + t, ln)                     '
      ELSE                                                        '
         t = RIGHT$(REPEAT$(ln, " ") + t, ln)                     '
      END IF                                                      '
      RETURN

   '----- Temporarily display an error message
   DispMsg:
      t = STRING$(ENV.ScrWidth - LEN(tt), "_")                    ' Build LH part of dash line
      sPrint (t, $$TxtLo, 2, 1)                                   ' Print LH part of line 2
      sPrint (tt, $$Error, 2, ENV.ScrWidth - LEN(tt) + 1)         ' Print rest of line
      GRAPHIC REDRAW                                              '
      SLEEP 1500                                                  '
      tt = "Entering PowerType mode, Press Enter to exit"         '
      sPrint (t, $$TxtLo, 2, 1)                                   ' Print LH part of line 2
      sPrint (tt, $$Error, 2, ENV.ScrWidth - LEN(tt) + 1)         ' Print rest of line
      GRAPHIC REDRAW                                              '
      RETURN                                                      '
   END METHOD

   METHOD  PTEOL()
   '---------- EOL key in PowerType mode
   LOCAL ix, i, j, k AS LONG
      '----- Ripple through the PT lines doing it
      FOR k = 1 TO gPTblCount                                     ' Loop for each PT line
         j = gPTbl(k).sRow                                        ' Get screen row
         ix = gPTbl(k).tLin                                       ' Get IX of real data line
         IF ISFALSE IsLData(ix) THEN ITERATE FOR                  ' Only Data lines
         me.ModSet(ix)                                            ' Remember we changed something
         i = CCol - gLNPadCol + Offset                            ' Calc index into string
         me.LTxtSet(ix, LEFT$(LTxtG(ix), i - 1))                  ' Truncate the line then
         LAttrS(ix) = LEFT$(LAttrG(ix), i - 1)                    ' Truncate the Attr
         IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) THEN      ' CAPS ON?
            me.LTxtSet(ix, UUCASE(LTxtG(ix)))                     ' UC if CAPS is ON
         END IF                                                   '
         me.AttrScan(ix)                                          ' Recolorize
         IF j > 0 THEN me.DispLine(ix, j)                         ' If visible on the screen Re-display the line
      NEXT k                                                      '
      TP.LastPTCurs = 0                                           ' Force cursor to draw
      sDoCursor                                                   ' To get back | | lines
   END METHOD

   METHOD  PTErase()
   '---------- Erase selected columns from the PT range
   LOCAL x, i, j, k, ix AS LONG, lclTxt, CBD AS STRING
      IF ISFALSE IsMarkActive THEN sDoBeep: EXIT METHOD           ' Better be something selected

      '----- OK, now do the erase
      me.MarkKill                                                 ' Reset the select frame
      FOR k = 1 TO gPTblCount                                     ' Loop for each PT line
         j = gPTbl(k).sRow                                        ' Get screen row
         ix = gPTbl(k).tLin                                       ' Get IX of real data line
         IF ISFALSE IsLData(ix) THEN ITERATE FOR                  ' Only Data lines
         me.ModSet(ix)                                            ' Remember we changed something
         lclTxt = LTxtG(ix)                                       ' Get the text
         IF LEN(lclTxt) < MarkECol THEN lclTxt = LSET$(lclTxt, MarkECol)' Padd if needed
         MID$(lclTxt, MarkSCol TO MarkECol) = SPACE$(MarkECol - MarkSCol + 1) ' Blank it
         me.LTxtSet(ix, lclTxt)                                   ' Put work string back
         me.AttrScan(ix)                                          ' Recolorize
         IF j > 0 THEN me.DispLine(ix, j)                         ' If visible on the screen Re-display the line
      NEXT k                                                      '
      TP.LastPTCurs = 0                                           ' Force cursor to draw
      sDoCursor                                                   ' To get back | | lines
   END METHOD

   METHOD  PTISODate()
   '---------- Paste ISO Date at cursor location in PowerType mode
   LOCAL ln AS STRING
      ln = DATE$                                                  ' Go get the date
      ln = MID$(ln, 7) + "-" + LEFT$(ln, 2) + "-" + MID$(ln, 4, 2)' Reformat it to ISO standard
      me.PTTDPaste(ln)                                            ' Go to common routine
   END METHOD

   METHOD  PTISOTime()
   '---------- Paste ISO Time at cursor location in PowerType mode
   LOCAL ln AS STRING
      ln = TIME$                                                  ' Go get the time
      me.PTTDPaste(ln)                                            ' Go to common routine
   END METHOD

   METHOD  PTJustify()
   '---------- Justify a selected block in PowerType mode
   LOCAL x, i, j, k, ix AS LONG, lclTxt, jtxt, ttxt AS STRING
      me.MarkKill                                                 ' Reset the select frame
      '----- OK, now do the justify
      FOR k = 1 TO gPTblCount                                     ' Loop for each PT line
         j = gPTbl(k).sRow                                        ' Get screen row
         ix = gPTbl(k).tLin                                       ' Get IX of real data line
         lclTxt = LTxtG(ix)                                       ' Get the text
         me.ModSet(ix)                                            ' Remember we changed something
         IF LEN(lclTxt) < MarkECol THEN _                         ' Lengthen text if needed
            lclTxt = LSET$(lclTxt, MarkECol)                      '
         jtxt = MID$(lclTxt, MarkSCol TO MarkECol)                ' Get the string to be justified

         '----- Select which Justify is needed
         SELECT CASE AS CONST$ gKeyChr                            ' Do the right type of justify
            CASE "JUSTIFYL": jtxt = LSET$(LTRIM$(jtxt), LEN(jtxt))'
            CASE "JUSTIFYR": jtxt = RSET$(RTRIM$(jtxt), LEN(jtxt))'
            CASE "JUSTIFYC": ttxt = SPACE$(LEN(jtxt)): jtxt = TRIM$(jtxt)
                             MID$(ttxt, (LEN(ttxt) - LEN(jtxt)) \ 2 + 1, LEN(jtxt)) = jtxt
                             jtxt = ttxt                          '
         END SELECT                                               '
         MID$(lclTxt, MarkSCol TO MarkECol) = jtxt                ' Re-insert the string
         me.LTxtSet(ix, lclTxt)                                   ' Put work string back
         me.AttrScan(ix)                                          ' Recolorize
         IF j > 0 THEN me.DispLine(ix, j)                         ' If visible on the screen Re-display the line
      NEXT k                                                      '
      TP.LastPTCurs = 0                                           ' Force cursor to draw
      sDoCursor                                                   ' To get back | | lines
   END METHOD

   METHOD  PTLastTab()
   '---------- Tab To the Last Tab marker
   LOCAL i, j, k AS LONG, lclTabs AS STRING
      j = gPTbl(1).tLin                                           ' Get Line number of cursor line
      lclTabs = me.TabsSimple(LEN(L(j).@LTxt))                    ' Get a working tabs line
      i = INSTR(Prf.TabsLine, "+")                                   ' + type tabs

      '----- If extensible tabs, just use last *
      IF ISFALSE i THEN                                           ' Can't do LastTab with extensible tabs
         k = 0                                                    ' Say no next ine
         IF ISFALSE Prf.Tabs THEN EXIT METHOD                     ' If not tabbing
         i = INSTR(-1, lclTabs, "*")                              ' look for last tab marker
      END IF                                                      '

      '----- If no + (or adjusted to last *)
      IF i THEN                                                   ' we have one
         IF i < Offset + 1 OR i > Offset + gDataLen - 1 THEN      ' Still on current screen
            me.CurSetReq(%Position, j, i, %True)                  ' No, Set cursor set attempt
            DoSet(%Refresh)                                       ' Have it looked at
            MExitMeth                                             '
         ELSE                                                     '
            me.CsrCol = i + gLNPadCol - Offset                    ' Just move cursor on the screen
            MExitMeth                                             '
         END IF                                                   '
      ELSE                                                        ' ?? No tabs?
         me.CsrCol = gLNData1 - Offset                            ' Just move cursor on the screen
      END IF                                                      '
   END METHOD

   METHOD  PTLeft()
   '---------- Decrement CsrCol in PowerType mode
   LOCAL ix AS LONG
   LOCAL t AS STRING
      me.CsrColSub(1)                                             ' Simple DECR
      IF CCol = gLNPadCol AND Offset > 0 THEN                     ' At left side of screen and room to scroll?
         me.CsrColAdd(1)                                          ' Yes, Undo the DECR for now
         ix = gPTbl(1).tLin                                       ' Locate data line cursor was on
         me.CurSetReq(%Position, ix, CCol + Offset - gLNPadCol - 1, %False) ' Set cursor set attempt
         me.OffsetSub(1)                                          ' Drop the offset
         DoSet(%Refresh)                                          ' Have it looked at
      ELSEIF CCol = gLNPadCol THEN                                ' At left end and no scroll possible
         me.CsrColAdd(1)                                          ' Put things back
         sDoBeep                                                  '
      END IF                                                      '
   END METHOD

   METHOD  PTLift()
   '---------- Copy and Erase selected columns from the PT range
   LOCAL x, i, j, k, ix AS LONG, lclTxt, ltxt2, CBD AS STRING

      MEntry
      '----- If nothing highlighted, null the clipboard
      IF ISFALSE IsMarkActive THEN                                ' Better be something selected
         CBD = ""                                                 ' No? Make a NULL string
         sWriteClipboard(CBD)                                     ' Write it
         MExitMeth                                                '
      END IF                                                      '

      CBD = ""                                                    ' Start as null

      '----- OK, now do the erase
      FOR k = 1 TO gPTblCount                                     ' Loop for each PT line
         j = gPTbl(k).sRow                                        ' Get screen row
         ix = gPTbl(k).tLin                                       ' Get IX of real data line
         IF ISFALSE IsLData(ix) THEN ITERATE FOR                  ' Only Data lines
         me.ModSet(ix)                                            ' Remember we changed something
         lclTxt = LTxtG(ix)                                       ' Get the text
         IF LEN(lclTxt) < MarkECol THEN lclTxt = LSET$(lclTxt, MarkECol)' Padd if needed
         lTxt2 = MID$(lclTxt, MarkSCol TO MarkECol)               ' Extract the columns
         CBD += lTxt2 + $CR + $LF                                 ' Add each line with CR/LF
         MID$(lclTxt, MarkSCol TO MarkECol) = SPACE$(MarkECol - MarkSCol + 1) ' Blank it
         me.LTxtSet(ix, lclTxt)                                   ' Put work string back
         me.AttrScan(ix)                                          ' Recolorize
         IF j > 0 THEN me.DispLine(ix, j)                         ' If visible on the screen Re-display the line
      NEXT k                                                      '

      '----- Stuff it in the clibboard
      IF sWriteClipboard(CBD) THEN _                              ' Write it
         me.MarkKill                                              ' Reset the select frame
      TP.LastPTCurs = 0                                           ' Force cursor to draw
      sDoCursor                                                   ' To get back | | lines
      MExit
   END METHOD

   METHOD  PTLEFTScroll()
   '---------- Scroll LEFT in PowerType mode
   LOCAL ScrAmt AS LONG, ScrCmd AS STRING
      IF Offset = 0 THEN sDoBeep: EXIT METHOD                     ' Already at the left?  Bail oput

      '----- Get the amount to scroll
      ScrCmd = Prf.Scroll                                         ' Pick up default scroll amount
      IF VAL(ScrCmd) <> 0 THEN                                    ' A Number?
         ScrAmt = VAL(ScrCmd)                                     ' Yes, use it
      ELSEIF ScrCmd = "HALF" THEN                                 ' Half?
         ScrAmt = INT(gDataLen / 2)                               ' Yes, set it
      ELSEIF ScrCmd = "PAGE" THEN                                 ' Page?
         ScrAmt = gDataLen                                        ' Yes, set it
      ELSEIF ScrCmd = "FULL" THEN                                 ' Full?
         ScrAmt = gDataLen                                        ' Yes, set it
      ELSEIF ScrCmd = "DATA" THEN                                 ' Data?
         ScrAmt = gDataLen - 1                                    ' Yes, set it
      ELSEIF IsCData THEN                                         ' No, In data area?
         ScrAmt = CCol - gLNData1                                 ' Yes, move cursor to left of screen
         IF ScrAmt = 0 THEN ScrAmt = gDataLen - 1                 ' If already there go a default amount
      END IF                                                      '

      '----- Now adjust things
      Offset = MAX(Offset - ScrAmt, 0)                            '
      me.CurSetReq(%Position, gPTbl(1).tLin, CCol - gLNPadCol + Offset + ScrAmt, %False)
      DoSet(%Refresh)                                             ' Have it looked at
   END METHOD

   METHOD  PTLowerCase()
   '---------- LowerCase a selected block in PowerType mode
   LOCAL x, i, j, k, ix AS LONG, lclTxt AS STRING
      IF ISFALSE IsMarkActive THEN sDoBeep: EXIT METHOD           ' Better be something selected, else exit
      me.MarkKill                                                 ' Reset the select frame
      '----- OK, now do the Lowercase
      FOR k = 1 TO gPTblCount                                     ' Loop for each PT line
         j = gPTbl(k).sRow                                        ' Get screen row
         ix = gPTbl(k).tLin                                       ' Get IX of real data line
         lclTxt = LTxtG(ix)                                       ' Get the text
         me.ModSet(ix)                                            ' Remember we changed something
         IF MarkSCol <= LEN(lclTxt) THEN                          ' Do only if long enough
            FOR i = MarkSCol TO MIN(MarkECol, LEN(lclTxt))        ' Just ones that are there
               MID$(lclTxt, i, 1) = LLCASE(MID$(lclTxt, i, 1))    '
            NEXT i                                                '
            me.LTxtSet(ix, lclTxt)                                ' Put work string back
            me.AttrScan(ix)                                       ' Recolorize
            IF j > 0 THEN me.DispLine(ix, j)                      ' If visible on the screen Re-display the line
         END IF                                                   '
      NEXT k                                                      '
      TP.LastPTCurs = 0                                           ' Force cursor to draw
      sDoCursor                                                   ' To get back | | lines
   END METHOD

   METHOD  PTMark(Dir AS STRING, Pass AS DWORD)
   '---------- Mark a block
   LOCAL i, Lin, sCol, OldLin, NewLin AS LONG
   LOCAL t AS STRING
   LOCAL sl, el, ll AS LONG
      MEntry
      CsrLinDX = 0                                                ' Reset possible hidden DX line location
      IF ISFALSE IsCData OR Prf.HexMode = &4  THEN CALL DWORD Pass: MExitMeth ' Not in the Data area or HEX? Treat as non-shifted key

      '----- Only handle valid ones in PT mode
      IF Dir <> "L" AND Dir <> "R" AND Dir <> "E" THEN MExitMeth  ' Ignore non L/R/E
      Lin = sGetIX(CRow)                                          ' Get IX of real data line
      sCol = CCol - gLNPadCol + Offset                            ' Calc index into string

      '----- Better be a marked area
      IF IsMarkActive THEN                                        ' If we already have a Marked area
         IF Lin < MarkSLin OR Lin > MarkELin OR _                 ' Somehow outside it's boundaries?
            sCol < MarkSCol OR sCol > MarkECol THEN               '
            me.MarkKill                                           ' Kill it and leave cursor alone
            MExitMeth                                             '
         END IF                                                   '
         IF Lin > MarkSLin AND Lin < MarkELin AND _               ' Somehow within the bounds?
            sCol > MarkSCol AND sCol < MarkECol THEN _            '
            CALL DWORD Pass: MExitMeth                            ' Just move the cursor

         '----- Move boundary based on direction asked for
         SELECT CASE dir                                          ' Adjust boundaries
            CASE "L"                                              ' Left
               IF sCol = MarkSCol AND MarkSCol > 1 THEN           ' At left and got room?
                  MarkSCol = MarkSCol - 1                         ' Move left boundary
                  me.MarkScr                                      ' Redraw the hilight
                  CALL DWORD Pass: MExitMeth                      ' Then just move the cursor
               ELSEIF sCol = MarkECol AND MarkECol > MarkSCol THEN ' At right and have room
                  MarkECol = MarkECol - 1                         ' Move right boundary
                  me.MarkScr                                      ' Redraw the hilight
                  CALL DWORD Pass: MExitMeth                      ' Then just move the cursor
               ELSE                                               '
                  CALL DWORD Pass: MExitMeth                      ' Just move the cursor
               END IF                                             '
               IF IsMarkActive THEN me.MarkScr                    ' Possibly redraw the highlight

            CASE "R"                                              ' Right
               IF sCol = MarkSCol AND MarkSCol < MarkECol THEN    ' At left end and have room?
                  MarkSCol = MarkSCol + 1                         ' Move left boundary
                  me.MarkScr                                      ' Redraw the hilight
                  CALL DWORD Pass: MExitMeth                      ' Then just move the cursor
               ELSEIF sCol = MarkECol THEN                        ' At right
                  MarkECol = MarkECol + 1                         ' Move right boundary
                  me.MarkScr                                      ' Redraw the hilight
                  CALL DWORD Pass: MExitMeth                      ' Then just move the cursor
               ELSE                                               '
                  CALL DWORD Pass: MExitMeth                      ' Just move the cursor
               END IF                                             '

            CASE "E"                                              ' End
               MarkECol = LEN(RTRIM$(LTxtG(lin)))                 ' Move right to text length column
               me.MarkScr                                         ' Redraw the hilight
               CALL DWORD Pass: MExitMeth                         ' Then just move the cursor

         END SELECT                                               '

      '----- No current marked area
      ELSE                                                        '
         IF dir <> "E" THEN                                       ' If not the END case
            MarkSLin = Lin: MarkELin = Lin                        ' Save 1 character area
            MarkSCol = sCol: MarkECol = sCol                      '
         ELSE                                                     ' END
            MarkSLin = Lin: MarkELin = Lin                        ' Save area to end of text
            MarkSCol = sCol: MarkECol = LEN(RTRIM$(LTxtG(lin)))   '
            CALL DWORD Pass                                       ' Then just move the cursor
         END IF                                                   '
         OnMarkActive                                             ' We're starting a marked area
         me.MarkScr                                               '
      END IF                                                      '
      MExit
   END METHOD

   METHOD  PTPaste()
   '---------- Paste text at cursor location in PowerType mode
   LOCAL ix, i, j, k AS LONG
   LOCAL ln, t, tt, MyDLM AS STRING
      MEntry
      CsrLinDX = 0                                                ' Reset possible hidden DX line location

      '----- Get clipboard data (1st line)
      sReadClipboard(t, MyDLM, %False)                            ' Go get Clipboard, keep the file
      ln = EXTRACT$(t, MyDLM)                                     ' Get 1st line from Clipboard

      '----- Loop now through the PT lines
      FOR k = 1 TO gPTblCount                                     ' Loop for each PT line
         j = gPTbl(k).sRow                                        ' Get screen row
         ix = gPTbl(k).tLin                                       ' Get IX of real data line
         IF ISFALSE IsLData(ix) THEN ITERATE FOR                  ' Only Data lines
         me.ModSet(ix)                                            ' Remember we changed something
         i = CCol - gLNPadCol + Offset                            ' Calc index into string
         tt = LTxtG(ix)                                           ' Get working copy of text
         IF IsNotNsrtFlag THEN                                    ' Do the non-Insert mode handling
            IF i + LEN(ln) - 1 > LEN(tt) THEN tt = LSET$(tt , i + LEN(ln) - 1)  ' Lengthen if needed
            MID$(tt, i, LEN(ln)) = ln                             ' Overlay the CBD text
         ELSE                                                     ' Do the Insert method
            tt = me.DataInsert(tt, ln, i)                         ' Insert the data
            me.LAttrAdjust(ix, i, LEN(ln))                        ' Adjust any Attr hilite
         END IF                                                   '
         IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) THEN      ' CAPS ON?
            tt = UUCASE(tt)                                       ' UC if CAPS is ON
         END IF                                                   '
         me.LTxtSet(ix, tt)                                       ' Save it back
         me.AttrScan(ix)                                          ' Recolorize
         IF j > 0 THEN me.DispLine(ix, j)                         ' If visible on the screen Re-display the line
      NEXT k                                                      '
      i += LEN(ln)                                                ' Move cursor
      IF i > Offset + gDataLen - 1 THEN                           ' Off current screen?
         me.CurSetReq(%Position, gPTbl(1).tLin, i, %True)         ' No, Set cursor set attempt
         DoSet(%Refresh)                                          ' Have it looked at
      ELSE                                                        '
         me.CsrCol = i - Offset + gLNPadCol                       '
      END IF                                                      '
      TP.LastPTCurs = 0                                           ' Force cursor to draw
      sDoCursor                                                   ' To get back | | lines
      MExit
   END METHOD

   METHOD  PTPowerCopy()
   '---------- Copy selected columns from the PT range to the clipboard
   LOCAL x, i, j, k, ix AS LONG, lclTxt, CBD AS STRING
      '----- OK, now build the Clipboard string
      CBD = ""                                                    ' Start as ""
      FOR k = 1 TO gPTblCount                                     ' Loop for each PT line
         j = gPTbl(k).sRow                                        ' Get screen row
         ix = gPTbl(k).tLin                                       ' Get IX of real data line
         IF ISFALSE IsLData(ix) THEN ITERATE FOR                  ' Only Data lines
         lclTxt = LTxtG(ix)                                       ' Get the text
         IF LEN(lclTxt) < MarkECol THEN lclTxt = LSET$(lclTxt, MarkECol)' Padd if needed
         lclTxt = MID$(lclTxt, MarkSCol TO MarkECol)              ' Extract the columns
         CBD += lclTxt + $CRLF                                    ' Add each line with CR/LF
      NEXT k                                                      '
      IF sWriteClipboard(CBD) THEN _                              ' Write it
         me.MarkKill                                              ' Reset the select frame
   END METHOD

   METHOD  PTPowerCut()
   '---------- Cut selected columns from the PT range to the clipboard
   LOCAL x, i, j, k, ix AS LONG, lclTxt, CBD AS STRING
      '----- OK, now build the Clipboard string
      CBD = ""                                                    ' Start as ""
      me.MarkKill                                                 ' Reset the select frame
      FOR k = 1 TO gPTblCount                                     ' Loop for each PT line
         j = gPTbl(k).sRow                                        ' Get screen row
         ix = gPTbl(k).tLin                                       ' Get IX of real data line
         IF ISFALSE IsLData(ix) THEN ITERATE FOR                  ' Only Data lines
         me.ModSet(ix)                                            ' Remember we changed something
         lclTxt = LTxtG(ix)                                       ' Get the text
         IF LEN(lclTxt) < MarkECol THEN lclTxt = LSET$(lclTxt, MarkECol)' Padd if needed
         lclTxt = MID$(lclTxt, MarkSCol TO MarkECol)              ' Extract the columns
         CBD += lclTxt + $CRLF                                    ' Add each line with CR/LF
         lclTxt = LTxtG(ix)                                       ' Get the text again
         lclTxt = STRDELETE$(lclTxt, MarkSCol, MarkECol - MarkSCol + 1) ' Delete the characters
         me.LAttrAdjust(ix, MarkSCol, -(MarkECol - MarkSCol + 1)) ' Adjust any Attr hilite
         IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) THEN      ' CAPS ON?
            lclTxt = UUCASE(lclTxt)                               ' UC if CAPS is ON
         END IF                                                   '
         me.LTxtSet(ix, lclTxt)                                   ' Put work string back
         me.AttrScan(ix)                                          ' Recolorize
         IF j > 0 THEN me.DispLine(ix, j)                         ' If visible on the screen Re-display the line
      NEXT k                                                      '
      sWriteClipboard(CBD)                                        ' Write it
      TP.LastPTCurs = 0                                           ' Force cursor to draw
      sDoCursor                                                   ' To get back | | lines
   END METHOD

   METHOD  PTPowerPaste()
   '---------- Paste whole clipboard at cursor location in PowerType mode
   LOCAL ix, i, j, k, d AS LONG
   LOCAL ln, t, tt, MyDLM AS STRING
   DIM CBD() AS STRING
      MEntry
      IF gPTblCount = 0 THEN MExitMeth                            ' No PT count, then nothing to do
      CsrLinDX = 0                                                ' Reset possible hidden DX line location

      '----- Get data from clipboard, exit if none
      sReadClipboard(t, MyDLM, %False)                            ' Go get Clipboard, keep the file
      IF LEN(t) < 1 THEN MExitMeth                                ' Error if nothing there
      REDIM CBD(1 TO PARSECOUNT(t, MyDLM))                        ' Get the lines in the clipboard
      PARSE t, CBD(), MyDLM                                       '

      '----- Paste into the PT lines
      FOR k = 1 TO gPTblCount                                     ' Loop for each PT line
         j = gPTbl(k).sRow                                        ' Get screen row
         ix = gPTbl(k).tLin                                       ' Get IX of real data line
         IF ISFALSE IsLData(ix) THEN ITERATE FOR                  ' Only Data lines
         me.ModSet(ix)                                            ' Remember we changed something
         INCR d: IF d > UBOUND(CBD) THEN EXIT FOR                 ' If data line # > lines in clipboard, we're done
         ln = CBD(d)                                              ' Set ln to CBD data
         i = CCol - gLNPadCol + Offset                            ' Calc index into string
         tt = LTxtG(ix)                                           ' Get working copy of text
         IF IsNotNsrtFlag THEN                                    ' Do the non-Insert mode handling
            IF i + LEN(ln) - 1 > LEN(tt) THEN tt = LSET$(tt , i + LEN(ln) - 1)  ' Lengthen if needed
            MID$(tt, i, LEN(ln)) = ln                             ' Overlay the CBD text
         ELSE                                                     ' Do the Insert method
            tt = IIF$(i = 1, ln + tt, LEFT$(tt, i - 1) + ln + MID$(tt, i))
            me.LAttrAdjust(ix, i, LEN(ln))                        ' Adjust any Attr hilite
         END IF                                                   '
         IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) THEN      ' CAPS ON?
            tt = UUCASE(tt)                                       ' UC if CAPS is ON
         END IF                                                   '
         me.LTxtSet(ix, tt)                                       ' Save it back
         me.AttrScan(ix)                                          ' Recolorize
         IF j > 0 THEN me.DispLine(ix, j)                         ' If visible on the screen Re-display the line
      NEXT k                                                      '
      i += LEN(ln)                                                ' Move cursor
      IF i > Offset + gDataLen - 1 THEN                           ' Off current screen?
         me.CurSetReq(%Position, gPTbl(1).tLin, i, %True)         ' No, Set cursor set attempt
         DoSet(%Refresh)                                          ' Have it looked at
      ELSE                                                        '
         me.CsrCol = i - Offset + gLNPadCol                       '
      END IF                                                      '
      TP.LastPTCurs = 0                                           ' Force cursor to draw
      sDoCursor                                                   ' To get back | | lines
      MExit
   END METHOD

   METHOD  PTRight()
   '---------- Increment Column in PowerType mode
   LOCAL ix AS LONG
   LOCAL t AS STRING
      me.CsrColAdd(1)                                             ' Simple INCR
      IF CCol > ENV.ScrWidth THEN                                 ' At right side of screen?
         me.CsrColSub(1)                                          ' Yes, Undo the INCR for now
         ix = gPTbl(1).tLin                                       ' Locate data line cursor was on
         me.CurSetReq(%Position, ix, CCol + Offset - gLNPadCol + 1, %False) ' Set cursor set attempt
         me.OffsetAdd(1)                                          ' Bump the offset
         DoSet(%Refresh)                                          ' Have it looked at
      END IF                                                      '
   END METHOD

   METHOD  PTRightScroll()
   '---------- Scroll RIGHT in Powertype mode
   DIM ScrAmt AS LONG, ScrCmd AS STRING
      '----- Get scroll amount
      ScrCmd = Prf.Scroll                                         ' Pick up the default scroll amount
      IF VAL(ScrCmd) <> 0 THEN                                    ' A Number?
         ScrAmt = VAL(ScrCmd)                                     ' Yes, use it
      ELSEIF ScrCmd = "HALF" THEN                                 ' Half?
         ScrAmt = INT(gDataLen / 2)                               ' Yes, set it
      ELSEIF ScrCmd = "PAGE" THEN                                 ' Page?
         ScrAmt = gDataLen                                        ' Yes, set it
      ELSEIF ScrCmd = "FULL" THEN                                 ' Full?
         ScrAmt = gDataLen                                        ' Yes, set it
      ELSEIF ScrCmd = "DATA" THEN                                 ' Data?
         ScrAmt = gDataLen - 1                                    ' Yes, set it
      ELSEIF IsCData THEN                                         ' No, In data area?
         ScrAmt = CCol - gLNData1                                 ' Yes, move cursor to left of screen
         IF ScrAmt = 0 THEN ScrAmt = gDataLen - 1                 ' If already there go a default amount
      END IF                                                      '

      '----- Apply the scroll amount
      me.OffsetAdd(ScrAmt)                                        '
      me.CurSetReq(%Position, gPTbl(1).tLin, CCol - gLNPadCol + Offset - ScrAmt, %False)

      DoSet(%Refresh)                                             ' Have it looked at
   END METHOD

   METHOD  PTSentenceCase()
   '---------- Sentence case a block in Power Type mode
   DIM words() AS STRING
   LOCAL Txt1, txt2 AS STRING
   LOCAL a, b, i, j, k, m, x, Caps, ix AS LONG
   LOCAL FromCol, ToCol AS LONG

      MEntry
      me.MarkKill                                                 ' Reset the select frame
      Caps = %True                                                ' 1st word overall gets capitalized

      '----- Loop through the PT lines
      FOR k = 1 TO gPTblCount                                     ' Loop for each PT line
         j = gPTbl(k).sRow                                        ' Get screen row
         ix = gPTbl(k).tLin                                       ' Get IX of real data line
         me.ModSet(ix)                                            ' Remember we changed something
         Txt1 = LLCASE(MID$(LTxtG(ix), MarkSCol TO MarkECol))     ' Get copy of substring to process
         GOSUB DoAChunk                                           ' Go process this piece
         txt2 = LTxtG(ix)                                         ' Get whole text line
         MID$(txt2, MarkSCol TO MarkECol) = Txt1                  ' Stuff back in the processed Txt1
         me.LTxtSet(ix, txt2)                                     ' Replace the whole line
         me.AttrScan(ix)                                          ' Recolorize
         IF j > 0 THEN me.DispLine(ix, j)                         ' If visible on the screen Re-display the line
      NEXT k                                                      '
      TP.LastPTCurs = 0                                           ' Force cursor to draw
      sDoCursor                                                   ' To get back | | lines
      MExitMeth                                                   '

   DoAChunk:
      i = PARSECOUNT(Txt1, " ")                                   ' Count words
      REDIM words(1 TO i) AS STRING                               ' Make table correct size
      PARSE Txt1, words(), " "                                    ' Parse the words into the table
      m = 1                                                       ' Start scan at 1
      FOR b = 1 TO i                                              ' Loop for each word
         IF ISNOTNULL(words(b)) THEN                              ' Ignore null 'words'
            a = INSTR(m, Txt1, words(b))                          ' Find location of this word
            IF caps THEN                                          ' Capitalize this word?
               IF a THEN MID$(Txt1, a, 1) = UUCASE(MID$(Txt1, a, 1))' Uppercase the 1st letter
               Caps = %False                                      ' Off till next period
            END IF                                                '
            IF RIGHT$(words(b), 1) = "."  OR _                    ' Look for sentence trigger
               RIGHT$(words(b), 1) = "?"  OR _                    '
               RIGHT$(words(b), 1) = "!"  THEN                    '
               Caps = %True                                       ' Restart sentence if required
            END IF                                                '
            m = a + 1                                             ' Start scan adjust
         END IF                                                   '
      NEXT b                                                      '
      RETURN                                                      '
   END METHOD

   METHOD  PTSeq()
   '---------- Do the Seq function
   LOCAL ix, i, j, k, ln, enumz AS LONG
   LOCAL t, tt, mprefix, msuffix, nv AS STRING
      MEntry
      '----- Do some basic setup
      CsrLinDX = 0                                                ' Reset possible hidden DX line location
      ix = gPTbl(1).tLin                                          ' Get IX of 1st PT line
      me.MarkKill                                                 ' Reset the select frame
      t = LTxtG(ix)                                               ' Get the text line
      IF LEN(t) < MarkECol THEN t = LSET$(t, MarkECol)            ' Padd if needed
      t = MID$(t, MarkSCol TO MarkECol)                           ' Extract the columns containg the model

      '----- See if there's a suffix
      IF INSTR($Numeric, RIGHT$(t, 1)) > 0 THEN                   ' If it ends in a digit,
         msuffix = ""                                             ' then there is NO suffix
      ELSE                                                        ' Non-digit, we must extract the suffix
         i = INSTR(-1, t, ANY $Numeric)                           ' Look for the last digit
         IF i = 0 THEN                                            ' None? whoops
            tt = "Seq model has no starting value"                '
            GOSUB DispMsg                                         '
            MExitMeth                                             '
         END IF                                                   '
         msuffix = MID$(t, i + 1)                                 ' Found digit, extract suffix
         t = LEFT$(t, i)                                          ' Strip it off the model
      END IF                                                      '

      '----- Now extract the prefix
      j = 0                                                       ' Reset j
      FOR i = LEN(t) TO 1 STEP -1                                 ' Look backward for 1st non-digit
         IF INSTR($Numeric, MID$(t, i, 1)) > 0 THEN ITERATE FOR   ' If a digit, keep looking
         j = i: EXIT FOR                                          ' Found a non digit, save its loc and exit for
      NEXT i                                                      '
      IF j = 0 THEN                                               ' All were digits
         mprefix = ""                                             ' Therefor no prefix
      ELSE                                                        ' Non-digit, we have a prefix
         mprefix = LEFT$(t, j)                                    ' Save prefix
         t = MID$(t, j + 1)                                       ' Extract it leaving t as the eval string.
      END IF                                                      '

      '----- Extract the enum value
      ln = LEN(t)                                                 ' Save length for rebuilding
      IF LEN(t) <> LEN(FORMAT$(VAL(t))) THEN enumz = %True        ' Entered as leading zeros?, remember that

      '----- Now modify the PT lines
      FOR k = 1 TO gPTblCount                                     ' Loop for each PT line
         j = gPTbl(k).sRow                                        ' Get screen row
         ix = gPTbl(k).tLin                                       ' Get IX of real data line
         IF ISFALSE IsLData(ix) THEN ITERATE FOR                  ' Only Data lines
         me.ModSet(ix)                                            ' Remember we changed something
         tt = LTxtG(ix)                                           ' Get working copy of text
         IF LEN(tt) < MarkECol THEN tt = LSET$(tt, MarkECol)      ' Padd if needed
         t = FORMAT$(VAL(LLnumG(ix)))                                          '
         GOSUB AdjustPad                                          ' Do padding etc.

         '----- Build the value and stuff it in
         t = mprefix + t + msuffix                                ' Add prefix and suffix
         t = RIGHT$(RSET$(t, MarkECol - MarkSCol + 1), MarkECol - MarkScol + 1) ' Right align and trim it
         MID$(tt, MarkSCol, LEN(t)) = t                           ' Overlay the enum text
         me.LTxtSet(ix, tt)                                       ' Save it back
         me.AttrScan(ix)                                          ' Recolorize

         '----- If on the screen, display it
         IF j > 0 THEN me.DispLine(ix, j)                         ' If visible on the screen Re-display the line
      NEXT k                                                      '
      TP.LastPTCurs = 0                                           ' Force cursor to draw
      sDoCursor                                                   ' To get back | | lines
      MExitMeth

   '----- Fiddle with the prefix padding
   AdjustPad:
      IF LEN(t) > ln AND RIGHT$(mprefix, 1) = " " THEN            ' If enum needs more room, and Prefix has some
         mprefix = LEFT$(mprefix, LEN(mprefix) - 1)               ' Steal right hand blank from prefix
         INCR ln                                                  ' Make enum longer
      ELSEIF LEN(t) > ln AND LEFT$(mprefix, 1) = " " THEN         ' If not right end, try left end
         mprefix = MID$(mprefix, 2)                               ' Steal left hand blank from prefix
         INCR ln                                                  ' Make enum longer
      END IF                                                      '
      IF LEN(t) > ln AND RIGHT$(msuffix, 1) = " " THEN            ' If enum needs more room, and Suffix has some
         msuffix = LEFT$(msuffix, LEN(msuffix) - 1)               ' Steal right hand blank from suffix
         INCR ln                                                  ' Make enum longer
      ELSEIF LEN(t) > ln AND LEFT$(msuffix, 1) = " " THEN         ' If not right end, try left end
         msuffix = MID$(msuffix, 2)                               ' Steal left hand blank from suffix
         INCR ln                                                  ' Make enum longer
      END IF                                                      '
      IF enumz THEN                                               ' Pad appropriately
         t = RIGHT$(REPEAT$(ln, "0") + t, ln)                     '
      ELSE                                                        '
         t = RIGHT$(REPEAT$(ln, " ") + t, ln)                     '
      END IF                                                      '
      RETURN

   '----- Temporarily display an error message
   DispMsg:
      t = STRING$(ENV.ScrWidth - LEN(tt), "_")                    ' Build LH part of dash line
      sPrint (t, $$TxtLo, 2, 1)                                   ' Print LH part of line 2
      sPrint (tt, $$Error, 2, ENV.ScrWidth - LEN(tt) + 1)         ' Print rest of line
      GRAPHIC REDRAW                                              '
      SLEEP 1500                                                  '
      tt = "Entering PowerType mode, Press Enter to exit"         '
      sPrint (t, $$TxtLo, 2, 1)                                   ' Print LH part of line 2
      sPrint (tt, $$Error, 2, ENV.ScrWidth - LEN(tt) + 1)         ' Print rest of line
      GRAPHIC REDRAW                                              '
      RETURN                                                      '
   END METHOD

   METHOD  PTSwap()
   '---------- Swap chunks on a line
   LOCAL x, i, j, k, ix AS LONG, lclTxt, CBD AS STRING
   LOCAL Lin, sCol, OldLin, NewLin, S1, S2, SNum AS LONG
   LOCAL l1, l2, t1, t2, p1, p2, p3, p4 AS STRING
   LOCAL sl, el, ll AS LONG

      MEntry
      '----- If both areas marked, we can do the Swap
      IF IsMarkActive AND IsSwapActive THEN                       ' We have two marked areas ready

         '----- Init for line loop
         me.MarkKill                                              ' Kill the mark area
         me.SwapKill                                              ' Kill the Swap area
         FOR k = 1 TO gPTblCount                                  ' Loop for each PT line
            j = gPTbl(k).sRow                                     ' Get screen row
            s1 = gPTbl(k).tLin                                    ' Get IX of real data line
            IF ISFALSE IsLData(S1) THEN ITERATE FOR               ' Only Data lines

            '----- Extract the two strings
            IF LEN(L(s1).@LTxt) < MAX(SwapECol, MarkECol) THEN me.LTxtSet(s1, LSET$(LTxtG(s1), MAX(SwapECol, MarkECol)))
            l1 = LTxtG(s1)                                        ' Get the text from the Swap line
            t1 = MID$(l1, SwapSCol TO SwapECol)                   ' Extract the Swap columns
            t2 = MID$(l1, MarkSCol TO MarkECol)                   ' Extract the columns
            me.ModSet(S1)                                         ' Remember we changed something

            '----- Swap on a single line
            IF SwapSCol < MarkSCol THEN                           ' Swap is left of Mark
               '----- Swap area is Left of the Mark area
               p1 = LEFT$(l1, SwapSCol - 1)                       ' Get part 1, the left of Swap string
               p2 = IIF$(SwapEcol + 1 = MarkSCol, "", MID$(l1, SwapECol + 1, MarkSCol - SwapECol -1)) ' Get part 2, between the strings
               p3 = MID$(l1, MarkECol + 1)                        ' Get part 3, After the strings
               l1 = BUILD$(p1, t2, p2, t1, p3)                    ' Rebuild the line

            ELSE                                                  ' Mark is left of swap
               '----- Swap area is Right of the Mark area
               p1 = LEFT$(l1, MarkSCol - 1)                       ' Get part 1, the left of Mark string
               p2 = IIF$(MarkEcol + 1 = SwapSCol, "", MID$(l1, MarkECol + 1, SwapSCol - MarkECol -1)) ' Get part 2, between the strings
               p3 = MID$(l1, SwapECol + 1)                        ' Get part 3, After the strings
               l1 = BUILD$(p1, t1, p2, t2, p3)                    ' Rebuild the line
            END IF                                                '
            IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) THEN              ' CAPS ON?
               lclTxt = UUCASE(lclTxt)                            ' UC if CAPS is ON
            END IF                                                '
            me.LTxtSet(S1, l1)                                    ' Save the Swap line back
            me.AttrScan(S1)                                       ' Recolorize
            IF j > 0 THEN me.DispLine(s1, j)                      ' If visible on the screen Re-display the line

         NEXT k                                                   '

      '----- Only Mark area active, start a Swap activity
      ELSEIF IsMarkActive AND ISFALSE IsSwapActive THEN           ' We have a MARK but no SWAP, start a Swap
         SwapSCol = MarkSCol                                      ' Copy marked area as a swap area
         SwapECol = MarkECol                                      '
         SwapSLin = MarkSLin                                      '
         SwapELin = MarkELin                                      '
         SwapLines = 1                                            ' Save # of data lines
         me.MarkKill                                              ' Kill the mark area
         OnSwapActive                                             ' Make Swap active
         me.SwapScr                                               ' Draw the underline

      '----- No Mark area, See if Move more, or Kill time
      ELSEIF ISFALSE IsMarkActive AND IsSwapActive THEN           ' We have swap and no Mark, See if maybe Move mode
         j = sGetIX(CRow)                                         ' Get line number of the cursor
         i = CCol - gLNPadCol + Offset                            ' Calc column index into line
         IF i >= SwapSCol AND i <= SwapECol THEN GOSUB KillSwap: MExitMeth ' Kill if

         '----- This looks like a SWAP Move mode request
         '----- Init for line loop
         me.MarkKill                                              ' Kill the mark area
         me.SwapKill                                              ' Kill the Swap area
         FOR k = 1 TO gPTblCount                                  ' Loop for each PT line
            j = gPTbl(k).sRow                                     ' Get screen row
            s1 = gPTbl(k).tLin                                    ' Get IX of real data line

            '----- Extract the two line
            l1 = LTxtG(S1)                                        ' Get the text from the Swap line
            IF LEN(l1) < j THEN l1 = LSET$(l1, j)                 ' Make sure lines are as long as the mark point

            '----- Swap on a single line
            IF SwapSCol < i THEN                                  ' Swap is left of Mark

               '----- Swap area is Left of the Mark area
               p1 = LEFT$(l1, SwapSCol - 1)                       ' Get part 1, the left of Swap string
               p2 = MID$(l1, SwapECol + 1 TO i - 1)               ' Get part 2, between the strings
               p3 = MID$(l1, SwapSCol TO SwapECol)                ' Get part 3, The swap string
               p4 = MID$(l1, i)                                   ' Get part 4, remainder of line
               l1 = BUILD$(p1, p2, p3, p4)                        ' Rebuild the line

            ELSE                                                  ' Mark is left of swap

               '----- Swap area is Right of the Mark point
               p1 = LEFT$(l1, i - 1)                              ' Get part 1, the left of Mark string
               p2 = MID$(l1, SwapSCol TO SwapECol)                ' Get part 2, the swap string
               p3 = MID$(l1, i TO SwapSCol - 1)                   ' Get part 3, Between mark and swap string
               p4 = MID$(l1, SwapECol + 1)                        ' Get part 4, After the Swap string
               l1 = BUILD$(p1, p2, p3, p4)                        ' Rebuild the line
            END IF                                                '
            IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) THEN              ' CAPS ON?
               lclTxt = UUCASE(lclTxt)                            ' UC if CAPS is ON
            END IF                                                '
            me.LTxtSet(S1, l1)                                    ' Save the Swap line back
            me.ModSet(S1)                                         ' Remember we changed something
            me.AttrScan(S1)                                       ' Recolorize
            IF j > 0 THEN me.DispLine(S1, j)                      ' If visible on the screen Re-display the line
         NEXT k                                                   ' Onward
         DoSet(%Refresh)                                          ' Have it looked at

      '----- Anything else, ignore
      ELSE                                                        ' Nothing going on, just exit
         MExitMeth                                                '

      END IF                                                      '
      TP.LastPTCurs = 0                                           ' Force cursor to draw
      sDoCursor                                                   ' To get back | | lines
      MExitMeth

   KillSwap:
      IF IsSwapDrawn THEN me.SwapKill                             ' Kill underline if drawn
      OffSwapActive                                               ' Now kill Swap
      RETURN
   END METHOD

   METHOD  PTTab()
   '---------- Tab forward To Next field in PowerType mode
   LOCAL i, j, k AS LONG, lclTabs AS STRING
      CsrLinDX = 0                                                ' Reset possible hidden DX line location
      j = gPTbl(1).tLin                                           ' Get Line number of cursor line
      lclTabs = me.TabsSimple(LEN(L(j).@LTxt))                    ' Get a working tabs line
      k = 0                                                       ' Say no next line
      IF ISFALSE Prf.Tabs THEN EXIT METHOD                        ' If not tabbing, exit
      i = CCol - gLNPadCol + Offset                               ' Calc index into data string
      i = INSTR(i + 1, lclTabs, "*")                              ' Look for next tab marker
      IF i THEN                                                   ' We have one
         IF i < Offset + 1 OR i > Offset + gDataLen - 1 THEN      ' Still on current screen
            me.CurSetReq(%Position, j, i, %True)                  ' No, Set cursor set attempt
            DoSet(%Refresh)                                       ' Have it looked at
            MExitMeth                                             '
         ELSE                                                     '
            me.CsrCol = i + gLNPadCol - Offset                    ' Just move cursor on the screen
            MExitMeth                                             '
         END IF                                                   '
      END IF                                                      '
   END METHOD

   METHOD  PTTime()
   '---------- Paste Date at cursor location in PowerType mode
   LOCAL ln AS STRING
      ln = sTime()                                                ' Go get the date
      me.PTTDPaste(ln)                                            ' Go do common code
   END METHOD

   METHOD  PTTitleCase()
   '---------- TitleCase a selected block in Power Type mode
   LOCAL x, i, j, k, ix AS LONG, lclTxt AS STRING
      me.MarkKill                                                 ' Reset the select frame

      '----- OK, now do the Titlecase
      FOR k = 1 TO gPTblCount                                     ' Loop for each PT line
         j = gPTbl(k).sRow                                        ' Get screen row
         ix = gPTbl(k).tLin                                       ' Get IX of real data line
         me.ModSet(ix)                                            ' Remember we changed something
         me.TitleCase(ix, MarkSCol, MarkECol)                     ' Go process it
         IF j > 0 THEN me.DispLine(ix, j)                         ' If visible on the screen Re-display the line
      NEXT k                                                      '
      TP.LastPTCurs = 0                                           ' Force cursor to draw
      sDoCursor                                                   ' To get back | | lines
   END METHOD

   METHOD  PTTDPaste(ln AS STRING)
   '---------- Paste in the passed string in PowerType mode
   LOCAL ix, i, j, k AS LONG
   LOCAL t, tt AS STRING
      CsrLinDX = 0                                                ' Reset possible hidden DX line location

      '----- Paste in the passed parameter into each PT line
      FOR k = 1 TO gPTblCount                                     ' Loop for each line
         j = gPTbl(k).sRow                                        ' Get screen row
         ix = gPTbl(k).tLin                                       ' Get IX of real data line
         me.ModSet(ix)                                            ' Remember we changed something
         IF ISFALSE IsLData(ix) THEN ITERATE FOR                  ' Only Data lines
         i = CCol - gLNPadCol + Offset                            ' Calc index into string
         tt = LTxtG(ix)                                           ' Get working copy of text
         IF IsNotNsrtFlag THEN                                    ' Do the non-Insert mode handling
            IF LEN(tt) - i < LEN(ln) THEN tt = tt + SPACE$(i + LEN(ln) - LEN(tt))  ' Lengthen if needed
            MID$(tt, i, LEN(ln)) = ln                             ' Overlay the CBD text
         ELSE                                                     ' Do the Insert method
            tt = IIF$(i = 1, ln + tt, LEFT$(tt, i - 1) + ln + MID$(tt, i))
            me.LAttrAdjust(ix, i, LEN(ln))                        ' Adjust any Attr hilite
         END IF                                                   '
         IF (Prf.CapsDesired = 1 OR Prf.CapsActual = 1) THEN      ' CAPS ON?
            tt = UUCASE(tt)                                       ' UC if CAPS is ON
         END IF                                                   '
         me.LTxtSet(ix, tt)                                       ' Save it back
         me.AttrScan(ix)                                          ' Recolorize
         IF j > 0 THEN me.DispLine(ix, j)                         ' If visible on the screen Re-display the line
         IF CCol < Offset + 1 OR CCol > Offset + gDataLen - 1 THEN' Still on current screen
            me.CurSetReq(%Position, ix, CCol, %True)              ' No, Set cursor set attempt
            DoSet(%Refresh)                                       ' Have it looked at
         END IF                                                   '
      NEXT k                                                      '
      me.CsrCol = CCol + LEN(ln)                                  ' Move cursor
      TP.LastPTCurs = 0                                           ' Force cursor to draw
      sDoCursor                                                   ' To get back | | lines
   END METHOD

   METHOD  PTUpperCase()
   '---------- UppeCase a selected block in PowerType mode
   LOCAL x, i, j, k, ix AS LONG, lclTxt AS STRING
      me.MarkKill                                                 ' Reset the select frame

      '----- OK, now do the Lowercase
      FOR k = 1 TO gPTblCount                                     ' Loop for each PT line
         j = gPTbl(k).sRow                                        ' Get screen row
         ix = gPTbl(k).tLin                                       ' Get IX of real data line
         lclTxt = LTxtG(ix)                                       ' Get the text
         me.ModSet(ix)                                            ' Remember we changed something
         IF MarkSCol <= LEN(lclTxt) THEN                          ' Do only if long enough
            FOR i = MarkSCol TO MIN(MarkECol, LEN(lclTxt))        ' Just ones that are there
               MID$(lclTxt, i, 1) = UUCASE(MID$(lclTxt, i, 1))    '
            NEXT i                                                '
            me.LTxtSet(ix, lclTxt)                                ' Put work string back
            me.AttrScan(ix)                                       ' Recolorize
            IF j > 0 THEN me.DispLine(ix, j)                      ' If visible on the screen Re-display the line
         END IF                                                   '
      NEXT k                                                      '
      TP.LastPTCurs = 0                                           ' Force cursor to draw
      sDoCursor                                                   ' To get back | | lines
   END METHOD
