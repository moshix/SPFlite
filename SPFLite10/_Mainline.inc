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

'------------------------------+
'           Keyboard Routines  |
'------------------------------+

FUNCTION MainAKeyHook(BYVAL iCode AS INTEGER, BYVAL WPARAM AS LONG, BYVAL LPARAM AS LONG) AS DWORD
'---------- Examine all keyboard traffic
DIM KeyStateArray(0 TO 255) AS BYTE
LOCAL KeyCode AS LONG
LOCAL ScanCode AS LONG
LOCAL ShiftState AS INTEGER
LOCAL ExtState AS INTEGER
LOCAL AltState AS INTEGER
LOCAL NumState AS INTEGER
LOCAL CtlState AS INTEGER
LOCAL Chars AS STRING * 2
LOCAL i, j AS LONG
LOCAL KMsg AS kbMsg
LOCAL KeyLookup, t AS STRING
STATIC LastMode AS STRING, LastKey AS LONG

   MEntry
   '----- Housekeeping
   CallNextHookEx(hKbrdHook, iCode, WPARAM, LPARAM)               ' Pass it on
   IF iCode < 0 THEN MExitFunc                                    ' Not for us

   '----- If we have a popup active, ignore keyboard
   IF ISTRUE(gfDoingMsg) THEN MExitFunc                           ' Popup active, bail out
   IF iCode <> %HC_ACTION THEN MExitFunc                          ' Exit if not a Keyboard action

   '----- Looks like we have something to do
   KeyCode = WPARAM                                               ' Get the keycode
   ScanCode = LPARAM                                              ' Get the ScanCode
   SHIFT LEFT ScanCode, 8                                         '
   SHIFT RIGHT ScanCode, 24                                       '

   '---------- Get flags set for special keys                     '
   gKeyChr = ""                                                   ' Clear our ultimate answer
   KeyLookup = "  ..."                                            ' Reset lookup area
   GetKeyboardState(BYVAL VARPTR(KeyStateArray(0)))               ' Get all key states etc
   ShiftState = ISTRUE(LOWRD(GetKeyState(%VK_SHIFT)) AND &H8000)  ' Collect all the status
   CtlState = ISTRUE(GetKeyState(%VK_CONTROL) AND &H8000)         '
   AltState = ISTRUE((LPARAM AND &H20000000) <> 0)                '
   NumState = ISTRUE GetKeyState(%VK_NUMLOCK)                     '
   ExtState = ISTRUE((LPARAM AND &H01000000) <> 0)                '

   '----- Somehow nothing, exit quickly
   IF KeyCode = 0 THEN MExitFunc                                  ' Bail out if somehow nothing

   '----- Build the key to lookup in the KB table
   MID$(KeyLookup, 1, 1) = CHR$(KeyCode)                          ' Start building the KeyLookup
   IF ExtState THEN MID$(KeyLookup, 2, 1) = "E"                   ' Add External Flag if needed
   IF KeyCode = %VK_SNAPSHOT AND AltState THEN                    ' PrtScrn is special
      MID$(KeyLookup, 2, 1) = "E"                                 ' Since Alt=PrtScr has no ExtState ????
   END IF                                                         '
   IF ShiftState THEN MID$(KeyLookup, 3, 1) = "S"                 ' Add Shift Flag if needed
   IF CtlState THEN MID$(KeyLookup, 4, 1) = "C"                   ' Add Ctl Flag if needed
   IF AltState THEN MID$(KeyLookup, 5, 1) = "A"                   ' Add Alt Flag if needed

   '----- First of many fudges for KB peculiarities
   IF KeyLookup = CHR$(&H3) + "E.C." THEN                         ' Fudge Ctl-Pause
      KeyLookUp = CHR$(&H13) + " .C."                             '
   END IF

'   debug  "(" + KeyLookup + ") NumState " + hex$(asc(mid$(KeyLookup, 1,1))) + HEX$(ASC(MID$(KeyLookup, 2,1))) + HEX$(ASC(MID$(KeyLookup, 3,1))) + HEX$(ASC(MID$(KeyLookup, 4,1)))

   '----- Another fudge for KeyPad handling
   IF ISTRUE NumState THEN                                        ' Try and overcome Keypad stupidities
      IF KeyCode = &H6F THEN KeyLookup = CHR$(&HBF) + " ..."      ' / Key
      IF KeyCode = &H6A THEN KeyLookup = CHR$(&H38) + " S.."      ' * Key
      IF KeyCode = &H6D THEN KeyLookup = CHR$(&HBD) + " ..."      ' - Key
      IF KeyCode = &H67 THEN KeyLookup = CHR$(&H37) + " ..."      ' 7 Key
      IF KeyCode = &H68 THEN KeyLookup = CHR$(&H38) + " ..."      ' 8 Key
      IF KeyCode = &H69 THEN KeyLookup = CHR$(&H39) + " ..."      ' 9 Key
      IF KeyCode = &H64 THEN KeyLookup = CHR$(&H34) + " ..."      ' 4 Key
      IF KeyCode = &H65 THEN KeyLookup = CHR$(&H35) + " ..."      ' 5 Key
      IF KeyCode = &H66 THEN KeyLookup = CHR$(&H36) + " ..."      ' 6 Key
      IF KeyCode = &H61 THEN KeyLookup = CHR$(&H31) + " ..."      ' 1 Key
      IF KeyCode = &H62 THEN KeyLookup = CHR$(&H32) + " ..."      ' 2 Key
      IF KeyCode = &H63 THEN KeyLookup = CHR$(&H33) + " ..."      ' 3 Key
      IF KeyCode = &H60 THEN KeyLookup = CHR$(&H30) + " ..."      ' 0 Key
      IF KeyCode = &H6B THEN KeyLookup = CHR$(&HBB) + " S.."      ' + Key
      IF KeyCode = &H6E THEN KeyLookup = CHR$(&HBE) + " ..."      ' . Key
   END IF                                                         '

   '----- On KeyDown watch out for Ctl and Alt keys
   IF (LPARAM AND &H80000000) = 0 THEN                            ' If Key Down Event, remember Ctrl and Alt
      IF KeyCode = &H11 OR KeyCode = &H12 THEN                    ' If Ctrl or Alt
         LastKey = KeyCode                                        ' Remember we saw them go down
      ELSE                                                        '
         LastKey = 0                                              ' Intervening key, clear last
      END IF                                                      '
   END IF                                                         '

'   debug "(" + KeyLookup + ") Lookup " + HEX$(ASC(MID$(KeyLookup, 1,1))) + HEX$(ASC(MID$(KeyLookup, 2,1))) + HEX$(ASC(MID$(KeyLookup, 3,1))) + HEX$(ASC(MID$(KeyLookup, 4,1))) + HEX$(ASC(MID$(KeyLookup, 5,1)))

   '----- Go see if the key has been assigned a mapping
   j = KbdT.Scan(KeyLookup)                                       ' See if mapped?

   '----- If not mapped, maybe it's some weird OEM key.
   IF j = 0 THEN                                                  ' Not mapped, Hmmm, OEM type key?
      j = KbdT.ScanFull(KeyLookup)                                ' See if mapped?

      '----- See if in Master table, OEM key's won't be there, if there, it's a simple unmapped key
      IF j THEN                                                   ' Yes, then just this combo not wanted
         FUNCTION = 1                                             '
         MExitFunc                                                ' If not wanted, just exit

      '----- Not in master table, must be an OEM key, send onward with dummy (Passthru) entry
      ELSE                                                        ' If not found, possible OEM key
         j = 1                                                    ' Point at the dummy all (Passthru) entry
      END IF                                                      '
   END IF                                                         '

   '----- Another fudge for PrintScrn; Do the KeyUp/KeyDown handling
   IF (LPARAM AND &H80000000) <> 0 THEN                           ' If Key Up Event
      IF KeyCode <> %VK_SNAPSHOT THEN                             ' Print Screen is different
         KbdT.SetSupSent(j, "N")                                  ' Reset 1st Sent flag
         IF KeyCode <> LastKey THEN MExitFunc                     ' If not the Ctrl or Alt just going up
      END IF                                                      '
   ELSE                                                           ' Key Down event
      IF Kbdt.SupRepeat(j) = "U" THEN                             ' An attention or special key?
         IF Kbdt.SupSent(j) =  "Y" THEN                           ' If we've already sent the kb-Down
            MExitFunc                                             ' Ignore these on Key Down
         ELSE                                                     ' Else this is the 1st one
            KbdT.SetSupSent(j, "Y")                               ' Remember we're sending it
         END IF                                                   '
      END IF                                                      '
   END IF                                                         '

   '----- Finally let Windows figure out what ANSI character it is
   i = ToAsciiEx(KeyCode, ScanCode, BYREF KeyStateArray(0), _     ' Let Windows decypher the character
                       BYVAL VARPTR(Chars), 0, GetKeyboardLayout(0))
   '----- Ack! Foreign KB mumbo jumbo
   IF i < 0 THEN
      FUNCTION = 1: MExitFunc                                     ' Diacritic Dead key?  Set RC = 1 and Exit
   END IF
   IF i = 2 THEN                                                  ' Unused dead key
      KMsg.kbString = CHR$(0) + LEFT$(Chars, 1) + "  "            ' Pass it back
      i = PostMessage(hWnd, %WM_USER, KMsg.MsgwParam, 0)          '
      i = 1: MID$(Chars, 1, 1) = MID$(Chars, 2, 1)                ' Make remainder look like only one key now
   END IF                                                         '

   '----- If a key mapped to something OTHER THAN (Passthru) send the KB table Index onward
   IF UUCASE(KbdT.KData(j)) <> "(PASSTHRU)" THEN                  ' If not just PASSTHRU
      KMsg.kbString = CHR$(1, 0, 0, 0)                            ' Pass back pointer to Kbd() entry
      KMsg.kbInt(1) = j                                           ' Stuff gKbd() index into message
      i = PostMessage(hWnd, %WM_USER, KMsg.MsgwParam, 0)          '
      FUNCTION = 1                                                '
      MExitFunc                                                   '
   END IF                                                         '

   '----- If a key mapped to (Passthru) send the key onward
   IF i = 1 AND UUCASE(KbdT.KData(j)) = "(PASSTHRU)" THEN         ' Maybe a normal key
      KMsg.kbString = CHR$(0) + LEFT$(Chars, 1) + "  "            ' Pass back the key
      i = PostMessage(hWnd, %WM_USER, KMsg.MsgwParam, 0)          '
      FUNCTION = 1                                                '
   ELSE                                                           '
      FUNCTION = 0                                                '
   END IF                                                         '
   MExit
END FUNCTION

SUB      MainBUserMsg()
'---------- Handle a message block from the Keyboard hook
LOCAL currtab, i, j, k, m, ourtab AS LONG, KeyData, t AS STRING
LOCAL ix AS LONG, MSG AS kbMsg
   Msg.kbString = gKeyChr                                         ' Move string into Union
   IF Msg.kbString = CHR$(0,0,0,0) THEN EXIT SUB                  ' Ignore Zero case ????
   sCaretHide                                                     ' Turn off the Cursor

   SELECT CASE AS CONST$ LEFT$(Msg.kbString, 1)                   ' Which user message
      CASE CHR$(0)                                                ' Just a simple KB character
         GOSUB ChkMarkReset                                       ' See if we should reset Mark action
         '----- Reduce gKeyChr to that character
         gKeyChr = MID$(gKeyChr, 2, 1)                            ' Extract it

         '----- If we're in KB Record mode, record it
         IF gKbdRecFlag THEN                                      ' Are we recording?
            IF ISFALSE gKbdRecTxtFlag THEN                        ' Have to start?
               gKbdRecTxtFlag = %True                             ' No. Start it now
               gKbdRecording += "["                               ' Insert start frame character
            END IF                                                '
            IF INSTR("[]", gKeyChr) <> 0 THEN                     ' One of our frame chars
               gKbdRecording += gKeyChr                           ' Yes, add it twice
               gKbdRecording += gKeyChr                           '
            ELSE                                                  '
               gKbdRecording += gKeyChr                           ' No, just add it once
            END IF                                                '
         END IF                                                   '
         MainDKeyProcess                                          ' Pass to the normal process

      CASE CHR$(1)                                                ' gKbd() lookup entry message
         IF gKbdRecFlag THEN                                      ' Doing a recording?
            IF gKbdRecTxtFlag THEN                                ' Doing a text string?
               gKbdRecTxtFlag = %False                            ' Yes, Kill it.
               gKbdRecording += "]"                               ' Close off the text string
            END IF                                                '
         END IF                                                   '

         '----- Extract the assigned key map string, send it for processing
         Msg.kbString = gKeyChr                                   ' Extract gKbd() index
         ix = Msg.kbInt(1)                                        '
         KeyData = KbdT.KData(ix)                                 '
         MainCDoKeyString(KeyData)                                '

      CASE CHR$(2)                                                ' File Change notice
         IF TabsNum = 0 THEN EXIT SUB                             ' Prevent Memory Access violations at termination
         CurrTab = TP.PgNumber                                    ' Save Tab we are on
         IF ISFALSE gfActive THEN                                 ' If we're not active
            IF Msg.kbInt(1) > TabsNum THEN EXIT SUB               ' Is Tab gone? (e.g. a CAN DEL)
            IF TP.PgNumber <> Msg.kbInt(1) THEN                   ' Are we on a different Tab?
               TP = Tabs(msg.kbInt(1))                            ' Switch to the tab being notified
               IF ISNOTHING(TP) THEN EXIT SUB                     ' Bail out if we didn't set TP to a valid Tab
               TP.WatchFlag = MID$(Msg.kbString, 2, 1)            ' Copy the post type to WatchFlag
               TP = Tabs(CurrTab)                                 ' Swap back the previous tab
             ELSE                                                 ' We're on the correct tab
                TP.WatchFlag = MID$(Msg.kbString, 2, 1)           ' Copy the post type to WatchFlag
             END IF                                               '

         ELSE                                                     ' We are active
            IF TP.PgNumber <> Msg.kbInt(1) THEN                   ' We're active, Are we on a different Tab?
               TP = Tabs(msg.kbInt(1))                            ' Switch to the tab being notified
               IF ISNOTHING(TP) THEN EXIT SUB                     ' Bail out if we didn't set TP to a valid Tab
               gTabSwitch = msg.kbInt(1)                          ' Set tab to go to
            END IF                                                '
            TP.WatchFlag = MID$(Msg.kbString, 2, 1)               ' Copy the post type to WatchFlag
         END IF                                                   '

      CASE CHR$(3)                                                ' Mouse keyboard message
         KeyData = MID$(gKeyChr, 2)                               ' Strip off the ID byte
         MainCDoKeyString(KeyData)                                '

      CASE CHR$(4)                                                ' Dialog forced Attention
         TP.AttnDo = (TP.AttnDo OR %Attention)                    ' Set Attention required
         TP.PostKeyBoard                                          ' Go process the event

      CASE CHR$(5)                                                ' DO command list to perform
         FOR i = 1 TO UBOUND(gCmdList())                          ' For each item
            IF TRIM$(gCmdList(i)) = "" THEN ITERATE FOR           ' Ignore blank lines
            IF LEFT$(gCmdList(i), 1) = ";" THEN ITERATE FOR       ' Ignore ; in col 1 comment lines
            IF uucase(LEFT$(LTRIM$(gCmdList(i)), 3)) = "DO " THEN ITERATE FOR ' Ignore nested DO commands

            '----- Look for comments ON the line
            t = gCmdList(i): k = 1: m = 0                         ' Start in Col 1, use working copy
            j = INSTR(-1, t, " ;")                                ' Find last " ;" on the line
            IF j = 0 THEN                                         ' None?  Then no comment
               MainCDoKeyString(gCmdList(i))                      ' Pass this entrys on
            ELSE                                                  ' Potential comment
               DO UNTIL k = 0                                     ' Find last ]
                  k = INSTR(k, t, "[")                            ' Look for [
                  IF k THEN                                       ' Found [, look for ]
                     k = INSTR(k + 1, t, "]")                     '
                     IF k THEN                                    ' Got the ]
                        m = k                                     ' Save its location
                        INCR k                                    ' step over
                     END IF                                       ' Continue
                  END IF                                          '
               LOOP                                               ' On exit m contains last good ]
               IF m < j THEN                                      ' Is comment past all the [..] stuff
                  t = LEFT$(t, j - 1)                             ' Strip off the comment
                  MainCDoKeyString(RTRIM$(t))                     ' Pass the stripped entry on
               ELSE                                               '
                  MainCDoKeyString(gCmdList(i))                   ' Just pass this entry on
               END IF                                             '
            END IF                                                '
         NEXT i                                                   '

   END SELECT                                                     '

   sDoPendingTabDels                                              ' Do any pending Tab deletes
   sDoPendingTabSwitch                                            ' Do any pending Tab switch
   IF TabsNum = 0 THEN EXIT SUB                                   ' We're closing down
   TP.MarkLineNumb                                                '
   sCaretDestroy                                                  '
   sCaretCreate                                                   '
   sDoCursor                                                      ' Set the cursor
   sCaretShow                                                     '
   OurTab = TP.PgNumber                                           ' Save our page number
   sTabStackAdd(OurTab)                                           '
   EXIT SUB                                                       '

ChkMarkReset:                                                     ' See if we should reset Mark action
   '----- User started marking, but marked nothing
   IF IsTPMarkActive AND ISFALSE gfLeftDown AND ISFALSE IsTPMarkDrawn THEN OffTPMarkActive    ' MarkActive but nothing marked? Kill it
   IF IsTPMiscActive AND ISFALSE gfLeftDown AND ISFALSE IsTPMiscDrawn THEN OffTPMiscActive    ' MiscActive but nothing marked? Kill it
   RETURN                                                         '
END SUB

SUB MainCDoKeyString(KeyData AS STRING)
'---------- Handle a key string
LOCAL i, j, DataIX, Done AS LONG, lcmd AS STRING
   DO WHILE LEN(KeyData) > 0                                      '
      SELECT CASE LEFT$(KeyData, 1)                               ' Process the next entry
         RESET gKeyPrimOper                                       ' Start as no primitive operand

         '----- One of the Primitives
         CASE "("                                                 ' Primitive key?
            i = INSTR(KeyData, ")")                               ' Find closing ) character
            IF i = 0 THEN KeyData = "": ITERATE DO                ' None? Ignore this

            '----- If KB recording, do it
            IF gKbdRecFlag THEN                                   ' Doing a recording?
               gKbdRecording += LEFT$(KeyData, i)                 ' Yes, add primitive to the recording
            END IF                                                '

            '----- Process the primitive
            gKeyChr = MID$(KeyData, 2, i - 2)                     ' Extract the primitive key
            KeyData = MID$(KeyData, i + 1)                        ' Shrink KeyData
            i = INSTR(gKeyChr, "/")                               ' A primitive operand?
            IF i THEN                                             ' Yes, lets strip it off
               IF i <> LEN(gKeyChr) THEN                          ' If not in the last position
                  gKeyPrimOper = MID$(gKeyChr, i + 1)             ' Extract the operand
               END IF                                             '
               gKeyChr = LEFT$(gKeyChr, i - 1)                    ' Just strip it off and ignore it
            END IF                                                '
            IF IsEQ(gKeyChr, "NULL") THEN ITERATE DO              ' Shrug off NULL

            i = INSTR(gKeyChr, ":")                               ' A Repeat request character
            IF i THEN j = VAL(LEFT$(gKeyChr, i - 1))              ' A number like 3: ?
            IF j THEN gKeyChr = MID$(gKeyChr, i + 1)              ' Remove the 3: prefix
            IF LEFT$(gKeyChr, 1) <> "[" THEN                      ' If not a character string [abc] type
               gKeyChr = UUCASE(gKeyChr)                          ' Uppercase it
               j = MAX(1, j)                                      ' Do at least 1 time
               DO WHILE j                                         ' Repeat the primitive
                  DECR j                                          '
                  MainDKeyProcess                                 ' Pass to the normal process
                  IF TabsNum < 1 THEN MExitSub                    ' Don't crash on shutdown
               LOOP                                               '
            ELSE                                                  '
               REPLACE "[[" WITH "[" IN gKeyChr                   ' Handle escaped [ ] chars
               REPLACE "]]" WITH "]" IN gKeyChr                   '
               gKeyChr = UNWRAP$(gKeyChr, "[", "]")               ' Extract the text
               gInternalCB = REPEAT$(j, gKeyChr)                  ' Apply repeat factor
               gKeyPrimOper = "|InternalCB"                       ' Point at internal CB
               gKeyChr = "PASTE"                                  ' Make into a PASTE command
               MainDKeyProcess                                    ' Pass to the normal process
            END IF                                                '

         '----- A Line command string
         CASE "{"                                                 ' Line command
            IF IsTPPTypeMode THEN MExitSub                        ' No commands in PowerType mode
            i = INSTR(KeyData, "}")                               ' Find closing } character
            IF i = 0 THEN KeyData = "": ITERATE DO                ' None? Ignore this

            '----- If KB recording, do it
            IF gKbdRecFlag THEN                                   ' Doing a recording?
               gKbdRecording += LEFT$(KeyData, i)                 ' Yes, add line command to the recording
            END IF                                                '

            '----- Extract and insert the command into the line
            lcmd = MID$(KeyData,2, i - 2)                         ' Extract the line command
            KeyData = MID$(KeyData, i + 1)                        ' Shrink KeyData
            IF ISFALSE IsFMTab THEN                               ' If not FM screen
               DataIX = sGetIX(TP.CsrRow)                         ' Get index to the Data Line
               IF DataIX < 0 THEN                                 ' A Hex line
                  i = DataIX                                      ' Save DataIX
                  DataIx = sGetIX(TP.CsrRow - ABS(DataIX))        ' Convert to real line index
                  TP.CsrRow = TP.CsrRow - ABS(i)                  ' Move cursor to real line
               END IF                                             '
               IF DataIX <> 0 THEN                                '
                  IF TP.LLCtlGet(DataIX) = LSET$(lcmd, 8) THEN    ' Is it already the same?
                     TP.LLCtlSet(DataIX, $BlankLNo)               ' Yes, blank it
                     IF TP.LLblGet(DataIX) <> $BlankLNo THEN      ' If a label, put it back
                        TP.LLCtlSet(DataIX, TP.LLblGet(DataIX))   '
                     ELSEIF TP.LTagGet(DataIX) <> $BlankLNo THEN  ' If a tag, put it back
                        TP.LLCtlSet(DataIX, TP.LTagGet(DataIX))   '
                     END IF                                       '
                  ELSE                                            '
                     TP.LLCtlSet(DataIX, LSET$(lcmd, 8))          ' Stuff it all in as a line command
                  END IF                                          '
                  TP.TTblAdd(DataIX, TP.LLCtlGet(DataIX))         ' Add touched items
                  sPrint (LEFT$(TP.LLCtlGet(DataIX), ENV.LinNoSize), $$LnoHi, TP.CsrRow, 1)   ' Print the Line area
                  IF LEN(KeyData) = 0 THEN KeyData = "(Enter)"    ' If no chained, insert an Enter
               END IF                                             '

            END IF                                                '

         '----- A raw text string, 'type' it in
         CASE "["                                                 ' Raw text data
            Done = %False                                         ' Loop breakout
            KeyData = MID$(KeyData, 2)                            ' Step over opening [

            '----- If KB recording, start the raw frame
            IF gKbdRecFlag THEN gKbdRecording += "["              ' Start text frame if recording

            '----- Process the characters one by one
            DO UNTIL Done OR ISNULL(KeyData)                      ' Pass characters onward
               gKeyChr = LEFT$(KeyData, 1)                        ' Extract next character
               KeyData = MID$(KeyData, 2)                         ' Step over extracted character

               IF gKeyChr <> "[" AND gKeyChr <> "]" THEN          ' Non-escapable character?
                  IF gKbdRecFlag THEN gKbdRecording += gKeyChr    ' Add gKeyChr if we're recording
                  MainDKeyProcess                                 ' Pass to the normal process
                  IF TabsNum < 1 THEN MExitSub                    ' Don't crash on shutdown

               ELSE                                               ' It's a [ or ] character
                  IF gKeyChr = "[" THEN                           ' Another opening bracket?
                     IF LEFT$(KeyData, 1) <> "[" THEN             ' Not duplicated
                     IF gKbdRecFlag THEN gKbdRecording += gKeyChr ' Add gKeyChr if we're recording
                        MainDKeyProcess                           ' Pass to the normal process
                        IF TabsNum < 1 THEN MExitSub              ' Don't crash on shutdown
                     ELSE                                         ' Dup'ed
                        IF gKbdRecFlag THEN gKbdRecording += gKeyChr ' Add gKeyChr if we're recording
                        MainDKeyProcess                           ' Pass to the normal process
                        IF TabsNum < 1 THEN MExitSub              ' Don't crash on shutdown
                        KeyData = MID$(KeyData, 2)                ' ignore it
                     END IF                                       '

                  ELSE                                            ' We have a closing bracket
                     IF ISNULL(KeyData) OR gKeyChr <> LEFT$(KeyData, 1) THEN ' It's a proper closing ]
                        Done = %True                              ' We're done
                     ELSE                                         ' Else it's a dup
                        IF gKbdRecFlag THEN gKbdRecording += gKeyChr ' Add gKeyChr if we're recording
                        MainDKeyProcess                           ' Pass to the normal process
                        IF TabsNum < 1 THEN MExitSub              ' Don't crash on shutdown
                        KeyData = MID$(KeyData, 2)                ' ignore one
                     END IF                                       '
                  END IF                                          '
               END IF                                             '
            LOOP                                                  '
            IF gKbdRecFlag THEN gKbdRecording += "]"              ' Close text frame if we're recording

         '----- PFSHOW label preceeding a command
         CASE "<"                                                 ' Labelled Primary Command
            i = INSTR(KeyData, ">")                               ' Look for end of label
            IF i = 0 THEN KeyData = "": ITERATE DO                ' None? Ignore this
            KeyData = MID$(KeyData, i + 1)                        ' Strip off PFSHOW label

         '----- Anything else better be a Primary command
         CASE ELSE                                                ' Primary Command

            '----- If just a single character, treat is as just that
            IF LEN(KeyData) = 1 THEN                              ' Just a single character?
               gKeyChr = KeyData                                  ' Extract it
               IF gKbdRecFlag THEN gKbdRecording += "[" + gKeyChr + "]" ' Add the 1 char string if recording
               MainDKeyProcess                                    ' Pass to the normal process
               IF TabsNum < 1 THEN MExitSub                       ' Don't crash on shutdown
               KeyData = ""                                       '

            '----- Normal command
            ELSE                                                  '
               '----- In PowerType mode reject nearly all commands
               IF IsTPPTypeMode THEN                              ' In Powertype mode
                  IF LEFT$(KeyData, 1) = "!" THEN                 ' See if ! type
                     IF IsNE(TRIM$(MID$(KeyData, 2)), "LEFT") AND _ ' Only LEFT/RIGHT can pass through
                        IsNE(TRIM$(MID$(KeyData, 2)), "RIGHT") THEN MExitSub
                  ELSE                                            '
                     IF IsNE(TRIM$(KeyData), "LEFT") AND _        '
                        IsNE(TRIM$(KeyData), "RIGHT") THEN MExitSub
                  END IF                                          '
               END IF                                             '

               '----- If KB record, do it
               IF gKbdRecFlag THEN gKbdRecording += KeyData       ' Add as a PFK type command if we're recording

               '----- Process the command
               OnTPPFKInsert                                      ' Say pCommand came from a PFK
               IF LEFT$(KeyData, 1) = "!" THEN                    ' Request for no appending?
                  TP.pCommand = MID$(KeyData, 2)                  ' Replace pCommand with everything after the !
                  TP.pCommandRaw = ""                             ' Null the raw version
               ELSE                                               ' Else, normal append
                  TP.pCommandRaw = TP.pCommand                    ' Save the un-prepended version
                  TP.pCommand = KeyData + " " + TP.pCommand       ' Insert it into the command line
               END IF                                             '
               KeyData = "(ENTER)"                                ' Force the (Enter)
            END IF                                                '
      END SELECT                                                  '
   LOOP                                                           '
END SUB

SUB MainDKeyProcess()
LOCAL i, j AS LONG, CallAddr AS DWORD

'---------- Keyboard processor
   MEntry
   gLoopFlag = %True                                              ' Start timing this transaction
   gLoopCtr = 0                                                   ' Reset counter
   gLastKBTime = sOneSecondTimer                                  ' Save time of last KB activity

   '----- Just a single typed character? Pass off quickly
   IF LEN(gKeyChr) = 1 THEN                                       ' If a 'normal character
      IF IsFMTab THEN                                             ' If FM Mode
         FMkbChar                                                 ' Let FM routine have a go
      ELSEIF IsTPPTypeMode THEN                                   ' Maybe PType?
         PTChar                                                   ' Let PType routine have a go
      ELSE                                                        ' Just normal
         krChar                                                   ' Let normal routine have a go
      END IF                                                      '

   '----- We have a Primitive, split off by Primitive name
   ELSE
      i = PrimT.GetPrimIX(gKeyChr)                                ' Search for the Primitive

      IF i THEN                                                   ' We found it
         gCrashLastPrim = gKeyChr                                 ' Last Primitive
         IF Primt.NdMarkd(i) <> 0 THEN                            ' Does command need a marked area?
            IF IsTPMiscActive OR IsTPMarkActive THEN              ' Something marked - OK
               '                                                  ' Do nothing
            ELSEIF TP.CursData THEN                               ' In the data area
               '----- Fudge a 1 char marked area
               krMarkL                                            ' Try to mark 1 char
            ELSE                                                  '
               i = 0                                              ' Kill it
            END IF                                                '
         END IF                                                   '
      END IF                                                      '

      IF i THEN                                                   ' Still have a command?
         IF IsFMTab THEN                                          ' If this is the FM tab
            IF Primt.FPath(i) <> 0 THEN                           ' If allowable in FM Mode
               CallAddr = PrimT.FPath(i)                          ' Get Routine address
               CALL DWORD CallAddr                                ' Call the routine
            END IF                                                '
         ELSEIF IsTPPTypeMode THEN                                ' If this is PType mode
            IF Primt.PPath(i) <> 0 THEN                           ' If allowable in PT Mode
               CallAddr = PrimT.PPath(i)                          ' Get Routine address
               CALL DWORD CallAddr                                ' Call the routine
            END IF                                                '
         ELSE                                                     ' Must be normal mode
            IF IsBrowse AND PrimT.BrowseOK(i) = %False THEN       ' If BROWSE, is it allowable in browse?
               sDoBeep                                            ' No? Then no dispatch
            ELSE                                                  ' Yes? Do the dispatch
               IF Primt.NPath(i) <> 0 THEN                        ' If allowable in normal Mode
                  CallAddr = PrimT.NPath(i)                       ' Get Routine address
                  CALL DWORD CallAddr                             ' Call the routine
               END IF                                             '
            END IF                                                '
         END IF                                                   '
      ELSE                                                        ' ???  just Beep
         sDoBeep                                                  '
      END IF                                                      '
   END IF                                                         '

   IF ISFALSE gMacroMode AND TabsNum > 0 THEN _                   ' No more in macro mode or termination
      TP.PostKeyBoard                                             ' See what, if any, post KB processing
   gLoopFlag = %False                                             ' Stop timing this transaction
   MExit
END SUB
