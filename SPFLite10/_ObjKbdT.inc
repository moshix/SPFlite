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

CLASS cKbdTable

   '----- Keyboard Table Data
   INSTANCE Kbd107()         AS iKbMastEntry                      ' Master key definition table
   INSTANCE KeyCode()        AS STRING                            ' KeyCode for each Kbd107 entry
   INSTANCE Kbd()            AS iKBActive                         ' Final lookup table built from Kbd107()
   INSTANCE KMode()          AS STRING                            ' KMode for each Kbd entry
   INSTANCE KbdPFShow()      AS STRING * 32                       ' PFSHOW data strings
   INSTANCE KbdKPEnter       AS LONG                              ' Kbd Index of the Enter Key
   INSTANCE KbdCount         AS LONG                              ' Kbd number of entries
   INSTANCE KbdPFShowNum     AS LONG                              ' PFSHOW number of entries
   INSTANCE KbdPFShowMax     AS LONG                              ' Max length of PFSHOW entry
   INSTANCE KIO              AS iIO                               ' For our I/O stuff

   CLASS METHOD CREATE()                                          ' Constructor - Initialize Class stuff
      DIM Kbd107(0 TO 107)   AS INSTANCE iKbMastEntry             ' Master physical key table
      DIM KeyCode(0 TO 107)  AS INSTANCE STRING                   ' Keycode for each entry
      DIM Kbd(1 TO 750)      AS INSTANCE iKBActive                ' Final Kbd mapping table
      DIM KMode(1 TO 750)    AS INSTANCE STRING                   ' Final Kbd mapping table
      DIM KbdPFShow(200)     AS INSTANCE STRING * 32              ' PFSHOW strings
      LET KIO = CLASS "cIO"                                       '
      Kbdt.LoadKBTable                                            ' Go Load the tables

   END METHOD                                                     '

   INTERFACE iKbdTable: INHERIT IUNKNOWN                          ' Define the interface

      GProp(KbdKPEnter, LONG)                                     ' Kbd Index of the Enter Key
      GProp(KbdPFShowNum, LONG)                                   ' PFSHOW number of entries
      GProp(KbdPFShowMax, LONG)                                   ' Max length of PFSHOW entry
      PROPERTY GET KbdPFShow(ix AS LONG) AS STRING:      PROPERTY = KbdPFShow(ix):                END PROPERTY
      PROPERTY GET Labl(ix AS LONG) AS STRING:           PROPERTY = Kbd107(ix).Labl:              END PROPERTY
      PROPERTY GET UpKey(ix AS LONG) AS STRING:          PROPERTY = Kbd107(ix).UpKey:             END PROPERTY
      PROPERTY GET NData(ix AS LONG) AS STRING:          PROPERTY = Kbd107(ix).NData:             END PROPERTY
      PROPERTY GET SData(ix AS LONG) AS STRING:          PROPERTY = Kbd107(ix).SData:             END PROPERTY
      PROPERTY GET CData(ix AS LONG) AS STRING:          PROPERTY = Kbd107(ix).CData:             END PROPERTY
      PROPERTY GET AData(ix AS LONG) AS STRING:          PROPERTY = Kbd107(ix).AData:             END PROPERTY
      PROPERTY GET SAData(ix AS LONG) AS STRING:         PROPERTY = Kbd107(ix).SAData:            END PROPERTY
      PROPERTY GET SCData(ix AS LONG) AS STRING:         PROPERTY = Kbd107(ix).SCData:            END PROPERTY
      PROPERTY GET CAData(ix AS LONG) AS STRING:         PROPERTY = Kbd107(ix).CAData:            END PROPERTY
      PROPERTY GET SCAData(ix AS LONG) AS STRING:        PROPERTY = Kbd107(ix).SCAData:           END PROPERTY
      PROPERTY GET NAllow(ix AS LONG) AS STRING:         PROPERTY = Kbd107(ix).NAllow:            END PROPERTY
      PROPERTY GET SAllow(ix AS LONG) AS STRING:         PROPERTY = Kbd107(ix).SAllow:            END PROPERTY
      PROPERTY GET CAllow(ix AS LONG) AS STRING:         PROPERTY = Kbd107(ix).CAllow:            END PROPERTY
      PROPERTY GET AAllow(ix AS LONG) AS STRING:         PROPERTY = Kbd107(ix).AAllow:            END PROPERTY
      PROPERTY GET NUp(ix AS LONG) AS STRING:            PROPERTY = Kbd107(ix).NUp:               END PROPERTY
      PROPERTY GET SUp(ix AS LONG) AS STRING:            PROPERTY = Kbd107(ix).SUp:               END PROPERTY
      PROPERTY GET CUp(ix AS LONG) AS STRING:            PROPERTY = Kbd107(ix).CUp:               END PROPERTY
      PROPERTY GET AUp(ix AS LONG) AS STRING:            PROPERTY = Kbd107(ix).AUp:               END PROPERTY
      PROPERTY GET SAUp(ix AS LONG) AS STRING:           PROPERTY = Kbd107(ix).SAUp:              END PROPERTY
      PROPERTY GET SCUp(ix AS LONG) AS STRING:           PROPERTY = Kbd107(ix).SCUp:              END PROPERTY
      PROPERTY GET CAUp(ix AS LONG) AS STRING:           PROPERTY = Kbd107(ix).CAUp:              END PROPERTY
      PROPERTY GET SCAUp(ix AS LONG) AS STRING:          PROPERTY = Kbd107(ix).SCAUp:             END PROPERTY
      PROPERTY GET ForceUp(ix AS LONG) AS STRING:        PROPERTY = Kbd107(ix).ForceUp:           END PROPERTY

      METHOD SupSent(ix AS LONG) AS STRING:              METHOD   = Kbd(ix).SupSent:              END METHOD
      METHOD SetSupSent(ix AS LONG, v AS STRING):        Kbd(ix).SupSent = v:                     END METHOD
      METHOD SupRepeat(ix AS LONG) AS STRING:            METHOD   = Kbd(ix).SupRepeat:            END METHOD
      METHOD SetSupRepeat(ix AS LONG, v AS STRING):      Kbd(ix).SupRepeat = v:                   END METHOD
      METHOD kData(ix AS LONG) AS STRING:                METHOD = Kbd(ix).KData:                  END METHOD

      METHOD SCAN(Lookup AS STRING) AS LONG
      '----- Look for the key/mode combo
      REGISTER i AS LONG
         ARRAY SCAN KMode() FOR KbdCount, =Lookup, TO i
         METHOD = i                                               ' Return ARRAY SCAN result
      END METHOD

      METHOD ScanFull(Lookup AS STRING) AS LONG
         REGISTER i AS LONG
         LOCAL k AS STRING
         k = LEFT$(Lookup, 1)                                     ' Get 1st byte
         ARRAY SCAN KeyCode(), =k, TO i                           ' Look for KeyCode
         METHOD = i                                               ' Return Array scan result
      END METHOD

      METHOD LoadKbTable()
      LOCAL fn, kIX, aIX, phase, kbDIX, hashok, i, j, RWHash AS LONG, hash, nhash AS DWORD, t, TX, TX2, kbData, pfk, pfpfx AS STRING
      DIM args(10) AS STRING
         '----- Build dummy entry as Key 0 (Used as a PASSTHRU for extra keys on foreign keyboards)
         IF ISNOTHING(Kbd107(0)) THEN LET Kbd107(0) = CLASS "cKBMastEntry" ' Build a dummy key entry
         KeyCode(0) = CHR$(0)                                     '
         Kbd107(0).Extern = " "                                   '
         Kbd107(0).KIcon = " "                                    '
         Kbd107(0).W = 0                                          '
         Kbd107(0).H = 0                                          '
         Kbd107(0).X = 0                                          '
         Kbd107(0).Y = 0                                          '
         Kbd107(0).NAllow = "N"                                   '
         Kbd107(0).SAllow = "S"                                   '
         Kbd107(0).CAllow = "C"                                   '
         Kbd107(0).AAllow = "A"                                   '
         Kbd107(0).UpKey = " "                                    '
         Kbd107(0).ForceUp = " "                                  '
         Kbd107(0).Labl = " "                                     '
         Kbd107(0).NData = "(PASSTHRU)"                           '
         Kbd107(0).NUp = " "                                      '
         Kbd107(0).SData = "(PASSTHRU)"                           '
         Kbd107(0).SUp = " "                                      '
         Kbd107(0).CData = "(PASSTHRU)"                           '
         Kbd107(0).CUp = " "                                      '
         Kbd107(0).AData = "(PASSTHRU)"                           '
         Kbd107(0).AUp = " "                                      '
         Kbd107(0).SCData = "(PASSTHRU)"                          '
         Kbd107(0).SCUp = " "                                     '
         Kbd107(0).SAData = "(PASSTHRU)"                          '
         Kbd107(0).SAUp = " "                                     '
         Kbd107(0).SCAData = "(PASSTHRU)"                         '
         Kbd107(0).SCAUp = " "                                    '
         Kbd107(0).CAData = "(PASSTHRU)"                          '
         Kbd107(0).CAUp = " "                                     '

         '----- Get our Master table from the EXE resource data
         kbData = RESOURCE$(RCDATA, "KBDATA")                     ' Get kbData from our RES file default
         j = PARSECOUNT(kbData, $CRLF)                            ' Loop through the data finally
         FOR kbDix = 1 TO j                                       '
            TX = PARSE$(kbData, $CRLF, kbDix)                     ' Extract a line

            '----- Comment lines
            IF LEFT$(TX, 1) = ";" OR ISNULL(TRIM$(TX)) THEN       ' Skip comments, reset phase
               phase = 1                                          '
               ITERATE FOR                                        '
            END IF                                                '
            SELECT CASE AS LONG phase                             ' Which line of Set?

               '----- First line of a set
               CASE 1                                             ' First
                  PARSE TX, Args()                                ' Break out the 'words'
                  INCR kIX                                        ' Bump Table index
                  IF ISNOTHING(Kbd107(kIX)) THEN  LET Kbd107(kIX) = CLASS "cKBMastEntry" ' Assign an object
                  KeyCode(kIX) = CHR$(VAL(Args(0)))               ' Save the parsed data
                  Kbd107(kIX).Extern = LEFT$(Args(1), 1)          '
                  Kbd107(kIX).KIcon = TRIM$(Args(2))              '
                  Kbd107(kIX).W = VAL(Args(3))                    '
                  Kbd107(kIX).H = VAL(Args(4))                    '
                  Kbd107(kIX).X = VAL(Args(5))                    '
                  Kbd107(kIX).Y = VAL(Args(6))                    '
                  Kbd107(kIX).NAllow = MID$(Args(7), 1, 1)        '
                  Kbd107(kIX).SAllow = MID$(Args(7), 2, 1)        '
                  Kbd107(kIX).CAllow = MID$(Args(7), 3, 1)        '
                  Kbd107(kIX).AAllow = MID$(Args(7), 4, 1)        '
                  Kbd107(kIX).UpKey = MID$(Args(7), 5, 1)         '
                  Kbd107(kIX).ForceUp = MID$(Args(7), 6, 1)       '
                  Kbd107(kIX).Labl = Args(8)                      '
                  INCR phase                                      ' On to next phase

               '----- Second line of a set (Normal Key)
               CASE 2                                             ' Normal command
                  Kbd107(kIX).NData = MID$(TX, 3)                 ' Save key data
                  Kbd107(kIX).NUp = LEFT$(TX, 1)                  '
                  INCR phase                                      ' Next phase

               '----- Third line of a set (Shift Key)
               CASE 3                                             ' Normal command
                  Kbd107(kIX).SData = MID$(TX, 3)                 ' Save key data
                  Kbd107(kIX).SUp = LEFT$(TX,1)                   '
                  INCR phase                                      ' Next phase

               '----- Fourth line of a set (Control Key)
               CASE 4                                             ' Normal command
                  Kbd107(kIX).CData = MID$(TX, 3)                 ' Save key data
                  Kbd107(kIX).CUp = LEFT$(TX, 1)                  '
                  INCR phase                                      ' Next phase

               '----- Fifth line of a set (ALT Key)
               CASE 5                                             ' Normal command
                  Kbd107(kIX).AData = MID$(TX, 3)                 ' Save key data
                  Kbd107(kIX).AUp = LEFT$(TX, 1)                  '
                  INCR phase                                      ' Next phase

               '----- Sixth line of a set (Shift-Control Key)
               CASE 6                                             ' Normal command
                  Kbd107(kIX).SCData = MID$(TX, 3)                ' Save key data
                  Kbd107(kIX).SCUp = LEFT$(TX, 1)                 '
                  INCR phase                                      ' Next phase

               '----- Seventh line of a set (Shift-ALT Key)
               CASE 7                                             ' Normal command
                  Kbd107(kIX).SAData = MID$(TX, 3)                ' Save key data
                  Kbd107(kIX).SAUp = LEFT$(TX, 1)                 '
                  INCR phase                                      ' Next phase

               '----- Eighth line of a set (Shift-Control-ALT Key)
               CASE 8                                             ' Normal command
                  Kbd107(kIX).SCAData = MID$(TX, 3)               ' Save Key Data
                  Kbd107(kIX).SCAUp = LEFT$(TX, 1)                '
                  INCR phase                                      ' Next phase

               '----- Ninth line of a set (Control-ALT Key)
               CASE 9                                             ' Normal command
                  Kbd107(kIX).CAData = MID$(TX, 3)                ' Save key data
                  Kbd107(kIX).CAUp = LEFT$(TX, 1)                 '
                  INCR phase                                      ' Next phase
            END SELECT                                            '
         NEXT kbDIX                                               ' Loop through kbData

         '----- If this is the very initial run, set the user's choice for Enter/NewLine
         IF ENV.WelcomeOpt > 0 THEN                               ' Is this the initial setup?
            IF ENV.WelcomeOpt = 1 THEN                            ' Enter and Backspace
               Kbd107(44).NData = "(Enter)"                       '
               Kbd107(30).NData = "(NewLine)"                     '
            ELSEIF ENV.WelcomeOpt = 2 THEN                        ' Enter and Ctrl-Enter
               Kbd107(44).NData = "(Enter)"                       '
               Kbd107(44).NData = "(NewLine)"                     '
            ELSEIF ENV.WelcomeOpt = 3 THEN                        ' R-Ctrl and Enter
               Kbd107(77).NData = "(Enter)"                       '
               Kbd107(44).NData = "(NewLine)"                     '
            END IF                                                '
            KbdT.WriteKBUser                                      ' Save things
         END IF                                                   '

         '----- If a user file, load it in

         IF ISFILE(ENV.INIPath + "SPFLite.Kbd") THEN              ' See if Customized file exists, if not, Exit

            '----- First try it using the new hash routine

            '----- Get the user's custom key defs, hash for safety
            KIO.Setup("IE", "", "", ENV.INIPath + "SPFLite.Kbd")  ' Tell IO what we're opening
            KIO.EXEC                                              '
            kbData = "": i = 0                                    ' Null kbData to start
            sFCS32Init(nHash)                                     ' Initialize the Hash count
            DO WHILE ISFALSE EOF(KIO.FNum)                        ' Read the data
               LINE INPUT # KIO.FNum, tx                          ' Get a line
               INCR i                                             ' Count lines
               IF i = 1071 THEN EXIT DO                           ' This should be the Hash line
               nHash = sFCS32Update(nHash, BYVAL STRPTR(tx), LEN(tx)) ' Accum the hash
               kbData = kbData + TX + $CRLF                       '
            LOOP                                                  '
            KIO.Close                                             ' Close file
            sFCS32Final(nHash)                                    ' Finalize Hash
            TX2 = HEX$(nhash, 8)
            IF TX2 = TX THEN hashok = %True:                      ' OK? Say so.

            '----- If hash OK, skip trying it the old way
            IF ISFALSE hashok THEN                                ' OK?

               '----- Get the user's custom key defs, hash for safety
               KIO.Setup("IE", "", "", ENV.INIPath + "SPFLite.Kbd")  ' Tell IO what we're opening
               KIO.EXEC                                           '
               kbData = "": i = 0                                 ' Null kbData to start
               DO WHILE ISFALSE EOF(KIO.FNum)                     ' Read the data
                  LINE INPUT # KIO.FNum, tx                       ' Get a line
                  INCR i                                          ' Count lines
                  IF i = 1071 THEN                                ' This should be the Hash line
                     IF HEX$(hash, 8) = TX THEN hashok = %True: EXIT DO ' OK? Say so.
                  END IF                                          '
                  hash = sHash(TX, hash)                          ' Simple Hash
                  kbData = kbData + TX + $CRLF                    '
               LOOP                                               '
               KIO.Close                                          ' Close file

               '----- Ignore if hash fails
               IF ISFALSE hashok THEN                             ' OK?
                  sDoMsgBox "User Keyboard file (|K" + ENV.INIPath + "SPFLite.Kbd" + "|B) is invalid, ignoring it", _
                            %MB_USERICON OR %MB_OK, "SPFLite"     '
                  GOTO SkipUser                                   '
               END IF
               RWHash = %True                                     ' Force re-write of the KBD file
            END IF                                                '

            kbData = LEFT$(kbData, LEN(kbData) - 1)               '

            '----- OK, override the Master table with the user's values
            kbData = LEFT$(kbData, LEN(kbData) - 1)               ' Drop last $CRLF
            j = 1                                                 ' Let j point at Kbd107()
            DO WHILE ISNOTNULL(kbdata)
               tx = EXTRACT$(kbdata, $CRLF): kbdata = REMAIN$(kbdata, $CRLF) ' Get ; comment line
               tx = EXTRACT$(kbdata, $CRLF): kbdata = REMAIN$(kbdata, $CRLF) ' Get Upkey data
               Kbd107(j).UpKey = tx                               '
               tx = EXTRACT$(kbdata, $CRLF): kbdata = REMAIN$(kbdata, $CRLF) ' Get NData
               Kbd107(j).NUp = LEFT$(TX, 1)                       '
               Kbd107(j).NData = MID$(TX, 3)                      '
               tx = EXTRACT$(kbdata, $CRLF): kbdata = REMAIN$(kbdata, $CRLF) ' Get SData
               Kbd107(j).SUp = LEFT$(TX, 1)                       '
               Kbd107(j).SData = MID$(TX, 3)                      '
               tx = EXTRACT$(kbdata, $CRLF): kbdata = REMAIN$(kbdata, $CRLF) ' Get CData
               Kbd107(j).CUp = LEFT$(TX, 1)                       '
               Kbd107(j).CData = MID$(TX, 3)                      '
               tx = EXTRACT$(kbdata, $CRLF): kbdata = REMAIN$(kbdata, $CRLF) ' Get AData
               Kbd107(j).AUp = LEFT$(TX, 1)                       '
               Kbd107(j).AData = MID$(TX, 3)                      '
               tx = EXTRACT$(kbdata, $CRLF): kbdata = REMAIN$(kbdata, $CRLF) ' Get SCData
               Kbd107(j).SCUp = LEFT$(TX, 1)                      '
               Kbd107(j).SCData = MID$(TX, 3)                     '
               tx = EXTRACT$(kbdata, $CRLF): kbdata = REMAIN$(kbdata, $CRLF) ' Get SAData
               Kbd107(j).SAUp = LEFT$(TX, 1)                      '
               Kbd107(j).SAData = MID$(TX, 3)                     '
               tx = EXTRACT$(kbdata, $CRLF): kbdata = REMAIN$(kbdata, $CRLF) ' Get SCAData
               Kbd107(j).SCAUp = LEFT$(TX, 1)                     '
               Kbd107(j).SCAData = MID$(TX, 3)                    '
               tx = EXTRACT$(kbdata, $CRLF): kbdata = REMAIN$(kbdata, $CRLF) ' Get CAData
               Kbd107(j).CAUp = LEFT$(TX, 1)                      '
               Kbd107(j).CAData = MID$(TX, 3)                     '
               INCR j                                             '
            LOOP                                                  '
         END IF                                                   '

         SkipUser:
         '----- Build the final active table
         KbdPFShowNum = 0                                         '
         KbdPFShowMax = 0                                         '
         aIX = 0                                                  '
         FOR kIX = 0 TO 107                                       ' Loop through the master table
            IF Kbd107(kIX).ForceUp = "D" THEN                     ' Untouchable key?
               ITERATE FOR                                        '
            END IF                                                '

            IF Kbd107(kIX).NAllow = "N" AND _                     ' Do normal (if not NULL)
               UUCASE(TRIM$(Kbd107(kIX).NData)) <> "(NULL)" THEN  '
               INCR aIX                                           ' Incr active table index
               IF ISNOTHING(Kbd(aIX)) THEN LET Kbd(aIX) = CLASS "cKBActive" ' Assign an object
               KMode(aIX) = KeyCode(kIX) + Kbd107(kIX).Extern + "..."
               Kbd(aIX).SupSent = "N"                             '       1st Sent
               Kbd(aIX).SupRepeat = IIF$(Kbd107(kIX).ForceUp = "U", "U", Kbd107(kIX).NUp)
               Kbd(aIX).KData = Kbd107(kIX).NData                 '   "   Data
               pfpfx = "1": GOSUB AddPFShow                       ' Set prefix and go add PFShow data
               IF KeyCode(kIX) = CHR$(%VK_RETURN) THEN KbdKPEnter = aIX ' Save index of Enter key
            END IF                                                '

            IF Kbd107(kIX).SAllow = "S" AND _                     ' Do Shifted (if not NULL)
                UUCASE(TRIM$(Kbd107(kIX).SData)) <> "(NULL)" THEN '
               INCR aIX                                           ' Incr active table index
               IF ISNOTHING(Kbd(aIX)) THEN LET Kbd(aIX) = CLASS "cKBActive" ' Assign an object
               KMode(aIX) = KeyCode(kIX) + Kbd107(kIX).Extern + "S.."
               Kbd(aIX).SupSent = "N"                             '       1st Sent
               Kbd(aIX).SupRepeat = IIF$(Kbd107(kIX).ForceUp = "U", "U", Kbd107(kIX).SUp)
               Kbd(aIX).KData = Kbd107(kIX).SData                 '   "   Data
               pfpfx = "2S-": GOSUB AddPFShow                     ' Set prefix and go add PFShow data
            END IF                                                '
                                                                  '
            IF Kbd107(kIX).CAllow = "C" AND _                     ' Do Control (if not NULL)
                UUCASE(TRIM$(Kbd107(kIX).CData)) <> "(NULL)" THEN '
               INCR aIX                                           ' Incr active table index
               IF ISNOTHING(Kbd(aIX)) THEN LET Kbd(aIX) = CLASS "cKBActive" ' Assign an object
               KMode(aIX) = KeyCode(kIX) + Kbd107(kIX).Extern + ".C."
               Kbd(aIX).SupSent = "N"                             '       1st Sent
               Kbd(aIX).SupRepeat = IIF$(Kbd107(kIX).ForceUp = "U", "U", Kbd107(kIX).CUp)
               Kbd(aIX).KData = Kbd107(kIX).CData                 '   "   Data
               pfpfx = "3C-": GOSUB AddPFShow                     ' Set prefix and go add PFShow data
            END IF                                                '

            IF Kbd107(kIX).AAllow = "A" AND _                     ' Do Alt (if not NULL)
                UUCASE(TRIM$(Kbd107(kIX).AData)) <> "(NULL)" THEN '
               INCR aIX                                           ' Incr active table index
               IF ISNOTHING(Kbd(aIX)) THEN LET Kbd(aIX) = CLASS "cKBActive" ' Assign an object
               KMode(aIX) = KeyCode(kIX) + Kbd107(kIX).Extern + "..A"
               Kbd(aIX).SupSent = "N"                             '       1st Sent
               Kbd(aIX).SupRepeat = IIF$(Kbd107(kIX).ForceUp = "U", "U", Kbd107(kIX).AUp)
               Kbd(aIX).KData = Kbd107(kIX).AData                 '   "   Data
               pfpfx = "4A-": GOSUB AddPFShow                     ' Set prefix and go add PFShow data
            END IF                                                '

            IF Kbd107(kIX).SAllow = "S" AND Kbd107(kIX).CAllow = "C" AND _ ' Do Shift-Control (if not NULL)
                UUCASE(TRIM$(Kbd107(kIX).SCData)) <> "(NULL)" THEN'
               INCR aIX                                           ' Incr active table index
               IF ISNOTHING(Kbd(aIX)) THEN LET Kbd(aIX) = CLASS "cKBActive" ' Assign an object
               KMode(aIX) = KeyCode(kIX) + Kbd107(kIX).Extern + "SC."
               Kbd(aIX).SupSent = "N"                             '       1st Sent
               Kbd(aIX).SupRepeat = IIF$(Kbd107(kIX).ForceUp = "U", "U", Kbd107(kIX).SCUp)
               Kbd(aIX).KData = Kbd107(kIX).SCData                '   "   Data
               pfpfx = "5SC-": GOSUB AddPFShow                    ' Set prefix and go add PFShow data

            END IF                                                '

            IF Kbd107(kIX).SAllow = "S" AND Kbd107(kIX).AAllow = "A" AND _ ' Do Shift-Alt (if not NULL)
                UUCASE(TRIM$(Kbd107(kIX).SAData)) <> "(NULL)" THEN'
               INCR aIX                                           ' Incr active table index
               IF ISNOTHING(Kbd(aIX)) THEN LET Kbd(aIX) = CLASS "cKBActive" ' Assign an object
               KMode(aIX) = KeyCode(kIX) + Kbd107(kIX).Extern + "S.A"
               Kbd(aIX).SupSent = "N"                             '       1st Sent
               Kbd(aIX).SupRepeat = IIF$(Kbd107(kIX).ForceUp = "U", "U", Kbd107(kIX).SAUp)
               Kbd(aIX).KData = Kbd107(kIX).SAData                '   "   Data
               pfpfx = "7SA-": GOSUB AddPFShow                    ' Set prefix and go add PFShow data
            END IF                                                '

            IF Kbd107(kIX).SAllow = "S" AND Kbd107(kIX).CAllow = "C" AND Kbd107(kIX).AAllow = "A" AND _ ' Do Shift-Control-Alt (if not NULL)
                UUCASE(TRIM$(Kbd107(kIX).SCAData)) <> "(NULL)" THEN
               INCR aIX                                           ' Incr active table index
               IF ISNOTHING(Kbd(aIX)) THEN LET Kbd(aIX) = CLASS "cKBActive" ' Assign an object
               KMode(aIX) = KeyCode(kIX) + Kbd107(kIX).Extern + "SCA"
               Kbd(aIX).SupSent = "N"                             '       1st Sent
               Kbd(aIX).SupRepeat = IIF$(Kbd107(kIX).ForceUp = "U", "U", Kbd107(kIX).SCAUp)
               Kbd(aIX).KData = Kbd107(kIX).SCAData               '   "   Data
               pfpfx = "8SCA-": GOSUB AddPFShow                   ' Set prefix and go add PFShow data
            END IF                                                '

            IF Kbd107(kIX).CAllow = "C" AND Kbd107(kIX).AAllow = "A" AND _ ' Do Control-Alt (if not NULL)
                UUCASE(TRIM$(Kbd107(kIX).CAData)) <> "(NULL)" THEN'
               INCR aIX                                           ' Incr active table index
               IF ISNOTHING(Kbd(aIX)) THEN LET Kbd(aIX) = CLASS "cKBActive" ' Assign an object
               KMode(aIX) = KeyCode(kIX) + Kbd107(kIX).Extern + ".CA"
               Kbd(aIX).SupSent = "N"                             '       1st Sent
               Kbd(aIX).SupRepeat = IIF$(Kbd107(kIX).ForceUp = "U", "U", Kbd107(kIX).CAUp)
               Kbd(aIX).KData = Kbd107(kIX).CAData                '   "   Data
               pfpfx = "6CA-": GOSUB AddPFShow                    ' Set prefix and go add PFShow data
            END IF                                                '

         NEXT kIX                                                 '
         KbdCount = aIX + 1                                       ' Save size of table
         ARRAY SORT KbdPFShow() FOR (KbdPFShowNum + 1), CALL sKbdSortEx() ' Sort PFShow able
         KbdPFShowMax = MAX(8, KbdPFShowMax)                      ' Just to avoid div by zero
         KbdPFShowMax = MIN(20, KbdPFShowMax)                     ' To avoid being too big
         IF RWHash THEN me.WriteKBUser                            ' Force rewrite if this is the hash conversion run
         EXIT METHOD                                              '

      AddPFShow:
         IF LEFT$(Kbd(aIX).KData, 1) = "<" THEN                   ' Possibly labelled?
            i = INSTR(Kbd(aIX).KData, ">")                        ' Find the end of the comment
            IF i THEN                                             ' Got it
               INCR KbdPFShowNum                                  ' Incr count
               pfk = sAdd128(pfpfx + Kbd107(kIX).Labl)            ' Get an upper128 version of the text
               KbdPFShow(KbdPFShowNum) = pfk + "=" + MID$(Kbd(aIX).KData, 2, i - 2)
               KbdPFShowMax = MAX(KbdPFShowMax, LEN(KbdPFShow(KbdPFShowNum))) 'Save MAX length
            END IF                                                '
         END IF                                                   '
         RETURN

      END METHOD

      METHOD WriteKBUser()
      '---------- Save the gKbd107 table
      LOCAL TX AS STRING, i, hash AS DWORD
         MEntry
         '----- Write the users custom KB settings out
         KIO.Setup("O", "", "", ENV.INIPath + "SPFLite.Kbd")      ' Tell KIO what er're doing
         KIO.EXEC                                                 ' Open it
         sFCS32Init(Hash)                                         ' Initialize the Hash count
         FOR i = 1 TO 107                                         ' Dump the Table
            TX = "; ------------- " + Kbd107(i).Labl              ' Delimiter line
            GOSUB HashPrint                                       ' Write it
            TX = Kbd107(i).UpKey                                  ' Repeat
            GOSUB HashPrint                                       ' Write it
            TX = Kbd107(i).NUp + "/" + Kbd107(i).NData            ' Insert the data
            GOSUB HashPrint                                       '
            TX = Kbd107(i).SUp + "/" + Kbd107(i).SData            '
            GOSUB HashPrint                                       '
            TX = Kbd107(i).CUp + "/" + Kbd107(i).CData            '
            GOSUB HashPrint                                       '
            TX = Kbd107(i).AUp + "/" + Kbd107(i).AData            '
            GOSUB HashPrint                                       '
            TX = Kbd107(i).SCUp + "/" + Kbd107(i).SCData          '
            GOSUB HashPrint                                       '
            TX = Kbd107(i).SAUp + "/" + Kbd107(i).SAData          '
            GOSUB HashPrint                                       '
            TX = Kbd107(i).SCAUp + "/" + Kbd107(i).SCAData        '
            GOSUB HashPrint                                       '
            TX = Kbd107(i).CAUp + "/" + Kbd107(i).CAData          '
            GOSUB HashPrint                                       '
         NEXT i                                                   '
         sFCS32Final(Hash)                                        ' Finalize Hash
         TX = HEX$(hash, 8)                                       ' Create hash line
         PRINT # KIO.FNum, TX                                     ' Print the hash
         KIO.Close                                                ' Close file
         MExitMeth

      HashPrint:
         Hash = sFCS32Update(Hash, BYVAL STRPTR(tx), LEN(tx))     ' Accum the hash
         PRINT # KIO.FNum, TX                                     ' Print it
         RETURN                                                   '

      END METHOD

      METHOD DispKeyMap(MsgMode AS LONG)
      '---------- Build and display the Keyboard Mapper
      LOCAL i, lclRepeat AS LONG, NormalText, ShiftText, ControlText, AltText, SCText, SAText, SCAText, CAText AS STRING
         MEntry
         '----- Build the screen
         DIALOG FONT DEFAULT "Tahoma", 10 / gFontScale,0, 0
         DIALOG NEW PIXELS, hWnd, "SPFLite KeyMap", , , 755, 538, %WS_POPUP OR %WS_BORDER OR %WS_DLGFRAME OR %WS_CAPTION OR _
                                  %WS_CLIPSIBLINGS OR %WS_VISIBLE OR %DS_MODALFRAME OR %DS_3DLOOK OR %DS_NOFAILCREATE OR _
                                  %DS_SETFONT, %WS_EX_CONTROLPARENT OR %WS_EX_LEFT OR %WS_EX_LTRREADING OR %WS_EX_RIGHTSCROLLBAR, TO hKey
         CONTROL ADD FRAME, hKey, %KEYMAP_KBFRAME, "", 0, 0, 753, 537, %BS_GROUPBOX, %WS_EX_WINDOWEDGE
         DIALOG SET COLOR hKey, %RGB_GAINSBORO, %RGB_GAINSBORO

         '----- Add the bottom DONE and CANCEL buttons
         CONTROL ADD BUTTON,  hKey, %KEYMAP_DONE, "Done", 487, 507, 75, 24, %WS_BORDER
         sToolTipSet (GetDlgItem(hKey, %KEYMAP_DONE), " Press Done once you have made all desiLAttrG changes. ")

         CONTROL ADD BUTTON,  hKey, %KEYMAP_CANCEL, "Cancel", 387, 507, 75, 24, %WS_BORDER
         sToolTipSet (GetDlgItem(hKey, %KEYMAP_CANCEL), " Press Cancel to Discard all changes. ")

         '----- Add the graphic keyboard map. X/Y coordinates are stoLAttrG in the Keyboard Master Table
         FOR i = 1 TO 107                                               ' Add the Key Images
            CONTROL ADD IMAGEX, hKey, INT(5000 + i), TRIM$(Kbd107(i).KIcon), Kbd107(i).X, Kbd107(i).Y, _
                                      Kbd107(i).W, Kbd107(i).H, %SS_ICON OR %SS_CENTERIMAGE OR %SS_NOTIFY', CALL DlgKeyMapCallback
            CONTROL SET USER hKey, INT(5000 + i), 1, i

         NEXT i                                                         '
         CONTROL ADD LINE, hKey, INT(5000 + i), "", 1, 225, 751, 3      '

         '----- Add some various text messages
         IF MsgMode = 0 THEN                                            ' Which prompt do we issue
            CONTROL ADD LABEL, hKey, %KEYMAP_PROMPT, "Select a key above to be customized", 3, 230, 450, 16
         ELSE                                                           '
            CONTROL ADD LABEL, hKey, %KEYMAP_PROMPT, "Select a key to receive the recording", 3, 230, 450, 16
         END IF                                                         '
         CONTROL SET COLOR  hKey, %KEYMAP_PROMPT, %RED, %RGB_GAINSBORO

         CONTROL ADD LABEL, hKey, %KEYMAP_CURRENT_TEXT, "Current Key - None", 3, 270, 165, 16
         CONTROL SET COLOR  hKey, %KEYMAP_CURRENT_TEXT, %BLACK, %RGB_GAINSBORO

         CONTROL ADD LABEL, hKey, %KEYMAP_HINT_1, "Hints", 500, 230, 45, 16
         CONTROL SET FONT   hKey, %KEYMAP_HINT_1, hFixedFont
         CONTROL SET COLOR  hKey, %KEYMAP_HINT_1, %RGB_DARKGREEN, -2
         CONTROL ADD LABEL, hKey, %KEYMAP_HINT_2, "Un-Bracketed - Primary Command", 500, 250, 275, 16
         CONTROL SET FONT   hKey, %KEYMAP_HINT_2, hFixedFont
         CONTROL SET COLOR  hKey, %KEYMAP_HINT_2, %RGB_MEDIUMSEAGREEN, -2
         CONTROL ADD LABEL, hKey, %KEYMAP_HINT_3, "{----} - Line Command", 500, 270, 275, 16
         CONTROL SET FONT   hKey, %KEYMAP_HINT_3, hFixedFont
         CONTROL SET COLOR  hKey, %KEYMAP_HINT_3, %RGB_MEDIUMSEAGREEN, -2
         CONTROL ADD LABEL, hKey, %KEYMAP_HINT_4, "(----) - Keyboard Primitives", 500, 290, 275, 16
         CONTROL SET FONT   hKey, %KEYMAP_HINT_4, hFixedFont
         CONTROL SET COLOR  hKey, %KEYMAP_HINT_4, %RGB_MEDIUMSEAGREEN, -2
         CONTROL ADD LABEL, hKey, %KEYMAP_HINT_5, "[----] - Text Strings", 500, 310, 275, 16
         CONTROL SET FONT   hKey, %KEYMAP_HINT_5, hFixedFont
         CONTROL SET COLOR  hKey, %KEYMAP_HINT_5, %RGB_MEDIUMSEAGREEN, -2
         CONTROL ADD LABEL, hKey, %KEYMAP_HINT_6, "<----> - PFSHOW Labels", 500, 330, 275, 16
         CONTROL SET FONT   hKey, %KEYMAP_HINT_6, hFixedFont
         CONTROL SET COLOR  hKey, %KEYMAP_HINT_6, %RGB_MEDIUMSEAGREEN, -2
         CONTROL ADD LABEL, hKey, %KEYMAP_HINT_7, "Select hint below to place in Clipboard", 500, 360, 275, 16
         CONTROL SET COLOR  hKey, %KEYMAP_HINT_7, %RGB_DARKGREEN, -2

         '----- Add the Primitive list and CharSet button
         CONTROL ADD COMBOBOX, hKey, %KEYMAP_PRIMITIVEHINT, gPrimTable(), 500, 385, 250, 160, %CBS_DROPDOWNLIST OR %WS_VSCROLL
         COMBOBOX SELECT    hKey, %KEYMAP_PRIMITIVEHINT, 1
         sToolTipSet (GetDlgItem(hKey, %KEYMAP_PRIMITIVEHINT), " List of valid Keyboard Primitives. ")

         CONTROL ADD BUTTON,  hKey, %KEYMAP_ANSI, "Show Character map", 500, 420, 250, 24, %WS_BORDER
         sToolTipSet (GetDlgItem(hKey, %KEYMAP_ANSI), " Click to display a full character map popup. ")


         '----- Add the 8 boxes for the various key modes
         CONTROL ADD LABEL, hKey, %KEYMAP_NORMAL_TEXT, "Normal", 5, 315, 100, 16
         CONTROL SET COLOR  hKey, %KEYMAP_NORMAL_TEXT, %BLACK, -2
         CONTROL DISABLE    hKey, %KEYMAP_NORMAL_TEXT
         CONTROL ADD TEXTBOX, hKey, %KEYMAP_NORMAL, NormalText, 125, 315, 335, 20

         CONTROL SET FONT   hKey, %KEYMAP_NORMAL, hScrFont
         CONTROL DISABLE    hKey, %KEYMAP_NORMAL
         sToolTipSet (GetDlgItem(hKey, %KEYMAP_NORMAL), " Enter the keyboard command string for the Normal Key Mode. ")

         CONTROL ADD CHECKBOX, hKey, %KEYMAP_NUP, " ", 100, 318, 16, 16, %BS_LEFT OR %BS_VCENTER
         CONTROL SET CHECK     hKey, %KEYMAP_NUP, 0
         sToolTipSet (GetDlgItem(hKey, %KEYMAP_NUP), " Should keyboard AutoRepeat be allowed for this key? ")
         CONTROL DISABLE hKey, %KEYMAP_NUP

         CONTROL ADD LABEL, hKey, %KEYMAP_SHIFT_TEXT, "Shift", 5, 338, 100, 16
         CONTROL SET COLOR  hKey, %KEYMAP_SHIFT_TEXT, %BLACK, -2
         CONTROL DISABLE    hKey, %KEYMAP_SHIFT_TEXT
         CONTROL ADD TEXTBOX, hKey, %KEYMAP_SHIFT, ShiftText, 125, 338, 335, 20
         CONTROL SET FONT   hKey, %KEYMAP_SHIFT, hScrFont
         CONTROL DISABLE    hKey, %KEYMAP_SHIFT
         sToolTipSet (GetDlgItem(hKey, %KEYMAP_SHIFT), " Enter the keyboard command string for the Shifted Key Mode. ")

         CONTROL ADD CHECKBOX, hKey, %KEYMAP_SUP, " ", 100, 341, 16, 16, %BS_LEFT OR %BS_VCENTER
         CONTROL SET CHECK     hKey, %KEYMAP_SUP, 0
         sToolTipSet (GetDlgItem(hKey, %KEYMAP_SUP), " Should keyboard AutoRepeat be allowed for this key? ")
         CONTROL DISABLE hKey, %KEYMAP_SUP

         CONTROL ADD LABEL, hKey, %KEYMAP_CONTROL_TEXT, "Ctrl", 5, 361, 100, 16
         CONTROL SET COLOR  hKey, %KEYMAP_CONTROL_TEXT, %BLACK, -2
         CONTROL DISABLE    hKey, %KEYMAP_CONTROL_TEXT
         CONTROL ADD TEXTBOX, hKey, %KEYMAP_CONTROL, ControlText, 125, 361, 335, 20
         CONTROL SET FONT   hKey, %KEYMAP_CONTROL, hScrFont
         CONTROL DISABLE    hKey, %KEYMAP_CONTROL
         sToolTipSet (GetDlgItem(hKey, %KEYMAP_CONTROL), " Enter the keyboard command string for the Control Key Mode. ")

         CONTROL ADD CHECKBOX, hKey, %KEYMAP_CUP, " ", 100, 364, 16, 16, %BS_LEFT OR %BS_VCENTER
         CONTROL SET CHECK     hKey, %KEYMAP_CUP, 0
         sToolTipSet (GetDlgItem(hKey, %KEYMAP_CUP), " Should keyboard AutoRepeat be allowed for this key? ")
         CONTROL DISABLE hKey, %KEYMAP_CUP

         CONTROL ADD LABEL, hKey, %KEYMAP_ALT_TEXT, "Alt", 5, 384, 100, 16
         CONTROL SET COLOR  hKey, %KEYMAP_ALT_TEXT, %BLACK, -2
         CONTROL DISABLE    hKey, %KEYMAP_ALT_TEXT
         CONTROL ADD TEXTBOX, hKey, %KEYMAP_ALT, AltText, 125, 384, 335, 20
         CONTROL SET FONT   hKey, %KEYMAP_ALT, hScrFont
         CONTROL DISABLE    hKey, %KEYMAP_ALT
         sToolTipSet (GetDlgItem(hKey, %KEYMAP_ALT), " Enter the keyboard command string for the Alt Key Mode. ")

         CONTROL ADD CHECKBOX, hKey, %KEYMAP_AUP, " ", 100, 387, 16, 16, %BS_LEFT OR %BS_VCENTER
         CONTROL SET CHECK     hKey, %KEYMAP_AUP, 0
         sToolTipSet (GetDlgItem(hKey, %KEYMAP_AUP), " Should keyboard AutoRepeat be allowed for this key? ")
         CONTROL DISABLE hKey, %KEYMAP_AUP

         CONTROL ADD LABEL, hKey, %KEYMAP_SHIFTCONTROL_TEXT, "Shift-Ctrl", 5, 407, 100, 16
         CONTROL SET COLOR  hKey, %KEYMAP_SHIFTCONTROL_TEXT, %BLACK, -2
         CONTROL DISABLE    hKey, %KEYMAP_SHIFTCONTROL_TEXT
         CONTROL ADD TEXTBOX, hKey, %KEYMAP_SHIFTCONTROL, SCText, 125, 407, 335, 20
         CONTROL SET FONT   hKey, %KEYMAP_SHIFTCONTROL, hScrFont
         CONTROL DISABLE    hKey, %KEYMAP_SHIFTCONTROL
         sToolTipSet (GetDlgItem(hKey, %KEYMAP_SHIFTCONTROL), " Enter the keyboard command string for the Shift-Control Key Mode. ")

         CONTROL ADD CHECKBOX, hKey, %KEYMAP_SCUP, " ", 100, 410, 16, 16, %BS_LEFT OR %BS_VCENTER
         CONTROL SET CHECK     hKey, %KEYMAP_SCUP, 0
         sToolTipSet (GetDlgItem(hKey, %KEYMAP_SCUP), " Should keyboard AutoRepeat be allowed for this key? ")
         CONTROL DISABLE hKey, %KEYMAP_SCUP

         CONTROL ADD LABEL, hKey, %KEYMAP_SHIFTALT_TEXT, "Shift-Alt", 5, 430, 100, 16
         CONTROL SET COLOR  hKey, %KEYMAP_SHIFTALT_TEXT, %BLACK, -2
         CONTROL DISABLE    hKey, %KEYMAP_SHIFTALT_TEXT
         CONTROL ADD TEXTBOX, hKey, %KEYMAP_SHIFTALT, SAText, 125, 430, 335, 20
         CONTROL SET FONT   hKey, %KEYMAP_SHIFTALT, hScrFont
         CONTROL DISABLE    hKey, %KEYMAP_SHIFTALT
         sToolTipSet (GetDlgItem(hKey, %KEYMAP_SHIFTALT), " Enter the keyboard command string for the Shift-Alt Key Mode. ")

         CONTROL ADD CHECKBOX, hKey, %KEYMAP_SAUP, " ", 100, 433, 16, 16, %BS_LEFT OR %BS_VCENTER
         CONTROL SET CHECK     hKey, %KEYMAP_SAUP, 0
         sToolTipSet (GetDlgItem(hKey, %KEYMAP_SAUP), " Should keyboard AutoRepeat be allowed for this key? ")
         CONTROL DISABLE hKey, %KEYMAP_SAUP

         CONTROL ADD LABEL, hKey, %KEYMAP_SHIFTCONTALT_TEXT, "Shift-Ctrl-Alt", 5, 453, 100, 16
         CONTROL SET COLOR  hKey, %KEYMAP_SHIFTCONTALT_TEXT, %BLACK, -2
         CONTROL DISABLE    hKey, %KEYMAP_SHIFTCONTALT_TEXT
         CONTROL ADD TEXTBOX, hKey, %KEYMAP_SHIFTCONTALT, SCAText, 125, 453, 335, 20
         CONTROL SET FONT   hKey, %KEYMAP_SHIFTCONTALT, hScrFont
         CONTROL DISABLE    hKey, %KEYMAP_SHIFTCONTALT
         sToolTipSet (GetDlgItem(hKey, %KEYMAP_SHIFTCONTALT), " Enter the keyboard command string for the Shift-Control-Alt Key Mode. ")

         CONTROL ADD CHECKBOX, hKey, %KEYMAP_SCAUP, " ", 100, 456, 16, 16, %BS_LEFT OR %BS_VCENTER
         CONTROL SET CHECK     hKey, %KEYMAP_SCAUP, 0
         sToolTipSet (GetDlgItem(hKey, %KEYMAP_SCAUP), " Should keyboard AutoRepeat be allowed for this key? ")
         CONTROL DISABLE hKey, %KEYMAP_SCAUP

         CONTROL ADD LABEL, hKey, %KEYMAP_CONTROLALT_TEXT, "Ctrl-Alt", 5, 476, 100, 16
         CONTROL SET COLOR  hKey, %KEYMAP_CONTROLALT_TEXT, %BLACK, -2
         CONTROL DISABLE    hKey, %KEYMAP_CONTROLALT_TEXT
         CONTROL ADD TEXTBOX, hKey, %KEYMAP_CONTROLALT, SCAText, 125, 476, 335, 20
         CONTROL SET FONT   hKey, %KEYMAP_CONTROLALT, hScrFont
         CONTROL DISABLE    hKey, %KEYMAP_CONTROLALT
         sToolTipSet (GetDlgItem(hKey, %KEYMAP_CONTROLALT), " Enter the keyboard command string for the Control-Alt Key Mode. ")

         CONTROL ADD CHECKBOX, hKey, %KEYMAP_CAUP, " ", 100, 479, 16, 16, %BS_LEFT OR %BS_VCENTER
         CONTROL SET CHECK     hKey, %KEYMAP_CAUP, 0
         sToolTipSet (GetDlgItem(hKey, %KEYMAP_CAUP), " Should keyboard AutoRepeat be allowed for this key? ")
         CONTROL DISABLE hKey, %KEYMAP_CAUP

         CONTROL ADD LABEL, hKey, %KEYMAP_REPEAT_TEXT, "^ Allow key combination to repeat?",100, 507, 250, 16
         CONTROL SET COLOR  hKey, %KEYMAP_REPEAT_TEXT, %BLACK, %RGB_GAINSBORO

         '----- DISPLAY the KeyMap DIALOG BOX
         GRAPHIC REDRAW                                                 '

         sPopReady                                                      ' Ready for pop-up
         DIALOG SHOW MODAL hKey, CALL DlgKeyMapCallBack                 '
         sPopReset                                                      ' Reset popup state
         IF hANSI THEN  DIALOG END hANSI, 0                             ' Now shut the ANSI dialog if used
         MExit
      END METHOD

      METHOD DlgKeyMapCallBack(BYVAL hDlg AS DWORD,  _            ' CB.HNDL
                               BYVAL wMsg AS LONG,   _            ' CB.MSG
                               BYVAL WPARAM AS LONG, _            ' CB.WPARAM    CB.CTLMSG = HI(WORD, wParam) CB.CTL = LO(WORD, wParam)
                               BYVAL LPARAM AS LONG) AS LONG      ' CB.LPARAM
      '--------------------
      ' Callback function used by the KeyMap Dialog
      '--------------------
      LOCAL i, j, Repeat AS LONG, t, tpf, u, kfield, MSG, sname, IconName AS STRING
      LOCAL BrkCount, SqrCount, SqgCount AS LONG
      STATIC lastID, lclmap, lclMapP AS LONG

         SELECT CASE AS LONG wMsg                                 '

            '----- KEYMAP is starting up
            CASE %WM_INITDIALOG                                   '
               lclMap = 0: lclMapP = 0: lastID = 0                ' Kill Previous pointers

            '----- Handle the various interactions
            CASE %WM_COMMAND                                      '
               SELECT CASE AS LONG LO(WORD, WPARAM)               ' Equiv to  CB.CTL

                  '----- CANCEL button pressed
                  CASE %KEYMAP_CANCEL                             '
                     IF HI(WORD, WPARAM) = %BN_CLICKED THEN       ' Equiv CB.CTLMSG - Cancel?
                        RESET hToolTips                           ' Clear so Tooltips get re-created correctly
                        DIALOG END hDlg                           ' Shut things down
                        me.LoadKBTable                            ' Reload Normal Keyboard stuff
                     END IF                                       '

                  '----- ANSI button, bring up the ANSI dialog
                  CASE %KEYMAP_ANSI                               '
                     IF HI(WORD, WPARAM) = %BN_CLICKED THEN       ' Equiv CB.CTLMSG - Done button pressed
                        TP.DispANSI (%False)                      ' Pop up the ANSI table in non-column mode
                     END IF                                       '

                  '----- DONE buton, lots of work now
                  CASE %KEYMAP_DONE                               '
                     IF HI(WORD, WPARAM) = %BN_CLICKED THEN       ' Equiv CB.CTLMSG - Done button pressed

                        '----- If a key currently displayed update the Repeat flags
                        IF lclMap <> 0 THEN                             ' If a previous key being displayed
                           CONTROL GET CHECK hDlg, %KEYMAP_NUP TO Repeat' Copy the previous key's Repeat key data
                           Kbd107(lclMap).NUp = IIF$(Repeat, ".", "U")  '
                           CONTROL GET CHECK hDlg, %KEYMAP_SUP TO Repeat
                           Kbd107(lclMap).SUp = IIF$(Repeat, ".", "U")  '
                           CONTROL GET CHECK hDlg, %KEYMAP_CUP TO Repeat
                           Kbd107(lclMap).CUp = IIF$(Repeat, ".", "U")  '
                           CONTROL GET CHECK hDlg, %KEYMAP_AUP TO Repeat
                           Kbd107(lclMap).AUp = IIF$(Repeat, ".", "U")  '
                           CONTROL GET CHECK hDlg, %KEYMAP_SCUP TO Repeat
                           Kbd107(lclMap).SCUp = IIF$(Repeat, ".", "U") '
                           CONTROL GET CHECK hDlg, %KEYMAP_SAUP TO Repeat
                           Kbd107(lclMap).SAUp = IIF$(Repeat, ".", "U") '
                           CONTROL GET CHECK hDlg, %KEYMAP_SCAUP TO Repeat
                           Kbd107(lclMap).SCAUp = IIF$(Repeat, ".", "U")'
                           CONTROL GET CHECK hDlg, %KEYMAP_CAUP TO Repeat
                           Kbd107(lclMap).CAUp = IIF$(Repeat, ".", "U") '
                        END IF                                          '

                        '----- Validate we have an (ENTER) key defined
                        j = 0                                           ' Clear J
                        FOR i = 1 TO 104                                ' Search the Table
                           IF UUCASE(Kbd107(i).NData) = "(ENTER)" THEN J += 1     ' Look for an Enter key
                           IF UUCASE(Kbd107(i).SData) = "(ENTER)" THEN J += 1     '
                           IF UUCASE(Kbd107(i).CData) = "(ENTER)" THEN J += 1     '
                           IF UUCASE(Kbd107(i).AData) = "(ENTER)" THEN J += 1     '
                           IF UUCASE(Kbd107(i).SCData) = "(ENTER)" THEN J += 1    '
                           IF UUCASE(Kbd107(i).SAData) = "(ENTER)" THEN J += 1    '
                           IF UUCASE(Kbd107(i).SCAData) = "(ENTER)" THEN J += 1 '
                           IF UUCASE(Kbd107(i).CAData) = "(ENTER)" THEN J += 1    '
                        NEXT i                                          '
                        IF j = 0 THEN                                   ' No Enter key found?
                           sDoMsgBox "You have no key defined as |K(Enter)|B, please assign one.", _
                           %MB_OK OR %MB_USERICON, "SPFLite KeyMap"    ' Issue a msg
                           EXIT METHOD                                  '
                        END IF                                          '

                        '----- Kill the dialog and save the new data
                        RESET hToolTips                                 ' Clear so Tooltips get re-created correctly
                        DIALOG END hDlg                                 ' Shut things down
                        me.WriteKBUser                                  ' Save things
                        me.LoadKBTable                                  ' Build active table
                     END IF                                             '

                  '----- We're in the NORMAL key area
                  CASE %KEYMAP_NORMAL                                   ' Normal data area?
                     IF HI(WORD, WPARAM) = %EN_KILLFOCUS THEN           ' Losing focus?
                        GOSUB ValNormal                                 ' Go process the key

                     END IF                                             '
                     IF HI(WORD, WPARAM) = %EN_SETFOCUS THEN            '
                        CONTROL SEND hDlg, LO(WORD, WPARAM), %EM_SETSEL, 0, 0    '
                     END IF                                             '

                  '----- We're in the SHIFT key area
                  CASE %KEYMAP_SHIFT                                    ' Shift data area?
                     IF HI(WORD, WPARAM) = %EN_KILLFOCUS THEN           ' Data changed?
                        GOSUB ValShift                                  ' Go process the key
                     END IF                                             '
                     IF HI(WORD, WPARAM) = %EN_SETFOCUS THEN            '
                        CONTROL SEND hDlg, LO(WORD, WPARAM), %EM_SETSEL, 0, 0    '
                     END IF                                             '

                  '----- We're in the CONTROL key area
                  CASE %KEYMAP_CONTROL                                  ' Control data area?
                     IF HI(WORD, WPARAM) = %EN_KILLFOCUS THEN           ' Data changed?
                        GOSUB ValControl                                ' Go process the key
                     END IF                                             '
                     IF HI(WORD, WPARAM) = %EN_SETFOCUS THEN            '
                        CONTROL SEND hDlg, LO(WORD, WPARAM), %EM_SETSEL, 0, 0    '
                     END IF                                             '

                  '----- We're in the ALT key area
                  CASE %KEYMAP_ALT                                      ' Normal data area?
                     IF HI(WORD, WPARAM) = %EN_KILLFOCUS THEN           ' Data changed?
                        GOSUB ValAlt                                    ' Go process the key
                     END IF                                             '
                     IF HI(WORD, WPARAM) = %EN_SETFOCUS THEN            '
                        CONTROL SEND hDlg, LO(WORD, WPARAM), %EM_SETSEL, 0, 0    '
                     END IF                                             '

                  '----- We're in the SHIFT CONTOL key area
                  CASE %KEYMAP_SHIFTCONTROL                             ' Shift-Control Data area?
                     IF HI(WORD, WPARAM) = %EN_KILLFOCUS THEN           ' Losing focus?
                        GOSUB ValShiftControl                           ' Go process the key
                     END IF                                             '
                     IF HI(WORD, WPARAM) = %EN_SETFOCUS THEN            '
                        CONTROL SEND hDlg, LO(WORD, WPARAM), %EM_SETSEL, 0, 0    '
                     END IF                                             '

                  '----- We're in the SHIFT ALT key area
                  CASE %KEYMAP_SHIFTALT                                 ' Shift-Alt data area?
                     IF HI(WORD, WPARAM) = %EN_KILLFOCUS THEN           ' Data changed?
                        GOSUB ValShiftAlt                               ' Go process the key
                     END IF                                             '
                     IF HI(WORD, WPARAM) = %EN_SETFOCUS THEN            '
                        CONTROL SEND hDlg, LO(WORD, WPARAM), %EM_SETSEL, 0, 0    '
                     END IF                                             '

                  '----- We're in the SHIFT CONTROL ALT key area
                  CASE %KEYMAP_SHIFTCONTALT                             ' Shift-Control-Alt data area?
                     IF HI(WORD, WPARAM) = %EN_KILLFOCUS THEN           ' Data changed?
                        GOSUB ValShiftCtlAlt                            ' Go process the key
                     END IF                                             '
                     IF HI(WORD, WPARAM) = %EN_SETFOCUS THEN            '
                        CONTROL SEND hDlg, LO(WORD, WPARAM), %EM_SETSEL, 0, 0    '
                     END IF                                             '

                  '----- We're in the CONTROL ALT key area
                  CASE %KEYMAP_CONTROLALT                               ' Control-Alt data area?
                     IF HI(WORD, WPARAM) = %EN_KILLFOCUS THEN           ' Data changed?
                        GOSUB ValControlAlt                             ' Go process the key
                     END IF                                             '
                     IF HI(WORD, WPARAM) = %EN_SETFOCUS THEN            '
                        CONTROL SEND hDlg, LO(WORD, WPARAM), %EM_SETSEL, 0, 0    '
                     END IF                                             '

                  '----- One of the Key Icons was pressed
                  CASE > 5000                                           ' A keyboard key?
                     IF HI(WORD, WPARAM) = %BN_CLICKED THEN             ' Button pressed

                        '----- If a previous key being displayed, validate it
                        IF lclMap <> 0 THEN                             ' Was a key displayed?
                           GOSUB ValNormal                              ' Process key data
                           GOSUB ValShift                               '
                           GOSUB ValControl                             '
                           GOSUB ValAlt                                 '
                           GOSUB ValShiftControl                        '
                           GOSUB ValShiftAlt                            '
                           GOSUB ValShiftCtlAlt                         '
                           GOSUB ValControlAlt                          '

                           CONTROL GET CHECK hDlg, %KEYMAP_NUP TO Repeat' Copy the previous key's Repeat key data
                           Kbd107(lclMap).NUp = IIF$(Repeat, ".", "U")  ' Stuff back in Allow area

                           CONTROL GET CHECK hDlg, %KEYMAP_SUP TO Repeat
                           Kbd107(lclMap).SUp = IIF$(Repeat, ".", "U")

                           CONTROL GET CHECK hDlg, %KEYMAP_CUP TO Repeat
                           Kbd107(lclMap).CUp = IIF$(Repeat, ".", "U")

                           CONTROL GET CHECK hDlg, %KEYMAP_AUP TO Repeat
                           Kbd107(lclMap).AUp = IIF$(Repeat, ".", "U")

                           CONTROL GET CHECK hDlg, %KEYMAP_SCUP TO Repeat
                           Kbd107(lclMap).SCUp = IIF$(Repeat, ".", "U")

                           CONTROL GET CHECK hDlg, %KEYMAP_SAUP TO Repeat
                           Kbd107(lclMap).SAUp = IIF$(Repeat, ".", "U")

                           CONTROL GET CHECK hDlg, %KEYMAP_SCAUP TO Repeat
                           Kbd107(lclMap).SCAUp = IIF$(Repeat, ".", "U")

                           CONTROL GET CHECK hDlg, %KEYMAP_CAUP TO Repeat
                           Kbd107(lclMap).CAUp = IIF$(Repeat, ".", "U")

                        END IF                                    '

                        '----- See if the key is even mappable
                        CONTROL GET USER hDlg, LO(WORD, WPARAM), 1 TO lclMap ' Get pointer to KeyMaster area
                        IF Kbd107(lclMap).ForceUp = "D" THEN      ' Disallowed key?
                           sDoMsgBox "The |K" + Kbd107(lclMap).Labl + "|B key is not mappable", _
                                     %MB_OK OR %MB_USERICON, "SPFLite KeyMap"
                           EXIT METHOD                            '
                        END IF                                    '

                        '----- Alter the ICON if needed
                        IF lclMapP <> 0 AND lclMapP <> lclMap THEN' Switching keys?
                           IconName = Kbd107(lclMapP).kIcon       ' Restore the normal Icon
                           CONTROL SET IMAGEX hDlg, lastID, IconName ' Set the Selected Icon image
                        END IF                                    '

                        '----- Get the key's data displayed
                        lclMapP = lclMap                          ' Save as previous
                        lastID = LO(WORD, WPARAM)                 '
                        IconName = "S" + FORMAT$(Kbd107(lclMapP).W) + FORMAT$(Kbd107(lclMapP).H) ' Build the name of the Blob Icon
                        CONTROL SET IMAGEX hDlg, lastID, IconName ' Set the Selected Icon image
                        CONTROL SET TEXT hDlg, %KEYMAP_PROMPT, "" ' Kill Prompt string
                        CONTROL SET TEXT hDlg, %KEYMAP_CURRENT_TEXT, "Current Key: " + Kbd107(lclMapP).Labl ' Set in Dialog
                        CONTROL KILL hDlg, %KEYMAP_CURRENT_ICON         ' Kill Previous
                        t = Kbd107(lclMap).KIcon                        '
                        CONTROL ADD IMAGEX, hDlg,%KEYMAP_CURRENT_ICON, Kbd107(lclMap).KIcon, 175, 280 - (Kbd107(lclMap).H / 2), _
                                      Kbd107(lclMap).W, Kbd107(lclMap).H, %SS_ICON OR %SS_CENTERIMAGE OR %SS_NOTIFY
                        sname = Kbd107(lclMapP).Labl                    ' Save trimmed key label
                        sToolTipSet (GetDlgItem(hDlg, %KEYMAP_SCAUP), " Should keyboard AutoRepeat be allowed for this key? ")
                        sToolTipSet (GetDlgItem(hDlg, %KEYMAP_SAUP), " Should keyboard AutoRepeat be allowed for this key? ")
                        sToolTipSet (GetDlgItem(hDlg, %KEYMAP_CAUP), " Should keyboard AutoRepeat be allowed for this key? ")
                        sToolTipSet (GetDlgItem(hDlg, %KEYMAP_SCUP), " Should keyboard AutoRepeat be allowed for this key? ")
                        sToolTipSet (GetDlgItem(hDlg, %KEYMAP_SUP), " Should keyboard AutoRepeat be allowed for this key? ")
                        sToolTipSet (GetDlgItem(hDlg, %KEYMAP_CUP), " Should keyboard AutoRepeat be allowed for this key? ")
                        sToolTipSet (GetDlgItem(hDlg, %KEYMAP_AUP), " Should keyboard AutoRepeat be allowed for this key? ")
                        sToolTipSet (GetDlgItem(hDlg, %KEYMAP_NUP),  " Should keyboard AutoRepeat be allowed for this key? ")
                        CONTROL SET TEXT hDlg, %KEYMAP_REPEAT_TEXT, "^ Allow key combination to repeat?"
                        IF sname = "LMB" OR sname = "MMB" OR sname = "RMB" THEN
                           CONTROL SET TEXT hDlg, %KEYMAP_REPEAT_TEXT, "^ Position Cursor when mouse clicked?"
                           sToolTipSet (GetDlgItem(hDlg, %KEYMAP_SCAUP), " Should cursor be set when the mouse button is clicked? ")
                           sToolTipSet (GetDlgItem(hDlg, %KEYMAP_SAUP), " Should cursor be set when the mouse button is clicked? ")
                           sToolTipSet (GetDlgItem(hDlg, %KEYMAP_CAUP), " Should cursor be set when the mouse button is clicked? ")
                           sToolTipSet (GetDlgItem(hDlg, %KEYMAP_SCUP), " Should cursor be set when the mouse button is clicked? ")
                           sToolTipSet (GetDlgItem(hDlg, %KEYMAP_SUP), " Should cursor be set when the mouse button is clicked? ")
                           sToolTipSet (GetDlgItem(hDlg, %KEYMAP_CUP), " Should cursor be set when the mouse button is clicked? ")
                           sToolTipSet (GetDlgItem(hDlg, %KEYMAP_AUP), " Should cursor be set when the mouse button is clicked? ")
                           sToolTipSet (GetDlgItem(hDlg, %KEYMAP_NUP), " Should cursor be set when the mouse button is clicked? ")
                        END IF                                          '


                        IF Kbd107(lclMap).NAllow = "N" THEN             ' Open up the Normal key input?
                           CONTROL ENABLE hDlg, %KEYMAP_NORMAL_TEXT     '
                           CONTROL ENABLE hDlg, %KEYMAP_NORMAL          '
                           CONTROL ENABLE hDlg, %KEYMAP_NUP             '
                           CONTROL SET CHECK hDlg, %KEYMAP_NUP, IIF(Kbd107(lclMap).NUp = ".", 1, 0)
                           CONTROL SET TEXT hDlg, %KEYMAP_NORMAL, Kbd107(lclMap).NData
                        ELSE                                            '
                           CONTROL DISABLE hDlg, %KEYMAP_NORMAL_TEXT    '
                           CONTROL DISABLE hDlg, %KEYMAP_NORMAL         '
                           CONTROL SET CHECK hDlg, %KEYMAP_NUP, 0       '
                           CONTROL DISABLE hDlg, %KEYMAP_NUP            '
                           CONTROL SET TEXT hDlg, %KEYMAP_NORMAL, ""    '
                        END IF                                          '
                        IF Kbd107(lclMap).ForceUp = "U" THEN            ' Forced Upkey mode?
                           CONTROL DISABLE hDlg, %KEYMAP_NUP            '
                           CONTROL SET CHECK hDlg, %KEYMAP_NUP, 0       '
                        END IF                                          '

                        IF Kbd107(lclMap).SAllow = "S" THEN             ' Open up the Shift key input?
                           CONTROL ENABLE hDlg, %KEYMAP_SHIFT_TEXT      '
                           CONTROL ENABLE hDlg, %KEYMAP_SHIFT           '
                           CONTROL ENABLE hDlg, %KEYMAP_SUP             '
                           CONTROL SET CHECK hDlg, %KEYMAP_SUP, IIF(Kbd107(lclMap).SUp = ".", 1, 0)
                           CONTROL SET TEXT hDlg, %KEYMAP_SHIFT, Kbd107(lclMap).SData
                        ELSE                                            '
                           CONTROL DISABLE hDlg, %KEYMAP_SHIFT_TEXT     '
                           CONTROL DISABLE hDlg, %KEYMAP_SHIFT          '
                           CONTROL SET CHECK hDlg, %KEYMAP_SUP, 0       '
                           CONTROL DISABLE hDlg, %KEYMAP_SUP            '
                           CONTROL SET TEXT hDlg, %KEYMAP_SHIFT, ""     '
                        END IF                                          '
                        IF Kbd107(lclMap).ForceUp = "U" THEN            ' Forced Upkey mode?
                           CONTROL DISABLE hDlg, %KEYMAP_SUP            '
                           CONTROL SET CHECK hDlg, %KEYMAP_SUP, 0       '
                        END IF                                          '

                        IF Kbd107(lclMap).CAllow = "C" THEN             ' Open up the Control key input?
                           CONTROL ENABLE hDlg, %KEYMAP_CONTROL_TEXT    '
                           CONTROL ENABLE hDlg, %KEYMAP_CONTROL         '
                           CONTROL ENABLE hDlg, %KEYMAP_CUP             '
                           CONTROL SET CHECK hDlg, %KEYMAP_CUP, IIF(Kbd107(lclMap).CUp = ".", 1, 0)
                           CONTROL SET TEXT hDlg, %KEYMAP_CONTROL, Kbd107(lclMap).CData
                        ELSE                                            '
                           CONTROL DISABLE hDlg, %KEYMAP_CONTROL_TEXT   '
                           CONTROL DISABLE hDlg, %KEYMAP_CONTROL        '
                           CONTROL SET CHECK hDlg, %KEYMAP_CUP, 0       '
                           CONTROL DISABLE hDlg, %KEYMAP_CUP            '
                           CONTROL SET TEXT hDlg, %KEYMAP_CONTROL, ""   '
                        END IF                                          '
                        IF Kbd107(lclMap).ForceUp = "U" THEN            ' Forced Upkey mode?
                           CONTROL DISABLE hDlg, %KEYMAP_CUP            '
                           CONTROL SET CHECK hDlg, %KEYMAP_CUP, 0       '
                        END IF                                          '

                        IF Kbd107(lclMap).AAllow = "A" THEN             ' Open up the Alt key input?
                           CONTROL ENABLE hDlg, %KEYMAP_ALT_TEXT        '
                           CONTROL ENABLE hDlg, %KEYMAP_ALT             '
                           CONTROL ENABLE hDlg, %KEYMAP_AUP             '
                           CONTROL SET CHECK hDlg, %KEYMAP_AUP, IIF(Kbd107(lclMap).AUp = ".", 1, 0)
                           CONTROL SET TEXT hDlg, %KEYMAP_ALT, Kbd107(lclMap).AData
                        ELSE                                            '
                           CONTROL DISABLE hDlg, %KEYMAP_ALT_TEXT       '
                           CONTROL DISABLE hDlg, %KEYMAP_ALT            '
                           CONTROL SET CHECK hDlg, %KEYMAP_AUP, 0       '
                           CONTROL DISABLE hDlg, %KEYMAP_AUP            '
                           CONTROL SET TEXT hDlg, %KEYMAP_ALT, ""       '
                        END IF                                          '
                        IF Kbd107(lclMap).ForceUp = "U" THEN            ' Forced Upkey mode?
                           CONTROL DISABLE hDlg, %KEYMAP_AUP            '
                           CONTROL SET CHECK hDlg, %KEYMAP_AUP, 0       '
                        END IF                                          '

                        IF Kbd107(lclMap).SAllow = "S" AND Kbd107(lclMap).CAllow = "C" THEN  ' Open up the Shift-Control?
                           CONTROL ENABLE hDlg, %KEYMAP_SHIFTCONTROL_TEXT
                           CONTROL ENABLE hDlg, %KEYMAP_SHIFTCONTROL    '
                           CONTROL ENABLE hDlg, %KEYMAP_SCUP            '
                           CONTROL SET CHECK hDlg, %KEYMAP_SCUP, IIF(Kbd107(lclMap).SCUp = ".", 1, 0)
                           CONTROL SET TEXT hDlg, %KEYMAP_SHIFTCONTROL, Kbd107(lclMap).SCData
                        ELSE                                            '
                           CONTROL DISABLE hDlg, %KEYMAP_SHIFTCONTROL_TEXT
                           CONTROL DISABLE hDlg, %KEYMAP_SHIFTCONTROL   '
                           CONTROL SET CHECK hDlg, %KEYMAP_SCUP, 0      '
                           CONTROL DISABLE hDlg, %KEYMAP_SCUP           '
                           CONTROL SET TEXT hDlg, %KEYMAP_SHIFTCONTROL, ""
                        END IF                                          '
                        IF Kbd107(lclMap).ForceUp = "U" THEN            ' Forced Upkey mode?
                           CONTROL DISABLE hDlg, %KEYMAP_SCUP           '
                           CONTROL SET CHECK hDlg, %KEYMAP_SCUP, 0      '
                        END IF                                          '

                        IF Kbd107(lclMap).SAllow = "S" AND Kbd107(lclMap).AAllow = "A" THEN ' Open up the Shift-Alt?
                           CONTROL ENABLE hDlg, %KEYMAP_SHIFTALT_TEXT   '
                           CONTROL ENABLE hDlg, %KEYMAP_SHIFTALT        '
                           CONTROL ENABLE hDlg, %KEYMAP_SAUP            '
                           CONTROL SET CHECK hDlg, %KEYMAP_SAUP, IIF(Kbd107(lclMap).SAUp = ".", 1, 0)
                           CONTROL SET TEXT hDlg, %KEYMAP_SHIFTALT, Kbd107(lclMap).SAData
                        ELSE                                            '
                           CONTROL DISABLE hDlg, %KEYMAP_SHIFTALT_TEXT  '
                           CONTROL DISABLE hDlg, %KEYMAP_SHIFTALT       '
                           CONTROL SET CHECK hDlg, %KEYMAP_SAUP, 0      '
                           CONTROL DISABLE hDlg, %KEYMAP_SAUP           '
                           CONTROL SET TEXT hDlg, %KEYMAP_SHIFTALT, ""  '
                        END IF                                          '
                        IF Kbd107(lclMap).ForceUp = "U" THEN            ' Forced Upkey mode?
                           CONTROL DISABLE hDlg, %KEYMAP_SAUP           '
                           CONTROL SET CHECK hDlg, %KEYMAP_SAUP, 0      '
                        END IF                                          '

                        IF Kbd107(lclMap).SAllow = "S" AND Kbd107(lclMap).CAllow = "C" AND Kbd107(lclMap).AAllow = "A" THEN ' Open up the Shift-Control-Alt?
                           CONTROL ENABLE hDlg, %KEYMAP_SHIFTCONTALT_TEXT
                           CONTROL ENABLE hDlg, %KEYMAP_SHIFTCONTALT    '
                           CONTROL ENABLE hDlg, %KEYMAP_SCAUP           '
                           CONTROL SET CHECK hDlg, %KEYMAP_SCAUP, IIF(Kbd107(lclMap).SCAUp = ".", 1, 0)
                           CONTROL SET TEXT hDlg, %KEYMAP_SHIFTCONTALT, Kbd107(lclMap).SCAData
                        ELSE                                            '
                           CONTROL DISABLE hDlg, %KEYMAP_SHIFTCONTALT_TEXT
                           CONTROL DISABLE hDlg, %KEYMAP_SHIFTCONTALT   '
                           CONTROL SET CHECK hDlg, %KEYMAP_SCAUP, 0     '
                           CONTROL DISABLE hDlg, %KEYMAP_SCAUP          '
                           CONTROL SET TEXT hDlg, %KEYMAP_SHIFTCONTALT, ""
                        END IF                                          '
                        IF Kbd107(lclMap).ForceUp = "U" THEN            ' Forced Upkey mode?
                           CONTROL DISABLE hDlg, %KEYMAP_SCAUP          '
                           CONTROL SET CHECK hDlg, %KEYMAP_SCAUP, 0     '
                        END IF                                          '

                        IF Kbd107(lclMap).CAllow = "C" AND Kbd107(lclMap).AAllow = "A" THEN ' Open up the Control-Alt?
                           CONTROL ENABLE hDlg, %KEYMAP_CONTROLALT_TEXT '
                           CONTROL ENABLE hDlg, %KEYMAP_CONTROLALT      '
                           CONTROL ENABLE hDlg, %KEYMAP_CAUP            '
                           CONTROL SET CHECK hDlg, %KEYMAP_CAUP, IIF(Kbd107(lclMap).CAUp = ".", 1, 0)
                           CONTROL SET TEXT hDlg, %KEYMAP_CONTROLALT, Kbd107(lclMap).CAData
                        ELSE                                            '
                           CONTROL DISABLE hDlg, %KEYMAP_CONTROLALT_TEXT
                           CONTROL DISABLE hDlg, %KEYMAP_CONTROLALT     '
                           CONTROL SET CHECK hDlg, %KEYMAP_CAUP, 0      '
                           CONTROL DISABLE hDlg, %KEYMAP_CAUP           '
                           CONTROL SET TEXT hDlg, %KEYMAP_CONTROLALT, ""
                        END IF                                          '
                        IF Kbd107(lclMap).ForceUp = "U" THEN            ' Forced Upkey mode?
                           CONTROL DISABLE hDlg, %KEYMAP_CAUP           '
                           CONTROL SET CHECK hDlg, %KEYMAP_CAUP, 0      '
                        END IF

                     END IF

                  '----- Show the Primitive keys hint list
                  CASE %KEYMAP_PRIMITIVEHINT                            ' Keymap Primitive hint list
                     IF HI(WORD, WPARAM) = %CBN_SELENDOK THEN           ' Item Selected
                        COMBOBOX GET SELECT hDlg, %KEYMAP_PRIMITIVEHINT TO i  ' Get the selected Index
                        IF i > 1 THEN                                   ' Ignore the heading
                           t = gPrimTable(i - 1)                        '
                           sWinclip_set(t)                              ' Stuff it into clipboard
                        END IF                                          '
                     END IF                                             '


               END SELECT                                               '
         END SELECT                                                     '
         EXIT METHOD

      '----- Validate a key's entry
      ValNormal:
         CONTROL GET TEXT hDlg, %KEYMAP_NORMAL TO t                     ' Get the modified data
         t = IIF$(ISNULL(TRIM$(t)), "(Passthru)", TRIM$(t))             '
         kfield = "Normal mode key"                                     '
         GOSUB ValKey                                                   ' Do some basic validation
         Kbd107(lclMap).NData = t                                       ' Replace in case ValKey 'fixed' it
         CONTROL SET TEXT hDlg, %KEYMAP_NORMAL, t                       ' Re-display it
         RETURN                                                         '

      ValShift:
         CONTROL GET TEXT hDlg, %KEYMAP_SHIFT TO t                      ' Get the modified data
         t = IIF$(ISNULL(TRIM$(t)), "(Passthru)", TRIM$(t))             '
         kfield = "Shift mode key"                                      '
         GOSUB ValKey                                                   ' Do some basic validation
         Kbd107(lclMap).SData = t                                       ' Replace in case ValKey 'fixed' it
         CONTROL SET TEXT hDlg, %KEYMAP_SHIFT, t                        ' Re-display it
         RETURN                                                         '

      ValControl:
         CONTROL GET TEXT hDlg, %KEYMAP_CONTROL TO t                    ' Get the modified data
         t = IIF$(ISNULL(TRIM$(t)), "(Null)", TRIM$(t))                 '
         kfield = "Control mode key"                                    '
         GOSUB ValKey                                                   ' Do some basic validation
         Kbd107(lclMap).CData = t                                       ' Replace in case ValKey 'fixed' it
         CONTROL SET TEXT hDlg, %KEYMAP_CONTROL, t                      ' Re-display it
         RETURN                                                         '

      ValAlt:
         CONTROL GET TEXT hDlg, %KEYMAP_ALT TO t                        ' Get the modified data
         t = IIF$(ISNULL(TRIM$(t)), "(Null)", TRIM$(t))                 '
         kfield = "Alt mode key"                                        '
         GOSUB ValKey                                                   ' Do some basic validation
         Kbd107(lclMap).AData = t                                       ' Replace in case ValKey 'fixed' it
         CONTROL SET TEXT hDlg, %KEYMAP_ALT, t                          ' Re-display it
         RETURN                                                         '

      ValShiftControl:
         CONTROL GET TEXT hDlg, %KEYMAP_SHIFTCONTROL TO t               ' Get the modified data
         t = IIF$(ISNULL(TRIM$(t)), "(Null)", TRIM$(t))                 '
         kfield = "Shift-Control mode key"                              '
         GOSUB ValKey                                                   ' Do some basic validation
         Kbd107(lclMap).SCData = t                                      ' Replace in case ValKey 'fixed' it
         CONTROL SET TEXT hDlg, %KEYMAP_SHIFTCONTROL, t                 ' Re-display it
         RETURN                                                         '

      ValShiftAlt:
         CONTROL GET TEXT hDlg, %KEYMAP_SHIFTALT TO t                   ' Get the modified data
         t = IIF$(ISNULL(TRIM$(t)), "(Null)", TRIM$(t))                 '
         kfield = "Shift-Alt mode key"                                  '
         GOSUB ValKey                                                   ' Do some basic validation
         Kbd107(lclMap).SAData = t                                      ' Replace in case ValKey 'fixed' it
         CONTROL SET TEXT hDlg, %KEYMAP_SHIFTALT, t                     ' Re-display it
         RETURN                                                         '

      ValShiftCtlAlt:
         CONTROL GET TEXT hDlg, %KEYMAP_SHIFTCONTALT TO t               ' Get the modified data
         t = IIF$(ISNULL(TRIM$(t)), "(Null)", TRIM$(t))                 '
         kfield = "Shift-Control-Alt mode key"                          '
         GOSUB ValKey                                                   ' Do some basic validation
         Kbd107(lclMap).SCAData = t                                     ' Replace in case ValKey 'fixed' it
         CONTROL SET TEXT hDlg, %KEYMAP_SHIFTCONTALT, t                 ' Re-display it
         RETURN                                                         '

      ValControlAlt:
         CONTROL GET TEXT hDlg, %KEYMAP_CONTROLALT TO t                 ' Get the modified data
         t = IIF$(ISNULL(TRIM$(t)), "(Null)", TRIM$(t))                 '
         kfield = "Control-Alt mode key"                                '
         GOSUB ValKey                                                   ' Do some basic validation
         Kbd107(lclMap).CAData = t                                      ' Replace in case ValKey 'fixed' it
         CONTROL SET TEXT hDlg, %KEYMAP_CONTROLALT, t                   ' Re-display it
         RETURN                                                         '

      '----- Common routine called by all the key types above
      ValKey:
         IF ISNULL(t) THEN RETURN                                       ' Nulls are always valid
         tpf = ""                                                       ' Null <...> prefix save
         IF LEFT$(t, 1) = "<" THEN                                      ' PFSHOW labelled?
            i = INSTR(t, ">")                                           ' Look for closing
            IF i THEN                                                   ' Found it, strip it off
               tpf = LEFT$(t, i)                                        ' And save in tpf
               t = MID$(t, i + 1)                                       '
            ELSE                                                        '
               MSG = "invalid PFSHOW label"                             '
               GOSUB ValError                                           '
               RETURN                                                   '
            END IF                                                      '
         END IF                                                         '

         IF LEFT$(t, 1) = ":" AND INSTR(t, ANY "(){}") = 0 THEN         ' Probable ISPF style line command?
            u = t: t = "{" + MID$(t, 2) + "}"                           ' Convert it
            sDoMsgBox("ISPF style line command detected." + $CRLF + $CRLF + _
                      "The string: |K" + u + "|B has been converted to: |K" + t, _
                      %MB_OK OR %MB_USERICON, "Line Command converted")
         END IF                                                         '

         SELECT CASE CONST$ LEFT$(t, 1)                                 ' See if a bracketd command string?
            CASE "{"
               IF INSTR(t, "}") = 0 THEN GOSUB BrkErr: RETURN           ' Make sure closing } is there
               IF LEN(PARSE$(t, ANY "{}", 2)) > 8 THEN                  ' Line command too long?
                  MSG = "line command too long"                         '
                  GOSUB ValError                                        '
                  RETURN                                                '
               END IF                                                   '
            CASE "(": IF INSTR(t, ")") = 0 THEN GOSUB BrkErr: RETURN    ' Make sure closing ) is there
            CASE "[": IF INSTR(t, "]") = 0 THEN GOSUB BrkErr: RETURN           ' Make sure closing ] is there
         END SELECT

         BrkCount = TALLY(t, ANY "()")                                  ' Count ()
         IF INSTR(UUCASE(t), "(PASSTHRU)") > 0 THEN                     ' (Passthru) should be alone
            IF BrkCount > 2 OR INSTR(t, ")") <> LEN(t)  THEN            '
               MSG = "(Passthru) may not be combined with anything else"'
               GOSUB ValError                                           '
               RETURN                                                   '
            END IF                                                      '
         END IF                                                         '
         IF (BrkCount MOD 2) <> 0 THEN                                  ' Test ( ) count
            MSG = "may have mismatched Primitive parentheses ( )"       '
            GOSUB ValError                                              '
            RETURN                                                      '
         END IF                                                         '

         SqgCount = TALLY(t, ANY "{}")                                  ' Count various brackets
         IF (SqgCount MOD 2) <> 0 THEN                                  ' Test { } count
            MSG = "may have mismatched Line Command braces { }"         '
            GOSUB ValError                                              '
            RETURN                                                      '
         END IF                                                         '
         t = tpf + t                                                    ' Restore <...>
         RETURN                                                         '

      BrkErr:
         MSG = "missing closing for " + LEFT$(t, 1)                     ' Finalize message
         GOSUB ValError                                                 ' Issue it
         RETURN                                                         '

      ValError:
         MSGBOX "Warning! " + kfield + " - " + MSG, _                   ' Issue a msg
                   %MB_OK OR %MB_ICONERROR, "SPFLite KeyMap"
         RETURN
      END METHOD

   END INTERFACE

END CLASS

'----- Keyboard Table Class definitions

CLASS cKBMastEntry
   INSTANCE KIcon           AS STRING                             ' Name of the Icon in the resource file
   INSTANCE W               AS LONG                               ' Icon width
   INSTANCE H               AS LONG                               ' Icon height
   INSTANCE X               AS LONG                               ' Icon X location
   INSTANCE Y               AS LONG                               ' Icon Y location
   INSTANCE Extern          AS STRING                             ' Extended status
   INSTANCE NAllow          AS STRING                             ' Allowable Normal key
   INSTANCE SAllow          AS STRING                             ' Allowable Shift key
   INSTANCE CAllow          AS STRING                             ' Allowable Control key
   INSTANCE AAllow          AS STRING                             ' Allowable Alt key
   INSTANCE UpKey           AS STRING                             ' Allowable UpKey detect only
   INSTANCE ForceUp         AS STRING                             ' Forced UpKey / Disable All
   INSTANCE NUp             AS STRING                             ' Normal  Up Only
   INSTANCE SUp             AS STRING                             ' Shift   Up Only
   INSTANCE CUp             AS STRING                             ' Control Up Only
   INSTANCE AUp             AS STRING                             ' Alt     Up only
   INSTANCE SCUp            AS STRING                             ' Shift Control Up Only
   INSTANCE SAUp            AS STRING                             ' Shift Alt Up Only
   INSTANCE SCAUp           AS STRING                             ' Shift Control Alt Up Only
   INSTANCE CAUp            AS STRING                             ' Control Alt Up Only
   INSTANCE Labl            AS STRING                             ' External name
   INSTANCE NData           AS STRING                             ' Normal  Data
   INSTANCE SData           AS STRING                             ' Shift   Data
   INSTANCE CData           AS STRING                             ' Control Data
   INSTANCE AData           AS STRING                             ' Alt     Data
   INSTANCE SCData          AS STRING                             ' Shift Control     Data
   INSTANCE SAData          AS STRING                             ' Shift Alt         Data
   INSTANCE SCAData         AS STRING                             ' Shift Control Alt Data
   INSTANCE CAData          AS STRING                             ' Control Alt       Data

   INTERFACE iKBMastEntry: INHERIT IUNKNOWN
      gsProp(KIcon, STRING)                                       ' Icon Name
      gsProp(W,LONG)                                              ' Icon width
      gsProp(H,LONG)                                              ' Icon height
      gsProp(X,LONG)                                              ' Icon X location
      gsProp(Y,LONG)                                              ' Icon Y location
      gsProp(Extern,STRING)                                       ' Extended status
      gsProp(NAllow,STRING)                                       ' Allowable Normal key
      gsProp(SAllow,STRING)                                       ' Allowable Shift key
      gsProp(CAllow,STRING)                                       ' Allowable Control key
      gsProp(AAllow,STRING)                                       ' Allowable Alt key
      gsProp(UpKey,STRING)                                        ' Allowable UpKey detect only
      gsProp(ForceUp,STRING)                                      ' Forced UpKey / Disable All
      gsProp(NUp,STRING)                                          ' Normal  Up Only
      gsProp(SUp,STRING)                                          ' Shift   Up Only
      gsProp(CUp,STRING)                                          ' Control Up Only
      gsProp(AUp,STRING)                                          ' Alt     Up only
      gsProp(SCUp,STRING)                                         ' Shift Control Up Only
      gsProp(SAUp,STRING)                                         ' Shift Alt Up Only
      gsProp(SCAUp,STRING)                                        ' Shift Control Alt Up Only
      gsProp(CAUp,STRING)                                         ' Control Alt Up Only
      gsProp(Labl,STRING)                                         ' External name
      gsProp(NData,STRING)                                        ' Normal  Data
      gsProp(SData,STRING)                                        ' Shift   Data
      gsProp(CData,STRING)                                        ' Control Data
      gsProp(AData,STRING)                                        ' Alt     Data
      gsProp(SCData,STRING)                                       ' Shift Control     Data
      gsProp(SAData,STRING)                                       ' Shift Alt         Data
      gsProp(SCAData,STRING)                                      ' Shift Control Alt Data
      gsProp(CAData,STRING)                                       ' Control Alt       Data
   END INTERFACE
END CLASS

CLASS cKBactive
   INSTANCE SupRepeat       AS STRING                             ' Suppress repeat (Up signal only)
   INSTANCE SupSent         AS STRING                             ' Suppressed, 1st has been sent
   INSTANCE KData           AS STRING                             ' Keyboard data string
   INTERFACE iKBActive: INHERIT IUNKNOWN
      gsProp(SupRepeat,STRING)                                    ' Suppress repeat (Up signal only)
      gsProp(SupSent,STRING)                                      ' Suppressed, 1st has been sent
      gsProp(KData,STRING)                                        ' Keyboard data string
   END INTERFACE
END CLASS

CALLBACK FUNCTION DlgKeyMapCallBack()
   FUNCTION = KbdT.DlgKeyMapCallBack(CB.HNDL, CB.MSG, CB.WPARAM, CB.LPARAM)
END FUNCTION
