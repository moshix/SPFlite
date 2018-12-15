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

CLASS cIO

   '----- IO Data areas
   INSTANCE PromptTxt      AS STRING                              ' Prompt string
   INSTANCE FNum           AS LONG                                ' File number
   INSTANCE TabFound       AS LONG                                ' Where found if Already asked for
   INSTANCE ResultMsg      AS STRING                              ' Result message text
   INSTANCE FilePath       AS STRING                              ' Full file name and path
   INSTANCE SetProfile     AS STRING                              ' User Provided Profile
   INSTANCE FD             AS DIRDATA                             ' Full DIRDATA area
   INSTANCE RQAlready      AS INTEGER                             ' Already Open in a tab test
   INSTANCE RQBinary       AS INTEGER                             ' OPEN for Binary
   INSTANCE RQCreateIf     AS INTEGER                             ' Create if it doesn't exist
   INSTANCE RQExist        AS INTEGER                             ' Just see if it exists
   INSTANCE RQForcePrompt  AS INTEGER                             ' Prompt even if FN is populated
   INSTANCE RQInput        AS INTEGER                             ' OPEN for Input
   INSTANCE RQOutPut       AS INTEGER                             ' OPEN for Output
   INSTANCE RQParse        AS INTEGER                             ' Parse (Get FD data) (Implied with IBEOR)
   INSTANCE RQPrompt       AS INTEGER                             ' Prompt if needed
   INSTANCE RQROTest       AS INTEGER                             ' Test Read Only
   INSTANCE RQTemp         AS INTEGER                             ' Asign a Temp filename
   INSTANCE RSExistFail    AS INTEGER                             ' Doesn't exist
   INSTANCE RSROnlyFail    AS INTEGER                             ' RO Failure
   INSTANCE RSTabFound     AS INTEGER                             ' A TabFound was set
   INSTANCE RSUserCancel   AS INTEGER                             ' Cancelled by user

   CLASS METHOD CREATE()                                          ' Constructor - Initialize Class stuff
   LOCAL i AS LONG
   END METHOD                                                     '

   INTERFACE iIO: INHERIT IUNKNOWN                                ' Define the interface

      gProp(FilePath, STRING)
      gProp(FNum, LONG)
      gProp(ResultMsg, STRING)
      gProp(RSExistFail, INTEGER)
      gProp(RSROnlyFail, INTEGER)
      gProp(RSTabFound, INTEGER)
      gProp(RSUserCancel, INTEGER)
      gProp(TabFound, LONG)

      METHOD Setup(flags AS STRING, prof AS STRING, prompt AS STRING, fname AS STRING)
      '----- Setup the IO request
         IF flags <> "*" THEN                                     ' If not 'Leave alone'
            FNum          = FREEFILE                              ' Assign an FNum to use
            RQAlready     = IIF(INSTR(flags, "A"), %True, %False) ' Set flags based on the request
            RQBinary      = IIF(INSTR(flags, "B"), %True, %False) '
            RQCreateIf    = IIF(INSTR(flags, "C"), %True, %False) '
            RQExist       = IIF(INSTR(flags, "E"), %True, %False) '
            RQForcePrompt = IIF(INSTR(flags, "F"), %True, %False) '
            RQInput       = IIF(INSTR(flags, "I"), %True, %False) '
            RQOutput      = IIF(INSTR(flags, "O"), %True, %False) '
            RQPrompt      = IIF(INSTR(flags, "P"), %True, %False) '
            RQROTest      = IIF(INSTR(flags, "R"), %True, %False) '
            RQTemp        = IIF(INSTR(flags, "T"), %True, %False) '
            RQParse       = IIF(RQBinary OR RQExist OR RQInput OR RQOutput OR RQROTest, %True, %False)
            RESET RSTabFound, RSUserCancel, RSROnlyFail, RSExistFail, SetProfile
         END IF
         IF prof <> "*" THEN SetProfile = prof                    '
         IF prompt <> "*" THEN PromptTxt = Prompt                 '
         IF fname <> "*" THEN
            FilePath = IIF$(RQTemp, sGetNewTempFile(fname), TRIM$(fname)) ' Set the passed fname
            sUnQuote(FilePath)                                    ' Remove quotes (if any)
         END IF                                                   '
      END METHOD

      METHOD EXEC() AS LONG
      '---------- Exec the requests
      LOCAL lclFN AS STRING

         '----- Do Prompt if needed
         IF RQPrompt THEN                                         ' Prompt if needed?
            IF ISNULL(FilePath) OR RQForcePrompt THEN             ' And a need? or Forced?
               FilePath = sDoOpenFile(PromptTxt)                  ' Go get a filename
               IF ISNULL(FilePath) THEN                           ' Nothing?
                  ResultMsg = "File selection cancelled"          ' Say some words
                  RSUserCancel = %True                            ' Set a flag
                  METHOD = %True: EXIT METHOD                     ' Set fail RC
               END IF                                             '
            END IF                                                '
         END IF                                                   '

         '----- Check if a Shortcut to a file
         IF FilePath <> $Empty THEN                               ' Ignore the (Empty) call
            IF ISNOTNULL(FilePath) AND INSTR(FilePath, "\") = 0 THEN _  ' If we have something with no path
               FilePath = sGetDefDir + FilePath                   ' Add a path if it doesn't have one
            IF IsEQ(RIGHT$(FilePath, 4), ".LNK") THEN             ' Passed a LNK file?
               FilePath = sLNKConvert(FilePath)                   ' Convert to a full path
               IF ISNULL(FilePath) THEN                           ' If null returned, we failed
                  ResultMsg = "Shortcut lookup failed"            ' Say why we failed
                  METHOD = %True: EXIT METHOD                     ' Exit, we're done
               END IF                                             '
            END IF                                                '
         END IF                                                   '

         '----- Make sure file is not a directory
         IF (RQInput OR RQBinary OR RQExist) AND ISFOLDER(FilePath) THEN
            ResultMsg = "Filename is a Folder, not a File"        ' Say some words
            RSExistFail = %True                                   ' Set a flag
            METHOD = %True: EXIT METHOD                           ' Set fail RC
         END IF                                                   '

         '----- Check if open already?
         IF RQAlready AND FilePath <> $Empty THEN                 ' See if it exists?
            TabFound = VAL(sFileQueue("S", " ", FilePath))        ' Returns tab number if open, else zero
            IF TabFound > 0 THEN                                  ' Tab number?
               ResultMsg = "File already open in this tab"        ' Say some words
               RSTabfound = %True                                 ' Say Tabfound is valid
               METHOD = %True: EXIT METHOD                        ' Exit, we're done
            END IF                                                '
         END IF                                                   '

         '----- Check if it exists
         IF RQExist AND FilePath <> $Empty THEN                   ' See if it exists?
            lclFN = PATHSCAN$(FULL, FilePath)                     ' Let PATHSCAN have a go
            IF ISNULL(lclFN) THEN                                 ' Returns null if not found
               IF RQCreateIf THEN                                 ' Create if not present?
                  TRY                                             ' Just in case
                     sMakeNullFile(FilePath)                      ' Create it
                  CATCH                                           '
                     ResultMsg = "File OPEN failed"               ' Say some words
                     METHOD = %True: EXIT METHOD                  ' Exit, we're done
                  END TRY                                         '
               ELSE                                               ' Kill, not present
                  ResultMsg = "File does not exist"               ' Say some words
                  RSExistFail = %True                             ' Set flag
                  METHOD = %True: EXIT METHOD                     ' Exit, we're done
               END IF                                             '
            END IF                                                '
         END IF                                                   '

         IF RQParse THEN lclFN = DIR$(FilePath, TO FD)            ' If Parse then get all the DIR info

         '----- Do the READ ONLY check
         IF RQROTest THEN                                         ' Test for RO?
            IF (FD.FileAttributes AND %FILE_ATTRIBUTE_READONLY) = %FILE_ATTRIBUTE_READONLY THEN ' Test it
               ResultMsg = "File is a READ ONLY file"             ' Say some words
               RSROnlyFail = %True                                '
               METHOD = %True: EXIT METHOD                        ' Exit, we're done
            END IF                                                '
         END IF                                                   '

         '----- Open for Input
         IF RQInput AND FilePath <> $Empty THEN                   ' Open for Input?
            TRY                                                   ' Just in case
               OPEN FilePath FOR INPUT ACCESS READ LOCK SHARED AS # FNum ' Open the File
            CATCH                                                 '
               ResultMsg = "File OPEN failed"                     ' Say some words
               METHOD = %True: EXIT METHOD                        ' Exit, we're done
            END TRY                                               '
         END IF                                                   '

         '----- Open for Output
         IF RQOutput AND FilePath <> $Empty THEN                  ' Open for Output?
            TRY                                                   ' Just in case
               '----- Next see if it is 'In Use' elsewhere
               TRY                                                ' See if file is in use
                  OPEN FilePath FOR BINARY ACCESS READ WRITE LOCK READ WRITE AS # FNum
                  CLOSE # FNum                                    '
               CATCH
                  ResultMsg = "File is in use or protected"       ' Say some words
                  METHOD = %True: EXIT METHOD                     ' Exit, we're done
               END TRY                                            '

               '----- Now try the real normal OPEN for output
               OPEN FilePath FOR OUTPUT ACCESS WRITE LOCK WRITE AS # FNum ' Open the File
            CATCH                                                 '
               ResultMsg = "File OPEN failed"                     ' Say some words
               METHOD = %True: EXIT METHOD                        ' Exit, we're done
            END TRY                                               '
         END IF                                                   '

         '----- Open for Binary
         IF RQBinary AND FilePath <> $Empty THEN                  ' Open for Binary?
            TRY                                                   ' Just in case
               OPEN FilePath FOR BINARY ACCESS READ LOCK SHARED AS # FNum ' Open the File
            CATCH                                                 '
               ResultMsg = "File OPEN failed"                     ' Say some words
               METHOD = %True: EXIT METHOD                        ' Exit, we're done
            END TRY                                               '
         END IF                                                   '

         METHOD = %False                                          ' Say we're done
      END METHOD                                                  '

      METHOD CLOSE() AS LONG
      '---------- Close the file now
      LOCAL lclFN AS STRING
         '----- Close the file
         TRY                                                      ' Just in case
            IF FNum <> 0 THEN                                     ' As long as we have FNum
               IF RQOutput AND ISFALSE RQBinary THEN SETEOF # FNum'
               CLOSE # FNum                                       ' Close it
               METHOD = %False                                    '
            END IF                                                '
         CATCH                                                    '
            ResultMsg = "File CLOSE failed"                       ' Say some words
            METHOD = %True                                        ' Exit, we're done
         END TRY                                                  '
         FNum = 0                                                 '
      END METHOD

      METHOD TimeDateRefresh()
      LOCAL lclFN AS STRING
         lclFN = DIR$(FilePath, TO FD)                            ' Re-load the FD data
      END METHOD

      METHOD DATE() AS STRING
      '----- Return File Date
         METHOD = IIF$(FilePath <> $Empty, LEFT$(sMakePrettyTime(FD.LastWriteTime), 10), "          ") ' Save file date
      END METHOD

      METHOD TIME() AS STRING
      '----- Return Time
         METHOD = IIF$(FilePath <> $Empty, MID$(sMakePrettyTime(FD.LastWriteTime), 13, 8), "        ") ' Save file time
      END METHOD

      METHOD Drive() AS STRING
      '----- Return Drive
         METHOD = IIF$(MID$(FilePath, 2, 1) = ":", LEFT$(FilePath, 2), "")  ' Extract drive if present
      END METHOD

      METHOD PATH() AS STRING
      '----- Return Path portion
      LOCAL i AS LONG
         i = INSTR(-1, FilePath, "\")                             '
         METHOD = IIF$(i, LEFT$(FilePath, i), sGetDefDir)         '
      END METHOD

      METHOD File() AS STRING
      '----- Return File portion
      LOCAL i AS LONG
         i = INSTR(-1, FilePath, "\")                             '
         METHOD = IIF$(i, MID$(FilePath, i + 1), FilePath)        '
      END METHOD

      METHOD BASE() AS STRING
      '----- Return Base portion
      LOCAL i AS LONG, t AS STRING
         i = INSTR(-1, FilePath, "\")                             '
         t = IIF$(i, MID$(FilePath, i + 1), FilePath)             '
         i = INSTR(-1, t, ".")                                    '
         METHOD = IIF$(i, LEFT$(t, i - 1), t)                     '
      END METHOD

      METHOD EXTN() AS STRING
      '----- Return Extn portion
      LOCAL i AS LONG
         i = INSTR(-1, FilePath, ".")                             '
         METHOD = IIF$(i, MID$(FilePath, i), "")                  '
      END METHOD

      METHOD ROStat() AS LONG
      '----- Return READ ONLY status
         METHOD = IIF((FD.FileAttributes AND %FILE_ATTRIBUTE_READONLY) = %FILE_ATTRIBUTE_READONLY, %True, %False)
      END METHOD

      METHOD PROFILE() AS STRING
      '----- Return Profile Name
         METHOD = IIF$(ISNULL(SetProfile), sParseProfile(FilePath), SetProfile) ' Return SetProfile if provided, else normal fn version
      END METHOD                                                  '

   END INTERFACE

END CLASS
