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

METHOD  FMLCmdAdd(i AS LONG)
'---------- FM Add Favorite line command
LOCAL llist AS STRING
   MEntry
   llist = "Favorite Files"                                       ' Set default
   IF TRIM$(LCmdOps(2)) <> "" THEN llist = LCmdOps(2)             ' Is provided use the provided name
   SELECT CASE AS LONG AFList(i).Flag                             ' Split by entry type
      CASE %FEntry, %FFLEntry, %FDirDown                          ' Normal file or Down Dir
         sFileListAdd(llist, RTRIM$(AFList(i).Path) + RTRIM$(AFList(i).FD.FileName)) ' Go add it

      CASE %FOpen                                                 ' Open files
         IF TRIM$(AFList(i).FD.FileName) = "(Clip)"  OR _         ' Eliminate Special entries
            TRIM$(AFList(i).FD.FileName) = "(SET Edit)"  OR _     '
            TRIM$(AFList(i).FD.FileName) = $Empty  THEN           '
            AFList(i).Msg = "Entry not eligible for Add to FILELIST"
            DoSet(%Msg): MExitMeth                                ' Get msg issued
         END IF                                                   '
         sFileListAdd(llist, TRIM$(AFList(i).Path) + TRIM$(AFList(i).FD.FileName)) ' Go add it
   END SELECT                                                     '
   AFList(i).Msg = "File added to " + UUCASE(llist) + ".FLIST"    '
   AFList(i).Cmd = ""                                             ' Clear Cmd
   DoSet(%LoadData OR %Msg)                                       ' Get msg issued and data refreshed
   MExit                                                          '
END METHOD

METHOD  FMLCmdAll(i AS LONG)
'---------- FM ALL line command
LOCAL fn, mask,  rlist(), tlist(), FMMedit, tstyle, DirPath, DirTypes, RQPath, RQMask, RQFlags, RQNote AS STRING, ix, RC, delctr AS LONG
LOCAL j, hFile AS LONG, hfile2 AS DWORD, fname AS WSTRINGZ * 260
LOCAL CreationTime, LastAccessTime, LastWriteTime AS FILETIME
LOCAL lquad AS QUAD
LOCAL SystemTimeArea AS SYSTEMTIME
LOCAL aIO AS iIO                                                  ' For our I/O stuff
   MEntry
   LET aIO = CLASS "cIO"                                          '

   '----- Only certain entries supported
   IF AFList(i).Flag <> %FFileList THEN MLErrExit(i, "ALL is not supported for this entry type") ' Kill if not a valid FILELIST

   '----- Better have a command operand
   IF TRIM$(LCmdOps(2)) = "" THEN MLErrExit(i, "ALL requires a command operand") ' Kill if no operand

   '----- Load and validate the FILELIST contents
   GOSUB LoadList                                                 ' Go load and validate the contents

   SELECT CASE AS CONST$ UUCASE(TRIM$(LCmdOps(2)))                ' What command have we got?

      CASE "B", "BROWSE", "BRO"                                   ' Browse
         FOR ix = 1 TO UBOUND(tlist())                            ' OK loop actually DOING it
            me.CallTab("BROWSE", $DQ + tlist(ix) + $DQ)           '
         NEXT ix                                                  '
         DoSet(%LoadData)                                         ' Ask for refresh

      CASE "C", "CLONE"                                           ' Clone
         FOR ix = 1 TO UBOUND(tlist())                            ' OK loop actually DOING it
            me.CallTab("CLONE", $DQ + tlist(ix) + $DQ)            '
         NEXT ix                                                  '
         DoSet(%LoadData)                                         ' Ask for refresh

      CASE "D", "DELETE", "DEL", "U", "PURGE", "PUR"              ' Delete
         FOR ix = 1 TO UBOUND(tlist())                            ' First check if not Open elsewhere
            IF VAL(sFileQueue("S", " ", tlist(ix))) <> 0 THEN MLErrExit(i, "File is open elsewhere") ' Open in some tab?
         NEXT ix                                                  '

         tstyle = "Purge: "                                       ' Pick message text
         IF ENV.UseRecycle AND _                                  '
            (IsEQ(TRIM$(LCmdOps(2)), "D") OR _                    '
             IsEQ(TRIM$(LCmdOps(2)), "DEL") OR _                  '
             IsEQ(TRIM$(LCmdOps(2)), "DELETE")) THEN _            '
            tstyle = "Recycle: "  '                               '

         j = 99                                                   ' Set dummy value
         IF ISFALSE ENV.ConfDelFlag THEN                          ' Are we going to individually prompt?
            j = sDoMsgBox("OK to |K" + tstyle + FORMAT$(UBOUND(tlist())) + "|B files", _
                          %MB_YESNO OR %MB_DEFBUTTON2 OR %MB_USERICON, "SPFLite Recycle/Purge")
            IF j = %IDNO THEN MExitMeth                           ' Ignore if told not to
         END IF                                                   '

         FOR ix = 1 TO UBOUND(tlist())                            ' Now actually Do it
            IF ENV.ConfDelFlag AND j <> %IDOK AND j <> %IDYES THEN' Need we prompt user for an OK?
               j = sDoMsgBox("OK to |K" + tstyle + tlist(ix), %MB_YESNOCANCEL OR %MB_DEFBUTTON2 OR %MB_USERICON, "SPFLite Recycle/Purge")
               IF j = %IDNO THEN ITERATE FOR                      ' Ignore if told not to
               IF j = %IDCANCEL THEN MExitMeth                    ' Bail out
            END IF                                                '
            RC = sRecycleBin(tlist(ix), IIF$(LEFT$(tstyle, 1) = "R", "D", "U"))
         NEXT ix                                                  '
         DoSet(%LoadData)                                         ' Refresh Data
         MLErrExit(i, FORMAT$(delctr) + " files deleted")         ' Issue message and exit

      CASE "E", "EDIT"                                            ' Edit
         FOR ix = 1 TO UBOUND(tlist())                            ' OK loop actually DOING it
            me.CallTab("EDIT", $DQ + tlist(ix) + $DQ)             '
         NEXT ix                                                  '
         DoSet(%LoadData)                                         ' Ask for refresh

      CASE "M", "MEDIT"                                           ' Multi-Edit
         FOR ix = 1 TO UBOUND(tlist())                            ' OK loop actually DOING it
            IF ISNULL(MEditCmd) THEN                              ' First M command?
               MEditCmd += $DQ + tlist(ix) + $DQ + " "            '
               MEditCmd += " NEW "                                '
            ELSE                                                  ' Else just add a filename on
               MEditCmd += $DQ + tlist(ix) + $DQ + " "            '
            END IF                                                '
         NEXT ix                                                  '
         me.CallTab("MEDIT", MEditCmd)                            '
         DoSet(%LoadData)                                         ' Ask for refresh

      CASE "O", "OPEN"                                            ' OPEN Edit
         FOR ix = 1 TO UBOUND(tlist())                            ' OK loop actually DOING it
            me.CallTab("OPEN", $DQ + tlist(ix) + $DQ)             '
         NEXT ix                                                  '
         DoSet(%LoadData)                                         ' Ask for refresh

      CASE "OV", "OPENV"                                          ' OPEN View
         FOR ix = 1 TO UBOUND(tlist())                            ' OK loop actually DOING it
            me.CallTab("OPENV", $DQ + tlist(ix) + $DQ)            '
         NEXT ix                                                  '
         DoSet(%LoadData)                                         ' Ask for refresh

      CASE "OB", "OPENB"                                          ' OPEN Browse
         FOR ix = 1 TO UBOUND(tlist())                            ' OK loop actually DOING it
            me.CallTab("OPENB", $DQ + tlist(ix) + $DQ)            '
         NEXT ix                                                  '
         DoSet(%LoadData)                                         ' Ask for refresh

      CASE "P", "PRINT"                                           ' Print
         FOR ix = 1 TO UBOUND(tlist())                            ' OK loop actually DOING it
            me.FMLCmdPrintIt($DQ + tlist(ix) + $DQ, 0)            '
         NEXT ix                                                  '

      CASE "T", "TOUCH"                                           ' Touch
         FOR ix = 1 TO UBOUND(tlist())                            ' OK loop actually DOING it
            ON ERROR GOTO TchError                                ' Ignore errors
            hFile = FREEFILE                                      ' Touch the file now
            OPEN tlist(ix) FOR APPEND ACCESS READ WRITE LOCK SHARED AS #hFile
            hFile2 = FILEATTR(hFile,2)                            ' Get Windows handle for the file
            GetFileTime hFile2, CreationTime, LastAccessTime, LastWriteTime' Fetch timestamps
            GetSystemTime SystemTimeArea                          ' Get the current time
            SystemtimeToFileTime SystemTimeArea, LastWriteTime    ' Put it in LastWriteTime
            SetFileTime hFile2, CreationTime, LastAccessTime, LastWriteTime
            TchExit:                                              '
            CLOSE #hFile                                          ' Close it
         NEXT ix                                                  '
         DoSet(%LoadData)                                         ' Ask for refresh
         MLErrExit(i, "All files Touched")

      CASE "V", "VIEW"                                            ' View
         FOR ix = 1 TO UBOUND(tlist())                            ' OK loop actually DOING it
            me.CallTab("VIEW", $DQ + tlist(ix) + $DQ)             '
         NEXT ix                                                  '
         DoSet(%LoadData)                                         ' Ask for refresh

      CASE ELSE
         MLErrExit(i, "Unsupported ALL command: " + UUCASE(TRIM$(LCmdOps(2))))
   END SELECT                                                     '
   MExitMeth                                                      '

TchError:
   RESUME TchExit

LoadList:
   '----- Load FILELIST and validate the contents
   fn = ENV.FileListPath + TRIM$(AFList(i).FD.FileName)           ' Build the filename
   IF ISFALSE ISFILE(fn) THEN MLErrExit(i, fn  + " was not found")' Oops, not there?
   '----- File exists, open it
   aIO.Setup("I", "", "", fn)                                     ' Tell IO what we're opening
   IF aIO.EXEC THEN _                                             '
      MErrExit(%eFail, aIO.ResultMsg)                             ' Oops?  Bail out

   '----- Get # records, DIM our table and load the file
   FILESCAN # aIO.FNum, RECORDS TO ix                             ' Get the number of records
   IF ix > 0 THEN                                                 ' Some records?
      REDIM rlist(1 TO ix) AS STRING                              ' Redim array to match save data
      LINE INPUT # aIO.FNum, rlist() TO ix                        ' Read it all
   END IF                                                         '

   '----- Close the file now
   aIO.Close                                                      ' Close the File

   '----- Do various validation stuff
   IF ix = 0 THEN MLErrExit(i, "The FILELIST was empty")          ' Nothing?

   '----- Process the loaded entries
   REDIM tlist(1 TO 1) AS STRING                                  ' Reset tlist
   tlist(1) = ""                                                  ' Null 1st entry
   ix = 0                                                         ' Reset index
   FOR j = 1 TO UBOUND(rlist())                                   ' Loop through the loaded FILELIST entries
      me.RQSplit(rlist(j), RQPath, RQMask, RQFlags, RQNote)       ' Split it up
      IF RIGHT$(RQPath, 1) <> "\" THEN                            ' A simple file name?
         IF ISFILE(RQPath) THEN                                   ' File still exist?
            INCR ix                                               ' Count found entry
            REDIM PRESERVE tlist(1 TO ix) AS STRING               ' Resize table
            tlist(ix) = RQPath                                    ' Save this filename
         ELSE                                                     '
            MLErrExit(i, "FILELIST entry: " + RQPath + " no longer exists")
         END IF                                                   '

      ELSE                                                        ' Must be a path entry
         DirPath = RQPath                                         ' Save Dir path
         Fn = DIR$(BUILD$(DirPath, "*"), 22)                      ' Search for 1st * filename  (22 = Normal, hidden, system, Dirs)
         DO WHILE ISNOTNULL(Fn)                                   ' Loop
            IF ISTRUE me.TestMask(TRIM$(fn), RQMask) THEN         ' If we pass the Mask test
               INCR ix                                            ' Count found entry
               REDIM PRESERVE tlist(1 TO ix) AS STRING            ' Resize table
               tlist(ix) = DirPath + fn                           ' Save this filename
            END IF                                                '
            Fn = DIR$(NEXT)                                       '
         LOOP                                                     '
      END IF                                                      '
   NEXT j                                                         '
   RETURN                                                         ' We're OK if we made it to here

END METHOD

METHOD   FMLCmdBrowse(i AS LONG)
'---------- FM Browse line command
LOCAL SaveProf, p AS STRING
   MEntry
   SELECT CASE AS LONG AFList(i).Flag                             ' Split by entry type
      CASE %FEntry, %FFLEntry, %FFileList                         ' Existing File?
         me.DoBROWSE(i)                                           ' Go do Browse
         DoSet(%LoadData)                                         ' Ask for refresh

      CASE %FDirDown                                              ' Dir Down
         me.DoDirDown(i)                                          ' Go do it

      CASE %FDirUp                                                ' Dir Up
         me.DoDirUp(i)                                            ' Go do it

      CASE %FProfile                                              ' A Profile?
         IF ISFALSE ISFILE(ENV.PROFPath + AFList(i).FD.FileName) THEN _  ' See if it exists
            MLErrExit(i, "Specified Profile name does not exist") '
         SaveProf = IIF$(Prf.ProfName <> "", Prf.ProfName, "DEFAULT") ' Save where we are
         p = TRIM$(AFList(i).FD.FileName)                         ' Get Profile name
         p = LEFT$(p, LEN(p) - 4)                                 ' Strip off .INI
         Prf.SetProfName(p, %False)                               ' Setup
         Prf.ReadAll(%True)                                       '
         Prf.EditProf                                             ' Let user Have a look
         Prf.SetProfName(SaveProf, %False)                        '
         Prf.ReadAll(%True)                                       '
         CONTROL SET FOCUS hWnd, %IDC_SPFLiteTAB                  ' Set focus

   END SELECT                                                     '
   AFList(i).Cmd = " "                                            ' Blank the command area
   MExit                                                          '
END METHOD

METHOD  FMLCmdCancel(i AS LONG)
'---------- FM CANCEL line command
LOCAL MSG AS STRING, j AS LONG
   MEntry
   IF AFList(i).Flag <> %FOpen THEN MLErrExit(i, "CANCEL only usable against OPEN files") ' An OPEN entry?
   sCallTabCmd(AFList(i).Uniq, "CANCEL ")                         '
   DoSet(%LoadData)                                               ' Ask for refresh
   AFList(i).Cmd = " "                                            ' Blank the command area
   MExit
END METHOD

METHOD  FMLCmdClone(i AS LONG)
'---------- FM Clone line command
LOCAL lcmd, cmd, SetProf, fn1 AS STRING
   MEntry
   lcmd = AFlist(i).Cmd                                           ' Save command string
   WHILE lcmd <> ""                                               ' Process lcmd for .profile name
      cmd = sGetWord(lcmd, %Strip, %QuoteNotSig)                  ' Get a word from the command
      IF LEFT$(cmd, 1) = "." THEN SetProf = UUCASE(cmd)           '
   WEND                                                           '
   AFList(i).Cmd = " "                                            ' Blank the command area

   SELECT CASE AS LONG AFList(i).Flag                             ' Split off
      CASE %FEntry, %FFLEntry, %FFileList, %FProfile              ' Normal type File
         cmd = $DQ + RTRIM$(AFList(i).Path) + RTRIM$(AFList(i).FD.FileName) + $DQ + " " + SetProf

      CASE %FOpen                                                 ' Open file list
         IF TRIM$(AFList(i).FD.FileName) = "(Clip)" OR _          ' Eliminate Special entries
            TRIM$(AFList(i).FD.FileName) = "(SET Edit)" OR _      '
            TRIM$(AFList(i).FD.FileName) = $Empty THEN MExitMeth
         cmd = $DQ + RTRIM$(AFList(i).Path) + RTRIM$(AFList(i).FD.FileName) + $DQ + " " + SetProf

      CASE  ELSE                                                  '
         MLErrExit(i, "CLONE is not supported for this entry type") ' Get msg issued
   END SELECT                                                     '
   me.CallTab("CLONE", cmd)                                       '
   MExit
END METHOD

METHOD  FMLCmdDelete(i AS LONG, which AS STRING)
'---------- FM Delete line command
LOCAL lclPath, lclDrive, DelDrive, DelPath, DelFn, Errmsg, tstyle, tdir, lfPath, lfFile AS STRING
LOCAL j, RC AS LONG

   MEntry
   SELECT CASE AS LONG AFList(i).Flag                             ' Split by entry type
      CASE %FEntry, %FFLEntry, %FFileList, %FDirDown, %FProfile   ' Normal type file
      CASE ELSE                                                   ' Ignore all others
         MLErrExit(i, "Can't Delete this entry type")             ' Get msg issued
   END SELECT

   IF (AFList(i).FD.FileAttributes AND %FILE_ATTRIBUTE_READONLY) = %FILE_ATTRIBUTE_READONLY THEN _   ' If RO
      MLErrExit(i, "Can't Delete a Read-Only file")               ' Get msg issued

   lfPath = TRIM$(AFList(i).Path)                                 ' Simplify coding
   lfFile = TRIM$(AFList(i).FD.FileName)                          '

   lclPath = CURDIR$                                              ' Get current path
   IF MID$(lclPath, 2, 1) = ":" THEN _                            ' Extract Drive if present
      lclDrive = LEFT$(lclPath, 2)                                ' and save it
   IF MID$(lfPath, 2, 1) = ":" THEN                               ' Extract Drive if present
      DelDrive = LEFT$(lfPath, 2)                                 ' and save it
      DelPath = MID$(lfPath, 3)                                   '
      IF ISNULL(DelPath) THEN DelPath = "\"                       ' In case we're at the root
   END IF                                                         '
   DelFn = lfPath + lfFile                                        '
   IF RIGHT$(TRIM$(DelFn), 1) = "\" THEN                          ' Is this the DIR name?
      DelFn = LEFT$(DelFn, LEN(TRIM$(DelFn)) - 1)                 ' Remove the last \
   END IF                                                         '
   IF VAL(sFileQueue("S", " ", lfPath + lfFile)) <> 0 THEN MLErrExit(i, "File open elsewhere: " + Delfn)  ' Open in some tab?
   IF AFlist(i).Flag <> %FDirDown THEN                            ' Not a Downward Directory?
      IF ISFALSE ISFILE(DelFn) THEN MLErrExit(i, "File no longer exists: " + DelFn)
   END IF                                                         '

   IF ENV.ConfDelFlag THEN                                        ' Need we prompt user for an OK?
      tstyle = "Purge: "                                          ' Pick message text
      IF ENV.UseRecycle AND LEFT$(which, 1) = "D" THEN tstyle = "Recycle: " '
      j = sDoMsgBox("OK to |K" + tstyle + DelFn, %MB_YESNO + %MB_DEFBUTTON2 + %MB_USERICON, "SPFLite Recycle/Purge")
      IF j = %IDYES THEN                                          ' OK?
         GOSUB DoADel                                             ' Actually Do the Delete
      ELSE                                                        '
         MExitMeth                                                '
      END IF                                                      '
   ELSE                                                           '
      GOSUB DoADel                                                ' Actually Do the Delete
   END IF                                                         '
   IF DelDrive <> lclDrive THEN CHDRIVE lclDrive                  ' Switch drive if needed
   CHDIR lclPath                                                  ' put back the original path

   IF RC = 75 THEN                                                '
      MLErrExit(i, DelFn + " is not Empty")                       '
   ELSEIF RC <> 0 THEN                                            '
      MLErrExit(i, "Recycle/Purge of: " + DelFn + " unsuccessful")'
   END IF                                                         '

   IF ISNOTNULL(FileListNm) AND FileListNm <> "Named Favorites" THEN  ' A real FILELIST entry
      sFileListDel(FileListNm, TRIM$(AFList(i).Path) + TRIM$(AFList(i).FD.FileName)) ' Remove from FILELIST
   END IF                                                         '
   AFList(i).Cmd = " "                                            ' Blank the command area
   MExitMeth

   DoADel:
      IF DelDrive <> lclDrive THEN CHDRIVE DelDrive               '
      CHDIR DelPath                                               '
      IF AFList(i).Flag <> %FDirDown THEN                         ' Not a Downward Directory?
         RC = sRecycleBin(DelFn, LEFT$(which, 1))                 '
      ELSE                                                        '
         RMDIR DelFn                                              ' Dir DEL
         RC = ERR                                                 '
      END IF
      DoSet(%LoadData)                                            ' Ask for refresh
      RETURN
END METHOD

METHOD  FMLCmdDir(i AS LONG)
'---------- Switch to this file's DIR Folder
   MEntry
   AFList(i).Cmd = " "                                            ' Blank the command area
   SELECT CASE AS LONG AFList(i).Flag                             ' Split by type
      CASE %FEntry, %FFLEntry, %FOpen                             ' Normal file entry
         FMask = "*"                                              ' Set the FileMask for FM
         FPath = RTRIM$(AFList(i).Path)                           ' Set the path
         FileListNm = ""                                          ' Null any FILELIST setting
         sIniSetString("FManager", "DefDir1", FPath)              '
         sIniSetString("FManager", "DefMask", FMask)              '
         sIniSetString("FManager", "Recent", FileListNm)          '
         DoSet(%LoadReq)                                          ' Full refresh
      CASE ELSE                                                   '
         MLErrExit(i, "DIR is only for normal File entries")      '
   END SELECT                                                     '
   MExit                                                          '
END METHOD

METHOD   FMLCmdEdit(i AS LONG)
'---------- FM Edit line command
LOCAL SaveProf, p AS STRING
   MEntry
   SELECT CASE AS LONG AFList(i).Flag                             ' Split by entry type
      CASE %FEntry, %FFLEntry, %FFileList                         ' Existing File?
         me.DoEDIT(i)                                             ' Go do Edit

      CASE %FDirDown                                              ' Dir Down
         me.DoDirDown(i)                                          ' Go do it

      CASE %FDirUp                                                ' Dir Up
         me.DoDirUp(i)                                            ' Go do it

      CASE %FProfile                                              ' Profile
         SaveProf = IIF$(Prf.ProfName <> "", Prf.ProfName, "DEFAULT") ' Save where we are
         p = TRIM$(AFList(i).FD.FileName)                         ' Get Profile name
         p = LEFT$(p, LEN(p) - 4)                                 ' Strip off .INI
         Prf.SetProfName(p, %False)                               ' Setup
         Prf.ReadAll(%True)                                       '
         Prf.EditProf                                             ' Let user Have a look
         Prf.SetProfName(SaveProf, %False)                        '
         Prf.ReadAll(%True)                                       '
         CONTROL SET FOCUS hWnd, %IDC_SPFLiteTAB                  ' Set focus

   END SELECT                                                     '
   AFList(i).Cmd = " "                                            ' Blank the command area
   DoSet(%LoadData)                                               ' Ask for refresh
   MExit                                                          '
END METHOD

METHOD  FMLCmdEmpty()
'---------- Start an Empty File
LOCAL tFn, t AS STRING, j, FNum, ForceBrowse AS LONG, lFD AS DIRDATA
   MEntry
   CmdEmpty = " "                                                 ' Clear command char
   me.CallTab("EDIT", "NEW")                                      '
   DoSet(%LoadData)                                               ' Ask for refresh
   MExit                                                          '
END METHOD

METHOD  FMLCmdEnd(i AS LONG)
'---------- FM END line command
   MEntry
   IF AFList(i).Flag <> %FOpen THEN MLErrExit(i, "END only usable against OPEN files") ' An OPEN entry?
   sCallTabCmd(AFList(i).Uniq, "END ")                            ' Call END
   DoSet(%LoadData)                                               ' Ask for refresh
   AFList(i).Cmd = " "                                            ' Blank the command area
   MExit
END METHOD

METHOD  FMLCmdExclude(i AS LONG)
'---------- FM Exclude line command
LOCAL fn AS STRING, FNum AS LONG
   MEntry                                                         '
   IF AFList(i).Flag = %FEntry OR AFList(i).Flag = %FFLEntry OR AFList(i).Flag = %FFileList THEN
      AFList(i).Cmd = " "                                         ' Clear command char
      me.IncRXSize                                                ' Allocate an RX slot
      RXList(RXCount) = TRIM$(AFList(i).Path) + TRIM$(AFList(i).FD.FileName) ' Create an Exclude entry
      DoSet(%LoadData)                                            ' Have it take effect
   ELSE                                                           '
      MLErrExit(i, "Exclude cannot be used on this Entry type")   '
   END IF                                                         '
   AFList(i).Cmd = " "                                            ' Blank the command area
   MExit                                                          '
END METHOD

METHOD  FMLCmdFavorite()
'---------- FM Favorite line command
   MEntry                                                         '
   CmdFavorite = " "                                              ' Clear command char
   FileListNm = "Favorite Files"                                  ' Switch to it
   NestCtr = 0                                                    ' Kill nesting then
   me.WindowTitle                                                 ' Re-do Window Title
   DoSet(%LoadReq)                                                '
   MExit                                                          '
END METHOD

METHOD  FMLCmdFileList()
'---------- FM FileList lineand
   MEntry                                                         '
   CmdFileList = " "                                              ' Clear command char
   FileListNm = "Named Favorites"                                 ' Switch to it
   NestCtr = 0                                                    ' Kill nesting then
   me.WindowTitle                                                 ' Re-do Window Title
   DoSet(%LoadData)                                               ' Ask for refresh
   MExit                                                          '
END METHOD

METHOD  FMLCmdFilePath()
'---------- Return to Normal FilePath display
LOCAL tFn, t AS STRING, j, FNum, ForceBrowse AS LONG, lFD AS DIRDATA
   MEntry
   CmdFilePath = " "                                              ' Clear command char
   FileListNm = ""                                                ' Clear any FileListNm
   NestCtr = 0                                                    ' Kill nesting then
   me.WindowTitle                                                 ' Re-do Window Title
   DoSet(%LoadData)                                               ' Ask for refresh
   MExit                                                          '
END METHOD

METHOD  FMLCmdForget(i AS LONG)
'---------- FM Forget line command
LOCAL j AS LONG, fn, RQPath, RQMask, RQFlags, RQNote AS STRING
   MEntry
   IF FMode <> %FMFilelist THEN MLErrExit(i, "FORGET can only be used with FILELIST displays")   '
   fn = TRIM$(AFList(i).Path) + TRIM$(AFList(i).FD.FileName)      ' Extract the filename
   me.IncRFSize                                                   ' Allocate an RF slot
   RFList(RFCount) = fn                                           ' Create a Forget entry

   '----- See if it's a specific name in the Request list
   FOR j = 1 TO RQCount                                           ' Spin through RQ table
      me.RQSplit(RQList(j), RQPath, RQMask, RQFlags, RQNote)      ' Split operands
      IF RIGHT$(RQPath, 1) = "\" THEN ITERATE FOR                 ' If a Path entry, ignore it
      IF IsEQ(fn, RQPath) THEN                                    ' We have a detailed filename request
         '----- Add F to the requests flags
         IF INSTR(RQFlags, "F") = 0 THEN RQFlags += "F"           ' Add an F flag if not already there
         RQList(j) = BUILD$(fn, "|", RQMask, "|", RQFlags, "|", RQNote)
         DoSet(%SaveReq + %LoadData)                              ' Ask for FILELIST to be saved
         MExitMeth                                                '
      END IF                                                      '
   NEXT j                                                         '

   '----- Not in the Request list, must have been added as a directory search
   me.IncRQSize                                                   ' Allocate an RQ slot
   RQList(RQCount) = fn + "|*|GF|"                                ' Add a generic FORGET entry
   DoSet(%SaveReq OR %LoadData)                                   ' Ask for FILELIST to be saved
   AFList(i).Cmd = " "                                            ' Blank the command area
   MExit                                                          '
END METHOD

METHOD  FMLCmdFound()
'---------- FM Found line command
   MEntry                                                         '
   CmdFound = " "                                                 ' Clear command char
   FileListNm = "Found Files"                                     ' Switch to it
   NestCtr = 0                                                    ' Kill nesting then
   me.WindowTitle                                                 ' Re-do Window Title
   DoSet(%LoadData)                                               ' Ask for refresh
   MExit                                                          '
END METHOD

METHOD  FMLCmdLines(i AS LONG)
'---------- FM LINES line command
LOCAL fn AS STRING, FNum, SaveTop AS LONG
   AFList(i).Cmd = " "                                            ' Blank the command area
   SELECT CASE AS LONG AFList(i).Flag                             ' Split off

      CASE %FEntry, %FFLEntry                                     ' Eligible File?
         '----- Look at it's data
         fn = RTRIM$(AFList(i).Path) + RTRIM$(AFList(i).FD.FileName)' Get the filename
         SaveTop = TopScrn                                        ' Save TopScrn around LInitTxtData
         me.LInitTxtData(fn)                                      ' Initialize our Text area
         TIP.Setup("E", "", "", fn)                               ' Set filename in TIP area
         TIP.EXEC                                                 ' Go do it
         me.InitaFile(%True)                                      ' Initialize file stuff
         TopScrn = SaveTop                                        ' Save TopScrn around LInitTxtData
         IF Prf.PState = %StateOff THEN                           ' If STATE OFF for this filetype
            AFList(i).Msg = "File has PROFILE STATE OFF, LINES skipped"
            DoSet(%Msg)                                           ' Get msg issued
         ELSE                                                     '
            me.StateSave(fn, 1, LastLine, %False)                 ' Go setup/refresh STATE data
            AFList(i).LinesInt = sGetLines(fn)                    ' Get Line count from STATE
            AFList(i).Msg = "File STATE information updated."     '
            DoSet(%Msg)                                           ' Get msg issued
         END IF                                                   '
         sFileQueue("D", " ", TIP.FilePath)                       ' Remove from Open queue
      CASE ELSE                                                   '
         AFList(i).Msg = "LINES cannot be used on this type of entry"
         DoSet(%Msg)                                              ' Get msg issued
   END SELECT                                                     '
END METHOD

METHOD  FMLCmdMEdit(i AS LONG)
'---------- FM MEdit line command
LOCAL fn AS STRING, FNum AS LONG
   MEntry                                                         '
   AFList(i).Cmd = " "                                            ' Clear command char
   SELECT CASE AS LONG AFList(i).Flag                             ' Split off
      CASE %FEntry, %FFLEntry                                     ' A valid entry?
         IF ISNULL(MEditCmd) THEN                                 ' First M command?
            MEditCmd += $DQ + RTRIM$(AFList(i).Path) + RTRIM$(AFList(i).FD.FileName) + $DQ + " "
            IF LastLine > 2 THEN MEditCmd += " NEW "              ' If we can't use this tab, add NEW to the command
         ELSE                                                     ' Else just add a filename on
            MEditCmd += $DQ + RTRIM$(AFList(i).Path) + RTRIM$(AFList(i).FD.FileName) + $DQ + " "
         END IF                                                   '
      CASE ELSE                                                   '
   END SELECT                                                     '
   DoSet(%LoadData)                                               ' Ask for refresh
   MExit                                                          '
END METHOD

METHOD  FMLCmdNorm(i AS LONG)
'---------- FM Normalize line command
LOCAL tfn, fn, ofile AS STRING, FD AS DIRDATA, FNum, j, k, fremoved, premoved, altered AS LONG, tlist() AS STRING
LOCAL RQPath, RQMask, RQFlags, RQNote AS STRING
LOCAL fmvIO AS iIO                                                ' For our I/O stuff
   MEntry
   LET fmvIO = CLASS "cIO"                                        '
   AFList(i).Cmd = " "                                            ' Blank the command area
   SELECT CASE AS LONG AFList(i).Flag                             ' Split off

      CASE %FFileList                                             ' A FILELIST?
         '----- Look at it's data
         fn = TRIM$(AFList(i).Path) + TRIM$(AFList(i).FD.FileName)' Get the filename
         fmvIO.Setup("IE", "", "", fn)                            ' Tell IO what we're opening
         fmvIO.EXEC                                               '
         FILESCAN # fmvIO.FNum, RECORDS TO k                      ' Get the number of records
         IF k > 0 THEN                                            ' Some records?
            REDIM tlist(1 TO k) AS STRING                         ' Redim array to match save data
            LINE INPUT # fmvIO.FNum, tlist() TO j                 ' Read it all
         END IF                                                   '
         fmvIO.Close                                              ' Close it

         '----- Scan the lines in the file for our verify
         IF j THEN                                                ' Did we get some records?
            FOR k = j TO 1 STEP -1                                ' OK, let's scan the records
               IF TRIM$(tlist(k)) = "" THEN                       ' Null record?
                  ARRAY DELETE tlist(k)                           ' Remove it
                  DECR j: altered = %True                         ' Reduce count, remember we altered it
                  ITERATE FOR                                     '
               END IF                                             '
               IF INSTR(tlist(k), "|") = 0 THEN                   ' If not yet in the new | format
                  REPLACE ANY "," WITH "|" IN tlist(k)            ' Swap commas to |
               END IF                                             '

               me.RQSplit(tlist(k), RQPath, RQMask, RQFlags, RQNote) ' Split out operands
               IF RIGHT$(RQPath, 1) = "\" THEN                    ' Is this a Path request?
                  IF ISFALSE ISFOLDER(RQPath) THEN                ' If it doesn't still exist?
                     ARRAY DELETE tlist(k)                        ' Remove it
                     DECR j: altered = %True                      ' Reduce count, remember we altered it
                     INCR premoved                                ' Count paths removed
                  END IF                                          '

               ELSE                                               ' This is a specific file entry
                  IF ISFALSE ISFILE(RQPath) OR INSTR(RQFlags, "F") THEN ' If it doesn't still exist or was Forgotten?
                     ARRAY DELETE tlist(k)                        ' Remove it
                     DECR j: altered = %True                      ' Reduce count, remember we altered it
                     INCR fremoved                                ' Count files removed
                  END IF                                          '
               END IF                                             '
            NEXT k                                                '
         END IF                                                   '

         '----- Re-write the FILELIST is it still contains something
         IF j > 0 AND altered THEN                                '
            fmvIO.Setup("O", "", "", fn)                          ' Tell IO what we're opening
            fmvIO.EXEC                                            '
            FOR k = 1 TO j                                        ' Write things
               IF TRIM$(tlist(k)) <> "" THEN PRINT # fmvIO.FNum, tlist(k) '
            NEXT k                                                '
            fmvIO.Close                                           '
            IF premoved = 0 AND fremoved = 0 THEN                 ' Choose message to issue
               AFList(i).Msg = "All entries normalized"           '
            ELSEIF premoved > 0 AND fremoved = 0 THEN             '
               AFList(i).Msg = FORMAT$(premoved) + " empty path request(s) removed"
            ELSEIF premoved = 0 AND fremoved > 0 THEN             '
               AFList(i).Msg = FORMAT$(fremoved) + " empty/forgotten filename request(s) removed"
            ELSE                                                  '
               AFList(i).Msg = FORMAT$(premoved) + " empty path and " + FORMAT$(fremoved) + " empty/forgotten file request(s) removed"
            END IF                                                '
            DoSet(%Msg)                                           ' Get msg issued

         ELSEIF j = 0 THEN                                        ' FILELIST is empty, delete it
            sRecycleBin(fn, "D")                                  ' Delete empty files
            AFList(i).Msg = "No valid entries remain, the FLIST will be removed"
            DoSet(%Msg)                                           ' Get msg issued
         ELSE
            AFList(i).Msg = "All FLIST entries validated, no changes were made"
            DoSet(%Msg)                                           ' Get msg issued
         END IF                                                   '
         DoSet(%LoadData)                                         ' Refresh the display

      CASE ELSE                                                   '
         AFList(i).Msg = "NORM cannot be used on this type of entry"
         DoSet(%Msg)                                              ' Get msg issued
   END SELECT                                                     '

   MExit                                                          '
END METHOD

METHOD  FMLCmdODel(i AS LONG, dtype AS STRING)
'---------- FM OPEN file DELETE command
LOCAL MSG, cmd AS STRING, j AS LONG
   MEntry
   TP = Tabs(AFList(i).Uniq)                                      ' Swap the TP pointer
   IF TIP.ROStat THEN                                             ' RO file?
      TP = Tabs(1)                                                ' Return to FM
      MLErrExit(i, dtype + " cannot be used on a Read Only file") '
   END IF                                                         '

   IF IsMEdit THEN                                                ' MEdit?
      TP = Tabs(1)                                                ' Return to FM
      MLErrExit(i, dtype + " cannot be used for a MEdit session") '
   END IF                                                         '

   IF dtype = "DELETE" OR dtype = "DEL" OR dtype = "D" THEN       ' Choose command flavor
      cmd = "DEL "                                                '
   ELSEIF dtype = "PURGE" OR dtype = "PUR" OR dtype = "U" THEN    '
      cmd = "PURGE "                                              '
   ELSE                                                           '
      cmd = ""                                                    '
   END IF                                                         '
   me.CallTab("CANCEL", cmd)                                      ' Call it
   DoSet(%LoadData)                                               ' Ask for refresh
   AFList(i).Cmd = " "                                            ' Blank the command area
   MExit
END METHOD

METHOD  FMLCmdOpen()
'---------- FM Open line command
   MEntry                                                         '
   CmdOpen = " "                                                  ' Clear command char
   FileListNm = "Open Files"                                      ' Switch to it
   NestCtr = 0                                                    ' Kill nesting then
   me.WindowTitle                                                 ' Re-do Window Title
   DoSet(%LoadData)                                               ' Ask for refresh
   MExit                                                          '
END METHOD

METHOD   FMLCmdOpenF(i AS LONG)
'---------- FM Open File line command
LOCAL SaveProf, p, lclCmd, lclOper AS STRING
   MEntry
   SELECT CASE AS LONG AFList(i).Flag                             ' Split by entry type
      CASE %FEntry, %FFLEntry, %FFileList                         ' Existing File?
         lclCmd = UUCASE(TRIM$(AFList(i).Cmd))                    ' Get original command
         IF lclCmd = "O" OR lclCmd = "OPEN" THEN                  ' Setup type
            lclOper = "OPEN"                                      '
         ELSEIF lclCmd = "OB" OR lclCmd = "OPENB" THEN            '
            lclOper = "OPENB"                                     '
         ELSEIF lclCmd = "OV" OR lclCmd = "OPENV" THEN            '
            lclOper = "OPENV"                                     '
         END IF
         me.DoOpen(i, lclOper)                                    ' Go do Open
         DoSet(%LoadData)                                         ' Ask for refresh

   END SELECT                                                     '
   AFList(i).Cmd = " "                                            ' Blank the command area
   MExit                                                          '
END METHOD

METHOD  FMLCmdPrint(i AS LONG)
'---------- FM Print line command
LOCAL Fn, Errmsg, t, tFn AS STRING
LOCAL FD AS DIRDATA, j AS LONG
LOCAL SystemTimeArea AS SYSTEMTIME
   MEntry
   AFList(i).Cmd = " "                                            ' Clear command char
   IF AFList(i).Flag = %FEntry OR _                               ' File Entry?
      AFList(i).Flag = %FOpen OR _                                '
      AFList(i).Flag = %FFLEntry THEN                             '
      fn = RTRIM$(AFList(i).Path) + RTRIM$(AFList(i).FD.FileName) ' Get filename
   ELSE                                                           '
      MExitMeth                                                   '
   END IF                                                         '

   me.FMLCmdPrintit(fn, i)                                        ' Do the actual print
   AFList(i).Msg = "File printed"                                 '
   DoSet(%Msg)                                                    ' Get msg issued
   AFList(i).Cmd = " "                                            ' Blank the command area
   MExit
END METHOD

METHOD  FMLCmdPrintIt(fn AS STRING, i AS LONG)
'---------- Print a specified filename
LOCAL j, lctr AS LONG, pLine AS STRING

   '----- Open the Printer
   IF ISFALSE sOpenPrinter("") THEN                               ' Get printer ready if not already
      IF i <> 0 THEN                                              ' Got table index?
         AFList(i).Msg = "OPEN of printer failed"                 '
         DoSet(%Msg)                                              ' Get msg issued
      END IF                                                      '
      MExitMeth                                                   '
   END IF                                                         '
   sPrtPrint(%PRTReset, " ", " ", %False)                         ' Tell sPRTPrint to reset

   '----- Open the file and process it
   TIP.Setup("IE", "", "", fn)                                    ' Tell IO what we're opening
   IF TIP.EXEC THEN                                               '
      IF i <> 0 THEN                                              ' Got an index?
         AFList(i).Msg = TIP.ResultMsg                            '
         DoSet(%Msg)                                              ' Get msg issued
      END IF                                                      '
      MExitMeth                                                   '
   END IF                                                         '

   DO WHILE ISFALSE EOF(TIP.FNum)                                 ' If not EOF
      LINE INPUT # TIP.FNum, PLine                                ' Read a line
      INCR lctr                                                   ' Incr line count
      sPRTPrint(%PRTLine, PLine, $$TxtHi, %False)                 ' Print each line in NONUM mode
      sPRTPrint(%PRTNewLine, " ", " ", %False)                    '
   LOOP                                                           '
   sPRTPrint(%PRTFlushClose, " ", " ", %False)                    ' Tell sPRTPrint to flush page
   TIP.Close                                                      ' Fiddle flags for close

END METHOD

METHOD FMLCmdProcess() AS LONG
'---------- Process the FM Line commands
LOCAL lcmd, Deferred(), t, tt, ttt, tttt AS STRING, i, DefCtr AS LONG
DIM Deferred(1 TO AFCount) AS STRING
   MEntry
   MEditCmd = ""                                                  ' Start MEdit command as null

   '----- First check the Quick entries
   IF CmdEmpty    <> "" AND CmdEmpty    <> " " THEN me.FMLCmdEmpty   : GOTO LookAtLines ' Process the Quick entries
   IF CmdFilePath <> "" AND CmdFilePath <> " " THEN me.FMLCmdFilePath: GOTO LookAtLines '
   IF CmdRecent   <> "" AND CmdRecent   <> " " THEN me.FMLCmdRecent  : GOTO LookAtLines '
   IF CmdFound    <> "" AND CmdFound    <> " " THEN me.FMLCmdFound   : GOTO LookAtLines '
   IF CmdOpen     <> "" AND CmdOpen     <> " " THEN me.FMLCmdOpen    : GOTO LookAtLines '
   IF CmdFavorite <> "" AND CmdFavorite <> " " THEN me.FMLCmdFavorite: GOTO LookAtLines '
   IF CmdFileList <> "" AND CmdFileList <> " " THEN me.FMLCmdFileList: GOTO LookAtLines '
   IF CmdRPaths   <> "" AND CmdRPaths   <> " " THEN me.FMLCmdRPaths  : GOTO LookAtLines '
   IF CmdProfiles <> "" AND CmdProfiles <> " " THEN me.FMLCmdProfiles: GOTO LookAtLines '

   LookAtLines:
   FOR i = 1 TO AFCount                                           ' Lets see if any lines were selected
      IF ISNOTNULL(TRIM$(AFList(i).Cmd)) THEN                     ' We have a command
         lcmd = LTRIM$(AFList(i).Cmd)                             '
         IF LEFT$(lcmd, 1) = "." THEN lcmd = "S " + lcmd          ' If just a profile request, prefix it with "S"
         LCmdOpsNum = 0: RESET LCmdOps()                          ' Get into normal string
         DO WHILE TRIM$(lcmd) <> ""                               ' Break it down
            INCR LCmdOpsNum                                       '
            LCmdOps(LCmdOpsNum) = sGetWord(lcmd, %Strip, %QuoteSig)
         LOOP                                                     '
         DoSet(%LoadData)                                         ' Request Data reload
         SELECT CASE AS CONST$ UUCASE(LCmdOps(1))                 ' What command?
            CASE "A", "ADD":                 me.FMLCmdAdd(i)
            CASE "ALL":                      me.FMLCmdAll(i)
            CASE "B", "BROWSE", "BRO":       me.FMLCmdBrowse(i)
            CASE "C", "CLONE":               me.FMLCmdClone(i)
            CASE "CANCEL", "CAN"
               INCR DefCtr                                        '
               Deferred(DefCtr) = FORMAT$(AFlist(i).Uniq, "000") + " CANCEL " + FORMAT$(i, "000")

            CASE "DIR":                      me.FMLCmdDIR(i)
            CASE "D", "U", "DEL", "DELETE", "PURGE", "PUR"
               INCR DefCtr                                        '
               Deferred(DefCtr) = FORMAT$(AFlist(i).Uniq, "000") + " DELETE " + FORMAT$(i, "000") + " " + UUCASE(LCmdOps(1))
            CASE "E", "EDIT":                me.FMLCmdEdit(i)
            CASE "END"
               INCR DefCtr                                        '
               Deferred(DefCtr) = FORMAT$(AFList(i).Uniq, "000") + " END " + FORMAT$(i, "000")
            CASE "F", "FORGET":              me.FMLCmdForget(i)
            CASE "J", "JOB", "SUBMIT", "SUB":me.FMLCmdSubmit(i)
            CASE "L", "LINES", "LINE":       me.FMLCmdLines(i)
            CASE "M", "MEDIT":               me.FMlCmdMedit(i)
            CASE "N", "NORM":                me.FMLCmdNorm(i)
            CASE "OPEN", "OPENV", "OPENB", "O", "OV","OB": me.FMLCmdOpenF(i)
            CASE "P", "PRINT":               me.FMLCmdPrint(i)
            CASE "R", "REN", "RENAME":       me.FMLCmdRename(i)
            CASE "SAVE":                     me.FMLCmdSave(i)
            CASE "S", "SELECT", "SEL":       me.FMLCmdSelect(i)
            CASE "T", "TOUCH":               me.FMLCmdTouch(i)
            CASE "V", "VIEW":                me.FMLCmdView(i)
            CASE "X":                        me.FMLCmdExclude(i)
            CASE "WDIR":                     me.FMLCmdWDIR(i)
         END SELECT                                               '
      END IF                                                      '
   NEXT i                                                         '
   IF DefCtr > 0 THEN                                             ' Any deferred commands?
      ARRAY SORT Deferred(1) FOR DefCtr, DESCEND                  ' Reverse sort
      FOR i = 1 TO DefCtr                                         ' Do the deferred commands
         t = sGetWord(Deferred(i), %Strip, %QuoteSig)             ' Get the 1st word (Uniq)
         tt = sGetWord(Deferred(i), %Strip, %QuoteSig)            ' Get the 2nd word 'command'
         ttt = sGetWord(Deferred(i), %Strip, %QuoteSig)           ' Get the 3rd word 'operand'
         tttt = sGetWord(Deferred(i), %Strip, %QuoteSig)          ' Get the 4th word 'operand'
         IF tt = "END"    THEN me.FMLCmdEnd(VAL(ttt))             ' END?
         IF tt = "CANCEL" THEN me.FMLCmdCancel(VAL(ttt))          ' CANCEL?
         IF tt = "ODEL"   THEN me.FMLCmdODEL(VAL(ttt), tttt)      ' OPEN DELETE?
         IF tt = "DELETE" THEN me.FMLCmdDelete(VAL(ttt), tttt)    ' DELETE/PURGE
      NEXT i                                                      '
   END IF                                                         '

   '----- Fire off MEdit if we build a command
   IF ISNOTNULL(MEditCmd) THEN                                    ' A pending MEDIT command
      me.CallTab("MEDIT", MEditCmd)                               ' Let MEDIT start it
   END IF                                                         '
   MExit                                                          '
END METHOD

METHOD  FMLCmdProfiles()
'---------- FM Profiles line command
   MEntry                                                         '
   CmdProfiles = " "                                              ' Clear command char
   FileListNm = "Profiles"                                        ' Switch to it
   NestCtr = 0                                                    ' Kill nesting then
   me.WindowTitle                                                 ' Re-do Window Title
   DoSet(%LoadData)                                               ' Ask for refresh
   MExit                                                          '
END METHOD

METHOD  FMLCmdRecent()
'---------- FM Recent line command
   MEntry                                                         '
   CmdRecent = " "                                                ' Clear command char
   FileListNm = "Recent Files"                                    ' Switch to it
   NestCtr = 0                                                    ' Kill nesting then
   me.WindowTitle                                                 ' Re-do Window Title
   DoSet(%LoadReq)                                                ' Ask for refresh
   MExit                                                          '
END METHOD

METHOD  FMLCmdRename(i AS LONG)
'---------- FM Rename line command
LOCAL lclPath, lclDrive, RenDrive, RenPath, OldFn, RenFn, ErrMsg, lclOldState, lclNewState AS STRING

   MEntry
   AFList(i).Cmd = " "                                            ' Clear command char
   IF ISNOTNULL(LCmdOps(2)) THEN RenFn = LCmdOps(2)               ' New name provided?  Save it

   lclPath = CURDIR$                                              ' Get current path
   IF MID$(lclPath, 2, 1) = ":" THEN _                            ' Extract Drive if present
      lclDrive = LEFT$(lclPath, 2)                                ' and save it
   RenPath = RTRIM$(AFList(i).Path)                               ' Build Path
   RenPath = LEFT$(RenPath, LEN(RenPath) - 1)                     ' Remove trailing \
   IF MID$(RenPath, 2, 1) = ":" THEN                              ' Extract Drive if present
      RenDrive = LEFT$(RenPath, 2)                                ' and save it
      RenPath = MID$(RenPath, 3)                                  ' Now remove Drive portion
      RenPath = IIF$(ISNULL(RenPath), "\", RenPath)               ' If null, make it \
   END IF                                                         '
   OldFn = RTRIM$(AFList(i).FD.FileName)                          ' Get old file Name

   SELECT CASE AS LONG AFList(i).Flag                             ' Split by type

      CASE %FEntry, %FFLEntry, %FFileList, %FProfile              ' File
         IF (AFList(i).FD.FileAttributes AND %FILE_ATTRIBUTE_READONLY) = %FILE_ATTRIBUTE_READONLY THEN _
            MLErrExit(i, "Can't Rename a Read-Only file")         '
         IF VAL(sFileQueue("S", " ", TRIM$(AFList(i).Path) + TRIM$(AFList(i).FD.FileName))) <> 0 THEN _   ' Open in some tab?
            MLErrExit(i, "File open elsewhere")                   '

         IF ISNULL(RenFn) THEN RenFn = sDoInputBox("Enter the new FileName: ", "Rename", RTRIM$(OldFn))
         IF ISNULL(RenFn) OR IsEQ(RenFn, OldFn) THEN MExitMeth    ' Bail out if none provided

         '----- Get STATE names if available
         lclOldState = BUILD$(RenDrive, Renpath, "\", OldFN, ".STATE")  ' Build STATE file name
         REPLACE ANY ":\/" WITH "```" IN lclOldState              ' Make : / and \ into `
         lclOldState = ENV.StatePath + lclOldState                ' Add our STATE folder
         IF ISFALSE ISFILE(lclOldState) THEN                      ' Does STATE even exist?
            lclOldState = ""                                      ' Kill it if no STATE file
         ELSE
            lclNewState = BUILD$(RenDrive, Renpath, "\", RenFN, ".STATE")  ' Build new STATE file name
            REPLACE ANY ":\/" WITH "```" IN lclNewState           ' Make : / and \ into `
            lclNewState = ENV.StatePath + lclNewState             ' Add our STATE folder
            IF ISFILE(lclNewState) THEN lclOldState = ""          ' Kill it all if new STATE already exists
         END IF

         '----- OK, Finally try the rename
         ERRCLEAR                                                 ' Reset ERR
         IF RenDrive <> lclDrive THEN CHDRIVE RenDrive            '
         CHDIR RenPath                                            '
         NAME OldFn AS RenFn                                      ' Do the rename
         sFileListRename(BUILD$(RenDrive, Renpath, "\", OldFN), BUILD$(RenDrive, Renpath, "\", RenFN))  ' Global rename
         IF lclOldState <> "" THEN                                ' Doing STATE rename?
            NAME lclOldState AS lclNewState                       ' Yes, do it
         END IF                                                   '


      CASE %FDirDown                                              ' Downward Directory
         OldFn = TRIM$(OldFn)                                     ' Trim
         OldFn = LEFT$(OldFn, LEN(OldFn) - 1)                     ' Remove trailing \ character
         IF ISNULL(RenFn) THEN RenFn = sDoInputBox("Enter the new Directory Name: ", "Rename Directory", TRIM$(OldFn))
         IF ISNULL(RenFn) OR IsEQ(RenFn, OldFn) THEN MExitMeth    ' Bail out if none provided

         ERRCLEAR                                                 ' Reset ERR
         IF RenDrive <> lclDrive THEN CHDRIVE RenDrive            '
         CHDIR RenPath                                            '
         NAME OldFn AS RenFn                                      ' Do the rename
         IF RenDrive <> lclDrive THEN CHDRIVE lclDrive            ' Switch drive if needed
         CHDIR lclPath                                            ' put back the original path

      CASE ELSE                                                   '
         MLErrExit(i, "RENAME is not supported for this entry type")
   END SELECT                                                     '

   '----- See if an error msg to issue
   ErrMsg = ""                                                    '
   IF ERR = %ERR_FILENOTFOUND THEN                                '
      ErrMsg = "File/Dir no longerfound??"                        '
   ELSEIF ERR = %ERR_FILEALREADYEXISTS THEN                       '
      ErrMsg = "File/Dir already exists."                         '
   ELSEIF ERR = %ERR_PERMISSIONDENIED THEN                        '
      ErrMsg = "File/Dir is Read Only."                           '
   ELSEIF ERR = %ERR_PATHFILEACCESSERROR THEN                     '
      ErrMsg = "Path/File Access error."                          '
   ELSEIF ERR <> 0 THEN                                           '
      ErrMsg = "UnKnown Error - Err: " + FORMAT$(ERR)             '
   END IF                                                         '
   IF ISNOTNULL(ErrMsg) THEN                                      ' Any error?
      AFList(i).Msg = ErrMsg                                      '
      DoSet(%Msg)                                                 ' Get msg issued
   END IF                                                         '
   DoSet(%LoadData)                                               ' Ask for refresh
   AFList(i).Cmd = " "                                            ' Blank the command area
   MExit
END METHOD

METHOD  FMLCmdRPaths()
'---------- FM Recent Paths line command
   MEntry                                                         '
   CmdRPaths = " "                                                ' Clear command char
   FileListNm = "Recent Paths"                                    ' Switch to it
   NestCtr = 0                                                    ' Kill nesting then
   me.WindowTitle                                                 ' Re-do Window Title
   DoSet(%LoadReq)                                                '
   MExit                                                          '
END METHOD

METHOD  FMLCmdSave(i AS LONG)
'---------- FM SAVE line command
LOCAL MSG AS STRING, j AS LONG
   MEntry
   IF AFList(i).Flag <> %FOpen THEN  MLErrExit(i, "SAVE only usable against OPEN files") ' An OPEN entry?
   sCallTabCmd(AFList(i).Uniq, "SAVE ")                           ' Try the command
   AFList(i).Cmd = " "                                            ' Blank the command area
   MExit                                                          '
END METHOD

METHOD   FMLCmdSelect(i AS LONG)
'---------- FM Select line command
LOCAL cmd, fn, lcmd, SetProf, tFn, t, pfn, SaveProf, p AS STRING, j, FNum AS LONG
   MEntry
   lcmd = AFList(i).Cmd                                           ' Save command string
   WHILE lcmd <> ""                                               ' Process lcmd for .profile name
      cmd = sGetWord(lcmd, %Strip, %QuoteNotSig)                  ' Get a word from the command
      IF LEFT$(cmd, 1) = "." THEN SetProf = UUCASE(cmd)           '
   WEND                                                           '

   SELECT CASE AS LONG AFList(i).Flag                             ' Split by type of entry

      CASE %FEntry, %FFLEntry                                     ' A file entry
         me.DoEDIT(i)                                             ' go do it

      CASE %FOpen                                                 ' OPEN file entry
         gTabSwitch = AFList(i).Uniq                              ' Set to cause a switch to this tab

      CASE %FPath                                                 ' Recent path entry
         FPath = AFList(i).FD.FileName                            '
         FMask = "*"                                              '
         FileListNm = ""                                          ' Kill FileList name
         sIniSetString("FManager", "DefDir1", FPath)              '
         sIniSetString("FManager", "DefMask", FMask)              '
         sIniSetString("FManager", "Recent", FileListNm)          '
         DoSet(%LoadReq)                                          ' Full refresh
         me.WindowTitle                                           ' Alter Window Title

      CASE %FDirDown                                              ' Dir Down
         me.NestAdd(FPath, FMask, "")                             ' Add it to the nest level so we can return here
         me.DoDirDown(i)                                          ' Go do it

      CASE %FDirUp                                                ' Dir Up
         IF NestCtr > 0 THEN                                      '
            me.NestDel: MExitMeth                                 ' If one is available
         ELSE                                                     '
            me.DoDirUp(i)                                         ' Go do it
         END IF                                                   '

      CASE %FFileList                                             ' FILELIST
         IF FileListNm = "Named Favorites" THEN                   ' Selected from the Named Favorites?
            me.NestAdd(FPath, FMask, "Named Favorites")           ' Add it to the nest level so we return here
         ELSE                                                     ' Something else?
            NestCtr = 0                                           ' Kill nesting then
         END IF                                                   '
         FileListNm = LEFT$(AFList(i).FD.FileName, LEN(TRIM$(AFList(i).FD.FileName)) - 6)
         sIniSetString("FManager", "Recent", FileListNm)          '
         me.WindowTitle                                           ' Alter Window Title

      CASE %FProfile                                              ' Profile
         SaveProf = IIF$(Prf.ProfName <> "", Prf.ProfName, "DEFAULT") ' Save where we are
         p = TRIM$(AFList(i).FD.FileName)                         ' Get Profile name
         p = LEFT$(p, LEN(p) - 4)                                 ' Strip off .INI
         Prf.SetProfName(p, %False)                               ' Setup
         Prf.ReadAll(%True)                                       '
         Prf.EditProf                                             ' Let user Have a look
         Prf.SetProfName(SaveProf, %False)                        '
         Prf.ReadAll(%True)                                       '
         CONTROL SET FOCUS hWnd, %IDC_SPFLiteTAB                  ' Set focus

      CASE ELSE
         MLErrExit(i, "Select not supported for this entry type") ' An OPEN entry?

   END SELECT                                                     '
   AFList(i).Cmd = " "                                            ' Blank the command area
   MExit                                                          '
END METHOD

METHOD FMLCmdSpread() AS LONG
'---------- Spread out range commands, Checking was all done by LCmdValidate
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL lcmd1, lcmd2 AS STRING, x, y, numval AS LONG
   MEntry

   '----- Loop through the AFList table
   FOR i = 1 TO AFCount                                           ' Loop through AFList
      lcmd1 = UUCASE(TRIM$(AFList(i).Cmd))                        ' Get command on this line
      IF ISNULL(TRIM$(lcmd1)) THEN ITERATE FOR                    ' Nothing, move on
      RESET numval                                                ' Reset
      x = INSTR(lcmd1, ANY $Numeric)                              ' Get loc of any trailing numeric
      IF x THEN numval = VAL(MID$(lcmd1, x))                      ' Save numeric value in numval
      lcmd2 = IIF$(x = 0, lcmd1, LEFT$(lcmd1, x - 1))             ' Strip off numeric if present

      IF LEN(lcmd2) = 2 AND LEFT$(lcmd2, 1) = RIGHT$(lcmd2, 1) THEN ' A doubled command
         AFList(i).Cmd = LEFT$(lcmd2, 1)                          ' Make it the single char version
         FOR x = i + 1 TO AFCount                                 ' Loop till we see the other
            IF TRIM$(AFList(x).Cmd) = "" THEN                     ' A blank in between
               AFList(x).Cmd = LEFT$(lcmd2, 1)                    ' Make it the single char version
            ELSE                                                  ' This HAS to be the closing double entry
               AFList(x).Cmd = LEFT$(lcmd2, 1)                    ' Make it the single char version
               i = x                                              ' Continue on
               EXIT FOR                                           ' Done this loop
            END IF                                                '
         NEXT x                                                   '
      ELSEIF LEN(lcmd2) = 1 AND numval > 0 THEN                   ' A numeric repeat?
         FOR x = i TO i + numval - 1                              ' Fill them in
            AFList(x).Cmd = lcmd2                                 ' Make it the single char version
         NEXT x                                                   '
         i = i + numval - 1                                       ' Set continue point
      END IF                                                      '
   NEXT i ' AFList                                                '
   MExit
END METHOD

METHOD  FMLCmdSubmit(i AS LONG)
'---------- FM SUBMIT process
LOCAL fn AS STRING
   MEntry
   AFList(i).Cmd = " "                                            ' Blank the command area
   SELECT CASE AS LONG AFList(i).Flag                             ' Split by file entry type
      CASE %FEntry, %FFLEntry, %FOpen                             ' File Entry?
         fn = RTRIM$(AFList(i).Path) + RTRIM$(AFList(i).FD.FileName) ' Get filename
      CASE ELSE                                                   '
         MLErrExit(i, "SUBMIT cannot be used on this entry type") '
   END SELECT                                                     '
   pCmdXSubmit("XSUBMIT " + $DQ + fn + $DQ)                       ' Run it
   DoSet(%Refresh)                                                '
   MExit                                                          ' We're done
END METHOD

METHOD  FMLCmdTouch(i AS LONG)
'---------- FM Touch line command
LOCAL lclPath, lclDrive, TchDrive, TchPath, TchFn, Errmsg AS STRING
LOCAL j, hFile AS LONG, hfile2 AS DWORD
LOCAL CreationTime, LastAccessTime, LastWriteTime AS FILETIME
LOCAL lquad AS QUAD
LOCAL SystemTimeArea AS SYSTEMTIME
   MEntry
   AFList(i).Cmd = " "                                            ' Blank the command area
   SELECT CASE AS LONG AFList(i).Flag                             ' Split off

      CASE %FEntry, %FFLEntry, %FFileList, %FProfile              ' Existing file
         lclPath = CURDIR$                                        ' Get current path
         IF MID$(lclPath, 2, 1) = ":" THEN _                      ' Extract Drive if present
            lclDrive = LEFT$(lclPath, 2)                          ' and save it
         TchPath = TRIM$(AFList(i).Path)                          ' Build Path
         IF MID$(TchPath, 2, 1) = ":" THEN                        ' Extract Drive if present
            TchDrive = LEFT$(TchPath, 2)                          ' and save it
            TchPath = MID$(TchPath, 3)                            '
            IF ISNULL(TchPath) THEN TchPath = "\"                 ' In case we're at the root
         END IF                                                   '

         TchFn = RTRIM$(AFList(i).FD.FileName)                    ' Get file Name
         IF VAL(sFileQueue("S", " ", TchPath + "\" + TchFn)) <> 0 THEN _ ' Open in some tab?
            MLErrExit(i, "File open elsewhere")                   '

         '----- Do the Touch
         IF TchDrive <> lclDrive THEN CHDRIVE TchDrive            '
         CHDIR TchPath                                            '

         ON ERROR GOTO TchError                                   '
         hFile = FREEFILE                                         ' Touch the file now
         OPEN TchFn FOR APPEND ACCESS READ WRITE LOCK SHARED AS #hFile
         hFile2 = FILEATTR(hFile,2)                               ' Get Windows handle for the file
         GetFileTime hFile2, CreationTime, LastAccessTime, LastWriteTime' Fetch timestamps
         GetSystemTime SystemTimeArea                             ' Get the current time
         SystemtimeToFileTime SystemTimeArea, LastWriteTime       ' Put it in LastWriteTime
         SetFileTime hFile2, CreationTime, LastAccessTime, LastWriteTime
         CLOSE #hFile                                             ' Close it
         IF TchDrive <> lclDrive THEN CHDRIVE lclDrive            ' Switch drive if needed
         CHDIR lclPath                                            ' Put back the original path
         AFList(i).Msg = "File touched"                           '
         DoSet(%Msg)                                              ' Get msg issued

      CASE ELSE                                                   '
         AFList(i).Msg = "TOUCH cannot be used on this type of entry"
         DoSet(%Msg)                                              ' Get msg issued
   END SELECT                                                     '

   TchExit:                                                       '
      DoSet(%LoadData)                                            ' Ask for refresh
      MExitMeth                                                   '

   TchError:                                                      '
      AFList(i).Msg = "Touch File failed"                         '
      DoSet(%Msg)                                                 ' Get msg issued
      RESUME TchExit                                              '
END METHOD

METHOD FMLCmdValidate() AS LONG
'---------- Validate the raw line command strings
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL lcmd1, lcmd2 AS STRING, x, y, numval AS LONG
DATA "A", "AA", "ADD","ALL", "E", "EE", "EDIT", "ED", "C", "CC", "CLONE", "CANCEL", "CAN", "B", "BB", "BROWSE", "BRO", "BR"
DATA "DIR", "D", "DD", "U", "UU", "DEL", "DELETE", "PURGE", "PUR", "END", "F", "FF", "FORGET"
DATA "J", "JJ", "JOB", "SUBMIT", "SUB", "L", "LL", "LINE", "LINES", "M", "MM", "MEDIT", "N", "NN", "NORM", "OPEN", "OPENV", "OPENB"
DATA "O", "OV", "OB", "P", "PP", "PRINT", "R", "RR", "REN", "RENAME"
DATA "SAVE", "S", "SS", "SELECT", "SEL", "T", "TT", "TOUCH", "V", "VIEW", "WDIR", "X", "XX"
   MEntry

   '----- Loop through the AFList table
   FOR i = 1 TO AFCount                                           ' Loop through AFList
      lcmd1 = UUCASE(TRIM$(AFList(i).Cmd))                        ' Get command on this line
      IF ISNULL(TRIM$(lcmd1)) THEN ITERATE FOR                    ' Nothing, move on
      lcmd1 = sGetWord(lcmd1, %NoStrip, %QuoteNotSig)             ' Get 1st word
      RESET y, numval                                             ' Reset
      x = INSTR(lcmd1, ANY $Numeric)                              ' Get loc of any trailing numeric
      IF x THEN numval = VAL(MID$(lcmd1, x))                      ' Save numeric value in numval
      lcmd2 = IIF$(x = 0, lcmd1, LEFT$(lcmd1, x - 1))             ' Strip off numeric if present
      IF LEFT$(lcmd2, 1) = "." THEN y = 1                         ' Let .profile get through
      FOR j = 1 TO DATACOUNT                                      ' See if valid
         IF lcmd2 = READ$(j) THEN y = i: EXIT FOR                 ' Found, save and exit loop
      NEXT j ' DATA                                               '
      IF y = 0 THEN                                               ' Not found
         METHOD = %True                                           ' Bail out
         MLErrExit(i, "Invalid line command: " + lcmd1)           '
      END IF                                                      '

      '----- Validate block command
      IF LEN(lcmd2) = 2 AND LEFT$(lcmd2, 1) = RIGHT$(lcmd2, 1) THEN ' A doubled command
         IF numval > 0 THEN                                       '
            METHOD = %True                                        ' Bail out
            MLErrExit(i, "Invalid use of numeric operand: " + lcmd1)
         END IF                                                   '

         FOR x = i + 1 TO AFCount                                 ' See if the other is here
            IF UUCASE(TRIM$(AFList(x).Cmd)) = lcmd2 THEN          ' Found the match?
               i = x: GOTO rangeOK                                ' Outer loop to resume at matching line
            ELSEIF TRIM$(AFList(x).Cmd) <> "" THEN                ' Some other non-blank?
               METHOD = %True                                     ' Bail out
               MLErrExit(i, "No matching line command for: " + lcmd1)
            END IF                                                '
         NEXT x                                                   '
         METHOD = %True                                           ' Bail out
         MLErrExit(i, "No matching line command for: " + lcmd1)   '

      '----- Validate numeric range command
      ELSEIF LEN(lcmd2) = 1 AND numval > 0 THEN                   ' A repeat value
         FOR x = i + 1 TO MIN(i + numval - 1, AFCount)            ' See if all blank within range
            IF UUCASE(TRIM$(AFList(x).Cmd)) <> "" THEN            ' Something there?
               METHOD = %True                                     ' Bail out
               MLErrExit(i, "Command repeat overlaps next command: " + lcmd1)
            END IF                                                '
          NEXT x                                                  '
          i = x                                                   ' Outer loop to resume after repeat range
      END IF                                                      '
      RangeOK:
   NEXT i ' AFList                                                '
   MExit
END METHOD

METHOD   FMLCmdView(i AS LONG)
'---------- FM View line command
LOCAL pfn AS STRING
   MEntry
   SELECT CASE AS LONG AFList(i).Flag                             ' Split by entry type
      CASE %FEntry, %FFLEntry, %FFileList                         ' Existing File?
         me.DoVIEW(i)                                             ' Go do View
         DoSet(%LoadData)                                         ' Ask for refresh

   END SELECT                                                     '
   AFList(i).Cmd = " "                                            ' Blank the command area
   MExit                                                          '
END METHOD

METHOD  FMLCmdWDir(i AS LONG)
'---------- Open a Windows Explorer to this file's DIR Folder
LOCAL lclpath AS STRING, j AS LONG
   MEntry
   SELECT CASE AS LONG AFList(i).Flag                             ' Split by Entry type
      CASE %FEntry, %FFLEntry, %FOpen                             ' File type?
         lclPath = TRIM$(AFList(i).Path) + TRIM$(AFList(i).FD.FileName) ' Build the full Set the filename
         SHELL "EXPLORER.EXE /SELECT, " + $DQ + lclPath + $DQ, 1  '

      CASE %FDirDown                                              ' Dir Entry name
        lclPath = TRIM$(AFList(i).Path)                           ' Set the path
        SHELL "EXPLORER.EXE " + $DQ + lclPath + $DQ, 1            '

      CASE %FPath                                                 ' Dir Entry name
         lclPath = TRIM$(AFList(i).FD.Filename)                   '
         SHELL "EXPLORER.EXE " + $DQ + lclPath + $DQ, 1           '

      CASE %FDirUp                                                ' Dir Entry name
         lclPath = TRIM$(AFList(i).Path)                          '
         j = INSTR(-1, lclPath, "\")                              ' Find last \
         IF j <> 0 THEN                                           ' Got one
            lclPath = LEFT$(lclPath, j - 1)                       ' Cut off the last level
         END IF                                                   ' Else do nothing
         SHELL "EXPLORER.EXE " + $DQ + lclPath + $DQ, 1           '

      CASE ELSE                                                   '
         MLErrExit(i, "WDIR is not supported for this entry type")'
   END SELECT
   AFList(i).Cmd = " "                                            ' Blank the command area
   MExit
END METHOD
