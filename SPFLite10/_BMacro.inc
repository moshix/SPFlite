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


FUNCTION MB_Call() AS LONG
'---------- Execute the macro in a seperate thread in case it loops
LOCAL idThread AS DWORD
LOCAL Arg, RetVal, i, WC AS LONG
LOCAL MSG AS STRING
   gMacNoLoopTest = %False: gMacroTrace = %False                  ' Reset Users no loop check flag and Trace flag
   gMacTempVars = %False                                          ' Reset Temp vars need cleanup
   gMacThread = CreateThread(BYVAL %NULL, 0, CODEPTR(MB_CallThread), BYREF Arg, _
                        %THREAD_TERMINATE OR %THREAD_QUERY_INFORMATION, _
                        BYVAL VARPTR(idThread))

   IF gMacThread <> %INVALID_HANDLE_VALUE THEN
      DO
         WC = WaitForSingleObject(gMacThread, 10000)              ' Give macro 10 seconds to execute
         SELECT CASE WC                                           ' What came back?

            CASE %WAIT_ABANDONED                                  '
               '                                                  '
            CASE %WAIT_OBJECT_0                                   ' Thread did its job and ended by itself
               GetExitCodeThread(gMacThread, Retval)              '
               TP.CsrRow = 1: TP.CsrCol = 11                      ' reset the command line
               TP.ErrFlag = 0: TP.ErrMsg = ""                     ' Clear intermediate errors
               IF gMacroRC > 0 OR gMacroMsg <> "" THEN _          ' Something to issue?
                  scError(gMacroRC, gMacroMsg)                    ' Issue the RC and Msg
               CloseHandle gMacThread                             ' Clear handle
               gMacroMode = %False                                ' Turn off Macro mode
               sPopReset                                          ' Reset popup state
               TP.AttnDo = (TP.AttnDo OR %Refresh)                ' Ask for refresh
               EXIT DO                                            ' Done

            CASE %WAIT_TIMEOUT                                    ' Delay is over, see what to do
               IF gMacNoLoopTest THEN ITERATE DO                  ' User says to ignore loops
               IF sFindWindow("thinBasic :: Script RunTime error!") THEN ITERATE DO ' If thinbasic telling of an error, also exit
               MSG = "Macro may be looping," + $CRLF + "Select OK, to ignore loop and continue execution," + $CRLF + _
                     "Select Cancel to terminate the macro"       '
               i = sDoMsgBox(MSG, %MB_USERICON OR %MB_OKCANCEL, "SPFLite Macro Loop Intercept")
               IF i = %IDOK THEN                                  ' Continue?
                  ITERATE DO                                      ' Just keep waiting
               ELSE                                               '
                  TerminateThread BYVAL gMacThread, 999           ' Second value is thread exit code
                  SLEEP 500                                       ' Give thread time to end
                  GetExitCodeThread(gMacThread, Retval)           '
                  CloseHandle gMacThread                          ' Call CloseHandle after TerminateThread
                  scError(%eFail, "Macro was terminated")
                  sPopReset                                       ' Do final reset
                  TP.AttnDo = (TP.AttnDo OR %Refresh)             ' Ask for refresh
                  gMacroMode = %False                             ' Turn off Macro mode
                  EXIT DO                                         '
               END IF                                             '
        END SELECT
      LOOP
   ELSE
     scError(%eFail, "Unable to start macro thread")
   END IF

   '----- Clear any transient labels
   IF gMacTempVars THEN                                           ' Need temp variables cleanup?
      FOR i = 1 TO TP.LastLine                                    ' Must search all lines
         IF LEFT$(TP.LLblGet(i), 2) = "._" THEN                   ' Contain a temporary label?
            TP.LLblSet(i, $BlankLNo)                              ' Clear the line
            TP.UpdLControl(i)                                     ' Redo LLCtl
         END IF                                                   '
      NEXT i                                                      '
   END IF                                                         '

   TP.CsrRow = 1: TP.CsrCol = 11                                  ' reset the command line
   LTblRange = %False                                             ' Reset line ranges
END FUNCTION

SUB      MB_CallThread()
'---------- Handle the new MACRO BASIC format
LOCAL BasicPgm AS STRING
LOCAL hLib_thinCore                  AS LONG                      ' Handle of thinCore.dll library
LOCAL hProc_thinBasic_Init           AS LONG                      ' Handle to Init function
LOCAL hProc_thinBasic_LoadSymbol     AS LONG                      ' Handle to LoadSymbol function
LOCAL hProc_thinBasic_AddIncludePath AS LONG                      ' Handle to AddIncludePath function
LOCAL hProc_thinBasic_AddVariable    AS LONG                      ' Handle to AddVariable function
LOCAL hProc_thinBasic_AddEquate      AS LONG                      ' Handle to AddEquate function
LOCAL hProc_thinBasic_Run            AS LONG                      ' Handle to Run function
LOCAL hProc_thinBasic_Release        AS LONG                      ' Handle to Release function
LOCAL lRet                           AS LONG
'   MEntry
   gMacroRC = 0: gMacroMsg = ""                                   ' Clear the RC and Msg

   '----- Save the Line control context at start of macro
   gMacRange  = LTblRange                                         ' Line control range at Macro start
   gMacSCmd   = TRIM$(LTblSCmd)                                   '
   gMacSFrom  = LTblSFrom                                         '
   gMacSTo    = LTblSTo                                           '
   gMacSRpt   = LTblSRpt                                          '
   gMacSFlag  = LTblSFlag                                         '
   gMacDCmd   = TRIM$(LTblDCmd)                                   '
   gMacDFrom  = LTblDFrom                                         '
   gMacDTo    = LTblDTo                                           '
   gMacDRpt   = LTblDRpt                                          '
   gMacDFlag  = LTblDFlag                                         '

   '---Open and load thinCore.dll library
   hLib_thinCore = LoadLibraryA( BYCOPY "thinCore.Dll" )

   '---If all went fine
   IF hLib_thinCore THEN                                          ' thinCore exists
      gMacCore = hLib_thinCore                                    ' Save Core address

      '---Try to load the functions
      hProc_thinBasic_Init           = GetProcAddress(hLib_thinCore, BYCOPY "thinBasic_Init")
      hProc_thinBasic_LoadSymbol     = GetProcAddress(hLib_thinCore, BYCOPY "thinBasic_LoadSymbol")
      hProc_thinBasic_AddVariable    = GetProcAddress(hLib_thinCore, BYCOPY "thinBasic_AddVariable")
      hProc_thinBasic_AddIncludePath = GetProcAddress(hLib_thinCore, BYCOPY "thinBasic_AddIncludePath")
      hProc_thinBasic_AddEquate      = GetProcAddress(hLib_thinCore, BYCOPY "thinBasic_AddEquate")
      hProc_thinBasic_Run            = GetProcAddress(hLib_thinCore, BYCOPY "thinBasic_Run"     )
      hProc_thinBasic_Release        = GetProcAddress(hLib_thinCore, BYCOPY "thinBasic_Release" )
      gMacRelease = hProc_thinBasic_Release                       ' Save release for global use

      '---If all went fine ...
      IF hProc_thinBasic_Init           AND _
         hProc_thinBasic_LoadSymbol     AND _
         hProc_thinBasic_AddVariable    AND _
         hProc_thinBasic_AddIncludePath AND _
         hProc_thinBasic_AddEquate      AND _
         hProc_thinBasic_Run            AND _
         hProc_thinBasic_Release        THEN

         '---Initialise all necessary functions etc.
         CALL DWORD hProc_thinBasic_Init USING thinBasic_Init(0, hInstance, "thinBasic")

         '----- Add our Macro path the thinBasic's #INCLUDE path
         CALL DWORD hProc_thinBasic_AddIncludePath USING thinBasic_AddIncludePath(ENV.MacrosPath)

         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_ARG$",              %thinBasic_ReturnString, CODEPTR(GET_ARG))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_ARG_COUNT",         %thinBasic_ReturnNumber, CODEPTR(GET_ARGCOUNT))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_ARG_KW",            %thinBasic_ReturnNumber, CODEPTR(GET_ARGKW))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_ARG_KWGROUP$",      %thinBasic_ReturnString, CODEPTR(GET_ARGKWGROUP))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_ARG_TEXTLIT$",      %thinBasic_ReturnString, CODEPTR(GET_ARGTextLit))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_ARG_LREF$",         %thinBasic_ReturnString, CODEPTR(GET_ARGLRef))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_ARG_NUMLIT$",       %thinBasic_ReturnString, CODEPTR(GET_ARGNumLit))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_ARG_TAG$",          %thinBasic_ReturnString, CODEPTR(GET_ARGTag))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_CLR_LINE$",         %thinBasic_ReturnString, CODEPTR(Get_ClrLine))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_E_SOURCE$",         %thinBasic_ReturnString, CODEPTR(Get_ESource))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_CSR_COL",           %thinBasic_ReturnNumber, CODEPTR(Get_CSRCOL))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_CSR_LPTR",          %thinBasic_ReturnNumber, CODEPTR(Get_CSRLIN))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_CURR_LINE$",        %thinBasic_ReturnString, CODEPTR(Get_CURRLINE))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_CURR_WORD$",        %thinBasic_ReturnString, CODEPTR(Get_CURRWORD))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_CURR_PATH$",        %thinBasic_ReturnString, CODEPTR(Get_CURRPATH))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_DEST_LCMD$",        %thinBasic_ReturnString, CODEPTR(Get_DESTLCMD))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_DEST1_LPTR",        %thinBasic_ReturnNumber, CODEPTR(Get_DESTL1))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_DEST2_LPTR",        %thinBasic_ReturnNumber, CODEPTR(Get_DESTL2))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_DEST_OP",           %thinBasic_ReturnNumber, CODEPTR(Get_DESTOP))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_DEST_LMOD$",        %thinBasic_ReturnString, CODEPTR(Get_DESTLMOD))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_ENVVAR$",           %thinBasic_ReturnString, CODEPTR(Get_SYSVAR))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_EXE_PATH$",         %thinBasic_ReturnString, CODEPTR(Get_EXEPath))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_FILEBASE$",         %thinBasic_ReturnString, CODEPTR(Get_FileBase))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_FILEDATE$",         %thinBasic_ReturnString, CODEPTR(Get_FileDate))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_FILEEXT$",          %thinBasic_ReturnString, CODEPTR(Get_FileExt))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_FILENAME$",         %thinBasic_ReturnString, CODEPTR(Get_FileName))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_FILEPATH$",         %thinBasic_ReturnString, CODEPTR(Get_FilePath))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_FILETIME$",         %thinBasic_ReturnString, CODEPTR(Get_FileTime))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_FIND_COL",          %thinBasic_ReturnNumber, CODEPTR(Get_FINDCOL))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_FIND_LEN",          %thinBasic_ReturnNumber, CODEPTR(Get_FINDLEN))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_FIND_LPTR",         %thinBasic_ReturnNumber, CODEPTR(Get_FINDLIN))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_FIRST_LPTR",        %thinBasic_ReturnNumber, CODEPTR(Get_FIRST))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_FULLPATH$",         %thinBasic_ReturnString, CODEPTR(Get_FullPath))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_GBL_NUM",           %thinBasic_ReturnNumber, CODEPTR(Get_GBLNUM))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_GBL_NUM_COUNT",     %thinBasic_ReturnNumber, CODEPTR(Get_GBLNUMCOUNT))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_GBL_NUM_NAME$",     %thinBasic_ReturnString, CODEPTR(Get_GBLNUMNAME))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_GBL_NUM_TABLENAME$",%thinBasic_ReturnString, CODEPTR(Get_GBLNUMTABLENAME))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_GBL_STR$",          %thinBasic_ReturnString, CODEPTR(Get_GBLSTR))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_GBL_STR_COUNT",     %thinBasic_ReturnNumber, CODEPTR(Get_GBLSTRCOUNT))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_GBL_STR_NAME$",     %thinBasic_ReturnString, CODEPTR(Get_GBLSTRNAME))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_GBL_STR_TABLENAME$",%thinBasic_ReturnString, CODEPTR(Get_GBLSTRTABLENAME))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_INI_PATH$",         %thinBasic_ReturnString, CODEPTR(Get_INIPath))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_LABEL$",            %thinBasic_ReturnString, CODEPTR(Get_Label))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_LAST_LPTR",         %thinBasic_ReturnNumber, CODEPTR(Get_LAST))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_LAST_REAL",         %thinBasic_ReturnNumber, CODEPTR(Get_LASTREAL))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_LBOUND",            %thinBasic_ReturnNumber, CODEPTR(Get_LBOUND))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_LNUM",              %thinBasic_ReturnNumber, CODEPTR(GET_LNum))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_LINE$",             %thinBasic_ReturnString, CODEPTR(Get_LINE))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_LINE_TYPE$",        %thinBasic_ReturnString, CODEPTR(Get_LINETYPE))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_LINE_LEN",          %thinBasic_ReturnNumber, CODEPTR(Get_LINELEN))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_LOC_LPTR",          %thinBasic_ReturnNumber, CODEPTR(Get_LOCLIN))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_MACNAME$",          %thinBasic_ReturnString, CODEPTR(Get_MacName))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_MODIFIED",          %thinBasic_ReturnNumber, CODEPTR(Get_MODIFIED))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_MODIFIED_FILENAME", %thinBasic_ReturnNumber, CODEPTR(Get_MODIFIEDFILENAME))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_MSG$",              %thinBasic_ReturnString, CODEPTR(Get_MSG))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_LPTR",              %thinBasic_ReturnNumber, CODEPTR(GET_LPtr))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_NEXT_LPTR",         %thinBasic_ReturnNumber, CODEPTR(GET_NextLPtr))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_PROFILE$",          %thinBasic_ReturnString, CODEPTR(GET_ProfString))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_RBOUND",            %thinBasic_ReturnNumber, CODEPTR(Get_RBOUND))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_RC",                %thinBasic_ReturnNumber, CODEPTR(Get_RC))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_SELECT_COL",        %thinBasic_ReturnNumber, CODEPTR(Get_SelCol))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_SELECT_LEN",        %thinBasic_ReturnNumber, CODEPTR(Get_SelLen))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_SELECT_FIRST_LPTR", %thinBasic_ReturnNumber, CODEPTR(Get_SelFirst))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_SELECT_LAST_LPTR",  %thinBasic_ReturnNumber, CODEPTR(Get_SelLast))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_SESSION_TYPE$",     %thinBasic_ReturnString, CODEPTR(Get_SESSIONTYPE))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_SETVAR$",           %thinBasic_ReturnString, CODEPTR(Get_SETVAR))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_SRC_LCMD$",         %thinBasic_ReturnString, CODEPTR(Get_SRCLCMD))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_SRC1_LPTR",         %thinBasic_ReturnNumber, CODEPTR(Get_SRCL1))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_SRC2_LPTR",         %thinBasic_ReturnNumber, CODEPTR(Get_SRCL2))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_SRC_OP",            %thinBasic_ReturnNumber, CODEPTR(Get_SRCOP))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_SRC_LMOD$",         %thinBasic_ReturnString, CODEPTR(Get_SRCLMOD))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_TAG$",              %thinBasic_ReturnString, CODEPTR(Get_Tag))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_TOPSCRN_LPTR",      %thinBasic_ReturnNumber, CODEPTR(Get_TOPSCREEN))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_Uniq_ID",           %thinBasic_ReturnNumber, CODEPTR(Get_UniqueID))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_XLines",            %thinBasic_ReturnNumber, CODEPTR(Get_XLines))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("GET_XStatus$",          %thinBasic_ReturnString, CODEPTR(Get_XStatus))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("HALT",                  %thinBasic_ReturnNumber, CODEPTR(HALT))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("IS_BOTTOM",             %thinBasic_ReturnNumber, CODEPTR(IS_BOTTOM))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("IS_DATA",               %thinBasic_ReturnNumber, CODEPTR(IS_DATA))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("IS_WORD",               %thinBasic_ReturnNumber, CODEPTR(IS_WORD))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("IS_MARK",               %thinBasic_ReturnNumber, CODEPTR(IS_MARK))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("IS_MASK",               %thinBasic_ReturnNumber, CODEPTR(IS_MASK))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("IS_NOTE",               %thinBasic_ReturnNumber, CODEPTR(IS_NOTE))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("IS_PROF",               %thinBasic_ReturnNumber, CODEPTR(IS_PROF))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("IS_TABS",               %thinBasic_ReturnNumber, CODEPTR(IS_TABS))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("IS_TOP",                %thinBasic_ReturnNumber, CODEPTR(IS_TOP))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("IS_ULINE",              %thinBasic_ReturnNumber, CODEPTR(IS_ULINE))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("IS_XMARKER",            %thinBasic_ReturnNumber, CODEPTR(IS_XMARKER))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("IS_BNDS",               %thinBasic_ReturnNumber, CODEPTR(IS_BNDS))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("IS_COLS",               %thinBasic_ReturnNumber, CODEPTR(IS_COLS))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("IS_FILE",               %thinBasic_ReturnNumber, CODEPTR(IS_FILE))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("IS_XLINE",              %thinBasic_ReturnNumber, CODEPTR(IS_XLine))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("IS_PRIMARY_CMD",        %thinBasic_ReturnNumber, CODEPTR(IS_PrimCmd))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("IS_LINE_CMD",           %thinBasic_ReturnNumber, CODEPTR(IS_LineCmd))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("NOP",                   %thinBasic_ReturnNumber, CODEPTR(NOP))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("OF",                    %thinBasic_ReturnNumber, CODEPTR(OV))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("REQUEST_LABEL$",        %thinBasic_ReturnString, CODEPTR(REQUEST_LABEL))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("RELEASE_LABEL$",        %thinBasic_ReturnString, CODEPTR(RELEASE_LABEL))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("RESET_GBL_NUM",         %thinBasic_ReturnNumber, CODEPTR(Reset_GBLNUM))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("RESET_GBL_STR",         %thinBasic_ReturnNumber, CODEPTR(Reset_GBLSTR))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("SET_E_RESULT",          %thinBasic_ReturnNumber, CODEPTR(SET_EResult))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("SET_CLR_LINE",          %thinBasic_ReturnNumber, CODEPTR(SET_ClrLine))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("SET_CSR",               %thinBasic_ReturnNumber, CODEPTR(SET_CSR))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("SET_LINE",              %thinBasic_ReturnNumber, CODEPTR(SET_LINE))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("SET_MSG",               %thinBasic_ReturnNumber, CODEPTR(SET_MSG))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("SET_SETVAR",            %thinBasic_ReturnNumber, CODEPTR(SET_SETVAR))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("SET_GBL_NUM",           %thinBasic_ReturnNumber, CODEPTR(SET_GBLNUM))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("SET_GBL_STR",           %thinBasic_ReturnNumber, CODEPTR(SET_GBLSTR))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("SET_TOPSCRN_LPTR",      %thinBasic_ReturnNumber, CODEPTR(SET_TOPSCREEN))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("SPF_CMD",               %thinBasic_ReturnNumber, CODEPTR(SPF_CMD))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("SPF_DEBUG",             %thinBasic_ReturnNumber, CODEPTR(SPF_Debug))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("SPF_TRACE",             %thinBasic_ReturnNumber, CODEPTR(SPF_Trace))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("SPF_EXEC",              %thinBasic_ReturnNumber, CODEPTR(SPF_EXEC))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("SPF_EXEMPT_FILE",       %thinBasic_ReturnNumber, CODEPTR(SPF_Exempt_File))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("SPF_INS",               %thinBasic_ReturnNumber, CODEPTR(SPF_INS))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("SPF_OVR",               %thinBasic_ReturnNumber, CODEPTR(SPF_OVR))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("SPF_LOOP_CHECK",        %thinBasic_ReturnNumber, CODEPTR(SPF_LOOPCHECK))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("SPF_OVR_REP",           %thinBasic_ReturnNumber, CODEPTR(SPF_OVR_REP))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("SPF_PARSE",             %thinBasic_ReturnNumber, CODEPTR(SPF_PARSE))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("SPF_REP",               %thinBasic_ReturnNumber, CODEPTR(SPF_REP))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("SPF_QUOTE$",            %thinBasic_ReturnString, CODEPTR(SPF_QUOTE))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("SPF_UNQUOTE$",          %thinBasic_ReturnString, CODEPTR(SPF_UNQUOTE))
         CALL DWORD hProc_thinBasic_LoadSymbol USING thinBasic_LoadSymbol("SPF_SHELL",             %thinBasic_ReturnNumber, CODEPTR(SPF_SHELL))
         CALL DWORD hProc_thinBasic_AddVariable USING thinBasic_AddVariable("OK",    "", 0,        %EquateTypeNumber)
         CALL DWORD hProc_thinBasic_AddVariable USING thinBasic_AddVariable("WARN",  "", 4,        %EquateTypeNumber)
         CALL DWORD hProc_thinBasic_AddVariable USING thinBasic_AddVariable("FAIL",  "", 8,        %EquateTypeNumber)
         CALL DWORD hProc_thinBasic_AddVariable USING thinBasic_AddVariable("ON",    "", 1,        %EquateTypeNumber)
         CALL DWORD hProc_thinBasic_AddVariable USING thinBasic_AddVariable("OFF",   "", 0,        %EquateTypeNumber)
         CALL DWORD hProc_thinBasic_AddVariable USING thinBasic_AddVariable("ERROR", "", 2,        %EquateTypeNumber)
         CALL DWORD hProc_thinBasic_AddVariable USING thinBasic_AddVariable("YES",   "", 1,        %EquateTypeNumber)
         CALL DWORD hProc_thinBasic_AddVariable USING thinBasic_AddVariable("NO",    "", 0,        %EquateTypeNumber)
         CALL DWORD hProc_thinBasic_AddVariable USING thinBasic_AddVariable("SYNC",  "", 0,        %EquateTypeNumber)
         CALL DWORD hProc_thinBasic_AddVariable USING thinBasic_AddVariable("ASYNC", "", 1,        %EquateTypeNumber)
         CALL DWORD hProc_thinBasic_AddVariable USING thinBasic_AddVariable("NORMAL","", 1,        %EquateTypeNumber)
         CALL DWORD hProc_thinBasic_AddVariable USING thinBasic_AddVariable("HIDDEN","", 0,        %EquateTypeNumber)
         CALL DWORD hProc_thinBasic_AddVariable USING thinBasic_AddVariable("ARG_OPT","", 1024,    %EquateTypeNumber)
         CALL DWORD hProc_thinBasic_AddVariable USING thinBasic_AddVariable("ARG_DEF","", 2048,    %EquateTypeNumber)
         CALL DWORD hProc_thinBasic_AddVariable USING thinBasic_AddVariable("ARG_VAR","", 4096,    %EquateTypeNumber)

         '---Run the program
         BasicPgm = JOIN$(gMaclines(), $CRLF) + $CRLF             ' Join the array together
         gMacroMode = %True                                       ' Flag we are in Macro Mode
         sPopReady                                                ' Ready for pop-up
         CALL DWORD hProc_thinBasic_Run USING thinBasic_RUN( _
                                     0&                             , _     'Will be used in future release
                                     BasicPgm                       , _     'File name or string Buffer
                                     %thinBasic_BufferType_IsScript , _     'Previous op is a File Name, or previous op is a string Buffer
                                     (1 OR 2)                       , _     'Options (See below)
                                     %FALSE                         , _     'Debug Mode      %TRUE/%FALSE
                                     %FALSE                         , _     'Log Mode        %TRUE/%FALSE
                                     %FALSE                         , _     'Obfuscate Mode  %TRUE/%FALSE
                                     1&                             , _     'Calling Program 1=thinBasic, 2=Console
                                     %FALSE                           _     'Dependancy Mode %TRUE/%FALSE
                                   ) TO lRet
                                                                  ' Option 1 = use base directory from the executable running thinBasic_Run
                                                                  '            else thinBasic itself
                                                                  ' Option 2 = In case of script run-time ERROR, DO NOT CLOSE calling application too.
                                                                  '            If this option is omitted, thinBasic will fire a PostQuitMessage

         '---Unload all modules, release all memory
         CALL DWORD hProc_thinBasic_Release USING thinBasic_Release(0&  _     'Will be used in future release
                                                                    ) TO lRet
         '---Free loaded thinCore library
         FreeLibrary(hLib_thinCore)
      ELSE                                                        '
         scError(%eFail, "Internal thinBasic functions not found")' Error
      END IF                                                      '
   ELSE                                                           '
      scError(%eFail, "thinBasic does not appear to be installed")' Say why we didn't do it
   END IF                                                         '
'   MexitSub                                                      ' Done
END SUB

SUB       DoTrace()
'---------- Display a trace line
LOCAL lText AS STRING
   IF gMacroTrace = 0 THEN EXIT SUB                               ' Leave quickly if not tracing
   IF gmacroTrace = 2 AND TP.Errflag = 0 THEN EXIT SUB            ' Also if ERR only and no error
   ltext = gMacroTHeader + " RC=" + FORMAT$(TP.ErrFlag) + " " + TP.ErrMsg
   DEBUG lText                                                    ' Display it
END SUB

FUNCTION GET_ARG() AS STRING
'---------- Get MACRO Paremeter(s)
LOCAL lIndex, lIndex2 AS EXT, t AS STRING
REGISTER i AS LONG
REGISTER j AS LONG
   lIndex = 0: lIndex2 = 999
   IF thinBasic_CheckOpenParens THEN                              ' Better have a Open (
      thinBasic_ParseNumber lIndex                                ' Get the Index
      IF thinBasic_CheckComma(%True, %True) THEN                  ' See if optional comma
         thinBasic_ParseNumber lIndex2                            ' Get 2nd operand if present
         IF thinBasic_CheckCloseParens THEN                       ' Better be a closing ) now
            gMacroTHeader = "Get_Arg$(" + FORMAT$(lIndex) + ", " + FORMAT$(lIndex2) + ") "
            IF lIndex2 <> 0 AND lIndex2 < lIndex THEN             ' In sequence
               FUNCTION = ""                                      '
               TP.ErrFlag = 8: TP.ErrMsg = "Invalid ARG indexes"  ' ZRC = Fail
               DoTrace                                            ' Do trace if needed
               EXIT FUNCTION                                      '
            END IF                                                '
         ELSE                                                     '
            EXIT FUNCTION                                         '
         END IF                                                   '
      ELSEIF ISFALSE thinBasic_CheckCloseParens THEN              ' Better be a closing )
         EXIT FUNCTION                                            '
      END IF
      TP.ErrFlag = 0: TP.ErrMsg = ""                              ' ZRC = OK default
      IF lIndex2 = 999 THEN                                       ' Just one operand?
         gMacroTHeader = "Get_Arg$(" + FORMAT$(lIndex) + ") "     ' Set trace header
         IF lIndex <= UBOUND(mOprands()) AND lIndex > 0 THEN      ' If valid index
            FUNCTION = mOprands(lIndex)                           '
         ELSEIF lIndex = 0 THEN                                   ' If index = 0
            FUNCTION = TRIM$(JOIN$(mOprands(), $SPC))             '
         ELSE                                                     ' Else return null
            FUNCTION = "": TP.ErrFlag = 8: TP.ErrMsg = "Invalid ARG index"
            DoTrace                                               ' Do trace if needed
            EXIT FUNCTION                                         '
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      ELSE                                                        ' Two argument numbers
         IF lIndex <= UBOUND(mOprands()) AND lIndex2 <= UBOUND(mOprands()) THEN      ' If valid index
            IF lIndex2 = 0 THEN                                   ' 2nd operand zero?
               j = 0                                              '
               FOR i = 1 TO 50                                    ' See how many operands
                  IF ISNOTNULL(mOprands(i)) THEN j = i            ' Save hi-water non-null parameter
               NEXT i                                             '
               lIndex2 = j                                        ' Set lIndex 2 to max operand number
            END IF                                                '
            IF lIndex > lIndex2  OR lIndex = 0 THEN               ' Illogical ?
               FUNCTION = ""                                      '
               TP.ErrFlag = 8: TP.ErrMsg = "Invalid ARG indexes"  ' ZRC = Fail
               DoTrace                                            ' Do trace if needed
               EXIT FUNCTION                                      '
            END IF                                                '
            FOR i = lIndex TO lIndex2                             ' Add operands together
               IF ISNOTNULL(mOprands(i)) THEN t += mOprands(i) + " "
            NEXT i                                                '
            FUNCTION = RTRIM$(t)                                  '
         ELSE                                                     ' Else return null
            FUNCTION = "": TP.ErrFlag = 8: TP.ErrMsg = "Invalid ARG indexes"
            DoTrace                                               ' Do trace if needed
            EXIT FUNCTION                                         '
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF
   END IF
END FUNCTION

FUNCTION GET_ARGCOUNT() AS EXT
'---------- Get number of macro parameters
REGISTER i AS LONG
REGISTER j AS LONG
   gMacroTHeader = "Get_ArgCount() "
   j = 0
   FOR i = 1 TO 50                                                ' See how many operands
      IF ISNOTNULL(mOprands(i)) THEN j = i                        ' Save hi-water non-null parameter
   NEXT i                                                         '
   FUNCTION = j: QErrorNull                                       ' Return value
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION GET_ARGKW() AS EXT
'---------- Get MACRO Keyword True/False
LOCAL lIndex AS EXT, t AS STRING
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL wvarname, wvarvalue AS WSTRING
LOCAL vvarvalue AS VARIANT
   lIndex = 0
   IF thinBasic_CheckOpenParens THEN                              ' Better have a Open (
      thinBasic_ParseString t                                     ' Get the Keyword
      IF thinBasic_CheckCloseParens THEN                          ' Better be a closing )
         QErrorNull                                               ' ZRC = OK default
         gMacroTHeader = "Get_Arg_KW(" + t + ") "                 ' Set trace header
         wvarname = UUCASE(t)                                     ' Build key
         vvarvalue = gParseTbl.Item(wvarname)                     ' Try the retrieval
         j = OBJRESULT                                            ' look at answer
         IF j > 0 THEN                                            ' Not found
            TP.ErrMsg = "Invalid Keyword: " + t                   ' KW not defined
            TP.ErrFlag = 8: FUNCTION = %False                     '
         ELSE                                                     ' Retrieved OK
            t = VARIANT$(vvarvalue)                               ' Return string
            TP.Errmsg = "": TP.ErrFlag = 0                        ' Set good RC
            IF LEFT$(t, 1) = "9" THEN                             ' An ALIAS entry?
               wvarname = MID$(t, 2)                              ' Pick up the truename
               vvarvalue = gParseTbl.Item(wvarname)               ' Retrieve the truename entry
               j = OBJRESULT                                      ' look at answer
               IF j > 0 THEN
                  FUNCTION = %False: TP.ErrFlag = 8: TP.ErrMsg = "Parse Table structure error"
                  DoTrace                                         '
               END IF                                             '
               t = VARIANT$(vvarvalue)                            ' Convert string
            ELSE                                                  ' OK, swap the truename status
               t = VARIANT$(vvarvalue)                            ' Convert string
            END IF                                                '
            FUNCTION = IIF(LEFT$(t, 1) = "1", %True, %False)      ' Present? Set True/False
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF
   END IF
END FUNCTION

FUNCTION GET_ARGKWGROUP() AS STRING
'---------- Get MACRO Keyword Group result
LOCAL lIndex AS EXT, t AS STRING
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL wvarname, wvarvalue AS WSTRING
LOCAL vvarvalue AS VARIANT
   lIndex = 0
   IF thinBasic_CheckOpenParens THEN                              ' Better have a Open (
      thinBasic_ParseString t                                     ' Get the Keyword
      IF thinBasic_CheckCloseParens THEN                          ' Better be a closing )
         QErrorNull                                               ' ZRC = OK default
         gMacroTHeader = "Get_Arg_KWGroup(" + t + ") "            ' Set trace header
         wvarname = "$$" + UUCASE(t)                              ' Build key
         vvarvalue = gParseTbl.Item(wvarname)                     ' Try the retrieval
         j = OBJRESULT                                            ' look at answer
         IF j > 0 THEN                                            ' Not found
            TP.ErrMsg = "Invalid Groupname: " + t                 ' Groupname not defined
            TP.ErrFlag = 8: FUNCTION = ""                         '
         ELSE                                                     ' Retrieved OK
            t = VARIANT$(vvarvalue)                               ' Return string
            t = PARSE$(t, "|", 2)                                 ' Get 2nd value
            TP.Errmsg = "": TP.ErrFlag = 0                        ' Set good RC
            FUNCTION = t                                          ' Return it
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF
   END IF
END FUNCTION

FUNCTION GET_ARGLRef() AS STRING
'---------- Get MACRO Parsed Line Reference
LOCAL lIndex AS EXT, t AS STRING
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL wvarname, wvarvalue AS WSTRING
LOCAL vvarvalue AS VARIANT
   lIndex = 0
   IF thinBasic_CheckOpenParens THEN                              ' Better have a Open (
      thinBasic_ParseNumber lIndex                                ' Get the Index
      IF thinBasic_CheckCloseParens THEN                          ' Better be a closing )
         QErrorNull                                               ' ZRC = OK default
         gMacroTHeader = "Get_Arg_LRef$(" + FORMAT$(lIndex) + ") "' Set trace header
         i = lIndex                                               ' Convert to LONG
         IF i > mgotlptr THEN                                     '
            FUNCTION = "": TP.ErrFlag = 8: TP.ErrMsg = "Invalid line reference index"
         ELSE                                                     ' Within range
            wvarname = "." + FORMAT$(i)                           ' Build key
            vvarvalue = gParseTbl.Item(wvarname)                  ' Try the retrieval
            j = OBJRESULT                                         ' look at answer
            IF j > 0 THEN                                         ' Not found
               TP.ErrMsg = "Invalid Parse Table structure"        ' Oops, should never happen
               TP.ErrFlag = 8: FUNCTION = ""                      '
            ELSE                                                  ' Retrieved OK
               FUNCTION = VARIANT$(vvarvalue)                     ' Return string
            END IF                                                '
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF
   END IF
END FUNCTION

FUNCTION GET_ARGNumLit() AS STRING
'---------- Get MACRO Parsed Numeric literal
LOCAL lIndex AS EXT, t AS STRING
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL wvarname, wvarvalue AS WSTRING
LOCAL vvarvalue AS VARIANT
   lIndex = 0
   IF thinBasic_CheckOpenParens THEN                              ' Better have a Open (
      thinBasic_ParseNumber lIndex                                ' Get the Index
      IF thinBasic_CheckCloseParens THEN                          ' Better be a closing )
         QErrorNull                                               ' ZRC = OK default
         gMacroTHeader = "Get_Arg_NumLit$(" + FORMAT$(lIndex) + ") " ' Set trace header
         i = lIndex                                               ' Convert to LONG
         IF i > mgotnum THEN                                      '
            FUNCTION = "": TP.ErrFlag = 8: TP.ErrMsg = "Invalid numeric literal index"
         ELSE                                                     ' Within range
            wvarname = "#" + FORMAT$(i)                           ' Build key
            vvarvalue = gParseTbl.Item(wvarname)                  ' Try the retrieval
            j = OBJRESULT                                         ' look at answer
            IF j > 0 THEN                                         ' Not found
               TP.ErrMsg = "Invalid Parse Table structure"        ' Oops, should never happen
               TP.ErrFlag = 8: FUNCTION = ""                      '
            ELSE                                                  ' Retrieved OK
               FUNCTION = VARIANT$(vvarvalue)                     ' Return string
            END IF                                                '
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF
   END IF
END FUNCTION

FUNCTION GET_ARGTag() AS STRING
'---------- Get MACRO Parsed Tag operand
LOCAL lIndex AS EXT, t AS STRING
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL wvarname, wvarvalue AS WSTRING
LOCAL vvarvalue AS VARIANT
   lIndex = 0
   IF thinBasic_CheckOpenParens THEN                              ' Better have a Open (
      thinBasic_ParseNumber lIndex                                ' Get the Index
      IF thinBasic_CheckCloseParens THEN                          ' Better be a closing )
         QErrorNull                                               ' ZRC = OK default
         gMacroTHeader = "Get_Arg_Tag$(" + FORMAT$(lIndex) + ") " ' Set trace header
         i = lIndex                                               ' Convert to LONG
         IF i > mgottag THEN                                      '
            FUNCTION = "": TP.ErrFlag = 8: TP.ErrMsg = "Invalid tag operand index"
         ELSE                                                     ' Within range
            wvarname = ":" + FORMAT$(i)                           ' Build key
            vvarvalue = gParseTbl.Item(wvarname)                  ' Try the retrieval
            j = OBJRESULT                                         ' look at answer
            IF j > 0 THEN                                         ' Not found
               TP.ErrMsg = "Invalid Parse Table structure"        ' Oops, should never happen
               TP.ErrFlag = 8: FUNCTION = ""                      '
            ELSE                                                  ' Retrieved OK
               FUNCTION = VARIANT$(vvarvalue)                     ' Return string
            END IF                                                '
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF
   END IF
END FUNCTION

FUNCTION GET_ARGTextLit() AS STRING
'---------- Get MACRO Parsed Text literal
LOCAL lIndex AS EXT, t AS STRING
REGISTER i AS LONG
REGISTER j AS LONG
LOCAL wvarname, wvarvalue AS WSTRING
LOCAL vvarvalue AS VARIANT
   lIndex = 0
   IF thinBasic_CheckOpenParens THEN                              ' Better have a Open (
      thinBasic_ParseNumber lIndex                                ' Get the Index
      IF thinBasic_CheckCloseParens THEN                          ' Better be a closing )
         QErrorNull                                               ' ZRC = OK default
         gMacroTHeader = "Get_Arg_TextLit$(" + FORMAT$(lIndex) + ") " ' Set trace header
         i = lIndex                                               ' Convert to LONG
         IF i > mgotlit THEN                                      '
            FUNCTION = "": TP.ErrFlag = 8: TP.ErrMsg = "Invalid text literal index"
         ELSE                                                     ' Within range
            wvarname = "$" + FORMAT$(i)                           ' Build key
            vvarvalue = gParseTbl.Item(wvarname)                  ' Try the retrieval
            j = OBJRESULT                                         ' look at answer
            IF j > 0 THEN                                         ' Not found
               TP.ErrMsg = "Invalid Parse Table structure"        ' Oops, should never happen
               TP.ErrFlag = 8: FUNCTION = ""                      '
            ELSE                                                  ' Retrieved OK
               FUNCTION = VARIANT$(vvarvalue)                     ' Return string
            END IF                                                '
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF
   END IF
END FUNCTION

FUNCTION Get_ClrLine() AS STRING
'---------- Get color text for a single line number
LOCAL lText AS STRING, lno AS EXT, i AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Get_Clr_Line$(" + FORMAT$(lno) + ") "
         i = INT(lno)                                             ' Make long
         IF i > 0 AND i <= TP.LastLine THEN                       ' Within range?
            IF ISFALSE TP.LFlagData(i) THEN                       ' Ineligible line?
               FUNCTION = "": TP.ErrFlag = 8:TP.ErrMsg = "Not a Data line" ' Else pass back null
            ELSE                                                  '
               LText = TP.LAttrGet(i)                             ' An Attr line?
               IF TRIM$(LText) = "" THEN LText = SPACE$(TP.LTxtLen(i)) ' fudge one
               REPLACE ANY "1234" WITH "YBGR" IN lText            ' Convert colors to 'english'
               FUNCTION = lText: QErrorNull                       ' Yes, fetch the text
            END IF                                                '
         ELSE                                                     '
            FUNCTION = "": TP.ErrFlag = 8: TP.ErrMsg = "Invalid line pointer"  ' Else pass back null
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION Get_CSRCOL() AS EXT
'---------- Get the column number of the cursor
   gMacroTHeader = "Get_CsrCol() "
   FUNCTION = IIF(ISFALSE TP.CursData, 0, TP.CsrCol - gLNPadCol + TP.Offset) ' Return the cursor column
   QErrorNull                                                     ' ZRC = OK
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_CSRLIN() AS EXT
'---------- Get the line number of the cursor
   gMacroTHeader = "Get_Csr_LPtr() "
   FUNCTION = TP.sGet(TP.CsrRow): QErrorNull                      ' Return the cursor line
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_CURRLINE() AS STRING
'---------- Get the Current text line
   gMacroTHeader = "Get_CurrLine$() "
   FUNCTION  = TP.CursLine: QErrorNull                            ' Return Current line
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_CURRPATH() AS STRING
'---------- Get the Current path
   gMacroTHeader = "Get_CurrPath$() "
   FUNCTION  = CURDIR$: QErrorNull                                ' Return Current Path
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_CURRWORD() AS STRING
'---------- Get the Current cursor word
   gMacroTHeader = "Get_CurrWord$() "
   FUNCTION  = TRIM$(TP.CursWord): QErrorNull                     ' Return Current Word
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_DESTL1() AS EXT
'---------- Get the Dest line command start line number
LOCAL lno AS EXT
   gMacroTHeader = "Get_Dest1_LPtr() "
   lno = IIF(gMacRange, IIF(gMacDFrom = 0, 0, gMacDFrom), 0)      ' If a range waiting, return the FromLin
   IF lno <> 0 THEN TP.UpdLControl(INT(lno))                      ' Reset the LLCtl area
   FUNCTION = lno: QErrorNull                                     ' Return the result
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_DESTL2() AS EXT
'---------- Get the Dest line command end line number
LOCAL lno AS EXT
   gMacroTHeader = "Get_Dest2_LPtr() "
   lno = IIF(gMacRange, IIF(gMacDTo = 0, 0, gMacDTo), 0)          ' If a range waiting, return the ToLin
   IF lno <> 0 THEN TP.UpdLControl(INT(lno))                      ' Reset the LLCtl area
   FUNCTION = lno: QErrorNull                                     ' Return the result
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_DESTLCMD() AS STRING
'---------- Get the dest line command
   gMacroTHeader = "Get_Dest_LCmd$() "
   FUNCTION = IIF$(gMacRange, gMacDCmd, "")                       ' If a range waiting, return the LCmd
   QErrorNull                                                     ' ZRC = OK
END FUNCTION

FUNCTION Get_DESTLMOD$() AS STRING
'---------- Get the Dest line command Post Operand
LOCAL op AS STRING
   gMacroTHeader = "Get_Dest_LMod$() "
   op = "  "                                                      ' Set default
   IF gMacRange THEN                                              ' If a range
      IF BIT(gMacDFlag, %lCmdX) THEN MID$(op, 1, 1) = "-"         ' Set +/- as appropriate
      IF BIT(gMacDFlag, %lCmdNX) THEN MID$(op, 1, 1) = "+"        '
      IF BIT(gMacDFlag, %lCmdR) THEN MID$(op, 2, 1) = "R"         '
   END IF                                                         '
   FUNCTION = op: QErrorNull                                      '
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_DESTOP() AS EXT
'---------- Get the Dest command repeat value
LOCAL rpt AS EXT
   gMacroTHeader = "Get_Dest_Op() "
   rpt = IIF(gMacRange, IIF(gMacDRpt = 0, 0, gMacDRpt), 0)        ' If a range waiting, return the FromLin
   FUNCTION = rpt: QErrorNull                                     ' Return the result
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_ESource() AS STRING
'---------- Get the find string for a CHANGE
LOCAL lText AS STRING, lno AS EXT, i AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Get_E_Source$()"                        '
         FUNCTION = gMacFString                                   ' Yes, fetch the text
         QErrorNull                                               ' Set ZRC
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION Get_EXEPath() AS STRING
'---------- Get the current EXE Path
   gMacroTHeader = "Get_EXEPath$() "
   FUNCTION = ENV.EXEPath: QErrorNull                             ' Pass back EXE path
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_FileBase() AS STRING
'---------- Get the current File Base name
   gMacroTHeader = "Get_FileBase$() "
   FUNCTION = TP.TIPBase: QErrorNull                              ' Pass back File Base name
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_FileDate() AS STRING
'---------- Get the current File Date
   gMacroTHeader = "Get_FileDate$() "
   FUNCTION = TP.TIPDate: QErrorNull                              ' Pass back File date
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_FileExt() AS STRING
'---------- Get the current File Extension
   gMacroTHeader = "Get_FileExt$() "
   FUNCTION = TP.TIPExtn: QErrorNull                              ' Pass back File extension
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_FileName() AS STRING
'---------- Get the current File Name
   gMacroTHeader = "Get_FileName$() "
   FUNCTION = TP.TIPFile: QErrorNull                              ' Pass back File name
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_FilePath() AS STRING
'---------- Get the current File Path
   gMacroTHeader = "Get_FilePath$() "
   FUNCTION = TP.TIPPath: QErrorNull                              ' Pass back File path
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_FileTime() AS STRING
'---------- Get the current File Time
   gMacroTHeader = "Get_FileTime$() "
   FUNCTION = TP.TIPTime: QErrorNull                              ' Pass back File time
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_FINDCOL() AS EXT
'---------- Get the last FIND column number
   gMacroTHeader = "Get_Find_Col() "
   FUNCTION = TP.cfFCol: QErrorNull                               ' Return ZFINDCOL
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_FINDLEN() AS EXT
'---------- Get the last FIND found length
   gMacroTHeader = "Get_Find_Len() "
   FUNCTION = TP.cfFLen: QErrorNull                               ' Return ZFINDLEN
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_FINDLIN() AS EXT
'---------- Get the last FIND line number
   gMacroTHeader = "Get_Find_LPtr() "
   FUNCTION = TP.cfFLine: QErrorNull                              ' Return the line number
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_FIRST() AS EXT
'---------- Get the top line
   gMacroTHeader = "Get_First_LPtr() "
   FUNCTION = IIF(TP.LastLine = 2, 0, 1): QErrorNull              ' Return top line or 0 if empty file
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_FullPath() AS STRING
'---------- Return the fully qualified filename
LOCAL lText AS STRING, i, j AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' A Open (
      thinBasic_ParseString lText                                 ' Then get the base filename
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "SPF_FullPath$(" + ltext + ") "          ' Set up trace
         IF INSTR(lText, ANY ":\") > 0 OR ISNULL(lText) THEN      ' Null or already qualified
            FUNCTION = lText: TP.ErrFlag = 8: TP.ErrMsg = "Name already appears qualified" ' Return code
         ELSE                                                     '
            FUNCTION = sGetDefDir + lText: QErrorNull             '
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION Get_GBLNUM() AS EXT
'---------- Get a GBL variable
LOCAL varprefix, varname AS STRING, i AS LONG, wvarname AS WSTRING, vvarvalue AS VARIANT, level, varvalue AS EXT
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      IF thinBasic_DetermineType = %thinBasic_ReturnNumber THEN   ' Is next param a number?
         thinBasic_ParseNumber level                              ' Then get the table number
         varprefix = FORMAT$(level) + ","                         ' Format it
         IF thinBasic_CheckComma() THEN                           ' A comma?
            '                                                     ' Better be there
         END IF                                                   '
      ELSE                                                        ' No table prefix - assign zero
         varprefix = "0,"                                         '
      END IF
      thinBasic_ParseString varname                               ' Then get the variable name
      varname = varprefix + varname                               ' Add the table prefix
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Get_Gbl_Num(" + varname + ") "
         QErrorNull                                               ' Default answers
         IF ISNULL(varname) THEN                                  ' Ignore null requests
            FUNCTION = 0                                          ' Return OK
            DoTrace                                               ' Do trace if needed
            EXIT FUNCTION                                         '
         END IF
         wvarname = varname                                       ' Convert to WSTRING
         vvarvalue = gNumVar.Item(wvarname)                       ' Try the retrieval
         i = OBJRESULT                                            ' look at answer
         IF i > 0 THEN                                            ' Error?
            FUNCTION = 0: TP.ErrFlag = 8: TP.ErrMsg = "Unable to find variable" ' Return code
         ELSE                                                     '
            varvalue = VARIANT#(vvarvalue)                        '
            FUNCTION = varvalue                                   ' Return OK
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION Get_GBLNUMCOUNT() AS EXT
'---------- Get count of GLB NUM variables
   gMacroTHeader = "Get_Gbl_Num_Count() "
   FUNCTION = gNumVar.Count: QErrorNull                           ' Return count of NumVars
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_GBLNUMNAME() AS STRING
'---------- Get name of a variable entry
LOCAL varname, varvalue AS STRING, i AS LONG, lno AS EXT, wvarname AS WSTRING, vvarvalue AS VARIANT
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the index number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Get_Gbl_Num_Name$(" + FORMAT$(lno) + ") "
         i = INT(lno)                                             ' Make long
         gNumVar.Entry(i, wvarname, vvarvalue)                    '
         i = OBJRESULT                                            ' Get result
         IF i = 0 THEN                                            ' If OK
            varname = wvarname                                    ' Return name
            varname = MID$(varname, INSTR(varname, ",") + 1)      ' Strip off table number
            FUNCTION = varname: QErrorNull                        '
         ELSE                                                     '
            FUNCTION = "": TP.ErrFlag = 8: TP.ErrMsg = "No NUM Variable entry"
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION Get_GBLNUMTABLENAME() AS STRING
'---------- Get name plus table of a variable entry
LOCAL varname, varvalue AS STRING, i AS LONG, lno AS EXT, wvarname AS WSTRING, vvarvalue AS VARIANT
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the index number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Get_Gbl_Num_TableName$(" + FORMAT$(lno) + ") "
         i = INT(lno)                                             ' Make long
         gNumVar.Entry(i, wvarname, vvarvalue)                    '
         i = OBJRESULT                                            ' Get result
         IF i = 0 THEN                                            ' If OK
            varname = wvarname                                    ' Return name
            FUNCTION = varname: QErrorNull                        '
         ELSE                                                     '
            FUNCTION = "": TP.ErrFlag = 8: TP.ErrMsg = "No NUM Variable entry"
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION Get_GBLSTR() AS STRING
'---------- Get a GBL variable
LOCAL varprefix, varname, varvalue AS STRING, i AS LONG, level AS EXT, wvarname AS WSTRING, vvarvalue AS VARIANT
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      IF thinBasic_DetermineType = %thinBasic_ReturnNumber THEN   ' Is next param a number?
         thinBasic_ParseNumber level                              ' Then get the table number
         varprefix = FORMAT$(level) + ","                         ' Format it
         IF thinBasic_CheckComma() THEN                           ' A comma?
            '                                                     ' Better be there
         END IF                                                   '
      ELSE                                                        ' No table prefix - assign zero
         varprefix = "0,"                                         '
      END IF
      thinBasic_ParseString varname                               ' Then get the variable name
      varname = varprefix + varname                               ' Add the table prefix
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Get_Gbl_Str$(" + varname + ") "
         QErrorNull                                               ' Default answers
         IF ISNULL(varname) THEN                                  ' Ignore null requests
            FUNCTION = "":                                        ' Return OK
            DoTrace                                               ' Do trace if needed
            EXIT FUNCTION                                         '
         END IF
         wvarname = varname                                       ' Convert to WSTRING
         vvarvalue = gStrVar.Item(wvarname)                       ' Try the retrieval
         i = OBJRESULT                                            ' look at answer
         IF i > 0 THEN                                            ' Error?
            FUNCTION = "": TP.ErrFlag = 8: TP.ErrMsg = "Unable to find variable"
         ELSE                                                     '
            varvalue = VARIANT$(vvarvalue)                        '
            FUNCTION = varvalue                                   ' Return OK
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION Get_GBLSTRCOUNT() AS EXT
'---------- Get count of GLB STR variables
   gMacroTHeader = "Get_Gbl_Str_Count() "
   FUNCTION = gStrVar.Count: QErrorNull                           ' Return count of StrVars
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_GBLSTRNAME() AS STRING
'---------- Get name of a variable entry
LOCAL varname, varvalue AS STRING, i AS LONG, lno AS EXT, wvarname AS WSTRING, vvarvalue AS VARIANT
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the index number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Get_Gbl_Str_Name$(" + FORMAT$(lno) + ") "
         i = INT(lno)                                             ' Make long
         gStrVar.Entry(i, wvarname, vvarvalue)                    '
         i = OBJRESULT                                            ' Get result
         IF i = 0 THEN                                            ' If OK
            varname = wvarname                                    ' Return name
            varname = MID$(varname, INSTR(varname, ",") + 1)      ' Strip off table number
            FUNCTION = varname: QErrorNull                        '
         ELSE                                                     '
            FUNCTION = "": TP.ErrFlag = 8: TP.ErrMsg = "No STR Variable entry"
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION Get_GBLSTRTABLENAME() AS STRING
'---------- Get name plus table of a variable entry
LOCAL varname, varvalue AS STRING, i AS LONG, lno AS EXT, wvarname AS WSTRING, vvarvalue AS VARIANT
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the index number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Get_Gbl_Str_TableName$(" + FORMAT$(lno) + ") "
         i = INT(lno)                                             ' Make long
         gStrVar.Entry(i, wvarname, vvarvalue)                    '
         i = OBJRESULT                                            ' Get result
         IF i = 0 THEN                                            ' If OK
            varname = wvarname                                    ' Return name
            FUNCTION = varname: QErrorNull                        '
         ELSE                                                     '
            FUNCTION = "": TP.ErrFlag = 8: TP.ErrMsg = "No STR Variable entry"
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION Get_INIPath() AS STRING
'---------- Get the current INI Path
   gMacroTHeader = "Get_INIPath$() "
   FUNCTION = ENV.INIPath: QErrorNull                             ' Pass back INI path
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_Label() AS STRING
'---------- Get label for a single line number
LOCAL lText AS STRING, lno AS EXT, i AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Get_Label$(" + FORMAT$(lno) + ") "
         i = INT(lno)                                             ' Make long
         IF i > 0 AND i <= TP.LastLine THEN                       ' Within range?
            IF ISFALSE TP.LFlagData(i) THEN                       ' Not a data line?
               FUNCTION = "": TP.ErrFlag = 8                      ' Else pass back null
            ELSE                                                  '
               FUNCTION = TRIM$(TP.LLblGet(i))                    ' Yes, fetch the label
               QErrorNull                                         ' Set ZRC
            END IF                                                '
         ELSE                                                     '
            FUNCTION = "": TP.ErrFlag = 8                         ' Else pass back null
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION Get_LAST() AS EXT
'---------- Get the last line
   gMacroTHeader = "Get_Last_LPtr() "
   FUNCTION = IIF(TP.LastLine = 2, 0, TP.LastLine): QErrorNull    ' Return last line or 0 if empty file
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_LASTREAL() AS EXT
'---------- Get the last real line number
   gMacroTHeader = "Get_Last_Real() "
   FUNCTION = TP.LastReal: QErrorNull                             ' Return the real line count
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_LBOUND() AS EXT
'---------- Get the left BOUNDS value
   gMacroTHeader = "Get_LBound() "
   FUNCTION = TP.PrfBndLeft: QErrorNull                           ' Return the left bounds
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION GET_LNum() AS EXT
'---------- Convert Internal to External line number
LOCAL lno AS EXT, i AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Get_LNum(" + FORMAT$(lno) + ") "
         i = INT(lno)                                             ' Make long
         IF i > 0 AND i < TP.LastLine THEN                        ' Valid range?
            FUNCTION = VAL(TP.LLNumGet(i)): QErrorNull            ' Get the external line number
         ELSE                                                     '
            FUNCTION = 0: TP.ErrFlag = 8: TP.ErrMsg = "Invalid line pointer"
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION Get_LINE() AS STRING
'---------- Get text for a single line number
LOCAL lText AS STRING, lno AS EXT, i AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Get_Line$(" + FORMAT$(lno) + ") "
         i = INT(lno)                                             ' Make long
         IF i > 0 AND i <= TP.LastLine THEN                       ' Within range?
            FUNCTION = TP.LTxtGet(i)                              ' Yes, fetch the text
            QErrorNull                                            ' Set ZRC
         ELSE                                                     '
            FUNCTION = "": TP.ErrMsg = "Invalid line pointer"     ' Else pass back null
            TP.ErrFlag = 8                                        ' Set ZRC
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION Get_LINELen() AS EXT
'---------- Get text length a single line number
LOCAL lText AS STRING, lno AS EXT, i AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Get_Line_Len(" + FORMAT$(lno) + ") "
         i = INT(lno)                                             ' Make long
         IF i > 0 AND i <= TP.LastLine THEN                       ' Within range?
            IF ISTRUE (TP.LFlagGet(i) AND %NonTypable) THEN       ' Ineligible line?
               FUNCTION = 0                                       ' Else pass back null
               TP.ErrFlag = 8: TP.ErrMsg = "Non modifyable line"  ' Set ZRC
            ELSE                                                  '
               FUNCTION = TP.LTxtLen(i)                           ' Yes, fetch the text length
               QErrorNull                                         ' Set ZRC
            END IF                                                '
         ELSE                                                     '
            FUNCTION = 0: TP.ErrMsg = "Invalid line pointer"      ' Else pass back null
            TP.ErrFlag = 8                                        ' Set ZRC
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION Get_LINETYPE() AS STRING
'---------- Get text version of Line Type
LOCAL lText AS STRING, lno AS EXT, i, j AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Get_Line_Type$(" + FORMAT$(lno) + ") "
         lText = "UNKN"                                           ' Default unknown
         i = INT(lno)                                             ' Make long
         IF i > 0 AND i <= TP.LastLine THEN                       ' Within range?
            QErrorNull                                            ' ZRC = OK
            lText = SWITCH$( _                                    ' Return the type
               ISTRUE (TP.LFlagGet(i) AND %Data),       "DATA", _ '
               ISTRUE (TP.LFlagGet(i) AND %Xclude),     "EXCL", _ '
               ISTRUE (TP.LFlagGet(i) AND %Top),        "TOP", _  '
               ISTRUE (TP.LFlagGet(i) AND %Bottom),     "BOT", _  '
               ISTRUE (TP.LFlagGet(i) AND %Tabs),       "TABS", _ '
               ISTRUE (TP.LFlagGet(i) AND %Bounds),     "BNDS", _ '
               ISTRUE (TP.LFlagGet(i) AND %Cols),       "COLS", _ '
               ISTRUE (TP.LFlagGet(i) AND %Word),       "WORD", _ '
               ISTRUE (TP.LFlagGet(i) AND %Mark),       "MARK", _ '
               ISTRUE (TP.LFlagGet(i) AND %Mask),       "MASK", _ '
               ISTRUE (TP.LFlagGet(i) AND %Prof),       "PROF", _ '
               ISTRUE (TP.LFlagGet(i) AND %File),       "FILE")   '
            IF ISTRUE (TP.LFlagGet(i) AND %Note) THEN             ' Enhance NOTE with the Note type
               j = TP.LWrk2Get(i)                                 ' Get the LWrk2 value
               IF j THEN                                          ' Other than default note?
                  lText = MID$($Upper, j, 1) + "NOTE"             ' Add type of xNOTE
               ELSE                                               ' Just a plain NOTE
                  lText = "NOTE"                                  '
               END IF                                             '
            END IF                                                '

         ELSE                                                     '
            TP.ErrFlag = 8: TP.ErrMsg = "Invalid line pointer"    ' ZRC = fail
         END IF                                                   '
         FUNCTION = lText                                         ' Return answer
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION Get_LOCLIN() AS EXT
'---------- Get the LOCATE line number
   gMacroTHeader = "Get_LOC_Lptr() "
   FUNCTION = IIF(TP.LocLineR = 0, 0, TP.LocLineR)                ' Return zero if no ZLOC else the line number
   QErrorNull                                                     ' ZRC = OK
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_MacName() AS STRING
'---------- Get the current Macro name
   gMacroTHeader = "Get_MacName$() "
   FUNCTION  = gMacroName                                         ' Return Macro Name
   QErrorNull                                                     ' ZRC = OK
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_MODIFIED() AS EXT
'---------- Get the modified status
   gMacroTHeader = "Get_Modified() "
   FUNCTION = IIF(IsTPModdFlag, %True, %False)                    ' Return modified status
   QErrorNull                                                     ' ZRC = OK
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_MODIFIEDFILENAME() AS EXT
'---------- Get the modified status of a file
LOCAL lText, XList AS STRING, i AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseString lText                                 ' Then get the string
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "SPF_Modified_FileName(" + ltext + ") "  ' Build trace
         IF ISTRUE (TP.TMode AND %MMEdit) THEN                    ' In MEdit mode
            i = TP.MeditTbl("S", ltext)                           ' Get MEdit table index
            IF i = 0 THEN                                         ' Not found?
               FUNCTION = %False: TP.ErrFlag = 8: TP.ErrMsg = "Filename not found"  ' Return
            ELSE                                                  '
               FUNCTION = TP.MEditFlagGet(i): QErrorNull          ' Return the modified flag
            END IF                                                '
         ELSE                                                     '
            FUNCTION = IIF(IsTPModdFlag, %True, %False)           ' Return modified status
            QErrorNull                                            ' ZRC = OK
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION Get_MSG() AS STRING
'---------- Get the Error message
   gMacroTHeader = "Get_Msg$() "
   FUNCTION = TRIM$(TP.ErrMsg)                                    ' Error message
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION GET_LPtr() AS EXT
'---------- Convert External to Internal line number
LOCAL lno AS STRING, i AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseString lno                                   ' Then get the line number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Get_LPtr(" + lno + ") "
         i = TP.LineNoRef(lno)                                    ' Go look it up
         IF i = -1 THEN i = 0                                     ' Make the -1 into a zero
         FUNCTION = i                                             ' Return it
         TP.ErrFlag = IIF(i <> 0, 0, 8)                           ' Set ZRC
         TP.ErrMsg = IIF$(i <> 0, "", "Invalid line number")      ' Set msg
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION GET_NextLPtr() AS EXT
'---------- Get the Next +/- line pointer of a type
LOCAL lptr, lincr AS EXT, ltype, ltype2, t AS STRING
REGISTER i AS LONG
REGISTER j AS LONG
   lincr = 1: ltype = "DATA"
   IF thinBasic_CheckOpenParens THEN                              ' Better have a Open (
      thinBasic_ParseNumber lptr                                  ' Get the current line pointer
      IF thinBasic_CheckComma(%True, %True) THEN                  ' See if optional comma
         thinBasic_ParseNumber lincr                              ' Get 2nd operand into the increment
         IF thinBasic_CheckComma(%True, %True) THEN               ' See if optional data type
            thinBasic_ParseString ltype                           ' Fetch the optional type
            IF thinBasic_CheckCloseParens THEN                    ' End of a 3 parameter set?
               gMacroTHeader = "Get_Next_Lptr(" + FORMAT$(lptr) + ", " + FORMAT$(lincr) + ", " + ltype + ") "
               IF IsNE(ltype, "DATA") AND IsNE(ltype, "SPECIAL") AND LEFT$(ltype, 1) <> ":" THEN  ' Valid type?
                  FUNCTION = 0                                    '
                  TP.ErrFlag = 8: TP.ErrMsg = "Invalid Line type argument"  ' Fail it
                  DoTrace                                         ' Do trace if needed
                  EXIT FUNCTION                                   '
               END IF                                             '
            ELSE                                                  '
               EXIT FUNCTION                                      '
            END IF                                                '
         ELSEIF thinBasic_CheckCloseParens THEN                   ' Closing ) after 2nd param
            gMacroTHeader = "Get_Next_Lptr(" + FORMAT$(lptr) + ", " + FORMAT$(lincr) + ") "
         END IF                                                   '
      ELSEIF thinBasic_CheckCloseParens THEN                      ' Better be a closing ) after 1st param
         gMacroTHeader = "Get_Next_Lptr(" + FORMAT$(lptr) + ") "
      END IF                                                      '

      IF lptr < 1 OR lptr >= TP.LastLine THEN                     ' Valid starting line pointer?
         FUNCTION = 0: TP.ErrFlag = 8: TP.ErrMsg = "Invalid starting line pointer argument"  ' Fail it
         DoTrace                                                  ' Do trace if needed
         EXIT FUNCTION                                            '
      END IF                                                      '
      IF lincr = 0 THEN                                           ' Must not be zero
         FUNCTION = 0: TP.ErrFlag = 8: TP.ErrMsg = "Invalid number of lines argument"  ' Fail it
         DoTrace                                                  ' Do trace if needed
         EXIT FUNCTION                                            '
      END IF                                                      '

      IF LEFT$(ltype, 1) = ":" THEN ltype2 = LSET$(UUCASE(ltype), 8) ' Pad to 8
      i = lincr: j = lptr                                         ' Copy params

      IF i > 0 THEN                                               ' Moving forward?
         DO WHILE i                                               ' While more lines to go forward
            IF LEFT$(ltype, 1) = ":" THEN                         ' Doing Tag search?
               DO                                                 ' Yes, loop foreward till a Data line is seen
                  j = MIN(j + 1, TP.LastLine)                     '
                  t = TP.LTagGet(j)
               LOOP UNTIL TP.LTagGet(j) = ltype2 OR j = TP.LastLine '
            ELSEIF IsEQ(ltype, "DATA") THEN                       ' Doing Data lines?
               DO                                                 ' Yes, loop foreward till a Data line is seen
                  j = MIN(j + 1, TP.LastLine)                     '
               LOOP UNTIL TP.LFlagData(j) OR j = TP.LastLine      '
            ELSE                                                  '
               DO                                                 ' Yes, loop foreward till a non-Data line is seen
                  j = MIN(j + 1, TP.LastLine)                     '
               LOOP UNTIL ISFALSE TP.LFlagData(j) OR j = TP.LastLine '
            END IF                                                '

            IF j = TP.LastLine THEN EXIT DO                       ' OK if Row is at the end, stop
            DECR i                                                ' Else DECR the number we're supposed to do
         LOOP                                                     ' and loop till done

         IF j = TP.LastLine THEN                                  ' End of file?
            FUNCTION = 0: TP.ErrFlag = 8: TP.ErrMsg = "Bottom of file reached"
         ELSE                                                     '
            FUNCTION = j: QErrorNull                              '
         END IF                                                   '
         DoTrace                                                  ' Dump trace

      ELSE                                                        ' Moving backward
         DO WHILE i                                               ' While more lines to go forward
            IF LEFT$(ltype, 1) = ":" THEN                         ' Doing Tag search?
               DO                                                 ' Yes, loop foreward till a Data line is seen
                  j = MAX(j - 1, 1)                               '
               LOOP UNTIL TP.LTagGet(j) = ltype2 OR j = 1         '
            ELSEIF IsEQ(ltype, "DATA") THEN                       ' Doing Data lines?
               DO                                                 ' Yes, loop backward till a Data line is seen
                  j = MAX(j - 1, 1)                               '
               LOOP UNTIL TP.LFlagData(j) OR j = 1                '
            ELSE                                                  '
               DO                                                 ' Yes, loop backward till a Data line is seen
                  j = MAX(j - 1, 1)                               '
               LOOP UNTIL ISFALSE TP.LFlagData(j) OR j = 1        '
            END IF                                                '

            IF j = 1 THEN EXIT DO                                 ' OK if Row is at the top, stop
            INCR i                                                ' Else INCR the number we're supposed to do
         LOOP                                                     ' and loop till done

         IF j = 1 THEN                                            ' End of file?
            FUNCTION = 0: TP.ErrFlag = 8: TP.ErrMsg = "Top of file reached"
         ELSE                                                     '
            FUNCTION = j: QErrorNull                              '
         END IF                                                   '
         DoTrace                                                  ' Dump trace
      END IF
   END IF
END FUNCTION

FUNCTION Get_ProfString() AS STRING
'---------- Get a Profile string
LOCAL lText AS STRING
   IF thinBasic_CheckOpenParens() THEN                            ' A Open (
      thinBasic_ParseString lText                                 ' Then get the string
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Get_Profile$(" + ltext + ") "           '
         QErrorNull                                               ' Assume all OK for now
         FUNCTION = TP.PrfQuery(lText)                            ' Query Profile for an answer
         DoTrace                                                  ' Do trace if needed
      END IF
   END IF
END FUNCTION

FUNCTION Get_RBOUND() AS EXT
'---------- Get the right BOUNDS value
   gMacroTHeader = "Get_RBound() "
   FUNCTION = IIF(TP.PrfBndRight = 0, TP.MaxLength, TP.PrfBndRight) ' Return the right bounds, or maxlength if >>
   QErrorNull                                                     ' ZRC = OK
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_RC() AS EXT
'---------- Get the Error code
   gMacroTHeader = "Get_RC() "
   FUNCTION = TP.ErrFlag                                          ' Return Code
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_SelCol() AS EXT
'---------- Get the selected starting column
   gMacroTHeader = "Get_SELECT_Col() "
   FUNCTION = IIF(IsTPMarkActive, TP.MarkSCol, 0): QErrorNull     ' Copy starting column
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_SelFirst() AS EXT
'---------- Get the selected starting line
   gMacroTHeader = "Get_SELECT_First_LPtr() "
   FUNCTION = IIF(IsTPMarkActive, TP.MarkSLin, 0): QErrorNull     ' Copy starting line
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_SelLast() AS EXT
'---------- Get the selected Last line
   gMacroTHeader = "Get_SELECT_Last_LPtr() "
   FUNCTION = IIF(IsTPMarkActive, TP.MarkELin, 0): QErrorNull     ' Copy ending line
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_SelLen() AS EXT
'---------- Get the selected length
   gMacroTHeader = "Get_SELECT_Len() "
   FUNCTION = IIF(IsTPMarkActive, TP.MarkECol - TP.MarkSCol + 1, 0): QErrorNull ' Return length
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_SESSIONTYPE() AS STRING
'---------- Get Tab type
LOCAL lText AS STRING
   gMacroTHeader = "Get_Session_Type$() "
   lText = SWITCH$(IsBrowse, "BROWSE", IsView, "VIEW", IsClip, "CLIP-EDIT", IsSetEdit, "SET-EDIT", %True, "EDIT")
   IF lText = "VIEW" AND TP.TIPROStat THEN lText = "RDONLY"
   IF IsMedit THEN lText = "MEDIT"                                ' If MEdit, re-do it all
   FUNCTION = lText
   QErrorNull                                                     ' ZRC = OK
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_SETVAR() AS STRING
'---------- Get a SET variable
LOCAL varname, varvalue AS STRING
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseString varname                               ' Then get the variable name
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Get_SETVar$(" + varname + ") "
         varvalue = sSetTable("GET", varname)                     ' Retrieve substitution value
         IF VAL(LEFT$(varvalue, 1)) > 0 THEN                      ' Not found?
            FUNCTION = ""                                         ' Return ""
            TP.ErrFlag = 8: TP.ErrMSg = "Unknown SET name"        ' Set RC
         ELSE                                                     '
           FUNCTION = MID$(varvalue, 2)                           ' Return looked up value
           QErrorNull                                             ' ZRC = OK
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION Get_SRCL1() AS EXT
'---------- Get the line command start line number
LOCAL lno AS EXT
   gMacroTHeader = "Get_Src1_LPtr() "
   lno = IIF(gMacRange, IIF(gMacSFrom = 0, 0, gMacSFrom), 0)      ' If a range waiting, return the FromLin
   IF lno <> 0 THEN TP.UpdLControl(INT(lno))                      ' Reset the LLCtl area
   FUNCTION = lno                                                 ' Return the result
   QErrorNull                                                     ' ZRC = OK
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_SRCL2() AS EXT
'---------- Get the line command end line number
LOCAL lno AS EXT
   gMacroTHeader = "Get_Src2_LPtr() "
   lno = IIF(gMacRange, IIF(gMacSTo = 0, 0, gMacSTo), 0)          ' If a range waiting, return the ToLin
   IF lno <> 0 THEN TP.UpdLControl(INT(lno))                      ' Reset the LLCtl area
   FUNCTION = lno                                                 ' Return the result
   QErrorNull                                                     ' ZRC = OK
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_SRCLCMD() AS STRING
'---------- Get the last line command
   gMacroTHeader = "Get_SRC_LCmd$() "
   FUNCTION = IIF$(gMacRange, gMacSCmd, ""): QErrorNull           ' If a range waiting, return the LCmd
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_SRCOP() AS EXT
'---------- Get the line command repeat value
LOCAL rpt AS EXT
   gMacroTHeader = "Get_SRC_Op() "
   rpt = IIF(gMacRange, IIF(gMacSRpt = 0, 0, gMacSRpt), 0)        ' If a range waiting, return the FromLin
   FUNCTION = rpt                                                 ' Return the result
   QErrorNull                                                     ' ZRC = OK
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_SRCLMOD$() AS STRING
'---------- Get the line command Post Operand
LOCAL op AS STRING
   gMacroTHeader = "Get_SRC_LMOD$() "
   op = "  "                                                      ' Set default
   IF gMacRange THEN                                              ' If a range
      IF BIT(gMacSFlag, %lCmdX) THEN MID$(op, 1, 1) = "-"         ' Set +/- as appropriate
      IF BIT(gMacSFlag, %lCmdNX) THEN MID$(op, 1, 1) = "+"        '
      IF BIT(gMacSFlag, %lCmdR) THEN MID$(op, 2, 1) = "R"         '
   END IF                                                         '
   FUNCTION = op                                                  '
   QErrorNull                                                     ' ZRC = OK
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_SYSVAR() AS STRING
'---------- Get a System Env. variable
LOCAL varname, varvalue AS STRING
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseString varname                               ' Then get the variable name
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Get_ENVVar$(" + varname + ") "
         FUNCTION = ENVIRON$(varname)                             ' Fetch environ string
         QErrorNull                                               ' ZRC = OK
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION Get_Tag() AS STRING
'---------- Get tag for a single line number
LOCAL lText AS STRING, lno AS EXT, i AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Get_Tag$(" + FORMAT$(lno) + ") "
         i = INT(lno)                                             ' Make long
         IF i > 0 AND i <= TP.LastLine THEN                       ' Within range?
            IF ISFALSE TP.LFlagData(i) THEN                       ' Not a data line?
               FUNCTION = "": TP.ErrFlag = 8: TP.ErrMsg = "Invalid line pointer" ' Else pass back null
            ELSE                                                  '
               FUNCTION = TRIM$(TP.LTagGet(i)): QErrorNull        ' Yes, fetch the tag
            END IF                                                '
         ELSE                                                     '
            FUNCTION = "": TP.ErrFlag = 8: TP.ErrMsg = "Invalid line pointer" ' Else pass back null
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION Get_TOPSCREEN() AS EXT
'---------- Get Line pointer of Top of Screen
   gMacroTHeader = "Get_TopScrn_LPtr() "
   FUNCTION  = TP.TopScrn: QErrorNull                             ' Return TopScrn
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION Get_UniqueID() AS EXT
'---------- Get Unique Tab ID
   gMacroTHeader = "Get_Unique_ID() "
   FUNCTION  = TP.WindowID: QErrorNull                            ' Return the Unique ID
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION IS_BOTTOM() AS EXT
'---------- Get the line's Bottom status
LOCAL lno AS EXT, i AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Is_Bottom(" + FORMAT$(lno) + ") "
         i = INT(lno)                                             ' Make long
         IF i > 0 AND i <= TP.LastLine THEN                       ' Within range?
            FUNCTION = ISTRUE TP.LFlagBottom(i): QErrorNull       ' Return Bottom status
         ELSE                                                     '
            FUNCTION = %False: TP.ErrFlag = 8: TP.ErrMsg = "Invalid line pointer"
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION Get_XLines() AS EXT
'---------- Get # Excluded lines in an X block
LOCAL lText AS STRING, lno AS EXT, i AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Get_XLines(" + FORMAT$(lno) + ") "
         i = INT(lno)                                             ' Make long
         IF i > 0 AND i <= TP.LastLine THEN                       ' Within range?
            IF ISTRUE TP.LFlagXClude(i) THEN                      ' Xclude line?
               FUNCTION = TP.LWrk1Get(i): QErrorNull              ' Yes, pass back the number
            ELSE                                                  '
               FUNCTION = 0: TP.ErrFlag = 8: TP.ErrMsg = "LPtr is not an Exclude Marker line"
            END IF                                                '
         ELSE                                                     '
            FUNCTION = 0: TP.ErrFlag = 8: TP.ErrMsg = "Invalid line pointer"
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION Get_XStatus() AS STRING
'---------- Get Exclude status for a single line number
LOCAL lText AS STRING, lno AS EXT, i AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Get_XStatus$(" + FORMAT$(lno) + ") "
         i = INT(lno)                                             ' Make long
         IF i > 0 AND i <= TP.LastLine THEN                       ' Within range?
            IF ISTRUE TP.LFlagInvisible(i) THEN                   ' Invisible line?
               FUNCTION = "X": QErrorNull                         ' Else pass back X
            ELSE                                                  '
               FUNCTION = "NX": QErrorNull                       ' Yes, pass back NX
            END IF                                                '
         ELSE                                                     '
            FUNCTION = "": TP.ErrFlag = 8: TP.ErrMsg = "Invalid line pointer"
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION HALT() AS EXT
'---------- Set ErrFlag and ErrMsg and terminate
LOCAL lText, tText AS STRING, level AS EXT, lRet AS LONG
   IF thinBasic_CheckOpenParens(%True, %True) THEN                ' An Open (
      IF thinBasic_CheckCloseParens(%True, %True) THEN            ' And if closed directly by a ) properly
         gMacroTHeader = "Halt() "                                ' Set trace
         DoTrace                                                  ' Dump it
         '---Unload all modules, release all memory
         CALL DWORD gMacRelease USING thinBasic_Release(0&) TO lRet ' Will be used in future release

         '---Free loaded thinCore library
         FreeLibrary(gMacCore)                                    '
         TerminateThread BYVAL gMacThread, 900                    ' Second value is thread exit code

      ELSE                                                        ' Not an immediate )
         IF thinBasic_DetermineType = %thinBasic_ReturnNumber THEN' Is next param a number?
            thinBasic_ParseNumber level                           ' Then get the error level
            IF thinBasic_CheckComma(%True, %True) THEN            ' A comma?
               thinBasic_ParseString lText                        ' Then get the string
               DO WHILE thinBasic_CheckComma(%True, %True)        ' A Comma?
                  thinBasic_ParseString tText                     ' Then get the next string
                  lText += " " + tText                            ' Add it to the string
               LOOP                                               '
            END IF                                                '
         ELSE                                                     '
            level = 0                                             ' Default level to zero
            thinBasic_ParseString lText                           ' Then get the string
            DO WHILE thinBasic_CheckComma(%True, %True)           ' A Comma?
               thinBasic_ParseString tText                        ' Then get the next string
               lText += " " + tText                               ' Add it to the string
            LOOP                                                  '
         END IF                                                   '
         IF thinBasic_CheckCloseParens THEN                       ' And if closed by ) properly
            gMacroTHeader = "Halt(" + FORMAT$(level) + ", " + lText + ") "
            IF level = 0 THEN                                     '
               gMacroRC = 0: gMacroMsg = lText                    '
            ELSE                                                  '
               gMacroRC = INT(level): gMacroMsg = "RC=" + FORMAT$(level) + ": " + lText
            END IF                                                '
            DoTrace                                               ' Do trace if needed
            '---Unload all modules, release all memory
            CALL DWORD gMacRelease USING thinBasic_Release(0&) TO lRet ' Will be used in future release

            '---Free loaded thinCore library
            FreeLibrary(gMacCore)                                 '
            TerminateThread BYVAL gMacThread, INT(level) + 900    ' Second value is thread exit code
         END IF                                                   '
      END IF                                                      '

   ELSE                                                           ' No Open bracket at all
      gMacroTHeader = "Halt "                                     ' Set trace
      DoTrace                                                     ' Dump it
      '---Unload all modules, release all memory
       CALL DWORD gMacRelease USING thinBasic_Release(0&) TO lRet '
      FreeLibrary(gMacCore)                                       '
      TerminateThread BYVAL gMacThread, 900                       ' Second value is thread exit code
   END IF                                                         '
   FUNCTION = 0                                                   ' All is well.
END FUNCTION

FUNCTION IS_BNDS() AS EXT
'---------- Get the line's BNDS status
LOCAL lno AS EXT, i AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Is_BNDS(" + FORMAT$(lno) + ") "
         i = INT(lno)                                             ' Make long
         IF i > 0 AND i <= TP.LastLine THEN                       ' Within range?
            FUNCTION = ISTRUE TP.LFlagBounds(i): QErrorNull       ' Return BNDS status
         ELSE                                                     '
            FUNCTION = %False: TP.ErrFlag = 8: TP.ErrMsg = "Invalid line pointer"
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION IS_COLS() AS EXT
'---------- Get the line's COLS status
LOCAL lno AS EXT, i AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Is_Cols(" + FORMAT$(lno) + ") "
         i = INT(lno)                                             ' Make long
         IF i > 0 AND i <= TP.LastLine THEN                       ' Within range?
            FUNCTION = ISTRUE TP.LFlagCols(i): QErrorNull         ' Return COLS status
         ELSE                                                     '
            FUNCTION = %False: TP.ErrFlag = 8: TP.ErrMsg = "Invalid line pointer"
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION IS_DATA() AS EXT
'---------- Get the line's Data status
LOCAL lno AS EXT, i AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Is_Data(" + FORMAT$(lno) + ") "
         i = INT(lno)                                             ' Make long
         IF i > 0 AND i <= TP.LastLine THEN                       ' Within range?
            FUNCTION = ISTRUE TP.LFlagData(i): QErrorNull         ' Return Data status
         ELSE                                                     '
            FUNCTION = %False: TP.ErrFlag = 8: TP.ErrMsg = "Invalid line pointer"
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION IS_FILE() AS EXT
'---------- Get the line's FILE status
LOCAL lno AS EXT, i AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Is_File(" + FORMAT$(lno) + ") "
         i = INT(lno)                                             ' Make long
         IF i > 0 AND i <= TP.LastLine THEN                       ' Within range?
            FUNCTION = ISTRUE TP.LFlagFile(i): QErrorNull         ' Return FILE status
         ELSE                                                     '
            FUNCTION = %False: TP.ErrFlag = 8: TP.ErrMsg = "Invalid line pointer"
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION IS_MARK() AS EXT
'---------- Get the line's Mark status
LOCAL lno AS EXT, i AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Is_Mark(" + FORMAT$(lno) + ") "
         i = INT(lno)                                             ' Make long
         IF i > 0 AND i <= TP.LastLine THEN                       ' Within range?
            FUNCTION = ISTRUE TP.LFlagMark(i): QErrorNull         ' Return Mark status
         ELSE                                                     '
            FUNCTION = %False: TP.ErrFlag = 8: TP.ErrMsg = "Invalid line pointer"
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION IS_MASK() AS EXT
'---------- Get the line's Mask status
LOCAL lno AS EXT, i AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Is_Mask(" + FORMAT$(lno) + ") "
         i = INT(lno)                                             ' Make long
         IF i > 0 AND i <= TP.LastLine THEN                       ' Within range?
            FUNCTION = ISTRUE TP.LFlagMask(i): QErrorNull         ' Return Mask status
         ELSE                                                     '
            FUNCTION = %False: TP.ErrFlag = 8: TP.ErrMsg = "Invalid line pointer"
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION IS_NOTE() AS EXT
'---------- Get the line's Note status
LOCAL lno AS EXT, i AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Is_Note(" + FORMAT$(lno) + ") "
         i = INT(lno)                                             ' Make long
         IF i > 0 AND i <= TP.LastLine THEN                       ' Within range?
            FUNCTION = ISTRUE TP.LFlagNote(i): QErrorNull         ' Return Note status
         ELSE                                                     '
            FUNCTION = %False: TP.ErrFlag = 8: TP.ErrMsg = "Invalid line pointer"
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION IS_prof() AS EXT
'---------- Get the line's pROF status
LOCAL lno AS EXT, i AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Is_Prof(" + FORMAT$(lno) + ") "
         i = INT(lno)                                             ' Make long
         IF i > 0 AND i <= TP.LastLine THEN                       ' Within range?
            FUNCTION = ISTRUE TP.LFlagpROF(i): QErrorNull         ' Return Prof status
         ELSE                                                     '
            FUNCTION = %False: TP.ErrFlag = 8: TP.ErrMsg = "Invalid line pointer"
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION IS_TABS() AS EXT
'---------- Get the line's Tabs status
LOCAL lno AS EXT, i AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Is_Tabs(" + FORMAT$(lno) + ") "
         i = INT(lno)                                             ' Make long
         IF i > 0 AND i <= TP.LastLine THEN                       ' Within range?
            FUNCTION = ISTRUE TP.LFlagTabs(i): QErrorNull         ' Return Tabs status
         ELSE                                                     '
            FUNCTION = %False: TP.ErrFlag = 8: TP.ErrMsg = "Invalid line pointer"
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION IS_TOP() AS EXT
'---------- Get the line's Top status
LOCAL lno AS EXT, i AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Is_Top(" + FORMAT$(lno) + ") "
         i = INT(lno)                                             ' Make long
         IF i > 0 AND i <= TP.LastLine THEN                       ' Within range?
            FUNCTION = ISTRUE TP.LFlagTop(i): QErrorNull          ' Return Top status
         ELSE                                                     '
            FUNCTION = %False: TP.ErrFlag = 8: TP.ErrMsg = "Invalid line pointer"
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION IS_ULINE() AS EXT
'---------- Get the line's User status
LOCAL lno AS EXT, i AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Is_ULine(" + FORMAT$(lno) + ") "
         i = INT(lno)                                             ' Make long
         IF i > 0 AND i <= TP.LastLine THEN                       ' Within range?
            FUNCTION = ISTRUE TP.LFlagUser(i): QErrorNull         ' Return User status
         ELSE                                                     '
            FUNCTION = %False: TP.ErrFlag = 8: TP.ErrMsg = "Invalid line pointer"
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION IS_WORD() AS EXT
'---------- Get the line's WORD status
LOCAL lno AS EXT, i AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Is_Word(" + FORMAT$(lno) + ") "
         i = INT(lno)                                             ' Make long
         IF i > 0 AND i <= TP.LastLine THEN                       ' Within range?
            FUNCTION = ISTRUE TP.LFlagWord(i): QErrorNull         ' Return WORD status
         ELSE                                                     '
            FUNCTION = %False: TP.ErrFlag = 8: TP.ErrMsg = "Invalid line pointer"
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION IS_XLine() AS EXT
'---------- Get the line's Invisible status
LOCAL lno AS EXT, i AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Is_XLine(" + FORMAT$(lno) + ") "
         i = INT(lno)                                             ' Make long
         IF i > 0 AND i <= TP.LastLine THEN                       ' Within range?
            FUNCTION = ISTRUE TP.LFlagInvisible(i): QErrorNull    ' Return Excl status
         ELSE                                                     '
            FUNCTION = %False: TP.ErrFlag = 8: TP.ErrMsg = "Invalid line pointer"
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION IS_XMARKER() AS EXT
'---------- Get the line's EXCL status
LOCAL lno AS EXT, i AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Is_XMarker(" + FORMAT$(lno) + ") "
         i = INT(lno)                                             ' Make long
         IF i > 0 AND i <= TP.LastLine THEN                       ' Within range?
            FUNCTION = ISTRUE TP.LFlagXclude(i): QErrorNull       ' Return Excl status
         ELSE                                                     '
            FUNCTION = %False: TP.ErrFlag = 8: TP.ErrMsg = "Invalid line pointer"
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION IS_PrimCmd() AS EXT
'---------- Get the macro execution mode
LOCAL lcmd AS STRING, i AS LONG
   gMacroTHeader = "Is_Primary_Cmd() "
   lcmd = TRIM$(LTBLSCmd)                                         ' Format the cmd name
   IF lcmd = "" THEN                                              ' If no Lcmd
      FUNCTION = %True: QErrorNull                                ' Pass back true
   ELSE                                                           '
      i = LCmdT.GetCmdIX(lcmd)                                    ' See if source command is an internal line command
      FUNCTION = IIF(i, %True, %False): QErrorNull                ' Pass back answer
   END IF                                                         '
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION IS_LineCmd() AS EXT
'---------- Get the macro execution mode
LOCAL lcmd AS STRING, i AS LONG
   gMacroTHeader = "Is_Line_Cmd() "
   lcmd = TRIM$(LTBLSCmd)                                         ' Format the cmd name
   IF lcmd = "" THEN                                              ' If no Lcmd
      FUNCTION = %False: QErrorNull                               ' Pass back false
   ELSE                                                           '
      i = LCmdT.GetCmdIX(lcmd)                                    ' See if source command is an internal line command
      FUNCTION = IIF(i, %False, %True): QErrorNull                ' Pass back answer
   END IF                                                         '
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION NOP() AS EXT
'---------- Do nothing
   gMacroTHeader = "NOP() "
   FUNCTION = 0: QErrorNull                                       ' Pass back all OK
   DoTrace                                                        ' Do trace if needed
END FUNCTION

FUNCTION OV() AS EXT
'---------- Return -1
   FUNCTION = -1                                                  ' Pass back answer
END FUNCTION

FUNCTION REQUEST_LABEL() AS STRING
'---------- Get a label for the requested line
LOCAL lText AS STRING, lno AS EXT, i, j, k AS LONG

   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      IF thinBasic_DetermineType = %thinBasic_ReturnNumber THEN   ' Is next param a number?
         thinBasic_ParseNumber lno                                ' Then get the number
         gMacroTHeader = "Request_Label$(" + FORMAT$(lno) + ") "  ' Setup trace data
         IF LEFT$(FORMAT$(lno), 1) = "." THEN                     ' Is this an unquoted .nnn operand?
            i = TP.LineNoRef(FORMAT$(lno))                        ' Do a lookup on it
         ELSE                                                     ' Else just a number
            i = lno                                               ' Put line number in i
         END IF                                                   '
      ELSE                                                        '
         thinBasic_ParseString lText                              ' Then get the string
         gMacroTHeader = "Request_Label$(" + lText + ") "         ' Setup trace data
         IF LEFT$(lText, 1) = "." THEN                            ' A dotted line number?
            i = TP.LineNoRef(lText)                               ' Go look it up
         ELSEIF LEFT$(lText, 1) = "!" THEN                        ' A pure line reference
            i = VAL(MID$(lText, 2))                               ' Extract the line number
         ELSE                                                     ' Alphabetic, reject it
            i = 0                                                 ' Force failure
         END IF                                                   '
      END IF                                                      '
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         GOSUB ReturnLabel                                        ' Go get a label for this line
      END IF                                                      '
      DoTrace                                                     ' Dump the trace
   END IF                                                         '
   EXIT FUNCTION                                                  ' We're done

ReturnLabel:
   IF i > 0 AND i <= TP.LastLine AND ISTRUE TP.LFlagData(i) THEN  ' Within range and a Data Line?
      lText = TP.LLblGet(i)                                       ' Get the line's label
      IF ISNULL(TRIM$(lText)) THEN                                ' No Label?
         k = 1                                                    ' Start k as true
         DO WHILE K > 0                                           ' Loop till an unused temp label is found
            TP.FixCtr = TP.FixCtr + 1                             ' Bump unique counter
            lText = FORMAT$(TP.FixCtr, "0000")                    ' Make it 4 digits
            FOR j = 1 TO 4                                        ' Lets make it alphabetic
               MID$(lText, j, 1) = MID$("ABCDEFGHIJ", VAL(MID$(lText, j, 1)) + 1, 1)
            NEXT j                                                '
            lText = "._" + lText                                  ' Make it a Temp Label format
            k = TP.LineNoRef(lText)                               ' Make sure it doesn't already exist
         LOOP                                                     '
         TP.LLblSet(i, lText)                                     ' Set it into the line
         TP.UpdLControl(i)                                        ' Redo LLCtl
         gMacTempVars = %True                                     ' Set so cleanup will remove these
         FUNCTION = lText: QErrorNull                             '
      ELSE                                                        ' Got label, pass it back
         FUNCTION = TRIM$(lText): QErrorNull                      '
      END IF                                                      '
   ELSE                                                           '
      FUNCTION = "": TP.ErrMsg = "Invalid line pointer"           ' Else pass back null
      TP.ErrFlag = 8                                              ' Set ZRC
   END IF                                                         '
   RETURN                                                         ' We're done here
END FUNCTION

FUNCTION RELEASE_LABEL() AS STRING
'---------- Release a label
LOCAL lText AS STRING, lno AS EXT, i, j, k AS LONG

   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseString lText                                 ' Then get the label
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Release_Label$(" + lText + ") "         ' Setup trace data
         IF IsEQ(lText, "ALL") THEN                               ' The ALL request
            GOSUB ClearAll                                        ' Go do it
         ELSEIF LEFT$(lText, 1) <> "." OR LEN(lText) > 8 THEN     ' Look valid?
            FUNCTION = "": TP.ErrMsg = "Invalid label format"     ' Else pass back null
            TP.ErrFlag = 8                                        ' Set ZRC
         ELSE
            i = TP.LineNoRef(lText)                               ' Do a lookup on it
            IF i <> -1 THEN                                       ' It's a defined label
               IF LEFT$(lText, 2) = "._" THEN                     ' A temporary label?
                  TP.LLblSet(i, $BlankLNo)                        ' Clear the line
                  TP.UpdLControl(i)                               ' Redo LLCtl
                  FUNCTION = "": QErrorNull                       ' Else pass back null
               ELSE                                               ' A permanent label
                  FUNCTION = lText: QErrorNull                    ' Else pass back the label
               END IF                                             '
            ELSE                                                  ' Not a valid label
               FUNCTION = "": TP.ErrMsg = "Not a defined label"   ' Else pass back null and error message
               TP.ErrFlag = 8                                     ' Bad RC
            END IF                                                '
         END IF                                                   '
      END IF                                                      '
      DoTrace                                                     ' Dump the trace
   END IF                                                         '
   EXIT FUNCTION                                                  ' We're done

ClearAll:
   FOR i = 1 TO TP.LastLine                                       ' Must search all lines
      IF LEFT$(TP.LLblGet(i), 2) = "._" THEN                      ' Contain a temporary label?
         TP.LLblSet(i, $BlankLNo)                                 ' Clear the line
         TP.UpdLControl(i)                                        ' Redo LLCtl
      END IF                                                      '
   NEXT i                                                         '
   FUNCTION = "": QErrorNull                                      ' Else pass back null
   RETURN                                                         ' We're done here
END FUNCTION

FUNCTION Reset_GBLNum() AS EXT
'---------- Clear a table
LOCAL i, j, k, o, gottable AS LONG, level AS EXT, varprefix, lclprefix, varname AS STRING, wvarname AS WSTRING, vvarvalue AS VARIANT
   IF thinBasic_CheckOpenParens(%True, %True) THEN                ' An Open (
      IF thinBasic_CheckCloseParens(%True, %True) THEN            ' And if closed directly by a ) properly
         gMacroTHeader = "Reset_GBLNum() "                        '

      ELSE                                                        ' Not an immediate )
         thinBasic_ParseNumber level                              ' Then get the table number
         IF thinBasic_CheckCloseParens THEN                       ' And if closed by ) properly
            gMacroTHeader = "Reset_GBLNum(" + FORMAT$(level) + ")"
            gottable = %True                                      ' Remember we are doing just one table
         END IF                                                   '
      END IF                                                      '
   END IF                                                         '

   '----- Doing all or just one table
   IF ISFALSE gottable THEN
      gNumVar.Clear                                               ' Clear the collection
      i = OBJRESULT                                               ' look at answer
      IF i > 0 THEN                                               ' Error?
         FUNCTION = 0: TP.ErrFlag = 8: TP.ErrMsg = "Unable to clear GBLNum" ' Return code
      ELSE                                                        '
         FUNCTION = 0: QErrorNull                                 ' Pass back answer
      END IF                                                      '
      DoTrace                                                     ' Do trace if needed
   ELSE                                                           '
      '----- Clear just one table
      varprefix = FORMAT$(level) + ","                            ' Build table prefix string
      j = gNumVar.Count                                           ' Get # items in the collection
      IF j > 0 THEN                                               ' If anything in the collection
         FOR i = 1 TO j                                           ' Loop through it
            gNumVar.Entry(i, wvarname, vvarvalue)                 '
            o = OBJRESULT                                         ' Get result
            IF o = 0 THEN                                         ' If OK
               varname = wvarname                                 ' Return name
               lclprefix = LEFT$(varname, INSTR(varname, ","))    ' Extract the table number
               IF lclprefix = varprefix THEN                      ' Is this one we want?
                  gNumVar.Remove(wvarname)                        ' Delete it
                  o = OBJRESULT                                   ' OK?
                  IF o = 0 THEN                                   '
                     DECR j                                       ' Reduce limit count
                     INCR k                                       ' Count deleted
                  ELSE                                            ' Else
                     EXIT FOR                                     ' Bail out
                  END IF                                          '
               END IF                                             '
            ELSE                                                  ' Error
               EXIT FOR                                           ' Bail out
            END IF                                                '
         NEXT i                                                   '
         FUNCTION = 0: TP.ErrFlag = 8: TP.ErrMsg = "Deleted Item count: " + FORMAT$(k) ' Return code
         DoTrace
      END IF                                                      '

   END IF

END FUNCTION

FUNCTION Reset_GBLSTR() AS EXT
'---------- Clear a table
LOCAL i, j, k, o, gottable AS LONG, level AS EXT, varprefix, lclprefix, varname AS STRING, wvarname AS WSTRING, vvarvalue AS VARIANT
   IF thinBasic_CheckOpenParens(%True, %True) THEN                ' An Open (
      IF thinBasic_CheckCloseParens(%True, %True) THEN            ' And if closed directly by a ) properly
         gMacroTHeader = "Reset_GBLStr() "                        '

      ELSE                                                        ' Not an immediate )
         thinBasic_ParseNumber level                              ' Then get the table number
         IF thinBasic_CheckCloseParens THEN                       ' And if closed by ) properly
            gMacroTHeader = "Reset_GBLStr(" + FORMAT$(level) + ")"
            gottable = %True                                      ' Remember we are doing just one table
         END IF                                                   '
      END IF                                                      '
   END IF                                                         '

   '----- Doing all or just one table
   IF ISFALSE gottable THEN
      gStrVar.Clear                                               ' Clear the collection
      i = OBJRESULT                                               ' look at answer
      IF i > 0 THEN                                               ' Error?
         FUNCTION = 0: TP.ErrFlag = 8: TP.ErrMsg = "Unable to clear GBLStr" ' Return code
      ELSE                                                        '
         FUNCTION = 0: QErrorNull                                 ' Pass back answer
      END IF                                                      '
      DoTrace                                                     ' Do trace if needed
   ELSE                                                           '
      '----- Clear just one table
      varprefix = FORMAT$(level) + ","                            ' Build table prefix string
      j = gStrVar.Count                                           ' Get # items in the collection
      IF j > 0 THEN                                               ' If anything in the collection
         FOR i = 1 TO j                                           ' Loop through it
            gStrVar.Entry(i, wvarname, vvarvalue)                 '
            o = OBJRESULT                                         ' Get result
            IF o = 0 THEN                                         ' If OK
               varname = wvarname                                 ' Return name
               lclprefix = LEFT$(varname, INSTR(varname, ","))    ' Extract the table number
               IF lclprefix = varprefix THEN                      ' Is this one we want?
                  gStrVar.Remove(wvarname)                        ' Delete it
                  o = OBJRESULT                                   ' OK?
                  IF o = 0 THEN                                   '
                     DECR j                                       ' Reduce limit count
                     INCR k                                       ' Count deleted
                  ELSE                                            ' Else
                     EXIT FOR                                     ' Bail out
                  END IF                                          '
               END IF                                             '
            ELSE                                                  ' Error
               EXIT FOR                                           ' Bail out
            END IF                                                '
         NEXT i                                                   '
         FUNCTION = 0: TP.ErrFlag = 8: TP.ErrMsg = "Deleted Item count: " + FORMAT$(k) ' Return code
         DoTrace
      END IF                                                      '
   END IF
END FUNCTION

FUNCTION SET_ClrLINE() AS EXT
'---------- Set color text for a line
LOCAL lText AS STRING, lwText AS WSTRING,lno AS EXT, i, j, k AS LONG, aTxt AS WSTRING POINTER, AttrAsc AS WORD
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckComma() THEN                              ' A Comma?
         thinBasic_ParseString lText                              ' Then get the string
         lwText = UUCASE(lText)                                   ' Uppercase it
         IF thinBasic_CheckCloseParens THEN                       ' And if closed by ) properly
            gMacroTHeader = "Set_Clr_Line(" + FORMAT$(lno) + ", " + LEFT$(lText, 32) + ") "
            i = INT(lno)                                          ' Make long
            IF i > 0 AND i <= TP.LastLine THEN                    ' Within range?
               IF ISFALSE TP.LFlagData(i) THEN                    ' Ineligible line?
                  TP.ErrFlag = 8: TP.ErrMsg = "Non-Data line"     ' Set ZRC
               ELSE                                               '
                  IF VERIFY(LText,nHiLitesChrs) > 0 THEN          ' Invalid characters?
                     TP.ErrFlag = 8: TP.ErrMsg = "Invalid color request characters" ' Set ZRC
                     FUNCTION = 8                                 '
                  ELSE                                            '
                     lText = LSET$(lText, TP.LTxtLen(i))          ' Make = text length
                     lwText = LSET$(lwText, TP.LTxtLen(i) USING CHR$$(0)) ' Make = text length
                     aTxt = TP.LAttrPtr(i)                        ' Get address of existing Attr data
                     FOR j = 1 TO LEN(lwText)                     ' Update the Attribute line now
                        AttrAsc = ASC(MID$(@aTxt, j)) AND (%AttrAll - %AttrHiLite) ' Get Attr without the HiLite bits
                        k = INSTR(nHiLitesChrs, MID$(lText, j, 1)) - 1 ' Get HiLite Index
                        SHIFT LEFT k, 8                           ' Align it properly
                        AttrAsc = AttrAsc OR k                    ' Stuff it in the Attribute byte
                        MID$(@aTxt, j, 1) = CHR$$(AttrAsc)        '
                     NEXT j                                       '
                     TP.ModSet(i)                                 ' Remember we changed something
                  END IF                                          '
               END IF                                             '
            ELSE                                                  '
               TP.ErrFlag = 8: TP.ErrMsg = "Invalid line pointer" ' Set ZRC
            END IF                                                '
            DoTrace                                               ' Do trace if needed
         END IF                                                   '
      END IF                                                      '
   END IF                                                         '
   FUNCTION = TP.ErrFlag                                          ' Set return
END FUNCTION

FUNCTION SET_CSR() AS EXT
'---------- Set next cursor position
LOCAL lno, colno, lgth AS EXT, ilno, icol, ilgth AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckComma() THEN                              ' A Comma?
         thinBasic_ParseNumber colno                              ' Then get the column number
         IF thinBasic_CheckComma(%True, %True) THEN               ' A Comma?
            thinBasic_ParseNumber lgth                            ' Then get the length
            IF thinBasic_CheckCloseParens THEN                    ' And if closed by ) properly
               '                                                  '
            END IF                                                '
            gMacroTHeader = "Set_Csr(" + FORMAT$(lno) + ", " + FORMAT$(colno) + ", " + FORMAT$(lgth) + ") "
         ELSE                                                     ' No length
            lgth = 0                                              ' Default to zero
            IF thinBasic_CheckCloseParens THEN                    ' And if closed by ) properly
               '
            END IF                                                '
            gMacroTHeader = "Set_Csr(" + FORMAT$(lno) + ", " + FORMAT$(colno) + ") "
         END IF                                                   '
         ilno = INT(lno): icol = INT(colno): ilgth = INT(lgth)    ' Make long
         IF ilno > 0 AND ilno < TP.LastLine THEN                  ' Within range?
            TP.CurSetReq(%CurMacro, ilno, icol, %True)            '
            QErrorNull                                            '
         ELSE                                                     '
            TP.Errflag = 8: TP.ErrMsg = "Invalid line pointer"    ' Bad line pointer
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
   FUNCTION = TP.ErrFlag                                          ' Set return
END FUNCTION

FUNCTION SET_EResult() AS EXT
'---------- Set the replacement CHANGE string
LOCAL lText AS STRING, lno AS EXT, i AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseString lText                                 ' Then get the string
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Set_E_Result(" + LEFT$(lText, 32) + ") "'
         gMacCString = lText                                      ' Yes, set the text
         QErrorNull                                               ' Set ZRC
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
   FUNCTION = TP.ErrFlag                                          ' Set return
END FUNCTION

FUNCTION Set_GBLNUM() AS EXT
'---------- Set a GBL NUM variable
LOCAL varprefix, varname AS STRING, i AS LONG, wvarname AS WSTRING, level, varvalue AS EXT, wvarvalue AS VARIANT
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      IF thinBasic_DetermineType = %thinBasic_ReturnNumber THEN   ' Is next param a number?
         thinBasic_ParseNumber level                              ' Then get the table number
         varprefix = FORMAT$(level) + ","                         ' Format it
         IF thinBasic_CheckComma() THEN                           ' A comma?
            '                                                     ' Better be there
         END IF                                                   '
      ELSE                                                        ' No table prefix - assign zero
         varprefix = "0,"                                         '
      END IF
      thinBasic_ParseString varname                               ' Then get the variable name
      varname = varprefix + varname                               ' Add the table prefix
      IF thinBasic_CheckComma() THEN                              ' A Comma?
         thinBasic_ParseNumber varvalue                           ' Then get the string
         IF thinBasic_CheckCloseParens THEN                       ' And if closed by ) properly
            gMacroTHeader = "Set_Gbl_Num(" + varname + ", " + FORMAT$(varvalue) + ") "
            QErrorNull                                            ' Default answers
            IF ISNULL(varname) THEN                               ' Null?  Ignore it
               FUNCTION = 0                                       ' Return OK
               EXIT FUNCTION                                      '
            END IF                                                '
            wvarname = varname: wvarvalue = varvalue              ' Convert to WSTRING / Variant, always uppercase name
            gNumvar.Remove(wvarname)                              ' Delete it in case it's there
            gNumVar.Add(wvarname, wvarvalue)                      ' Try the Add
            i = OBJRESULT                                         ' look at answer
            IF i > 0 THEN                                         ' Error?
               FUNCTION = 8: TP.ErrFlag = 8: TP.ErrMsg = "Unable to ADD variable" ' Return code
            ELSE                                                  '
               FUNCTION = 0                                       ' Return OK
            END IF                                                '
            DoTrace                                               ' Do trace if needed
         END IF                                                   '
      END IF                                                      '
   END IF
END FUNCTION

FUNCTION Set_GBLSTR() AS EXT
'---------- Set a GBL STR variable
LOCAL varprefix, varname, varvalue AS STRING, i AS LONG, level AS EXT, wvarname, wvarvalue AS WSTRING
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      IF thinBasic_DetermineType = %thinBasic_ReturnNumber THEN   ' Is next param a number?
         thinBasic_ParseNumber level                              ' Then get the table number
         varprefix = FORMAT$(level) + ","                         ' Format it
         IF thinBasic_CheckComma() THEN                           ' A comma?
            '                                                     ' Better be there
         END IF                                                   '
      ELSE                                                        ' No table prefix - assign zero
         varprefix = "0,"                                         '
      END IF
      thinBasic_ParseString varname                               ' Then get the variable name
      varname = varprefix + varname                               ' Prefix it
      IF thinBasic_CheckComma() THEN                              ' A Comma?
         thinBasic_ParseString varvalue                           ' Then get the string
         IF thinBasic_CheckCloseParens THEN                       ' And if closed by ) properly
            gMacroTHeader = "Set_Gbl_Str(" + varname + ", " + varvalue + ") "
            QErrorNull                                            ' Default answers
            IF ISNULL(varname) THEN                               ' Null?  Ignore it
               FUNCTION = 0                                       ' Return OK
               EXIT FUNCTION                                      '
            END IF                                                '
            wvarname = varname: wvarvalue = varvalue              ' Convert to WSTRINGs
            gStrvar.Remove(wvarname)                              ' Delete it in case it's there
            gStrVar.Add(wvarname, wvarvalue)                      ' Try the Add
            i = OBJRESULT                                         ' look at answer
            IF i > 0 THEN                                         ' Error?
               FUNCTION = 8: TP.ErrFlag = 8: TP.ErrMsg = "Unable to ADD variable" ' Return code
            ELSE                                                  '
               FUNCTION = 0                                       ' Return OK
            END IF                                                '
            DoTrace                                               ' Do trace if needed
         END IF                                                   '
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION SET_LINE() AS EXT
'---------- Set text for a line
LOCAL lText AS STRING, lno AS EXT, i AS LONG, aTxt AS WSTRING POINTER
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckComma() THEN                              ' A Comma?
         thinBasic_ParseString lText                              ' Then get the string
         IF thinBasic_CheckCloseParens THEN                       ' And if closed by ) properly
            gMacroTHeader = "Set_Line(" + FORMAT$(lno) + ", " + LEFT$(lText, 32) + ") "
            i = INT(lno)                                          ' Make long
            IF i = 0 THEN                                         ' Request for Cmd Line?
              TP.PCommand = lText                                 ' Stuff it in
              TP.PCommandRaw = lText                              '
              QErrorNull                                          ' Set ZRC
            ELSEIF i > 0 AND i <= TP.LastLine THEN                ' Within range?
               IF ISTRUE (TP.LFlagGet(i) AND %NonTypable) THEN    ' Ineligible line?
                  TP.ErrFlag = 8: TP.ErrMsg = "Non-modifyable line type" ' Set ZRC
               ELSE                                               '
                  TP.LTxtSet(i, lText)                            ' Yes, set the text
                  TP.ModSet(i)                                    ' Remember we changed something
                  TP.ErrFlag = 0: TP.ErrMSg = ""                  ' Set ZRC
                  IF TP.LFlagWord(i) THEN TP.WordSave             '
                  IF TP.LFlagMark(i) THEN TP.MarkSave             '
                  IF TP.LFlagMask(i) THEN TP.MaskSave             '
                  IF TP.LFlagBounds(i) THEN TP.BndsSave           '
                  IF TP.LFlagTabs(i) THEN TP.TabsSave             '
                  aTxt = TP.LAttrPtr(i)                           ' Get it
                  @aTxt = LSET$(@aTxt, TP.LTxtLen(i) USING CHR$$(0)) ' Make = text length
               END IF                                             '
            ELSE                                                  '
               TP.ErrFlag = 8: TP.ErrMsg = "Invalid line pointer" ' Set ZRC
            END IF                                                '
            DoTrace                                               ' Do trace if needed
         END IF                                                   '
      END IF                                                      '
   END IF                                                         '
   FUNCTION = TP.ErrFlag                                          ' Set return
END FUNCTION

FUNCTION SET_MSG() AS EXT
'---------- Set ErrFlag and ErrMsg
LOCAL lText, tText AS STRING, level AS EXT
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      IF thinBasic_DetermineType = %thinBasic_ReturnNumber THEN   ' Is next param a number?
         thinBasic_ParseNumber level                              ' Then get the error level
         IF thinBasic_CheckComma() THEN                           ' Better be a comma
            thinBasic_ParseString lText                           ' Then get the string
            DO WHILE thinBasic_CheckComma(%True, %True)           ' A Comma?
               thinBasic_ParseString tText                        ' Then get the next string
               lText += " " + tText                               ' Add it to the string
            LOOP                                                  '
         END IF                                                   '
      ELSE                                                        '
         level = 0                                                ' Default level to zero
         thinBasic_ParseString lText                              ' Then get the string
         DO WHILE thinBasic_CheckComma(%True, %True)              ' A Comma?
            thinBasic_ParseString tText                           ' Then get the next string
            lText += " " + tText                                  ' Add it to the string
         LOOP                                                     '
      END IF                                                      '
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Set_Msg(" + FORMAT$(level) + ", " + lText + ") "
         IF level = 0 THEN                                                     '
            gMacroRC = INT(level): gMacroMsg = lText              '
         ELSE                                                     '
            gMacroRC = INT(level): gMacroMsg = "RC=" + FORMAT$(level) + ": " + lText
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
   FUNCTION = 0                                                   ' All is well.
END FUNCTION

FUNCTION Set_SETVAR() AS EXT
'---------- Set a SET variable
LOCAL varname, varvalue, SetRet AS STRING, ctab, i AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseString varname                               ' Then get the variable name
      IF thinBasic_CheckComma() THEN                              ' A Comma?
         thinBasic_ParseString varvalue                           ' Then get the string
         IF thinBasic_CheckCloseParens THEN                       ' And if closed by ) properly
            gMacroTHeader = "Set_SETVar(" + varname + ", " + varvalue + ") "
            ctab = TP.PgNumber                                    ' Save where we are
            FOR i = 1 TO TabsNum                                  ' Do for each tab
               TP = Tabs(i)                                       ' Pick the Tab
               IF IsSetEdit THEN                                  ' We have a SetEdit tab active
                  TP = Tabs(ctab)                                 ' Replace TP
                  FUNCTION = 8: TP.ErrFlag = 8: TP.ErrMsg = "A SET-EDIT session is active" ' Set fail RC
                  EXIT FUNCTION                                   '
               END IF                                             '
            NEXT i                                                '
            TP = Tabs(ctab)                                       ' Replace TP
            SetRet = sSetTable("SET", varname + " " + varvalue)   ' Do the variable SET
            IF VAL(LEFT$(SetRet, 1)) > 0 THEN                     ' Error?
               FUNCTION = VAL(LEFT$(SetRet, 1))                   ' Return code""
               TP.ErrFlag = VAL(LEFT$(SetRet, 1))                 '
               TP.ErrMsg = MID$(SetRet, 2)                        ' And error text
            ELSE                                                  '
               FUNCTION = 0: QErrorNull                           ' Return OK
            END IF                                                '
            DoTrace                                               ' Do trace if needed
         END IF                                                   '
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION Set_TOPSCREEN() AS EXT
'---------- Set Line pointer of Top of Screen
LOCAL lno, colno, lgth AS EXT, ilno, icol, ilgth AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "Set_TopScrn_Lptr(" + FORMAT$(lno) + ") "
         ilno = INT(lno)                                          ' Make long
         IF ilno > 0 AND ilno < TP.LastLine THEN                  ' Within range?
            TP.TopScrn = ilno                                     '
            QErrorNull                                            '
         ELSE                                                     '
            TP.Errflag = 8: TP.ErrMsg = "Invalid line pointer"    ' Bad line pointer
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
   FUNCTION = TP.ErrFlag                                          ' Set return
END FUNCTION

FUNCTION SPF_CMD() AS EXT
'---------- Issue an SPFLite Primary command
LOCAL lText, errMsg AS STRING
LOCAL errRC AS LONG
   lText = sGetStrConcat                                          ' Go get the operands
   IF lText <> "" THEN                                            ' Null "" means errors
      gMacroTHeader = "SPF_Cmd(" + ltext + ") "                   ' Build trace header
      TP.pCommand = lText                                         ' Set string in pCommand
      TP.SetSubst()                                               ' Do SET substitution
      QErrorNull                                                  ' Start off clean
      TP.lCtlProcess                                              ' Get any line commands ready
      TP.CmdAssign                                                ' Try to run it
      errRC = TP.ErrFlag: errMsg = TP.ErrMsg                      ' Copy RC and message
      TP.DispScreen                                               ' Let cursor positioing take place
      TP.RangeSet("RESET", %False, %False)                        ' Ensure default range
      TP.ErrFlag = errRC: TP.ErrMsg = errMsg                      ' Restore so trace has correct values
      DoTrace                                                     ' Do trace if needed
      FUNCTION = errRC                                            ' Return error flag
   END IF                                                         '
END FUNCTION

FUNCTION SPF_Debug() AS EXT
'---------- Display a string
LOCAL lText AS STRING
   lText = sGetStrConcat                                          ' Go get the operands
   IF lText <> "" THEN                                            ' Null "" means errors
      gMacroTHeader = "SPF_Debug(" + ltext + ") "                 '
      QErrorNull                                                  '
      DEBUG lText                                                 ' Display it
      DoTrace                                                     ' Do trace if needed
   END IF                                                         '
END FUNCTION

FUNCTION SPF_EXEC() AS EXT
'---------- Issue a DOS execute command
LOCAL lText, tText AS STRING, i, j AS LONG, sync, normal AS EXT
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      IF thinBasic_DetermineType = %thinBasic_ReturnNumber THEN   ' Is next param a number?
         thinBasic_ParseNumber sync                               ' Then get the SYNC/ASYNC flag
         IF thinBasic_CheckComma() THEN                           ' Better be a comma
            IF thinBasic_DetermineType = %thinBasic_ReturnNumber THEN   ' Is next also a number?
               thinBasic_ParseNumber normal                       ' Then get the SHOW/HIDE  flag
               IF thinBasic_CheckComma() THEN                     ' Better be a another comma
                  thinBasic_ParseString lText                     ' Then get the string
                  DO WHILE thinBasic_CheckComma(%True, %True)     ' A Comma?
                     thinBasic_ParseString tText                  ' Then get the next string
                     lText += " " + tText                         ' Add it to the string
                  LOOP                                            '
               END IF                                             '
            ELSE                                                  ' Not a SHOW/HIDE flag
               normal = 1                                         ' Set to NORMAL
               thinBasic_ParseString lText                        ' Then get the string
               DO WHILE thinBasic_CheckComma(%True, %True)        ' A Comma?
                  thinBasic_ParseString tText                     ' Then get the next string
                  lText += " " + tText                            ' Add it to the string
               LOOP                                               '
            END IF                                                '
         END IF                                                   '
      ELSE                                                        '
         sync = 0                                                 ' Default SYNC to zero  (SYNC)
         normal = 1                                               ' Default SHOW/HIDE to 1 (NORMAL)
         thinBasic_ParseString lText                              ' Then get the string
         DO WHILE thinBasic_CheckComma(%True, %True)              ' A Comma?
            thinBasic_ParseString tText                           ' Then get the next string
            lText += " " + tText                                  ' Add it to the string
         LOOP                                                     '
      END IF                                                      '
      j = normal                                                  ' Convert SHOW/HIDE to integer
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "SPF_Exec(" + IIF$(sync = 0, "SYNC", "ASYNC") + ", " + IIF$(j = 1, "NORMAL", "HIDDEN") + ", " + lText + ") "
         IF sync = 0 THEN                                         ' Normal SYNC?
            SHELL lText, j, EXIT TO i                             '
            TP.ErrFlag = i: TP.ErrMsg = IIF$(i =0, "", ERROR$): FUNCTION = i ' Set RC
         ELSE                                                     '
            i = SHELL(lText, j)                                   '
            QErrorNull: FUNCTION = i                              ' Set RC
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
   FUNCTION = 0                                                   ' All is well.
END FUNCTION

FUNCTION SPF_Exempt_File() AS EXT
'---------- See if a file extension is Exempt from text processing
LOCAL lText, XList AS STRING, i AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseString lText                                 ' Then get the string
      IF LEFT$(lText, 1) = "." THEN lText = MID$(lText, 2)        ' Strip off any leading period.
      lText = "," + uucase(lText) + ","                           ' Upper case it and surround bu commas
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "SPF_Exempt_File(" + ltext + ") "        ' Build trace
         XList = "," + UUCASE(ENV.FMFFList) + ","                 ' Build XList with leading/trailing commas
         XList = REMOVE$(XList, " ")                              ' Eliminate any spaces
         IF INSTR(XList, lText) THEN                              ' No there, then not exempt
            FUNCTION = %True: TP.ErrFlag = 8: TP.ErrMsg = "File IS in Exempt list"  ' Return
         ELSE                                                     '
            FUNCTION = %False: TP.ErrFlag = 0: TP.ErrMsg = "File is NOT in Exempt list"  ' Return
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
   FUNCTION = TP.ErrFlag                                          ' All is well.
END FUNCTION

FUNCTION SPF_INS() AS EXT
'---------- Insert text into a line
LOCAL lText, ltext2 AS STRING, lno, lclCol AS EXT, i AS LONG, aTxt AS WSTRING POINTER
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckComma() THEN                              ' A Comma?
         thinBasic_ParseNumber lclCol                             ' Then get the column number
         IF thinBasic_CheckComma() THEN                           ' A Comma?
            thinBasic_ParseString ltext                           ' Then get the text
            IF thinBasic_CheckCloseParens THEN                    ' And if closed by ) properly
               gMacroTHeader = "SPF_INS(" + FORMAT$(lno) + ", " + FORMAT$(lclCol) + ", " + ltext + ") "
               QErrorNull                                         ' Start off as no error
               i = INT(lno)                                       ' Make long
               IF i > 0 AND i <= TP.LastLine THEN                 ' Within range?
                  IF ISTRUE (TP.LFlagGet(i) AND %NonTypable) THEN ' Ineligible line?
                     TP.ErrFlag = 8: TP.ErrMsg = "Non-modifyable line type"  ' Set ZRC
                  ELSE                                            '
                     ltext2 = TP.LTxtGet(i)                       ' Get the original text
                     IF LEN(ltext2) < lclCol - 1 THEN ltext2 = LSET$(ltext2, lclCol - 1) ' Make minimum needed length
                     ltext2 = LEFT$(ltext2, lclCol - 1) + ltext + MID$(ltext2, lclCol)
                     TP.LTxtSet(i, lText2)                        ' Yes, set back the text
                     TP.ModSet(i)                                 ' Remember we changed something
                     IF TP.LFlagWord(i) THEN TP.WordSave          '
                     IF TP.LFlagMark(i) THEN TP.MarkSave          '
                     IF TP.LFlagMask(i) THEN TP.MaskSave          '
                     IF TP.LFlagBounds(i) THEN TP.BndsSave        '
                     IF TP.LFlagTabs(i) THEN TP.TabsSave          '
                     aTxt = TP.LAttrPtr(i)                        ' Get it
                     @aTxt = LSET$(@aTxt, TP.LTxtLen(i) USING CHR$$(0)) ' Make = text length
                  END IF                                          '
               ELSE                                               '
                  TP.ErrFlag = 8: TP.ErrMsg = "Invalid line pointer"  ' Set ZRC to fail
               END IF                                             '
               DoTrace                                            ' Do trace if needed
            END IF                                                '
         END IF                                                   '
      END IF                                                      '
   END IF                                                         '
   FUNCTION = TP.ErrFlag                                          ' All is well.
END FUNCTION

FUNCTION SPF_LOOPCHECK() AS EXT
'---------- Tell SPFLite no loop checking
LOCAL OnOff, nvalue AS EXT, i AS LONG
   OnOff = IIF(gMacNoLoopTest, %False, %True)                     ' Save current state (inverted)
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber nvalue                                ' Then get the new value
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "SPF_Loop_Check(" + FORMAT$(nvalue) + ") "
         i = IIF(nvalue, %False, %True)                           ' Make long and invert
         gMacNoLoopTest = i                                       ' Set in the new value
         FUNCTION = OnOff: QErrorNull                             ' Return previous setting
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION SPF_OVR() AS EXT
'---------- Overlay text in a line
LOCAL lText, ltext2 AS STRING, lno, lclCol AS EXT, i, j AS LONG, aTxt AS WSTRING POINTER
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckComma() THEN                              ' A Comma?
         thinBasic_ParseNumber lclCol                             ' Then get the column number
         IF thinBasic_CheckComma() THEN                           ' A Comma?
            thinBasic_ParseString ltext                           ' Then get the text
            IF thinBasic_CheckCloseParens THEN                    ' And if closed by ) properly
               gMacroTHeader = "SPF_OVR(" + FORMAT$(lno) + ", " + FORMAT$(lclCol) + ", " + ltext + ") "
               QErrorNull                                         ' Start off as no error
               i = INT(lno)                                       ' Make long
               IF i > 0 AND i <= TP.LastLine THEN                 ' Within range?
                  IF ISTRUE (TP.LFlagGet(i) AND %NonTypable) THEN ' Ineligible line?
                     TP.ErrFlag = 8: TP.ErrMsg = "Non-modifyable line type" ' Set ZRC
                  ELSE                                            '
                     ltext2 = TP.LTxtGet(i)                       ' Get the original text
                     IF LEN(ltext2) < lclCol + LEN(ltext) - 1 THEN ltext2 = LSET$(ltext2, lclCol + LEN(ltext) - 1) ' Make minimum needed length
                     FOR j = lclCol TO lclCol + LEN(ltext) - 1    ' Loop doing the overlay
                        IF MID$(ltext2, j, 1) = " " THEN _        ' If blank in original text
                           MID$(ltext2, j, 1) = MID$(ltext, j - lclCol + 1, 1) ' Copy character from overlay text
                     NEXT j                                       '
                     TP.LTxtSet(i, lText2)                        ' Yes, set back the text
                     TP.ModSet(i)                                 ' Remember we changed something
                     IF TP.LFlagWord(i) THEN TP.WordSave          '
                     IF TP.LFlagMark(i) THEN TP.MarkSave          '
                     IF TP.LFlagMask(i) THEN TP.MaskSave          '
                     IF TP.LFlagBounds(i) THEN TP.BndsSave        '
                     IF TP.LFlagTabs(i) THEN TP.TabsSave          '
                     aTxt = TP.LAttrPtr(i)                        ' Get it
                     @aTxt = LSET$(@aTxt, TP.LTxtLen(i) USING CHR$$(0)) ' Make = text length
                  END IF                                          '
               ELSE                                               '
                  TP.ErrFlag = 8: TP.ErrMsg = "Invalid line pointer" ' Set ZRC to fail
               END IF                                             '
               DoTrace                                            ' Do trace if needed
            END IF                                                '
         END IF                                                   '
      END IF                                                      '
   END IF                                                         '
   FUNCTION = TP.ErrFlag                                          ' All is well.
END FUNCTION

FUNCTION SPF_OVR_REP() AS EXT
'---------- Force Overlay text in a line
LOCAL lText, ltext2 AS STRING, lno, lclCol AS EXT, i, j AS LONG, aTxt AS WSTRING POINTER
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckComma() THEN                              ' A Comma?
         thinBasic_ParseNumber lclCol                             ' Then get the column number
         IF thinBasic_CheckComma() THEN                           ' A Comma?
            thinBasic_ParseString ltext                           ' Then get the text
            IF thinBasic_CheckCloseParens THEN                    ' And if closed by ) properly
               gMacroTHeader = "SPF_OVR_REP(" + FORMAT$(lno) + ", " + FORMAT$(lclCol) + ", " + ltext + ") "
               QErrorNull                                         ' Start off as no error
               i = INT(lno)                                       ' Make long
               IF i > 0 AND i <= TP.LastLine THEN                 ' Within range?
                  IF ISTRUE (TP.LFlagGet(i) AND %NonTypable) THEN ' Ineligible line?
                     TP.ErrFlag = 8: TP.ErrMsg = "Non-modifyable line type" ' Set ZRC
                  ELSE                                            '
                     ltext2 = TP.LTxtGet(i)                       ' Get the original text
                     IF LEN(ltext2) < lclCol + LEN(ltext) - 1 THEN ltext2 = LSET$(ltext2, lclCol + LEN(ltext) - 1) ' Make minimum needed length
                     FOR j = lclCol TO lclCol + LEN(ltext) - 1    ' Loop doing the overlay
                        IF MID$(ltext, j - lclCol + 1, 1) <> " " THEN _ ' If non-blank in overlay
                           MID$(ltext2, j, 1) = MID$(ltext, j - lclCol + 1, 1) ' Copy character from overlay text
                     NEXT j                                       '
                     TP.LTxtSet(i, lText2)                        ' Yes, set back the text
                     TP.ModSet(i)                                 ' Remember we changed something
                     IF TP.LFlagWord(i) THEN TP.WordSave          '
                     IF TP.LFlagMark(i) THEN TP.MarkSave          '
                     IF TP.LFlagMask(i) THEN TP.MaskSave          '
                     IF TP.LFlagBounds(i) THEN TP.BndsSave        '
                     IF TP.LFlagTabs(i) THEN TP.TabsSave          '
                     aTxt = TP.LAttrPtr(i)                        ' Get it
                     @aTxt = LSET$(@aTxt, TP.LTxtLen(i) USING CHR$$(0)) ' Make = text length
                  END IF                                          '
               ELSE                                               '
                  TP.ErrFlag = 8: TP.ErrMSg = "Invalid line pointer" ' Set ZRC to fail
               END IF                                             '
               DoTrace                                            ' Do trace if needed
            END IF                                                '
         END IF                                                   '
      END IF                                                      '
   END IF                                                         '
   FUNCTION = TP.ErrFlag                                          ' All is well.
END FUNCTION

FUNCTION SPF_PARSE() AS EXT
'---------- Parse the Get_Arg$ array
LOCAL num AS EXT
LOCAL maxlit, maxnum, maxlptr, maxtag, tv AS LONG
LOCAL KWCtl, KWCtl2, KWGroup, KWEntry, OrigKWCtl, KWMastName, t, aliasname, KWUse AS STRING
LOCAL wvarname, wvarvalue AS WSTRING
LOCAL vvarvalue AS VARIANT
LOCAL i, j, k, parsedone, brkted AS LONG
   gParseTbl.Clear                                                ' Clear the Parse table
   RESET mgotlit, mgotnum, mgotlptr, mgottag

   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber num                                   ' Then get the maxlit value
      maxlit = num                                                '
      IF thinBasic_CheckComma() THEN                              ' A Comma?
         thinBasic_ParseNumber num                                ' Then get the maxnum value
         maxnum = num                                             '
         IF thinBasic_CheckComma() THEN                           ' A Comma?
            thinBasic_ParseNumber num                             ' Then get the maxlptr value
            maxlptr = num                                         '
            IF thinBasic_CheckComma() THEN                        ' A Comma?
               thinBasic_ParseNumber num                          ' Then get the maxtag value
               maxtag = num                                       '
               gMacroTHeader = "SPF_Parse(" + FORMAT$(maxlit) + ", " + FORMAT$(maxnum) + ", " + FORMAT$(maxlptr) + ", " + FORMAT$(maxtag)
               DO WHILE thinBasic_CheckComma(%True, %True)        ' A Comma?
                  thinBasic_ParseString KWCtl                     ' Then get the KWCtl string
                  gMacroTHeader += ", " + KWCtl                   ' Add to trace string
                  GOSUB DoKWCtl                                   ' Go process it
               LOOP                                               '
               IF thinBasic_CheckCloseParens THEN                 ' And if closed by ) properly
                  gMacroTHeader += ")"                            ' Add to trace string
                  parsedone = %True                               ' Remember we've reached the )
               END IF                                             '
            END IF                                                '
         END IF                                                   '
      END IF                                                      '
   END IF                                                         '

   '----- Handle old style -ve operands, convert to new flag oriented style
   IF maxlit  < 0 THEN maxlit  = ABS(maxlit)  + %ARG_VAR          ' Do each one
   IF (maxlit AND (%ARG_VAR + %ARG_OPT)) = (%ARG_VAR + %ARG_OPT) THEN TP.ErrMsg = "Cannot use both ARG_VAR and ARG_OPT": GOSUB ErrorExit
   IF maxnum  < 0 THEN maxnum  = ABS(maxnum)  + %ARG_VAR          '
   IF (maxnum AND (%ARG_VAR + %ARG_OPT)) = (%ARG_VAR + %ARG_OPT) THEN TP.ErrMsg = "Cannot use both ARG_VAR and ARG_OPT": GOSUB ErrorExit
   IF maxlptr < 0 THEN maxlptr = ABS(maxlptr) + %ARG_VAR          '
   IF (maxlptr AND (%ARG_VAR + %ARG_OPT)) = (%ARG_VAR + %ARG_OPT) THEN TP.ErrMsg = "Cannot use both ARG_VAR and ARG_OPT": GOSUB ErrorExit
   IF maxtag  < 0 THEN maxtag  = ABS(maxtag)  + %ARG_VAR          '
   IF (maxtag AND (%ARG_VAR + %ARG_OPT)) = (%ARG_VAR + %ARG_OPT) THEN TP.ErrMsg = "Cannot use both ARG_VAR and ARG_OPT": GOSUB ErrorExit

   '----- Parse parameters have been done, now process the Arg$ array
   FOR j = 1 TO 50                                                ' OK, now process the arguments
      IF ISNOTNULL(mOprands(j)) THEN                              ' Only do non-null operands

         '--- First see if it's a Keyword
         wvarname = UUCASE(mOprands(j))                           ' Convert to WSTRING, always uppercase the name
         KWUse = UUCASE(mOprands(j))                              ' Save as name to use
         vvarvalue = gParseTbl.Item(wvarname)                     ' Try the retrieval
         k = OBJRESULT                                            ' look at answer
         IF k = 0 THEN                                            ' Found, validate it
            KWMastName = VARIANT$(vvarvalue)                      ' Make normal string
            IF LEFT$(KWMastName, 1) = "9" THEN                    ' An Alias entry?
               wvarname = UUCASE(MID$(KWMastName, 2))             ' Swap names
               KWUse = UUCASE(MID$(KWMastName, 2))                ' Save this as the name to use
               vvarvalue = gParseTbl.Item(wvarname)               ' Try the retrieval
               k = OBJRESULT                                      ' look at answer
               IF k > 0 THEN TP.ErrMsg = "ParseTbl Structure error": GOSUB ErrorExit ' Bail out, shold never happen
               KWMastName = VARIANT$(vvarvalue)                   ' Make normal string
            END IF                                                '
            KWMastName = MID$(KWMastName, 2)                      ' Extract the master name
            wvarname = KWMastName                                 ' Lets see if MasterName has a set value
            vvarvalue = gParseTbl.Item(wvarname)                  ' Try the retrieval
            k = OBJRESULT                                         ' look at answer
            IF k > 0 THEN TP.ErrMsg = "ParseTbl Structure error": GOSUB ErrorExit ' Bail out, shold never happen
            t = VARIANT$(vvarvalue)                               ' Make a string
            IF ISNULL(t) THEN                                     ' If still null, goodness, no collision
               wvarvalue = UUCASE(mOprands(j) + "|" + KWUse)      ' Set it to the original Operand followed by main name
               gParseTbl.Replace(wvarname, wvarvalue)             '
               k = OBJRESULT                                      ' look at answer
               IF k > 0 THEN TP.ErrMsg = "ParseTbl REPLACE error": GOSUB ErrorExit ' Bail out
               wvarvalue = "1" + UUCASE(KWMastName)               ' Set the KW Entry to True
               wvarname = UUCASE(KWUse)                           '
               gParseTbl.Replace(wvarname, wvarvalue)             '
               k = OBJRESULT                                      ' look at answer
               IF k > 0 THEN TP.ErrMsg = "ParseTbl REPLACE error": GOSUB ErrorExit ' Bail out
            ELSE                                                  '
               IF IsEQ(mOprands(j), PARSE$(t, "|", 1)) THEN       ' Simple duplicate
                  TP.ErrMsg = "Keyword: " + UUCASE(mOprands(j)) + " has been entered twice"
               ELSE                                               '
                  TP.ErrMsg = "Keyword: " + UUCASE(mOprands(j)) + " cannot be entered along with: " + UUCASE(PARSE$(t, "|", 1))
               END IF                                             '
               GOSUB ErrorExit                                    ' Bail out
            END IF                                                '

         ELSE                                                     ' Not a KW
            IF LEFT$(mOprands(j), 1) = "." THEN                   ' Is this a dotted label?
               IF mgotlptr = (maxlptr AND &H000000FF) THEN TP.ErrMsg = mOprands(j) + " exceeds the maximum allowed Line References": GOSUB ErrorExit ' At the maximum
               GOSUB vLRef                                        ' Go Validate it
               INCR mgotlptr                                      ' Count it
               wvarvalue = mOprands(j)                            ' Make it Wide
               wvarname = "." + FORMAT$(mgotlptr)                 ' Build the key
               gParseTbl.Add(wvarname, wvarvalue)                 ' Try to Add the .n entry
               k = OBJRESULT                                      ' look at answer
               IF k > 0 THEN TP.ErrMsg = "ParseTbl ADD error": GOSUB ErrorExit ' Bail out

            ELSEIF LEFT$(mOprands(j), 1) = ":" THEN               ' Is this a Tag??
               IF mgottag = (maxtag AND &H000000FF) THEN TP.ErrMsg = mOprands(j) + " exceeds the maximum allowed Tag References": GOSUB ErrorExit ' At the maximum
               GOSUB vTag                                         ' Go Validate it
               INCR mgottag                                       ' Count it
               wvarvalue = mOprands(j)                            ' Make it Wide
               wvarname = ":" + FORMAT$(mgottag)                  ' Build the key
               gParseTbl.Add(wvarname, wvarvalue)                 ' Try to Add the .n entry
               k = OBJRESULT                                      ' look at answer
               IF k > 0 THEN TP.ErrMsg = "ParseTbl ADD error": GOSUB ErrorExit ' Bail out

            ELSEIF VERIFY(mOprands(j), "0123456789") = 0 THEN     ' Is this a numeric literal
               IF mgotnum = (maxnum AND &H000000FF) THEN TP.ErrMsg = mOprands(j) + " exceeds the maximum allowed numeric literals": GOSUB ErrorExit ' At the maximum
               INCR mgotnum                                       ' Count it
               wvarvalue = mOprands(j)                            ' Make it Wide
               wvarname = "#" + FORMAT$(mgotnum)                  ' Build the key
               gParseTbl.Add(wvarname, wvarvalue)                 ' Try to Add the .n entry
               k = OBJRESULT                                      ' look at answer
               IF k > 0 THEN TP.ErrMsg = "ParseTbl ADD error": GOSUB ErrorExit ' Bail out

            ELSE                                                  ' This is a string
               IF mgotlit = (maxlit AND &H000000FF) THEN TP.ErrMsg = mOprands(j) + " exceeds the maximum allowed text literals": GOSUB ErrorExit ' At the maximum
               INCR mgotlit                                       ' Count it
               wvarvalue = mOprands(j)                            ' Make it Wide
               wvarname = "$" + FORMAT$(mgotlit)                  ' Build the key
               gParseTbl.Add(wvarname, wvarvalue)                 ' Try to Add the .n entry
               k = OBJRESULT                                      ' look at answer
               IF k > 0 THEN TP.ErrMsg = "ParseTbl ADD error": GOSUB ErrorExit ' Bail out

            END IF                                                '
         END IF                                                   '
      END IF                                                      '
   NEXT j                                                         '

   '----- Test for overall limit requirements
   tv = maxlit AND &H000000FF                                     ' Get value without flag bits
   IF tv > 0 THEN                                                 ' Something is wanted
      IF (maxlit AND %ARG_VAR) THEN                               ' If VAR specified, anything is fine
         '                                                        '
      ELSEIF (maxlit AND %ARG_OPT) THEN                           ' If OPT specified, must be equal or zero
         IF mgotlit <> tv AND mgotlit <> 0 THEN                   '
            TP.ErrMsg = "Command requires 0 or " + FORMAT$(tv) + " TEXT literals, " + FORMAT$(mgotlit) + IIF$(mgotlit = 1, " was ", " were ") + "entered": GOSUB ErrorExit
         END IF                                                   '
      ELSE                                                        ' Else must be equal
         IF mgotlit <> tv THEN                                    '
             TP.ErrMsg = "Command requires " + FORMAT$(tv) + " TEXT literals, " + FORMAT$(mgotlit) + IIF$(mgotlit = 1, " was ", " were ") + "entered": GOSUB ErrorExit
          END IF                                                  '
       END IF                                                     '
   END IF                                                         '

   tv = maxnum AND &H000000FF                                     ' Get value without flag bits
   IF tv > 0 THEN                                                 ' Something is wanted
      IF (maxnum AND %ARG_VAR) THEN                               ' If VAR specified, anything is fine
         '                                                        '
      ELSEIF (maxnum AND %ARG_OPT) THEN                           ' If OPT specified, must be equal or zero
         IF mgotnum <> tv AND mgotnum <> 0 THEN                   '
            TP.ErrMsg = "Command requires 0 or " + FORMAT$(tv) + " Numeric literals, " +  FORMAT$(mgotnum) + IIF$(mgotnum = 1, " was ", " were ") + "entered": GOSUB ErrorExit
         END IF                                                   '
      ELSE                                                        '
         IF mgotnum <> tv THEN                                    '
            TP.ErrMsg = "Command requires " + FORMAT$(tv) + " Numeric literals, " +  FORMAT$(mgotnum) + IIF$(mgotnum = 1, " was ", " were ") + "entered": GOSUB ErrorExit
         END IF                                                   '
      END IF                                                      '
   END IF                                                         '

   tv = maxtag AND &H000000FF                                     ' Get value without flag bits
   IF tv > 0 THEN                                                 ' Something is wanted
      IF (maxtag AND %ARG_VAR) THEN                               ' If VAR specified, anything is fine
         '                                                        '
      ELSEIF (maxtag AND %ARG_OPT) THEN                           ' If OPT specified, must be equal or zero
         IF mgottag <> tv AND mgottag <> 0 THEN                   '
            TP.ErrMsg = "Command requires 0 or " + FORMAT$(tv) + " Tag operandss, " +  FORMAT$(mgottag) + IIF$(mgottag = 1, " was ", " were ") + "entered": GOSUB ErrorExit
         END IF                                                   '
      ELSE                                                        '
         IF mgottag <> tv THEN                                    '
            TP.ErrMsg = "Command requires " + FORMAT$(tv) + " Tag operandss, " +  FORMAT$(mgottag) + IIF$(mgottag = 1, " was ", " were ") + "entered": GOSUB ErrorExit
         END IF                                                   '
      END IF                                                      '
   END IF                                                         '

   tv = maxlptr AND &H000000FF                                    ' Get value without flag bits
   IF tv > 0 THEN                                                 ' Something is wanted
      IF (maxlptr AND %ARG_VAR) THEN                              ' If VAR specified, anything is fine
         '                                                        '
      ELSEIF (maxlptr AND %ARG_OPT) THEN                          ' If OPT specified, must be equal or zero
         IF mgotlptr <> tv AND mgotlptr <> 0 THEN                 '
            TP.ErrMsg = "Command requires 0 or " + FORMAT$(tv) + " Line References, " +  FORMAT$(mgotlptr) + IIF$(mgotlptr = 1, " was ", " were ") + "entered": GOSUB ErrorExit
         END IF                                                   '
      ELSE                                                        '
         IF mgotlptr <> tv THEN                                   '
            TP.ErrMsg = "Command requires " + FORMAT$(tv) + " Line References, " +  FORMAT$(mgotlptr) + IIF$(mgotlptr = 1, " was ", " were ") + "entered": GOSUB ErrorExit
         END IF                                                   '
      END IF                                                      '
   END IF                                                         '

   FUNCTION = 0: QErrorNull                                       ' Set return
   DoTrace                                                        ' Do Trace if needed
   EXIT FUNCTION

'----- Process one KWCtl string
DoKWCtl:
   origKWCtl = KWCtl: KWGroup = ""                                ' Save original, reset the groupname

   '----- Process the Groupname if present
   i = INSTR(KWCtl, ":")                                          ' Is a groupname present?
   IF i THEN                                                      '
      IF i = 1 THEN                                               ' Oops?
         TP.ErrMsg = "Invalid/Missing Groupname for a Keyword list" ' Set error text
         GOSUB ErrorExit                                          ' Bail out
      END IF
      KWGroup = LEFT$(KWCtl, i - 1)                               ' Extract it
      KWCtl = MID$(KWCtl, i + 1)                                  ' Strip from KWCtl
      IF ISNULL(KWCtl) THEN
         TP.ErrMsg = "No keywords in the keyword list - " + origKWCtl
         GOSUB ErrorExit                                          ' Bail out
      END IF                                                      '
      GOSUB AddGroup                                              ' Add the parse Group entry
   END IF                                                         '

   '----- Process the keywords now
   KWCtl += ","                                                   ' Force trailing ,
   DO WHILE KWCtl <> ""                                           ' Do while still some data left
      i = INSTR(KWCtl, ",")                                       ' Get comma location
      j = INSTR(KWCtl, "(")                                       ' Get any ( location
      IF j <> 0 AND j < i THEN                                    ' Bracketed operand?
         i = INSTR(KWCtl, ")")                                    ' Better have a trailing )
         IF i = 0 THEN                                            ' Oops
            TP.ErrMsg = "Unmatched parenthesis in keyword entry - " + origKWCtl
            GOSUB ErrorExit                                       ' Bail out
         END IF                                                   '
         i = INSTR(i, KWCtl, ",")                                 ' position at , following )
      END IF                                                      ' i now marks , following next operand
      KWEntry = LEFT$(KWCtl, i - 1)                               ' Extract one entry
      KWCtl = MID$(KWCtl, i + 1)                                  ' Shorten KWCtl
      IF ISNOTNULL(KWEntry) THEN                                  ' If we got something
         GOSUB DoKWEntry                                          ' Go process it
      END IF                                                      '
   LOOP                                                           '
   RETURN                                                         '

'----- Add Group entry
AddGroup:
   wvarname = UUCASE("$$" + KWGroup)                              ' Convert to WSTRING, always uppercase the name
   vvarvalue = gParseTbl.Item(wvarname)                           ' Try the retrieval
   k = OBJRESULT                                                  ' look at answer
   IF k = 0 THEN                                                  ' Groupname already in use?
      TP.ErrMsg = "KW Definition Groupname (" + KWGroup + ") has been used more than once" ' Oops, shouldn't already exist
      GOSUB ErrorExit                                             ' Kill it
   END IF                                                         '
   wvarvalue = ""                                                 ' Initialize as null
   gParseTbl.Add(wvarname, wvarvalue)                             ' Try to Add the $$Grpname entry
   k = OBJRESULT                                                  ' look at answer
   IF k > 0 THEN TP.ErrMsg = "ParseTbl ADD error": GOSUB ErrorExit' Bail out
   RETURN

'----- Process one Keyword entry
DoKWEntry:
   brkted = %False                                                ' Reset flag
   IF LEFT$(KWEntry, 1) = "(" THEN                                ' Bracketed type?
      brkted = %true                                              ' Remember brackets
      KWEntry = MID$(KWEntry, 2, LEN(KWEntry) - 2)                ' Remove them
   END IF                                                         '
   IF ISNULL(KWGroup) THEN                                        ' Any Group entry yet?
      k = INSTR(KWEntry, ",")                                     ' Any comma?
      KWGroup = IIF$(k, LEFT$(KWEntry, k - 1), KWEntry)           ' Use 1st KW as the Groupname
      GOSUB AddGroup                                              ' Add the parse Group entry
   END IF                                                         '
   IF ISFALSE(brkted) THEN                                        ' Normal, non-aliased KW?
      wvarname = UUCASE(KWEntry)                                  ' Convert to WSTRING, always uppercase the name
      vvarvalue = gParseTbl.Item(wvarname)                        ' Try the retrieval
      k = OBJRESULT                                               ' look at answer
      IF k = 0 THEN                                               ' Groupname already in use?
         TP.ErrMsg = "Keyword Definition (" + KWEntry + ") has been used more than once" ' Oops, shouldn't already exist
         GOSUB ErrorExit                                          ' Kill it
      END IF                                                      '
      wvarvalue = "0$$" + KWGroup                                 ' Initialize as 0 + the $$grpname
      gParseTbl.Add(wvarname, wvarvalue)                          ' Try to Add the keyword entry
      k = OBJRESULT                                               ' look at answer
      IF k > 0 THEN TP.ErrMsg = "ParseTbl ADD error": GOSUB ErrorExit ' Bail out
   ELSE                                                           ' Bracketed
      KWEntry += ","                                              ' Ensure trailing comma
      DO WHILE KWEntry <> ""                                      ' Do each operand
         k = INSTR(KWEntry, ",")                                  ' Find next comma
         wvarname = UUCASE(LEFT$(KWEntry, k - 1))                 ' Convert to WSTRING, always uppercase the name
         KWEntry = MID$(KWEntry, k + 1)                           ' Reduce string
         vvarvalue = gParseTbl.Item(wvarname)                     ' Try the retrieval
         k = OBJRESULT                                            ' look at answer
         IF k = 0 THEN                                            ' Groupname already in use?
            TP.ErrMsg = "Keyword Definition (" + wvarname + ") has been used more than once" ' Oops, shouldn't already exist
            GOSUB ErrorExit                                       ' Kill it
         END IF                                                   '
         IF brkted THEN                                           ' Use brkted as a one time switch
            brkted = %False                                       '
            wvarvalue = "0$$" + KWGroup                           ' Build an entry
            gParseTbl.Add(wvarname, wvarvalue)                    ' Try to Add the keyword entry
            aliasname = wvarname                                  ' Save the alias name
            k = OBJRESULT                                         ' look at answer
            IF k > 0 THEN TP.ErrMsg = "ParseTbl ADD error": GOSUB ErrorExit ' Bail out
         ELSE                                                     ' Non-first entry
            wvarvalue = "9" + UUCASE(aliasname)                   ' Build an alias entry
            gParseTbl.Add(wvarname, wvarvalue)                    ' Try to Add the keyword entry
            k = OBJRESULT                                         ' look at answer
            IF k > 0 THEN TP.ErrMsg = "ParseTbl ADD error": GOSUB ErrorExit ' Bail out
         END IF
      LOOP                                                        '
   END IF
   RETURN

'----- Validate a Line Reference
vLRef:
   k = VERIFY(2, mOprands(j), "<>=\" + CHR$(172))                 ' Find 1st Label/Tagname character
   IF k = 2 THEN                                                  ' No condition
      t = mOprands(j)                                             ' Take whole operand
   ELSE                                                           '
      t = "." + MID$(mOprands(j), k)                              ' Else make one without the condition
   END IF                                                         '
   IF (maxlptr AND %ARG_DEF) <> 0 THEN                            ' If ARG_DEF, then it must exist
       IF TP.LineNoRef(t) = -1 THEN                               ' If invalid
          TP.ErrMsg = "Line Reference: " + mOprands(j) + " is invalid"
          GOSUB ErrorExit                                         '
       END IF                                                     '
   END IF                                                         '
   RETURN

vTag:
   k = VERIFY(2, mOprands(j), "<>=\" + CHR$(172))                 ' Find 1st Label/Tagname character
   IF k = 2 THEN                                                  ' No condition
      t = mOprands(j)                                             ' Take whole operand
   ELSE                                                           '
      t = "." + MID$(mOprands(j), k)                              ' Else make one without the condition
   END IF                                                         '
   IF (maxtag AND %ARG_DEF) <> 0 THEN                             ' If ARG_DEF, make sure it's valid
       IF sTagVal(t, %True) THEN                                  ' If invalid
          TP.ErrMsg = "Tag Operand: " + mOprands(j) + " is invalid"   '
          GOSUB ErrorExit                                         '
       END IF                                                     '
   END IF                                                         '
   RETURN

'----- Kill the function and return
ErrorExit:
   TP.ErrFlag = 8: FUNCTION = 8                                   ' Set error level
   IF ISFALSE parsedone THEN                                      ' Have we reached the ) ?
      DO                                                          ' Spin to closing )
         thinBasic_GetToken                                       '
         IF thinBasic_CheckCloseParens(%True, %False) THEN EXIT DO'
      LOOP                                                        '
   END IF                                                         '
   DoTrace                                                        ' Do Trace if needed
   EXIT FUNCTION                                                  '
   RETURN                                                         '

END FUNCTION

FUNCTION SPF_QUOTE() AS STRING
'---------- Wrap quotes around a string
LOCAL lText, lText2 AS STRING, q1, q2, q3 AS LONG
   IF thinBasic_CheckOpenParens() THEN                            ' A Open (
      thinBasic_ParseString lText                                 ' Then get the string
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "SPF_Quote$(" + ltext + ") "

         '----- See if existing string is quoted
         IF INSTR($DQ + "'`", LEFT$(lText, 1)) <> 0 AND _         ' If a leading " ' OR `
            LEFT$(lText, 1) = RIGHT$(lText, 1)  AND _             ' and trailing = leading
            INSTR(MID$(lText, 2, LEN(lText) - 2), LEFT$(lText, 1)) = 0 THEN
            FUNCTION = lText: TP.ErrFlag = 0: TP.ErrMsg = ""      ' Return string unchanged
         ELSE                                                     '
            q1 = TALLY(lText, $DQ)                                ' Count types of quotes
            q2 = TALLY(lText, "'")                                '
            q3 = TALLY(lText, "`")                                '
            IF q1 AND q2 AND q3 THEN                              ' No available quote?
               FUNCTION = lText: TP.ErrFlag = 8: TP.ErrMsg = "String uses all three quotes" ' Return string unchanged
            ELSE                                                  '
               IF q1 = 0 THEN                                     ' " available?
                  lText2 = WRAP$(lText, $DQ, $DQ)                 ' Use it
               ELSEIF q2 = 0 THEN                                 ' ' available?
                  lText2 = WRAP$(lText, "'", "'")                 ' Use it
               ELSE                                               ' ` must be available
                  lText2 = WRAP$(lText, "'", "'")                 ' Use it
               END IF                                             '
               FUNCTION = lText2: QErrorNull                      ' Return the quoted operand
            END IF                                                '
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION SPF_REP() AS EXT
'---------- Replace text in a line
LOCAL lText, ltext2 AS STRING, lno, lclCol AS EXT, i AS LONG, aTxt AS WSTRING POINTER
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      thinBasic_ParseNumber lno                                   ' Then get the line number
      IF thinBasic_CheckComma() THEN                              ' A Comma?
         thinBasic_ParseNumber lclCol                             ' Then get the column number
         IF thinBasic_CheckComma() THEN                           ' A Comma?
            thinBasic_ParseString ltext                           ' Then get the text
            IF thinBasic_CheckCloseParens THEN                    ' And if closed by ) properly
               gMacroTHeader = "SPF_REP(" + FORMAT$(lno) + ", " + FORMAT$(lclCol) + ", " + ltext + ") "
               QErrorNull                                         ' Start off as no error
               i = INT(lno)                                       ' Make long
               IF i > 0 AND i <= TP.LastLine THEN                 ' Within range?
                  IF ISTRUE (TP.LFlagGet(i) AND %NonTypable) THEN ' Ineligible line?
                     TP.ErrFlag = 8: TP.ErrMsg = "Non-modifyable line type" ' Set ZRC
                  ELSE                                            '
                     ltext2 = TP.LTxtGet(i)                       ' Get the original text
                     IF LEN(ltext2) < lclCol + LEN(ltext) - 1 THEN ltext2 = LSET$(ltext2, lclCol + LEN(ltext) - 1) ' Make minimum needed length
                     MID$(ltext2, lclCol, LEN(ltext)) = ltext     '
                     TP.LTxtSet(i, lText2)                        ' Yes, set back the text
                     TP.ModSet(i)                                 ' Remember we changed something
                     IF TP.LFlagWord(i) THEN TP.WordSave          '
                     IF TP.LFlagMark(i) THEN TP.MarkSave          '
                     IF TP.LFlagMask(i) THEN TP.MaskSave          '
                     IF TP.LFlagBounds(i) THEN TP.BndsSave        '
                     IF TP.LFlagTabs(i) THEN TP.TabsSave          '
                     aTxt = TP.LAttrPtr(i)                        ' Get it
                     @aTxt = LSET$(@aTxt, TP.LTxtLen(i) USING CHR$$(0)) ' Make = text length
                  END IF                                          '
               ELSE                                               '
                  TP.ErrFlag = 8: TP.ErrMSg = "Invalid line pointer" ' Set ZRC to fail
               END IF                                             '
               DoTrace                                            ' Do trace if needed
            END IF                                                '
         END IF                                                   '
      END IF                                                      '
   END IF                                                         '
   FUNCTION = TP.ErrFlag                                          ' All is well.
END FUNCTION

FUNCTION SPF_SHELL() AS EXT
'---------- Issue a DOS command
LOCAL lText, tText AS STRING, i, j AS LONG, sync, normal AS EXT
   IF thinBasic_CheckOpenParens() THEN                            ' An Open (
      IF thinBasic_DetermineType = %thinBasic_ReturnNumber THEN   ' Is next param a number?
         thinBasic_ParseNumber sync                               ' Then get the SYNC/ASYNC flag
         IF thinBasic_CheckComma() THEN                           ' Better be a comma
            IF thinBasic_DetermineType = %thinBasic_ReturnNumber THEN   ' Is next also a number?
               thinBasic_ParseNumber normal                       ' Then get the SHOW/HIDE  flag
               IF thinBasic_CheckComma() THEN                     ' Better be a another comma
                  thinBasic_ParseString lText                     ' Then get the string
                  DO WHILE thinBasic_CheckComma(%True, %True)     ' A Comma?
                     thinBasic_ParseString tText                  ' Then get the next string
                     lText += " " + tText                         ' Add it to the string
                  LOOP                                            '
               END IF                                             '
            ELSE                                                  ' Not a SHOW/HIDE flag
               normal = 1                                         ' Set to NORMAL
               thinBasic_ParseString lText                        ' Then get the string
               DO WHILE thinBasic_CheckComma(%True, %True)        ' A Comma?
                  thinBasic_ParseString tText                     ' Then get the next string
                  lText += " " + tText                            ' Add it to the string
               LOOP                                               '
            END IF                                                '
         END IF                                                   '
      ELSE                                                        '
         sync = 0                                                 ' Default SYNC to zero  (SYNC)
         normal = 1                                               ' Default SHOW/HIDE to 1 (NORMAL)
         thinBasic_ParseString lText                              ' Then get the string
         DO WHILE thinBasic_CheckComma(%True, %True)              ' A Comma?
            thinBasic_ParseString tText                           ' Then get the next string
            lText += " " + tText                                  ' Add it to the string
         LOOP                                                     '
      END IF                                                      '
      j = normal                                                  ' Convert SHOW/HIDE to integer
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "SPF_Shell(" + IIF$(sync = 0, "SYNC", "ASYNC") + ", " + IIF$(j = 1, "NORMAL", "HIDDEN") + ", " + lText + ") "
         IF sync = 0 THEN                                         ' Normal SYNC?
            SHELL ENVIRON$("COMSPEC") + " /D /K" + lText, j, EXIT TO i
            TP.ErrFlag = i: TP.ErrMsg = "": FUNCTION = i          ' Set RC
         ELSE                                                     '
            i = SHELL(ENVIRON$("COMSPEC") + " /D /K" + lText, j)  '
            QErrorNull: FUNCTION = i                              ' Set RC
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION SPF_Trace() AS EXT
'---------- Set Trace mode
LOCAL lno AS EXT
   IF thinBasic_CheckOpenParens() THEN                            ' A Open (
      thinBasic_ParseNumber lno                                   ' Then get the flag value
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "SPF_Trace$(" + FORMAT$(lno) + ") "
         IF lno <> 0 AND lno <> 1 AND lno <> 2 THEN               ' Validate it
            FUNCTION = 8: TP.ErrFlag = 8: TP.ErrMSg = "Invalid Trace value"     ' Set ZRC to fail
         ELSE                                                     '
            gMacroTrace = lno                                     ' Save the value
            FUNCTION = 0: QErrorNull                              '
         END IF                                                   '
      END IF                                                      '
      DoTrace                                                     ' Do trace if needed
   END IF                                                         '
END FUNCTION

FUNCTION SPF_UNQUOTE() AS STRING
'---------- UnWrap quotes from a string
LOCAL lText AS STRING
   IF thinBasic_CheckOpenParens() THEN                            ' A Open (
      thinBasic_ParseString lText                                 ' Then get the string
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         gMacroTHeader = "SPF_UnQuote$(" + ltext + ") "

         '----- See if existing string is quoted
         QErrorNull                                               ' Set RC = 0
         IF INSTR($DQ + "'`", LEFT$(lText, 1)) <> 0 AND _         ' If a leading " ' OR `
            LEFT$(lText, 1) = RIGHT$(lText, 1) THEN               ' and trailing = leading
            FUNCTION = MID$(lText, 2, LEN(LText) - 2)             ' Return string stripped
         ELSE                                                     '
            FUNCTION = lText                                      ' Return the quoted operand
         END IF                                                   '
         DoTrace                                                  ' Do trace if needed
      END IF                                                      '
   END IF                                                         '
END FUNCTION

FUNCTION sGetStrConcat() AS STRING
'---------- Get and concatenate the remaining operands as a string
LOCAL lText, tText AS STRING
   FUNCTION = ""                                                  ' Default a failing answer
   IF thinBasic_CheckOpenParens() THEN                            ' A Open (
      thinBasic_ParseString lText                                 ' Then get the string
      DO WHILE thinBasic_CheckComma(%True, %True)                 ' A Comma?
         thinBasic_ParseString tText                              ' Then get the next string
         lText += " " + tText                                     ' Add it to the string
      LOOP                                                        '
      IF thinBasic_CheckCloseParens THEN                          ' And if closed by ) properly
         FUNCTION = lText                                         ' pass back the result
      END IF                                                      '
   END IF
END FUNCTION
