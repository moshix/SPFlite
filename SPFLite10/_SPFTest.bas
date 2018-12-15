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

#COMPILE EXE "SPFTest.EXE"
#DIM ALL
#DEBUG DISPLAY ON
#DEBUG ERROR ON
#TOOLS OFF
#OPTIMIZE CODE ON
#INCLUDE ONCE "Win32Api.inc"                                      ' Windows standard stuff
#INCLUDE ONCE "CommCtrl.inc"                                      ' Common Controls
#INCLUDE ONCE "_PCRE.inc"                                         ' RegEx stuff
#INCLUDE ONCE "_Types.inc"                                        ' Standard Types etc.
#INCLUDE ONCE "_ASMDATA.inc"                                      ' ASM data tables

' Fake structures to make mapping.inc compatible with main SPFLite source

TYPE TP_block
   PrfPCase                   AS STRING * 1
   cfFLine                    AS LONG                             ' Current found line
   cfFCol                     AS LONG                             ' Current found column
'  cfChange    AS STRING                                          ' Current Change string
   cfCLen      AS LONG                                            ' Current real length of change string
'  cfFind      AS STRING                                          ' Current Find string
   cfFlag      AS QUAD                                            ' Current Operand flags
   cfFLen      AS LONG                                            ' Current real length of find string
   mapstr_sequence_num AS LONG
END TYPE

' Dialogue equates
ENUM A1 SINGULAR
   Dlg_Window                 = 1000
   Dlg_Icon
   Dlg_H1
   Dlg_H2
   Dlg_RegEx_Case_Str
   Dlg_RegEx_Case_Str_Text
   Dlg_RegEx_Test_Str
   Dlg_RegEx_Test_Str_Text
   Dlg_RegEx_Str
   Dlg_RegEx_Str_Text
   Dlg_Map_Case_Str
   Dlg_Map_Case_Str_Text
   Dlg_Map_Source_Str
   Dlg_Map_Source_Str_Text
   Dlg_Map_Map_Str
   Dlg_Map_Map_Str_Text
   Dlg_Calc_S_Variable
   Dlg_Calc_S_Variable_Text
   Dlg_Calc_S_Variable_Text2
   Dlg_Calc_S_Incr
   Dlg_Calc_RX_Variable
   Dlg_Calc_RX_Variable_Text
   Dlg_Calc_Calc_Str
   Dlg_Calc_Calc_Str_Text
   Dlg_Result_Str
   Dlg_Result_Str_Text
   Dlg_Error_Str
   Dlg_Error_Str_Text
   Dlg_Test_Button
   Dlg_Tab
END ENUM

'----- Resource Stuff
#RESOURCE ICON, A,        "Resource File\SPFLite8.ICO"

'---------- Dialog I/O areas
GLOBAL RegEx_Case_Str         AS STRING
GLOBAL RegEx_Test_Str         AS STRING
GLOBAL RegEx_Str              AS STRING

GLOBAL Map_Case_Str           AS STRING
GLOBAL Map_Source_Str         AS STRING
GLOBAL Map_Map_Str            AS STRING

GLOBAL Calc_S_Variable        AS STRING
GLOBAL Calc_RX_Variable       AS STRING
GLOBAL Calc_Calc_Str          AS STRING
GLOBAL result_Str             AS STRING

GLOBAL err_Str                AS STRING
GLOBAL retcode                AS LONG
GLOBAL pageno                 AS LONG

'---------- Dialog Handles
GLOBAL hWnd                   AS DWORD
GLOBAL hTab                   AS DWORD
GLOBAL hRegEx_Tab             AS DWORD
GLOBAL hMap_Tab               AS DWORD
GLOBAL hCalc_Tab              AS DWORD

GLOBAL hFixedFont             AS DWORD
GLOBAL hTabFont               AS DWORD
GLOBAL hHeadFont              AS DWORD
GLOBAL hToolTips              AS DWORD

'---------- RegEx areas
GLOBAL hPCRE                  AS DWORD
GLOBAL PCRE_Regex_Str2        AS STRING
GLOBAL PCRE_ErrPtr            AS ASCIIZ PTR
GLOBAL PCRE_ErrOffsetPtr      AS DWORD
GLOBAL PCRE_Options           AS LONG
GLOBAL PCRE_ExecRC            AS LONG
GLOBAL PCRE_Offsets()         AS LONG
GLOBAL PCRE_errMsg            AS STRING
GLOBAL PCRE_lperrMsg          AS STRING PTR

'---------- Calc areas
GLOBAL Calc_S                 AS QUAD
GLOBAL Calc_S_Incr            AS LONG
GLOBAL Calc_RX                AS QUAD
GLOBAL Calc_calc              AS STRING
GLOBAL Calc_Error_str         AS STRING
GLOBAL Calc_result            AS QUAD
GLOBAL Calc_line              AS LONG
GLOBAL Calc_col               AS LONG
GLOBAL gCrashList()           AS STRING                           ' Module trace
GLOBAL gCrashCtr              AS LONG                             ' Module trace Index

'---------- PCRE links
GLOBAL hLib_PCRE              AS LONG                             ' Handle of PCRE3.dll library
GLOBAL hProc_PCRE_Compile     AS LONG                             ' Handle to Compile function
GLOBAL hProc_PCRE_Exec        AS LONG                             ' Handle to Exec function
GLOBAL hProc_PCRE_Free        AS LONG                             ' Handle to Free function
GLOBAL hProc_PCRE_Free_Ptr    AS LONG                             ' Handle to Real Free function
GLOBAL lptr                   AS LONG PTR                         ' Temp

GLOBAL TP                     AS TP_block


FUNCTION PBMAIN () AS LONG

'---------- DIM the global arrays
DIM    PCRE_Offsets(12)       AS GLOBAL LONG
DIM    gCrashList(0 TO 200)   AS GLOBAL STRING

'---------- Get PCRE reqady to use

   '---Open and load PCRE3.dll library
   hLib_PCRE = LoadLibraryA( BYCOPY "PCRE3.Dll" )

   '---If all went fine
   IF hLib_PCRE THEN                                              ' PCRE exists

      '---Try to load the functions
      hProc_PCRE_Compile          = GetProcAddress(hLib_PCRE, BYCOPY "pcre_compile")
      hProc_PCRE_Exec             = GetProcAddress(hLib_PCRE, BYCOPY "pcre_exec")
      hProc_PCRE_Free             = GetProcAddress(hLib_PCRE, BYCOPY "pcre_free")
      lptr = hProc_PCRE_Free                                      ' Free returns a POINTER to the real free routine
      hProc_PCRE_Free_Ptr = @lptr                                 ' so chain to it as the REAL entry point

      '---If all went fine ...
      IF hProc_PCRE_Compile        AND _                          ' All three better be non-zero
         hProc_PCRE_Exec           AND _                          '
         hProc_PCRE_Free_Ptr       THEN                           '
         ' All is well
      ELSE
         MSGBOX "Internal PCRE functions not found"               ' Error
         EXIT FUNCTION                                            '
      END IF                                                      '

   ELSE                                                           '
      MSGBOX "PCRE DLL does not appear to be installed"           ' Say why we didn't do it
      EXIT FUNCTION                                               '
   END IF                                                         '

'---------- Init some fields with defaults
   RegEx_Case_Str = "T"                                           ' Default CASE for Regex
   Map_Case_Str = "T"                                             ' Default CASE for Mapping
   Calc_S_Variable = "0"                                          ' Default CALC value for variable 'S'
   Calc_S_Incr = 0                                                ' Default Incr S before test?
   pageno = VAL(COMMAND$)                                         ' Get page number from command line
   pageno = IIF(pageno, pageno, 1)                                ' Default to 1 if none
   Build_Dialog                                                   ' Build the Dialog
END FUNCTION

CALLBACK FUNCTION DlgCallBack
'--------------------
' Callback function used by the Dialog
'--------------------
LOCAL lclText AS STRING
LOCAL txtP AS ASCIIZ PTR

   SELECT CASE AS LONG CB.MSG                                     '

      '----- SYSCOMMAND
      CASE %WM_SYSCOMMAND                                         '
         IF CB.HNDL <> hWnd THEN EXIT FUNCTION                    '
         IF (CBWPARAM AND &HFFF0) = %SC_CLOSE THEN                ' Trap the [x] button and Alt-F4
            DIALOG END HWnd                                       '
         END IF                                                   '

      '----- A NOTIFY message coming in
      CASE %WM_NOTIFY                                             ' Notify message
         SELECT CASE AS LONG CB.NMCODE                            ' What type?
            '----- A tab is getting control
            CASE %TCN_SELCHANGE                                   ' We're getting control
               CONTROL SEND CB.HNDL, %Dlg_Tab, %TCM_GETCURSEL, 0, 0 TO pageno
               INCR pageno                                        ' Convert from relative to actual
               result_Str = ""                                    ' Clear the result fields
               err_Str = ""                                       '
               CONTROL SET TEXT hWnd, %Dlg_Result_Str, ""         '
               CONTROL SET TEXT hWnd, %Dlg_Error_Str, ""          '

         END SELECT                                               '

      '----- A user draw item request
      CASE %WM_DRAWITEM
         IF CB.WPARAM = %Dlg_Tab THEN                             '
            TabHighLight CBHNDL, CBWPARAM, CBLPARAM               '
         END IF
         FUNCTION = 1                                             '

      '----- COMMAND                                              '
      CASE %WM_COMMAND                                            '
         SELECT CASE AS LONG CB.CTL                               '

            '----- Test button pressed (or Enter)
            CASE %Dlg_Test_Button                                 ' Test button?
               IF CB.CTLMSG = %BN_CLICKED THEN                    '
                  SELECT CASE pageno                              ' Split by which page number
                     CASE 1                                       ' Regex page

                        CONTROL GET TEXT hRegex_Tab, %Dlg_RegEx_Case_Str TO Regex_Case_Str ' Get the modified data
                        Regex_Case_Str = TRIM$(UUCASE(Regex_Case_Str))
                        IF RegEx_Case_Str <> "T" AND RegEx_Case_Str <> "C" THEN  ' Valid?               '
                           CONTROL SET TEXT hWnd, %Dlg_Error_Str, "CASE is not 'C' or 'T'"
                           EXIT FUNCTION                          '
                        END IF                                    '
                        CONTROL GET TEXT hRegex_Tab, %Dlg_RegEx_Test_Str TO Regex_Test_Str' Get the modified data
                        IF TRIM$(Regex_Test_Str) = "" THEN        '
                           CONTROL SET TEXT hWnd, %Dlg_Error_Str, "Test Source string is empty"
                           EXIT FUNCTION                          '
                        END IF                                    '
                        CONTROL GET TEXT hRegex_Tab, %Dlg_RegEx_Str TO RegEx_Str  ' Get the modified data
                        IF TRIM$(RegEx_Str) = "" THEN             '
                           CONTROL SET TEXT hWnd, %Dlg_Error_Str, "RegEx expression string is empty"
                           EXIT FUNCTION                          '
                        END IF                                    '

                        '----- Actually run the test now          '
                        result_Str = ""                           ' Clear the result fields
                        err_Str = ""                              '
                        CONTROL SET TEXT hWnd, %Dlg_Result_Str, ""'
                        CONTROL SET TEXT hWnd, %Dlg_Error_Str, "" '
                        PCRE_Options = IIF(RegEx_Case_Str = "C", 0, %PCRE_CASELESS)' Set Options to match CASE

                        PCRE_Regex_Str2 = RegEx_Str + CHR$(0)     ' Make into pseudo ASCIIZ
                        PCRE_lperrMsg = STRPTR(PCRE_errMsg)       ' Setup pointer

                        '----- Call PCRE Compile
                        CALL DWORD hProc_PCRE_Compile USING pcre_compile( _
                                       STRPTR(PCRE_Regex_Str2), _ ' Regex string
                                       PCRE_Options,            _ ' Options
                                       VARPTR(PCRE_ErrPtr),     _ ' Pointer to error string
                                       VARPTR(PCRE_ErrOffsetPtr),_' Error offset
                                       &0) _                      ' Character tables
                                       TO hPCRE
                        IF hPCRE = 0 THEN                         ' OK?
                           txtp = PCRE_ErrPtr                     '
                           err_Str = "Error found at Col: " + FORMAT$(PCRE_ErrOffsetPtr + 1) + " : " +  @txtp
                           CONTROL SET TEXT hWnd, %Dlg_Error_Str, err_Str
                           EXIT FUNCTION                          '
                        END IF                                    '

                        '----- Now call for an Exec of the compiled RegEx
                        CALL DWORD hProc_PCRE_Exec USING pcre_exec( _
                                    hPCRE,                      _ ' Compile handle
                                    &0,                         _ ' extra-data
                                    STRPTR(RegEx_Test_Str),     _ ' Test-string
                                    LEN(Regex_Test_Str),        _ ' length of Test-srtring
                                    0,                          _ ' Starting position
                                    &0,                         _ ' Options
                                    VARPTR(PCRE_Offsets(0)),    _ ' PCRE_Offsets array
                                    &12) _                        ' Size of offsets array
                                    TO PCRE_ExecRC

                        IF PCRE_ExecRC < 1 THEN                   ' How'd search go?
                           err_Str = "Not found"                  ' Not found
                           CONTROL SET TEXT hWnd, %Dlg_Result_Str, "||" ' Do messages
                           CONTROL SET TEXT hWnd, %Dlg_Error_Str, err_Str
                           CALL DWORD hProc_PCRE_Free_Ptr USING pcre_free( _
                                    hPCRE)                        ' Compile handle
                           hPCRE = 0                              '
                           EXIT FUNCTION                          '

                        ELSE                                      '
                           CONTROL SET TEXT hWnd, %Dlg_Result_Str, "|" + MID$(Regex_Test_Str, PCRE_Offsets(0) + 1 TO PCRE_Offsets(1)) + "|"
                           CONTROL SET TEXT hWnd, %Dlg_Error_Str, "Found at Col: " + FORMAT$(PCRE_Offsets(0) + 1) + ", Lgth: " + _
                                                                FORMAT$(PCRE_Offsets(1) - PCRE_Offsets(0))
                           CALL DWORD hProc_PCRE_Free_Ptr USING pcre_free( _
                                    hPCRE)                        ' Compile handle
                           hPCRE = 0                              '
                        END IF                                    '

                     CASE 2                                       ' Mapping test page

                        CONTROL GET TEXT hMap_Tab, %Dlg_Map_Case_Str TO Map_Case_Str' Get the modified data
                        Map_Case_Str = TRIM$(UUCASE(Map_Case_Str))' Test fields
                        IF Map_Case_Str <> "C" AND Map_Case_Str <> "T" THEN
                           MSGBOX "Case is not 'C' or 'T'"        '
                           EXIT FUNCTION                          '
                        END IF                                    '

                        CONTROL GET TEXT hMap_Tab, %Dlg_Map_Source_Str TO Map_Source_Str' Get the modified data
                        IF TRIM$(Map_Source_Str) = "" THEN        '
                           MSGBOX "Source string is empty"        '
                           EXIT FUNCTION                          '
                        END IF                                    '

                        CONTROL GET TEXT hMap_Tab, %Dlg_Map_Map_Str TO Map_Map_Str  ' Get the modified data
                        IF TRIM$(Map_Map_Str) = "" THEN           '
                           MSGBOX "Mapping string is empty"       '
                           EXIT FUNCTION                          '
                        END IF                                    '

                        '----- Actually run the test now          '
                        result_str = ""                           ' Clear the result
                        err_Str = "DIAG"                          '

                        retcode = mapstr_process(map_source_str, Map_map_str, result_str, err_str) ' Do a DIAG call
                        IF retcode = 0 THEN                       ' If no errors
                           result_str = ""                        ' Clear the result
                           err_Str = ""                           ' Do a full call
                           retcode = mapstr_process(map_source_str, Map_map_str, result_str, err_str)
                        END IF                                    '

                        CONTROL SET TEXT hWnd, %Dlg_Result_Str, "¦" + result_str + "¦"
                        CONTROL SET TEXT hWnd, %Dlg_Error_Str, "RC=" + FORMAT$(retcode) + " " + Err_Str

                     CASE 3                                       ' Calc test page

                        CONTROL GET TEXT hCalc_Tab, %Dlg_Calc_S_Variable TO Calc_S_Variable' Get the modified data
                        Calc_S_Variable = TRIM$(UUCASE(Calc_S_Variable))' Test fields
                        IF TRIM$(Calc_S_Variable) = "" THEN Calc_S_Variable = "1"
                        IF VERIFY(Calc_S_Variable, "01234567890-+") = 0 THEN
                           Calc_S = VAL(Calc_S_Variable)          ' Do it the decimal way
                        ELSEIF VERIFY(Calc_S_Variable, "01234567890ABCDEF.") = 0 AND LEFT$(Calc_S_Variable, 1) = "." THEN
                           Calc_S = VAL("&H" + MID$(Calc_S_Variable, 2))' Do it the Hex way
                        ELSE
                           CONTROL SET TEXT hWnd, %Dlg_Error_Str, "Illegal characters in the Initial s string"
                           EXIT FUNCTION                          '
                        END IF                                    '
                        CONTROL GET CHECK hCalc_Tab, %Dlg_Calc_S_Incr TO Calc_S_Incr

                        CONTROL GET TEXT hCalc_Tab, %Dlg_Calc_RX_Variable TO Calc_RX_Variable' Get the modified data
                        Calc_RX_Variable = TRIM$(UUCASE(Calc_RX_Variable)) ' Test fields
                        IF Calc_RX_Variable = "" THEN             '
                           CONTROL SET TEXT hWnd, %Dlg_Error_Str, "Init value is empty."
                           EXIT FUNCTION                          '
                        END IF                                    '
                        IF VERIFY(Calc_RX_Variable, "01234567890-+") = 0 THEN
                           Calc_RX = VAL(Calc_RX_Variable)        ' Do it the decimal way
                        ELSEIF VERIFY(Calc_RX_Variable, "01234567890ABCDEF.") = 0 AND LEFT$(Calc_RX_Variable, 1) = "." THEN
                           Calc_RX = VAL("&H" + MID$(Calc_RX_Variable, 2))' Do it the Hex way
                        ELSE                                      '
                           CONTROL SET TEXT hWnd, %Dlg_Error_Str, "Illegal characters in the Initial R/X string"
                           EXIT FUNCTION                          '
                        END IF                                    '

                        CONTROL GET TEXT hCalc_Tab, %Dlg_Calc_Calc_Str TO Calc_Calc_Str  ' Get the modified data
                        IF TRIM$(Calc_Calc_Str) = "" THEN         '
                           CONTROL SET TEXT hWnd, %Dlg_Error_Str, "Calculation string is empty. "
                           EXIT FUNCTION                          '
                        END IF                                    '

                        '----- Actually run the test now
                        IF Calc_s_Incr THEN INCR Calc_S           ' Bump sequence if asked for
                        result_str = ""                           ' Clear the result
                        err_Str = ""                              '
                        Calc_Line = RND (1, 1000)                 '
                        Calc_Col  = RND (1, 80)                   '
                        Calc_Result = Calc_S                      '

                        retcode = mapstr_calc(Calc_RX, Calc_Calc_str, Calc_Error_str, Calc_Result, Calc_Line, Calc_Col)

                        CONTROL SET TEXT hCalc_Tab, %Dlg_Calc_S_Variable, FORMAT$(Calc_S)
                        CONTROL SET TEXT hWnd, %Dlg_Result_Str, DEC$(Calc_Result) + "  Hex --> ." & HEX$(Calc_Result, 16)
                        CONTROL SET TEXT hWnd, %Dlg_Error_Str, "RC=" + FORMAT$(retcode) + " " + Calc_Error_str

                  END SELECT                                      '
               END IF                                             '
         END SELECT
   END SELECT                                                     '
END FUNCTION

FUNCTION ToolTipCreate (BYVAL Wnd AS LONG) AS LONG
'---------- Create tooltips control if needed.                    '
   IF hToolTips = 0 THEN                                          '
      IF Wnd = 0 THEN Wnd = GetActiveWindow()                     '
      IF Wnd = 0 THEN EXIT FUNCTION                               '
      InitCommonControls                                          '
      hToolTips = CreateWindowEx(0, "tooltips_class32", "", %TTS_ALWAYSTIP OR %TTS_BALLOON, _
             0, 0, 0, 0, Wnd, BYVAL 0&, GetModuleHandle(""), BYVAL %NULL)
   END IF                                                         '
   FUNCTION = hToolTips                                           '
END FUNCTION

FUNCTION ToolTipSet (BYVAL Wnd AS LONG, BYVAL TXT AS STRING) AS LONG
'---------- Add a tooltip to a window/control
LOCAL ti AS TOOLINFO                                              '
   IF ToolTipCreate(GetParent(Wnd)) = 0 THEN EXIT FUNCTION        ' Ensure creation
   ti.cbSize   = LEN(ti)                                          '
   ti.uFlags   = %TTF_SUBCLASS OR %TTF_IDISHWND                   '
   ti.hWnd     = GetParent(Wnd)                                   '
   ti.uId      = hWnd                                             '

   '---------- Remove existing tooltip                            '
   IF SENDMESSAGE (hToolTips, %TTM_GETTOOLINFO, 0, BYVAL VARPTR(ti)) THEN
      SENDMESSAGE hToolTips, %TTM_DELTOOL, 0, BYVAL VARPTR(ti)    '
   END IF                                                         '
   ti.cbSize   = LEN(ti)                                          '
   ti.uFlags   = %TTF_SUBCLASS OR %TTF_IDISHWND                   '
   ti.hWnd     = GetParent(Wnd)                                   '
   ti.uId      = Wnd                                              '
   ti.lpszText = STRPTR(TXT)                                      '
   FUNCTION = SENDMESSAGE(hToolTips, %TTM_ADDTOOL, 0, BYVAL VARPTR(ti)) 'add tooltip
END FUNCTION

SUB Build_Dialog()
'---------- Build and start the Dialog
LOCAL hFixedFont AS DWORD
   FONT NEW "Courier New", 10, 1, 1, 1 TO hFixedFont              ' Build font for our Dialog text boxes
   FONT NEW "Tahoma", 9, 1, 1, 1 TO hTabFont                      ' Build font for tabs
   FONT NEW "Tahoma", 12, 0, 1, 1 TO hHeadFont                    ' Build font for heading
   DIALOG FONT DEFAULT "Tahoma", 10, 0, 0

   DIALOG NEW PIXELS, 0, "SPFLite Test RegEx / Mapping / Calculator Engines", 0, 0, 600, 375, _
          %WS_CAPTION OR %WS_THICKFRAME OR %WS_MINIMIZEBOX OR %WS_SYSMENU OR %WS_OVERLAPPEDWINDOW OR %WS_CLIPCHILDREN, 0 _
          TO hWnd
   DIALOG SET COLOR hWnd, %BLUE, %RGB_GAINSBORO
   CONTROL ADD IMAGEX, hWnd, %Dlg_Icon, "A", 5, 5, 32, 32, %SS_ICON OR %SS_CENTERIMAGE OR %SS_NOTIFY

   CONTROL ADD LABEL,  hWnd, %Dlg_H1, "Test SPFLite RegEx / Mapping / Calculator Engines ", 50, 10, 550, 30
   CONTROL SET COLOR   hWnd, %Dlg_H1, %BLUE, %RGB_GAINSBORO
   CONTROL SET FONT hWnd, %Dlg_H1, hHeadFont                      '

   '----- Now add a Tab control
   CONTROL ADD TAB, hWnd, %Dlg_Tab, "", 0, 40, 600, 170, %TCS_OWNERDRAWFIXED
   CONTROL HANDLE hWnd, %Dlg_Tab TO hTab                          ' Get handle for Tab
   CONTROL SET FONT hWnd, %Dlg_Tab, hTabFont                      '
   CONTROL SET COLOR hWnd, %Dlg_Tab, %BLUE, %RGB_GAINSBORO        ' Default color it

   '----- Insert a page for the Regex testing
   TAB INSERT PAGE hWnd, %Dlg_Tab, 1, 0, "RegEx Strings" CALL DlgCallback TO hRegEx_Tab
   DIALOG SET COLOR hRegEx_Tab, %BLUE, %RGB_GAINSBORO
   CONTROL ADD TEXTBOX, hRegEx_Tab, %Dlg_RegEx_Case_Str, Regex_Case_Str, 10, 10, 20, 20
   CONTROL SET FONT     hRegEx_Tab, %Dlg_RegEx_Case_Str, hFixedFont
   ToolTipSet (GetDlgItem(hRegEx_Tab, %Dlg_RegEx_Case_Str), " Enter the CASE value - C / T to be used as a default. ")
   CONTROL ADD LABEL, hRegEx_Tab, %Dlg_RegEx_Case_Str_Text, "Default CASE value - C / T", 40, 10, 300, 16
   CONTROL SET COLOR  hRegEx_Tab, %Dlg_RegEx_Case_Str_Text, %BLUE, %RGB_GAINSBORO


   CONTROL ADD TEXTBOX, hRegEx_Tab, %Dlg_RegEx_Test_Str, Regex_Test_Str, 10, 45, 575, 20
   CONTROL SET FONT     hRegEx_Tab, %Dlg_RegEx_Test_Str, hFixedFont
   ToolTipSet (GetDlgItem(hRegEx_Tab, %Dlg_RegEx_Test_Str), " Source string for the test. ")
   CONTROL ADD LABEL, hRegEx_Tab, %Dlg_RegEx_Test_Str_Text, "Source string for the test", 10, 68, 450, 16
   CONTROL SET COLOR  hRegEx_Tab, %Dlg_RegEx_Test_Str_Text, %BLUE, %RGB_GAINSBORO

   CONTROL ADD TEXTBOX, hRegEx_Tab, %Dlg_RegEx_Str, RegEx_Str, 10, 95, 575, 20
   CONTROL SET FONT     hRegEx_Tab, %Dlg_RegEx_Str, hFixedFont
   ToolTipSet (GetDlgItem(hRegEx_Tab, %Dlg_RegEx_Str), " Enter the Regex expression string. ")
   CONTROL ADD LABEL, hRegEx_Tab, %Dlg_RegEx_Str_Text, " Regex expression string - No R'...' framing required ", 10, 118, 450, 16
   CONTROL SET COLOR  hRegEx_Tab, %Dlg_RegEx_Str_Text, %BLUE, %RGB_GAINSBORO

   '----- Insert a page for the Mapping testing
   TAB INSERT PAGE hWnd, %Dlg_Tab, 2, 0, "Mapping Strings" CALL DlgCallback TO hMap_Tab
   DIALOG SET COLOR hMap_Tab, %BLUE, %RGB_GAINSBORO
   CONTROL ADD TEXTBOX, hMap_Tab, %Dlg_Map_Case_Str, Map_Case_Str, 10, 10, 20, 20
   CONTROL SET FONT     hMap_Tab, %Dlg_Map_Case_Str, hFixedFont
   ToolTipSet (GetDlgItem(hMap_Tab, %Dlg_Map_Case_Str), " Enter the CASE value - C / T to be used as a default. ")
   CONTROL ADD LABEL, hMap_Tab, %Dlg_Map_Case_Str_Text, "Default CASE value - C / T", 40, 10, 300, 16
   CONTROL SET COLOR  hMap_Tab, %Dlg_Map_Case_Str_Text, %BLUE, %RGB_GAINSBORO

   CONTROL ADD TEXTBOX, hMap_Tab, %Dlg_Map_Source_Str, Map_Source_Str, 10, 45, 575, 20
   CONTROL SET FONT     hMap_Tab, %Dlg_Map_Source_Str, hFixedFont
   ToolTipSet (GetDlgItem(hMap_Tab, %Dlg_Map_Source_Str), " Enter the source string to be re-mapped ")
   CONTROL ADD LABEL, hMap_Tab, %Dlg_Map_Source_Str_Text, "Source string for the test", 10, 68, 450, 16
   CONTROL SET COLOR  hMap_Tab, %Dlg_Map_Source_Str_Text, %BLUE, %RGB_GAINSBORO

   CONTROL ADD TEXTBOX, hMap_Tab, %Dlg_Map_Map_Str, Map_Map_Str, 10, 95, 575, 20
   CONTROL SET FONT     hMap_Tab, %Dlg_Map_Map_Str, hFixedFont
   ToolTipSet (GetDlgItem(hMap_Tab, %Dlg_Map_Map_Str), " Enter the Mapping string for the test. ")
   CONTROL ADD LABEL, hMap_Tab, %Dlg_Map_Map_Str_Text, " Mapping string - No M'...' framing required ", 10, 118, 450, 16
   CONTROL SET COLOR  hMap_Tab, %Dlg_Map_Map_Str_Text, %BLUE, %RGB_GAINSBORO

   '----- Insert a page for the Calculator testing
   TAB INSERT PAGE hWnd, %Dlg_Tab, 3, 0, "Calculator Strings" CALL DlgCallback TO hCalc_Tab
   DIALOG SET COLOR hCalc_Tab, %BLUE, %RGB_GAINSBORO
   CONTROL ADD TEXTBOX, hCalc_Tab, %Dlg_Calc_S_Variable, Calc_S_Variable, 10, 10, 75, 20
   CONTROL SET FONT     hCalc_Tab, %Dlg_Calc_S_Variable, hFixedFont
   ToolTipSet (GetDlgItem(hCalc_Tab, %Dlg_Calc_S_Variable), " Enter the initial value for variable S if needed. ")
   CONTROL ADD LABEL, hCalc_Tab, %Dlg_Calc_S_Variable_Text, "Initial value for the S variable", 90, 11, 200, 16
   CONTROL SET COLOR  hCalc_Tab, %Dlg_Calc_S_Variable_Text, %BLUE, %RGB_GAINSBORO
   CONTROL ADD CHECKBOX, hCalc_Tab, %Dlg_Calc_S_Incr, "Increment value before eash test?", 300, 10, 250, 16,
   CONTROL SET COLOR  hCalc_Tab, %Dlg_Calc_S_Incr, %BLUE, %RGB_GAINSBORO
   CONTROL SET CHECK  hCalc_Tab, %Dlg_Calc_S_Incr, Calc_S_Incr


   CONTROL ADD TEXTBOX, hCalc_Tab, %Dlg_Calc_RX_Variable, Calc_RX_Variable, 10, 45, 575, 20
   CONTROL SET FONT     hCalc_Tab, %Dlg_Calc_RX_Variable, hFixedFont
   ToolTipSet (GetDlgItem(hCalc_Tab, %Dlg_Calc_RX_Variable), " Enter the initial value for the R and X variables. ")
   CONTROL ADD LABEL, hCalc_Tab, %Dlg_Calc_RX_Variable_Text, "Initial value for the R and X variables ", 10, 68, 450, 16
   CONTROL SET COLOR  hCalc_Tab, %Dlg_Calc_RX_Variable_Text, %BLUE, %RGB_GAINSBORO

   CONTROL ADD TEXTBOX, hCalc_Tab, %Dlg_Calc_Calc_Str, Calc_Calc_Str, 10, 95, 575, 20
   CONTROL SET FONT     hCalc_Tab, %Dlg_Calc_Calc_Str, hFixedFont
   ToolTipSet (GetDlgItem(hCalc_Tab, %Dlg_Calc_Calc_Str), " Enter the Calculation string to be evaluated. ")
   CONTROL ADD LABEL, hCalc_Tab, %Dlg_Calc_Calc_Str_Text, " Calculation string to be evaluated ", 10, 118, 450, 16
   CONTROL SET COLOR  hCalc_Tab, %Dlg_Calc_Calc_Str_Text, %BLUE, %RGB_GAINSBORO

   CONTROL ADD BUTTON,  hWnd, %Dlg_Test_Button, "Run Test", 265, 220, 75, 24, %WS_BORDER OR %BS_DEFAULT
   ToolTipSet (GetDlgItem(hWnd, %Dlg_Test_Button), " Click to run the test. ")

   CONTROL ADD LABEL,  hWnd, %Dlg_H2, "Results", 10, 245, 550, 30
   CONTROL SET COLOR   hWnd, %Dlg_H2, %BLUE, %RGB_GAINSBORO
   CONTROL SET FONT hWnd, %Dlg_H2, hHeadFont                      '

   CONTROL ADD TEXTBOX, hWnd, %Dlg_Result_Str, Result_Str, 10, 275, 575, 20
   CONTROL SET FONT     hWnd, %Dlg_Result_Str, hFixedFont
   ToolTipSet (GetDlgItem(hWnd, %Dlg_Result_Str), " Result string for the last test will appear here. ")
   CONTROL ADD LABEL, hWnd, %Dlg_Result_Str_Text, "Returned result string", 10, 298, 300, 16
   CONTROL SET COLOR  hWnd, %Dlg_Result_Str_Text, %BLUE, %RGB_GAINSBORO

   CONTROL ADD TEXTBOX, hWnd, %Dlg_Error_Str, Result_Str, 10, 325, 575, 20
   CONTROL SET FONT     hWnd, %Dlg_Error_Str, hFixedFont
   ToolTipSet (GetDlgItem(hWnd, %Dlg_Error_Str), " Any Error message for the last test will appear here. ")
   CONTROL ADD LABEL, hWnd, %Dlg_Error_Str_Text, "Returned error message", 10, 348, 300, 16
   CONTROL SET COLOR  hWnd, %Dlg_Error_Str_Text, %BLUE, %RGB_GAINSBORO

   IF pageno > 0 THEN _                                            ' If a page request
      TAB SELECT hWnd, %Dlg_Tab, pageno                            ' set it

   DIALOG SHOW MODAL hWnd CALL DlgCallback                         ' Display it all

END SUB

SUB TabHighLight(HDLG AS LONG, wParm AS LONG, lParm AS LONG)
'---------- Highlight the selected Dialog tab
LOCAL lDISPtr AS DRAWITEMSTRUCT PTR, zCap AS ASCIIZ * 50
LOCAL ti AS TC_ITEM
LOCAL nColor, hBrush AS LONG
   lDisPtr = lparm                                                '
   ti.mask = %TCIF_TEXT                                           '
   ti.pszText = VARPTR(zCap)                                      '
   ti.cchTextMax = SIZEOF(zCap)                                   '
   TabCtrl_GetItem(GetDlgItem(HDLG, wParm), @lDisptr.itemID, ti)
   @lDisptr.rcItem.nTop = @lDisptr.rcItem.nTop + 2                '

   SetTextColor @lDisPtr.hDc, %BLACK                              '
   IF @lDisPtr.ItemState = %ODS_SELECTED THEN                     '
      SetBkColor @lDisPtr.hDc, %YELLOW                            '
      hBrush = CreateSolidBrush(%YELLOW)                          '
      SelectObject @lDisptr.hDc, hBrush                           '
      FillRect @lDisptr.hDc, @lDisptr.rcItem, hBrush              '
   ELSE                                                           '
      SetBkColor @lDisPtr.hDc, %WHITE                             '
      hBrush = CreateSolidBrush(%WHITE)                           '
      SelectObject @lDisptr.hDc, hBrush                           '
      FillRect @lDisptr.hDc, @lDisptr.rcItem, hBrush              '
   END IF                                                         '

   DrawText @lDisptr.hDc, zCap, LEN(zCap), @lDisptr.rcItem, %DT_SINGLELINE OR %DT_CENTER
   DeleteObject hBrush                                            '
 END SUB

'/-----------------------------------------------------------------------------
'/ FAKE function - sSetTable ("GET", var_name)
'/
'/ for testing, simulate sSetTable() with ENVIRON$() = GET parm is ignored
'/-----------------------------------------------------------------------------

FUNCTION                      sSetTable                                       _
(  BYVAL spflite_GET          AS STRING                                       _
,  BYVAL var_name             AS STRING                                       _
   )                          AS STRING

   LOCAL v                    AS STRING
   MEntry
   '/ PB's ENVIRON$ is broken, so we have to call Windows API to get it
   v = mapstr_getenv (var_name)
   FUNCTION = v
   MExit
END FUNCTION

#INCLUDE ONCE "_mapping.inc"
#INCLUDE ONCE "_mapping_calc.inc"
