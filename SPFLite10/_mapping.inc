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


'/*****************************************************************************
'/****   ----------------------------   ***************************************
'/****   INTERFACE TO MAPPING STRINGS   ***************************************
'/****   ----------------------------   ***************************************
'/*****************************************************************************


FUNCTION                      mapstr_process                                  _
(  BYVAL source_arg           AS STRING                                       _
,  BYVAL map_arg              AS STRING                                       _
,  BYREF result_arg           AS STRING                                       _
,  BYREF err_arg              AS STRING                                       _
   )                          AS LONG                     '/ --> AS RETURN CODE

   LOCAL case_CT_arg          AS STRING
   LOCAL rc                   AS LONG
   LOCAL diag_mode            AS LONG
   LOCAL sequence_arg         AS QUAD

'/ STATIC static_sequence_num AS QUAD                         '/ no longer used


   MEntry


   case_CT_arg = TP.PrfPCase

   '/--------------------------------------------------------------------------
   '/ THE STATEMENT ABOVE IS VALID ONLY IN A METHOD, NOT IN A FUNCTION
   '/--------------------------------------------------------------------------


   '/ set up dummy args and reset enumeration count on each DIAG call

   IF err_arg = "DIAG" THEN                        '/ this is a validation call
      err_arg = ""                                 '/ clear DIAG flag by caller
      TP.mapstr_sequence_num = 0            '/ reset counter for each DIAG call
      source_arg = "0000"                      '/ dummy value for source string
      diag_mode = 1
   ELSE
      TP.mapstr_sequence_num = TP.mapstr_sequence_num + 1  '/ get next sequence
      diag_mode = 0
   END IF

   sequence_arg = TP.mapstr_sequence_num

   rc =  mapstr_execute             _
         ( source_arg               _
         , map_arg                  _
         , result_arg               _
         , err_arg                  _
         , case_CT_arg              _
         , CODEPTR (tran_A2E_tab)   _
         , CODEPTR (tran_E2A_tab)   _
         , sequence_arg             _
         , diag_mode                _
         )

   FUNCTION = rc

   MExitFunc
END FUNCTION



'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   IMPLEMENTATION OF MAPPING STRINGS   **********************************
'/****   ---------------------------------   **********************************
'/*****************************************************************************
'/
'/    FUNCTION                      mapstr_execute
'/    (  BYVAL source_arg           AS STRING
'/    ,  BYVAL map_arg              AS STRING
'/    ,  BYREF result_arg           AS STRING
'/    ,  BYREF err_arg              AS STRING
'/    ,  BYVAL case_CT_arg          AS STRING
'/    ,  BYVAL tran_A2E_arg         AS BYTE PTR
'/    ,  BYVAL tran_E2A_arg         AS BYTE PTR
'/    ,  BYVAL sequence_arg         AS QUAD
'/    ,  BYVAL diag_mode            AS LONG
'/       )                          AS LONG               '/ --> AS RETURN CODE
'/
'/
'/ This is the mapping-string function to be called by SPFLite
'/
'/
'/ Arguments:
'/
'/    source_arg     String DATA as found by CHANGE command string-1 operand
'/
'/    map_arg        String-2 OPERAND as parsed by CHANGE command
'/                      M'xxx'  or  'xxx'M  is reduced to just the  xxx  value
'/
'/    result_arg     RETURNED: contents of source_arg string after reformatting
'/
'/    err_arg        RETURNED: error message, or "" if mapping was successful
'/
'/    case_CT_arg    Current state of PROFILE CASE setting, either "C" or "T"
'/                   If CASE unavailable, caller should pass " "
'/
'/    tran_A2E_arg   BYTE PTR to 256-byte Ansi to EBCDIC translation table
'/
'/    tran_E2A_arg   BYTE PTR to 256-byte EBCDIC to Ansi translation table
'/
'/    sequence_arg   sequential number for SD and SX commands; 0 = DIAG
'/
'/    diag_mode      1 = DIAG mode, 0 - normal mode
'/
'/ Return code:
'/
'/    0        OK    - Mapping was successful
'/    8        ERROR - Mapping string is malformed or other error detected
'/
'/ Notes:      1. On return code 0, caller should check whether returned
'/                source_arg = result_arg.
'/
'/             2. If supplied string is not changed from original value,
'/                update of file could be skipped to save time.
'/
'/             3. If translation tables are unavailable, caller should
'/                pass null pointers (0).  Either both pointers should be
'/                valid, or both null; do not mix null with non-null.
'/
'/             4. When edit file is Ansi or Unicode, translation table pointers
'/                should point to the default EBCDIC 1140 code page tables.
'/                If it is not convenient to do that, pass null pointers.
'/
'/                When edit file is explicitly EBCDIC or is a user-provided
'/                code page, those code pages should be supplied.
'/
'/                For example, if a user defined a code page of MYTABLE,
'/                the pointers for the translation tables should be:
'/
'/                   tran_A2E_arg = Ansi to MYTABLE translation table
'/                   tran_E2A_arg = MYTABLE to Ansi translation table
'/
'/*****************************************************************************


'/*****************************************************************************
'/****   ------------------------   *******************************************
'/****   BEGINNING OF MAPSTR CODE   *******************************************
'/****   ------------------------   *******************************************
'/*****************************************************************************


'/ to clarify that a function returns a PTR, since we can't say it directly:

MACRO RETURNED_BYTE_PTR       = DWORD


'/ a way to better document what a VERIFY call is doing:

MACRO VERIFIED                = 0


'/ an idea that didn't get used:

'  MACRO mapstr_bool(x)       = ABS(SGN(x))
'  MACRO mapstr_not(x)        = (1 - mapstr_bool(x))


'/-----------------------------------------------------------------------------
'/ mapstr_dest_t is a data structure used to simplify calls to parse a
'/ destination reference, as used in the Delete command code.
'/-----------------------------------------------------------------------------

TYPE mapstr_dest_t

   dest_num1                  AS LONG
   dest_num2                  AS LONG

   dest_num1_used             AS LONG
   dest_num2_used             AS LONG

   end_relative               AS LONG
   suppresss                  AS LONG
   dest_diag_str              AS STRINGZ * 128

END TYPE '/ mapstr_dest_t



'/-----------------------------------------------------------------------------
'/ mapstr_cmd_t is a data structure used to simplify calls to various parsing
'/ routines, to gather in one place the data needed for a command, and to avoid
'/ complex calling sequences.
'/-----------------------------------------------------------------------------

TYPE mapstr_cmd_t

   cmd_num1                   AS LONG
   cmd_num2                   AS LONG

   cmd_num1_used              AS LONG
   cmd_num2_used              AS LONG

   cmd_str1_used              AS LONG
   cmd_str2_used              AS LONG

   cmd_ID                     AS LONG
   end_relative               AS LONG
   cmd_sep                    AS LONG
   cmd_type                   AS LONG
   cmd_rls_mode               AS LONG          '/ has RTL, LTR or STD scan mode

   use_result_str             AS LONG
   uses_str2                  AS LONG
   null_left_str              AS LONG       '/ 1 = string found like PD='right'
   takes_col_ref              AS LONG
   suppress_auto_ref          AS LONG        '/ if this cmd should not auto ref

   dxd_ordinal                AS LONG     '/ 1-9 if scanning, 0 if non-scanning
   dxd_fmt_width              AS LONG     '/          number preceding F Z or V
   dxd_fmt_width_used         AS LONG     '/ dxd_fmt_width appears before F/Z/V
   dxd_sign_pos               AS LONG     '/     where +/- found in valud, or 0
   dxd_sign_len               AS LONG     '/    length of +/- prefix inc. space
   dxd_numeric                AS QUAD     '/      extracted value of dxd source

   cmd_name                   AS STRINGZ * 8

   cmd_diag_str               AS STRINGZ * 128
   txt_diag_str               AS STRINGZ * 128
   num_diag_str               AS STRINGZ * 128

   cmd_str2_code              AS STRING * 1              '/ if used, "=" or ":"
   txt_case_CTX_1             AS STRING * 1             '/ C T or X suffix code
   txt_case_CTX_2             AS STRING * 1             '/ CTX if 2-string pair

   dxd_fmt_type               AS STRING * 1                         '/ F Z or V
   dxd_sign_code              AS STRING * 1           '/ sign code S C or blank

END TYPE '/ mapstr_cmd_t


%mapstr_retcode_OK            = 0      '/        return code from main function
%mapstr_retcode_ERR           = 8      '/        return code from main function
%mapstr_DIAG                  = (-1)   '/          if code running in DIAG mode

%mapstr_none                  = 0      '/ general init and <> default case code

%mapstr_upper                 = 1      '/    state for UC and > upper case code
%mapstr_lower                 = 2      '/    state for LC and < lower case code                                '
%mapstr_title                 = 3      '/                          state for TC
%mapstr_sentence              = 4      '/                          state for SC
%mapstr_invert                = 5      '/  state for IC and >< invert case code
%mapstr_dash                  = 6      '/                column-range specified
%mapstr_plus                  = 7      '/           remaining-columns specified
%mapstr_slash                 = 8      '/  mixed-mode columns, forward ordering
%mapstr_backslash             = 9      '/ mixed-mode columns, backward ordering
%mapstr_other_cmd             = 10     '/    general classification of cmd type
%mapstr_copying_cmd           = 11     '/    general classification of cmd type
%mapstr_modifying_cmd         = 12     '/    general classification of cmd type

%mapstr_dxd_dec               = 1
%mapstr_dxd_hex               = 2

%mapstr_not_sig               = 0
%mapstr_sig                   = 1

%mapstr_rls_STD               = 0       '/ left-or-right scan mode is STD
%mapstr_rls_RTL               = 1       '/ left-or-right scan mode is RTL
%mapstr_rls_LTR               = 2       '/ left-or-right scan mode is LTR

%mapstr_str2_none             = 0       '/ handling of string-pair operands
%mapstr_str2_allowed          = 1
%mapstr_str2_required         = 2

%mapstr_m_cmd_none            = 0       '/ m_cmd_move_mode setting
%mapstr_m_cmd_append          = 1       '/ m_cmd_move_mode setting
%mapstr_m_cmd_insert          = 2       '/ m_cmd_move_mode setting
%mapstr_m_cmd_replace         = 3       '/ m_cmd_move_mode setting

%mapstr_esv_max               = 10      '/ num of EV/SV nesting levels allowed

%esv_mode_DV                  = 1
%esv_mode_SV                  = 3

'/  %esv_mode_EV                  = 2

%esv_state_beg                = 1
%esv_state_end                = 2

'/-----------------------------------------------------------------------------
'/ command code enumerations
'/-----------------------------------------------------------------------------


ENUM mapstr_cmd SINGULAR

   mapstr_cmd_none

   mapstr_cmd_ID_AQ                     '/ accent-quote literal
   mapstr_cmd_ID_AX                     '/ Ansi to hex, prefer UC
   mapstr_cmd_ID_C                      '/ Center the result string
   mapstr_cmd_ID_CC                     '/ CASE C mode
   mapstr_cmd_ID_CT                     '/ CASE T mode
   mapstr_cmd_ID_D                      '/ Delete char columns from result str
   mapstr_cmd_ID_DC                     '/ Delete char values from result str
   mapstr_cmd_ID_DD                     '/ decimal to decimal numeric convert
   mapstr_cmd_ID_DQ                     '/ double-quote literal
   mapstr_cmd_ID_DV                     '/ Dynamic Value
   mapstr_cmd_ID_DX                     '/ decimal to hex numeric convert

'/ mapstr_cmd_ID_EV                     '/ include Env Variable

   mapstr_cmd_ID_EX                     '/ EBCDIC to hex, prefer UC
   mapstr_cmd_ID_IC                     '/ Invert Case
   mapstr_cmd_ID_L                      '/ Left justify the result string
   mapstr_cmd_ID_LC                     '/ Lower Case
   mapstr_cmd_ID_M                      '/ Move characters within result
   mapstr_cmd_ID_P                      '/ Pad the result string - both sides
   mapstr_cmd_ID_PA                     '/ Pivot data and Align
   mapstr_cmd_ID_PD                     '/ Pivot Data
   mapstr_cmd_ID_PL                     '/ Pad the result string - left side
   mapstr_cmd_ID_PR                     '/ Pad the result string - right side
   mapstr_cmd_ID_QD                     '/ quit use of Dynamic Value - private
   mapstr_cmd_ID_QE                     '/ quit use of Env Variable - private
   mapstr_cmd_ID_QS                     '/ quit use of Set Variable - private
   mapstr_cmd_ID_R                      '/ Right justify the result string
   mapstr_cmd_ID_RA                     '/ Replace All
   mapstr_cmd_ID_RC                     '/ Replace Characters via trans table
   mapstr_cmd_ID_RL                     '/ Replace Leftmost
   mapstr_cmd_ID_RP                     '/ Repeat the result string
   mapstr_cmd_ID_RR                     '/ Replace Rightmost
   mapstr_cmd_ID_RV                     '/ Reverse the result string
   mapstr_cmd_ID_SC                     '/ Sentence Case
   mapstr_cmd_ID_SD                     '/ Sequence in decimal
   mapstr_cmd_ID_SQ                     '/ single-quote literal
   mapstr_cmd_ID_SV                     '/ include Set Variable
   mapstr_cmd_ID_SX                     '/ Sequence in hex
   mapstr_cmd_ID_T                      '/ Trim the result string - both sides
   mapstr_cmd_ID_TC                     '/ Title Case
   mapstr_cmd_ID_TL                     '/ Trim the result string - left side
   mapstr_cmd_ID_TR                     '/ Trim the result string - right side
   mapstr_cmd_ID_UC                     '/ Upper Case
   mapstr_cmd_ID_X                      '/ Exchange strings
   mapstr_cmd_ID_XA                     '/ hex to Ansi
   mapstr_cmd_ID_XD                     '/ hex to decimal numeric convert
   mapstr_cmd_ID_XE                     '/ hex to EBCDIC
   mapstr_cmd_ID_XX                     '/ hex to hex numeric convert
   mapstr_cmd_ID_Z                      '/ Zero-suppress the result field

END ENUM



FUNCTION mapstr_is_quote_b (x AS BYTE) AS LONG

   MEntry

   IF x = ASC("`") _
   OR x = ASC($SQ) _
   OR x = ASC($DQ) THEN
      FUNCTION = 1
   ELSE
      FUNCTION = 0
   END IF

   MExit

END FUNCTION



FUNCTION mapstr_is_alnum_b (x AS BYTE) AS LONG

   MEntry

   '/ hex DF is German ess-zett letter, which has no upper case variant
   '/ ess-zett is only a letter if ENGLISH-only is disabled

   IF x = &hDF _                                      '/ German ess-zett letter
   AND ISFALSE ENV.ENGchars THEN       '/ we are not restricted to English only
      FUNCTION = 1

   ELSEIF UUCASE(CHR$(x)) <> LLCASE(CHR$(x)) _
   OR     (x >= ASC("0") AND x <= ASC("9")) THEN
      FUNCTION = 1

   ELSE
      FUNCTION = 0
   END IF

   MExit

END FUNCTION



FUNCTION mapstr_is_alpha_b (x AS BYTE) AS LONG

   MEntry

   '/ hex DF is German ess-zett letter, which has no upper case variant
   '/ ess-zett is only a letter if ENGLISH-only is disabled

   IF x = &hDF _                                      '/ German ess-zett letter
   AND ISFALSE ENV.ENGchars THEN       '/ we are not restricted to English only
      FUNCTION = 1

   ELSEIF UUCASE(CHR$(x)) <> LLCASE(CHR$(x)) THEN
      FUNCTION = 1

   ELSE
      FUNCTION = 0
   END IF

   MExit

END FUNCTION



FUNCTION mapstr_is_digit_b (x AS BYTE) AS LONG

   MEntry

   IF (x >= ASC("0") AND x <= ASC("9")) THEN
      FUNCTION = 1
   ELSE
      FUNCTION = 0
   END IF

   MExit

END FUNCTION



FUNCTION mapstr_is_symbol_c (x AS STRING) AS LONG

   MEntry

   IF (x >= "A" AND x <= "Z") _
   OR (x >= "a" AND x <= "z") _
   OR (x >= "0" AND x <= "9") _
   OR (x = "_") THEN
      FUNCTION = 1
   ELSE
      FUNCTION = 0
   END IF

   MExit

END FUNCTION



FUNCTION mapstr_is_alnum_c (x AS STRING) AS LONG

   MEntry

   IF (x >= "A" AND x <= "Z") _
   OR (x >= "a" AND x <= "z") _
   OR (x >= "0" AND x <= "9") THEN
      FUNCTION = 1
   ELSE
      FUNCTION = 0
   END IF

   MExit

END FUNCTION



FUNCTION mapstr_is_alpha_c (x AS STRING) AS LONG

   MEntry

   IF (x >= "A" AND x <= "Z") _
   OR (x >= "a" AND x <= "z") THEN
      FUNCTION = 1
   ELSE
      FUNCTION = 0
   END IF

   MExit

END FUNCTION



FUNCTION mapstr_is_digit_c (x AS STRING) AS LONG

   MEntry

   IF (x >= "0" AND x <= "9") THEN
      FUNCTION = 1
   ELSE
      FUNCTION = 0
   END IF

   MExit

END FUNCTION



'/-----------------------------------------------------------------------------
'/ mapstr_getenv
'/
'/ PB's ENVIRON$ is broken when the value has embedded equal signs
'/ use #  for =
'/ use !# for #
'/ use !! for !
'/-----------------------------------------------------------------------------

FUNCTION                      mapstr_getenv                                   _
(  BYVAL var_name             AS STRING                                       _
   )                          AS STRING

   LOCAL env_text             AS STRING
   LOCAL new_text             AS STRING
   LOCAL n                    AS LONG
   LOCAL c                    AS STRING
   LOCAL esc                  AS BYTE


   MEntry

   env_text = ENVIRON$ (var_name)

   IF env_text = "" THEN
      FUNCTION = ""
      MExitFunc
   END IF

   esc = 0
   new_text = ""

   FOR n = 1 TO LEN (env_text)
      c = MID$ (env_text, n, 1)

      IF esc THEN
         esc = 0
         new_text &= c                      '/ any escaped char is copied as is

      ELSE
         IF c = "!" THEN
            esc = 1
         ELSEIF c = "#" THEN
            new_text &= "="                     '/ make substitution of = for #
         ELSE
            new_text &= c
         END IF
      END IF
   NEXT

   '/ a trailing ! is copied as us

   IF esc THEN                      '/ only possible if ! is last char in value
      new_text &= c
   END IF

   FUNCTION = new_text

   MExit

END FUNCTION



'/-----------------------------------------------------------------------------
'/ mapstr_fetch_env_var
'/
'/ get value of Windows Environment variable based on supplied name
'/-----------------------------------------------------------------------------

FUNCTION                      mapstr_fetch_env_var                            _
(  BYVAL var_name             AS STRING                                       _
   )                          AS STRING

   LOCAL v                    AS STRING

   MEntry

   v = mapstr_getenv (var_name)

   FUNCTION = v

   MExit

END FUNCTION



'/-----------------------------------------------------------------------------
'/ mapstr_fetch_set_var
'/
'/ get value of SPFLite SET variable based on supplied name
'/-----------------------------------------------------------------------------

FUNCTION                      mapstr_fetch_set_var                            _
(  BYVAL var_name             AS STRING                                       _
   )                          AS STRING

   LOCAL v                    AS STRING

   MEntry

   v = sSetTable ("GET", var_name)

   FUNCTION = v

   MExit

END FUNCTION



'/-----------------------------------------------------------------------------
'/ mapstr_valid_env_name
'/
'/ verify that supplied base name for SV or EV is valid
'/-----------------------------------------------------------------------------

FUNCTION                      mapstr_valid_env_name                           _
(  BYVAL var_name             AS STRING                                       _
   )                          AS LONG
   LOCAL n                    AS LONG
   LOCAL c                    AS STRING


   MEntry

   IF LEN(var_name) = 0 THEN
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc
   END IF

   FOR n = 1 TO LEN(var_name)
      c = MID$ (var_name, n, 1)

      IF mapstr_is_symbol_c (c) = 0 THEN
         FUNCTION = %mapstr_retcode_ERR
         MExitFunc
      END IF
   NEXT

   FUNCTION = %mapstr_retcode_OK

   MExit

END FUNCTION



'/-----------------------------------------------------------------------------
'/ mapstr_form_env_name
'/
'/ given a string from an EV or QE command, create the corresponding var name
'/-----------------------------------------------------------------------------

FUNCTION                      mapstr_form_env_name                            _
(  BYVAL var_name             AS STRING                                       _
   )                          AS STRING
   MEntry

   FUNCTION = UUCASE ("MAPSTR.EV." & var_name)

   MExit

END FUNCTION



'/-----------------------------------------------------------------------------
'/ mapstr_form_set_name
'/
'/ given a string from an SV or QS command, create the corresponding var name
'/-----------------------------------------------------------------------------

FUNCTION                      mapstr_form_set_name                            _
(  BYVAL var_name             AS STRING                                       _
   )                          AS STRING
   MEntry

   FUNCTION = UUCASE ("MAPSTR.SV." & var_name)

   MExit

END FUNCTION



'/-----------------------------------------------------------------------------
'/ mapstr_insert_into_col_range
'/
'/ take a source value and insert it into specific columns of target string
'/ in case of error, return "" string
'/-----------------------------------------------------------------------------

FUNCTION                      mapstr_insert_into_col_range                    _
(  BYVAL target_val           AS STRING                                       _
,  BYVAL col1                 AS LONG                                         _
,  BYVAL col2                 AS LONG                                         _
,  BYVAL source_val           AS STRING                                       _
   )                          AS STRING

   LOCAL first_str            AS STRING
   LOCAL last_str             AS STRING
   LOCAL answer               AS STRING

   MEntry


   IF LEN(target_val) = 0     _
   OR LEN(source_val) = 0     _
   OR col1 < 1                _
   OR col1 > LEN(target_val)  _
   OR col2 < 1                _
   OR col2 > LEN(target_val)  THEN

      FUNCTION = target_val       '/ parameters are bad, leave target unchanged
      MExitFunc

   END IF


   IF col1 = 1 THEN
      first_str = ""
   ELSE
      first_str = LEFT$ (target_val, col1-1)
   END IF

   IF col2 = LEN(target_val) THEN
      last_str = ""
   ELSE
      last_str = MID$ (target_val, col2+1)
   END IF

   answer = BUILD$ (first_str, source_val, last_str)
   FUNCTION = answer

   MExit

END FUNCTION                                    '/ mapstr_insert_into_col_range



'/-----------------------------------------------------------------------------
'/ mapstr_insert_at_col_position
'/
'/ take a source value and insert it into at a specific location in the
'/ target string.  if the insertion point is after the end, do an append.
'/ if the insertion point is 0, do a prepend.
'/-----------------------------------------------------------------------------

FUNCTION                      mapstr_insert_at_col_position                   _
(  BYVAL target_val           AS STRING                                       _
,  BYVAL target_col           AS LONG                                         _
,  BYVAL source_val           AS STRING                                       _
   )                          AS STRING

   MEntry

   IF target_col <= 1 THEN                     '/ source goes before the target
      FUNCTION = source_val & target_val

   ELSEIF target_col > LEN(target_val) THEN     '/ source goes after the target
      FUNCTION = target_val & source_val

   ELSE
      FUNCTION = STRINSERT$ (target_val, source_val, target_col)
   END IF

   MExit

END FUNCTION                                   '/ mapstr_insert_at_col_position



'/-----------------------------------------------------------------------------
'/ mapstr_get_dxd_source
'/
'/ extract a "floating" numeric field from a string for DD/DX/XD/XX command
'/ in case of error, return "" string
'/-----------------------------------------------------------------------------

FUNCTION                      mapstr_get_dxd_source                           _
(  BYVAL move_text            AS STRING                                       _
,  BYVAL dxd_source_type      AS BYTE                                         _
,  BYREF cmd                  AS mapstr_cmd_t                                 _
,  BYREF sign_arg             AS STRING                                       _
   )                          AS STRING

   LOCAL digit_set            AS STRING

   LOCAL n_str                AS STRING
   LOCAL u_str                AS STRING
   LOCAL scan_str             AS STRING
   LOCAL result_str           AS STRING
   LOCAL c                    AS STRING
   LOCAL n                    AS LONG

   LOCAL left_num_col         AS LONG
   LOCAL right_num_col        AS LONG

   LOCAL left_sign_col        AS LONG
   LOCAL right_sign_col       AS LONG

   LOCAL curr_ordinal         AS LONG
   LOCAL prev_ordinal         AS LONG

   LOCAL curr_valid           AS BYTE
   LOCAL prev_valid           AS BYTE

   LOCAL loc_ordinal          AS LONG
   LOCAL loc_rls_mode         AS LONG

   MEntry

   sign_arg = ""

   '/ the ordinal and RTL tell us how to look for the number
   '/ an ordinal of 2 means the second number in the string
   '/ RTL is 1 if scanning right to left; 0 means LTR

   loc_ordinal = cmd.dxd_ordinal
   loc_rls_mode = cmd.cmd_rls_mode

   IF dxd_source_type = %mapstr_dxd_hex THEN
      digit_set = "0123456789ABCDEF"
   ELSE
      digit_set = "0123456789"
   END IF

   cmd.dxd_numeric = 0
   left_num_col = 0
   right_num_col = 0

   '/ get rid of extra blanks and check for empty value

   n_str = TRIM$(move_text)

   IF ISNULL(n_str) THEN                    '/ after trim there is no data left
      FUNCTION = ""
      MexitFunc
   END IF

   u_str = UUCASE(n_str)         '/ simplify check for hex digits and 0x prefix


   '/ for now, we are going to ignore the distinction between a command like
   '/ DD and 1DD, and treat them the same.  at some point, we may have to be
   '/ more subtle than that, but for now, it helps to get the design finished.

   IF loc_ordinal = 0 THEN
      loc_rls_mode = %mapstr_rls_LTR     '/ if looking for 1 value, must be LTR
      loc_ordinal = 1                 '/ treat as if looking for first one, LTR
   END IF


   '/ we are looking for a positional value, from 1st to 9th position.
   '/ it's hard to parse the variable-format dxd commands as it is, so we can't
   '/ handle any more than a 1-digit positional indicator; that's why only 1-9

   IF loc_rls_mode = %mapstr_rls_RTL THEN      '/ RTL scan needs fancy footwork
      scan_str = STRREVERSE$ (u_str)      '/ this lets us scan it left to right
   ELSE
      scan_str = u_str
   END IF

   curr_ordinal = 0                '/ which number within field have we reached
   curr_valid = 0   '/ tracks transition from valid to invalid digit characters

   n = 1                                         '/ initial index into scan_str

   DO WHILE n <= LEN(scan_str)        '/ note that scanned string is upper case
      prev_valid = curr_valid
      c = MID$ (scan_str, n, 1)                        '/ grab a possible digit

      '/-----------------------------------------------------------------------
      '/ if this is a C number, a prefix of 0x or 0X is bypassed
      '/
      '/ SNAG 1: if we are scanning right-to-left, the 0x prefix has
      '/ become an x0 suffix, so we have to check this correctly
      '/ there is no need to check this if 0 (or X) is last char of string
      '/
      '/ SNAG 2: in LTR, if 0x was (most likely) preceded by a non-number, then
      '/ finding a 0x doesn't change the non-number state.  but in RTL, we are
      '/ scanning a number, and x0 ends that.  so, we have to check for 0x at
      '/ the beginning, but for x0 at the end.
      '/
      '/ this hack is not required for languages like BASIC, since they use a
      '/ prefix like &H which has no digits in it and we won't get confused.
      '/-----------------------------------------------------------------------

      IF  cmd.dxd_sign_code = "C" _                       '/ look for 0x prefix
      AND n < LEN(scan_str) _                   '/ enough space to skip 2 chars
      AND loc_rls_mode <> %mapstr_rls_RTL THEN
         IF c = "0" AND MID$ (scan_str, n+1, 1) = "X" THEN
            n += 2                               '/ count the forward 0x prefix
            ITERATE DO
         END IF
      END IF


      IF VERIFY (c, digit_set) = VERIFIED THEN             '/ valid digit found

         IF prev_valid = 0 THEN          '/ this is the start of a number field
            left_num_col = 0            '/ so code below will start a new field
         END IF

         curr_valid = 1
         right_num_col = n

         IF left_num_col = 0 THEN                                  '/ new field
            left_num_col = n                         '/ remember start of field
            curr_ordinal += 1         '/ this is the next "nth" number in order
         END IF

      ELSE
         curr_valid = 0       '/ end of a number field, or never got to one yet

         '/ if this is 2nd number, and we're looking for the 2nd number ...

         IF loc_ordinal = curr_ordinal THEN
            EXIT DO                        '/ ...then this is the one we wanted
         END IF

         IF  cmd.dxd_sign_code = "C" _           '/ look for reversed 0x prefix
         AND n < LEN(scan_str) _                '/ enough space to skip 2 chars
         AND loc_rls_mode = %mapstr_rls_RTL THEN
            IF c = "X" AND MID$ (scan_str, n+1, 1) = "0" THEN
               n += 2                           '/ count the reversed 0x prefix
               ITERATE DO
            END IF
         END IF
      END IF

      INCR n                   '/ COUNT THE DIGIT - OTHERWISE GET STUCK IN LOOP
   LOOP

   '/ if the nth field was not found after a complete scan for it, we're done

   IF left_num_col = 0 _                                 '/ never found a value
   OR curr_ordinal <> loc_ordinal THEN              '/ or not the one we wanted
      FUNCTION = ""
      MexitFunc
   END IF

   '/--------------------------------------------------------------------------
   '/ if we did a right-to-left scan, the column numbers are end-relative
   '/ and backwards, so we need to correct them.  example:
   '/
   '/                0  000   1     111  2
   '/ fwd cols:      123456789012345678901
   '/                   ---         ---
   '/ orig data:     ...123...456...789...      LEN = 21
   '/                   ***         ***
   '/ rev data:      ...987...654...321...
   '/                   ---         ---
   '/                2  111     1   000  0
   '/ bwd col:       109876543210987654321
   '/
   '/ task: find the first number going right to left (the number 789).
   '/
   '/ the "left" and "right" column appears to be 4 and 6 after the scan
   '/ because we used forward-relative column numbering in the scan, but
   '/ the real column numbers are 16 and 18.
   '/
   '/ subtracting the columns from (LEN(u_str)+1) gives,
   '/
   '/    left:  (21+1) - 4 = 22 - 4 = 18    actually the right column
   '/    right: (21+1) - 6 = 22 - 6 = 16    actually the left column
   '/
   '/ the correct results are thus obtained by swapping the left and right
   '/ columns and applying the calculation.
   '/
   '/--------------------------------------------------------------------------

   IF loc_rls_mode = %mapstr_rls_RTL THEN
      SWAP left_num_col, right_num_col
      left_num_col = (LEN(u_str) + 1) - left_num_col
      right_num_col = (LEN(u_str) + 1) - right_num_col
   END IF

   '/ see if there is a sign, and capture it if the user cares about it

   left_sign_col  = 0
   right_sign_col = 0

   n = left_num_col

   '/--------------------------------------------------------------------------
   '/ determine the presence of a C prefix like 0x and +/- sign.
   '/ method: starting tothe left of the column number of the left-most digit,
   '/ scan backwards, skipping blanks, until +/- found, or the start of the
   '/ string, or a non-sign character.  if sign found, capture it.
   '/
   '/ if this is a C number, there could be both a 0x prefix and a +/- sign
   '/ first check for the prefix, then check any signed type for +/-
   '/ example:  ... 0x77    n = 7         0x99   n = 3
   '/ n-:       6543210                   210
   '/--------------------------------------------------------------------------

   IF  cmd.dxd_sign_code = "C" _
   AND n > 2 _                                '/ it is possible for 0x to exist
   AND MID$ (u_str, n-2, 2) = "0X" THEN                           '/ found prefix
      left_sign_col = n-2
      right_sign_col = left_num_col - 1
      n - = 2                                '/ make n the index to the 0 of 0x
   END IF

   IF cmd.dxd_sign_code <> " " THEN                           '/ sign is C or S
      DECR n                          '/ skip any leading space going backwards
      DO WHILE n > 1 AND MID$ (u_str, n, 1) = " "
         DECR n
      LOOP

      IF n >= 1 THEN
         c =  MID$ (u_str, n, 1)
         IF c = "+" _
         OR c = "-" THEN
            left_sign_col = n
            right_sign_col = left_num_col - 1
         END IF
      END IF
   END IF

   '/ pass back the sign prefix, if any was found

   IF left_sign_col > 0 THEN
      sign_arg = MID$(n_str, left_sign_col TO right_sign_col)
   END IF

   result_str = MID$(n_str, left_num_col TO right_num_col)

   '/ pass back the converted value
   '/ the QUAD numeric value is captured in case conversion between hex and dec

   IF dxd_source_type = %mapstr_dxd_hex THEN
      cmd.dxd_numeric = VAL ("&H" & result_str)
   ELSE
      cmd.dxd_numeric = VAL (result_str)
   END IF

   FUNCTION = result_str

   MExit

END FUNCTION                                           '/ mapstr_get_dxd_source


'/-----------------------------------------------------------------------------
'/ mapstr_CT_equal()
'/
'/ do compare as case sensitive or not, based on CT code
'/ if equal, return 1, else 0
'/-----------------------------------------------------------------------------

FUNCTION                      mapstr_CT_equal                                 _
(  BYVAL one                  AS STRING                                       _
,  BYVAL two                  AS STRING                                       _
,  BYVAL ct_arg               AS STRING                                       _
   )                          AS LONG
   LOCAL result_eq            AS LONG

   MEntry


   IF  ct_arg = "T" _
   AND UUCASE(one) = UUCASE(two) THEN
      result_eq = 1

   ELSEIF one = two THEN
      result_eq = 1

   ELSE
      result_eq = 0

   END IF

   FUNCTION = result_eq

   MExit

END FUNCTION                                                 '/ mapstr_CT_equal


'/-----------------------------------------------------------------------------
'/ mapstr_conform_case()
'/
'/ propagate case from first argument to second
'/ case-conformant processing propagates the casing from the last-encountered
'/ letter.  if for some reason no letter was found, it's treated as if the last
'/ one was lower case.
'/-----------------------------------------------------------------------------

SUB                           mapstr_conform_case                             _
(  BYREF x1                   AS STRING                                       _
,  BYREF x2                   AS STRING                                       _
   )

   LOCAL n                    AS LONG
   LOCAL len_1                AS LONG
   LOCAL len_2                AS LONG

   LOCAL c2                   AS STRING

   LOCAL test_c               AS STRING
   LOCAL test_uc              AS STRING
   LOCAL test_lc              AS STRING

   LOCAL make_lower           AS BYTE

   MEntry

   len_1 = LEN(x1)
   len_2 = LEN(x2)
   make_lower = 1                        '/ create default case for conformancy

   FOR n = 1 TO len_2
      c2 = MID$ (x2, n, 1)                    '/ grab a character from string 2

      '/ we always have a c2, but maybe not a c1
      '/ we only capture a character from x1 if it's a letter, otherwise leave
      '/ it alone and use its prior state.
      '/ hex DF is German es-zett letter which has no upper case variant

      IF n <= len_1 THEN
         test_c  = MID$ (x1, n, 1)
         test_uc = UUCASE (test_c)
         test_lc = LLCASE (test_c)

         IF test_c = CHR$(&hDF) _             '/ it's the German es-zett letter
         AND ISFALSE ENV.ENGchars THEN        '/ not restricted to English only
            make_lower = 1                         '/ so use  lower case for x2

         ELSEIF test_uc <> test_lc THEN                    '/ x1[n] is a letter
            IF test_c = test_uc THEN               '/ it's an upper case letter
               make_lower = 0                      '/ so use  upper case for x2
            ELSE                                   '/ it's an lower case letter
               make_lower = 1                      '/ so use  lower case for x2
            END IF
         END IF

      END IF

      IF make_lower = 1 THEN
         c2 = LLCASE (c2)
      ELSE
         c2 = UUCASE (c2)
      END IF

      MID$ (x2, n, 1) = c2    '/ replace a U/L cased character back to string 2
   NEXT

   MExit

END SUB                                                  '/ mapstr_conform_case


'/-----------------------------------------------------------------------------
'/ mapstr_SCASE()
'/
'/ implement Sentence Case function, since PB's MCASE$ won't do it
'/-----------------------------------------------------------------------------

FUNCTION                      mapstr_SCASE                                    _
(  BYVAL inpstr               AS STRING                                       _
   )                          AS STRING

   LOCAL c                    AS STRING
   LOCAL answer               AS STRING
   LOCAL prev_c               AS STRING
   LOCAL in_sentence          AS BYTE
   LOCAL n                    AS LONG

   MEntry

   '/ method: beginning of string is assumed to be beginning of a sentence.
   '/ from that point on, spans of ". " start new sentences.
   '/ a "word" is a letter at beginning of string, or else preceded by a space.

   answer = LLCASE (inpstr)
   in_sentence = 1
   c = " "                   '/ so first time thru, prev_c will look like blank

   FOR n = 1 TO LEN(answer)
      prev_c = c
      c = MID$(answer, n, 1)

      IF  prev_c = "." _
      AND c = " " THEN
         in_sentence = 1

      ELSEIF UUCASE(c) <> LLCASE(c) THEN                    '/ this is a letter
         IF  in_sentence = 1 _                        '/ first word of sentence
         AND prev_c = " " THEN
            in_sentence = 0
            MID$(answer, n, 1) = UUCASE(c)
         END IF
      END IF

   NEXT

   FUNCTION = answer

   MExit

END FUNCTION                                                    '/ mapstr_SCASE



'/-----------------------------------------------------------------------------
'/ mapstr_TCASE()
'/
'/ implement Title Case function, since PB's MCASE$ won't do it
'/-----------------------------------------------------------------------------

FUNCTION                      mapstr_TCASE                                    _
(  BYVAL inpstr               AS STRING                                       _
   )                          AS STRING

   LOCAL c                    AS STRING
   LOCAL answer               AS STRING
   LOCAL prev_c               AS STRING
   LOCAL n                    AS LONG

   MEntry

   '/ method: beginning of string is assumed to be beginning of a word.
   '/ a "word" is a letter at beginning of string, or else preceded by a space.

   answer = LLCASE (inpstr)
   c = " "                   '/ so first time thru, prev_c will look like blank

   FOR n = 1 TO LEN(answer)
      prev_c = c
      c = MID$(answer, n, 1)

      IF  UUCASE(c) <> LLCASE(c) _                          '/ this is a letter
      AND prev_c = " " THEN                             '/ at beginning of word
         MID$(answer, n, 1) = UUCASE(c)
      END IF

   NEXT

   FUNCTION = answer

   MExit

END FUNCTION                                                    '/ mapstr_TCASE



'/-----------------------------------------------------------------------------
'/ mapstr_ICASE()
'/
'/ invert case: make upper case letters lower case, and vice versa
'/-----------------------------------------------------------------------------

FUNCTION                      mapstr_ICASE                                    _
(  BYVAL inpstr               AS STRING                                       _
   )                          AS STRING

   LOCAL answer               AS STRING
   LOCAL c                    AS STRING
   LOCAL uc                   AS STRING
   LOCAL lc                   AS STRING
   LOCAL n                    AS LONG

   MEntry

   answer = ""

   FOR n = 1 TO LEN(inpstr)
      c = MID$(inpstr, n, 1)
      uc = UUCASE(c)
      lc = LLCASE(c)

      IF uc = lc THEN                                   '/ non-alpha characters
         answer &= c                                              '/ keep as is

      ELSEIF c = uc THEN                                           '/ was upper
         answer &= lc                                             '/ make lower

      ELSE                                                         '/ was lower
         answer &= uc                                             '/ make upper
      END IF
   NEXT

   FUNCTION = answer

   MExit

END FUNCTION                                                    '/ mapstr_ICASE



'/-----------------------------------------------------------------------------
'/ mapstr_handle_case()
'/
'/ adjust case based on case code
'/-----------------------------------------------------------------------------

FUNCTION                      mapstr_handle_case                              _
(  BYVAL move_text            AS STRING                                       _
,  BYVAL mapstr_case          AS BYTE                                         _
   )                          AS STRING

   MEntry

   IF mapstr_case = %mapstr_upper THEN
      FUNCTION = UUCASE (move_text)

   ELSEIF mapstr_case = %mapstr_lower THEN
      FUNCTION = LLCASE (move_text)

   ELSEIF mapstr_case = %mapstr_sentence THEN
      FUNCTION = mapstr_SCASE (move_text)

   ELSEIF mapstr_case = %mapstr_title THEN
      FUNCTION = mapstr_TCASE (move_text)

   ELSEIF mapstr_case = %mapstr_invert THEN
      FUNCTION = mapstr_ICASE (move_text)

   ELSE
      FUNCTION = move_text

   END IF

   MExit

END FUNCTION                                              '/ mapstr_handle_case



'/-----------------------------------------------------------------------------
'/ mapstr_is_delim()
'/
'/ return 1 if valid mapping string delimiter, else 0
'/-----------------------------------------------------------------------------

FUNCTION                      mapstr_is_delim                                 _
(  BYVAL c                    AS STRING                                       _
   )                          AS LONG

   MEntry

   '/ comma was removed as a delimiter; it's not really needed ...

   IF c = " " _
   OR c = ";" _
   OR c = CHR$(0) THEN                     '/ be sure to stop at nul terminator
      FUNCTION = 1
   ELSE
      FUNCTION = 0
   END IF

   MExit

END FUNCTION                                                 '/ mapstr_is_delim



'/-----------------------------------------------------------------------------
'/ mapstr_is_hex_delim()
'/
'/ return 1 if valid mapping hex literal string delimiter, else 0.
'/ this is a test for "readability" characters that can be ignored.
'/ to allow string lists in hex, we cannot allow comma as a "noise" delimiter
'/ for compatibility with pivot operands, only underscore is now allowed.
'/ blank was formerly allowed, but that would make pivots values ambiguous.
'/-----------------------------------------------------------------------------

FUNCTION                      mapstr_is_hex_delim                             _
(  BYVAL c                    AS STRING                                       _
   )                          AS LONG

   MEntry

   IF c = "_" THEN
      FUNCTION = 1
   ELSE
      FUNCTION = 0
   END IF

   MExit

END FUNCTION                                             '/ mapstr_is_hex_delim



'/-----------------------------------------------------------------------------
'/ mapstr_validate_text_range
'/
'/ confirm that a text-range operand like [x-y] is in correct format
'/ to be sure we don't run off the string, at any point in the scan if we get a
'/ nul terminator, it's an error.  this should not happen.
'/
'/ if valid, return the byte offset needed to bypass the trailing ] character
'/ if invalid, return 0
'/-----------------------------------------------------------------------------

FUNCTION                      mapstr_validate_text_range                      _
(  BYVAL cp                   AS BYTE PTR                                     _
   )                          AS LONG

   LOCAL n                    AS LONG
   LOCAL fmt_error            AS LONG

   MEntry

   IF @cp <> ASC("[") THEN                      '/ didn't get leading [ bracket
      FUNCTION = 0                     '/ should not happen, just being careful
      MexitFunc
   END IF

   '/ format we are looking for is [x-y] or [x-y,a-b,p-q] etc.
   '/                              01234    01234   01234
   '/
   '/ at each point, we need a middle dash, and either a comma or end bracket

   n = 0
   fmt_error = 0

   DO WHILE @cp <> ASC("]")

      IF @cp[1] = 0         _
      OR @cp[2] <> ASC("-") _
      OR @cp[3] = 0         THEN
         fmt_error = 1

      ELSEIF @cp[4] <> ASC(",") _
      AND    @cp[4] <> ASC("]") THEN
         fmt_error = 1
      END IF

      IF fmt_error THEN
         FUNCTION = 0                  '/ should not happen, just being careful
         MexitFunc
      END IF

      cp += 4
      n += 4
   LOOP

   FUNCTION = n + 1                          '/ n value is 1 past the ] bracket

   MExit

END FUNCTION                                      '/ mapstr_validate_text_range



'/-----------------------------------------------------------------------------
'/ mapstr_X2A()
'/
'/ convert a string of hex-digit characters into ASCII data
'/-----------------------------------------------------------------------------

FUNCTION                      mapstr_X2A                                      _
(  BYVAL hexstr               AS STRING                                       _
,  BYREF hex_str_diag         AS STRING                                       _
   )                          AS STRING

   LOCAL work                 AS STRING
   LOCAL answer               AS STRING
   LOCAL c                    AS STRING
   LOCAL val_str              AS STRING
   LOCAL chr_str              AS STRING
   LOCAL n                    AS LONG

   MEntry

   work = ""


   '/ verify that hexstr is valid, omit readability chars, and store work value
   '/ allow an embedded '=' and treat as hex 3D

   '/ that way users don't have to embed hex 3D in their RC hex strings
   '/ we want to let users be able to say RC'1122=3344'X
   '/ even though '=' is not a hex digit - we are going to fake it

   FOR n = 1 TO LEN(hexstr)
      c = MID$ (hexstr, n, 1)

      IF mapstr_is_hex_delim(c) = 0 THEN                '/ skip separator chars
         work &= c                                        '/ append a hex digit
      END IF
   NEXT

   IF VERIFY (work, "0123456789ABCDEFabcdef") <> VERIFIED THEN     '/ bad value
      hex_str_diag = "Mapping string has invalid hex digits"
      FUNCTION = ""                                       '/ value is malformed
      MExitFunc
   END IF

   IF LEN(work) < 2 _
   OR LEN(work) MOD 2 <> 0 THEN                '/ value not pairs of hex digits
      hex_str_diag = "Mapping hex-string is null or odd number of digits"
      FUNCTION = ""                                       '/ value is malformed
      MExitFunc
   END IF

   answer = ""

   FOR n = 1 TO LEN(work)-1 STEP 2
      val_str = "&H" & MID$(work, n, 2)                    '/ form "&Hxx" value
      chr_str = CHR$ (VAL (val_str))
      answer &= chr_str
   NEXT

   FUNCTION = answer

   MExit

END FUNCTION                                                      '/ mapstr_X2A



'/-----------------------------------------------------------------------------
'/ mapstr_normalize_columns()
'/
'/ when column numbers are mixed mode like x/y* or x\y*, convert to normal
'/ column numbers and change separator type to "dash" mode.
'/-----------------------------------------------------------------------------

SUB                           mapstr_normalize_columns                        _
(  BYREF cmd                  AS mapstr_cmd_t                                 _
,  BYVAL col_max              AS LONG                                         _
   )

   MEntry


   '/ reject conditions we can't handle here

   IF  cmd.cmd_sep <> %mapstr_slash _
   AND cmd.cmd_sep <> %mapstr_backslash THEN    '/ not one of the special types
      MExitSub                                 '/ this does not need adjustment
   END IF

   IF cmd.cmd_num1_used = 0 _
   OR cmd.cmd_num2_used = 0 _
   OR cmd.end_relative  = 0 THEN         '/ cmd does not have the expected data
      MExitSub                                 '/ this does not need adjustment
   END IF


   '/ if both numbers are larger than the column max, the associated command is
   '/ going to be rejected eventually.  if so, just change the type to dash
   '/ mode and get out.

   IF  cmd.cmd_num1 > col_max _
   AND cmd.cmd_num2 > col_max THEN
      cmd.cmd_sep = %mapstr_dash
      MExitSub
   END IF


   '/ we can now individually adjust the columns to fit within the max

   cmd.cmd_num1 = MIN& (cmd.cmd_num1, col_max)
   cmd.cmd_num2 = MIN& (cmd.cmd_num2, col_max)


   '/--------------------------------------------------------------------------
   '/ since we know that the second number is end-relative (it's in x/y* or
   '/ x\y format) and the first column is in normal format, we want to change
   '/ the mixed-mode format to fully normal columns, and clear the end relative
   '/ flag.  since the / and \ don't adhere to low-high vs. high-low order for
   '/ indicating reversal, we have to fix this.
   '/
   '/ there is no easy wasy to avoid two swaps without complicating the code.
   '/--------------------------------------------------------------------------

   cmd.cmd_num2 = (col_max + 1) - cmd.cmd_num2   '/ now both numbers are normal

   IF cmd.cmd_num1 > cmd.cmd_num2 THEN
      SWAP cmd.cmd_num1, cmd.cmd_num2                 '/ now in ascending order
   END IF

   IF cmd.cmd_sep = %mapstr_backslash THEN      '/ they wanted reversed columns
      SWAP cmd.cmd_num1, cmd.cmd_num2                '/ now in descending order
   END IF

   cmd.cmd_sep = %mapstr_dash         '/ column reference is now a normal range
   cmd.end_relative  = 0                       '/ and is no longer end relative

   MExit

END SUB                                             '/ mapstr_normalize_columns



'/-----------------------------------------------------------------------------
'/ mapstr_parse_numbers()
'/
'/ parse and extract any mapping-string compatible numeric field parameter
'/
'/ valid formats are:   n   n+   x-y   x-y*   x/y*   x\y*
'/
'/ the associated "separator" of "+", "-" or none is recorded
'/
'/ if the number or number pair ends in * it is "end relative notation"
'/ we note the existence of this flag and pass it back to the caller.
'/
'/ we allow the form of x/y* in end-relative notation.  when used, this causes
'/ the x value to be negated.  this serves as a flag so that when end relative
'/ notation is being calculated, the x value is made positive again, and is not
'/ part of the end-relative calculations.  this allows column ranges to have
'/ mixed normal and end-relative column numbers.
'/
'/ the notation x\y* is like x/y* but takes the columns in reverse order.
'/ the / and \ notations are transformed later into - notation.
'/
'/ if the user fails to use * with / or \ it is treated like a normal - range
'/-----------------------------------------------------------------------------

FUNCTION                      mapstr_parse_numbers                            _
(  BYVAL cp                   AS BYTE PTR                                     _
,  BYREF cmd                  AS mapstr_cmd_t                                 _
   )                          AS RETURNED_BYTE_PTR

   LOCAL end_relative_allowed AS BYTE
   LOCAL c                    AS STRING * 1

   MEntry

   cmd.cmd_num1 = 0
   cmd.cmd_num2 = 0

   cmd.cmd_num1_used = 0
   cmd.cmd_num2_used = 0

   cmd.cmd_sep = %mapstr_none                         '/ assume no - or + found
   cmd.end_relative = 0                         '/ assume no * suffix on number
   end_relative_allowed = 1

   cmd.num_diag_str = ""

   '/ for commands like L-9, it means same as L0-9
   '/ for ordinary columns, a leading dash like -9 is illegal
   '/ column 0 is illegal in cmd.end_relative notation; we check for that later
   '/ cmd.end_relative notation is illegal for values like L-9

   IF @cp = ASC("-") THEN
      end_relative_allowed = 0
      GOTO Dash_number_format
   END IF


   IF mapstr_is_digit_b (@cp) = 0 THEN
      cmd.num_diag_str = "Mapping number expected digit, got '" _
         & CHR$(@cp) & "'"
      FUNCTION = 0                                  '/ initial char not a digit
      MExitFunc                                       '/ JIC; should not happen

   END IF

   cmd.cmd_num1_used = 1

   '/ extract the first (and possibly only) number field

   DO WHILE mapstr_is_digit_b (@cp)
      cmd.cmd_num1 = (cmd.cmd_num1 * 10) + (@cp - ASC("0"))
      INCR cp
   LOOP


   '/ see if any separator followed the first number

   IF mapstr_is_delim (CHR$ (@cp)) THEN                         '/ no separator
      FUNCTION = cp                                '/ a single number ends here
      MExitFunc

   END IF


   '/ look for 'remainder number' notation, like 5+
   '/ note that a value of number+* is illegal - code is already handling this

   IF @cp = ASC("+") THEN
      cmd.cmd_sep = %mapstr_plus
      INCR cp                                               '/ skip over + sign

      IF mapstr_is_delim (CHR$ (@cp)) THEN              '/ good delimiter found
         FUNCTION = cp
      ELSE
         cmd.num_diag_str = "Mapping number expected delimiter after +, " _
            & "got '" & CHR$(@cp) & "'"
         FUNCTION = 0                                   '/ malformed + notation
      END IF

      MExitFunc
   END IF


   '/ look for 'range' notation

   Dash_number_format:
   '/----------------

   c = CHR$ (@cp)
   IF @cp = ASC("-") THEN                                       '/  x-y or x-y*
      cmd.cmd_sep = %mapstr_dash

   ELSEIF @cp = ASC("/") THEN                                   '/  x/y or x/y*
      cmd.cmd_sep = %mapstr_slash

   ELSEIF @cp = ASC("\") THEN                                   '/  x\y or x\y*
      cmd.cmd_sep = %mapstr_backslash
   END IF


   IF cmd.cmd_sep <> %mapstr_none THEN
      INCR cp                                            '/ skip over separator

      '/ see if second number if present

      IF mapstr_is_digit_b (@cp) = 0 THEN
         cmd.num_diag_str = "Mapping number expected digit after " & c & _
            ", got '" & CHR$(@cp) & "'"
         FUNCTION = 0                             '/ dash not followed by digit
         MExitFunc                                  '/ malformed range notation

      END IF

      cmd.cmd_num2_used = 1

      '/ extract the second number field

      DO WHILE mapstr_is_digit_b (@cp)
         cmd.cmd_num2 = (cmd.cmd_num2 * 10) + (@cp - ASC("0"))
         INCR cp
      LOOP

   END IF


   '/ see if end_relative notation is used

   IF  end_relative_allowed = 1 _
   AND @cp = ASC("*") THEN
      INCR cp                                           '/ skip over * asterisk
      cmd.end_relative = 1                                    '/ pass back flag
   END IF


   '/ a notation like x/y or x\y without * for end relative is invalid
   '/ pretending that these are dash notation would just be confusing

   IF cmd.end_relative = 0 THEN
      IF cmd.cmd_sep = %mapstr_slash _
      OR cmd.cmd_sep = %mapstr_backslash THEN
         cmd.num_diag_str = _
            "Mapping number uses / or \ without * end-relative notation"
         FUNCTION = 0
         MExitFunc
      END IF
   END IF


   '/ see if a good delimiter followed the last number parsed
   '/ a \ reversal code or alpha action code is not allowed as a delimiter

   IF mapstr_is_delim (CHR$ (@cp)) _                       '/ general delimiter
   OR @cp = ASC("<") _                                  '/ case conversion code
   OR @cp = ASC(">") _                                  '/ case conversion code
   OR mapstr_is_quote_b (@cp) THEN                             '/ inserted text
      cmd.cmd_sep = cmd.cmd_sep                   '/ store the - or / or \ flag
      FUNCTION = cp                                '/ number properly delimited

   ELSE
      cmd.num_diag_str = _
         "Mapping number improperly followed by " & CHR$ (@cp)
      FUNCTION = 0                             '/ number not properly delimited
   END IF

   MExit

END FUNCTION                                            '/ mapstr_parse_numbers



'/-----------------------------------------------------------------------------
'/ mapstr_parse_txt()
'/
'/ parse and extract any mapping-string compatible quoted string parameter
'/ the data may be ASCII, or hex if a trailing "X" or "x" is used
'/ if hex specified, call our hex converter to get the corresponding value
'/
'/ for special operands like PD='right' there is no string1 just string2
'/-----------------------------------------------------------------------------

FUNCTION                      mapstr_parse_txt                                _
(  BYVAL cp                   AS BYTE PTR                                     _
,  BYREF txt_str1             AS STRING                                       _
,  BYREF txt_str2             AS STRING                                       _
,  BYREF cmd                  AS mapstr_cmd_t                                 _
   )                          AS RETURNED_BYTE_PTR

   LOCAL quote                AS BYTE

   LOCAL work_str             AS STRING
   LOCAL work_CTX             AS STRING

   LOCAL hex_str              AS STRING
   LOCAL hex_str_diag         AS STRING
   LOCAL rep_num              AS LONG
   LOCAL n                    AS LONG
   LOCAL p                    AS LONG

   MEntry

   txt_str1 = ""
   txt_str2 = ""

   cmd.txt_diag_str = ""                         '/ reason why txt parse failed

   cmd.txt_case_CTX_1 = " "                  '/ assume C/T/X suffix not present
   cmd.txt_case_CTX_2 = " "                  '/ assume C/T/X suffix not present
   cmd.cmd_str2_code  = " "                      '/ saves "=" or ":" string sep
   cmd.cmd_str1_used  = 0
   cmd.cmd_str2_used  = 0

   '/ if cp pointer is at an "=" or ":" right now, we have the special format
   '/ if so, we don't even try to parse string1, but go right to string 2

   IF  @cp <> ASC("=") _                             '/ we have a normal string
   AND @cp <> ASC(":") THEN                     '/ it DOESN'T start with = or :

      '/ string might be 'xxx' or 'xxx'='xxx' or 'xxx':'xxx'

      '/ for PD etc. string can be ='xxx' with string 2 but no string 1
      '/ but since we are here, they didn't do it that way ...

      GOSUB sub_get_string_item

      '/-----------------------------------------------------------------------
      '/ for open strings, not part of a command, we allow a replication factor
      '/ this works for plain strings, as well as C, T and X
      '/ sub_get_string_item has already skipped the type code if present
      '/-----------------------------------------------------------------------

      IF  cmd.cmd_ID = %mapstr_none _
      AND mapstr_is_digit_b (@cp) THEN
         GOSUB sub_get_string_replication
      END IF

      txt_str1           = work_str
      cmd.txt_case_CTX_1 = work_CTX
      cmd.cmd_str1_used  = 1

   END IF


   '/--------------------------------------------------------------------------
   '/ after gosub sub_get_string_item, @cp will point to character that ended
   '/ the string.  if it's an "=" or ":", get second item.
   '/ the = notation and the : colon notation are alternative syntax
   '/ the colon notation means that the two values both have type T codes.
   '/
   '/ 'one'='two' is used for normal string handling
   '/ 'one':'two' is used for case-conformant string changes, where applicable
   '/ for commands like PA/PD, case-conformance doesn't apply, but it still
   '/ processes the type codes, in case the pivot or boundary are alphabetic.
   '/--------------------------------------------------------------------------

   IF @cp = ASC("=") _                            '/ there is an 'xx'='xx' pair
   OR @cp = ASC(":") THEN                         '/ there is an 'xx':'xx' pair

      IF cmd.uses_str2 = %mapstr_str2_none THEN
         cmd.txt_diag_str = "Mapping command " & cmd.cmd_name _
               & " does not use a string-pair"
         FUNCTION = 0                                            '/ parse error
         MExitFunc
      END IF

      cmd.cmd_str2_code = CHR$(@cp)            '/ save the "=" or ":" separator

      IF mapstr_is_quote_b (@cp[1]) THEN
         INCR cp                    '/ skip over = or : and point to next quote

         GOSUB sub_get_string_item

         txt_str2           = work_str
         cmd.txt_case_CTX_2 = work_CTX
         cmd.cmd_str2_used  = 1

      ELSE                                 '/ got  'value'=  but no second half
         cmd.txt_diag_str = "Mapping command string-pair malformed"
         FUNCTION = 0                                            '/ parse error
         MExitFunc
      END IF

   '/ here, neither a ":" or "=" was not found, so there is ONLY STRING 1

   ELSEIF cmd.uses_str2 <> %mapstr_str2_none THEN

      '/-----------------------------------------------------------------------
      '/ STRING SPLITTING LOGIC - ONLY STRING 1 IS PRESENT
      '/
      '/ if only one string is used but it splittable, convert it into two
      '/ if possible.  for commands like RA, string 2 can be omitted.
      '/ for RC, they can only have one = sign in one string, otherwise two
      '/ strings are needed.
      '/-----------------------------------------------------------------------

      IF LEN(txt_str1) = 0 THEN
         cmd.txt_diag_str = "Mapping string-pair cannot be null"
         FUNCTION = 0                                            '/ parse error
         MExitFunc
      END IF

      n = TALLY (txt_str1, "=")

      IF  n = 0 _
      AND cmd.uses_str2 = %mapstr_str2_required THEN
         cmd.txt_diag_str = "Mapping string-pair value requires an = sign"
         FUNCTION = 0                                            '/ parse error
         MExitFunc
      END IF

      IF n > 1 THEN
         cmd.txt_diag_str = "Mapping string-pair value allows only one = sign"
         FUNCTION = 0                                            '/ parse error
         MExitFunc
      END IF

      '/ one equal sign found, but it cannot appear first, because we never
      '/ allow for a null string-1 value

      p = 0                                           '/ location of equal sign

      IF n = 1 THEN
         p = INSTR (txt_str1, "=")
         IF p < 2 THEN
            cmd.txt_diag_str = "Mapping string-pair cannot be null on left"
            FUNCTION = 0                                         '/ parse error
            MExitFunc
         END IF
      END IF


      '/-----------------------------------------------------------------------
      '/ at this point, we create string 2.  we don't demand that the string 2
      '/ value be non-null, we just make it, and let the command routines
      '/ decide if a null string-2 is OK or not.
      '/-----------------------------------------------------------------------

      IF p > 1 THEN
         cmd.cmd_str2_used = 1              '/ we are synthsizing string-2 here

         cmd.txt_case_CTX_2 = cmd.txt_case_CTX_1           '/ RULE 2 and RULE 3

         IF p < LEN (txt_str1) THEN          '/ = sign is not at right hand end
            txt_str2 = MID$ (txt_str1, p+1)          '/ everything after = sign
         ELSE
            txt_str2 = ""           '/ value like 'abc=' means string 2 is null
         END IF

         txt_str1 = LEFT$ (txt_str1, p-1)           '/ everything before = sign

      END IF
   END IF



   '/ if we have two explicit strings, or one that got split in two, deal with
   '/ the value pair here

   IF cmd.cmd_str2_used = 1 THEN

      '/-----------------------------------------------------------------------
      '/ when string 2 type is T, status is always case conformant
      '/ if string 1 type is blank if string 2 is T, T is copied to string 1
      '/-----------------------------------------------------------------------

      IF  cmd.txt_case_CTX_2 = "T" THEN
         cmd.cmd_str2_code = ":"
         IF  cmd.txt_case_CTX_1 = " " THEN
             cmd.txt_case_CTX_1 = "T"
         END IF
      END IF


      '/ if string-2 is hex, convert to char and treat as type C

      IF cmd.txt_case_CTX_2 = "X" THEN

         hex_str = mapstr_X2A (txt_str2, hex_str_diag)

         IF ISNULL (hex_str) THEN                     '/ hex value is malformed
            cmd.txt_diag_str = hex_str_diag
            FUNCTION = 0                                         '/ parse error
            MExitFunc
         END IF

         cmd.txt_case_CTX_2 = "C"       '/ after hex conversion, type is like C
         txt_str2 = hex_str
      END IF
   END IF


   '/ if string-1 is hex, convert to char and treat as type C

   IF cmd.txt_case_CTX_1 = "X" THEN

      hex_str = mapstr_X2A (txt_str1, hex_str_diag)

      IF ISNULL (hex_str) THEN                        '/ hex value is malformed
         cmd.txt_diag_str = hex_str_diag
         FUNCTION = 0                                            '/ parse error
         MExitFunc
      END IF

      cmd.txt_case_CTX_1 = "C"          '/ after hex conversion, type is like C
      txt_str1 = hex_str
   END IF



   FUNCTION = cp                   '/ point to end of string when parse is good

   MExitFunc


'/*****************************************************************************
'/*****************************************************************************


sub_get_string_item:

   work_str = ""
   work_CTX = " "
   quote = @cp                                        '/ remember opening quote

   DO
      INCR cp

      '/ if nul terminator reached, we passed end of string; end-quote missing

      IF @cp = 0 THEN
         cmd.txt_diag_str = _
            "Mapping string is missing close " & CHR$(quote) & " quote"
         FUNCTION = 0                                            '/ parse error
         MExitFunc
      END IF

      IF @cp = quote THEN                           '/ found closing-quote char
         INCR cp                 '/ do look ahead to see if doubled or hex mode

         '/ handle suffixed string: format is 'abcd'C or 'abcd'C or 'abcd'X
         '/ a type code cannot be followed by a letter
         '/ if no type code present, final quote cannot be followed by a letter

         IF @cp = ASC("C") _
         OR @cp = ASC("c") _
         OR @cp = ASC("T") _
         OR @cp = ASC("t") _
         OR @cp = ASC("X") _
         OR @cp = ASC("x") THEN
            work_CTX = UUCASE(CHR$(@cp))
            INCR cp                                           '/ skip over type

            IF mapstr_is_alpha_b (@cp) THEN
               cmd.txt_diag_str = _
                  "Mapping string quoted value not properly delimited: "

               '/ add some characters to the message, but stop doing it if we
               '/ reach the end of the line.  we do up to 9 chars of lookahead
               '/ to make the error msg more informative.

               FOR n = 0 TO 9
                  IF @cp[n] = 0 THEN EXIT FOR
                  cmd.txt_diag_str &= CHR$(@cp[n])
               NEXT

               FUNCTION = 0                                      '/ parse error
               MExitFunc
            END IF

            RETURN
         END IF

         IF @cp <> quote THEN          '/ quote is not doubled, string is ended

            IF mapstr_is_alpha_b (@cp) THEN
               cmd.txt_diag_str = _
                  "Mapping string quoted value not properly delimited: "

               '/ add some characters to the message, but stop doing it if we
               '/ reach the end of the line.  we do up to 9 chars of lookahead
               '/ to make the error msg more informative.

               FOR n = 0 TO 9
                  IF @cp[n] = 0 THEN EXIT FOR
                  cmd.txt_diag_str &= CHR$(@cp[n])
               NEXT

               FUNCTION = 0                                      '/ parse error
               MExitFunc
            END IF

            RETURN                                                  '/ parse OK
         END IF
      END IF

      work_str &= CHR$(@cp)                       '/ append a literal data char
   LOOP


   RETURN ' from sub_get_string_item


'/*****************************************************************************
'/*****************************************************************************


sub_get_string_replication:

   '/ grab the replication factor from a string that is not associated with a
   '/ command, and make 'n' copies of it.  to make this easy to understand,
   '/ a value like '*'2 will produce '**', so values of 0 and 1 will not
   '/ have any effect.  upon entry, @cp is pointing to the first digit.

   rep_num = 0

   DO WHILE mapstr_is_digit_b (@cp)
      rep_num = (rep_num * 10) + (@cp - ASC("0"))
      INCR cp
   LOOP

   IF rep_num > 1 THEN
      work_str = REPEAT$(rep_num, work_str)
   END IF


   RETURN ' from sub_get_string_replication


END FUNCTION                                                '/ mapstr_parse_txt



'/-----------------------------------------------------------------------------
'/ mapstr_parse_dxd_opcode()
'/
'/ 'dxd' is a generic term for DD, DX, XD or XX commands in general or
'/ extended format.  because general format dxd's can also have optional codes
'/ after them, they are all parsed here rather than elsewhere.
'/
'/ perform lookahead scan to see if extended-syntax versions of DX and XD
'/ commands are present.  if so, extract the relevant fields, including an
'/ indicator of implied scan direction.  a code like 1DX means LTR, while D1X
'/ means RTL.
'/
'/ if the current cp is not pointing to a valid dxd code, return 0
'/ otherwise return length of dxd prefix, which now is always 3.
'/
'/ we also parse SD and SX codes since they share the same syntax as dxd,
'/ except that the don't take a column reference and the are not scannable.
'/-----------------------------------------------------------------------------



FUNCTION                      mapstr_parse_dxd_opcode                         _
(  BYVAL cp                   AS BYTE PTR                                     _
,  BYREF cmd                  AS mapstr_cmd_t                                 _
   )                          AS LONG

   LOCAL u1                   AS STRING
   LOCAL u2                   AS STRING
   LOCAL u3                   AS STRING

   LOCAL u12                  AS STRING
   LOCAL u13                  AS STRING
   LOCAL u23                  AS STRING

   LOCAL dxd                  AS STRING
   LOCAL ordinal_str          AS STRING
   LOCAL width_str            AS STRING

   LOCAL c                    AS STRING
   LOCAL loc_cmd_text         AS STRING
   LOCAL loc_cmd_ID           AS LONG
   LOCAL loc_dxd_ordinal      AS LONG
   LOCAL dxdlen               AS LONG
   LOCAL n                    AS LONG

   MEntry


   u1 = UUCASE (CHR$(@cp[0]))
   u2 = UUCASE (CHR$(@cp[1]))
   u3 = UUCASE (CHR$(@cp[2]))

   u12 = u1 & u2
   u13 = u1 & u3
   u23 = u2 & u3

   cmd.cmd_rls_mode = %mapstr_rls_STD
   ordinal_str = "0"                               '/ non-scanning-mode ordinal


   '/ an extended dxd code must consist of a digit 1-9 followed by D or X
   '/ or a D or X followed by a digit.  if we don't see this prefix, it is
   '/ either a normal DX or XD or some other command.

   dxd = u1 & u2 & u3
   dxdlen = 3                                         '/ assume extended format


   IF u12 =                      "DD" THEN     '/ decimal to decimal processing
      loc_cmd_ID = %mapstr_cmd_ID_DD
      dxdlen = 2                                         '/ non-scanning format
      loc_cmd_text = u12                            '/ capture the command name

   ELSEIF u12 =                  "DX" THEN            '/ convert decimal to hex
      loc_cmd_ID = %mapstr_cmd_ID_DX
      dxdlen = 2                                         '/ non-scanning format
      loc_cmd_text = u12                            '/ capture the command name

   ELSEIF u12 =                  "XD" THEN            '/ convert hex to decimal
      loc_cmd_ID = %mapstr_cmd_ID_XD
      dxdlen = 2                                         '/ non-scanning format
      loc_cmd_text = u12                            '/ capture the command name

   ELSEIF u12 =                  "XX" THEN             '/ hex to hex processing
      loc_cmd_ID = %mapstr_cmd_ID_XX
      dxdlen = 2                                         '/ non-scanning format
      loc_cmd_text = u12                            '/ capture the command name

   ELSEIF u12 =                  "SD" THEN            '/ sequence number to dec
      loc_cmd_ID = %mapstr_cmd_ID_SD
      dxdlen = 2                                         '/ non-scanning format
      loc_cmd_text = u12                            '/ capture the command name

   ELSEIF u12 =                  "SX" THEN            '/ sequence number to hex
      loc_cmd_ID = %mapstr_cmd_ID_SX
      dxdlen = 2                                         '/ non-scanning format
      loc_cmd_text = u12                            '/ capture the command name


   '/ look for RTL scanning versions of command codes

   ELSEIF u13 =                  "DD" THEN                      '/ look for DnD
      loc_cmd_ID = %mapstr_cmd_ID_DD
      loc_cmd_text = u13
      ordinal_str = u2                              '/ presume u2 has the digit
      cmd.cmd_rls_mode = %mapstr_rls_RTL

   ELSEIF u13 =                  "DX" THEN                      '/ look for DnX
      loc_cmd_ID = %mapstr_cmd_ID_DX
      loc_cmd_text = u13
      ordinal_str = u2                              '/ presume u2 has the digit
      cmd.cmd_rls_mode = %mapstr_rls_RTL

   ELSEIF u13 =                  "XD" THEN                      '/ look for XnD
      loc_cmd_ID = %mapstr_cmd_ID_XD
      loc_cmd_text = u13
      ordinal_str = u2                              '/ presume u2 has the digit
      cmd.cmd_rls_mode = %mapstr_rls_RTL

   ELSEIF u13 =                  "XX" THEN                      '/ look for XnX
      loc_cmd_ID = %mapstr_cmd_ID_XX
      loc_cmd_text = u13
      ordinal_str = u2                              '/ presume u2 has the digit
      cmd.cmd_rls_mode = %mapstr_rls_RTL


   '/ look for LTR scanning versions of command codes

   ELSEIF u23 =                  "DD" THEN                      '/ look for nDD
      loc_cmd_ID = %mapstr_cmd_ID_DD
      loc_cmd_text = u23
      ordinal_str = u1                              '/ presume u1 has the digit

   ELSEIF u23 =                  "DX" THEN                      '/ look for nDX
      loc_cmd_ID = %mapstr_cmd_ID_DX
      loc_cmd_text = u23
      ordinal_str = u1                              '/ presume u1 has the digit

   ELSEIF u23 =                  "XD" THEN                      '/ look for nXD
      loc_cmd_ID = %mapstr_cmd_ID_XD
      loc_cmd_text = u23
      ordinal_str = u1                              '/ presume u1 has the digit

   ELSEIF u23 =                  "XX" THEN                      '/ look for nXX
      loc_cmd_ID = %mapstr_cmd_ID_XX
      loc_cmd_text = u23
      ordinal_str = u1                              '/ presume u1 has the digit


   ELSE
      FUNCTION = 0
      MExitFunc                                   '/ not a valid prefix for dxd

   END IF


   IF VERIFY(ordinal_str, "0123456789") <> VERIFIED THEN
      FUNCTION = 0
      MExitFunc                               '/ not a valid dxd ordinal number
   END IF

   '/ dxd not supposed to use explicit 0 as ordinal, but we'll accept it
   '/ ordinal 0 makes the dxd non-scanning mode

   loc_dxd_ordinal = VAL(ordinal_str)  '/ relative position of num within field

   IF loc_dxd_ordinal = 0 THEN
      cmd.cmd_rls_mode = %mapstr_rls_STD        '/ can't be RTL if not scanning
   END IF

   '/ we found a good dxd command

   cmd.cmd_ID      = loc_cmd_ID
   cmd.cmd_name    = loc_cmd_text
   cmd.dxd_ordinal = loc_dxd_ordinal


   '/--------------------------------------------------------------------------
   '/ look ahead and parse any additional DXD formatting codes
   '/ the returned length includes all of these codes
   '/
   '/ format syntax is:  [ [ w ] t ] [ S ]
   '/ 'w' is width number
   '/ 't' is type code "F" "Z" or "V"
   '/ "S" is sign specifier
   '/
   '/ command code may not have a format suffix.  that is not an error.
   '/ there may be 'nothing' or there may be just a column selection operand.
   '/ note that width/format and sign are independent operands.
   '/--------------------------------------------------------------------------

   '/ see if possible width is present
   '/ example: 1DX2F
   '/          0123

   n = dxdlen
   width_str = ""

   DO WHILE VERIFY (CHR$(@cp[n]), "0123456789") = VERIFIED
      width_str &=  CHR$(@cp[n])
      INCR n
   LOOP

   c = UUCASE (CHR$(@cp[n]))

   IF c = "F" _
   OR c = "Z" _
   OR c = "V" THEN                            '/ found a valid format type code

      IF ISNOTNULL(width_str) THEN
         cmd.dxd_fmt_width_used = 1
         cmd.dxd_fmt_width = VAL (width_str)
      END IF

      cmd.dxd_fmt_type = c
      INCR n                                             '/ skip over type code
      dxdlen = n                           '/ commit to this much of the format

      c = UUCASE (CHR$(@cp[n]))
   END IF

   IF c = "S" _                                       '/ found sign code S or C
   OR c = "C" THEN
      cmd.dxd_sign_code = c
      INCR n                                             '/ skip over sign code
      dxdlen = n                           '/ commit to this much of the format
   END IF

   IF cmd.dxd_fmt_type = " " THEN                     '/ if format code omitted
      cmd.dxd_fmt_type = "V"                  '/ we will assume variable length
   END IF

   FUNCTION = dxdlen                              '/ length of found dxd prefix

   MExit

END FUNCTION                                         '/ mapstr_parse_dxd_opcode



'/-----------------------------------------------------------------------------
'/ mapstr_parse_opcode()
'/
'/ parse and extract mapping cmd codes, which are alpha codes like L, TR and AX
'/ after detecting the cmd code, we do a lookahead to confirm that the cmd code
'/ was used in the correct context and any following tokens are valid, and are
'/ present if required by the particular cmd code found.
'/-----------------------------------------------------------------------------

FUNCTION                      mapstr_parse_opcode                             _
(  BYVAL cp                   AS BYTE PTR                                     _
,  BYREF cmd                  AS mapstr_cmd_t                                 _
,  BYVAL len_result_str       AS LONG                                         _
   )                          AS RETURNED_BYTE_PTR

   LOCAL cid                  AS STRING

   LOCAL c                    AS STRING
   LOCAL u1                   AS STRING
   LOCAL u12                  AS STRING
   LOCAL u123                 AS STRING

   LOCAL fake_str1            AS STRING
   LOCAL fake_str2            AS STRING
   LOCAL fake_cmd             AS mapstr_cmd_t

   LOCAL optlen               AS LONG
   LOCAL dxdlen               AS LONG

   '/ flags

   LOCAL has_number           AS BYTE           '/ if number operand is present
   LOCAL plain_cmd_only       AS BYTE                  '/ if plain cmd_ID is OK
   LOCAL dot_notation_allowed AS BYTE             '/ if dot notation valid here
   LOCAL copying_cmd          AS BYTE                     '/ describes cmd type
   LOCAL modifying_cmd        AS BYTE                     '/ describes cmd type
   LOCAL need_str1            AS BYTE        '/ if string operand 1 is required

   LOCAL cmd_found            AS BYTE         '/ if string operand 2 is allowed


   LOCAL lap                  AS BYTE PTR                 '/ look-ahead pointer

   MEntry


   cmd.cmd_ID = %mapstr_none           '///////// --- is this needed ??? \\\\\\
   cmd.cmd_type = %mapstr_other_cmd         '/ if not modifying or copying type

   cmd.dxd_fmt_type  = " "
   cmd.dxd_sign_code = " "

   has_number     = 0
   need_str1      = 0
   copying_cmd    = 0
   modifying_cmd  = 0
   plain_cmd_only = 0

   dot_notation_allowed = 0


   u1   =       UUCASE (CHR$ (@cp))
   u12  = u1  & UUCASE (CHR$ (@cp[1]))
   u123 = u12 & UUCASE (CHR$ (@cp[2]))               '/ for long pivot commands



   '/ check first if there is a dxd opcode which has a different syntax

   dxdlen = mapstr_parse_dxd_opcode (cp, cmd)

   IF dxdlen > 0 THEN                                     '/ extended dxd found
      optlen = dxdlen                                '/ it has a varying length
      cmd_found = 1
      cid = cmd.cmd_name

      IF cmd.cmd_ID = %mapstr_cmd_ID_SD THEN
         cmd.cmd_type = %mapstr_other_cmd

      ELSEIF cmd.cmd_ID = %mapstr_cmd_ID_SX THEN
         cmd.cmd_type = %mapstr_other_cmd

      ELSE
         cmd.cmd_type = %mapstr_copying_cmd
         cmd.takes_col_ref = 1      '/ the number operand is a column reference

      END IF

   ELSE
      cmd_found = 0
   END IF



   '/ look for 3-character commands

   IF cmd_found = 0 THEN
      cmd_found = 1
      optlen = 3                                     '/ look for 3 letter codes
      cid = u123
      cmd.cmd_name = cid                 '/ preliminary value, used in messages

   SELECT CASE AS CONST$ u123

      '/ LTR version of PA

      CASE                          "PAL"              '/ special pivot command
         cmd.cmd_ID = %mapstr_cmd_ID_PA              '/ STRING 2 CAN BE OMITTED
         cmd.cmd_rls_mode = %mapstr_rls_LTR
         cmd.cmd_type = %mapstr_modifying_cmd
         optlen = 3                                    '/ special 3 letter code
         cmd.cmd_name = u123                                '/ save longer name
         '/ pivot/boundary string is optional          '/ *** need_str1 = 1 ***
         cmd.uses_str2 = %mapstr_str2_allowed       '/ CAN use str-pair operand
         cmd.takes_col_ref = 1      '/ the number operand is a column reference

      '/ LTR version of PD

      CASE                          "PDL"              '/ special pivot command
         cmd.cmd_ID = %mapstr_cmd_ID_PD              '/ STRING 2 CAN BE OMITTED
         cmd.cmd_rls_mode = %mapstr_rls_LTR
         cmd.cmd_type = %mapstr_modifying_cmd
         optlen = 3                                    '/ special 3 letter code
         cmd.cmd_name = u123                                '/ save longer name
         '/ pivot/boundary string is optional          '/ *** need_str1 = 1 ***
         cmd.uses_str2 = %mapstr_str2_allowed       '/ CAN use str-pair operand
         cmd.takes_col_ref = 1      '/ the number operand is a column reference

      '/ RTL version of PA

      CASE                          "PAR"              '/ special pivot command
         cmd.cmd_ID = %mapstr_cmd_ID_PA              '/ STRING 2 CAN BE OMITTED
         cmd.cmd_rls_mode = %mapstr_rls_RTL
         cmd.cmd_type = %mapstr_modifying_cmd
         optlen = 3                                    '/ special 3 letter code
         cmd.cmd_name = u123                                '/ save longer name
         '/ pivot/boundary string is optional          '/ *** need_str1 = 1 ***
         cmd.uses_str2 = %mapstr_str2_allowed       '/ CAN use str-pair operand
         cmd.takes_col_ref = 1      '/ the number operand is a column reference

      '/ RTL version of PD

      CASE                          "PDR"              '/ special pivot command
         cmd.cmd_ID = %mapstr_cmd_ID_PD              '/ STRING 2 CAN BE OMITTED
         cmd.cmd_rls_mode = %mapstr_rls_RTL
         cmd.cmd_type = %mapstr_modifying_cmd
         optlen = 3                                    '/ special 3 letter code
         cmd.cmd_name = u123                                '/ save longer name
         '/ pivot/boundary string is optional          '/ *** need_str1 = 1 ***
         cmd.uses_str2 = %mapstr_str2_allowed       '/ CAN use str-pair operand
         cmd.takes_col_ref = 1      '/ the number operand is a column reference

      CASE ELSE
         cmd_found = 0

   END SELECT
   END IF



   '/ look for 2-character commands

   IF cmd_found = 0 THEN
      cmd_found = 1
      optlen = 2                                     '/ look for 2 letter codes
      cid = u12
      cmd.cmd_name = cid                 '/ preliminary value, used in messages

   SELECT CASE AS CONST$ u12

      CASE                          "AQ"                '/ accent quote literal
         cmd.cmd_ID = %mapstr_cmd_ID_AQ

      CASE                          "AX"                '/ convert ascii to hex
         cmd.cmd_ID = %mapstr_cmd_ID_AX           '/ hex digits rendered as A-F
         cmd.cmd_type = %mapstr_copying_cmd
         cmd.takes_col_ref = 1      '/ the number operand is a column reference

      CASE                          "CC"                         '/ CASE C mode
         cmd.cmd_ID = %mapstr_cmd_ID_CC
         plain_cmd_only = 1                '/ string and number must be omitted

      CASE                          "CT"                         '/ CASE T mode
         cmd.cmd_ID = %mapstr_cmd_ID_CT
         plain_cmd_only = 1                '/ string and number must be omitted

      CASE                          "DC"             '/ delete character values
         cmd.cmd_ID = %mapstr_cmd_ID_DC
         cmd.cmd_type = %mapstr_modifying_cmd
         need_str1 = 1
         cmd.takes_col_ref = 1      '/ the number operand is a column reference

      CASE                          "DQ"                '/ double quote literal
         cmd.cmd_ID = %mapstr_cmd_ID_DQ

      CASE                          "DV"                   '/ include Dyn Value
         cmd.cmd_ID = %mapstr_cmd_ID_DV
         cmd.cmd_type = %mapstr_modifying_cmd
         cmd.takes_col_ref = 1      '/ the number operand is a column reference
         '/ DV defaults to delete option D if there is no string.


  '/ environment variables with embedded = signs in their value proved to be
  '/ unreliable, so support for this feature has been disabled.

   '/ CASE                          "EV"                '/ include Env Variable
   '/    cmd.cmd_ID = %mapstr_cmd_ID_EV
   '/    cmd.cmd_type = %mapstr_other_cmd
   '/    need_str1 = 1

      CASE                          "EX"               '/ convert ebcdic to hex
         cmd.cmd_ID = %mapstr_cmd_ID_EX           '/ hex digits rendered as A-F
         cmd.cmd_type = %mapstr_copying_cmd
         cmd.takes_col_ref = 1      '/ the number operand is a column reference

      CASE                          "IC"                '/ alter to invert case
         cmd.cmd_ID = %mapstr_cmd_ID_IC
         cmd.cmd_type = %mapstr_modifying_cmd
         cmd.takes_col_ref = 1      '/ the number operand is a column reference

      CASE                          "LC"                 '/ alter to lower case
         cmd.cmd_ID = %mapstr_cmd_ID_LC
         cmd.cmd_type = %mapstr_modifying_cmd
         cmd.takes_col_ref = 1      '/ the number operand is a column reference

      CASE                          "PA"       '/ standard Pivot data and Align
         cmd.cmd_ID = %mapstr_cmd_ID_PA
         cmd.cmd_rls_mode = %mapstr_rls_STD
         cmd.cmd_type = %mapstr_modifying_cmd
         '/ pivot/boundary string is optional          '/ *** need_str1 = 1 ***
         cmd.uses_str2 = %mapstr_str2_allowed       '/ CAN use str-pair operand
         cmd.takes_col_ref = 1      '/ the number operand is a column reference

      CASE                          "PD"       '/ standard Pivot data and Align
         cmd.cmd_ID = %mapstr_cmd_ID_PD
         cmd.cmd_rls_mode = %mapstr_rls_STD
         cmd.cmd_type = %mapstr_modifying_cmd
         '/ pivot/boundary string is optional          '/ *** need_str1 = 1 ***
         cmd.uses_str2 = %mapstr_str2_allowed       '/ CAN use str-pair operand
         cmd.takes_col_ref = 1      '/ the number operand is a column reference

      CASE                          "PL"                         '/ pad on left
         cmd.cmd_ID = %mapstr_cmd_ID_PL
         cmd.cmd_type = %mapstr_modifying_cmd

      CASE                          "PR"                        '/ pad on right
         cmd.cmd_ID = %mapstr_cmd_ID_PR
         cmd.cmd_type = %mapstr_modifying_cmd

      CASE                          "QD"     '/ quit use of Dyn Value - private
         cmd.cmd_ID = %mapstr_cmd_ID_QD
         cmd.cmd_type = %mapstr_other_cmd

      CASE                          "QE"  '/ quit use of Env Variable - private
         cmd.cmd_ID = %mapstr_cmd_ID_QE
         cmd.cmd_type = %mapstr_other_cmd
         need_str1 = 1

      CASE                          "QS"  '/ quit use of Set Variable - private
         cmd.cmd_ID = %mapstr_cmd_ID_QS
         cmd.cmd_type = %mapstr_other_cmd
         need_str1 = 1

      CASE                          "RA"                         '/ Replace All
         cmd.cmd_ID = %mapstr_cmd_ID_RA              '/ STRING 2 CAN BE OMITTED
         cmd.cmd_type = %mapstr_modifying_cmd
         cmd.cmd_type = %mapstr_modifying_cmd
         need_str1 = 1
         cmd.uses_str2 = %mapstr_str2_allowed       '/ CAN use str-pair operand
         cmd.takes_col_ref = 1      '/ the number operand is a column reference

      CASE                          "RC"          '/ character replacement item
         cmd.cmd_ID = %mapstr_cmd_ID_RC
         cmd.cmd_type = %mapstr_modifying_cmd
         need_str1 = 1
         cmd.uses_str2 = %mapstr_str2_required     '/ MUST use str-pair operand
         cmd.takes_col_ref = 1      '/ the number operand is a column reference


      CASE                          "RL"                    '/ Replace Leftmost
         cmd.cmd_ID = %mapstr_cmd_ID_RL              '/ STRING 2 CAN BE OMITTED
         cmd.cmd_type = %mapstr_modifying_cmd
         cmd.cmd_type = %mapstr_modifying_cmd
         need_str1 = 1
         cmd.uses_str2 = %mapstr_str2_allowed       '/ CAN use str-pair operand
         cmd.takes_col_ref = 1      '/ the number operand is a column reference

      CASE                          "RP"            '/ repeat the result string
         cmd.cmd_ID = %mapstr_cmd_ID_RP
         cmd.cmd_type = %mapstr_modifying_cmd

      CASE                          "RR"                   '/ Replace Rightmost
         cmd.cmd_ID = %mapstr_cmd_ID_RR              '/ STRING 2 CAN BE OMITTED
         cmd.cmd_type = %mapstr_modifying_cmd
         cmd.cmd_type = %mapstr_modifying_cmd
         need_str1 = 1
         cmd.uses_str2 = %mapstr_str2_allowed       '/ CAN use str-pair operand
         cmd.takes_col_ref = 1      '/ the number operand is a column reference

      CASE                          "RV"           '/ reverse the result string
         cmd.cmd_ID = %mapstr_cmd_ID_RV
         cmd.cmd_type = %mapstr_modifying_cmd
         cmd.takes_col_ref = 1      '/ the number operand is a column reference

      CASE                          "SC"              '/ alter to sentence case
         cmd.cmd_ID = %mapstr_cmd_ID_SC
         cmd.cmd_type = %mapstr_modifying_cmd
         cmd.takes_col_ref = 1      '/ the number operand is a column reference

      CASE                          "SQ"                '/ single quote literal
         cmd.cmd_ID = %mapstr_cmd_ID_SQ

      CASE                          "SV"                '/ include Set Variable
         cmd.cmd_ID = %mapstr_cmd_ID_SV
         cmd.cmd_type = %mapstr_other_cmd
         need_str1 = 1

      CASE                          "TC"                 '/ alter to title case
         cmd.cmd_ID = %mapstr_cmd_ID_TC
         cmd.cmd_type = %mapstr_modifying_cmd
         cmd.takes_col_ref = 1      '/ the number operand is a column reference

      CASE                          "TL"                        '/ trim on left
         cmd.cmd_ID = %mapstr_cmd_ID_TL
         cmd.cmd_type = %mapstr_modifying_cmd

      CASE                          "TR"                       '/ trim on right
         cmd.cmd_ID = %mapstr_cmd_ID_TR
         cmd.cmd_type = %mapstr_modifying_cmd

      CASE                          "UC"                 '/ alter to upper case
         cmd.cmd_ID = %mapstr_cmd_ID_UC

         cmd.cmd_type = %mapstr_modifying_cmd
         cmd.takes_col_ref = 1      '/ the number operand is a column reference

      CASE                          "XA"                '/ convert hex to ascii
         cmd.cmd_ID = %mapstr_cmd_ID_XA
         cmd.cmd_type = %mapstr_copying_cmd
         cmd.takes_col_ref = 1      '/ the number operand is a column reference

      CASE                          "XE"               '/ convert hex to ebcdic
         cmd.cmd_ID = %mapstr_cmd_ID_XE
         cmd.cmd_type = %mapstr_copying_cmd
         cmd.takes_col_ref = 1      '/ the number operand is a column reference


      CASE ELSE
         cmd_found = 0

   END SELECT
   END IF



   '/ look for 1-character commands

   IF cmd_found = 0 THEN
      cmd_found = 1
      optlen = 1                                     '/ look for 1 letter codes
      cid = u1
      cmd.cmd_name = cid                 '/ preliminary value, used in messages

   SELECT CASE AS CONST$ u1

      CASE                          "C"                          '/ center text
         cmd.cmd_ID = %mapstr_cmd_ID_C
         cmd.cmd_type = %mapstr_modifying_cmd

      CASE                          "D"            '/ delete result string text
         cmd.cmd_ID = %mapstr_cmd_ID_D
         cmd.cmd_type = %mapstr_modifying_cmd
         cmd.takes_col_ref = 1      '/ the number operand is a column reference
         cmd.suppress_auto_ref = 1  '/ this command cannot auto-ref; it's "bad"

      CASE                          "L"                         '/ left justify
         cmd.cmd_ID = %mapstr_cmd_ID_L
         cmd.cmd_type = %mapstr_modifying_cmd

      CASE                          "M"        '/ move characters within result
         cmd.cmd_ID = %mapstr_cmd_ID_M
         cmd.cmd_type = %mapstr_modifying_cmd
         cmd.takes_col_ref = 1      '/ the number operand is a column reference
         cmd.suppress_auto_ref = 1  '/ this command cannot auto-ref; it's "bad"

      CASE                          "P"                '/ pad on left and right
         cmd.cmd_ID = %mapstr_cmd_ID_P
         cmd.cmd_type = %mapstr_modifying_cmd

      CASE                          "R"                        '/ right justify
         cmd.cmd_ID = %mapstr_cmd_ID_R
         cmd.cmd_type = %mapstr_modifying_cmd

      CASE                          "T"               '/ trim on left and right
         cmd.cmd_ID = %mapstr_cmd_ID_T
         cmd.cmd_type = %mapstr_modifying_cmd

      CASE                          "X"                     '/ exchange strings
         cmd.cmd_ID = %mapstr_cmd_ID_X
         cmd.cmd_type = %mapstr_copying_cmd
         need_str1 = 1
         cmd.uses_str2 = %mapstr_str2_required     '/ MUST use str-pair operand
         cmd.takes_col_ref = 1      '/ the number operand is a column reference

      CASE                          "Z"                     '/ zero suppression
         cmd.cmd_ID = %mapstr_cmd_ID_Z
         cmd.cmd_type = %mapstr_modifying_cmd

      CASE ELSE
         cmd_found = 0

   END SELECT
   END IF



   IF cmd_found = 0 THEN
      cmd.cmd_ID = %mapstr_none
      cmd.cmd_diag_str = "Mapping command code at " & TRIM$(u12) & " undefined"
      FUNCTION = 0                                  '/ cmd_ID code is undefined
      MExitFunc
   END IF


   '/ skip past 1 to 3 letters depending on code length, then check next char
   '/ dxd-type commands skip 3 characters

   cp += optlen


   '/--------------------------------------------------------------------------
   '/ a command code must be followed by a delimiter, or number, or string,
   '/ that is, either the command is ended, or there are operands.
   '/ if the code is followed by a letter, the command is undefined.
   '/--------------------------------------------------------------------------

   IF mapstr_is_alpha_b (@cp) THEN
      cmd.cmd_ID = %mapstr_none
      cmd.cmd_diag_str = "Mapping command code at " & _
         cid & c & " undefined"

      FUNCTION = 0                                  '/ cmd_ID code is undefined
      MExitFunc

   END IF


   '/--------------------------------------------------------------------------
   '/ most commands are either copying mode or modifying mode
   '/ if they are, a dot command is allowed.  the few that are not, like CT
   '/ are neither one, and don't take a dot command.
   '/ but if this is a modifying mode, dot is assumed even if not actually used
   '/
   '/ modifying commands always use the result string, an copying commands use
   '/ the result string if they have dot notation.  that part we determine
   '/ later on.
   '/--------------------------------------------------------------------------

   IF cmd.cmd_type = %mapstr_modifying_cmd THEN
      cmd.use_result_str = 1
   END IF

   IF cmd.cmd_type = %mapstr_modifying_cmd _
   OR cmd.cmd_type = %mapstr_copying_cmd THEN
      dot_notation_allowed = 1
   END IF


   '/--------------------------------------------------------------------------
   '/ we may have to skip over a string value.  to handle this, we leave the
   '/ cp where it is, and continue on with the look ahead pointer
   '/ all command codes allow string operands, and none require them
   '/ if a string is present, we (redundantly) parse it, just to skip over to
   '/ the end of it.  then we check for a number if present and required
   '/
   '/ if command code can end right here, and there's a good delimiter, we
   '/ are done at this point
   '/--------------------------------------------------------------------------

   lap = cp

   '/--------------------------------------------------------------------------
   '/ if the cmd_ID is one that can appear by itself ("plain") like T, we
   '/ check if it has a valid delimiter, which can be blank/comma or a case
   '/ conversion code.
   '/
   '/ allowing a \ reversal code to end a plain cmd_ID may or may not be the
   '/ correct thing to do.  once testing is underway, this may need to be
   '/ revisited to confirm this is the right thing here.
   '/--------------------------------------------------------------------------

   IF plain_cmd_only = 1 THEN
      IF mapstr_is_delim (CHR$ (@lap)) _
      OR @lap = ASC("\") _                 '/ not sure if allowing "\" is valid
      OR @lap = ASC("<") _
      OR @lap = ASC(">") THEN

         FUNCTION = cp
         MExitFunc

      END IF
   END IF


   IF mapstr_is_quote_b (@lap) THEN

      '/ we want a realistic 'cmd' structure, but don't want the real one to be
      '/ modified right now.  so, just copy the real one to the fake one.

      fake_cmd = cmd

      lap = mapstr_parse_txt (lap, fake_str1, fake_str2, fake_cmd)

      IF lap = 0 THEN

         IF ISNULL(fake_cmd.txt_diag_str) THEN     '/ don't know why str is bad
            cmd.cmd_diag_str = _
               "Mapping command " & cid & " string operand malformed"

         ELSE                                           '/ parser gave a reason
            cmd.cmd_diag_str = fake_cmd.txt_diag_str
         END IF

         FUNCTION = 0                             '/ string not properly closed
         MExitFunc

      END IF

   '/--------------------------------------------------------------------------
   '/ if command demands a string operand and it's missing, report the error
   '/ if command can use a string but it's optional, or command doesn't need a
   '/ string operand at all, but one is present, we will do a look-ahead parse
   '/ to make sure the string appears valid.  it may or may not get used later.
   '/--------------------------------------------------------------------------

   ELSE                                        '/ cmd_ID not followed by string
      IF need_str1 = 1 THEN                                 '/  but we need one
         cmd.cmd_diag_str = "Mapping command " & cid & " needs string operand"
         FUNCTION = 0                             '/ required string is missing
         MExitFunc

      END IF
   END IF


   '/--------------------------------------------------------------------------
   '/ if command uses a column reference and a + appears, we will allow it
   '/ as a shortcut for 1+.  there has to be a good delim after the +
   '/--------------------------------------------------------------------------

   IF  @lap = ASC("+") _
   AND cmd.takes_col_ref _
   AND mapstr_is_delim (CHR$ (@lap[1])) = 1  THEN
      INCR lap                                                '/ skip over plus
      FUNCTION = cp                         '/ successfully parsed command code
      MExitFunc
   END IF


   IF @lap = ASC(".") THEN
      IF dot_notation_allowed = 1 THEN
         INCR lap                                              '/ skip over dot

         '/ dot can by followed by blank/delim or number field.  if those are
         '/ not present, they forgot the number or the map string is bad

         IF  mapstr_is_delim (CHR$ (@lap)) = 0 _
         AND VERIFY(CHR$(@lap), "-0123456789") <> VERIFIED THEN
            cmd.cmd_diag_str = "Mapping command " & cid & " " & CHR$ (@lap) & _
               " . notation used improperly"
            FUNCTION = 0                          '/ required number is missing
            MExitFunc

         ELSE
            has_number = 1                      '/ an explicit number was found
            cmd.use_result_str = 1   '/ dot notation means we use result string
         END IF


      '/ there are a few commands that cannot use dot notation, like CC

      ELSE
         cmd.cmd_diag_str = "Mapping command " & cid & " cannot use . notation"
         FUNCTION = 0                             '/ required number is missing
         MExitFunc

      END IF
   END IF

   FUNCTION = cp                            '/ successfully parsed command code

   MExitFunc

END FUNCTION                                             '/ mapstr_parse_opcode



'/*****************************************************************************
'/****   -----------------------------   **************************************
'/****   MAIN LOGIC OF MAPPING STRINGS   **************************************
'/****   -----------------------------   **************************************
'/*****************************************************************************
'/
'/    FUNCTION                      mapstr_execute
'/    (  BYVAL source_arg           AS STRING
'/    ,  BYVAL map_arg              AS STRING
'/    ,  BYREF result_arg           AS STRING
'/    ,  BYREF err_arg              AS STRING
'/    ,  BYVAL case_CT_arg          AS STRING
'/    ,  BYVAL tran_A2E_arg         AS BYTE PTR
'/    ,  BYVAL tran_E2A_arg         AS BYTE PTR
'/    ,  BYVAL sequence_arg         AS QUAD
'/    ,  BYVAL diag_mode            AS LONG
'/       )                          AS LONG               '/ --> AS RETURN CODE
'/
'/
'/ This is the mapping-string function to be called by SPFLite
'/
'/
'/ Arguments:
'/
'/    source_arg     String DATA as found by CHANGE command string-1 operand
'/
'/    map_arg        String-2 OPERAND as parsed by CHANGE command
'/                      M'xxx'  or  'xxx'M  is reduced to just the  xxx  value
'/
'/    result_arg     RETURNED: contents of source_arg string after reformatting
'/
'/    err_arg        RETURNED: error message, or "" if mapping was successful
'/
'/    case_CT_arg    Current state of PROFILE CASE setting, either "C" or "T"
'/                   If CASE unavailable, caller should pass " "
'/
'/    tran_A2E_arg   BYTE PTR to 256-byte Ansi to EBCDIC translation table
'/
'/    tran_E2A_arg   BYTE PTR to 256-byte EBCDIC to Ansi translation table
'/
'/    sequence_arg   sequential number for SD and SX commands; 0 = DIAG
'/
'/    diag_mode      1 = DIAG mode, 0 - normal mode
'/
'/ Return code:
'/
'/    0        OK    - Mapping was successful
'/    8        ERROR - Mapping string is malformed or other error detected
'/
'/ Notes:      1. On return code 0, caller should check whether returned
'/                source_arg = result_arg.
'/
'/             2. If supplied string is not changed from original value,
'/                update of file could be skipped to save time.
'/
'/             3. If translation tables are unavailable, caller should
'/                pass null pointers (0).  Either both pointers should be
'/                valid, or both null; do not mix null with non-null.
'/
'/             4. When edit file is Ansi or Unicode, translation table pointers
'/                should point to the default EBCDIC 1140 code page tables.
'/                If it is not convenient to do that, pass null pointers.
'/
'/                When edit file is explicitly EBCDIC or is a user-provided
'/                code page, those code pages should be supplied.
'/
'/                For example, if a user defined a code page of MYTABLE,
'/                the pointers for the translation tables should be:
'/
'/                   tran_A2E_arg = Ansi to MYTABLE translation table
'/                   tran_E2A_arg = MYTABLE to Ansi translation table
'/
'/*****************************************************************************
'/*****************************************************************************


FUNCTION                      mapstr_execute                                  _
(  BYVAL source_arg           AS STRING                                       _
,  BYVAL map_arg              AS STRING                                       _
,  BYREF result_arg           AS STRING                                       _
,  BYREF err_arg              AS STRING                                       _
,  BYVAL case_CT_arg          AS STRING                                       _
,  BYVAL tran_A2E_arg         AS BYTE PTR                                     _
,  BYVAL tran_E2A_arg         AS BYTE PTR                                     _
,  BYVAL sequence_arg         AS QUAD                                         _
,  BYVAL diag_mode            AS LONG                                         _
   )                          AS LONG                     '/ --> AS RETURN CODE


   LOCAL sequence_num         AS QUAD

   LOCAL calc_first_num       AS QUAD                '/ for call to mapstr_calc
   LOCAL calc_final_num       AS QUAD                '/ for call to mapstr_calc
   LOCAL calc_line_arg        AS LONG
   LOCAL calc_col_arg         AS LONG

   LOCAL result_str           AS STRING

   '/ if dot notation is used, the result_str is selected
   '/ if dot notation is not used, the source_arg is selected
   '/ the VARPTR function is used to set the appropriate address

   LOCAL select_str           AS STRING PTR

   LOCAL temp_result          AS STRING
   LOCAL inp_len              AS LONG
   LOCAL map_str              AS STRING
   LOCAL map_char             AS STRING

   LOCAL cp                   AS BYTE PTR
   LOCAL c                    AS STRING
   LOCAL u1                   AS STRING
   LOCAL u2                   AS STRING
   LOCAL u12                  AS STRING

   LOCAL inp_dim              AS LONG             '/ max possible mapping items
   LOCAL inp_ndx              AS LONG
   LOCAL retcode              AS LONG

   LOCAL parse_str1           AS STRING
   LOCAL parse_str2           AS STRING

   LOCAL move_text            AS STRING

   LOCAL scan_text            AS STRING
   LOCAL scan_data            AS STRING
   LOCAL scan_str1            AS STRING
   LOCAL scan_str2            AS STRING
   LOCAL scan_result          AS STRING
   LOCAL scan_prefix          AS STRING
   LOCAL scan_suffix          AS STRING
   LOCAL scan_infix           AS STRING
   LOCAL scan_len1            AS LONG
   LOCAL scan_len2            AS LONG
   LOCAL scan_last            AS LONG
   LOCAL scan_limit           AS LONG

   LOCAL move_text_valid      AS LONG
   LOCAL move_text_replaced   AS LONG

   LOCAL hex_result           AS STRING
   LOCAL hex_text             AS STRING
   LOCAL hex_temp             AS STRING
   LOCAL hex_diag             AS STRING
   LOCAL hex_n                AS LONG
   LOCAL hex_c                AS STRING
   LOCAL hex_xx               AS STRING
   LOCAL hex_c1               AS STRING
   LOCAL hex_c2               AS STRING
   LOCAL range_len            AS LONG                '/ for [x-y] range strings


   LOCAL cmd                  AS mapstr_cmd_t          '/ parsing control block


   LOCAL dxd_source_type      AS BYTE
   LOCAL dxd_result_type      AS BYTE

   LOCAL dxd_sign             AS STRING
   LOCAL dxd_source           AS STRING
   LOCAL dxd_result           AS STRING
   LOCAL dxd_diag             AS STRING
   LOCAL dxd_rc               AS LONG


   LOCAL reverse_mode         AS BYTE
   LOCAL reverse_err          AS BYTE
   LOCAL reverse_lookahead    AS STRING
   LOCAL string_err           AS BYTE
   LOCAL mapstr_case          AS BYTE
   LOCAL hexstr_case          AS BYTE

   LOCAL curr_case_CT         AS STRING
   LOCAL this_case_CT         AS STRING
   LOCAL fake_case_CT         AS STRING


   LOCAL col_left             AS LONG
   LOCAL col_right            AS LONG
   LOCAL col_first            AS LONG
   LOCAL col_final            AS LONG
   LOCAL col_step             AS LONG
   LOCAL col_max              AS LONG

   LOCAL n                    AS LONG

   LOCAL pad_len              AS LONG
   LOCAL pad_len1             AS LONG
   LOCAL pad_len2             AS LONG
   LOCAL pad_count            AS LONG
   LOCAL pad_str              AS STRING
   LOCAL pad_str1             AS STRING
   LOCAL pad_str2             AS STRING
   LOCAL cmd_str1             AS STRING
   LOCAL cmd_str2             AS STRING

   LOCAL del_len              AS LONG

   LOCAL trim_len             AS LONG
   LOCAL trim_max             AS LONG
   LOCAL trim_len1            AS LONG
   LOCAL trim_len2            AS LONG
   LOCAL trim_pos1            AS LONG
   LOCAL trim_pos2            AS LONG
   LOCAL trim_count           AS LONG

   LOCAL txt_middle           AS LONG
   LOCAL txt_before           AS STRING
   LOCAL txt_after            AS STRING

   LOCAL xc_txt_before        AS STRING
   LOCAL xc_txt_after         AS STRING
   LOCAL xc_txt_result        AS STRING

   LOCAL usage_flag()         AS BYTE

   LOCAL byte_tab()           AS BYTE     '/ PowerBasic requires this LOCAL/DIM
   DIM   byte_tab(0 TO 255)   AS BYTE       '/ pairing; we can't use just LOCAL

   LOCAL dot_notation_used    AS BYTE

   '/ data for M command

   LOCAL m_cmd_cp             AS BYTE PTR

   LOCAL m_cmd_dest_str       AS STRING

   LOCAL m_cmd                AS mapstr_cmd_t          '/ parsing control block

   LOCAL m_cmd_move_mode      AS LONG
   LOCAL dst_cols_valid       AS LONG

   LOCAL dst_first            AS LONG
   LOCAL dst_final            AS LONG
   LOCAL dst_step             AS LONG
   LOCAL dst_left             AS LONG
   LOCAL dst_right            AS LONG

   LOCAL dst_min              AS LONG
   LOCAL dst_max              AS LONG

   LOCAL m_txt_seg_1          AS STRING
   LOCAL m_txt_seg_2          AS STRING
   LOCAL m_txt_seg_3          AS STRING
   LOCAL m_txt_seg_4          AS STRING
   LOCAL m_txt_seg_5          AS STRING

   '/ data for DV/EV/SV QD/QE/QS commands

   LOCAL esv_name             AS STRING               '/ prefixed EV or SV name
   LOCAL esv_quit             AS STRING                '/ QE or QS quit command
   LOCAL esv_text             AS STRING         '/ text from SET var or ENV var
   LOCAL esv_tail             AS STRING          '/ tail end of current command
   LOCAL esv_cmd              AS STRING         '/ new command build from EV/SV

   LOCAL esv_mode             AS BYTE
   LOCAL esv_state            AS BYTE                   '/ %esv_state_beg / end

   LOCAL esv_nest_level       AS LONG               '/ number of slots occupied
   DIM   esv_nest_name (%mapstr_esv_max) AS STRING

   LOCAL esv_get_rc           AS LONG
                                   '/ %mapstr_retcode_OK or %mapstr_retcode_ERR



'/*****************************************************************************
'/****  --------------  *******************************************************
'/****  mapstr_execute  *******************************************************
'/****  --------------  *******************************************************
'/*****************************************************************************

   MEntry

   result_arg = source_arg          '/ default output is unchanged input string
   err_arg = ""

   IF sequence_arg = %mapstr_DIAG THEN                  '/ a DIAG call was made
      sequence_num = 1
   ELSE
      sequence_num = sequence_arg
   END IF

   '/ FROM THIS POINT ON, ONLY USE sequence_num

   '/--------------------------------------------------------------------------
   '/ we should never get a null source string, but we can't proceed if we did
   '/--------------------------------------------------------------------------

   '/ IF ISNULL(source_arg) THEN
   '/    err_arg = "Mapping source string is empty, internal error"
   '/    FUNCTION = %mapstr_retcode_ERR
   '/    MExitFunc
   '/ END IF


   '/--------------------------------------------------------------------------
   '/ if map_arg is blank, return data_arg unchanged
   '/ map should not be blank, but syntax of command does not prevent this
   '/ this would be considered unusual usage
   '/
   '/ a blank map implies the output is null, but copying the input would be
   '/ the least surprising result
   '/--------------------------------------------------------------------------

   IF ISNULL(map_arg)  _
   OR VERIFY(map_arg, " ") = VERIFIED THEN
      err_arg = "Mapping string is empty, no CHANGE performed"
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc
   END IF


   dot_notation_used = 0
   select_str = VARPTR (source_arg)   '/ assumed unless/until dot notation used

   result_str   = ""
   retcode      = 0
   mapstr_case  = %mapstr_none
   reverse_mode = 0

   '/ normalize the profile CASE C or CASE T in case the value not available

   case_CT_arg = UUCASE (LEFT$ (case_CT_arg & " ", 1))

   IF  case_CT_arg <> "C" _
   AND case_CT_arg <> "T" THEN
       case_CT_arg =  "T"        '/ default to Text compares if C/T unavailable
   END IF

   curr_case_CT = case_CT_arg    '/ currently active C/T code in mapping string

   inp_len = LEN(source_arg)

   REDIM usage_flag (0 TO inp_len) AS BYTE

   '/ keep track of which data columns have been referenced by column numbers

   FOR n = 1 TO inp_len
      usage_flag(n) = 0
   NEXT


   '/ initialize EV/SV variables

   FOR n = 1 TO %mapstr_esv_max
      esv_nest_name(n) = ""
   NEXT

   esv_nest_level = 0
   esv_name = ""
   esv_text = ""




   '/ add safety zone and nul terminators to mapping string to simplify scan

   map_str = CHR$(map_arg, "  ", 0, 0)
   cp = STRPTR (map_str)

   GOSUB sub_command_reset

   '/ skip any leading blanks

   DO WHILE @cp = ASC(" ")  OR  @cp = ASC(";")
      INCR cp
   LOOP


   DO
      SELECT CASE AS CONST$ CHR$(@cp)

         '/ scan continues until we run into the nul terminator added above

         CASE CHR$(0)                                          '/ end of string
            EXIT DO                                             '/ done parsing


         CASE " ", ";"
            INCR cp                              '/ skip over blanks delimiters
            ITERATE DO


         CASE "\"                                               '/ reverse mode
            reverse_mode = 1
            INCR cp                                          '/ skip over slash

            '/-----------------------------------------------------------------
            '/ the reversal code must be followed by a digit, else it is being
            '/ used in an improper context
            '/ if digit found, field is parsed in next pass
            '/-----------------------------------------------------------------

            IF mapstr_is_digit_b (@cp) = 0 THEN
               err_arg = "Mapping reversal code \ used improperly"
               FUNCTION = %mapstr_retcode_ERR
               MExitFunc

            END IF

            ITERATE DO



         CASE "<"                                 '/ handle < and <> case codes
            INCR cp                                         '/ skip over < code
            c = CHR$(@cp)

            IF mapstr_is_delim(c) THEN          '/ code < was followed by delim
               mapstr_case = %mapstr_lower               '/ lower case in force
               ITERATE DO

            ELSEIF c = ">" THEN                           '/ code <> recognized
               INCR cp                                        '/ skip over code
               c = CHR$(@cp)

               IF mapstr_is_delim(c) THEN      '/ code <> was followed by delim
                  mapstr_case = %mapstr_none             '/ as-is case in force
                  ITERATE DO

               ELSE
                  err_arg = "Mapping case code <> used improperly"
                  FUNCTION = %mapstr_retcode_ERR
                  MExitFunc

               END IF

            ELSE
               err_arg = "Mapping case code < used improperly"
               FUNCTION = %mapstr_retcode_ERR
               MExitFunc

            END IF



         CASE ">"                                 '/ handle > and >< case codes
            INCR cp                                         '/ skip over > code
            c = CHR$(@cp)

            IF mapstr_is_delim(c) THEN          '/ code > was followed by delim
               mapstr_case = %mapstr_upper               '/ upper case in force
               ITERATE DO

            ELSEIF c = "<" THEN                           '/ code >< recognized
               INCR cp                                        '/ skip over code
               c = CHR$(@cp)

               IF mapstr_is_delim(c) THEN      '/ code >< was followed by delim
                  mapstr_case = %mapstr_invert          '/ invert case in force
                  ITERATE DO

               ELSE
                  err_arg = "Mapping case code >< used improperly"
                  FUNCTION = %mapstr_retcode_ERR
                  MExitFunc

               END IF

            ELSE
               err_arg = "Mapping case code > used improperly"
               FUNCTION = %mapstr_retcode_ERR
               MExitFunc

            END IF



         CASE "["                                '/ inserted text range operand

            range_len = mapstr_validate_text_range (cp)

            IF range_len = 0 THEN
               err_arg = "Mapping text-range operand is malformed"
               FUNCTION = %mapstr_retcode_ERR
               MExitFunc
            END IF

            GOSUB sub_generate_text_range

            cp += range_len
            ITERATE DO


         CASE $DQ, $SQ, "`"                                    '/ inserted text
         CASE_DQ_SQ_AQ:

            IF @cp = ASC("=") _                     '/ see if PD='right' format
            OR @cp = ASC(":") THEN
               c = CHR$(@cp[1])                      '/ quote is next char over
            ELSE
               c = CHR$(@cp)                       '/ save quote if diag needed
            END IF

            cp = mapstr_parse_txt (cp, parse_str1, parse_str2, cmd)

            IF cp = 0 THEN                       '/ something wrong with string

               IF ISNULL (cmd.txt_diag_str) THEN     '/ no custom msg available
                  err_arg = "Mapping text literal is malformed or " _
                          & "missing close " & c & " quote"
               ELSE
                  err_arg = cmd.txt_diag_str
               END IF

               FUNCTION = %mapstr_retcode_ERR
               MExitFunc

            END IF


            IF cmd.cmd_ID = %mapstr_none THEN
               '/ there is no active cmd_ID, just some inserted text
               '/ copy it to the result_str now

               result_str &= parse_str1

            ELSE

               '/ the extracted string was part of a larger operation
               '/ the string was being used for padding or similar reason
               '/ the second parse string is for RC and XC/XT command codes

               cmd_str1 = parse_str1
               cmd_str2 = parse_str2


               '/ if we are pointing to a digit, the operation has both a pad
               '/ string and a numeric string, so go handle that part next
               '/ if there's no digit, opcode has already been determined
               '/ so go select command handlers

               IF VERIFY (CHR$(@cp), "0123456789") = VERIFIED THEN
                  GOTO CASE_DIGITS

               ELSEIF CHR$(@cp) = "." THEN
                  GOTO CASE_DOT

               ELSE
                  GOTO CASE_OPCODE
               END IF

            END IF

            ITERATE DO       '/ only iterates if a "plain" column range command


         CASE "."
         CASE_DOT:

            INCR cp                                            '/ skip over dot

            IF  cmd.cmd_ID = %mapstr_none _        '/ there is no active cmd_ID
            AND mapstr_is_delim (CHR$ (@cp)) THEN      '/ illegal lone dot used
               err_arg = "Mapping symbol . notation used improperly"
               FUNCTION = %mapstr_retcode_ERR
               MExitFunc

            END IF

            '/ at this point, the single dot was found that is not followed by
            '/ a delimiter.  if there is a digit present, we can simply go to
            '/ the number handler.  if there is a dash, we can allow it for
            '/ now if we have a command, otherwise it's an invalid dash.

            IF @cp = ASC("-") THEN
               IF cmd.cmd_ID = %mapstr_none THEN   '/ there is no active cmd_ID

                  '/ a regular column range can't start with a dash, so neither
                  '/ can dot notation be followed by a dash if there is no cmd

                  err_arg = "Mapping symbol . notation used improperly"
                  FUNCTION = %mapstr_retcode_ERR
                  MExitFunc

               ELSE
                  dot_notation_used = 1
                  select_str = VARPTR (result_str)
                  GOTO CASE_DIGITS
               END IF

            ELSEIF mapstr_is_digit_b (@cp) THEN        '/ dot followed by digit
                  dot_notation_used = 1
                  select_str = VARPTR (result_str)
                  GOTO CASE_DIGITS

            ELSE
               '/ at this point, the dot notation is not followed by a digit or
               '/ dash if we have an active command, let the command handler
               '/ deal with it.  if there's no command, the . dot notation
               '/ does not have a any delimiter after it, and we documented
               '/ that they can't do that.

               IF cmd.cmd_ID <> %mapstr_none THEN  '/ there is an active cmd_ID
                  dot_notation_used = 1
                  select_str = VARPTR (result_str)
                  GOTO CASE_OPCODE
               END IF
            END IF

            err_arg = "Mapping symbol . notation used improperly"
            FUNCTION = %mapstr_retcode_ERR
            MExitFunc


         '/ a lone "+" is treated as a shorthand for 1+

         CASE "+"
         CASE_PLUS:
            IF mapstr_is_delim (CHR$ (@cp[1])) THEN
               INCR cp
               cmd.cmd_sep = %mapstr_plus

               cmd.cmd_num1 = 1
               cmd.cmd_num2 = 0

               cmd.cmd_num1_used = 1
               cmd.cmd_num2_used = 0

               GOTO CASE_OPCODE

            ELSE
               err_arg = "Mapping symbol + notation used improperly"
               FUNCTION = %mapstr_retcode_ERR
               MExitFunc
            END IF


         '/ for CONST$ cases, we have to enumerate each value ...

         CASE "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
         CASE_DIGITS:


            '/-----------------------------------------------------------------
            '/ if a command has not been identified yet, we need to a look-
            '/ ahead to see if a dxd command of the form nDD, nDX, nXD or nXX
            '/ is present.  if so, treat the leading digit as the beginning of
            '/ the command, and go to where we parse command codes.
            '/
            '/ the check for cmd.cmd_ID = %mapstr_none is needed to prevent a
            '/ crazy loop from occurring.  we are already going out on a limb
            '/ by using these GOTO's to begin with, but since this is a state
            '/ machine, it's mostly a legitimate way of doing it.
            '/-----------------------------------------------------------------

            IF cmd.cmd_ID = %mapstr_none THEN
               u1 = UUCASE (CHR$(@cp[1]))
               u2 = UUCASE (CHR$(@cp[2]))
               u12 = u1 & u2

               IF u12 = "DD" _
               OR u12 = "DX" _
               OR u12 = "XD" _
               OR u12 = "XX" THEN
                  GOTO CASE_LETTERS
                  '/ parse nDX etc. as if it started with letter

               END IF

               '/--------------------------------------------------------------
               '/ now that we confirmed there is no dxd command with a leading
               '/ digit, we have a stand-alone column reference.  we have to
               '/ mark this as a copying command (which it basically is) so
               '/ that our more strict categorization of copying vs. modifying
               '/ commands will work properly when it comes time to actually
               '/ move the columns of data.
               '/--------------------------------------------------------------

               cmd.cmd_type = %mapstr_copying_cmd

            END IF


            '/-----------------------------------------------------------------
            '/ if we get here from a leading dash, like L-9, then we 'fake' the
            '/ fact that two numbers are present, by assuming the first one is
            '/ zero.  that works well because 'cmd_num1' is already initialized
            '/ to zero when the code gets here, and that's what we want to
            '/ assume, so even though we don't really have the first number, it
            '/ works OK as if there was an implied 0 here.
            '/-----------------------------------------------------------------

            cmd.cmd_num1_used = 1                 '/ remember the field is here
            c = CHR$(@cp)                          '/ save digit if diag needed

            cp = mapstr_parse_numbers (cp, cmd)


            '/-----------------------------------------------------------------
            '/ if a mixed-mode range was used, we have to adjust it
            '/ since this must be done
            '/-----------------------------------------------------------------


            IF cmd.cmd_sep = %mapstr_dash THEN            '/ -x or x-y was used
               cmd.cmd_num2_used = 1              '/ remember the field is here
            END IF

            IF cp = 0 THEN

               IF cmd.num_diag_str = "" THEN                   '/ no custom msg
                  err_arg = "Mapping number malformed starting at '" & c & "'"
               ELSE
                  err_arg = cmd.num_diag_str                  '/ use custom msg
               END IF

               FUNCTION = %mapstr_retcode_ERR
               MExitFunc

            END IF


            '/-----------------------------------------------------------------
            '/ since a range may be used both for column numbers and size
            '/ ranges, we cannot reject ranges with 0 numbers just yet
            '/
            '/ remainder-column mode must not start at 0.  column selection
            '/ of 0+ is illegal, but might be valid if with cmd.cmd_ID
            '/-----------------------------------------------------------------

            IF cmd.cmd_sep = %mapstr_plus THEN
               IF cmd.cmd_num1 = 0 _
               AND cmd.cmd_ID = %mapstr_none THEN
                  err_arg = "Mapping remainder column-number cannot be 0+"
                  FUNCTION = %mapstr_retcode_ERR
                  MExitFunc

               END IF
            END IF


            '/-----------------------------------------------------------------
            '/ make sure \ reverse mode is used properly
            '/ reverse mode is allowed on \0 and \n+ where n > 0
            '/ reverse mode is not allowed on \xx-yy or on \n
            '/ this only applies to open col references, not to command codes
            '/-----------------------------------------------------------------

            reverse_err = 0

            IF reverse_mode = 1 THEN
               IF cmd.cmd_sep = %mapstr_none THEN             '/ legal usage \0
                  IF cmd.cmd_num1 <> 0 THEN                 '/ illegal usage \n
                     reverse_err = 1
                  END IF

               ELSEIF cmd.cmd_sep = %mapstr_plus THEN        '/ legal usage \n+
                  IF cmd.cmd_num1 = 0 THEN                 '/ illegal usage \0+
                     reverse_err = 1
                  END IF

               ELSEIF cmd.cmd_sep = %mapstr_dash THEN   '/ illegal usage \xx-yy
                     reverse_err = 1

               END IF
            END IF

            IF reverse_err = 1 THEN
               err_arg = "Mapping column reversal code \ used improperly"
               FUNCTION = %mapstr_retcode_ERR
               MExitFunc

            END IF

            '/ decide what to do once we reach the end of a number field
            '/ if there is an active cmd_ID, that gets processed
            '/ otherwise the number is a column reference


         CASE_OPCODE:

            '/-----------------------------------------------------------------
            '/ when a command takes a column range but one was not supplied
            '/ we imply an "auto reference".  when a command uses the result
            '/ string before one is defined, we imply an "auto copy".
            '/-----------------------------------------------------------------

            GOSUB sub_auto_copy_auto_ref


            '/-----------------------------------------------------------------
            '/ if the opcode is a modifying command, there is implied dot
            '/ notation, even if the dot is not used.  if so, be sure that
            '/ the select_ptr is pointing to the result string.
            '/-----------------------------------------------------------------

            IF cmd.cmd_type = %mapstr_modifying_cmd THEN
               select_str = VARPTR (result_str)
            END IF


            '/ perform some common code before selecting based on opcode

            IF cmd.cmd_ID = %mapstr_cmd_ID_EX _
            OR cmd.cmd_ID = %mapstr_cmd_ID_XE THEN

               '/--------------------------------------------------------------
               '/ if we are not supplied with ASCII/EBCDIC tran tables, EX
               '/ and XE, the command codes cannot be done.
               '/
               '/ At this point it could be possible to supply a set of
               '/ default translation tables, but with some risk.  that's
               '/ why we didn't do that as part of the initial
               '/ implementation of mapping strings.
               '/--------------------------------------------------------------

               IF tran_A2E_arg = 0 _
               OR tran_E2A_arg = 0 THEN
                  err_arg = "Mapping command " & _
                     cmd.cmd_name & " not available"

                  FUNCTION = %mapstr_retcode_ERR
                  MExitFunc
               END IF

            END IF


            IF mapstr_case = %mapstr_upper _
            OR mapstr_case = %mapstr_lower THEN

                  '/ explicit > or < mapping code was used
                hexstr_case = mapstr_case
            ELSE
                hexstr_case = %mapstr_upper
            END IF

            IF cmd.cmd_sep = %mapstr_dash THEN
               IF cmd.cmd_num1 = 0 _
               OR cmd.cmd_num2 = 0 THEN

                  '/ a column range can never be specified as 0-0, 0-5 or 5-0

                  err_arg = "Mapping number column range "                    _
                          & DEC$(cmd.cmd_num1) & "-"                          _
                          & DEC$(cmd.cmd_num2)                                _
                          & " malformed"

                  FUNCTION = %mapstr_retcode_ERR
                  MExitFunc

               END IF
            END IF

            SELECT CASE AS CONST cmd.cmd_ID

               CASE %mapstr_none            '/ no cmd code, just a number field
                  GOSUB sub_move_columns

               CASE %mapstr_cmd_ID_AQ                  '/ generate accent quote
                  GOSUB sub_opcode_AQ

               CASE %mapstr_cmd_ID_AX         '/ Do Ansi to Hex char conversion
                  GOSUB sub_opcode_AX

               CASE %mapstr_cmd_ID_C            '/ Center text in result string
                  GOSUB sub_opcode_C

               CASE %mapstr_cmd_ID_CC   '/ Set CASE C default for char searches
                  curr_case_CT = "C"        '/ <== NO SUB FOR THIS COMMAND CODE

               CASE %mapstr_cmd_ID_CT   '/ Set CASE T default for char searches
                  curr_case_CT = "T"        '/ <== NO SUB FOR THIS COMMAND CODE

               CASE %mapstr_cmd_ID_D        '/ Delete char cols from result str
                  GOSUB sub_opcode_D

               CASE %mapstr_cmd_ID_DC     '/ Delete char values from result str
                  GOSUB sub_opcode_DC

               CASE %mapstr_cmd_ID_DD       '/ Do Dec to Dec numeric conversion
                  dxd_source_type = %mapstr_dxd_dec
                  dxd_result_type = %mapstr_dxd_dec
                  GOSUB sub_opcode_DD

               CASE %mapstr_cmd_ID_DQ                  '/ generate double quote
                  GOSUB sub_opcode_DQ

               CASE %mapstr_cmd_ID_DV                        '/ include Dyn Var
                  esv_mode = %esv_mode_DV
                  esv_state = %esv_state_beg
                  GOSUB sub_opcode_DV

               CASE %mapstr_cmd_ID_DX       '/ Do Dec to Hex numeric conversion
                  dxd_source_type = %mapstr_dxd_dec
                  dxd_result_type = %mapstr_dxd_hex
                  GOSUB sub_opcode_DX

            '/ CASE %mapstr_cmd_ID_EV                        '/ include Env Var
            '/    esv_mode = %esv_mode_EV
            '/    esv_state = %esv_state_beg               '/ 1 = def, 0 = quit
            '/    GOSUB sub_opcode_EV

               CASE %mapstr_cmd_ID_EX       '/ Do EBCDIC to Hex char conversion
                  GOSUB sub_opcode_EX

               CASE %mapstr_cmd_ID_IC           '/ Invert Case of result string
                  GOSUB sub_opcode_IC

               CASE %mapstr_cmd_ID_L          '/ Left justify the result string
                  GOSUB sub_opcode_L

               CASE %mapstr_cmd_ID_LC        '/ Set result string to Lower Case
                  GOSUB sub_opcode_LC

               CASE %mapstr_cmd_ID_M                '/ Move chars within result
                  GOSUB sub_opcode_M

               CASE %mapstr_cmd_ID_P     '/ Pad the result string on both sides
                  GOSUB sub_opcode_P

               CASE %mapstr_cmd_ID_PA   '/ Pivot data and Align - also PLA, PRA
                  GOSUB sub_opcode_PA

               CASE %mapstr_cmd_ID_PD             '/ Pivot Data - also PLD, PRD
                  GOSUB sub_opcode_PD

               CASE %mapstr_cmd_ID_PL     '/ Pad the result string on Left side
                  GOSUB sub_opcode_PL

               CASE %mapstr_cmd_ID_PR    '/ Pad the result string on Right side
                  GOSUB sub_opcode_PR

               CASE %mapstr_cmd_ID_QD          '/ quit use of Dyn Val - private
                  esv_mode = %esv_mode_DV
                  esv_state = %esv_state_end
                  GOSUB sub_opcode_QD

            '/ CASE %mapstr_cmd_ID_QE          '/ quit use of Env Var - private
            '/    esv_mode = %esv_mode_EV         '/ 1 = SET mode, 0 = ENV mode
            '/    esv_state = %esv_state_end
            '/    GOSUB sub_opcode_QE

               CASE %mapstr_cmd_ID_QS          '/ quit use of Set Var - private
                  esv_mode = %esv_mode_SV         '/ 1 = SET mode, 0 = ENV mode
                  esv_state = %esv_state_end
                  GOSUB sub_opcode_QS

               CASE %mapstr_cmd_ID_R         '/ Right justify the result string
                  GOSUB sub_opcode_R

               CASE %mapstr_cmd_ID_RA                    '/ Replace All strings
                  GOSUB sub_opcode_RA

               CASE %mapstr_cmd_ID_RC    '/ Replace Chars via translation table
                  GOSUB sub_opcode_RC

               CASE %mapstr_cmd_ID_RL                       '/ Replace Leftmost
                  GOSUB sub_opcode_RL

               CASE %mapstr_cmd_ID_RP               '/ repeat the result string
                  GOSUB sub_opcode_RP

               CASE %mapstr_cmd_ID_RR                      '/ Replace Rightmost
                  GOSUB sub_opcode_RR

               CASE %mapstr_cmd_ID_RV     '/ Reverse order of the result string
                  GOSUB sub_opcode_RV

               CASE %mapstr_cmd_ID_SC     '/ Set result string to Sentence Case
                  GOSUB sub_opcode_SC

               CASE %mapstr_cmd_ID_SD                 '/ Sequence number to dec
                  dxd_source_type = %mapstr_dxd_dec
                  dxd_result_type = %mapstr_dxd_dec
                  GOSUB sub_opcode_SD

               CASE %mapstr_cmd_ID_SV                        '/ include Set Var
                  esv_mode = %esv_mode_SV         '/ 1 = SET mode, 0 = ENV mode
                  esv_state = %esv_state_beg
                  GOSUB sub_opcode_SV

               CASE %mapstr_cmd_ID_SX                 '/ Sequence number to hex
                  dxd_source_type = %mapstr_dxd_hex
                  dxd_result_type = %mapstr_dxd_hex
                  GOSUB sub_opcode_SX

               CASE %mapstr_cmd_ID_SQ                  '/ generate single quote
                  GOSUB sub_opcode_SQ

               CASE %mapstr_cmd_ID_T    '/ Trim the result string on both sides
                  GOSUB sub_opcode_T

               CASE %mapstr_cmd_ID_TC        '/ Set result string to Title Case
                  GOSUB sub_opcode_TC

               CASE %mapstr_cmd_ID_TL    '/ Trim the result string on Left side
                  GOSUB sub_opcode_TL

               CASE %mapstr_cmd_ID_TR   '/ Trim the result string on Right side
                  GOSUB sub_opcode_TR

               CASE %mapstr_cmd_ID_UC        '/ Set result string to Upper Case
                  GOSUB sub_opcode_UC

               CASE %mapstr_cmd_ID_XA         '/ Do Hex to Ansi char conversion
                  GOSUB sub_opcode_XA

               CASE %mapstr_cmd_ID_X                        '/ Exchange strings
                  GOSUB sub_opcode_X

               CASE %mapstr_cmd_ID_XD       '/ Do Hex to Dec numeric conversion
                  dxd_source_type = %mapstr_dxd_hex
                  dxd_result_type = %mapstr_dxd_dec
                  GOSUB sub_opcode_XD

               CASE %mapstr_cmd_ID_XE       '/ Do Hex to EBCDIC char conversion
                  GOSUB sub_opcode_XE

               CASE %mapstr_cmd_ID_XX       '/ Do Hex to Hex numeric conversion
                  dxd_source_type = %mapstr_dxd_hex
                  dxd_result_type = %mapstr_dxd_hex
                  GOSUB sub_opcode_XX

               CASE %mapstr_cmd_ID_Z          '/ Zero-suppress the result field
                  GOSUB sub_opcode_Z


               CASE ELSE              '/ JIC; should only happen in development

                  err_arg = "Mapping item " & cmd.cmd_name & " not supported"
                  FUNCTION = %mapstr_retcode_ERR
                  MExitFunc

            END SELECT

            GOSUB sub_command_reset


         '/ for CONST$ cases, we have to enumerate each value ...

         CASE "A", "B", "C", "D", "E", "F", "G", "H", "I", _
              "J", "K", "L", "M", "N", "O", "P", "Q", "R", _
                   "S", "T", "U", "V", "W", "X", "Y", "Z", _
              "a", "b", "c", "d", "e", "f", "g", "h", "i", _
              "j", "k", "l", "m", "n", "o", "p", "q", "r", _
                   "s", "t", "u", "v", "w", "x", "y", "z"

         CASE_LETTERS:

            cp = mapstr_parse_opcode (cp, cmd, LEN(result_str))

            IF cp = 0 _
            OR cmd.cmd_ID = %mapstr_none THEN         '/ undefined mapping code

               IF ISNULL(cmd.cmd_diag_str) THEN
                  '/ no custom message returned, use default one
                  err_arg = "Mapping string has illegal operation at " _
                          & cmd.cmd_name
               ELSE
                  err_arg = cmd.cmd_diag_str           '/ use custom diagnostic
               END IF

               FUNCTION = %mapstr_retcode_ERR
               MExitFunc

            END IF

            '/ if cmd_ID followed by quote, handle it

            IF mapstr_is_quote_b (@cp) THEN
               GOTO CASE_DQ_SQ_AQ

            '/ for commands like PD, allow null 'left' pair, like PD='right'
            '/ there is not much call for ":" but we allow it to be consistent

            ELSEIF mapstr_is_quote_b (@cp[1]) THEN
               IF @cp = ASC("=") _
               OR @cp = ASC(":") THEN
                  cmd.null_left_str = 1                  '/ note special format
                  GOTO CASE_DQ_SQ_AQ
               END IF
            END IF

            '/ if cmd_ID followed by number but no string, extract the
            '/ numeric field.  allow for commands like L-9

            '/ the 'numeric' field may be a . dot notation, or a dot followed
            '/ by a conventional column operand.  if we have a lone dot,
            '/ we have to 'fake' the attributes of a real column range.
            '/ that is OK if there is a defined command, else not

            IF @cp = ASC(".") THEN


               IF  cmd.cmd_ID = %mapstr_none _
               AND LEN(result_str) = 0 THEN
                  err_arg = _
                     "Mapping symbol . notation used before result defined"

                  FUNCTION = %mapstr_retcode_ERR
                  MExitFunc

               END IF

               dot_notation_used = 1
               select_str = VARPTR (result_str)
               INCR cp

               '/ example: if current result_str = "abcd" then a . dot notation
               '/ is treated like a "column" reference of 1-4, where both
               '/ 'num1' and 'num2' are defined with a 'dash' separator, but
               '/ the value is the result string, not the source string.

               IF mapstr_is_delim (CHR$ (@cp)) THEN
                  cmd.cmd_num1_used = 1
                  cmd.cmd_num2_used = 1
                  cmd.cmd_num1 = 1
                  cmd.cmd_num2 = LEN(result_str)

                  '/ we are looking too early for result, so look at source

                  IF cmd.cmd_num2 = 0 THEN
                     cmd.cmd_num2 = LEN(source_arg)
                  END IF

                  cmd.cmd_sep = %mapstr_dash
               END IF

            END IF

            IF VERIFY (CHR$(@cp), "-0123456789") = VERIFIED THEN
               GOTO CASE_DIGITS
            ELSEIF @cp = ASC("+") THEN
               GOTO CASE_PLUS
            ELSE
               GOTO CASE_OPCODE                     '/ cmd_ID is "plain" like T
            END IF

         CASE ELSE
            err_arg = "Mapping string is malformed at " & CHR$(@cp)
            FUNCTION = %mapstr_retcode_ERR
            MExitFunc

      END SELECT

   LOOP

   '/ at this point, the mapping string has been successfully processed

   result_arg = result_str

   FUNCTION = %mapstr_retcode_OK

   MExitFunc


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_command_reset                   **********************************
'/****   ---------------------------------   **********************************
'/*****************************************************************************

sub_command_reset:

   RESET cmd

   cmd.cmd_ID           = %mapstr_none
   cmd.cmd_sep          = %mapstr_none
   cmd.cmd_type         = %mapstr_other_cmd

   cmd.dxd_sign_code    = " "
   cmd.txt_case_CTX_1   = " "
   cmd.txt_case_CTX_2   = " "
   cmd.dxd_fmt_type     = " "

   reverse_err          = 0
   reverse_mode         = 0
   reverse_err          = 0
   dot_notation_used    = 0

   select_str = VARPTR (source_arg)


   RETURN ' from sub_command_reset


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_move_columns                    ****   COPYING COMMAND   *********
'/****   ---------------------------------   **********************************
'/*****************************************************************************

sub_move_columns:

   '/ handle column operands of the form 0, \0, n, n+, \n+, aa-bb and bb-aa

   '/ in case column references are out of bounds, we have to shorten them
   '/ to available data, done elsewhere; if both are too high, move is rejected
   '/ single numbers and ranges may have an end-relative column reference
   '/ indicated by a trailing * asterisk.


   GOSUB sub_get_move_text

   IF move_text_valid = 1 THEN
      '/ columns requested are within range, after possible adjustements
      '/ copy 'move_text' to result_str, adjusting case as requested

      result_str &= mapstr_handle_case (move_text, mapstr_case)
   END IF


   RETURN ' from sub_move_columns


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_opcode_AQ / DQ / SQ             ****   MODIFYING COMMAND   *******
'/****   ---------------------------------   ****  does not use columns  ******
'/*****************************************************************************

'/ add one or more quotes of the specified type to the result string

sub_opcode_AQ:
   c = "`"
   GOTO sub_opcode_AQ_DQ_SQ

sub_opcode_DQ:
   c = $DQ
   GOTO sub_opcode_AQ_DQ_SQ

sub_opcode_SQ:
   c = $SQ

sub_opcode_AQ_DQ_SQ:

   GOSUB sub_validate_replication_number

   IF  cmd.cmd_num1_used = 1 _
   AND cmd.cmd_num1 > 0 THEN
      c = REPEAT$(cmd.cmd_num1, c)
      result_str &= c
   END IF


   RETURN ' from sub_opcode_AQ / DQ / SQ


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_opcode_RP                       ****   MODIFYING COMMAND   *******
'/****   ---------------------------------   ****  does not use columns  ******
'/*****************************************************************************

'/ repeat the result string

sub_opcode_RP:

   GOSUB sub_validate_replication_number


   IF  cmd.cmd_num1_used = 1 _
   AND cmd.cmd_num1 > 0 THEN
      c = REPEAT$(cmd.cmd_num1, result_str)
      result_str &= c
   END IF


   RETURN ' from sub_opcode_RP


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_validate_replication_number     **********************************
'/****   ---------------------------------   **********************************
'/*****************************************************************************

'/ make sure a replication number is not used improperly, like a col ref

sub_validate_replication_number:


   IF cmd.cmd_num2_used = 1 _
   OR cmd.cmd_sep <> %mapstr_none THEN
      err_arg = "Mapping code " & cmd.cmd_name & _
         " has invalid replication factor"
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc
   END IF

   '/ imply a replication factor of 1 if number omitted.
   '/ don't override an explicit 0 if they happen to use one.

   IF cmd.cmd_num1_used = 0 THEN
      cmd.cmd_num1_used = 1
      cmd.cmd_num1 = 1
   END IF


   RETURN ' from sub_validate_replication_number


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_opcode_AX / XA / EX / XE        ****   COPYING COMMAND   *********
'/****   ---------------------------------   **********************************
'/*****************************************************************************

sub_opcode_AX:
sub_opcode_XA:
sub_opcode_EX:
sub_opcode_XE:

   '/ handle hex conversions

   GOSUB sub_get_move_text

   IF move_text_valid = 0 _
   OR diag_mode = 1       _                        '/ SPFLite sent us DIAG flag
   OR col_left = 0 THEN
      RETURN                                   '/ unable to perform the command
   END IF

   '/ determine CASE C/T to use.  method: if a C/T code was present on the
   '/ parsed string, use it, otherwise use the 'current' C/T code, which was
   '/ set by a CC or CT cmd, or defaulted from the caller's CT code which
   '/ is supposed to be C or T from the user's edit profile.

   IF cmd.cmd_str1_used = 0 THEN
      this_case_CT = "C"
      cmd_str1 = ""             '/ make sure code below doesn't bypass anything

   ELSEIF cmd.txt_case_CTX_1 <> " " THEN
      this_case_CT = cmd.txt_case_CTX_1    '/ use C/T suffix on str if supplied

   ELSEIF curr_case_CT <> " " THEN
      this_case_CT = curr_case_CT      '/ use C/T from CC or CT cmd, or profile

   ELSE
      this_case_CT = "C"                  '/ should not occur, but failsafe JIC

   END IF


   '/ create default bypass table: 1 = bypass, 0 = convert data

   FOR n = 0 TO 255
      byte_tab(n) = 0                            '/ assume data is not bypassed
   NEXT

   FOR n = 1 TO LEN(cmd_str1)
      c = MID$ (cmd_str1, n, 1)                 '/ grab char from bypass string

      IF this_case_CT = "T" THEN                '/ bypass chars are CASE T mode
         c = UUCASE(c)
         byte_tab(ASC(c)) = 1                      '/ mark this char for bypass

         c = LLCASE(c)
         byte_tab(ASC(c)) = 1                      '/ mark this char for bypass

      ELSE                                      '/ bypass chars are CASE C mode
         byte_tab(ASC(c)) = 1                      '/ mark this char for bypass

      END IF
   NEXT


   hex_result = ""                       '/ overall result_str from hex command
   hex_text = ""              '/ individual pieces if delimited by bypass chars


   '/ scan 'move_text'.  if any chars are in 'bypass' table, accumulate them
   '/ in 'hex_result', otherwise store in 'hex_text' until end of move_text
   '/ or the next bypass/delimiter char is found.  once the hex_text has
   '/ been built up, convert it as needed, and if in the correct format.

   n = 0
   DO
      n += 1

      IF n > LEN (move_text) THEN
         EXIT DO
      END IF

      c = MID$ (move_text, n, 1)

      IF byte_tab(ASC(c)) = 1 THEN                     '/ this char is bypassed

         IF LEN(hex_text) > 0 THEN             '/ first convert accumulated hex
            GOSUB sub_convert_hex_text
            hex_text = ""
         END IF

         hex_result &= c                    '/ then store curr char unconverted

      ELSE
         hex_text &= c
      END IF

   LOOP

   '/ handle last (or only) segment of to-be-converted characters

   IF LEN(hex_text) > 0 THEN
      GOSUB sub_convert_hex_text
      hex_text = ""
   END IF

   result_str &= hex_result


   RETURN ' from sub_opcode_AX/XA/EX/XE


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_convert_hex_text                **********************************
'/****   ---------------------------------   **********************************
'/*****************************************************************************

sub_convert_hex_text:

      '/ the dispatcher that go us here has already checked whether the
      '/ needed EBCDIC translation tables are available; we don't need to
      '/ check this again


      IF cmd.cmd_ID = %mapstr_cmd_ID_AX _
      OR cmd.cmd_ID = %mapstr_cmd_ID_EX THEN

         FOR hex_n = 1 TO LEN(hex_text)
            hex_c = MID$ (hex_text, hex_n, 1)

            '/ if we want to show the hex digits of the EBCDIC equivalent of
            '/ the current char, we have to translate it from ASCII to EBCDIC
            '/ first.  recall that tran_A2E_arg and tran_E2A_arg are passed as
            '/ BYTE PTR values, so we need to use @ and [] notation to get to
            '/ the passed translation tables.

            IF cmd.cmd_ID = %mapstr_cmd_ID_EX THEN
               hex_c = CHR$ (@tran_A2E_arg [ASC (hex_c)])
            END IF

            hex_xx = HEX$ (ASC (hex_c), 2, 0, 0)

            IF hexstr_case = %mapstr_lower THEN
               hex_xx = LLCASE (hex_xx)

            ELSEIF hexstr_case = %mapstr_upper THEN
               hex_xx = UUCASE (hex_xx)

            END IF

            hex_result &= hex_xx

         NEXT

      ELSE                                                   '/ handle XA or XE

         hex_temp = mapstr_X2A (hex_text, hex_diag)

         IF ISNULL (hex_temp) THEN
            err_arg = hex_diag              '/ has reason why hex number is bad
            FUNCTION = %mapstr_retcode_ERR
            MExitFunc

         END IF

         IF cmd.cmd_ID = %mapstr_cmd_ID_XE THEN

            FOR hex_n = 1 TO LEN(hex_temp)
               hex_c = MID$ (hex_temp, hex_n, 1)
               hex_c = CHR$ (@tran_E2A_arg [ASC (hex_c)])
               MID$ (hex_temp, hex_n, 1) = hex_c
            NEXT

         END IF

         hex_result &= hex_temp

      END IF


   RETURN ' from sub_convert_hex_text


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_opcode_C / L / R                ****   MODIFYING COMMAND   *******
'/****   ---------------------------------   ****  does not use columns  ******
'/*****************************************************************************

sub_opcode_C:
sub_opcode_L:
sub_opcode_R:

   '/ handle alignment codes C/L/R
   '/ the 'result_str' string is modifed, but the input string is not used

   '/ do common code first, then check cmd_ID for %mapstr_cmd_ID_L/R/C
   '/ formats: L9, L-9, L3-9, L9+

   IF LEN(result_str) = 0 THEN
      RETURN                                           '/ can't justify nothing
   END IF



   IF cmd.cmd_str1_used = 1 THEN
      cmd_str1 = LEFT$ (cmd_str1 & " ", 1)
   ELSE
      cmd_str1 = " "
   END IF


   SELECT CASE cmd.cmd_sep

      '/ handle cases where single number, or no number at all, was specified

      CASE %mapstr_none                                 '/ L9 - a single length

         IF cmd.cmd_num1 = 0 _
         OR cmd.cmd_num1_used = 0 THEN
            cmd.cmd_num1_used = 1
            cmd.cmd_num1 = LEN(result_str)
         END IF


         '/ first, remove all padding chars, then go back and justify the value
         '/ within the specified size, or within the former size of the result
         '/ string value.

         result_str = TRIM$(result_str, cmd_str1)


         IF LEN(result_str) < cmd.cmd_num1 THEN               '/ padding needed

            pad_len = cmd.cmd_num1 - LEN(result_str)
            pad_str = STRING$ (pad_len, cmd_str1)

            IF cmd.cmd_ID = %mapstr_cmd_ID_L THEN  '/ L just needs padding on R
               result_str &= pad_str

            ELSEIF cmd.cmd_ID = %mapstr_cmd_ID_R THEN   '/ R needs padding on L
               result_str = pad_str & result_str

            ELSEIF cmd.cmd_ID = %mapstr_cmd_ID_C THEN   '/ center needs padding

               '/ center justifying requires calculating two lengths, which may
               '/ be different if the total amount is an odd number

               pad_len1 = pad_len \ 2
               pad_len2 = pad_len - pad_len1

               pad_str1 = STRING$ (pad_len1, cmd_str1)
               pad_str2 = STRING$ (pad_len2, cmd_str1)

               result_str = pad_str1 & result_str & pad_str2

            END IF


         ELSEIF LEN(result_str) > cmd.cmd_num1 THEN        '/ truncation needed

            pad_len = cmd.cmd_num1 - LEN(result_str)

            IF cmd.cmd_ID = %mapstr_cmd_ID_L THEN  '/ L just needs padding on R
               result_str = LEFT$ (result_str, pad_len)

            ELSEIF cmd.cmd_ID = %mapstr_cmd_ID_R THEN   '/ R needs padding on L
               result_str = RIGHT$ (result_str, pad_len)

            ELSEIF cmd.cmd_ID = %mapstr_cmd_ID_C THEN   '/ center needs padding

               '/--------------------------------------------------------------
               '/ center justifying requires calculating two lengths, which may
               '/
               '/ be different if the total amount is an odd number
               '/ additional issue is that one of the lengths may be 0 if the
               '/ length being centered is 1; weird, but they may ask for it.
               '/
               '/ suppose string to be centered is xx1234xxx
               '/                  columns:        123456789
               '/ centering a 9-position string into 4 requires 5 positions
               '/ removed, 2 in front and 3 in back.  The starting column for
               '/ a MID$ will be 2+1=3, and the ending position is 9-3=6
               '/--------------------------------------------------------------

               pad_len1 = pad_len \ 2
               pad_len2 = pad_len - pad_len1

               result_str = _
                  MID$ (result_str, pad_len1 + 1 TO LEN(result_str) - pad_len2)

            END IF

         END IF                           '/ if already same length, do nothing



      CASE %mapstr_dash

         '/ L3-9 has min and max len, L0-9 and L-9 have no min

         IF cmd.cmd_num1 > cmd.cmd_num2 THEN
            '/ cannot use L9-3 or L3-0 for instance

            err_arg = "Mapping code " & cmd.cmd_name & " columns out of order"
            FUNCTION = %mapstr_retcode_ERR
            MExitFunc

         ELSEIF cmd.cmd_num2 = 0 THEN    '/ cannot use L-0 or L3-0 for instance
            err_arg = "Mapping code " & cmd.cmd_name & " illegal zero length"
            FUNCTION = %mapstr_retcode_ERR
            MExitFunc

         ELSEIF LEN(result_str) < cmd.cmd_num1 THEN

            '/ result_str shorter than min - padding needed

            pad_len = cmd.cmd_num1 - LEN(result_str)
            pad_str = STRING$ (pad_len, cmd_str1)

            IF cmd.cmd_ID = %mapstr_cmd_ID_L THEN  '/ L just needs padding on R
               result_str &= pad_str

            ELSEIF cmd.cmd_ID = %mapstr_cmd_ID_R THEN   '/ R needs padding on L
               result_str = pad_str & result_str

            ELSEIF cmd.cmd_ID = %mapstr_cmd_ID_C THEN   '/ center needs padding

               pad_len1 = pad_len \ 2
               pad_len2 = pad_len - pad_len1

               pad_str1 = STRING$ (pad_len1, cmd_str1)
               pad_str2 = STRING$ (pad_len2, cmd_str1)

               result_str = pad_str1 & result_str & pad_str2

            END IF

         ELSEIF LEN(result_str) > cmd.cmd_num2 THEN   '/ result longer than max

            pad_len = cmd.cmd_num2 - LEN(result_str)

            IF cmd.cmd_ID = %mapstr_cmd_ID_L THEN    '/ L just needs trunc on R
               result_str = LEFT$ (result_str, pad_len)

            ELSEIF cmd.cmd_ID = %mapstr_cmd_ID_R THEN    '/ R only trunc's on L
               result_str = RIGHT$ (result_str, pad_len)

            ELSEIF cmd.cmd_ID = %mapstr_cmd_ID_C THEN     '/ center needs trunc

               pad_len1 = pad_len \ 2
               pad_len2 = pad_len - pad_len1

               result_str = _
                  MID$ (result_str, pad_len1 + 1 TO LEN(result_str) - pad_len2)

            END IF

         END IF               '/ string within specified size range, do nothing


      CASE %mapstr_plus

         '/ L9+ is min len with no max, may pad but not trunc

         IF LEN(result_str) < cmd.cmd_num1 THEN

            '/ result_str shorter than min - padding needed

            pad_len = cmd.cmd_num1 - LEN(result_str)
            pad_str = STRING$ (pad_len, cmd_str1)

            IF cmd.cmd_ID = %mapstr_cmd_ID_L THEN  '/ L just needs padding on R
               result_str &= pad_str

            ELSEIF cmd.cmd_ID = %mapstr_cmd_ID_R THEN   '/ R needs padding on L
               result_str = pad_str & result_str

            ELSEIF cmd.cmd_ID = %mapstr_cmd_ID_C THEN   '/ center needs padding

               pad_len1 = pad_len \ 2
               pad_len2 = pad_len - pad_len1

               pad_str1 = STRING$ (pad_len1, cmd_str1)
               pad_str2 = STRING$ (pad_len2, cmd_str1)

               result_str = pad_str1 & result_str & pad_str2

            END IF

         END IF

   END SELECT


   RETURN ' from sub_opcode_C / L / R


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_opcode_D                        ****   MODIFYING COMMAND   *******
'/****   ---------------------------------   ****  uses result columns  *******
'/*****************************************************************************

sub_opcode_D:

   '/ delete characters from result string

   '/ we cannot use default auto-copy columns

   IF cmd.cmd_num1_used = 0 THEN
      err_arg = "Mapping code " & cmd.cmd_name & " requires a column operand"
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc
   END IF

   GOSUB sub_get_move_text

   IF move_text_valid = 0 _
   OR col_left < 1 THEN
      RETURN                                   '/ unable to perform the command
   END IF

   IF  col_left  = 1 _
   AND col_right = LEN(result_str) THEN
      result_str = ""            '/ they deleted the whole string; weird but OK
      RETURN
   END IF

   del_len = col_right - col_left + 1

   '/ carve out the part they want to delete
   result_str = CLIP$ (MID result_str, col_left, del_len)


   RETURN ' from sub_opcode_D


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_opcode_DC                       ****   MODIFYING COMMAND   *******
'/****   ---------------------------------   **** uses opt. result cols *******
'/*****************************************************************************

sub_opcode_DC:

   '/ perform character deletion based on character types

   IF LEN(result_str) = 0 THEN RETURN                     '/ nothing to work on

   GOSUB sub_get_move_text

   IF move_text_valid = 0 _
   OR col_left < 1 THEN
      RETURN                                   '/ unable to perform the command
   END IF


   '/ a string is required but not a number

   IF cmd.cmd_str1_used = 0 _
   OR LEN(cmd_str1) = 0 THEN
      err_arg = "Mapping code " & cmd.cmd_name & " requires a string operand"
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc

   END IF


   '/--------------------------------------------------------------------------
   '/ determine CASE C/T to use.  method: if a C/T code was present on the
   '/ parsed string, use it.  otherwise, use the 'current' C/T code, which was
   '/ set by a CC or CT command, or defaulted from the caller's CT code which
   '/ is supposed to be C or T from the user's edit profile.
   '/--------------------------------------------------------------------------

   IF cmd.txt_case_CTX_1 <> " " THEN
      this_case_CT = cmd.txt_case_CTX_1 '/ use C/T suffix on string if supplied
   ELSEIF curr_case_CT <> " " THEN
      this_case_CT = curr_case_CT      '/ use C/T from CC or CT cmd, or profile
   ELSE
      this_case_CT = "C"                  '/ should not occur, but failsafe JIC
   END IF



   '/ create default deletion table; we assume nothing is being deleted

   FOR n = 0 TO 255
      byte_tab(n) = 1                          '/ 1 means retain, 0 mean delete
   NEXT


   FOR n = 1 TO LEN (cmd_str1)
      c = MID$ (cmd_str1, n, 1)

      IF this_case_CT = "T" THEN
         c = UUCASE(c)
         byte_tab (ASC(c)) = 0             '/ mark character type to be deleted

         c = LLCASE(c)
         byte_tab (ASC(c)) = 0             '/ mark character type to be deleted

      ELSE
         byte_tab (ASC(c)) = 0             '/ mark character type to be deleted

      END IF
   NEXT

   '/ now, delete all chars appearing in the delete-list from the result_str
   '/ we always retain columns that are outside the column selection range,
   '/ if specified; that range defaults to the entire result string if omitted

   temp_result = ""

   FOR n = 1 TO LEN(result_str)
      c = MID$ (result_str, n, 1)
      IF n < col_left _
      OR n > col_right _
      OR byte_tab (ASC(c)) = 1 THEN     '/ this is a char not being deleted, so
         temp_result &= c                    '/ it's retained in the new result
      END IF
   NEXT

   result_str = temp_result


   RETURN ' from sub_opcode_DC


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_opcode_DD / DX / XD / XX        ****   COPYING COMMAND   *********
'/****   ---------------------------------   **********************************
'/*****************************************************************************

sub_opcode_DD:
sub_opcode_DX:
sub_opcode_XD:
sub_opcode_XX:

sub_opcode_SD:
sub_opcode_SX:

   '/ for SD/SX, just synthesize a string value from sequence number

   IF cmd.cmd_ID = %mapstr_cmd_ID_SD THEN
      dxd_source = DEC$ (sequence_num)
      cmd.dxd_numeric = sequence_num

   ELSEIF cmd.cmd_ID = %mapstr_cmd_ID_SX THEN
      dxd_source = HEX$ (sequence_num)
      cmd.dxd_numeric = sequence_num

   ELSE
      GOSUB sub_get_move_text

      IF move_text_valid = 0 _
      OR col_left = 0 THEN
         RETURN                                '/ unable to perform the command
      END IF

      '/ move_text now has the dxd source

      dxd_source = _
         mapstr_get_dxd_source (move_text, dxd_source_type, cmd, dxd_sign)


      IF dxd_source = "" THEN
         err_arg = "Mapping code " & cmd.cmd_name & " data value invalid"
         FUNCTION = %mapstr_retcode_ERR
         MExitFunc
      END IF
   END IF


   IF cmd.dxd_fmt_width_used = 0 _                    '/ no 'w' field specified
   OR cmd.dxd_fmt_width      = 0 THEN                 '/ or they asked for zero
      IF cmd.dxd_fmt_type = "V" THEN
         cmd.dxd_fmt_width = 1                           '/ assume 1, per specs
      ELSE
         cmd.dxd_fmt_width = LEN (dxd_source)
      END IF
   END IF


   '/ if a calculation operand was supplied, interpret the expression

   IF  cmd.cmd_str1_used = 1 _
   AND LEN(cmd_str1) > 0 THEN

      calc_first_num = cmd.dxd_numeric       '/ initial 'current value' in calc
      calc_final_num = sequence_num                '/ for SN read-only register
      calc_line_arg = TP.cfFLine
      calc_col_arg = TP.cfFCol

      dxd_rc = mapstr_calc         _
      (  calc_first_num            _
      ,  cmd_str1                  _
      ,  dxd_diag                  _
      ,  calc_final_num            _
      ,  calc_line_arg             _
      ,  calc_col_arg              _
         )

      IF dxd_rc <> %mapstr_retcode_OK THEN
         IF dxd_diag = "" THEN                       '/ no custom msg available
            err_arg = "Mapping code " & cmd.cmd_name & _
               " calculation operand malformed"
         ELSE
            err_arg = dxd_diag
         END IF

         FUNCTION = %mapstr_retcode_ERR
         MExitFunc
      END IF

      cmd.dxd_numeric = calc_final_num           '/ grab result from calculator

   END IF


   '/ convert numeric value back to a string

   IF dxd_result_type = %mapstr_dxd_hex THEN
      dxd_result = HEX$ (cmd.dxd_numeric)

   ELSE
      dxd_result = DEC$ (cmd.dxd_numeric)
   END IF


   '/ pad a V string with zeros if they ask for larger min width than needed

   IF cmd.dxd_fmt_type = "V" THEN
      IF LEN (dxd_result) < cmd.dxd_fmt_width THEN       '/ less than min width
         n = cmd.dxd_fmt_width - LEN (dxd_result)    '/ num of extras 0' needed
         dxd_result = REPEAT$ (n, "0") & dxd_result
      END IF

   ELSE
      IF cmd.dxd_fmt_type = "F" THEN
         pad_str = "0"                                   '/ F uses "0" fill/pad
      ELSE
         pad_str = " "                                   '/ Z uses " " fill/pad
      END IF

      IF LEN (dxd_result) > cmd.dxd_fmt_width THEN
         dxd_result = RIGHT$ (dxd_result, cmd.dxd_fmt_width)     '/ trunc value

      ELSEIF LEN (dxd_result) < cmd.dxd_fmt_width THEN
         n = cmd.dxd_fmt_width - LEN (dxd_result)
         dxd_result = REPEAT$ (n, pad_str) & dxd_result            '/ pad value
      END IF
   END IF


   '/ manage C 0x prefix if changing radix

   IF  cmd.dxd_sign_code = "C" _
   AND dxd_source_type <> dxd_result_type THEN
      IF dxd_result_type = %mapstr_dxd_hex THEN

         '/ going from dec to hex, we need to add 0x to "sign"
         '/ the standard for for C numbers is 0x12FF
         '/ keep the 'x' lower and hex digits upper unless they insist

         IF mapstr_case = %mapstr_upper THEN
            dxd_sign &= "0X"
         ELSE
            dxd_sign &= "0x"
         END IF

      ELSE
         '/ going from hex to dec, we need to remove 0x "sign"

         IF UUCASE (RIGHT$ (dxd_sign, 2)) = "0X" THEN
            dxd_sign = LEFT$ (dxd_sign, -2)              '/ remove rightmost 0X
         END IF
      END IF
   END IF

   IF mapstr_case = %mapstr_lower THEN
      dxd_result = LLCASE(dxd_result)

   ELSE
      dxd_result = UUCASE(dxd_result)
   END IF

   result_str &= (dxd_sign & dxd_result)


   RETURN ' from sub_opcode_DD / DX / XD / XX


'/-----------------------------------------------------------------------------
'/ FOR sub_opcode_DX --> SEE sub_opcode_DD   ****   COPYING COMMAND   *********
'/-----------------------------------------------------------------------------



'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_opcode_DV/EV/SV QD/QE/QS        ***   non-MODIFYING COMMAND   ****
'/****   ---------------------------------   **********************************
'/*****************************************************************************

sub_opcode_DV:
sub_opcode_QD:

sub_opcode_SV:
sub_opcode_QS:

'/  sub_opcode_EV:
'/  sub_opcode_QE:

   '/ handle inclusion of external mapping strings
   '/ a string is required but not a number
   '/ for DV, if string is omitted, it is assumed to be D

   IF  esv_mode = %esv_mode_DV _
   AND esv_state = %esv_state_beg THEN

      GOSUB sub_get_move_text

      '/ if there is no move text, we fake a single blank
      '/ later, a delete code will only apply if the move text is valid

      IF move_text_valid = 0 _
      OR col_left = 0 THEN
         col_left = 0
         col_right = 0
         move_text_valid = 0
         move_text = " "
      END IF
   END IF

   IF esv_mode = %esv_mode_DV THEN
      cmd_str1 = LEFT$ (UUCASE (TRIM$(cmd_str1)) & " ", 1)

      IF cmd.cmd_str1_used = 0 _
      OR cmd_str1 = ""         _
      OR cmd_str1 = " "        THEN
         cmd.cmd_str1_used = 1
         cmd_str1 = "D"               '/ assume delete code D is string omitted

      ELSE
         IF  cmd_str1 <> "D" _
         AND cmd_str1 <> "E" _
         AND cmd_str1 <> "K" THEN
            err_arg = "Mapping code " & cmd.cmd_name & " delete option " & _
               cmd_str1 & " is undefined"
            FUNCTION = %mapstr_retcode_ERR
            MExitFunc
         END IF
      END IF

   ELSE
      IF cmd.cmd_str1_used = 0 _
      OR LEN(cmd_str1) = 0 THEN
         err_arg = "Mapping code " & cmd.cmd_name & _
            " requires a string operand"
         FUNCTION = %mapstr_retcode_ERR
         MExitFunc
      END IF
   END IF

   cmd_str1 = UUCASE (cmd_str1)

   IF mapstr_valid_env_name(cmd_str1) = %mapstr_retcode_ERR THEN
      err_arg = "Mapping code " & cmd.cmd_name & " name " & _
         cmd_str1 & " is malformed"
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc
   END IF


   IF esv_mode = %esv_mode_SV THEN
      esv_name = mapstr_form_set_name (cmd_str1)


'/ ELSEIF esv_mode = %esv_mode_EV THEN
'/    esv_name = mapstr_form_env_name (cmd_str1)


   ELSEIF esv_mode = %esv_mode_DV THEN
      esv_name = cmd_str1                                          '/ D, E or K
   ELSE
      err_arg = "Mapping code " & cmd.cmd_name & " internal error esv_mode"
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc
   END IF


   IF esv_state = %esv_state_beg THEN

      '/ handle begin action - SV and EV
      '/ for DV, just grab the selected part of the result string

      IF esv_mode = %esv_mode_SV THEN
         esv_text = mapstr_fetch_set_var (esv_name)

'/    ELSEIF esv_mode = %esv_mode_EV THEN
'/       esv_text = mapstr_fetch_env_var (esv_name)

      ELSEIF esv_mode = %esv_mode_DV THEN
         esv_text = move_text
      END IF

      '/ if DV is used on selected text that is blank it just won't do anything
      '/ we won't treat that as an error

      IF esv_text = "" THEN
         IF esv_mode <> %esv_mode_DV THEN
            err_arg = "Mapping code " & cmd.cmd_name & _
               " variable name " & cmd_str1 & " is undefined"
         END IF

         FUNCTION = %mapstr_retcode_ERR
         MExitFunc
      END IF


      '/ for DV, delete part or all of the result string if requested

      '/ DV option code K means to keep the result str
      '/ DV option code E means to erase the result str
      '/ DV option code D means to delete the result str where it was used

      IF  esv_mode = %esv_mode_DV _
      AND esv_state = %esv_state_beg THEN
         IF cmd_str1 = "E" THEN
            result_str = ""

         ELSEIF cmd_str1 = "D"               _
         AND    move_text_valid              _
         AND    col_left  > 0                _
         AND    col_right <= LEN(result_str) _
         AND    col_left  <= col_right THEN

            '/ we have DV option D and a valid column range to delete

            IF  col_left  = 1 _
            AND col_right = LEN(result_str) THEN
               result_str = ""                 '/ they deleted the whole string

            ELSE
               del_len = col_right - col_left + 1
               '/ carve out the part they want to delete
               result_str = CLIP$ (MID result_str, col_left, del_len)
            END IF
         END IF
      END IF


      '/ create matching 'quit' command
      '/ QE/QS uses simple name not qualified one

      IF esv_mode = %esv_mode_SV THEN
         esv_quit = " QS'"  &  cmd_str1 & "' "

'/    ELSEIF esv_mode = %esv_mode_EV THEN
'/       esv_quit = " QE'"  &  cmd_str1 & "' "

      ELSEIF esv_mode = %esv_mode_DV THEN
         esv_quit = " QD'"  &  cmd_str1 & "' "
      END IF


      '/ be sure we are not using the given esv name recursively

      FOR n = 1 TO esv_nest_level
         IF esv_nest_name(n) = esv_name THEN

            IF esv_mode = %esv_mode_DV THEN
               err_arg = "Mapping code " & cmd.cmd_name & " used recursively"
            ELSE
               err_arg = "Mapping code " & cmd.cmd_name & _
                  " variable name " & cmd_str1 & " used recursively"
            END IF

            FUNCTION = %mapstr_retcode_ERR
            MExitFunc
         END IF
      NEXT


      INCR esv_nest_level

      IF esv_nest_level > %mapstr_esv_max THEN
         err_arg = "Mapping code " & cmd.cmd_name & _
            " variable name " & cmd_str1 & " exceeds nesting limit"

         FUNCTION = %mapstr_retcode_ERR
         MExitFunc
      END IF

      esv_nest_name(esv_nest_level) = esv_name      '/ save name for quit later

      '/ grab remaining portion of mapping command

      esv_tail = ""

      DO WHILE @cp <> 0                      '/ copy until nul terminator found
         esv_tail &= CHR$ (@cp)
         INCR cp
      LOOP

      esv_tail &= CHR$(0, 0)                          '/ BUILD$ won't take CHR$

      esv_cmd = BUILD$ ("  ", esv_text, esv_quit, esv_tail)

      '/ re-point mappping string char pointer to newly created string

      cp = STRPTR (esv_cmd)

   ELSE
      '/ handle end action = QD/QE/QS

      '/ for QE/QS, the top of the stack must match the esv_name, or the
      '/ closures do not nest properly.  the only way that could happen is a
      '/ logic error.

      IF esv_nest_level < 1 _
      OR esv_nest_level > %mapstr_esv_max _
      OR esv_nest_name(esv_nest_level) <> esv_name THEN
         err_arg = "Mapping code " & cmd.cmd_name & _
            " variable name " & cmd_str1 & " internal QE/QS error occurred"
         FUNCTION = %mapstr_retcode_ERR
         MExitFunc
      END IF

      esv_nest_name(esv_nest_level) = ""
      DECR esv_nest_level

   END IF


   RETURN ' from ub_opcode_DV/EV/SV QD/QE/QS


'/-----------------------------------------------------------------------------
'/ FOR sub_opcode_EX --> SEE sub_opcode_AX   ****   COPYING COMMAND   *********
'/-----------------------------------------------------------------------------


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_opcode_IC                       ****   MODIFYING COMMAND   *******
'/****   ---------------------------------   **** uses opt. result cols *******
'/*****************************************************************************

sub_opcode_IC:

   GOSUB sub_get_move_text

   IF move_text_valid = 0 _
   OR col_left = 0 THEN
      RETURN                                   '/ unable to perform the command
   END IF

   MID$(result_str, col_left TO col_right) = mapstr_ICASE(move_text)


   RETURN ' from sub_opcode_IC


'/-----------------------------------------------------------------------------
'/ FOR sub_opcode_L --> SEE sub_opcode_C     ****   MODIFYING COMMAND   *******
'/-----------------------------------------------------------------------------


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_opcode_LC                       ****   MODIFYING COMMAND   *******
'/****   ---------------------------------   **** uses opt. result cols *******
'/*****************************************************************************

sub_opcode_LC:

   GOSUB sub_get_move_text

   IF move_text_valid = 0 _
   OR col_left = 0 THEN
      RETURN                                   '/ unable to perform the command
   END IF

   MID$(result_str, col_left TO col_right) = LLCASE(move_text)


   RETURN ' from sub_opcode_LC


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_opcode_M                        ****   MODIFYING COMMAND   *******
'/****   ---------------------------------   ****  uses result columns  *******
'/*****************************************************************************

sub_opcode_M:

   '/ move character columns
   '/ we cannot use default auto-copy columns

   m_cmd_move_mode = %mapstr_m_cmd_none

   IF cmd.cmd_num1_used = 0 THEN
      err_arg = "Mapping code " & cmd.cmd_name _
         & " requires a column reference"
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc
   END IF

   GOSUB sub_get_move_text

   IF move_text_valid = 0 _
   OR col_left = 0 THEN
      err_arg = "Mapping code " & cmd.cmd_name _
         & " column reference is invalid"
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc
   END IF


   IF cmd.cmd_str1_used = 0 _
   OR LEN(cmd_str1)     = 0 _
   OR VERIFY(cmd_str1, " ") = VERIFIED THEN              '/ string is all blank
      err_arg = "Mapping code " & cmd.cmd_name & _
         " requires a destination reference"
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc
   END IF

   m_cmd = cmd              '/ initialize data structure with reasonable values

   IF cmd_str1 = "*" THEN
      m_cmd_move_mode = %mapstr_m_cmd_append

   ELSE
      cmd_str1 = TRIM$ (cmd_str1)
      m_cmd_dest_str = CHR$(cmd_str1, "  ", 0, 0)
      m_cmd_cp = STRPTR (m_cmd_dest_str)

      '/ parse the destination

      m_cmd_cp = mapstr_parse_numbers (m_cmd_cp, m_cmd)

      IF m_cmd_cp            = 0 _
      OR m_cmd.cmd_num1_used = 0 _
      OR m_cmd.cmd_num1      = 0 THEN

         err_arg = "Mapping code "        _
            & cmd.cmd_name              _
            & " destination reference " _
            & cmd_str1                  _
            & " is malformed"

         FUNCTION = %mapstr_retcode_ERR
         MExitFunc
      END IF

      GOSUB sub_get_dst_cols

      IF dst_cols_valid = 0 _
      OR dst_left = 0 THEN
         err_arg = "Mapping code " & cmd.cmd_name & _
            " destination reference is invalid"
         FUNCTION = %mapstr_retcode_ERR
         MExitFunc
      END IF

      IF m_cmd.cmd_num2_used THEN
         m_cmd_move_mode = %mapstr_m_cmd_replace
      ELSE
         m_cmd_move_mode = %mapstr_m_cmd_insert
      END IF
   END IF


   '/ cannot do further checking in DIAG mode

   IF diag_mode = 1 THEN
      RETURN
   END IF


   m_txt_seg_1 = ""
   m_txt_seg_2 = ""
   m_txt_seg_3 = ""
   m_txt_seg_4 = ""
   m_txt_seg_5 = ""


   '/ verify that columns do not overlap

   IF m_cmd_move_mode = %mapstr_m_cmd_append THEN
      IF col_right = LEN(result_str) THEN

         '/ they asked for the value to be appended, but the source string is
         '/ already located at the end.  a 'move' would result in no change

         RETURN
      END IF

      '/ the source data is in move_text

      IF col_left > 1 THEN
         m_txt_seg_1 = LEFT$ (result_str, col_left-1)
      END IF

      IF col_right < LEN(result_str) THEN
         m_txt_seg_2 = MID$ (result_str, col_right+1)
      END IF

      m_txt_seg_3 = move_text

      result_str = BUILD$ (m_txt_seg_1, m_txt_seg_2, m_txt_seg_3)

   ELSEIF m_cmd_move_mode = %mapstr_m_cmd_replace THEN

      IF (dst_left  >= col_left AND dst_left  <= col_right) _
      OR (dst_right >= col_left AND dst_right <= col_right) THEN

         err_arg = "Mapping code " & cmd.cmd_name & _
            " target string overlaps source string"
         FUNCTION = %mapstr_retcode_ERR
         MExitFunc
      END IF

      '/ the test above proves that the source and target do not overlap
      '/ that means either the source is to the left of the target, or it is
      '/ source is to the right of the target.

      '/ both sets of columns define areas of result_str that will be deleted
      '/ the string to be moved in is in move_text

      dst_min = MIN& (dst_left, col_left)

      IF dst_min > 1 THEN
         m_txt_seg_1 = LEFT$ (result_str, dst_min-1)
      END IF

      dst_max = MAX& (dst_right, col_right)

      IF dst_max < LEN(result_str) THEN
         m_txt_seg_5 = MID$ (result_str, dst_max+1)
      END IF

      IF dst_right < col_left THEN
         '/ dest on left side, source on right

         m_txt_seg_2 = move_text              '/ m_txt_seg_4 is left unassigned

         IF dst_right+1 <> col_left THEN
            '/ dest and source are not next to each other
            m_txt_seg_3 = MID$ (result_str, dst_right+1 TO col_left-1)
         END IF


      ELSEIF col_right < dst_left THEN
         '/ source on left side, dest on right

         m_txt_seg_4 = move_text              '/ m_txt_seg_2 is left unassigned

         IF col_right+1 <> dst_left THEN
            '/ dest and source are not next to each other
            m_txt_seg_3 = MID$ (result_str, col_right+1 TO dst_left-1)
         END IF

      ELSE
         err_arg = "Mapping code " & cmd.cmd_name & _
            " target string overlaps source string"
         FUNCTION = %mapstr_retcode_ERR
         MExitFunc
      END IF

      '/ reassemble all the segments

      result_str = BUILD$ ( _
         m_txt_seg_1, m_txt_seg_2, m_txt_seg_3, m_txt_seg_4, m_txt_seg_5)

   ELSEIF m_cmd_move_mode = %mapstr_m_cmd_insert THEN

      IF (dst_left >= col_left AND dst_left <= col_right) THEN

         err_arg = "Mapping code " & cmd.cmd_name & _
            " insertion point overlaps source string"
         FUNCTION = %mapstr_retcode_ERR
         MExitFunc
      END IF

      IF dst_left < col_left THEN                 '/ insertion is before source
         '/ before:  111222***444
         '/             ^  L R
         '/ result:  111***222444

         IF  dst_left > 1 THEN
            m_txt_seg_1 = LEFT$ (result_str, dst_left-1)
         END IF

         m_txt_seg_2 = MID$ (result_str, col_left TO col_right)

         m_txt_seg_3 = MID$ (result_str, dst_left TO col_left-1)

         IF col_right < LEN (result_str) THEN
            m_txt_seg_4 = MID$ (result_str, col_right+1)
         END IF

      ELSE                                         '/ insertion is after source
         '/ after:   111***333444
         '/             L R   ^
         '/ result:  111333***444

         IF  col_left > 1 THEN
            m_txt_seg_1 = LEFT$ (result_str, col_left-1)
         END IF

         m_txt_seg_2 = MID$ (result_str, col_right+1 TO dst_left-1)

         m_txt_seg_3 = MID$ (result_str, col_left TO col_right)

         m_txt_seg_4 = MID$ (result_str, dst_left)

      END IF

      result_str = BUILD$ (m_txt_seg_1, m_txt_seg_2, m_txt_seg_3, m_txt_seg_4)

   ELSE             '/ must still be '%mapstr_m_cmd_none' so something is wrong
      err_arg = "Mapping code " & cmd.cmd_name & _
         " destination reference is unresolved"
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc
   END IF


   RETURN ' from sub_opcode_M


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_opcode_P / PL / PR              ****   MODIFYING COMMAND   *******
'/****   ---------------------------------   ****  does not use columns  ******
'/*****************************************************************************

sub_opcode_P:
sub_opcode_PL:
sub_opcode_PR:

   '/ Pad the result_str on one or both sides with a fixed amount of the
   '/ defined pad characters, or blank.  The number operand defines the amount
   '/ of padding desired, which is not dependent on the result_str's current
   '/ length.

   '/ method: determine a span of padding characters, and then append that
   '/ value to the left side, right side, or both, of the result_str, to
   '/ produce a new result_str.


   IF cmd.cmd_sep <> %mapstr_none THEN
      err_arg = "Mapping code " & cmd.cmd_name & " has invalid number operand"
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc

   END IF


   IF cmd.cmd_num1_used THEN
      IF cmd.cmd_num1 = 0 THEN
         '/ user explicitly requested passing of zero
         '/ to pad with zero characters means to do nothing

         RETURN
      ELSE
         pad_len = cmd.cmd_num1
      END IF

   ELSE
      pad_len = 1
   END IF


   IF cmd.cmd_str1_used THEN
      cmd_str1 = LEFT$ (cmd_str1 & " ", 1)
   ELSE
      cmd_str1 = " "
   END IF

   pad_str = STRING$ (pad_len, cmd_str1)


   IF cmd.cmd_ID = %mapstr_cmd_ID_P THEN
      result_str = pad_str & result_str & pad_str

   ELSEIF cmd.cmd_ID = %mapstr_cmd_ID_PL THEN
      result_str = pad_str & result_str

   ELSEIF cmd.cmd_ID = %mapstr_cmd_ID_PR THEN
      result_str = result_str & pad_str
   END IF


   RETURN ' from sub_opcode_P/PL/PR


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_opcode_PD/PDL/PDR PA/PAL/PAR    ****   MODIFYING COMMAND   *******
'/****   ---------------------------------   **** uses opt. result cols *******
'/*****************************************************************************

'/ support for PD/PDL/PDR and PA/PAL/PAR

sub_opcode_PA:
   GOTO sub_opcode_generic_PDA

sub_opcode_PAL:
   GOTO sub_opcode_generic_PDA

sub_opcode_PAR:
   GOTO sub_opcode_generic_PDA

sub_opcode_PD:
   GOTO sub_opcode_generic_PDA

sub_opcode_PDL:
   GOTO sub_opcode_generic_PDA

sub_opcode_PDR:


sub_opcode_generic_PDA:                         '/ any sort of PD or PA command

   RETURN                                   '/ command has not been implemented


   '/ perform pivot operations

   IF LEN(result_str) = 0 THEN RETURN                     '/ nothing to work on

   '/ a string is required but not a number

   IF cmd.cmd_str1_used = 0 THEN
      err_arg = "Mapping code " & cmd.cmd_name & " requires a string operand"
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc

   END IF


   '/--------------------------------------------------------------------------
   '/ determine CASE C/T to use.  method: if a C/T code was present on the
   '/ parsed string, use it.  otherwise, use the 'current' C/T code, which was
   '/ set by a CC or CT command, or defaulted from the caller's CT code which
   '/ is supposed to be C or T from the user's edit profile.
   '/
   '/ if this was a string-pair, the parsing logic already reconciled the type
   '/ codes, so that string 1's type will suffice for both sides.
   '/
   '/ for pivot commands, C/T only applies to the pivot string itself, in case
   '/ an alphabetic pivot value was specified, which would not be typical.
   '/ here, the failsafe is T because, honestly, that seems like the right way.
   '/--------------------------------------------------------------------------


   IF cmd.txt_case_CTX_1 <> " " THEN
      this_case_CT = cmd.txt_case_CTX_1 '/ use C/T suffix on string if supplied
   ELSEIF curr_case_CT <> " " THEN
      this_case_CT = curr_case_CT      '/ use C/T from CC or CT cmd, or profile
   ELSE
      this_case_CT = "T"                  '/ should not occur, but failsafe JIC
   END IF

   '/--------------------------------------------------------------------------
   '/ pivot and boundary strings are not joined back together; they are just
   '/ two strings, and we are 'borrowing' the syntax of RC in order to grab the
   '/ two values, since the logic is already set up to handle that.
   '/--------------------------------------------------------------------------

   string_err = 0

   '/--------------------------------------------------------------------------
   '/ if a string pair used, join them together
   '/ we could let the later test check for proper sizes, but an odd
   '/ combination of values might escape the later test.  since we are sure
   '/ there are two strings, it's a stronger test to compare their lengths
   '/ individually now.
   '/--------------------------------------------------------------------------

   IF cmd.cmd_str2_used = 1 THEN
      IF LEN(cmd_str1) <> LEN(cmd_str2) THEN
         string_err = 1

      ELSE
         '/ convert RC'123'='abc' into RC'123=abc'
         cmd_str1 = cmd_str1 & "=" & cmd_str2
      END IF
   END IF


   IF LEN(cmd_str1) < 3 _                                   '/ string too short
   OR (LEN(cmd_str1)) MOD 2 = 0 THEN                    '/ even length is wrong
      string_err = 1

   ELSE
      txt_middle = (LEN(cmd_str1) \ 2) + 1

      c = MID$ (cmd_str1, txt_middle, 1)                  '/ grab midpoint char
      IF c <> "=" AND c <> ":" THEN                     '/ "=" in middle absent
         string_err = 1
      END IF
   END IF

   IF string_err = 1 THEN
      err_arg = "Mapping code " & cmd.cmd_name & " string malformed"
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc

   END IF

   '/ create default translation table

   FOR n = 0 TO 255
      byte_tab(n) = n
   NEXT

   '/ update translate table with before/after pairs:
   '/ if we are operating in CASE T mode, create UCASE/LCASE paits to map

   '/    RC'ABC=abc'    LEN(cmd_str1) = 7
   '/       1234567     txt_middle = 4  n = 1 to 3 = 1 to txt_middle-1

   FOR n = 1 TO txt_middle - 1
      txt_before = MID$ (cmd_str1, n, 1)
      txt_after = MID$ (cmd_str1, txt_middle + n, 1)

      IF this_case_CT = "T" THEN           '/ map upper and lower case versions
         txt_before = UUCASE (txt_before)
         byte_tab (ASC(txt_before)) = ASC(txt_after)

         txt_before = LLCASE (txt_before)
         byte_tab (ASC(txt_before)) = ASC(txt_after)

      ELSE                                           '/ use 'before' char as is
         byte_tab (ASC(txt_before)) = ASC(txt_after)

      END IF

   NEXT

   '/ finally, translate characters in the result_str

   FOR n = 1 TO LEN(result_str)
      c = MID$ (result_str, n, 1)
      MID$ (result_str, n, 1) = CHR$ (byte_tab (ASC(c)))
   NEXT


   RETURN ' from sub_opcode_PA/PAL/PAR PD/PDL/PDR


'/-----------------------------------------------------------------------------
'/ FOR sub_opcode_R --> SEE sub_opcode_C     ****   MODIFYING COMMAND   *******
'/-----------------------------------------------------------------------------


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_opcode_RA / RL /RR              ****   MODIFYING COMMAND   *******
'/****   ---------------------------------   **** uses opt. result cols *******
'/*****************************************************************************

sub_opcode_RA:                     '///--- UNFINISHED ----------->>>>>>>>>>> ??
sub_opcode_RL:                            '/  if cmd.cmd_ID = %mapstr_cmd_ID_RA
sub_opcode_RR:

   '/ perform string replacement

   IF cmd.cmd_str1_used = 0 THEN
      err_arg = "Mapping code " & cmd.cmd_name & " requires a string operand"
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc
   END IF

   '/ the RA/RL/RR commands can omit string 2 and it becomes a delete

   IF LEN(result_str) = 0 THEN RETURN                     '/ nothing to work on

   GOSUB sub_get_move_text

   scan_len1 = LEN (cmd_str1)
   scan_text = move_text

   '/ if the search string is longer than the found string, it cannot be found

   IF move_text_valid = 0 _
   OR col_left = 0 _
   OR scan_len1 > LEN (scan_text) THEN
      RETURN                                   '/ unable to perform the command
   END IF

   '/ string-2 can be length 0 if a deletion is requested

   GOSUB sub_resolve_case_CT                        '/ determine 'this_case_CT'

   scan_str1 = cmd_str1
   scan_result = ""


   '/--------------------------------------------------------------------------
   '/ scan the move_text in a sliding window, once character as a time
   '/ if a match is found, replace with string 2, otherwise peel off one
   '/ character at a time, then try again.
   '/
   '/ for RR, scan forward until last match is found, if any
   '/ last checkable location of 'n':
   '/
   '/ ABCDEFGHI      scan_text: LEN(scan_text) = 9
   '/ 123456789      values of n: max(n) = (LEN(scan_text) + 1) - scan_len1
   '/       GHI      value of scan_str1, len = 3      [ 10 - 3] = 7
   '/--------------------------------------------------------------------------

   IF cmd.cmd_ID = %mapstr_cmd_ID_RR THEN
      scan_limit = (LEN (scan_text) + 1) - scan_len1
      scan_last = 0

      FOR n = 1 TO scan_limit
         scan_data = MID$ (scan_text, n, scan_len1)
         IF mapstr_CT_equal (scan_data, scan_str1, this_case_CT) THEN
            scan_last = n
         END IF
      NEXT

      IF scan_last > 0 THEN
         n = scan_last          '/ to make this code look similar to other code

         IF n = 1 THEN                            '/ found at leftmost position
            scan_prefix = ""                          '/ nothing came before it
         ELSE
            scan_prefix = LEFT$ (scan_text, n-1)
         END IF

         '/ recapture data to make sure case conformance is correct
         scan_data = MID$ (scan_text, n, scan_len1)

         scan_str2 = cmd_str2                                 '/ get fresh copy

         IF cmd.cmd_str2_code = ":" THEN      '/ this is a case conformant move
            mapstr_conform_case (scan_data, scan_str2)
         END IF

         n += scan_len1                '/ column of remaining scan_text, if any

         IF n <= LEN (scan_text) THEN
            scan_suffix = MID$ (scan_text, n)
         ELSE
            scan_suffix = ""
         END IF

         '/ rebuild move text from these pieces

         scan_result = BUILD$ (scan_prefix, scan_str2, scan_suffix)

      END IF

   ELSE

      DO WHILE LEN (scan_text) >= scan_len1
         scan_data = LEFT$ (scan_text, scan_len1)

         IF mapstr_CT_equal (scan_data, scan_str1, this_case_CT) THEN

            scan_str2 = cmd_str2                              '/ get fresh copy

            IF cmd.cmd_str2_code = ":" THEN   '/ this is a case conformant move
               mapstr_conform_case (scan_data, scan_str2)
            END IF

            scan_result &= scan_str2

            IF LEN (scan_text) = scan_len1 THEN        '/ move text is consumed
               scan_text = ""
            ELSE
               scan_text = MID$ (scan_text, scan_len1+1)
            END IF

            '/ RL command stops replacement at first match

            IF cmd.cmd_ID = %mapstr_cmd_ID_RL THEN EXIT DO

         ELSE                  '/ scan_str1 not found at beginning of scan data

            IF LEN (scan_text) > 0 THEN
               scan_result &= LEFT$ (scan_text, 1)           '/ chop off a char
               scan_text = MID$ (scan_text, 2)             '/ give it to result
            END IF

         END IF
      LOOP

      '/ store any remaining final text that was too short to match
      scan_result &= scan_text

   END IF

   '/ store move_text into result_str from col_left to col_right of result_str

   IF col_left = 1 THEN
      scan_prefix = ""
   ELSE
      scan_prefix = LEFT$ (result_str, col_left-1)
   END IF

   IF col_right = LEN (result_str) THEN
      scan_suffix = ""
   ELSE
      scan_suffix = MID$ (result_str, col_right+1)
   END IF

   result_str = BUILD$ (scan_prefix, scan_result, scan_suffix)


   RETURN ' from sub_opcode_RA / RL / RR


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_opcode_RC                       ****   MODIFYING COMMAND   *******
'/****   ---------------------------------   **** uses opt. result cols *******
'/*****************************************************************************

sub_opcode_RC:

   '/ perform character replacement

   IF LEN(result_str) = 0 THEN RETURN                     '/ nothing to work on

   '/ command requires a string pair

   IF cmd.cmd_str1_used = 0 THEN
      err_arg = "Mapping code " & cmd.cmd_name & " requires a string operand"
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc
   END IF

   '/ if parse_text didn't find a good string pair ...

   IF cmd.cmd_str2_used = 0 THEN
      err_arg = "Mapping code " & cmd.cmd_name & " requires a string pair"
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc
   END IF

   IF LEN(cmd_str1) = 0 _
   OR LEN(cmd_str2) = 0 _
   OR LEN(cmd_str1) <> LEN(cmd_str2) THEN
      err_arg = "Mapping code " & cmd.cmd_name & " string malformed"
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc
   END IF

   GOSUB sub_resolve_case_CT

   '/ create default translation table

   FOR n = 0 TO 255
      byte_tab(n) = n
   NEXT

   '/ update translate table with before/after pairs:
   '/ if we are operating in CASE T mode, create UCASE/LCASE paits to map

   FOR n = 1 TO LEN(cmd_str1)
      txt_before = MID$ (cmd_str1, n, 1)
      txt_after  = MID$ (cmd_str2, n, 1)

      IF this_case_CT = "T" THEN           '/ map upper and lower case versions
         txt_before = UUCASE (txt_before)
         byte_tab (ASC(txt_before)) = ASC(txt_after)

         txt_before = LLCASE (txt_before)
         byte_tab (ASC(txt_before)) = ASC(txt_after)

      ELSE                                           '/ use 'before' char as is
         byte_tab (ASC(txt_before)) = ASC(txt_after)

      END IF

   NEXT

   '/ finally, translate characters in the result_str

   FOR n = 1 TO LEN(result_str)
      c = MID$ (result_str, n, 1)
      MID$ (result_str, n, 1) = CHR$ (byte_tab (ASC(c)))
   NEXT


   RETURN ' from sub_opcode_RC


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_opcode_RV                       ****   MODIFYING COMMAND   *******
'/****   ---------------------------------   **** uses opt. result cols *******
'/*****************************************************************************

sub_opcode_RV:

   '/ reverse the order of characters within the result string

   GOSUB sub_get_move_text

   IF move_text_valid = 0 _
   OR col_left = 0 THEN
      RETURN                                   '/ unable to perform the command
   END IF

   MID$(result_str, col_left TO col_right) = STRREVERSE$(move_text)


   RETURN ' from sub_opcode_RV


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_opcode_SC                       ****   MODIFYING COMMAND   *******
'/****   ---------------------------------   **** uses opt. result cols *******
'/*****************************************************************************

sub_opcode_SC:

   GOSUB sub_get_move_text

   IF move_text_valid = 0 _
   OR col_left = 0 THEN
      RETURN                                   '/ unable to perform the command
   END IF

   MID$(result_str, col_left TO col_right) = mapstr_SCASE(move_text)


   RETURN ' from sub_opcode_SC


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_opcode_TC                       ****   MODIFYING COMMAND   *******
'/****   ---------------------------------   **** uses opt. result cols *******
'/*****************************************************************************

sub_opcode_TC:

   GOSUB sub_get_move_text

   IF move_text_valid = 0 _
   OR col_left = 0 THEN
      RETURN                                   '/ unable to perform the command
   END IF

   MID$(result_str, col_left TO col_right) = mapstr_TCASE(move_text)


   RETURN ' from sub_opcode_TC


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_opcode_T / TL / TR              ****   MODIFYING COMMAND   *******
'/****   ---------------------------------   ****  does not use columns  ******
'/*****************************************************************************

sub_opcode_T:
sub_opcode_TL:
sub_opcode_TR:

   '/ Trim the result_str on one or both sides by a maximum number of
   '/ defined pad characters, or blank.  If the current result_str consists
   '/ entirely of pad characters, it is converted into a zero length string.

   '/ method: scan the result_str left to right, and right to left, finding the
   '/ number of leading and trailing pad characters present, if any.  Once this
   '/ is determined, use a MID$ function to extract the desired substring.


   IF cmd.cmd_sep <> %mapstr_none THEN
      err_arg = "Mapping code " & cmd.cmd_name & " has invalid number operand"
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc

   END IF

   IF cmd.cmd_num1_used THEN
      IF cmd.cmd_num1 = 0 THEN
         '/ user explicitly requested passing of zero
         '/ to trim with zero characters means to do nothing
         RETURN
      ELSE
         trim_max = cmd.cmd_num1
      END IF
   ELSE
      trim_max = LEN(result_str)        '/ this is a maximum amount to trim off
   END IF

   IF LEN(result_str) = 0 THEN
      '/ there is nothing available to trim, we cannot do this command code
      RETURN
   END IF

   IF cmd.cmd_str1_used THEN
      cmd_str1 = LEFT$ (cmd_str1 & " ", 1)
   ELSE
      cmd_str1 = " "
   END IF


   '/ determine number of leading pad chars

   IF cmd.cmd_ID = %mapstr_cmd_ID_T _
   OR cmd.cmd_ID = %mapstr_cmd_ID_TL THEN

      trim_count = 0
      FOR n = 1 TO trim_max
         IF MID$(result_str, n, 1) = cmd_str1 THEN
            trim_count = n
         ELSE
            EXIT FOR
         END IF
      NEXT

      trim_len1 = trim_count

   ELSE
      trim_len1 = 0                         '/ if TR, leading pad chars ignored
   END IF


   '/ determine number of trailing pad chars
   '/ method: flip the result_str string, and repeat code above

   IF cmd.cmd_ID = %mapstr_cmd_ID_T _
   OR cmd.cmd_ID = %mapstr_cmd_ID_TR THEN

      trim_count = 0
      temp_result = STRREVERSE$(result_str)

      FOR n = 1 TO trim_max
         IF MID$(temp_result, n, 1) = cmd_str1 THEN
            trim_count = n
         ELSE
            EXIT FOR
         END IF
      NEXT

      trim_len2 = trim_count

   ELSE
      trim_len2 = 0                        '/ if TL, trailing pad chars ignored
   END IF


   IF  trim_len1 = 0 _
   AND trim_len2 = 0 THEN
      '/ curr value of result_str does not begin or end with current pad char
      '/ there is nothing we need to do
      RETURN
   END IF

   '/                                  123   54321
   '/ get MID$ positions               xxx///xxxxx
   '/                                  12345678901
   '/
   '/ LEN(result_str) = 11, trim_len2 = 5, 11-5 = 6 = rightmost non-pad column

   trim_pos1 = trim_len1 + 1
   trim_pos2 = LEN(result_str) - trim_len2

   IF trim_pos1 > trim_pos2 THEN
      '/ this can only happen if entire string was all pad characters
      result_str = ""
      RETURN
   END IF

   '/ trim off requested chars, no more than maximum if one is specified

   result_str = MID$(result_str, trim_pos1 TO trim_pos2)


   RETURN ' from sub_opcode_T / TL / TR


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_opcode_UC                       ****   MODIFYING COMMAND   *******
'/****   ---------------------------------   **** uses opt. result cols *******
'/*****************************************************************************

sub_opcode_UC:

   GOSUB sub_get_move_text

   IF move_text_valid = 0 _
   OR col_left = 0 THEN
      RETURN                                   '/ unable to perform the command
   END IF

   MID$(result_str, col_left TO col_right) = UUCASE(move_text)


   RETURN ' from sub_opcode_UC


'/-----------------------------------------------------------------------------
'/ FOR sub_opcode_XA --> SEE sub_opcode_AX   ****   COPYING COMMAND   *********
'/-----------------------------------------------------------------------------


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_opcode_X                        ****   COPYING COMMAND   *********
'/****   ---------------------------------   **********************************
'/*****************************************************************************

sub_opcode_X:

   '/ Exchange strings

   GOSUB sub_get_move_text

   IF move_text_valid = 0 _
   OR col_left = 0 THEN
      RETURN                                   '/ unable to perform the command
   END IF

   '/ since X is a copying command, we are not dependent on 'result_str'
   '/ having any data in it yet, so we don't check it for ISNULL here.


   IF cmd.cmd_str1_used = 0 _
   OR cmd.cmd_str2_used = 0 THEN
      err_arg = "Mapping code " & cmd.cmd_name & " requires a string-pair"
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc
   END IF


   '/--------------------------------------------------------------------------
   '/ determine CASE C/T to use.  method: if a C/T code was present on the
   '/ parsed string, use it.  otherwise, use the 'current' C/T code, which was
   '/ set by a CC or CT command, or defaulted from the caller's CT code which
   '/ is supposed to be C or T from the user's edit profile.
   '/
   '/ if this was a string-pair, the parsing logic already reconciled the type
   '/ codes, so that string 1's type will suffice for both sides.
   '/ the parse text function forces the CTX type for the XT command to T.
   '/--------------------------------------------------------------------------


   IF cmd.txt_case_CTX_1 <> " " THEN
      this_case_CT = UUCASE (cmd.txt_case_CTX_1)  '/ from suffix on str if used

   ELSEIF curr_case_CT <> " " THEN
      this_case_CT = UUCASE (curr_case_CT)     '/ from CC or CT cmd, or profile

   ELSE
      this_case_CT = "C"                  '/ should not occur, but failsafe JIC

   END IF


   xc_txt_before = cmd_str1
   xc_txt_after  = cmd_str2
   xc_txt_result = ""

   '/ if move_text matches first string, replace with second, and vice versa

   IF mapstr_CT_equal (move_text, xc_txt_before, this_case_CT) THEN
      xc_txt_result = xc_txt_after
      move_text_replaced = 1

   ELSEIF mapstr_CT_equal (move_text, xc_txt_after, this_case_CT) THEN
      xc_txt_result = xc_txt_before
      move_text_replaced = 1

   ELSE
      move_text_replaced = 0
   END IF

   '/ for CASE CONFORMANCE, propagate sign from move_text to selected string

   IF  cmd.cmd_str2_code = ":" _    '/ ":" set when used or if both codes are T
   AND move_text_replaced = 1 THEN
      mapstr_conform_case (move_text, xc_txt_result)
   END IF


   IF move_text_replaced = 1 THEN
      IF cmd.use_result_str = 0 THEN

         '/ original value came from source_str so go ahead and concatenate it
         result_str &= xc_txt_result

      ELSE

         '/ update result_str in place; don't concatenate answer to it
         result_str = mapstr_insert_into_col_range ( _
            result_str, col_left, col_right, xc_txt_result)

      END IF
   ELSE                                                   '/ no match was found
      IF cmd.use_result_str = 0 THEN

         '/ original value came from source_str so go ahead and concatenate it
         result_str &= xc_txt_result

         '/ IF VALUE CAME FROM RESULT STRING, BUT NO MATCH WAS FOUND, LEAVE
         '/ THE RESULT STRING ALONE.  IT DOESN'T MAKE SENSE TO MAKE A DUPLICATE
         '/ COPY OF THE VALUE FOR A NOT-FOUND CONDITION.

      END IF
   END IF


   RETURN ' from sub_opcode_X


'/-----------------------------------------------------------------------------
'/ FOR sub_opcode_XD --> SEE sub_opcode_DD   ****   COPYING COMMAND   *********
'/-----------------------------------------------------------------------------


'/-----------------------------------------------------------------------------
'/ FOR sub_opcode_XE --> SEE sub_opcode_AX   ****   COPYING COMMAND   *********
'/-----------------------------------------------------------------------------


'/-----------------------------------------------------------------------------
'/ FOR sub_opcode_XX --> SEE sub_opcode_DD   ****   COPYING COMMAND   *********
'/-----------------------------------------------------------------------------


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_opcode_Z                        ****   MODIFYING COMMAND   *******
'/****   ---------------------------------   ****  does not use columns  ******
'/*****************************************************************************

sub_opcode_Z:
   '/ zero-suppress the result_str

   IF LEN(result_str) = 0 THEN RETURN                    '/ nothing to suppress

   '/ only an optional plain number is allowed

   IF cmd.cmd_num1_used THEN
      IF cmd.cmd_sep <> %mapstr_none THEN
         err_arg = "Mapping code " & cmd.cmd_name & _
            " has invalid number operand"

         FUNCTION = %mapstr_retcode_ERR
         MExitFunc

      END IF

      IF cmd.cmd_num1 >= LEN(result_str) THEN
         '/ user asked for entire string not be suppressed; nothing to do
         RETURN
      END IF

   ELSE
      cmd.cmd_num1 = 0              '/ entire field is eligible for suppression
   END IF

   col_final = LEN(result_str) - cmd.cmd_num1    '/ where suppression can occur

   '/ create list of characters to be significant vs. ignored

   FOR n = 0 TO 255
      byte_tab(n) = %mapstr_sig             '/ initially, all chars significant
   NEXT

   '/ add user-define chars to the list as ignored chars (not sig)

   IF cmd.cmd_str1_used = 1 THEN
      FOR n = 1 TO LEN(cmd_str1)
         c = MID$ (cmd_str1, n, 1)
         byte_tab (ASC(c)) = %mapstr_not_sig
      NEXT
   END IF

   '/ certain assumptions cannot be overridden by the user
   '/ digits 1-9 always significant, 0 and blank never significant

   byte_tab (ASC("0")) = %mapstr_not_sig
   byte_tab (ASC(" ")) = %mapstr_not_sig

   FOR n = ASC("1") TO ASC("9")
      byte_tab (n) = %mapstr_sig
   NEXT

   FOR n = 1 TO col_final
      c = MID$(result_str, n, 1)

      IF byte_tab(ASC(c)) = %mapstr_sig THEN
         '/ stop suppression at first significant character
         RETURN
      END IF

      MID$(result_str, n, 1) = " "
   NEXT


   RETURN ' from sub_opcode_Z


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_auto_copy_auto_ref              **********************************
'/****   ---------------------------------   **********************************
'/*****************************************************************************

sub_auto_copy_auto_ref:

   '/--------------------------------------------------------------------------
   '/ if we have a modifying command, or a copying command with dot notation,
   '/ a reference is being made to the result string.  if the current result
   '/ string is null, the source string gets copied over to the result.  that
   '/ is what the specs call auto-copy.
   '/
   '/ the command parse sets cmd.use_result_str so that is all we have to check
   '/
   '/ when an auto-copy occurs while a case conversion code is in effect, the
   '/ copied data needs to conform to the requested case.
   '/--------------------------------------------------------------------------

   IF  cmd.use_result_str = 1 _
   AND ISNULL (result_str) THEN               '/ perform the auto-copy by golly
      result_str = mapstr_handle_case (source_arg, mapstr_case)
   END IF


   '/--------------------------------------------------------------------------
   '/ convert mixed-mode references to normal ones.  mixed mode only occurs
   '/ for a column ranges, so when we are done, the "sep" mode is "dash"
   '/--------------------------------------------------------------------------

   IF cmd.use_result_str = 1 THEN
      col_max = LEN (result_str)
   ELSE
      col_max = LEN (source_arg)
   END IF

   mapstr_normalize_columns (cmd, col_max)


   '/--------------------------------------------------------------------------
   '/ if we have a command that uses column numbers, but none were provided,
   '/ we will assume that the entire source or result string is being
   '/ referenced.  for a few commands like C they take numeric arguments but
   '/ they don't take column ranges, so we can't adjust them in this way.
   '/
   '/ commands that are neither copying nor modifying will not have the
   '/ cmd.takes_col_ref flag set.
   '/
   '/ this is what the specs call auto-reference.
   '/
   '/ some commands cannot use auto-reference, like D and M
   '/ it would be too dangerous to allow a plain D to delete everything.
   '/
   '/ we fake out a "dash" reference, so that if a command like UC has not
   '/ column reference, and the result is length 5, we fake a ref of 1-5.
   '/--------------------------------------------------------------------------

   IF  cmd.takes_col_ref = 1 _                      '/ the command uses columns
   AND cmd.suppress_auto_ref = 0 _          '/ it's not special cmd like D or M
   AND cmd.cmd_num1_used = 0 THEN                  '/ no columns were specified
      cmd.cmd_sep       = %mapstr_dash
      cmd.cmd_num1_used = 1
      cmd.cmd_num2_used = 1
      cmd.cmd_num1      = 1

      IF cmd.use_result_str = 1 THEN
         cmd.cmd_num2 = LEN (result_str)
      ELSE
         cmd.cmd_num2 = LEN (source_arg)
      END IF
   END IF


   RETURN ' from sub_auto_copy_auto_ref


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_get_move_text                   **********************************
'/****   ---------------------------------   **********************************
'/*****************************************************************************

sub_get_move_text:

   '/--------------------------------------------------------------------------
   '/ handle standard columns of the form 0, \0, n, n+, \n+, aa-bb and bb-aa
   '/
   '/ in case column references are out of bounds, we have to shorten them
   '/ to available data, done elsewhere; if both are too high, move is rejected
   '/ single numbers and ranges may have an end-relative column reference
   '/ indicated by a trailing * asterisk.
   '/
   '/ this routine is called by a number of command-code actions.  not all of
   '/ the command codes use the 'move_text'.  it was decided to make this a
   '/ general purpose routine, even if not all of its results always get used.
   '/--------------------------------------------------------------------------

   '/ fake the move text in DIAG mode
   '/ this prevents data errors when there's no actual data present

   IF diag_mode = 1 THEN
      move_text       = "0000"
      col_left        = 1
      col_right       = LEN (move_text)
      col_first       = col_left
      col_final       = col_right
      col_step        = 1
      cmd.cmd_num1    = col_left
      cmd.cmd_num2    = col_right
      move_text_valid = 1

      RETURN
   END IF



   move_text = ""               '/ accumulation area, before it's case-adjusted
   col_left  = 0                            '/ resolved column range, low value
   col_right = 0                           '/ resolved column range, high value

   IF  cmd.cmd_num1 > LEN(@select_str) _
   AND cmd.cmd_num2 > LEN(@select_str) THEN

      '/ range completely outside of found string
      move_text_valid = 0
      RETURN
   END IF

   SELECT CASE CONST cmd.cmd_sep

      CASE %mapstr_none                                      '/ a simple column

         IF cmd.cmd_num1 = 0 THEN

            '/ leave col_left and col_right at 0; other cmds can't use this

            '/ gather all unreferenced columns and insert them now

            IF dot_notation_used THEN
               err_arg = "Mapping symbol . notation invalid with column 0"
               FUNCTION = %mapstr_retcode_ERR
               MExitFunc

            END IF


            '/-----------------------------------------------------------------
            '/ modifying commands always reference the result string.
            '/ column 0 means unreferenced columns of the source string
            '/ using column 0 on a modifying command makes no sense
            '/ this test does not apply to stand-alone column references,
            '/ which are implicitly a kind of copying command
            '/-----------------------------------------------------------------

            IF  cmd.cmd_type <> %mapstr_copying_cmd THEN
               err_arg = "Mapping command " & cmd.cmd_name & _
                  " cannot reference column 0"
               FUNCTION = %mapstr_retcode_ERR
               MExitFunc
            END IF

            IF reverse_mode = 1 THEN
               col_first = LEN(source_arg)
               col_final = 1
               col_step = -1
            ELSE
               col_first = 1
               col_final = LEN(source_arg)
               col_step  = 1
            END IF

            FOR n = col_first TO col_final STEP col_step

               '/--------------------------------------------------------------
               '/ we only take data that has not yet been referenced
               '/ these have a usage_flag of 0
               '/
               '/ taking them now means they are used, in case another
               '/ 0 or \0 action is requested later
               '/
               '/ we already rejected possible dot notation here, so marking
               '/ the usage flags is OK
               '/--------------------------------------------------------------

               IF usage_flag(n) = 0 THEN
                  usage_flag(n) = 1

                  c = MID$(source_arg, n, 1)
                  move_text &= c
               END IF

            NEXT

         ELSE                                 '/ %mapstr_none AND cmd_num1 <> 0

            '/ handle regular non-zero single-columm move here
            '/ reversal code is illegal here, but should have been caught
            '/ reversing one column would not change the move anyway ...

            IF cmd.cmd_num1 > LEN(@select_str) THEN   '/ col requested past end
               move_text_valid = 0
               RETURN
            END IF

            '/ handle single-column end-relative reference

            '/ normal =     123456789
            '/ example:     abcPDQxyz  'x' is end-relative 3, normal 7
            '/ end rel =    987654321

            IF cmd.end_relative = 1 THEN
               cmd.cmd_num1 = (LEN(@select_str) + 1) - cmd.cmd_num1
            END IF

            IF dot_notation_used = 0 THEN     '/ don't mark usage of result str
               usage_flag(cmd.cmd_num1) = 1
            END IF

            col_first = cmd.cmd_num1
            col_final = cmd.cmd_num1
            col_left  = cmd.cmd_num1
            col_right = cmd.cmd_num1
            col_step  = 1

            move_text = MID$(@select_str, cmd.cmd_num1, 1)

         END IF



      CASE %mapstr_plus                    '/ a remainder-columm in the form n+

         '/ a 0+ field is illegal; this should be handled elsewhere
         '/ if n+ is used and n is > field size, ignore it
         '/ a value of \n+ is legal, and copies the chars in reverse order

         IF cmd.cmd_num1 = 0 _
         OR cmd.cmd_num1 > LEN(@select_str) THEN
            move_text_valid = 0
            RETURN
         END IF

         IF cmd.end_relative = 1 THEN                 '/ we should not get this
            move_text_valid = 0
            RETURN
         END IF

         IF reverse_mode = 1 THEN
            col_first = LEN(@select_str)
            col_final = cmd.cmd_num1
            col_step = -1
         ELSE
            col_first = cmd.cmd_num1
            col_final = LEN(@select_str)
            col_step  = 1
         END IF

         col_left  = MIN& (col_first, col_final)
         col_right = MAX& (col_first, col_final)

         FOR n = col_first TO col_final STEP col_step

            '/ we take all data, and mark it as referenced
            '/ if this is not a copying command, don't change the flags
            '/ (references to result str don't reference the source str)

            IF cmd.cmd_type = %mapstr_copying_cmd THEN
               usage_flag(n) = 1
            END IF

            c = MID$(@select_str, n, 1)
            move_text &= c

         NEXT



      CASE %mapstr_dash                                              '/ a range

         '/ we already rejected the case where both numbers are beyond the end
         '/ of the string, but if just one extends past it, we limit it to the
         '/ actual data at hand

         col_first = MIN& (cmd.cmd_num1, LEN(@select_str))
         col_final = MIN& (cmd.cmd_num2, LEN(@select_str))

         IF  col_first = 0 THEN                       '/ somebody used column 0
            err_arg = "Mapping column range "    _
                    & DEC$(cmd.cmd_num1)         _
                    & "-"                        _
                    & DEC$(cmd.cmd_num2)         _
                    & " is invalid"

            FUNCTION = %mapstr_retcode_ERR
            MExitFunc
         END IF

         '/ handle end-relative column range reference
         '/ this translates end-rel numbers to normal numbers

         '/ normal =     123456789
         '/ example:     abcPDQxyz  'x' is end-relative 3, normal 7
         '/ end rel =    987654321

         IF cmd.end_relative = 1 THEN
            col_first = (LEN(@select_str) + 1) - col_first
            col_final = (LEN(@select_str) + 1) - col_final
         END IF


         '/ user is not supposed to say \10-20 to reverse, but we will look the
         '/ other way here; we are being vague about how this gets specified

         IF col_first > col_final THEN         '/ the 'right' way to reverse it
            reverse_mode = 1
            '/ since reverse mode causes STEP -1, we don't need to SWAP here

         ELSEIF reverse_mode = 1 _
         AND col_first < col_final THEN      '/ the bass-wackbards way to do it
            SWAP col_first, col_final          '/ we need to fix it so it works

         END IF

         IF reverse_mode = 1 THEN
            col_step = -1
         ELSE
            col_step  = 1
         END IF

         col_left  = MIN& (col_first, col_final)
         col_right = MAX& (col_first, col_final)


         FOR n = col_first TO col_final STEP col_step

            '/ we take all data, and mark it as referenced
            '/ if this is not a copying command, don't change the flags
            '/ (references to result str don't reference the source str)

            IF cmd.cmd_type = %mapstr_copying_cmd THEN
               usage_flag(n) = 1
            END IF

            c = MID$(@select_str, n, 1)
            move_text &= c

         NEXT

   END SELECT

   IF LEN(move_text) > 0 THEN
      move_text_valid = 1
   ELSE
      move_text_valid = 0
   END IF


   RETURN ' from sub_get_move_text


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_get_dst_cols                    **********************************
'/****   ---------------------------------   **********************************
'/*****************************************************************************

sub_get_dst_cols:

   '/--------------------------------------------------------------------------
   '/ normalize destination columns for Move command M
   '/
   '/ this is much like sub_get_move_text, except that the destination data is
   '/ being deleted, so we don't need to obtain the 'move_text'.
   '/--------------------------------------------------------------------------

   dst_left  = 0                            '/ resolved column range, low value
   dst_right = 0                           '/ resolved column range, high value

   IF  m_cmd.cmd_num1 > LEN(@select_str) _
   AND m_cmd.cmd_num2 > LEN(@select_str) THEN

      dst_cols_valid = 0
      RETURN
   END IF

   SELECT CASE CONST m_cmd.cmd_sep

      CASE %mapstr_none                                      '/ a simple column

         IF m_cmd.cmd_num1 = 0 THEN
               err_arg = "Mapping command " & m_cmd.cmd_name & _
                  " cannot have destination column of 0"

               FUNCTION = %mapstr_retcode_ERR
               MExitFunc
         END IF

         '/ reverse_mode does not apply to a destination

         dst_first = m_cmd.cmd_num1
         dst_final = m_cmd.cmd_num1
         dst_left  = m_cmd.cmd_num1
         dst_right = m_cmd.cmd_num1
         dst_step  = 1

      CASE %mapstr_plus                    '/ a remainder-columm in the form n+

         '/ a 0+ field is illegal; this should be handled elsewhere
         '/ if n+ is used and n is > field size, ignore it
         '/ a value of \n+ is legal, and copies the chars in reverse order

         IF m_cmd.cmd_num1 = 0 _
         OR m_cmd.cmd_num1 > LEN(@select_str) THEN
            dst_cols_valid = 0

            RETURN
         END IF

         IF m_cmd.end_relative = 1 THEN               '/ we should not get this
            dst_cols_valid = 0

            RETURN
         END IF

         '/ reverse_mode does not apply to a destination

         dst_first = m_cmd.cmd_num1
         dst_final = LEN(@select_str)
         dst_step  = 1

         dst_left  = MIN& (dst_first, dst_final)
         dst_right = MAX& (dst_first, dst_final)

      CASE %mapstr_dash                                              '/ a range

         '/ we already rejected the case where both numbers are beyond the end
         '/ of the string, but if just one extends past it, we limit it to the
         '/ actual data at hand

         dst_first = MIN& (m_cmd.cmd_num1, LEN(@select_str))
         dst_final = MIN& (m_cmd.cmd_num2, LEN(@select_str))

         IF  dst_first = 0 THEN                       '/ somebody used column 0
            err_arg = "Mapping column range "    _
                    & DEC$(m_cmd.cmd_num1) _
                    & "-"                        _
                    & DEC$(m_cmd.cmd_num2) _
                    & " is invalid"

            FUNCTION = %mapstr_retcode_ERR
            MExitFunc
         END IF

         '/ handle end-relative column range reference
         '/ this translates end-rel numbers to normal numbers

         '/ normal =     123456789
         '/ example:     abcPDQxyz  'x' is end-relative 3, normal 7
         '/ end rel =    987654321

         IF m_cmd.end_relative = 1 THEN
            dst_first = (LEN(@select_str) + 1) - dst_first
            dst_final = (LEN(@select_str) + 1) - dst_final
         END IF


         '/ reverse_mode does not apply to a destination
         '/ destination columns are deleted, so reversal is not relevant here

         IF dst_first > dst_final THEN
            SWAP dst_first, dst_final
         END IF

         dst_step  = 1
         dst_left  = MIN& (dst_first, dst_final)
         dst_right = MAX& (dst_first, dst_final)

   END SELECT


   IF  dst_left  > 0 _
   AND dst_right > 0 _
   AND dst_right >= dst_left THEN
      dst_cols_valid = 1

   ELSE

      dst_cols_valid = 0
   END IF


   RETURN ' from sub_get_dst_cols


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_resolve_case_CT                 **********************************
'/****   ---------------------------------   **********************************
'/*****************************************************************************

sub_resolve_case_CT:

   '/--------------------------------------------------------------------------
   '/ determine CASE C/T to use.  method: if a C/T code was present on the
   '/ parsed string, use it.  otherwise, use the 'current' C/T code, which was
   '/ set by a CC or CT command, or defaulted from the caller's CT code which
   '/ is supposed to be C or T from the user's edit profile.
   '/
   '/ if this was a string-pair, the parsing logic already reconciled the type
   '/ codes, so that string 1's type will suffice for both sides.
   '/--------------------------------------------------------------------------

   IF cmd.txt_case_CTX_1 <> " " THEN
      this_case_CT = cmd.txt_case_CTX_1 '/ use C/T suffix on string if supplied
   ELSEIF curr_case_CT <> " " THEN
      this_case_CT = curr_case_CT      '/ use C/T from CC or CT cmd, or profile
   ELSE
      this_case_CT = "C"                  '/ should not occur, but failsafe JIC
   END IF


   RETURN ' from sub_resolve_case_CT


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_generate_text_range             **********************************
'/****   ---------------------------------   **********************************
'/*****************************************************************************

sub_generate_text_range:

   '/ character at @cp is [ bracket of inserted text range, which has been
   '/ validated to be in correct format.  append the requested character ranges
   '/ to the result string.  even though validation was done, we will still
   '/ check for nuls to make sure we don't go burserk here ...

   '/ format: [x-y]  [a-b,c-d,x-y]
   '/         01234  01234

   n = 0

   DO
      IF @cp[n+0] = 0 _
      OR @cp[n+1] = 0 _
      OR @cp[n+2] = 0 _
      OR @cp[n+3] = 0 _
      OR @cp[n+4] = 0 THEN
         err_arg = "Mapping text-range operand is malformed"
         FUNCTION = %mapstr_retcode_ERR
         MExitFunc
      END IF

      result_str &= CHR$(@cp[n+1] TO @cp[n+3])
      n += 4
   LOOP UNTIL @cp[n+0] = ASC("]")


   RETURN ' from sub_generate_text_range


END FUNCTION                                                  '/ mapstr_execute
