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


'/-----------------------------------------------------------------------------
'/ mapstr_calc()
'/
'/ evaluate calculation operations
'/ if successful, return %mapstr_retcode_OK else %mapstr_retcode_ERR
'/ on error, diag_arg will contain error message, else ""
'/
'/ method:
'/
'/ the calculator implements a non-recursive precedence-climbing parser.
'/
'/ recursion is avoided by stepping the relative precedence level for each new
'/ nested parenthesis or bracket enclosures, and by biasing right-to-left
'/ operators to force RTL associativity.  values and operators are pushed onto
'/ a stack (a "shift" action, in parsing nomenclature) until the calculated
'/ precedence of a new operator is lower than the last operator pushed on the
'/ stack, at which time the stacks are popped and evaluted (a "reduce" action).
'/ the reduce action is repeated until the stack is empty or until the
'/ remaining operator is of lower precedence than the incoming new operator.
'/
'/ Note dual-use of result_arg.
'/
'/ arguments:
'/
'/    initial_arg       SUPPLIED: starting value for calculator
'/
'/    calc_arg          SUPPLIED: expression string
'/
'/    diag_arg          RETURNED: error messages
'/
'/    result_arg        SUPPLIED: current value of sequence num from CHANGE ALL
'/                      RETURNED: calculated results
'/-----------------------------------------------------------------------------

'ENUM var_enum SINGULAR
'   V_null
'   VA
'   VB
'   VC
'   VD
'   VE
'   VF
'   VG
'   VH
'   VI
'   VJ
'   VK
'   VL
'   VM
'   VN
'   VO
'   VP
'   VQ
'   VR
'   VS
'   VT
'   VU
'   VV
'   VW
'   VX
'   VY
'   VZ
'   V_dummy                                   '/ dummy slot for null set clauses
'   V_num                                      '/ flag for non-var numeric value
'
'   V_min = %VA
'   V_max = %V_num
'END ENUM
'
'
''/ token classes generically describe a token in terms of its functionality
''/ cls types:              {  ;  ,  (  )  [  ]  n  v  U  B  A  S  }
'
'ENUM cls_enum SINGULAR
'
'   cls_null          '/  00        for initializing undefined cls values
'
'   cls_begin         '/  01  {     start of expr
'   cls_semi          '/  02  ;     statement delimiter
'   cls_LP            '/  03  (     parenthesized subexpression
'   cls_RP            '/  04  )     parenthesized subexpression
'   cls_LB            '/  05  [     set expression
'   cls_RB            '/  06  ]     set expression
'   cls_comma         '/  07  ,     set clause delimiter
'   cls_n             '/  08  n     number - decimal or hex
'   cls_v             '/  09  v     variable name
'   cls_U             '/  10  U     unary operator
'   cls_B             '/  11  B     binary operator
'   cls_A             '/  12  A     assignment operator
'   cls_S             '/  13  S     swap operator
'   cls_end           '/  14  }     end of expression
'
'   cls_MIN = %cls_begin
'   cls_MAX = %cls_end
'
'END ENUM
'
'
''/ expression operator tokens
'
'ENUM opr_enum SINGULAR
'
'   opr_null          '/          for initializing undefined opr values
'
'   opr_A_B_SCALE     '/ 01
'   opr_A_B_SL        '/ 02
'   opr_A_B_SR        '/ 03
'   opr_A_L_AND       '/ 04
'   opr_A_L_OR        '/ 05
'   opr_A_L_XOR       '/ 06
'   opr_A_N_SR        '/ 07
'
'   opr_NGTE          '/ 08
'   opr_NLTE          '/ 09
'
'   opr_A_ADD         '/ 10
'   opr_A_B_AND       '/ 11
'   opr_A_B_OR        '/ 12
'   opr_A_B_XOR       '/ 13
'   opr_A_C_MOD       '/ 14
'   opr_A_DIV         '/ 15
'   opr_A_MOD_S       '/ 16
'   opr_A_MUL         '/ 17
'   opr_A_N_SCALE     '/ 18
'   opr_A_SUB         '/ 19
'
'   opr_B_SCALE       '/ 20
'   opr_B_SL          '/ 21
'   opr_B_SR          '/ 22
'   opr_EQ            '/ 23
'   opr_GE            '/ 24
'   opr_LE            '/ 25
'   opr_L_AND         '/ 26
'   opr_L_OR          '/ 27
'   opr_L_XOR         '/ 28
'   opr_MAKE_GE       '/ 29
'   opr_MAKE_LE       '/ 30
'   opr_NE_C          '/ 31
'   opr_NE_S          '/ 32
'   opr_NGT           '/ 33
'   opr_NLT           '/ 34
'   opr_N_SR          '/ 35
'   opr_SWAP          '/ 36
'   opr_U_SGN_X       '/ 37
'   opr_ASSIGN        '/ 38
'
'   opr_ADD           '/ 39
'   opr_ASSIGN_EQ     '/ 40
'   opr_B_AND         '/ 41
'   opr_B_OR          '/ 42
'   opr_B_XOR         '/ 43
'   opr_COMMA         '/ 44
'   opr_C_MOD         '/ 45
'   opr_DIV           '/ 46
'   opr_GT            '/ 47
'   opr_LT            '/ 48
'   opr_L_SET         '/ 49
'   opr_MOD_S         '/ 50
'   opr_MUL           '/ 51
'   opr_N_SCALE       '/ 52
'   opr_R_SET         '/ 53
'   opr_SUB           '/ 54
'   opr_U_ABS         '/ 55
'   opr_U_B_NOT       '/ 56
'   opr_U_L_NOT       '/ 57
'   opr_U_NEG         '/ 58
'   opr_U_POS         '/ 59
'   opr_U_SGN         '/ 60
'
'END ENUM



$mapstr_oplist                = "!#$%&*+,-/;<=>?@\^|~[]:(){}"

%mapstr_init_stack_size       = 20
%mapstr_stack_incr            = 10
%mapstr_incr_prec             = 100



FUNCTION                      mapstr_calc                                     _
(  BYVAL initial_arg          AS QUAD                                         _
,  BYVAL calc_arg             AS STRING                                       _
,  BYREF diag_arg             AS STRING                                       _
,  BYREF result_arg           AS QUAD                                         _
,  BYVAL line_arg             AS LONG                                         _
,  BYVAL col_arg              AS LONG                                         _
   )                          AS LONG                      '/ -> AS RETURN CODE

   '/ misc. items

   LOCAL n                    AS LONG
   LOCAL nn                   AS LONG
   LOCAL set_0                AS LONG         '/ calculating a set-clause index
   LOCAL set_n                AS LONG         '/ calculating a set-clause index
   LOCAL ref_n                AS LONG                      '/ for deref routine
   LOCAL initial_num          AS QUAD
   LOCAL sequence_num         AS QUAD


   '/ variable-storage items

   DIM   V(%V_max)            AS QUAD
   DIM   V_mod(%V_max)        AS BYTE


   '/ expression-scan related items

   LOCAL calc_str             AS STRING
   LOCAL calc_orig            AS STRING                  '/ unmodified calc_str
   LOCAL cp                   AS STRING PTR * 1
   LOCAL lap                  AS STRING PTR * 1                '/ lookahead ptr
   LOCAL assumed_X            AS STRING * 1                 '/ digit for suffix
   LOCAL clause_level         AS LONG
   LOCAL clause_selected      AS LONG
   LOCAL clause_found         AS LONG
   LOCAL var_name             AS STRING * 1
   LOCAL var_index            AS LONG


   '/ get_decimal / get_hex items

   LOCAL num_ndx              AS LONG
   LOCAL num_c                AS STRING * 1
   LOCAL num_suffix_len       AS LONG
   LOCAL num_suffix_max       AS LONG
   LOCAL num_value            AS QUAD
   LOCAL num_string           AS STRING
   LOCAL num_suffix_str       AS STRING
   LOCAL num_radix_desc       AS STRING
   LOCAL H_digit              AS STRING * 1            '/ high digit for suffix
   LOCAL R_digit              AS STRING * 1                 '/ digit for suffix
   LOCAL S_digit              AS STRING * 1                 '/ digit for suffix
   LOCAL num_suffix_type      AS STRING * 1                      '/ suffix code


   '/ precedence, associativity and related items

   LOCAL prev_opr             AS BYTE
   LOCAL curr_opr             AS BYTE

   LOCAL curr_cls             AS BYTE
   LOCAL prev_cls             AS BYTE

   LOCAL curr_opr_prec        AS LONG                       '/ precedence level
   LOCAL curr_opr_nOps        AS BYTE                     '/ number of operands
   LOCAL curr_opr_bias        AS BYTE             '/ bias (0 or 2, where 2=RTL)
   LOCAL curr_opr_base        AS BYTE           '/ base ID for assign operators

   LOCAL base_prec            AS LONG           '/ base precedence-level offset

   LOCAL new_num              AS QUAD
   LOCAL left_num             AS QUAD
   LOCAL right_num            AS QUAD
   LOCAL last_num             AS QUAD         '/ last value stored on the stack

   LOCAL new_type             AS BYTE
   LOCAL left_type            AS BYTE
   LOCAL right_type           AS BYTE


   '/ opr table search items

   DIM   opr_table_first(4)   AS DWORD
   DIM   opr_table_final(4)   AS DWORD
   DIM   opr_table_size(4)    AS DWORD
   LOCAL opr_table_base       AS DWORD

   LOCAL opr_len              AS LONG
   LOCAL opr_table_lo         AS DWORD
   LOCAL opr_table_hi         AS DWORD
   LOCAL opr_table_mid        AS DWORD
   LOCAL opr_table_ptr        AS BYTE PTR

   DIM   opr_byte(32)         AS BYTE AT opr_table_ptr  '/ it's shorter than 32
   DIM   opr_entry(0)         AS STRINGZ * 4 AT opr_table_ptr

   DIM   opr_args(3)          AS STRINGZ * 4      '/ operators being looked for
   LOCAL opr_found            AS STRINGZ * 4         '/ operator found in table


   '/ closure stack items

   DIM   closure_useEQ()      AS BYTE
   DIM   closure_delim()      AS STRING * 1
   DIM   closure_clauses()    AS LONG

   LOCAL closure_right        AS LONG        '/ right-most (highest) used entry
   LOCAL closure_alloc        AS LONG


   '/ opr stack items

   DIM   opr_token()          AS BYTE
   DIM   opr_prec()           AS LONG
   DIM   opr_nOps()           AS LONG
   DIM   opr_cls()            AS BYTE
   DIM   opr_right            AS LONG        '/ right-most (highest) used entry
   DIM   opr_alloc            AS LONG            '/ max allocated size of stack
   DIM   prev_opr_needed      AS BYTE


   '/ val stack items

   DIM   val_type()           AS BYTE             '/ has var ID or flag for num
   DIM   val_num()            AS QUAD
   DIM   val_left             AS LONG      '/ WHERE left operand is for bin ops
   DIM   val_right            AS LONG     '/ WHERE right operand is for bin ops
   DIM   val_alloc            AS LONG            '/ max allocated size of stack


   '/ follows-list items

   LOCAL fol_list_ptr         AS BYTE PTR
   DIM   fol_list(0)          AS ASCIIZ * 32 AT fol_list_ptr


   '/ diagnostic message items

   LOCAL curr_token           AS STRING                    '/ used for messages
   LOCAL curr_pos             AS LONG                 '/ location of curr_token
   LOCAL init_cp              AS DWORD                '/ location of curr_token
   LOCAL curr_cp              AS DWORD                '/ location of curr_token
   LOCAL err_loc              AS STRING


   '/**************************************************************************
   '/**************************************************************************

   MEntry

   '/**************************************************************************
   '/**************************************************************************


   '/--------------------------------------------------------------------------
   '/ validation
   '/ result_arg is used to pass sequence_num, so we have to grab it before
   '/ setting result_arg to the default value equal to the initial_arg.
   '/--------------------------------------------------------------------------

   initial_num  = initial_arg
   sequence_num = result_arg
   result_arg   = initial_arg
   last_num     = initial_arg
   diag_arg     = ""

   calc_str = calc_arg

   '/ if no calculation present, just return initial value

   IF ISNULL (calc_str) _                                        '/ null string
   OR VERIFY (calc_str, " ") = VERIFIED THEN                      '/ only blank
'bug " 355 calc_str is null str"
      FUNCTION = %mapstr_retcode_OK
      MExitFunc
   END IF


   '/--------------------------------------------------------------------------
   '/ user input is not supposed to contain our special start/end markers
   '/ also, reject input if they have blatently mismatched closures
   '/ the ? is used for dummy var references inside set expressions when the
   '/ given set clause is not being selected.
   '/--------------------------------------------------------------------------

   IF TALLY (calc_str, ANY "{}?") <> 0 THEN
      diag_arg = "Calc operand has illegal characters"

   ELSEIF TALLY (calc_str, "(") <> TALLY (calc_str, ")") THEN
      diag_arg = "Calc operand has mismatched () characters"

   ELSEIF TALLY (calc_str, "[") <> TALLY (calc_str, "]") THEN
      diag_arg = "Calc operand has mismatched [] characters"
   END IF

   IF diag_arg <> "" THEN
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc
   END IF


   '/--------------------------------------------------------------------------
   '/ calculator initialization
   '/--------------------------------------------------------------------------

   '/ grab starting addresses of each section in operator table.
   '/ sections 0 and 4 are dummy section addresses just to determine lengths.

   opr_table_base     = CODEPTR (opr_table_0_cp)

   opr_table_first(0) = CODEPTR (opr_table_0_cp)
   opr_table_first(1) = CODEPTR (opr_table_1_cp)
   opr_table_first(2) = CODEPTR (opr_table_2_cp)
   opr_table_first(3) = CODEPTR (opr_table_3_cp)
   opr_table_first(4) = CODEPTR (opr_table_4_cp)

   FOR n = 0 TO 3

      '/ determine size of each section by the difference in addresses
      '/ between start of the next section (n+1) and the start of curr one (n)

      opr_table_size(n) = opr_table_first(n+1) - opr_table_first(n)

      '/ final entry of each section is at start of next one, less the length
      '/ of a single entry, which is found in opr_table_size(0), set during the
      '/ first pass of the FOR loop

      opr_table_final(n) = opr_table_first(n+1) - opr_table_size(0)
   NEXT

   '/ change first/final to subscript form.
   '/ we need to do this because binary search won't work with address offsets,
   '/ but only with subscript numbers, so we calculate equivalent subscripts.

   FOR n = 0 TO 3
      opr_table_first(n) = (opr_table_first(n) - opr_table_base) \ opr_table_size(0)
      opr_table_final(n) = (opr_table_final(n) - opr_table_base) \ opr_table_size(0)
   NEXT

   '/ closure and closure-stack initialization

   closure_right = 0
   closure_alloc = %mapstr_init_stack_size

   REDIM closure_useEQ (closure_alloc)
   REDIM closure_delim (closure_alloc)
   REDIM closure_clauses (closure_alloc)

   closure_delim (0)   = " "
   closure_delim (1)   = " "
   closure_useEQ (0)   = 0
   closure_useEQ (1)   = 0
   closure_clauses (0) = 0


   '/ opr stack initialization

   opr_right = 0
   opr_alloc = %mapstr_init_stack_size

   REDIM opr_token (opr_alloc)
   REDIM opr_prec  (opr_alloc)
   REDIM opr_nOps  (opr_alloc)
   REDIM opr_cls   (opr_alloc)

   opr_token (0) = %opr_null
   opr_prec  (0) = 0
   opr_nOps  (0) = 0
   opr_cls   (0) = 0


   '/ val stack initialization

   val_left  = 0
   val_right = 0
   val_alloc = %mapstr_init_stack_size

   REDIM val_type (val_alloc)
   REDIM val_num  (val_alloc)

   val_type (0) = %V_null
   val_num  (0) = 0


   '/--------------------------------------------------------------------------
   '/ add a trailing ; to last expression if not present, then enclose with {}
   '/ this lets us process the expr in a uniform way even if they leave ; out
   '/ we don't use full TRIM$ because that could throw off col nums in msgs
   '/--------------------------------------------------------------------------

   calc_str = RTRIM$ (calc_str)


   '/--------------------------------------------------------------------------
   '/ if expression begins with an operator, we insert an assumed X
   '/ that makes it possible to have expressions like +5 mean X+5
   '/--------------------------------------------------------------------------

   lap = STRPTR (calc_str)

   DO WHILE @lap = " "
      INCR lap                                       '/ skip any leading blanks
   LOOP

   '/--------------------------------------------------------------------------
   '/ if the first nonblank is a unary operator, we don't assume X
   '/ we need to do a lookahead for !, because if it's followed by < or > it is
   '/ a binary operator, which we will allow and assume X for.
   '/ for operators ~ $ $$ and @ we do NOT assume an X
   '/--------------------------------------------------------------------------

   '/ oper list:    !#$%&*+-/<=>@\^|~:
   '/ assumed X:     # %&*+-/<=> \^| :

   IF VERIFY (@lap, "#%&*+-/<=>\^|:") = VERIFIED THEN
      assumed_X = "X"
   ELSE
      assumed_X = " "

      IF @lap = "!" THEN                       '/ binary !< !> !<= !>= assume X
         IF @lap[1] = "<" _
         OR @lap[1] = ">" THEN
            assumed_X = "X"
         END IF
      END IF
   END IF

   IF RIGHT$ (calc_str, 1) <> ";" THEN
      calc_str &= ";"
   END IF

   calc_str = "{" & assumed_X & UUCASE (calc_str) & " } "


'bug " 500 initial calc_str: |" & calc_str & "|"
   calc_str &= CHR$(0)
   calc_orig = calc_str

   cp = STRPTR (calc_str)
   curr_cp = cp            '/ to help calculate position were an error occurred
   init_cp = cp            '/ to help calculate position were an error occurred
   curr_pos = 1

   FOR n = %V_min TO %V_max
      V(n) = 0                                               '/ clear var array
      V_mod(n) = 0                                  '/ clear var-modified flags
   NEXT

   V(%VX) = initial_num
   V(%VR) = initial_num
   V(%VS) = sequence_num
   V(%VL) = line_arg
   V(%VC) = col_arg

'bug " 520 VX = " & DEC$(V(%VX))
'bug " 521 VR = " & DEC$(V(%VR))
'bug " 522 VS = " & DEC$(V(%VS))
'bug " 523 VL = " & DEC$(V(%VL))
'bug " 524 VC = " & DEC$(V(%VC))

   '/--------------------------------------------------------------------------
   '/ start the parse by allowing it to grab the start token "{"
   '/ if we don't get the start of the expression, get-token has malfunctioned
   '/ we do not validate the follow-list for begin, because there was nothing
   '/ prior to the beginning.  the first cls value after begin is validated as
   '/ being allowed after the "{" symbol.
   '/--------------------------------------------------------------------------

   curr_cls = %cls_null        '/ so first time thru, prev_cls will be cls_null
   curr_opr = %opr_null        '/ so first time thru, prev_opr will be opr_null

   GOSUB sub_get_token                               '/ we must get our { token

   IF curr_cls <> %cls_begin THEN
'bug " 540 invalid curr_cls: " & DEC$(curr_cls)
      diag_arg = "Calc operand internal error at expr start"
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc
   END IF

   base_prec = 0                         '/ adjusted as [] and () levels change


   '/--------------------------------------------------------------------------
   '/ main token processing loop
   '/--------------------------------------------------------------------------

   DO
      GOSUB sub_get_token

      '/-----------------------------------------------------------------------
      '/ determine follows-list from the cls of the prior token cls
      '/ we MUST use the prev_cls to test this, NOT the curr_cls.
      '/
      '/ that is because the follows-list are crafted with respect to what
      '/ follows them.  since the curr_cls follows the prev_cls, we are
      '/ asking if what we have now (curr_cls) can follow what preceded it
      '/ (prev_cls).  otherwise, we'd have to ask if the curr_cls was allowed
      '/ to be preceded by the prev_cls, which is a backwards kind of question
      '/ to be asking.
      '/-----------------------------------------------------------------------

'bug " 568 prev_cls: " & DEC$(prev_cls)
      SELECT CASE AS CONST prev_cls                           '/ see note above

         CASE %cls_begin:  fol_list_ptr = CODEPTR (fol_begin_cp)
'bug " 572 fol_list_ptr = fol_begin_cp   "
         CASE %cls_semi:   fol_list_ptr = CODEPTR (fol_semi_cp)
'bug " 574 fol_list_ptr = fol_semi_cp    "
         CASE %cls_comma:  fol_list_ptr = CODEPTR (fol_comma_cp)
'bug " 576 fol_list_ptr = fol_comma_cp   "
         CASE %cls_LP:     fol_list_ptr = CODEPTR (fol_LP_cp)
'bug " 578 fol_list_ptr = fol_LP_cp      "
         CASE %cls_RP:     fol_list_ptr = CODEPTR (fol_RP_cp)
'bug " 580 fol_list_ptr = fol_RP_cp      "
         CASE %cls_LB:     fol_list_ptr = CODEPTR (fol_LB_cp)
'bug " 582 fol_list_ptr = fol_LB_cp      "
         CASE %cls_RB:     fol_list_ptr = CODEPTR (fol_RB_cp)
'bug " 584 fol_list_ptr = fol_RB_cp      "
         CASE %cls_n:      fol_list_ptr = CODEPTR (fol_n_cp)
'bug " 586 fol_list_ptr = fol_n_cp       "
         CASE %cls_v:      fol_list_ptr = CODEPTR (fol_v_cp)
'bug " 588 fol_list_ptr = fol_v_cp       "
         CASE %cls_U:      fol_list_ptr = CODEPTR (fol_U_cp)
'bug " 590 fol_list_ptr = fol_U_cp       "
         CASE %cls_B:      fol_list_ptr = CODEPTR (fol_B_cp)
'bug " 592 fol_list_ptr = fol_B_cp       "
         CASE %cls_A:      fol_list_ptr = CODEPTR (fol_A_cp)
'bug " 594 fol_list_ptr = fol_A_cp       "
         CASE %cls_S:      fol_list_ptr = CODEPTR (fol_S_cp)
'bug " 596 fol_list_ptr = fol_S_cp       "

         '/ if prev_cls is not set properly, we have an error somewhere

         CASE ELSE:
'bug " 601 invalid prev_cls: " & DEC$(prev_cls)
            diag_arg = err_loc & "Calc operand internal error in follows list"
            FUNCTION = %mapstr_retcode_ERR
            MExitFunc

      END SELECT

      '/-----------------------------------------------------------------------
      '/ verify that the current token class can legitimately follow the
      '/ previous one.  if not, there is a syntax error in the expression.
      '/
      '/ the 'follows list' is encoded in ASM statements like STRINGZ values.
      '/ we do a VERIFY treating the current CLS like a one-character string.
      '/ if curr_cls is not in the 'string', it's an error.
      '/-----------------------------------------------------------------------

      REDIM fol_list(0) AS ASCIIZ * 32 AT fol_list_ptr
'bug " 618 fol_list = " & showstr(fol_list(0))

      IF VERIFY (CHR$(curr_cls), fol_list(0)) <> VERIFIED THEN
'bug " 621 prev_cls: " & DEC$(prev_cls)
'bug " 622 curr_cls: " & DEC$(curr_cls)
         diag_arg = err_loc & "Calc operand has syntax error at " & curr_token
         FUNCTION = %mapstr_retcode_ERR
         MExitFunc
      END IF


      '/-----------------------------------------------------------------------
      '/ select action based on current token class
      '/-----------------------------------------------------------------------

'bug " 633 select action, curr_cls = " & DEC$(curr_cls)
      SELECT CASE AS CONST curr_cls

         CASE %cls_semi                                                    '/ ;
'bug " 637 got cls_semi"
            GOSUB sub_evaluate_stack
            closure_useEQ(closure_right) = 0    '/ allow plain = operator again

         CASE %cls_comma                                                   '/ ,
'bug " 642 got cls_comma"
            GOSUB sub_evaluate_stack
            closure_useEQ(closure_right) = 0    '/ allow plain = operator again

         CASE %cls_LP                                                      '/ (
'bug " 647 got cls_LP"
            base_prec += %mapstr_incr_prec
'bug " 649 base_prec = " & DEC$(base_prec)

         CASE %cls_RP                                                      '/ )
'bug " 652 got cls_RP"
            base_prec -= %mapstr_incr_prec
'bug " 654 base_prec = " & DEC$(base_prec)

         CASE %cls_LB                                                      '/ [
'bug " 657 got cls_LB"
            GOSUB sub_evaluate_stack
            GOSUB sub_dereference_val_ref

            '/ determine set clause selector value
            '/ if the evaluated selector value is < 1 or more than the number
            '/ of set clauses, the value is forced to the number of clauses.
            '/ that provides the default-clause mechanism.

            IF val_num (val_right) < 1 _
            OR val_num (val_right) > closure_clauses(closure_right) THEN
'bug " 670 LB took default: selector = " & DEC$(val_num (val_right)) & " limit = " & DEC$(closure_clauses(closure_right))
'bug " 671 closure_right = " & DEC$(closure_right)
               clause_selected = closure_clauses(closure_right)
               val_num (val_right) = clause_selected     '/ adj. value on stack
               val_type (val_right) = %V_num      '/ new value is a num literal
            ELSE
'bug " 675 LB took request: selector = " & DEC$(val_num (val_right)) & " limit = " & DEC$(closure_clauses(closure_right))
'bug " 676 closure_right = " & DEC$(closure_right)
               clause_selected = val_num (val_right)
            END IF

            GOSUB sub_divert_set_vars
            base_prec += %mapstr_incr_prec
'bug " 682 base_prec = " & DEC$(base_prec)

         CASE %cls_RB                                                      '/ ]
'bug " 685 got cls_RB"

            GOSUB sub_evaluate_stack

            '/-----------------------------------------------------------------
            '/ execute a set selector operation
            '/
            '/ the number of set clauses is stored in
            '/ closure_clauses(closure_right).  for example, in a set expr of
            '/ x>5 [123,456] the closure_clauses count is 2.  on the val_num
            '/ stack, the right-most contents are ? 123 456, where ? is the
            '/ set selector, which has been adjusted to a value between 1 and
            '/ the number of set clauses, which is 2 in the example.
            '/
            '/ to do a set selection, we must do the following:
            '/ calculate a relative-0 index into the val_num stack.  that will
            '/ be equal to val_right minus the closure_clauses count.
            '/ the selector is used to copy the selected value to the relative
            '/ zero location. the stack is shortened by the number of set
            '/ clauses, leaving only the selected value remaining.  finally,
            '/ to make the return-value process work correctly, the selected
            '/ value must be saved in the last_num variable.
            '/-----------------------------------------------------------------

            nn = closure_clauses(closure_right)
'bug " 710 closure_clauses(closure_right) = " & DEC$(nn)
'bug " 711 val_right = " & DEC$(val_right)
            set_0 = val_right - nn              '/ relative 0 location in stack
'bug " 713 set_0 = " & DEC$(set_0)

            GOSUB sub_dereference_val_ref

'bug " 718 val_num (set_0) = " & DEC$(val_num (set_0))
            set_n = set_0 + val_num (set_0)     '/ form index to selected value
'bug " 720 set_n = " & DEC$(set_n)
'bug " 721 val_num (set_n) = " & DEC$(val_num (set_n))

            val_num (set_0) = val_num (set_n)            '/ copy selected value
            val_type (set_0) = %V_num          '/ value of set is a num literal

            DO WHILE nn > 0
               GOSUB sub_prev_val
               DECR nn
            LOOP

            last_num = val_num (set_0)
            base_prec -= %mapstr_incr_prec
            GOSUB sub_prev_closure
'bug " 734 base_prec = " & DEC$(base_prec)

         CASE %cls_n                                 '/ number - decimal or hex
'bug " 737 got cls_n"
            GOSUB sub_next_val                '/ push new value on to val stack
'bug " 739 sub_next_val - number " & DEC$(num_value) & " at stack " & DEC$(val_right)
            val_type (val_right) = %V_num
            val_num  (val_right) = num_value
            last_num = num_value         '/ save value in case R/X vars not set

         '/ when a variable is pushed on the stack, we don't store its value.
         '/ rather, we note its 'name' and then dereference it when needed for
         '/ its value, and use it directly when it is the target of a swap or
         '/ assignment operator.  val_num (val_right) is not set here.

         CASE %cls_v                                           '/ variable name
            GOSUB sub_next_val                '/ push new value on to val stack
'bug " 751 got cls_v - " & var(var_index) & " at stack " & DEC$(val_right)
            val_type (val_right) = var_index         '/ var identifier, 1 to 26

         CASE %cls_U                                          '/ unary operator
'bug " 755 got cls_U"
            GOTO CASE_EVALUATE

         CASE %cls_B                                         '/ binary operator
'bug " 759 got cls_B"
            GOTO CASE_EVALUATE

         CASE %cls_A                                     '/ assignment operator
'bug " 763 got cls_A"
            GOTO CASE_EVALUATE

         CASE %cls_S                                           '/ swap operator
'bug " 767 got cls_S"

         CASE_EVALUATE:

            GOSUB sub_evaluate_stack
            GOSUB sub_next_opr

            opr_token(opr_right) = curr_opr
            opr_prec (opr_right) = curr_opr_prec
            opr_nOps (opr_right) = curr_opr_nOps
            opr_cls  (opr_right) = curr_cls
'bug " 778 opr_right at stack " & DEC$(opr_right)

    ' curr_opr_bias = opr_byte(8)                 '/ bias (0 or 2, where 2=RTL)
    ' curr_opr_base = opr_byte(9)               '/ base ID for assign operators

         CASE %cls_end                                   '/ } end of expression
'bug " 784 got cls_end"

         '/--------------------------------------------------------------------
         '/ pass back computed value to caller
         '/--------------------------------------------------------------------

         IF V_mod (%VR) = 1 THEN                          '/ variable R was set
'bug " 791 passing back VR"
            result_arg = V(%VR)
         ELSEIF V_mod (%VX) = 1 THEN                      '/ variable X was set
'bug " 794 passing back VX"
            result_arg = V(%VX)
         ELSE
'bug " 797 passing back last_num"
            result_arg = last_num                      '/ last calculated value
         END IF

         diag_arg = ""
         FUNCTION = %mapstr_retcode_OK
         MExitFunc

      END SELECT

   LOOP

'bug " 809 reached default MExitFunc - error"
   MExitFunc


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_evaluate_stack                  **********************************
'/****   ---------------------------------   **********************************
'/*****************************************************************************

sub_evaluate_stack:
'bug " 820 entered sub_evaluate_stack"

   '/ if the last operator on the stack is of higher precedence than the new
   '/ operator, the operator and its operands must be popped, evaluated, and
   '/ the results stored back on the stack.
'bug " 825 prec: right = " & DEC$(opr_prec (opr_right)) & " curr = " & DEC$(curr_opr_prec) & " bias " & DEC$(curr_opr_bias)

   DO WHILE opr_right > 0 _
        AND opr_prec (opr_right) + 1 > curr_opr_prec + curr_opr_bias

'bug " 830 prec: right = " & DEC$(opr_prec (opr_right)) & " curr = " & DEC$(curr_opr_prec) & " bias " & DEC$(curr_opr_bias)
      GOSUB sub_get_operands
      GOSUB sub_exec_operator
   LOOP


   RETURN ' from sub_evaluate_stack


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_dereference_val_ref             **********************************
'/****   ---------------------------------   **********************************
'/*****************************************************************************

sub_dereference_val_ref:

   '/ the val_num slot for a register is normally unused.  we dereference it
   '/ to make it easier for other code to use the value on the stack, without
   '/ going through a formal get_operand process.

   FOR ref_n = 1 TO val_right
      IF val_type (ref_n) <> %V_num THEN
         val_num (ref_n) = V(val_type (ref_n))
      END IF
   NEXT


   RETURN ' from sub_dereference_val_ref


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_get_operands                    **********************************
'/****   ---------------------------------   **********************************
'/*****************************************************************************

sub_get_operands:

'bug " 872 entered sub_get_operands - nOps = " & DEC$ (opr_nOps (opr_right))

   IF opr_nOps (opr_right) = 0 THEN
'bug " 875 sub_get_operands nOps is zero"
      diag_arg = err_loc & "sub_get_operands internal error"
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc
   END IF


   right_type = val_type (val_right)              '/ determine if var or number

   IF right_type = %V_num THEN                                        '/ number
      right_num = val_num (val_right)
'bug " 886 right_num literal = " & DEC$(right_num) & " at stack " & DEC$(val_right)
   ELSE
      right_num = V(right_type)           '/ variable - fetch its current value
      val_num (val_right) = right_num          '/ do a dereference at this time
'bug " 889 right_num " & var(right_type) & " = " & DEC$(right_num) & " at stack " & DEC$(val_right)
   END IF

   IF opr_nOps (opr_right) = 2 THEN          '/ a two-operand (binary) operator
      left_type = val_type (val_left)             '/ determine if var or number

      IF left_type = %V_num THEN                                      '/ number
         left_num = val_num (val_left)
'bug " 897 left_num literal = " & DEC$(left_num) & " at stack " & DEC$(val_left)
      ELSE
         left_num = V(left_type)          '/ variable - fetch its current value
         val_num (val_left) = left_num         '/ do a dereference at this time
'bug " 900 left_num " & var(left_type) & " = " & DEC$(left_num) & " at stack " & DEC$(val_left)
      END IF
      GOSUB sub_prev_val         '/ release (pop) slot stack for 2 arg operator
   END IF
   GOSUB sub_prev_val       '/ release (pop) slot stack for 1 or 2 arg operator


   RETURN ' from sub_get_operands


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_exec_operator                   **********************************
'/****   ---------------------------------   **********************************
'/*****************************************************************************

sub_exec_operator:

   '/ set common result values

   new_num = 0
   new_type = left_type
   prev_opr_needed = 1
'bug " 923 exec opr = " & DEC$(opr_token(opr_right)) & " at stack " & DEC$(opr_right)
   SELECT CASE AS CONST opr_token(opr_right)

      CASE %opr_ASSIGN                          '/  %cls_A  :=  EXPLICIT ASSIGN
         new_num = right_num
'bug " 928 exec := new_num = " & DEC$(new_num)

      CASE %opr_ADD _                                       '/  %cls_B  +   ADD
      ,    %opr_A_ADD                                       '/  %cls_A  +=  ADD
         new_num = left_num + right_num
'bug " 933 exec + new_num = " & DEC$(new_num)


      CASE %opr_B_AND _                             '/  %cls_B  &   BITWISE AND
      ,    %opr_A_B_AND                             '/  %cls_A  &=  BITWISE AND
         new_num = left_num AND right_num
'bug " 939 exec & new_num = " & DEC$(new_num)


      CASE %opr_B_OR _                               '/  %cls_B  |   BITWISE OR
      ,    %opr_A_B_OR                               '/  %cls_A  |=  BITWISE OR
         new_num = left_num OR right_num
'bug " 945 exec | new_num = " & DEC$(new_num)


      CASE %opr_B_XOR _                             '/  %cls_B  ^   BITWISE XOR
      ,    %opr_A_B_XOR                             '/  %cls_A  ^=  BITWISE XOR
         new_num = left_num XOR right_num
'bug " 951 exec ^ new_num = " & DEC$(new_num)


      CASE %opr_B_SL _                       '/  %cls_B  <<  BITWISE SHIFT LEFT
      ,    %opr_A_B_SL                       '/  %cls_A  <<= BITWISE SHIFT LEFT
         '/ shifting left with negative amt is shifting right with ABS(amt)

         IF right_num < 0 THEN
            right_num = ABS (right_num)
            GOTO CASE_RIGHT
         END IF

      CASE_LEFT:
         right_num = MIN (right_num, 63)              '/ limit shift to 63 bits

         IF right_num > 0 THEN
            SHIFT LEFT left_num, right_num
         END IF

         new_num = left_num
'bug " 971 exec << new_num = " & DEC$(new_num)


      CASE %opr_B_SR _                      '/  %cls_B  >>  BITWISE SHIFT RIGHT
      ,    %opr_A_B_SR                      '/  %cls_A  >>= BITWISE SHIFT RIGHT
         '/ shifting right with negative amt is shifting left with ABS(amt)

         IF right_num < 0 THEN
            right_num = ABS (right_num)
            GOTO CASE_LEFT
         END IF

      CASE_RIGHT:
         right_num = MIN (right_num, 63)              '/ limit shift to 63 bits

         IF right_num > 0 THEN
            SHIFT RIGHT left_num, right_num
         END IF

         new_num = left_num
'bug " 991 exec >> new_num = " & DEC$(new_num)


      CASE %opr_N_SR _                      '/  %cls_B  +>  NUMERIC SHIFT RIGHT
      ,    %opr_A_N_SR                      '/  %cls_A  +>= NUMERIC SHIFT RIGHT
         '/ shifting right with negative amt is shifting left with ABS(amt)

         IF right_num < 0 THEN
            right_num = ABS (right_num)
            GOTO CASE_LEFT
         END IF

         right_num = MIN (right_num, 63)              '/ limit shift to 63 bits

         IF right_num > 0 THEN
            SHIFT SIGNED RIGHT left_num, right_num
         END IF

         new_num = left_num
'bug "1010 exec +> new_num = " & DEC$(new_num)


      '/-----------------------------------------------------------------------
      '/ we make sure divide and mod are arithmetically correct
      '/
      '/ examples (x / y) = quot and rem
      '/
      '/ +7 / +2 = +3 rem +1   because (+2 * +3) + (+1) = +7
      '/
      '/ +7 / -2 = -3 rem +1   because (-2 * -3) + (+1) = +7
      '/
      '/ -7 / +2 = -3 rem -1   because (+2 * -3) + (-1) = -7
      '/
      '/ -7 / -2 = +3 rem -1   because (-2 * +3) + (-1) = -7
      '/
      '/ so, quotient  = (abs(x) DIV abs(y)) * (sgn(x) * sgn(y))
      '/     remainder = (abs(x) MOD abs(y)) *  sgn(x)
      '/-----------------------------------------------------------------------


      CASE %opr_DIV _                                    '/  %cls_B  /   DIVIDE
      ,    %opr_A_DIV                                    '/  %cls_A  /=  DIVIDE

         IF left_type = %V_dummy THEN    '/ operator is \= or %= with dummy var
            new_type = %V_dummy                     '/ don't bother with divide
'bug "1036 exec / gets dummy"

         ELSEIF right_num = 0 THEN
            diag_arg = err_loc & "Calc operand has zero divisor"
            FUNCTION = %mapstr_retcode_ERR
            MExitFunc

         ELSE
            new_num = (ABS(left_num) \ ABS(right_num)) _
                    * (SGN(left_num) * SGN(right_num))
'bug "1046 exec / new_num = " & DEC$(new_num)
         END IF


      CASE %opr_MOD_S   _                  '/  %cls_B  \   MODULUS - STD SYNTAX
      ,    %opr_A_MOD_S _                  '/  %cls_A  \=  MODULUS - STD SYNTAX
      ,    %opr_C_MOD   _                    '/  %cls_B  %   MODULUS - C SYNTAX
      ,    %opr_A_C_MOD                      '/  %cls_A  %=  MODULUS - C SYNTAX

         IF left_type = %V_dummy THEN    '/ operator is \= or %= with dummy var
            new_type = %V_dummy               '/ don't bother with MOD operator
'bug "1057 exec \ gets dummy"

         ELSEIF right_num = 0 THEN
            diag_arg = err_loc & "Calc operand has zero divisor in modulus"
            FUNCTION = %mapstr_retcode_ERR
            MExitFunc

         ELSE
            new_num = (ABS(left_num) MOD ABS(right_num)) * SGN(left_num)
         END IF
'bug "1067 exec \ new_num = " & DEC$(new_num)


      CASE %opr_L_AND _                             '/  %cls_B  &&  LOGICAL AND
      ,    %opr_A_L_AND                             '/  %cls_A  &&= LOGICAL AND
         IF left_num <> 0 AND right_num <> 0 THEN
            new_num = 1
         END IF
'bug "1075 exec && new_num = " & DEC$(new_num)


      CASE %opr_L_OR _                               '/  %cls_B  ||  LOGICAL OR
      ,    %opr_A_L_OR                               '/  %cls_A  ||= LOGICAL OR
         IF left_num <> 0 OR right_num <> 0 THEN
            new_num = 1
         END IF
'bug "1083 exec || new_num = " & DEC$(new_num)


      CASE %opr_L_XOR _                             '/  %cls_B  ^^  LOGICAL XOR
      ,    %opr_A_L_XOR                             '/  %cls_A  ^^= LOGICAL XOR
         IF left_num <> 0 THEN
            new_num = 1
         END IF

         IF right_num <> 0 THEN
            new_num XOR= 1
         END IF
'bug "1095 exec ^^ new_num = " & DEC$(new_num)


      CASE %opr_MUL _                                  '/  %cls_B  *   MULTIPLY
      ,    %opr_A_MUL                                  '/  %cls_A  *=  MULTIPLY
         new_num = left_num * right_num
'bug "1101 exec * new_num = " & DEC$(new_num)


      CASE %opr_N_SCALE _                         '/  %cls_B  #   NUMERIC SCALE
      ,    %opr_A_N_SCALE                         '/  %cls_A  #=  NUMERIC SCALE
         IF right_num > 1 THEN
            DO
               left_num \= right_num
               INCR new_num
            LOOP UNTIL left_num = 0
         END IF
'bug "1112 exec # new_num = " & DEC$(new_num)


      CASE %opr_B_SCALE _                         '/  %cls_B  ##  BITWISE SCALE
      ,    %opr_A_B_SCALE                         '/  %cls_A  ##= BITWISE SCALE
         IF right_num >= 1 AND right_num <= 63 THEN
            DO
               SHIFT RIGHT left_num, right_num
               INCR new_num
            LOOP UNTIL left_num = 0
         END IF
'bug "1123 exec ## new_num = " & DEC$(new_num)


      CASE %opr_SUB _                                  '/  %cls_B  -   SUBTRACT
      ,    %opr_A_SUB                                  '/  %cls_A  -=  SUBTRACT
         new_num = left_num - right_num
'bug "1129 exec - new_num = " & DEC$(new_num)


      CASE %opr_NE_S _                '/  %cls_B  <>  NOT EQUAL TO - STD SYNTAX
      ,    %opr_NE_C                    '/  %cls_B  !=  NOT EQUAL TO - C SYNTAX
         IF left_num <> right_num THEN
            new_num = 1
         END IF
'bug "1137 exec <> new_num = " & DEC$(new_num)


      CASE %opr_EQ                          '/  %cls_B  ==  EQUAL TO - C SYNTAX
         IF left_num = right_num THEN
            new_num = 1
         END IF
'bug "1144 exec == new_num = " & DEC$(new_num)


      CASE %opr_GE                     '/  %cls_B  >=  GREATER THAN OR EQUAL TO
         IF left_num >= right_num THEN
            new_num = 1
         END IF
'bug "1151 exec >= new_num = " & DEC$(new_num)


      CASE %opr_GT                                 '/  %cls_B  >   GREATER THAN
         IF left_num > right_num THEN
            new_num = 1
         END IF
'bug "1158 exec > new_num = " & DEC$(new_num)


      CASE %opr_LE                        '/  %cls_B  <=  LESS THAN OR EQUAL TO
         IF left_num <= right_num THEN
            new_num = 1
         END IF
'bug "1165 exec <= new_num = " & DEC$(new_num)


      CASE %opr_LT                                    '/  %cls_B  <   LESS THAN
         IF left_num < right_num THEN
            new_num = 1
         END IF
'bug "1172 exec < new_num = " & DEC$(new_num)


      CASE %opr_NGT                            '/  %cls_B  !>  NOT GREATER THAN

         '/--------------------------------------------------------------------
         '/  Not Greater Than.
         '/
         '/  In the expression X !> Y, the intent is to limit the value of X
         '/  such that it is not greater than Y. If X > Y, the value
         '/  of X !> Y is Y, otherwise it is X.
         '/--------------------------------------------------------------------

         IF left_num > right_num THEN
            new_num = right_num
         ELSE
            new_num = left_num
         END IF
'bug "1190 exec !> new_num = " & DEC$(new_num)


      CASE %opr_NLT                               '/  %cls_B  !<  NOT LESS THAN

         '/--------------------------------------------------------------------
         '/  Not Less Than.
         '/
         '/  In the expression X !< Y, the intent is to limit the value of X
         '/  such that it is not less than Y. If X < Y, the value
         '/  of X !< Y is Y, otherwise it is X.
         '/--------------------------------------------------------------------

         IF left_num < right_num THEN
            new_num = right_num
         ELSE
            new_num = left_num
         END IF
'bug "1208 exec !< new_num = " & DEC$(new_num)


      CASE %opr_NGTE               '/  %cls_B  !>= NOT GREATER THAN OR EQUAL TO

         '/--------------------------------------------------------------------
         '/ Not Greater Than or Equal To
         '/
         '/ In the expression X !>= Y, the intent is to limit the value of X
         '/ such that it is not greater than or equal to Y.
         '/
         '/ If X >= Y, the value of X !>= Y is Y–1 .
         '/ If X is not >= Y, the value of X !>= Y is X.
         '/--------------------------------------------------------------------

         IF left_num >= right_num THEN
            new_num = right_num - 1
         ELSE
            new_num = left_num
         END IF
'bug "1228 exec !>= new_num = " & DEC$(new_num)


      CASE %opr_NLTE                  '/  %cls_B  !<= NOT LESS THAN OR EQUAL TO

         '/--------------------------------------------------------------------
         '/ Not Less Than or Equal To
         '/
         '/ In the expression X !<= Y, the intent is to limit the value of X
         '/ such that it is not less than or equal to Y.
         '/
         '/ If X <= Y, the value of X !<= Y is Y+1.
         '/ If X is not <= Y, the value of X !=< Y is X.
         '/--------------------------------------------------------------------

         IF left_num <= right_num THEN
            new_num = right_num + 1
         ELSE
            new_num = left_num
         END IF
'bug "1248 exec !<= new_num = " & DEC$(new_num)


      CASE %opr_MAKE_LE                     '/  %cls_S  <:  X <: Y MAKES X <= Y

         '/--------------------------------------------------------------------
         '/ Conditional swap variables for <=
         '/
         '/ Both the left-hand operand and the right-hand operand must be
         '/ variables.  The operator implements an ordering function.
         '/ In the expression X <: Y, the intent is to conditionally swap the
         '/ contents of the two variables, if necessary, so that the
         '/ relationship  X <= Y holds true.  If X <= Y is already true,
         '/ no swap occurs.  If X <= Y is false, a swap occurs.
         '/--------------------------------------------------------------------

         IF V(left_type) > V(right_type) THEN               '/  X <= Y is false
            GOTO CASE_SWAP
         END IF

         new_num  = V(left_type)
         new_type = left_type
'bug "1270 exec <: new_num = " & DEC$(new_num)


      CASE %opr_MAKE_GE                     '/  %cls_S  >:  X >: Y MAKES X >= Y

         '/--------------------------------------------------------------------
         '/ Conditional swap variables for >=
         '/
         '/ Both the left-hand operand and the right-hand operand must be
         '/ variables.  The operator implements an ordering function.
         '/ In the expression X >: Y, the intent is to conditionally swap the
         '/ contents of the two variables, if necessary, so that the
         '/ relationship  X >= Y holds true.  If X >= Y is already true,
         '/ no swap occurs.  If X >= Y is false, a swap occurs.
         '/--------------------------------------------------------------------

         IF V(left_type) < V(right_type) THEN               '/  X >= Y is false
            GOTO CASE_SWAP
         END IF

         new_num  = V(left_type)
         new_type = left_type
'bug "1292 exec >: new_num = " & DEC$(new_num)


      CASE %opr_SWAP                             '/  %cls_S  ::  SWAP VARIABLES
      CASE_SWAP:
         SWAP V(left_type), V(right_type)

         V_mod(left_type)  = 1                  '/ both variables were modified
         V_mod(right_type) = 1

         '/ a statement of A :: B makes the value of the expression equal to B
         '/ so, the new_num is the value of B and the new type is the type of A

         new_num  = V(left_type)
         new_type = left_type
'bug "1309 exec :: new_num = " & DEC$(new_num)


      CASE %opr_U_ABS                                 '/  %cls_U  @   UNARY ABS
         new_num = ABS(right_num)
         new_type = %V_num
'bug "1315 exec u@ new_num = " & DEC$(new_num)


      CASE %opr_U_B_NOT                  '/  %cls_U  ~   UNARY BITWISE NEGATION
         new_num = NOT(right_num)
         new_type = %V_num
'bug "1321 exec u~ new_num = " & DEC$(new_num)


      CASE %opr_U_L_NOT                       '/  %cls_U  !   UNARY LOGICAL NOT
         IF right_num = 0 THEN
            new_num = 1
         END IF
         new_type = %V_num
'bug "1329 exec u! new_num = " & DEC$(new_num)


      CASE %opr_U_NEG                              '/  %cls_U  -   UNARY NEGATE
         new_num = -right_num
         new_type = %V_num
'bug "1335 exec u- new_num = " & DEC$(new_num)


      CASE %opr_U_POS                            '/  %cls_U  +   UNARY POSITIVE
         new_num = right_num                                   '/ this is a NOP
         new_type = %V_num
'bug "1341 exec u+ new_num = " & DEC$(new_num)


      CASE %opr_U_SGN                                 '/  %cls_U  $   UNARY SGN
         new_num = SGN(right_num)
         new_type = %V_num
'bug "1347 exec u$ new_num = " & DEC$(new_num)


      CASE %opr_U_SGN_X                   '/  %cls_B  $$  UNARY SGN/INDEX = 2+$
         new_num = 2 + SGN(right_num)
         new_type = %V_num
'bug "1353 exec u$$ new_num = " & DEC$(new_num)

      CASE ELSE                                 '/ perhaps a comma or semicolon
         prev_opr_needed = 0   '/ don't pop opr because these are not operators
         new_type = %V_num

   END SELECT

   '/ store number back

   GOSUB sub_next_val                   '/ make room for new value on val stack

   val_type (val_right) = new_type
   val_num (val_right) = new_num                 '/ overall value of expression

   IF opr_cls (opr_right) = %cls_A THEN                '/ this is an assignment
      V(new_type) = new_num                   '/ variable - store its new value
      V_mod(new_type) = 1                          '/ mark variable as modified
'bug "1372 exec storing " & var(new_type) & " as new_num = " & DEC$(new_num) & " at stack " & DEC$(val_right)

   ELSEIF opr_cls (opr_right) = %cls_S THEN
      last_num = new_num                 '/ save value in case R/X vars not set

   ELSE
      last_num = new_num                 '/ save value in case R/X vars not set
      val_type (val_right) = %V_num                   '/ this is an expr result
'bug "1377 exec storing new_num = " & DEC$(new_num) & " at stack " & DEC$(val_right)

   END IF

   IF prev_opr_needed THEN
      GOSUB sub_prev_opr         '/ operator was processed, so pop it off stack
   END IF


   RETURN ' from sub_exec_operator


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_get_token                       **********************************
'/****   ---------------------------------   **********************************
'/*****************************************************************************

sub_get_token:

   curr_token = ""                                              '/ for messages

   DO WHILE @cp = " "
      INCR cp                                               '/ skip over blanks
   LOOP

   '/--------------------------------------------------------------------------
   '/ whatever non-blank we found has to be the beginning of some token
   '/ we use that as the value reported as the "position" where expr is bad
   '/ because of the "{ " prefix, the location will be off by two.
   '/ we don't want to mangle curr_pos, but we change in in the msg.
   '/--------------------------------------------------------------------------

   curr_cp  = cp
   curr_pos = (curr_cp - init_cp) + 1
   err_loc  = "(" & DEC$(curr_pos-2) & "): " _

   '/--------------------------------------------------------------------------
   '/ process-token loop should halt at "}".  if we get to a nul terminator,
   '/ the process-token logic has failed, and we need to halt with an error msg
   '/--------------------------------------------------------------------------

   IF @cp = CHR$(0) THEN
      diag_arg = "Calc operand internal error at expr terminator"
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc
   END IF

   opr_table_ptr = 0           '/ so we know if operator was found in the table
   opr_len = 0
   opr_found = ""

   prev_cls = curr_cls
   curr_cls = %cls_null                          '/ until we identify the input

   prev_opr = curr_opr
   curr_opr = %opr_null

   '/--------------------------------------------------------------------------
   '/ handle var-name tokens
   '/--------------------------------------------------------------------------

   IF @cp >= "A" AND @cp <= "Z" THEN
      curr_token = @cp
      var_name = @cp
      INCR cp                                               '/ skip over letter

      '/ var name cannot be followed by letter or digit

      IF (@cp >= "0" AND @cp <= "9") _
      OR (@cp >= "A" AND @cp <= "Z") THEN
         diag_arg = err_loc & "Calc operand variable name "  _
            & var_name & @cp & " invalid"

         FUNCTION = %mapstr_retcode_ERR
         MExitFunc
      END IF

      curr_cls = %cls_v                               '/ token class = variable
      var_index = ASC(var_name) - ASC("@")                     '/ "A" - "@" = 1

'bug "1458 got cls_v - " & var(var_index)
      RETURN
   END IF

   '/--------------------------------------------------------------------------
   '/ handle fake var-name ? tokens
   '/
   '/ a ? is inserted into variable references inside set clauses, when the
   '/ given set clause is determined not the the selected clause.  by doing
   '/ this, any set clauses that had assignment or swap operators will not
   '/ affect any variables.  we cannot optimize away these null clauses, but we
   '/ can prevent any undesired side effects.  we use the calc_orig string to
   '/ get the original non-modified var name in case they did something wrong.
   '/ that way, using fake var names doesn't prevent us from issuing accurate
   '/ error messages if they have a syntax error in a null set clause.
   '/--------------------------------------------------------------------------

   IF @cp = "?" THEN
      curr_token = MID$ (calc_orig, curr_pos, 1)
      var_name = curr_token
      INCR cp                                                    '/ skip over ?

      '/ var name cannot be followed by letter or digit

      IF (@cp >= "0" AND @cp <= "9") _
      OR (@cp >= "A" AND @cp <= "Z") THEN
         diag_arg = err_loc & "Calc operand variable name " _
            & var_name & @cp & " invalid"

         FUNCTION = %mapstr_retcode_ERR
         MExitFunc
      END IF

      curr_cls = %cls_v                               '/ token class = variable
      var_index = %V_dummy      '/ entry to dummy var slot for null set clauses
'bug "1493 got fake variable ? that replaced " & curr_token
      RETURN
   END IF

   '/--------------------------------------------------------------------------
   '/ handle decimal number token
   '/--------------------------------------------------------------------------

   IF @cp >= "0" AND @cp <= "9" THEN
      GOSUB sub_get_decimal                            '/ sets curr_cls = cls_n
'bug "1503 get decimal " & curr_token
      RETURN
   END IF

   '/--------------------------------------------------------------------------
   '/ handle hex number token
   '/--------------------------------------------------------------------------

   IF @cp = "." THEN
      GOSUB sub_get_hex                                '/ sets curr_cls = cls_n
'bug "1513 get hex " & curr_token
      RETURN
   END IF

   '/ the only thing left that could legally be here is an operator

   IF VERIFY (@cp, $mapstr_oplist) <> VERIFIED THEN
      diag_arg = err_loc & "Calc operand invalid character " & @cp

      FUNCTION = %mapstr_retcode_ERR
      MExitFunc
   END IF

   '/--------------------------------------------------------------------------
   '/ check for non-operator tokens.
   '/--------------------------------------------------------------------------

   SELECT CASE AS CONST$ @cp

      CASE "(":
'bug "1533 got ("
         curr_cls = %cls_LP
         GOSUB sub_next_closure
         closure_delim(closure_right) = @cp


      CASE ")":
'bug "1540 got )"
         curr_cls = %cls_RP
         IF  closure_right > 0 _
         AND closure_delim(closure_right) = "(" THEN          '/ closures match
            GOSUB sub_prev_closure
         ELSE
            diag_arg = err_loc & "Calc operand has extraneous or misplaced )"
            FUNCTION = %mapstr_retcode_ERR
            MExitFunc
         END IF

      CASE "[":
'bug "1552 got ["
         '/ the [ is an operator, cannot set curr_cls = %cls_LB here

         GOSUB sub_next_closure
         closure_delim(closure_right) = @cp
         GOSUB sub_count_clauses

      CASE "]":
'bug "1560 got ]"
         '/ the ] is an operator, cannot set curr_cls = %cls_RB here

         IF  closure_right > 0 _
         AND closure_delim(closure_right) = "[" THEN          '/ closures match
         '/ GOSUB sub_prev_closure            <<<--- CANNOT DO THIS HERE --->>>
         ELSE
            diag_arg = err_loc & "Calc operand has extraneous or misplaced ]"
            FUNCTION = %mapstr_retcode_ERR
            MExitFunc
         END IF

      CASE "{":
'bug "1573 got {"
         curr_cls = %cls_begin

      CASE "}":
'bug "1577 got }"
         curr_cls = %cls_end

      CASE ";":
'bug "1581 got ;"
         curr_cls = %cls_semi

         IF closure_delim (closure_right) <> " " THEN
            diag_arg = err_loc & _
               "Calc operand semicolon cannot be inside () or []"
            FUNCTION = %mapstr_retcode_ERR
            MExitFunc
         END IF

         closure_useEQ(0) = 0           '/ flag reset after end of statement
         closure_delim(0) = " "
         curr_opr      = 0                          '/ operator token ID number
         curr_opr_prec = 0            '/ force all pending operators to execute
         curr_opr_nOps = 0                                '/ number of operands
         curr_opr_bias = 0                        '/ bias (0 or 2, where 2=RTL)
         curr_opr_base = 0                      '/ base ID for assign operators

      CASE ",":
'bug "1600 got ,"
         '/ the , is an operator, cannot set curr_cls = %cls_comma here

         IF closure_delim (closure_right) <> "[" THEN
            diag_arg = err_loc & "Calc operand comma is only valid inside []"
            FUNCTION = %mapstr_retcode_ERR
            MExitFunc
         END IF
   END SELECT


   '/ if we found non-operator tokens, they all have an opr_len of 1

   IF curr_cls <> %cls_null THEN
      opr_len = 1
      opr_found = @cp
      curr_token = @cp
      INCR cp                              '/ skip over 1-character punctuation
'bug "1618 got non-operator token " & curr_token
      RETURN
   END IF

   '/--------------------------------------------------------------------------
   '/ fix unary plus/minus
   '/
   '/ method: if + or - is preceeded by any of:   n v ) ]
   '/ then the + or - is binary; if not, it is unary
   '/
   '/ to "fix" this, we directly set the opr_table_ptr to the desired entries
   '/ which are formatted the same but are outside the normal table.
   '/--------------------------------------------------------------------------

   IF @cp = "+" OR @cp = "-" THEN
      IF  prev_cls <> %cls_n _            '/ if prior token is not in:  n v ) ]
      AND prev_cls <> %cls_v _
      AND prev_cls <> %cls_RP _
      AND prev_cls <> %cls_RB THEN                   '/ we are in unary context
'bug "1637 got unary " & @cp
         opr_len = 1
         opr_found = @cp
         curr_token = @cp
         curr_cls = %cls_U

         IF @cp = "+" THEN                                       '/ fix unary +
            opr_table_ptr = CODEPTR (opr_table_unary_plus_cp)
'bug "1645 opr_table_ptr = opr_table_unary_plus_cp"
         ELSE                                                    '/ fix unary -
            opr_table_ptr = CODEPTR (opr_table_unary_minus_cp)
'bug "1648 opr_table_ptr = opr_table_unary_minus_cp"
         END IF

         REDIM opr_entry(0) AS STRINGZ * 4 AT opr_table_ptr
         REDIM opr_byte(32) AS BYTE AT opr_table_ptr
      END IF
   END IF

   '/--------------------------------------------------------------------------
   '/ at this point, @cp contains a character that could be an operator
   '/ operators come in 3 sizes so we have to find the longest one first
   '/ we have one character for certain, maybe more. we do a 3-character
   '/ lookahead.  cannot incr 'cp' yet until operator has been identified
   '/ so we know how long it is.
   '/--------------------------------------------------------------------------

   IF opr_table_ptr = 0 THEN          '/ skip lookup if unary +/- already found
      opr_args(1) = @cp
      opr_args(2) = ""
      opr_args(3) = ""
'bug "1668 opr_args(1) = " & opr_args(1)

      IF VERIFY (@cp[1], $mapstr_oplist) = VERIFIED THEN
         opr_args(2) = @cp[0] & @cp[1]
'bug "1672 opr_args(2) = " & opr_args(2)

         IF VERIFY (@cp[2], $mapstr_oplist) = VERIFIED THEN
            opr_args(3) = @cp[0] & @cp[1] & @cp[2]
'bug "1676 opr_args(3) = " & opr_args(3)
         END IF
      END IF
      GOSUB sub_get_opcode
   END IF

   '/--------------------------------------------------------------------------
   '/ handle illegal operators here
   '/--------------------------------------------------------------------------

   IF opr_len = 0 THEN
'bug "1687 opr_len = 0, could not find operator"
      diag_arg = err_loc _
         & "Calc operand has invalid operator at " & curr_token

      FUNCTION = %mapstr_retcode_ERR
      MExitFunc
   END IF

   '/--------------------------------------------------------------------------
   '/ disambiguate = equal sign.  inside parens, it's compare, else assign.
   '/ note that [] closures make plain = into := NOT into ==
   '/ we can't change opr_len, it has to stay at 1 since that's how long "=" is
   '/--------------------------------------------------------------------------

   IF opr_found = "=" THEN
      curr_token = opr_found

      IF closure_delim (closure_right) = "(" THEN
'bug "1705 opr = disambiguated to EQ"
         opr_table_ptr = CODEPTR (opr_table_EQ_cp)
         REDIM opr_entry(0) AS STRINGZ * 4 AT opr_table_ptr
         REDIM opr_byte(32) AS BYTE AT opr_table_ptr

      ELSE                                            '/ treat like := operator
         '/--------------------------------------------------------------------
         '/ multiple equal signs in a statement are invalid
         '/ a stmt like X = Y = Z; is ambiguous.  what does the second = mean?
         '/ no matter what meaning we gave it, it could be misunderstood.
         '/ so, we disallow it; a stmt like this must be rewritten to be clear.
         '/
         '/    X = Y == Z;             result of comparing Y to Z assigned to X
         '/    X = (Y = Z);            result of comparing Y to Z assigned to X
         '/    X = Y := Z;            Z is assigned to Y which is assigned to X
         '/
         '/ the flag that is checked for this is reset at comma and semicolon
         '/--------------------------------------------------------------------

         IF closure_useEQ(closure_right) = 0 THEN
'bug "1725 opr = disambiguated to ASSIGN"
            closure_useEQ(closure_right) = 1
            opr_table_ptr = CODEPTR (opr_table_ASSIGN_cp)
            REDIM opr_entry(0) AS STRINGZ * 4 AT opr_table_ptr
            REDIM opr_byte(32) AS BYTE AT opr_table_ptr

         ELSE
'bug "1732 opr = disambiguate failed"
            diag_arg = err_loc & _
               "Calc operand has extraneous or ambiguous = sign"

            FUNCTION = %mapstr_retcode_ERR
            MExitFunc
         END IF
      END IF
   END IF

   cp += opr_len          '/ skip over the recognized operator (1,2 or 3 chars)

   '/--------------------------------------------------------------------------
   '/ if we have found the operator in the table, extract the values for use
   '/ by the shift/reduce logic.
   '/
   '/ the precedence level is multiplied by 4 so that when we do calculations
   '/ on relative precedence, and add in the bias factor, we don't want to
   '/ "collide" with the next precedence level up.  the calculation used is:
   '/
   '/    prev_prec + 1 < next_prec + bias
   '/
   '/ if the prev_prec was one level below the next_prec, and the next operator
   '/ had a bias of zero, the two would appear the same, which would appear as
   '/ if a non-associative relationship existed, which would not be true.
   '/ by avoiding numeric collisions, false non-associative relationships will
   '/ not occur.
   '/
   '/ to make the precedence of operators inside parentheses higher than those
   '/ outside, we add the 'base_prec' level.  this is incremented for each
   '/ ( and [ reached, and decremented for each ) and ] reached.
   '/--------------------------------------------------------------------------

   IF opr_table_ptr <> 0 THEN
      curr_opr      = opr_byte(5)                   '/ operator token ID number

      curr_opr_prec = (opr_byte(6) * 4) + base_prec           '/ see note above

      curr_opr_nOps = opr_byte(7)                         '/ number of operands
      curr_opr_bias = opr_byte(8)                 '/ bias (0 or 2, where 2=RTL)
      curr_opr_base = opr_byte(9)               '/ base ID for assign operators
      curr_cls      = opr_byte(10)                  '/ cls code of the operator
   END IF


   RETURN ' from sub_get_token


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_count_clauses                   **********************************
'/****   ---------------------------------   **********************************
'/*****************************************************************************

sub_count_clauses:

   '/--------------------------------------------------------------------------
   '/ this sub is entered when a [ is encountered in the lexical scan.
   '/ we do a lookahead scan to see how many clauses exist in the set
   '/ the value is stored in closure_clauses(closure_right)
   '/
   '/ method: count commas and closing ]
   '/
   '/ we must account for nested [] and () and check for nesting errors and
   '/ missing close ] in case the } or nul terminator is reached.
   '/ we cannot accurately detect mismatched [] () but that will be found
   '/ elsewhere.  the technique here will work if the expr is not malformed.
   '/ if it is malformed, they have bigger problems.
   '/--------------------------------------------------------------------------

   lap = cp                      '/ set lookahead pointer to curr scan location
   clause_level = 1                           '/ count the current [ as level 1
'bug "1804 initial @lap = hex " & HEX$(ASC(@lap)) & " = " & @lap & " at " & err_loc
'bug "1805 curr calc_str: |" & calc_str & "|"

   nn = 0
   DO
      INCR nn
      INCR lap
'bug "1811 in loop " & DEC$(nn) & " @lap = hex " & HEX$(ASC(@lap)) & " = " & @lap
      SELECT CASE @lap

         CASE "}", CHR$(0)
'bug "1815 @lap = hex " & HEX$(ASC(@lap)) & " = " & @lap
            diag_arg = err_loc & "Calc operand open [ has no close ]"
            FUNCTION = %mapstr_retcode_ERR
            MExitFunc

         CASE "[", "("
            INCR clause_level

         CASE ")"
            DECR clause_level

            '/ clause_level could only be zero if they used unmatched ) without
            '/ corresponding ( inside the set clause

            IF clause_level = 0 THEN
               diag_arg = err_loc & "Calc operand has extraneous or misplaced )"
               FUNCTION = %mapstr_retcode_ERR
               MExitFunc
            END IF

         CASE "]"
            DECR clause_level
            IF clause_level = 0 THEN          '/ this is a non-nested ] closure
               INCR closure_clauses(closure_right)
'bug "1839 number of set clauses: " & DEC$(closure_clauses(closure_right) )
               RETURN                                      '/ all done scanning
            END IF

         CASE ","
            INCR closure_clauses(closure_right)

      END SELECT
   LOOP


   RETURN ' from sub_count_clauses


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_divert_set_vars                 **********************************
'/****   ---------------------------------   **********************************
'/*****************************************************************************

sub_divert_set_vars:

   '/--------------------------------------------------------------------------
   '/ this sub is entered when a [ is encountered in expression evaluation.
   '/ our expression evaluator cannot skip subexpressions, and set notation is
   '/ not conditional.  we just plow through and evaluate everything, then
   '/ select the desired set clause based on the selector expression.  the only
   '/ problem is that if any set clause has expressions with side effects, we
   '/ can't allow them to run, because the user's intent was that they wouldn't
   '/ get executed.  but, if we tried to bypass things, and they had any syntax
   '/ errors, they might not get accurately reported.  so, instead, we do a
   '/ preliminary lookahead at all the set clauses that are within the [] list.
   '/ if any of the set clause expressions contain a variable name, we replace
   '/ the name with a ? question mark.  the ? is symbolic for a dummy 27th var
   '/ name.  then, we just let the expression evaluate, and any variable
   '/ references go to the dummy one and the real variables don't get changed.
   '/ only thing we have to be careful of is in case there was a divide by zero
   '/ involving the dummy variable, we would not report it.  that is taken care
   '/ in the divide and modulus routines.
   '/
   '/ method: do a quick and dirty parse of the set clauses.  if we are NOT in
   '/ the selected clause, and any clause expression has a variable reference,
   '/ replace the variable name with ?.  Since the expression could have hex
   '/ numbers, and any kind of number with a suffix, we can't just replace all
   '/ letters, but we have to recognize and skip numbers to do this correctly.
   '/--------------------------------------------------------------------------

   '/ the [ was already scanned and cp is past it

   lap = cp                      '/ set lookahead pointer to curr scan location
   clause_level = 1                           '/ count the current [ as level 1
   clause_found = 1     '/ identifies what clause we are looking at in the scan
'bug "1891 initial @lap = hex " & HEX$(ASC(@lap)) & " = " & @lap & " at " & err_loc
'bug "1892 curr calc_str: |" & calc_str & "|"
   DECR lap                    '/ compensate for loop so code is like prior sub

   nn = 0
   DO
      INCR nn
      INCR lap
'bug "1899 in loop " & DEC$(nn) & " @lap = hex " & HEX$(ASC(@lap)) & " = " & @lap
      SELECT CASE @lap

         CASE "}", CHR$(0)
'bug "1903 @lap = hex " & HEX$(ASC(@lap)) & " = " & @lap
            diag_arg = err_loc & "Calc operand open [ has no close ]"
            FUNCTION = %mapstr_retcode_ERR
            MExitFunc

         CASE "[", "("
            INCR clause_level

         CASE ")"
            DECR clause_level

            '/ clause_level could only be zero if they used unmatched ) without
            '/ corresponding ( inside the set clause

            IF clause_level = 0 THEN
               diag_arg = err_loc & "Calc operand has extraneous or misplaced )"
               FUNCTION = %mapstr_retcode_ERR
               MExitFunc
            END IF

         CASE "]"
            DECR clause_level
            IF clause_level = 0 THEN          '/ this is a non-nested ] closure
               RETURN                                      '/ all done scanning
            END IF

         CASE ","
            '/ we only count clauses (and thus, commas) at the same level as
            '/ the initial [ was.  so, if they were ambitious and had inner
            '/ set expressions with commas, we don't count those.

            IF clause_level = 1 THEN
               INCR clause_found
            END IF

         CASE "."                      '/ a hex number
            INCR lap                   '/ skip over dot
            GOTO SKIP_ALPHANUMERIC

         CASE "0" TO "9"
            '/ numbers can't have variable names embedded in them.
            '/ so, we just scan until there are no more alphanumerics
            '/ this is not an accurate parse, be we don't have to be accurate
            '/ since the main loop will do an INCR lap, we have to backtrack
            '/ after the last INCR here

         SKIP_ALPHANUMERIC:
            DO WHILE (@lap >= "0" AND @lap <= "9") _
                  OR (@lap >= "A" AND @lap <= "Z")
               INCR lap
            LOOP
            DECR lap              '/ backtrack to make main loop work correctly

         CASE "A" TO "Z"

            '/-----------------------------------------------------------------
            '/ if we are outside of the selected set clause, and the expression
            '/ has a variable name, and it is not followed by a letter or digit
            '/ then replace the name with ?
            '/
            '/ the extra check to examine the following-char context might not
            '/ be strictly necessary, but we want to err on the side of caution
            '/ when we are in here mangling the expression.
            '/-----------------------------------------------------------------

            IF clause_found <> clause_selected THEN
               IF (@lap[1] >= "0" AND @lap[1] <= "9") _
               OR (@lap[1] >= "A" AND @lap[1] <= "Z") THEN
                  ITERATE DO       '/ letter doesn't look right, leave it alone
               END IF
'bug "1973 variable " & @lap & " diverted to ?"

               @lap = "?"         '/ divert the variable reference to the dummy
            END IF

      END SELECT
   LOOP


   RETURN ' sub_divert_set_vars


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_next_closure                    **********************************
'/****   ---------------------------------   **********************************
'/*****************************************************************************

sub_next_closure:

'bug "1993 entered sub_next_closure"
   closure_right += 1

   IF closure_right > closure_alloc THEN
      closure_alloc = %mapstr_stack_incr + MAX(closure_right, closure_alloc)
      REDIM PRESERVE closure_useEQ  (closure_alloc)
      REDIM PRESERVE closure_delim  (closure_alloc)
      REDIM PRESERVE closure_clauses(closure_alloc)
   END IF

   closure_useEQ  (closure_right) = 0
   closure_delim  (closure_right) = " "
   closure_clauses(closure_right) = 0


   RETURN ' from sub_next_closure


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_prev_closure                    **********************************
'/****   ---------------------------------   **********************************
'/*****************************************************************************

sub_prev_closure:
'bug "2018 entered sub_prev_closure"

   IF closure_right < 1 THEN
      diag_arg = err_loc & "Calc operand internal error at prev closure"
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc
   END IF

   closure_right -= 1


   RETURN ' from sub_prev_closure


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_next_val                        **********************************
'/****   ---------------------------------   **********************************
'/*****************************************************************************

sub_next_val:
'bug "2039 entered sub_next_val"

   val_right += 1
   val_left = val_right-1

   IF val_right > val_alloc THEN
      val_alloc = %mapstr_stack_incr + MAX(val_right, val_alloc)
      REDIM PRESERVE val_type(val_alloc)
      REDIM PRESERVE val_num (val_alloc)

      val_type (val_right) = %V_null
      val_num  (val_right) = 0
   END IF


   RETURN ' from sub_next_val


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_prev_val                        **********************************
'/****   ---------------------------------   **********************************
'/*****************************************************************************

sub_prev_val:
'bug "2064 entered sub_prev_val"

   IF val_right < 1 THEN
'bug "2067 in sub_prev_val, val_right = " & DEC$(val_right)
      diag_arg = err_loc & "Calc operand internal error at prev val"
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc
   END IF

   val_right -= 1
   val_left = val_right-1


   RETURN ' from sub_prev_val


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_next_opr                        **********************************
'/****   ---------------------------------   **********************************
'/*****************************************************************************

sub_next_opr:
'bug "2087 entered sub_next_opr"

   opr_right += 1

   IF opr_right > opr_alloc THEN
      opr_alloc = %mapstr_stack_incr + MAX(opr_right, opr_alloc)
      REDIM PRESERVE opr_token(opr_alloc)
      REDIM PRESERVE opr_prec (opr_alloc)
      REDIM PRESERVE opr_nOps (opr_alloc)

      opr_token(opr_right) = %opr_null
      opr_prec (opr_right) = 0
      opr_nOps (opr_right) = 0
   END IF


   RETURN ' from sub_next_opr


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_prev_opr                        **********************************
'/****   ---------------------------------   **********************************
'/*****************************************************************************

sub_prev_opr:
'bug "2113 entered sub_prev_opr"

   IF opr_right < 1 THEN
'bug "2116 in sub_prev_opr, opr_right = " & DEC$(opr_right)
      diag_arg = err_loc & "Calc operand internal error at prev opr"
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc
   END IF

   opr_right -= 1


   RETURN ' from sub_prev_opr


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_get_opcode                      **********************************
'/****   ---------------------------------   **********************************
'/*****************************************************************************

sub_get_opcode:
'bug "2135 entered sub_get_opcode"

   '/ perform a binary search on each section of opcode table
   '/ when the entry is found, return the offset to the table, else null

   curr_token = ""

   FOR opr_len = 3 TO 1 STEP -1
      IF opr_args(opr_len) = "" THEN ITERATE FOR
'bug "2144 looking for opr |" & opr_args(opr_len) & "|"

      IF curr_token = "" THEN
         curr_token = opr_args(opr_len)   '/ save longest token in case illegal
      END IF

      opr_table_lo = opr_table_first(opr_len)
      opr_table_hi = opr_table_final(opr_len)

      DO WHILE opr_table_lo <= opr_table_hi
         opr_table_mid = (opr_table_hi + opr_table_lo) \ 2

         opr_table_ptr = (opr_table_mid * opr_table_size(0)) + opr_table_base
         REDIM opr_entry(0) AS STRINGZ * 4 AT opr_table_ptr
         REDIM opr_byte(32) AS BYTE AT opr_table_ptr

         IF opr_args(opr_len) = opr_entry(0) THEN      '/ arg found at midpoint
            opr_found  = opr_args(opr_len)        '/ capture the operator found
            curr_token = opr_args(opr_len)          '/ this is the (good) token
'bug "2163 found opr |" & opr_args(opr_len) & "| at pos " & DEC$(opr_table_mid)
            RETURN

         ELSEIF opr_args(opr_len) < opr_entry(0) THEN '/ maybe before midpoint
            opr_table_hi = opr_table_mid - 1

         ELSE                                           '/ maybe after midpoint
            opr_table_lo = opr_table_mid + 1
         END IF
      LOOP
   NEXT

'bug "2175 not found opr |" & opr_args(opr_len) & "|"
   '/ set values for not-found condition

   opr_table_ptr = 0
   opr_len = 0
   opr_found = ""


   RETURN ' from sub_get_opcode


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_get_decimal                     **********************************
'/****   ---------------------------------   **********************************
'/*****************************************************************************

sub_get_decimal:
'bug "2193 entered sub_get_decimal"

   '/ sub_get_decimal when sub_get_token finds a digit

   curr_cls = %cls_n                                 '/ token item type: number
   num_string = ""
   H_digit = "9"                                   '/ highest digit for decimal
   num_radix_desc = "decimal"                                   '/ for messages
   num_suffix_max = 19

   DO WHILE @cp >= "0" AND @cp <= "9"
      R_digit = @cp
      num_string &= @cp
      curr_token &= @cp
      INCR cp
      DO WHILE @cp = "_": INCR cp: LOOP           '/ skip underscores in number
   LOOP

   GOSUB sub_get_number_suffix
'bug "2212 num_string:|" & num_string & "|"

   '/ VAL fails for large numbers, so we calculate this the hard way

   num_value = 0
   FOR num_ndx = 1 TO LEN (num_string)
      num_c = MID$(num_string, num_ndx, 1)
      num_value = (num_value * 10) + (ASC(num_c) - ASC("0"))
   NEXT


   RETURN ' from sub_get_decimal


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_get_hex                         **********************************
'/****   ---------------------------------   **********************************
'/*****************************************************************************

sub_get_hex:
'bug "2233 entered sub_get_hex"

   '/ sub_get_hex when sub_get_token finds a dot

   curr_cls = %cls_n                                 '/ token item type: number
   num_string = ""
   H_digit = "F"                                   '/ highest digit for decimal
   R_digit = " "                        '/ don't have a hex digit yet to repeat
   num_radix_desc = "hex"
   num_suffix_max = 16

   INCR cp                                             '/ skip over leading dot
   DO WHILE @cp = "_": INCR cp: LOOP              '/ skip underscores in number

   DO WHILE (@cp >= "0" AND @cp <= "9") _
         OR (@cp >= "A" AND @cp <= "F")
      R_digit = @cp
      num_string &= @cp
      curr_token &= @cp
      INCR cp
      DO WHILE @cp = "_": INCR cp: LOOP           '/ skip underscores in number
   LOOP

   GOSUB sub_get_number_suffix

   '/ it is possible to have a hex number like .H5 but if they use ._X or .*
   '/ then the hex "number" has no digits in it.

   IF num_string = "" THEN
      diag_arg = err_loc & "Calc operand hex number malformed"
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc
   END IF

   num_value = VAL ("&H0" & num_string)


   RETURN ' from sub_get_hex


'/*****************************************************************************
'/****   ---------------------------------   **********************************
'/****   sub_get_number_suffix               **********************************
'/****   ---------------------------------   **********************************
'/*****************************************************************************

sub_get_number_suffix:
'bug "2280 entered sub_get_number_suffix"

   '/ number can end with L H or R, or a non-letter
   '/ otherwise the number is malformed

   IF @cp < "A" OR @cp > "Z" THEN                         '/ there is no suffix
      RETURN
   END IF

   IF @cp = "L" THEN                 '/ request for one or more Low digits of 0
      num_suffix_type = "L"
      S_digit = "0"

   ELSEIF @cp = "H" THEN       '/ request for one or more High digits of 9 or F
      S_digit = H_digit
      num_suffix_type = "H"

   ELSEIF @cp = "R" THEN      '/ request to repeat last digit one or more times
      num_suffix_type = "R"
      IF R_digit = " " THEN                                '/ may happen if .R4
         diag_arg = err_loc & "Calc operand has malformed " & num_radix_desc _
            & " number"

         FUNCTION = %mapstr_retcode_ERR
         MExitFunc
      END IF

      S_digit = R_digit

   ELSE
      diag_arg = err_loc & "Calc operand invalid syntax after " & num_string
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc
   END IF

   curr_token &= @cp
   INCR cp                                                '/ skip over L H or R
   DO WHILE @cp = "_": INCR cp: LOOP              '/ skip underscores in number

   IF @cp < "0" OR @cp > "9" THEN
      diag_arg = err_loc & "Calc operand number suffix invalid after " _
         & num_string & S_digit

      FUNCTION = %mapstr_retcode_ERR
      MExitFunc
   END IF

   num_suffix_str = ""

   DO WHILE @cp >= "0" AND @cp <= "9"
      num_suffix_str &= @cp
      curr_token &= @cp
      INCR cp
      DO WHILE @cp = "_": INCR cp: LOOP           '/ skip underscores in number
   LOOP

   '/ suffix count cannot be ended by a letter

   IF @cp >= "A" AND @cp <= "Z" THEN
      diag_arg = err_loc & "Calc operand number suffix invalid after " _
         & num_suffix_str

      FUNCTION = %mapstr_retcode_ERR
      MExitFunc
   END IF

   num_suffix_len = VAL (num_suffix_str)

   IF num_suffix_len < 1 _
   OR (LEN(num_string) + num_suffix_len) > num_suffix_max THEN
      diag_arg = err_loc & "Calc operand number suffix length invalid"
      FUNCTION = %mapstr_retcode_ERR
      MExitFunc
   END IF

   '/ for suffix R, one digit is already present, so reduce the count by 1
   '/ so that the count after R works the same as L and H

   IF num_suffix_type = "R" THEN
      DECR num_suffix_len
   END IF

   '/ add the suffix by golly

   DO WHILE num_suffix_len > 0
      num_string &= S_digit
      DECR num_suffix_len
   LOOP


   RETURN ' from sub_get_number_suffix


   EXIT FUNCTION                            '/ put here just to be conservative

END FUNCTION                                                      '/ mapstr_calc
