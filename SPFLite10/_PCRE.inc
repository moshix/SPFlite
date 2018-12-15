' /*************************************************
' *       Perl-Compatible Regular Expressions      *
' *************************************************/
'
' /* This is the public header file for the PCRE library, to be #included by
' applications that call the PCRE functions.
'
'            Copyright (c) 1997-2014 University of Cambridge
'
' -----------------------------------------------------------------------------
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'
'     * Redistributions of source code must retain the above copyright notice,
'       this list of conditions and the following disclaimer.
'
'     * Redistributions in binary form must reproduce the above copyright
'       notice, this list of conditions and the following disclaimer in the
'       documentation and/or other materials provided with the distribution.
'
'     * Neither the name of the University of Cambridge nor the names of its
'       contributors may be used to endorse or promote products derived from
'       this software without specific prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
' IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
' ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE
' LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
' CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
' SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
' INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
' CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
' ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
' POSSIBILITY OF SUCH DAMAGE.
' -----------------------------------------------------------------------------
' */

' /* The current PCRE version information. */

$PCRE_MAJOR       =  "8"
$PCRE_MINOR       =  "33"
$PCRE_PRERELEASE  =  ""
$PCRE_DATE        =  "2015-04-28"

' /* Public options. Some are compile-time only, some are run-time only, and some
' are both. Most of the compile-time options are saved with the compiled regex so
' that they can be inspected during studying (and therefore JIT compiling). Note
' that pcre_study() has its own set of options. Originally, all the options
' defined here used distinct bits. However, almost all the bits in a 32-bit word
' are now used, so in order to conserve them, option bits that were previously
' only recognized at matching time (i.e. by pcre_exec() or pcre_dfa_exec()) may
' also be used for compile-time options that affect only compiling and are not
' relevant for studying or JIT compiling.
'
' Some options for pcre_compile() change its behaviour but do not affect the
' behaviour of the execution functions. Other options are passed through to the
' execution functions and affect their behaviour, with or without affecting the
' behaviour of pcre_compile().
'
' Options that can be passed to pcre_compile() are tagged Cx below, with these
' variants:
'
' C1   Affects compile only
' C2   Does not affect compile; affects exec, dfa_exec
' C3   Affects compile, exec, dfa_exec
' C4   Affects compile, exec, dfa_exec, study
' C5   Affects compile, exec, study
'
' Options that can be set for pcre_exec() and/or pcre_dfa_exec() are flagged with
' E and D, respectively. They take precedence over C3, C4, and C5 settings passed
' from pcre_compile(). Those that are compatible with JIT execution are flagged
' with J. */

%PCRE_CASELESS                = &H00000001  ' /* C1     */
%PCRE_MULTILINE               = &H00000002  ' /* C1     */
%PCRE_DOTALL                  = &H00000004  ' /* C1     */
%PCRE_EXTENDED                = &H00000008  ' /* C1     */
%PCRE_ANCHORED                = &H00000010  ' /* C4 E D */
%PCRE_DOLLAR_ENDONLY          = &H00000020  ' /* C2     */
%PCRE_EXTRA                   = &H00000040  ' /* C1     */
%PCRE_NOTBOL                  = &H00000080  ' /*  E D J */
%PCRE_NOTEOL                  = &H00000100  ' /*  E D J */
%PCRE_UNGREEDY                = &H00000200  ' /* C1     */
%PCRE_NOTEMPTY                = &H00000400  ' /*  E D J */
%PCRE_UTF8                    = &H00000800  ' /* C4      )          */
%PCRE_UTF16                   = &H00000800  ' /* C4      ) Synonyms */
%PCRE_UTF32                   = &H00000800  ' /* C4      )          */
%PCRE_NO_AUTO_CAPTURE         = &H00001000  ' /* C1     */
%PCRE_NO_UTF8_CHECK           = &H00002000  ' /* C1 E D J )          */
%PCRE_NO_UTF16_CHECK          = &H00002000  ' /* C1 E D J ) Synonyms */
%PCRE_NO_UTF32_CHECK          = &H00002000  ' /* C1 E D J )          */
%PCRE_AUTO_CALLOUT            = &H00004000  ' /* C1     */
%PCRE_PARTIAL_SOFT            = &H00008000  ' /*  E D J  ) Synonyms */
%PCRE_PARTIAL                 = &H00008000  ' /*  E D J  )          */

' /* This pair use the same bit. */
%PCRE_NEVER_UTF               = &H00010000  ' /* C1      ) Overlaid */
%PCRE_DFA_SHORTEST            = &H00010000  ' /*    D    ) Overlaid */

' /* This pair use the same bit. */
%PCRE_NO_AUTO_POSSESS         = &H00020000  ' /* C1      ) Overlaid */
%PCRE_DFA_RESTART             = &H00020000  ' /*    D    ) Overlaid */

%PCRE_FIRSTLINE               = &H00040000  ' /* C3     */
%PCRE_DUPNAMES                = &H00080000  ' /* C1     */
%PCRE_NEWLINE_CR              = &H00100000  ' /* C3 E D */
%PCRE_NEWLINE_LF              = &H00200000  ' /* C3 E D */
%PCRE_NEWLINE_CRLF            = &H00300000  ' /* C3 E D */
%PCRE_NEWLINE_ANY             = &H00400000  ' /* C3 E D */
%PCRE_NEWLINE_ANYCRLF         = &H00500000  ' /* C3 E D */
%PCRE_BSR_ANYCRLF             = &H00800000  ' /* C3 E D */
%PCRE_BSR_UNICODE             = &H01000000  ' /* C3 E D */
%PCRE_JAVASCRIPT_COMPAT       = &H02000000  ' /* C5     */
%PCRE_NO_START_OPTIMIZE       = &H04000000  ' /* C2 E D  ) Synonyms */
%PCRE_NO_START_OPTIMISE       = &H04000000  ' /* C2 E D  )          */
%PCRE_PARTIAL_HARD            = &H08000000  ' /*  E D J */
%PCRE_NOTEMPTY_ATSTART        = &H10000000  ' /*  E D J */
%PCRE_UCP                     = &H20000000  ' /* C3     */

' /* Exec-time and get/set-time error codes */

%PCRE_ERROR_NOMATCH           =     (-1)
%PCRE_ERROR_NULL              =     (-2)
%PCRE_ERROR_BADOPTION         =     (-3)
%PCRE_ERROR_BADMAGIC          =     (-4)
%PCRE_ERROR_UNKNOWN_OPCODE    =     (-5)
%PCRE_ERROR_UNKNOWN_NODE      =     (-5)  ' /* For backward compatibility */
%PCRE_ERROR_NOMEMORY          =     (-6)
%PCRE_ERROR_NOSUBSTRING       =     (-7)
%PCRE_ERROR_MATCHLIMIT        =     (-8)
%PCRE_ERROR_CALLOUT           =     (-9)  ' /* Never used by PCRE itself */
%PCRE_ERROR_BADUTF8           =    (-10)  ' /* Same for 8/16/32 */
%PCRE_ERROR_BADUTF16          =    (-10)  ' /* Same for 8/16/32 */
%PCRE_ERROR_BADUTF32          =    (-10)  ' /* Same for 8/16/32 */
%PCRE_ERROR_BADUTF8_OFFSET    =    (-11)  ' /* Same for 8/16 */
%PCRE_ERROR_BADUTF16_OFFSET   =    (-11)  ' /* Same for 8/16 */
%PCRE_ERROR_PARTIAL           =    (-12)
%PCRE_ERROR_BADPARTIAL        =    (-13)
%PCRE_ERROR_INTERNAL          =    (-14)
%PCRE_ERROR_BADCOUNT          =    (-15)
%PCRE_ERROR_DFA_UITEM         =    (-16)
%PCRE_ERROR_DFA_UCOND         =    (-17)
%PCRE_ERROR_DFA_UMLIMIT       =    (-18)
%PCRE_ERROR_DFA_WSSIZE        =    (-19)
%PCRE_ERROR_DFA_RECURSE       =    (-20)
%PCRE_ERROR_RECURSIONLIMIT    =    (-21)
%PCRE_ERROR_NULLWSLIMIT       =    (-22)  ' /* No longer actually used */
%PCRE_ERROR_BADNEWLINE        =    (-23)
%PCRE_ERROR_BADOFFSET         =    (-24)
%PCRE_ERROR_SHORTUTF8         =    (-25)
%PCRE_ERROR_SHORTUTF16        =    (-25)  ' /* Same for 8/16 */
%PCRE_ERROR_RECURSELOOP       =    (-26)
%PCRE_ERROR_JIT_STACKLIMIT    =    (-27)
%PCRE_ERROR_BADMODE           =    (-28)
%PCRE_ERROR_BADENDIANNESS     =    (-29)
%PCRE_ERROR_DFA_BADRESTART    =    (-30)
%PCRE_ERROR_JIT_BADOPTION     =    (-31)
%PCRE_ERROR_BADLENGTH         =    (-32)
%PCRE_ERROR_UNSET             =    (-33)

' /* Specific error codes for UTF-8 validity checks */

%PCRE_UTF8_ERR0               =      0
%PCRE_UTF8_ERR1               =      1
%PCRE_UTF8_ERR2               =      2
%PCRE_UTF8_ERR3               =      3
%PCRE_UTF8_ERR4               =      4
%PCRE_UTF8_ERR5               =      5
%PCRE_UTF8_ERR6               =      6
%PCRE_UTF8_ERR7               =      7
%PCRE_UTF8_ERR8               =      8
%PCRE_UTF8_ERR9               =      9
%PCRE_UTF8_ERR10              =     10
%PCRE_UTF8_ERR11              =     11
%PCRE_UTF8_ERR12              =     12
%PCRE_UTF8_ERR13              =     13
%PCRE_UTF8_ERR14              =     14
%PCRE_UTF8_ERR15              =     15
%PCRE_UTF8_ERR16              =     16
%PCRE_UTF8_ERR17              =     17
%PCRE_UTF8_ERR18              =     18
%PCRE_UTF8_ERR19              =     19
%PCRE_UTF8_ERR20              =     20
%PCRE_UTF8_ERR21              =     21
%PCRE_UTF8_ERR22              =     22  ' /* Unused (was non-character) */

' /* Request types for pcre_fullinfo() */

%PCRE_INFO_OPTIONS            =      0
%PCRE_INFO_SIZE               =      1
%PCRE_INFO_CAPTURECOUNT       =      2
%PCRE_INFO_BACKREFMAX         =      3
%PCRE_INFO_FIRSTBYTE          =      4
%PCRE_INFO_FIRSTCHAR          =      4  ' /* For backwards compatibility */
%PCRE_INFO_FIRSTTABLE         =      5
%PCRE_INFO_LASTLITERAL        =      6
%PCRE_INFO_NAMEENTRYSIZE      =      7
%PCRE_INFO_NAMECOUNT          =      8
%PCRE_INFO_NAMETABLE          =      9
%PCRE_INFO_STUDYSIZE          =     10
%PCRE_INFO_DEFAULT_TABLES     =     11
%PCRE_INFO_OKPARTIAL          =     12
%PCRE_INFO_JCHANGED           =     13
%PCRE_INFO_HASCRORLF          =     14
%PCRE_INFO_MINLENGTH          =     15
%PCRE_INFO_JIT                =     16
%PCRE_INFO_JITSIZE            =     17
%PCRE_INFO_MAXLOOKBEHIND      =     18
%PCRE_INFO_FIRSTCHARACTER     =     19
%PCRE_INFO_FIRSTCHARACTERFLAGS=     20
%PCRE_INFO_REQUIREDCHAR       =     21
%PCRE_INFO_REQUIREDCHARFLAGS  =     22
%PCRE_INFO_MATCHLIMIT         =     23
%PCRE_INFO_RECURSIONLIMIT     =     24
%PCRE_INFO_MATCH_EMPTY        =     25

' /* Request types for pcre_config(). Do not re-arrange, in order to remain
' compatible. */

%PCRE_CONFIG_UTF8                    =      0
%PCRE_CONFIG_NEWLINE                 =      1
%PCRE_CONFIG_LINK_SIZE               =      2
%PCRE_CONFIG_POSIX_MALLOC_THRESHOLD  =      3
%PCRE_CONFIG_MATCH_LIMIT             =      4
%PCRE_CONFIG_STACKRECURSE            =      5
%PCRE_CONFIG_UNICODE_PROPERTIES      =      6
%PCRE_CONFIG_MATCH_LIMIT_RECURSION   =      7
%PCRE_CONFIG_BSR                     =      8
%PCRE_CONFIG_JIT                     =      9
%PCRE_CONFIG_UTF16                   =     10
%PCRE_CONFIG_JITTARGET               =     11
%PCRE_CONFIG_UTF32                   =     12
%PCRE_CONFIG_PARENS_LIMIT            =     13

' /* Request types for pcre_study(). Do not re-arrange, in order to remain
' compatible. */

%PCRE_STUDY_JIT_COMPILE              =    &H0001
%PCRE_STUDY_JIT_PARTIAL_SOFT_COMPILE =    &H0002
%PCRE_STUDY_JIT_PARTIAL_HARD_COMPILE =    &H0004
%PCRE_STUDY_EXTRA_NEEDED             =    &H0008

' /* Bit flags for the pcre[16|32]_extra structure. Do not re-arrange or redefine
' these bits, just add new ones on the end, in order to remain compatible. */

%PCRE_EXTRA_STUDY_DATA               =    &H0001
%PCRE_EXTRA_MATCH_LIMIT              =    &H0002
%PCRE_EXTRA_CALLOUT_DATA             =    &H0004
%PCRE_EXTRA_TABLES                   =    &H0008
%PCRE_EXTRA_MATCH_LIMIT_RECURSION    =    &H0010
%PCRE_EXTRA_MARK                     =    &H0020
%PCRE_EXTRA_EXECUTABLE_JIT           =    &H0040

' /* The structure for passing additional data to pcre_exec(). This is defined in
' such as way as to be extensible. Always add new fields at the end, in order to
' remain compatible. */

TYPE pcre_extra
  flags        AS DWORD           ' /* Bits for which fields are set */
  study_data   AS DWORD           ' /* Opaque data from pcre_study() */
  match_limit  AS DWORD           ' /* Maximum number of calls to match() */
  callout_data AS DWORD           ' /* Data passed back in callouts */
  tables       AS DWORD           ' /* Pointer to character tables */
  match_limit_recursion AS DWORD  ' /* Max recursive calls to match() */
  MARK         AS DWORD           ' /* For passing back a mark pointer */
  executable_jit AS DWORD         ' /* Contains a pointer to a compiled jit code */
END TYPE

' /* The structure for passing out data via the pcre_callout_function. We use a
' structure so that new fields can be added on the end in future versions,
' without changing the API of the function, thereby allowing old clients to work
' without modification. */

TYPE pcre_callout_block
  version AS INTEGER              ' /* Identifies version of block */
  ' /* ------------------------ Version 0 ------------------------------- */
  callout_number   AS LONG         ' /* Number compiled into pattern */
  offset_vector    AS LONG PTR     ' /* The offset vector */
  subject          AS ASCIIZ PTR   ' /* The subject being matched */
  subject_length   AS LONG         ' /* The length of the subject */
  start_match      AS LONG         ' /* Offset to start of this match attempt */
  current_position AS LONG         ' /* Where we currently are in the subject */
  capture_top      AS LONG         ' /* Max current capture */
  capture_last     AS LONG         ' /* Most recently closed capture */
  callout_data     AS DWORD        ' /* Data passed in with the call */
  ' /* ------------------- Added for Version 1 -------------------------- */
  pattern_position AS LONG         ' /* Offset to next item in the pattern */
  next_item_length AS LONG         ' /* Length of next item in the pattern */
  ' /* ------------------- Added for Version 2 -------------------------- */
  mark             AS BYTE PTR     ' /* Pointer to current mark or NULL  */
  ' /* ------------------------------------------------------------------ */
END TYPE


'************************************************
'        Compile a Regular Expression           *
'************************************************

' This function takes a string and returns a pointer to a PCRE data block
' holding a compiled version of the expression. The original API for this
' function had no error code return variable; it is retained for backwards
' compatibility. The new function is given a new name.
'
'Returns:        pointer to compiled data block, or NULL on error,
'                with errorptr and erroroffset set
'
' Original Declares
'PCRE_EXP_DECL pcre *pcre_compile(const char *, int, const char **, int *,
'                  const unsigned char *);
'DECLARE FUNCTION pcre_compile CDECL LIB "pcre3.dll" ALIAS "pcre_compile" _ ' DLL Name
'                                (BYVAL  ASCIIZ PTR,                      _ ' Regex string
'                                 BYVAL LONG,                             _ ' Options
'                                 BYVAL  DWORD,                           _ ' Pointer to error text
'                                 BYVAL LONG PTR,                         _ ' Offset to error location
'                                 BYVAL BYTE PTR) AS DWORD                  ' Character tables
DECLARE FUNCTION pcre_compile CDECL _                                      '
                                (BYVAL  ASCIIZ PTR,                      _ ' Regex string
                                 BYVAL LONG,                             _ ' Options
                                 BYVAL  DWORD,                           _ ' Pointer to error text
                                 BYVAL LONG PTR,                         _ ' Offset to error location
                                 BYVAL BYTE PTR) AS DWORD                  ' Character tables


'************************************************
'         Execute a Regular Expression          *
'************************************************

' This function applies a compiled re to a subject string and picks out
'portions of the string if it matches. Two elements in the vector are set for
'each substring: the offsets to the start and end of the substring.
'
'Returns:          > 0 => success; value is the number of elements filled in
'                  = 0 => success, but offsets is not big enough
'                   -1 => failed to match
'                 < -1 => some kind of unexpected problem
'
'
'Original Declare
'PCRE_EXP_DECL int  pcre_exec(const pcre *, const pcre_extra *, PCRE_SPTR,
'                   int, int, int, int *, int);


'DECLARE FUNCTION pcre_exec CDECL LIB "pcre3.dll" ALIAS "pcre_exec"       _ ' DLL Name
'                             (BYVAL DWORD PTR,                           _ ' Ptr to compile area
'                              BYVAL pcre_extra PTR,                      _ ' Ptr to extra data
'                              BYVAL ASCIIZ PTR,                          _ ' Test string
'                              BYVAL LONG,                                _ ' Length of test string
'                              BYVAL LONG,                                _ ' Start scan offset
'                              BYVAL LONG,                                _ ' Options
'                              BYVAL LONG PTR,                            _ ' Start of Offset Table
'                              BYVAL LONG) AS LONG                          ' # entries in Offset Table
DECLARE FUNCTION pcre_exec CDECL _                                       _ '
                             (BYVAL DWORD PTR,                           _ ' Ptr to compile area
                              BYVAL pcre_extra PTR,                      _ ' Ptr to extra data
                              BYVAL ASCIIZ PTR,                          _ ' Test string
                              BYVAL LONG,                                _ ' Length of test string
                              BYVAL LONG,                                _ ' Start scan offset
                              BYVAL LONG,                                _ ' Options
                              BYVAL LONG PTR,                            _ ' Start of Offset Table
                              BYVAL LONG) AS LONG                          ' # entries in Offset Table


'************************************************
'         Free the PCRE compile area            *
'************************************************

' This function frees the memory allocated to the PCRE compile area.
'
'Original Declare
'PCRE_EXP_DECL int  pcre_free(const pcre *);

'DECLARE SUB pcre_free   LIB "pcre3.dll" ALIAS "pcre_free" _         ' DLL name
'                             (argument_re AS DWORD)
DECLARE SUB pcre_free   (argument_re AS DWORD)

'------- Note pcre_free returns address of a POINTER to the real Free routine
'                                            =======

'----- Unconverted (since we're not using them) PCRE API calls

'PCRE_EXP_DECL pcre *pcre_compile2(const char *, int, int *, const char **,
'                  int *, const unsigned char *);
'PCRE_EXP_DECL int  pcre_config(int, void *);
'PCRE_EXP_DECL int  pcre_copy_named_substring(const pcre *, const char *,
'                  int *, int, const char *, char *, int);
'PCRE_EXP_DECL int  pcre_copy_substring(const char *, int *, int, int,
'                  char *, int);
'PCRE_EXP_DECL int  pcre_dfa_exec(const pcre *, const pcre_extra *,
'                  const char *, int, int, int, int *, int , int *, int);
'PCRE_EXP_DECL int  pcre_jit_exec(const pcre *, const pcre_extra *,
'                   PCRE_SPTR, int, int, int, int *, int,
'                   pcre_jit_stack *);
'PCRE_EXP_DECL void pcre_free_substring(const char *);
'PCRE_EXP_DECL void pcre_free_substring_list(const char **);
'PCRE_EXP_DECL int  pcre_fullinfo(const pcre *, const pcre_extra *, int,
'                  void *);
'PCRE_EXP_DECL int  pcre_get_named_substring(const pcre *, const char *,
'                  int *, int, const char *, const char **);
'PCRE_EXP_DECL int  pcre_get_stringnumber(const pcre *, const char *);
'PCRE_EXP_DECL int  pcre_get_stringtable_entries(const pcre *, const char *,
'                  char **, char **);
'PCRE_EXP_DECL int  pcre_get_substring(const char *, int *, int, int,
'                  const char **);
'PCRE_EXP_DECL int  pcre_get_substring_list(const char *, int *, int,
'                  const char ***);
'PCRE_EXP_DECL const unsigned char *pcre_maketables(void);
'PCRE_EXP_DECL int  pcre_refcount(pcre *, int);
'PCRE_EXP_DECL pcre_extra *pcre_study(const pcre *, int, const char **);
'PCRE_EXP_DECL void pcre_free_study(pcre_extra *);
'PCRE_EXP_DECL const char *pcre_version(void);
'PCRE_EXP_DECL int  pcre_pattern_to_host_byte_order(pcre *, pcre_extra *,
'                  CONST unsigned char *);
'PCRE_EXP_DECL pcre_jit_stack *pcre_jit_stack_alloc(int, int);
'PCRE_EXP_DECL void pcre_jit_stack_free(pcre_jit_stack *);
'PCRE_EXP_DECL void pcre_assign_jit_stack(pcre_extra *,
'                  pcre_jit_callback, void *);
'PCRE_EXP_DECL void pcre_jit_free_unused_memory(void);
