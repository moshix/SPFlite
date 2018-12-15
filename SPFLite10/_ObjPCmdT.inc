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

CLASS cPCmdTable

   '----- Primary Command Table Data
   INSTANCE CmdNumber    AS LONG                                  ' Number of Primary commands
   INSTANCE CodeNumber   AS LONG                                  ' Number of Primary routines
   INSTANCE PCmdTable()   AS CmdTable                             ' Primary Command Table
   INSTANCE PCodeTable()  AS DWORD                                ' Code table (Temporary)

   CLASS METHOD CREATE()                                          ' Constructor - Initialize Class stuff
   LOCAL i AS LONG
      CmdNumber = 209                                             ' Set current number of primary commands
      CodeNumber = 146                                            ' Set current number of command routines
      DIM PCmdTable(1 TO CmdNumber) AS INSTANCE CmdTable          ' Dim the Primary command table
      DIM PCodeTable(1 TO CodeNumber) AS DWORD                    ' Dim the Primary code table

      '----- Add the CodePtr values (CODEPTR cannot exist in DATA statements
      PCodeTable(001) = CODEPTR(PCmdBounds)
      PCodeTable(002) = CODEPTR(PCmdBrowse)
      PCodeTable(003) = CODEPTR(PCmdCancel)
      PCodeTable(004) = CODEPTR(PCmdCaps)
      PCodeTable(005) = CODEPTR(PCmdChange)
      PCodeTable(006) = CODEPTR(PCmdClip)
      PCodeTable(007) = CODEPTR(PCmdCopy)
      PCodeTable(008) = CODEPTR(PCmdCreate)
      PCodeTable(009) = CODEPTR(PCmdEOL)
      PCodeTable(010) = CODEPTR(PCmdCut)
      PCodeTable(011) = CODEPTR(PCmdDelete)
      PCodeTable(012) = CODEPTR(PCmdDown)
      PCodeTable(013) = CODEPTR(PCmdEdit)
      PCodeTable(014) = CODEPTR(PCmdSource)
      PCodeTable(015) = CODEPTR(PCmdEnd)
      PCodeTable(016) = CODEPTR(sDoEndAll)
      PCodeTable(017) = CODEPTR(PCmdExclude)
      PCodeTable(018) = CODEPTR(PCmdFind)
      PCodeTable(019) = CODEPTR(PCmdFlip)
      PCodeTable(020) = CODEPTR(PCmdHelp)
      PCodeTable(021) = CODEPTR(PCmdHex)
      PCodeTable(022) = CODEPTR(PCmdHide)
      PCodeTable(023) = CODEPTR(PCmdKeyMap)
      PCodeTable(024) = CODEPTR(PCmdLeft)
      PCodeTable(025) = CODEPTR(PCmdLocate)
      PCodeTable(026) = CODEPTR(PCmdMacro)
      PCodeTable(027) = CODEPTR(PCmdQuery)
      PCodeTable(028) = CODEPTR(PCmdOpen)
      PCodeTable(029) = CODEPTR(PCmdPaste)
      PCodeTable(030) = CODEPTR(PCmdOpen)
      PCodeTable(031) = CODEPTR(PCmdPrint)
      PCodeTable(032) = CODEPTR(PCmdOptions)
      PCodeTable(033) = CODEPTR(PCmdReplace)
      PCodeTable(034) = CODEPTR(PCmdReset)
      PCodeTable(035) = CODEPTR(PCmdRetrieve)
      PCodeTable(036) = CODEPTR(PCmdRight)
      PCodeTable(037) = CODEPTR(PCmdRChange)
      PCodeTable(038) = CODEPTR(PCmdRFind)
      PCodeTable(039) = CODEPTR(PCmdCmd)
      PCodeTable(040) = CODEPTR(PCmdSave)
      PCodeTable(041) = CODEPTR(PCmdSaveAs)
      PCodeTable(042) = CODEPTR(PCmdScroll)
      PCodeTable(043) = CODEPTR(PCmdSort)
      PCodeTable(044) = CODEPTR(PCmdSwap)
      PCodeTable(045) = CODEPTR(PCmdTabs)
      PCodeTable(046) = CODEPTR(PCmdUndo)
      PCodeTable(047) = CODEPTR(PCmdTop)
      PCodeTable(048) = CODEPTR(PCmdUp)
      PCodeTable(049) = CODEPTR(PCmdProfile)
      PCodeTable(050) = CODEPTR(pCmdAUTOBKUP)
      PCodeTable(051) = CODEPTR(PCmdAUTOSAVE)
      PCodeTable(052) = CODEPTR(PCmdXTABS)
      PCodeTable(053) = CODEPTR(PCmdCOLS)
      PCodeTable(054) = CODEPTR(pCmdPRESERVE)
      PCodeTable(055) = CODEPTR(PCmdSetUndo)
      PCodeTable(056) = CODEPTR(PCmdHILITE)
      PCodeTable(057) = CODEPTR(PcmdState)
      PCodeTable(058) = CODEPTR(pCmdLRECL)
      PCodeTable(059) = CODEPTR(PCmdTag)
      PCodeTable(060) = CODEPTR(PCmdNFIND)
      PCodeTable(061) = CODEPTR(pCmdNExclude)
      PCodeTable(062) = CODEPTR(pCmdDROPKEEP)
      PCodeTable(063) = CODEPTR(PCmdCASE)
      PCodeTable(064) = CODEPTR(PCmdLCUC)
      PCodeTable(065) = CODEPTR(pCmdRLoc)
      PCodeTable(066) = CODEPTR(pCmdRLocFind)
      PCodeTable(067) = CODEPTR(pCmdSHOW)
      PCodeTable(068) = CODEPTR(PCmdNSHow)
      PCodeTable(069) = CODEPTR(pCmdRECFM)
      PCodeTable(070) = CODEPTR(pCmdRECALL)
      PCodeTable(071) = CODEPTR(pCmdFav)
      PCodeTable(072) = CODEPTR(PCmdEnd)
      PCodeTable(073) = CODEPTR(pCmdSet)
      PCodeTable(074) = CODEPTR(pCmdNFLIP)
      PCodeTable(075) = CODEPTR(pCmdSubmit)
      PCodeTable(076) = CODEPTR(PCmdSubArg)
      PCodeTable(077) = CODEPTR(pCmdReLoad)
      PCodeTable(078) = CODEPTR(pCmdRename)
      PCodeTable(079) = CODEPTR(pCmdNDelete)
      PCodeTable(080) = CODEPTR(PCmdRedo)
      PCodeTable(081) = CODEPTR(pCmdCompress)
      PCodeTable(082) = CODEPTR(pCmdRetFwd)
      PCodeTable(083) = CODEPTR(pCmdDO)
      PCodeTable(084) = CODEPTR(PCmdPType)
      PCodeTable(085) = CODEPTR(FMpCmdFF)
      PCodeTable(086) = CODEPTR(pCmdMEdit)
      PCodeTable(087) = CODEPTR(pCmdGlueWith)
      PCodeTable(088) = CODEPTR(PCmdVSave)
      PCodeTable(089) = CODEPTR(pCmdStart)
      PCodeTable(090) = CODEPTR(pCmdSaveAll)
      PCodeTable(091) = CODEPTR(pCmdClone)
      PCodeTable(092) = CODEPTR(PCmdAppend)
      PCodeTable(093) = CODEPTR(pCmdPrepend)
      PCodeTable(094) = CODEPTR(pCmdEnumWith)
      PCodeTable(095) = CODEPTR(FMPCmdBottom)
      PCodeTable(096) = CODEPTR(FMpCmdBrowse)
      PCodeTable(097) = CODEPTR(FMPCmdCase)
      PCodeTable(098) = CODEPTR(FMPCmdEnd)
      PCodeTable(099) = CODEPTR(FMPCmdMakeList)
      PCodeTable(100) = CODEPTR(FMpCmdDown)
      PCodeTable(101) = CODEPTR(FMPCmdEdit)
      PCodeTable(102) = CODEPTR(FMPCmdFind)
      PCodeTable(103) = CODEPTR(FMPCmdLocate)
      PCodeTable(104) = CODEPTR(FMPCmdRecall)
      PCodeTable(105) = CODEPTR(FMPCmdTop)
      PCodeTable(106) = CODEPTR(FMPCmdUp)
      PCodeTable(107) = CODEPTR(PCmdBottom)
      PCodeTable(108) = CODEPTR(pCmdOrder)
      PCodeTable(109) = CODEPTR(pCmdMark)
      PCodeTable(110) = CODEPTR(FMPCmdRFind)
      PCodeTable(111) = CODEPTR(pCmdSplit)
      PCodeTable(112) = CODEPTR(pCmdJoin)
      PCodeTable(113) = CODEPTR(pCmdRSplit)
      PCodeTable(114) = CODEPTR(pCmdRJoin)
      PCodeTable(115) = CODEPTR(FMPCmdExclude)
      PCodeTable(116) = CODEPTR(pCmdPlease)
      PCodeTable(117) = CODEPTR(pCmdNotify)
      PCodeTable(118) = CODEPTR(pCmdLine)
      PCodeTable(119) = CODEPTR(FMPCmdView)
      PCodeTable(120) = CODEPTR(pCmdMINLEN)
      PCodeTable(121) = CODEPTR(pCmdDIR)
      PCodeTable(122) = CODEPTR(pCmdPAGE)
      PCodeTable(123) = CODEPTR(pCmdRUN)
      PCodeTable(124) = CODEPTR(pCmdFOLD)
      PCodeTable(125) = CODEPTR(pCmdCOLLATE)
      PCodeTable(126) = CODEPTR(pCmdView)
      PCodeTable(127) = CODEPTR(pCmdXSUBMIT)
      PCodeTable(128) = CODEPTR(pCmdWDIR)
      PCodeTable(129) = CODEPTR(pCmdSUBCMD)
      PCodeTable(130) = CODEPTR(pCmdADD)
      PCodeTable(131) = CODEPTR(pCmdULine)
      PCodeTable(132) = CODEPTR(pCmdNULine)
      PCodeTable(133) = CODEPTR(pCmdRevert)
      PCodeTable(134) = CODEPTR(pCmdNRevert)
      PCodeTable(135) = CODEPTR(FMpCmdDefault)
      PCodeTable(136) = CODEPTR(FMpCmdReset)
      PCodeTable(137) = 0
      PCodeTable(138) = CODEPTR(pCmdAUTOCAPS)
      PCodeTable(139) = CODEPTR(pCmdACTION)
      PCodeTable(140) = CODEPTR(FMPCmdAll)
      PCodeTable(141) = CODEPTR(pCmdTest)
      PCodeTable(142) = CODEPTR(pCmdNumber)
      PCodeTable(143) = CODEPTR(pCmdNumType)
      PCodeTable(144) = CODEPTR(pCmdUnNumber)
      PCodeTable(145) = CODEPTR(pCmdAutonum)
      PCodeTable(146) = CODEPTR(pCmdRenumber)

      ' Now do the rest of the fields

      '----- Primary command data
      '----- Fields are:
      '-----   Command name
      '-----   Editor Index into pCodetable
      '-----   Max number of operands
      '-----   Are Line operands allowed
      '-----   Is Scrolling allowed
      '-----   Allowed in FM
      '-----   FM Index into PCodeTable
      '-----   Dual Line Ops accepted
      '-----   Allowed in MACRO Call
      '-----   Eligible for Command Chaining
      '-----   Valid KW list
      '-----   Valid in Browse Mode
      '-----   Valid keyword operands
      '                PCmd#  NmOp LnOp Scr  FMPcmd LinDual Macro Chain Brws
      DATA "=X",          16,    1,  N,  0,   16,     N,      N,    N,    Y, "NOREOPEN"
      DATA "ACTION",     139,    2,  N,  0,    0,     N,      Y,    N,    Y, "SAVE,VSAVE"
      DATA "ADD",        130,    3,  S,  0,    0,     N,      Y,    Y,    N, "AND,OR"
      DATA "ALL",          0,    1,  N,  0,  140,     N,      N,    N,    N, "B,BRO,BROWSE,C,CLONE,E,EDIT,S,SEL,SELECT,M,MEDIT,P,PRINT,T,TOUCH,V,VIEW"
      DATA "APPEND",      92,    5,  S,  0,    0,     N,      Y,    Y,    N, "X,NX,U,NU,ALL,TOP,AND,OR"
      DATA "AUTOBKUP",    50,    1,  N,  0,    0,     N,      Y,    N,    Y, "ON,OFF"
      DATA "AUTOCAPS",   138,    1,  N,  0,    0,     N,      Y,    N,    Y, "ON,OFF"
      DATA "AUTONUM",    145,    5,  N,  0,    0,     N,      N,    N,    N, "ON,OFF"
      DATA "AUTOSAVE",    51,    2,  N,  0,    0,     N,      Y,    N,    Y, "ON,OFF,PROMPT,NOPROMPT"
      DATA "B",            2,    2,  N,  0,   96,     N,      N,    N,    Y, ""
      DATA "BND",          1,    2,  N,  0,    0,     N,      Y,    N,    Y, "MAX"
      DATA "BNDS",         1,    2,  N,  0,    0,     N,      Y,    N,    Y, "MAX"
      DATA "BOT",        107,    0,  N,  1,   95,     N,      N,    N,    Y, ""
      DATA "BOTTOM",     107,    0,  N,  1,   95,     N,      N,    N,    Y, ""
      DATA "BOU",          1,    2,  N,  0,    0,     N,      Y,    N,    Y, "MAX"
      DATA "BOUND",        1,    2,  N,  0,    0,     N,      Y,    N,    Y, "MAX"
      DATA "BOUNDS",       1,    2,  N,  0,    0,     N,      Y,    N,    Y, "MAX"
      DATA "BR",           2,    2,  N,  0,   96,     N,      N,    N,    Y, ""
      DATA "BRO",          2,    1,  N,  0,   96,     N,      N,    N,    Y, ""
      DATA "BROWSE",       2,    2,  N,  0,   96,     N,      N,    N,    Y, ""
      DATA "C",            5,   11,  S,  0,    0,     N,      Y,    Y,    N, "FIRST,LAST,NEXT,PREV,ALL,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,WORDS,CHAR,CHARS,LEFT,RIGHT,TRUNC,DS,CS,X,NX,U,NU,MX,DX,TOP,AND,OR,~C1,~C2"
      DATA "CAN",          3,    1,  N,  0,    0,     N,      N,    N,    Y, "DELETE,DEL,PURGE,PUR"
      DATA "CANCEL",       3,    1,  N,  0,    0,     N,      N,    N,    Y, "DELETE,DEL,PURGE,PUR"
      DATA "CAPS",         4,    1,  N,  0,    0,     N,      Y,    N,    N, "ON,OFF,AUTO"
      DATA "CASE",        63,    1,  N,  0,   97,     N,      Y,    N,    Y, "C,T"
      DATA "CHA",          5,   11,  S,  0,    0,     N,      Y,    Y,    N, "FIRST,LAST,NEXT,PREV,ALL,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,WORDS,CHAR,CHARS,LEFT,RIGHT,TRUNC,DS,CS,X,NX,U,NU,MX,DX,TOP,AND,OR,~C1,~C2"
      DATA "CHANGE",       5,   11,  S,  0,    0,     N,      Y,    Y,    N, "FIRST,LAST,NEXT,PREV,ALL,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,WORDS,CHAR,CHARS,LEFT,RIGHT,TRUNC,DS,CS,X,NX,U,NU,MX,DX,TOP,AND,OR,~C1,~C2"
      DATA "CHG",          5,   11,  S,  0,    0,     N,      Y,    Y,    N, "FIRST,LAST,NEXT,PREV,ALL,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,WORDS,CHAR,CHARS,LEFT,RIGHT,TRUNC,DS,CS,X,NX,U,NU,MX,DX,TOP,AND,OR,~C1,~C2"
      DATA "CLIP",         6,    1,  N,  0,    6,     N,      N,    N,    Y, ""
      DATA "CLONE",       91,    2,  N,  0,   91,     N,      N,    N,    Y, ""
      DATA "CMD",         39,   20,  N,  0,   39,     N,      Y,    N,    Y, ""
      DATA "COL",         53,    1,  N,  0,    0,     N,      N,    N,    Y, "ON,OFF"
      DATA "COLS",        53,    1,  N,  0,    0,     N,      N,    N,    Y, "ON,OFF"
      DATA "COLLATE",    125,    1,  N,  0,    0,     N,      Y,    N,    Y, "ANSI,ASCII,EBCDIC"
      DATA "COMPRESS",    81,   15,  S,  0,    0,     N,      Y,    Y,    N, "X,NX,U,NU,MX,DX,ALL,TOP,AND,OR,~C1,~C2"
      DATA "CP",          81,   15,  S,  0,    0,     N,      Y,    Y,    N, "X,NX,U,NU,MX,DX,ALL,TOP,AND,OR,~C1,~C2"
      DATA "COPY",         7,    5,  D,  0,    0,     N,      Y,    Y,    N, "BEFORE,AFTER"
      DATA "CRE",          8,    4,  S,  0,    0,     N,      Y,    Y,    Y, "X,NX,U,NU,AND,OR"
      DATA "CREATE",       8,    6,  S,  0,    0,     N,      Y,    Y,    Y, "X,NX,U,NU,AND,OR"
      DATA "CRETRIEV",    35,   50,  N,  1,   35,     N,      N,    N,    Y, ""
      DATA "CUT",         10,    4,  S,  0,    0,     N,      Y,    Y,    Y, "X,NX,U,NU,REPLACE,APPEND,RAW,AND,OR"
      '                PCmd#  NmOp LnOp Scr  FMPcmd LinDual Macro Chain Brws
      DATA "DEF",          0,    3,  N,  0,  135,     N,      N,    N,    N, ""
      DATA "DEFAULT",      0,    3,  N,  0,  135,     N,      N,    N,    N, ""
      DATA "DEL",         11,    5,  S,  0,    0,     N,      Y,    Y,    N, "DUP,~N1,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,WORDS,CHAR,CHARS,X,NX,U,NU,ALL,FIRST,LAST,NEXT,PREV,TOP,AND,OR"
      DATA "DELETE",      11,    5,  S,  0,    0,     N,      Y,    Y,    N, "DUP,~N1,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,WORDS,CHAR,CHARS,X,NX,U,NU,ALL,FIRST,LAST,NEXT,PREV,TOP,AND,OR"
      DATA "DO",          83,    1,  S,  0,   83,     N,      N,    N,    Y, ""
      DATA "DOWN",        12,    1,  N,  1,  100,     N,      N,    N,    Y, "PAGE,HALF,DATA,MAX,CSR"
      DATA "DIR",        121,    1,  N,  0,  100,     N,      N,    N,    Y, ""
      DATA "DROP",        62,    3,  N,  0,    0,     N,      Y,    N,    N, "FIRST,LAST,ALL"
      DATA "E",           13,    2,  N,  0,  101,     N,      N,    N,    Y, "NEW"
      DATA "ED",          13,    2,  N,  0,  101,     N,      N,    N,    Y, "NEW"
      DATA "EDIT",        13,    2,  N,  0,  101,     N,      N,    N,    Y, "NEW"
      DATA "EDITSET",     32,    1,  N,  0,   32,     N,      N,    N,    Y, "GENERAL,FILEMGR,FM,SUBMIT,SCREEN,MOUSE,KEYBOARD,KB"
      DATA "EDSET",       32,    1,  N,  0,   32,     N,      N,    N,    Y, "GENERAL,FILEMGR,FM,SUBMIT,SCREEN,MOUSE,KEYBOARD,KB"
      DATA "END",         15,    1,  N,  0,   98,     N,      N,    N,    Y, ""
      DATA "ENUMWITH",    94,    1,  N,  0,    0,     N,      Y,    N,    N, ""
      DATA "EOL",          9,    1,  N,  0,    0,     N,      Y,    N,    Y, "CRLF,CR,LF,NL,AUTO,AUTONL,NONE"
      DATA "EX",          17,   10,  S,  0,  115,     N,      Y,    Y,    Y, "FIRST,LAST,NEXT,PREV,ALL,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,WORDS,CHAR,CHARS,U,NU,AND,OR~C1,TOP"
      DATA "EXC",         17,   10,  S,  0,  115,     N,      Y,    Y,    Y, "FIRST,LAST,NEXT,PREV,ALL,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,WORDS,CHAR,CHARS,U,NU,AND,OR,~C1,TOP"
      DATA "EXCLUDE",     17,   10,  S,  0,  115,     N,      Y,    Y,    Y, "FIRST,LAST,NEXT,PREV,ALL,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,WORDS,CHAR,CHARS,U,NU,AND,OR,~C1,TOP"
      DATA "EXCLUDED",    17,   10,  S,  0,  115,     N,      Y,    Y,    Y, "FIRST,LAST,NEXT,PREV,ALL,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,WORDS,CHAR,CHARS,U,NU,AND,OR,~C1,TOP"
      DATA "EXIT",        16,    3,  N,  0,   16,     N,      N,    N,    Y, "NOREOPEN,END,CANCEL,CAN,PURGE,PUR,DELETE,DEL"
      DATA "F",           18,   10,  S,  1,  102,     N,      Y,    Y,    Y, "WORDS,WORD,CHARS,CHAR,FIRST,LAST,NEXT,PREV,ALL,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,CHAR,LEFT,RIGHT,~C1,~C2,X,NX,U,NU,MX,DX,TOP,AND,OR"
      DATA "FAV",         71,    1,  N,  0,    0,     N,      Y,    N,    Y, ""
      DATA "FAVORITE",    71,    1,  N,  0,    0,     N,      Y,    N,    Y, ""
      DATA "FAVOURITE",   71,    1,  N,  0,    0,     N,      Y,    N,    Y, ""
      DATA "FF",          18,   10,  S,  1,   85,     N,      N,    Y,    Y, "CHARS,CHAR,WORD,WORDS,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,NF"
      DATA "FIND",        18,   10,  S,  1,  102,     N,      Y,    Y,    Y, "WORDS,WORD,CHARS,CHAR,FIRST,LAST,NEXT,PREV,ALL,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,CHAR,LEFT,RIGHT,~C1,~C2,X,NX,U,NU,MX,DX,TOP,AND,OR"
      DATA "FLIP",        19,    5,  S,  0,    0,     N,      Y,    Y,    Y, "FIRST,LAST,NEXT,PREV,ALL,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,WORDS,CHAR,CHARS,U,NU,AND,OR,~C1,TOP"
      DATA "FOLD",       124,    1,  N,  0,    0,     N,      Y,    N,    Y, "ON,OFF"
      DATA "GLUEWITH",    87,    1,  N,  0,    0,     N,      Y,    N,    N, ""
      DATA "H",           20,    1,  N,  0,   20,     N,      N,    N,    Y, ""
      DATA "HELP",        20,    1,  N,  0,   20,     N,      N,    N,    Y, ""
      DATA "HEX",         21,    1,  N,  0,    0,     N,      Y,    N,    Y, "ON,OFF"
      DATA "HI",          56,    3,  N,  0,    0,     N,      Y,    N,    Y, "ON,OFF,AUTO,FIND"
      DATA "HIDE",        22,    1,  N,  0,    0,     N,      Y,    N,    Y, "ON,OFF"
      DATA "HILIGHT",     56,    3,  N,  0,    0,     N,      Y,    N,    Y, "ON,OFF,AUTO,FIND"
      DATA "HILITE",      56,    3,  N,  0,    0,     N,      Y,    N,    Y, "ON,OFF,AUTO,FIND"
      DATA "ISCROLL",     42,    3,  N,  1,   42,     N,      N,    N,    Y, ""
      DATA "JOIN",       112,   18,  S,  0,    0,     N,      Y,    Y,    N, "FIRST,LAST,NEXT,PREV,ALL,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,WORDS,CHAR,CHARS,~C1,X,NX,U,NU,TOP,AND,OR"
      DATA "KEEP",        62,    3,  N,  0,    0,     N,      Y,    N,    N, "FIRST,LAST"
      DATA "KEY",         23,    1,  N,  0,   23,     N,      N,    N,    Y, "LIST"
      DATA "KBD",         23,    1,  N,  0,   23,     N,      N,    N,    Y, "LIST"
      '                PCmd#  NmOp LnOp Scr  FMPcmd LinDual Macro Chain Brws
      DATA "KEYMAP",      23,    1,  N,  0,   23,     N,      N,    N,    Y, "LIST"
      DATA "KEYS",        23,    1,  N,  0,   23,     N,      N,    N,    Y, "LIST"
      DATA "L",           25,    7,  N,  1,  103,     N,      Y,    Y,    Y, "NOT,CHANGE,COMMAND,ERROR,EXCLUDED,X,LABEL,~N1,FIND,SPECIAL,KEEP,TAG,LONG,PAGE,FILE,SIZE,NEXT,FIRST,LAST,PREV,CURRENT,ALL,NX,MX,~C1,TOP,U,NU"
      DATA "LC",          64,    5,  S,  0,    0,     N,      Y,    Y,    N, "X,NX,ALL,U,NU,MX,DX,AND,OR"
      DATA "LEFT",        24,    1,  N,  1,    0,     N,      N,    N,    Y, "PAGE,HALF,DATA,MAX,CSR"
      DATA "LINE",       118,   50,  N,  1,    0,     N,      Y,    Y,    N, "FIRST,LAST,ALL,X,NX,U,NU,TOP,AND,OR"
      DATA "LOC",         25,    7,  N,  1,  103,     N,      Y,    Y,    Y, "NOT,CHANGE,COMMAND,ERROR,EXCLUDED,X,LABEL,~N1,FIND,SPECIAL,KEEP,TAG,LONG,PAGE,FILE,SIZE,NEXT,FIRST,LAST,PREV,CURRENT,ALL,NX,MX,~C1,TOP,U,NU"
      DATA "LOCATE",      25,    7,  N,  1,  103,     N,      Y,    Y,    Y, "NOT,CHANGE,COMMAND,ERROR,EXCLUDED,X,LABEL,~N1,FIND,SPECIAL,KEEP,TAG,LONG,PAGE,FILE,SIZE,NEXT,FIRST,LAST,PREV,CURRENT,ALL,NX,MX,~C1,TOP,U,NU"
      DATA "LREC",        58,    1,  N,  0,    0,     N,      Y,    N,    Y, ""
      DATA "LRECL",       58,    1,  N,  0,    0,     N,      Y,    N,    Y, ""
      DATA "MAC",         26,   20,  N,  0,    0,     N,      N,    N,    N, ""
      DATA "MACRO",       26,   20,  N,  0,    0,     N,      N,    N,    N, ""
      DATA "MAKELIST",     0,    3,  N,  0,   99,     N,      N,    N,    N, "SYM,REPLACE,APPEND"
      DATA "MARK",       109,    1,  N,  0,    0,     N,      Y,    N,    Y, "ON,OFF"
      DATA "MEDIT",       86, 9999,  N,  0,    0,     N,      N,    N,    Y, ""
      DATA "MINLEN",     120,    1,  N,  0,    0,     N,      Y,    N,    N, ""
      DATA "ML",           0,    3,  N,  0,   99,     N,      N,    N,    N, "SYM,REPLACE,APPEND"
      DATA "NDEL",        79,   10,  S,  0,    0,     N,      Y,    Y,    N, "PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,WORDS,CHAR,CHARS,~C1,X,NX,U,NU,ALL,FIRST,LAST,NEXT,PREV,TOP,AND,OR"
      DATA "NDELETE",     79,   10,  S,  0,    0,     N,      Y,    Y,    N, "PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,WORDS,CHAR,CHARS,~C1,X,NX,U,NU,ALL,FIRST,LAST,NEXT,PREV,TOP,AND,OR"
      DATA "NEXCLUDE",    61,   10,  S,  0,    0,     N,      Y,    Y,    Y, "FIRST,LAST,NEXT,PREV,ALL,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,WORDS,CHAR,CHARS,U,NU,AND,OR,~C1,TOP"
      DATA "NF",          60,   10,  S,  1,    0,     N,      Y,    Y,    Y, "FIRST,LAST,NEXT,PREV,ALL,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,WORDS,CHAR,CHARS,X,NX,U,NU,MX,DX,TOP,AND,OR"
      DATA "NFIND",       60,   10,  S,  1,    0,     N,      Y,    Y,    Y, "FIRST,LAST,NEXT,PREV,ALL,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,WORDS,CHAR,CHARS,X,NX,U,NU,MX,DX,TOP,AND,OR"
      DATA "NFLIP",       74,   10,  S,  0,    0,     N,      Y,    Y,    Y, "FIRST,LAST,NEXT,PREV,ALL,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,WORDS,CHAR,CHARS,U,NU,AND,OR,~C1,TOP"
      DATA "NOTIFY",     117,    1,  N,  0,  117,     N,      Y,    N,    Y, "ALL,NONE,EDIT,RESET"
      DATA "NREVERT",    134,   10,  S,  0,    0,     N,      Y,    Y,    Y, "CHAR,CHARS,WORD,WORDS,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,~C1,MX,DX,X,NX,ALL,TOP,AND,OR"
      DATA "NV",         134,   10,  S,  0,    0,     N,      Y,    Y,    Y, "CHAR,CHARS,WORD,WORDS,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,~C1,MX,DX,X,NX,ALL,TOP,AND,OR"
      DATA "NSHOW",       68,   10,  S,  0,    0,     N,      Y,    Y,    Y, "FIRST,LAST,NEXT,PREV,ALL,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,WORDS,CHAR,CHARS,U,NU,AND,OR,~C1,TOP"
      DATA "NU",         132,   10,  S,  0,    0,     N,      Y,    Y,    Y, "CHAR,CHARS,WORD,WORDS,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,~C1,MX,DX,X,NX,ALL,TOP,AND,OR"
      DATA "NULINE",     132,   10,  S,  0,    0,     N,      Y,    Y,    Y, "CHAR,CHARS,WORD,WORDS,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,~C1,MX,DX,X,NX,ALL,TOP,AND,OR"
      DATA "NUMBER",     142,    0,  N,  0,    0,     N,      N,    N,    N, ""
      DATA "NUMB",       142,    0,  N,  0,    0,     N,      N,    N,    N, ""
      DATA "NUM",        142,    0,  N,  0,    0,     N,      N,    N,    N, ""
      DATA "NUMTYPE",    143,    4,  N,  0,    0,     N,      Y,    N,    N, "TRUNC,REAL,NONE"
      DATA "NX",          61,   10,  S,  0,    0,     N,      Y,    Y,    Y, "FIRST,LAST,NEXT,PREV,ALL,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,WORDS,CHAR,CHARS,U,NU,AND,OR,~C1,TOP"
      DATA "OPEN",        28,    1,  N,  0,   28,     N,      Y,    N,    Y, "NEW"
      DATA "OPENV",       28,    1,  N,  0,   28,     N,      Y,    N,    Y, "NEW"
      DATA "OPENB",       28,    1,  N,  0,   28,     N,      Y,    N,    Y, "NEW"
      DATA "OPT",         32,    1,  N,  0,   32,     N,      N,    N,    Y, "GENERAL,FILEMGR,FM,SUBMIT,SCREEN,MOUSE,KEYBOARD,KB"
      DATA "OPTION",      32,    1,  N,  0,   32,     N,      N,    N,    Y, "GENERAL,FILEMGR,FM,SUBMIT,SCREEN,MOUSE,KEYBOARD,KB"
      DATA "OPTIONS",     32,    1,  N,  0,   32,     N,      N,    N,    Y, "GENERAL,FILEMGR,FM,SUBMIT,SCREEN,MOUSE,KEYBOARD,KB"
      DATA "ORDER",      108,    0,  N,  0,    0,     N,      Y,    N,    N, ""
      DATA "PAGE",       122,    2,  N,  0,    0,     N,      Y,    N,    Y, "ON,OFF,SCROLL"
      DATA "PASTE",       29,    2,  D,  0,    0,     N,      Y,    Y,    N, "BEFORE,AFTER,ERASE"
      '                PCmd#  NmOp LnOp Scr  FMPcmd LinDual Macro Chain Brws
      DATA "PLEASE",     116,    1,  N,  0,  116,     N,      N,    N,    N, ""
      DATA "PREPEND",     93,    5,  S,  0,    0,     N,      Y,    Y,    N, "~C1,X,NX,U,NU,ALL,TOP,AND,OR"
      DATA "PRESERVE",    54,    1,  N,  0,    0,     N,      Y,    N,    Y, "ON,OFF,C"
      DATA "PRINT",       31,    3,  S,  0,   31,     N,      Y,    Y,    Y, "SETUP,X,NX,U,NU,NUM,NONUM,AND,OR"
      DATA "PRO",         49,    3,  N,  0,   49,     N,      N,    N,    Y, "LOCK,UNLOCK,RESET,COPY,USE,NEW,EDIT"
      DATA "PROF",        49,    3,  N,  0,   49,     N,      N,    N,    Y, "LOCK,UNLOCK,RESET,COPY,USE,NEW,EDIT"
      DATA "PROFILE",     49,    3,  N,  0,   49,     N,      N,    N,    Y, "LOCK,UNLOCK,RESET,COPY,USE,NEW,EDIT"
      DATA "PTYPE",       84,    5,  S,  0,    0,     N,      N,    Y,    N, "X,NX,U,NU,AND,OR"
      DATA "PT",          84,    5,  S,  0,    0,     N,      N,    Y,    N, "X,NX,U,NU,AND,OR"
      DATA "Q",           27,    2,  N,  0,    0,     N,      Y,    N,    Y, ""
      DATA "QUERY",       27,    2,  N,  0,    0,     N,      Y,    N,    Y, ""
      DATA "RC",          70,    1,  N,  0,  104,     N,      N,    N,    Y, "RECENT"
      DATA "RCHANGE",     37,   11,  N,  0,    0,     N,      Y,    Y,    N, "FIRST,LAST,NEXT,PREV,ALL,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,WORDS,CHAR, CHARS,LEFT,RIGHT,TRUNC,DS,CS,X,NX,U,NU,MX,DX,TOP,AND,OR,~C1,~C2"
      DATA "RECALL",      70,    1,  N,  0,  104,     N,      N,    N,    Y, "RECENT"
      DATA "RECALL",      70,    1,  N,  0,  104,     N,      N,    N,    Y, "RECENT"
      DATA "RECFM",       69,    1,  N,  0,    0,     N,      Y,    N,    Y, "U,F,V,VBI,VLI"
      DATA "REDO",        80,    0,  N,  0,    0,     N,      N,    N,    N, ""
      DATA "RELOAD",      77,    0,  N,  0,    0,     N,      N,    N,    N, ""
      DATA "RENAME",      78,    1,  N,  0,    0,     N,      Y,    N,    N, ""
      DATA "RENUM",      146,    1,  N,  0,    0,     N,      N,    N,    N, "QUIET"
      DATA "RENUMBER",   146,    1,  N,  0,    0,     N,      N,    N,    N, "QUIET"
      DATA "REN",        146,    1,  N,  0,    0,     N,      N,    N,    N, "QUIET"
      DATA "REP",         33,    4,  S,  0,    0,     N,      Y,    Y,    Y, "X,NX,U,NU,AND,OR"
      DATA "REPL",        33,    3,  S,  0,    0,     N,      Y,    Y,    Y, "X,NX,U,NU,AND,OR"
      DATA "REPLACE",     33,    4,  S,  0,    0,     N,      Y,    Y,    Y, "X,NX,U,NU,AND,OR"
      DATA "RES",         34,    5,  N,  1,  136,     N,      Y,    Y,    Y, "ALL,CHANGE,C,CHG,COMMAND,CMD,COM,FIND,WORD,WORDS,EXCLUDED,X,EX,EXC,EXCLUDE,USER,U,HIDE,LABEL,LAB,LABELS,SPECIAL,SPE,STATE,TAGS,TAG,STD,COLOR,AND,OR"
      DATA "RESET",       34,    5,  S,  1,  136,     N,      Y,    Y,    Y, "ALL,CHANGE,C,CHG,COMMAND,CMD,COM,FIND,WORD,WORDS,EXCLUDED,X,EX,EXC,EXCLUDE,USER,U,HIDE,LABEL,LAB,LABELS,SPECIAL,SPE,STATE,TAGS,TAG,STD,COLOR,AND,OR"
      DATA "RETF",        82,   50,  N,  1,   82,     N,      N,    N,    Y, ""
      DATA "RETRIEVE",    35,   50,  N,  1,   35,     N,      N,    N,    Y, ""
      DATA "REVERT",     133,   10,  S,  0,    0,     N,      Y,    Y,    Y, "FIRST,LAST,NEXT,PREV,ALL,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,WORDS,CHAR,CHARS,U,NU,AND,OR,~C1,TOP"
      DATA "RFIND",       38,   11,  N,  1,  110,     N,      Y,    Y,    Y, "REVERSE,WORDS,WORD,CHARS,CHAR,FIRST,LAST,NEXT,PREV,ALL,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,CHAR,LEFT,RIGHT,~C1,~C2,X,NX,U,NU,MX,DX,TOP,AND,OR"
      DATA "RIGHT",       36,    1,  N,  1,    0,     N,      N,    N,    Y, "PAGE,HALF,DATA,MAX,CSR"
      DATA "RJOIN",      114,    0,  S,  0,    0,     N,      Y,    Y,    N, ""
      DATA "RLOC",        65,    7,  N,  1,    0,     N,      Y,    Y,    Y, "REVERSE"
      DATA "RLOCFIND",    66,   11,  N,  1,  110,     N,      Y,    Y,    Y, "REVERSE"
      DATA "RSPLIT",     113,    0,  S,  0,    0,     N,      Y,    Y,    N, ""
      DATA "RUN",        123,   50,  N,  0,    0,     N,      N,    N,    Y, ""
      DATA "SAV",         40,    2,  N,  0,    0,     N,      Y,    N,    N, ""
      '                PCmd#  NmOp LnOp Scr  FMPcmd LinDual Macro Chain Brws
      DATA "SAVE",        40,    2,  N,  0,    0,     N,      Y,    N,    N, ""
      DATA "SAVEALL",     90,    1,  N,  0,   90,     N,      Y,    N,    N, "COND"
      DATA "SAVEAS",      41,    1,  N,  0,    0,     N,      Y,    N,    N, ""
      DATA "SC",          64,    5,  S,  0,    0,     N,      Y,    Y,    N, "X,NX,ALL,U,NU,MX,DX,AND,OR"
      DATA "SET",         73,    4,  N,  0,   73,     N,      Y,    N,    Y, "OFF,PUSH,POP"
      DATA "SETU",        55,    2,  N,  0,    0,     N,      Y,    N,    Y, ""
      DATA "SETUNDO",     55,    1,  N,  0,    0,     N,      Y,    N,    Y, ""
      DATA "SHOW",        67,   10,  S,  0,    0,     N,      Y,    Y,    Y, "FIRST,LAST,NEXT,PREV,ALL,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,WORDS,CHAR,CHARS,AND,OR"
      DATA "SORT",        43,   20,  S,  0,    0,     N,      Y,    Y,    N, "UNIQ,X,NX,MX,DX,AND,OR,MARKUNIQ"
      DATA "SOURCE",      14,    1,  N,  0,    0,     N,      Y,    N,    Y, "ANSI,UTF8,UTF16,UTF16LE,UTF16BE,EBCDIC"
      DATA "SPLIT",      111,   18,  S,  0,    0,     N,      Y,    Y,    N, "FIRST,LAST,NEXT,PREV,ALL,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,WORDS,CHAR,CHARS,LEFT,RIGHT,X,NX,AND,OR"
      DATA "START",       89,    1,  N,  0,    0,     N,      Y,    N,    Y, "FIRST,LAST,PRIOR,LABEL,NEW"
      DATA "ST",          57,    1,  N,  0,    0,     N,      Y,    N,    N, "ON,OFF,FEW,MOST,CREATE,CRE,DELETE,DEL"
      DATA "STATE",       57,    1,  N,  0,    0,     N,      Y,    N,    N, "ON,OFF,FEW,MOST,CREATE,CRE,DELETE,DEL"
      DATA "STATS",       57,    1,  N,  0,    0,     N,      Y,    N,    N, "ON,OFF,FEW,MOST,CREATE,CRE,DELETE,DEL"
      DATA "SUB",         75,   50,  S,  0,    0,     N,      Y,    Y,    Y, "X,NX,DEBUG,AND,OR"
      DATA "SUBARG",      76,    1,  N,  0,    0,     N,      Y,    N,    Y, "OFF"
      DATA "SUBCMD",     129,    1,  N,  0,    0,     N,      Y,    N,    Y, "OFF"
      DATA "SUBMIT",      75,   50,  S,  0,    0,     N,      Y,    Y,    Y, "X,NX,DEBUG,AND,OR"
      DATA "SWAP",        44,    1,  N,  0,   44,     N,      N,    N,    Y, "PREV,NEXT,PRIOR,FIRST,LAST,HOME,LIST"
      DATA "TAB",         45,    1,  N,  0,    0,     N,      Y,    N,    Y, "ON,OFF"
      DATA "TABS",        45,    1,  N,  0,    0,     N,      Y,    N,    Y, "ON,OFF"
      DATA "TAG",         59,   10,  S,  0,    0,     N,      Y,    N,    Y, "NF,ON,OFF,TOGGLE,ASSERT,SET,FIRST,LAST,NEXT,PREV,ALL,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,WORDS,CHAR,CHARS,X,NX,MX,DX,AND,OR"
      DATA "TC",          64,    5,  S,  0,    0,     N,      Y,    Y,    N, "X,NX,ALL,MX,DX,AND,OR"
      DATA "TEST",       141,    1,  N,  0,  141,     N,      N,    N,    Y, "REGEX,MAP,MAPPING,CALC"
      DATA "TOP",         47,    0,  N,  1,  105,     N,      N,    N,    Y, ""
      DATA "UC",          64,    5,  S,  0,    0,     N,      Y,    Y,    N, "X,NX,ALL,MX,DX,AND,OR"
      DATA "ULINE",      131,   10,  S,  0,    0,     N,      Y,    Y,    Y, "FIRST,LAST,NEXT,PREV,ALL,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,WORDS,CHAR,CHARS,U,NU,AND,OR,~C1,TOP"
      DATA "UU",         131,   10,  S,  0,    0,     N,      Y,    Y,    Y, "FIRST,LAST,NEXT,PREV,ALL,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,WORDS,CHAR,CHARS,U,NU,AND,OR,~C1,TOP"
      DATA "UNDO",        46,    0,  N,  0,    0,     N,      N,    N,    N, ""
      DATA "UP",          48,    1,  N,  1,  106,     N,      N,    N,    Y, "PAGE,HALF,DATA,MAX,CSR"
      DATA "UNNUMBER",   144,    3,  N,  0,    0,     N,      N,    N,    N, ""
      DATA "UNNUMB",     144,    3,  N,  0,    0,     N,      N,    N,    N, ""
      DATA "UNNUM",      144,    3,  N,  0,    0,     N,      N,    N,    N, ""
      DATA "UNN",        144,    3,  N,  0,    0,     N,      N,    N,    N, ""
      DATA "V",          126,    2,  N,  0,  119,     N,      N,    N,    Y, ""
      DATA "VIEW",       126,    2,  N,  0,  119,     N,      N,    N,    Y, ""
      DATA "VSAVE",       88,    1,  N,  0,    0,     N,      Y,    N,    N, "QUIET"
      DATA "VV",         133,   10,  S,  0,    0,     N,      Y,    Y,    Y, "FIRST,LAST,NEXT,PREV,ALL,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,WORDS,CHAR,CHARS,U,NU,AND,OR,~C1,TOP"
      DATA "WDIR",       128,    0,  N,  0,  128,     N,      N,    N,    Y, ""
      DATA "XSUBMIT",    127,    1,  N,  0,  127,     N,      Y,    N,    Y, ""
      DATA "XSUB",       127,    1,  N,  0,  127,     N,      Y,    N,    Y, ""
      DATA "XTABS",       52,    1,  N,  0,    0,     N,      Y,    N,    Y, ""
      DATA "XTAB",        52,    1,  N,  0,    0,     N,      Y,    N,    Y, ""
      DATA "X",           17,   10,  S,  0,  115,     N,      Y,    Y,    Y, "FIRST,LAST,NEXT,PREV,ALL,PREFIX,PRE,PFX,SUFFIX,SUF,SFX,WORD,WORDS,CHAR,CHARS,U,NU,AND,OR,~C1,TOP"
      '                PCmd#  NmOp LnOp Scr  FMPcmd LinDual Macro Chain Brws
      '
      '
      '----- Load the Command table
      FOR i = 1 TO  CmdNumber                                     ' Load the primary command table
         PCmdTable(i).NAME     = READ$((i - 1) * 11 + 1)          ' Command name
         PCmdTable(i).EditCode = VAL(READ$((i - 1) * 11 + 2))     ' Command Index into PCodeTable
         PCmdTable(i).NumOps   = VAL(READ$((i - 1) * 11 + 3))     ' Max number of operands
         PCmdTable(i).LinOps   = READ$((i - 1) * 11 + 4)          ' Lin operands allowed
         PCmdTable(i).Scroll   = VAL(READ$((i - 1) * 11 + 5))     ' Scrolling allowed
         PCmdTable(i).FMCode   = VAL(READ$((i - 1) * 11 + 6))     ' Command index into PCodeTable
         PCmdTable(i).LinDual  = READ$((i - 1) * 11 + 7)          ' Accepts Dual line Ops (CC/MM and A/B)
         PCmdTable(i).MacOK    = READ$((i - 1) * 11 + 8)          ' OK to run from MACRO SPFCommand call
         PCmdTable(i).ChainOK  = READ$((i - 1) * 11 + 9)          ' Eligible for Command Chaining
         PCmdTable(i).BrowseOK = READ$((i - 1) * 11 + 10)         ' Valid for Browse modeEligible for Command Chaining
         PCmdTable(i).Oprands  = READ$((i - 1) * 11 + 11)         ' Valid KW Oprands
      NEXT i
      PCmdTable(125).Oprands = _                                  ' Add Q operands (too long for DATA)
         "AUTOBKUP,AUTOCAPS,AUTOSAVE,BOUNDS,CAPS,CASE,CHANGE,COLLATE,COLS,ENUMWITH,EOL,FOLD,GLUEWITH,HEX,HIDE,HILITE,LRECL,MARK,MINLEN,NOTIFY,PAGE,PRESERVE,RECFM,SETUNDO,SOURCE,START,STATE,SUBARG,SUBCMD,TABS,XTABS"
      PCmdTable(126).Oprands = _                                  ' Add QUERY operands (too long for DATA)
         "AUTOBKUP,AUTOCAPS,AUTOSAVE,BOUNDS,CAPS,CASE,CHANGE,COLLATE,COLS,ENUMWITH,EOL,FOLD,GLUEWITH,HEX,HIDE,HILITE,LRECL,MARK,MINLEN,NOTIFY,PAGE,PRESERVE,RECFM,SETUNDO,SOURCE,START,STATE,SUBARG,SUBCMD,TABS,XTABS"

      '----- Now copy the CODEPTR values to the main table
      FOR i = 1 TO CmdNumber                                      ' Loop through
         IF PCmdTable(i).EditCode <> 0 THEN                       ' If there's an EDIT code routine number
            PCmdTable(i).EditCode = PCodeTable(PCmdTable(i).EditCode) ' Swap in a real code address instead of a number
         END IF                                                   '
         IF PCmdTable(i).FMCode <> 0 THEN                         ' If there's an FM code routine number
            PCmdTable(i).FMCode = PCodeTable(PCmdTable(i).FMCode) ' Swap in a real code address instead of a number
         END IF                                                   '
      NEXT i                                                      '

      REDIM PCodeTable(1) AS INSTANCE DWORD                       ' Free the table area
   END METHOD                                                     '

   INTERFACE iPCmdTable: INHERIT IUNKNOWN                         ' Define the interface

      METHOD GetCmdIX(CmdStr AS STRING) AS LONG
      '----- Validate Command name and return Index (0 if invalid)
      LOCAL i AS LONG
         ARRAY SCAN PCmdTable() FOR UBOUND(PCmdTable()), FROM 1 TO 16, COLLATE UCASE, =LSET$(Cmdstr, 16), TO i   ' Can we find the command?
         METHOD = i                                               ' Return the ARRAY SCAN result
      END METHOD

      METHOD GetEditCode(IX AS LONG) AS DWORD
      '----- Get Editor Code address for a command
         METHOD = PCmdTable(IX).EditCode                          ' Pass back address
      END METHOD

      METHOD GetFMCode(IX AS LONG) AS DWORD
      '----- Get FM Code address for a command
         METHOD = PCmdTable(IX).FMCode                            ' Pass back address
      END METHOD

      METHOD GetLinOps(IX AS LONG) AS STRING
      '----- Get Line Type Operands
         METHOD = TRIM$(PCmdTable(IX).LinOps)                     ' Pass back name
      END METHOD

      METHOD GetName(IX AS LONG) AS STRING
      '----- Get name of a command
         METHOD = TRIM$(PCmdTable(IX).Name)                       ' Pass back name
      END METHOD

      METHOD GetNumOps(IX AS LONG) AS LONG
      '----- Get number of allowed operands
         METHOD = PCmdTable(IX).NumOps                            ' Pass back number
      END METHOD

      METHOD GetOprands(IX AS LONG) AS STRING
      '----- Get number of allowed operands
      REGISTER i AS LONG
      LOCAL t, cNames AS STRING
         t = TRIM$(PCmdTable(IX).Oprands)                         ' Get the basic string
         '----- fudge since DATA statements can't be longer than 256
         cNames = "STD,SOLID,-STD,"                               ' Start with the basic ones
         FOR i = 1 TO 15                                          ' Build valid color names
            cNames += nHiLites(i) + ",-" + nHiLites(i) + ","      ' Add the color and the -color version
         NEXT i                                                   '
         REPLACE "~C1" WITH cNames IN t
         cNames = "+STD,"                                         ' Start with the basic one
         FOR i = 1 TO 15                                          ' Build valid color names
            cNames += nHiLites(i) + ",+" + nHiLites(i) + ","      ' Add the color and the +color version
         NEXT i                                                   '
         REPLACE "~C2" WITH cNames IN t                           '
         REPLACE "~N1"  WITH "NOTE,ANOTE,BNOTE,CNOTE,DNOTE,ENOTE,FNOTE,GNOTE,HNOTE,INOTE,JNOTE,KNOTE,LNOTE,MNOTE,~N2," IN t
         REPLACE "~N2"  WITH "NNOTE,ONOTE,PNOTE,QNOTE,RNOTE,SNOTE,TNOTE,UNOTE,VNOTE,WNOTE,XNOTE,WNOTE,ZNOTE," IN t
         METHOD = t                                               ' Pass back operand string
      END METHOD

      METHOD IsBrowseOK(IX AS LONG) AS LONG
      '----- See if command supported in Browse Node
         METHOD = IIF(PCmdTable(IX).BrowseOK = "Y", %True, %False)' Return true/false
      END METHOD

      METHOD IsChainOK(IX AS LONG) AS LONG
      '----- See if command supported in Chain mode
         METHOD = IIF(PCmdTable(IX).ChainOK = "Y", %True, %False) ' Return true/false
      END METHOD

      METHOD IsDualOK(IX AS LONG) AS LONG
      '----- See if Dual mode supported
         METHOD = IIF(PCmdTable(IX).LinDual = "Y", %True, %False) ' Return true/false
      END METHOD

      METHOD IsMacOK(IX AS LONG) AS LONG
      '----- See if command supported in Macros
         METHOD = IIF(PCmdTable(IX).MacOK = "Y", %True, %False)   ' Return true/false
      END METHOD

      METHOD IsScrollOK(IX AS LONG) AS LONG
      '----- See if command is a scrolling command
         METHOD = IIF(PCmdTable(IX).Scroll <> 0, %True, %False)   ' Return true/false
      END METHOD

      METHOD IsFMAllowed(IX AS LONG) AS LONG
      '----- See if command supported in File Manager
         METHOD = IIF(PCmdTable(IX).FMCode <> 0, %True, %False)   ' Return true/false
      END METHOD

      METHOD IsEditAllowed(IX AS LONG) AS LONG
      '----- See if command supported in Editor Mode
         METHOD = IIF(PCmdTable(IX).EditCode <> 0, %True, %False) ' Return true/false
      END METHOD

   END INTERFACE

END CLASS
