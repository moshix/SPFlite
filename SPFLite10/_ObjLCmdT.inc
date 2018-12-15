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

CLASS cLCmdTable

   '----- Primary Command Table Data
   INSTANCE LCmdNumber    AS LONG                                 ' Number of Line commands
   INSTANCE LCodeNumber   AS LONG                                 ' Number of Line code routines
   INSTANCE LCmdTable()   AS LCtlValEnt                           ' Line Command Table
   INSTANCE LCodeTable()  AS DWORD                                ' Code table (Temporary)

   CLASS METHOD CREATE()                                          ' Constructor - Initialize Class stuff
   LOCAL i AS LONG
      LCmdNumber = 99                                             ' Set current number of line commands
      LCodeNumber = 41                                            ' Set current number of command routines
      DIM LCmdTable(1 TO LCmdNumber) AS INSTANCE LCtlValEnt       ' Dim the Line command table
      DIM LCodeTable(1 TO LCodeNumber) AS DWORD                   ' Dim the Primary code table

      '----- Add the CodePtr values (CODEPTR cannot exist in DATA statements
      LCodeTable(01) = CODEPTR(lCmdI)
      LCodeTable(02) = CODEPTR(lCmdDD)
      LCodeTable(03) = CODEPTR(lCmdMM)
      LCodeTable(04) = CODEPTR(lCmdCC)
      LCodeTable(05) = CODEPTR(lCmdRR)
      LCodeTable(06) = CODEPTR(lCmdLShift)
      LCodeTable(07) = CODEPTR(lCmdRShift)
      LCodeTable(08) = CODEPTR(lCmdXX)
      LCodeTable(09) = CODEPTR(lCmdF)
      LCodeTable(10) = CODEPTR(lCmdL)
      LCodeTable(11) = CODEPTR(lCmdSI)
      LCodeTable(12) = CODEPTR(lCmdTS)
      LCodeTable(13) = CODEPTR(lCmdTF)
      LCodeTable(14) = CODEPTR(lCmdUCUC)
      LCodeTable(15) = CODEPTR(lCmdLCLC)
      LCodeTable(16) = CODEPTR(lCmdSC)
      LCodeTable(17) = CODEPTR(lCmdTC)
      LCodeTable(18) = CODEPTR(lCmdCOLS)
      LCodeTable(19) = CODEPTR(lCmdBNDS)
      LCodeTable(20) = CODEPTR(lCmdTABS)
      LCodeTable(21) = CODEPTR(lCmdWord)
      LCodeTable(22) = CODEPTR(lCmdMark)
      LCodeTable(23) = CODEPTR(lCmdLDShift)
      LCodeTable(24) = CODEPTR(lCmdRDShift)
      LCodeTable(25) = CODEPTR(lCmdJJ)
      LCodeTable(26) = CODEPTR(lCmdSS)
      LCodeTable(27) = CODEPTR(lCmdTR)
      LCodeTable(28) = CODEPTR(lCmdPL)
      LCodeTable(29) = CODEPTR(lCmdTL)
      LCodeTable(30) = CODEPTR(lCmdTJJ)
      LCodeTable(31) = CODEPTR(lCmdNote)
      LCodeTable(32) = CODEPTR(lCmdMakeNote)
      LCodeTable(33) = CODEPTR(lCmdMakeData)
      LCodeTable(34) = CODEPTR(lCmdMask)
      LCodeTable(35) = CODEPTR(lCmdLIShift)
      LCodeTable(36) = CODEPTR(lCmdRIShift)
      LCodeTable(37) = CODEPTR(lCmdTT)
      LCodeTable(38) = CODEPTR(lCmdUU)
      LCodeTable(39) = CODEPTR(lCmdVV)
      LCodeTable(40) = CODEPTR(lCmdTX)
      LCodeTable(41) = CODEPTR(lCmdTU)

      ' Line command table
      '     Cmnd     - The 4 char command
      '     Ix       - The Index of the code routine in the Line routine table. 999 = no routine a Dest operand
      '     R        - Y/N - Is it a range command type
      '     #        - Y/N - Does it accept a numeric operand
      '     T        - Command type.  I = Immediate, stand alone  S = Source lines  D = Destination lines
      '     K        - Y/N - Can it be 'kept' with a K suffix
      '     +/-      - Y/N - Can it accept post processing exclude requests
      '     \/       - Y/N - Can it accept / \ line range requests
      '     RF       - Y/N - Repeat/Spacing/Numeric operand
      '     PC       - Y/N - Valid for Immediate Primary Command
      '     BR       - Y/N - Valid for Browse mode
      '
      '     Cmnd        Ix  R  #  T  K +/- \/ RF PC BR
      DATA "        ", 999, N, Y, S, Y, Y, Y, Y, Y, N             ' Line Macro Prototype
      DATA "        ", 999, Y, Y, S, Y, Y, N, Y, Y, N             ' Line Macro Block Prototype
      DATA "I       ",  01, N, Y, I, N, N, N, Y, Y, N             ' Insert
      DATA "N       ",  01, N, Y, I, N, N, N, Y, Y, N             ' NewLine
      DATA "D       ",  02, N, Y, I, N, N, Y, N, Y, N             ' Delete
      DATA "A       ", 999, N, Y, D, Y, Y, N, Y, N, N             ' After
      DATA "B       ", 999, N, Y, D, Y, Y, N, Y, N, N             ' Before
      DATA "F       ",  09, N, Y, I, N, N, N, Y, N, Y             ' First
      DATA "L       ",  10, N, Y, I, N, N, N, Y, N, Y             ' Last
      DATA "M       ",  03, N, Y, S, N, N, Y, N, N, N             ' Move
      DATA "C       ",  04, N, Y, S, Y, Y, Y, N, N, N             ' Copy
      DATA "R       ",  05, N, Y, I, N, Y, Y, Y, Y, N             ' Replicate
      DATA "(       ",  06, N, Y, I, N, Y, N, Y, Y, N             ' Shift Left
      DATA ")       ",  07, N, Y, I, N, Y, N, Y, Y, N             ' Shift Right
      DATA "O       ", 999, N, Y, D, Y, Y, Y, N, N, N             ' Overlay
      DATA "X       ",  08, N, Y, I, N, N, Y, N, Y, Y             ' eXclude
      DATA "CC      ",  04, Y, Y, S, Y, Y, N, Y, N, N             ' Copy
      DATA "RR      ",  05, Y, Y, I, N, Y, N, Y, N, N             ' Replicate
      DATA "DD      ",  02, Y, N, I, N, N, N, N, N, N             ' Delete
      DATA "MM      ",  03, Y, Y, S, N, N, N, Y, N, N             ' Move
      DATA "TS      ",  12, N, Y, I, N, N, N, Y, N, N             ' Text Split
      DATA "TF      ",  13, N, Y, I, N, N, N, Y, N, N             ' Text Flow
      DATA "((      ",  06, Y, Y, I, N, Y, N, Y, Y, N             ' Shift Left
      DATA "))      ",  07, Y, Y, I, N, Y, N, Y, Y, N             ' Shift Right
      DATA "UC      ",  14, N, Y, I, N, Y, Y, N, Y, N             ' Uppercase
      DATA "LC      ",  15, N, Y, I, N, Y, Y, N, Y, N             ' Lowercase
      DATA "XX      ",  08, Y, N, I, N, N, N, N, N, Y             ' eXclude
      DATA "OO      ", 999, Y, N, D, Y, Y, N, N, N, N             ' Overlay
      DATA "COLS    ",  18, N, N, I, N, N, N, N, N, Y             ' Columns
      DATA "LCC     ",  15, Y, N, I, N, Y, N, N, N, N             ' Lowercase
      DATA "UCC     ",  14, Y, N, I, N, Y, N, N, N, N             ' Uppercase
      DATA "BNDS    ",  19, N, N, I, N, N, N, N, N, Y             ' Bounds
      DATA "TABS    ",  20, N, N, I, N, N, N, N, N, Y             ' Tabs
      '     Cmnd        Ix  R  #  T  K +/- \/ RF PC BR
      DATA "SI      ",  11, N, Y, I, N, N, N, Y, N, Y             ' Show Indent
      DATA "<<      ",  23, Y, Y, I, N, N, N, Y, N, N             ' Shift Data Left
      DATA ">>      ",  24, Y, Y, I, N, N, N, Y, N, N             ' Shift Data Right
      DATA "<       ",  23, N, Y, I, N, N, N, Y, Y, N             ' Shift Data Left
      DATA ">       ",  24, N, Y, I, N, N, N, Y, Y, N             ' Shift Data Right
      DATA "J       ",  25, N, Y, I, N, N, Y, N, Y, N             ' Join
      DATA "JJ      ",  25, Y, N, I, N, N, N, N, N, N             ' Join
      DATA "G       ",  25, N, Y, I, N, N, Y, N, Y, N             ' Glue
      DATA "GG      ",  25, Y, N, I, N, N, N, N, N, N             ' Glue
      DATA "SC      ",  16, N, Y, I, N, Y, Y, N, Y, N             ' Sentence Case
      DATA "TC      ",  17, N, Y, I, N, Y, Y, N, Y, N             ' Title Case
      DATA "WORD    ",  21, N, N, I, N, N, N, N, N, Y             ' WORD line
      DATA "TAB     ",  20, N, N, I, N, N, N, N, N, Y             ' TABS line
      DATA "BND     ",  19, N, N, I, N, N, N, N, N, Y             ' BNDS line
      DATA "DLM     ",  21, N, N, I, N, N, N, N, N, Y             ' WORD line
      DATA "MARK    ",  22, N, N, I, N, N, N, N, N, Y             ' Mark line
      DATA "MASK    ",  34, N, N, I, N, N, N, N, N, Y             ' Mask line
      DATA "COL     ",  18, N, N, I, N, N, N, N, N, Y             ' Columns
      DATA "BOU     ",  19, N, N, I, N, N, N, N, N, Y             ' BNDS line
      DATA "S       ",  26, N, Y, I, N, N, Y, N, N, Y             ' Show
      DATA "SS      ",  26, Y, Y, I, N, N, N, N, N, Y             ' Show
      DATA "H       ", 999, N, Y, D, Y, Y, Y, N, N, N             ' Here
      DATA "HH      ", 999, Y, Y, D, Y, Y, N, Y, N, N             ' Here
      '     Cmnd        Ix  R  #  T  K +/- \/ RF PC BR
      DATA "AA      ", 999, Y, Y, D, N, Y, N, Y, N, N             ' AA Block
      DATA "BB      ", 999, Y, Y, D, N, Y, N, Y, N, N             ' BB Block
      DATA "TR      ",  27, N, Y, I, N, N, Y, N, Y, N             ' Trim Line
      DATA "TRR     ",  27, Y, Y, I, N, N, N, N, N, N             ' Trim Line
      DATA "PL      ",  28, N, Y, I, N, N, Y, Y, Y, N             ' Pad Line
      DATA "PLL     ",  28, Y, Y, I, N, N, N, N, N, N             ' Pad Line
      DATA "TL      ",  29, N, Y, I, N, Y, N, N, Y, N             ' Truncate Line
      DATA "TLL     ",  29, Y, Y, I, N, Y, N, N, N, N             ' Truncate Line
      DATA "SCC     ",  16, Y, N, I, N, Y, N, N, N, N             ' Sentence case
      DATA "TCC     ",  17, Y, N, I, N, Y, N, N, N, N             ' Title Case
      DATA "TABS    ",  20, N, N, I, N, N, N, N, N, Y             ' TABS line
      DATA "TJ      ",  30, N, Y, I, N, N, Y, N, Y, N             ' Text Join
      DATA "TJJ     ",  30, Y, N, I, N, N, N, N, N, N             ' Text Join
      DATA "TG      ",  30, N, Y, I, N, N, Y, N, Y, N             ' Text Glue
      DATA "TGG     ",  30, Y, N, I, N, N, N, N, N, N             ' Text Glue
      DATA "OR      ", 999, N, Y, D, Y, Y, Y, N, N, N             ' Overlay Replace
      DATA "ORR     ", 999, Y, N, D, Y, Y, N, N, N, N             ' Overlay
      DATA "W       ", 999, N, Y, D, N, N, Y, N, N, N             ' Swap
      DATA "WW      ", 999, Y, N, D, N, N, N, N, N, N             ' Swap
      DATA "NOTE    ",  31, N, Y, I, N, N, N, Y, Y, N             ' Note
      DATA "MN      ",  32, N, Y, I, N, N, Y, N, Y, N             ' MakeNote
      DATA "MNN     ",  32, Y, N, I, N, N, N, N, N, N             ' MakeNote
      DATA "MD      ",  33, N, Y, I, N, N, Y, N, N, N             ' MakeData
      '     Cmnd        Ix  R  #  T  K +/- \/ RF PC BR
      DATA "MDD     ",  33, Y, N, I, N, N, N, N, N, N             ' MakeData
      DATA "TFF     ",  13, Y, Y, I, N, N, N, N, N, N             ' Text Flow
      DATA "[[      ",  35, Y, Y, I, N, N, N, Y, N, N             ' Shift Indent Left
      DATA "]]      ",  36, Y, Y, I, N, N, N, Y, N, N             ' Shift Indent Right
      DATA "[       ",  35, N, Y, I, N, N, N, Y, Y, N             ' Shift Indent Left
      DATA "]       ",  36, N, Y, I, N, N, N, Y, Y, N             ' Shift Indent Right
      DATA "TM      ",  13, N, Y, I, N, N, N, Y, N, N             ' Text Margin
      DATA "TMM     ",  13, Y, Y, I, N, N, N, N, N, N             ' Text Flow
      DATA "TB      ",  12, N, Y, I, N, N, N, Y, N, N             ' Text Break
      DATA "TBB     ",  12, Y, N, I, N, N, N, N, N, N             ' Text Break
      DATA "T       ",  37, N, Y, I, N, N, Y, N, N, Y             ' Text Select
      DATA "TT      ",  37, Y, N, I, N, N, N, N, N, Y             ' Text Select
      DATA "U       ",  38, N, Y, I, N, Y, Y, N, N, Y             ' U Mark User line
      DATA "UU      ",  38, Y, N, I, N, Y, N, N, N, Y             ' UU Block Mark User line
      DATA "V       ",  39, N, Y, I, N, Y, Y, N, N, Y             ' V ReVert Mark User line
      DATA "VV      ",  39, Y, N, I, N, Y, N, N, N, Y             ' VV ReVert Block Mark User line
      DATA "TX      ",  40, N, Y, I, N, Y, Y, N, N, Y             ' TX Toggle X lines
      DATA "TXX     ",  40, Y, N, I, N, Y, N, N, N, Y             ' TXX Toggle X lines Block
      DATA "TU      ",  41, N, Y, I, N, Y, Y, N, N, Y             ' TU Toggle User Lines
      DATA "TUU     ",  41, Y, N, I, N, Y, N, N, N, Y             ' TUU Toggle User Lines Block
      '     Cmnd        Ix  R  #  T  K +/- \/ RF PC BR

      FOR i = 1 TO 99                                             ' Load the Table
         LCmdTable(i).lcTxt   = READ$((i-1)*11+1)                 '
         LCmdTable(i).lcIndex = VAL(READ$((i-1)*11+2))            '
         LCmdTable(i).lcRange = READ$((i-1)*11+3)                 '
         LCmdTable(i).lcNumV  = READ$((i-1)*11+4)                 '
         LCmdTable(i).lcType  = READ$((i-1)*11+5)                 '
         LCmdTable(i).lcKeep  = READ$((i-1)*11+6)                 '
         LCmdTable(i).lcPM    = READ$((i-1)*11+7)                 '
         LCmdTable(i).lcSlsh  = READ$((i-1)*11+8)                 '
         LCmdTable(i).lcRFac  = READ$((i-1)*11+9)                 '
         LCmdTable(i).lcPCOK  = READ$((i-1)*11+10)                '
         LCmdTable(i).lcbrOK  = READ$((i-1)*11+11)                '
      NEXT i                                                      '

      '----- Now copy the CODEPTR values to the main table
      FOR i = 1 TO 99                                             ' Loop through
         IF LCmdTable(i).lcIndex <> 999 THEN                      ' If there's an code routine number other than 999
            LCmdTable(i).lcIndex = LCodeTable(LCmdTable(i).lcIndex) ' Swap in a real code address instead of a number
         END IF                                                   '
      NEXT i                                                      '
      REDIM LCodeTable(1) AS INSTANCE DWORD                       ' Free the code table area
   END METHOD                                                     '

   INTERFACE iLCmdTable: INHERIT IUNKNOWN                         ' Define the interface

      METHOD GetCmdIX(CmdStr AS STRING) AS LONG
      '----- Validate Command name and return Index (0 if invalid)
      LOCAL i AS LONG
         ARRAY SCAN LCmdTable() FOR UBOUND(LCmdTable()), FROM 1 TO 8, =LSET$(Cmdstr, 8), TO i   ' Can we find the command?
         METHOD = i                                               ' Return the ARRAY SCAN result
      END METHOD

      METHOD GetEditCode(IX AS LONG) AS DWORD
      '----- Get Editor Code address for a line command
         METHOD = LCmdTable(IX).lcIndex                           ' Pass back address
      END METHOD

      METHOD GetName(IX AS LONG) AS STRING
      '----- Get name of a command
         METHOD = TRIM$(LCmdTable(IX).lcTxt)                      ' Pass back name
      END METHOD

      METHOD GetType(IX AS LONG) AS STRING
      '----- Get command type
         METHOD = TRIM$(LCmdTable(IX).lcType)                     ' Pass back line command type
      END METHOD

      METHOD IsBrowseOK(IX AS LONG) AS LONG
      '----- Get Line command Browse status
         METHOD = IIF(LCmdTable(IX).lcBROK = "Y", %True, %False)  ' Pass back true/false
      END METHOD

      METHOD IsKeepOK(IX AS LONG) AS LONG
      '----- Get Line command Keep status
         METHOD = IIF(LCmdTable(IX).lcKeep = "Y", %True, %False)  ' Pass back true/false
      END METHOD

      METHOD IsNumVOK(IX AS LONG) AS LONG
      '----- Get Line command Numeric Value status
         METHOD = IIF(LCmdTable(IX).lcNumV = "Y", %True, %False)  ' Pass back true/false
      END METHOD

      METHOD IsPCmdOK(IX AS LONG) AS LONG
      '----- Get Line command Primary command status
         METHOD = IIF(LCmdTable(IX).lcPCOK = "Y", %True, %False)  ' Pass back true/false
      END METHOD

      METHOD IsPMOK(IX AS LONG) AS LONG
      '----- Get Line command PM status
         METHOD = IIF(LCmdTable(IX).lcPM = "Y", %True, %False)    ' Pass back true/false
      END METHOD

      METHOD IsRangeOK(IX AS LONG) AS LONG
      '----- Get Line command Range status
         METHOD = IIF(LCmdTable(IX).lcRange = "Y", %True, %False) ' Pass back true/false
      END METHOD

      METHOD IsRFactorOK(IX AS LONG) AS LONG
      '----- Get Line command Repeat factor status
         METHOD = IIF(LCmdTable(IX).lcRfac = "Y", %True, %False)  ' Pass back true/false
      END METHOD

      METHOD IsSlashOK(IX AS LONG) AS LONG
      '----- Get Line command Slash status
         METHOD = IIF(LCmdTable(IX).lcSlsh = "Y", %True, %False)  ' Pass back true/false
      END METHOD

   END INTERFACE

END CLASS
