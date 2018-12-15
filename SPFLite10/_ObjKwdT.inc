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

CLASS cKwdTable

   '----- Primitive Data
   INSTANCE KWTable()         AS STRING                           ' Valid keywords
   INSTANCE KWCodes()         AS LONG                             ' Corresponding KW Code
   INSTANCE KwCtr             AS LONG                             ' No. items in table

   CLASS METHOD CREATE()                                          ' Constructor - Initialize Class stuff
   LOCAL i AS LONG
      DIM KWTable(153) AS INSTANCE STRING                         ' KW Operands     |__ DIM'ed the same
      DIM KWCodes(153) AS INSTANCE LONG                           ' KW Unique code  |
      Me.InitKW("A", %KWA)
      Me.InitKW("ALL", %KWALL)
      Me.InitKW("AND", %KWAND)
      Me.InitKW("ANOTE", %KWANOTE)
      Me.InitKW("APPEND", %KWAPPEND)
      Me.InitKW("ASSERT", %KWASSERT)
      Me.InitKW("AUTO", %KWAUTO)
      Me.InitKW("BNOTE", %KWBNOTE)
      Me.InitKW("C", %KWCHANGE)
      Me.InitKW("CAN", %KWCANCEL)
      Me.InitKW("CANCEL", %KWCANCEL)
      Me.InitKW("CHANGE", %KWCHANGE)
      Me.InitKW("CHAR", %KWCHARS)
      Me.InitKW("CHARS", %KWCHARS)
      Me.InitKW("CHG", %KWCHANGE)
      Me.InitKW("CMD", %KWCOMMAND)
      Me.InitKW("CNOTE", %KWCNOTE)
      Me.InitKW("COLOR", %KWCOLOR)
      Me.InitKW("COLOUR", %KWCOLOR)
      Me.InitKW("COM", %KWCOMMAND)
      Me.InitKW("COMMAND", %KWCOMMAND)
      Me.InitKW("COPY", %KWCOPY)
      Me.InitKW("CR", %KWCR)
      Me.InitKW("CREATE", %KWCREATE)
      Me.InitKW("CRLF", %KWCRLF)
      Me.InitKW("CS", %KWCS)
      Me.InitKW("CSR", %KWCSR)
      Me.InitKW("CURR", %KWCURRENT)
      Me.InitKW("CURRENT", %KWCURRENT)
      Me.InitKW("D", %KWD)
      Me.InitKW("DATA", %KWDATA)
      Me.InitKW("DELETE", %KWDEL)
      Me.InitKW("DEL", %KWDEL)
      Me.InitKW("DNOTE", %KWDNOTE)
      Me.InitKW("DOWN", %KWDOWN)
      Me.InitKW("DS", %KWDS)
      Me.InitKW("DX", %KWDX)
      Me.InitKW("EDIT", %KWEDIT)
      Me.InitKW("ENOTE", %KWENOTE)
      Me.InitKW("ERROR", %KWERROR)
      Me.InitKW("EXCLUDE", %KWEXCLUDE)
      Me.InitKW("EXCLUDED", %KWEXCLUDE)
      Me.InitKW("F", %KWFULL)
      Me.InitKW("FILE", %KWFILE)
      Me.InitKW("FIND", %KWFIND)
      Me.InitKW("FIRST", %KWFIRST)
      Me.InitKW("FNOTE", %KWFNOTE)
      Me.InitKW("FULL", %KWFULL)
      Me.InitKW("GNOTE", %KWGNOTE)
      Me.InitKW("H", %KWHALF)
      Me.InitKW("HALF", %KWHALF)
      Me.InitKW("HIDE", %KWHIDE)
      Me.InitKW("HNOTE", %KWHNOTE)
      Me.InitKW("HOME", %KWHOME)
      Me.InitKW("INOTE", %KWINOTE)
      Me.InitKW("JNOTE", %KWJNOTE)
      Me.InitKW("KEEP", %KWKEEP)
      Me.InitKW("KNOTE", %KWKNOTE)
      Me.InitKW("LAB", %KWLABEL)
      Me.InitKW("LABEL", %KWLABEL)
      Me.InitKW("LABELS", %KWLABEL)
      Me.InitKW("LAST", %KWLAST)
      Me.InitKW("LEFT", %KWLEFT)
      Me.InitKW("LF", %KWLF)
      Me.InitKW("LIST", %KWLIST)
      Me.InitKW("LM", %KWLM)
      Me.InitKW("LNOTE", %KWLNOTE)
      Me.InitKW("LOAD", %KWLOAD)
      Me.InitKW("LOCK", %KWLOCK)
      Me.InitKW("LOCKED", %KWLOCK)
      Me.InitKW("LONG", %KWLONG)
      Me.InitKW("M", %KWMAX)
      Me.InitKW("MAX", %KWMAX)
      Me.InitKW("MNOTE", %KWMNOTE)
      Me.InitKW("MX", %KWMX)
      Me.InitKW("NEW", %KWNEW)
      Me.InitKW("NEXT", %KWNEXT)
      Me.InitKW("NF", %KWNF)
      Me.InitKW("NL", %KWNL)
      Me.InitKW("NNOTE", %KWNNOTE)
      Me.InitKW("NONE", %KWNONE)
      Me.InitKW("NONUM", %KWNONUM)
      Me.InitKW("NOPROMPT", %KWNOPROMPT)
      Me.InitKW("NOT", %KWNOT)
      Me.InitKW("NOTE", %KWNOTE)
      Me.InitKW("NUM", %KWNUM)
      Me.InitKW("NU", %KWNU)
      Me.InitKW("NX", %KWNX)
      Me.InitKW("OFF", %KWOFF)
      Me.InitKW("ON", %KWON)
      Me.InitKW("ONOTE", %KWONOTE)
      Me.InitKW("OPEN", %KWOPEN)
      Me.InitKW("OR", %KWOR)
      Me.InitKW("P", %KWPAGE)
      Me.InitKW("PAGE", %KWPAGE)
      Me.InitKW("PNOTE", %KWPNOTE)
      Me.InitKW("PRE", %KWPREFIX)
      Me.InitKW("PFX", %KWPREFIX)
      Me.InitKW("PREFIX", %KWPREFIX)
      Me.InitKW("PREV", %KWPREV)
      Me.InitKW("PRIOR", %KWPRIOR)
      Me.InitKW("PROMPT", %KWPROMPT)
      Me.InitKW("QNOTE", %KWQNOTE)
      Me.InitKW("RAW", %KWRAW)
      Me.InitKW("REPLACE", %KWREPLACE)
      Me.InitKW("RESET", %KWRESET)
      Me.InitKW("RETRIEVE", %KWRETRIEVE)
      Me.InitKW("REV", %KWREVERSE)
      Me.InitKW("REVERSE", %KWREVERSE)
      Me.InitKW("RIGHT", %KWRIGHT)
      Me.InitKW("RM", %KWRM)
      Me.InitKW("RNOTE", %KWRNOTE)
      Me.InitKW("SET", %KWSET)
      Me.InitKW("SETUP", %KWSETUP)
      Me.InitKW("SIZE", %KWSIZE)
      Me.InitKW("SNOTE", %KWSNOTE)
      Me.InitKW("SOLID", %KWSOLID)
      Me.InitKW("-SOLID", %KWMSOLID)
      Me.InitKW("SOURCE", %KWSOURCE)
      Me.InitKW("SPE", %KWSPECIAL)
      Me.InitKW("SPECIAL", %KWSPECIAL)
      Me.InitKW("STATE", %KWSTATE)
      Me.InitKW("STD", %KWSTD)
      Me.InitKW("-STD", %KWMSTD)
      Me.InitKW("+STD", %KWPSTD)
      Me.InitKW("SUF", %KWSUFFIX)
      Me.InitKW("SFX", %KWSUFFIX)
      Me.InitKW("SUFFIX", %KWSUFFIX)
      Me.InitKW("T", %KWT)
      Me.InitKW("TAG", %KWTAG)
      Me.InitKW("TAGS", %KWTAG)
      Me.InitKW("TARGET", %KWTARGET)
      Me.InitKW("TNOTE", %KWTNOTE)
      Me.InitKW("TOG", %KWTOGGLE)
      Me.InitKW("TOGGLE", %KWTOGGLE)
      Me.InitKW("TOP", %KWTOP)
      Me.InitKW("TRUNC", %KWTRUNC)
      Me.InitKW("U", %KWU)
      Me.InitKW("UNLOCK", %KWUNLOCK)
      Me.InitKW("UNLOCKED", %KWUNLOCK)
      Me.InitKW("UNOTE", %KWUNOTE)
      Me.InitKW("UP", %KWUP)
      Me.InitKW("USE", %KWUSING)
      Me.InitKW("USER", %KWU)
      Me.InitKW("USING", %KWUSING)
      Me.InitKW("VNOTE", %KWVNOTE)
      Me.InitKW("WNOTE", %KWWNOTE)
      Me.InitKW("WORD", %KWWORD)
      Me.InitKW("WORDS", %KWWORD)
      Me.InitKW("X", %KWEXCLUDE)
      Me.InitKW("XNOTE", %KWXNOTE)
      Me.InitKW("YNOTE", %KWYNOTE)
      Me.InitKW("ZNOTE", %KWZNOTE)                                ' 153
   END METHOD                                                     '

   INTERFACE iKwdTable: INHERIT IUNKNOWN                          ' Define the interface

      METHOD KWLookup(Cmd AS STRING) AS LONG
      '----- Validate Keyword return Index (0 if invalid)
      REGISTER i AS LONG
         ARRAY SCAN KWTable() FOR KWCtr, COLLATE UCASE, =Cmd, TO i
         METHOD = IIF(i, KWCodes(i - 1), 0)                       ' Return ARRAY SCAN result
      END METHOD

   END INTERFACE

   CLASS METHOD InitKW(kw AS STRING, eq AS LONG)
      '---------- Add a single KW to the table
      INCR KWCtr
      KWTable(KWCtr) = kw
      KWCodes(KWCtr) = eq
   END METHOD

END CLASS
