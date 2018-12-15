#COMPILE EXE
#DIM ALL
$V1 = "10"
$V2 = "1"
FUNCTION PBMAIN () AS LONG
LOCAL FNum, i, j, k, Sign, DayNum AS LONG, Julian, t AS STRING
LOCAL RTxt() AS STRING
LOCAL Jan1 AS IPOWERTIME
LET   Jan1 = CLASS "PowerTime
LOCAL Now AS IPOWERTIME
LET   Now = CLASS "PowerTime
DIM RTxt(1 TO 20) AS STRING
   Jan1.Now: Now.Now                                              ' Set both dates to NOW
   Jan1.NewDate(Jan1.Year, 1, 1)                                  ' Set Day 1 back to Jan 1
   Now.TimeDiff(Jan1, Sign, DayNum): INCR DayNum                  ' Get # days difference (Day Number)
   Julian = FORMAT$(Now.Year - 2000, "00") + FORMAT$(DayNum, "000") ' Make it YYDDD format

   '----- Setup the basic resource strings
   RTxt(01) = "#RESOURCE VERSIONINFO"
   RTxt(02) = "#RESOURCE FILEVERSION X, Y, 0, ZZZZ"
   RTxt(03) = "#RESOURCE PRODUCTVERSION X, Y, 0, ZZZZ"
   RTxt(04) = "#RESOURCE STRINGINFO  '0409', '04E4'"
   RTxt(05) = "#RESOURCE VERSION$ 'CompanyName',      'SPFLite'"
   RTxt(06) = "#RESOURCE VERSION$ 'FileDescription',  'SPFLite Editor'"
   RTxt(07) = "#RESOURCE VERSION$ 'FileVersion',      'X.Y.ZZZZ'"
   RTxt(08) = "#RESOURCE VERSION$ 'InternalName',     'SPFLite'"
   RTxt(09) = "#RESOURCE VERSION$ 'OriginalFilename', 'SPFLite'"
   RTxt(10) = "#RESOURCE VERSION$ 'LegalCopyright',   '© 2014-2018 George D. Deluca, Robert L. Hodge'"
   RTxt(11) = "#RESOURCE VERSION$ 'ProductName',      'SPFLite'"
   RTxt(12) = "#RESOURCE VERSION$ 'ProductVersion',   'X.Y.ZZZZ'"
   RTxt(13) = "#RESOURCE VERSION$ 'Comments',         'None'"
   RTxt(14) = "#RESOURCE VERSION$ 'Author',           'George D Deluca'"
   RTxt(15) = "#RESOURCE VERSION$ 'VersionDate',      'YYYYMMDD'"
   RTxt(16) = "#RESOURCE VERSION$ 'Compiler',         'PowerBasic 10.04.0108'"

   '----- Update version and date in the resource strings
   REPLACE "X" WITH $V1 IN RTxt(02)
   REPLACE "X" WITH $V1 IN RTxt(03)
   REPLACE "X" WITH $V1 IN RTxt(07)
   REPLACE "X" WITH $V1 IN RTxt(12)
   REPLACE "Y" WITH $V2 IN RTxt(02)
   REPLACE "Y" WITH $V2 IN RTxt(03)
   REPLACE "Y" WITH $V2 IN RTxt(07)
   REPLACE "Y" WITH $V2 IN RTxt(12)
   REPLACE "ZZZZ" WITH MID$(Julian, 2) IN RTxt(02)
   REPLACE "ZZZZ" WITH MID$(Julian, 2) IN RTxt(03)
   REPLACE "ZZZZ" WITH MID$(Julian, 2) IN RTxt(07)
   REPLACE "ZZZZ" WITH MID$(Julian, 2) IN RTxt(12)
   REPLACE "YYYY" WITH MID$(DATE$, 7, 4) IN RTxt(15)
   REPLACE "MM" WITH MID$(DATE$, 1, 2) IN RTxt(15)
   REPLACE "DD" WITH MID$(DATE$, 4, 2) IN RTxt(15)

   FOR i = 1 TO 16
      t = RTxt(i)
      REPLACE "'" WITH $DQ IN t
      #DEBUG PRINT t
   NEXT i

   FNum = FREEFILE                                                ' Load the Recent table
   OPEN "_Version.inc" FOR OUTPUT AS #FNum                        ' Open the INC File
   FOR i = 1 TO 16                                                ' Loop dumping the lines
      t = RTxt(i)                                                 '
      REPLACE "'" WITH $DQ IN t                                   ' Single into double quotes
      PRINT #FNum, t                                              '
   NEXT i                                                         '
   SETEOF #FNum                                                   '
   CLOSE #FNum                                                    '

END FUNCTION
