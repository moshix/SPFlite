/* REXX */
Say "Additional Text: "
PARSE PULL FTxt
if Ftxt = "" then FTxt = "Normal"
FDate = DATE("S")
FTime = TIME("N")
FTime = translate(FTime, ".", ":")
ADDRESS CMD '"C:\Program Files\7-Zip\7z.exe" a -r "E:\GDrive\Backups\SPFLite\SPFLite10.'FDATE'.'FTIME'.'FTxt'.ZIP" *.*'
