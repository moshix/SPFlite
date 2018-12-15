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

'----- Get only Property Macro
MACRO GProp(pName, varType) = PROPERTY GET pName() AS varType: PROPERTY = pName: END PROPERTY

'----- Set only Property Macro
MACRO SProp(pName, varType) = PROPERTY SET pName(BYREF v AS varType): pName = v: END PROPERTY

'----- Get/Set Property Macro
MACRO GSProp(pName, varType) = GProp(pName, varType): SProp(pName, varType)

'----- Set Debug module trace
MACRO MEntry() = gCrashList(gCrashCtr) = FUNCNAME$ : INCR gCrashCtr

'----- Remove Debug module trace
MACRO MExit() = gcrashCtr = MAX(0, gCrashCtr - 1)
MACRO MExitSub() = gcrashCtr = MAX(0, gCrashCtr - 1): EXIT SUB
MACRO MExitFunc() = gcrashCtr = MAX(0, gCrashCtr - 1): EXIT FUNCTION
MACRO MExitMeth() = gcrashCtr = MAX(0, gCrashCtr - 1): EXIT METHOD
MACRO MErrExit(a, b)    = scError(a, b): MExitMeth
MACRO MLErrExit(a, b)   = AFList(a).Msg = b: DoSet(%Msg): MExitMeth

'----- Running in debugger?
MACRO IsPBDebuggerOn = (ISTRUE LEN(ENVIRON$("PBDEBUG")))

'----- scError code if not gMacroMode
MACRO NMac(errcd) = IIF(gMacroMode, errcd, %eNone)

'----- Misc macros
MACRO IsEQ(a,b)         = ISTRUE (StrCmpr(a, b) = 0)
MACRO IsNE(a,b)         = ISTRUE (StrCmpr(a, b) <> 0)
MACRO IsMacro(mname)    = ISFILE(ENV.MacrosPath + mname + ".MACRO")
MACRO IsWordChr(char)   = INSTR(TP.PrfPWord + "._", char)
MACRO IsCCmnd           = Zone = %CursCmnd
MACRO IsCScrl           = Zone = %CursScrl
MACRO IsCData           = Zone = %CursData
MACRO IsCLinN           = Zone = %CursLinN
MACRO IsCLLin           = Zone = %CursLLin
MACRO IsCLCmd           = Zone = %CursLCmd
MACRO IsCLNote          = Zone = %CursLNote
MACRO IsCBad            = Zone = %CursBad
MACRO IsCfL1Picture     = ISTRUE BIT(cfFlag, %CRTL1Picture)
MACRO IsCfL1CaseComp    = ISTRUE BIT(cfFlag, %CRTL1CaseComp)
MACRO IsCfL2Picture     = ISTRUE BIT(cfFlag, %CRTL2Picture)
MACRO IsCfL2Format      = ISTRUE BIT(cfFlag, %CRTL2Format)
MACRO IsCfL2CaseInComp  = ISTRUE BIT(cfFlag, %CRTL2CaseInComp)
MACRO IsCfL2Trunc       = ISTRUE BIT(cfFlag, %CRTL2Trunc)
MACRO IsCfL2Map         = ISTRUE BIT(cfFlag, %CRTL2Map)
MACRO IsCfL2Exec        = ISTRUE BIT(cfFlag, %CRTL2Exec)
MACRO IsCfRangePass     = ISTRUE BIT(cfFlag, %CRTRangePass)
MACRO IsCfL1RegEx       = ISTRUE BIT(cfFlag, %CRTL1RegEx)
MACRO IsCfLit1          = ISTRUE BIT(cfFlag, %CRTLit1)
MACRO IsCfLit2          = ISTRUE BIT(cfFlag, %CRTLit2)
MACRO IsCfL1Hex         = ISTRUE BIT(cfFlag, %CRTL1Hex)
MACRO IsCfAll           = ISTRUE BIT(cfFlag, %CRTAll)
MACRO IsCfNF            = ISTRUE BIT(cfFlag, %CRTNF)
MACRO IsCfU             = ISTRUE BIT(cfFlag, %CRTU)
MACRO IsCfNU            = ISTRUE BIT(cfFlag, %CRTNU)
MACRO IsCfX             = ISTRUE BIT(cfFlag, %CRTX)
MACRO IsCfNX            = ISTRUE BIT(cfFlag, %CRTNX)
MACRO IsCfMX            = ISTRUE BIT(cfFlag, %CRTMX)
MACRO IsCfWord          = ISTRUE BIT(cfFlag, %CRTWord)
MACRO IsCfPrefix        = ISTRUE BIT(cfFlag, %CRTPrefix)
MACRO IsCfSuffix        = ISTRUE BIT(cfFlag, %CRTSuffix)
MACRO IsCfNext          = ISTRUE BIT(cfFlag, %CRTNext)
MACRO IsCfPrev          = ISTRUE BIT(cfFlag, %CRTPrev)
MACRO IsCfFirst         = ISTRUE BIT(cfFlag, %CRTFirst)
MACRO IsCfLast          = ISTRUE BIT(cfFlag, %CRTLast)
MACRO IsCfLM            = ISTRUE BIT(cfFlag, %CRTLM)
MACRO IsCfRM            = ISTRUE BIT(cfFlag, %CRTRM)
MACRO IsCfOff           = ISTRUE BIT(cfFlag, %CRTOff)
MACRO IsCfOn            = ISTRUE BIT(cfFlag, %CRTOn)
MACRO IsCfSet           = ISTRUE BIT(cfFlag, %CRTSet)
MACRO IsCfAssert        = ISTRUE BIT(cfFlag, %CRTAssert)
MACRO IsCfToggle        = ISTRUE BIT(cfFlag, %CRTToggle)
MACRO IsCfFCol          = ISTRUE BIT(cfFlag, %CRTFCol)
MACRO IsCfTCol          = ISTRUE BIT(cfFlag, %CRTTCol)
MACRO IsCfTop           = ISTRUE BIT(cfFlag, %CRTTop)
MACRO IsCfPStd          = ISTRUE BIT(cfFlag, %CRTPStd)
MACRO IsCfStd           = ISTRUE BIT(cfFlag, %CRTStd)
MACRO IsCfSolid         = ISTRUE BIT(cfFlag, %CRTSolid)
MACRO IsCfMSolid        = ISTRUE BIT(cfFlag, %CRTMSolid)
MACRO IsCfLeft          = ISTRUE BIT(cfFlag, %CRTLeft)
MACRO IsCfRight         = ISTRUE BIT(cfFlag, %CRTRight)
MACRO IsCfCS            = ISTRUE BIT(cfFlag, %CRTCS)
MACRO IsCfDS            = ISTRUE BIT(cfFlag, %CRTDS)
MACRO IsCfDX            = ISTRUE BIT(cfFlag, %CRTDX)

MACRO IsLfAll           = ISTRUE BIT(LocFlag, %LocAll)
MACRO IsLfTop           = ISTRUE BIT(LocFlag, %LocTop)
MACRO IsLfPrev          = ISTRUE BIT(LocFlag, %LocPrev)
MACRO IsLfFirst         = ISTRUE BIT(LocFlag, %LocFirst)
MACRO IsLfLast          = ISTRUE BIT(LocFlag, %LocLast)
MACRO IsLfNot           = ISTRUE BIT(LocFlag, %LocNot)
MACRO IsLfChange        = ISTRUE BIT(LocFlag, %LocChange)
MACRO IsLfCurrent       = ISTRUE BIT(LocFlag, %LocCurrent)
MACRO IsLfFind          = ISTRUE BIT(LocFlag, %LocFind)
MACRO IsLfCommand       = ISTRUE BIT(LocFlag, %LocCommand)
MACRO IsLfError         = ISTRUE BIT(LocFlag, %LocError)
MACRO IsLfExcluded      = ISTRUE BIT(LocFlag, %LocExcluded)
MACRO IsLfU             = ISTRUE BIT(LocFlag, %LocU)
MACRO IsLfNU            = ISTRUE BIT(LocFlag, %LocNU)
MACRO IsLfLabel         = ISTRUE BIT(LocFlag, %LocLabel)
MACRO IsLfNote          = ISTRUE BIT(LocFlag, %LocNote)
MACRO IsLfSpecial       = ISTRUE BIT(LocFlag, %LocSpecial)
MACRO IsLfFile          = ISTRUE BIT(LocFlag, %LocFile)
MACRO IsLfPage          = ISTRUE BIT(LocFlag, %LocPage)
MACRO IsLfKeep          = ISTRUE BIT(LocFlag, %LocKeep)
MACRO IsLfTag           = ISTRUE BIT(LocFlag, %LocTag)
MACRO IsLfLong          = ISTRUE BIT(LocFlag, %LocLong)
MACRO IsLfSize          = ISTRUE BIT(LocFlag, %LocSize)
MACRO IsLfColor         = ISTRUE BIT(LocFlag, %LocColor)
MACRO IsLfStd           = ISTRUE BIT(LocFlag, %LocStd)
MACRO IsLfMX            = ISTRUE BIT(LocFlag, %LocMX)
MACRO IsLfNX            = ISTRUE BIT(LocFlag, %LocNX)

MACRO CfSet(flag)       = BIT SET cfFlag, flag
MACRO CfClear(flag)     = BIT RESET cfFlag, flag

MACRO LfSet(flag)       = BIT SET LocFlag, flag
MACRO LfClear(flag)     = BIT RESET LocFlag, flag

MACRO DOSet(flag)       = AttnDo OR= flag
MACRO TPDOSet(flag)     = TP.AttnDo = TP.AttnDo OR flag
MACRO DOClear(flag)     = AttnDo = AttnDo AND (&HFFFFFFFF - flag)

MACRO IsDOLoadReq       = (AttnDo AND %LoadReq)
MACRO IsDOLoadData      = (AttnDo AND %LoadData)
MACRO IsDOSaveReq       = (AttnDo AND %SaveReq)
MACRO IsDORefresh       = (AttnDo AND %Refresh)
MACRO IsDOAttention     = (AttnDo AND %Attention)
MACRO IsDOMsg           = (AttnDo AND %Msg)
MACRO IsDOMarkKill      = (AttnDo AND %MarkKill)
MACRO IsDOMarkKillSkip  = (AttnDo AND %MarkKillSkip)

MACRO IsLData(i)        = ISTRUE (L(i).LFlag AND %Data)
MACRO IsLTop(i)         = ISTRUE (L(i).LFlag AND %Top)
MACRO IsLBottom(i)      = ISTRUE (L(i).LFlag AND %Bottom)
MACRO IsLFile(i)        = ISTRUE (L(i).LFlag AND %File)
MACRO IsLXClude(i)      = ISTRUE (L(i).LFlag AND %XClude)
MACRO IsLTabs(i)        = ISTRUE (L(i).LFlag AND %Tabs)
MACRO IsLBounds(i)      = ISTRUE (L(i).LFlag AND %Bounds)
MACRO IsLCols(i)        = ISTRUE (L(i).LFlag AND %Cols)
MACRO IsLInvisible(i)   = ISTRUE (L(i).LFlag AND %Invisible)
MACRO IsLInsertLine(i)  = ISTRUE (L(i).LFlag AND %InsertLine)
MACRO IsLEQChange(i)    = ISTRUE (L(i).LFlag AND %EQChange)
MACRO IsLWord(i)        = ISTRUE (L(i).LFlag AND %Word)
MACRO IsLMark(i)        = ISTRUE (L(i).LFlag AND %Mark)
MACRO IsLMask(i)        = ISTRUE (L(i).LFlag AND %Mask)
MACRO IsLProf(i)        = ISTRUE (L(i).LFlag AND %Prof)
MACRO IsLCursor(i)      = ISTRUE (L(i).LFlag AND %Cursor)
MACRO IsLTopScreen(i)   = ISTRUE (L(i).LFlag AND %TopScreen)
MACRO IsLScroll(i)      = ISTRUE (L(i).LFlag AND %Scroll)
MACRO IsLXPtr(i)        = ISTRUE (L(i).LFlag AND %XPtr)
MACRO IsLPage(i)        = ISTRUE (L(i).LFlag AND %Page)
MACRO IsLNote(i)        = ISTRUE (L(i).LFlag AND %Note)
MACRO IsLPopped(i)      = ISTRUE (L(i).LFlag AND %Popped)
MACRO IsLUser(i)        = ISTRUE (L(i).LFlag AND %User)
MACRO IsLChange(i)      = ISTRUE (L(i).LFlag AND %EQChange)

MACRO IsFMTab           = ISTRUE (TP.TMode AND %MFMTab)
MACRO IsEdit            = ISTRUE (TP.TMode AND %MEdit)
MACRO IsMEdit           = ISTRUE (TP.TMode AND %MMEdit)
MACRO IsSetEdit         = ISTRUE (TP.TMode AND %MSetEdit)
MACRO IsClip            = ISTRUE (TP.TMode AND %MClip)
MACRO IsView            = ISTRUE (TP.TMode AND %MView)
MACRO IsBrowse          = ISTRUE (TP.TMode AND %MBrowse)
MACRO IsOpen            = ISTRUE (TP.TMode AND %MOpen)
MACRO IsClipLoad        = ISTRUE (TP.TMode AND %MClipLoad)
MACRO IsClipSave        = ISTRUE (TP.TMode AND %MClipSave)
MACRO IsKeyMap          = ISTRUE (TP.TMode AND %MKeyMap)

MACRO IsENVFMTab        = ISTRUE (Env.PMode AND %MFMTab)
MACRO IsENVEdit         = ISTRUE (Env.PMode AND %MEdit)
MACRO IsENVMEdit        = ISTRUE (Env.PMode AND %MMEdit)
MACRO IsENVSetEdit      = ISTRUE (Env.PMode AND %MSetEdit)
MACRO IsENVClip         = ISTRUE (Env.PMode AND %MClip)
MACRO IsENVView         = ISTRUE (Env.PMode AND %MView)
MACRO IsENVBrowse       = ISTRUE (Env.PMode AND %MBrowse)
MACRO IsENVOpen         = ISTRUE (Env.PMode AND %MOpen)
MACRO IsENVClipLoad     = ISTRUE (Env.PMode AND %MClipLoad)
MACRO IsENVClipSave     = ISTRUE (Env.PMode AND %MClipSave)
MACRO IsENVKeyMap       = ISTRUE (Env.PMode AND %MKeyMap)

MACRO IsBndsAFlag        = ISTRUE (Flag AND %BndsAFlag)
MACRO OnBndsAFlag        = Flag OR= %BndsAFlag
MACRO OffBndsAFlag       = Flag AND= (&HFFFFFFFF - %BndsAFlag)
MACRO IsWordAFlag        = ISTRUE (Flag AND %WordAFlag)
MACRO OnWordAFlag        = Flag OR= %WordAFlag
MACRO OffWordAFlag       = Flag AND= (&HFFFFFFFF - %WordAFlag)
MACRO IsMarkAFlag        = ISTRUE (Flag AND %MarkAFlag)
MACRO OnMarkAFlag        = Flag OR= %MarkAFlag
MACRO OffMarkAFlag       = Flag AND= (&HFFFFFFFF - %MarkAFlag)
MACRO IsMaskAFlag        = ISTRUE (Flag AND %MaskAFlag)
MACRO OnMaskAFlag        = Flag OR= %MaskAFlag
MACRO OffMaskAFlag       = Flag AND= (&HFFFFFFFF - %MaskAFlag)
MACRO IsTabsAFlag        = ISTRUE (Flag AND %TabsAFlag)
MACRO OnTabsAFlag        = Flag OR= %TabsAFlag
MACRO OffTabsAFlag       = Flag AND= (&HFFFFFFFF - %TabsAFlag)
MACRO IsUndoFlag         = ISTRUE (Flag AND %UndoFlag)
MACRO OnUndoFlag         = Flag OR= %UndoFlag
MACRO OffUndoFlag        = Flag AND= (&HFFFFFFFF - %UndoFlag)
MACRO IsClrCase          = ISTRUE (Flag AND %ClrCase)
MACRO OnClrCase          = Flag OR= %ClrCase
MACRO OffClrCase         = Flag AND= (&HFFFFFFFF - %ClrCase)
MACRO IsBackupDone       = ISTRUE (Flag AND %BackupDone)
MACRO OnBackupDone       = Flag OR= %BackupDone
MACRO OffBackupDone      = Flag AND= (&HFFFFFFFF - %BackupDone)
MACRO IsHideFlag         = ISTRUE (Flag AND %HideFlag)
MACRO IsTPHideFlag       = ISTRUE (TP.Flag AND %HideFlag)
MACRO OnHideFlag         = Flag OR= %HideFlag
MACRO OffHideFlag        = Flag AND= (&HFFFFFFFF - %HideFlag)
MACRO IsInsClnSupp       = ISTRUE (Flag AND %InsClnSupp)
MACRO OnInsClnSupp       = Flag OR= %InsClnSupp
MACRO OnTPInsClnSupp     = TP.Flag = (TP.Flag OR %InsClnSupp)
MACRO OffInsClnSupp      = Flag AND= (&HFFFFFFFF - %InsClnSupp)
MACRO IsModdFlag         = ISTRUE (Flag AND %ModdFlag)
MACRO IsTPModdFlag       = ISTRUE (TP.Flag AND %ModdFlag)
MACRO OnModdFlag         = Flag OR= %ModdFlag
MACRO OffModdFlag        = Flag AND= (&HFFFFFFFF - %ModdFlag)
MACRO IsModdEver         = ISTRUE (Flag AND %ModdEver)
MACRO OnModdEver         = Flag OR= %ModdEver
MACRO OffModdEver        = Flag AND= (&HFFFFFFFF - %ModdEver)
MACRO IsModdLast         = ISTRUE (Flag AND %ModdLast)
MACRO OnModdLast         = Flag OR= %ModdLast
MACRO OffModdLast        = Flag AND= (&HFFFFFFFF - %ModdLast)
MACRO IsOvrTextDel       = ISTRUE (Flag AND %OvrTextDel)
MACRO OnOvrTextDel       = Flag OR= %OvrTextDel
MACRO OffOvrTextDel      = Flag AND= (&HFFFFFFFF - %OvrTextDel)
MACRO IsColSupp          = ISTRUE (Flag AND %ColSupp)
MACRO OnColSupp          = Flag OR= %ColSupp
MACRO OffColSupp         = Flag AND= (&HFFFFFFFF - %ColSupp)
MACRO IsPFKInsert        = ISTRUE (Flag AND %PFKInsert)
MACRO OnPFKInsert        = Flag OR= %PFKInsert
MACRO OnTPPFKInsert      = TP.Flag = (TP.Flag OR %PFKInsert)
MACRO OffPFKInsert       = Flag AND= (&HFFFFFFFF - %PFKInsert)
MACRO IsPTypeMode        = ISTRUE (Flag AND %PTypeMode)
MACRO IsTPPTypeMode      = ISTRUE (TP.Flag AND %PTypeMode)
MACRO OnPTypeMode        = Flag OR= %PTypeMode
MACRO OffPTypeMode       = Flag AND= (&HFFFFFFFF - %PTypeMode)
MACRO OffTPPTypeMode     = TP.Flag = (TP.Flag AND (&HFFFFFFFF - %PTypeMode))
MACRO IsRenumFlag        = ISTRUE (Flag AND %RenumFlag)
MACRO OnRenumFlag        = Flag OR= %RenumFlag
MACRO OffRenumFlag       = Flag AND= (&HFFFFFFFF - %RenumFlag)
MACRO IsRevOnce          = ISTRUE (Flag AND %RevOnce)
MACRO OnRevOnce          = Flag OR= %RevOnce
MACRO OffRevOnce         = Flag AND= (&HFFFFFFFF - %RevOnce)
MACRO IsStateExist       = ISTRUE (Flag AND %StateExist)
MACRO IsTPStateExist     = ISTRUE (TP.Flag AND %StateExist)
MACRO OnStateExist       = Flag OR= %StateExist
MACRO OffStateExist      = Flag AND= (&HFFFFFFFF - %StateExist)
MACRO IsStateForce       = ISTRUE (Flag AND %StateForce)
MACRO OnStateForce       = Flag OR= %StateForce
MACRO OffStateForce      = Flag AND= (&HFFFFFFFF - %StateForce)
MACRO IsClrFlag          = ISTRUE (Flag AND %ClrFlag)
MACRO IsTPClrFlag        = ISTRUE (TP.Flag AND %ClrFlag)
MACRO OnClrFlag          = Flag OR= %ClrFlag
MACRO OffClrFlag         = Flag AND= (&HFFFFFFFF - %ClrFlag)
MACRO IsMarkActive       = ISTRUE (Flag AND %MarkActive)
MACRO IsTPMarkActive     = ISTRUE (TP.Flag AND %MarkActive)
MACRO OnMarkActive       = Flag OR= %MarkActive
MACRO OffMarkActive      = Flag AND= (&HFFFFFFFF - %MarkActive)
MACRO OffTPMarkActive    = TP.Flag =  TP.Flag AND (&HFFFFFFFF - %MarkActive)
MACRO IsMarkDrawn        = ISTRUE (Flag AND %MarkDrawn)
MACRO IsTPMarkDrawn      = ISTRUE (TP.Flag AND %MarkDrawn)
MACRO OnMarkDrawn        = Flag OR= %MarkDrawn
MACRO OffMarkDrawn       = Flag AND= (&HFFFFFFFF - %MarkDrawn)
MACRO IsMiscActive       = ISTRUE (Flag AND %MiscActive)
MACRO IsTPMiscActive     = ISTRUE (TP.Flag AND %MiscActive)
MACRO OnMiscActive       = Flag OR= %MiscActive
MACRO OffMiscActive      = Flag AND= (&HFFFFFFFF - %MiscActive)
MACRO OffTPMiscActive    = TP.Flag =  TP.Flag AND (&HFFFFFFFF - %MiscActive)
MACRO IsMiscDrawn        = ISTRUE (Flag AND %MiscDrawn)
MACRO IsTPMiscDrawn      = ISTRUE (TP.Flag AND %MiscDrawn)
MACRO OnMiscDrawn        = Flag OR= %MiscDrawn
MACRO OffMiscDrawn       = Flag AND= (&HFFFFFFFF - %MiscDrawn)
MACRO IsSwapActive       = ISTRUE (Flag AND %SwapActive)
MACRO IsTPSwapActive     = ISTRUE (TP.Flag AND %SwapActive)
MACRO OnSwapActive       = Flag OR= %SwapActive
MACRO OffSwapActive      = Flag AND= (&HFFFFFFFF - %SwapActive)
MACRO OffTPSwapActive    = TP.Flag =  TP.Flag AND (&HFFFFFFFF - %SwapActive)
MACRO IsSwapDrawn        = ISTRUE (Flag AND %SwapDrawn)
MACRO IsTPSwapDrawn      = ISTRUE (TP.Flag AND %SwapDrawn)
MACRO OnSwapDrawn        = Flag OR= %SwapDrawn
MACRO OffSwapDrawn       = Flag AND= (&HFFFFFFFF - %SwapDrawn)
MACRO IsSlecSet          = ISTRUE (Flag AND %SlecSet)
MACRO IsTPSlecSet        = ISTRUE (TP.Flag AND %SlecSet)
MACRO OnSlecSet          = Flag OR= %SlecSet
MACRO OffSlecSet         = Flag AND= (&HFFFFFFFF - %SlecSet)
MACRO IsSlecActive       = ISTRUE (Flag AND %SlecActive)
MACRO IsTPSlecActive     = ISTRUE (TP.Flag AND %SlecActive)
MACRO OnSlecActive       = Flag OR= %SlecActive
MACRO OffSlecActive      = Flag AND= (&HFFFFFFFF - %SlecActive)


MACRO IsNSrtData         = ISTRUE (Flag2 AND %NSrtData)
MACRO IsTPNSrtData       = ISTRUE (TP.Flag2 AND %NSrtData)
MACRO OnNSrtData         = Flag2 OR= %NSrtData
MACRO OffNSrtData        = Flag2 AND= (&HFFFFFFFF - %NSrtData)
MACRO OffTPNSrtData      = TP.Flag2 = TP.Flag2  AND (&HFFFFFFFF - %NSrtData)
MACRO IsNsrtFlag         = ISTRUE (Flag2 AND %NsrtFlag)
MACRO IsNotNsrtFlag      = ISFALSE (Flag2 AND %NsrtFlag)
MACRO IsTPNsrtFlag       = ISTRUE (TP.Flag2 AND %NsrtFlag)
MACRO OnNsrtFlag         = Flag2 OR= %NsrtFlag
MACRO OffNsrtFlag        = Flag2 AND= (&HFFFFFFFF - %NsrtFlag)
MACRO IsNsrtSave         = ISTRUE (Flag2 AND %NsrtSave)
MACRO OnNsrtSave         = Flag2 OR= %NsrtSave
MACRO OnTPNsrtSave       = TP.Flag2 = (TP.Flag2 OR %NsrtSave)
MACRO OffNsrtSave        = Flag2 AND= (&HFFFFFFFF - %NsrtSave)
MACRO OffTPNsrtSave      = TP.Flag2 = TP.Flag2  AND (&HFFFFFFFF - %NsrtSave)


MACRO LFOn(i, fl)       = L(i).LFlag = (L(i).LFlag OR fl)
MACRO LFOff(i, fl)      = L(i).LFlag = (L(i).LFlag AND (&HFFFFFFFF - fl))
MACRO LFTog(i, fl)      = L(i).LFlag = (L(i).LFlag XOR fl)

MACRO LTxtG(i)          = L(i).@LTxt
MACRO LTxtGP(i)         = L(i).LTxt
MACRO LTxtGLen(i)       = LEN(L(i).@LTxt)
MACRO LTxtSPtr(i)       = STRPTR(L(i).@LTxt)
MACRO LAttrG(i)         = L(i).@LAttr
MACRO LAttrGP(i)        = L(i).LAttr
MACRO LAttrGLen(i)      = LEN(L(i).@LAttr)
MACRO LAttrSPtr(i)      = STRPTR(L(i).@LAttr)
MACRO LMIXG(i)          = L(i).LMix
MACRO LFIXG(i)          = L(i).LFix
MACRO LFlagG(i)         = L(i).LFlag
MACRO LTagG(i)          = L(i).LTag
MACRO LLNumG(i)         = L(i).LLNum
MACRO LLCtlG(i)         = L(i).LLCtl
MACRO LLblG(i)          = L(i).LLbl
MACRO LWrk1G(i)         = L(i).LWrk1
MACRO LWrk2G(i)         = L(i).LWrk2
MACRO LPCmdG(i)         = L(i).LPCmd

MACRO LAttrS(i)         = L(i).@LAttr
MACRO LMIXS(i)          = L(i).LMix
MACRO LFixS(i)          = L(i).LFix
MACRO LFlagS(i)         = L(i).LFlag
MACRO LTagS(i)          = L(i).LTag
MACRO LLNumS(i)         = L(i).LLNum
MACRO LLCtlS(i)         = L(i).LLCtl
MACRO LLblS(i)          = L(i).LLbl
MACRO LWrk1S(i)         = L(i).LWrk1
MACRO LWrk2S(i)         = L(i).LWrk2
MACRO LPCmdS(i)         = L(i).LPCmd

'----- Main Text storage Data
TYPE DataLine
   LLbl    AS STRING * 8                                          ' Label
   LLNum   AS STRING * 8                                          ' Original line number at last AttnFlag refresh
   LLCtl   AS STRING * 8                                          ' Line control after KB interaction
   LTag    AS STRING * 8                                          ' Tag
   LFlag   AS LONG                                                ' Type of entry
   LWrk1   AS LONG                                                ' Misc. work value (e.g. XX count)
   LWrk2   AS LONG                                                ' Work value. =PAGE> ctr; Note type
   LCol    AS LONG                                                ' Column number if %Cursor line
   LFix    AS LONG                                                ' FIND index counter
   LMix    AS LONG                                                ' MultiEdit File Index
   LPCmd   AS LONG                                                ' ID of last PCommand to access the line
   LTxtIX  AS LONG                                                ' Index of allocated T() dynamic string for Txt
   LTxt    AS STRING POINTER                                      ' Ptr to real text string
   LAttrIX AS LONG                                                ' Index of allocated T() dynamic string for Attr
   LAttr   AS WSTRING POINTER                                     ' Ptr to Text attribute string
END TYPE

%Data       = &H00000001                                          ' Equates for DataLine.LFlag
%Top        = &H00000002                                          '
%Bottom     = &H00000004                                          '
%Tabs       = &H00000008                                          '
%Bounds     = &H00000010                                          '
%Cols       = &H00000020                                          '
%Xclude     = &H00000040                                          '
%Invisible  = &H00000080                                          '
%InsertLine = &H00000100                                          '
%EQChange   = &H00000200                                          '
%Word       = &H00000400                                          '
%Mark       = &H00000800                                          '
%Mask       = &H00001000                                          '
%Prof       = &H00002000                                          '
%Cursor     = &H00004000                                          '
%TopScreen  = &H00008000                                          '
%Scroll     = &H00010000                                          '
%XPtr       = &H00020000                                          '
%Page       = &H00040000                                          '
%File       = &H00080000                                          '
%Note       = &H00100000                                          '
%Popped     = &H00200000                                          '
%User       = &H00400000                                          '
%NonTypable = %Top OR %Bottom OR %Cols OR %File OR %Xclude OR %Prof
%SpecialLine= %Top OR %Bottom OR %Tabs OR %Bounds OR %Cols OR %Xclude OR %Mask OR _
              %InsertLine OR %EQChange OR %Word OR %Mark OR %Prof OR %File OR %Note OR %Page
%NonLineNum = %Top OR %Bottom OR %Tabs OR %Bounds OR %Cols OR %Xclude OR %Mask OR _
              %Word OR %Mark OR %Prof OR %File OR %Note

'----- Undo Stuff
TYPE UndoType
   UFn        AS ASCIIZ * %MAX_PATH                               ' Filename of Udtsave data
   TFn        AS ASCIIZ * %MAX_PATH                               ' Filename of Txtsave data
   TWFn       AS ASCIIZ * %MAX_PATH                               ' Filename of WTxtsave data
   IXFn       AS ASCIIZ * %MAX_PATH                               ' Filename of IXsave data
   UBoundL    AS LONG                                             ' UBoundL
   UBoundT    AS LONG                                             ' UBoundT
   TopScr     AS LONG                                             ' Copy of TopScr
   LastLine   AS LONG                                             ' LastLine
   LastReal   AS LONG                                             ' LastReal
   Modified   AS LONG                                             ' Modified flag
   MEditMod   AS STRING * 255                                     ' MEdit Modified string
   TIME       AS DOUBLE                                           ' Time save was done
   pTIDX      AS STRING POINTER                                   ' Pointer to TIDX
   pL         AS LONG                                             ' Pointer to L() array
   pT         AS STRING POINTER                                   ' Pointer to T() array
   pTW        AS WSTRING POINTER                                  ' Pointer to TW() array
   Busy       AS LONG                                             ' Busy doing the write
   MCpyBusy   AS LONG                                             ' Memory copy Busy
END TYPE
%Valid = %True                                                    ' Valid Undo entry

'----- Line Control Data Areas

TYPE TouchEntry
   LinNo      AS LONG                                             ' Line Number
   LinCtl     AS STRING * 8                                       ' Line Ctl data
   LinCmd     AS STRING * 8                                       ' Line Ctl extracted command
   LinRpt     AS LONG                                             ' Line Ctl extracted numeric
   LinFlag    AS LONG                                             ' Line Ctl flags (K+-\/)
   LinCmdIX   AS LONG                                             ' Index to LCmd table entry
   LinCType   AS STRING * 1                                       ' Command type (I/S/D)
   LinCVar    AS STRING * 1                                       ' Command variation (e,g, ANote = A, MNOTE = M, etc.)
END TYPE

UNION LAsString
   LVar       AS LONG
   CVar       AS STRING * 4
END UNION

'---------- Status Bar stuff
TYPE SBarEntry
   SBName     AS STRING * 6                                       ' External name for the box
   SBID       AS STRING * 1                                       ' $ID character
   SBActive   AS STRING * 1                                       ' Listed in SBLayout
   SBAlign    AS STRING * 1                                       ' Box Alignment
   SBPosBar   AS STRING * 1                                       ' This is the Position Bar box
   SBMySelf   AS STRING * 2                                       ' My index number
   SBWidth    AS LONG                                             ' Width of this box
   SBDfScheme AS LONG                                             ' Default scheme number
   SBOVScheme AS LONG                                             ' Single Override Scheme
   SBText     AS STRING * 32                                      ' Message text
END TYPE                                                          '

%SBMode     = 01                                                  ' Statusbar box IDs
%SBLinNo    = 02                                                  '
%SBLines    = 03                                                  '
%SBCols     = 04                                                  '
%SBBnds     = 05                                                  '
%SBInsOvr   = 06                                                  '
%SBCaseWord = 07                                                  '
%SBChange   = 08                                                  '
%SBState    = 09                                                  '
%SBMisc     = 10                                                  '
%SBSelect   = 11                                                  '
%SBCaps     = 12                                                  '
%SBSource   = 13                                                  '
%SBEOL      = 14                                                  '
%SBPad      = 15                                                  '
$AllStatusBarBoxes = "ABCDEFGHIJKLMN"                             ' All boxes
$SBMode     = "A"                                                 ' Statusbar box call IDs
$SBLinNo    = "B"                                                 '
$SBLines    = "C"                                                 '
$SBCols     = "D"                                                 '
$SBBnds     = "E"                                                 '
$SBInsOvr   = "F"                                                 '
$SBCaseWord = "G"                                                 '
$SBChange   = "H"                                                 '
$SBState    = "I"                                                 '
$SBMisc     = "J"                                                 '
$SBSelect   = "K"                                                 '
$SBCaps     = "L"                                                 '
$SBSource   = "M"                                                 '
$SBEOL      = "N"                                                 '
$SBPad      = "O"                                                 '

%SCTxtLo    = 0                                                   ' Text scheme numbers
%SCTxtHi    = 1                                                   ' 1st 2 are locked
%SCUser3    = 2                                                   ' User defined
%SCUser4    = 3                                                   '
%SCUser5    = 4                                                   '
%SCUser6    = 5                                                   '
%SCUser7    = 6                                                   '
%SCUser8    = 7                                                   '
%SCUser9    = 8                                                   '
%SCUser10   = 9                                                   '
%SCUser11   = 10                                                  '
%SCUser12   = 11                                                  '
%SCUser13   = 12                                                  '
%SCUser14   = 13                                                  '
%SCUser15   = 14                                                  '
%SCUser16   = 15                                                  '

%SCLnoHi    = 16                                                  ' Internal basic screen colors
%SCLnoLo    = 17                                                  '
%SCATabMod  = 18                                                  '
%SCATabNMod = 19                                                  '
%SCITabMod  = 20                                                  '
%SCITabNMod = 21                                                  '
%SCPFK      = 22                                                  '
%SCStatus   = 23                                                  '
%SCFMTool   = 24                                                  '
%SCError    = 25                                                  '
%SCRsvrd6   = 26                                                  '
%SCRsvrd5   = 27                                                  '
%SCRsvrd4   = 28                                                  '
%SCRsvrd3   = 29                                                  '
%SCRsvrd2   = 30                                                  '
%SCRsvrd1   = 31

%SCHiBlue   = 32                                                  ' Hi-Lite colors
%SCHiGreen  = 33                                                  ' 1st 4 locked
%SCHiYellow = 34                                                  '
%SCHiRed    = 35                                                  '
%SCHiBlack  = 36                                                  '
%SCHiNavy   = 37                                                  '
%SCHiTeal   = 38                                                  '
%SCHiViolet = 39                                                  '
%SCHiOrange = 40                                                  '
%SCHiGray   = 41                                                  '
%SCHiLime   = 42                                                  '
%SCHiCyan   = 43                                                  '
%SCHiPink   = 44                                                  '
%SCHiMagenta= 45                                                  '
%SCHiWhite  = 46                                                  '

%SCCust     = -1                                                  '
%SCFG       = 1                                                   '
%SCBG1      = 2                                                   '
%SCBG2      = 3                                                   '

%AttrScheme = &H001F                                              ' Scheme bits
%AttrUC     = &H0080                                              ' Attr UpperCase
%AttrInv    = &H0040                                              ' Attr Invert Hi-Lite
%AttrUL     = &H0020                                              ' Attr Underline
%AttrHiLite = &H0F00                                              ' HiLite bits
%AttrHiSTD  = &H0000                                              ' HiLite STD
%AttrAll    = &HFFFF                                              ' All bits

$$TxtLo     = CHR$$(&H0000)                                       ' Normal Text Lo
$$TxtHi     = CHR$$(&H0001)                                       ' Normal Text Hi
$$TxtLoUL   = CHR$$(&H0020)                                       ' Normal Text Lo Underlined
$$TxtHiUL   = CHR$$(&H0021)                                       ' Normal Text Hi Underlined
$$LNoLo     = CHR$$(&H0011)                                       ' Line Number Lo
$$LNoHi     = CHR$$(&H0010)                                       ' Line Number Hi
$$LNoLoUL   = CHR$$(&H0031)                                       ' Line Number Lo Underlined
$$LNoHiUL   = CHR$$(&H0030)                                       ' Line Number Hi Underlined
$$PFK       = CHR$$(&H0016)                                       ' PFK text
$$PFKUL     = CHR$$(&H0036)                                       ' PFK Text underlined
$$FMTool    = CHR$$(&H0018)                                       ' FM Tool text
$$Error     = CHR$$(&H0019)                                       ' Error text
$$Selected  = CHR$$(&H001A)                                       ' Selected text
$$Blue      = CHR$$(&H0100)                                       ' Blue hilite
$$Green     = CHR$$(&H0200)                                       ' Green hilite
$$Yellow    = CHR$$(&H0300)                                       ' Yellow hilite
$$Red       = CHR$$(&H0400)                                       ' Red Hilite

TYPE LCtlScan                                                     ' Initial Scan data table entry (Pre-Process)
   FromIX     AS LONG                                             ' From line
   ToIX       AS LONG                                             ' To line
   ToIXOrig   AS LONG                                             ' To line (original)
   Cmd        AS STRING * 8                                       ' Command
   Repeat     AS LONG                                             ' Repeat factor
   CmdType    AS STRING * 1                                       ' Command Type
   CmdFlag    AS LONG                                             ' Command Flag (&, -, etc.)
   CmdIX      AS LONG                                             ' Index to LCmd table
   CmdVar     AS STRING * 1                                       ' Command Type
END TYPE

TYPE LCtlCmd                                                      ' Final Line Command entry (After validation)
   SrcCmd     AS STRING * 8                                       ' Source Command (C, M, CC, etc.)
   SrcFrom    AS LONG                                             ' Source Start Line number
   SrcTo      AS LONG                                             ' Source End Line number
   SrcRepeat  AS LONG                                             ' Source Repeat
   SrcFlag    AS LONG                                             ' Source Flag
   SrcVar     AS STRING * 1                                       ' Source Cmd Var
   DstCmd     AS STRING * 8                                       ' Dest Command (A, B, O/OO, H/HH etc.)
   DstFrom    AS LONG                                             ' Dest Start Line number
   DstTo      AS LONG                                             ' Dest End Line number
   DstRepeat  AS LONG                                             ' Dest Repeat
   DstFlag    AS LONG                                             ' Dest Flag
   DstVar     AS STRING * 1                                       ' Dest Cmd Var
END TYPE

%lCmdR         = 1                                                ' Line Cmd '&' modifier
%lCmdX         = 2                                                ' Line Cmd '+'    "
%lCmdNX        = 3                                                ' Line Cmd '-'    "
%lCmdBSlash    = 4                                                ' Line Cmd '\'    "
%lCmdFSlash    = 5                                                ' Line Cmd '/'    "
%lCmdForce0    = 6                                                ' Line Cmd has a specific zero numeric value

TYPE LCtlValEnt
   lcTxt   AS STRING * 8                                          ' Text
   lcIndex AS LONG                                                ' Index to gLCmdCode() table
   lcRange AS STRING * 1                                          ' Range Cmnd? Y=Yes, N=No
   lcNumV  AS STRING * 1                                          ' Numbers valid? Y=Yes, N=No
   lcType  AS STRING * 1                                          ' Cmd Type S=Source D=Destination I=Immediate
   lcKeep  AS STRING * 1                                          ' Line Keep allowed
   lcpm    AS STRING * 1                                          ' Line +/- allowed
   lcSlsh  AS STRING * 1                                          ' Line \ / allowed
   lcRFac  AS STRING * 1                                          ' Repeat/Num. operand etc.
   lcPCOK  AS STRING * 1                                          ' Primary command LINE OK
   lcbrOK  AS STRING * 1                                          ' Valid in Browse mode
END TYPE

TYPE SortCtl                                                      ' Sort Control
   sFROM   AS INTEGER
   sTO     AS INTEGER
   AorD    AS STRING * 1
   CASE    AS STRING * 1
END TYPE

TYPE CmdTable
   NAME     AS STRING * 16                                        ' Command name
   EditCode AS DWORD                                              ' Address of Edit mode command handler
   NumOps   AS LONG                                               ' Max number of operands
   LinOps   AS STRING * 1                                         ' Accept Line Operands (S/D/B/N)
   LinDual  AS STRING * 1                                         ' Accept Dual Line Operands (CC/MM and A/B)
   MacOK    AS STRING * 1                                         ' OK to run from MACRO SPFCommand call
   ChainOK  AS STRING * 1                                         ' Y/N eligible for Command Chaining
   SCROLL   AS LONG                                               ' A screen movement command Up, Down, Locate etc.
   FMCode   AS DWORD                                              ' Address of Edit mode command handler
   BrowseOK AS STRING * 1                                         ' Valid in Browse mode
   Oprands  AS STRING * 256                                       ' Valid KW operands, comma delimited
END TYPE


TYPE MacType
   MName   AS STRING * 16
   MType   AS STRING * 1
END TYPE

'----- ProfState settings
%StateOff     = 0
%StateOn      = 1
%StateFew     = 2
%StateMost    = 3

'----- AttnDo flags
%Attention    = &H00000001                                        ' Attention processing needed
%Refresh      = &H00000002                                        ' TP.DispScreen needed
%LoadReq      = &H00000004                                        ' FM load request data
%LoadData     = &H00000008                                        ' FM Load data based on Req
%SaveReq      = &H00000010                                        ' FM Save current Req data back to FILELIST
%Msg          = &H00000020                                        ' FM Line Message waiting
%MarkKill     = &H00000040                                        ' Kill any marked area
%MarkKillSkip = &H00000080                                        ' Skip killing once

'----- Cmd operand types
%OpMaxKW = 150                                                    ' Recognized KW
%OpStr   = 151                                                    ' Simple character string
%OpNum   = 152                                                    ' Numeric string
%OpQStr  = 154                                                    ' Quoted string
%OpQStrD = 155                                                    ' Quoted string was a "
%OpQStrS = 156                                                    ' Quoted string was a '
%OpQStrB = 157                                                    ' Quoted string was a `
%OpPStr  = 160                                                    ' Picture string
%OpRStr  = 162                                                    ' RegEx string
%OpXStr  = 164                                                    ' Hex string
%OpCStr  = 166                                                    ' Char string C'aaa'
%OpTStr  = 168                                                    ' Text string T'aaa'
%OpMStr  = 170                                                    ' Text string M'aaa'
%OpEStr  = 172                                                    ' Text string E'aaa'
%OpFStr  = 180                                                    ' Text string F'aaa'
%OpDotd  = 182                                                    ' Dotted string (.Label)
%OpTag   = 184                                                    ' Tagged string (:Tag)
%OpSqB   = 186                                                    ' [ @ left or right
%OpBrk   = 188                                                    ' ( @ left or right
%OpLPtr  = 200                                                    ' LPtr string !nnnn

'----- Macro function equates
%ARG_OPT = 1024
%ARG_DEF = 2048
%ARG_VAR = 4096

'----- Cmd Keyword equates
ENUM Z1 SINGULAR
   KWA         = 01
   KWALL
   KWAND
   KWANOTE
   KWAPPEND
   KWASSERT
   KWAUTO
   KWBNOTE
   KWC
   KWCANCEL
   KWCHANGE
   KWCHARS
   KWCNOTE
   KWCOLOR
   KWCOMMAND
   KWCOPY
   KWCR
   KWCREATE
   KWCRLF
   KWCS
   KWCSR
   KWCURRENT
   KWD
   KWDATA
   KWDEL
   KWDNOTE
   KWDOWN
   KWDS
   KWDX
   KWEDIT
   KWENOTE
   KWERROR
   KWEXCLUDE
   KWFILE
   KWFIND
   KWFIRST
   KWFNOTE
   KWFULL
   KWGNOTE
   KWHALF
   KWHIDE
   KWHNOTE
   KWHOME
   KWINOTE
   KWJNOTE
   KWKEEP
   KWKNOTE
   KWLABEL
   KWLAST
   KWLEFT
   KWLF
   KWLIST
   KWLM
   KWLNOTE
   KWLOAD
   KWLOCK
   KWLONG
   KWMAX
   KWMNOTE
   KWMX
   KWNEW
   KWNEXT
   KWNF
   KWNL
   KWNONE
   KWNONUM
   KWNOPROMPT
   KWNOT
   KWNNOTE
   KWNOTE
   KWNUM
   KWNU
   KWNX
   KWOFF
   KWON
   KWONOTE
   KWOPEN
   KWOR
   KWPAGE
   KWPNOTE
   KWPREFIX
   KWPREV
   KWPRIOR
   KWPROMPT
   KWQNOTE
   KWRAW
   KWREPLACE
   KWRESET
   KWREVERSE
   KWRETRIEVE
   KWRIGHT
   KWRM
   KWRNOTE
   KWSETUP
   KWSET
   KWSIZE
   KWSNOTE
   KWSOLID
   KWMSOLID
   KWSOURCE
   KWSPECIAL
   KWSTATE
   KWSTD
   KWMSTD
   KWPSTD
   KWSUFFIX
   KWT
   KWTAG
   KWTARGET
   KWTNOTE
   KWTOGGLE
   KWTOP
   KWTRUNC
   KWU
   KWUNLOCK
   KWUNOTE
   KWUP
   KWUSING
   KWVNOTE
   KWWNOTE
   KWWORD
   KWXNOTE
   KWYNOTE
   KWZNOTE
END ENUM

'----- Cursor setting stuff
%CurMacro = 24                                                    ' Positioning by a macro
%PosEnd   = 20                                                    ' Positioning by the kbEND function
%CError   = 18                                                    ' Command error
%LError   = 18                                                    ' Line Command error
%Insert   = 14                                                    ' Insert lines
%Position = 16                                                    ' Positioning, Scrolling etc.
%Change   = 12                                                    ' Change
%FChange  = 10                                                    ' Forced Change
%FLocate  = 9                                                     ' Forced Locate
%Find     = 8                                                     ' Find
%Locate   = 7                                                     ' Locate
%LineCmd  = 6                                                     ' Line commands
%EOL      = 4                                                     ' EOL position
%Low      = 1                                                     ' Anything else

'----- Criteria Parse Flags
$CRTAllOK = "A"                                                   ' ALL Subset, allow next keywords
%CRTTop = 0                                                       ' TOP
%CRTAll = 1                                                       ' ALL

$CRTSubSet = "S"                                                  ' Line Subset, allow next keywords
%CRTX = 2                                                         ' X
%CRTNX = 3                                                        ' NX

$CRTDirect = "D"                                                  ' Direction, allow next keywords
%CRTFirst = 4                                                     ' FIRST
%CRTLast = 5                                                      ' LAST
%CRTPrev = 6                                                      ' PREV
%CRTNext = 7                                                      ' NEXT

$CRTLR = "L"                                                      ' Left / Right
%CRTLeft = 8                                                      ' LEFT
%CRTRight = 9                                                     ' RIGHT

$CRTModifier = "M"                                                ' String find modifier, allow next keywords
%CRTWord = 10                                                     ' WORD
%CRTChars = 11                                                    ' CHARS
%CRTPrefix = 12                                                   ' PREFIX
%CRTSuffix = 13                                                   ' SUFFIX
%CRTLM = 14                                                       ' LM
%CRTRM = 15                                                       ' RM

$CRTShift = "H"                                                   ' Change Shift type modifier, allow next keywords
%CRTCS = 16                                                       ' CS
%CRTDS = 17                                                       ' DS

$CRTExclude = "E"                                                 ' Exclude line handling, allow next keywords
%CRTMX = 18                                                       ' MX
%CRTDX = 19                                                       ' DX

$CRTNegative = "N"                                                ' Not Found option, allow next keywords
%CRTNF = 20                                                       ' NF

$CRTCols = "C"                                                    ' Cols allowed, allow next keywords
%CRTFCol = 21                                                     ' From Col
%CRTTCol = 22                                                     ' To Col

$CRTL1 = "1"                                                      ' Allow literal 1, allow next keywords
%CRTLit1 = 23                                                     ' Literal 1
%CRTL1Picture = 24                                                ' Literal 1 is a P'...'
%CRTL1Hex = 25                                                    ' Literal 1 is a X'....'
%CRTL1RegEx = 26                                                  ' Literal 1 is a R'......'
%CRTL1CaseComp = 27                                               ' Forced Case sensitive compare
%CRTL1CaseInComp = 28                                             ' Forced Case insensitive compare

$CRTL2 = "2"                                                      ' Allow literal 2, allow next keywords
%CRTLit2 = 29                                                     ' Literal 2
%CRTL2Picture = 30                                                ' Literal 2 is a P'...'
%CRTL2Hex = 31                                                    ' Literal 2 is a X'....'
%CRTL2Format = 32                                                 ' Literal 2 is a F'...'
%CRTL2Map = 33                                                    ' Literal 2 is a M'...'
%CRTL2Exec = 34                                                   ' Literal 2 is a E'...'
%CRTL2CaseComp = 35                                               ' Forced Case sensitive compare
%CRTL2CaseInComp = 36                                             ' Forced Case insensitive compare
%CRTL2Trunc = 37                                                  ' TRUNC requested

$CRTTag = "T"                                                     ' Allow Tag values
%CRTOn = 38                                                       ' ON
%CRTOff = 39                                                      ' OFF
%CRTToggle = 40                                                   ' TOGGLE
%CRTAssert = 41                                                   ' ASSERT
%CRTSet = 42                                                      ' SET

$CRTPen = "P"                                                     ' Allow Pen Search values
%CRTHiClr = 45                                                    ' HiLite search color
%CRTStd = 51                                                      ' Std
%CRTMStd = 52                                                     ' -Std
%CRTSolid = 53                                                    ' Solid
%CRTMSolid = 54                                                   ' -Solid

$CRTCPen = "Q"                                                    ' Allow Pen change values
%CRTHiOff = 57                                                    ' HiLite Off color
%CRTHiOn  = 58                                                    ' HiLite ON color
%CRTPStd = 59                                                     ' +Std

$CRTUser = "U"
%CRTU = 60                                                        ' U
%CRTNU = 61                                                       ' NU

%CrtRangePass = 62                                                ' Internal RangePass flag

'----- LOCATE Flag Equates
%LocNext     = 1                                                  '
%LocPrev     = 2                                                  '
%LocFirst    = 3                                                  '
%LocLast     = 4                                                  '
%LocAll      = 5                                                  '
%LocNot      = 6                                                  '
%LocChange   = 7                                                  '
%LocCommand  = 8                                                  '
%LocError    = 9                                                  '
%LocExcluded = 10                                                 '
%LocLabel    = 11                                                 '
%LocSpecial  = 12                                                 '
%LocKeep     = 13                                                 '
%LocTag      = 14                                                 '
%LocLong     = 15                                                 '
%LocMX       = 16                                                 '
%LocPage     = 17                                                 '
%LocFind     = 18                                                 '
%LocCurrent  = 19                                                 '
%LocFile     = 20                                                 '
%LocNote     = 21                                                 '
%LocSize     = 22                                                 '
%LocColor    = 23
%LocStd      = 27                                                 '
%LocTop      = 28                                                 '
%LocU        = 29                                                 '
%LocNU       = 30                                                 '
%LocNX       = 31                                                 '

TYPE ClrCommModel
   Txt1       AS STRING * 32                                      ' Comment string
   sCol       AS STRING * 32                                      ' Comment column/dlm
   Clr        AS LONG                                             ' Comment color
   Fnd        AS LONG                                             ' Comment found
END TYPE

TYPE ClrExcIncModel
   Txt1       AS STRING * 32                                      ' Comment string
   sCol       AS LONG                                             ' Comment column
END TYPE

TYPE TxtEntry
   Length     AS LONG                                             ' Length of an entry
   Offset     AS LONG                                             ' Offset to 1st byte
   Clr        AS LONG                                             ' Color index
END TYPE

'----- Equates for CursMode                                       '
   %CursCmnd     = 1                                              ' In pCommand area
   %CursData     = 2                                              ' In Data text ares
   %CursLinN     = 3                                              ' In Line number area
   %CursScrl     = 4                                              ' In Scroll Amt. area
   %CursInternal = 5                                              ' sLin and sCol are set internally
   %CursBad      = 99                                             ' In non-typing area
   '----- FM entries
   %CursPath     = 6                                              ' In FM Path
   %CursMask     = 7                                              ' In FM Mask
   %CursEmpty    = 8                                              ' In FM New Empty
   %CursFavorite = 9                                              ' In FM Favorites
   %CursFileList = 10                                             ' In FM FileList
   %CursFound    = 11                                             ' In FM Found
   %CursRecent   = 12                                             ' In FM Recent Files
   %CursRPaths   = 13                                             ' In FM Recent Paths
   %CursOpen     = 14                                             ' In FM Open Files
   %CursFilePath = 15                                             ' In FM FilePath
   %CursProfiles = 16                                             ' In FM Profiles
   %CursLCmd     = 17                                             ' In FM LCmd segment
   %CursLNote    = 18                                             ' In FM Line Note segment
   %CursLLin     = 19                                             ' In FM Line elsewhere

'--------- Equates for errFlag
   %eNone = 0                                                     ' None
   %eRetrieve = 1                                                 ' RETRIEVE'd command
   %ePending = 2                                                  ' Line Command pending
   %eCmdPend = 4                                                  ' Command pending
   %eFail = 8                                                     ' Failure
   %eMultiple = 32                                                ' Multiple error override
   %eTerminate = 99                                               ' In termination

'--------- Equates for Tab Mode Flag and Env.PMode
%MFMTab     = &H00000001                                          ' Equates for Tab Mode Flag
%MBrowse    = &H00000002                                          '
%MView      = &H00000004                                          '
%MEdit      = &H00000008                                          '
%MClip      = &H00000010                                          '
%MSetEdit   = &H00000020                                          '
%MMedit     = &H00000040                                          '
%MClipLoad  = &H00000080                                          '
%MClipSave  = &H00000100                                          '
%MOpen      = &H00000200                                          '
%MKeyMap    = &H00000400                                          '

'--------- Equates for TP.Flag - Misc flags
%BndsAFlag  = &H00000001                                          ' BNDS line has been modified
%WordAFlag  = &H00000002                                          ' WORD line has been modified
%MarkAFlag  = &H00000004                                          ' MARK line has been modified
%MaskAFlag  = &H00000008                                          ' MASK line has been modified
%TabsAFlag  = &H00000010                                          ' TABS line has been modified
%UndoFlag   = &H00000020                                          ' UNDO save needed at next Attn
%ClrCase    = &H00000040                                          ' Clr KW's honor case
%BackupDone = &H00000100                                          ' File Backup has been done
%HideFlag   = &H00000200                                          ' Hide X'd lines
%InsClnSupp = &H00000400                                          ' Suppress Insert Line cleanup one time
%ModdFlag   = &H00000800                                          ' File is in Modified state
%ModdEver   = &H00001000                                          ' File was modified sometime since open
%ModdLast   = &H00002000                                          ' Modd status at last TabTitle
%OvrTextDel = &H00004000                                          ' OverlayTextDel flag
%ColSupp    = &H00008000                                          ' Suppress maintain col mode
%PFKInsert  = &H00010000                                          ' Command inserted by a PFK command
%PTypeMode  = &H00020000                                          ' PType Mode
%RenumFlag  = &H00040000                                          ' Renum needed
%RevOnce    = &H00080000                                          ' Reverse just once
%StateExist = &H00100000                                          ' STATE data exists
%StateForce = &H00200000                                          ' STATE force (set by STATE CREATE)
%ClrFlag    = &H00400000                                          ' AUTO color file available
%MarkActive = &H00800000                                          ' Mark - Marked area is active
%MarkDrawn  = &H01000000                                          ' Mark area has been drawn
%MiscActive = &H02000000                                          ' MiscMark - Marked area is active
%MiscDrawn  = &H04000000                                          ' MiscMark - Area is Marked on screen
%SwapActive = &H08000000                                          ' Swap - Marked area is active
%SwapDrawn  = &H10000000                                          ' Swap - Area is Marked on screen
%SlecSet    = &H20000000                                          ' Select is set
%SlecActive = &H40000000                                          ' Select is active

'--------- Equates for TP.Flag2 - Misc flags
%NSrtData   = &H00000001                                          ' KB Insert Data mode or not
%NSrtFlag   = &H00000002                                          ' KB Insert mode or not
%NSrtSave   = &H00000004                                          ' KB Insert mode saved


'----- Equates from shlobj.h not found in Win32API.INC
   %CIS            = 1
   %shgfp_type_current     = 0&                                   ' Current value for user, verify it exists
   %shgfp_type_default     = 1&                                   ' Default value, may not exist

   DECLARE FUNCTION Sub1( p1 AS ANY ) AS DWORD
   DECLARE FUNCTION Sub2( p1 AS ANY, p2 AS ANY ) AS DWORD
   DECLARE FUNCTION Sub3( p1 AS ANY, p2 AS ANY, p3 AS ANY ) AS DWORD
   DECLARE FUNCTION Sub5( p1 AS ANY, p2 AS ANY, p3 AS ANY, p4 AS ANY, p5 AS ANY ) AS DWORD

'----- Misc. Equates
%Strip = %True                                                    ' For sGetWord
%NoStrip = %False                                                 '
%QuoteSig = %True                                                 '
%QuoteNotSig = %False                                             '
'                                                                 ' Pseudo bit flag for DlgEditCallback mouse logic
%MK_ALT = &H0080                                                  ' ALT key pressed

$Numbers = "0123456789 *"                                         ' For Line control parsing
$Numeric = "0123456789"                                           '
$Hex     = "0123456789ABCDEF"                                     '
$NL      = CHR$(21)                                               ' EBCDIC NL character
$BlankLNo  = "        "                                           ' For line number stuff
$Quotes = $DQ + $SQ + CHR$(96)                                    ' Our 3 supported quotes
$AlphaNum = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_"
$DefCharSet = "!" + $DQ + "#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`" + _
                "abcdefghijklmnopqrstuvwxyz{|}~¬¢£€"
$Upper  = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
$Lower  = "abcdefghijklmnopqrstuvwxyz"
$UpperNat = CHR$(138,140,142,159,192,193,194,195,196,197,198,199,200, _  ' Š Œ Ž Ÿ À Á Â Ã Ä Å Æ Ç È
                 201,202,203,204,205,206,207,208,209,210,211,212,213, _  ' É Ê Ë Ì Í Î Ï Ð Ñ Ò Ó Ô Õ
                 214,216,217,218,219,220,221,222)                        ' Ö Ø Ù Ú Û Ü Ý Þ

$LowerNat = CHR$(154,156,158,223,224,225,226,227,228,229,230,231,232, _  ' š œ ž ß à á â ã ä å æ ç è
                 233,234,235,236,237,238,239,240,241,242,243,244,245, _  ' é ê ë ì í î ï ð ñ ò ó ô õ
                 246,249,250,251,252,253,254,248)                        ' ö ø ù ú û ü ý þ


$WORD    = "A-Z a-z 0-9"                                          ' For 'word', 'prefix' and 'suffix' tests
$EOLDLM = CHR$(10, 12, 13, 26)                                    ' $LF, $FF, $CR, $EOF

$RegUpper = "A-Z"                                                 ' Reg Expresssion UC
$RegUpperNat = "\x8A\x8C\x8E\x9F\xC0-\xD6\xD8-\xDE"               ' Reg Expresssion UC European chars
$RegLower = "a-z"                                                 ' Reg Expresssion LC
$RegLowerNat = "\x9A\x9C\x9E\xDF-\xF6\xF8-\xFF"                   ' Reg Expresssion LC European chars
$ConClose = "$CON:CLOSE"
$UpperSpec = CHR$(171, 173, 176 TO 185, 183 TO 218, 225 TO 250)   ' +, -, 0-9, A-Z, a-z
$NonBlank = CHR$(0 TO 31, 33 TO 255)                              ' Everything but blanks
$Empty = "(Empty)"                                                '

$AttrColor = " BGR"

'----- Constant equates for some SUB/FUNCTION calls

'----- sCRPFwd MoveMode values
%mVisible = 1
%mData    = 2
%mText    = 3
%mVisData = 4
%mPage    = 5
%mOData   = 62

'----- sResetFunc
%ResetExcluded = &H0001
%ResetCommand  = &H0002
%ResetCommandK = &H0004
%ResetLabel    = &H0008
%ResetSpecial  = &H0010
%ResetChange   = &H0020
%ResetTag      = &H0040
%ResetRetrieve = &H0080
%ResetUser     = &H0100
%ResetFind     = &H0200
%ResetClr      = &H0400
%ResetAll      = &H07FF

'----- Symbol scan states
%Normal_State = 0
%Quoted_State = 1
%Symbol_State = 3
%Format_State = 4

'----- sPrtPrint modes
%PrtReset      = 0
%PrtLine       = 1
%PrtNewLine    = 2
%PrtNewPage    = 3
%PrtFlushClose = 4

'----- Declare stuff for ShellLink (Cnvt .LNK to full path)
%CIS                  = 1                                               'Component is allowed in the same process space.
$CLSID_ShellLink      = GUID$("{00021401-0000-0000-C000-000000000046}") 'CHR$(  1, 20, 2, 0, 0, 0, 0, 0, 192, 0, 0, 0, 0, 0, 0, 70)
$IIDG_IShellLink      = GUID$("{000214EE-0000-0000-C000-000000000046}") 'CHR$(238, 20, 2, 0, 0, 0, 0, 0, 192, 0, 0, 0, 0, 0, 0, 70)
$IID_IPersistFile     = GUID$("{0000010B-0000-0000-C000-000000000046}") 'CHR$( 11,  1, 0, 0, 0, 0, 0, 0, 192, 0, 0, 0, 0, 0, 0, 70)

'----- Used in color selection dialogue
TYPE CustColor
   cc(16) AS LONG
END TYPE

'----- For searching window titles
TYPE WinInfoType
    WinTitle AS STRING * 256
    WinHandle AS LONG
END TYPE

'----- For PowerType support
TYPE PTypeTable
    tLin AS LONG
    sRow AS LONG
END TYPE

'----- FM data structures
TYPE FMFList
   FD            AS DIRDATA
   SizeInt       AS QUAD
   Flag          AS LONG
   CmdOff        AS LONG
   NoteOff       AS LONG
   Uniq          AS LONG
   RQIX          AS LONG
   LinesInt      AS LONG
   EXT           AS STRING * 32
   MSG           AS STRING * 64
   SIZE          AS STRING * 15
   LWTime        AS STRING * 17
   Note          AS STRING * %MAX_PATH
   PATH          AS ASCIIZ * %MAX_PATH
   Cmd           AS ASCIIZ * %MAX_PATH
END TYPE

'----- FM.Flag Equates
%FDirUp          = 4
%FDirDown        = 6
%FEntry          = 8
%FFLEntry        = 10
%FOpen           = 12
%FPath           = 14
%FFileList       = 16
%FProfile        = 18
%FDirLow         = 20
%FTotals         = 22
%FEnd            = 24

'----- FM.FMode Equates
%FMPath       = 1                                                 ' Normal Path/Mask mode
%FMOpen       = 2                                                 ' OPEN     files list
%FMProfiles   = 3                                                 ' Profiles files list
%FMRecPaths   = 4                                                 ' Recent   paths list
%FMNamed      = 5                                                 ' Named Favorites list
%FMFilelist   = 6                                                 ' Filelist type list

'----- FM Display field lengths
%FMLExt       = 8
%FMLDate      = 10
%FMLDateT     = 17
%FMLSizeS     = 6
%FMLSizeL     = 15
%FMLLines     = 9

'----- FM Help strings
$FMHAdd       = "Ádd to favorites  "                              ' A + 128
$FMHBrowse    = "Ârowse  "                                        ' B + 128
$FMHCancel    = "ÃÁÎcel  "                                        ' C + 128, A + 128, N + 128
$FMHClone     = "Ãlone  "                                         ' C + 128
$FMHDelete    = "Äelete  "                                        ' D + 128
$FMHDIR       = "ÄÉÒ  "                                           ' D + 128, I + 128, R + 128
$FMHEdit      = "Ådit  "                                          ' E + 128
$FMHEND       = "ÅÎÄ  "                                           ' E + 128, N + 128, D + 128
$FMHForget    = "Æorget  "                                        ' F + 128
$FMHJob       = "Êob submit  "                                    ' J + 128
$FMHLines     = "Ìines  "                                         ' L + 128
$FMHMEdit     = "ÍultiEdit  "                                     ' M + 128
$FMHNorm      = "Îorm  "                                          ' N + 128
$FMHOpen      = "Ïpen  "                                          ' O + 128
$FMHPrint     = "Ðrint  "                                         ' P + 128
$FMHPurge     = "PÕrge  "                                         ' U + 128
$FMHRename    = "Òename  "                                        ' R + 128
$FMHSave      = "ÓÁÖÅ  "                                          ' S + 128, A + 128, V + 128, E + 128
$FMHSelect    = "Óelect  "                                        ' S + 128
$FMHTouch     = "Ôouch  "                                         ' T + 128
$FMHWDIR      = "×ÄÉÒ  "                                          ' W + 128, D + 128, I + 128, R + 128
$FMHAll       = "ÁÌÌ  "                                           ' A + 128, L + 128, L + 128
$FMHView      = "Öiew  "                                          ' V + 128
$FMHExclude   = "EØclude  "                                       ' X + 128

'----- Misc stuff
UNION ClrCnvt
   Clr  AS LONG
   c(1 TO 4) AS BYTE
END UNION

UNION Both
   lft AS FILETIME
   lq  AS QUAD
END UNION

'----- Directory Watching data
TYPE WatchData                                                    ' Parameter list for the Thread
   gWatchDir         AS ASCIIZ * %MAX_PATH                        ' The directory to watch
   gWatchFile        AS ASCIIZ * %MAX_PATH                        ' The file to watch
   gWatchThread      AS LONG                                      ' Thread address
   gEvent            AS LONG                                      ' Event to be signalled to stop the thread
   gChanged          AS LONG                                      ' Answer area from thread to signal a change
   gActive           AS LONG                                      ' Thread is active
   gFileAttrib       AS DWORD                                     ' File Attributes
   gFileCRTime       AS QUAD                                      ' Creation Timestamp when last seen
   gFileLWTime       AS QUAD                                      ' Modified Timestamp when last seen
   gFileSizeHigh     AS DWORD                                     ' File Size when last seen
   gFileSizeLow      AS DWORD                                     '
   gPgNumber         AS INTEGER                                   ' Tab # we're watching for
   gInUse            AS LONG                                      ' InUse - %True, Available - %False
   gFlag             AS STRING * 1                                ' Modified flag
END TYPE
%WatchStart = 1                                                   ' Start watching directory
%WatchCheck = 2                                                   ' Get current Changed status
%WatchEnd   = 3                                                   ' Stop watching, return result

'----- Size of Key Map Table (in INIBASIC)
%Max_KeyMap = 219
%KPad_Start = 166
%KPad_End   = 182
UNION kbMsg                                                       ' To support Keyboard hook passing a
   MsgWparam         AS LONG                                      ' text string via %WM_USER param
   kbString          AS STRING * 4                                '
   kbInt(0 TO 1)     AS INTEGER                                   '
END UNION                                                         '


'----- Picture mapping equates
%PicMap_Literal            = &H00000001                           ' Literal data, non-alphabetic
%PicMap_Any                = &H00000002                           ' P'='    Any char
%PicMap_NonBlank           = &H00000004                           ' P'¬'    Any nonblank char; also P'^'
%PicMap_NonDisplay         = &H00000008                           ' P'.'    Any non-displayable char
%PicMap_Numeric            = &H00000010                           ' P'#'    Any numeric character, 0-9
%PicMap_NonNumeric         = &H00000020                           ' P'-'    Any nonnumeric char
%PicMap_Alpha              = &H00000040                           ' P'@'    Any alphabetic char, UC or LC
%PicMap_Lower              = &H00000080                           ' P'<'    Any lowercase alphabetic char
%PicMap_Upper              = &H00000100                           ' P'>'    Any uppercase alphabetic char
%PicMap_Special            = &H00000200                           ' P'$'    Any char not alpha or numeric
%PicMap_WORD_Char          = &H00000400                           ' P'%'    Any WORD char
%PicMap_Non_WORD_Char      = &H00000800                           ' P'&'    Any non-WORD char
%PicMap_LMargin            = &H00001000                           ' P'{'    LM + leading spaces
%PicMap_RMargin            = &H00002000                           ' P'}'    RM + trailing spaces
%PicMap_LiteralAlpha       = &H00004000                           ' Literal data, alphabetic
%PicMap_Escape             = &H00008000                           ' To mark backslash escaped pic codes

'----- Picture mapping globals
TYPE PIC_type

   data_mask(0 TO 255)     AS WORD
   pic_mask(0 TO 255)      AS WORD
   mapping_initialized     AS LONG
   defaults_initialized    AS LONG
   WORD_WORD_changed       AS LONG                                ' set to 1 when WORD changed
   CharSet_changed         AS LONG                                ' set to 1 when CharSet changed

END TYPE

TYPE MY_EXCEPTION_POINTERS
   ExceptionRecord      AS EXCEPTION_RECORD POINTER
   ContextRecord        AS CONTEXT          POINTER
END TYPE

'/-----------------------------------------------------------------------------/
'/  SPFLiteCodePageTypes.inc                                                   /
'/                                                                             /
'/  TYPES AND CONSTANTS FOR NEW CODEPAGE SUPPORT                               /
'/-----------------------------------------------------------------------------/
%TX_ASCII  = 1
%TX_EBCDIC = 2
%TA_INDEX  = %TX_ASCII
%TE_INDEX  = %TX_EBCDIC
%TX_MAX    = 2
%AE_MODE   = 1
%EA_MODE   = 2

'/-----------------------------------------------------------------------------/
'/  TYPE CODEPAGE_TT_T                                                         /
'/                                                                             /
'/  COMMON TRAITS FOR A CODEPAGE                                               /
'/  THIS INCLUDES CODEPAGE SUMMARY INFORMATION AND VALIDATION STATUS           /
'/-----------------------------------------------------------------------------/
TYPE CODEPAGE_TT_T         '/ CODEPAGE TRAITS FOR 'TT' LINES
    TT_ERRORS               AS LONG                '/ NUMBER OF ERRORS FOUND
    TT_REASON               AS ASCIIZ * (80+1)     '/ DESCRIPTION OF ERROR
    TT_AUTHOR               AS ASCIIZ * (64+1)     '/ CREATOR OF TABLE
    TT_GENDATE              AS ASCIIZ * (20+1)     '/ "2002-12-03 00:00:00"
    TT_MODE                 AS ASCIIZ * (2+1)      '/ RT/ES ROUND TRIP/SUBSET
    TT_NAME                 AS ASCIIZ * (64+1)     '/ NAME  OF TRANSLATION TABLE
    TT_TITLE                AS ASCIIZ * (64+1)     '/ TITLE OF TRANSLATION TABLE
    TT_OTHER                AS ASCIIZ * (80+1)     '/ UNKNOWN KEYWORD VALUE
END TYPE

'/-----------------------------------------------------------------------------/
'/  TYPE CODEPAGE_TX_T                                                         /
'/                                                                             /
'/  STORES ATTRIBUTE FIELDS FOR A GIVEN CODEPAGE, WHICH CHARACTERIZES HOW THE  /
'/  CODEPAGE GOT GENERATED.  VALIDATION STATUS IS ALSO INCLUDED.               /
'/-----------------------------------------------------------------------------/
TYPE CODEPAGE_TX_T         '/ CODEPAGE TRAITS FOR 'TA' AND 'TE' LINES
    TX_DEFINED              AS BYTE                '/ IF AE/EA TABLE DEFINED
    TX_ERRORS               AS LONG                '/ NUMBER OF ERRORS FOUND
    TX_VALUES               AS LONG                '/ NUMBER OF VALUES STORED
    TX_REASON               AS ASCIIZ * (80+1)     '/ DESCRIPTION OF ERROR
    TX_ENTRY (0 TO 15)      AS BYTE                '/ FLAGS FOR AE/EA LINES
    TX_CCSID                AS ASCIIZ * (9+1)      '/ PIECES OF 'NUMBER' AS INT
    TX_CGCSGID              AS ASCIIZ * (10+1)     '/ CGCSGID   "00695"
    TX_CODESET              AS ASCIIZ * (64+1)     '/ FULL CODESET NAME
    TX_CPGID                AS ASCIIZ * (10+1)     '/ CPGID  "01140"
    TX_EURO                 AS ASCIIZ * (2+1)      '/ VALUE OF EURO OR 00
    TX_NUMBER               AS ASCIIZ * (20+1)     '/ "1140", "8859_1", ETC.
    TX_ORIGIN               AS ASCIIZ * (16+1)     '/ "IBM", "ISO" ETC.
    TX_RELATED              AS ASCIIZ * (10+1)     '/ RELATED EURO CCSID "-37"
    TX_SCHEME               AS ASCIIZ * (10+1)     '/ ENCODING SCHEME "1100"
    TX_SIZE                 AS ASCIIZ * (3+1)      '/ NUM OF DEFINED CHARS
    TX_SUB                  AS ASCIIZ * (2+1)      '/ SUBSTITUTION CHAR
    TX_TYPE                 AS ASCIIZ * (6+1)      '/ "ASCII", "EBCDIC"
    TX_UCM                  AS ASCIIZ * (64+1)     '/ .UCM FILE NAME
    TX_UCMDATE              AS ASCIIZ * (20+1)     '/ "2002-12-03 00:00:00"
    TX_VERSION              AS ASCIIZ * (16+1)     '/ "2.3.3", "1995", ETC.
    TX_OTHER                AS ASCIIZ * (80+1)     '/ UNKNOWN KEYWORD VALUE
    TX_TABLE (0 TO 255)     AS BYTE                '/ FINAL TRANSLATION TABLE
END TYPE

'/-----------------------------------------------------------------------------/
'/  TYPE CODEPAGE_CP_T                                                         /
'/                                                                             /
'/  CONTAINING STRUCTURE FOR THE CODEPAGE IDENTIFICATION                       /
'/  AND THE TWO TRAN TABLES (AE AND EA TABLES) REFERRED TO AS 'TX' INFORMATION /
'/-----------------------------------------------------------------------------/
TYPE CODEPAGE_CP_T         '/ CODEPAGE DATA STRUCTURE
    CP_LINENO               AS LONG                '/ LINE NUMBER FOR ERRORS
    CP_ERRORS               AS LONG                '/ NUMBER OF ERRORS FOUND
    CP_REASON               AS ASCIIZ * (80+1)     '/ DESCRIPTION OF ERROR
    TT                      AS CODEPAGE_TT_T
    TX (1 TO %TX_MAX)       AS CODEPAGE_TX_T
END TYPE

'/-----------------------------------------------------------------------------/
'/  PARSE_KEYWORD_VALUE_DATA                                                   /
'/  USED BY PARSE_KEYWORD_VALUE_DATA SUB                                       /
'/-----------------------------------------------------------------------------/
TYPE                        KEYWORD_VALUE_DATA_T
    KV_DATA                 AS ASCIIZ * (256+1)
    KV_KEYWORD              AS ASCIIZ * (32+1)
    KV_VALUE                AS ASCIIZ * (128+1)
    KV_REASON               AS ASCIIZ * (80+1)
    KV_STATUS               AS LONG
END TYPE '                  KEYWORD_VALUE_DATA_T

'/-----------------------------------------------------------------------------
'/
'/ Writeup from ISPF manual about Data Shifting (with modifications at "-->"):
'/
'/
'/ Data shifting attempts to shift the body of a program statement without
'/ shifting the label or comments, and prevents loss of data.
'/
'/ This shift is non-destructive because it stops before it shifts a nonblank
'/ character past the bound.  This shift is explicitly done with the < and >
'/ line commands, and the SHIFT < and SHIFT > macro commands.
'/
'/ In ISPF, the CHANGE command can cause an implicit shift of the same
'/ nature.  In SPFLite, data shifting occurs when CHANGE DS is effect, either
'/ by default, or by a DS keyword.  Column shifting occurs when CHANGE CS is
'/ in effect, either by default, or by a CS keyword.
'/
'/ For data shift left attempts that exceed the current BOUNDS setting, text
'/ stops at the left bound and PDF marks the shifted lines with ==ERR> flags.
'/
'/ If an error occurs in an excluded line, you can find the error with LOCATE
'/ ERROR|ERR, and remove the error flag by using RESET ERROR|ERR.
'/
'/ --> SPFLite does not currently support these LOCATE or RESET options.
'/
'/ Data shifts are designed to work with typical program sources.
'/
'/ In doing so, it makes certain general assumptions about the format of the
'/ source code.
'/
'/
'/ For instance, the editor ASSUMES:
'/
'/ *  Anything beginning at the left bound is a label and should not be shifted
'/
'/ *  If there are TWO OR MORE consecutive blanks, one can be added or deleted
'/
'/ --> This means by default, spans of 2 or more blanks can be collapsed or
'/ --> expanded.  If the minimum span of blanks has been changed from the
'/ --> default of 1, then 'adjustable' spans are of length (minimum span) + 1.
'/
'/ *  Blanks within quotes (’ or ") are to be treated as non_blanks
'/
'/ --> The code here treats an escaped quote \' or \" as not ending the quote.
'/ --> This is necessary for C-language strings to be handled properly.
'/ --> This usage is enabled only if OPT.DS.ESCAPE is set to 1.
'/
'/ *  Source statements appear on the left followed by comments on the right
'/
'/ *  Single blanks are used between source code and comment words
'/
'/ --> This means by default, a single blank CANNOT be deleted or expanded.
'/ --> If the minimum span of blanks has been changed from the default of 1,
'/ --> then 'non-adjustable' spans are of length (minimum span) or less, with
'/ --> the "or less" being as short as one blank.
'/
'/ --> These are ASSUMPTIONS that ISPF makes as a "model" of what they expect
'/ --> a typical source line to look like.  The assumptions are not the RULES
'/ --> of how the shifts work, only the rationale of how the rules came to be.
'/
'/
'/ Therefore [ EFFECTIVELY ] the only strings of multiple blanks appear between
'/ the source code and the comment, and between the comment and its ending
'/ delimiter (if there is one).
'/
'/ --> The word EFFECTIVELY was added so the IBM explanation woule make sense.
'/
'/ In this example, LABEL and */ are at the left and right bounds,
'/ respectively:
'/
'/     LABEL: DO I=1 TO 5;       /* The comment... */
'/          A=A+B(I);            /* The comment... */
'/     END;
'/
'/ Keeping the previous assumptions in mind, the editor attempts to move only
'/ the source code statement when shifting data. The label and comments are
'/ left unchanged.  However, if necessary, it shifts the comment also.
'/
'/ --> What IBM does not mention is that, to the data shift logic, a "label"
'/ --> means a continguous span of nonblanks.  In the example above, the ISPF
'/ --> label-field is not "LABEL" but "LABEL:".  The shift logic does NOT look
'/ --> at programming syntax - only spans of blanks vs. nonblanks.  So, in the
'/ --> example, it's not just the letters in "LABEL" but the colon counts too.
'/
'/ Although the editor always uses these assumptions, data shifting is not
'/ language-sensitive.  It only makes generalities about syntax and individual
'/ code entry style.
'/
'/-----------------------------------------------------------------------------
'/
'/ Additional comments:
'/
'/ 1. The < left shift and > right shift are not symmetrical.  In a left
'/    shift, the first span of two or more spaces is a "detach point".  That
'/    is, after that span, no further data is pulled left because of the
'/    shift.  This means the first of two blanks is the 'stopping point' or
'/    'barrier' for the left shift, and it can be immediately found, even
'/    without knowing the number of columns being shifted.
'/
'/ 2. In a right shift, as many spans of blanks are collapsed as needed to
'/    make the shift possible, so there is no obvious 'stopping point' beyond
'/    which the shift does not happen.  Instead, this must be carefully
'/    determined, and the number of columns shifted is critical to this.
'/
'/ 3. As an extension, the minimum number of spaces that can be collapsed to
'/    is provided as a parameter.  If this were an ISPF feature, the ISPF
'/    default for the minimum gap size would be 1.  Allowance is made for
'/    gaps of 1 to 9.  For demo purposes, this value is fetched by getting the
'/    value of SET symbol OPT.DS.MINSIZE.     A better way may be to have some
'/    means like a primary command or global option to set this, but using the
'/    SET symbol allows us to experiment while deciding how to handle this idea
'/    going forward.  A new Global Options field would have the least impact.
'/
'/ 4. It was found by experimentation that what SPFLite calls CHANGE DS mode
'/    and what ISPF does all the time with CHANGE when the result string is a
'/    different size, is comparable to the > and < shifts, if the column that
'/    corresponded to the left-hand character of the found string where set
'/    (temporarily) as the left-bound.  In particular, ISPF respects quoted
'/    strings, and will not collapse quoted spans of quoted blanks if a changed
'/    string "pushes" against it.  In the future, consideration should be
'/    given to incorporating these routines, or logic like them, in SPFLite's
'/    implementation of CHANGE.  However, performance issues and the amount of
'/    effort needed to integrate these functions into existing code might make
'/    that impractical.
'/
'/ 5. IBM's quoted-string detection logic is fragile and non-robust.  They will
'/    recognize string in single quotes but not double quotes, even thought the
'/    ISPF manual says they allow both.  Also, if a data line contains a string
'/    whose beginning and end straddles the left bound, the behavior of shifts
'/    is surprising and unexpected. The code here attempts to handle these
'/    situations in a consistent and rationale manner.
'/
'/ 6. In order to reconstitute the color line, when the data line has colors,
'/    colors associated with nonblanks are extended using the color of the
'/    last blank in each span of blanks.
'/
'/-----------------------------------------------------------------------------

TYPE DataShift_Segment_t
   gap_len                 AS LONG
   data_pos                AS LONG
   data_len                AS LONG
   blank_pos               AS LONG
   blank_len               AS LONG
END TYPE

' Mapping support stuff

ENUM var_enum SINGULAR
   V_null
   VA
   VB
   VC
   VD
   VE
   VF
   VG
   VH
   VI
   VJ
   VK
   VL
   VM
   VN
   VO
   VP
   VQ
   VR
   VS
   VT
   VU
   VV
   VW
   VX
   VY
   VZ
   V_dummy                                   '/ dummy slot for null set clauses
   V_num                                      '/ flag for non-var numeric value

   V_min = %VA
   V_max = %V_num
END ENUM


'/ token classes generically describe a token in terms of its functionality
'/ cls types:              {  ;  ,  (  )  [  ]  n  v  U  B  A  S  }

ENUM cls_enum SINGULAR

   cls_null          '/  00        for initializing undefined cls values

   cls_begin         '/  01  {     start of expr
   cls_semi          '/  02  ;     statement delimiter
   cls_LP            '/  03  (     parenthesized subexpression
   cls_RP            '/  04  )     parenthesized subexpression
   cls_LB            '/  05  [     set expression
   cls_RB            '/  06  ]     set expression
   cls_comma         '/  07  ,     set clause delimiter
   cls_n             '/  08  n     number - decimal or hex
   cls_v             '/  09  v     variable name
   cls_U             '/  10  U     unary operator
   cls_B             '/  11  B     binary operator
   cls_A             '/  12  A     assignment operator
   cls_S             '/  13  S     swap operator
   cls_end           '/  14  }     end of expression

   cls_MIN = %cls_begin
   cls_MAX = %cls_end

END ENUM


'/ expression operator tokens

ENUM opr_enum SINGULAR

   opr_null          '/          for initializing undefined opr values

   opr_A_B_SCALE     '/ 01
   opr_A_B_SL        '/ 02
   opr_A_B_SR        '/ 03
   opr_A_L_AND       '/ 04
   opr_A_L_OR        '/ 05
   opr_A_L_XOR       '/ 06
   opr_A_N_SR        '/ 07

   opr_NGTE          '/ 08
   opr_NLTE          '/ 09

   opr_A_ADD         '/ 10
   opr_A_B_AND       '/ 11
   opr_A_B_OR        '/ 12
   opr_A_B_XOR       '/ 13
   opr_A_C_MOD       '/ 14
   opr_A_DIV         '/ 15
   opr_A_MOD_S       '/ 16
   opr_A_MUL         '/ 17
   opr_A_N_SCALE     '/ 18
   opr_A_SUB         '/ 19

   opr_B_SCALE       '/ 20
   opr_B_SL          '/ 21
   opr_B_SR          '/ 22
   opr_EQ            '/ 23
   opr_GE            '/ 24
   opr_LE            '/ 25
   opr_L_AND         '/ 26
   opr_L_OR          '/ 27
   opr_L_XOR         '/ 28
   opr_MAKE_GE       '/ 29
   opr_MAKE_LE       '/ 30
   opr_NE_C          '/ 31
   opr_NE_S          '/ 32
   opr_NGT           '/ 33
   opr_NLT           '/ 34
   opr_N_SR          '/ 35
   opr_SWAP          '/ 36
   opr_U_SGN_X       '/ 37
   opr_ASSIGN        '/ 38

   opr_ADD           '/ 39
   opr_ASSIGN_EQ     '/ 40
   opr_B_AND         '/ 41
   opr_B_OR          '/ 42
   opr_B_XOR         '/ 43
   opr_COMMA         '/ 44
   opr_C_MOD         '/ 45
   opr_DIV           '/ 46
   opr_GT            '/ 47
   opr_LT            '/ 48
   opr_L_SET         '/ 49
   opr_MOD_S         '/ 50
   opr_MUL           '/ 51
   opr_N_SCALE       '/ 52
   opr_R_SET         '/ 53
   opr_SUB           '/ 54
   opr_U_ABS         '/ 55
   opr_U_B_NOT       '/ 56
   opr_U_L_NOT       '/ 57
   opr_U_NEG         '/ 58
   opr_U_POS         '/ 59
   opr_U_SGN         '/ 60

END ENUM

'---------- For WinVer usage
%VER_NT_WORKSTATION   = &H0000001
%VER_PRODUCT_TYPE     = &H0000080
%WIN32_WINNT_NT4      = &H0400???                                 ' Windows NT4
%WIN32_WINNT_WIN2K    = &H0500???                                 ' Windows 2000
%WIN32_WINNT_WINXP    = &H0501???                                 ' Windows Server 2003, Windows XP
%WIN32_WINNT_WS03     = &H0502???                                 ' Windows Server 2003 with SP1, Windows XP with SP2
%WIN32_WINNT_WIN6     = &H0600???                                 ' Windows Vista
%WIN32_WINNT_VISTA    = &H0600???                                 ' Windows Vista
%WIN32_WINNT_WS08     = &H0600???                                 ' Windows Server 2008
%WIN32_WINNT_LONGHORN = &H0600???                                 ' Windows
%WIN32_WINNT_WIN7     = &H0601???                                 ' Windows Seven
%WIN32_WINNT_WIN8     = &H0602???                                 ' Windows Eight
%WIN32_WINNT_WINBLUE  = &H0603???                                 ' Windows Eight.1
%WIN32_WINNT_WIN10    = &H0A00???                                 ' Windows Ten

TYPE utab_t                                                       ' Private Unicode table
   valid            AS LONG
   uchar (0 TO 255) AS WSTRING * 1
END TYPE

'/ User Data Area indexes for (CharSet) and (CharSetCol)

%uda_colMode          = 1                                         ' index to tls_colMode
%uda_PrfGetX2APtr     = 2                                         ' index to tls_rfGetX2APtr
%uda_ansiSource       = 3                                         ' index to tls_ansiSource
