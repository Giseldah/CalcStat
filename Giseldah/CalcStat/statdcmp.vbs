Const adSaveCreateNotExist = 1
Const adSaveCreateOverWrite = 2
Const adTypeBinary = 1 'Binary data
Const adTypeText = 2 '(Default) Text data
Const adLF = 10
Const adCR = 13
Const adCRLF = -1

Const adReadLine = -2
Const adWriteChar = 0
Const adWriteLine = 1

'indexes StatData record fields
Const SD_STAT = 0
Const SD_SCRIPT = 1
Const SD_LSTART = 2
Const SD_LEND = 3
Const SD_TYPE = 4
Const SD_USEN = 5
Const SD_DECIMALS = 6
Const SD_P1 = 7
Const SD_P2 = 8
Const SD_P3 = 9
Const SD_P4 = 10
Const SD_P5 = 11
Const SD_P6 = 12
Const SD_COMPF = 13
Const SDFLDCNT = 14

Const SDFLDSEP = ","
Const SDSTRQUOT = """"

Const InFileFull = "statdata.csv"
Const OutFileName = "CalcStat"
Const OutFileExtBas = ".bas"
Const OutFileExtLua = ".lua"
Const OutFileExtVbs = ".vbs"

Const optOutFileBas = 1	'output file choice StarBasic
Const optOutFileLua = 2	'output file choice Lua-script
Const optOutFileVbs = 3	'output file choice VB-script

Dim OutFileOption 'output file choice

'indexes Binary Search Tree node variables
Const BST_KEY = 0
Const BST_VALUE = 1
Const BST_LEFT = 2
Const BST_RIGHT = 3
Const BSTFLDCNT = 4

'scripting types
Const SCRIPT_DEFINE = "D" 'define $name for replacements
Const SCRIPT_SEQUENTIAL = "S" 'in sequence L <= Lend
Const SCRIPT_INTERVAL = "I" 'by interval Lstart <= L <= Lend

Const DF_DEF = 0 'defined text to replace
Const DF_REPL = 1 'replacement text

Const optCompileFull = 0 'full version with all stats
Const optCompilePercentages = 1 'percentage calculations only

Dim CompileOption 'compile option choice

Main
Sub Main

	OutFileOption = CInt(InputBox("Enter a Number:"&Chr(10)&"1 = StarBasic"&Chr(10)&"2 = Lua-script <- for plugins"&Chr(10)&"3 = VB-script","Choose output script type",2))

	If Not (OutFileOption = optOutFileBas Or OutFileOption = optOutFileLua Or OutFileOption = optOutFileVbs) Then
		MsgBox "Invalid option. Operation cancelled."
		Exit Sub
	End If
	
	CompileOption = CInt(InputBox("Enter a Number:"&Chr(10)&"1 = Full version with all stats"&Chr(10)&"2 = Percentage calculations only <- normal for plugins","Choose output version",2))-1

	If Not (CompileOption = optCompileFull Or CompileOption = optCompilePercentages) Then
		MsgBox "Invalid option. Operation cancelled."
		Exit Sub
	End If
	
	Dim SDInfoLst

	ReadSDFile InFileFull,SDInfoLst
	
	If Not IsEmpty(SDInfoLst) Then
		MsgBox Replace("Found #N# data records.","#N#",CStr(UBound(SDInfoLst)+1))
	End If

	Dim aDefines
	aDefines = ReadDefines(SDInfoLst)
		
	Dim aStatsIndex
	aStatsIndex = CreateStatsIndex(SDInfoLst)

	Dim aBalancedBST
	aBalancedBST = CreateBalancedBST(aStatsIndex)
	
	Dim OutStream
	Set OutStream = CreateObject("ADODB.Stream")
	With OutStream
		.Type = adTypeText
		.CharSet = "us-ascii"
		.LineSeparator = adCRLF
		.Open

		WriteScrStart OutStream

		Dim iNodeIndex
		iNodeIndex = LBound(aBalancedBST)	'Head node of the tree is first in the list

		WriteScrStatsNode OutStream, SDInfoLst, aDefines, aBalancedBST, iNodeIndex, "	"

		WriteScrEnd OutStream

		Dim OutFileFull
If OutFileOption = optOutFileBas Then
		OutFileFull = OutFileName+OutFileExtBas
ElseIf OutFileOption = optOutFileLua Then
		OutFileFull = OutFileName+OutFileExtLua
ElseIf OutFileOption = optOutFileVbs Then
		OutFileFull = OutFileName+OutFileExtVbs
End If
		.SaveToFile OutFileFull, adSaveCreateOverWrite
		.Close
	End With
	Set OutStream = Nothing

	MsgBox Replace(Replace("Compilation of #S1# to #S2# completed.","#S1#",InFileFull),"#S2#",OutFileFull)
	
End Sub

Function ReadDefines(theSDInfoLst)

	Dim aDefines

	Dim sStat, sLastStat
	sLastStat = ""

	For SI = LBound(theSDInfoLst) To UBound(theSDInfoLst)
		If theSDInfoLst(SI)(SD_SCRIPT) = SCRIPT_DEFINE Then
			sStat = UCase(theSDInfoLst(SI)(SD_STAT))
			If sStat <> sLastStat Then
				If sLastStat <> "" Then
					Redim Preserve aDefines(UBound(aDefines)+1)
					aDefines(UBound(aDefines)) = Array(sStat,theSDInfoLst(SI)(SD_P1))
				Else
					aDefines = Array(Array(sStat,theSDInfoLst(SI)(SD_P1)))
				End If
				sLastStat = sStat
			End If
		End If
	Next

	If VarType(theDefines) = 8204 Then
		Dim temp
	
		For I = UBound(aDefines)-1 To LBound(aDefines) Step -1
			For J = 0 to I
				If Len(aDefines(J)(0)) < Len(aDefines(J+1)(0)) Then
					temp = aDefines(J+1)
					aDefines(J+1) = aDefines(J)
					aDefines(J) = temp
				End If
			Next
		Next 
	End If
	
	''MsgBox Replace("Number of defines: #N#","#N#",CStr(UBound(aDefines)+1))

	ReadDefines = aDefines

End Function

Function ReplaceDefines(theText, theDefines)

	Dim sNewText
	sNewText = theText

	If VarType(theDefines) = 8204 Then
	
		Dim sSearchText
		sSearchText = UCase(theText)

		Dim SDefine
		Dim sRepl
	
		For DI = LBound(theDefines) To UBound(theDefines)
			sDefine = theDefines(DI)(DF_DEF)
			sRepl = theDefines(DI)(DF_REPL)
			nSearch = InStr(sSearchText,sDefine)
			While nSearch > 0
				sNewText = Left(sNewText,nSearch-1)+sRepl+Right(sNewText,Len(sSearchText)-nSearch+1-Len(sDefine))
				sSearchText = Left(sSearchText,nSearch-1)+sRepl+Right(sSearchText,Len(sSearchText)-nSearch+1-Len(sDefine))
				nSearch = InStr(sSearchText,sDefine)
			Wend
		Next
	End If
	
	ReplaceDefines = sNewText

End Function

Sub WriteScrStart(theOutStream)

	With theOutStream

If OutFileOption = optOutFileBas Then
.WriteText "Function CalcStat(ByVal SName As String, ByVal SLvl As Double, Optional SParam As Variant) As Variant", adWriteLine
.WriteText "", adWriteLine
.WriteText "	Dim SN As String", adWriteLine
.WriteText "	Dim L As Double", adWriteLine
.WriteText "	Dim N As Double", adWriteLine
.WriteText "	Dim C As String", adWriteLine
.WriteText "", adWriteLine
.WriteText "	SN = UCase(LTrim(RTrim(SName)))", adWriteLine
.WriteText "", adWriteLine
.WriteText "	L = SLvl", adWriteLine
.WriteText "", adWriteLine
.WriteText "	If IsMissing(SParam) Then", adWriteLine
.WriteText "		N = 1#", adWriteLine
.WriteText "		C = """"", adWriteLine
.WriteText "	ElseIf IsNumeric(SParam) Then", adWriteLine
.WriteText "		N = SParam", adWriteLine
.WriteText "		C = """"", adWriteLine
.WriteText "	Else", adWriteLine
.WriteText "		N = 1#", adWriteLine
.WriteText "		C = SParam", adWriteLine
.WriteText "	EndIf", adWriteLine
ElseIf OutFileOption = optOutFileLua Then
.WriteText "_G.CalcStat = function(sname, slvl, sparam)", adWriteLine
.WriteText "", adWriteLine
.WriteText "	local sn = string.upper(string.match(sname,""(%w+)""));", adWriteLine
.WriteText "", adWriteLine
.WriteText "	local L = slvl;", adWriteLine
.WriteText "	local N = 1;", adWriteLine
.WriteText "	local C = """";", adWriteLine
.WriteText "", adWriteLine
.WriteText "	if sparam ~= nil then", adWriteLine
.WriteText "		if type(sparam) == ""number"" then", adWriteLine
.WriteText "			N = sparam;", adWriteLine
.WriteText "		else", adWriteLine
.WriteText "			C = sparam;", adWriteLine
.WriteText "		end", adWriteLine
.WriteText "	end", adWriteLine
ElseIf OutFileOption = optOutFileVbs Then
.WriteText "Function CalcStat(ByVal SName, ByVal SLvl, ByVal SParam)", adWriteLine
.WriteText "", adWriteLine
.WriteText "	Dim SN", adWriteLine
.WriteText "	Dim L", adWriteLine
.WriteText "	Dim N", adWriteLine
.WriteText "	Dim C", adWriteLine
.WriteText "", adWriteLine
.WriteText "	SN = UCase(LTrim(RTrim(SName)))", adWriteLine
.WriteText "", adWriteLine
.WriteText "	L = SLvl", adWriteLine
.WriteText "", adWriteLine
.WriteText "	If VarType(SParam) = 9 Then", adWriteLine
.WriteText "		N = 1", adWriteLine
.WriteText "		C = """"", adWriteLine
.WriteText "	ElseIf VarType(SParam) <> 8 Then", adWriteLine
.WriteText "		N = SParam", adWriteLine
.WriteText "		C = """"", adWriteLine
.WriteText "	Else", adWriteLine
.WriteText "		N = 1", adWriteLine
.WriteText "		C = SParam", adWriteLine
.WriteText "	End If", adWriteLine
End If
.WriteText "", adWriteLine

	End With

End Sub

Sub WriteScrStatsNode(theOutStream,theSDInfoLst,theDefines,aBST,iNodeIndex,sIndent)
	
	With theOutStream
		Dim sStatName
		sStatName = aBST(iNodeIndex)(BST_KEY)
		
		Dim iStatIndex
		iStatIndex = aBST(iNodeIndex)(BST_VALUE)
		
		Dim bIfOpen
		bIfOpen = False
		
		Dim iLeftIndex
		iLeftIndex = aBST(iNodeIndex)(BST_LEFT)
		
		If iLeftIndex <> -1 Then
			If bIfOpen Then
If OutFileOption = optOutFileBas Then
.WriteText sIndent+"ElseIf SN < """+sStatName+""" Then", adWriteLine
ElseIf OutFileOption = optOutFileLua Then
.WriteText sIndent+"elseif sn < """+sStatName+""" then", adWriteLine
ElseIf OutFileOption = optOutFileVbs Then
.WriteText sIndent+"ElseIf SN < """+sStatName+""" Then", adWriteLine
End If
			Else
If OutFileOption = optOutFileBas Then
.WriteText sIndent+"If SN < """+sStatName+""" Then", adWriteLine
ElseIf OutFileOption = optOutFileLua Then
.WriteText sIndent+"if sn < """+sStatName+""" then", adWriteLine
ElseIf OutFileOption = optOutFileVbs Then
.WriteText sIndent+"If SN < """+sStatName+""" Then", adWriteLine
End If
				bIfOpen = True
			End If
			WriteScrStatsNode theOutStream, theSDInfoLst, theDefines, aBST, iLeftIndex, sIndent+"	"
		End If
		
		Dim iRightIndex
		iRightIndex = aBST(iNodeIndex)(BST_RIGHT)
		
		If iRightIndex <> -1 Then
			If bIfOpen Then
If OutFileOption = optOutFileBas Then
.WriteText sIndent+"ElseIf SN > """+sStatName+""" Then", adWriteLine
ElseIf OutFileOption = optOutFileLua Then
.WriteText sIndent+"elseif sn > """+sStatName+""" then", adWriteLine
ElseIf OutFileOption = optOutFileVbs Then
.WriteText sIndent+"ElseIf SN > """+sStatName+""" Then", adWriteLine
End If
			Else
If OutFileOption = optOutFileBas Then
.WriteText sIndent+"If SN > """+sStatName+""" Then", adWriteLine
ElseIf OutFileOption = optOutFileLua Then
.WriteText sIndent+"if sn > """+sStatName+""" then", adWriteLine
ElseIf OutFileOption = optOutFileVbs Then
.WriteText sIndent+"If SN > """+sStatName+""" Then", adWriteLine
End If
				bIfOpen = True
			End If
			WriteScrStatsNode theOutStream, theSDInfoLst, theDefines, aBST, iRightIndex, sIndent+"	"
		End If

		If iLeftIndex = -1 Or iRightIndex = -1 Then
			If bIfOpen Then
If OutFileOption = optOutFileBas Then
.WriteText sIndent+"ElseIf SN = """+sStatName+""" Then", adWriteLine
ElseIf OutFileOption = optOutFileLua Then
.WriteText sIndent+"elseif sn == """+sStatName+""" then", adWriteLine
ElseIf OutFileOption = optOutFileVbs Then
.WriteText sIndent+"ElseIf SN = """+sStatName+""" Then", adWriteLine
End If
			Else
If OutFileOption = optOutFileBas Then
.WriteText sIndent+"If SN = """+sStatName+""" Then", adWriteLine
ElseIf OutFileOption = optOutFileLua Then
.WriteText sIndent+"if sn == """+sStatName+""" then", adWriteLine
ElseIf OutFileOption = optOutFileVbs Then
.WriteText sIndent+"If SN = """+sStatName+""" Then", adWriteLine
End If
				bIfOpen = True
			End If
			WriteScrStat theOutStream,sStatName,theSDInfoLst,theDefines,sIndent+"	"
If OutFileOption = optOutFileBas Then
.WriteText sIndent+"Else", adWriteLine
.WriteText sIndent+"	CalcStat = 0", adWriteLine
ElseIf OutFileOption = optOutFileLua Then
.WriteText sIndent+"else", adWriteLine
.WriteText sIndent+"	return 0;", adWriteLine
ElseIf OutFileOption = optOutFileVbs Then
.WriteText sIndent+"Else", adWriteLine
.WriteText sIndent+"	CalcStat = 0", adWriteLine
End If
		ElseIf iLeftIndex <> -1 And iRightIndex <> -1 Then
If OutFileOption = optOutFileBas Then
.WriteText sIndent+"Else", adWriteLine
ElseIf OutFileOption = optOutFileLua Then
.WriteText sIndent+"else", adWriteLine
ElseIf OutFileOption = optOutFileVbs Then
.WriteText sIndent+"Else", adWriteLine
End If
			WriteScrStat theOutStream,sStatName,theSDInfoLst,theDefines,sIndent+"	"
		End If

		If bIfOpen Then
If OutFileOption = optOutFileBas Then
.WriteText sIndent+"End If", adWriteLine
ElseIf OutFileOption = optOutFileLua Then
.WriteText sIndent+"end", adWriteLine
ElseIf OutFileOption = optOutFileVbs Then
.WriteText sIndent+"End If", adWriteLine
End If
		End If
	End With

End Sub

Sub WriteScrEnd(theOutStream)

	With theOutStream

.WriteText "", adWriteLine
If OutFileOption = optOutFileBas Then
.WriteText "End Function", adWriteLine
ElseIf OutFileOption = optOutFileLua Then
.WriteText "end", adWriteLine
ElseIf OutFileOption = optOutFileVbs Then
.WriteText "End Function", adWriteLine
End If

	End With

End Sub

Function CreateStatsIndex(theSDInfoLst)

	Dim aStatsIndex

	Dim sStat, sLastStat
	sLastStat = ""

	For SI = LBound(theSDInfoLst) To UBound(theSDInfoLst)
		If theSDInfoLst(SI)(SD_SCRIPT) <> SCRIPT_DEFINE Then
			sStat = theSDInfoLst(SI)(SD_STAT)
			If sStat <> sLastStat Then
				If sLastStat <> "" Then
					Redim Preserve aStatsIndex(UBound(aStatsIndex)+1)
					aStatsIndex(UBound(aStatsIndex)) = Array(sStat,SI)
				Else
					aStatsIndex = Array(Array(sStat,SI))
				End If
				sLastStat = sStat
			End If
		End If
	Next

	MsgBox Replace("Number of stats: #N#","#N#",CStr(UBound(aStatsIndex)+1))

	CreateStatsIndex = aStatsIndex

End Function

Function CreateBalancedBST(aStatsIndex)

	Dim iLowIndex
	iLowIndex = LBound(aStatsIndex)
	
	Dim iHighIndex
	iHighIndex = UBound(aStatsIndex)

	Dim iStatIndex
	iStatIndex = Int((iLowIndex+(iHighIndex-iLowIndex)/2))

	Dim aBalancedBST
	aBalancedBST = Array(Array(-1,-1,-1,-1))
	
	Dim iNodeIndex
	iNodeIndex = UBound(aBalancedBST)

	aBalancedBST(iNodeIndex)(BST_KEY) = aStatsIndex(iStatIndex)(0)
	aBalancedBST(iNodeIndex)(BST_VALUE) = aStatsIndex(iStatIndex)(1)
	If iStatIndex-1 >= iLowIndex Then
		aBalancedBST(iNodeIndex)(BST_LEFT) = AddBSTnode(aBalancedBST,aStatsIndex,iLowIndex,iStatIndex-1)
	End If
	If iStatIndex+1 <= iHighIndex Then
		aBalancedBST(iNodeIndex)(BST_RIGHT) = AddBSTnode(aBalancedBST,aStatsIndex,iStatIndex+1,iHighIndex)
	End If

	CreateBalancedBST = aBalancedBST
	
	MsgBox Replace("Height of Binary Search Tree: #N#","#N#",GetHeightBSTnode(aBalancedBST,0))
		
End Function

Function AddBSTnode(ByRef aBST, ByRef aStatsIndex, ByVal iLowIndex, ByVal iHighIndex)

	Dim iStatIndex
	iStatIndex = Int((iLowIndex+(iHighIndex-iLowIndex)/2))
	
	Redim Preserve aBST(UBound(aBST)+1)
	aBST(UBound(aBST)) = Array(-1,-1,-1,-1)

	Dim iNodeIndex
	iNodeIndex = UBound(aBST)

	aBST(iNodeIndex)(BST_KEY) = aStatsIndex(iStatIndex)(0)
	aBST(iNodeIndex)(BST_VALUE) = aStatsIndex(iStatIndex)(1)
	If iStatIndex-1 >= iLowIndex Then
		aBST(iNodeIndex)(BST_LEFT) = AddBSTnode(aBST,aStatsIndex,iLowIndex,iStatIndex-1)
	End If
	If iStatIndex+1 <= iHighIndex Then
		aBST(iNodeIndex)(BST_RIGHT) = AddBSTnode(aBST,aStatsIndex,iStatIndex+1,iHighIndex)
	End If
	
	AddBSTnode = iNodeIndex
	
End Function

Function GetHeightBSTnode(ByRef aBST, ByVal iStatIndex)

	Dim iLeftIndex
	iLeftIndex = aBST(iStatIndex)(BST_LEFT)
	Dim iLeftHeight
	iLeftHeight = 0
	If iLeftIndex <> -1 Then
		iLeftHeight = GetHeightBSTnode(aBST,iLeftIndex)
	End If

	Dim iRightIndex
	iRightIndex = aBST(iStatIndex)(BST_RIGHT)
	Dim iRightHeight
	iRightHeight = 0
	If iRightIndex <> -1 Then
		iRightHeight = GetHeightBSTnode(aBST,iRightIndex)
	End If

	If iLeftHeight >= iRightHeight Then
		GetHeightBSTnode = iLeftHeight+1
	Else
		GetHeightBSTnode = iRightHeight+1
	End If
	
End Function

Sub WriteScrStat(theOutStream,theStat,theSDInfoLst,theDefines,sIndent)

	With theOutStream

		Dim StatPos
		StatPos = -1
		
		For SI = LBound(theSDInfoLst) To UBound(theSDInfoLst)
			If theSDInfoLst(SI)(SD_STAT) = theStat Then
				StatPos = SI
				Exit For
			End If
		Next

		If StatPos >= 0 Then
			Dim Build

			For SI = StatPos To UBound(theSDInfoLst)
				If theSDInfoLst(SI)(SD_STAT) <> theStat Then
					Exit For
				End If
				
				For FI = LBound(theSDInfoLst(SI)) To UBound(theSDInfoLst(SI))
					If FI <> SD_STAT Then
						theSDInfoLst(SI)(FI) = ReplaceDefines(theSDInfoLst(SI)(FI),theDefines)
					End If
				Next
				
				If theSDInfoLst(SI)(SD_TYPE) <> "" And theSDInfoLst(SI)(SD_LEND) <> "" And SI = StatPos Then
If theSDInfoLst(SI)(SD_SCRIPT) = SCRIPT_SEQUENTIAL Then
If OutFileOption = optOutFileBas Then
.WriteText sIndent+"Select Case L-DblCalcDev", adWriteLine
.WriteText sIndent+"	Case <= "+theSDInfoLst(SI)(SD_LEND)+"#", adWriteLine
ElseIf OutFileOption = optOutFileLua Then
.WriteText sIndent+"if L-DblCalcDev <= "+theSDInfoLst(SI)(SD_LEND)+" then", adWriteLine
ElseIf OutFileOption = optOutFileVbs Then
.WriteText sIndent+"If L-DblCalcDev <= "+theSDInfoLst(SI)(SD_LEND)+" Then", adWriteLine
End If
ElseIf theSDInfoLst(SI)(SD_SCRIPT) = SCRIPT_INTERVAL Then
If OutFileOption = optOutFileBas Then
.WriteText sIndent+"If "+theSDInfoLst(SI)(SD_LSTART)+" <= L+DblCalcDev And L-DblCalcDev <= "+theSDInfoLst(SI)(SD_LEND)+" Then", adWriteLine
ElseIf OutFileOption = optOutFileLua Then
.WriteText sIndent+"if "+theSDInfoLst(SI)(SD_LSTART)+" <= L+DblCalcDev and L-DblCalcDev <= "+theSDInfoLst(SI)(SD_LEND)+" then", adWriteLine
ElseIf OutFileOption = optOutFileVbs Then
.WriteText sIndent+"If "+theSDInfoLst(SI)(SD_LSTART)+" <= L+DblCalcDev And L-DblCalcDev <= "+theSDInfoLst(SI)(SD_LEND)+" Then", adWriteLine
End If
End If
				ElseIf theSDInfoLst(SI)(SD_TYPE) <> "" And theSDInfoLst(SI)(SD_LEND) <> "" Then
If theSDInfoLst(SI)(SD_SCRIPT) = SCRIPT_SEQUENTIAL Then
If OutFileOption = optOutFileBas Then
.WriteText sIndent+"	Case <= "+theSDInfoLst(SI)(SD_LEND)+"#", adWriteLine
ElseIf OutFileOption = optOutFileLua Then
.WriteText sIndent+"elseif L-DblCalcDev <= "+theSDInfoLst(SI)(SD_LEND)+" then", adWriteLine
ElseIf OutFileOption = optOutFileVbs Then
.WriteText sIndent+"ElseIf L-DblCalcDev <= "+theSDInfoLst(SI)(SD_LEND)+" Then", adWriteLine
End If
ElseIf theSDInfoLst(SI)(SD_SCRIPT) = SCRIPT_INTERVAL Then
If OutFileOption = optOutFileBas Then
.WriteText sIndent+"ElseIf "+theSDInfoLst(SI)(SD_LSTART)+" <= L+DblCalcDev And L-DblCalcDev <= "+theSDInfoLst(SI)(SD_LEND)+" Then", adWriteLine
ElseIf OutFileOption = optOutFileLua Then
.WriteText sIndent+"elseif "+theSDInfoLst(SI)(SD_LSTART)+" <= L+DblCalcDev and L-DblCalcDev <= "+theSDInfoLst(SI)(SD_LEND)+" then", adWriteLine
ElseIf OutFileOption = optOutFileVbs Then
.WriteText sIndent+"ElseIf "+theSDInfoLst(SI)(SD_LSTART)+" <= L+DblCalcDev And L-DblCalcDev <= "+theSDInfoLst(SI)(SD_LEND)+" Then", adWriteLine
End If
End If
				ElseIf theSDInfoLst(SI)(SD_TYPE) <> "" And ((theSDInfoLst(SI)(SD_SCRIPT) = SCRIPT_SEQUENTIAL And theSDInfoLst(SI)(SD_LSTART) <> "") Or (theSDInfoLst(SI)(SD_SCRIPT) = SCRIPT_INTERVAL And theSDInfoLst(SI)(SD_LSTART) = "" And theSDInfoLst(SI)(SD_LEND) = "")) Then
If theSDInfoLst(SI)(SD_SCRIPT) = SCRIPT_SEQUENTIAL Then
If OutFileOption = optOutFileBas Then
.WriteText sIndent+"	Case Else", adWriteLine
ElseIf OutFileOption = optOutFileLua Then
.WriteText sIndent+"else", adWriteLine
ElseIf OutFileOption = optOutFileVbs Then
.WriteText sIndent+"Else", adWriteLine
End If
ElseIf theSDInfoLst(SI)(SD_SCRIPT) = SCRIPT_INTERVAL Then
If OutFileOption = optOutFileBas Then
.WriteText sIndent+"Else", adWriteLine
ElseIf OutFileOption = optOutFileLua Then
.WriteText sIndent+"else", adWriteLine
ElseIf OutFileOption = optOutFileVbs Then
.WriteText sIndent+"Else", adWriteLine
End If
End If
				End If

				Build = ""

				If theSDInfoLst(SI)(SD_TYPE) = "A" Then
					Build = theSDInfoLst(SI)(SD_P1)
				ElseIf theSDInfoLst(SI)(SD_TYPE) = "C" Then
					Build = theSDInfoLst(SI)(SD_P1)
				ElseIf theSDInfoLst(SI)(SD_TYPE) = "D" Then
If OutFileOption = optOutFileBas Then
					Build = "DataTableValue(Array("+theSDInfoLst(SI)(SD_P1)
ElseIf OutFileOption = optOutFileLua Then
					Build = "DataTableValue({"+theSDInfoLst(SI)(SD_P1)
ElseIf OutFileOption = optOutFileVbs Then
					Build = "DataTableValue(Array("+theSDInfoLst(SI)(SD_P1)
End If
					Dim DI
					For DI = SI+1 To UBound(theSDInfoLst)
						If theSDInfoLst(DI)(SD_STAT) <> theStat Or theSDInfoLst(DI)(SD_TYPE) <> "" Then
							Exit For
						End If
						Build = Build+","+theSDInfoLst(DI)(SD_P1)
					Next
If OutFileOption = optOutFileBas Then
						Build = Build+")"
ElseIf OutFileOption = optOutFileLua Then
						Build = Build+"}"
ElseIf OutFileOption = optOutFileVbs Then
						Build = Build+")"
End If
					If theSDInfoLst(SI)(SD_LSTART) <> "" And theSDInfoLst(SI)(SD_LSTART) <> "1" Then
						Build = Build+",L-"+CStr(theSDInfoLst(SI)(SD_LSTART)-1)+")"
					Else
						Build = Build+",L)"
					End If
				ElseIf theSDInfoLst(SI)(SD_TYPE) = "E" Then
					If theSDInfoLst(SI)(SD_LSTART) <> "" Then
						Build = "ExpFmod("+theSDInfoLst(SI)(SD_P1)+","+theSDInfoLst(SI)(SD_LSTART)+","+theSDInfoLst(SI)(SD_P2)+","
					Else
						Build = "ExpFmod("+theSDInfoLst(SI)(SD_P1)+",1,"+theSDInfoLst(SI)(SD_P2)+","
					End If
					IF theSDInfoLst(SI)(SD_P3) <> "" Then
						Build = Build+"L+"+theSDInfoLst(SI)(SD_P3)+")"
					Else
						Build = Build+"L)"
					End If
				ElseIf theSDInfoLst(SI)(SD_TYPE) = "F" Then
					Build = theSDInfoLst(SI)(SD_P1)+"("
					If theSDInfoLst(SI)(SD_P2) <> "" Then
						Build = Build+theSDInfoLst(SI)(SD_P2)
					If theSDInfoLst(SI)(SD_P3) <> "" Then
						Build = Build+","+theSDInfoLst(SI)(SD_P3)
					If theSDInfoLst(SI)(SD_P4) <> "" Then
						Build = Build+","+theSDInfoLst(SI)(SD_P4)
					If theSDInfoLst(SI)(SD_P5) <> "" Then
						Build = Build+","+theSDInfoLst(SI)(SD_P5)
					If theSDInfoLst(SI)(SD_P6) <> "" Then
						Build = Build+","+theSDInfoLst(SI)(SD_P6)
					End If
					End If
					End If
					End If
					End If
					Build = Build+")"
				ElseIf theSDInfoLst(SI)(SD_TYPE) = "L" Then
					If InStr(theSDInfoLst(SI)(SD_P1),"/") Or InStr(theSDInfoLst(SI)(SD_P1),"+") Or InStr(theSDInfoLst(SI)(SD_P1),"-") Then
						Build = "("+theSDInfoLst(SI)(SD_P1)+")*L"
					ElseIf theSDInfoLst(SI)(SD_P1) = "1" Then
						Build = "L"
					Else
						Build = theSDInfoLst(SI)(SD_P1)+"*L"
					End If
					If theSDInfoLst(SI)(SD_P2) <> "" Then
						Build = Build+"+"+theSDInfoLst(SI)(SD_P2)
					End If
				ElseIf theSDInfoLst(SI)(SD_TYPE) = "M" Then
					If theSDInfoLst(SI)(SD_P2) <> theSDInfoLst(SI)(SD_P3) Then
						Build = "LinFmod(1,"
						If InStr(theSDInfoLst(SI)(SD_P2),"/") Or InStr(theSDInfoLst(SI)(SD_P2),"+") Or InStr(theSDInfoLst(SI)(SD_P2),"-") Then
							Build = Build+"("+theSDInfoLst(SI)(SD_P2)+")*"
						ElseIf theSDInfoLst(SI)(SD_P2) <> "1" Then
							Build = Build+theSDInfoLst(SI)(SD_P2)+"*"
						End If
						Build = Build+"CalcStat("""+theSDInfoLst(SI)(SD_P1)+""","+theSDInfoLst(SI)(SD_P4)+",N),"
						If InStr(theSDInfoLst(SI)(SD_P3),"/") Or InStr(theSDInfoLst(SI)(SD_P3),"+") Or InStr(theSDInfoLst(SI)(SD_P3),"-") Then
							Build = Build+"("+theSDInfoLst(SI)(SD_P3)+")*"
						ElseIf theSDInfoLst(SI)(SD_P3) <> "1" Then
							Build = Build+theSDInfoLst(SI)(SD_P3)+"*"
						End If
						Build = Build+"CalcStat("""+theSDInfoLst(SI)(SD_P1)+""","+theSDInfoLst(SI)(SD_P5)+",N),"
						Build = Build+theSDInfoLst(SI)(SD_P4)+","+theSDInfoLst(SI)(SD_P5)+",L"
If OutFileOption = optOutFileVbs Then
						Build = Build+",Nothing"
End If 
						Build = Build+")/CalcStat("""+theSDInfoLst(SI)(SD_P1)+""",L,N)"
					Else
						Build = theSDInfoLst(SI)(SD_P2)
					End If
				ElseIf theSDInfoLst(SI)(SD_TYPE) = "N" Then
					If theSDInfoLst(SI)(SD_LSTART) <> "" And theSDInfoLst(SI)(SD_LSTART) <> "1" Then
						Build = "NamedRangeValue("""+theSDInfoLst(SI)(SD_P1)+""",L-"+CStr(theSDInfoLst(SI)(SD_LSTART)-1)+","+theSDInfoLst(SI)(SD_P2)+")"
					Else
						Build = "NamedRangeValue("""+theSDInfoLst(SI)(SD_P1)+""",L,"+theSDInfoLst(SI)(SD_P2)+")"
					End If
				ElseIf theSDInfoLst(SI)(SD_TYPE) = "P" Then
					Build = "CalcPercAB("+theSDInfoLst(SI)(SD_P1)+","+theSDInfoLst(SI)(SD_P2)+","+theSDInfoLst(SI)(SD_P3)+",N)"
				ElseIf theSDInfoLst(SI)(SD_TYPE) = "R" Then
					Build = "CalcRatAB("+theSDInfoLst(SI)(SD_P1)+","+theSDInfoLst(SI)(SD_P2)+","+theSDInfoLst(SI)(SD_P3)+",N)"
				ElseIf theSDInfoLst(SI)(SD_TYPE) = "T" Then
					If theSDInfoLst(SI)(SD_P6) <> "" Then
						Build = "LinFmod("+theSDInfoLst(SI)(SD_P1)+","+theSDInfoLst(SI)(SD_P2)+","+theSDInfoLst(SI)(SD_P3)+","+theSDInfoLst(SI)(SD_P4)+","+theSDInfoLst(SI)(SD_P5)+",L,"+theSDInfoLst(SI)(SD_P6)+")"
					Else
If OutFileOption = optOutFileVbs Then
						Build = "LinFmod("+theSDInfoLst(SI)(SD_P1)+","+theSDInfoLst(SI)(SD_P2)+","+theSDInfoLst(SI)(SD_P3)+","+theSDInfoLst(SI)(SD_P4)+","+theSDInfoLst(SI)(SD_P5)+",L,Nothing)"
Else
						Build = "LinFmod("+theSDInfoLst(SI)(SD_P1)+","+theSDInfoLst(SI)(SD_P2)+","+theSDInfoLst(SI)(SD_P3)+","+theSDInfoLst(SI)(SD_P4)+","+theSDInfoLst(SI)(SD_P5)+",L)"
End If 
					End If
				End If

				If Build <> "" Then

					If Ucase(theSDInfoLst(SI)(SD_USEN)) <> "N" And InStr("CDEFLNT",theSDInfoLst(SI)(SD_TYPE)) Then
						If Build = "0" Then
						ElseIf Build = "1" Then
							Build = "N"
						ElseIf InStr("DEFNT",theSDInfoLst(SI)(SD_TYPE)) Then
							Build = Build+"*N"
						ElseIf InStr(Build,"+") Or InStr(Build,"-") Or InStr(Build,"/") Then
							Build = "("+Build+")*N"
						Else
							Build = Build+"*N"
						End If
					End If

					If theSDInfoLst(SI)(SD_DECIMALS) <> "" Then
						If theSDInfoLst(SI)(SD_DECIMALS) = "0" Then
If OutFileOption = optOutFileVbs Then
							Build = "RoundDbl("+Build+",Nothing)"
Else
							Build = "RoundDbl("+Build+")"
End If
						Else
							Build = "RoundDbl("+Build+","+theSDInfoLst(SI)(SD_DECIMALS)+")"
						End If
					End If
	
					Build = Replace(Replace(Build,"!",""""),"|",",")
	
					Build = Replace(Build,"+-","-")
					Build = Replace(Build,"--","+")
If OutFileOption <> optOutFileVbs Then
					Build = Replace(Build,",Nothing","")
End If
					Build = ReplStatRefs(Build)

					If theSDInfoLst(SI)(SD_LEND) <> "" Or theSDInfoLst(SI)(SD_LSTART) <> "" Then
If theSDInfoLst(StatPos)(SD_SCRIPT) = SCRIPT_SEQUENTIAL Then
If OutFileOption = optOutFileBas Then
.WriteText sIndent+"		CalcStat = "+Build, adWriteLine
ElseIf OutFileOption = optOutFileLua Then
.WriteText sIndent+"	return "+Build+";", adWriteLine
ElseIf OutFileOption = optOutFileVbs Then
.WriteText sIndent+"	CalcStat = "+Build, adWriteLine
End If
ElseIf theSDInfoLst(StatPos)(SD_SCRIPT) = SCRIPT_INTERVAL Then
If OutFileOption = optOutFileBas Then
.WriteText sIndent+"	CalcStat = "+Build, adWriteLine
ElseIf OutFileOption = optOutFileLua Then
.WriteText sIndent+"	return "+Build+";", adWriteLine
ElseIf OutFileOption = optOutFileVbs Then
.WriteText sIndent+"	CalcStat = "+Build, adWriteLine
End If
End If
					Else
If theSDInfoLst(StatPos)(SD_SCRIPT) = SCRIPT_SEQUENTIAL Then
If OutFileOption = optOutFileBas Then
.WriteText sIndent+"CalcStat = "+Build, adWriteLine
ElseIf OutFileOption = optOutFileLua Then
.WriteText sIndent+"return "+Build+";", adWriteLine
ElseIf OutFileOption = optOutFileVbs Then
.WriteText sIndent+"CalcStat = "+Build, adWriteLine
End If
ElseIf theSDInfoLst(StatPos)(SD_SCRIPT) = SCRIPT_INTERVAL Then
If OutFileOption = optOutFileBas Then
.WriteText sIndent+"	CalcStat = "+Build, adWriteLine
ElseIf OutFileOption = optOutFileLua Then
.WriteText sIndent+"	return "+Build+";", adWriteLine
ElseIf OutFileOption = optOutFileVbs Then
.WriteText sIndent+"	CalcStat = "+Build, adWriteLine
End If
End If
					End If

				End If
			Next	

			If theSDInfoLst(StatPos)(SD_LSTART) <> "" Or theSDInfoLst(StatPos)(SD_LEND) <> "" Then
If theSDInfoLst(StatPos)(SD_SCRIPT) = SCRIPT_SEQUENTIAL Then
If OutFileOption = optOutFileBas Then
.WriteText sIndent+"End Select", adWriteLine
ElseIf OutFileOption = optOutFileLua Then
.WriteText sIndent+"end", adWriteLine
ElseIf OutFileOption = optOutFileVbs Then
.WriteText sIndent+"End If", adWriteLine
End If
ElseIf theSDInfoLst(StatPos)(SD_SCRIPT) = SCRIPT_INTERVAL Then
If OutFileOption = optOutFileBas Then
.WriteText sIndent+"End If", adWriteLine
ElseIf OutFileOption = optOutFileLua Then
.WriteText sIndent+"end", adWriteLine
ElseIf OutFileOption = optOutFileVbs Then
.WriteText sIndent+"End If", adWriteLine
End If
End If
			End If

		End If

	End With

End Sub

Function ReplStatRefs(ByVal theBuild)

	Dim Result
	Result = theBuild

	Dim nSearch
	Dim BI
	Dim sStartToken, sNumToken, sEncapStartToken, sEncapEndToken
	Dim nEncapOpen
	Dim sStatName
	Dim cStatChar
	Dim sLevel
	Dim sNumber

	sStartToken = "@"
	sNumToken = "#"
	sEncapStartToken = "{"
	sEncapEndToken = "}"

	nSearch = InStr(Result,sStartToken)
	While nSearch > 0
		Result = Left(Result,nSearch-1)+Right(Result,Len(Result)-nSearch+1-Len(sStartToken))

		sStatName = ""
		For BI = nSearch To Len(Result)
			cStatChar = Mid(Result,BI,1)
			If ("A" <= UCase(cStatChar) And UCase(cStatChar) <= "Z") Or (("0" <= UCase(cStatChar) And UCase(cStatChar) <= "9") And ("A" <= UCase(Mid(Result,BI+1,1)) And UCase(Mid(Result,BI+1,1)) <= "Z"))Then
				sStatName = sStatName+cStatChar
			Else
				Exit For	
			End If
		Next

		If Len(sStatName) > 0 Then
			Result = Left(Result,nSearch-1)+"CalcStat("""+sStatName+""""+Right(Result,Len(Result)-nSearch+1-Len(sStatName))
			nSearch = nSearch+Len("CalcStat("""+sStatName+"""")

			sLevel = ""
			If Mid(Result,nSearch,Len(sEncapStartToken)) = sEncapStartToken Then
				Result = Left(Result,nSearch-1)+Right(Result,Len(Result)-nSearch+1-Len(sEncapStartToken))
				nEncapOpen = 1
				For BI = nSearch To Len(Result)
					cStatChar = Mid(Result,BI,1)
					If Mid(Result,BI,Len(sEncapStartToken)) = sEncapStartToken Then
						nEncapOpen = nEncapOpen+1
					End If
					If Mid(Result,BI,Len(sEncapEndToken)) = sEncapEndToken Then
						nEncapOpen = nEncapOpen-1
					End If
					If nEncapOpen > 0 Then
						sLevel = sLevel+cStatChar
					Else
						Result = Left(Result,BI-1)+Right(Result,Len(Result)-BI+1-Len(sEncapEndToken))
						Exit For
					End If
				Next
			ElseIf "0" <= Mid(Result,nSearch,1) And Mid(Result,nSearch,1) <= "9" Then
				For BI = nSearch To Len(Result)
					cStatChar = Mid(Result,BI,1)
					If "0" <= cStatChar And cStatChar <= "9" Then
						sLevel = sLevel+cStatChar
					Else
						Exit For
					End If
				Next
			End If
			If sLevel = "" Then
				Result = Left(Result,nSearch-1)+",L"+Right(Result,Len(Result)-nSearch+1)
				nSearch = nSearch+Len(",L")
			Else
				Result = Left(Result,nSearch-1)+","+sLevel+Right(Result,Len(Result)-nSearch+1-Len(sLevel))
				nSearch = nSearch+Len(","+sLevel)
			End If

			If Mid(Result,nSearch,Len(sNumToken)) = sNumToken Then
				Result = Left(Result,nSearch-1)+Right(Result,Len(Result)-nSearch+1-Len(sNumToken))

				sNumber = ""
				If Mid(Result,nSearch,Len(sEncapStartToken)) = sEncapStartToken Then
					Result = Left(Result,nSearch-1)+Right(Result,Len(Result)-nSearch+1-Len(sEncapStartToken))
					nEncapOpen = 1
					For BI = nSearch To Len(Result)
						cStatChar = Mid(Result,BI,1)
						If Mid(Result,BI,Len(sEncapStartToken)) = sEncapStartToken Then
							nEncapOpen = nEncapOpen+1
						End If
						If Mid(Result,BI,Len(sEncapEndToken)) = sEncapEndToken Then
							nEncapOpen = nEncapOpen-1
						End If
						If nEncapOpen > 0 Then
							sNumber = sNumber+cStatChar
						Else
							Result = Left(Result,BI-1)+Right(Result,Len(Result)-BI+1-Len(sEncapEndToken))
							rem MsgBox "["+Result+"]"
							Exit For
						End If
					Next
				ElseIf "0" <= Mid(Result,nSearch,1) And Mid(Result,nSearch,1) <= "9" Then
					For BI = nSearch To Len(Result)
						cStatChar = Mid(Result,BI,1)
						If ("0" <= cStatChar And cStatChar <= "9") Or cStatChar = "." Then
							sNumber = sNumber+cStatChar
						Else
							Exit For
						End If
					Next
				End If
				If sNumber = "" Then
					Result = Left(Result,nSearch-1)+",N)"+Right(Result,Len(Result)-nSearch+1)
					nSearch = nSearch+Len(",N)")
				Else
					Result = Left(Result,nSearch-1)+","+sNumber+")"+Right(Result,Len(Result)-nSearch+1-Len(sNumber))
					nSearch = nSearch+Len(","+sNumber+")")
				End If
			Else
If OutFileOption = optOutFileVbs Then
				Result = Left(Result,nSearch-1)+",Nothing)"+Right(Result,Len(Result)-nSearch+1)
				nSearch = nSearch+Len(",Nothing)")
Else
				Result = Left(Result,nSearch-1)+")"+Right(Result,Len(Result)-nSearch+1)
				nSearch = nSearch+Len(")")
End If
			End If

		End If

		nSearch = InStr(Result,sStartToken)
	Wend

	ReplStatRefs = Result

End Function

Sub ReadSDFile(ByVal theInFile, ByRef theSDInfoLst)

	Dim InStream
	Set InStream = CreateObject("ADODB.Stream")
	With InStream
		.Type = adTypeText
		.CharSet = "UTF-8"
		.LineSeparator = adLF
		.Open
		.LoadFromFile theInFile

		Dim ReadError
		ReadError = False

		Dim RowText
		RowText = ""

		'skip header line
		RowText = .ReadText(adReadLine)

		Do Until .EOS OR ReadError
			RowText = .ReadText(adReadLine)

			RowText = Replace(RowText,Chr(10),"")
			RowText = Replace(RowText,Chr(13),"")

			Dim NewSDInfo
			InitSDInfo NewSDInfo

			For FldI = LBound(NewSDInfo) To UBound(NewSDInfo)
				ReadSDField RowText, NewSDInfo(FldI)
			Next

			If NewSDInfo(SD_COMPF) = "" Then
				NewSDInfo(SD_COMPF) = "0"
			End If
			
			If NewSDInfo(SD_STAT) = "-VERSION" Then
				Select Case CompileOption
					Case optCompileFull
						NewSDInfo(SD_P1) = Replace(NewSDInfo(SD_P1),"v","f")
					Case optCompilePercentages
						NewSDInfo(SD_P1) = Replace(NewSDInfo(SD_P1),"v","p")
				End Select
			End If
			
			If NewSDInfo(SD_STAT) = "" Then
				Break
			ElseIf CompileOption = optCompileFull Or (CInt(NewSDInfo(SD_COMPF)) And CompileOption) Then
				If IsEmpty(theSDInfoLst) Then
					ReDim theSDInfoLst(0)
				Else
					ReDim Preserve theSDInfoLst(UBound(theSDInfoLst)+1)
				End If
				theSDInfoLst(UBound(theSDInfoLst)) = NewSDInfo
			End If
		Loop

		.Close
	End With
	Set InStream = Nothing

End Sub

Sub InitSDInfo(ByRef theSDInfo)

	theSDInfo = Split(String(SDFLDCNT-1,","),",",SDFLDCNT)
				
End Sub

Sub ReadSDField(ByRef RowText, ByRef NewSDFld)

	Dim SepPos
	SepPos = InStr(RowText,SDFLDSEP)
	If SepPos > 0 Then
		NewSDFld = RTrim(Mid(RowText,1,SepPos-1))
		RowText = LTrim(Mid(RowText,SepPos+1,Len(RowText)-SepPos))
	Else
		NewSDFld = Rtrim(RowText)
		RowText = ""
	End If

	If Len(NewSDFld) >= 2 Then
		If Mid(NewSDFld,1,1) = SDSTRQUOT AND Mid(NewSDFld,Len(NewSDFld),1) = SDSTRQUOT Then
			NewSDFld = Mid(NewSDFld,2,Len(NewSDFld)-2)
		End If
	End If

End Sub
