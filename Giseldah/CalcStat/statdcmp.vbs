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

Const InFileName = "statdata.csv"
Const DataFileName = "statdcmp.dat"

'script data
Dim ScriptDescr
Dim OutFileName
Dim STextStart
Dim STextEnd
Dim STextSNSearchCondStartSm
Dim STextSNSearchCondSm
Dim STextSNSearchCondStartEq
Dim STextSNSearchCondEq
Dim STextSNSearchCondStartGr
Dim STextSNSearchCondGr
Dim STextSNSearchAlternative
Dim STextSNSearchEnd
Dim STextSNSearchResult
Dim STextSeqLSearchCondStart
Dim STextSeqLSearchCond
Dim STextSeqLSearchAlternative
Dim STextSeqLSearchEnd
Dim STextSeqLSearchResult
Dim STextIntLSearchCondStart
Dim STextIntLSearchCond
Dim STextIntLSearchAlternative
Dim STextIntLSearchEnd
Dim STextIntLSearchResult
Dim	STextDefineC
Dim	STextDefineL
Dim	STextDefineN
Dim	STextDefinePNULL
Dim	STextDefineSN
Dim STextArrayStart
Dim STextArraySep
Dim STextArrayEnd

'script text variables
Const STV_SN = "$STATNAME"
Const STV_LVL = "$LEVEL"
Const STV_LSTART = "$LVLSTART"
Const STV_LEND = "$LVLEND"
Const STV_CALC = "$CALC"

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

'defines
Const DF_DEF = 0 'defined text to replace
Const DF_REPL = 1 'replacement text

Const optCompileFull = 0 'full version with all stats
Const optCompilePercentages = 1 'percentage calculations only

Dim CompileOption 'compile option choice

Main
Sub Main

	Dim DataArray

	If Not ReadDataFile(DataFileName,DataArray) Then
		MsgBox Replace("Could not read #S# file. Operation cancelled.","#S#",DataFileName)
		Exit Sub
	End If
	
	Dim ScrSelText
	Dim ScrSelDefault
	Dim ScrSelCount
	
	GetScrSelData DataArray,ScrSelText,ScrSelDefault,ScrSelCount
	
	Dim OutFileOption

	OutFileOption = CInt(InputBox(ScrSelText,"Choose output script type",ScrSelDefault))

	If Not (1 <= OutFileOption And OutFileOption <= ScrSelCount) Then
		MsgBox "Invalid option. Operation cancelled."
		Exit Sub
	End If
	
	CompileOption = CInt(InputBox("Enter a Number:"&Chr(10)&"1 = Full version with all stats"&Chr(10)&"2 = Percentage calculations only <- normal for plugins","Choose output version",2))-1

	If Not (CompileOption = optCompileFull Or CompileOption = optCompilePercentages) Then
		MsgBox "Invalid option. Operation cancelled."
		Exit Sub
	End If
	
	GetScriptData DataArray,OutFileOption
	
	Dim SDInfoLst

	ReadSDFile InFileName,SDInfoLst
	
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

		WriteSText OutStream,STextStart,""

		Dim iNodeIndex
		iNodeIndex = LBound(aBalancedBST)	'Head node of the tree is first in the list

		WriteScrStatsNode OutStream,SDInfoLst,aDefines,aBalancedBST,iNodeIndex,""

		WriteSText OutStream,STextEnd,""

		.SaveToFile OutFileName,adSaveCreateOverWrite
		.Close
	End With
	Set OutStream = Nothing

	MsgBox Replace(Replace("Compilation of #S1# to #S2# completed.","#S1#",InFileName),"#S2#",OutFileName)
	
End Sub

Function ReadDataFile(ByVal theDataFileName, ByRef theDataArray)

	Dim Result
	Result = True

	Dim DataStream
	Set DataStream = CreateObject("ADODB.Stream")
	With DataStream
		.Type = adTypeText
		.CharSet = "UTF-8"
		.LineSeparator = adLF
		.Open
		.LoadFromFile theDataFileName

		Dim RowText

		Do Until .EOS
			RowText = .ReadText(adReadLine)

			RowText = Replace(RowText,Chr(10),"")
			RowText = Replace(RowText,Chr(13),"")

			If IsEmpty(theDataArray) Then
				ReDim theDataArray(0)
			Else
				ReDim Preserve theDataArray(UBound(theDataArray)+1)
			End If
			theDataArray(UBound(theDataArray)) = RowText
		Loop

		.Close
	End With
	Set DataStream = Nothing

	If IsEmpty(theDataArray) Then
		Result = False
	End If
	
	ReadDataFile = Result
	
End Function

Sub	GetScrSelData(ByVal theDataArray, ByRef theScrSelText, ByRef theScrSelDefault, ByRef theScrSelCount)

	Dim SelText
	SelText = "Enter a Number:"

	Dim SelCount
	SelCount = 0
	
	Dim I
	I = LBound(theDataArray)
	Dim MaxI
	MaxI = UBound(theDataArray)
	
	Dim ScrTextLen
	
	'default script number
	theScrSelDefault = CInt(theDataArray(I))
	I = I+1
	
	Do Until MaxI-I < 18
		'script description
		SelCount = SelCount+1
		SelText = SelText & Chr(10) & CStr(SelCount) & " = " & theDataArray(I)
		I = I+1

		'outfilename
		I = I+1

		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script start text
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script end text
		I = I+ScrTextLen

		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script SN search condition start smaller text
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script SN search condition smaller text
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script SN search condition start equal text
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script SN search condition equal text
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script SN search condition start greater text
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script SN search condition greater text
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script SN search alternative text
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script SN search end text
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script SN search result text
		I = I+ScrTextLen

		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script Sequential Level Search condition start text
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script Sequential Level Search condition text
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script Sequential Level Search alternative text
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script Sequential Level Search end text
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script Sequential Level Search result text
		I = I+ScrTextLen
		
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script Interval Level Search condition start text
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script Interval Level Search condition text
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script Interval Level Search alternative text
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script Interval Level Search end text
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script Interval Level Search result text
		I = I+ScrTextLen
		
		'script Define C text
		I = I+1
		'script Define L text
		I = I+1
		'script Define N text
		I = I+1
		'script Define PNULL text
		I = I+1
		'script Define SN text
		I = I+1

		'script Array start text
		I = I+1
		'script Array separator text
		I = I+1
		'script Array end text
		I = I+1
	Loop
	
	theScrSelText = SelText
	theScrSelCount = SelCount
	
End Sub

Sub	GetScriptData(ByVal theDataArray, ByVal theOutFileOption)

	Dim ScriptCount
	ScriptCount = 0
	
	Dim I
	I = LBound(theDataArray)
	Dim MaxI
	MaxI = UBound(theDataArray)
	
	Dim ScrTextLen
	
	'default script number
	I = I+1
	
	Do Until MaxI-I < 18
		ScriptCount = ScriptCount+1

		'script description
		If ScriptCount = theOutFileOption Then
			ScriptDescr = theDataArray(I)
		End If
		I = I+1

		'outfilename
		If ScriptCount = theOutFileOption Then
			OutFileName = theDataArray(I)
		End If
		I = I+1

		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script start text
		If ScriptCount = theOutFileOption Then
			STextStart = GetScriptText(theDataArray,I,ScrTextLen)
		End If
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script end text
		If ScriptCount = theOutFileOption Then
			STextEnd = GetScriptText(theDataArray,I,ScrTextLen)
		End If
		I = I+ScrTextLen

		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script SN search condition start smaller text
		If ScriptCount = theOutFileOption Then
			STextSNSearchCondStartSm = GetScriptText(theDataArray,I,ScrTextLen)
		End If
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script SN search condition smaller text
		If ScriptCount = theOutFileOption Then
			STextSNSearchCondSm = GetScriptText(theDataArray,I,ScrTextLen)
		End If
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script SN search condition start equal text
		If ScriptCount = theOutFileOption Then
			STextSNSearchCondStartEq = GetScriptText(theDataArray,I,ScrTextLen)
		End If
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script SN search condition equal text
		If ScriptCount = theOutFileOption Then
			STextSNSearchCondEq = GetScriptText(theDataArray,I,ScrTextLen)
		End If
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script SN search condition start greater text
		If ScriptCount = theOutFileOption Then
			STextSNSearchCondStartGr = GetScriptText(theDataArray,I,ScrTextLen)
		End If
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script SN search condition greater text
		If ScriptCount = theOutFileOption Then
			STextSNSearchCondGr = GetScriptText(theDataArray,I,ScrTextLen)
		End If
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script SN search alternative text
		If ScriptCount = theOutFileOption Then
			STextSNSearchAlternative = GetScriptText(theDataArray,I,ScrTextLen)
		End If
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script SN search end text
		If ScriptCount = theOutFileOption Then
			STextSNSearchEnd = GetScriptText(theDataArray,I,ScrTextLen)
		End If
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script SN search result text
		If ScriptCount = theOutFileOption Then
			STextSNSearchResult = GetScriptText(theDataArray,I,ScrTextLen)
		End If
		I = I+ScrTextLen

		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script Sequential Level Search condition start text
		If ScriptCount = theOutFileOption Then
			STextSeqLSearchCondStart = GetScriptText(theDataArray,I,ScrTextLen)
		End If
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script Sequential Level Search condition text
		If ScriptCount = theOutFileOption Then
			STextSeqLSearchCond = GetScriptText(theDataArray,I,ScrTextLen)
		End If
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script Sequential Level Search alternative text
		If ScriptCount = theOutFileOption Then
			STextSeqLSearchAlternative = GetScriptText(theDataArray,I,ScrTextLen)
		End If
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script Sequential Level Search end text
		If ScriptCount = theOutFileOption Then
			STextSeqLSearchEnd = GetScriptText(theDataArray,I,ScrTextLen)
		End If
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script Sequential Level Search result text
		If ScriptCount = theOutFileOption Then
			STextSeqLSearchResult = GetScriptText(theDataArray,I,ScrTextLen)
		End If
		I = I+ScrTextLen
		
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script Interval Level Search condition start text
		If ScriptCount = theOutFileOption Then
			STextIntLSearchCondStart = GetScriptText(theDataArray,I,ScrTextLen)
		End If
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script Interval Level Search condition text
		If ScriptCount = theOutFileOption Then
			STextIntLSearchCond = GetScriptText(theDataArray,I,ScrTextLen)
		End If
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script Interval Level Search alternative text
		If ScriptCount = theOutFileOption Then
			STextIntLSearchAlternative = GetScriptText(theDataArray,I,ScrTextLen)
		End If
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script Interval Level Search end text
		If ScriptCount = theOutFileOption Then
			STextIntLSearchEnd = GetScriptText(theDataArray,I,ScrTextLen)
		End If
		I = I+ScrTextLen
		ScrTextLen = CInt(theDataArray(I))
		I = I+1
		'script Interval Level Search result text
		If ScriptCount = theOutFileOption Then
			STextIntLSearchResult = GetScriptText(theDataArray,I,ScrTextLen)
		End If
		I = I+ScrTextLen
		
		'script Define C text
		If ScriptCount = theOutFileOption Then
			STextDefineC = theDataArray(I)
		End If
		I = I+1
		'script Define L text
		If ScriptCount = theOutFileOption Then
			STextDefineL = theDataArray(I)
		End If
		I = I+1
		'script Define N text
		If ScriptCount = theOutFileOption Then
			STextDefineN = theDataArray(I)
		End If
		I = I+1
		'script Define PNULL text
		If ScriptCount = theOutFileOption Then
			STextDefinePNULL = theDataArray(I)
		End If
		I = I+1
		'script Define SN text
		If ScriptCount = theOutFileOption Then
			STextDefineSN = theDataArray(I)
		End If
		I = I+1

		'script Array start text
		If ScriptCount = theOutFileOption Then
			STextArrayStart = theDataArray(I)
		End If
		I = I+1
		'script Array separator text
		If ScriptCount = theOutFileOption Then
			STextArraySep = theDataArray(I)
		End If
		I = I+1
		'script Array end text
		If ScriptCount = theOutFileOption Then
			STextArrayEnd = theDataArray(I)
		End If
		I = I+1
	Loop
	
	theScrSelText = SelText
	theScrSelCount = SelCount
	
End Sub

Function GetScriptText(theDataArray,I,ScrTextLen)

	Dim ScriptText
	
	For II = I To I+ScrTextLen-1
		If IsEmpty(ScriptText) Then
			ReDim ScriptText(0)
		Else
			ReDim Preserve ScriptText(UBound(ScriptText)+1)
		End If
		ScriptText(UBound(ScriptText)) = theDataArray(II)
	Next
	
	GetScriptText = ScriptText
	
End Function

Function ReadDefines(theSDInfoLst)

	Dim aDefines

	Dim sStat, sLastStat
	sLastStat = ""

	Redim aDefines(4)
	aDefines(0) = Array("$C",STextDefineC)
	aDefines(1) = Array("$L",STextDefineL)
	aDefines(2) = Array("$N",STextDefineN)
	aDefines(3) = Array("$PNULL",STextDefinePNULL)
	aDefines(4) = Array("$SN",STextDefineSN)
	
	For SI = LBound(theSDInfoLst) To UBound(theSDInfoLst)
		If theSDInfoLst(SI)(SD_SCRIPT) = SCRIPT_DEFINE Then
			sStat = UCase(theSDInfoLst(SI)(SD_STAT))
			If sStat <> sLastStat Then
				Redim Preserve aDefines(UBound(aDefines)+1)
				aDefines(UBound(aDefines)) = Array(sStat,theSDInfoLst(SI)(SD_P1))
				sLastStat = sStat
			End If
		End If
	Next

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
	
	ReadDefines = aDefines

End Function

Function ReplaceDefines(theText, theDefines)

	Dim sNewText
	sNewText = theText

	Dim sSearchText
	sSearchText = UCase(theText)

	Dim SDefine
	Dim sRepl
	
	Dim nSearch
	
	For DI = LBound(theDefines) To UBound(theDefines)
		sDefine = theDefines(DI)(DF_DEF)
		sRepl = theDefines(DI)(DF_REPL)
		If UCase(sDefine) <> UCase(sRepl) Then
			nSearch = InStr(sSearchText,sDefine)
			While nSearch > 0
				sNewText = Left(sNewText,nSearch-1)+sRepl+Right(sNewText,Len(sSearchText)-nSearch+1-Len(sDefine))
				sSearchText = Left(sSearchText,nSearch-1)+UCase(sRepl)+Right(sSearchText,Len(sSearchText)-nSearch+1-Len(sDefine))
				nSearch = InStr(sSearchText,sDefine)
			Wend
		End If
	Next
	
	ReplaceDefines = sNewText

End Function

Sub WriteSText(theOutStream,theSText,theIndent)

	If VarType(theSText) <> 8204 Then
		Exit Sub
	End If

	With theOutStream
		For I = LBound(theSText) To UBound(theSText)
			.WriteText theIndent & theSText(I), adWriteLine
		Next
	End With
	
End Sub

Function ReplSText(theSText,theSearch,theRepl)

	Dim NewSText
	NewSText = theSText

	For I = LBound(NewSText) To UBound(NewSText)
		NewSText(I) = Replace(NewSText(I),theSearch,theRepl)
	Next

	ReplSText = NewSText
	
End Function

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
				WriteSText theOutStream,ReplSText(STextSNSearchCondSm,STV_SN,sStatName),sIndent
			Else
				WriteSText theOutStream,ReplSText(STextSNSearchCondStartSm,STV_SN,sStatName),sIndent
				bIfOpen = True
			End If
			WriteScrStatsNode theOutStream,theSDInfoLst,theDefines,aBST,iLeftIndex,sIndent+"	"
		End If
		
		Dim iRightIndex
		iRightIndex = aBST(iNodeIndex)(BST_RIGHT)
		
		If iRightIndex <> -1 Then
			If bIfOpen Then
				WriteSText theOutStream,ReplSText(STextSNSearchCondGr,STV_SN,sStatName),sIndent
			Else
				WriteSText theOutStream,ReplSText(STextSNSearchCondStartGr,STV_SN,sStatName),sIndent
				bIfOpen = True
			End If
			WriteScrStatsNode theOutStream,theSDInfoLst,theDefines,aBST,iRightIndex,sIndent+"	"
		End If

		If iLeftIndex = -1 Or iRightIndex = -1 Then
			If bIfOpen Then
				WriteSText theOutStream,ReplSText(STextSNSearchCondEq,STV_SN,sStatName),sIndent
			Else
				WriteSText theOutStream,ReplSText(STextSNSearchCondStartEq,STV_SN,sStatName),sIndent
				bIfOpen = True
			End If
			WriteScrStat theOutStream,sStatName,theSDInfoLst,theDefines,sIndent
			WriteSText theOutStream,STextSNSearchAlternative,sIndent
			WriteSText theOutStream,ReplSText(STextSNSearchResult,STV_CALC,"0"),sIndent
		ElseIf iLeftIndex <> -1 And iRightIndex <> -1 Then
			WriteSText theOutStream,STextSNSearchAlternative,sIndent
			WriteScrStat theOutStream,sStatName,theSDInfoLst,theDefines,sIndent
		End If

		If bIfOpen Then
			WriteSText theOutStream,STextSNSearchEnd,sIndent
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
						WriteSText theOutStream,ReplSText(STextSeqLSearchCondStart,STV_LVL,theSDInfoLst(SI)(SD_LEND)),sIndent
					ElseIf theSDInfoLst(SI)(SD_SCRIPT) = SCRIPT_INTERVAL Then
						WriteSText theOutStream,ReplSText(ReplSText(STextIntLSearchCondStart,STV_LSTART,theSDInfoLst(SI)(SD_LSTART),theSDInfoLst(SI)(SD_LEND))),sIndent
					End If
				ElseIf theSDInfoLst(SI)(SD_TYPE) <> "" And theSDInfoLst(SI)(SD_LEND) <> "" Then
					If theSDInfoLst(SI)(SD_SCRIPT) = SCRIPT_SEQUENTIAL Then
						WriteSText theOutStream,ReplSText(STextSeqLSearchCond,STV_LVL,theSDInfoLst(SI)(SD_LEND)),sIndent
					ElseIf theSDInfoLst(SI)(SD_SCRIPT) = SCRIPT_INTERVAL Then
						WriteSText theOutStream,ReplSText(ReplSText(STextIntLSearchCond,STV_LSTART,theSDInfoLst(SI)(SD_LSTART),theSDInfoLst(SI)(SD_LEND))),sIndent
					End If
				ElseIf theSDInfoLst(SI)(SD_TYPE) <> "" And ((theSDInfoLst(SI)(SD_SCRIPT) = SCRIPT_SEQUENTIAL And theSDInfoLst(SI)(SD_LSTART) <> "") Or (theSDInfoLst(SI)(SD_SCRIPT) = SCRIPT_INTERVAL And theSDInfoLst(SI)(SD_LSTART) = "" And theSDInfoLst(SI)(SD_LEND) = "")) Then
					If theSDInfoLst(SI)(SD_SCRIPT) = SCRIPT_SEQUENTIAL Then
						WriteSText theOutStream,STextSeqLSearchAlternative,sIndent
					ElseIf theSDInfoLst(SI)(SD_SCRIPT) = SCRIPT_INTERVAL Then
						WriteSText theOutStream,STextIntLSearchAlternative,sIndent
					End If
				End If

				Build = ""

				If theSDInfoLst(SI)(SD_TYPE) = "A" Then
					Build = theSDInfoLst(SI)(SD_P1)
				ElseIf theSDInfoLst(SI)(SD_TYPE) = "C" Then
					Build = theSDInfoLst(SI)(SD_P1)
				ElseIf theSDInfoLst(SI)(SD_TYPE) = "D" Then
					Build = "DataTableValue("+STextArrayStart+theSDInfoLst(SI)(SD_P1)
					Dim DI
					For DI = SI+1 To UBound(theSDInfoLst)
						If theSDInfoLst(DI)(SD_STAT) <> theStat Or theSDInfoLst(DI)(SD_TYPE) <> "" Then
							Exit For
						End If
						Build = Build+STextArraySep+theSDInfoLst(DI)(SD_P1)
					Next
					Build = Build+STextArrayEnd
					If theSDInfoLst(SI)(SD_LSTART) <> "" And theSDInfoLst(SI)(SD_LSTART) <> "1" Then
						Build = Build+","+STextDefineL+"-"+CStr(theSDInfoLst(SI)(SD_LSTART)-1)+")"
					Else
						Build = Build+","+STextDefineL+")"
					End If
				ElseIf theSDInfoLst(SI)(SD_TYPE) = "E" Then
					If theSDInfoLst(SI)(SD_LSTART) <> "" Then
						Build = "ExpFmod("+theSDInfoLst(SI)(SD_P1)+","+theSDInfoLst(SI)(SD_LSTART)+","+theSDInfoLst(SI)(SD_P2)+","
					Else
						Build = "ExpFmod("+theSDInfoLst(SI)(SD_P1)+",1,"+theSDInfoLst(SI)(SD_P2)+","
					End If
					IF theSDInfoLst(SI)(SD_P3) <> "" Then
						Build = Build+STextDefineL+"+"+theSDInfoLst(SI)(SD_P3)+")"
					Else
						Build = Build+STextDefineL+")"
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
						Build = "("+theSDInfoLst(SI)(SD_P1)+")*"+STextDefineL
					ElseIf theSDInfoLst(SI)(SD_P1) = "1" Then
						Build = STextDefineL
					Else
						Build = theSDInfoLst(SI)(SD_P1)+"*"+STextDefineL
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
						Build = Build+"CalcStat("""+theSDInfoLst(SI)(SD_P1)+""","+theSDInfoLst(SI)(SD_P4)+","+STextDefineN+"),"
						If InStr(theSDInfoLst(SI)(SD_P3),"/") Or InStr(theSDInfoLst(SI)(SD_P3),"+") Or InStr(theSDInfoLst(SI)(SD_P3),"-") Then
							Build = Build+"("+theSDInfoLst(SI)(SD_P3)+")*"
						ElseIf theSDInfoLst(SI)(SD_P3) <> "1" Then
							Build = Build+theSDInfoLst(SI)(SD_P3)+"*"
						End If
						Build = Build+"CalcStat("""+theSDInfoLst(SI)(SD_P1)+""","+theSDInfoLst(SI)(SD_P5)+","+STextDefineN+"),"
						Build = Build+theSDInfoLst(SI)(SD_P4)+","+theSDInfoLst(SI)(SD_P5)+","+STextDefineL
						Build = Build+STextDefinePNULL
						Build = Build+")/CalcStat("""+theSDInfoLst(SI)(SD_P1)+""","+STextDefineL+","+STextDefineN+")"
					Else
						Build = theSDInfoLst(SI)(SD_P2)
					End If
				ElseIf theSDInfoLst(SI)(SD_TYPE) = "N" Then
					If theSDInfoLst(SI)(SD_LSTART) <> "" And theSDInfoLst(SI)(SD_LSTART) <> "1" Then
						Build = "NamedRangeValue("""+theSDInfoLst(SI)(SD_P1)+""","+STextDefineL+"-"+CStr(theSDInfoLst(SI)(SD_LSTART)-1)+","+theSDInfoLst(SI)(SD_P2)+")"
					Else
						Build = "NamedRangeValue("""+theSDInfoLst(SI)(SD_P1)+""","+STextDefineL+","+theSDInfoLst(SI)(SD_P2)+")"
					End If
				ElseIf theSDInfoLst(SI)(SD_TYPE) = "P" Then
					Build = "CalcPercAB("+theSDInfoLst(SI)(SD_P1)+","+theSDInfoLst(SI)(SD_P2)+","+theSDInfoLst(SI)(SD_P3)+","+STextDefineN+")"
				ElseIf theSDInfoLst(SI)(SD_TYPE) = "R" Then
					Build = "CalcRatAB("+theSDInfoLst(SI)(SD_P1)+","+theSDInfoLst(SI)(SD_P2)+","+theSDInfoLst(SI)(SD_P3)+","+STextDefineN+")"
				ElseIf theSDInfoLst(SI)(SD_TYPE) = "T" Then
					If theSDInfoLst(SI)(SD_P6) <> "" Then
						Build = "LinFmod("+theSDInfoLst(SI)(SD_P1)+","+theSDInfoLst(SI)(SD_P2)+","+theSDInfoLst(SI)(SD_P3)+","+theSDInfoLst(SI)(SD_P4)+","+theSDInfoLst(SI)(SD_P5)+","+STextDefineL+","+theSDInfoLst(SI)(SD_P6)+")"
					Else
						Build = "LinFmod("+theSDInfoLst(SI)(SD_P1)+","+theSDInfoLst(SI)(SD_P2)+","+theSDInfoLst(SI)(SD_P3)+","+theSDInfoLst(SI)(SD_P4)+","+theSDInfoLst(SI)(SD_P5)+","+STextDefineL+STextDefinePNULL+")"
					End If
				End If

				If Build <> "" Then

					If Ucase(theSDInfoLst(SI)(SD_USEN)) <> "N" And InStr("CDEFLNT",theSDInfoLst(SI)(SD_TYPE)) Then
						If Build = "0" Then
						ElseIf Build = "1" Then
							Build = STextDefineN
						ElseIf InStr("DEFNT",theSDInfoLst(SI)(SD_TYPE)) Then
							Build = Build+"*"+STextDefineN
						ElseIf InStr(Build,"+") Or InStr(Build,"-") Or InStr(Build,"/") Then
							Build = "("+Build+")*"+STextDefineN
						Else
							Build = Build+"*"+STextDefineN
						End If
					End If

					If theSDInfoLst(SI)(SD_DECIMALS) <> "" Then
						If theSDInfoLst(SI)(SD_DECIMALS) = "0" Then
							Build = "RoundDbl("+Build+STextDefinePNULL+")"
						Else
							Build = "RoundDbl("+Build+","+theSDInfoLst(SI)(SD_DECIMALS)+")"
						End If
					End If
	
					Build = Replace(Build,"!","""")
					Build = Replace(Build,"|",",")
					Build = Replace(Build,"+-","-")
					Build = Replace(Build,"--","+")
					Build = ReplStatRefs(Build)

					If theSDInfoLst(SI)(SD_LEND) <> "" Or theSDInfoLst(SI)(SD_LSTART) <> "" Then
						If theSDInfoLst(StatPos)(SD_SCRIPT) = SCRIPT_SEQUENTIAL Then
							WriteSText theOutStream,ReplSText(STextSeqLSearchResult,STV_CALC,Build),sIndent
						ElseIf theSDInfoLst(StatPos)(SD_SCRIPT) = SCRIPT_INTERVAL Then
							WriteSText theOutStream,ReplSText(STextIntLSearchResult,STV_CALC,Build),sIndent
						End If
					Else
						WriteSText theOutStream,ReplSText(STextSNSearchResult,STV_CALC,Build),sIndent
					End If

				End If
			Next	

			If theSDInfoLst(StatPos)(SD_LSTART) <> "" Or theSDInfoLst(StatPos)(SD_LEND) <> "" Then
				If theSDInfoLst(StatPos)(SD_SCRIPT) = SCRIPT_SEQUENTIAL Then
					WriteSText theOutStream,STextSeqLSearchEnd,sIndent
				ElseIf theSDInfoLst(StatPos)(SD_SCRIPT) = SCRIPT_INTERVAL Then
					WriteSText theOutStream,STextIntLSearchEnd,sIndent
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
				Result = Left(Result,nSearch-1)+","+STextDefineL+Right(Result,Len(Result)-nSearch+1)
				nSearch = nSearch+Len(","+STextDefineL)
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
					Result = Left(Result,nSearch-1)+","+STextDefineN+")"+Right(Result,Len(Result)-nSearch+1)
					nSearch = nSearch+Len(","+STextDefineN+")")
				Else
					Result = Left(Result,nSearch-1)+","+sNumber+")"+Right(Result,Len(Result)-nSearch+1-Len(sNumber))
					nSearch = nSearch+Len(","+sNumber+")")
				End If
			Else
				Result = Left(Result,nSearch-1)+STextDefinePNULL+")"+Right(Result,Len(Result)-nSearch+1)
				nSearch = nSearch+Len(STextDefinePNULL+")")
			End If

		End If

		nSearch = InStr(Result,sStartToken)
	Wend

	ReplStatRefs = Result

End Function

Sub ReadSDFile(ByVal theInFileName, ByRef theSDInfoLst)

	Dim InStream
	Set InStream = CreateObject("ADODB.Stream")
	With InStream
		.Type = adTypeText
		.CharSet = "UTF-8"
		.LineSeparator = adLF
		.Open
		.LoadFromFile theInFileName

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
