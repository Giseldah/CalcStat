' Support functions for CalcStat. These consist of implementations of more complex calculation types, decode methods for parameter "C" and rounding/min/max/compare functions for floating point numbers.
' Created by Giseldah

' Floating point numbers bring errors into the calculation, both inside the Lotro-client and in this function collection. This is why a 100% match with the stats in Lotro is impossible.
' Anyway, to compensate for some errors, we use a calculation deviation correction value. This makes for instance 24.49999999 round to 25, as it's assumed that 24.5 was intended as outcome of a formula.
Const DblCalcDev = 0.00000001

' ****************** Calculation Type support functions ******************

' DDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDD
' DataTableValue: Takes a value from an array table.

Function DataTableValue(ByVal vDataArray, ByVal lIndex)

	Dim lI
	lI = lIndex-1
	If lI < LBound(vDataArray) Then
		lI = LBound(vDataArray)
	End If
	If lI > UBound(vDataArray) Then
		lI = UBound(vDataArray)
	End If

	DataTableValue = vDataArray(lI)

End Function

' EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
' ExpFmod: Exponential function based on percentage.
' Common percentage values are around ~5.5% for between levels and ~20% jumps between level segments.

Function ExpFmod(ByVal dVal, ByVal dLstart, ByVal dPlvl, ByVal dLvl)

	Dim Result

	If SmallerDbl(dLvl,dLstart) Then
		Result = dVal
	Else
		Result = dVal*(1+dPlvl/100)^(dLvl-dLstart+1)
	End If

	ExpFmod = Result

End Function

' NNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNN
' NamedRangeValue: Takes a value from a named spreadsheet table.
' This function doesn't have a meaning outside a spreadsheet and so is not implemented here.

Function NamedRangeValue(ByVal RName, ByVal RowIndex, ByVal ColIndex)

	NamedRangeValue = 0

End Function

' PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
' CalcPercAB: Calculates the percentage out of a rating based on the AB formula.

Function CalcPercAB(ByVal dA, ByVal dB, ByVal dPCap, ByVal dR)

	Dim Result

	If dR <= 0 Then
		Result = 0
	Else
		Result = MinDbl(dA/(1+dB/dR),dPCap)
	End If

	CalcPercAB = Result

End Function

' RRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRR
' CalcRatAB: Calculates the rating out of a percentage based on the AB formula.

Function CalcRatAB(ByVal dA, ByVal dB, ByVal dCapR, ByVal dP)

	Dim Result

	If dP <= 0 Then
		Result = 0
	Else
		Result = MinDbl(dB/(dA/dP-1),dCapR)
	End If

	CalcRatAB = Result

End Function

' TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT
' LinFmod: Linear line function between 2 points with some optional modifications.
' Connects point (dLstart,dVal*dFstart) with (dLend,dVal*dFend).
' Usually used with dVal=1 and dFstart/dFend containing unrelated points or dVal=# and dFstart/dFend containing multiplier factors.
' Modification for in-between points on the line: rounding (in general to multiples of dRound=10).

Function LinFmod(ByVal dVal, ByVal dFstart, ByVal dFend, ByVal dLstart, ByVal dLend, ByVal dLvl, ByVal dRound)

	Dim Result

	If IsDbl(dLvl,dLstart) Then
		Result = dVal*dFstart
	ElseIf IsDbl(dLvl,dLend) Then
		Result = dVal*dFend
	ElseIf VarType(dRound) = 9 Then
		Result = dVal*(dFstart*(dLend-dLvl)+(dLvl-dLstart)*dFend)/(dLend-dLstart)
	Else
		Result = RoundDbl((dVal*(dFstart*(dLend-dLvl)+(dLvl-dLstart)*dFend)/((dLend-dLstart)*dRound)),Nothing)*dRound
	End If

	LinFmod = Result

End Function

' ****************** Parameter "C" decode support functions ******************

' ArmCodeIndex: returns a specified index from an Armour Code.
' acode string:
' 1st position: H=heavy, M=medium, L=light
' 2nd position: H=head, S=shoulders, CL=cloak/back, C=chest, G=gloves, L=leggings, B=boots, Sh=shield
' 3rd position: W=white/common, Y=yellow/uncommon, P=purple/rare, T=teal/blue/incomparable, G=gold/legendary/epic
' Note: no such thing exists as a heavy, medium or light cloak, so no H/M/L in cloak codes (cloaks go automatically in the M class since U23, although historically this was L)

Function ArmCodeIndex(ByVal ACode, ByVal iI)

	Dim ArmourCode
	ArmourCode = UCase(LTrim(RTrim(ACode)))

	Dim sArmClass
	Dim sArmType
	Dim sArmCol
	sArmClass = Left(ArmourCode,1)
	sArmType = Mid(ArmourCode,2,1)
	sArmCol = Mid(ArmourCode,3,1)
	If sArmType = "S" AND sArmCol = "H" Then
		sArmType = "SH"
		sArmCol = Mid(ArmourCode,4,1)
	ElseIf sArmClass = "C" AND sArmType = "L" Then
		sArmClass = "M"
		sArmType = "CL"
		sArmCol = Mid(ArmourCode,3,1)
	Else
		sArmType = " "+sArmType
	End If
	
	Select Case iI
		Case 1
			ArmCodeIndex = InStr("HML",sArmClass)
		Case 2
			ArmCodeIndex = (InStr(" H SCL C G L BSH",sArmType)+1)/2
		Case 3
			ArmCodeIndex = InStr("WYPTG",sArmCol)
		Case Else
			ArmCodeIndex = 0
	End Select
	
End Function

' RomanRankDecode: converts a string with a Roman number in characters, to an integer number.
' used for Legendary Item Title calculation.

Function RomanRankDecode(ByVal number)

	Dim sLetters
	sLetters = Array("M","CM","D","CD","C","XC","L","XL","X","IX","V","IV","I")
	Dim iValues
	iValues = Array(1000,900,500,400,100,90,50,40,10,9,5,4,1)
	Dim Result
	Result = 0

	If number <> "" Then
		For I = LBound(sLetters) To UBound(sLetters)
			If Left(UCase(number),Len(sLetters(I))) = sLetters(I) Then
				Result = iValues(I)+RomanRankDecode(Right(number,Len(number)-Len(sLetters(I))))
				Exit For
			End If
		Next
	End If

	RomanRankDecode = Result

End Function

' ****************** Misc. floating point support functions ******************

' Misc. functions for floating point: rounding etc.
' For roundings: iDec is number of decimals.

Function RoundDbl(ByVal dNum, ByVal iDec)

	If VarType(iDec) = 9 Then
		iDec = 0
	End If
	If iDec = 0 Then
		RoundDbl = INT(dNum+0.5+DblCalcDev)
	Else
		RoundDbl = INT(dNum*10^iDec+0.5+DblCalcDev)/10^iDec
	End If
	
End Function

Function CeilDbl(ByVal dNum, ByVal iDec)

	If VarType(iDec) = 9 Then
		iDec = 0
	End If
	If iDec = 0 Then
		CeilDbl = INT(dNum+1-DblCalcDev)
	Else
		CeilDbl = INT(dNum*10^iDec+1-DblCalcDev)/10^iDec
	End If
	
End Function

Function FloorDbl(ByVal dNum, ByVal iDec)

	If VarType(iDec) = 9 Then
		iDec = 0
	End If
	If iDec = 0 Then
		FloorDbl = INT(dNum+DblCalcDev)
	Else
		FloorDbl = INT(dNum*10^iDec+DblCalcDev)/10^iDec
	End If

End Function

Function IsDbl(ByVal dNum1, ByVal dNum2)

	IsDbl = (ABS(dNum1-dNum2) <= DblCalcDev)

End Function

Function SmallerDbl(ByVal dNum1, ByVal dNum2)

	SmallerDbl = ((dNum2-dNum1) > DblCalcDev)

End Function

Function GreaterDbl(ByVal dNum1, ByVal dNum2)

	GreaterDbl = ((dNum1-dNum2) > DblCalcDev)

End Function

Function MinDbl(ByVal dNum1, ByVal dNum2)

	If dNum1 <= dNum2 Then
		MinDbl = dNum1
	Else
		MinDbl = dNum2
	End If

End Function

Function MaxDbl(ByVal dNum1, ByVal dNum2)

	If dNum1 >= dNum2 Then
		MaxDbl = dNum1
	Else
		MaxDbl = dNum2
	End If

End Function
