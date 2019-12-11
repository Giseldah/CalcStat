Dim ilvl
ilvl = CInt(InputBox("Enter item level:","CalcStat Example",0))

WScript.echo "The armour value for a ilvl " & ilvl & " incomparable heavy helmet is: " & CalcStat("Armour",ilvl,"HHT")