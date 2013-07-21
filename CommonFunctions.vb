'MacroName: CommonFunctions
'MacroDescription: A macrobook of common functions using in macros
' USAGE:  '$Include: "<macroBk>!CommonFunctions"

'=======================
' Function declarations
'=======================
Declare Function ReplaceAllInstances(sStr$, sOldStr$, sNewStr$) As String
Declare Function RemoveAllInstances(sStr$, sStrToRemove$) As String

'----------------------
' REMOVE ALL INSTANCES
'----------------------
Function RemoveAllInstances(sStr$, sStrToRemove$)
  Dim nIndex as Integer
 
  nIndex = InStr(sStr$, sStrToRemove$)
  Do While nIndex > 0
     sStr$ = Left(sStr$, nIndex - 1) & Mid(sStr$, nIndex + Len(sStrToRemove$))
     nIndex = InStr(sStr$, sStrToRemove$)
  Loop
 
  RemoveAllInstances = sStr$
End Function


'---------------------------
' REPLACE ALL INSTANCES
'---------------------------
Function ReplaceAllInstances(sStr$, sOldStr$, sNewStr$)
  Dim nIndex as Integer
 
  nIndex = InStr(sStr$, sOldStr$)
  Do While nIndex > 0
     sStr$ = Left(sStr$, nIndex - 1) & sNewStr$ & Mid(sStr$, nIndex + Len(sOldStr$))
     nIndex = InStr(sStr$, sOldStr$)
  Loop
 
  ReplaceAllInstances = sStr$
End Function