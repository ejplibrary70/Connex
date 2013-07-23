'MacroName:AddRelatorTerms
'MacroDescription:Add relator terms

'$Include "Public!ArrayFunctions"

Declare Function FindEndMainString( TargetHeading$ ) As Integer

Sub Main

   Dim CS As Object
   Set CS = CreateObject("Connex.Client")
   Dim nBool%, nRow%
   Dim sHdg$, sLab$, sTag$, sHdgNoSF$, endSF$, x%
   
   ' Get field

   nRow = CS.CursorRow
   ' Check that it is a bibliographic record
      If CS.GetFieldLine (nRow, sHdg) = True Then
         sTag = Mid$(sHdg, 1, 3)
         sHdg = Mid$(sHdg, 6)
         If InStr("100,110,111,130,700,710,711,730,800,810,811,830", sTag) <> 0 Then
            x = FindEndMainString( sHdg$ )
            If x <> 0 Then 
               endSF$ = Mid$( sHdg$, x )
               sHdgNoSF$ = Left$( sHdg$, x-2 )
            Else 
               endSF$ = ""
               sHdgNoSF$ = sHdg
            End If

         Else
            Msgbox "Place the cursor in 100, 110, 111, 130, 700, 710, 711, 730, 800, 810, 811, or 830 to run relator term macro", 0, "Relator Term Macro"
         End If
      Else
         Msgbox "Place the cursor in 100, 110, 111, 130, 700, 710, 711, 730, 800, 810, 811, or 830 to run relator term macro", 0, "Relator Term Macro"
      End If


   msgBox sHdgNoSF$ & "END and x=" & x 
   msgBox "BEG" & endSF$ & "END"

   ' TODO Convert delimiter 4 to e
   
   ' TODO Convert delimiter e with abbreviations to RDA relator
   
   ' TODO Display dialog box with current heading
   
   ' TODO Allow additional relators to be added from pull down
   

End Sub


Function FindEndMainString( TargetHeading$ ) As Integer

' If the current heading ends in subfields $e, $4, or $5, this function marks the
' beginning of those subfields so they can be copied and added back on to the new heading

Dim xE As Integer, x4 As Integer, x5 As Integer
Dim Lowest%
Dim indices%(3)

indices(1) = InStr( TargetHeading$, Chr$( 223 ) & "e" )
indices(2) = InStr( TargetHeading$, Chr$( 223 ) & "4" )
indices(3) = InStr( TargetHeading$, Chr$( 223 ) & "5" )

' Find the first occurrence of these subfields so all characters following that point are
' copied and retained

Lowest% = NonZeroLowerBound(indices)
    
FindEndMainString = Lowest%

End Function




