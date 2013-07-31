'MacroName:AddRelatorTerms
'MacroDescription:Add relator terms by a dialog box

' Written by E.J. Petersen
' Email: ejplibrary70@gmail.com
' Version: 0.3 (31 Jul 2013)
' In active development

' To begin, place the cursor in a heading and run the macro.

' Use at your own risk
' E.J. Petersen shall not be liable for any loss or damage, lost profits, loss of business, loss of or
' damage to data, downtime or unavailability, of or in connection with use of materials. E.J. Petersen
' shall have no liability for any claims arising from use of the materials, based on
' infringement of copyright, patent, trade secret or other right, libel, slander or invasion of
' privacy or claims based on errors, inaccuracies or omissions in or loss of the data.

' E.J. Petersen makes no express warranties or representations and disclaims all implied warranties with
' respect to materials as to their accuracy, merchantability or fitness for a particular purpose.
' Macros are supplied "as is."


' *** MODIFY for local set-up ***
'$Include "RDA!ArrayFunctions"
'         for NonZeroLowerBound function
'$Include "RDA!getRelatorTerms"
'         for loadTerms procedure

Declare Function FindEndMainString( TargetHeading$ ) As Integer
Declare Function RTDlgFunction( identifier$, action, suppvalue )
Declare Sub UpdateListBoxes ( xArray$(), xArrayN$() )
Declare Function StripEndingPunct ( str$ )

Dim sHdgBase, endSF$, sNewHdg$

Sub Main

   Dim CS As Object
   Set CS = CreateObject("Connex.Client")
   Dim nBool%, nRow%
   Dim sHdg$, sLab$, sTag$, sInd$, x%, bool%
   
   On Error Resume Next
   
   ' Get field

   nRow = CS.CursorRow
   ' Check that it is a bibliographic record
      If CS.GetFieldLine (nRow, sHdg) = True Then
         sTag = Mid$(sHdg, 1, 3)
         sInd = Mid$(sHdg, 4, 2)  ' Used to reconstruct heading at end
         sHdg = Mid$(sHdg, 6)
         If InStr("100,110,111,130,700,710,711,800,810,811", sTag) <> 0 Then
            x = FindEndMainString( sHdg$ )   ' separate existing ending subfields
            If x <> 0 Then 
               endSF$ = Mid$( sHdg$, x )
               sHdgBase = Left$( sHdg$, x-2 )
            Else 
               endSF$ = ""
               sHdgBase = sHdg
            End If
         Else
            Msgbox "Place the cursor in 100, 110, 111, 130, 700, 710, 711, 800, 810, or 811 to run relator term macro", 0, "Relator Term Macro"
            Goto QuitSub
         End If
      Else
         Msgbox "Place the cursor in 100, 110, 111, 130, 700, 710, 711, 800, 810, or 811 to run relator term macro", 0, "Relator Term Macro"
         Goto QuitSub
      End If
      
   ' Strip ending off of base heading
   sHdgBase =  StripEndingPunct ( sHdgBase )

   ' TODO Convert delimiter 4 to e
   
   ' TODO Convert delimiter e with abbreviations to RDA relator
   
   ' Create the arrays for the dialog box
   loadTerms
   
   ' Dialog box definition
   Begin Dialog newdlg 251, 291, "Add Relator Terms", .RTDlgFunction
      OkButton  194, 249, 50, 15
      CancelButton  195, 272, 50, 15
      DropListBox  99, 44, 108, 211, "", .TheDropList
      ListBox  3, 47, 86, 81, typeDB(), .TheTypes
      TextBox  3, 3, 245, 31, .Hdg
      Text  14, 145, 85, 73, "Notes go here", .Note
      Text  1, 132, 87, 12, "Notes on term usage:"
      Text  0, 228, 127, 10, "Previous end subfields:"
      Text  11, 246, 139, 16, endSF, .endSFDisplay
      PushButton  211, 45, 35, 15, "&Add", .AddButton
   End Dialog

   Dim dlg as newdlg
   Dim response as Integer
   response= Dialog(dlg)
   Select Case response
      Case -1            ' OK
         'msgBox "OK and newHeading is " & dlg.Hdg
         repHdg = sTag & sInd & dlg.Hdg  & "."        'Add a period at the end
         'msgBox repHdg
         bool = CS.SetFieldLine (nRow, repHdg)
      Case 0             ' Cancel
         msgBox "You cancelled.  Heading not updated."
   End Select
   
QuitSub:

End Sub


Function FindEndMainString( TargetHeading$ ) As Integer

' If the current heading ends in subfields $e, $4, or $5, this function marks the
' beginning of those subfields so they can be copied and added back on to the new heading

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


Function RTDlgFunction(identifier$, action, suppvalue)

Dim newDBList$, newHdgAdd$
Dim dlmE as String

dlmE = Chr(223) & "e "

   Select Case action
     Case 1                        'dialog box initialized
       DlgValue "TheTypes", 0
       DlgListBoxArray "TheDropList", dropboxArray0()
       DlgValue "TheDropList", 0
       DlgText "Hdg", sHdgBase
       DlgText "Note", dropboxArrayN0( 0 )

     Case 2                        'button or control value changed
       Select Case identifier$ 
         Case "TheTypes"
          Select Case DlgValue("TheTypes")
            Case 0
               UpdateListBoxes dropboxArray0$(), dropboxArrayN0$()
               DlgFocus ("TheDropList")
            Case 1
               UpdateListBoxes dropboxArray1$(), dropboxArrayN1$()
               DlgFocus ("TheDropList")
            Case 2
               UpdateListBoxes dropboxArray2$(), dropboxArrayN2$()
               DlgFocus ("TheDropList")
            Case 3
               UpdateListBoxes dropboxArray3$(), dropboxArrayN3$()
               DlgFocus ("TheDropList")
            Case 4
               UpdateListBoxes dropboxArray4$(), dropboxArrayN4$()
               DlgFocus ("TheDropList")
            Case 5
               UpdateListBoxes dropboxArray5$(), dropboxArrayN5$()
               DlgFocus ("TheDropList")
            Case 6
               UpdateListBoxes dropboxArray6$(), dropboxArrayN6$()
               DlgFocus ("TheDropList")
            Case 7
               UpdateListBoxes dropboxArray7$(), dropboxArrayN7$()
               DlgFocus ("TheDropList")
           End Select
         Case "TheDropList" 
          Select Case DlgValue("TheTypes")
            Case 0
               DlgText "Note", dropboxArrayN0( DlgValue("TheDropList") )
            Case 1
               DlgText "Note", dropboxArrayN1( DlgValue("TheDropList") )
            Case 2
               DlgText "Note", dropboxArrayN2( DlgValue("TheDropList") )
            Case 3
               DlgText "Note", dropboxArrayN3( DlgValue("TheDropList") )
            Case 4
               DlgText "Note", dropboxArrayN4( DlgValue("TheDropList") )
            Case 5
               DlgText "Note", dropboxArrayN5( DlgValue("TheDropList") )
            Case 6
               DlgText "Note", dropboxArrayN6( DlgValue("TheDropList") )
            Case 7
               DlgText "Note", dropboxArrayN7( DlgValue("TheDropList") )
          End Select
         Case "AddButton"
           ' Check if hyphen at the end-- no comma
           newHdgAdd = DlgText("Hdg")
           If InStr("-", Mid$(newHdgAdd, len(newHdgAdd)) ) Then
              DlgText "Hdg", newHdgAdd & " " & dlmE & DlgText ("TheDropList")
           Else
              DlgText "Hdg", newHdgAdd & ", " & dlmE & DlgText ("TheDropList")
           End If
           
           DlgFocus ("TheDropList")
           RTDlgFunction = 1
       End Select
     Case 3                        'text or combo box changed
        ' Do nothing
     Case 4                        'control focus changed
        ' Do nothing
      
     Case 5                        'idle
  End Select
End Function

' Updates the dropboxList of attributes when a new type is selected
Sub UpdateListBoxes ( xArray$(), xArrayN$() )
    DlgListBoxArray "TheDropList", xArray()
    DlgValue "TheDropList", 0
    DlgText "Note", xArrayN( 0 )
End Sub

Function StripEndingPunct ( str$ )
   Dim newStr$
   
   newStr = Trim( str )
   If InStr(".,", Mid$(newStr, len(newStr)) ) Then
      newStr = Mid$(newStr, 1, len(newStr) - 1 )
   End If

   StripEndingPunct = newStr$
End Function

