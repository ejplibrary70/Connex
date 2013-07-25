'MacroName:getRelatorTerms
'MacroDescription:Get terms from files

'  *** MODIFY FileBase directory to where datafiles are ***

' Written by E.J. Petersen
' Email: ejplibrary70@gmail.com
' Version: 1.0 (24 Jul 2013)

' Use at your own risk
' E.J. Petersen shall not be liable for any loss or damage, lost profits, loss of business, loss of or
' damage to data, downtime or unavailability, of or in connection with use of materials. E.J. Petersen
' shall have no liability for any claims arising from use of the materials, based on
' infringement of copyright, patent, trade secret or other right, libel, slander or invasion of
' privacy or claims based on errors, inaccuracies or omissions in or loss of the data.

' E.J. Petersen makes no express warranties or representations and disclaims all implied warranties with
' respect to materials as to their accuracy, merchantability or fitness for a particular purpose.
' Macros are supplied "as is."

Declare Sub LoadTerms
Declare Sub loadFile ( xArray$(), xArrayN$(), fileName$ )

'====================
'Routine to retrieve relator terms and populate arrays for dialog box.
'  from tab-delimited file
'  Format of file:
'     1st line:  number of records
'     rest of file: relator term /tab/ notes on usage
'====================

'=====================
' Global variable declaration
'=====================

    Dim dropboxArray0$(), dropboxArray1$(), dropboxArray2$(), dropboxArray3$()
    Dim dropboxArray4$(), dropboxArray5$(), dropboxArray6$(), dropboxArray7$()
    Dim dropboxArrayN0$(), dropboxArrayN1$(), dropboxArrayN2$(), dropboxArrayN3$()
    Dim dropboxArrayN4$(), dropboxArrayN5$(), dropboxArrayN6$(), dropboxArrayN7$()
    Dim typeDB$(7)
    
'====================
' DATA FILE LOCATIONS
'====================
   Const FileBase = "F:\Relators\"


Sub LoadTerms()

   Dim CS As Object
   Set CS = CreateObject("Connex.Client")
   
   Dim fileLine As String
   
   ' Name of files with data.  Assumes ".txt" suffix
   typeDB$(0) = "I_2_1_Creators"
   typeDB$(1) = "I_2_2_Other"
   typeDB$(2) = "I_3_1_Contributors"
   typeDB$(3) = "I_4_1_Manufacturers"
   typeDB$(4) = "I_4_2_Publishers"
   typeDB$(5) = "I_4_3_Distributors"
   typeDB$(6) = "I_5_1_Owners"
   typeDB$(7) = "I_5_2_Item"
   
   
' dropboxArray0
   loadFile dropboxArray0$, dropboxArrayN0$, typeDB$(0)
   
' dropboxArray1
   loadFile dropboxArray1$, dropboxArrayN1$, typeDB$(1)
  
' dropboxArray2
   loadFile dropboxArray2$, dropboxArrayN2$, typeDB$(2)
   
' dropboxArray3
   loadFile dropboxArray3$, dropboxArrayN3$, typeDB$(3)
   
' dropboxArray4
   loadFile dropboxArray4$, dropboxArrayN4$, typeDB$(4)
   
' dropboxArray5
   loadFile dropboxArray5$, dropboxArrayN5$, typeDB$(5)

' dropboxArray6
   loadFile dropboxArray6$, dropboxArrayN6$, typeDB$(6)
   
' dropboxArray7
   loadFile dropboxArray7$, dropboxArrayN7$, typeDB$(7)

End Sub


Sub loadFile ( xArray$(), xArrayN$(), fileName$ )

   Dim fileLine$, file$

   file = FileBase & fileName & ".txt"
   recordno = 0
   Open file For input As #1
   
   ' read 1st line to detemine how many codes are in file
   Line Input #1, fileLine
   numRec = Val( fileLine )
   
   ' define dimensions of arrays
   Redim xArray( numRec ) 
   Redim xArrayN( numRec )

   Do While Not Eof(1)
      Line Input #1, fileLine
      xArray( recordno ) = GetField( fileLine, 1, Chr$(9) )
      xArrayN( recordno ) = GetField( fileLine, 2, Chr$(9) )
      recordno = recordno + 1
   Loop
   Close #1  

End Sub
