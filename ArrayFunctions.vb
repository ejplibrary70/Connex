'MacroName:ArrayFunctions
'MacroDescription:Array Functions

Declare Function NonZeroLowerBound( arrayvar%() ) As Integer

' Returns zero or the lowest non-zero value of an Integer array
Function NonZeroLowerBound( arrayvar%() ) As Integer

    Dim lowerBound%, count%, i%

    i=LBound(arrayvar,1)
    count=UBound(arrayvar,1)
   
    lowerBound% = 0

    Do While i <= count
        If lowerBound = 0 Then
            lowerBound = arrayvar(i)
        ElseIF (lowerBound > arrayvar(i)) Then
            If arrayvar(i) <> 0 Then
                lowerBound = arrayvar(i)
            End If
        End If
        i = i+1
    Loop
   
    NonZeroLowerBound = lowerBound%
End Function