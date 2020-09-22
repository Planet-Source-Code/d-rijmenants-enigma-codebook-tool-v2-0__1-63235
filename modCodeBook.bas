Attribute VB_Name = "modCodeBook"
Option Explicit

Public Function GetRefAndFourth(M4 As Boolean, Compat As Boolean) As String
If Not M4 Then
    '3 rotors
    If Int((2 * Rnd) + 1) = 1 Then
        GetRefAndFourth = "B |"
        Else
        GetRefAndFourth = "C |"
        End If
    Else
    '4 rotors
    If Not Compat Then
        'not compat
        If Int((2 * Rnd) + 1) = 1 Then
            GetRefAndFourth = "B |"
            Else
            GetRefAndFourth = "C |"
            End If
        If Int((2 * Rnd) + 1) = 1 Then
            GetRefAndFourth = GetRefAndFourth & " Beta  "
            Else
            GetRefAndFourth = GetRefAndFourth & " Gamma "
            End If
        Else
        'compat with 3-rotor
        If Int((2 * Rnd) + 1) = 1 Then
            GetRefAndFourth = "B | Beta  "
            Else
            GetRefAndFourth = "C | Gamma "
            End If
        
        End If
    End If
End Function

Public Function GetRotors(ByVal Choice As Integer) As String
'select normal rotors
Dim k As Integer
Dim tmp As String
For k = 1 To 3
    Do
        tmp = ConvRot(Int((Choice * Rnd) + 1))
    Loop While InStr(1, GetRotors, tmp) <> 0
    GetRotors = GetRotors & tmp
Next k
End Function

Public Function GetRings(M4 As Boolean, Compat As Boolean) As String
'select rings
Dim nrRings As Integer
Dim k As Integer
Dim v As Integer
Dim rn As String

If Not M4 Then
    '3-rotor rings
    nrRings = 3
ElseIf M4 = True And Compat = True Then
    '4 rotor compatible (ring 4th rotor always A)
    nrRings = 3
    GetRings = "01 "
Else
    '4 rings
    nrRings = 4
End If
For k = 1 To nrRings
    rn = Trim(Str(Int((26 * Rnd) + 1)))
    If Len(rn) = 1 Then rn = "0" & rn
    GetRings = GetRings & rn & " "
Next k
End Function

Public Function GetNewPair(ByVal aList As String) As String
'find new rnd pair that isn't used
Dim p1 As String
Dim p2 As String
Do
    p1 = Chr(Int((26 * Rnd) + 1) + 64)
    DoEvents
Loop While InStr(1, aList, p1) <> 0
Do
    p2 = Chr(Int((26 * Rnd) + 1) + 64)
    DoEvents
Loop While InStr(1, aList, p2) <> 0 Or p1 = p2
If Asc(p1) > Asc(p2) Then
    GetNewPair = p2 & p1 & " "
    Else
    GetNewPair = p1 & p2 & " "
    End If
End Function


Public Function GetKennGruppen()
Dim k As Integer
Dim i As Integer
For k = 1 To 4
    For i = 1 To 3
    GetKennGruppen = GetKennGruppen & Chr(Int((26 * Rnd) + 1) + 64)
    Next
    GetKennGruppen = GetKennGruppen & " "
Next
End Function


Public Sub InsertPair(ByVal aPair As String, aList As String)
'insert pair in list in alphabetic order
Dim pos As Integer
If aList = "" Then aList = aPair: Exit Sub
pos = 1
Do
    If Asc(Mid(aList, pos, 1)) > Asc(Left(aPair, 1)) Then Exit Do
pos = pos + 3
Loop While pos < Len(aList)
If pos = 1 Then
    aList = aPair & aList
ElseIf pos > Len(aList) Then
    aList = aList & aPair
Else
    aList = Left(aList, pos - 1) & aPair & Mid(aList, pos)
End If
End Sub
