Attribute VB_Name = "Pi"
Option Explicit

Const N As Long = 10000

Public Sub Pi()
    Dim pi_&(), reminders&(), heldDigits&, carriedOver&, sum&, length&, i&, j&, k&, q&, outputString$
    
    length = Int(N * 10 / 3)
    ReDim reminders(length)
    ReDim pi_(N - 1)
    
    If Len(Dir(Application.ActiveWorkbook.Path & "\Pi_" & N & ".txt")) <> 0 Then Kill Application.ActiveWorkbook.Path & "\Pi_" & N & ".txt"
    Dim iFile As Integer: iFile = FreeFile
    Open Application.ActiveWorkbook.Path & "\Pi_" & N & ".txt" For Append As #iFile
    
    For i = 0 To length
        reminders(i) = 2
    Next
    
    For i = 0 To N - 1
        carriedOver = 0
        For j = length To 0 Step -1
            reminders(j) = reminders(j) * 10
            sum = reminders(j) + carriedOver
            carriedOver = Int(sum / (j * 2 + 1)) * j
            reminders(j) = sum Mod (j * 2 + 1)
        Next j
        
        reminders(0) = sum Mod 10
        q = Int(sum / 10)
        
        If q = 9 Then
            heldDigits = heldDigits + 1
        ElseIf q = 10 Then
            q = 0
            For k = 1 To heldDigits
                If pi_(i - k) = 9 Then
                    pi_(i - k) = 0
                Else
                    pi_(i - k) = pi_(i - k) + 1
                End If
            Next k
            heldDigits = 1
        Else
            heldDigits = 1
        End If
        pi_(i) = q
    Next i
    
    For i = 0 To N - 1
        If i = 0 Then
            outputString = outputString & pi_(i) & "."
        Else
            outputString = outputString & pi_(i)
        End If
    Next i
    
    Print #iFile, outputString
    Close #iFile
End Sub

Sub TestCase_1000()
    Dim Pi_1000_Wolfram As String
    Dim iFile As Integer: iFile = FreeFile
    Open Application.ActiveWorkbook.Path & "\Pi_1000_Wolfram.txt" For Input As #iFile
    Pi_1000_Wolfram = Input(LOF(iFile), iFile)
    Close #iFile
    
    Dim Pi_1000 As String
    iFile = FreeFile
    Open Application.ActiveWorkbook.Path & "\Pi_1000.txt" For Input As #iFile
    Pi_1000 = Input(LOF(iFile), iFile)
    Close #iFile
    
    Dim i&
    
    For i = 1 To Len(Pi_1000_Wolfram)
        If Mid(Pi_1000, i, 1) <> Mid(Pi_1000_Wolfram, i, 1) Then
            Debug.Print "Error in " & (i - 1) & "th digit: " & Mid(Pi_1000, i, 1) & " is not equal to " & Mid(Pi_1000_Wolfram, i, 1)
        End If
    Next i
End Sub

Sub TestCase_10000()
    Dim Pi_10000_MathTOOLS As String
    Dim iFile As Integer: iFile = FreeFile
    Open Application.ActiveWorkbook.Path & "\Pi_10000_MathTOOLS.txt" For Input As #iFile
    Pi_10000_MathTOOLS = Input(LOF(iFile), iFile)
    Close #iFile
    
    Dim Pi_10000 As String
    iFile = FreeFile
    Open Application.ActiveWorkbook.Path & "\Pi_10000.txt" For Input As #iFile
    Pi_10000 = Input(LOF(iFile), iFile)
    Close #iFile
    
    Dim i&
    
    For i = 1 To Len(Pi_10000_MathTOOLS)
        If Mid(Pi_10000, i, 1) <> Mid(Pi_10000_MathTOOLS, i, 1) Then
            Debug.Print "Error in " & (i - 1) & "th digit: " & Mid(Pi_10000, i, 1) & " is not equal to " & Mid(Pi_10000_MathTOOLS, i, 1)
        End If
    Next i
End Sub
