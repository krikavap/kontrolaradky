Sub vlozeniPevnychMezer()
'
' tvrde_mezery Macro
' zkontroluje osamocené předložky na konci řádku a vyřeší to vložením tvrdých mezer
'

    Application.ScreenUpdating = False
    Dim AscA, AscE, AscI, AscK, AscO, AscS, AscU, AscV, AscZ, AscSpace As Integer
    AscA = Asc("A")
    AscE = Asc("E")
    AscI = Asc("I")
    AscK = Asc("K")
    AscO = Asc("O")
    AscS = Asc("S")
    AscU = Asc("U")
    AscV = Asc("V")
    AscZ = Asc("Z")
    AscSpace = Asc(" ")

    Dim doc As Document
    Set doc = ActiveDocument
    
    Dim prevWord As Range
    Set prevWord = doc.Range.Words.First
    Dim prevLine As Integer
    prevLine = prevWord.Information(wdFirstCharacterLineNumber)
    
    Dim currWord As Range
    For Each currWord In doc.Range.Words
        Dim currLine As Integer
        currLine = currWord.Information(wdFirstCharacterLineNumber)
        If (prevLine <> currLine) Then
            Dim prevWordChars As Characters
            Set prevWordChars = prevWord.Characters
            If (prevWordChars.Count = 2) Then
                If (Asc(prevWordChars.Last) = AscSpace) Then
                    Select Case (Asc(UCase(prevWordChars.First.Text)))
                        Case AscA, AscE, AscI, AscK, AscO, AscS, AscU, AscV, AscZ
                            prevWordChars.Last.Text = Chr$(160)
                    End Select
                End If
            End If
        End If
        prevLine = currLine
        Set prevWord = currWord
    Next
    
    Application.ScreenUpdating = True
End Sub
