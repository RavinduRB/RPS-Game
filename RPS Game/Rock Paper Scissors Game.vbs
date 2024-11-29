' Rock Paper Scissors Game in VBScript
Option Explicit

' Declare variables
Dim myDict, userChoice, compChoice, userCount, compCount, gamesPlayed, totalGames, result, isValidChoice

' Initialize the dictionary for choices
Set myDict = CreateObject("Scripting.Dictionary")
myDict.Add "R", "Rock"
myDict.Add "P", "Paper"
myDict.Add "S", "Scissors"

' Initialize scores
userCount = 0
compCount = 0
gamesPlayed = 0

' Prompt user for the total number of games
totalGames = InputBox("Enter the number of rounds:", "Rock Paper Scissors Game")
If Not IsNumeric(totalGames) Or totalGames <= 0 Then
    MsgBox "Invalid input. Please enter a positive number.", vbExclamation, "Error"
    WScript.Quit
End If

totalGames = CInt(totalGames)

' Main game loop
Do While gamesPlayed < totalGames
    ' Prompt user for choice
    userChoice = InputBox("Enter your choice (R for Rock, P for Paper, S for Scissors):", "Your Turn")
    userChoice = UCase(Trim(userChoice))

    ' Validate input
    isValidChoice = myDict.Exists(userChoice)
    If Not isValidChoice Then
        MsgBox "Invalid choice. Please enter R, P, or S.", vbExclamation, "Error"
    Else
        ' Generate computer choice
        Randomize
        Dim compChoices
        compChoices = Array("R", "P", "S")
        compChoice = compChoices(Int(3 * Rnd()))

        ' Display choices
        MsgBox "You chose: " & myDict(userChoice) & vbCrLf & "Computer chose: " & myDict(compChoice), vbInformation, "Choices"

        ' Determine the winner
        If (userChoice = "R" And compChoice = "P") Or (userChoice = "P" And compChoice = "S") Or (userChoice = "S" And compChoice = "R") Then
            compCount = compCount + 1
            result = "Computer Wins!"
        ElseIf (userChoice = "P" And compChoice = "R") Or (userChoice = "S" And compChoice = "P") Or (userChoice = "R" And compChoice = "S") Then
            userCount = userCount + 1
            result = "You Win!"
        Else
            result = "It's a Tie!"
        End If

        MsgBox result, vbInformation, "Round Result"

        ' Update round count
        gamesPlayed = gamesPlayed + 1
    End If
Loop

' Display final results
If userCount > compCount Then
    MsgBox "Congratulations! You won the game!" & vbCrLf & "Final Score: You " & userCount & " - Computer " & compCount, vbInformation, "Game Over"
ElseIf userCount < compCount Then
    MsgBox "Sorry, you lost the game!" & vbCrLf & "Final Score: You " & userCount & " - Computer " & compCount, vbInformation, "Game Over"
Else
    MsgBox "It's a tie!" & vbCrLf & "Final Score: You " & userCount & " - Computer " & compCount, vbInformation, "Game Over"
End If

' Clean up
Set myDict = Nothing
