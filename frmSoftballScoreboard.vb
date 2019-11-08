'Program Name: Softball Scoreboard
'Author: Brian Day
'Date: March 18
'Purpose: This application calculates the score of each inning during a softball game for one team. 
'        Also calculates a running total And a final total after completing seven innings.

Option Strict On
Public Class frmSoftballScoreboard

    'btnEnterScore event handling
    Private Sub btnEnterScore_Click(sender As Object, e As EventArgs) Handles btnEnterScore.Click
        '
        'Declare variables
        Dim strUserInningResult As String
        Dim intInningResult As Integer
        Dim intFinalScore As Integer = 0
        Dim strInputMessage As String = "Enter the runs earned from inning #"
        Dim strInputHeading As String = "Score"
        Dim strNormalMessage As String = "Enter the runs earned from inning #"
        Dim strNonNumericError As String = "Error - Please use whole, numeric values only"
        Dim strNegativeError As String = "Error - Please use positive values only"

        'Declare and initialize the loop incremeneter/accumulator variables
        Dim strCancelClicked As String = ""
        Dim intMaxInnings As Integer = 7
        Dim intCurrentInning As Integer = 1

        'This loop takes all seven innings inputs and ends when it reaches the last inning or when user clicks Clear or Exit
        'Sets strInputMessage accordingly
        strUserInningResult = InputBox(strInputMessage & intCurrentInning, strInputHeading, " ")

        Do Until intCurrentInning > intMaxInnings Or strUserInningResult = strCancelClicked
            If IsNumeric(strUserInningResult) Then
                intInningResult = Convert.ToInt32(strUserInningResult)
                If intInningResult > 0 Then
                    lstScore.Items.Add(intInningResult)
                    intFinalScore += intInningResult
                    lstScore.Items.Add(intFinalScore)
                    intCurrentInning += 1
                    strInputMessage = strNormalMessage
                Else
                    strInputMessage = strNegativeError
                End If
            Else
                strInputMessage = strNonNumericError
            End If

            If intCurrentInning <= intMaxInnings Then
                strUserInningResult = InputBox(strInputMessage & intCurrentInning, strInputHeading, " ")
            End If
        Loop

        'This loop calculates and displays Final Score
        If intCurrentInning > 1 Then
            lblFinalScore.Visible = True
            lblFinalScore.Text = "Final Score of the Game " & intFinalScore.ToString("G")
        Else
            MsgBox("No score value for this inning entered, please select Clear from the menu")
        End If

        'Disable the Enter Score button
        btnEnterScore.Enabled = False
    End Sub

    'mnuClear click event handling
    Private Sub mnuClear_Click(sender As Object, e As EventArgs) Handles mnuClear.Click
        'Clear ListBox object and turns visibility back off for the Final Score label.
        'Enables the Enter Score button
        lstScore.Items.Clear()
        lblFinalScore.Visible = False
        btnEnterScore.Enabled = True
    End Sub

    'mnuExit click event handling
    Private Sub mnuExit_Click(sender As Object, e As EventArgs) Handles mnuExit.Click
        'Close window and exits application
        Close()
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
End Class
