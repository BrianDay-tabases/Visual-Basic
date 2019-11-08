'Program: Campus Cafe
'Author: Brian Day
'Date: Feb 3
'Purpose: Display a welcome screen, and details option to all college students about the Open Mic Event


Public Class frmCampusCafe
    Private Sub frmCampusCafe_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub lblDetails_Click(sender As Object, e As EventArgs) Handles lblDetails.Click

    End Sub

    'USER SELECTS TO VIEW DETAILS BUTTON EVENT
    Private Sub btnDetails_Click(sender As Object, e As EventArgs) Handles btnDetails.Click
        'Set lblDetails to Visible
        'Enable btnExit
        'Set btnExit to Visible
        lblDetails.Visible = True
        btnExit.Enabled = True
        btnExit.Visible = True
    End Sub

    'USER SELECTS THE EXIT BUTTON EVENT
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Closes the program
        Close()
    End Sub
End Class
