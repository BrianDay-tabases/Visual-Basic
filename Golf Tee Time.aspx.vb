'Project:   Golf Tee Time Reservation Form
'Author:    Brian Day
'Date:      April 19
'Purpose:   This form takes user reservation information. Then data validation is performed and a
'           message will display if the reservation is made successfully.

Option Strict On

Partial Class About
    Inherits Page

    'btnSubmit event handling
    Protected Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click
        'Declare variables
        Dim strName As String
        Dim strMessage As String

        'Trim procedure for String
        strName = txtName.Text.Trim

        'clear reservation confirmation label
        lblReservation.Text = ""

        'Data validation
        If cldDate.SelectedDate >= cldDate.TodaysDate Then
            lblCalendarError.Visible = False
            strMessage = "Thank you " & strName & "for booking your " & ddlTimes.SelectedItem.Text & " tee time on " & cldDate.SelectedDate.ToShortDateString()
            lblReservation.Text = strMessage
        Else
            lblCalendarError.Visible = True
        End If



    End Sub
End Class