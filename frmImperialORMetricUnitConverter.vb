'Program: Building Assistant
'Author: Brian Day
'Date: 2-21
'Purpose: Take users text input and radio selection, and do a conversion to Imperial units or Metric Units accordingly

Option Strict On

Public Class frmUnitConverter
    'Declare Constants
    Const _cdecToInches As Decimal = 39.37008D
    Const _cdecToMeters As Decimal = 0.0254D


    'Form load event
    Private Sub frmUnitConverter_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Set Focus on user input textbox
        'Clear placeholder values from the equivalent results textbox
        'radToMetric is Selected by default
        txtEquals.Clear()
        txtUserInput.Focus()
        radToMetric.Select()
    End Sub

    'Calculate Button event handling
    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click
        'Declare Variables
        Dim decInputValue As Decimal
        Dim decOutputValue As Decimal

        If IsNumeric(txtUserInput.Text) Then
            'Program ensures user input is acceptable numeric; and if so, change its data type to decimal
            decInputValue = Convert.ToDecimal(txtUserInput.Text)
            'Program ensures user input is non-negative
            If decInputValue > 0 Then
                'Calculate the desired value using radToMetric/radToImperial value to decide (radToMetric by default, so will be ELSE portion)
                'Output final conversion value (three decimals)
                If radToImperial.Checked Then
                    decOutputValue = (decInputValue * _cdecToInches)
                    txtEquals.Text = (decOutputValue.ToString("F3") & " in")
                Else
                    decOutputValue = (decInputValue * _cdecToMeters)
                    txtEquals.Text = (decOutputValue.ToString("F3") & " m")
                End If
            Else
                'Display appropriate error message (non-negative)
                MsgBox("You entered " & decInputValue.ToString() & ". Enter a positive value.", MsgBoxStyle.Exclamation, "Input Error")
                txtUserInput.Clear()
                txtUserInput.Focus()
            End If
        Else
            'Display appropriate error message (user input is NULL or an unaceptable numeric)
            MsgBox("Please only use numerical values", MsgBoxStyle.Exclamation, "Input Error")
            txtUserInput.Clear()
            txtUserInput.Focus()
        End If

    End Sub



    'Clear button event handling
    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        'Clear user input textbox values, and set as focus again
        'Clear output textbox values
        'Resets radio selection to ToMetric by default
        txtUserInput.Clear()
        txtUserInput.Focus()
        txtEquals.Clear()
        radToMetric.Checked = True
        radToImperial.Checked = False
    End Sub

    'Exit Button event handling
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Closes program
        Close()
    End Sub
End Class
