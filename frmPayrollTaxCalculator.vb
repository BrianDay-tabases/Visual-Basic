'Program: Payroll Calculator
'Author: Brian Day
'Date: 2-10
'Purpose: Generate the appropriate tax amounts based upon the users input(a gross sum of biweekly pay)

Option Strict On

Public Class frmPayroll
    'Declare tax constants
    Const _cdecFICA As Decimal = 0.0765D
    Const _cdecFed As Decimal = 0.22D
    Const _cdecState As Decimal = 0.04D

    'Event handling of initial form load
    Private Sub frmPayroll_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Display heading by default, but draw tax labels from constants
        lblFICA.Text = "FICA Tax (" & _cdecFICA.ToString("P2") & "):"
        lblFed.Text = "Federal Tax (" & _cdecFed.ToString("P2") & "):"
        lblState.Text = "State Tax (" & _cdecState.ToString("P2") & "):"

        'Clear the textbox placeholder digits.
        txtGrossPay.Text = ""
        txtFICA.Text = ""
        txtFed.Text = ""
        txtState.Text = ""
        txtNetPay.Text = ""
    End Sub

    'Event handling of Calulate button
    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click
        'Take user input And convert to numerical.
        'Calculate corresponding tax amounts.
        'Output final netPay value (two decimals).
        'Display tax amounts (two decimals).
        'All values in currency format

        Dim strGrossPay As String
        strGrossPay = txtGrossPay.Text
        Dim decGrossPay As Decimal
        decGrossPay = Convert.ToDecimal(strGrossPay)

        Dim decFICA As Decimal
        decFICA = decGrossPay * _cdecFICA
        Dim decFed As Decimal
        decFed = decGrossPay * _cdecFed
        Dim decState As Decimal
        decState = decGrossPay * _cdecState

        Dim decNetPay As Decimal
        decNetPay = (decGrossPay - decFICA - decFed - decState)
        txtNetPay.Text = decNetPay.ToString("C2")

        txtFICA.Text = decFICA.ToString("C2")
        txtFed.Text = decFed.ToString("C2")
        txtState.Text = decState.ToString("C2")
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        'Clear netPay, as well as tax text box values.
        'Clear initial gross pay value.
        'Put Focus on the gross pay text box.

        txtFICA.Clear()
        txtFed.Clear()
        txtState.Clear()
        txtNetPay.Clear()
        txtGrossPay.Clear()
        txtGrossPay.Focus()
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Closes program
        Close()
    End Sub

End Class