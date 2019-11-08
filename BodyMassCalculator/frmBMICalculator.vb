'Program Name: BMI Calculator
'Author: Brian Day
'Date: April 6
'Purpose: Calculates BMI Index rate for user based on their Height and Weight inputs and selection of Metric or Imperial

Option Strict On

Public Class frmBMICalculator
    Dim _intMeasurementChoice As Integer

    'Event Handling for System of Measurement
    Private Sub lstSystemofMeasurement_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstSystemofMeasurement.SelectedIndexChanged
        Dim intUserChoice = lstSystemofMeasurement.SelectedIndex
        'lstSystemofMeasurement.Items.Clear()
        Select Case intUserChoice
            Case 0
                _intMeasurementChoice = 1 'Metric
            Case 1
                _intMeasurementChoice = 2 'Imperial
        End Select
        lblBMI.Visible = True
        lblBMI.Text = ""
        lblWeightStatus.Visible = True
        lblWeightStatus.Text = ""
        btnCalculate.Visible = True
        btnClear.Visible = True
        txtHeight.Focus()
    End Sub

    'Calculate button event handling
    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click
        'Declare conversion variables and MsgBox variables
        Dim decHeight As Decimal
        Dim decWeight As Decimal
        Dim decBMI As Decimal
        Dim strMessageBoxTitle As String = "Error"

        'Try/Catch for validating numerical values
        Try
            decHeight = Convert.ToDecimal(txtHeight.Text)
            decWeight = Convert.ToDecimal(txtWeight.Text)
            'Catch Non-numericals
        Catch Exception As FormatException
            MsgBox("Please use only numerical values", , strMessageBoxTitle)
            txtHeight.Clear()
            txtWeight.Clear()
            txtHeight.Focus()
            'Catch Divsion by Zero
        Catch Exception As DivideByZeroException
            MsgBox("Division by Zero not allowed", , strMessageBoxTitle)
            txtHeight.Clear()
            txtWeight.Clear()
            txtHeight.Focus()
            'Catch superfluous numericals
        Catch Exception As OverflowException
            MsgBox("Please use only numbers a normal human would possess for height and weight", , strMessageBoxTitle)
            txtHeight.Clear()
            txtWeight.Clear()
            txtHeight.Focus()
            'Catch all
        Catch Exception As SystemException
            MsgBox("Unknown solution to issues with values entered, Please try again", , strMessageBoxTitle)
            txtHeight.Clear()
            txtWeight.Clear()
            txtHeight.Focus()
        End Try

        'decBMI = DisplayWeightStatus(decHeight, decWeight)

        'Call appropriate procedures and display final values to labels
        If _intMeasurementChoice = 1 Then
            decBMI = MetricBMI(decHeight, decWeight)
        Else
            decBMI = ImperialBMI(decHeight, decWeight)
        End If

        If (decBMI >= 30) Then
            lblWeightStatus.Text = "Obese"
        ElseIf (decBMI <= 29.9 And decBMI >= 25.0) Then
            lblWeightStatus.Text = "Overweight"
        ElseIf (decBMI <= 24.9 And decBMI >= 18.5) Then
            lblWeightStatus.Text = "Normal Weight"
        Else
            lblWeightStatus.Text = "Underweight"
        End If

        lblBMI.Text = decBMI.ToString("F2")
    End Sub

    ' Private Sub DisplayWeightStatus()
    'End Sub

    'Function returning BMI for Metric Selection
    Private Function MetricBMI(ByVal decHeight As Decimal, ByVal decWeight As Decimal) As Decimal
        Dim decBMI As Decimal
        decBMI = decWeight / (decHeight * decHeight)
        Return decBMI
    End Function

    'Function returning BMI for Imperial Selection
    Private Function ImperialBMI(ByVal decHeight As Decimal, ByVal decWeight As Decimal) As Decimal
        Dim decBMI As Decimal
        decBMI = (decWeight / (decHeight * decHeight)) * 703
        Return decBMI
    End Function

    'Clear button event handling
    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        'Clear form labels and reset focus
        lblBMI.Text = ""
        lblWeightStatus.Text = ""
        btnCalculate.Visible = True
        btnClear.Visible = True
        txtHeight.Focus()
    End Sub

    Private Sub frmBMICalculator_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Opening splash screen should be displayed for 2 seconds
        Threading.Thread.Sleep(2000)
    End Sub

End Class
