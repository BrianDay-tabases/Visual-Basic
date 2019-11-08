' Program Name  Patient Cholesterol Info
' Author        Brian Day
' Date          April 13
' Purpose       Read patient info, display patient info, calculate if patient has high cholesterol level, write output info, 
'               compute and display average level along with running total of cases

Option Strict On

Public Class frmCholesterol
    ' Class Level variables
    'Patient Name Array
    Public Shared _strPatientName(15) As String
    'Cholesterol Level Array
    Public Shared _decPatientLevel(15) As Decimal

    Dim strPatientFile As String = "C:\Users\aday\Documents\Spring18\BIS305\DayAssignment8\patient.txt"
    Dim strConsultFile As String = "C:\Users\aday\Documents\Spring18\BIS305\DayAssignment8\consult.txt"

    ' The frmDepreciation loads, reads the patient text file
    Private Sub frmCholesterol_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Allow SplashScreen time to display
        Threading.Thread.Sleep(2000)
        lblDisplayTotal.Text = ""
        lblDisplayAverage.Text = ""
        'Initialize StreamReader/Writer objects and declare variables
        Dim objReader As IO.StreamReader
        Dim objWriter As New IO.StreamWriter(strConsultFile)
        'Loop counter
        Dim intCount As Integer = 0
        'Total patients with high level
        Dim intPatientTotal As Integer = 0
        'Sum of cholesterol levels
        Dim decPatientSum As Decimal


        'Verify the file exists
        If IO.File.Exists(strPatientFile) Then
            objReader = IO.File.OpenText(strPatientFile)

            ' Read the file line by line until the file is completed
            Do While objReader.Peek <> -1
                _strPatientName(intCount) = objReader.ReadLine()

                _decPatientLevel(intCount) = Convert.ToDecimal(objReader.ReadLine())

                'add current cholesterol level to total cholesterol amount variable
                decPatientSum = decPatientSum + _decPatientLevel(intCount)



                If IO.File.Exists(strConsultFile) Then
                    If (_decPatientLevel(intCount) >= 200) Then
                        'Write to consult text file the patient name of patient array
                        objWriter.WriteLine(_strPatientName(intCount))
                        'Write to consult text file the cholesterol level of cholestrol array
                        objWriter.WriteLine(_decPatientLevel(intCount))
                        intPatientTotal += 1
                    End If
                End If

                'increment loop counter
                intCount += 1

            Loop
            objWriter.Close()
            objReader.Close()

        Else
            MsgBox("The file is not available. Please correct and restart application.", , "Error")
            Close()
        End If


        'set the label text to show the Average Cholesterol
        lblDisplayAverage.Text = (decPatientSum / _strPatientName.Length).ToString("F2")

        'set the label text to she the count of patients with high cholesterol above 200
        lblDisplayTotal.Text = intPatientTotal.ToString("G")


    End Sub

    'Menu Object Display Data
    Private Sub mnuDisplayData_Click(sender As Object, e As EventArgs) Handles mnuDisplayData.Click
        ' The mnuDisplay click event creates an instance of the frmDisplayInfo
        Dim frmSecond As New frmDisplayInfo

        'Hide this form and show the Display Info form
        Hide()
        frmSecond.ShowDialog()
    End Sub

    'Menu Object Clears and Resets Form
    Private Sub mnuClear_Click(sender As Object, e As EventArgs) Handles mnuClear.Click
        lblDisplayTotal.Visible = False
        lblTotal.Visible = False
        lblDisplayAverage.Visible = False


        frmDisplayInfo.lstName.Items.Clear()
        frmDisplayInfo.lstLevel.Items.Clear()
    End Sub

    'Menu Object Exit Application
    Private Sub mnuExit_Click(sender As Object, e As EventArgs) Handles mnuExit.Click
        Application.Exit()
    End Sub

End Class
