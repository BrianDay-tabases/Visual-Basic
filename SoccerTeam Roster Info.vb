'Title:     Soccer Team Roster Info
'Author:    Brian Day
'Date:      April 28
'Purpose:   This application displays the information retrieved by the Soccer DB as player linked data. Age Query button runs a query and then
'           the application counts the number of team members who are 12, 13, 14 years old, and calculates the average age of the team.

Option Strict On

Public Class Soccer
    Private Sub TeamBindingNavigatorSaveItem_Click(sender As Object, e As EventArgs) Handles TeamBindingNavigatorSaveItem.Click
        Me.Validate()
        Me.TeamBindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(Me.SoccerDataSet)
    End Sub

    'Form Load autofills dataset
    Private Sub Soccer_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'SoccerDataSet.Team' table. You can move, or remove it, as needed.
        Me.TeamTableAdapter.Fill(Me.SoccerDataSet.Team)

    End Sub

    Private Sub btnAgesQuery_Click(sender As Object, e As EventArgs) Handles btnAgesQuery.Click
        'SQL Select String
        Dim strSQL As String = "SELECT * FROM Team"

        'strPath provides the database type and path of the Soccer database
        Dim strPath As String = "Provider=Microsoft.ACE.OLEDB.12.0 ;" & "Data Source=C:\Users\aday\Documents\Spring18\BIS305\Soccer.accdb"
        'pass in the SQL String and String Path to the Db into the OleDbDataAdapter
        Dim odaPlayer As New OleDb.OleDbDataAdapter(strSQL, strPath)
        'Dim the DataTable
        Dim datValue As New DataTable

        'Declare Variables
        Dim decTotalAge As Decimal = 0D
        Dim intPlayerCount As Integer
        Dim decAverageAge As Decimal = 0D
        Dim int12Count As Integer
        Dim int13Count As Integer
        Dim int14Count As Integer

        'the DataTable name datValue is filled with the table data
        odaPlayer.Fill(datValue)
        'the connection to the database is disconnected
        odaPlayer.Dispose()
        'start the for/next loop to interate the datatable rows and calculate age totals
        For intPlayerCount = 0 To datValue.Rows.Count - 1
            'create variable to hold the age of the player in the current record in the loop
            Dim decAge As Decimal = Convert.ToDecimal(datValue.Rows(intPlayerCount)("Age"))

            'if condition that checks if the player is 12 years of age, then increase the 12 year old total count variable by 1
            If decAge = 12D Then
                int12Count += 1
            End If
            'if condition that checks if the player is 13 years of age, increase the 13 year old total count variable by 1
            If decAge = 13D Then
                int13Count += 1
            End If
            'if condition that checks if the player is 14 years of age, increase the 14 year old total count variable by 1
            If decAge = 14D Then
                int14Count += 1
            End If
            'increase the total age variable by the age of the current loop player
            decTotalAge += decAge
        Next

        'Calculate average age
        decAverageAge = decTotalAge / intPlayerCount

        'set the average age and player age labels visible property to True to display the labels
        'set the label text properties for the total number of 12, 13, and 14 year olds. In addition to the team average.
        lblPlayersAge12.Visible = True
        lblPlayersAge13.Visible = True
        lblPlayersAge14.Visible = True
        lblAvg.Visible = True
        lbl12Count.Visible = True
        lbl12Count.Text = int12Count.ToString("G")
        lbl13Count.Visible = True
        lbl13Count.Text = int13Count.ToString("G")
        lbl14Count.Visible = True
        lbl14Count.Text = int14Count.ToString("G")
        lblTeamAvg.Visible = True
        lblTeamAvg.Text = decAverageAge.ToString("F2")
    End Sub
End Class