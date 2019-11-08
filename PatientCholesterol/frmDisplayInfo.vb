Option Strict On
Public Class frmDisplayInfo

    Private Sub frmDisplayInfo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' The frmDisplayInfo load event is a second form that displays the sorted name and level

        Dim strItem As String

        ' Sorts the arrays
        Array.Sort(frmCholesterol._strPatientName)
        Array.Sort(frmCholesterol._decPatientLevel)

        ' Displays the name array
        For Each strItem In frmCholesterol._strPatientName
            lstName.Items.Add(strItem)
        Next
        ' Displays the cholesterol array
        For Each strItem In frmCholesterol._decPatientLevel
            lstLevel.Items.Add(strItem).ToString()
        Next
    End Sub

    Private Sub btnReturn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReturn.Click
        ' This Sub procedure opens frmCholesterol
        Dim frmFirst As New frmCholesterol
        Hide()
        frmFirst.ShowDialog()
    End Sub
End Class