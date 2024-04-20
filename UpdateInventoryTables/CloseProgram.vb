'Title:         Close Program
'Date:          2-13-15
'Author:        Terry Holmes

'Description:   This will close the program

Option Strict On

Public Class CloseProgram

    Private Sub btnNo_Click(sender As Object, e As EventArgs) Handles btnNo.Click

        'Returns back to main form
        Me.Close()

    End Sub

    Private Sub btnYes_Click(sender As Object, e As EventArgs) Handles btnYes.Click

        'This will close the program
        Logon.Close()
        Me.Close()

    End Sub
End Class