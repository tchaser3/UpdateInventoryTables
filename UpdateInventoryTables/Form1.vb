'Title:         Update Inventory Tables
'Date:          3-9-15
'Author:        Terry Holmes

'Description:   This application will copy all the part numbers form the Receive, Issue, and BOM tables into both the Inventory
'               and Warehouse Inventory Tables.  I will also add the warehouse ID for the warehouse selected

Option Strict On

Public Class Form1

    Dim TheModuleUnderDevlopment As New ModuleUnderDevelopment

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click

        'This will close the program
        CloseProgram.ShowDialog()

    End Sub

    Private Sub btnAbout_Click(sender As Object, e As EventArgs) Handles btnAbout.Click

        'This will display the about box
        AboutBox1.Show()

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        'This will set up for the last transaction
        Logon.mstrLastTransactionSummary = "LOADED MAIN MENU FORM FOR UPDATE INVENTORY TABLE PARTS"

    End Sub

    Private Sub btnWarehouseInventory_Click(sender As Object, e As EventArgs) Handles btnWarehouseInventory.Click

        'this will load the warehouse inventory table
        LastTransaction.Show()
        UpdateWarehouseInventory.Show()
        Me.Close()

    End Sub

    Private Sub btnTWCInventory_Click(sender As Object, e As EventArgs) Handles btnTWCInventory.Click

        'this will update the twc inventory table
        LastTransaction.Show()
        UpdateInventoryTable.Show()
        Me.Close()

    End Sub
End Class
