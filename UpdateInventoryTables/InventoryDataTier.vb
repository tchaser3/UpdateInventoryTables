'Title:         Inventory Data Tier
'Date:          3-11-15
'Author:        Terry Holmes

'Description:   Data Tier for the Inventory Table

Option Strict On

Public Class InventoryDataTier

    Private aInventoryDataSetTableAdapter As InventoryDataSetTableAdapters.InventoryTableAdapter
    Private aInventoryDataSet As InventoryDataSet

    Public Function GetInventoryInformation() As InventoryDataSet

        'Setting up the Datatier
        Try
            aInventoryDataSet = New InventoryDataSet
            aInventoryDataSetTableAdapter = New InventoryDataSetTableAdapters.InventoryTableAdapter
            aInventoryDataSetTableAdapter.Fill(aInventoryDataSet.Inventory)
            Return aInventoryDataSet

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Check", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return aInventoryDataSet
        End Try
    End Function

    Public Sub UpdateInventoryDB(ByVal aInventoryDataSet As InventoryDataSet)

        'This will update the database
        Try
            aInventoryDataSetTableAdapter.Update(aInventoryDataSet.Inventory)
        Catch ex As Exception

        End Try
    End Sub

End Class
