'Title:         Warehouse Inventory Data Tier
'Date:          2-13-15
'Author:        Terry Holmes

'Description:   This is the data tier for the warehouse inventory data set

Public Class WarehouseInventoryDataTier

    Private aWarehouseInventoryDataSetTableAdapter As WarehouseInventoryDataSetTableAdapters.WarehouseInventoryTableAdapter
    Private aWarehouseInventoryDataSet As WarehouseInventoryDataSet

    Public Function GetWarehouseInventoryInformation() As WarehouseInventoryDataSet

        'Setting up the Datatier
        Try
            aWarehouseInventoryDataSet = New WarehouseInventoryDataSet
            aWarehouseInventoryDataSetTableAdapter = New WarehouseInventoryDataSetTableAdapters.WarehouseInventoryTableAdapter
            aWarehouseInventoryDataSetTableAdapter.Fill(aWarehouseInventoryDataSet.WarehouseInventory)
            Return aWarehouseInventoryDataSet

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Check", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return aWarehouseInventoryDataSet
        End Try
    End Function

    Public Sub UpdateWarehouseInventoryDB(ByVal aWarehouseInventoryDataSet As WarehouseInventoryDataSet)

        'This will update the database
        Try
            aWarehouseInventoryDataSetTableAdapter.Update(aWarehouseInventoryDataSet.WarehouseInventory)
        Catch ex As Exception

        End Try
    End Sub

End Class
