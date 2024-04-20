'Title:         Update Inventory Table
'Date:          3-11-15
'Author:        Terry Holmes

'Description:   This will update the Inventory Table for TWC

Option Strict On

Public Class UpdateInventoryTable

    'Creating the variables for the data
    Private addingBoolean As Boolean = False
    Private editingBoolean As Boolean = False
    Private previousSelectedIndex As Integer

    'Setting up the employee
    Private TheEmployeeDataSet As EmployeesDataSet
    Private TheEmployeeDataTier As EmployeeDataTier
    Private WithEvents TheEmployeeBindingSource As BindingSource

    'Setting up the warehouse
    Private TheInventoryDataSet As InventoryDataSet
    Private TheInventoryDataTier As InventoryDataTier
    Private WithEvents TheInventoryBindingSource As BindingSource

    'Setting up the received parts
    Private TheReceivePartsDataSet As ReceivePartsDataSet
    Private TheReceivePartsDataTier As ReceivePartsDataTier
    Private WithEvents TheReceivePartsBindingSource As BindingSource

    'Setting up the issued parts
    Private TheIssuedPartsDataSet As IssuedPartsDataSet
    Private TheIssuedPartsDataTier As IssuedPartsDataTier
    Private WithEvents TheIssuedPartsBindingSource As BindingSource

    'Setting up the Used Parts
    Private TheUsedPartsDataSet As BOMPartsDataSet
    Private TheUsedPartsDataTier As BOMPartsDataTier
    Private WithEvents TheUsedPartsBindingSource As BindingSource

    'setting up the warehouse array
    Dim mstrWarehousePartNumber() As String
    Dim mintWarehouseCounter As Integer
    Dim mintWarehouseUpperLimit As Integer

    'Setting up added part numbers in added array
    Dim mstrNewPartNumberAdded() As String
    Dim mintNewPartCounter As Integer
    Dim mintNewPartUpperLimit As Integer

    'Setting up the employee array
    Dim mintEmployeeSelectedIndex() As Integer
    Dim mintEmployeeCounter As Integer
    Dim mintEmployeeUpperLimit As Integer

    Dim mintNewPrintCounter As Integer
    Dim LogDate As Date = Date.Now
    Dim mstrDate As String

    Private Sub SetWarehouseInventoryDataBindings()

        'try catch for exceptions
        Try

            TheInventoryDataTier = New InventoryDataTier
            TheInventoryDataSet = TheInventoryDataTier.GetInventoryInformation
            TheInventoryBindingSource = New BindingSource

            'Setting up the binding source
            With TheInventoryBindingSource
                .DataSource = TheInventoryDataSet
                .DataMember = "Inventory"
                .MoveFirst()
                .MoveLast()
            End With

            'Setting up the combo box
            With cboPartID
                .DataSource = TheInventoryBindingSource
                .DisplayMember = "PartID"
                .DataBindings.Add("text", TheInventoryBindingSource, "PartID", False, DataSourceUpdateMode.Never)
            End With

            'Setting the rest of the controls
            txtPartNumber.DataBindings.Add("text", TheInventoryBindingSource, "PartNumber")
            txtQTYOnHand.DataBindings.Add("text", TheInventoryBindingSource, "QTYResponible")
            txtWarehouseID.DataBindings.Add("Text", TheInventoryBindingSource, "WarehouseID")

        Catch ex As Exception

            'Message to user
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub SetEmployeeDataBindings()

        'Setting local variables
        Dim intCounter As Integer
        Dim intNumberOfRecords As Integer
        Dim strLastNameForSearch As String
        Dim strLastNameFromTable As String

        'try catch to handle execptions
        Try
            TheEmployeeDataTier = New EmployeeDataTier
            TheEmployeeDataSet = TheEmployeeDataTier.GetEmployeesInformation
            TheEmployeeBindingSource = New BindingSource

            'Setting up the binding source
            With TheEmployeeBindingSource
                .DataSource = TheEmployeeDataSet
                .DataMember = "employees"
                .MoveFirst()
                .MoveLast()
            End With

            'setting up the combo box
            With cboWarehouseID
                .DataSource = TheEmployeeBindingSource
                .DisplayMember = "EmployeeID"
                .DataBindings.Add("text", TheEmployeeBindingSource, "EmployeeID", False, DataSourceUpdateMode.Never)
            End With

            'Setting up the rest of the controls
            txtFirstName.DataBindings.Add("text", TheEmployeeBindingSource, "FirstName")
            txtLastName.DataBindings.Add("text", TheEmployeeBindingSource, "LastName")

            'Getting ready to find the warehouses
            intNumberOfRecords = cboWarehouseID.Items.Count - 1
            ReDim mintEmployeeSelectedIndex(intNumberOfRecords)
            strLastNameForSearch = "PARTS"
            mintEmployeeCounter = 0

            'Beginning loop
            For intCounter = 0 To intNumberOfRecords

                'Incrementing the combo box
                cboWarehouseID.SelectedIndex = intCounter

                'getting the last name
                strLastNameFromTable = txtLastName.Text

                'If statements
                If strLastNameForSearch = strLastNameFromTable Then

                    'Loading the array
                    mintEmployeeSelectedIndex(mintEmployeeCounter) = intCounter
                    mintEmployeeCounter += 1

                End If
            Next

            'setting the variables
            mintEmployeeUpperLimit = mintEmployeeCounter - 1
            mintEmployeeCounter = 0
            If mintEmployeeUpperLimit > 0 Then
                btnNext.Enabled = True
            End If

            'setting location of the warehouse
            cboWarehouseID.SelectedIndex = mintEmployeeSelectedIndex(0)

        Catch ex As Exception

            'this will alert the user if there is a problem
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click

        'This will close the program
        CloseProgram.ShowDialog()

    End Sub
    Private Sub SetControlsVisible(ByVal ValueBoolean As Boolean)

        'This will make the controls visible
        cboPartID.Visible = ValueBoolean
        cboTransactionID.Visible = ValueBoolean
        txtPartNumber.Visible = ValueBoolean
        txtQTYOnHand.Visible = ValueBoolean
        txtTransactionPartNumber.Visible = ValueBoolean
        txtWarehouseID.Visible = ValueBoolean

    End Sub

    Private Sub btnMainMenu_Click(sender As Object, e As EventArgs) Handles btnMainMenu.Click

        'This will show the main menu
        LastTransaction.Show()
        ClearMainDataBindings()
        ClearTransactionalDataBindings()
        Form1.Show()
        Me.Close()

    End Sub
    Private Sub ClearMainDataBindings()

        'This will clear the data bindings
        cboPartID.DataBindings.Clear()
        cboWarehouseID.DataBindings.Clear()
        txtLastName.DataBindings.Clear()
        txtFirstName.DataBindings.Clear()
        txtPartNumber.DataBindings.Clear()
        txtQTYOnHand.DataBindings.Clear()
        txtWarehouseID.DataBindings.Clear()

    End Sub
    Private Sub ClearTransactionalDataBindings()

        'this will clear the transactional data bindings
        cboTransactionID.DataBindings.Clear()
        txtTransactionPartNumber.DataBindings.Clear()

    End Sub
    Private Sub SetComboBoxBinding()

        With cboPartID
            If (addingBoolean Or editingBoolean) Then
                .DataBindings!text.DataSourceUpdateMode = DataSourceUpdateMode.OnValidation
                .DropDownStyle = ComboBoxStyle.Simple
            Else
                .DataBindings!text.DataSourceUpdateMode = DataSourceUpdateMode.Never
                .DropDownStyle = ComboBoxStyle.DropDownList
            End If
        End With
    End Sub

    Private Sub UpdateWarehouseInventory_Load(sender As Object, e As EventArgs) Handles Me.Load

        'setting initial conditions
        SetEmployeeDataBindings()
        SetWarehouseInventoryDataBindings()
        SetControlsVisible(False)

    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click

        'Incrementing the combo box
        mintEmployeeCounter += 1
        cboWarehouseID.SelectedIndex = mintEmployeeSelectedIndex(mintEmployeeCounter)

        btnBack.Enabled = True

        If mintEmployeeCounter = mintEmployeeUpperLimit Then
            btnNext.Enabled = False
        End If
    End Sub

    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click

        'Incrementing the combo box
        mintEmployeeCounter -= 1
        cboWarehouseID.SelectedIndex = mintEmployeeSelectedIndex(mintEmployeeCounter)

        btnNext.Enabled = True

        If mintEmployeeCounter = 0 Then
            btnBack.Enabled = False
        End If
    End Sub
    Private Sub LoadWarehousePartNumberArray()

        'This sub-routine will load the part number array
        'Setting local variables
        Dim intCounter As Integer
        Dim intNumberOfRecords As Integer
        Dim intWarehouseIDForSearch As Integer
        Dim intWarehouseIDFromTable As Integer

        'Setting up for loop
        intNumberOfRecords = cboWarehouseID.Items.Count - 1
        ReDim mstrWarehousePartNumber(intNumberOfRecords)
        intWarehouseIDForSearch = CInt(cboWarehouseID.Text)
        mintWarehouseCounter = 0

        'Preforming loop
        For intCounter = 0 To intNumberOfRecords

            'incrementing the combo box
            cboPartID.SelectedIndex = intCounter

            'Getting warehouse id
            intWarehouseIDFromTable = CInt(txtWarehouseID.Text)

            If intWarehouseIDForSearch = intWarehouseIDFromTable Then

                'Loading the array
                mstrWarehousePartNumber(mintWarehouseCounter) = txtPartNumber.Text
                mintWarehouseCounter += 1

            End If
        Next

        'Setting up for the array
        If mintWarehouseCounter > 0 Then
            mintWarehouseUpperLimit = mintWarehouseCounter - 1
        Else
            mintWarehouseUpperLimit = 0
        End If

        mintWarehouseCounter = 0

    End Sub
    Private Function SetReceivedPartsDataBindings() As Boolean

        'Setting local variable
        Dim blnFatalError As Boolean = False

        Try

            'setting up the global data variables
            TheReceivePartsDataTier = New ReceivePartsDataTier
            TheReceivePartsDataSet = TheReceivePartsDataTier.GetReceivePartsInformation
            TheReceivePartsBindingSource = New BindingSource

            'setting up the binding source
            With TheReceivePartsBindingSource
                .DataSource = TheReceivePartsDataSet
                .DataMember = "ReceivedParts"
                .MoveFirst()
                .MoveLast()
            End With

            'Setting up the combo box
            With cboTransactionID
                .DataSource = TheReceivePartsBindingSource
                .DisplayMember = "TransactionID"
                .DataBindings.Add("text", TheReceivePartsBindingSource, "TransactionID", False, DataSourceUpdateMode.Never)
            End With

            'Binding the rest of the controls
            txtTransactionPartNumber.DataBindings.Add("text", TheReceivePartsBindingSource, "PartNumber")

            'Return to calling 
            Return blnFatalError

        Catch ex As Exception

            'Message to the user
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return blnFatalError

        End Try

    End Function
    Private Function SetIssuedPartsDataBindings() As Boolean

        'Setting local variable
        Dim blnFatalError As Boolean = False

        Try

            'setting up the global data variables
            TheIssuedPartsDataTier = New IssuedPartsDataTier
            TheIssuedPartsDataSet = TheIssuedPartsDataTier.GetIssuedPartsInformation
            TheIssuedPartsBindingSource = New BindingSource

            'setting up the binding source
            With TheIssuedPartsBindingSource
                .DataSource = TheIssuedPartsDataSet
                .DataMember = "IssuedParts"
                .MoveFirst()
                .MoveLast()
            End With

            'Setting up the combo box
            With cboTransactionID
                .DataSource = TheIssuedPartsBindingSource
                .DisplayMember = "TransactionID"
                .DataBindings.Add("text", TheIssuedPartsBindingSource, "TransactionID", False, DataSourceUpdateMode.Never)
            End With

            'Binding the rest of the controls
            txtTransactionPartNumber.DataBindings.Add("text", TheIssuedPartsBindingSource, "PartNumber")

            'Return to calling 
            Return blnFatalError

        Catch ex As Exception

            'Message to the user
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return blnFatalError

        End Try

    End Function
    Private Function SetUsedPartsDataBindings() As Boolean

        'Setting local variable
        Dim blnFatalError As Boolean = False

        Try

            'setting up the global data variables
            TheUsedPartsDataTier = New BOMPartsDataTier
            TheUsedPartsDataSet = TheUsedPartsDataTier.GetBOMPartsInformation
            TheUsedPartsBindingSource = New BindingSource

            'setting up the binding source
            With TheUsedPartsBindingSource
                .DataSource = TheUsedPartsDataSet
                .DataMember = "BOMParts"
                .MoveFirst()
                .MoveLast()
            End With

            'Setting up the combo box
            With cboTransactionID
                .DataSource = TheUsedPartsBindingSource
                .DisplayMember = "TransactionID"
                .DataBindings.Add("text", TheUsedPartsBindingSource, "TransactionID", False, DataSourceUpdateMode.Never)
            End With

            'Binding the rest of the controls
            txtTransactionPartNumber.DataBindings.Add("text", TheUsedPartsBindingSource, "PartNumber")

            'Return to calling 
            Return blnFatalError

        Catch ex As Exception

            'Message to the user
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return blnFatalError

        End Try

    End Function

    Private Sub btnProcess_Click(sender As Object, e As EventArgs) Handles btnProcess.Click

        'setting local variables
        Dim blnFatalError As Boolean
        Dim intNumberOfRecords As Integer

        'This sub routine perform the copy
        ClearTransactionalDataBindings()
        SetControlsVisible(True)
        LoadWarehousePartNumberArray()
        mintNewPartCounter = 0
        mintNewPartUpperLimit = 0

        'First set of data bindings called
        blnFatalError = SetReceivedPartsDataBindings()

        intNumberOfRecords = cboTransactionID.Items.Count - 1
        ReDim mstrNewPartNumberAdded(intNumberOfRecords)

        If blnFatalError = True Then
            Exit Sub
        End If

        CheckPartNumbers()

        'Setting up for issed parts
        ClearTransactionalDataBindings()
        blnFatalError = SetIssuedPartsDataBindings()

        If blnFatalError = True Then
            Exit Sub
        End If

        CheckPartNumbers()

        'setting up for used parts
        ClearTransactionalDataBindings()
        blnFatalError = SetUsedPartsDataBindings()

        If blnFatalError = True Then
            Exit Sub
        End If

        CheckPartNumbers()

        'ending sub-routine
        ClearTransactionalDataBindings()
        Logon.mstrLastTransactionSummary = "UPDATE WAREHOUSE INVENTORY TABLE WITH PART NUMBERS"
        LastTransaction.Show()
        SetControlsVisible(False)
        MessageBox.Show("Procedure is Done, a Report is Being Generated for Part Numbers Added", "Thank You", MessageBoxButtons.OK, MessageBoxIcon.Information)

        If PrintDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            PrintDocument1.PrinterSettings = PrintDialog1.PrinterSettings
        End If

        mintNewPrintCounter = 0

        PrintDocument1.Print()

    End Sub
    Private Sub CheckPartNumbers()

        'This sub routine will check the part number to see if it exists
        'Setting local variables
        Dim intTransactionCounter As Integer
        Dim intTransactionNumberOfRecords As Integer
        Dim intWarehouseCounter As Integer
        Dim strPartNumberForSearch As String
        Dim strPartNumberFromTable As String
        Dim blnItemNotFound As Boolean
        Dim intPartNumberLength As Integer

        'Getting ready for loop
        intTransactionNumberOfRecords = cboTransactionID.Items.Count - 1

        'Performing Loop
        For intTransactionCounter = 0 To intTransactionNumberOfRecords

            'Incrementing the combo box
            cboTransactionID.SelectedIndex = intTransactionCounter

            'Getting the part number
            strPartNumberForSearch = txtTransactionPartNumber.Text

            'Setting boolean modifier
            blnItemNotFound = True

            'Seaching the array
            For intWarehouseCounter = 0 To mintWarehouseUpperLimit

                'Getting part number
                strPartNumberFromTable = mstrWarehousePartNumber(intWarehouseCounter)

                'Checking Part Number
                If strPartNumberForSearch = strPartNumberFromTable Then

                    'setting boolean modifier
                    blnItemNotFound = False
                End If

            Next

            If blnItemNotFound = True Then

                'Setting up loop
                If mintNewPartUpperLimit > 0 Then

                    For intWarehouseCounter = 0 To mintNewPartUpperLimit - 1

                        'getting the part number
                        strPartNumberFromTable = mstrNewPartNumberAdded(intWarehouseCounter)

                        If strPartNumberForSearch = strPartNumberFromTable Then

                            'Setting boolean modifier
                            blnItemNotFound = False
                        End If
                    Next

                End If
            End If

            'If statements to see if the item was found
            If blnItemNotFound = True Then

                'getting string length
                intPartNumberLength = strPartNumberForSearch.Length

                If intPartNumberLength = 7 Then

                    'Setting variablesReDim mstrNewPartNumberAdded(mintNewPartUpperLimit)
                    mstrNewPartNumberAdded(mintNewPartCounter) = strPartNumberForSearch
                    AddNewPartNumberRecord(strPartNumberForSearch)
                    mintNewPartCounter += 1
                    mintNewPartUpperLimit = mintNewPartCounter

                End If
            End If
        Next

    End Sub

    Private Sub AddNewPartNumberRecord(ByVal strPartNumberToAdd As String)

        'this sub routine will create a record within the warehouse inventory table
        Try

            With TheInventoryBindingSource
                .EndEdit()
                .AddNew()
            End With

            'Setting up to bind the combo box
            addingBoolean = True
            SetComboBoxBinding()

            'Setting the rest of the controls
            CreatePartID.Show()
            cboPartID.Text = CStr(Logon.mintPartID)
            txtPartNumber.Text = strPartNumberToAdd
            txtWarehouseID.Text = cboWarehouseID.Text
            txtQTYOnHand.Text = "0"

            'updating the table
            TheInventoryBindingSource.EndEdit()
            TheInventoryDataTier.UpdateInventoryDB(TheInventoryDataSet)
            addingBoolean = False
            editingBoolean = False
            SetComboBoxBinding()

        Catch ex As Exception

            'Message to user
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage

        'This will print the document

        'Setting local variables
        Dim intCounter As Integer
        Dim intStartingPageCounter As Integer
        Dim blnNewPage As Boolean = False

        'Setting up variables for the reports
        Dim PrintHeaderFont As New Font("Arial", 18, FontStyle.Bold)
        Dim PrintSubHeaderFont As New Font("Arial", 14, FontStyle.Bold)
        Dim PrintItemsFont As New Font("Arial", 10, FontStyle.Regular)
        Dim PrintX As Single = e.MarginBounds.Left
        Dim PrintY As Single = e.MarginBounds.Top
        Dim HeadingLineHeight As Single = PrintHeaderFont.GetHeight + 18
        Dim ItemLineHeight As Single = PrintItemsFont.GetHeight + 10

        'Getting the date
        mstrDate = CStr(LogDate)

        'Setting up for default position
        PrintY = 100

        'Setting up for the print header
        PrintX = 150
        e.Graphics.DrawString("Blue Jay Communications Inventory Report", PrintHeaderFont, Brushes.Black, PrintX, PrintY)
        PrintY = PrintY + HeadingLineHeight
        PrintX = 162
        e.Graphics.DrawString("Part Numbers To Counted:  " + mstrDate, PrintSubHeaderFont, Brushes.Black, PrintX, PrintY)
        PrintY = PrintY + HeadingLineHeight

        'Setting up the columns
        PrintX = 100
        e.Graphics.DrawString("Part Number", PrintItemsFont, Brushes.Black, PrintX, PrintY)
        PrintY = PrintY + HeadingLineHeight

        'Performing Loop
        For intCounter = mintNewPrintCounter To mintNewPartUpperLimit

            PrintX = 100
            e.Graphics.DrawString(mstrNewPartNumberAdded(intCounter), PrintItemsFont, Brushes.Black, PrintX, PrintY)
            PrintY = PrintY + ItemLineHeight

            If PrintY > 900 Then
                If intStartingPageCounter = mintNewPartUpperLimit Then
                    e.HasMorePages = False
                Else
                    e.HasMorePages = True
                    blnNewPage = True
                End If
            End If

            If blnNewPage = True Then
                mintNewPrintCounter = intCounter + 1
                intCounter = mintNewPartUpperLimit
            End If

        Next

    End Sub

End Class