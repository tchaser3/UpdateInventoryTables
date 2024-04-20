'title:         Create Part Id
'Date:          3-10-15
'Author:        Terry Holmes

'Description:   This form will create a new Part ID

Public Class CreatePartID

    'Creating the data set variables
    Private ThePartNumberIDDataTier As PartNumberDataTier
    Private ThePartNumberIDDataSet As PartNumberIDDataSet
    Private WithEvents ThePartNumberIDBindingSource As BindingSource


    Private Sub CreatePartID_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Setting local variables
        Dim intNewID As Integer

        'try catch to catch exceptions
        Try

            'Setting up the main controls
            ThePartNumberIDDataTier = New PartNumberDataTier
            ThePartNumberIDDataSet = ThePartNumberIDDataTier.GetPartNumberIDInformation
            ThePartNumberIDBindingSource = New BindingSource

            'Setting up the binding source
            With ThePartNumberIDBindingSource
                .DataSource = ThePartNumberIDDataSet
                .DataMember = "PartNumberID"
                .MoveFirst()
                .MoveLast()
            End With

            'Setting up the combo box
            With cboTransactionID
                .DataSource = ThePartNumberIDBindingSource
                .DisplayMember = "TransactionID"
                .DataBindings.Add("text", ThePartNumberIDBindingSource, "TransactionID", False, DataSourceUpdateMode.Never)
            End With

            'Setting up the rest of the controls
            txtPartID.DataBindings.Add("Text", ThePartNumberIDBindingSource, "PartID")

            'Getting ID
            intNewID = CInt(txtPartID.Text)
            intNewID += 1
            txtPartID.Text = CStr(intNewID)
            Logon.mintPartID = intNewID

            'Saving information
            ThePartNumberIDBindingSource.EndEdit()
            ThePartNumberIDDataTier.UpdatePartNumberIDDB(ThePartNumberIDDataSet)

            Me.Close()

        Catch ex As Exception

            'Message to user if there is a failure
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub
End Class