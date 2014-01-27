' Andrew Thompson
' 01/26/2014
' PC 9-1
Imports System.IO

Public Class Form1
    Dim DataFile As StreamWriter

    Private Sub btnSaveRecord_Click(sender As Object, e As EventArgs) Handles btnSaveRecord.Click
        If Not AllFieldsValid() Then
            MessageBox.Show("Please check the input, all fields are required.")
            Return
        End If

        For Each field In GetFields()
            DataFile.WriteLine(field.Trim())
        Next

        ClearFields()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        While IsNothing(DataFile)
            Try
                DataFile = File.CreateText(InputBox("Enter data file name."))
            Catch ex As Exception
                MessageBox.Show("Error creating file, try again.")
            End Try
        End While
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        ClearFields()
    End Sub

    Private Sub ClearFields()
        txtFirstName.Text = String.Empty
        txtMiddleName.Text = String.Empty
        txtLastName.Text = String.Empty
        txtEmployeeNumber.Text = String.Empty
        cbxDepartment.SelectedIndex = -1
        txtTelephoneNumber.Text = String.Empty
        txtTelephoneExtension.Text = String.Empty
        txtEmailAddress.Text = String.Empty
    End Sub

    Private Function AllFieldsValid() As Boolean
        If String.IsNullOrEmpty(txtFirstName.Text) Then Return False
        If String.IsNullOrEmpty(txtMiddleName.Text) Then Return False
        If String.IsNullOrEmpty(txtLastName.Text) Then Return False
        If String.IsNullOrEmpty(txtEmployeeNumber.Text) Then Return False
        If String.IsNullOrEmpty(txtTelephoneNumber.Text) Then Return False
        If String.IsNullOrEmpty(txtTelephoneExtension.Text) Then Return False
        If String.IsNullOrEmpty(txtEmailAddress.Text) Then Return False
        Return True
    End Function

    Private Iterator Function GetFields() As IEnumerable(Of String)
        Yield txtFirstName.Text
        Yield txtMiddleName.Text
        Yield txtLastName.Text
        Yield txtEmployeeNumber.Text
        Yield cbxDepartment.SelectedItem
        Yield txtTelephoneNumber.Text
        Yield txtTelephoneExtension.Text
        Yield txtEmailAddress.Text
    End Function

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        DataFile.Close()
    End Sub
End Class
