Imports System.Data.SqlClient

Public Class frmUser
    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        ' Dim id As Integer = txtID.Text
        Dim name As String = txtName.Text
        Dim address As String = txtAddress.Text
        Dim city As String = cmbCity.Text
        Dim age As Integer = txtAge.Text
        Dim sex As String = ""

        If rdbMale.Checked = True Then
            sex = rdbMale.Text
        ElseIf rdbFemale.Checked = True Then
            sex = rdbFemale.Text
        End If

        'se realiza la consulta de captura y la configuracion de la base de datos
        Dim query As String = "INSERT INTO UserInfo(Name,Address,city,Age,Sex) VALUES(@name,@address,@city,@age,@sex)"
        Using con As SqlConnection = New SqlConnection("Data Source=DESKTOP-JT4DFR6;Initial Catalog=Users;Integrated Security=True")
            Using cmd As SqlCommand = New SqlCommand(query, con)
                'cmd.Parameters.AddWithValue("@id", id)
                cmd.Parameters.AddWithValue("@name", name)
                cmd.Parameters.AddWithValue("@address", address)
                cmd.Parameters.AddWithValue("@city", city)
                cmd.Parameters.AddWithValue("@age", age)
                cmd.Parameters.AddWithValue("@sex", sex)
                con.Open()
                cmd.ExecuteNonQuery()
                con.Close()
                clear_form()
                MessageBox.Show("Successfully added")
            End Using
        End Using
    End Sub

    Private Sub clear_form()
        ' Limpiar formulario
        txtName.Text = ""
        txtAddress.Text = ""
        cmbCity.Items.Clear()
        txtAge.Text = ""
        rdbMale.Checked = False
        rdbFemale.Checked = False
    End Sub

End Class
