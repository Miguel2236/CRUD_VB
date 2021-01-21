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
                getData()
                MessageBox.Show("Successfully added")
            End Using
        End Using
    End Sub

    Private Sub clear_form()
        ' Limpiar formulario
        txtID.Text = ""
        txtName.Text = ""
        txtAddress.Text = ""
        cmbCity.Items.Clear()
        txtAge.Text = ""
        rdbMale.Checked = False
        rdbFemale.Checked = False
    End Sub

    Private Sub getData()
        ' obtener los datos y pegarlos a la tabla
        Dim sqlQuery As String = "SELECT * FROM UserInfo"
        Using con As SqlConnection = New SqlConnection("Data Source=DESKTOP-JT4DFR6;Initial Catalog=Users;Integrated Security=True")
            Using cmd As SqlCommand = New SqlCommand(sqlQuery, con)
                Using da As New SqlDataAdapter()
                    da.SelectCommand = cmd
                    Using dt As New DataTable()
                        da.Fill(dt)
                        grdData.DataSource = dt
                    End Using
                End Using
            End Using
        End Using
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Dim id As Integer = txtID.Text
        Dim sqlQuery As String = "SELECT * FROM UserInfo WHERE id = @id"
        Using con As SqlConnection = New SqlConnection("Data Source=DESKTOP-JT4DFR6;Initial Catalog=Users;Integrated Security=True")
            Using cmd As SqlCommand = New SqlCommand(sqlQuery, con)
                cmd.Parameters.AddWithValue("@id", id)
                Using da As New SqlDataAdapter()
                    da.SelectCommand = cmd
                    Using dt As New DataTable()
                        da.Fill(dt)
                        If dt.Rows.Count > 0 Then

                            txtID.Text = dt.Rows(0)(0).ToString()
                            txtName.Text = Trim(dt.Rows(0)(1).ToString())
                            txtAddress.Text = Trim(dt.Rows(0)(2).ToString())
                            cmbCity.SelectedItem = Trim(dt.Rows(0)(3).ToString())
                            txtAge.Text = dt.Rows(0)(4).ToString()

                            If Trim(dt.Rows(0)(5)) = "Male" Then
                                rdbMale.Checked = True
                            Else
                                rdbFemale.Checked = True
                            End If
                        Else
                            MessageBox.Show("No hay registro")
                        End If
                    End Using
                End Using

            End Using
        End Using
    End Sub

    Private Sub btnList_Click(sender As Object, e As EventArgs) Handles btnList.Click
        ' obtener la lista de todos los registros
        getData()
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        Dim id As Integer = txtID.Text
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

        Dim sqlQuery As String = "UPDATE UserInfo SET Name = @name, Address = @address, City = @city, Age = @age, Sex = @sex WHERE ID = @id"
        Using con As SqlConnection = New SqlConnection("Data Source=DESKTOP-JT4DFR6;Initial Catalog=Users;Integrated Security=True")
            Using cmd As SqlCommand = New SqlCommand(sqlQuery, con)
                cmd.Parameters.AddWithValue("@id", id)
                cmd.Parameters.AddWithValue("@name", name)
                cmd.Parameters.AddWithValue("@address", address)
                cmd.Parameters.AddWithValue("@city", city)
                cmd.Parameters.AddWithValue("@age", age)
                cmd.Parameters.AddWithValue("@sex", sex)
                con.Open()
                cmd.ExecuteNonQuery()
                con.Close()
                clear_form()
                getData()
                MessageBox.Show("Successfully edited")
            End Using
        End Using
    End Sub
End Class
