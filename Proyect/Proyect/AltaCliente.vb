Imports Entidad
Imports Negocio
Imports System.Windows.Forms
Public Class AltaCliente
    Dim registro As Integer
    Dim objCliente As New NCliente
    Dim objProvincia As New NProvincia


    Private Sub txtApellido_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtApellido.KeyPress
        Validacion.SoloLetras(txtApellido.Text, e)
    End Sub

    Private Sub btnAtras_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAtras.Click

        Close()

    End Sub


    Private Sub btnAgregar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAgregar.Click
        If Trim(txtApellido.Text) = "" Or Trim(txtNombre.Text) = "" Or Trim(txtDireccion.Text) = "" Or Trim(txtDNI.Text) = "" Or Trim(txtEmail.Text) = "" Or Trim(txtTelefono.Text) = "" Or Trim(CmbProvincia.Text) = "" Then
            MsgBox("Debe completar todos los campos", MsgBoxStyle.Critical, "Advertencia"
                   )

        ElseIf validarEmail(txtEmail.Text) = False Then
            MsgBox("El EMAIL no cumple con el formato especificado", MsgBoxStyle.Exclamation, "Advertencia")
            txtEmail.Focus()
            txtEmail.BackColor = Color.Coral
        ElseIf txtDNI.TextLength < 8 Then
            MsgBox("El DNI no cumple con el formato especificado (8 digitos)", MsgBoxStyle.Exclamation, "Advertencia")
            txtDNI.Focus()
            txtDNI.BackColor = Color.Coral
        ElseIf objCliente.getClienteVerificarDNI(txtDNI.Text) Then
            Dim ecliente As ECliente = New ECliente
            With ecliente
                .Dni = txtDNI.Text
                .Nombre = txtNombre.Text
                .Apellido = txtApellido.Text
                .Telefono = txtTelefono.Text
                .Direccion = txtDireccion.Text
                .Email = txtEmail.Text
                .Provincia = CmbProvincia.SelectedValue
                .Estado = 1
            End With

            If objCliente.guardarCliente(ecliente) Then
                MessageBox.Show("Cliente Agregado Correctamente", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
                cargarDGTodos()
                txtDNI.BackColor = Color.White
                txtEmail.BackColor = Color.White
                txtDNI.Clear()
                txtNombre.Clear()
                txtApellido.Clear()
                txtDireccion.Clear()
                txtEmail.Clear()
                txtTelefono.Clear()
                CmbProvincia.ResetText()
                CmbProvincia.SelectedText = ""
            Else
                MessageBox.Show("Error al agregar el Cliente", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtDNI.Focus()
                txtDNI.BackColor = Color.Coral
            End If
            'Else
            '    MessageBox.Show("Error", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

        ''Close()

    End Sub

    Private Sub txtNombre_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtNombre.KeyPress
        Validacion.SoloLetras(txtNombre.Text, e)
    End Sub

    Private Sub txtDNI_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDNI.KeyPress
        If txtDNI.TextLength < 8 Then
            If Not (Char.IsDigit(e.KeyChar) Or Char.IsControl(e.KeyChar)) Then
                e.Handled = True
            Else
                e.Handled = False
            End If
        ElseIf (Char.IsControl(e.KeyChar)) Then
            e.Handled = False
        Else
            e.Handled = True
            MsgBox("Formato Incorrecto de DNI (8 digitos)", MsgBoxStyle.Exclamation, "Advertencia")
        End If
    End Sub

    Private Sub txtTelefono_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTelefono.KeyPress
        Validacion.SoloNumeros(txtTelefono.Text, e)
    End Sub


    Private Sub frmAltaCliente_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Limpiar(Me)
        'cargarDG()
        cargarDGTodos()
        cargarCBO()
        txtDNI.Focus()
    End Sub


    Sub cargarDG()

        dvgDatos.DataSource = objCliente.getClienteTodos

    End Sub

    Sub cargarCBO()

        With CmbProvincia
            .DataSource = objProvincia.getProvincia
            .DisplayMember = "descripcion"
            .ValueMember = "provincia"
            .SelectedIndex = -1
        End With

    End Sub

    Sub cargarDGTodos()

        dvgDatos.DataSource = objCliente.getClienteTodos

        For Each Row As DataGridViewRow In (dvgDatos.Rows)

            If Row.Cells("estado").Value = 0 Then
                Row.DefaultCellStyle.BackColor = Color.Red
            End If

        Next

        Me.dvgDatos.Columns("estado").Visible = False

        dvgDatos.ClearSelection()

    End Sub


    Private Sub txtApellido_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtApellido.TextChanged
        txtApellido.Text = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(txtApellido.Text)
        txtApellido.SelectionStart = txtApellido.Text.Length
    End Sub

    Private Sub txtNombre_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNombre.TextChanged
        txtNombre.Text = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(txtNombre.Text)
        txtNombre.SelectionStart = txtNombre.Text.Length
    End Sub

    Private Sub txtDireccion_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDireccion.TextChanged
        txtDireccion.Text = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(txtDireccion.Text)
        txtDireccion.SelectionStart = txtDireccion.Text.Length
    End Sub


    Private Sub Click_boton(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dvgDatos.CellClick
        btnModificar.Enabled = True
        btnActualizar.Enabled = True
        btnBorrar.Enabled = True
        btnAgregar.Enabled = False

    End Sub

    Private Sub btnModificar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnModificar.Click
        registro = 0
        registro = Me.dvgDatos.CurrentRow.Index
        Dim estado As Integer = Me.dvgDatos.Rows(registro).Cells(7).Value

        If estado = 0 Then

            MsgBox("El producto ya se encuentra dado de baja")

            txtDNI.Enabled = True
            txtDNI.Focus()
            txtDNI.Clear()
            txtNombre.Clear()
            txtApellido.Clear()
            txtDireccion.Clear()
            txtEmail.Clear()
            txtTelefono.Clear()
            CmbProvincia.SelectedIndex = -1
            cmbEstado.SelectedIndex = -1

            btnModificar.Enabled = False
            btnActualizar.Enabled = False
            btnBorrar.Enabled = False
            btnAgregar.Enabled = True

            Exit Sub
        Else

            txtDNI.Enabled = False
            Dim Fila As Integer = Me.dvgDatos.CurrentRow.Index
            txtDNI.Text = Me.dvgDatos.Rows(Fila).Cells(0).Value
            txtNombre.Text = dvgDatos.Rows(Fila).Cells(1).Value
            txtApellido.Text = Me.dvgDatos.Rows(Fila).Cells(2).Value
            txtTelefono.Text = Me.dvgDatos.Rows(Fila).Cells(3).Value
            txtDireccion.Text = Me.dvgDatos.Rows(Fila).Cells(4).Value
            txtEmail.Text = Me.dvgDatos.Rows(Fila).Cells(5).Value
            CmbProvincia.Text = Me.dvgDatos.Rows(Fila).Cells(6).Value

        End If
    End Sub

    Private Sub btnActualizar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnActualizar.Click
        If Trim(txtApellido.Text) = "" Or Trim(txtNombre.Text) = "" Or Trim(txtDireccion.Text) = "" Or Trim(txtDNI.Text) = "" Or Trim(txtEmail.Text) = "" Or Trim(txtTelefono.Text) = "" Or Trim(CmbProvincia.Text) = "" Then
            MsgBox("Debe completar todos los campos", MsgBoxStyle.Critical, "Advertencia"
                   )
        ElseIf validarEmail(txtEmail.Text) = False Then
            MsgBox("El EMAIL no cumple con el formato especificado", MsgBoxStyle.Exclamation, "Advertencia")
            txtEmail.Focus()
            txtEmail.BackColor = Color.Coral
        Else
            Dim ecliente As ECliente = New ECliente
            With ecliente
                .Dni = txtDNI.Text
                .Nombre = txtNombre.Text
                .Apellido = txtApellido.Text
                .Telefono = txtTelefono.Text
                .Direccion = txtDireccion.Text
                .Email = txtEmail.Text
                .Provincia = CmbProvincia.SelectedValue
                .Estado = 1
            End With

            If objCliente.ActualizarCliente(ecliente) Then
                MessageBox.Show("Cliente Actualizado Correctamente", "", MessageBoxButtons.OK, MessageBoxIcon.Information)

                cargarDGTodos()
                txtDNI.Enabled = True
                txtEmail.BackColor = Color.White
                txtDNI.Focus()
                txtDNI.Clear()
                txtNombre.Clear()
                txtApellido.Clear()
                txtDireccion.Clear()
                txtEmail.Clear()
                txtTelefono.Clear()
                CmbProvincia.ResetText()
                CmbProvincia.SelectedText = ""

                btnModificar.Enabled = False
                btnActualizar.Enabled = False
                btnBorrar.Enabled = False
                btnAgregar.Enabled = True

            Else
                MessageBox.Show("Error al actualizar Cliente", "", MessageBoxButtons.OK, MessageBoxIcon.Error)

            End If
        End If

    End Sub

    Private Sub cargarDatos()

        registro = Me.dvgDatos.CurrentRow.Index

        Dim dni As String = Me.dvgDatos.Rows(registro).Cells(0).Value
        Dim nombre As String = Me.dvgDatos.Rows(registro).Cells(1).Value
        Dim apellido As String = Me.dvgDatos.Rows(registro).Cells(2).Value
        Dim telefono As Integer = Me.dvgDatos.Rows(registro).Cells(3).Value
        Dim direccion As String = Me.dvgDatos.Rows(registro).Cells(4).Value
        Dim email As String = Me.dvgDatos.Rows(registro).Cells(5).Value


        Dim provincia As String = Me.dvgDatos.Rows(registro).Cells(6).Value
        Dim estado As Integer = Me.dvgDatos.Rows(registro).Cells(7).Value

        txtDNI.Text = dni
        txtNombre.Text = nombre
        txtApellido.Text = apellido
        txtTelefono.Text = telefono
        txtDireccion.Text = direccion
        txtEmail.Text = email

        CmbProvincia.Text = provincia

        cmbEstado.SelectedIndex = estado


    End Sub



    Private Sub btnBorrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBorrar.Click

        cargarDatos()
        Dim Fila As Integer = Me.dvgDatos.CurrentRow.Index
        Dim estado As Integer = Me.dvgDatos.Rows(Fila).Cells(7).Value

        If estado = 0 Then

            MsgBox("El producto ya se encuentra dado de baja")

            txtDNI.Enabled = True
            txtDNI.Focus()
            txtDNI.Clear()
            txtNombre.Clear()
            txtApellido.Clear()
            txtDireccion.Clear()
            txtEmail.Clear()
            txtTelefono.Clear()
            CmbProvincia.SelectedIndex = -1
            cmbEstado.SelectedIndex = -1

            btnModificar.Enabled = False
            btnActualizar.Enabled = False
            btnBorrar.Enabled = False
            btnAgregar.Enabled = True
            Exit Sub

        End If


        Dim ask As MsgBoxResult

        ask = MsgBox("Los datos serán eliminados. ¿Está seguro?", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 + MsgBoxStyle.Information, "Borrar registro")
        If MsgBoxResult.Yes = ask Then
            objCliente.borrarCliente(txtDNI.Text)
            cargarDGTodos()
            dvgDatos.Rows(registro).Selected = True
        End If
        txtDNI.Enabled = True
        txtDNI.Focus()
        txtDNI.Clear()
        txtNombre.Clear()
        txtApellido.Clear()
        txtDireccion.Clear()
        txtEmail.Clear()
        txtTelefono.Clear()
        CmbProvincia.SelectedIndex = -1
        cmbEstado.SelectedIndex = -1

        btnModificar.Enabled = False
        btnActualizar.Enabled = False
        btnBorrar.Enabled = False
        btnAgregar.Enabled = True

    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub cargarTextbox(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dvgDatos.CellContentDoubleClick


        Dim Fila As Integer = Me.dvgDatos.CurrentRow.Index


        Dim estado As Integer = Me.dvgDatos.Rows(Fila).Cells(7).Value

        If estado = 0 Then

            MsgBox("El producto ya se encuentra dado de baja")
        Else
            frmVentaProducto.txtDniCliente.Text = Me.dvgDatos.Rows(Fila).Cells(0).Value
            Dim aux As String = Me.dvgDatos.Rows(Fila).Cells(1).Value & " " & Me.dvgDatos.Rows(Fila).Cells(2).Value
            frmVentaProducto.txtNomCliente.Text = aux
            Close()
        End If
    End Sub

End Class