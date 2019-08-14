Imports System.IO
Imports System.Text

'Este programa funciona de la siguiente forma
'Creamos un vector v que nos va a servir para crear nuestra clave basandonos en la codificacion ASCII
'El vector va a ser tan grande como el usuario ponga en el programa--> ((Int(TextBox1.Text)) - 1)

Public Class Form1
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        tbDigitos.Text = "10"
        tbClave.ReadOnly = True
        cbLetras.Checked = True
        btnGuardar.Enabled = False
        rbMayusculas.Checked = True
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btGenerar.Click
        Dim i As Integer
        Randomize()
        'Si no ponemos un numero, salta el mensaje.
        If (tbDigitos.Text = "") Then
            MessageBox.Show("Tienes que escribir una cifra entre 1 y 99")
        Else
            'Si no seleccionamos ninguna opcion, salta:
            If (cbLetras.Checked = False And cbNumeros.Checked = False And cbCaracteres.Checked = False) Then
                MessageBox.Show("Seleccione al menos una opcion")
            Else
                btnCopy.Enabled = True
                Dim v(((Int(tbDigitos.Text)) - 1)) As Integer
                tbClave.Text = ""

                'Esto nos borra toda la matriz a 0 para no tener problemas si queremos crear mas claves.
                v.Initialize()

                'Clave con letras en mayusculas
                If (cbLetras.Checked = True And cbNumeros.Checked = False And cbCaracteres.Checked = False And rbMayusculas.Checked) Then
                    v = Funciones.letrasMayus(Int(tbDigitos.Text) - 1, v)
                End If

                'Clave con letras en minusculas
                If (cbLetras.Checked = True And cbNumeros.Checked = False And cbCaracteres.Checked = False And rbMinusculas.Checked = True) Then
                    v = Funciones.letrasMinusculas(Int(tbDigitos.Text) - 1, v)
                End If

                'Clave con letras en mayusculas y minusculas
                If (cbLetras.Checked = True And cbNumeros.Checked = False And cbCaracteres.Checked = False And rbAmbas.Checked = True) Then
                    v = Funciones.letrasAmbas(Int(tbDigitos.Text) - 1, v)
                End If

                'Clave con numeros
                If (cbLetras.Checked = False And cbNumeros.Checked = True And cbCaracteres.Checked = False) Then
                    v = Funciones.numeros(Int(tbDigitos.Text) - 1, v)
                End If

                'Clave con numeros y letras en mayusculas
                If (cbLetras.Checked = True And cbNumeros.Checked = True And cbCaracteres.Checked = False And rbMayusculas.Checked = True) Then
                    v = Funciones.numerosLetraMay(Int(tbDigitos.Text) - 1, v)
                End If

                'Clave con numeros y letras en minusculas
                If (cbLetras.Checked = True And cbNumeros.Checked = True And cbCaracteres.Checked = False And rbMinusculas.Checked = True) Then
                    v = Funciones.numerosLetraMin(Int(tbDigitos.Text) - 1, v)
                End If

                'Clave con numeros y ambos tipos de letras
                If (cbLetras.Checked = True And cbNumeros.Checked = True And cbCaracteres.Checked = False And rbAmbas.Checked = True) Then
                    v = Funciones.numerosLetraAmbas(Int(tbDigitos.Text) - 1, v)
                End If
                'Clave con caracteres
                If (cbLetras.Checked = False And cbNumeros.Checked = False And cbCaracteres.Checked = True) Then
                    v = Funciones.caracteres(Int(tbDigitos.Text) - 1, v)
                End If

                'Clave con caracteres y letras mayusculas
                If (cbLetras.Checked = True And cbNumeros.Checked = False And cbCaracteres.Checked = True And rbMayusculas.Checked = True) Then
                    v = Funciones.caracteresLetrasMayus(Int(tbDigitos.Text) - 1, v)
                End If

                'Clave con caracteres y letras minusculas
                If (cbLetras.Checked = True And cbNumeros.Checked = False And cbCaracteres.Checked = True And rbMinusculas.Checked = True) Then
                    v = Funciones.caracteresLetrasMinus(Int(tbDigitos.Text) - 1, v)
                End If

                'Clave con caracteres y letras ambas
                If (cbLetras.Checked = True And cbNumeros.Checked = False And cbCaracteres.Checked = True And rbAmbas.Checked = True) Then
                    v = Funciones.caracteresLetrasAmbas(Int(tbDigitos.Text) - 1, v)
                End If

                'Clave con caracteres y numeros
                If (cbLetras.Checked = False And cbNumeros.Checked = True And cbCaracteres.Checked = True) Then
                    v = Funciones.caracteresNumeros(Int(tbDigitos.Text) - 1, v)
                End If

                'Clave con caracteres, numeros y letras mayusculas
                If (cbLetras.Checked = True And cbNumeros.Checked = True And cbCaracteres.Checked = True And rbMayusculas.Checked = True) Then
                    v = Funciones.caracteresNumerosLetrasMayus(Int(tbDigitos.Text) - 1, v)
                End If

                'Clave con caracteres, numeros y letras minusculas
                If (cbLetras.Checked = True And cbNumeros.Checked = True And cbCaracteres.Checked = True And rbMinusculas.Checked = True) Then
                    v = Funciones.caracteresNumerosLetrasMinus(Int(tbDigitos.Text) - 1, v)
                End If

                'Clave con caracteres, numeros y letras ambas
                If (cbLetras.Checked = True And cbNumeros.Checked = True And cbCaracteres.Checked = True And rbAmbas.Checked = True) Then
                    v = Funciones.caracteresNumerosLetrasAmbas(Int(tbDigitos.Text) - 1, v)
                End If

                'Generamos la clave
                For i = 0 To (Int(tbDigitos.Text) - 1)
                    tbClave.Text &= Chr(v(i))
                Next
            End If
        End If
        If (tbClave.Text = "288") Then
            MsgBox("Magic number")
        End If
        btnGuardar.Enabled = True
        tbTitulo.Text = ""
        tbUsuario.Text = ""
        tbDescripcion.Text = ""
    End Sub
    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbDigitos.KeyPress
        If InStr(1, "0123456789" & Chr(8), e.KeyChar) = 0 Then
            e.KeyChar = ""
        End If
    End Sub
    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Dim ruta As String = ""
        Dim caracteres As Integer
        caracteres = tbTitulo.TextLength
        Dim sep As String = StrDup(caracteres, "=")
        sfdArchivo.Filter = "Documento de texto (*.txt)|*.txt"
        sfdArchivo.ShowDialog()
        ruta = sfdArchivo.FileName
        If (ruta <> "") Then
            Dim fs As FileStream = File.Create(ruta)
            Dim info As Byte() = New UTF8Encoding(True).GetBytes(tbTitulo.Text + vbCrLf + sep + vbCrLf + "Cuenta:" + vbTab + tbUsuario.Text + vbCrLf + "Pass:" + vbTab + tbClave.Text + vbCrLf + vbCrLf + "Descripcion:" + vbCrLf + tbDescripcion.Text)
            fs.Write(info, 0, info.Length)
            fs.Close()
        End If
    End Sub

    Private Sub version_Click(sender As Object, e As EventArgs) Handles version.Click
        MessageBox.Show("Build by Wirfen")
    End Sub

    Private Sub btnCopy_Click(sender As Object, e As EventArgs) Handles btnCopy.Click
        My.Computer.Clipboard.SetText(tbClave.Text)
    End Sub
End Class
