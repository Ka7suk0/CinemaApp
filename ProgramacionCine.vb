' This file is not excecutable, it's just a placeholder to make GitHub recognize some code and set this project as coded in Visual Basic.
' The original project (Cine.xlsm) was built entirely using Visual Basic For Excel, and thus it is a shortened version of language embedded into .xlms files, GitHub does not recognizes it.
' Cheers!


ACCARTELERA
Private Sub Btn_AcCartelera_Click()
Dim Vacío As Boolean

If txt_Película1 = "" Or txt_Película2 = "" Or txt_Película3 = "" Or txt_Película4 = "" Or txt_Película5 = "" Or txt_Película6 = "" Or txt_Estreno1 = "" Or txt_Estreno2 = "" Or txt_Estreno3 = "" Or txt_Estreno4 = "" Or txt_Estreno5 = "" Or txt_Estreno6 = "" Or txt_Eliminación1 = "" Or txt_Eliminación2 = "" Or txt_Eliminación3 = "" Or txt_Eliminación4 = "" Or txt_Eliminación5 = "" Or txt_Eliminación6 = "" Then
Vacío = True
End If

If Vacío = True Then
    MsgBox ("Por favor, complete la cartelera para realizar los cambios.")
End If

If Vacío = False Then
    Worksheets("Películas").Range("B2").Value = txt_Película1
    Worksheets("Películas").Range("B3").Value = txt_Película2
    Worksheets("Películas").Range("B4").Value = txt_Película3
    Worksheets("Películas").Range("B5").Value = txt_Película4
    Worksheets("Películas").Range("B6").Value = txt_Película5
    Worksheets("Películas").Range("B7").Value = txt_Película6
    Worksheets("Películas").Range("C2").Value = txt_Estreno1
    Worksheets("Películas").Range("C3").Value = txt_Estreno2
    Worksheets("Películas").Range("C4").Value = txt_Estreno3
    Worksheets("Películas").Range("C5").Value = txt_Estreno4
    Worksheets("Películas").Range("C6").Value = txt_Estreno5
    Worksheets("Películas").Range("C7").Value = txt_Estreno6
    Worksheets("Películas").Range("D2").Value = txt_Eliminación1
    Worksheets("Películas").Range("D3").Value = txt_Eliminación2
    Worksheets("Películas").Range("D4").Value = txt_Eliminación3
    Worksheets("Películas").Range("D5").Value = txt_Eliminación4
    Worksheets("Películas").Range("D6").Value = txt_Eliminación5
    Worksheets("Películas").Range("D7").Value = txt_Eliminación6
End If
End Sub

Private Sub Btn_Cancelar_Click()
    ACCARTELERA.Hide
    ADMINISTRADOR.Show
End Sub

Private Sub Btn_CerrarSesión_Click()
    ACCARTELERA.Hide
    INICIO.Show
End Sub

Private Sub txt_Eliminación1_Change()
Dim Estreno11 As Date
Dim Eliminación11 As Date
Dim FechaH As Date

If txt_Estreno1 <> "" And txt_Eliminación1 <> "" Then
Estreno11 = txt_Estreno1
Eliminación11 = txt_Eliminación1
FechaH = Worksheets("Películas").Range("K2").Value

If Estreno11 > FechaH Then
    txt_Disponibilidad1 = "PRÓXIMAMENTE"
Else
    If Eliminación11 <= FechaH Then
        txt_Disponibilidad1 = "NO DISPONIBLE"
    Else
        txt_Disponibilidad1 = "PRESENTANDO"
    End If
End If
End If
End Sub

Private Sub txt_Eliminación2_Change()
Dim Estreno22 As Date
Dim Eliminación22 As Date
Dim FechaH As Date

If txt_Estreno2 <> "" And txt_Eliminación2 <> "" Then
Estreno22 = txt_Estreno2
Eliminación22 = txt_Eliminación2
FechaH = Worksheets("Películas").Range("K2").Value

If Estreno22 > FechaH Then
    txt_Disponibilidad2 = "PRÓXIMAMENTE"
Else
    If Eliminación22 <= FechaH Then
        txt_Disponibilidad2 = "NO DISPONIBLE"
    Else
        txt_Disponibilidad2 = "PRESENTANDO"
    End If
End If
End If
End Sub

Private Sub txt_Eliminación3_Change()
Dim Estreno33 As Date
Dim Eliminación33 As Date
Dim FechaH As Date

If txt_Estreno3 <> "" And txt_Eliminación3 <> "" Then
Estreno33 = txt_Estreno3
Eliminación33 = txt_Eliminación3
FechaH = Worksheets("Películas").Range("K2").Value

If Estreno33 > FechaH Then
    txt_Disponibilidad3 = "PRÓXIMAMENTE"
Else
    If Eliminación33 <= FechaH Then
        txt_Disponibilidad3 = "NO DISPONIBLE"
    Else
        txt_Disponibilidad3 = "PRESENTANDO"
    End If
End If
End If
End Sub

Private Sub txt_Eliminación4_Change()
Dim Estreno44 As Date
Dim Eliminación44 As Date
Dim FechaH As Date

If txt_Estreno4 <> "" And txt_Eliminación4 <> "" Then
Estreno44 = txt_Estreno4
Eliminación44 = txt_Eliminación4
FechaH = Worksheets("Películas").Range("K2").Value

If Estreno44 > FechaH Then
    txt_Disponibilidad4 = "PRÓXIMAMENTE"
Else
    If Eliminación44 <= FechaH Then
        txt_Disponibilidad4 = "NO DISPONIBLE"
    Else
        txt_Disponibilidad4 = "PRESENTANDO"
    End If
End If
End If
End Sub

Private Sub txt_Eliminación5_Change()
Dim Estreno55 As Date
Dim Eliminación55 As Date
Dim FechaH As Date

If txt_Estreno5 <> "" And txt_Eliminación5 <> "" Then
Estreno55 = txt_Estreno5
Eliminación55 = txt_Eliminación5
FechaH = Worksheets("Películas").Range("K2").Value

If Estreno55 > FechaH Then
    txt_Disponibilidad5 = "PRÓXIMAMENTE"
Else
    If Eliminación55 <= FechaH Then
        txt_Disponibilidad5 = "NO DISPONIBLE"
    Else
        txt_Disponibilidad5 = "PRESENTANDO"
    End If
End If
End If
End Sub

Private Sub txt_Eliminación6_Change()
Dim Estreno66 As Date
Dim Eliminación66 As Date
Dim FechaH As Date

If txt_Estreno6 <> "" And txt_Eliminación6 <> "" Then
Estreno66 = txt_Estreno6
Eliminación66 = txt_Eliminación6
FechaH = Worksheets("Películas").Range("K2").Value

If Estreno66 > FechaH Then
    txt_Disponibilidad6 = "PRÓXIMAMENTE"
Else
    If Eliminación66 <= FechaH Then
        txt_Disponibilidad6 = "NO DISPONIBLE"
    Else
        txt_Disponibilidad6 = "PRESENTANDO"
    End If
End If
End If
End Sub


Private Sub txt_Estreno1_Change()
Dim Estreno11 As Date
Dim Eliminación11 As Date
Dim FechaH As Date

If txt_Estreno1 <> "" And txt_Eliminación1 <> "" Then
Estreno11 = txt_Estreno1
Eliminación11 = txt_Eliminación1
FechaH = Worksheets("Películas").Range("K2").Value

If Estreno11 > FechaH Then
    txt_Disponibilidad1 = "PRÓXIMAMENTE"
Else
    If Eliminación11 <= FechaH Then
        txt_Disponibilidad1 = "NO DISPONIBLE"
    Else
        txt_Disponibilidad1 = "PRESENTANDO"
    End If
End If
End If
End Sub

Private Sub txt_Estreno2_Change()
Dim Estreno22 As Date
Dim Eliminación22 As Date
Dim FechaH As Date

If txt_Estreno2 <> "" And txt_Eliminación2 <> "" Then
Estreno22 = txt_Estreno2
Eliminación22 = txt_Eliminación2
FechaH = Worksheets("Películas").Range("K2").Value

If Estreno22 > FechaH Then
    txt_Disponibilidad2 = "PRÓXIMAMENTE"
Else
    If Eliminación22 <= FechaH Then
        txt_Disponibilidad2 = "NO DISPONIBLE"
    Else
        txt_Disponibilidad2 = "PRESENTANDO"
    End If
End If
End If
End Sub

Private Sub txt_Estreno3_Change()
Dim Estreno33 As Date
Dim Eliminación33 As Date
Dim FechaH As Date

If txt_Estreno3 <> "" And txt_Eliminación3 <> "" Then
Estreno33 = txt_Estreno3
Eliminación33 = txt_Eliminación3
FechaH = Worksheets("Películas").Range("K2").Value

If Estreno33 > FechaH Then
    txt_Disponibilidad3 = "PRÓXIMAMENTE"
Else
    If Eliminación33 <= FechaH Then
        txt_Disponibilidad3 = "NO DISPONIBLE"
    Else
        txt_Disponibilidad3 = "PRESENTANDO"
    End If
End If
End If
End Sub

Private Sub txt_Estreno4_Change()
Dim Estreno44 As Date
Dim Eliminación44 As Date
Dim FechaH As Date

If txt_Estreno4 <> "" And txt_Eliminación4 <> "" Then
Estreno44 = txt_Estreno4
Eliminación44 = txt_Eliminación4
FechaH = Worksheets("Películas").Range("K2").Value

If Estreno44 > FechaH Then
    txt_Disponibilidad4 = "PRÓXIMAMENTE"
Else
    If Eliminación44 <= FechaH Then
        txt_Disponibilidad4 = "NO DISPONIBLE"
    Else
        txt_Disponibilidad4 = "PRESENTANDO"
    End If
End If
End If
End Sub

Private Sub txt_Estreno5_Change()
Dim Estreno55 As Date
Dim Eliminación55 As Date
Dim FechaH As Date

If txt_Estreno5 <> "" And txt_Eliminación5 <> "" Then
Estreno55 = txt_Estreno5
Eliminación55 = txt_Eliminación5
FechaH = Worksheets("Películas").Range("K2").Value

If Estreno55 > FechaH Then
    txt_Disponibilidad5 = "PRÓXIMAMENTE"
Else
    If Eliminación55 <= FechaH Then
        txt_Disponibilidad5 = "NO DISPONIBLE"
    Else
        txt_Disponibilidad5 = "PRESENTANDO"
    End If
End If
End If
End Sub

Private Sub txt_Estreno6_Change()
Dim Estreno66 As Date
Dim Eliminación66 As Date
Dim FechaH As Date

If txt_Estreno6 <> "" And txt_Eliminación6 <> "" Then
Estreno66 = txt_Estreno6
Eliminación66 = txt_Eliminación6
FechaH = Worksheets("Películas").Range("K2").Value

If Estreno66 > FechaH Then
    txt_Disponibilidad6 = "PRÓXIMAMENTE"
Else
    If Eliminación66 <= FechaH Then
        txt_Disponibilidad6 = "NO DISPONIBLE"
    Else
        txt_Disponibilidad6 = "PRESENTANDO"
    End If
End If
End If
End Sub



Private Sub UserForm_Activate()
txt_Estreno1 = Format(txt_Estreno1, "dd/mm/yyyy")
txt_Estreno2 = Format(txt_Estreno2, "dd/mm/yyyy")
txt_Estreno3 = Format(txt_Estreno3, "dd/mm/yyyy")
txt_Estreno4 = Format(txt_Estreno4, "dd/mm/yyyy")
txt_Estreno5 = Format(txt_Estreno5, "dd/mm/yyyy")
txt_Estreno6 = Format(txt_Estreno6, "dd/mm/yyyy")
txt_Eliminación1 = Format(txt_Eliminación1, "dd/mm/yyyy")
txt_Eliminación2 = Format(txt_Eliminación2, "dd/mm/yyyy")
txt_Eliminación3 = Format(txt_Eliminación3, "dd/mm/yyyy")
txt_Eliminación4 = Format(txt_Eliminación4, "dd/mm/yyyy")
txt_Eliminación5 = Format(txt_Eliminación5, "dd/mm/yyyy")
txt_Eliminación6 = Format(txt_Eliminación6, "dd/mm/yyyy")
    txt_Película1 = Worksheets("Películas").Range("B2").Value
    txt_Película2 = Worksheets("Películas").Range("B3").Value
    txt_Película3 = Worksheets("Películas").Range("B4").Value
    txt_Película4 = Worksheets("Películas").Range("B5").Value
    txt_Película5 = Worksheets("Películas").Range("B6").Value
    txt_Película6 = Worksheets("Películas").Range("B7").Value
    txt_Estreno1 = Worksheets("Películas").Range("C2").Value
    txt_Estreno2 = Worksheets("Películas").Range("C3").Value
    txt_Estreno3 = Worksheets("Películas").Range("C4").Value
    txt_Estreno4 = Worksheets("Películas").Range("C5").Value
    txt_Estreno5 = Worksheets("Películas").Range("C6").Value
    txt_Estreno6 = Worksheets("Películas").Range("C7").Value
    txt_Eliminación1 = Worksheets("Películas").Range("D2").Value
    txt_Eliminación2 = Worksheets("Películas").Range("D3").Value
    txt_Eliminación3 = Worksheets("Películas").Range("D4").Value
    txt_Eliminación4 = Worksheets("Películas").Range("D5").Value
    txt_Eliminación5 = Worksheets("Películas").Range("D6").Value
    txt_Eliminación6 = Worksheets("Películas").Range("D7").Value

Dim Estreno1 As Integer
Dim Estreno2 As Integer
Dim Estreno3 As Integer
Dim Estreno4 As Integer
Dim Estreno5 As Integer
Dim Estreno6 As Integer
Dim Eliminación1 As Integer
Dim Eliminación2 As Integer
Dim Eliminación3 As Integer
Dim Eliminación4 As Integer
Dim Eliminación5 As Integer
Dim Eliminación6 As Integer

Estreno1 = Worksheets("Películas").Range("E2").Value
Estreno2 = Worksheets("Películas").Range("E3").Value
Estreno3 = Worksheets("Películas").Range("E4").Value
Estreno4 = Worksheets("Películas").Range("E5").Value
Estreno5 = Worksheets("Películas").Range("E6").Value
Estreno6 = Worksheets("Películas").Range("E7").Value
Eliminación1 = Worksheets("Películas").Range("F2").Value
Eliminación2 = Worksheets("Películas").Range("F3").Value
Eliminación3 = Worksheets("Películas").Range("F4").Value
Eliminación4 = Worksheets("Películas").Range("F5").Value
Eliminación5 = Worksheets("Películas").Range("F6").Value
Eliminación6 = Worksheets("Películas").Range("F7").Value

If Estreno1 > 0 Then
    txt_Disponibilidad1 = "PRÓXIMAMENTE"
Else
    If Eliminación1 <= 0 Then
        txt_Disponibilidad1 = "NO DISPONIBLE"
    Else
        txt_Disponibilidad1 = "PRESENTANDO"
    End If
End If

If Estreno2 > 0 Then
    txt_Disponibilidad2 = "PRÓXIMAMENTE"
Else
    If Eliminación2 <= 0 Then
        txt_Disponibilidad2 = "NO DISPONIBLE"
    Else
        txt_Disponibilidad2 = "PRESENTANDO"
    End If
End If

If Estreno3 > 0 Then
    txt_Disponibilidad3 = "PRÓXIMAMENTE"
Else
    If Eliminación3 <= 0 Then
        txt_Disponibilidad3 = "NO DISPONIBLE"
    Else
        txt_Disponibilidad3 = "PRESENTANDO"
    End If
End If

If Estreno4 > 0 Then
    txt_Disponibilidad4 = "PRÓXIMAMENTE"
Else
    If Eliminación4 <= 0 Then
        txt_Disponibilidad4 = "NO DISPONIBLE"
    Else
        txt_Disponibilidad4 = "PRESENTANDO"
    End If
End If

If Estreno5 > 0 Then
    txt_Disponibilidad5 = "PRÓXIMAMENTE"
Else
    If Eliminación5 <= 0 Then
        txt_Disponibilidad5 = "NO DISPONIBLE"
    Else
        txt_Disponibilidad5 = "PRESENTANDO"
    End If
End If

If Estreno6 > 0 Then
    txt_Disponibilidad6 = "PRÓXIMAMENTE"
Else
    If Eliminación6 <= 0 Then
        txt_Disponibilidad6 = "NO DISPONIBLE"
    Else
        txt_Disponibilidad6 = "PRESENTANDO"
    End If
End If
End Sub
ACTUALIZARDATOS
Private Sub Btn_Aceptar_Click()
    Dim Duplicado As Boolean
    Duplicado = False
    Dim Vacío As Boolean
    Vacío = False
    
    If txt_Correo.Value = "" Then
        Vacío = True
    End If
    If CB_Pregunta = "" Then
        Vacío = True
    End If
    If CB_Respuesta = "" Then
        Vacío = True
    End If
    
    If Vacío = True Then
        MsgBox ("Por favor, completa el formulario. Recuerda que todos los campos son obligatorios.")
    End If
    
Dim c As Double

    For c = 2 To 9999 Step 1
    If Worksheets("Usuarios").Range("B" & c).Value = txt_Correo Then
        MsgBox "Ya existe una cuenta con ese correo."
        Duplicado = True
    End If
    Next c
    
    
Dim User As String
Dim UserOr As String
Dim Reng As Double

If Duplicado = False Then
    If Vacío = False Then
        User = Worksheets("Boletos").Range("M2").Value
        For Reng = 2 To 9999 Step 1
            UserOr = Worksheets("Usuarios").Range("C" & Reng).Value
                If UserOr = User Then
                    Worksheets("Usuarios").Range("B" & Reng).Value = txt_Correo
                    Worksheets("Usuarios").Range("J" & Reng).Value = CB_Pregunta
                    Worksheets("Usuarios").Range("K" & Reng).Value = CB_Respuesta
                    MsgBox ("Los datos han sido actualizados exitosamente.")
                    ACTUALIZARDATOS.Hide
                    CONFIGURACIÓN.Show
                End If
        Next Reng
    End If
End If
End Sub

Private Sub Btn_Cancelar_Click()
    ACTUALIZARDATOS.Hide
    CONFIGURACIÓN.Show
End Sub


Private Sub CB_Pregunta_Change()
    If CB_Pregunta.Text = "¿Cuál es tu color favorito?" Then
        CB_Respuesta.Clear
        CB_Respuesta.AddItem "Amarillo"
        CB_Respuesta.AddItem "Azul"
        CB_Respuesta.AddItem "Blanco"
        CB_Respuesta.AddItem "Café"
        CB_Respuesta.AddItem "Gris"
        CB_Respuesta.AddItem "Morado"
        CB_Respuesta.AddItem "Naranja"
        CB_Respuesta.AddItem "Negro"
        CB_Respuesta.AddItem "Rojo"
        CB_Respuesta.AddItem "Rosa"
        CB_Respuesta.AddItem "Turqueza"
        CB_Respuesta.AddItem "Verde"
    End If
    If CB_Pregunta.Text = "¿Cuál es tu animal favorito?" Then
        CB_Respuesta.Clear
        CB_Respuesta.AddItem "Águila"
        CB_Respuesta.AddItem "Delfín"
        CB_Respuesta.AddItem "Elefante"
        CB_Respuesta.AddItem "Gato"
        CB_Respuesta.AddItem "Hámster"
        CB_Respuesta.AddItem "Koala"
        CB_Respuesta.AddItem "León"
        CB_Respuesta.AddItem "Lobo"
        CB_Respuesta.AddItem "Pato"
        CB_Respuesta.AddItem "Perro"
        CB_Respuesta.AddItem "Pingüino"
        CB_Respuesta.AddItem "Venado"
    End If
    If CB_Pregunta.Text = "¿Cuál es el mes de nacimiento de tu mamá?" Then
        CB_Respuesta.Clear
        CB_Respuesta.AddItem "Enero"
        CB_Respuesta.AddItem "Febrero"
        CB_Respuesta.AddItem "Marzo"
        CB_Respuesta.AddItem "Abril"
        CB_Respuesta.AddItem "Mayo"
        CB_Respuesta.AddItem "Junio"
        CB_Respuesta.AddItem "Julio"
        CB_Respuesta.AddItem "Agosto"
        CB_Respuesta.AddItem "Septiembre"
        CB_Respuesta.AddItem "Octubre"
        CB_Respuesta.AddItem "Noviembre"
        CB_Respuesta.AddItem "Diciembre"
    End If
    If CB_Pregunta.Text = "¿Cuál es tu comida favorita?" Then
        CB_Respuesta.Clear
        CB_Respuesta.AddItem "Asado"
        CB_Respuesta.AddItem "Burritos"
        CB_Respuesta.AddItem "Hamburguesas"
        CB_Respuesta.AddItem "Hot Dog"
        CB_Respuesta.AddItem "Lasaña"
        CB_Respuesta.AddItem "Mole"
        CB_Respuesta.AddItem "Paella"
        CB_Respuesta.AddItem "Pasta"
        CB_Respuesta.AddItem "Pizza"
        CB_Respuesta.AddItem "Sushi"
        CB_Respuesta.AddItem "Tacos"
        CB_Respuesta.AddItem "Tamales"
    End If
End Sub

Private Sub UserForm_Activate()
Dim Usuario As String
Dim UsuarioOr As String
Dim i As Double
Usuario = Worksheets("Boletos").Range("M2").Value

For i = 2 To 9999 Step 1
    UsuarioOr = Worksheets("Usuarios").Range("C" & i)
    If UsuarioOr = Usuario Then
        txt_Correo = Worksheets("Usuarios").Range("B" & i).Value
        CB_Pregunta = Worksheets("Usuarios").Range("J" & i).Value
        CB_Respuesta = Worksheets("Usuarios").Range("K" & i).Value
    End If
Next i

    

    CB_Pregunta.AddItem "¿Cuál es tu color favorito?"
    CB_Pregunta.AddItem "¿Cuál es tu animal favorito?"
    CB_Pregunta.AddItem "¿Cuál es el mes de nacimiento de tu mamá?"
    CB_Pregunta.AddItem "¿Cuál es tu comida favorita?"
    
    
    If CB_Pregunta.Text = "¿Cuál es tu color favorito?" Then
        CB_Respuesta.AddItem "Amarillo"
        CB_Respuesta.AddItem "Azul"
        CB_Respuesta.AddItem "Blanco"
        CB_Respuesta.AddItem "Café"
        CB_Respuesta.AddItem "Gris"
        CB_Respuesta.AddItem "Morado"
        CB_Respuesta.AddItem "Naranja"
        CB_Respuesta.AddItem "Negro"
        CB_Respuesta.AddItem "Rojo"
        CB_Respuesta.AddItem "Rosa"
        CB_Respuesta.AddItem "Turqueza"
        CB_Respuesta.AddItem "Verde"
    End If
    If CB_Pregunta.Text = "¿Cuál es tu animal favorito?" Then
        CB_Respuesta.AddItem "Águila"
        CB_Respuesta.AddItem "Delfín"
        CB_Respuesta.AddItem "Elefante"
        CB_Respuesta.AddItem "Gato"
        CB_Respuesta.AddItem "Hámster"
        CB_Respuesta.AddItem "Koala"
        CB_Respuesta.AddItem "León"
        CB_Respuesta.AddItem "Lobo"
        CB_Respuesta.AddItem "Pato"
        CB_Respuesta.AddItem "Perro"
        CB_Respuesta.AddItem "Pingüino"
        CB_Respuesta.AddItem "Venado"
    End If
    If CB_Pregunta.Text = "¿Cuál es el mes de nacimiento de tu mamá?" Then
        CB_Respuesta.AddItem "Enero"
        CB_Respuesta.AddItem "Febrero"
        CB_Respuesta.AddItem "Marzo"
        CB_Respuesta.AddItem "Abril"
        CB_Respuesta.AddItem "Mayo"
        CB_Respuesta.AddItem "Junio"
        CB_Respuesta.AddItem "Julio"
        CB_Respuesta.AddItem "Agosto"
        CB_Respuesta.AddItem "Septiembre"
        CB_Respuesta.AddItem "Octubre"
        CB_Respuesta.AddItem "Noviembre"
        CB_Respuesta.AddItem "Diciembre"
    End If
    If CB_Pregunta.Text = "¿Cuál es tu comida favorita?" Then
        CB_Respuesta.AddItem "Asado"
        CB_Respuesta.AddItem "Burritos"
        CB_Respuesta.AddItem "Hamburguesas"
        CB_Respuesta.AddItem "Hot Dog"
        CB_Respuesta.AddItem "Lasaña"
        CB_Respuesta.AddItem "Mole"
        CB_Respuesta.AddItem "Paella"
        CB_Respuesta.AddItem "Pasta"
        CB_Respuesta.AddItem "Pizza"
        CB_Respuesta.AddItem "Sushi"
        CB_Respuesta.AddItem "Tacos"
        CB_Respuesta.AddItem "Tamales"
    End If
    
    
End Sub

ACUÑA
Private Sub Btn_Volver_Click()
    ACUÑA.Hide
End Sub
ADMINISTRACIÓN
Private Sub Btn_CerrarSesión_Click()
    ADMINISTRACIÓN.Hide
    INICIO.Show
End Sub

Private Sub Btn_Ventas_Click()
    ADMINISTRACIÓN.Hide
    VENTAS.Show
End Sub

Private Sub BtnCartelera_Click()
    ADMINISTRACIÓN.Hide
    ACCARTELERA.Show
End Sub

Private Sub lbl_CambioAdministrador_Click()
    ADMINISTRACIÓN.Hide
    CAMADMINISTRADOR.Show
End Sub


Private Sub txt_CamUsuario_Click()
    ADMINISTRACIÓN.Hide
    CAMUSUARIO.Show
End Sub
ADMINISTRADOR
Private Sub Btn_Cancelar_Click()
    ADMINISTRADOR.Hide
    INICIO.Show
End Sub

Private Sub Btn_Aceptar_Click()
    If txt_Usuario = "" Or txt_Contraseña = "" Then
        MsgBox ("Contacta al gerente para que te proporcione el usuario y la contraseña.")
        ADMINISTRADOR.Hide
        INICIO.Show
    Else
        If txt_Usuario = Worksheets("Películas").Range("M2").Value And txt_Contraseña = Worksheets("Películas").Range("N2").Value Then
            txt_Usuario = ""
            txt_Contraseña = ""
            ADMINISTRADOR.Hide
            ADMINISTRACIÓN.Show
        Else
            MsgBox ("Los datos ingresados no coinciden con el usuario y contraseña actuales. Por favor, ponte en contacto con el gerente para que te los proporcione.")
            ADMINISTRADOR.Hide
            INICIO.Show
        End If
    End If
End Sub

Private Sub UserForm_Activate()
    txt_Usuario = ""
    txt_Contraseña = ""
End Sub
BIENVENIDO
Private Sub Btn_Cartelera_Click()
    BIENVENIDO.Hide
    CARTELERA.Show
End Sub

Private Sub Btn_CerrarSesión_Click()
    Worksheets("Boletos").Range("N1").Value = ""
    BIENVENIDO.Hide
    INICIO.Show
End Sub

Private Sub Btn_Contacto_Click()
    BIENVENIDO.Hide
    CONTACTO.Show
End Sub

Private Sub Btn_MisCompras_Click()
    BIENVENIDO.Hide
    MISCOMPRAS.Show
End Sub

Private Sub Btn_Sucursales_Click()
    BIENVENIDO.Hide
    SUCURSALES.Show
End Sub


Private Sub lbl_configuración_Click()
    BIENVENIDO.Hide
    CONFIGURACIÓN.Show
End Sub


Private Sub UserForm_Activate()
    Dim i As Double
    Dim UsuarioOr As String
    Dim UsuarioAc As String
    UsuarioAc = Worksheets("Boletos").Range("M2").Value
        For i = 2 To 9999
            UsuarioOr = Worksheets("Usuarios").Range("C" & i).Value
            If UsuarioOr = UsuarioAc Then
            Dim Intentos As Double
            Intentos = Worksheets("Usuarios").Range("S" & i).Value
            If Intentos <> "0" Then
                MsgBox ("Alguien ha tratado de ingresar a tu cuenta. Número de intentos:" & Intentos & "Si no eres tú, contacta o asiste a las oficinas de Administración para cambiar tu nombre de usuario.")
                Dim f As Double
                Dim Cell As String
                For f = 19 To 999
                    Cell = Worksheets("Usuarios").Cells(i & f)
                    If Cell = "" Then
                    Worksheets("Usuarios").Cells(i & f).Value = Worksheets("Usuarios").Range("S" & i).Value
                    Exit For
                    End If
                Next f
                Worksheets("Usuarios").Range("S" & i).Value = "0"
                Exit For
            End If
            End If
        Next i
End Sub
BOLETO
Private Sub Btn_Cancelar_Click()
    CB_Película = ""
    txt_Sala = ""
    CB_Fecha = ""
    CB_Hora = ""
    txt_Precio = "0"
    txt_Cantidad = "1"
    Spin_Cantidad = 1
    txt_Total = "0"
End Sub

Private Sub Btn_CerrarSesión_Click()
    Worksheets("Boletos").Range("N1").Value = ""
    BOLETO.Hide
    INICIO.Show
End Sub

Private Sub Btn_Comprar_Click()
Dim Vacío As Boolean
Vacío = False
Dim Más As Boolean
Más = False

If CB_Sucursal = "" Then
    Vacío = True
End If
If CB_Película = "" Then
    Vacío = True
End If
If CB_Fecha = "" Then
    Vacío = True
End If
If CB_Hora = "" Then
    Vacío = True
End If
Dim Cantidad As Integer
Dim ADisponibles As Integer
Cantidad = txt_Cantidad.Value


If Vacío = True Then
    MsgBox ("Por favor, completa el formulario para realizar la compra.")
Else
ADisponibles = txt_ADisponibles.Value

If Cantidad > ADisponibles Then
    Más = True
End If
    If Más = True Then
        MsgBox ("Lo sentimos, no hay suficientes asientos disponibles para esa función. Puedes probar a otra sucursal, otra hora, o elegir otra película")
    Else
        Worksheets("Boletos").Range("O2") = CB_Hora.Value
        Worksheets("Boletos").Range("P2") = CB_Película.Value
        Worksheets("Boletos").Range("Q2") = txt_Sala.Value
        Worksheets("Boletos").Range("R2") = txt_Precio.Value
        Worksheets("Boletos").Range("S2") = txt_Cantidad.Value
        Worksheets("Boletos").Range("T2") = txt_Total.Value
        Worksheets("Boletos").Range("U2") = CB_Sucursal.Value
        CB_Sucursales = ""
        txt_ADisponibles = ""
        CB_Película = ""
        txt_Sala = ""

        CB_Hora = ""
        txt_Precio = "0"
        txt_Cantidad = "1"
        Spin_Cantidad = 1
        txt_Total = "0"
        BOLETO.Hide
        COMPRA.Show
    End If
End If

End Sub

Private Sub Btn_VerCartelera_Click()
    BOLETO.Hide
    CARTELERA.Show
End Sub

Private Sub Btn_VolverInicio_Click()
    BOLETO.Hide
    BIENVENIDO.Show
End Sub

Private Sub CB_Fecha_Change()
    txt_Precio = "0"
    txt_Cantidad = "1"
    Spin_Cantidad = 1
    txt_Total = "0"
    txt_ADisponibles = ""
    CB_Hora.Clear
Dim FechaSel As Date

If CB_Fecha = "" Then
    CB_Hora.Clear
End If

Dim Película As String
Dim PelículaSelec As String
Dim i As Integer
PelículaSelec = CB_Película.Value
Dim Fecha1 As Date
Dim FechaH As Date


If CB_Fecha <> "" Then
CB_Hora.Clear
    FechaSel = CB_Fecha
Dim X As Integer
Dim Dia(1 To 7) As String
Dia(1) = "Domingo"
Dia(2) = "Lunes"
Dia(3) = "Martes"
Dia(4) = "Miercoles"
Dia(5) = "Jueves"
Dia(6) = "Viernes"
Dia(7) = "Sabado"
X = Weekday(FechaSel)


Dim Hora12 As String
Dim Hora14 As String
Dim Hora16 As String
Dim Hora18 As String
Dim Hora20 As String
Dim Hora22 As String

Hora12 = Worksheets("Películas").Range("H10").Value
Hora14 = Worksheets("Películas").Range("H11").Value
Hora16 = Worksheets("Películas").Range("H12").Value
Hora18 = Worksheets("Películas").Range("H13").Value
Hora20 = Worksheets("Películas").Range("H14").Value
Hora22 = Worksheets("Películas").Range("H15").Value

    For i = 2 To 7 Step 1
        Película = Worksheets("Películas").Range("B" & i)
        If Película = PelículaSelec Then
            Fecha1 = Worksheets("Películas").Range("C" & i).Value
            FechaH = Worksheets("Películas").Range("K2").Value
            If FechaH < FechaSel Then
                If FechaSel = Fecha1 Then
                    CB_Hora.AddItem "00:00"
                End If
                If X = 2 Or X = 3 Or X = 4 Or X = 5 Then
                    CB_Hora.AddItem "14:00"
                    CB_Hora.AddItem "16:00"
                    CB_Hora.AddItem "18:00"
                    CB_Hora.AddItem "20:00"
                End If
                If X = 6 Or X = 7 Or X = 1 Then
                    CB_Hora.AddItem "12:00"
                    CB_Hora.AddItem "14:00"
                    CB_Hora.AddItem "16:00"
                    CB_Hora.AddItem "18:00"
                    CB_Hora.AddItem "20:00"
                    CB_Hora.AddItem "22:00"
                End If
            Else
                If FechaH = FechaSel Then
                    If X = 2 Or X = 3 Or X = 4 Or X = 5 Then
                            CB_Hora.AddItem "14:00"
                            CB_Hora.AddItem "16:00"
                            CB_Hora.AddItem "18:00"
                            CB_Hora.AddItem "20:00"
                            If Time$ > Hora14 Then
                                CB_Hora.Clear
                                CB_Hora.AddItem "16:00"
                                CB_Hora.AddItem "18:00"
                                CB_Hora.AddItem "20:00"
                            End If
                            If Time$ > Hora16 Then
                                CB_Hora.Clear
                                CB_Hora.AddItem "18:00"
                                CB_Hora.AddItem "20:00"
                            End If
                            If Time$ > Hora18 Then
                                CB_Hora.Clear
                                CB_Hora.AddItem "20:00"
                            End If
                            If Time$ > Hora20 Then
                                CB_Hora.Clear
                                MsgBox ("Lo sentimos, no quedan funciones para el día de hoy.")
                            End If
                    End If
                        If X = 6 Or X = 7 Or X = 1 Then
                            CB_Hora.AddItem "12:00"
                            CB_Hora.AddItem "14:00"
                            CB_Hora.AddItem "16:00"
                            CB_Hora.AddItem "18:00"
                            CB_Hora.AddItem "20:00"
                            CB_Hora.AddItem "22:00"
                            If Time$ > Hora12 Then
                                CB_Hora.Clear
                                CB_Hora.AddItem "14:00"
                                CB_Hora.AddItem "16:00"
                                CB_Hora.AddItem "18:00"
                                CB_Hora.AddItem "20:00"
                                CB_Hora.AddItem "22:00"
                            End If
                            If Time$ > Hora14 Then
                                CB_Hora.Clear
                                CB_Hora.AddItem "16:00"
                                CB_Hora.AddItem "18:00"
                                CB_Hora.AddItem "20:00"
                                CB_Hora.AddItem "22:00"
                            End If
                            If Time$ > Hora16 Then
                                CB_Hora.Clear
                                CB_Hora.AddItem "18:00"
                                CB_Hora.AddItem "20:00"
                                CB_Hora.AddItem "22:00"
                            End If
                            If Time$ > Hora18 Then
                                CB_Hora.Clear
                                CB_Hora.AddItem "20:00"
                                CB_Hora.AddItem "22:00"
                            End If
                            If Time$ > Hora20 Then
                                CB_Hora.Clear
                                CB_Hora.AddItem "22:00"
                            End If
                            If Time$ > Hora22 Then
                                CB_Hora.Clear
                                MsgBox ("Lo sentimos, no quedan funciones para el día de hoy.")
                            End If
                        End If
                End If
            End If
        End If
    Next i
End If

Worksheets("Boletos").Range("N2").Value = CB_Fecha

End Sub

Private Sub CB_Hora_Change()
    txt_Precio = "0"
    txt_Cantidad = "1"
    Spin_Cantidad = 1
    txt_Total = "0"
    txt_ADisponibles = ""

Dim PelículaSelec As String
PelículaSelec = CB_Película.Value
FechaSel = 0

If CB_Fecha <> "" Then

FechaSel = CB_Fecha

Dim X As Integer
Dim Dia(1 To 7) As String
Dia(1) = "Domingo"
Dia(2) = "Lunes"
Dia(3) = "Martes"
Dia(4) = "Miercoles"
Dia(5) = "Jueves"
Dia(6) = "Viernes"
Dia(7) = "Sabado"
X = Weekday(FechaSel)



If CB_Hora = "00:00" Then
    txt_Precio = Worksheets("Películas").Range("I2").Value
Else
    If X = 2 Or X = 3 Or X = 4 Or X = 5 Then
       txt_Precio = Worksheets("Películas").Range("I3").Value
    End If
    If X = 6 Or X = 7 Or X = 1 Then
        txt_Precio = Worksheets("Películas").Range("I4").Value
    End If
End If
End If


txt_Precio = Format(txt_Precio, "$ #,##0")
txt_Total = Format(txt_Total, "$ #,##0")

Dim Venta As Double
Dim ADisponibles As Integer
Dim Hora As String

Dim Sucursal As String
Dim Película As String
Dim HoraSel As String
Dim Hora00 As String
Dim Hora12 As String
Dim Hora14 As String
Dim Hora16 As String
Dim Hora18 As String
Dim Hora20 As String
Dim Hora22 As String
Dim HoraCB As String

Hora00 = Worksheets("Películas").Range("I9").Value
Hora12 = Worksheets("Películas").Range("I10").Value
Hora14 = Worksheets("Películas").Range("I11").Value
Hora16 = Worksheets("Películas").Range("I12").Value
Hora18 = Worksheets("Películas").Range("I13").Value
Hora20 = Worksheets("Películas").Range("I14").Value
Hora22 = Worksheets("Películas").Range("I15").Value

HoraCB = CB_Hora
Dim PelículaSel As String
Dim SucursalSel As String

If HoraCB = "00:00" Then
    HoraSel = Hora00
End If
If HoraCB = "12:00" Then
    HoraSel = Hora12
End If
If HoraCB = "14:00" Then
    HoraSel = Hora14
End If
If HoraCB = "16:00" Then
    HoraSel = Hora16
End If
If HoraCB = "18:00" Then
    HoraSel = Hora18
End If
If HoraCB = "20:00" Then
    HoraSel = Hora20
End If
If HoraCB = "22:00" Then
    HoraSel = Hora22
End If


FechaSele = Worksheets("Boletos").Range("N2").Value
PelículaSel = CB_Película
SucursalSel = CB_Sucursal
Dim Fecha As Date


ADisponibles = 300
For Venta = 2 To 10000
    Hora = Worksheets("Boletos").Range("B" & Venta).Value
    Fecha = Worksheets("Boletos").Range("A" & Venta).Value
    Película = Worksheets("Boletos").Range("C" & Venta).Value
    Sucursal = Worksheets("Boletos").Range("H" & Venta).Value
    If Hora = HoraSel And Fecha = FechaSele And Sucursal = SucursalSel And Película = PelículaSel Then
        ADisponibles = ADisponibles - Worksheets("Boletos").Range("F" & Venta).Value
    End If
Next Venta

txt_ADisponibles = ADisponibles

End Sub

Private Sub CB_Película_Change()
    txt_Precio = "0"
    txt_Cantidad = "1"
    Spin_Cantidad = 1
    txt_Total = "0"
    txt_ADisponibles = ""

Dim Película As String
Dim PelículaSelec As String
Dim i As Integer
Dim Fecha1 As Date
Dim Fecha2 As Date
Dim FechaH As Date
Dim Estreno As Integer



PelículaSelec = CB_Película.Value

For i = 2 To 7 Step 1
        Película = Worksheets("Películas").Range("B" & i)
        If Película = PelículaSelec Then
            txt_Sala = Worksheets("Películas").Range("A" & i).Value
            CB_Fecha.Clear
            Estreno = Worksheets("Películas").Range("E" & i).Value
            Fecha1 = Worksheets("Películas").Range("C" & i).Value
            Fecha2 = Worksheets("Películas").Range("D" & i).Value
            FechaH = Worksheets("Películas").Range("K2").Value
            If Estreno < 0 Then
                Do While FechaH < Fecha2
                    CB_Fecha.AddItem FechaH
                    FechaH = FechaH + 1
                Loop
            Else
                Do While Fecha1 < Fecha2
                    CB_Fecha.AddItem Fecha1
                    Fecha1 = Fecha1 + 1
                Loop
            End If
        End If
    Next i


End Sub

Private Sub CB_Sucursal_Change()
    txt_Sala = ""
    txt_Precio = "0"
    txt_Cantidad = "1"
    Spin_Cantidad = 1
    txt_Total = "0"
    txt_ADisponibles = ""
    CB_Película = ""
Dim i As Integer
Dim Eliminación As Integer


For i = 2 To 7 Step 1
        Eliminación = Worksheets("Películas").Range("F" & i).Value
        If Eliminación > 0 Then
            CB_Película.AddItem Worksheets("Películas").Range("B" & i).Value
        End If
Next i


End Sub



Private Sub Spin_Cantidad_Change()
    txt_Cantidad = Spin_Cantidad.Value
End Sub

Private Sub txt_Cantidad_Change()

Dim Total As Single
Dim Precio As Single
Dim Cantidad As Single
Precio = txt_Precio.Value
Cantidad = txt_Cantidad.Value
Total = Precio * Cantidad
txt_Total = Total
txt_Total = Format(txt_Total, "$ #,##0")

End Sub

Private Sub txt_Precio_Change()
txt_Cantidad = "1"
Dim Total As Single
Dim Precio As Single
Dim Cantidad As Single
Precio = txt_Precio.Value
Cantidad = txt_Cantidad
Total = Precio * Cantidad
txt_Total = Total
txt_Total = Format(txt_Total, "$ #,##0")

End Sub


Private Sub UserForm_Activate()

    lbl_Fecha = Worksheets("Películas").Range("K2").Value
    
    CB_Sucursales = ""
    txt_ADisponibles = ""
    CB_Película = ""
    txt_Sala = ""
    CB_Fecha = ""
    CB_Hora = ""
    txt_Precio = "0"
    txt_Cantidad = "1"
    Spin_Cantidad = 1
    txt_Total = "0"
    

    CB_Fecha = Format(CB_Fecha, "dd/mm/yyyy")

End Sub
CAMADMINISTRADOR
Private Sub Btn_AceptarAd_Click()
Dim Contraseña1 As String
Dim Contraseña2 As String
Dim ContraseñaG As String
Dim Vacío As Boolean
Vacío = False

ContraseñaG = "Gerente1"
Contraseña1 = txt_Contraseña
Contraseña2 = txt_Contraseña2

If txt_Usuario = "" Then
    Vacío = True
End If
If txt_Contraseña = "" Then
    Vacío = True
End If
If txt_Contraseña2 = "" Then
    Vacío = True
End If
If txt_ContraseñaP = "" Then
    Vacío = True
End If

If Vacío = True Then
MsgBox ("Por favor, complete el formulario para realizar el cambio de cuenta de administración.")
End If

If Vacío = False Then
If Contraseña1 = Contraseña2 Then
    If txt_ContraseñaP = ContraseñaG Then
        Worksheets("Películas").Range("M2").Value = txt_Usuario
        Worksheets("Películas").Range("N2").Value = txt_Contraseña
        txt_Usuario = ""
        txt_Contraseña = ""
        txt_Contraseña2 = ""
        txt_ContraseñaP = ""
        MsgBox ("El Usuario y la Contraseña de administración se han cambiado exitosamente")
        CAMADMINISTRADOR.Hide
        ADMINISTRACIÓN.Show
    Else
        MsgBox ("La contraseña del gerente no es correcta.")
        CAMADMINISTRADOR.Hide
        ADMINISTRACIÓN.Show
    End If
Else
    MsgBox ("Las contraseñas no coinciden.")
End If
End If
End Sub

Private Sub Btn_AceptarCS_Click()
Dim Contraseña1 As String
Dim Contraseña2 As String
Dim ContraseñaG As String
Dim Vacío As Boolean
Vacío = False

ContraseñaG = "Gerente1"
Contraseña1 = txt_Contraseña
Contraseña2 = txt_Contraseña2

If txt_Usuario = "" Then
    Vacío = True
End If
If txt_Contraseña = "" Then
    Vacío = True
End If
If txt_Contraseña2 = "" Then
    Vacío = True
End If
If txt_ContraseñaP = "" Then
    Vacío = True
End If

If Vacío = True Then
MsgBox ("Por favor, complete el formulario para realizar el cambio de cuenta de administración.")
End If

If Vacío = False Then
If Contraseña1 = Contraseña2 Then
    If txt_ContraseñaP = ContraseñaG Then
        Worksheets("Películas").Range("M2").Value = txt_Usuario
        Worksheets("Películas").Range("N2").Value = txt_Contraseña
        txt_Usuario = ""
        txt_Contraseña = ""
        txt_Contraseña2 = ""
        txt_ContraseñaP = ""
        MsgBox ("El Usuario y la Contraseña de administración se han cambiado exitosamente")
        CAMADMINISTRADOR.Hide
        INICIO.Show
    Else
        MsgBox ("La contraseña del gerente no es correcta.")
        CAMADMINISTRADOR.Hide
        INICIO.Show
    End If
Else
    MsgBox ("Las contraseñas no coinciden.")
End If
End If
End Sub

Private Sub Btn_CancelarAd_Click()
    txt_Usuario = ""
    txt_Contraseña = ""
    txt_Contraseña2 = ""
    txt_ContraseñaP = ""
    CAMADMINISTRADOR.Hide
    ADMINISTRACIÓN.Show
End Sub

Private Sub Btn_CancelarCS_Click()
    txt_Usuario = ""
    txt_Contraseña = ""
    txt_Contraseña2 = ""
    txt_ContraseñaP = ""
    CAMADMINISTRADOR.Hide
    INICIO.Show
End Sub



Private Sub UserForm_Activate()
    txt_Usuario = ""
    txt_Contraseña = ""
    txt_Contraseña2 = ""
    txt_ContraseñaP = ""
End Sub
CAMBIARCONTRASEÑA
Private Sub Btn_CambiarContraseña_Click()
Dim i As Integer
Dim Usuario As String
Usuario = Worksheets("Boletos").Range("M2")
Dim UsuarioOr As String
Dim ContraseñaOr As String
Dim ContraseñaAc As String
ContraseñaAc = txt_ContraseñaAc

Dim Vacío As Boolean
Vacío = False

If txt_ContraseñaAc = "" Or txt_NContraseña = "" Or txt_NContraseña2 = "" Then
Vacío = True
End If

If Vacío = True Then
    MsgBox ("Por Favor, completa los datos para realizar el cambio de contraseña.")
End If

If Vacío = False Then
    For i = 2 To 9999
        UsuarioOr = Worksheets("Usuarios").Range("C" & i)
        If UsuarioOr = Usuario Then
            ContraseñaOr = Worksheets("Usuarios").Range("D" & i)
            If ContraseñaOr = ContraseñaAc Then
                If txt_NContraseña = txt_NContraseña2 Then
                    Worksheets("Usuarios").Range("D" & i) = NContraseña
                    MsgBox ("La contraseña ha sido cambiada exitosamente.")
                    CAMBIARCONTRASEÑA.Hide
                    CONFIGURACIÓN.Show
                Else
                    MsgBox ("Las contraseñas nuevas no coinciden.")
                End If
            Else
                MsgBox ("La contraseña actual ingresada es incorrecta.")
                CAMBIARCONTRASEÑA.Hide
                BIENVENIDO.Show
            End If
        End If
    Next i
End If
                
            
End Sub

Private Sub Btn_Cancelar_Click()
    CAMBIARCONTRASEÑA.Hide
    CONFIGURACIÓN.Show
End Sub
CAMCONTRASEÑA
Private Sub Btn_CambiarContraseña_Click()
    Dim Nombre As String
    Dim Correo As String
    Dim Usuario As String
    Dim Contraseña As String
    Dim Contraseña2 As String
    Dim Día As String
    Dim Mes As String
    Dim Año As String
    Dim Estado As String
    Dim Ciudad As String
    Dim Pregunta As String
    Dim Respuesta As String
    
    Dim NombreOr As String
    Dim CorreoOr As String
    Dim UsuarioOr As String
    Dim DíaOr As String
    Dim MesOr As String
    Dim AñoOr As String
    Dim EstadoOr As String
    Dim CiudadOr As String
    Dim PreguntaOr As String
    Dim RespuestaOr As String
    
    Dim Vacío As Boolean
    Vacío = False
    Dim i As Double
    Dim Encontrado As Boolean
    Encontrado = False
    
    
    Nombre = txt_Nombre
    Correo = txt_Correo
    Usuario = txt_Usuario
    Contraseña = txt_Contraseña
    Contraseña2 = txt_Contraseña2
    Día = CB_Día
    Mes = CB_Mes
    Año = CB_Año
    Estado = CB_Estado
    Ciudad = CB_Ciudad
    Pregunta = CB_Pregunta
    Respuesta = CB_Respuesta
    
    If Nombre = "" Then
        Vacío = True
    Else
        Vacío = False
    End If
    If Correo = "" Then
        Vacío = True
    Else
        Vacío = False
    End If
    If Usuario = "" Then
        Vacío = True
    Else
        Vacío = False
    End If
    If Contraseña = "" Then
        Vacío = True
    Else
        Vacío = False
    End If
    If Contraseña2 = "" Then
        Vacío = True
    Else
        Vacío = False
    End If
    If Día = "" Then
        Vacío = True
    Else
        Vacío = False
    End If
    If Mes = "" Then
        Vacío = True
    Else
        Vacío = False
    End If
    If Año = "" Then
        Vacío = True
    Else
        Vacío = False
    End If
    If Estado = "" Then
        Vacío = True
    Else
        Vacío = False
    End If
    If Ciudad = "" Then
        Vacío = True
    Else
        Vacío = False
    End If
    If Pregunta = "" Then
        Vacío = True
    Else
        Vacío = False
    End If
    If Respuesta = "" Then
        Vacío = True
    Else
        Vacío = False
    End If
    If Contraseña = "" Then
        Vacío = True
    Else
        Vacío = False
    End If
    If Contraseña2 = "" Then
        Vacío = True
    Else
        Vacío = False
    End If
    
    If Vacío = True Then
        MsgBox ("Por favor, completa el formulario. Todos los campos son obligatorios para realizar el cambio de contraseña.")
    End If
    
    If Vacío = False Then
        For i = 2 To 9999 Step 1
             NombreOr = Worksheets("Usuarios").Range("A" & i)
             CorreoOr = Worksheets("Usuarios").Range("B" & i)
             UsuarioOr = Worksheets("Usuarios").Range("C" & i)
             DíaOr = Worksheets("Usuarios").Range("E" & i)
             MesOr = Worksheets("Usuarios").Range("F" & i)
             AñoOr = Worksheets("Usuarios").Range("G" & i)
             EstadoOr = Worksheets("Usuarios").Range("H" & i)
             CiudadOr = Worksheets("Usuarios").Range("I" & i)
             PreguntaOr = Worksheets("Usuarios").Range("J" & i)
             RespuestaOr = Worksheets("Usuarios").Range("K" & i)
            If UsuarioOr = Usuario And CorreoOr = Correo Then
                Encontrado = True
                If NombreOr = Nombre And DíaOr = Día And MesOr = Mes And AñoOr = Año And EstadoOr = Estado And CiudadOr = Ciudad And PreguntaOr = Pregunta And RespuestaOr = Respuesta Then
                    If txt_Contraseña.Value = txt_Contraseña2.Value Then
                        Worksheets("Usuarios").Range("D" & i).Value = txt_Contraseña.Value
                        MsgBox ("La contraseña se ha cambiado exitosamente.")
                            txt_Nombre.Value = ""
                            txt_Correo.Value = ""
                            txt_Usuario.Value = ""
                            txt_Contraseña.Value = ""
                            txt_Contraseña2.Value = ""
                            CB_Día = ""
                            CB_Mes = ""
                            CB_Año = ""
                            CB_Estado = ""
                            CB_Ciudad = ""
                            CB_Pregunta = ""
                            CB_Respuesta = ""
                        CAMCONTRASEÑA.Hide
                        INICIO.Show
                    Else
                        MsgBox ("Las contraseñas no coinciden.")
                    End If
                Else
                    MsgBox ("Los datos ingresados no coinciden con los registrados, por favor comprueba que hayas contestado correctamente todos los campos.")
                End If
            Else
            End If
        Next i
    If Encontrado = False Then
        MsgBox ("El correo electrónico y/o el nombre de usuario no coinciden con ninguna cuenta existente.")
    End If
End If
End Sub

Private Sub Btn_Cancelar_Click()
    txt_Nombre.Value = ""
    txt_Correo.Value = ""
    txt_Usuario.Value = ""
    txt_Contraseña.Value = ""
    txt_Contraseña2.Value = ""
    CB_Día = ""
    CB_Mes = ""
    CB_Año = ""
    CB_Estado = ""
    CB_Ciudad = ""
    CB_Pregunta = ""
    CB_Respuesta = ""
    CAMCONTRASEÑA.Hide
    INICIO.Show
End Sub

Private Sub CB_Estado_Change()
   If CB_Estado.Text = "Baja California Norte" Then
        CB_Ciudad.Clear
        CB_Ciudad.AddItem "Ensenada"
        CB_Ciudad.AddItem "La Paz"
        CB_Ciudad.AddItem "Mexicali"
        CB_Ciudad.AddItem "Tijuana"
    End If
    If CB_Estado.Text = "Chihuahua" Then
        CB_Ciudad.Clear
        CB_Ciudad.AddItem "Ciudad Juárez"
        CB_Ciudad.AddItem "Chihuahua"
    End If
    If CB_Estado.Text = "Coahuila" Then
        CB_Ciudad.Clear
        CB_Ciudad.AddItem "Monclova"
        CB_Ciudad.AddItem "Saltillo"
        CB_Ciudad.AddItem "Torreón"
    End If
    If CB_Estado.Text = "Durango" Then
        CB_Ciudad.Clear
        CB_Ciudad.AddItem "Durango"
        CB_Ciudad.AddItem "Gómez Palacio"
    End If
    If CB_Estado.Text = "Nuevo León" Then
        CB_Ciudad.Clear
        CB_Ciudad.AddItem "Apodaca"
        CB_Ciudad.AddItem "Guadalupe"
        CB_Ciudad.AddItem "Monterrey"
    End If
    If CB_Estado.Text = "Sinaloa" Then
        CB_Ciudad.Clear
        CB_Ciudad.AddItem "Culiacán"
        CB_Ciudad.AddItem "Los Mochis"
        CB_Ciudad.AddItem "Mazatlán"
    End If
    If CB_Estado.Text = "Sonora" Then
        CB_Ciudad.Clear
        CB_Ciudad.AddItem "Ciudad Obregón"
        CB_Ciudad.AddItem "Hermosillo"
        CB_Ciudad.AddItem "Nogales"
    End If
    If CB_Estado.Text = "Tamaulipas" Then
        CB_Ciudad.Clear
        CB_Ciudad.AddItem "Ciudad Victoria"
        CB_Ciudad.AddItem "Matamoros"
        CB_Ciudad.AddItem "Nuevo Laredo"
        CB_Ciudad.AddItem "Reynosa"
        CB_Ciudad.AddItem "Tampico"
    End If
End Sub

Private Sub CB_Pregunta_Change()
    If CB_Pregunta.Text = "¿Cuál es tu color favorito?" Then
        CB_Respuesta.Clear
        CB_Respuesta.AddItem "Amarillo"
        CB_Respuesta.AddItem "Azul"
        CB_Respuesta.AddItem "Blanco"
        CB_Respuesta.AddItem "Café"
        CB_Respuesta.AddItem "Gris"
        CB_Respuesta.AddItem "Morado"
        CB_Respuesta.AddItem "Naranja"
        CB_Respuesta.AddItem "Negro"
        CB_Respuesta.AddItem "Rojo"
        CB_Respuesta.AddItem "Rosa"
        CB_Respuesta.AddItem "Turqueza"
        CB_Respuesta.AddItem "Verde"
    End If
    If CB_Pregunta.Text = "¿Cuál es tu animal favorito?" Then
        CB_Respuesta.Clear
        CB_Respuesta.AddItem "Águila"
        CB_Respuesta.AddItem "Delfín"
        CB_Respuesta.AddItem "Elefante"
        CB_Respuesta.AddItem "Gato"
        CB_Respuesta.AddItem "Hámster"
        CB_Respuesta.AddItem "Koala"
        CB_Respuesta.AddItem "León"
        CB_Respuesta.AddItem "Lobo"
        CB_Respuesta.AddItem "Pato"
        CB_Respuesta.AddItem "Perro"
        CB_Respuesta.AddItem "Pingüino"
        CB_Respuesta.AddItem "Venado"
    End If
    If CB_Pregunta.Text = "¿Cuál es el mes de nacimiento de tu mamá?" Then
        CB_Respuesta.Clear
        CB_Respuesta.AddItem "Enero"
        CB_Respuesta.AddItem "Febrero"
        CB_Respuesta.AddItem "Marzo"
        CB_Respuesta.AddItem "Abril"
        CB_Respuesta.AddItem "Mayo"
        CB_Respuesta.AddItem "Junio"
        CB_Respuesta.AddItem "Julio"
        CB_Respuesta.AddItem "Agosto"
        CB_Respuesta.AddItem "Septiembre"
        CB_Respuesta.AddItem "Octubre"
        CB_Respuesta.AddItem "Noviembre"
        CB_Respuesta.AddItem "Diciembre"
    End If
    If CB_Pregunta.Text = "¿Cuál es tu comida favorita?" Then
        CB_Respuesta.Clear
        CB_Respuesta.AddItem "Asado"
        CB_Respuesta.AddItem "Burritos"
        CB_Respuesta.AddItem "Hamburguesas"
        CB_Respuesta.AddItem "Hot Dog"
        CB_Respuesta.AddItem "Lasaña"
        CB_Respuesta.AddItem "Mole"
        CB_Respuesta.AddItem "Paella"
        CB_Respuesta.AddItem "Pasta"
        CB_Respuesta.AddItem "Pizza"
        CB_Respuesta.AddItem "Sushi"
        CB_Respuesta.AddItem "Tacos"
        CB_Respuesta.AddItem "Tamales"
    End If
End Sub

Private Sub UserForm_Activate()
    txt_Nombre.Value = ""
    txt_Correo.Value = ""
    txt_Usuario.Value = ""
    txt_Contraseña.Value = ""
    txt_Contraseña2.Value = ""
    CB_Día = ""
    CB_Mes = ""
    CB_Año = ""
    CB_Estado = ""
    CB_Ciudad = ""
    CB_Pregunta = ""
    CB_Respuesta = ""

    
    CB_Día.AddItem "1"
    CB_Día.AddItem "2"
    CB_Día.AddItem "3"
    CB_Día.AddItem "4"
    CB_Día.AddItem "5"
    CB_Día.AddItem "6"
    CB_Día.AddItem "7"
    CB_Día.AddItem "8"
    CB_Día.AddItem "9"
    CB_Día.AddItem "10"
    CB_Día.AddItem "11"
    CB_Día.AddItem "12"
    CB_Día.AddItem "13"
    CB_Día.AddItem "14"
    CB_Día.AddItem "15"
    CB_Día.AddItem "16"
    CB_Día.AddItem "17"
    CB_Día.AddItem "18"
    CB_Día.AddItem "19"
    CB_Día.AddItem "20"
    CB_Día.AddItem "21"
    CB_Día.AddItem "22"
    CB_Día.AddItem "23"
    CB_Día.AddItem "24"
    CB_Día.AddItem "25"
    CB_Día.AddItem "26"
    CB_Día.AddItem "27"
    CB_Día.AddItem "28"
    CB_Día.AddItem "29"
    CB_Día.AddItem "30"
    CB_Día.AddItem "31"
    
    CB_Mes.AddItem "Enero"
    CB_Mes.AddItem "Febrero"
    CB_Mes.AddItem "Marzo"
    CB_Mes.AddItem "Abril"
    CB_Mes.AddItem "Mayo"
    CB_Mes.AddItem "Junio"
    CB_Mes.AddItem "Julio"
    CB_Mes.AddItem "Agosto"
    CB_Mes.AddItem "Septiembre"
    CB_Mes.AddItem "Octubre"
    CB_Mes.AddItem "Noviembre"
    CB_Mes.AddItem "Diciembre"
    
    CB_Año.AddItem "2019"
    CB_Año.AddItem "2018"
    CB_Año.AddItem "2017"
    CB_Año.AddItem "2016"
    CB_Año.AddItem "2015"
    CB_Año.AddItem "2014"
    CB_Año.AddItem "2013"
    CB_Año.AddItem "2012"
    CB_Año.AddItem "2011"
    CB_Año.AddItem "2010"
    CB_Año.AddItem "2009"
    CB_Año.AddItem "2008"
    CB_Año.AddItem "2007"
    CB_Año.AddItem "2006"
    CB_Año.AddItem "2005"
    CB_Año.AddItem "2004"
    CB_Año.AddItem "2003"
    CB_Año.AddItem "2002"
    CB_Año.AddItem "2001"
    CB_Año.AddItem "2000"
    CB_Año.AddItem "1999"
    CB_Año.AddItem "1998"
    CB_Año.AddItem "1997"
    CB_Año.AddItem "1996"
    CB_Año.AddItem "1995"
    CB_Año.AddItem "1994"
    CB_Año.AddItem "1993"
    CB_Año.AddItem "1992"
    CB_Año.AddItem "1991"
    CB_Año.AddItem "1990"
    CB_Año.AddItem "1989"
    CB_Año.AddItem "1988"
    CB_Año.AddItem "1987"
    CB_Año.AddItem "1986"
    CB_Año.AddItem "1985"
    CB_Año.AddItem "1984"
    CB_Año.AddItem "1983"
    CB_Año.AddItem "1982"
    CB_Año.AddItem "1981"
    CB_Año.AddItem "1980"
    CB_Año.AddItem "1979"
    CB_Año.AddItem "1978"
    CB_Año.AddItem "1977"
    CB_Año.AddItem "1976"
    CB_Año.AddItem "1975"
    CB_Año.AddItem "1974"
    CB_Año.AddItem "1973"
    CB_Año.AddItem "1972"
    CB_Año.AddItem "1971"
    CB_Año.AddItem "1970"
    CB_Año.AddItem "1969"
    CB_Año.AddItem "1968"
    CB_Año.AddItem "1967"
    CB_Año.AddItem "1966"
    CB_Año.AddItem "1965"
    CB_Año.AddItem "1964"
    CB_Año.AddItem "1963"
    CB_Año.AddItem "1962"
    CB_Año.AddItem "1961"
    CB_Año.AddItem "1960"
    CB_Año.AddItem "1959"
    CB_Año.AddItem "1958"
    CB_Año.AddItem "1957"
    CB_Año.AddItem "1956"
    CB_Año.AddItem "1955"
    CB_Año.AddItem "1954"
    CB_Año.AddItem "1953"
    CB_Año.AddItem "1952"
    CB_Año.AddItem "1951"
    CB_Año.AddItem "1950"
    CB_Año.AddItem "1949"
    CB_Año.AddItem "1948"
    CB_Año.AddItem "1947"
    CB_Año.AddItem "1946"
    CB_Año.AddItem "1945"
    CB_Año.AddItem "1944"
    CB_Año.AddItem "1943"
    CB_Año.AddItem "1942"
    CB_Año.AddItem "1941"
    CB_Año.AddItem "1940"
    CB_Año.AddItem "1939"
    CB_Año.AddItem "1938"
    CB_Año.AddItem "1937"
    CB_Año.AddItem "1936"
    CB_Año.AddItem "1935"
    CB_Año.AddItem "1934"
    CB_Año.AddItem "1933"
    CB_Año.AddItem "1932"
    CB_Año.AddItem "1931"
    CB_Año.AddItem "1930"
    CB_Año.AddItem "1929"
    CB_Año.AddItem "1928"
    CB_Año.AddItem "1927"
    CB_Año.AddItem "1926"
    CB_Año.AddItem "1925"
    CB_Año.AddItem "1924"
    CB_Año.AddItem "1923"
    CB_Año.AddItem "1922"
    CB_Año.AddItem "1921"
    CB_Año.AddItem "1920"

    CB_Estado.AddItem "Baja California Norte"
    CB_Estado.AddItem "Chihuahua"
    CB_Estado.AddItem "Coahuila"
    CB_Estado.AddItem "Durango"
    CB_Estado.AddItem "Nuevo León"
    CB_Estado.AddItem "Sinaloa"
    CB_Estado.AddItem "Sonora"
    CB_Estado.AddItem "Tamaulipas"

    CB_Pregunta.AddItem "¿Cuál es tu color favorito?"
    CB_Pregunta.AddItem "¿Cuál es tu animal favorito?"
    CB_Pregunta.AddItem "¿Cuál es el mes de nacimiento de tu mamá?"
    CB_Pregunta.AddItem "¿Cuál es tu comida favorita?"
End Sub
CAMTARJETA
Private Sub Btn_Aceptar_Click()
Dim i As Double
Dim Usuario As String
Dim UsuarioAc As String
UsuarioAc = Worksheets("Boletos").Range("M2").Value
Dim Vacío As Boolean
Vacío = False

If txt_Tarjeta = "" Or CB_Mes = "" Or CB_Año = "" Or txt_Clave = "" Then
    Vacío = True
End If

If Vacío = True Then
    MsgBox ("Por favor, llene el formulario para actualizar los datos de su tarjeta.")
End If

If Vacío = False Then
    For i = 2 To 9999
        Usuario = Worksheets("Usuarios").Range("C" & i).Value
        If Usuario = UsuarioAc Then
            Worksheets("Usuarios").Range("N" & i) = txt_Tarjeta
            Worksheets("Usuarios").Range("O" & i) = CB_Mes
            Worksheets("Usuarios").Range("P" & i) = CB_Año
            Worksheets("Usuarios").Range("Q" & i) = txt_Clave
            MsgBox ("Los datos de su tarjeta se han actualizado exitosamente.")
        End If
    Next i
End If
End Sub

Private Sub Btn_Cancelar_Click()
    CAMTARJETA.Hide
    CONFIGURACIÓN.Show
End Sub

Private Sub CB_Mes_Change()
Dim Fecha As String
Fecha = Worksheets("Películas").Range("K2").Value
Dim Mes As Double
Mes = Month(Fecha)
Dim MesSelec As Integer

If CB_Mes = "Enero" Then
    MesSelec = 1
End If
If CB_Mes = "Febrero" Then
    MesSelec = 2
End If
If CB_Mes = "Marzo" Then
    MesSelec = 3
End If
If CB_Mes = "Abril" Then
    MesSelec = 4
End If
If CB_Mes = "Mayo" Then
    MesSelec = 5
End If
If CB_Mes = "Junio" Then
    MesSelec = 6
End If
If CB_Mes = "Julio" Then
    MesSelec = 7
End If
If CB_Mes = "Agosto" Then
    MesSelec = 8
End If
If CB_Mes = "Septiembre" Then
    MesSelec = 9
End If
If CB_Mes = "Octubre" Then
    MesSelec = 10
End If
If CB_Mes = "Noviembre" Then
    MesSelec = 11
End If
If CB_Mes = "Diciembre" Then
    MesSelec = 12
End If

Año = Year(Fecha)
Límite = Año + 10

    If MesSelec < Mes Then
    CB_Año.Clear
    Año = Año + 1
    Do While Año < Límite
        CB_Año.AddItem Año
        Año = Año + 1
    Loop
    End If
    
    If MesSelec >= Mes Then
    CB_Año.Clear
    Do While Año < Límite
        CB_Año.AddItem Año
        Año = Año + 1
    Loop
    End If
End Sub

Private Sub UserForm_Activate()
Dim Fecha As String
Fecha = Worksheets("Películas").Range("K2").Value
Dim Año As Double
Año = Year(Fecha)
Límite = Año + 10

    Do While Año < Límite
        CB_Año.AddItem Año
        Año = Año + 1
    Loop

CB_Mes.AddItem "Enero"
CB_Mes.AddItem "Febrero"
CB_Mes.AddItem "Marzo"
CB_Mes.AddItem "Abril"
CB_Mes.AddItem "Mayo"
CB_Mes.AddItem "Junio"
CB_Mes.AddItem "Julio"
CB_Mes.AddItem "Agosto"
CB_Mes.AddItem "Septiembre"
CB_Mes.AddItem "Octubre"
CB_Mes.AddItem "Noviembre"
CB_Mes.AddItem "Diciembre"


Dim i As Double
Dim Usuario As String
Dim UsuarioAc As String
UsuarioAc = Worksheets("Boletos").Range("M2").Value

    For i = 2 To 9999
        Usuario = Worksheets("Usuarios").Range("C" & i).Value
        If Usuario = UsuarioAc Then
            txt_Tarjeta = Worksheets("Usuarios").Range("N" & i)
            CB_Mes = Worksheets("Usuarios").Range("O" & i)
            CB_Año = Worksheets("Usuarios").Range("P" & i)
            txt_Clave = Worksheets("Usuarios").Range("Q" & i)
        End If
    Next i

End Sub
CAMUSUARIO
Private Sub Btn_CamUsuario_Click()
Dim Libre As Boolean
Libre = True
UsuarioAc = txt_UsuarioAc

Dim Vacío As Boolean
Vacío = False

If txt_NUsuario = "" Or txt_Contraseña = "" Then
Vacío = True
End If

If Vacío = False Then
For i = 2 To 9999
    UsuarioOr = Worksheets("Usuarios").Range("C" & i).Value
    If UsuarioOr = UsuarioAc Then
        Dim ContraseñaOr As String
        Dim Contraseña As String
        Contraseña = txt_Contraseña
        ContraseñaOr = orksheets("Usuarios").Range("D" & i).Value
        If ContraeñaOr = Contraseña Then
            Dim f As String
            Dim User As String
            Dim NUsuario As String
                For f = 2 To 9999
                    User = Worksheets("Usuarios").Range("C" & f).Value
                    If User = Usuario Then
                        MsgBox ("Ya existe una cuenta con es nombre de usuario. Por favor, prueba con un nombre de usuario diferente")
                        Libre = False
                        Exit For
                    End If
                Next i
        End If
    End If
Next i
End If

If Libre = True Then
    Worksheets("Usuarios").Range("C" & i).Value = txt_NUsuario
    MsgBox ("El nombre de Usuario se  ha cambiado exitosamente.")
End If
End Sub

Private Sub Btn_Cancelar_Click()
    CAMUSUARIO.Hide
    ADMINISTRACIÓN.Show
End Sub

Private Sub Brn_Verificar_Click()
Dim i As Double
Dim UsuarioAc As String
Dim UsuarioOr As String
UsuarioAc = txt_UsuarioAc
For i = 2 To 9999
    UsuarioOr = Worksheets("Usuarios").Range("C" & i).Value
    If UsuarioOr = UsuarioAc Then
        Dim f As Double
        Dim Cell As String
        Dim Intentos As Double
        Intentos = 0
            For f = 19 To 9999
                Cell = Worksheets("Usuarios").Cells(i & f).Value
                If Cell <> "" Then
                Intentos = Intentos + Cell.Value
                End If
                If Cell = "" Then
                Exit For
            Next f
    Exit For
    End If
Next i

txt_Intentos = Intentos

End Sub
CARTELERA
Private Sub Btn_Boletos_Click()
    CARTELERA.Hide
    BOLETO.Show
End Sub

Private Sub Btn_CerrarSesión_Click()
    Worksheets("Boletos").Range("N1").Value = ""
    CARTELERA.Hide
    INICIO.Show
End Sub

Private Sub Btn_VolverInicio_Click()
    CARTELERA.Hide
    BIENVENIDO.Show
End Sub

Private Sub UserForm_Activate()
    txt_Película1 = Worksheets("Películas").Range("B2").Value
    txt_Película2 = Worksheets("Películas").Range("B3").Value
    txt_Película3 = Worksheets("Películas").Range("B4").Value
    txt_Película4 = Worksheets("Películas").Range("B5").Value
    txt_Película5 = Worksheets("Películas").Range("B6").Value
    txt_Película6 = Worksheets("Películas").Range("B7").Value
    txt_Estreno1 = Worksheets("Películas").Range("C2").Value
    txt_Estreno2 = Worksheets("Películas").Range("C3").Value
    txt_Estreno3 = Worksheets("Películas").Range("C4").Value
    txt_Estreno4 = Worksheets("Películas").Range("C5").Value
    txt_Estreno5 = Worksheets("Películas").Range("C6").Value
    txt_Estreno6 = Worksheets("Películas").Range("C7").Value
    txt_Eliminación1 = Worksheets("Películas").Range("D2").Value
    txt_Eliminación2 = Worksheets("Películas").Range("D3").Value
    txt_Eliminación3 = Worksheets("Películas").Range("D4").Value
    txt_Eliminación4 = Worksheets("Películas").Range("D5").Value
    txt_Eliminación5 = Worksheets("Películas").Range("D6").Value
    txt_Eliminación6 = Worksheets("Películas").Range("D7").Value

Dim Estreno1 As Integer
Dim Estreno2 As Integer
Dim Estreno3 As Integer
Dim Estreno4 As Integer
Dim Estreno5 As Integer
Dim Estreno6 As Integer
Dim Eliminación1 As Integer
Dim Eliminación2 As Integer
Dim Eliminación3 As Integer
Dim Eliminación4 As Integer
Dim Eliminación5 As Integer
Dim Eliminación6 As Integer

Estreno1 = Worksheets("Películas").Range("E2").Value
Estreno2 = Worksheets("Películas").Range("E3").Value
Estreno3 = Worksheets("Películas").Range("E4").Value
Estreno4 = Worksheets("Películas").Range("E5").Value
Estreno5 = Worksheets("Películas").Range("E6").Value
Estreno6 = Worksheets("Películas").Range("E7").Value
Eliminación1 = Worksheets("Películas").Range("F2").Value
Eliminación2 = Worksheets("Películas").Range("F3").Value
Eliminación3 = Worksheets("Películas").Range("F4").Value
Eliminación4 = Worksheets("Películas").Range("F5").Value
Eliminación5 = Worksheets("Películas").Range("F6").Value
Eliminación6 = Worksheets("Películas").Range("F7").Value

If Estreno1 > 0 Then
    txt_Disponibilidad1 = "PRÓXIMAMENTE"
Else
    If Eliminación1 <= 0 Then
        txt_Disponibilidad1 = "NO DISPONIBLE"
    Else
        txt_Disponibilidad1 = "PRESENTANDO"
    End If
End If

If Estreno2 > 0 Then
    txt_Disponibilidad2 = "PRÓXIMAMENTE"
Else
    If Eliminación2 <= 0 Then
        txt_Disponibilidad2 = "NO DISPONIBLE"
    Else
        txt_Disponibilidad2 = "PRESENTANDO"
    End If
End If

If Estreno3 > 0 Then
    txt_Disponibilidad3 = "PRÓXIMAMENTE"
Else
    If Eliminación3 <= 0 Then
        txt_Disponibilidad3 = "NO DISPONIBLE"
    Else
        txt_Disponibilidad3 = "PRESENTANDO"
    End If
End If

If Estreno4 > 0 Then
    txt_Disponibilidad4 = "PRÓXIMAMENTE"
Else
    If Eliminación4 <= 0 Then
        txt_Disponibilidad4 = "NO DISPONIBLE"
    Else
        txt_Disponibilidad4 = "PRESENTANDO"
    End If
End If

If Estreno5 > 0 Then
    txt_Disponibilidad5 = "PRÓXIMAMENTE"
Else
    If Eliminación5 <= 0 Then
        txt_Disponibilidad5 = "NO DISPONIBLE"
    Else
        txt_Disponibilidad5 = "PRESENTANDO"
    End If
End If

If Estreno6 > 0 Then
    txt_Disponibilidad6 = "PRÓXIMAMENTE"
Else
    If Eliminación6 <= 0 Then
        txt_Disponibilidad6 = "NO DISPONIBLE"
    Else
        txt_Disponibilidad6 = "PRESENTANDO"
    End If
End If
    
End Sub
COMPRA
Private Sub Btn_Aceptar_Click()
Dim Vacío As Boolean
Vacío = False

If CB_Método = "" Then
Vacío = True
End If

If Vacío = True Then
    MsgBox ("Por favor, elija un método de pago para continuar con su compra.")
End If

Dim r As Double
Dim Cell As String

If Vacío = False Then
If CB_Método = "Caja" Then
For r = 2 To 9999
    Cell = Worksheets("Boletos").Range("A" & r)
    If Cell = "" Then
        Worksheets("Boletos").Range("A" & r) = Worksheets("Boletos").Range("N2")
        Worksheets("Boletos").Range("B" & r) = Worksheets("Boletos").Range("O2")
        Worksheets("Boletos").Range("C" & r) = Worksheets("Boletos").Range("P2")
        Worksheets("Boletos").Range("D" & r) = Worksheets("Boletos").Range("Q2")
        Worksheets("Boletos").Range("E" & r) = Worksheets("Boletos").Range("R2")
        Worksheets("Boletos").Range("F" & r) = Worksheets("Boletos").Range("S2")
        Worksheets("Boletos").Range("G" & r) = Worksheets("Boletos").Range("T2")
        Worksheets("Boletos").Range("H" & r) = Worksheets("Boletos").Range("U2")
        Worksheets("Boletos").Range("I" & r) = Worksheets("Boletos").Range("M2")
        Worksheets("Boletos").Range("J" & r) = CB_Método.Text
        Dim Código As Double
        Dim Arriba As Double
        Arriba = r - 1
        Código = Worksheets("Boletos").Range("K" & Arriba).Value
        Código = Código + 1
        Worksheets("Boletos").Range("K" & r) = Código
        COMPRA.Hide
        FIN.Show
        Exit For
    End If
Next r
Else
    Dim i As Double
    Dim UsuarioOr As String
    Dim UsuarioAc As String
    UsuarioAc = Worksheets("Boletos").Range("M2")
    For i = 2 To 9999
        UsuarioOr = Worksheets("Usuarios").Range("C" & i)
        If UsuarioOr = UsuarioAc Then
            Dim TARJETA As String
            TARJETA = Worksheets("Usuarios").Range("N" & i)
                If TARJETA = "" Then
                    MsgBox ("No has igresado los datos de tu tarjeta de crédito o débito.")
                    COMPRA.Hide
                    REGTARJETA.Show
                Else
                    Dim AñoAc As Double
                    Dim AñoV As Double
                    Dim Fecha As Date
                    Fecha = Worksheets("Películas").Range("K2").Value
                    AñoV = Worksheets("Usuarios").Range("P" & i)
                    AñoAc = Year(Fecha)
                    If AñoV < AñoAc Then
                        MsgBox ("Su tarjeta de crédito o débito ya se ha vencido, por favor seleccione otro método de pago o acceda a configuración para actualizar los datos de tarjeta")
                    End If
                    If AñoV = AñoAc Then
                    Dim MesV As String
                    MesV = Worksheets("Usuarios").Range("O" & i)
                    Dim MesSelec As Integer
                    Dim MesAc As Integer
                    MesAc = Month(Fecha)
                        If MesV = "Enero" Then
                            MesSelec = 1
                        End If
                        If MesV = "Febrero" Then
                            MesSelec = 2
                        End If
                        If MesV = "Marzo" Then
                            MesSelec = 3
                        End If
                        If MesV = "Abril" Then
                            MesSelec = 4
                        End If
                        If MesV = "Mayo" Then
                            MesSelec = 5
                        End If
                        If MesV = "Junio" Then
                            MesSelec = 6
                        End If
                        If MesV = "Julio" Then
                            MesSelec = 7
                        End If
                        If MesV = "Agosto" Then
                            MesSelec = 8
                        End If
                        If MesV = "Septiembre" Then
                            MesSelec = 9
                        End If
                        If MesV = "Octubre" Then
                            MesSelec = 10
                        End If
                        If MesV = "Noviembre" Then
                            MesSelec = 11
                        End If
                        If MesV = "Diciembre" Then
                            MesSelec = 12
                        End If
                    If MesSelec < MesAc Then
                        MsgBox ("Su tarjeta de crédito o débito ya se ha vencido, por favor seleccione otro método de pago o acceda a configuración para actualizar los datos de tarjeta")
                    Else
                        For r = 2 To 9999
                            Cell = Worksheets("Boletos").Range("A" & r)
                            If Cell = "" Then
                                Worksheets("Boletos").Range("A" & r) = Worksheets("Boletos").Range("N2")
                                Worksheets("Boletos").Range("B" & r) = Worksheets("Boletos").Range("O2")
                                Worksheets("Boletos").Range("C" & r) = Worksheets("Boletos").Range("P2")
                                Worksheets("Boletos").Range("D" & r) = Worksheets("Boletos").Range("Q2")
                                Worksheets("Boletos").Range("E" & r) = Worksheets("Boletos").Range("R2")
                                Worksheets("Boletos").Range("F" & r) = Worksheets("Boletos").Range("S2")
                                Worksheets("Boletos").Range("G" & r) = Worksheets("Boletos").Range("T2")
                                Worksheets("Boletos").Range("H" & r) = Worksheets("Boletos").Range("U2")
                                Worksheets("Boletos").Range("I" & r) = Worksheets("Boletos").Range("M2")
                                Worksheets("Boletos").Range("J" & r) = CB_Método.Text
                                Arriba = r - 1
                                Código = Worksheets("Boletos").Range("K" & Arriba).Value
                                Código = Código + 1
                                Worksheets("Boletos").Range("K" & r) = Código
                                COMPRA.Hide
                                FIN.Show
                                Exit For
                            End If
                        Next r
                    End If
                    End If
                    If AñoV > AñoAc Then
                        For r = 2 To 9999
                            Cell = Worksheets("Boletos").Range("A" & r)
                            If Cell = "" Then
                                Worksheets("Boletos").Range("A" & r) = Worksheets("Boletos").Range("N2")
                                Worksheets("Boletos").Range("B" & r) = Worksheets("Boletos").Range("O2")
                                Worksheets("Boletos").Range("C" & r) = Worksheets("Boletos").Range("P2")
                                Worksheets("Boletos").Range("D" & r) = Worksheets("Boletos").Range("Q2")
                                Worksheets("Boletos").Range("E" & r) = Worksheets("Boletos").Range("R2")
                                Worksheets("Boletos").Range("F" & r) = Worksheets("Boletos").Range("S2")
                                Worksheets("Boletos").Range("G" & r) = Worksheets("Boletos").Range("T2")
                                Worksheets("Boletos").Range("H" & r) = Worksheets("Boletos").Range("U2")
                                Worksheets("Boletos").Range("I" & r) = Worksheets("Boletos").Range("M2")
                                Worksheets("Boletos").Range("J" & r) = CB_Método.Text
                                Arriba = r - 1
                                Código = Worksheets("Boletos").Range("K" & Arriba).Value
                                Código = Código + 1
                                Worksheets("Boletos").Range("K" & r) = Código
                                COMPRA.Hide
                                FIN.Show
                                Exit For
                            End If
                        Next r
                    End If
                End If
               
        End If
    Next i
    End If
    End If
End Sub

Private Sub Btn_Cancelar_Click()
    COMPRA.Hide
    BOLETO.Show
End Sub


Private Sub UserForm_Activate()

    txt_Sucursal = Worksheets("Boletos").Range("U2")
    txt_Película = Worksheets("Boletos").Range("P2")
    txt_Fecha = Worksheets("Boletos").Range("N2")
    txt_Hora = Worksheets("Boletos").Range("O2").Text
    txt_Sala = Worksheets("Boletos").Range("Q2")
    txt_Cantidad = Worksheets("Boletos").Range("T2")
    txt_Boletos = Worksheets("Boletos").Range("S2")
        
    txt_Cantidad = Format(txt_Cantidad, "$ #,##0")
    
Dim i As Double
Dim Usuario As String
Usuario = Worksheets("Boletos").Range("M2").Value
Dim UsuarioOr As String
Dim FechaAc As Date
Dim FechaNac As Date
Dim Años As Single



    For i = 2 To 9999
        UsuarioOr = Worksheets("Usuarios").Range("C" & i).Value
        If UsuarioOr = Usuario Then
            FechaAc = Worksheets("Películas").Range("K2").Value
            FechaNac = Worksheets("Usuarios").Range("L" & i).Value
            Años = DateDiff("yyyy", FechaNac, FechaAc)
            If Años < 18 Then
                MsgBox ("La cuenta pertenece a un menor de edad, por lo que el pago de los boletos solo se puede realizar en caja.")
                CB_Método = "Caja"
                CB_Método.AddItem "Caja"
            Else
                CB_Método.AddItem "Caja"
                CB_Método.AddItem "Tarjeta de Crédito o Débito"
            End If
            Exit For
        End If
    Next i
End Sub
CONFIGURACIÓN
Private Sub Btn_Actualizar_Click()
    PREACTUALIZARDATOS.Show
End Sub

Private Sub Btn_CambioContraseña_Click()
    CONFIGURACIÓN.Hide
    CAMBIARCONTRASEÑA.Show
End Sub

Private Sub Btn_CamTarjeta_Click()
    PWTARJETA.Show
End Sub

Private Sub Btn_Términos_Click()
    TERMINOS.Show
End Sub

Private Sub Btn_VolverInicio_Click()
    CONFIGURACIÓN.Hide
    BIENVENIDO.Show
End Sub
CONTACTO
Private Sub Btn_VolverInicio_Click()
    CONTACTO.Hide
    BIENVENIDO.Show
End Sub
CUENTA
Option Explicit

Private Sub Btn_Cancelar_Click()
    txt_Nombre.Value = ""
    txt_Correo.Value = ""
    txt_Usuario.Value = ""
    txt_Contraseña.Value = ""
    txt_Contraseña2.Value = ""
    CB_Día = ""
    CB_Mes = ""
    CB_Año = ""
    CB_Estado = ""
    CB_Ciudad = ""
    CB_Pregunta = ""
    CB_Respuesta = ""
    CheckBox1 = False
    CUENTA.Hide
    INICIO.Show
End Sub

Private Sub Btn_CrearCuenta_Click()
    Dim i As Double
    Dim Duplicado As Boolean
    Duplicado = False
    Dim Vacío As Boolean
    Vacío = False
    
    If txt_Nombre.Value = "" Then
        Vacío = True
    End If
    If txt_Correo.Value = "" Then
        Vacío = True
    End If
    If txt_Usuario.Value = "" Then
        Vacío = True
    End If
    If txt_Contraseña.Value = "" Then
        Vacío = True
    End If
    If txt_Contraseña2.Value = "" Then
        Vacío = True
    End If
    If CB_Día = "" Then
        Vacío = True
    End If
    If CB_Mes = "" Then
        Vacío = True
    End If
    If CB_Año = "" Then
        Vacío = True
    End If
    If CB_Estado = "" Then
        Vacío = True
    End If
    If CB_Ciudad = "" Then
        Vacío = True
    End If
    If CB_Pregunta = "" Then
        Vacío = True
    End If
    If CB_Respuesta = "" Then
        Vacío = True
    End If
    
    If Vacío = True Then
        MsgBox ("Por favor, completa el formulario. Recuerda que todos los campos son obligatorios.")
    End If
    
    If Not Vacío Then
        For i = 2 To 9999
            If Worksheets("Usuarios").Range("B" & i).Value = txt_Correo Then
                MsgBox "Ya existe una cuenta con ese correo."
                Duplicado = True
            End If
            If Worksheets("Usuarios").Range("C" & i).Value = txt_Usuario Then
                MsgBox "Ya existe una cuenta con ese nombre de usuario, por favor, prueba con uno diferente."
                Duplicado = True
            End If
        Next i
        If Not Duplicado Then
            If txt_Contraseña = txt_Contraseña2 Then
                If CheckBox1 = False Then
                    MsgBox ("Por favor, lee y acepta nuestros términos y condiciones")
                Else
                    Dim Reng As Double
                    Dim Nombre As String
                    For Reng = 2 To 9999
                        Nombre = Worksheets("Usuarios").Range("A" & Reng).Value
                        If Nombre = "" Then
                        Worksheets("Usuarios").Range("A" & Reng).Value = txt_Nombre
                        Worksheets("Usuarios").Range("B" & Reng).Value = txt_Correo
                        Worksheets("Usuarios").Range("C" & Reng).Value = txt_Usuario
                        Worksheets("Usuarios").Range("D" & Reng).Value = txt_Contraseña
                        Worksheets("Usuarios").Range("E" & Reng).Value = CB_Día
                        Worksheets("Usuarios").Range("F" & Reng).Value = CB_Mes
                        Worksheets("Usuarios").Range("G" & Reng).Value = CB_Año
                        Worksheets("Usuarios").Range("H" & Reng).Value = CB_Estado
                        Worksheets("Usuarios").Range("I" & Reng).Value = CB_Ciudad
                        Worksheets("Usuarios").Range("J" & Reng).Value = CB_Pregunta
                        Worksheets("Usuarios").Range("K" & Reng).Value = CB_Respuesta
                        MsgBox ("La cuenta se ha creado exitosamente")
                        txt_Nombre.Value = ""
                        txt_Correo.Value = ""
                        txt_Usuario.Value = ""
                        txt_Contraseña.Value = ""
                        txt_Contraseña2.Value = ""
                        CB_Día = ""
                        CB_Mes = ""
                        CB_Año = ""
                        CB_Estado = ""
                        CB_Ciudad = ""
                        CB_Pregunta = ""
                        CB_Respuesta = ""
                        CheckBox1 = False
                        CUENTA.Hide
                        INICIO.Show
                        Exit For
                        
                        End If
                    Next Reng
                End If
            Else
                MsgBox ("Las contraseñas no coinciden.")
            End If
        End If
   End If
End Sub



Private Sub CB_Estado_Change()
    If CB_Estado.Text = "Baja California Norte" Then
        CB_Ciudad.Clear
        CB_Ciudad.AddItem "Ensenada"
        CB_Ciudad.AddItem "La Paz"
        CB_Ciudad.AddItem "Mexicali"
        CB_Ciudad.AddItem "Tijuana"
    End If
    If CB_Estado.Text = "Chihuahua" Then
        CB_Ciudad.Clear
        CB_Ciudad.AddItem "Ciudad Juárez"
        CB_Ciudad.AddItem "Chihuahua"
    End If
    If CB_Estado.Text = "Coahuila" Then
        CB_Ciudad.Clear
        CB_Ciudad.AddItem "Monclova"
        CB_Ciudad.AddItem "Saltillo"
        CB_Ciudad.AddItem "Torreón"
    End If
    If CB_Estado.Text = "Durango" Then
        CB_Ciudad.Clear
        CB_Ciudad.AddItem "Durango"
        CB_Ciudad.AddItem "Gómez Palacio"
    End If
    If CB_Estado.Text = "Nuevo León" Then
        CB_Ciudad.Clear
        CB_Ciudad.AddItem "Apodaca"
        CB_Ciudad.AddItem "Guadalupe"
        CB_Ciudad.AddItem "Monterrey"
    End If
    If CB_Estado.Text = "Sinaloa" Then
        CB_Ciudad.Clear
        CB_Ciudad.AddItem "Culiacán"
        CB_Ciudad.AddItem "Los Mochis"
        CB_Ciudad.AddItem "Mazatlán"
    End If
    If CB_Estado.Text = "Sonora" Then
        CB_Ciudad.Clear
        CB_Ciudad.AddItem "Ciudad Obregón"
        CB_Ciudad.AddItem "Hermosillo"
        CB_Ciudad.AddItem "Nogales"
    End If
    If CB_Estado.Text = "Tamaulipas" Then
        CB_Ciudad.Clear
        CB_Ciudad.AddItem "Ciudad Victoria"
        CB_Ciudad.AddItem "Matamoros"
        CB_Ciudad.AddItem "Nuevo Laredo"
        CB_Ciudad.AddItem "Reynosa"
        CB_Ciudad.AddItem "Tampico"
    End If

End Sub

Private Sub CB_Pregunta_Change()
    If CB_Pregunta.Text = "¿Cuál es tu color favorito?" Then
        CB_Respuesta.Clear
        CB_Respuesta.AddItem "Amarillo"
        CB_Respuesta.AddItem "Azul"
        CB_Respuesta.AddItem "Blanco"
        CB_Respuesta.AddItem "Café"
        CB_Respuesta.AddItem "Gris"
        CB_Respuesta.AddItem "Morado"
        CB_Respuesta.AddItem "Naranja"
        CB_Respuesta.AddItem "Negro"
        CB_Respuesta.AddItem "Rojo"
        CB_Respuesta.AddItem "Rosa"
        CB_Respuesta.AddItem "Turqueza"
        CB_Respuesta.AddItem "Verde"
    End If
    If CB_Pregunta.Text = "¿Cuál es tu animal favorito?" Then
        CB_Respuesta.Clear
        CB_Respuesta.AddItem "Águila"
        CB_Respuesta.AddItem "Delfín"
        CB_Respuesta.AddItem "Elefante"
        CB_Respuesta.AddItem "Gato"
        CB_Respuesta.AddItem "Hámster"
        CB_Respuesta.AddItem "Koala"
        CB_Respuesta.AddItem "León"
        CB_Respuesta.AddItem "Lobo"
        CB_Respuesta.AddItem "Pato"
        CB_Respuesta.AddItem "Perro"
        CB_Respuesta.AddItem "Pingüino"
        CB_Respuesta.AddItem "Venado"
    End If
    If CB_Pregunta.Text = "¿Cuál es el mes de nacimiento de tu mamá?" Then
        CB_Respuesta.Clear
        CB_Respuesta.AddItem "Enero"
        CB_Respuesta.AddItem "Febrero"
        CB_Respuesta.AddItem "Marzo"
        CB_Respuesta.AddItem "Abril"
        CB_Respuesta.AddItem "Mayo"
        CB_Respuesta.AddItem "Junio"
        CB_Respuesta.AddItem "Julio"
        CB_Respuesta.AddItem "Agosto"
        CB_Respuesta.AddItem "Septiembre"
        CB_Respuesta.AddItem "Octubre"
        CB_Respuesta.AddItem "Noviembre"
        CB_Respuesta.AddItem "Diciembre"
    End If
    If CB_Pregunta.Text = "¿Cuál es tu comida favorita?" Then
        CB_Respuesta.Clear
        CB_Respuesta.AddItem "Asado"
        CB_Respuesta.AddItem "Burritos"
        CB_Respuesta.AddItem "Hamburguesas"
        CB_Respuesta.AddItem "Hot Dog"
        CB_Respuesta.AddItem "Lasaña"
        CB_Respuesta.AddItem "Mole"
        CB_Respuesta.AddItem "Paella"
        CB_Respuesta.AddItem "Pasta"
        CB_Respuesta.AddItem "Pizza"
        CB_Respuesta.AddItem "Sushi"
        CB_Respuesta.AddItem "Tacos"
        CB_Respuesta.AddItem "Tamales"
    End If
End Sub

Private Sub lbl_Términos_Click()
    TERMINOS.Show
End Sub

Private Sub UserForm_Activate()
    txt_Nombre.Value = ""
    txt_Correo.Value = ""
    txt_Usuario.Value = ""
    txt_Contraseña.Value = ""
    txt_Contraseña2.Value = ""
    CB_Día = ""
    CB_Mes = ""
    CB_Año = ""
    CB_Estado = ""
    CB_Ciudad = ""
    CB_Pregunta = ""
    CB_Respuesta = ""
    CheckBox1 = False
    
    CB_Día.AddItem "1"
    CB_Día.AddItem "2"
    CB_Día.AddItem "3"
    CB_Día.AddItem "4"
    CB_Día.AddItem "5"
    CB_Día.AddItem "6"
    CB_Día.AddItem "7"
    CB_Día.AddItem "8"
    CB_Día.AddItem "9"
    CB_Día.AddItem "10"
    CB_Día.AddItem "11"
    CB_Día.AddItem "12"
    CB_Día.AddItem "13"
    CB_Día.AddItem "14"
    CB_Día.AddItem "15"
    CB_Día.AddItem "16"
    CB_Día.AddItem "17"
    CB_Día.AddItem "18"
    CB_Día.AddItem "19"
    CB_Día.AddItem "20"
    CB_Día.AddItem "21"
    CB_Día.AddItem "22"
    CB_Día.AddItem "23"
    CB_Día.AddItem "24"
    CB_Día.AddItem "25"
    CB_Día.AddItem "26"
    CB_Día.AddItem "27"
    CB_Día.AddItem "28"
    CB_Día.AddItem "29"
    CB_Día.AddItem "30"
    CB_Día.AddItem "31"
    
    CB_Mes.AddItem "Enero"
    CB_Mes.AddItem "Febrero"
    CB_Mes.AddItem "Marzo"
    CB_Mes.AddItem "Abril"
    CB_Mes.AddItem "Mayo"
    CB_Mes.AddItem "Junio"
    CB_Mes.AddItem "Julio"
    CB_Mes.AddItem "Agosto"
    CB_Mes.AddItem "Septiembre"
    CB_Mes.AddItem "Octubre"
    CB_Mes.AddItem "Noviembre"
    CB_Mes.AddItem "Diciembre"
    
    CB_Año.AddItem "2019"
    CB_Año.AddItem "2018"
    CB_Año.AddItem "2017"
    CB_Año.AddItem "2016"
    CB_Año.AddItem "2015"
    CB_Año.AddItem "2014"
    CB_Año.AddItem "2013"
    CB_Año.AddItem "2012"
    CB_Año.AddItem "2011"
    CB_Año.AddItem "2010"
    CB_Año.AddItem "2009"
    CB_Año.AddItem "2008"
    CB_Año.AddItem "2007"
    CB_Año.AddItem "2006"
    CB_Año.AddItem "2005"
    CB_Año.AddItem "2004"
    CB_Año.AddItem "2003"
    CB_Año.AddItem "2002"
    CB_Año.AddItem "2001"
    CB_Año.AddItem "2000"
    CB_Año.AddItem "1999"
    CB_Año.AddItem "1998"
    CB_Año.AddItem "1997"
    CB_Año.AddItem "1996"
    CB_Año.AddItem "1995"
    CB_Año.AddItem "1994"
    CB_Año.AddItem "1993"
    CB_Año.AddItem "1992"
    CB_Año.AddItem "1991"
    CB_Año.AddItem "1990"
    CB_Año.AddItem "1989"
    CB_Año.AddItem "1988"
    CB_Año.AddItem "1987"
    CB_Año.AddItem "1986"
    CB_Año.AddItem "1985"
    CB_Año.AddItem "1984"
    CB_Año.AddItem "1983"
    CB_Año.AddItem "1982"
    CB_Año.AddItem "1981"
    CB_Año.AddItem "1980"
    CB_Año.AddItem "1979"
    CB_Año.AddItem "1978"
    CB_Año.AddItem "1977"
    CB_Año.AddItem "1976"
    CB_Año.AddItem "1975"
    CB_Año.AddItem "1974"
    CB_Año.AddItem "1973"
    CB_Año.AddItem "1972"
    CB_Año.AddItem "1971"
    CB_Año.AddItem "1970"
    CB_Año.AddItem "1969"
    CB_Año.AddItem "1968"
    CB_Año.AddItem "1967"
    CB_Año.AddItem "1966"
    CB_Año.AddItem "1965"
    CB_Año.AddItem "1964"
    CB_Año.AddItem "1963"
    CB_Año.AddItem "1962"
    CB_Año.AddItem "1961"
    CB_Año.AddItem "1960"
    CB_Año.AddItem "1959"
    CB_Año.AddItem "1958"
    CB_Año.AddItem "1957"
    CB_Año.AddItem "1956"
    CB_Año.AddItem "1955"
    CB_Año.AddItem "1954"
    CB_Año.AddItem "1953"
    CB_Año.AddItem "1952"
    CB_Año.AddItem "1951"
    CB_Año.AddItem "1950"
    CB_Año.AddItem "1949"
    CB_Año.AddItem "1948"
    CB_Año.AddItem "1947"
    CB_Año.AddItem "1946"
    CB_Año.AddItem "1945"
    CB_Año.AddItem "1944"
    CB_Año.AddItem "1943"
    CB_Año.AddItem "1942"
    CB_Año.AddItem "1941"
    CB_Año.AddItem "1940"
    CB_Año.AddItem "1939"
    CB_Año.AddItem "1938"
    CB_Año.AddItem "1937"
    CB_Año.AddItem "1936"
    CB_Año.AddItem "1935"
    CB_Año.AddItem "1934"
    CB_Año.AddItem "1933"
    CB_Año.AddItem "1932"
    CB_Año.AddItem "1931"
    CB_Año.AddItem "1930"
    CB_Año.AddItem "1929"
    CB_Año.AddItem "1928"
    CB_Año.AddItem "1927"
    CB_Año.AddItem "1926"
    CB_Año.AddItem "1925"
    CB_Año.AddItem "1924"
    CB_Año.AddItem "1923"
    CB_Año.AddItem "1922"
    CB_Año.AddItem "1921"
    CB_Año.AddItem "1920"

    CB_Estado.AddItem "Baja California Norte"
    CB_Estado.AddItem "Chihuahua"
    CB_Estado.AddItem "Coahuila"
    CB_Estado.AddItem "Durango"
    CB_Estado.AddItem "Nuevo León"
    CB_Estado.AddItem "Sinaloa"
    CB_Estado.AddItem "Sonora"
    CB_Estado.AddItem "Tamaulipas"

    CB_Pregunta.AddItem "¿Cuál es tu color favorito?"
    CB_Pregunta.AddItem "¿Cuál es tu animal favorito?"
    CB_Pregunta.AddItem "¿Cuál es el mes de nacimiento de tu mamá?"
    CB_Pregunta.AddItem "¿Cuál es tu comida favorita?"
    
End Sub
FIN
Private Sub Btn_CerrarSesión_Click()
    FIN.Hide
    INICIO.Show
End Sub

Private Sub Btn_VolverInicio_Click()
    FIN.Hide
    BIENVENIDO.Show
End Sub


Private Sub UserForm_Activate()
    Dim i As Double
    Dim Cell As String
    For i = 2 To 9999
        Cell = Worksheets("Boletos").Range("A" & i).Value
        If Cell = "" Then
            Dim Código As String
            i = i - 1
            Código = Worksheets("Boletos").Range("K" & i).Value
            txt_Código = Código
            Dim Método As String
            Método = Worksheets("Boletos").Range("J" & i).Value
            If Método = "Tarjeta de Crédito o Débito" Then
                lbl_Método = "Pase a la sala con el siguiente código:"
            End If
            If Método = "Caja" Then
                lbl_Método = "Pase a pagar a caja con el siguiente código:"
            End If
            Exit For
        End If
    Next i
            
        
End Sub
INICIO
Private Sub Btn_CrearUnaCuenta_Click()
    INICIO.Hide
    CUENTA.Show
End Sub

Private Sub Btn_OlvidéMiContraseña_Click()
    INICIO.Hide
    CAMCONTRASEÑA.Show
End Sub

Private Sub lbl_Administrador_Click()
    INICIO.Hide
    ADMINISTRADOR.Show
End Sub

Private Sub UserForm_Activate()
    txt_Usuario = ""
    txt_Contraseña = ""
End Sub

Private Sub Btn_IniciarSesión_Click()
Dim Usuario As String
Dim Contraseña As String
Dim UserOr As String
Dim PwOr As String
Dim Vacío As Boolean
Dim i As Double
Dim Encontrado As Boolean
Encontrado = False

Vacío = False
Usuario = txt_Usuario.Value
Contraseña = txt_Contraseña.Value

If Usuario = "" Then
    Vacío = True
End If
If Contraseña = "" Then
    Vacío = True
End If

If Vacío = True Then
    MsgBox ("Por favor, ingrese el usuario y la contraseña.")
End If

If Vacío = False Then
    For i = 2 To 9999 Step 1
        UserOr = Worksheets("Usuarios").Range("C" & i)
        If UserOr = Usuario Then
            Encontrado = True
            PwOr = Worksheets("Usuarios").Range("D" & i)
            If PwOr = Contraseña Then
                Worksheets("Boletos").Range("M2") = Usuario
                txt_Usuario = ""
                txt_Contraseña = ""
                INICIO.Hide
                BIENVENIDO.Show
            Else
                MsgBox ("El usuario y la contraseña no coinciden.")
                Dim Intentos As Double
                Intentos = Worksheets("Usuarios").Range("S" & i).Value
                Intentos = Intentos + 1
                Worksheets("Usuarios").Range("S" & i) = Intentos
            End If
            
        End If
    Next i
    If Encontrado = False Then
        MsgBox ("No existe una cuenta con ese nombre de usuario.")
    End If
End If

End Sub
MISCOMPRAS
Private Sub Btn_Buscar_Click()
ListBox1.Clear


Application.ScreenUpdating = False
Dim fila As Double
Dim a As Double

UsuarioAc = Worksheets("Boletos").Range("M2").Value
a = 0
fila = 2

While Worksheets("Boletos").Cells(fila, 1) <> Empty
Dim Fecha As Boolean
Dim Hora As Boolean
Dim Película As Boolean
Dim Sucursal As Boolean
Dim Método As Boolean
Dim FechaSel As String
Dim HoraSel As String
Dim PelículaSel As String
Dim SucursalSel As String
Dim MétodoSel As String

Fecha = False
Hora = False
Película = False
Sucursal = False
Método = False

FechaSel = CB_Fecha
HoraSel = CB_Hora
PelículaSel = CB_Película
SucursalSel = CB_Sucursal
MétodoSel = CB_Método

If FechaSel = "" Then
Fecha = True
End If
If HoraSel = "" Then
Hora = True
End If
If PelículaSel = "" Then
Película = True
End If
If SucursalSel = "" Then
Sucursal = True
End If
If MétodoSel = "" Then
Método = True
End If
    If Fecha = False Then
        If FechaSel = Worksheets("Boletos").Range("A" & fila) Then
            Fecha = True
        End If
    End If
    If Hora = False Then
        If HoraSel = Worksheets("Boletos").Range("B" & fila) Then
            Hora = True
        End If
    End If
    If Película = False Then
        If PelículaSel = Worksheets("Boletos").Range("C" & fila) Then
            Película = True
        End If
    End If
    If Sucursal = False Then
        If SucursalSel = Worksheets("Boletos").Range("H" & fila) Then
            Sucursal = True
        End If
    End If
    If Método = False Then
        If MétodoSel = Worksheets("Boletos").Range("J" & fila) Then
            Método = True
        End If
    End If
    If Fecha = True And Hora = True And Película = True And Sucursal = True And Método = True Then
        a = ListBox1.ListCount
        UsuarioOr = Worksheets("Boletos").Cells(fila, 9).Text
        If UsuarioOr = UsuarioAc Then
        ListBox1.AddItem
        ListBox1.List(a, 0) = Worksheets("Boletos").Cells(fila, 1).Text
        ListBox1.List(a, 1) = Worksheets("Boletos").Cells(fila, 2).Text
        ListBox1.List(a, 2) = Worksheets("Boletos").Cells(fila, 3).Text
        ListBox1.List(a, 3) = Worksheets("Boletos").Cells(fila, 5).Text
        ListBox1.List(a, 4) = Worksheets("Boletos").Cells(fila, 6).Text
        ListBox1.List(a, 5) = Worksheets("Boletos").Cells(fila, 7).Text
        ListBox1.List(a, 6) = Worksheets("Boletos").Cells(fila, 8).Text
        ListBox1.List(a, 7) = Worksheets("Boletos").Cells(fila, 10).Text
        ListBox1.List(a, 8) = Worksheets("Boletos").Cells(fila, 11).Text
        End If
fila = fila + 1
    End If
  
fila = fila + 1
Wend

Application.ScreenUpdating = True
End Sub

Private Sub Btn_CerrarSesión_Click()
    MISCOMPRAS.Hide
    INICIO.Show
End Sub

Private Sub Btn_Reiniciar_Click()
ListBox1.Clear
CB_Fecha = ""
CB_Hora = ""
CB_Película = ""
CB_Sucursal = ""
CB_Método = ""

a = 0
fila = 2
UsuarioAc = Worksheets("Boletos").Range("M2").Value
While Worksheets("Boletos").Cells(fila, 1) <> Empty
        a = ListBox1.ListCount
        UsuarioOr = Worksheets("Boletos").Cells(fila, 9).Text
        If UsuarioOr = UsuarioAc Then
        ListBox1.AddItem
        ListBox1.List(a, 0) = Worksheets("Boletos").Cells(fila, 1).Text
        ListBox1.List(a, 1) = Worksheets("Boletos").Cells(fila, 2).Text
        ListBox1.List(a, 2) = Worksheets("Boletos").Cells(fila, 3).Text
        ListBox1.List(a, 3) = Worksheets("Boletos").Cells(fila, 5).Text
        ListBox1.List(a, 4) = Worksheets("Boletos").Cells(fila, 6).Text
        ListBox1.List(a, 5) = Worksheets("Boletos").Cells(fila, 7).Text
        ListBox1.List(a, 6) = Worksheets("Boletos").Cells(fila, 8).Text
        ListBox1.List(a, 7) = Worksheets("Boletos").Cells(fila, 10).Text
        ListBox1.List(a, 8) = Worksheets("Boletos").Cells(fila, 11).Text
        End If
fila = fila + 1
Wend
End Sub

Private Sub Btn_Volver_Click()
    MISCOMPRAS.Hide
    BIENVENIDO.Show
End Sub

Private Sub UserForm_Activate()
ListBox1.ColumnCount = 9
ListBox1.ColumnWidths = "70pt;40pt;130pt;60pt;30pt;82pt;140pt;44pt;50pt"
Dim fila As Double
Dim a As Double
Dim UsuarioOr As String
Dim UsuarioAc As String
UsuarioAc = Worksheets("Boletos").Range("M2").Value

a = 0
fila = 2

While Worksheets("Boletos").Cells(fila, 1) <> Empty
        a = ListBox1.ListCount
        UsuarioOr = Worksheets("Boletos").Cells(fila, 9).Text
        If UsuarioOr = UsuarioAc Then
        ListBox1.AddItem
        ListBox1.List(a, 0) = Worksheets("Boletos").Cells(fila, 1).Text
        ListBox1.List(a, 1) = Worksheets("Boletos").Cells(fila, 2).Text
        ListBox1.List(a, 2) = Worksheets("Boletos").Cells(fila, 3).Text
        ListBox1.List(a, 3) = Worksheets("Boletos").Cells(fila, 5).Text
        ListBox1.List(a, 4) = Worksheets("Boletos").Cells(fila, 6).Text
        ListBox1.List(a, 5) = Worksheets("Boletos").Cells(fila, 7).Text
        ListBox1.List(a, 6) = Worksheets("Boletos").Cells(fila, 8).Text
        ListBox1.List(a, 7) = Worksheets("Boletos").Cells(fila, 10).Text
        ListBox1.List(a, 8) = Worksheets("Boletos").Cells(fila, 11).Text
        End If
fila = fila + 1
Wend

Dim Fila2 As Integer
Dim Final As Integer
Dim Registro As Integer
Fila2 = 2
Do While Worksheets("Boletos").Cells(Fila2, 1) <> ""
    Fila2 = Fila2 + 1
Loop

Final = Fila2 - 1

For Fila2 = 2 To Final
    Registro = WorksheetFunction.CountIf(Range(Worksheets("Boletos").Cells(1, 1), Worksheets("Boletos").Cells(Fila2, 1)), Worksheets("Boletos").Cells(Fila2, 1))
    If Registro = 1 Then
        CB_Fecha.AddItem Worksheets("Boletos").Cells(Fila2, 1).Text
    End If
Next Fila2
For Fila2 = 2 To Final
    Registro = WorksheetFunction.CountIf(Range(Worksheets("Boletos").Cells(1, 2), Worksheets("Boletos").Cells(Fila2, 2)), Worksheets("Boletos").Cells(Fila2, 2))
    If Registro = 1 Then
        CB_Hora.AddItem Worksheets("Boletos").Cells(Fila2, 2).Text
    End If
Next Fila2
For Fila2 = 2 To Final
    Registro = WorksheetFunction.CountIf(Range(Worksheets("Boletos").Cells(1, 3), Worksheets("Boletos").Cells(Fila2, 3)), Worksheets("Boletos").Cells(Fila2, 3))
    If Registro = 1 Then
        CB_Película.AddItem Worksheets("Boletos").Cells(Fila2, 3).Text
    End If
Next Fila2
For Fila2 = 2 To Final
    Registro = WorksheetFunction.CountIf(Range(Worksheets("Boletos").Cells(1, 8), Worksheets("Boletos").Cells(Fila2, 8)), Worksheets("Boletos").Cells(Fila2, 8))
    If Registro = 1 Then
        CB_Sucursal.AddItem Worksheets("Boletos").Cells(Fila2, 8).Text
    End If
Next Fila2
For Fila2 = 2 To Final
    Registro = WorksheetFunction.CountIf(Range(Worksheets("Boletos").Cells(1, 10), Worksheets("Boletos").Cells(Fila2, 10)), Worksheets("Boletos").Cells(Fila2, 10))
    If Registro = 1 Then
        CB_Método.AddItem Worksheets("Boletos").Cells(Fila2, 10).Text
    End If
Next Fila2
End Sub
PERIFLUISECHEVERRÍA
Private Sub CommandButton2_Click()
    PERIFLUISECHEVERRÍA.Hide
End Sub
PREACTUALIZARDATOS
Private Sub Btn_Aceptar_Click()
Dim Usuario As String
Dim Contraseña As String
Dim UserOr As String
Dim PwOr As String
Dim Vacío As Boolean
Dim i As Double

Vacío = False
Usuario = Worksheets("Boletos").Range("M2").Value
Contraseña = txt_Contraseña.Value

If Contraseña = "" Then
    Vacío = True
End If

If Vacío = True Then
    MsgBox ("Por favor, ingrese la contraseña.")
End If

If Vacío = False Then
    For i = 2 To 9999 Step 1
        UserOr = Worksheets("Usuarios").Range("C" & i)
        If UserOr = Usuario Then
            PwOr = Worksheets("Usuarios").Range("D" & i)
            If PwOr = Contraseña Then
                PREACTUALIZARDATOS.Hide
                CONFIGURACIÓN.Hide
                ACTUALIZARDATOS.Show
            Else
                MsgBox ("La contraseña ingresada es errónea.")
                Worksheets("Boletos").Range("M2").Value = ""
                PREACTUALIZARDATOS.Hide
                CONFIGURACIÓN.Hide
                INICIO.Show
            End If
        End If
    Next i
End If

End Sub

Private Sub Btn_Cancelar_Click()
    PREACTUALIZARDATOS.Hide
End Sub
PWTARJETA
Private Sub Btn_Aceptar_Click()
Dim i As Double
Dim UsuarioOr As String
Dim UsuarioAc As String
UsuarioAc = Worksheets("Boletos").Range("M2").Value
Dim Vacío As Boolean
Vacío = False

If txt_Contraseña = "" Then
Vacío = True
End If

If Vacío = True Then
    MsgBox ("Por favor, introduce tu contraseña.")
End If

If Vacío = False Then
For i = 2 To 9999
    UsuarioOr = Worksheets("Usuarios").Range("C" & i).Value
    If UsuarioOr = UsuarioAc Then
        Dim ContraseñaOr As String
        ContraseñaOr = Worksheets("Usuarios").Range("D" & i).Value
        Dim Contraseña2 As String
        Contraseña2 = txt_Contraseña.Value
        If ContraseñaOr = Contraseña2 Then
            PWTARJETA.Hide
            CONFIGURACIÓN.Hide
            CAMTARJETA.Show
        Else
            MsgBox ("La contraseña ingresada es incorrecta.")
        End If
    End If
Next i
End If
            
End Sub

Private Sub Btn_Cancelar_Click()
PWTARJETA.Hide
End Sub
REGTARJETA
Private Sub Btn_Aceptar_Click()
Dim i As Double
Dim Usuario As String
Dim UsuarioAc As String
UsuarioAc = Worksheets("Boletos").Range("M2").Value
Dim Vacío As Boolean
Vacío = False

If txt_Tarjeta = "" Or CB_Mes = "" Or CB_Año = "" Or txt_Clave = "" Then
    Vacío = True
End If

If Vacío = True Then
    MsgBox ("Por favor, llene el formulario con los  datos de su tarjeta para completar la compra.")
End If

If Vacío = False Then
    For i = 2 To 9999
        Usuario = Worksheets("Usuarios").Range("C" & i).Value
        If Usuario = UsuarioAc Then
            Worksheets("Usuarios").Range("N" & i) = txt_Tarjeta
            Worksheets("Usuarios").Range("O" & i) = CB_Mes
            Worksheets("Usuarios").Range("P" & i) = CB_Año
            Worksheets("Usuarios").Range("Q" & i) = txt_Clave
            MsgBox ("Los datos de su tarjeta se han registrado exitosamente.")
            Dim r As Double
            Dim Cell As String
                For r = 2 To 9999
                    Cell = Worksheets("Boletos").Range("A" & r)
                    If Cell = "" Then
                        Worksheets("Boletos").Range("A" & r) = Worksheets("Boletos").Range("N2")
                        Worksheets("Boletos").Range("B" & r) = Worksheets("Boletos").Range("O2")
                        Worksheets("Boletos").Range("C" & r) = Worksheets("Boletos").Range("P2")
                        Worksheets("Boletos").Range("D" & r) = Worksheets("Boletos").Range("Q2")
                        Worksheets("Boletos").Range("E" & r) = Worksheets("Boletos").Range("R2")
                        Worksheets("Boletos").Range("F" & r) = Worksheets("Boletos").Range("S2")
                        Worksheets("Boletos").Range("G" & r) = Worksheets("Boletos").Range("T2")
                        Worksheets("Boletos").Range("H" & r) = Worksheets("Boletos").Range("U2")
                        Worksheets("Boletos").Range("I" & r) = Worksheets("Boletos").Range("M2")
                        Worksheets("Boletos").Range("J" & r) = "Tarjeta de Crédito o Débito"
                        Dim Código As Double
                        Dim Arriba As Double
                        Arriba = r - 1
                        Código = Worksheets("Boletos").Range("K" & Arriba).Value
                        Código = Código + 1
                        Worksheets("Boletos").Range("K" & r) = Código
                        COMPRA.Hide
                        FIN.Show
                        Exit For
                    End If
                Next r
            REGTARJETA.Hide
            FIN.Show
            Exit For
        End If
    Next i
End If
End Sub

Private Sub Btn_Cancelar_Click()
    TARJETA.Hide
    COMPRA.Show
End Sub


Private Sub CB_Mes_Change()
Dim Fecha As String
Fecha = Worksheets("Películas").Range("K2").Value
Dim Mes As Double
Mes = Month(Fecha)
Dim MesSelec As Integer

If CB_Mes = "Enero" Then
    MesSelec = 1
End If
If CB_Mes = "Febrero" Then
    MesSelec = 2
End If
If CB_Mes = "Marzo" Then
    MesSelec = 3
End If
If CB_Mes = "Abril" Then
    MesSelec = 4
End If
If CB_Mes = "Mayo" Then
    MesSelec = 5
End If
If CB_Mes = "Junio" Then
    MesSelec = 6
End If
If CB_Mes = "Julio" Then
    MesSelec = 7
End If
If CB_Mes = "Agosto" Then
    MesSelec = 8
End If
If CB_Mes = "Septiembre" Then
    MesSelec = 9
End If
If CB_Mes = "Octubre" Then
    MesSelec = 10
End If
If CB_Mes = "Noviembre" Then
    MesSelec = 11
End If
If CB_Mes = "Diciembre" Then
    MesSelec = 12
End If

Año = Year(Fecha)
Límite = Año + 10

    If MesSelec < Mes Then
    CB_Año.Clear
    Año = Año + 1
    Do While Año < Límite
        CB_Año.AddItem Año
        Año = Año + 1
    Loop
    End If
    
    If MesSelec >= Mes Then
    CB_Año.Clear
    Do While Año < Límite
        CB_Año.AddItem Año
        Año = Año + 1
    Loop
    End If

End Sub


Private Sub UserForm_Activate()
Dim Fecha As String
Fecha = Worksheets("Películas").Range("K2").Value
Dim Año As Double
Año = Year(Fecha)
Límite = Año + 10

    Do While Año < Límite
        CB_Año.AddItem Año
        Año = Año + 1
    Loop

CB_Mes.AddItem "Enero"
CB_Mes.AddItem "Febrero"
CB_Mes.AddItem "Marzo"
CB_Mes.AddItem "Abril"
CB_Mes.AddItem "Mayo"
CB_Mes.AddItem "Junio"
CB_Mes.AddItem "Julio"
CB_Mes.AddItem "Agosto"
CB_Mes.AddItem "Septiembre"
CB_Mes.AddItem "Octubre"
CB_Mes.AddItem "Noviembre"
CB_Mes.AddItem "Diciembre"


End Sub
SUCURSALES

Private Sub Btn_Acuña_Click()
    ACUÑA.Show
End Sub

Private Sub Btn_CerrarSesión_Click()
    Worksheets("Boletos").Range("N1").Value = ""
    SUCURSALES.Hide
    INICIO.Show
End Sub

Private Sub Btn_PerifLuisEcheverría_Click()
    PERIFLUISECHEVERRÍA.Show
End Sub

Private Sub Btn_VolverInicio_Click()
    SUCURSALES.Hide
    BIENVENIDO.Show
End Sub

Private Sub Btn_Xicoténcatl_Click()
    XICOTÉNCATL.Show
End Sub
TÉRMINOS
Private Sub Btn_Aceptar_Click()
    TERMINOS.Hide
End Sub

Private Sub UserForm_Activate()
    TextBox1.SelStart = 0
End Sub
VENTAS
Private Sub Btn_Buscar_Click()
ListBox1.Clear


Application.ScreenUpdating = False
Dim fila As Double
Dim a As Double


a = 0
fila = 2

While Worksheets("Boletos").Cells(fila, 1) <> Empty
Dim Fecha As Boolean
Dim Hora As Boolean
Dim Película As Boolean
Dim Sucursal As Boolean
Dim Método As Boolean
Dim Usuario As Boolean
Dim FechaSel As String
Dim HoraSel As String
Dim PelículaSel As String
Dim SucursalSel As String
Dim MétodoSel As String
Dim UsuarioSel As String

Fecha = False
Hora = False
Película = False
Sucursal = False
Método = False
Usuario = False

FechaSel = CB_Fecha
HoraSel = CB_Hora
PelículaSel = CB_Película
SucursalSel = CB_Sucursal
MétodoSel = CB_Método
UsuarioSel = CB_Usuario

If FechaSel = "" Then
Fecha = True
End If
If HoraSel = "" Then
Hora = True
End If
If PelículaSel = "" Then
Película = True
End If
If SucursalSel = "" Then
Sucursal = True
End If
If MétodoSel = "" Then
Método = True
End If
If UsuarioSel = "" Then
Usuario = True
End If
    If Fecha = False Then
        If FechaSel = Worksheets("Boletos").Range("A" & fila) Then
            Fecha = True
        End If
    End If
    If Hora = False Then
        If HoraSel = Worksheets("Boletos").Range("B" & fila) Then
            Hora = True
        End If
    End If
    If Película = False Then
        If PelículaSel = Worksheets("Boletos").Range("C" & fila) Then
            Película = True
        End If
    End If
    If Sucursal = False Then
        If SucursalSel = Worksheets("Boletos").Range("H" & fila) Then
            Sucursal = True
        End If
    End If
    If Método = False Then
        If MétodoSel = Worksheets("Boletos").Range("J" & fila) Then
            Método = True
        End If
    End If
    If Usuario = False Then
        If UsuarioSel = Worksheets("Boletos").Range("I" & fila) Then
            Usuario = True
        End If
    End If
    If Fecha = True And Hora = True And Película = True And Sucursal = True And Método = True And Usuario = True Then
        a = ListBox1.ListCount
        ListBox1.AddItem
        ListBox1.List(a, 0) = Worksheets("Boletos").Cells(fila, 1).Text
        ListBox1.List(a, 1) = Worksheets("Boletos").Cells(fila, 2).Text
        ListBox1.List(a, 2) = Worksheets("Boletos").Cells(fila, 3).Text
        ListBox1.List(a, 3) = Worksheets("Boletos").Cells(fila, 5).Text
        ListBox1.List(a, 4) = Worksheets("Boletos").Cells(fila, 6).Text
        ListBox1.List(a, 5) = Worksheets("Boletos").Cells(fila, 7).Text
        ListBox1.List(a, 6) = Worksheets("Boletos").Cells(fila, 8).Text
        ListBox1.List(a, 7) = Worksheets("Boletos").Cells(fila, 9).Text
        ListBox1.List(a, 8) = Worksheets("Boletos").Cells(fila, 10).Text
        ListBox1.List(a, 9) = Worksheets("Boletos").Cells(fila, 11).Text
    End If
  
fila = fila + 1
Wend

Application.ScreenUpdating = True

End Sub

Private Sub Btn_CerrarSesión_Click()
    VENTAS.Hide
    INICIO.Show
End Sub

Private Sub Btn_Reiniciar_Click()
ListBox1.Clear
CB_Fecha = ""
CB_Hora = ""
CB_Película = ""
CB_Sucursal = ""
CB_Método = ""
CB_Usuario = ""

a = 0
fila = 2

While Worksheets("Boletos").Cells(fila, 1) <> Empty
        a = ListBox1.ListCount
        ListBox1.AddItem
        ListBox1.List(a, 0) = Worksheets("Boletos").Cells(fila, 1).Text
        ListBox1.List(a, 1) = Worksheets("Boletos").Cells(fila, 2).Text
        ListBox1.List(a, 2) = Worksheets("Boletos").Cells(fila, 3).Text
        ListBox1.List(a, 3) = Worksheets("Boletos").Cells(fila, 5).Text
        ListBox1.List(a, 4) = Worksheets("Boletos").Cells(fila, 6).Text
        ListBox1.List(a, 5) = Worksheets("Boletos").Cells(fila, 7).Text
        ListBox1.List(a, 6) = Worksheets("Boletos").Cells(fila, 8).Text
        ListBox1.List(a, 7) = Worksheets("Boletos").Cells(fila, 9).Text
        ListBox1.List(a, 8) = Worksheets("Boletos").Cells(fila, 10).Text
        ListBox1.List(a, 9) = Worksheets("Boletos").Cells(fila, 11).Text
fila = fila + 1
Wend
End Sub

Private Sub Btn_Volver_Click()
    VENTAS.Hide
    ADMINISTRACIÓN.Show
End Sub

Private Sub UserForm_Activate()
ListBox1.ColumnCount = 10
ListBox1.ColumnWidths = "70pt;40pt;130pt;60pt;30pt;82pt;80pt;60pt;44pt;50pt"
Dim fila As Double
Dim a As Double
a = 0
fila = 2

While Worksheets("Boletos").Cells(fila, 1) <> Empty
        a = ListBox1.ListCount
        ListBox1.AddItem
        ListBox1.List(a, 0) = Worksheets("Boletos").Cells(fila, 1).Text
        ListBox1.List(a, 1) = Worksheets("Boletos").Cells(fila, 2).Text
        ListBox1.List(a, 2) = Worksheets("Boletos").Cells(fila, 3).Text
        ListBox1.List(a, 3) = Worksheets("Boletos").Cells(fila, 5).Text
        ListBox1.List(a, 4) = Worksheets("Boletos").Cells(fila, 6).Text
        ListBox1.List(a, 5) = Worksheets("Boletos").Cells(fila, 7).Text
        ListBox1.List(a, 6) = Worksheets("Boletos").Cells(fila, 8).Text
        ListBox1.List(a, 7) = Worksheets("Boletos").Cells(fila, 9).Text
        ListBox1.List(a, 8) = Worksheets("Boletos").Cells(fila, 10).Text
        ListBox1.List(a, 9) = Worksheets("Boletos").Cells(fila, 11).Text
fila = fila + 1
Wend

Dim Fila2 As Integer
Dim Final As Integer
Dim Registro As Integer
Fila2 = 2
Do While Worksheets("Boletos").Cells(Fila2, 1) <> ""
    Fila2 = Fila2 + 1
Loop

Final = Fila2 - 1

For Fila2 = 2 To Final
    Registro = WorksheetFunction.CountIf(Range(Worksheets("Boletos").Cells(1, 1), Worksheets("Boletos").Cells(Fila2, 1)), Worksheets("Boletos").Cells(Fila2, 1))
    If Registro = 1 Then
        CB_Fecha.AddItem Worksheets("Boletos").Cells(Fila2, 1).Text
    End If
Next Fila2
For Fila2 = 2 To Final
    Registro = WorksheetFunction.CountIf(Range(Worksheets("Boletos").Cells(1, 2), Worksheets("Boletos").Cells(Fila2, 2)), Worksheets("Boletos").Cells(Fila2, 2))
    If Registro = 1 Then
        CB_Hora.AddItem Worksheets("Boletos").Cells(Fila2, 2).Text
    End If
Next Fila2
For Fila2 = 2 To Final
    Registro = WorksheetFunction.CountIf(Range(Worksheets("Boletos").Cells(1, 3), Worksheets("Boletos").Cells(Fila2, 3)), Worksheets("Boletos").Cells(Fila2, 3))
    If Registro = 1 Then
        CB_Película.AddItem Worksheets("Boletos").Cells(Fila2, 3).Text
    End If
Next Fila2
For Fila2 = 2 To Final
    Registro = WorksheetFunction.CountIf(Range(Worksheets("Boletos").Cells(1, 8), Worksheets("Boletos").Cells(Fila2, 8)), Worksheets("Boletos").Cells(Fila2, 8))
    If Registro = 1 Then
        CB_Sucursal.AddItem Worksheets("Boletos").Cells(Fila2, 8).Text
    End If
Next Fila2
For Fila2 = 2 To Final
    Registro = WorksheetFunction.CountIf(Range(Worksheets("Boletos").Cells(1, 9), Worksheets("Boletos").Cells(Fila2, 9)), Worksheets("Boletos").Cells(Fila2, 9))
    If Registro = 1 Then
        CB_Usuario.AddItem Worksheets("Boletos").Cells(Fila2, 9).Text
    End If
Next Fila2
For Fila2 = 2 To Final
    Registro = WorksheetFunction.CountIf(Range(Worksheets("Boletos").Cells(1, 10), Worksheets("Boletos").Cells(Fila2, 10)), Worksheets("Boletos").Cells(Fila2, 10))
    If Registro = 1 Then
        CB_Método.AddItem Worksheets("Boletos").Cells(Fila2, 10).Text
    End If
Next Fila2

End Sub
XICOTÉNCATL
Private Sub CommandButton2_Click()
    XICOTÉNCATL.Hide
End Sub





