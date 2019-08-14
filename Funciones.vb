Module Funciones
    Dim i As Integer
    Public Function numeros(ByVal n, ByVal vector)
        For i = 0 To n
            vector(i) = CInt(Int((10) * Rnd() + 48))
        Next
        Return vector
    End Function
    Public Function letrasMayus(ByVal n, ByVal vector)
        For i = 0 To n
            vector(i) = CInt(Int((26) * Rnd() + 65))
        Next
        Return vector
    End Function
    Public Function letrasMinusculas(ByVal n, ByVal vector)
        For i = 0 To n
            vector(i) = CInt(Int((26) * Rnd() + 97))
        Next
        Return vector
    End Function
    Public Function letrasAmbas(ByVal n, ByVal vector)
        'Vamos a rellenar el vector con 0 y 1, 0 para una letra en mayuscula y 1 para una letra minuscula
        For i = 0 To n
            vector(i) = CInt(Int((2) * Rnd()))
        Next
        For i = 0 To n
            If (vector(i) = 0) Then
                vector(i) = CInt(Int((26) * Rnd() + 65))
            Else
                vector(i) = CInt(Int((26) * Rnd() + 97))
            End If
        Next
        Return vector
    End Function
    Public Function numerosLetraMay(ByVal n, ByVal vector)
        'Vamos a rellenar el vector con 0 y 1, 0 para numeros y 1 letra mayuscula
        For i = 0 To n
            vector(i) = CInt(Int((2) * Rnd()))
        Next
        For i = 0 To n
            If (vector(i) = 0) Then
                vector(i) = CInt(Int((10) * Rnd() + 48))
            Else
                vector(i) = CInt(Int((26) * Rnd() + 65))
            End If
        Next
        Return vector
    End Function
    Public Function numerosLetraMin(ByVal n, ByVal vector)
        'Vamos a rellenar el vector con 0 y 1, 0 para numeros y 1 letra minuscula
        For i = 0 To n
            vector(i) = CInt(Int((2) * Rnd()))
        Next
        For i = 0 To n
            If (vector(i) = 0) Then
                vector(i) = CInt(Int((10) * Rnd() + 48))
            Else
                vector(i) = CInt(Int((26) * Rnd() + 97))
            End If
        Next
        Return vector
    End Function
    Public Function numerosLetraAmbas(ByVal n, ByVal vector)
        'Vamos a rellenar el vector con 0, 1 y 2. 0 para numeros, 1 letra mayuscula, 2 letras mayusculas.
        For i = 0 To n
            vector(i) = CInt(Int((3) * Rnd()))
        Next
        For i = 0 To n
            'Numeros
            If (vector(i) = 0) Then
                vector(i) = CInt(Int((10) * Rnd() + 48))
            End If
            If (vector(i) = 1) Then
                vector(i) = CInt(Int((26) * Rnd() + 65))
            End If
            If (vector(i) = 2) Then
                vector(i) = CInt(Int((26) * Rnd() + 97))
            End If
        Next
        Return vector
    End Function
    Public Function caracteres(ByVal n, ByVal vector)
        'Vamos a rellenar el vector con 0 y 1. 0 para la primera parte de caracteres 1 para la segunda parte y 2 para la tercera parte.
        For i = 0 To n
            vector(i) = CInt(Int((3) * Rnd()))
        Next
        For i = 0 To n
            If (vector(i) = 0) Then
                vector(i) = CInt(Int((15) * Rnd() + 33))
            End If
            If (vector(i) = 1) Then
                vector(i) = CInt(Int((7) * Rnd() + 58))
            End If
            If (vector(i) = 2) Then
                vector(i) = CInt(Int((6) * Rnd() + 91))
            End If
        Next
        Return vector
    End Function
    Public Function caracteresLetrasMayus(ByVal n, ByVal vector)
        'Vamos a rellenar el vector con 0, 1 y 2. 0 para la primera parte de caracteres 1 para la segunda parte y 2 para la tercera parte.
        For i = 0 To n
            vector(i) = CInt(Int((4) * Rnd()))
        Next
        For i = 0 To n
            If (vector(i) = 0) Then
                vector(i) = CInt(Int((15) * Rnd() + 33))
            End If
            If (vector(i) = 1) Then
                vector(i) = CInt(Int((7) * Rnd() + 58))
            End If
            If (vector(i) = 2) Then
                vector(i) = CInt(Int((6) * Rnd() + 91))
            End If
            If (vector(i) = 3) Then
                vector(i) = CInt(Int((26) * Rnd() + 65))
            End If
        Next
        Return vector
    End Function
    Public Function caracteresLetrasMinus(ByVal n, ByVal vector)
        'Vamos a rellenar el vector con 0 y 1. 0 para la primera parte de caracteres 1 para la segunda parte y 2 para la tercera parte.
        For i = 0 To n
            vector(i) = CInt(Int((4) * Rnd()))
        Next
        For i = 0 To n
            If (vector(i) = 0) Then
                vector(i) = CInt(Int((15) * Rnd() + 33))
            End If
            If (vector(i) = 1) Then
                vector(i) = CInt(Int((7) * Rnd() + 58))
            End If
            If (vector(i) = 2) Then
                vector(i) = CInt(Int((6) * Rnd() + 91))
            End If
            If (vector(i) = 3) Then
                vector(i) = CInt(Int((26) * Rnd() + 97))
            End If
        Next
        Return vector
    End Function
    Public Function caracteresLetrasAmbas(ByVal n, ByVal vector)
        For i = 0 To n
            vector(i) = CInt(Int((5) * Rnd()))
        Next
        For i = 0 To n
            'Caracteres especiales 1
            If (vector(i) = 0) Then
                vector(i) = CInt(Int((15) * Rnd() + 33))
            End If
            'Caracteres especiales 2
            If (vector(i) = 1) Then
                vector(i) = CInt(Int((7) * Rnd() + 58))
            End If
            'Caracteres especiales 3
            If (vector(i) = 2) Then
                vector(i) = CInt(Int((6) * Rnd() + 91))
            End If
            'Letras mayusculas
            If (vector(i) = 3) Then
                vector(i) = CInt(Int((26) * Rnd() + 65))
            End If
            'Letras minusculas
            If (vector(i) = 4) Then
                vector(i) = CInt(Int((26) * Rnd() + 97))
            End If
        Next
        Return vector
    End Function
    Public Function caracteresNumeros(ByVal n, ByVal vector)
        'Vamos a rellenar el vector con 0 y 1. 0 para la primera parte de caracteres 1 para la segunda parte y 2 para la tercera parte.
        For i = 0 To n
            vector(i) = CInt(Int((4) * Rnd()))
        Next
        For i = 0 To n
            'Caracteres especiales 1
            If (vector(i) = 0) Then
                vector(i) = CInt(Int((15) * Rnd() + 33))
            End If
            'Caracteres especiales 2
            If (vector(i) = 1) Then
                vector(i) = CInt(Int((7) * Rnd() + 58))
            End If
            'Caracteres especiales 3
            If (vector(i) = 2) Then
                vector(i) = CInt(Int((6) * Rnd() + 91))
            End If
            'Numeros
            If (vector(i) = 3) Then
                vector(i) = CInt(Int((10) * Rnd() + 48))
            End If
        Next
        Return vector
    End Function
    Public Function caracteresNumerosLetrasMayus(ByVal n, ByVal vector)
        'TODO
        'Vamos a rellenar el vector con 0 y 1. 0 para la primera parte de caracteres 1 para la segunda parte y 2 para la tercera parte.
        For i = 0 To n
            vector(i) = CInt(Int((5) * Rnd()))
        Next
        For i = 0 To n
            'Numeros
            If (vector(i) = 0) Then
                vector(i) = CInt(Int((10) * Rnd() + 48))
            End If
            'Caracteres especiales 1
            If (vector(i) = 1) Then
                vector(i) = CInt(Int((15) * Rnd() + 33))
            End If
            'Caracteres especiales 2
            If (vector(i) = 2) Then
                vector(i) = CInt(Int((7) * Rnd() + 58))
            End If
            'Caracteres especiales 3
            If (vector(i) = 3) Then
                vector(i) = CInt(Int((6) * Rnd() + 91))
            End If
            'Letras mayusculas
            If (vector(i) = 4) Then
                vector(i) = CInt(Int((26) * Rnd() + 65))
            End If
        Next
        Return vector
    End Function
    Public Function caracteresNumerosLetrasMinus(ByVal n, ByVal vector)
        'TODO
        'Vamos a rellenar el vector con 0 y 1. 0 para la primera parte de caracteres 1 para la segunda parte y 2 para la tercera parte.
        For i = 0 To n
            vector(i) = CInt(Int((5) * Rnd()))
        Next
        For i = 0 To n
            'Numeros
            If (vector(i) = 0) Then
                vector(i) = CInt(Int((10) * Rnd() + 48))
            End If
            'Caracteres especiales 1
            If (vector(i) = 1) Then
                vector(i) = CInt(Int((15) * Rnd() + 33))
            End If
            'Caracteres especiales 2
            If (vector(i) = 2) Then
                vector(i) = CInt(Int((7) * Rnd() + 58))
            End If
            'Caracteres especiales 3
            If (vector(i) = 3) Then
                vector(i) = CInt(Int((6) * Rnd() + 91))
            End If
            'Letras minusculas
            If (vector(i) = 4) Then
                vector(i) = CInt(Int((26) * Rnd() + 97))
            End If
        Next
        Return vector
    End Function
    Public Function caracteresNumerosLetrasAmbas(ByVal n, ByVal vector)
        'TODO
        'Vamos a rellenar el vector con 0 y 1. 0 para la primera parte de caracteres 1 para la segunda parte y 2 para la tercera parte.
        For i = 0 To n
            vector(i) = CInt(Int((6) * Rnd()))
        Next
        For i = 0 To n
            'Numeros
            If (vector(i) = 0) Then
                vector(i) = CInt(Int((10) * Rnd() + 48))
            End If
            'Caracteres especiales 1
            If (vector(i) = 1) Then
                vector(i) = CInt(Int((15) * Rnd() + 33))
            End If
            'Caracteres especiales 2
            If (vector(i) = 2) Then
                vector(i) = CInt(Int((7) * Rnd() + 58))
            End If
            'Caracteres especiales 3
            If (vector(i) = 3) Then
                vector(i) = CInt(Int((6) * Rnd() + 91))
            End If
            'Letras mayusculas
            If (vector(i) = 4) Then
                vector(i) = CInt(Int((26) * Rnd() + 65))
            End If
            'Letras minusculas
            If (vector(i) = 5) Then
                vector(i) = CInt(Int((26) * Rnd() + 97))
            End If
        Next
        Return vector
    End Function
End Module
