Rem Attribute VBA_ModuleType=VBAModule
Function CALIFICACION2STR(Numero As Double) As String

Dim Preposicion As String
Dim NumCentimos As Double
Dim Letra As String
Const Mxaimo = 1999999999.99

'************************************************************
' Parámetros
' -- CALIFICACION2STR --
' Créditos a fitorec
' https://gist.github.com/d2e29d81019610db49bb.git
'************************************************************
Preposicion = "punto"     'Preposición entre Decimos y Céntimos
'************************************************************

'Validar que el Numero está dentro de los límites
If (Numero >= 0) And (Numero <= Maximo) Then

    Letra = NUMERORECURSIVO((Fix(Numero)))              'Convertir el Numero en letras
    NumCentimos = Round((Numero - Fix(Numero)) * 100)   'Obtener los centimos del Numero

    'Si NumCentimos es mayor a cero inicar la conversión
    If NumCentimos >= 0 Then
        'Obtenemos en letras para los céntimos
        Letra = Letra & " " & Preposicion & " " & NUMERORECURSIVO(Fix(NumCentimos)) 'Convertir los céntimos en letra
    End If

    'Regresar el resultado final de la conversión
    CALIFICACION2STR = Letra

Else
    'Si el Numero no está dentro de los límites, entivar un mensaje de error
    CALIFICACION2STR = "ERROR: El número excede los límites."
End If

End Function

Function NUMERORECURSIVO(Numero As Long) As String

Dim Unidades, Decenas, Centenas
Dim Resultado As String

'**************************************************
' Nombre de los números
'**************************************************
Unidades = Array("", "Un", "Dos", "Tres", "Cuatro", "Cinco", "Seis", "Siete", "Ocho", "Nueve", "Diez", "Once", "Doce", "Trece", "Catorce", "Quince", "Dieciséis", "Diecisiete", "Dieciocho", "Diecinueve", "Veinte", "Veintiuno", "Veintidos", "Veintitres", "Veinticuatro", "Veinticinco", "Veintiseis", "Veintisiete", "Veintiocho", "Veintinueve")
Decenas = Array("", "Diez", "Veinte", "Treinta", "Cuarenta", "Cincuenta", "Sesenta", "Setenta", "Ochenta", "Noventa", "Cien")
Centenas = Array("", "Ciento", "Doscientos", "Trescientos", "Cuatrocientos", "Quinientos", "Seiscientos", "Setecientos", "Ochocientos", "Novecientos")
'**************************************************

Select Case Numero
    Case 0
        Resultado = "Cero"
    Case 1 To 29
        Resultado = Unidades(Numero)
    Case 30 To 100
        Resultado = Decenas(Numero \ 10) + IIf(Numero Mod 10 <> 0, " y " + NUMERORECURSIVO(Numero Mod 10), "")
    Case 101 To 999
        Resultado = Centenas(Numero \ 100) + IIf(Numero Mod 100 <> 0, " " + NUMERORECURSIVO(Numero Mod 100), "")
    Case 1000 To 1999
        Resultado = "Mil" + IIf(Numero Mod 1000 <> 0, " " + NUMERORECURSIVO(Numero Mod 1000), "")
    Case 2000 To 999999
        Resultado = NUMERORECURSIVO(Numero \ 1000) + " Mil" + IIf(Numero Mod 1000 <> 0, " " + NUMERORECURSIVO(Numero Mod 1000), "")
    Case 1000000 To 1999999
        Resultado = "Un Millón" + IIf(Numero Mod 1000000 <> 0, " " + NUMERORECURSIVO(Numero Mod 1000000), "")
    Case 2000000 To 1999999999
        Resultado = NUMERORECURSIVO(Numero \ 1000000) + " Millones" + IIf(Numero Mod 1000000 <> 0, " " + NUMERORECURSIVO(Numero Mod 1000000), "")
End Select

NUMERORECURSIVO = Resultado

End Function
