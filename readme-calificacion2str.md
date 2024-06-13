# e-scolar módulo VBA - conversion número a String
Créditos a [fitorec](https://github.com/fitorec) en [GitHub Gist](https://gist.github.com/d2e29d81019610db49bb.git)

Un módulo en Visual Basic para agregar al **excel de MS Office** y **calc de LibreOffice** la funcionalidad de convertir un número a letra en formato de calificación, ejemplos de conversiones realizadas:

	10.00	DIEZ PUNTO CERO
	9.80	NUEVE PUNTO OCHENTA
	7.30	SIETE PUNTO TREINTA
	9.20	NUEVE PUNTO VEINTE
	9.40	NUEVE PUNTO CUARENTA
	6.90	SEIS PUNTO NOVENTA
	6.70	SEIS PUNTO SETENTA
	10.00	DIEZ PUNTO CERO
	4.60	CUATRO PUNTO SESENTA
	7.75	SIETE PUNTO SETENTA Y CINCO
	9.60	NUEVE PUNTO SESENTA
	9.20	NUEVE PUNTO VEINTE

## Use

Agregar la función en el `editor de VisualBasic`, para esto te vas a Herramientas-Macro-Editor de VisualBasic- de ahí les abrirá el editor, ya en el editor se van a Insertar-Modulo, agregas el código y lo usas en la siguiente forma.

### Conversión simple

En una celda le das en formula y agregas la siguiente conversión:

	=CALIFICACION2STR(6.5)

### Conversión de referencia

Por ejemplo si queremos convertir la celda B3, entonces agregamos la siguiente formula:

	=CALIFICACION2STR(B3)


### Ejemplos mas complejos y anidaciones

Finalmente podemos anidas funciones, por ejemplo:

#### sacando un promedio de un columna(`F`) y mostrarlo en letra:

	=CALIFICACION2STR(MAYÚSC(=PROMEDIO(F2:F20))

#### Conversión a mayusculas de las letras de la celda F9:

	=MAYÚSC(CALIFICACION2STR(F9))
