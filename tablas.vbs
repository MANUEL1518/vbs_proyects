While c = False
a = Abs(InputBox("Oprime 1313 para salir" & chr(10) & "Escribe un numero para ver sus tablas: ", "Tablas de multiplicar", "", 1500,1500))
If a = "" or a = 1313 or a = 0 Then c = True
h= array ()
For i = 0 To 20 Step 1
	Redim Preserve h(Ubound(h)+1)
	h(ubound(h))=(i & " X " & a & " = " & (a * i))
Next
entr = chr(10)
tabla =  h(1) & entr & h(2) & entr & h(3) & entr & h(4) & entr & h(5) & entr & h(6) & entr & h(7) & entr & h(8) & entr & h(9) & entr &  h(10)
titulo = "********** Las tablas del " & a & " **********"
MsgBox(tabla), vbalert, titulo
Wend