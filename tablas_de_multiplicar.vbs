Do
a = InputBox("Introduce numero a sacar tabla", "Las tablas")
arr = Array()
For i = 0 To 10 Step 1
	Redim Preserve arr(Ubound(arr)+1)
	arr(ubound(arr))=(i & " x " & a & " = " & (i * a))
Next
titulo = "*************** Tabla del " & a & " ***************"
todo = arr(1) & chr(10) & arr(2) & chr(10) & arr(3) & chr(10) & arr(4) & chr(10) & arr(5) & chr(10) & arr(6) & chr(10) & arr(7) & chr(10) & arr(8) & chr(10) & arr(9) & chr(10) & arr(10)
MsgBox(todo), vbinform, titulo
Loop Until a = 1313