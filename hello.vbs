a = Abs(InputBox("Escribe a: ", "Calculadora"))
b = Abs(InputBox("Escribe b: ", "Calculadora"))

'Suma
sum_res = a + b
sum = a & " + " & b & " = " & sum_res & chr(10)

'Resta
res_res = a - b
res = a & " - " & b & " = " & res_res & chr(10)

'Multiplicación
mul_res = a * b
mul = a & " * " & b & " = " & mul_res & chr(10)

'Divicion
div_res = a / b
div = a & " / " & b & " = " & div_res & chr(10)

'Ciclo For
'For i = 0 To 20 Step 1	
'Next

MsgBox(sum & res & mul & div), vbalert, "********** Calculadora de todos los operadores **********"