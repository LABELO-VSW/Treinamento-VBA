Attribute VB_Name = "Repeticao"
Sub doWhile()
    Dim i As Integer: i = 0
    
    Do
        i = i + 1
        Debug.Print i
    Loop While i < 10 'Ir� contar de 1 a 10 e imprimir na janela de verifica��o
End Sub
Sub whileEx()
Dim i As Integer: i = 0
    'A diferen�a em rela��o ao do While, � que este verifica
    'se i � menor que 10 antes da primeira repeti��o
    While i < 10 'Ir� contar de 1 a 10 e imprimir na janela de verifica��o
        i = i + 1
        Debug.Print i
    Wend
End Sub

Sub forEx()
    Dim i As Integer
    'inica i em 1, incrementa 2 a cada loop at� i ser maior que 10
    For i = 1 To 10 Step 2 'incrementa de dois em dois
        Debug.Print i
    Next i
End Sub

Sub gotoEx()

    Dim i As Integer: i = 0
start:      'Nome do label escolhido pelo usu�rio
    Debug.Print i
    i = i + 1
If i < 10 Then GoTo start
End Sub
