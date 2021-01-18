Attribute VB_Name = "Operacoes_Aritm�ticas"
Function SOMA(a As Integer, b As Integer) As Integer
    'SOMA � uma fun��o que retorna um valor inteiro
    'que � o resultado da soma de a e b.
    SOMA = a + b
End Function

Function MULTIPLICACAO(a As Integer, b As Integer) As Integer
    'MULTIPLICACAO � uma fun��o que retorna um valor inteiro
    'que � o resultado da multiplica��o entre a e b.
    MULTIPLICACAO = a * b
End Function

Function SUBTRACAO(a As Integer, b As Integer) As Integer
    'SUBTRACAO � uma fun��o que retorna um valor inteiro
    'que � o resultado da subtra��o entre a e b.
    SUBTRACAO = a - b
End Function

Function DIVISAO(a As Integer, b As Integer) As Integer
    'DIVISAO � uma fun��o que retorna um valor inteiro
    'que � o resultado da divis�o de a por b.
    DIVISAO = a / b
End Function

Function DIVISAO_COM_VIRGULA(a As Integer, b As Integer) As Single
'Single = ponto flutuante de precis�o simples IEEE 32 bits - equivalente ao float do C
    DIVISAO_COM_VIRGULA = a / b
End Function

Sub Plot_de_Resultados()

    Debug.Print ("Soma: " & SOMA(10, 3)) 'Concaten��o da palavra Soma com o resultado
    
    Debug.Print ("Multiplica��o: " & MULTIPLICACAO(10, 3))
    
    Debug.Print ("Subtra��o: " & SUBTRACAO(10, 3))
    
    Debug.Print ("Divis�o (Fun��o 1): " & DIVISAO(10, 3)) ' Repare que o resultado � inteiro,
                                                          ' devido a forma como a fun��o foi declarada
    
    Debug.Print ("Divis�o (Fun��o 2): " & DIVISAO_COM_VIRGULA(10, 3))
    
    'Validando uma fun��o
    Debug.Assert (SOMA(5, 5) = 10) 'A 5+5 deve ser igual a 10
    
    
End Sub

