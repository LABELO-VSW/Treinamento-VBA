Attribute VB_Name = "Operacoes_Aritméticas"
Function SOMA(a As Integer, b As Integer) As Integer
    'SOMA é uma função que retorna um valor inteiro
    'que é o resultado da soma de a e b.
    SOMA = a + b
End Function

Function MULTIPLICACAO(a As Integer, b As Integer) As Integer
    'MULTIPLICACAO é uma função que retorna um valor inteiro
    'que é o resultado da multiplicação entre a e b.
    MULTIPLICACAO = a * b
End Function

Function SUBTRACAO(a As Integer, b As Integer) As Integer
    'SUBTRACAO é uma função que retorna um valor inteiro
    'que é o resultado da subtração entre a e b.
    SUBTRACAO = a - b
End Function

Function DIVISAO(a As Integer, b As Integer) As Integer
    'DIVISAO é uma função que retorna um valor inteiro
    'que é o resultado da divisão de a por b.
    DIVISAO = a / b
End Function

Function DIVISAO_COM_VIRGULA(a As Integer, b As Integer) As Single
'Single = ponto flutuante de precisão simples IEEE 32 bits - equivalente ao float do C
    DIVISAO_COM_VIRGULA = a / b
End Function

Sub Plot_de_Resultados()

    Debug.Print ("Soma: " & SOMA(10, 3)) 'Concatenção da palavra Soma com o resultado
    
    Debug.Print ("Multiplicação: " & MULTIPLICACAO(10, 3))
    
    Debug.Print ("Subtração: " & SUBTRACAO(10, 3))
    
    Debug.Print ("Divisão (Função 1): " & DIVISAO(10, 3)) ' Repare que o resultado é inteiro,
                                                          ' devido a forma como a função foi declarada
    
    Debug.Print ("Divisão (Função 2): " & DIVISAO_COM_VIRGULA(10, 3))
    
    'Validando uma função
    Debug.Assert (SOMA(5, 5) = 10) 'A 5+5 deve ser igual a 10
    
    
End Sub

