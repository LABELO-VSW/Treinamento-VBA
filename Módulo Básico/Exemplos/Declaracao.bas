Attribute VB_Name = "Declaracao"
Sub Declaracao_Variaveis_de_Texto() 'Sub-Rotina p�blica, portanto pode ser utilizada por outros m�dulos

    Dim texto As String 'Vari�vel din�mica do tipo string
                        'a qual pode-se alterar valor durante o programa
    
    texto = "Ol� Mundo!" 'Atribuindo um valor
    
    texto = texto & "Texto adcional" 'concatenando um valor adcional a vari�vel
    
    'Atribuindo valor na mesma linha
    Dim msgResult As VbMsgBoxResult: msgResult = MsgBox(texto, vbInformation, "T�tulo")
    
    'Fun��es de teste
    'Janela de verifica��o imediata
    Debug.Print msgResult = vbOK 'verifica se foi apertado o bot�o de OK
    
    'Assert: o debugador ir� parar o programa caso o teste reprove
    Debug.Assert (msgResult = vbOK) 'Trocar o sinal de igual por <> para ver funcionalide do Assert
    
End Sub

Sub declaracao_variaveis_numericas()
    Dim b As Boolean: b = True 'valor booleando true or false ( Verdadeiro ou Falso)
    
    Dim bb As Byte: bb = 1 'valor inteiro positivo de 0 a 255
    Dim i As Integer: i = 1 'valor inteiro de 32 bits
    Dim l As Long: d = 2.41241241241241E+28 'Valor inteiro com 64 bits
    
    Dim s As Single: b = 2.1 'valor com ponto flutuante de 32 bits
    Dim d As Double: c = 2.23124124214124 'valor com ponto flutuante de 64 bits
    
End Sub

