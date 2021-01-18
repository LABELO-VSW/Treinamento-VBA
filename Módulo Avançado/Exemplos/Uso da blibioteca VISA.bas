Attribute VB_Name = "VisaCom"
'**********************************************************************************************************************
'Autor: Jonathan Culau
'Versão: 01
'30/03/2019
'********************************************************************************************************************
Option Explicit

Enum TerminationCharacter
    CR = &HD
    LF = &HA
    NoCharacter = -1
End Enum


Enum VisaParity
    Even = ASRL_PAR_EVEN
    Mark = ASRL_PAR_MARK
    none = ASRL_PAR_NONE
    Space = ASRL_PAR_SPACE
    Odd = ASRL_PAR_ODD
End Enum

Enum VisaStopBits
    TWO = ASRL_STOP_TWO
    ONE = ASRL_STOP_ONE
    ONE5 = ASRL_STOP_ONE5
End Enum


Function VisaOpen(VisaAddress As String, Optional TimeOut As Long = 2000, _
Optional TerminationCharacter As TerminationCharacter = TerminationCharacter.LF, Optional Baudrate As Long, _
 Optional DataSize As Integer = 8, Optional Parity As VisaParity = none, Optional StopBits As VisaStopBits = ONE) As VisaComLib.FormattedIO488
    
    Dim Iomgr As New VisaComLib.ResourceManager
    
    Set VisaOpen = New VisaComLib.FormattedIO488
    
    
    Set VisaOpen.IO = Iomgr.Open(VisaAddress, NO_LOCK, TimeOut) 'open communication
    
    VisaOpen.IO.TerminationCharacter = TerminationCharacter
    
    If Not IsMissing(Baudrate) Then
        Dim serial As VisaComLib.ISerial
        Set serial = VisaOpen.IO
        With serial
            .Baudrate = Baudrate
            .Parity = Parity
            .DataBits = DataSize
            .StopBits = StopBits
            .EndIn = ASRL_END_TERMCHAR
            .EndOut = ASRL_END_TERMCHAR
           .FlowControl = ASRL_FLOW_RTS_CTS
        End With
    End If

    
End Function


' ************************************************************ Public methods *************************************************
Sub Visa_Snd_Cmd(Instrument As VisaComLib.FormattedIO488, ByVal cmd As String)

    On Error GoTo handler: 'error procedure

    If Instrument.IO.TerminationCharacter <> TerminationCharacter.NoCharacter Then 'with termination character
         Instrument.WriteString cmd 'cmd in ascii
    Else
        cmd = Replace(UCase(cmd), " ", "")
        Call Instrument.IO.Write(StrToBytes(cmd), Len(cmd) / 2) 'cmd in hex values
    End If
    
Exit Sub

handler:
        Call MsgBox("Erro ao enviar comando! Verifique as conexões! Verifique as conexões e endereços!", vbCritical)
        End
End Sub
Function VisaRead(Instrument As VisaComLib.FormattedIO488, Optional offsetParam As Integer = -1) As String
    
   ' On Error GoTo handler: 'error procedure
    If Instrument.IO.TerminationCharacter <> TerminationCharacter.NoCharacter Then 'with termination character
            VisaRead = Instrument.ReadString
    Else
        
        Dim header() As Byte, Checksum() As Byte
        Dim data() As Byte, DataSize As Integer
        
        header = Instrument.IO.Read(4)
        
        'radian: header = 2bytes pack + 2bytes data size
        DataSize = Application.WorksheetFunction.Bitlshift(header(2), 8) + header(3) 'calculates data size
        
        data = Instrument.IO.Read(DataSize)
        
        Checksum = Instrument.IO.Read(2) '16 bits checksum
        
        If CheckSumTest(header, data, Checksum) Then 'if the redundancy check uses another algorithm, it's necessary develops other tests procedures
            If offsetParam > 0 Then VisaRead = ByteArrayToSingle(data, offsetParam) Else: VisaRead = "Offset?" '
        End If
    End If
    
Exit Function

handler:
        Call MsgBox("Error ao executar a leitura!  Verifique as conexões e endereços!", vbCritical)
        End
End Function

Function Visa_query(Instrument As VisaComLib.FormattedIO488, ByVal cmd As String, Optional offsetParam As Integer = -1)
    
    Call Visa_Snd_Cmd(Instrument, cmd)
    
    Visa_query = VisaRead(Instrument, offsetParam)
    
End Function


Private Function StrToNumber(number As String) As Double

    StrToNumber = Replace(number, ".", ",")
    
End Function



Private Function StrToBytes(ByRef x As String) As Byte()
'converts hex values inside a string to a byte array
    
    Dim AscValue(1 To 2) As Integer
    Dim data() As Byte
    Dim cont As Integer, I As Integer, j As Integer
    
    
    cont = Len(x) / 2 - 1
    
    ReDim data(cont)
    
    For I = 0 To cont
        
        For j = 1 To 2
            AscValue(j) = Asc(Mid(x, j + 2 * I, 1))
            AscValue(j) = IIf(AscValue(j) > 64, AscValue(j) - 55, AscValue(j) - 48)
        Next j
        
        data(I) = AscValue(1) * 16 + AscValue(2)
    Next I
    
    StrToBytes = data
End Function




Private Function CheckSumTest(header() As Byte, data() As Byte, Checksum() As Byte) As Boolean 'checkSum verification

'this procedure calculate a new checksum using header and data values and compares with checksum receive from equipment
Dim x As Long
Dim I As Integer

For I = 0 To UBound(header)
    x = x + header(I)
Next I
For I = 0 To UBound(data)
    x = x + data(I)
Next I

Dim y As Long
    y = Application.WorksheetFunction.Bitlshift(Checksum(0), 8) Or Application.WorksheetFunction.Bitlshift(Checksum(1), 0)
    
If x = y Then CheckSumTest = True Else: CheckSumTest = False 'returns true if they are equal and false if otherwise

End Function



Private Function ByteArrayToSingle(data() As Byte, Optional ByVal offset As Integer = -1) As Single

'converts an array float-point with  DSP standard to float-point IEEE standard
'byte array to single-precision floating-point IEEE standard

Dim n As Single     'this is the resulting single number
Dim frac As Double  'fractional binary number from significant
Dim e As Integer    'exponent
Dim sig As Long     'significant integer
Dim I As Integer    'for for loops and whatnot
     
  If offset = -1 Then offset = LBound(data) Else: offset = LBound(data) + offset
  
  'get exponent
  e = 1 * (data(offset) And &HFF&)
  If e = &H80 Or e = &H81 Then
    ByteArrayToSingle = 0
    Exit Function
 End If

  If e > 127 Then e = e - 256
  
  sig = 0
  sig = data(offset + 1) * 2 ^ 16   'shift left 16.
  sig = sig Or (data(offset + 2) * 2 ^ 8) 'shift over 8 and put second byte in sig
  sig = sig Or (data(offset + 3))  'put third byte in sig.

  frac = 1 'start with 1 so result is 1.<fraction>
  If e = -127 Then 'denormalized (0.fraction rather than 1.fraction)
    e = -126
    frac = 0  'start with 0 so result is 0.<fraction>.
  End If

  'calculate binary fraction from significant.
  For I = 22 To 0 Step -1  'step through all 23 bits.
    If (sig And 2 ^ I) > 0 Then frac = frac + 2 ^ -(23 - I) 'if the bit is a 1, the number will be non zero.
  Next

  n = (2 ^ CLng(e)) * frac 'calculate the final number
  If (data(offset + 1) And 128) <> 0 Then n = -n    'if sign bit is set, make negative.
  
  ByteArrayToSingle = n 'return result

End Function
Function FinResources() As Variant
    
    Dim Iomgr As New VisaComLib.ResourceManager
    
    On Error GoTo handler:
        FinResources = Iomgr.FindRsrc("?*INSTR") 'Find resources
    
    Exit Function
handler:
    FinResources = ""

End Function

Function VISA_IDN(ByVal VisaAddress As String) As String
    
    On Error GoTo handler:
    
    Dim Instrument As New VisaComLib.FormattedIO488
    
    Set Instrument = VisaOpen(VisaAddress) 'open communication
    
     Instrument.WriteString "*IDN?"
     VISA_IDN = Instrument.ReadString
     
Exit Function
handler:
VISA_IDN = ""
End Function

