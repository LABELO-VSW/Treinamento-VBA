Attribute VB_Name = "Módulo1"
Sub Generate_HTMLList()
' Cria uma lista em html
    str2ClipBoard (listValues)
End Sub

Private Function listValues(Optional l As Integer = 1, Optional c As Integer = 1)
    
    If l = 1 Then
        If Cells(l, c) <> "" Then
            listValues = "<ul><li>" & Cells(l, c) & "</li>" & listValues(l + 1, 1)
        Else
            listValues = ""
            Call MsgBox("A lista deve ser iniciada pela célula A1!", vbExclamation, "Erro")
        End If
    Else
        Dim i As Integer
        For i = 1 To c + 1
            If Cells(l, i) <> "" Then
                listValues = repeatString("<ul>", i - c) & repeatString("</ul>", c - i) & _
                            "<li>" & Cells(l, i) & "</li>" & _
                            listValues(l + 1, i)
                Exit Function
            End If
        Next i
        listValues = repeatString("</ul>", c)
    End If

End Function

Private Function repeatString(str As String, n As Integer)
    'Replica um string n vezes
    If n > 0 Then repeatString = str & repeatString(str, n - 1)
End Function

Private Function str2ClipBoard(str As String)
    'Põe uma string dentro do clipboard
    Dim obj As Object
    Set obj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    obj.SetText str
    obj.PutInClipboard
End Function

Sub html_table()
    
    ' Converte área selecionada para uma tabela html
    
    Const className = "table-bordered"
    
    
    Dim str As String
    
    With Selection
        For i = 1 To .Rows.Count
            str = str & "<tr>"
            For j = 1 To .Columns.Count
                If .Cells(i, j).MergeArea.Cells(1, 1).Address = .Cells(i, j).Address Then
                Dim rSpan As String: rSpan = IIf(.Cells(i, j).MergeArea.Cells.Rows.Count > 1, " rowspan='" & .Cells(i, j).MergeArea.Cells.Rows.Count & "'", "")
                Dim cSpan As String: cSpan = IIf(.Cells(i, j).MergeArea.Cells.Columns.Count > 1, " colspan='" & .Cells(i, j).MergeArea.Cells.Columns.Count & "'", "")
                Dim span As String
                If i > 1 Then
                    str = str & "<td" & rSpan & cSpan & ">" & .Cells(i, j) & "</td>"
                Else
                    str = str & "<th" & rSpan & cSpan & ">" & .Cells(i, j) & "</th>"
                End If
                End If
            Next j
            str = str + "</tr>"
        Next i
    End With
    str = "<table class='" & className & "'>" & str & "</table>"
    str2ClipBoard (str)
End Sub

