Attribute VB_Name = "modMain"
Option Explicit


Public Sub Main()

    Load frmISBNTOCodeBAR
    frmISBNTOCodeBAR.Show
    
End Sub
Function RetornaCheckISBN(txtTexto As String) As String

    'By Nelson Guajardo N.
    'Email: nguajardo@submarino.com.mx
    'http://www.submarino.com.mx/
    
    On Error GoTo TrataErro
    
    Dim i As Integer, Tamanho As Integer
    Dim SumTemp As Integer, peso As Integer
    
    If IsNull(txtTexto) Then
        RetornaCheckISBN = Null
        Exit Function
    End If
        
    SumTemp = 0
    Tamanho = Len(txtTexto)
    If Tamanho = 0 Then
        RetornaCheckISBN = ""
        Exit Function
    End If
        
    peso = 2
    For i = Len(txtTexto) - 1 To 1 Step -1
        If IsNumeric(Mid(txtTexto, i, 1)) = True Then
            SumTemp = SumTemp + Val(Mid(txtTexto, i, 1)) * peso
            peso = peso + 1
        End If
    Next i
    
    SumTemp = Abs((SumTemp Mod 11) - 11)
    If SumTemp = 11 Then
        SumTemp = 0
    End If
        
    If SumTemp = 10 Then
        RetornaCheckISBN = Mid(txtTexto, 1, Len(txtTexto) - 1) & "X"
    Else
        RetornaCheckISBN = Mid(txtTexto, 1, Len(txtTexto) - 1) & Trim(Str(SumTemp))
    End If
    
Exit Function

TrataErro:
If Err = 94 Then Resume Next
If Err > 0 Then
    MsgBox Err.Description
    Exit Function
End If

End Function

 


Function RetornaCheckSum(txtTexto As String) As String
    
    'By Nelson Guajardo N.
    'Email: nguajardo@submarino.com.mx
    'http://www.submarino.com.mx/
    
    On Error GoTo TrataErro
    
    Dim i As Integer, Tamanho As Integer
    Dim SumTemp As Integer, par As Integer
    
    If IsNull(txtTexto) Then
        RetornaCheckSum = Null
        Exit Function
    End If
        
    SumTemp = 0
    Tamanho = Len(txtTexto)
    If Tamanho = 0 Then
        RetornaCheckSum = ""
        Exit Function
    End If
    
    par = 1
    For i = Len(txtTexto) - 1 To 1 Step -1
        If IsNumeric(Mid(txtTexto, i, 1)) = True Then
            If par = 1 Then
                SumTemp = SumTemp + Val(Mid(txtTexto, i, 1)) * 3
                par = 0
            Else
                SumTemp = SumTemp + Val(Mid(txtTexto, i, 1)) * 1
                par = 1
            End If
        End If
    Next i
    
    SumTemp = Abs((SumTemp Mod 10) - 10)
    If SumTemp = 10 Then
        SumTemp = 0
    End If
    RetornaCheckSum = Mid(txtTexto, 1, Len(txtTexto) - 1) & Trim(Str(SumTemp))
    
Exit Function

TrataErro:
If Err = 94 Then Resume Next
If Err > 0 Then
    MsgBox Err.Description
    Exit Function
End If

End Function


Function RetornaCheckSumDigito(txtTexto As String) As String
    
    'By Nelson Guajardo N.
    'Email: nguajardo@submarino.com.mx
    'http://www.submarino.com.mx/
    
    On Error GoTo TrataErro   '** Changed by: Nelson Guajardo N 05/31/01
    
    Dim i As Integer, Tamanho As Integer
    Dim SumTemp As Integer, par As Integer
    
    If IsNull(txtTexto) Then
        RetornaCheckSumDigito = Null
        Exit Function
    End If
        
    SumTemp = 0
    Tamanho = Len(txtTexto)
    If Tamanho = 0 Then
        RetornaCheckSumDigito = ""
        Exit Function
    End If
    
    par = 1
    For i = Len(txtTexto) To 1 Step -1
        If IsNumeric(Mid(txtTexto, i, 1)) = True Then
            If par = 1 Then
                SumTemp = SumTemp + Val(Mid(txtTexto, i, 1)) * 3
                par = 0
            Else
                SumTemp = SumTemp + Val(Mid(txtTexto, i, 1)) * 1
                par = 1
            End If
        End If
    Next i
    
    SumTemp = Abs((SumTemp Mod 10) - 10)
    If SumTemp = 10 Then
        SumTemp = 0
    End If
    RetornaCheckSumDigito = txtTexto & Trim(Str(SumTemp))
    
Exit Function

TrataErro:
If Err = 94 Then Resume Next
If Err > 0 Then
    MsgBox Err.Description
    Exit Function
End If

End Function



