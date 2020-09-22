VERSION 5.00
Begin VB.Form frmISBNTOCodeBAR 
   Caption         =   "ISBN To Code BaR (By Nelson Guajardo N)"
   ClientHeight    =   2805
   ClientLeft      =   9300
   ClientTop       =   5745
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   2805
   ScaleWidth      =   5415
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   735
      Left            =   1920
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Txt_ISBN 
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Text            =   "0000190284"
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox Txt_Cod_Barra 
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Text            =   "Txt_Cod_Barra"
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CodeBar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Index           =   1
      Left            =   3000
      TabIndex        =   4
      Top             =   240
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ISBN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   240
      Width           =   645
   End
End
Attribute VB_Name = "frmISBNTOCodeBAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''By Nelson Guajardo N.
    'Email: nguajardo@submarino.com.mx
    'http://www.submarino.com.mx/
    
    Private Sub cmdOK_Click()

    Verifica_ISBN

End Sub


Private Function Verifica_ISBN() As Boolean
 
On Error Resume Next

    Dim Texto As String

    If RetornaCheckISBN(Trim(Txt_ISBN)) <> Trim(Txt_ISBN) Then
        If RetornaCheckSum(Trim(Txt_ISBN)) <> Trim(Txt_ISBN) Then
            If MsgBox("ISBN o EAN Inválido Quieres Hacerlo?", vbYesNo + vbDefaultButton2) = vbNo Then
                'Codigo
                Verifica_ISBN = False
                Exit Function
            End If
        End If
  
    Else
        Texto = Trim(Str(978) & Mid(Txt_ISBN, 1, 9))
        Txt_Cod_Barra = RetornaCheckSumDigito(Texto)
    End If

End Function

'By Nelson Guajardo N.
    'Email: nguajardo@submarino.com.mx
    'http://www.submarino.com.mx/
    
Private Sub Form_Load()


End Sub
