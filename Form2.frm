VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H8000000D&
   Caption         =   "Form2"
   ClientHeight    =   6450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17850
   LinkTopic       =   "Form2"
   ScaleHeight     =   6450
   ScaleWidth      =   17850
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15720
      TabIndex        =   20
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14400
      TabIndex        =   18
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      TabIndex        =   16
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Siguiente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   12
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Anterior"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   11
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   10
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   9
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Otro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   960
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4- Presione otro para buscar por codigo"
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   240
      TabIndex        =   22
      Top             =   2880
      Width           =   2790
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   300
      Left            =   15720
      TabIndex        =   21
      Top             =   480
      Width           =   645
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Isv"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   300
      Left            =   14400
      TabIndex        =   19
      Top             =   480
      Width           =   360
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Precio costo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   300
      Left            =   12120
      TabIndex        =   17
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3 - Puede agregar nuevos registros desde aqui"
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   2520
      Width           =   3300
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2 - Puede usar los botones para controlar el registro"
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   2160
      Width           =   3630
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "1- Ingrese el codigo para buscar luego presione buscar en el boton de abajo"
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   1800
      Width           =   5535
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Precio compra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   300
      Left            =   9960
      TabIndex        =   7
      Top             =   480
      Width           =   1725
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   300
      Left            =   7680
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   300
      Left            =   2640
      TabIndex        =   5
      Top             =   480
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   300
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnnconexion As ADODB.Connection
Dim registron As ADODB.Recordset
Dim cmdcomando As ADODB.Command

Private Sub Command1_Click()
'este es un truco muy bonito te pregunta que quiere buscar con inputbox
Dim respuesta As String
'se declaro la variable respuesta y luego pregunto
respuesta = InputBox("Ingrese el codigo a buscar", "Sistema")
'codigo de conexion normal
Set cnnconexion = New ADODB.Connection
cnnconexion.ConnectionString = "Provider=SQLOLEDB.1;Server=DESKTOP-8CDH8HU;Uid=sa;pwd=CRISTO777;Database=Inventario;"
cnnconexion.CursorLocation = adUseClient
cnnconexion.ConnectionTimeout = 15
cnnconexion.Open
Set cmdcomando = New ADODB.Command
With cmdcomando
.ActiveConnection = cnnconexion
.CommandType = adCmdText
.CommandTimeout = 15
'esta es la consulta sql segun el codigo introducido en la variable respuesta
.CommandText = "Select * From Registro Where Codigo = " & Val(respuesta)
End With
Set registron = cmdcomando.Execute()
Call suscampos
'deshabilito los botones siguientes no estan programados todavia
Command4.Enabled = True
Command5.Enabled = True
End Sub

Public Sub cargardatos()
Set cnnconexion = New ADODB.Connection
cnnconexion.ConnectionString = "Provider=SQLOLEDB.1;Server=DESKTOP-8CDH8HU;Uid=sa;pwd=CRISTO777;Database=Inventario;"
cnnconexion.CursorLocation = adUseClient
cnnconexion.ConnectionTimeout = 15
cnnconexion.Open
Set cmdcomando = New ADODB.Command
With cmdcomando
.ActiveConnection = cnnconexion
.CommandType = adCmdText
.CommandTimeout = 15
.CommandText = "Select * From Registro"
End With
Set registron = cmdcomando.Execute()
Call suscampos

End Sub

Private Sub suscampos()
Text1.Text = registron.Fields("Codigo")
Text2.Text = registron.Fields("Nombre")
Text3.Text = registron.Fields("Cantidad")
Text4.Text = registron.Fields("Preciocompra")
Text5.Text = registron.Fields("Preciocosto")
Text6.Text = registron.Fields("Isv")
Text7.Text = registron.Fields("Valor")
End Sub

Private Sub Command2_Click()
If Len(Text1.Text) >= 1 And Len(Text2.Text) >= 1 And Len(Text3.Text) >= 1 And Len(Text4.Text) >= 1 And Len(Text5.Text) >= 1 And Len(Text6.Text) >= 1 And Len(Text7.Text) >= 1 Then
'se ejecuta la consulta para guardar
Set cnnconexion = New ADODB.Connection
cnnconexion.ConnectionString = "Provider=SQLOLEDB.1;Server=DESKTOP-8CDH8HU;Uid=sa;Pwd=CRISTO777;Database=Inventario;"
cnnconexion.CursorLocation = adUseClient
cnnconexion.ConnectionTimeout = 15
cnnconexion.Open
Set cmdcomando = New ADODB.Command
With cmdcomando
.ActiveConnection = cnnconexion
.CommandType = adCmdText
.CommandTimeout = 15
'esta es la instruccion sql que guarda
.CommandText = "Insert Into Registro (Codigo,Nombre,Cantidad,Preciocosto,Preciocompra,Isv,Valor) Values ('" + Text1.Text + "','" + Text2.Text + "','" + Text3.Text + "','" + Text4.Text + "','" + Text5.Text + "','" + Text6.Text + "','" + Text7.Text + "')"
End With
'si no hago esto no produce los cambios ni el guardado
Set registron = cmdcomando.Execute()
Else
MsgBox "No todos los campos estan rellenos", vbOKOnly, "Sistema"
'cierro el comando sql
End If

End Sub

Private Sub Command5_Click()
If registron.EOF = True Xor registron.BOF = True Then
MsgBox "CRISTO VIVE", vbOKOnly, "Sistema"
Else
registron.MoveNext
'Call suscampos
Text1.Text = registron.Fields("Codigo")
Text2.Text = registron.Fields("Nombre")
Text3.Text = registron.Fields("Cantidad")
Text4.Text = registron.Fields("Precicosto")
End If
End Sub

Private Sub Form_Load()
Command4.Enabled = False
Command5.Enabled = False
End Sub
