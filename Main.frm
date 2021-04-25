VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Praktik Jaringan; SMTP Client"
   ClientHeight    =   4905
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   7185
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Kondisi :"
      Height          =   615
      Left            =   960
      TabIndex        =   15
      Top             =   3480
      Width           =   5175
      Begin VB.Label StatusTxt 
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.TextBox txtServidor 
      Height          =   285
      Left            =   3600
      TabIndex        =   13
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox txtSuNombre 
      Height          =   285
      Left            =   3600
      TabIndex        =   11
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox txtNombreRemitente 
      Height          =   285
      Left            =   3600
      TabIndex        =   9
      Top             =   240
      Width           =   3375
   End
   Begin VB.CommandButton Cmd_salir 
      Caption         =   "&Keluar"
      Height          =   495
      Left            =   4440
      TabIndex        =   8
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox txtCuerpo 
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1920
      Width           =   6855
   End
   Begin VB.TextBox txtAsunto 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   3255
   End
   Begin VB.TextBox txtEmailDestino 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox txtEmailRemitente 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton Cmd_enviar 
      Caption         =   "&Kirim Email"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   4200
      Width           =   2175
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3480
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label6 
      Caption         =   "SMTP Server:"
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Label Label5 
      Caption         =   "Nama :"
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label4 
      Caption         =   "Nama:"
      Height          =   255
      Left            =   3600
      TabIndex        =   10
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "Subjek:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Kepada: (Email)"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Pengirim: (Email)"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Response As String, Reply As Integer, DateNow As String
Dim mail_from As String, rcpt_to As String, fecha_envio As String
Dim nombre_remitente As String, nombre_destinatario As String, asunto As String
Dim cuerpo_mensaje As String, mensaje_total As String
Dim Start As Single, Tmr As Single

Sub SendEmail(MailServerName As String, FromName As String, FromEmailAddress As String, ToName As String, ToEmailAddress As String, EmailSubject As String, EmailBodyOfMessage As String)
          
    Winsock1.LocalPort = 0
    
If Winsock1.State = sckClosed Then ' Comprueba si el socket está cerrado
    DateNow = Format(Date, "Ddd") & ", " & Format(Date, "dd Mmm YYYY") & " " & Format(Time, "hh:mm:ss") & "" & " -0600"
    mail_from = "mail from:" + Chr(32) + FromEmailAddress + vbCrLf ' E-mail del remitente
    rcpt_to = "rcpt to:" + Chr(32) + ToEmailAddress + vbCrLf ' E-mail destinatario
    fecha_envio = "Date:" + Chr(32) + DateNow + vbCrLf ' Fecha del envío (la coge del sistema
    nombre_remitente = "From:" + Chr(32) + FromName + vbCrLf ' Nombre del remitente
    nombre_destinatario = "To:" + Chr(32) + ToName + vbCrLf ' Nombre del destinatario
    asunto = "Subject:" + Chr(32) + EmailSubject + vbCrLf ' Asunto del mensaje
    cuerpo_mensaje = EmailBodyOfMessage + vbCrLf ' Cuerpo del mensaje
    mensaje_total = nombre_remitente + fecha_envio + nombre_destinatario + asunto  ' Combinamos las variables necesarias para enviar el mensaje

    Winsock1.Protocol = sckTCPProtocol ' Establecemos el protocolo para el envío
    Winsock1.RemoteHost = MailServerName ' Establecemos la dirección del servidor
    Winsock1.RemotePort = 25 ' Puerto SMTP
    Winsock1.Connect ' Comienza la conexión
    
    WaitFor ("220") ' Indica que la conexión está preparada
    
    StatusTxt.Caption = "Conectando...."
    StatusTxt.Refresh
           
    Winsock1.SendData ("HELO prueba de javi y ramon.com" + vbCrLf)

    WaitFor ("250") ' Acepta la conexión

    StatusTxt.Caption = "Conectado"
    StatusTxt.Refresh

    Winsock1.SendData (mail_from)

    StatusTxt.Caption = "Enviando mensaje"
    StatusTxt.Refresh

    WaitFor ("250")

    Winsock1.SendData (rcpt_to)

    WaitFor ("250")

    Winsock1.SendData ("data" + vbCrLf)
    
    WaitFor ("354") 'Indica que está listo para recibir los datos del mensaje

    Winsock1.SendData (mensaje_total + vbCrLf)
    Winsock1.SendData (cuerpo_mensaje + vbCrLf)
    Winsock1.SendData ("." + vbCrLf) 'Indica que acaba el cuerpo del mensaje

    WaitFor ("250")

    Winsock1.SendData ("quit" + vbCrLf) 'Se cierra la conexión con el servidor
    
    StatusTxt.Caption = "Desconectando"
    StatusTxt.Refresh

    WaitFor ("221") 'Servidor indica final de transmisión

    Winsock1.Close 'Cerramos el socket
Else
    MsgBox (Str(Winsock1.State))
End If
   
End Sub
Sub WaitFor(ResponseCode As String)
    Start = Timer ' contador de tiempo
    While Len(Response) = 0  ' Mientras el servidor no de señales de vida
        Tmr = Timer - Start
        DoEvents ' Hace que el sistema siga comprobando si llega respuesta
        If Tmr > 10 Then ' Tiempo de espera en segundos
            MsgBox "Error SMTP, time out mientras se esperaba respuesta por parte de " + txtServidor.Text, 64, MsgTitle
             End
        End If
    Wend
    While Left(Response, 3) <> ResponseCode   'comprueba el código de respuesta apropiado en cada momento
        DoEvents
        If Tmr > 10 Then
            MsgBox "Error SMTP, código de respuesta incorrecto. Código correcto debería ser: " + ResponseCode + " Código recibido: " + Response, 64, MsgTitle
            Exit Sub
        End If
    Wend
Response = "" ' Inicializa la variable a blancos
End Sub


Private Sub Cmd_enviar_Click()
    StatusTxt.Caption = ""
    SendEmail txtServidor.Text, txtNombreRemitente.Text, txtEmailRemitente.Text, txtSuNombre.Text, txtEmailDestino.Text, txtAsunto.Text, txtCuerpo.Text
    StatusTxt.Caption = "Mensaje enviado"
    StatusTxt.Refresh
    Beep
    
    Close
    txtServidor.Text = ""
    txtSuNombre.Text = ""
    txtEmailDestino.Text = ""
    txtAsunto.Text = ""
    txtCuerpo.Text = ""
End Sub

Private Sub Cmd_salir_Click()
    
    End
    
End Sub

'Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

   ' Winsock1.GetData Response ' Comprueba los datos recibidos

'End Sub
