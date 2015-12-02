VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmbusUsuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda de Usuarios"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   11955
   Icon            =   "frmbusUsuarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4905
   ScaleWidth      =   11955
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   7680
      OleObjectBlob   =   "frmbusUsuarios.frx":058A
      Top             =   240
   End
   Begin VB.CommandButton cmdbuscar 
      Caption         =   "&Buscar"
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   10320
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtbuscar 
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshusuario 
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   7435
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "frmbusUsuarios.frx":07BE
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmbusUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsUsuario As New ADODB.Recordset
Dim strUsuario As String

Private Sub cmdbuscar_Click()
If Me.txtbuscar <> "" Then
strUsuario = " select *  from Usuarios where u_apellido like '" & Me.txtbuscar.Text & "%'"
rsUsuario.Open strUsuario, Cn, adOpenDynamic, adCmdTable

Set Me.mshusuario.DataSource = rsUsuario
Me.mshusuario.Refresh
rsUsuario.Close
Set rsUsuario = Nothing
TitulosColumnas
End If
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()

strUsuario = " SELECT Usuarios.u_cod_usuario, Usuarios.u_apellido, Usuarios.u_nombre, " & _
             " Usuarios.u_fecha_nacimiento, Usuarios.u_domicilio, Usuarios.u_caracteristica_tel, " & _
             " Usuarios.u_telefono, Usuarios.u_usuario, Usuarios.u_contraseña, Usuarios.u_permiso " & _
             " FROM Usuarios "

               
rsUsuario.Open strUsuario, Cn, adOpenDynamic, adCmdTable

Set Me.mshusuario.DataSource = rsUsuario

TitulosColumnas

rsUsuario.Close
Set rsUsuario = Nothing

Skin1.LoadSkin App.Path & "\Skins\Dogmas2.skn"
Skin1.ApplySkin frmbusUsuarios.hwnd

End Sub

Sub TitulosColumnas()

    With Me.mshusuario

        .ColWidth(0) = 0
        
        .ColWidth(1) = 2500
        .TextMatrix(0, 1) = "Apellido"
        
        .ColWidth(2) = 2500
        .TextMatrix(0, 2) = "Nombre"
        
        .ColWidth(3) = 1000
        .TextMatrix(0, 3) = "Fecha"
                 
        .ColWidth(4) = 3500
        .TextMatrix(0, 4) = "Domicilio"
        
        .ColWidth(5) = 1200
        .TextMatrix(0, 5) = "Caract. Tel."
        
        .ColWidth(6) = 1000
        .TextMatrix(0, 6) = "Telefono"
        
        .ColWidth(7) = 0
        .TextMatrix(0, 7) = "Usuario"
        
        .ColWidth(8) = 0
        .TextMatrix(0, 8) = "Contraseña"
        
        .ColWidth(9) = 0
        .TextMatrix(0, 9) = "Permiso"
        
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

    UnHookForm Me.mshusuario
    UnHookForm Me.mshusuario

End Sub

Private Sub mshusuario_Click()
Me.mshusuario.SelectionMode = flexSelectionByRow
End Sub

Private Sub mshusuario_DblClick()

strUsuario = "select * from Usuarios"
rsUsuario.Open strUsuario, Cn, adOpenDynamic, adCmdTable

frmcargaUsuarios.txtcod = Me.mshusuario.TextMatrix(Me.mshusuario.Row, 0)
frmcargaUsuarios.txtapellido = Me.mshusuario.TextMatrix(Me.mshusuario.Row, 1)
frmcargaUsuarios.txtnombre = Me.mshusuario.TextMatrix(Me.mshusuario.Row, 2)
frmcargaUsuarios.mskfecha = Me.mshusuario.TextMatrix(Me.mshusuario.Row, 3)
frmcargaUsuarios.txtdomicilio = Me.mshusuario.TextMatrix(Me.mshusuario.Row, 4)
frmcargaUsuarios.txtcaracteristicatel = Me.mshusuario.TextMatrix(Me.mshusuario.Row, 5)
frmcargaUsuarios.txttelefono = Me.mshusuario.TextMatrix(Me.mshusuario.Row, 6)
frmcargaUsuarios.txtusuario = Me.mshusuario.TextMatrix(Me.mshusuario.Row, 7)
frmcargaUsuarios.txtcontraseña = Me.mshusuario.TextMatrix(Me.mshusuario.Row, 8)
frmcargaUsuarios.cbopermiso = Me.mshusuario.TextMatrix(Me.mshusuario.Row, 9)

Unload Me
rsUsuario.Close
Set rsUsuario = Nothing

frmcargaUsuarios.Avilitar
frmcargaUsuarios.cmdmodificar.Enabled = True
frmcargaUsuarios.cmdborrar.Enabled = True
frmcargaUsuarios.cmdlimpiar.Enabled = True

End Sub

Private Sub mshusuario_GotFocus()

    HookForm Me.mshusuario

End Sub

Private Sub mshusuario_LostFocus()

    UnHookForm Me.mshusuario

End Sub

