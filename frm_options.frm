VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_options 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frm_environment 
      Caption         =   "Snowflake Environment Options"
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5175
      Begin MSComctlLib.Slider sld_layers 
         Height          =   510
         Left            =   780
         TabIndex        =   6
         Top             =   240
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   900
         _Version        =   393216
         Min             =   1
         Max             =   100
         SelStart        =   26
         TickFrequency   =   2
         Value           =   26
      End
      Begin MSComctlLib.Slider sld_flakes 
         Height          =   510
         Left            =   780
         TabIndex        =   7
         Top             =   840
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   900
         _Version        =   393216
         Min             =   1
         Max             =   100
         SelStart        =   5
         TickFrequency   =   2
         Value           =   5
      End
      Begin VB.Label lbl_layers 
         Caption         =   "Layers"
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   300
         Width           =   555
      End
      Begin VB.Label lbl_flakes 
         Caption         =   "Flakes"
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   900
         Width           =   555
      End
   End
   Begin VB.Frame frm_movement 
      Caption         =   "Snowflake Movement Options"
      Height          =   1635
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   5175
      Begin MSComctlLib.Slider sld_speed 
         Height          =   510
         Left            =   780
         TabIndex        =   1
         Top             =   420
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   900
         _Version        =   393216
         Min             =   1
         Max             =   100
         SelStart        =   2
         TickFrequency   =   2
         Value           =   2
      End
      Begin MSComctlLib.Slider sld_size 
         Height          =   510
         Left            =   780
         TabIndex        =   3
         Top             =   1020
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   900
         _Version        =   393216
         Min             =   1
         Max             =   100
         SelStart        =   3
         TickFrequency   =   2
         Value           =   3
      End
      Begin VB.Label lbl_size 
         Caption         =   "Size"
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lbl_speed 
         Caption         =   "Speed"
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   480
         Width           =   555
      End
   End
End
Attribute VB_Name = "frm_options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub sld_flakes_Change()
    used_flakes = sld_flakes
    setup_snowflakes
End Sub

Private Sub sld_layers_Change()
    used_layers = sld_layers
    setup_snowflakes
End Sub

Private Sub sld_size_Change()
    max_asize = sld_size
End Sub

Private Sub sld_speed_Change()
    max_speed = sld_speed
End Sub
