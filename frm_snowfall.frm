VERSION 5.00
Begin VB.Form frm_snowfall 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Snowfall"
   ClientHeight    =   4485
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5730
   Icon            =   "frm_snowfall.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic_snowfall 
      AutoRedraw      =   -1  'True
      Height          =   4335
      Left            =   60
      Picture         =   "frm_snowfall.frx":0442
      ScaleHeight     =   4275
      ScaleWidth      =   5535
      TabIndex        =   0
      Top             =   60
      Width           =   5595
      Begin VB.Timer tim_refresh 
         Interval        =   1
         Left            =   60
         Top             =   120
      End
   End
   Begin VB.Menu mnu_file 
      Caption         =   "File"
      Begin VB.Menu mnu_start 
         Caption         =   "Start"
      End
      Begin VB.Menu mnu_stop 
         Caption         =   "Stop"
      End
      Begin VB.Menu mnu_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu men_options 
      Caption         =   "Options"
   End
End
Attribute VB_Name = "frm_snowfall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    used_layers = 26
    used_flakes = 5
    max_speed = 2
    max_asize = 3
    
    mnu_stop_Click
    pic_snowfall.BackColor = vbBlack
    size_form
    setup_snowflakes
    move_snowflakes
    draw_snowflakes
    mnu_start_Click
End Sub

Private Sub size_form()
    setup_snowflakes
    If frm_snowfall.WindowState <> vbMinimized Then
        If frm_snowfall.Height > 835 And frm_snowfall.Width > 250 Then
            pic_snowfall.Left = 60
            pic_snowfall.Top = 60
            pic_snowfall.Width = frm_snowfall.Width - 250
            pic_snowfall.Height = frm_snowfall.Height - 835
        End If
    End If
    floor = pic_snowfall.Height
    pic_snowfall.FillStyle = vbFSSolid
End Sub

Private Sub Form_Resize()
    size_form
End Sub

Private Sub setup_snowflake(flake_layer As Integer, fin_flake As Integer)
    snowflakes(flake_layer, fin_flake, x) = Int((pic_snowfall.Width * Rnd) + 1)
End Sub

Private Sub move_snowflakes()
    For layer = 0 To used_layers
        For flake = 0 To used_flakes
            Select Case snowflakes(layer, flake, y)
                Case Is >= floor
                    snowflakes(layer, flake, y) = ceiling
                    setup_snowflake layer, flake
                Case Else
                    snowflakes(layer, flake, y) = snowflakes(layer, flake, y) + ((used_layers - layer) + 1) * max_speed
            End Select
        Next flake
    Next layer
End Sub

Private Sub draw_snowflakes()
    pic_snowfall.Cls
    For layer = used_layers To 0 Step -1
        flake_colour = (used_layers - layer) * 25
        pic_snowfall.FillColor = RGB(flake_colour, flake_colour, flake_colour)
        For flake = 0 To used_flakes
            pic_snowfall.Circle (snowflakes(layer, flake, x), snowflakes(layer, flake, y)), (used_layers - layer) * max_asize, RGB(flake_colour, flake_colour, flake_colour)
        Next flake
    Next layer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frm_options
End Sub

Private Sub men_options_Click()
    frm_options.Show
End Sub

Private Sub mnu_exit_Click()
    Unload Me
End Sub

Private Sub mnu_options_Click()
    frm_options.Show
End Sub

Private Sub mnu_start_Click()
    mnu_start.Enabled = False
    With mnu_stop
        .Enabled = True
        tim_refresh.Enabled = .Enabled
    End With
End Sub

Private Sub mnu_stop_Click()
    mnu_start.Enabled = True
    With mnu_stop
        .Enabled = False
        tim_refresh.Enabled = .Enabled
    End With
End Sub

Private Sub tim_refresh_Timer()
    move_snowflakes
    draw_snowflakes
End Sub
