Attribute VB_Name = "mod_globals"
'**************
'SNOWFLAKE DATA
    Global Const layers = 100
    Global Const flakes = 100
    Global Const elements = 3
        Global Const x = 1
        Global Const y = 2
    Global snowflakes(layers, flakes, elements) As Long
    Global max_speed
    Global max_asize
    Global used_layers
    Global used_flakes
'**************
'ENVIRONMENTAL CONSTANTS
    Global Const ceiling = 0
    Global floor As Long
    Global layer As Integer
    Global flake As Integer
'**************

Public Sub setup_snowflakes()
    For layer = 0 To used_layers
        For flake = 0 To used_flakes
            snowflakes(layer, flake, x) = Int((frm_snowfall.pic_snowfall.Width * Rnd) + 1)
            snowflakes(layer, flake, y) = Int((floor * Rnd) + 1)
        Next flake
    Next layer
End Sub

