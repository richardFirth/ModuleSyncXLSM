Attribute VB_Name = "ZZZ_Color_1"
'$VERSIONCONTROL
'$*MINOR_VERSION*1.1
'$*DATE*23Jan18
'$*ID*Color

Option Explicit

'Red     #e6194b (230, 25, 75)   (0, 100, 66, 0)
'Green   #3cb44b (60, 180, 75)   (75, 0, 100, 0)
'Yellow  #ffe119 (255, 225, 25)  (0, 25, 95, 0)
'Blue    #0082c8 (0, 130, 200)   (100, 35, 0, 0)
'Orange  #f58231 (245, 130, 48)  (0, 60, 92, 0)
'Purple  #911eb4 (145, 30, 180)  (35, 70, 0, 0)
'Cyan    #46f0f0 (70, 240, 240)  (70, 0, 0, 0)
'Magenta #f032e6 (240, 50, 230)  (0, 100, 0, 0)
'Lime    #d2f53c (210, 245, 60)  (35, 0, 100, 0)
'Pink    #fabebe (250, 190, 190) (0, 30, 15, 0)
'Teal    #008080 (0, 128, 128)   (100, 0, 0, 50)
'Lavender    #e6beff (230, 190, 255) (10, 25, 0, 0)
'Brown   #aa6e28 (170, 110, 40)  (0, 35, 75, 33)
'Beige   #fffac8 (255, 250, 200) (5, 10, 30, 0)
'Maroon  #800000 (128, 0, 0) (0, 100, 100, 50)
'Mint    #aaffc3 (170, 255, 195) (33, 0, 23, 0)
'Olive   #808000 (128, 128, 0)   (0, 0, 100, 50)
'Coral   #ffd8b1 (255, 215, 180) (0, 15, 30, 0)
'Navy    #000080 (0, 0, 128) (100, 100, 0, 50)
'Grey    #808080 (128, 128, 128) (0, 0, 0, 50)
'White   #FFFFFF (255, 255, 255) (0, 0, 0, 0)
'Black   #000000 (0, 0, 0)   (0, 0, 0, 100)


Public Enum RFcolor
    A_Red
    B_Green
    C_Yellow
    D_Blue
    E_Orange
    F_Purple
    G_Cyan
    H_Magenta
    I_Lime
    J_Pink
    K_Teal
    L_Lavender
    M_Brown
    N_Beige
    O_Maroon
    P_Mint
    Q_Olive
    R_Coral
    S_Navy
    T_Grey
    U_White
    V_Black
    W_Nocolor
End Enum




 ' /-----------------------------------------\
 ' |interface for getting common colors      |
 ' \-----------------------------------------/
Public Function getRFColor(theC As RFcolor) As Long
    
    If theC = A_Red Then getRFColor = RGB(230, 25, 75): Exit Function
    If theC = B_Green Then getRFColor = RGB(60, 180, 75): Exit Function
    If theC = C_Yellow Then getRFColor = RGB(255, 225, 25): Exit Function
    If theC = D_Blue Then getRFColor = RGB(0, 130, 200): Exit Function
    If theC = E_Orange Then getRFColor = RGB(245, 130, 48): Exit Function
    If theC = F_Purple Then getRFColor = RGB(145, 30, 180): Exit Function
    If theC = G_Cyan Then getRFColor = RGB(70, 240, 240): Exit Function
    If theC = H_Magenta Then getRFColor = RGB(240, 50, 230): Exit Function
    If theC = I_Lime Then getRFColor = RGB(210, 245, 60): Exit Function
    If theC = J_Pink Then getRFColor = RGB(250, 190, 190): Exit Function
    If theC = K_Teal Then getRFColor = RGB(0, 128, 128): Exit Function
    If theC = L_Lavender Then getRFColor = RGB(230, 190, 255): Exit Function
    If theC = M_Brown Then getRFColor = RGB(170, 110, 40): Exit Function
    If theC = N_Beige Then getRFColor = RGB(255, 250, 200): Exit Function
    If theC = O_Maroon Then getRFColor = RGB(128, 0, 0): Exit Function
    If theC = P_Mint Then getRFColor = RGB(170, 255, 195): Exit Function
    If theC = Q_Olive Then getRFColor = RGB(128, 128, 0): Exit Function
    If theC = R_Coral Then getRFColor = RGB(255, 215, 180): Exit Function
    If theC = S_Navy Then getRFColor = RGB(0, 0, 128): Exit Function
    If theC = T_Grey Then getRFColor = RGB(128, 128, 128): Exit Function
    If theC = U_White Then getRFColor = RGB(255, 255, 255): Exit Function
    If theC = V_Black Then getRFColor = RGB(0, 0, 0): Exit Function
    If theC = W_Nocolor Then getRFColor = RGB(255, 255, 255): Exit Function
End Function


Public Function getColorTF(val As Boolean) As Long
    If val Then getColorTF = getRFColor(B_Green): Exit Function
    getColorTF = getRFColor(A_Red)
End Function








