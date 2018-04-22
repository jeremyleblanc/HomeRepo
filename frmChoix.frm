VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChoix 
   Caption         =   "Choix d'action"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7215
   OleObjectBlob   =   "frmChoix.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmChoix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnOk_Click()

'Ouvre le formulaire demandé et ferme le formulaire frmChoix
    Unload Me

    If cbxChoix.Value = "" Then
        frmAjoutCadenas.Show
    ElseIf cbxChoix.Value = "" Then
        frmEnDev.Show
    ElseIf cbxChoix.Value = "" Then
        frmMontagePoste.Show
    ElseIf cbxChoix.Value = "" Then
        frmEnDev.Show
    End If
    
End Sub

Private Sub btnQuitter_Click()

'ferme à partir du bouton quitter
Call UserForm_QueryClose(0, 0)

End Sub


Private Sub cbxChoix_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

'Accepte le choix avec la touche Enter
    If KeyCode = 13 Then
        Call btnOk_Click
    End If
    
'permet d'utiliser la flèche du bas pour faire une rotation de choix dans le menu
    If KeyCode = 40 And cbxChoix.ListIndex = (cbxChoix.ListCount - 1) Then
        cbxChoix.ListIndex = -1
    End If
    
End Sub

Private Sub UserForm_Initialize()

'Initialisation de la liste de choix offert
    With cbxChoix
        .AddItem "Embauche d'employé"
        .AddItem "D'épart d'employé"
        .AddItem "Modification d'employé"
        .AddItem ""
        .Value = ""
        .SetFocus
    End With
    
End Sub



'Début du code pour la fermeture d'un formulaire!!___________________________________________________________________________________________________________________



Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

Dim Response As Integer
Dim msg As String: Smsg = "Voulez-vouz vraiment quitter?"


If CloseMode <> 1 Then ' Si la fermeture n'est demandé pas par le code faire les actions suivantes.

    If CloseMode = 0 Then 'si la demande vient du formulaire, ouvre une msgbox.
        Response = MsgBox(Smsg, vbYesNo + vbDefaultButton2 + vbQuestion, "Quitter?")
    End If

    
    If Response = vbYes And CloseMode = 0 Then 'Si la réponse du messagebox est oui et que le mode de fermeture est celui du formulaire, on ferme le formulaire
        'Cancel = False
        On Error Resume Next
        Unload Me
        On Error GoTo 0
    End

    Else
    'sinon on ne ferme pas le formulaire.
    Cancel = True
    End If
    
    Else
    'sinon le code demande la fermeture,alors on ferme
    Unload Me
End If

' Affiche un message si l'utilisateur à appuyer sur le X du formulaire.

End Sub



'Fin du code pour la fermeture d'un formulaire!!___________________________________________________________________________________________________________________
