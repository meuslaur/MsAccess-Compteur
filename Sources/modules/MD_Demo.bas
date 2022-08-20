Attribute VB_Name = "MD_Demo"
'@Folder("Demo")
Option Compare Database
Option Explicit

    Private cCompte As C_UtCompteur

Public Function GetInstanceclsLabels() As C_UtCompteur
    If (cCompte Is Nothing) Then Set cCompte = New C_UtCompteur
    Set GetInstanceclsLabels = cCompte
End Function

Public Sub ResetInstanceclsLabels()
    Set cCompte = Nothing
End Sub

Public Sub LanceBoucle(lNbBoucle As Long)
    If (cCompte Is Nothing) Then Set cCompte = New C_UtCompteur

    Dim lCompte As Long
    Dim lRand   As Long
    Dim bRep    As Boolean

    lRand = 80
    bRep = cCompte.SetCompteurs(E_LblRang.Rang1, lNbBoucle)
    If (Not bRep) Then
        If cCompte.xDebugMode Then Stop              'TODO ERR msg erreur compteur
        Exit Sub
    End If

    cCompte.xDebugTimerStart
    DoCmd.Hourglass True

    For lCompte = lNbBoucle To 0 Step -1

        cCompte.UpdateLabels Rang1, "Boucle : " & CStr(lCompte)

        If (lCompte = 112 Or lCompte = 98 Or lCompte = 75 Or lCompte = 38) Then
            bRep = cCompte.SetCompteurs(E_LblRang.Rang2, 45)
            If (Not bRep) Then
                If cCompte.xDebugMode Then Stop      'TODO ERR msg erreur compteur
                Exit For
            End If
            Boucle2 (45)
        End If
        If (Not cCompte.xDebugMode) Then lRand = Int((450 * Rnd) + 1)
        Sleep (lRand)

    Next

    DoCmd.Hourglass False
    cCompte.xDebugTimerEnd

End Sub

Private Sub Boucle2(lNbBoucle As Long)

    Dim lCompte As Long
    Dim lRand   As Long
    Dim bRep    As Boolean

    bRep = cCompte.SetCompteurs(E_LblRang.Rang2, lNbBoucle)
    If (Not bRep) Then
        If cCompte.xDebugMode Then Stop      'TODO ERR msg erreur compteur
        Exit Sub
    End If

    lRand = 30
    For lCompte = lNbBoucle To 0 Step -1

        cCompte.UpdateLabels Rang2

        If (Not cCompte.xDebugMode) Then lRand = Int((200 * Rnd) + 1)
        Sleep (lRand)

    Next
    
'    cCompte.MasqueLabels Rang2

End Sub

