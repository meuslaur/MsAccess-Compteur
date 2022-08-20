VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_UtCompteur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Dev")
' DateMod: 18/08/2022 - 19:17
Option Compare Database
Option Explicit

    Private Const COUL_ORANGE   As Long = 3243501
    Private Const COUL_VERT     As Long = 3506772
    Private Const CODE_ASC      As Long = 116

    Private Type T_Labels
        lblU1  As Access.Label  '// Label Unités
        lblD1  As Access.Label  '// Label Dizaines
        lblC1  As Access.Label  '// Label Centaines
        Cpt1(2) As Long         '// Compteurs en cours pour U,D,C
        Compteur1 As Long       '// Compteur en cours.
        Rang1Isup As Boolean    '// Labels rang1 actif on/off.
        lblU2  As Access.Label
        lblD2  As Access.Label
        lblC2  As Access.Label
        Cpt2(2) As Long
        Compteur2 As Long
        Rang2Isup As Boolean
        lblTxt As Access.Label  '// Label texte.
        TxtIsUp As Boolean      '// Label text actif on/off.
    End Type
    Private thisLbl As T_Labels

    Public Enum E_LblRang
        Rang1 = 1
        Rang2 = 2
    End Enum

    Private mCoulDef  As Boolean

'//----------- Debug Mode ---------
    Private mDebugOn            As Boolean
    Private mDebugLblOn         As Boolean
    Private mTimerStartLblOn    As Single
    Private mTimerEndLblOn      As Single
    Private mTimerStartLblOff   As Single
    Private mTimerEndLblOff     As Single
'//--------------------------------

Private Sub Class_Initialize()
    mDebugLblOn = True
End Sub

Private Sub Class_Terminate()
    Set thisLbl.lblU1 = Nothing
    Set thisLbl.lblD1 = Nothing
    Set thisLbl.lblC1 = Nothing
    Set thisLbl.lblU2 = Nothing
    Set thisLbl.lblD2 = Nothing
    Set thisLbl.lblC2 = Nothing
    Set thisLbl.lblTxt = Nothing
End Sub

'----------------------------------------------------------------
' Procedure Nom   : SetLabels
' ----------------------------------------------------------------
' Sujet           : Iniitialise les labels du form.
'=== Paramètres ===
' RangLabel (E_LblRang): Indique quel labels utiliser.
' NewLblU (Label):
' NewLblD (Label):
' NewLblC (Label):
'==================
'
' Return Boolean : True si pas de problème.
' Author  : Laurent
' Date    : 19/08/2022 - 14:17
' ----------------------------------------------------------------
Public Function SetLabels(RangLabel As E_LblRang, NewLblU As Access.Label, NewLblD As Access.Label, NewLblC As Access.Label) As Boolean

    Dim bErr As Boolean

    bErr = ((NewLblU Is Nothing) Or (NewLblD Is Nothing) Or (NewLblC Is Nothing))

    Select Case RangLabel
        Case E_LblRang.Rang1
            thisLbl.Rang1Isup = Not bErr
            If (bErr And mDebugOn) Then Stop      'TODO ERR msg si a false
            Set thisLbl.lblU1 = NewLblU
            Set thisLbl.lblD1 = NewLblD
            Set thisLbl.lblC1 = NewLblC
        Case E_LblRang.Rang2
            thisLbl.Rang2Isup = Not bErr
            If (bErr And mDebugOn) Then Stop      'TODO ERR msg si a false
            Set thisLbl.lblU2 = NewLblU
            Set thisLbl.lblD2 = NewLblD
            Set thisLbl.lblC2 = NewLblC
    End Select

    SetLabels = Not bErr

End Function

Public Sub SetLabelTexte(NewLblTxt As Access.Label)
    Set thisLbl.lblTxt = NewLblTxt
    thisLbl.TxtIsUp = False
    If (Not NewLblTxt Is Nothing) Then
        thisLbl.TxtIsUp = True
        thisLbl.lblTxt.Visible = True
    End If
End Sub

'----------------------------------------------------------------
'// Défini le compteur de départ pour chaque rang de labels.
' ----------------------------------------------------------------
Public Function SetCompteurs(RangLabel As E_LblRang, NewCompteur As Long) As Boolean

    If ((NewCompteur < 5) Or (NewCompteur > 999)) Then
        If mDebugOn Then Stop
        Exit Function               'TODO ERR msg compteur non valide
    End If

    Select Case RangLabel
        Case E_LblRang.Rang1
            thisLbl.Compteur1 = NewCompteur
            InitialiseLabels Rang1
        Case E_LblRang.Rang2
            thisLbl.Compteur2 = NewCompteur
            InitialiseLabels Rang2
    End Select

    SetCompteurs = True

End Function

Public Sub UpdateLabels(LabelRang As E_LblRang, Optional NewTexte As String = vbNullString)
    Select Case LabelRang
        Case E_LblRang.Rang1
            UpdateLabels1 NewTexte
        Case E_LblRang.Rang2
            UpdateLabels2
   End Select
End Sub

Public Sub MasqueLabels(LabelRang As E_LblRang)
    Select Case LabelRang
        Case E_LblRang.Rang1
            If (Not thisLbl.Rang1Isup) Then
                If mDebugOn Then Stop       'TODO ERR msg labels rang1 non initialisés.
                Exit Sub
            End If
            thisLbl.lblU1.Visible = False
            thisLbl.lblD1.Visible = False
            thisLbl.lblC1.Visible = False
            thisLbl.Rang1Isup = False
            If thisLbl.TxtIsUp Then thisLbl.lblTxt.Visible = False
            thisLbl.TxtIsUp = False
            thisLbl.Compteur1 = -1
        Case E_LblRang.Rang2
            If (Not thisLbl.Rang2Isup) Then
                If mDebugOn Then Stop       'TODO ERR msg labels rang2 non initialisés.
                Exit Sub
            End If
            thisLbl.lblU2.Visible = False
            thisLbl.lblD2.Visible = False
            thisLbl.lblC2.Visible = False
            thisLbl.Compteur2 = -1
    End Select
End Sub
' ----------------------------------------------------------------
'// Mise a jout des labels rang 1, si Compteur non défini ou si mode debug sans label, sort.
' ----------------------------------------------------------------
Private Sub UpdateLabels1(Optional NewTexte As String = vbNullString)

    If (Not mDebugLblOn) Then Exit Sub  '// Mode debug n'affiche pas les labels.

    If (Not thisLbl.Rang1Isup) Then
        If mDebugOn Then Stop           'TODO ERR msg Compteur non initialisé
        Exit Sub
    End If

    If (thisLbl.Compteur1 < 99) And (Not mCoulDef) Then DefCoul (COUL_VERT)

    If ((Len(NewTexte) > 0) And thisLbl.TxtIsUp) Then thisLbl.lblTxt.Caption = NewTexte

    thisLbl.lblU1.Caption = ChrW$(CODE_ASC + thisLbl.Cpt1(0))
    If (thisLbl.Compteur1 > 8) Then thisLbl.lblD1.Caption = ChrW$(CODE_ASC + thisLbl.Cpt1(1))
    If (thisLbl.Compteur1 > 98) Then thisLbl.lblC1.Caption = ChrW$(CODE_ASC + thisLbl.Cpt1(2))
    DoEvents            '// Refresh labels.

    thisLbl.Cpt1(0) = thisLbl.Cpt1(0) - 1
    If (thisLbl.Cpt1(0) < 0) Then thisLbl.Cpt1(0) = 9: thisLbl.Cpt1(1) = thisLbl.Cpt1(1) - 1
    If (thisLbl.Cpt1(1) < 0) Then thisLbl.Cpt1(1) = 9: thisLbl.Cpt1(2) = thisLbl.Cpt1(2) - 1

    thisLbl.Compteur1 = thisLbl.Compteur1 - 1

End Sub

' ----------------------------------------------------------------
'// Idem ci-dessus pour le rang 2.
' ----------------------------------------------------------------
Private Sub UpdateLabels2()

    If (Not mDebugLblOn) Then Exit Sub
    If (Not thisLbl.Rang2Isup) Then
        If mDebugOn Then Stop           'TODO ERR msg Compteur non initialisé
        Exit Sub
    End If

    thisLbl.lblU2.Caption = ChrW$(CODE_ASC + thisLbl.Cpt2(0))
    If (thisLbl.Compteur2 > 8) Then thisLbl.lblD2.Caption = ChrW$(CODE_ASC + thisLbl.Cpt2(1))
    If (thisLbl.Compteur2 > 98) Then thisLbl.lblC2.Caption = ChrW$(CODE_ASC + thisLbl.Cpt2(2))
    DoEvents            '// Refresh labels.

    thisLbl.Cpt2(0) = thisLbl.Cpt2(0) - 1
    If (thisLbl.Cpt2(0) < 0) Then thisLbl.Cpt2(0) = 9: thisLbl.Cpt2(1) = thisLbl.Cpt2(1) - 1
    If (thisLbl.Cpt2(1) < 0) Then thisLbl.Cpt2(1) = 9: thisLbl.Cpt2(2) = thisLbl.Cpt2(2) - 1

    thisLbl.Compteur2 = thisLbl.Compteur2 - 1

End Sub

' ----------------------------------------------------------------
'// Initialisation des compteurs des labels suivant la valeur de départ (thislbl.CompteurX)
' ----------------------------------------------------------------
Private Sub InitialiseLabels(RangLabel As E_LblRang)

    If (Not mDebugLblOn) Then Exit Sub      '// Mode debug et n'affiche pas les labels

    Dim lCpt As Long

    Select Case RangLabel
        Case E_LblRang.Rang1
            If (Not thisLbl.Rang1Isup) Then
                If mDebugOn Then Stop                       'TODO ERR msg labels rang1 non initialisés
                Exit Sub
            End If
            lCpt = thisLbl.Compteur1
            thisLbl.Cpt1(0) = lCpt Mod 10
            thisLbl.Cpt1(1) = Fix((lCpt / 10)) Mod 10
            thisLbl.Cpt1(2) = Fix((lCpt / 100)) Mod 10
            thisLbl.lblU1.Visible = True
            thisLbl.lblD1.Visible = True
            thisLbl.lblC1.Visible = True
        Case E_LblRang.Rang2
            If (Not thisLbl.Rang2Isup) Then
                If mDebugOn Then Stop                       'TODO ERR msg labels rang1 non initialisés
                Exit Sub
            End If
            lCpt = thisLbl.Compteur2
            thisLbl.Cpt2(0) = lCpt Mod 10
            thisLbl.Cpt2(1) = Fix((lCpt / 10)) Mod 10
            thisLbl.Cpt2(2) = Fix((lCpt / 100)) Mod 10
            thisLbl.lblU2.Visible = True
            thisLbl.lblD2.Visible = True
            thisLbl.lblC2.Visible = True
    End Select

    If (lCpt > 99) Then DefCoul (COUL_ORANGE): mCoulDef = False

End Sub

Private Sub DefCoul(lCoul As Long)  'TODO a voir...
    mCoulDef = True
    thisLbl.lblU1.ForeColor = lCoul
    thisLbl.lblD1.ForeColor = lCoul
    thisLbl.lblC1.ForeColor = lCoul
End Sub

'//----------- Pour le Mode Debug ---------------------------------
'//----------------------------------------------------------------
Public Property Let xDebugMode(ActiveDebugMode As Boolean)           '// Mode debug actif on/off
    mDebugOn = ActiveDebugMode
End Property
Public Property Get xDebugMode() As Boolean
     xDebugMode = mDebugOn
End Property
Public Property Let xDebugActiveLabels(ActiveLabel As Boolean)   '// Mode debug affiche labels on/off
    mDebugLblOn = ActiveLabel
End Property
'Public Property Get xDebugModeActiveLabels() As Boolean
'    xDebugModeActiveLabels = mDebugLblOn
'End Property
Public Property Get xDebugTimerEndLabelOn() As Single                '// Timer labels on
    xDebugTimerEndLabelOn = mTimerEndLblOn
End Property
Public Property Get xDebugTimerEndLabelOff() As Single               '// Timer labels off
    xDebugTimerEndLabelOff = mTimerEndLblOff
End Property

Public Sub xDebugTimerStart()
    If (Not mDebugOn) Then Exit Sub
    If mDebugLblOn Then
        mTimerStartLblOn = Timer
    Else
        mTimerStartLblOff = Timer
    End If
End Sub
Public Sub xDebugTimerEnd()
    If (Not mDebugOn) Then Exit Sub
    If mDebugLblOn Then
        mTimerEndLblOn = Timer
        mTimerEndLblOn = Round((mTimerEndLblOn - mTimerStartLblOn), 3)
    Else
        mTimerEndLblOff = Timer
'        If mTimerEndLblOff < mTimerStartLblOff Then mTimerEndLblOff = mTimerEndLblOff + 86400
        mTimerEndLblOff = Round((mTimerEndLblOff - mTimerStartLblOff), 3)
    End If
End Sub
