VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMonsterAttackSim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"colMonsterAttacks"
Option Explicit
Option Base 0

Private m_sAtkName(0 To 4) As String
Private m_nAtkType(0 To 4) As Integer
Private m_nAtkEnergy(0 To 4) As Integer
Private m_nAtkMin(0 To 4) As Integer
Private m_nAtkMax(0 To 4) As Integer
Private m_nAtkChance(0 To 4) As Integer
Private m_nAtkSuccess(0 To 4) As Integer
Private m_nAtkHitSpellMin(0 To 4) As Integer
Private m_nAtkHitSpellMax(0 To 4) As Integer
Private m_nAtkResist(0 To 4) As Integer
Private m_nAtkMRdmgResist(0 To 4) As Integer
Private m_nAtkDuration(0 To 4) As Integer

Private m_nStatAtkAttempted(0 To 4) As Currency
Private m_nStatAtkHits(0 To 4) As Currency
Private m_nStatAtkDmgResisted(0 To 4) As Currency
Private m_nStatAtkAttemptDodgedOrResisted(0 To 4) As Currency
Private m_nStatAtkTotalDamage(0 To 4) As Currency

Private m_nUserAC As Integer
Private m_nUserDR As Integer
Private m_nUserDodge As Integer
Private m_nUserMR As Integer
Private m_nUserAntiMagic As Integer

Private m_nEnergyPerRound As Integer
Private m_nNumberOfRounds As Long
Private m_nTotalAttacks As Currency
Private m_nTotalDamage As Currency
Private m_nMaxRoundDamage As Currency

Private m_bUseCPU As Boolean

Private m_cProgressBar As ProgressBar
Private nProgressBarScale As Integer
Private nProgressBarScaleCount As Long

Private m_sCombatLog As String
Private m_nCombatLogMaxRounds As Long
Private nCombatLogRoundCount As Long

Private m_sBetweenRoundName(0 To 4) As String
Private m_nBetweenRoundMin(0 To 4) As Integer
Private m_nBetweenRoundMax(0 To 4) As Integer
Private m_nBetweenRoundChance(0 To 4) As Integer
Private m_nBetweenRoundResistType(0 To 4) As Integer
Private m_nBetweenRoundResistDmgMR(0 To 4) As Integer
Private m_nBetweenRoundDuration(0 To 4) As Integer

Private nActiveAtkSpells_Ticks(0 To 4) As Integer
Private nActiveAtkSpells_DurationLeft(0 To 4) As Currency
Private nActiveAtkSpells_Value(0 To 4) As Integer
Private nActiveAtkSpells_ValueOriginal(0 To 4) As Integer

Private nActiveBetweenSpells_Ticks(0 To 4) As Integer
Private nActiveBetweenSpells_DurationLeft(0 To 4) As Currency
Private nActiveBetweenSpells_Value(0 To 4) As Integer
Private nActiveBetweenSpells_ValueOriginal(0 To 4) As Integer

Public Property Get nAtkDuration(ByVal Index As Integer) As Integer
    nAtkDuration = m_nAtkDuration(Index)
End Property

Public Property Let nAtkDuration(ByVal Index As Integer, ByVal iData As Integer)
    m_nAtkDuration(Index) = iData
End Property

Public Property Get nBetweenRoundDuration(ByVal Index As Integer) As Integer
    nBetweenRoundDuration = m_nBetweenRoundDuration(Index)
End Property

Public Property Let nBetweenRoundDuration(ByVal Index As Integer, ByVal iData As Integer)
    m_nBetweenRoundDuration(Index) = iData
End Property

Public Property Get nBetweenRoundResistDmgMR(ByVal Index As Integer) As Integer
    nBetweenRoundResistDmgMR = m_nBetweenRoundResistDmgMR(Index)
End Property

Public Property Let nBetweenRoundResistDmgMR(ByVal Index As Integer, ByVal iData As Integer)
    m_nBetweenRoundResistDmgMR(Index) = iData
End Property

Public Property Get nBetweenRoundResistType(ByVal Index As Integer) As Integer
    nBetweenRoundResistType = m_nBetweenRoundResistType(Index)
End Property

Public Property Let nBetweenRoundResistType(ByVal Index As Integer, ByVal iData As Integer)
    m_nBetweenRoundResistType(Index) = iData
End Property

Public Property Get sBetweenRoundName(ByVal Index As Integer) As String
    sBetweenRoundName = m_sBetweenRoundName(Index)
End Property

Public Property Let sBetweenRoundName(ByVal Index As Integer, ByVal sData As String)
    m_sBetweenRoundName(Index) = sData
End Property

Public Property Get nBetweenRoundMin(ByVal Index As Integer) As Integer
    nBetweenRoundMin = m_nBetweenRoundMin(Index)
End Property

Public Property Get nBetweenRoundMax(ByVal Index As Integer) As Integer
    nBetweenRoundMax = m_nBetweenRoundMax(Index)
End Property

Public Property Get nBetweenRoundChance(ByVal Index As Integer) As Integer
    nBetweenRoundChance = m_nBetweenRoundChance(Index)
End Property

Public Property Let nBetweenRoundChance(ByVal Index As Integer, ByVal iData As Integer)
    m_nBetweenRoundChance(Index) = iData
End Property

Public Property Let nBetweenRoundMax(ByVal Index As Integer, ByVal iData As Integer)
    m_nBetweenRoundMax(Index) = iData
End Property

Public Property Let nBetweenRoundMin(ByVal Index As Integer, ByVal iData As Integer)
    m_nBetweenRoundMin(Index) = iData
End Property

Public Property Get nMaxRoundDamage() As Long
    nMaxRoundDamage = m_nMaxRoundDamage
End Property

Public Property Get nCombatLogMaxRounds() As Long
    nCombatLogMaxRounds = m_nCombatLogMaxRounds
End Property

Public Property Let nCombatLogMaxRounds(ByVal lNewValue As Long)
    m_nCombatLogMaxRounds = lNewValue
End Property

Public Property Get sCombatLog() As String
    sCombatLog = m_sCombatLog
End Property

Public Property Get nTotalDamage() As Currency
    nTotalDamage = m_nTotalDamage
End Property

Public Property Get nTotalAttacks() As Long
    nTotalAttacks = m_nTotalAttacks
End Property

Public Property Let bUseCPU(ByVal vData As Boolean)
    m_bUseCPU = vData
End Property

Public Property Get bUseCPU() As Boolean
    bUseCPU = m_bUseCPU
End Property

Public Function RandomNumber(startNum As Integer, endNum As Integer) As Integer
    RandomNumber = Int(((endNum - startNum + 1) * Rnd) + startNum)
End Function

Private Sub AddToCombatLog(sLeft As String, Optional sRight As String = "")
'On Error GoTo error:

If nCombatLogRoundCount >= m_nCombatLogMaxRounds Then Exit Sub
If Len(sRight) > 0 And Len(sLeft) > 29 Then
    sLeft = Mid(sLeft, 1, 26) & "..."
End If
m_sCombatLog = m_sCombatLog & vbCrLf & sLeft
If Len(sRight) > 0 Then m_sCombatLog = m_sCombatLog & String(30 - Len(sLeft), " ") & sRight

out:
On Error Resume Next
Exit Sub
error:
Call privHandleError("AddToCombatLog")
Resume out:
End Sub

Public Sub RunSim()
Dim nRound As Long, nAttempt As Integer, nAttack As Integer, x As Integer
Dim nRemainingEnergy As Long, nEnergyUsed As Integer, bShowHitSpell As Boolean
Dim nUserMR_Adj As Integer, nDR_DamageResisted As Currency, nMR_Reduction As Currency
Dim bAttackHit As Boolean, bDodged As Boolean, bGlanced As Boolean, bResisted As Boolean
Dim nLastAttackType As Integer, nLastAttackEnergy As Integer, nAttack_AdjSuccessChance As Currency
Dim nRoll_AttackChance As Integer, nRoll_HitChance As Integer, nRoll_DodgeChance As Integer
Dim nDamage As Currency, nHitSpellDamage As Currency, nRoundDamage As Long, nOriginalDamage As Currency
Dim bSepShown As Boolean, bDurSpellApplied As Boolean

'On Error GoTo error:

Randomize

If m_nNumberOfRounds < 1 Then Exit Sub
If m_nEnergyPerRound < 1 Then Exit Sub
m_nTotalAttacks = 0
m_nTotalDamage = 0
m_nMaxRoundDamage = 0

Call ResetActiveSpells

Call ProgressBarSetRange(m_nNumberOfRounds)

m_sCombatLog = "=================================================================="
nCombatLogRoundCount = 0
If m_nCombatLogMaxRounds < 0 Then m_nCombatLogMaxRounds = 0

nRemainingEnergy = 0
For nRound = 1 To m_nNumberOfRounds
    
    nRoundDamage = 0
    bSepShown = False
    
    '============================================================================
    ' CHECK FOR EXTRA SPELL ROUND TICKS
    ' (SPELL ROUND = 3s, COMBAT ROUND 5s ... EVERY 3 ROUNDS = EXTRA SPELL ROUND)
    '============================================================================
    For x = 0 To 4
    
        '============================================================================
        ' regular attack spell duration and hit spell duration
        '============================================================================
        If nActiveAtkSpells_DurationLeft(x) > 0 Then
            If nActiveAtkSpells_Ticks(x) = 3 Then '15 seconds has based, second tick goes off
                
                nRoundDamage = nRoundDamage + nActiveAtkSpells_Value(x)
                m_nTotalDamage = m_nTotalDamage + nActiveAtkSpells_Value(x)
                m_nStatAtkTotalDamage(x) = m_nStatAtkTotalDamage(x) + nActiveAtkSpells_Value(x)
                m_nStatAtkDmgResisted(x) = m_nStatAtkDmgResisted(x) + (nActiveAtkSpells_ValueOriginal(x) - nActiveAtkSpells_Value(x))
                
                If Not bSepShown Then
                    Call AddToCombatLog("")
                    bSepShown = True
                End If
                
                nActiveAtkSpells_DurationLeft(x) = nActiveAtkSpells_DurationLeft(x) - 1
                
                If m_nAtkType(x) = 2 Then 'spell
                    Call AddToCombatLog("[" & m_sAtkName(x) & ", spell tick] for " & nActiveAtkSpells_Value(x) _
                        & " -- " & nActiveAtkSpells_DurationLeft(x) & " rounds rem.")
                Else
                    Call AddToCombatLog("[attack " & (x + 1) & ", hit spell tick] for " & nActiveAtkSpells_Value(x) _
                        & " -- " & nActiveAtkSpells_DurationLeft(x) & " rounds rem.")
                End If
                
                If nActiveAtkSpells_DurationLeft(x) < 1 Then Call ResetActiveAtkSpell(x)
            End If
        End If
        
        '============================================================================
        ' between round spell durations
        '============================================================================
        If nActiveBetweenSpells_DurationLeft(x) > 0 Then
            If nActiveBetweenSpells_Ticks(x) = 3 Then '15 seconds has based, second tick goes off
                
                nRoundDamage = nRoundDamage + nActiveBetweenSpells_Value(x)
                m_nTotalDamage = m_nTotalDamage + nActiveBetweenSpells_Value(x)
                
                If Not bSepShown Then
                    Call AddToCombatLog("")
                    bSepShown = True
                End If
                
                nActiveBetweenSpells_DurationLeft(x) = nActiveBetweenSpells_DurationLeft(x) - 1
                
                Call AddToCombatLog("[" & m_sBetweenRoundName(x) & ", between spell tick] for " & nActiveBetweenSpells_Value(x) _
                    & " -- " & nActiveBetweenSpells_DurationLeft(x) & " rounds rem.")
                
                If nActiveBetweenSpells_DurationLeft(x) < 1 Then Call ResetActiveBetweenSpell(x)
            End If
        End If
    Next x
    
    If bSepShown Then
        Call AddToCombatLog("")
        bSepShown = False
    End If
    
    Call AddToCombatLog("ROUND " & nRound & " / Energy: " _
        & nRemainingEnergy & " + " & m_nEnergyPerRound & " = " & (nRemainingEnergy + m_nEnergyPerRound))
    Call AddToCombatLog("")
    
    nRemainingEnergy = nRemainingEnergy + m_nEnergyPerRound
    
    '============================================================================
    ' BEGIN REGULAR ATTACK ATTEMPTS
    '============================================================================
    For nAttempt = 1 To 6

        nRoll_AttackChance = Me.RandomNumber(1, 100)
        nLastAttackType = 0
        nLastAttackEnergy = 0
        
        For nAttack = 0 To 4
        
            nDamage = 0: nHitSpellDamage = 0: nOriginalDamage = 0: nMR_Reduction = 0
            bShowHitSpell = False: bDurSpellApplied = False
            If nRoll_AttackChance <= m_nAtkChance(nAttack) And m_nAtkType(nAttack) > 0 Then
                
                nLastAttackEnergy = m_nAtkEnergy(nAttack)
                nLastAttackType = m_nAtkType(nAttack)
                
                If nRemainingEnergy < m_nAtkEnergy(nAttack) Then
                    'no energy for attack
                Else
                    m_nStatAtkAttempted(nAttack) = m_nStatAtkAttempted(nAttack) + 1
                    m_nTotalAttacks = m_nTotalAttacks + 1
                    
                    bAttackHit = False: bDodged = False: bGlanced = False: bResisted = False
                    
                    ' CHANCE TO HIT / CAST
                    nAttack_AdjSuccessChance = m_nAtkSuccess(nAttack)
                    If m_nAtkType(nAttack) <> 2 Then 'not spell
                        If m_nUserAC > 0 Then
                            '=((AC*AC)/100)/((ACCY*ACCY)/140)=fail %
                            nAttack_AdjSuccessChance = Round(1 - (((m_nUserAC * m_nUserAC) / 100) / ((nAttack_AdjSuccessChance * nAttack_AdjSuccessChance) / 140)), 2) * 100
                        Else
                            nAttack_AdjSuccessChance = 99
                        End If
                        'PHYSICAL ATTACKS = 9% MIN HIT CHANCE, 99% MAX
                        If nAttack_AdjSuccessChance < 9 Then nAttack_AdjSuccessChance = 9
                        If nAttack_AdjSuccessChance > 99 Then nAttack_AdjSuccessChance = 99
                    End If
                    
                    nRoll_HitChance = Me.RandomNumber(1, 100)
                    If nRoll_HitChance <= nAttack_AdjSuccessChance Then
                        bAttackHit = True
                        If m_nAtkType(nAttack) <> 2 And m_nUserDodge > 0 Then
                            nRoll_DodgeChance = Me.RandomNumber(1, 100)
                            If nRoll_DodgeChance <= m_nUserDodge Then
                                bAttackHit = False: bDodged = True
                                m_nStatAtkAttemptDodgedOrResisted(nAttack) = m_nStatAtkAttemptDodgedOrResisted(nAttack) + 1
                            End If
                        End If
                    End If
                    
                    If bAttackHit Then
                    
                        nDamage = Me.RandomNumber(m_nAtkMin(nAttack), m_nAtkMax(nAttack))
                        nOriginalDamage = nDamage
                        
                        '============================================================================
                        ' SPELL ATTACK ?
                        '============================================================================
                        If m_nAtkType(nAttack) = 2 Then 'spell
                            
                            nActiveAtkSpells_ValueOriginal(nAttack) = nDamage
                            
                            If m_nAtkMRdmgResist(nAttack) = 1 And nDamage > 0 Then
                                nMR_Reduction = Round(CalcResistedDamage(nDamage, m_nUserMR, m_nUserAntiMagic))
                                'If m_nAtkDuration(nAttack) = 0 Then
                                '    m_nStatAtkDmgResisted(nAttack) = m_nStatAtkDmgResisted(nAttack) + (nDamage - nMR_Reduction)
                                'End If
                                nDamage = nMR_Reduction
                            End If
                            
                            If m_nAtkDuration(nAttack) > 0 Then
                                If nActiveAtkSpells_DurationLeft(nAttack) < 1 Or nActiveAtkSpells_Value(nAttack) <> nDamage Then
                                    
                                    If IsSpellResisted(m_nAtkResist(nAttack), m_nUserMR, m_nUserAntiMagic) Then
                                                                    
                                        m_nStatAtkAttemptDodgedOrResisted(nAttack) = m_nStatAtkAttemptDodgedOrResisted(nAttack) + 1
                                        bResisted = True
                                        bAttackHit = False
                                        nDamage = 0
                                        
                                    Else
                                    
                                        Call ResetActiveAtkSpell(nAttack)
                                        nActiveAtkSpells_DurationLeft(nAttack) = m_nAtkDuration(nAttack)
                                        nActiveAtkSpells_Value(nAttack) = nDamage
                                        bDurSpellApplied = True
                                        nDamage = 0
                                    
                                    End If
                                Else
                                    'DURATION CONTINUES, NOT REALLY AN ATTEMPT
                                    m_nStatAtkAttempted(nAttack) = m_nStatAtkAttempted(nAttack) - 1
                                    m_nTotalAttacks = m_nTotalAttacks - 1
                                    bAttackHit = False
                                End If
                            Else
                                If IsSpellResisted(m_nAtkResist(nAttack), m_nUserMR, m_nUserAntiMagic) Then
                                                                    
                                    m_nStatAtkAttemptDodgedOrResisted(nAttack) = m_nStatAtkAttemptDodgedOrResisted(nAttack) + 1
                                    'm_nStatAtkDmgResisted(nAttack) = m_nStatAtkDmgResisted(nAttack) + nDamage
                                    bResisted = True
                                    bAttackHit = False
                                    nDamage = 0
                                  
                                End If
                                
                            End If
                            
                        '============================================================================
                        ' NORMAL ATTACK
                        '============================================================================
                        Else
                            'normal/rob
                            nDR_DamageResisted = m_nUserDR '(m_nUserDR / 2)
                            If nDR_DamageResisted > nDamage Then nDR_DamageResisted = nDamage
                            m_nStatAtkDmgResisted(nAttack) = m_nStatAtkDmgResisted(nAttack) + nDR_DamageResisted
                            nDamage = nDamage - nDR_DamageResisted
                            
                            '============================================================================
                            ' HIT SPELL APPLICATION CHECK
                            '============================================================================
                            If m_nAtkHitSpellMin(nAttack) > 0 Or m_nAtkHitSpellMax(nAttack) > 0 Or m_nAtkDuration(nAttack) > 0 Then
                                
                                nHitSpellDamage = Me.RandomNumber(m_nAtkHitSpellMin(nAttack), m_nAtkHitSpellMax(nAttack))
                                nActiveAtkSpells_ValueOriginal(nAttack) = nHitSpellDamage
                                
                                If m_nAtkMRdmgResist(nAttack) = 1 And nHitSpellDamage > 0 Then
                                    nMR_Reduction = Round(CalcResistedDamage(nHitSpellDamage, m_nUserMR, m_nUserAntiMagic))
                                    If m_nAtkDuration(nAttack) = 0 Then
                                        m_nStatAtkDmgResisted(nAttack) = m_nStatAtkDmgResisted(nAttack) + (nHitSpellDamage - nMR_Reduction)
                                    End If
                                    nHitSpellDamage = nMR_Reduction
                                End If
                                
                                If m_nAtkDuration(nAttack) > 0 Then
                                    If nActiveAtkSpells_DurationLeft(nAttack) < 1 Or nActiveAtkSpells_Value(nAttack) <> nHitSpellDamage Then
                                        
                                        If IsSpellResisted(m_nAtkResist(nAttack), m_nUserMR, m_nUserAntiMagic) Then
                                                                        
                                            m_nStatAtkDmgResisted(nAttack) = m_nStatAtkDmgResisted(nAttack) + nHitSpellDamage
                                            bShowHitSpell = True
                                            bResisted = True
                                            nHitSpellDamage = 0
                                            
                                        Else
                                        
                                            Call ResetActiveAtkSpell(nAttack)
                                            nActiveAtkSpells_DurationLeft(nAttack) = m_nAtkDuration(nAttack)
                                            nActiveAtkSpells_Value(nAttack) = nHitSpellDamage
                                            bDurSpellApplied = True
                                            bShowHitSpell = True
                                            nHitSpellDamage = 0
                                        
                                        End If
                                    End If
                                Else
                                    If IsSpellResisted(m_nAtkResist(nAttack), m_nUserMR, m_nUserAntiMagic) Then
                                                                        
                                        m_nStatAtkDmgResisted(nAttack) = m_nStatAtkDmgResisted(nAttack) + nHitSpellDamage
                                        bResisted = True
                                        nHitSpellDamage = 0
                                      
                                    End If
                                    
                                    bShowHitSpell = True
                                End If
                                                                    
                                
                                If Not bShowHitSpell Then nHitSpellDamage = 0
                            End If
                        End If
                        
                        If nDamage <= 0 Then
                            If m_nAtkType(nAttack) <> 2 Then bGlanced = True
                            nDamage = 0
                        End If
                        
                        If bAttackHit Then
                        
                            If nOriginalDamage > nDamage And m_nAtkType(nAttack) = 2 Then 'spell
                                m_nStatAtkDmgResisted(nAttack) = m_nStatAtkDmgResisted(nAttack) + (nOriginalDamage - nDamage)
                            End If
                            
                            nRoundDamage = nRoundDamage + nDamage + nHitSpellDamage
                            m_nTotalDamage = m_nTotalDamage + nDamage + nHitSpellDamage
                            
                            nRemainingEnergy = nRemainingEnergy - m_nAtkEnergy(nAttack)
                            
                            m_nStatAtkHits(nAttack) = m_nStatAtkHits(nAttack) + 1
                            m_nStatAtkTotalDamage(nAttack) = m_nStatAtkTotalDamage(nAttack) + nDamage + nHitSpellDamage
                            
                            If m_nAtkType(nAttack) = 2 And m_nAtkDuration(nAttack) > 0 Then 'duration spell
                                Call AddToCombatLog(m_sAtkName(nAttack) & " applied (" & nActiveAtkSpells_ValueOriginal(nAttack) & ")", _
                                    "Energy used: " & m_nAtkEnergy(nAttack) & " ... Remaining: " & nRemainingEnergy)
                            Else
                                Call AddToCombatLog(m_sAtkName(nAttack) & " for " & Round(nDamage) _
                                    & IIf(bGlanced, " (GLANCE)", ""), _
                                    "Energy used: " & m_nAtkEnergy(nAttack) & " ... Remaining: " & nRemainingEnergy)
                            End If
                            
                            If bShowHitSpell Then
                                If bDurSpellApplied Then
                                    Call AddToCombatLog(" - duration spell applied (" & nActiveAtkSpells_Value(nAttack) & ")")
                                Else
                                    Call AddToCombatLog(" - hit spell " & IIf(bResisted, "(resist)", "for " & Round(nHitSpellDamage)))
                                End If
                            End If
                        End If
                    End If
                    
                    If Not bAttackHit Then
                        If m_nAtkType(nAttack) = 2 Then 'spell
                            If m_nAtkDuration(nAttack) = 0 Or nActiveAtkSpells_DurationLeft(nAttack) < 1 Then
                            'If m_nAtkDuration(nAttack) = 0 Or bResisted Then
                                nEnergyUsed = Round(m_nAtkEnergy(nAttack) / 2, 0)
                                
                                nRemainingEnergy = nRemainingEnergy - nEnergyUsed
                                
                                Call AddToCombatLog(m_sAtkName(nAttack) & " (" _
                                    & IIf(bResisted, "RESIST", "FAIL") & ")", _
                                    "Energy used: " & nEnergyUsed & " ... Remaining: " & nRemainingEnergy)
                            End If
                        Else
                            nRemainingEnergy = nRemainingEnergy - m_nAtkEnergy(nAttack)
                            
                            Call AddToCombatLog(m_sAtkName(nAttack) & " (" _
                                & IIf(bDodged, "DODGE", "MISS") & ")", _
                                "Energy used: " & m_nAtkEnergy(nAttack) & " ... Remaining: " & nRemainingEnergy)
                        End If
                    End If
                End If
                
                GoTo next_attempt:
            End If
        Next nAttack
        
next_attempt:
        If nLastAttackType = 1 Then 'spell
            If nRemainingEnergy < nLastAttackEnergy Then GoTo next_round:
        End If

    Next nAttempt
    
next_round:

    '============================================================================
    ' BETWEEN ROUND SPELL CAST
    '============================================================================
    bSepShown = False
    nRoll_AttackChance = Me.RandomNumber(1, 100)
    For x = 0 To 4
        bAttackHit = False: bResisted = False: bDurSpellApplied = False
        
        If m_nBetweenRoundChance(x) > 0 Then
            If nRoll_AttackChance <= m_nBetweenRoundChance(x) Then
            
                nDamage = Me.RandomNumber(m_nBetweenRoundMin(x), m_nBetweenRoundMax(x))
                nActiveBetweenSpells_ValueOriginal(x) = nDamage
                            
                If m_nBetweenRoundResistDmgMR(x) = 1 And nDamage > 0 Then
                    nMR_Reduction = Round(CalcResistedDamage(nDamage, m_nUserMR, m_nUserAntiMagic))
                    nDamage = nMR_Reduction
                End If
                
                If m_nBetweenRoundDuration(x) > 0 Then
                    If nActiveBetweenSpells_DurationLeft(x) < 1 Or nActiveBetweenSpells_Value(x) <> nDamage Then
                        
                        If IsSpellResisted(m_nBetweenRoundResistType(x), m_nUserMR, m_nUserAntiMagic) Then
                            bResisted = True
                            nDamage = 0
                        Else
                            Call ResetActiveBetweenSpell(x)
                            nActiveBetweenSpells_DurationLeft(x) = m_nBetweenRoundDuration(x)
                            nActiveBetweenSpells_Value(x) = nDamage
                            bDurSpellApplied = True
                            nDamage = 0
                        End If
                    End If
                Else
                    If IsSpellResisted(m_nBetweenRoundResistType(x), m_nUserMR, m_nUserAntiMagic) Then
                        bResisted = True
                        nDamage = 0
                    Else
                        bAttackHit = True
                    End If
                End If
            End If
        End If
        
        If Not bSepShown And (bDurSpellApplied Or bResisted Or bAttackHit) Then
            Call AddToCombatLog("")
            bSepShown = True
        End If
        
        If bDurSpellApplied Then
            Call AddToCombatLog("[between round] " & m_sBetweenRoundName(x) & " - duration spell applied (" & nActiveBetweenSpells_Value(x) & ")")
            GoTo duration_ticks:
        ElseIf bResisted Then
            Call AddToCombatLog("[between round] " & m_sBetweenRoundName(x) & " (RESIST)")
            GoTo duration_ticks:
        ElseIf bAttackHit Then
            Call AddToCombatLog("[between round] " & m_sBetweenRoundName(x) & " for " & Round(nDamage))
            nRoundDamage = nRoundDamage + nDamage
            m_nTotalDamage = m_nTotalDamage + nDamage
            GoTo duration_ticks:
        End If
    Next x
    
duration_ticks:

    '============================================================================
    ' ALL SPELL DURATION TICKS
    '============================================================================
    bSepShown = False
    For x = 0 To 4
        If nActiveAtkSpells_DurationLeft(x) > 0 Then
                
            nRoundDamage = nRoundDamage + nActiveAtkSpells_Value(x)
            m_nTotalDamage = m_nTotalDamage + nActiveAtkSpells_Value(x)
            m_nStatAtkTotalDamage(x) = m_nStatAtkTotalDamage(x) + nActiveAtkSpells_Value(x)
            m_nStatAtkDmgResisted(x) = m_nStatAtkDmgResisted(x) + (nActiveAtkSpells_ValueOriginal(x) - nActiveAtkSpells_Value(x))
            
            If Not bSepShown Then
                Call AddToCombatLog("")
                bSepShown = True
            End If
            
            nActiveAtkSpells_DurationLeft(x) = nActiveAtkSpells_DurationLeft(x) - 1
            
            If m_nAtkType(x) = 2 Then 'spell
                Call AddToCombatLog("[" & m_sAtkName(x) & ", attack spell tick] for " & nActiveAtkSpells_Value(x) _
                    & " -- " & nActiveAtkSpells_DurationLeft(x) & " rounds rem.")
            Else
                Call AddToCombatLog("[attack " & (x + 1) & ", hit spell tick] for " & nActiveAtkSpells_Value(x) _
                    & " -- " & nActiveAtkSpells_DurationLeft(x) & " rounds rem.")
            End If
            
            If nActiveAtkSpells_Ticks(x) >= 3 Then
                nActiveAtkSpells_Ticks(x) = 1
            Else
                nActiveAtkSpells_Ticks(x) = nActiveAtkSpells_Ticks(x) + 1
            End If
            
            If nActiveAtkSpells_DurationLeft(x) < 1 Then Call ResetActiveAtkSpell(x)
        End If
        
        If nActiveBetweenSpells_DurationLeft(x) > 0 Then
                
            nRoundDamage = nRoundDamage + nActiveBetweenSpells_Value(x)
            m_nTotalDamage = m_nTotalDamage + nActiveBetweenSpells_Value(x)
            
            If Not bSepShown Then
                Call AddToCombatLog("")
                bSepShown = True
            End If
            
            nActiveBetweenSpells_DurationLeft(x) = nActiveBetweenSpells_DurationLeft(x) - 1
            
            Call AddToCombatLog("[" & m_sBetweenRoundName(x) & ", between spell tick] for " & nActiveBetweenSpells_Value(x) _
                & " -- " & nActiveBetweenSpells_DurationLeft(x) & " rounds rem.")
            
            If nActiveBetweenSpells_Ticks(x) >= 3 Then
                nActiveBetweenSpells_Ticks(x) = 1
            Else
                nActiveBetweenSpells_Ticks(x) = nActiveBetweenSpells_Ticks(x) + 1
            End If
            
            If nActiveBetweenSpells_DurationLeft(x) < 1 Then Call ResetActiveBetweenSpell(x)
        End If
    Next x
    
    If nRoundDamage > m_nMaxRoundDamage Then m_nMaxRoundDamage = nRoundDamage
    
    Call AddToCombatLog("")
    Call AddToCombatLog("Damage for round: " & nRoundDamage, "Energy Remaining: " & nRemainingEnergy)
    Call AddToCombatLog("==================================================================")
    nCombatLogRoundCount = nCombatLogRoundCount + 1
    
    Call ProgressBarIncrease
    If Not m_bUseCPU Then DoEvents
Next nRound

out:
On Error Resume Next
Exit Sub
error:
Call privHandleError("RunSim")
Resume out:
End Sub

Private Sub ResetActiveAtkSpell(ByVal nAttack As Integer)
On Error GoTo error:

nActiveAtkSpells_DurationLeft(nAttack) = 0
nActiveAtkSpells_Ticks(nAttack) = 0
nActiveAtkSpells_Value(nAttack) = 0

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("ResetActiveAtkSpell")
Resume out:
End Sub

Private Sub ResetActiveBetweenSpell(ByVal nAttack As Integer)
On Error GoTo error:

nActiveBetweenSpells_DurationLeft(nAttack) = 0
nActiveBetweenSpells_Ticks(nAttack) = 0
nActiveBetweenSpells_Value(nAttack) = 0

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("ResetActiveBetweenSpell")
Resume out:
End Sub

Public Function IsSpellResisted(ByVal nAttackResistType As Integer, ByVal nMR As Integer, ByVal nAntiMagic As Integer) As Boolean
Dim nRoll_ResistChance As Integer
On Error GoTo error:

IsSpellResisted = False

If (nAttackResistType = 1 And nAntiMagic = 1) Or nAttackResistType = 2 Then
    'if( TYPEofRESIST=Never , 0 , IF( ANTI_MAGIC=Yes or TYPEofRESIST=Yes , IF( MR>196 , 0.98 , MR/200 ) , 0 ) )
    nRoll_ResistChance = Me.RandomNumber(1, 100)
    If nMR > 196 Then nMR = 196
    If nRoll_ResistChance <= nMR / 2 Then
        IsSpellResisted = True
    End If
End If

out:
On Error Resume Next
Exit Function
error:
Call privHandleError("IsSpellResisted")
Resume out:
End Function

Public Function CalcResistedDamage(ByVal nDamage As Currency, ByVal nMR As Integer, ByVal nAntiMagic As Integer) As Currency
Dim nPercentReduction As Currency
On Error GoTo error:

If nMR > 150 Then nMR = 150

If nAntiMagic = 0 Then
    'DAMAGE = DAMAGE - ( DAMAGE * IF( MR<50, (MR-50)/100, IF( MR>150, 0.5, (MR-50)/200 ) ) )
    If nMR < 50 Then
        nPercentReduction = Round((nMR - 50) / 100, 2)
    Else
        nPercentReduction = Round((nMR - 50) / 200, 2)
    End If
Else
    'DAMAGE = DAMAGE - ( DAMAGE * IF( MR>150 , 0.75 , MR/200 ) )
    nPercentReduction = Round(nMR / 200, 2)
End If

CalcResistedDamage = nDamage - (nDamage * nPercentReduction)

out:
On Error Resume Next
Exit Function
error:
Call privHandleError("CalcResistedDamage")
Resume out:
End Function

Public Property Set cProgressBar(ByVal vData As ProgressBar)
    Set m_cProgressBar = vData
End Property

Public Property Get cProgressBar() As ProgressBar
    Set cProgressBar = m_cProgressBar
End Property

Private Sub ProgressBarSetRange(ByVal MaxValue As Double)
Dim nNewMax As Long
On Error GoTo error:

If m_cProgressBar Is Nothing Then Exit Sub

nProgressBarScale = 0

If MaxValue > MaxInt Then
    nNewMax = MaxValue
    Do While nNewMax > MaxInt
        nProgressBarScale = nProgressBarScale + 2
        If MaxValue / nProgressBarScale < MaxInt Then
            nNewMax = MaxValue / nProgressBarScale
            GoTo done:
        End If
    Loop
Else
    nNewMax = MaxValue
End If

done:
nNewMax = Fix(nNewMax)

nProgressBarScaleCount = 1
m_cProgressBar.Value = 0
m_cProgressBar.Min = 0
m_cProgressBar.Max = nNewMax

out:
On Error Resume Next
Exit Sub
error:
Call privHandleError("ProgressBarSetRange")
Resume out:
End Sub

Private Sub ProgressBarIncrease()
On Error GoTo error:

If m_cProgressBar Is Nothing Then Exit Sub

If nProgressBarScale > 0 Then
    If nProgressBarScaleCount = nProgressBarScale Then
        If m_cProgressBar.Value + 1 < m_cProgressBar.Max Then m_cProgressBar.Value = m_cProgressBar.Value + 1
        nProgressBarScaleCount = 1
    Else
        nProgressBarScaleCount = nProgressBarScaleCount + 1
    End If
Else
    If m_cProgressBar.Value + 1 < m_cProgressBar.Max Then m_cProgressBar.Value = m_cProgressBar.Value + 1
End If

out:
On Error Resume Next
Exit Sub
error:
Call privHandleError("ProgressBarIncrease")
Resume out:
End Sub

Public Sub ResetValues()
Dim x As Integer
On Error GoTo error:

For x = 0 To 4
    m_sAtkName(x) = ""
    m_nAtkType(x) = 0
    m_nAtkEnergy(x) = 0
    m_nAtkMin(x) = 0
    m_nAtkMax(x) = 0
    m_nAtkChance(x) = 0
    m_nAtkSuccess(x) = 0
    m_nAtkHitSpellMin(x) = 0
    m_nAtkHitSpellMax(x) = 0
    m_nAtkResist(x) = 0
    m_nAtkMRdmgResist(x) = 0
    
    m_nStatAtkAttempted(x) = 0
    m_nStatAtkHits(x) = 0
    m_nStatAtkDmgResisted(x) = 0
    m_nStatAtkAttemptDodgedOrResisted(x) = 0
    m_nStatAtkTotalDamage(x) = 0
    
    m_nUserAC = 0
    m_nUserDR = 0
    m_nUserDodge = 0
    m_nUserMR = 0
    m_nUserAntiMagic = 0
    
    m_nEnergyPerRound = 0
    m_nNumberOfRounds = 0
    m_nTotalAttacks = 0
    m_nTotalDamage = 0
    m_nMaxRoundDamage = 0
    
    m_sCombatLog = ""
    nCombatLogRoundCount = 0
    
    m_sBetweenRoundName(x) = ""
    m_nBetweenRoundMin(x) = 0
    m_nBetweenRoundMax(x) = 0
    m_nBetweenRoundChance(x) = 0
    m_nBetweenRoundResistType(x) = 0
    m_nBetweenRoundResistDmgMR(x) = 0
    
Next x

out:
On Error Resume Next
Exit Sub
error:
Call privHandleError("ResetAttacks")
Resume out:
End Sub

Private Sub ResetActiveSpells()
Dim x As Integer
On Error GoTo error:

For x = 0 To 4
    nActiveAtkSpells_Ticks(x) = 0
    nActiveAtkSpells_DurationLeft(x) = 0
    nActiveAtkSpells_Value(x) = 0
    
    nActiveBetweenSpells_Ticks(x) = 0
    nActiveBetweenSpells_DurationLeft(x) = 0
    nActiveBetweenSpells_Value(x) = 0
Next x

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("ResetActiveSpells")
Resume out:
End Sub

Private Sub privHandleError(Optional ByVal sSource As String)

If Not sSource = "" Then
    sSource = "Error " & Err.Number & " in [" & sSource & "]:" & vbCrLf
Else
    sSource = "Error " & Err.Number & ": "
End If

Select Case Err.Number
    Case 6: MsgBox sSource & "Overflow (One of the values entered was too large)."
    Case Else: MsgBox sSource & Err.Source & vbCrLf & Err.Description, vbExclamation
End Select

Err.clear
End Sub

Public Property Let nNumberOfRounds(ByVal vData As Long)
    m_nNumberOfRounds = vData
End Property

Public Property Get nNumberOfRounds() As Long
    nNumberOfRounds = m_nNumberOfRounds
End Property

Public Property Let nEnergyPerRound(ByVal vData As Integer)
    m_nEnergyPerRound = vData
End Property

Public Property Get nEnergyPerRound() As Integer
    nEnergyPerRound = m_nEnergyPerRound
End Property

Public Property Let nUserAntiMagic(ByVal vData As Integer)
    m_nUserAntiMagic = vData
End Property

Public Property Get nUserAntiMagic() As Integer
    nUserAntiMagic = m_nUserAntiMagic
End Property

Public Property Let nUserMR(ByVal vData As Integer)
    m_nUserMR = vData
End Property

Public Property Get nUserMR() As Integer
    nUserMR = m_nUserMR
End Property

Public Property Let nUserDodge(ByVal vData As Integer)
    m_nUserDodge = vData
End Property

Public Property Get nUserDodge() As Integer
    nUserDodge = m_nUserDodge
End Property

Public Property Let nUserDR(ByVal vData As Integer)
    m_nUserDR = vData
End Property

Public Property Get nUserDR() As Integer
    nUserDR = m_nUserDR
End Property

Public Property Let nUserAC(ByVal vData As Integer)
    m_nUserAC = vData
End Property

Public Property Get nUserAC() As Integer
    nUserAC = m_nUserAC
End Property

Public Property Let nStatAtkTotalDamage(ByVal Index As Integer, ByVal vData As Currency)
    m_nStatAtkTotalDamage(Index) = vData
End Property

Public Property Get nStatAtkTotalDamage(ByVal Index As Integer) As Currency
    nStatAtkTotalDamage = m_nStatAtkTotalDamage(Index)
End Property

Public Property Let nStatAtkAttemptDodgedOrResisted(ByVal Index As Integer, ByVal vData As Currency)
    m_nStatAtkAttemptDodgedOrResisted(Index) = vData
End Property

Public Property Get nStatAtkAttemptDodgedOrResisted(ByVal Index As Integer) As Currency
    nStatAtkAttemptDodgedOrResisted = m_nStatAtkAttemptDodgedOrResisted(Index)
End Property

Public Property Let nStatAtkDmgResisted(ByVal Index As Integer, ByVal vData As Currency)
    m_nStatAtkDmgResisted(Index) = vData
End Property

Public Property Get nStatAtkDmgResisted(ByVal Index As Integer) As Currency
    nStatAtkDmgResisted = m_nStatAtkDmgResisted(Index)
End Property

Public Property Let nStatAtkHits(ByVal Index As Integer, ByVal vData As Currency)
    m_nStatAtkHits(Index) = vData
End Property

Public Property Get nStatAtkHits(ByVal Index As Integer) As Currency
    nStatAtkHits = m_nStatAtkHits(Index)
End Property

Public Property Let nStatAtkAttempted(ByVal Index As Integer, ByVal vData As Currency)
    m_nStatAtkAttempted(Index) = vData
End Property

Public Property Get nStatAtkAttempted(ByVal Index As Integer) As Currency
    nStatAtkAttempted = m_nStatAtkAttempted(Index)
End Property

Public Property Let nAtkMRdmgResist(ByVal Index As Integer, ByVal vData As Integer)
    m_nAtkMRdmgResist(Index) = vData
End Property

Public Property Get nAtkMRdmgResist(ByVal Index As Integer) As Integer
    nAtkMRdmgResist = m_nAtkMRdmgResist(Index)
End Property

Public Property Let nAtkResist(ByVal Index As Integer, ByVal vData As Integer)
    m_nAtkResist(Index) = vData
End Property

Public Property Get nAtkResist(ByVal Index As Integer) As Integer
    nAtkResist = m_nAtkResist(Index)
End Property

Public Property Let nAtkHitSpellMax(ByVal Index As Integer, ByVal vData As Integer)
    m_nAtkHitSpellMax(Index) = vData
End Property

Public Property Get nAtkHitSpellMax(ByVal Index As Integer) As Integer
    nAtkHitSpellMax = m_nAtkHitSpellMax(Index)
End Property

Public Property Let nAtkHitSpellMin(ByVal Index As Integer, ByVal vData As Integer)
    m_nAtkHitSpellMin(Index) = vData
End Property

Public Property Get nAtkHitSpellMin(ByVal Index As Integer) As Integer
    nAtkHitSpellMin = m_nAtkHitSpellMin(Index)
End Property

Public Property Let nAtkSuccess(ByVal Index As Integer, ByVal vData As Integer)
    m_nAtkSuccess(Index) = vData
End Property

Public Property Get nAtkSuccess(ByVal Index As Integer) As Integer
    nAtkSuccess = m_nAtkSuccess(Index)
End Property

Public Property Let nAtkChance(ByVal Index As Integer, ByVal vData As Integer)
    m_nAtkChance(Index) = vData
End Property

Public Property Get nAtkChance(ByVal Index As Integer) As Integer
    nAtkChance = m_nAtkChance(Index)
End Property

Public Property Let nAtkMax(ByVal Index As Integer, ByVal vData As Integer)
    m_nAtkMax(Index) = vData
End Property

Public Property Get nAtkMax(ByVal Index As Integer) As Integer
    nAtkMax = m_nAtkMax(Index)
End Property

Public Property Let nAtkMin(ByVal Index As Integer, ByVal vData As Integer)
    m_nAtkMin(Index) = vData
End Property

Public Property Get nAtkMin(ByVal Index As Integer) As Integer
    nAtkMin = m_nAtkMin(Index)
End Property

Public Property Let nAtkEnergy(ByVal Index As Integer, ByVal vData As Integer)
    m_nAtkEnergy(Index) = vData
End Property

Public Property Get nAtkEnergy(ByVal Index As Integer) As Integer
    nAtkEnergy = m_nAtkEnergy(Index)
End Property

Public Property Let nAtkType(ByVal Index As Integer, ByVal vData As Integer)
    m_nAtkType(Index) = vData
End Property

Public Property Get nAtkType(ByVal Index As Integer) As Integer
    nAtkType = m_nAtkType(Index)
End Property

Public Property Let sAtkName(ByVal Index As Integer, ByVal vData As String)
    m_sAtkName(Index) = vData
End Property

Public Property Get sAtkName(ByVal Index As Integer) As String
    sAtkName = m_sAtkName(Index)
End Property

Private Sub Class_Initialize()
m_nCombatLogMaxRounds = 10
Randomize
End Sub
