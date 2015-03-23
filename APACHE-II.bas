Option Explicit


'###############################################################################
'# Acute Physiology and Chronic Health Evaluation II (APACHEII)
'# This function calculates the APACHEII score for a patient. This score estimates ICU mortality.
'# You should use the worst value for each physiological variable within the past 24 hours.
'#
'# This macro was written using the formula provided by Knaus WA et al (1985).
'# This function ensures that all values are within a sane range. Upper and lower limits allow for some "world record" setting conditions
'# and might suggest conditions beyond known survivable ranges. These are meant to ensure that there are not obvious coding or data entry errors.
'#
'# References:
'# * https://en.wikipedia.org/wiki/APACHE_II
'# * http://www.sfar.org/scores2/apache22.html - Used for quality assurance of macro below
'# * Knaus WA et al. APACHE II : A severity of disease classification system. Crit Care Med. 1985;13:818-2 (https://www.ncbi.nlm.nih.gov/pubmed/3928249)
'###############################################################################
Function APACHEII(AGE, TEMP, MAP, HR, RR, AA, PAO2, PH, HCO3, NA, K, ARF, CR, HCT, WBC, GCS, COIIC) As String
    Dim AGE_POINTS As Integer, APS_POINTS As Integer, CHRONIC_POINTS As Integer, ERRMSG As String
    
    ' Calculate Age Points with basic error checking
    If AGE >= 0 And AGE <= 44 Then
        AGE_POINTS = 0
    ElseIf AGE >= 45 And AGE <= 54 Then
        AGE_POINTS = 2
    ElseIf AGE >= 55 And AGE <= 64 Then
        AGE_POINTS = 3
    ElseIf AGE >= 65 And AGE <= 74 Then
        AGE_POINTS = 5
    ElseIf AGE >= 75 And AGE <= 130 Then
        AGE_POINTS = 6
    Else
        ERRMSG = "Age must be a number >=0 and <=130"
    End If
     
    ' Verify COIIC column is 0, 2, or 5
    If COIIC = 0 Or COIIC = 2 Or COIIC = 5 Then
        CHRONIC_POINTS = COIIC
    Else
        ERRMSG = "Chronic points (COIIC) must be 0, 2, or 5"
    End If
    
    ' Calculate APS_POINTS based on various factors
    APS_POINTS = 0
    
    ' Rectal temperature in ºC
    If TEMP <= 50# And TEMP >= 41# Then
        APS_POINTS = APS_POINTS + 4
    ElseIf TEMP < 41# And TEMP >= 39# Then
        APS_POINTS = APS_POINTS + 3
    ElseIf TEMP < 39# And TEMP >= 38.5 Then
        APS_POINTS = APS_POINTS + 1
    ElseIf (TEMP < 38.5 And TEMP >= 36#) Or IsEmpty(TEMP) Then
        APS_POINTS = APS_POINTS + 0
    ElseIf TEMP < 36# And TEMP >= 34# Then
        APS_POINTS = APS_POINTS + 1
    ElseIf TEMP < 34# And TEMP >= 32# Then
        APS_POINTS = APS_POINTS + 2
    ElseIf TEMP < 32# And TEMP >= 30# Then
        APS_POINTS = APS_POINTS + 3
    ElseIf TEMP < 30# And TEMP >= 10# Then
        APS_POINTS = APS_POINTS + 4
    Else
        ERRMSG = "Temperature (TEMP) must be <= 50ºC and >=10ºC"
    End If
    
    ' Mean Aterial Pressure (mm Hg)
    If MAP < 300 And MAP >= 160 Then
        APS_POINTS = APS_POINTS + 4
    ElseIf MAP < 160 And MAP >= 130 Then
        APS_POINTS = APS_POINTS + 3
    ElseIf MAP < 130 And MAP >= 110 Then
        APS_POINTS = APS_POINTS + 2
    ElseIf (MAP < 110 And MAP >= 70) Or IsEmpty(MAP) Then
        APS_POINTS = APS_POINTS + 0
    ElseIf MAP < 70 And MAP >= 50 Then
        APS_POINTS = APS_POINTS + 2
    ElseIf MAP < 50 And MAP > 0 Then
        APS_POINTS = APS_POINTS + 4
    Else
        ERRMSG = "Mean arterial pressure (MAP) must be < 300 and > 0 mm Hg"
    End If
    
    ' Heart Rate
    If HR < 400 And HR >= 180 Then
        APS_POINTS = APS_POINTS + 4
    ElseIf HR < 180 And HR >= 140 Then
        APS_POINTS = APS_POINTS + 3
    ElseIf HR < 140 And HR >= 110 Then
        APS_POINTS = APS_POINTS + 2
    ElseIf (HR < 110 And HR >= 70) Or IsEmpty(HR) Then
        APS_POINTS = APS_POINTS + 0
    ElseIf HR < 70 And HR >= 55 Then
        APS_POINTS = APS_POINTS + 2
    ElseIf HR < 55 And HR >= 40 Then
        APS_POINTS = APS_POINTS + 3
    ElseIf HR < 40 And HR > 0 Then
        APS_POINTS = APS_POINTS + 4
    Else
        ERRMSG = "Heart Rate (HR) must be < 400 and > 0"
    End If
    
    ' Respiratory Rate
    If RR < 100 And RR >= 50 Then
        APS_POINTS = APS_POINTS + 4
    ElseIf RR < 50 And RR >= 35 Then
        APS_POINTS = APS_POINTS + 3
    ElseIf RR < 35 And RR >= 25 Then
        APS_POINTS = APS_POINTS + 1
    ElseIf (RR < 25 And RR >= 12) Or IsEmpty(RR) Then
        APS_POINTS = APS_POINTS + 0
    ElseIf RR < 12 And RR >= 10 Then
        APS_POINTS = APS_POINTS + 1
    ElseIf RR < 10 And RR >= 6 Then
        APS_POINTS = APS_POINTS + 2
    ElseIf RR < 6 And RR > 0 Then
        APS_POINTS = APS_POINTS + 4
    Else
        ERRMSG = "Respiratory Rate (RR) must be < 100 and > 0"
    End If
        
    ' Aa (FiO2>50) and PaO2 (FiO2 < 50)
    If Not IsEmpty(AA) And Not IsEmpty(PAO2) Then
        ERRMSG = "Both AA and PAO2 are both provided; please clear one cell"
    ElseIf Not IsEmpty(AA) And IsEmpty(PAO2) Then
        
        ' If FIO2 >= 0.5, use AA
        If AA >= 500 Then
            APS_POINTS = APS_POINTS + 4
        ElseIf AA < 500 And AA >= 350 Then
            APS_POINTS = APS_POINTS + 3
        ElseIf AA < 350 And AA >= 200 Then
            APS_POINTS = APS_POINTS + 2
        ElseIf (AA < 200 And AA >= 0) Or IsEmpty(AA) Then
            APS_POINTS = APS_POINTS + 0
        Else
            ERRMSG = "Oxygenation A-aDO2 (AA) must be > 0"
        End If
    
    ElseIf IsEmpty(AA) And Not IsEmpty(PAO2) Then
    
        ' If FIO2 < 0.5, use PAO2
        If PAO2 < 55 And PAO2 > 0 Then
            APS_POINTS = APS_POINTS + 4
        ElseIf PAO2 <= 60 And PAO2 >= 55 Then
            APS_POINTS = APS_POINTS + 3
        ElseIf PAO2 <= 70 And PAO2 >= 61 Then
            APS_POINTS = APS_POINTS + 1
        ElseIf (PAO2 > 70) Or IsEmpty(AA) Then
            APS_POINTS = APS_POINTS + 0
        Else
            ERRMSG = "Oxygenation PAO2 must be > 0"
        End If
    
    End If
    
    ' Arterial pH
    If PH < 9# And PH >= 7.7 Then
        APS_POINTS = APS_POINTS + 4
    ElseIf PH < 7.7 And PH >= 7.6 Then
        APS_POINTS = APS_POINTS + 3
    ElseIf PH < 7.6 And PH >= 7.5 Then
        APS_POINTS = APS_POINTS + 1
    ElseIf (PH < 7.5 And PH >= 7.33) Or IsEmpty(PH) Then
        APS_POINTS = APS_POINTS + 0
    ElseIf PH < 7.33 And PH >= 7.25 Then
        APS_POINTS = APS_POINTS + 2
    ElseIf PH < 7.25 And PH >= 7.15 Then
        APS_POINTS = APS_POINTS + 3
    ElseIf PH < 7.15 And PH > 6 Then
        APS_POINTS = APS_POINTS + 4
    Else
        ERRMSG = "Arterial pH (PH) must be < 9 and > 6"
    End If
    
    ' Serum HCO3
    If HCO3 < 150 And HCO3 >= 52 Then
        APS_POINTS = APS_POINTS + 4
    ElseIf HCO3 < 52 And HCO3 >= 41 Then
        APS_POINTS = APS_POINTS + 3
    ElseIf HCO3 < 41 And HCO3 >= 32 Then
        APS_POINTS = APS_POINTS + 1
    ElseIf (HCO3 < 32 And HCO3 >= 22) Or IsEmpty(HCO3) Then
        APS_POINTS = APS_POINTS + 0
    ElseIf HCO3 < 22 And HCO3 >= 18 Then
        APS_POINTS = APS_POINTS + 2
    ElseIf HCO3 < 18 And HCO3 >= 15 Then
        APS_POINTS = APS_POINTS + 3
    ElseIf HCO3 < 15 And HCO3 > 0 Then
        APS_POINTS = APS_POINTS + 4
    Else
        ERRMSG = "Serum HCO3 (HCO3) must be < 150 and > 0"
    End If
    
    ' Serum Sodium (NA)
    If NA < 400 And NA >= 180 Then
        APS_POINTS = APS_POINTS + 4
    ElseIf NA < 180 And NA >= 160 Then
        APS_POINTS = APS_POINTS + 3
    ElseIf NA < 160 And NA >= 155 Then
        APS_POINTS = APS_POINTS + 2
    ElseIf NA < 155 And NA >= 150 Then
        APS_POINTS = APS_POINTS + 1
    ElseIf (NA < 150 And NA >= 130) Or IsEmpty(NA) Then
        APS_POINTS = APS_POINTS + 0
    ElseIf NA < 130 And NA >= 120 Then
        APS_POINTS = APS_POINTS + 2
    ElseIf NA < 120 And NA >= 111 Then
        APS_POINTS = APS_POINTS + 3
    ElseIf NA <= 110 And NA > 0 Then
        APS_POINTS = APS_POINTS + 4
    Else
        ERRMSG = "Serum Sodium (NA) must be < 400 and > 0"
    End If
    
    ' Serum Potassium (K)
    If K < 20 And K >= 7 Then
        APS_POINTS = APS_POINTS + 4
    ElseIf K < 7 And K >= 6 Then
        APS_POINTS = APS_POINTS + 3
    ElseIf K < 6 And K >= 5.5 Then
        APS_POINTS = APS_POINTS + 1
    ElseIf (K < 5.5 And K >= 3.5) Or IsEmpty(K) Then
        APS_POINTS = APS_POINTS + 0
    ElseIf K < 3.5 And K >= 3 Then
        APS_POINTS = APS_POINTS + 1
    ElseIf K < 3 And K >= 2.5 Then
        APS_POINTS = APS_POINTS + 2
    ElseIf K < 2.5 And K > 0 Then
        APS_POINTS = APS_POINTS + 4
    Else
        ERRMSG = "Serum Sodium (K) must be < 20 and > 0"
    End If
    
    ' Serium Creatinne + Acute Renal Failure
    Dim ARF_MULTIPLIER As Integer
    If UCase(ARF) = "Y" Then
        ARF_MULTIPLIER = 2
    ElseIf UCase(ARF) = "N" Then
        ARF_MULTIPLIER = 1
    Else
        ERRMSG = "Acute Renal Failure (ARF) must be set to 'Y' or 'N'"
    End If
    
    ' Double point score for CR if Acute Renal Failure (ARF) is present ('Y')
    If CR < 20 And CR >= 3.5 Then
        APS_POINTS = APS_POINTS + (4 * ARF_MULTIPLIER)
    ElseIf CR < 3.5 And CR >= 2# Then
        APS_POINTS = APS_POINTS + (3 * ARF_MULTIPLIER)
    ElseIf CR < 2# And CR >= 1.5 Then
        APS_POINTS = APS_POINTS + (2 * ARF_MULTIPLIER)
    ElseIf (CR < 1.5 And CR >= 0.6) Or IsEmpty(CR) Then
        APS_POINTS = APS_POINTS + (0 * ARF_MULTIPLIER)
    ElseIf CR < 0.6 And CR > 0 Then
        APS_POINTS = APS_POINTS + (2 * ARF_MULTIPLIER)
    Else
        ERRMSG = "Serium Creatinne (CR) must be < 20 and > 0"
    End If
    
    ' Hematrocrit (HCT)
    If HCT <= 100 And HCT >= 60 Then
        APS_POINTS = APS_POINTS + 4
    ElseIf HCT < 60 And HCT >= 50 Then
        APS_POINTS = APS_POINTS + 2
    ElseIf HCT < 50 And HCT >= 46 Then
        APS_POINTS = APS_POINTS + 1
    ElseIf (HCT < 46 And HCT >= 30) Or IsEmpty(HCT) Then
        APS_POINTS = APS_POINTS + 0
    ElseIf HCT < 30 And HCT >= 20 Then
        APS_POINTS = APS_POINTS + 2
    ElseIf HCT < 20 And HCT > 0 Then
        APS_POINTS = APS_POINTS + 4
    Else
        ERRMSG = "Hematrocrit (HCT) must be <= 100 and > 0"
    End If

    ' White Blood Count (WBC)
    If WBC <= 200 And WBC >= 40 Then
        APS_POINTS = APS_POINTS + 4
    ElseIf WBC < 40 And WBC >= 20 Then
        APS_POINTS = APS_POINTS + 2
    ElseIf WBC < 20 And WBC >= 15 Then
        APS_POINTS = APS_POINTS + 1
    ElseIf (WBC < 15 And WBC >= 3) Or IsEmpty(WBC) Then
        APS_POINTS = APS_POINTS + 0
    ElseIf WBC < 3 And WBC >= 1 Then
        APS_POINTS = APS_POINTS + 2
    ElseIf WBC < 1 And WBC >= 0# Then
        APS_POINTS = APS_POINTS + 4
    Else
        ERRMSG = "White Blood Count (WBC) must be <= 200 and >= 0"
    End If
    
    ' Glasgow coma score (GCS) (valid 3 to 15)
    If GCS >= 3 And GCS <= 15 Then
        APS_POINTS = APS_POINTS + (15 - GCS)
    Else
        ERRMSG = "Glasgow coma score (GCS) must be a number >=3 and <=15"
    End If
    
    ' Add the points from each category to determine the APACHEII score
    APACHEII = AGE_POINTS + APS_POINTS + CHRONIC_POINTS
    
    ' If we have an error, return that instead
    If ERRMSG <> "" Then
        APACHEII = "ERROR: " & ERRMSG
    End If

End Function

'###############################################################################
'# Given an APACHEII Score, calculate the predicted death rate
'###############################################################################
Function APACHEII_DEATHRATE(APACHEII)
    Dim LOGIT As Double
    LOGIT = -3.517 + APACHEII * 0.146
    APACHEII_DEATHRATE = Exp(LOGIT) / (1 + Exp(LOGIT))
End Function


'###############################################################################
'# Given an APACHEII Score, calculate the predicted death rate using an adjustment
'###############################################################################
Function APACHEII_DEATHRATE_ADJUSTED(APACHEII, ADJUSTMENT)
    Dim LOGIT As Double
    LOGIT = -3.517 + APACHEII * 0.146 + ADJUSTMENT
    APACHEII_DEATHRATE_ADJUSTED = Exp(LOGIT) / (1 + Exp(LOGIT))
End Function

