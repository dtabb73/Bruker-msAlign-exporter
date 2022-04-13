'***** Bruker DataAnalysis msAlign Exporter
' Created by Matt Willetts (Bruker), Kyle A. Brown (U-Wisconsin), and David L. Tabb (Institut Pasteur)
' This software makes it possible to conduct TopPIC searches on data from Bruker Q-TOF instruments.
' The SNAP / AutoMSn code at the very top handles "deconvolution."
' Everything else crafts an msalign file from the list of Compound objects.

'***** To Do:
' How can we detect the precursor charge better?  The loop through MS1 peaks seems error-prone.
' How can we specify the settings for SNAP and AutoMSn directly rather than relying on users?
' How do we "rescue" MS/MS scans for which no precursor charge is reported?
' Do we need to write "FEATURE" files for better TopPIC compatibility?

' Localise for US, using '.' rather than ',' as a decimal separator 
SetLocale(1033)  
Const Proton              = 1.00727647   
Const IntensityMultiplier = 10 
Const MZTolerance         = 0.001 
Const ZCeiling		  = 100
Const PkCountCeiling      = 1000
Dim   ZDistn(100)
Dim   PkCountDistn(1000)
Dim   PathAndFile, objFSO, msalignFile, qcFile, NoZCount
Dim   ZMinimum, ZQuartile1, ZMedian, ZQuartile3, ZMaximum, ZMode
Dim   PkCountMinimum, PkCountQuartile1, PkCountMedian, PkCountQuartile3, PkCountMaximum

'***** Run deconvolution in SNAP / MaxEnt / AutoMSn
'***** Configure for SNAP peak picking 
Analysis.Method.MassListParameters.DetectionAlgorithm = 2 
'***** Just select everything for peak picking 
Analysis.ClearChromatogramRangeSelections 
Analysis.AddChromatogramRangeSelection 0, 1000
Analysis.FindAutoMSn  
On Error Resume Next  
Analysis.Save
 
'***** Set up our output file for writing 
Set objFSO = CreateObject("Scripting.FileSystemObject")  
PathAndFile = Left(Analysis.Path, Len(Analysis.Path)-2) & ".msalign"
Set msalignFile = objFSO.CreateTextFile(PathAndFile)  
PathAndFile = Left(Analysis.Path, Len(Analysis.Path)-2) & ".qc.tsv"
Set qcFile = objFSO.CreateTextFile(PathAndFile)  
  
For Looper = 0 to ZCeiling
    ZDistn(Looper) = 0
Next
For Looper = 0 to PkCountCeiling
    PkCountDistn(Looper) = 0
Next
 
For Each ThisCompound In Analysis.Compounds
    Dim MS1Spec, MS2Spec, MSMSType, MS1ScanNumber, MS2ScanNumber
    Dim PrecursorMass, PrecursorCharge, PrecursorMZ, PrecursorIntensity, PrecursorRT
    Dim StartPos, StopPos, ThisPeak, PeakCounter 
 
    Set MS1Spec = ThisCompound(1) 
    Set MS2Spec = ThisCompound(2) 
 
'***** Determine the dissociation type (based on whether or not component Name contains the string "ETD"). 
    If InStr(ThisCompound.Name,"ETD") Then  
        MSMSType = "ETD" 
    Else 
        MSMSType = "CID" 
    End If 
 
'***** Determine the precursor m/z and mass, since we will need that to determine its charge 
    If IsNumeric(ThisCompound.Precursor) Then 
       PrecursorMZ = ThisCompound.Precursor 
    Else 
       '***** Otherwise grab the mass from the string naming this component 
       StartPos = InStr(ThisCompound.Name,"(") + 1 
       StopPos = InStr(ThisCompound.Name,")") 
       PrecursorMZ = Mid(ThisCompound.Name,StartPos,StopPos-StartPos) 
    End If 

    If IsNumeric(ThisCompound.RetentionTime) Then
       PrecursorRT = ThisCompound.RetentionTime
    Else
       PrecursorRT = 0
    End If 

    If IsNumeric(ThisCompound.Intensity) Then
       PrecursorIntensity = ThisCompound.Intensity
    Else
       PrecursorIntensity = 0
    End If 

'***** Determine the precursor charge state, keeping track of how many precursors cannot be matched back to the MS scans. 
    PrecursorCharge = -1 
 
    For Each ThisPeak in MS1Spec.MSPeakList
       If Abs(ThisPeak.m_over_z - PrecursorMZ)< MZTolerance Then 
           PrecursorCharge = ThisPeak.ChargeState 
           Exit For 
       End If 
    Next

    If PrecursorCharge = -1 Then
       ZDistn(0) = ZDistn(0) + 1
       ' If we couldn't determine the precursor charge we call it a +1 in the msalign file.
       PrecursorCharge = 1
       PrecursorMass = PrecursorMZ - Proton 
    Else
       ' Record this precursor charge in the array of precursor charge frequences
       ZDistn(PrecursorCharge) = ZDistn(PrecursorCharge) + 1
       PrecursorMass = (PrecursorMZ*PrecursorCharge) - (PrecursorCharge*Proton) 
    End If 

'***** Get the first MS scan number to report for this compound 
    StartPos = InStr(MS1Spec.Name,"#") 
    StopPos = InStrRev(MS1Spec.Name,"-") 
    If StopPos = 0 Then 
       MS1ScanNumber=Mid(MS1Spec.Name, StartPos+1) 
    Else 
       MS1ScanNumber=Mid(MS1Spec.Name, StartPos+1, StopPos-(StartPos+1)) 
    End If 


'***** Get the first MS/MS scan number to report for this compound 
    StartPos = InStr(MS2Spec.Name,"#") 
    StopPos = InStrRev(MS2Spec.Name,"-") 
    If StopPos = 0 Then 
       MS2ScanNumber=Mid(MS2Spec.Name, StartPos+1) 
    Else 
       MS2ScanNumber=Mid(MS2Spec.Name, StartPos+1, StopPos-(StartPos+1)) 
    End If 

    ' Add the number of MS/MS peaks for this spectrum to the Peak Count distribution
    PkCountDistn(MS2Spec.MSPeakList.Count) = PkCountDistn(MS2Spec.MSPeakList.Count)+1
 
'***** Now write this compound header to the msAlign file 
    msalignFile.WriteLine("BEGIN IONS")  
    msalignFile.WriteLine("ID=" & ThisCompound.CompoundNumber)
    msalignFile.WriteLine("FRACTION_ID=1")
    msalignFile.WriteLine("FILE_NAME=" & Analysis.Name)
    msalignFile.WriteLine("SCANS=" & MS2ScanNumber)
    msalignFile.WriteLine("RETENTION_TIME=" & PrecursorRT)
    ' Naturally the following will require update when something beyond MS/MS is being used.
    msalignFile.WriteLine("LEVEL=2")
    msalignFile.WriteLine("ACTIVATION=" & MSMSType)
    msalignFile.WriteLine("MS_ONE_ID=" & ThisCompound.CompoundNumber)
    msalignFile.WriteLine("MS_ONE_SCAN=" & MS1ScanNumber)
    msalignFile.WriteLine("PRECURSOR_MZ=" & PrecursorMZ)  
    msalignFile.WriteLine("PRECURSOR_CHARGE=" & PrecursorCharge) 
    msalignFile.WriteLine("PRECURSOR_MASS=" & PrecursorMass)
    ' TopFD writes floating point precursor intensities, but Bruker gives us integers
    msalignFile.WriteLine("PRECURSOR_INTENSITY=" & PrecursorIntensity & ".00")

'***** Now write the compound MS/MS peaks to the msAlign file 
    For Each ThisPeak in MS2Spec.MSPeakList 
       msalignFile.Write(Round(ThisPeak.ChargeState * (ThisPeak.m_over_z - Proton),4))
       msalignFile.Write(" ") 
       msalignFile.Write(Round(ThisPeak.Intensity * IntensityMultiplier,2)) 
       msalignFile.Write(" ") 
       msalignFile.WriteLine(ThisPeak.ChargeState) 
    Next 
 
    msalignFile.WriteLine("END IONS" & vbCrLf)  
      
'***** We're done with this Compound.  Continue to the next one.
Next
'***** Now complete writing the msalign file
msalignFile.Close  

'***** QC METRIC COMPUTATION

'***** Compute quartiles for Precursor Z, skipping unknown precursor charges
ZMinimum = -1
ZQuartile1 = -1
ZMedian = -1
ZQuartile3 = -1
ZMaximum = -1
ZMode  = -1

PositiveBinSum = 0
BiggestBinFreq = 0
For Looper = 1 to ZCeiling
    If ZDistn(Looper) > 0 Then
       PositiveBinSum = PositiveBinSum + ZDistn(Looper)
       If ZMinimum = -1 Then
          ZMinimum = Looper
       End If
       If ZDistn(Looper) > BiggestBinFreq Then
          BiggestBinFreq=ZDistn(Looper)
	  ZMode = Looper
       End If
       ZMaximum = Looper
    End If
Next

' What target number of spectra must be taken into account to find the first, second, and third quartiles?
Q1BinSum = PositiveBinSum / 4
Q2BinSum = PositiveBinSum / 2
Q3BinSum = Q1BinSum + Q2BinSum
PositiveBinSum = 0
For Looper = 1 to ZCeiling
    PositiveBinSum = PositiveBinSum + ZDistn(Looper)
    If ZQuartile1 = -1 AND PositiveBinSum >= Q1BinSum Then
       ZQuartile1 = Looper
    End IF
    If ZMedian = -1 AND PositiveBinSum >= Q2BinSum Then
       ZMedian = Looper
    End IF
    If ZQuartile3 = -1 AND PositiveBinSum >= Q3BinSum Then
       ZQuartile3 = Looper
    End IF
Next

'***** Compute quartiles for MS/MS Peak Count
PkCountMinimum = -1
PkCountQuartile1 = -1
PkCountMedian = -1
PkCountQuartile3 = -1
PkCountMaximum = -1
PkCountMode  = -1

PositiveBinSum = 0
BiggestBinFreq = 0
For Looper = 0 to PkCountCeiling
    If PkCountDistn(Looper) > 0 Then
       PositiveBinSum = PositiveBinSum + PkCountDistn(Looper)
       If PkCountMinimum = -1 Then
          PkCountMinimum = Looper
       End If
       If PkCountDistn(Looper) > BiggestBinFreq Then
          BiggestBinFreq=PkCountDistn(Looper)
	  PkCountMode = Looper
       End If
       PkCountMaximum = Looper
    End If
Next

' What target number of spectra must be taken into account to find the first, second, and third quartiles?
Q1BinSum = PositiveBinSum / 4
Q2BinSum = PositiveBinSum / 2
Q3BinSum = Q1BinSum + Q2BinSum
PositiveBinSum = 0
For Looper = 0 to PkCountCeiling
    PositiveBinSum = PositiveBinSum + PkCountDistn(Looper)
    If PkCountQuartile1 = -1 AND PositiveBinSum >= Q1BinSum Then
       PkCountQuartile1 = Looper
    End IF
    If PkCountMedian = -1 AND PositiveBinSum >= Q2BinSum Then
       PkCountMedian = Looper
    End IF
    If PkCountQuartile3 = -1 AND PositiveBinSum >= Q3BinSum Then
       PkCountQuartile3 = Looper
    End IF
Next

qcFile.WriteLine("AllSpectraCount" & vbTab & Analysis.SpectraCount)
qcFile.WriteLine("AutoMSnCompoundCount" & vbTab & Analysis.Compounds.Count)
qcFile.WriteLine("CompoundsLackingZ" & vbTab & ZDistn(0))
qcFile.WriteLine()

qcFile.WriteLine("ZMinimum" & vbTab & ZMinimum)
qcFile.WriteLine("ZQuartile1" & vbTab & ZQuartile1)
qcFile.WriteLine("ZMedian" & vbTab & ZMedian)
qcFile.WriteLine("ZQuartile3" & vbTab & ZQuartile3)
qcFile.WriteLine("ZMaximum" & vbTab & ZMaximum)
qcFile.WriteLine("ZMode" & vbTab & ZMode)

qcFile.WriteLine()
qcFile.WriteLine("MS2PkCountMinimum" & vbTab & PkCountMinimum)
qcFile.WriteLine("MS2PkCountQuartile1" & vbTab & PkCountQuartile1)
qcFile.WriteLine("MS2PkCountMedian" & vbTab & PkCountMedian)
qcFile.WriteLine("MS2PkCountQuartile3" & vbTab & PkCountQuartile3)
qcFile.WriteLine("MS2PkCountMaximum" & vbTab & PkCountMaximum)
qcFile.WriteLine("MS2PkCountMode" & vbTab & PkCountMode)

qcFile.WriteLine()
qcFile.WriteLine("CompoundZ" & vbTab & "Frequency")
For Looper = 0 to ZMaximum
    qcFile.WriteLine(Looper & vbTab & ZDistn(Looper))
Next

qcFile.WriteLine()
qcFile.WriteLine("MS2PkCount" & vbTab & "Frequency")
For Looper = 0 to PkCountMaximum
    qcFile.WriteLine(Looper & vbTab & PkCountDistn(Looper))
Next

qcFile.Close
MsgBox ("Wrote .msalign and .qc.tsv file successfully.")
Form.Close  
