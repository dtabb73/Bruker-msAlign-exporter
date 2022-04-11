'***** Bruker DataAnalysis msAlign Exporter
' Created by Matt Willetts (Bruker), Kyle A. Brown (U-Wisconsin), and David L. Tabb (Institut Pasteur)
' This software makes it possible to conduct TopPIC searches on data from Bruker Q-TOF instruments.
' The SNAP / MaxEnt / AutoMSn code at the very top handles "deconvolution."
' Everything else crafts an msalign file from the list of Compound objects.

'***** To Do:
' How can we detect the precursor charge better?  The loop through MS1 peaks seems error-prone.
' How can we specify the settings for SNAP, MaxEnt, and AutoMSn directly rather than relying on users?
' How do we "rescue" MS/MS scans for which no precursor charge is reported?
' Let's improve QC reporting, such as charge state histograms.
' Do we need to write "FEATURE" files for better TopPIC compatibility?

' Localise for US, using '.' rather than ',' as a decimal separator 
SetLocale(1033)  
Const Proton              = 1.00727647   
Const sExtension          = ".msalign"   
Const IntensityMultiplier = 10 
Const MZTolerance         = 0.001 
Dim   PathAndFile, objFSO, objFile, CompoundCount, NoZCount

'***** Run deconvolution in SNAP / MaxEnt / AutoMSn
'***** Configure for SNAP peak picking 
Analysis.Method.MassListParameters.DetectionAlgorithm = 2 
'***** Just select everything for peak picking 
Analysis.ClearChromatogramRangeSelections 
Analysis.AddChromatogramRangeSelection 0, 1000
Analysis.FindAutoMSn  
On Error Resume Next  
 
'***** Set up our output file for writing 
PathAndFile = Left(Analysis.Path, Len(Analysis.Path)-2) & sExtension
CompoundCount = 0
NoZCount = 0
 
Set objFSO = CreateObject("Scripting.FileSystemObject")  
Set objFile = objFSO.CreateTextFile(PathAndFile)  
  
For Each comp In Analysis.compounds
    CompoundCount = CompoundCount + 1
    Dim MS1Spec, MS2Spec 
    Dim MSMSType, PrecursorMass, PrecursorCharge, PrecursorMZ, PrecursorIntensity, PrecursorRT, StartPos, StopPos, ThisPeak, MS1ScanNumber, MS2ScanNumber, PeakCounter 
 
    Set MS1Spec = comp(1) 
    Set MS2Spec = comp(2) 
 
'***** Determine the dissociation type (based on whether or not component Name contains the string "ETD"). 
    If InStr(comp.Name,"ETD") Then  
        MSMSType = "ETD" 
    Else 
        MSMSType = "CID" 
    End If 
 
'***** Determine the precursor m/z and mass, since we will need that to determine its charge 
    If IsNumeric(comp.Precursor) Then 
       PrecursorMZ = comp.Precursor 
    Else 
       '***** Otherwise grab the mass from the string naming this component 
       StartPos = InStr(comp.Name,"(") + 1 
       StopPos = InStr(comp.Name,")") 
       PrecursorMZ = Mid(comp.Name,StartPos,StopPos-StartPos) 
    End If 

    If IsNumeric(comp.RetentionTime) Then
       PrecursorRT = comp.RetentionTime
    Else
       PrecursorRT = 0
    End If 

    If IsNumeric(comp.Intensity) Then
       PrecursorIntensity = comp.Intensity
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
       NoZCount = NoZCount + 1
       PrecursorCharge = 1
       PrecursorMass = PrecursorMZ - Proton 
    Else 
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
 
'***** Now write this compound header to the msAlign file 
    objFile.WriteLine("BEGIN IONS")  
    objFile.WriteLine("ID=" & comp.CompoundNumber)
    objFile.WriteLine("FRACTION_ID=1")
    objFile.WriteLine("FILE_NAME=" & Analysis.Name)
    objFile.WriteLine("SCANS=" & MS2ScanNumber)
    objFile.WriteLine("RETENTION_TIME=" & PrecursorRT)
    ' Naturally the following will require update when something beyond MS/MS is being used.
    objFile.WriteLine("LEVEL=2")
    objFile.WriteLine("ACTIVATION=" & MSMSType)
    objFile.WriteLine("MS_ONE_ID=" & comp.CompoundNumber)
    objFile.WriteLine("MS_ONE_SCAN=" & MS1ScanNumber)
    objFile.WriteLine("PRECURSOR_MZ=" & PrecursorMZ)  
    objFile.WriteLine("PRECURSOR_CHARGE=" & PrecursorCharge) 
    objFile.WriteLine("PRECURSOR_MASS=" & PrecursorMass)
    ' TopFD writes floating point precursor intensities, but Bruker gives us integers
    objFile.WriteLine("PRECURSOR_INTENSITY=" & PrecursorIntensity & ".00")

'***** Now write the compound MS/MS peaks to the msAlign file 
    For Each ThisPeak in MS2Spec.MSPeakList 
       objFile.Write(Round(ThisPeak.ChargeState * (ThisPeak.m_over_z - Proton),4))
       objFile.Write(" ") 
       objFile.Write(Round(ThisPeak.Intensity * IntensityMultiplier,2)) 
       objFile.Write(" ") 
       objFile.WriteLine(ThisPeak.ChargeState) 
    Next 
 
    objFile.WriteLine("END IONS" & vbCrLf)  
      
Next  
  
objFile.Close  
Analysis.Save
MsgBox (CompoundCount & " compounds, " & NoZCount & " without precursor charges!")
Form.Close  
