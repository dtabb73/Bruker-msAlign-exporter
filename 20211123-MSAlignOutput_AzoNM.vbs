SetLocale(1033) ' Localise for US (. instead , accepted as decimal separator for calculations)    
 
Dim comp, spec, i, oPeak, prec, chargeState, PrecursorMZ, PrecursorMass, RTStart, RTEnd, bUseSNAP 
Dim MSMSSpec, MSMSoutput(), FileName, sExtension, Factor, MSMSType, counter 
Const Proton = 1.00727647 
bUseSNAP = True 
sExtension = ".msalign" 
counter = 1 
Const ETDData = "ETD" 
 
'************************************************************************************************** 
'************************************************************************************************* 
Factor = 10
RTStart = 1
RTEnd = 80
 
'************************************************************************************************* 
'************************************************************************************************ 
 
FileName = GetExportFileName(sExtension) 
Dim objFSO, objFile 
Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set objFile = objFSO.CreateTextFile(FileName) 
 
 
Call SetMethodParametersPeakfinder() 
Analysis.ClearChromatogramRangeSelections 
Analysis.AddChromatogramRangeSelection RTStart, RTEnd, 0, 0 
Analysis.FindAutoMSn 
On Error Resume Next 
      
 
 
 
For Each comp In Analysis.compounds 
    MSMSType = "CID" 
    PrecursorMass = 0 
    chargeState = 0 
    PrecursorMZ = 0 
    chargeState = CInt(GetChargeState(comp)) 
     
     
    If IsNumeric(comp.Precursor) Then 
        PrecursorMZ = Round(comp.precursor, 4) 
    Else 
        PrecursorMZ= GetPrecMassFromName(comp) 
    End If 
     
     
     
    PrecursorMass = Round(((PrecursorMZ*chargeState) - chargeState),4) 
 
 
 
    Set MSMSSpec = comp(2) 
     
    If IsETD(MSMSSpec) Then 
        MSMSType = "ETD" 
    End If 
 
    ReDim MSMSoutput(MSMSSpec.MSPeakList.Count, 2) 
    for i=1 to MSMSSpec.MSPeakList.Count  
      Set oPeak = MSMSSpec.MSPeakList(i) 
      MSMSoutput(i,0) = Round((oPeak.DeconvolutedMolweight - Proton),4) 
      MSMSoutput(i,1) = Round((oPeak.intensity * Factor),2) 
      MSMSoutput(i,2) = oPeak.chargestate 
    Next 
 
 
    objFile.WriteLine("BEGIN IONS") 
    objFile.WriteLine("ID=" & comp.CompoundNumber) 
    objFile.WriteLine("SCANS=" & GetScanNumbers(MSMSSpec)) 
    objFile.WriteLine("ACTIVATION=" & MSMSType) 
    objFile.WriteLine("PRECURSOR_MZ=" & PrecursorMZ) 
    objFile.WriteLine("PRECURSOR_CHARGE=" & chargeState) 
    objFile.WriteLine("PRECURSOR_MASS=" & PrecursorMass) 
 
    Dim count, line 
    For count = 1 To UBound(MSMSOutput) 
        line =  (MSMSoutput(count,0) & " " & MSMSoutput(count,1) & " " & MSMSoutput(count,2)) 
        objFile.WriteLine(line) 
    Next 
    objFile.WriteLine("END IONS" & vbCrLf) 
     
Next 
 
Call ResetPeakfinder() 
objFile.Close 
Analysis.Save 
Form.Close 
 
 
'*********************************************************************** 
Public Function GetChargeState (comp) 
    GetChargeState = 1 
    mChargeState = 1 
    Dim mSpec, mPrec, mPeak, mChargeState, p, y 
    Set mSpec = comp(1) 
    If IsNumeric(comp.Precursor) Then        'get the precursor mass. This checks to make sure the compound has an assigned precursor.  (A bug in DA 4.4 means this can be missing) 
        mPrec = Round(comp.precursor, 4) 
    Else 
        mPrec = GetPrecMassFromName(comp)    'if the precursor mass is missing then extract the information from the compound name 
    End If 
     
    for y=1 to mSpec.MSPeakList.Count         'As the charge state for the precursor is not included search through the MS spectrum to find the precursor mass and get the charge state from there 
          Set mPeak = mSpec.MSPeakList(y) 
           
              If Round(mPeak.m_over_z,4) = mPrec Then 
                  mChargeState = mPeak.ChargeState 
                  Exit For 
              End If 
    Next 
 
    GetChargeState = mChargeState    'return the chrage state (if for some reason the charge state can not be determined (SNAP was not able to deconvolute peak) then return "1" as charge state 
 
End Function 
 
Sub SetMethodParametersPeakfinder()  
    Dim oPeakFinderParameters  
    Set oPeakFinderParameters = Analysis.Method.MassListParameters()  
    If bUseSNAP = True Then  
        oPeakFinderParameters.DetectionAlgorithm = 2 ' 3 means Sum; 2 means SNAP; 0 means APEX  
    Else  
        oPeakFinderParameters.DetectionAlgorithm = 3 ' 3 means Sum; 2 means SNAP; 0 means APEX  
    End If  
End Sub   
 
Sub ResetPeakfinder  
    Dim oPeakFinderParameters  
    Set oPeakFinderParameters = Analysis.Method.MassListParameters()  
    oPeakFinderParameters.DetectionAlgorithm = 3 ' 3 means Sum; 2 means SNAP; 0 means APEX  
End Sub  
 
Function GetExportFileName(sExtension)  
    Dim aFileNames      
    aFileNames = split(Analysis.Name,".d",-1,1)  
      
    GetExportFileName = Analysis.Path & "\" & aFileNames(0) & sExtension  
End Function   
 
Function IsETD(comp) 
    Dim mComp, mCompName 
    IsETD = False 
    mCompName = comp.Name 
    If InStr(mCompName,ETDData) Then 
        IsETD= True 
    End If 
      
End Function 
 
Function GetScanNumbers(Spectrum) 
    Dim position, position2, initialScanPos 
    GetScanNumbers = "" 
    position = InStr(Spectrum.Name,"#") 
    position2 = InStrRev(Spectrum.Name,"-") 
    If position2 >0 Then 
        initialScanPos = position2-(position+1) 
        GetScanNumbers = (Mid(Spectrum.Name, position+1, initialScanPos))+1 
    Else 
        GetScanNumbers = (Mid(Spectrum.Name, position+1))+1 
    End If  
     
End Function 
 
Function GetPrecMassFromName(comp) 
    GetPrecMassFromName = 0 
    Dim compName 
    compName = comp.Name 
    Dim count, str, startPos, endPos 
    startPos = (InStr(compName,"(") + 1) 
    endPos = InStr(compName,")") 
    str = Mid(compName,(startPos),(endPos-startPos)) 
    GetPrecMassFromName = CDbl(str) 
 
End Function 
