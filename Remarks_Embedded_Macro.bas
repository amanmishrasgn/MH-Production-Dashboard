Attribute VB_Name = "Remarks_Display"
Option Explicit

Private Const msoShapeRoundedRectangle As Long = 5
Private Const msoTrue As Long = -1
Private Const msoFalse As Long = 0
Private Const msoBringToFront As Long = 0
Private Const msoAnchorTop As Long = 1
Private Const ppMouseClick As Long = 1
Private Const ppActionRunMacro As Long = 8
Private Const ppAlignLeft As Long = 1

' PUBLIC CALLABLE MACROS (these show in Alt+F8)

Public Sub SetupAllSlides()
    MsgBox "Date buttons are already set up!" & vbCrLf & vbCrLf & _
           "Just click any date button on the slides during slideshow (F5)" & vbCrLf & _
           "Remarks will appear at the bottom-left of the chart.", _
           vbInformation, "Setup Complete"
End Sub

Public Sub SetupInteractiveSlide()
    MsgBox "Date buttons are already set up!" & vbCrLf & vbCrLf & _
           "Just click any date button on this slide during slideshow (F5)" & vbCrLf & _
           "Remarks will appear at the bottom-left of the chart.", _
           vbInformation, "Setup Complete"
End Sub

' DATA STORAGE

Private Function GetRemarksData() As String
    Dim data As String
    data = ""
    data = data & "01-Oct-2026|Oil|MHN|NA#4yH: SSSV control line leakage.Wireline cargo picked up at NA Plat...|-158.0|BOPD§"
    data = data & "01-Oct-2026|Oil|NQO|Separator V-1220 taken offline for replacement of ILT lower tapping...|-60.0|BOPD§"
    data = data & "01-Oct-2026|Oil|NQO|NX#6: clsd for EMG POOH by wireline team for 1 Hr|-10.0|BOPD§"
    data = data & "01-Oct-2026|Oil|BHS|SP: Plat SD occurred low pneumatics in SDP|-150.0|BOPD§"
    data = data & "01-Oct-2026|Oil|BHS|12"" SE-MSP WFL leakage rectification job by MSV SEAMEC AGASTYA: Vessel...|-450.0|BOPD§"
    data = data & "01-Oct-2026|Oil|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-250.0|BOPD§"
    data = data & "01-Oct-2026|Oil|WO-16|WO-16A#3/#5/#9 , WO-5 #1 & #3 & B-121A#1 clsd for Solvent treatment job...|-1773.0|BOPD§"
    data = data & "01-Oct-2026|Oil|ICP|RS12#12 (O): flowing with less temperature, Observed annulus-tubing...|-104.0|BOPD§"
    data = data & "01-Oct-2026|Oil|ICP|ICD#9: 2.25"" held up @4920 FT (MMG-2@3763 ft and MMG-3@5768 FT) jarred...|-15.0|BOPD§"
    data = data & "01-Oct-2026|Gas|MHN|NA#4yH: SSSV control line leakage.Wireline cargo picked up at NA Plat...|-0.02|MMSCMD§"
    data = data & "01-Oct-2026|Gas|BHS|SP: Plat SD occurred low pneumatics in SDP|-0.01|MMSCMD§"
    data = data & "01-Oct-2026|Gas|BHS|12"" SE-MSP WFL leakage rectification job by MSV SEAMEC AGASTYA: Vessel...|-0.03|MMSCMD§"
    data = data & "01-Oct-2026|Gas|WO-16|WO-16A#3/#5/#9 , WO-5 #1 & #3 & B-121A#1 clsd for Solvent treatment job...|-0.201|MMSCMD§"
    data = data & "02-Oct-2026|Oil|MHN|NA#4yH S/choke installation: was clsd SSSV control line leakage|-70.0|BOPD§"
    data = data & "02-Oct-2026|Oil|MHN|N19#4H: Acid job carried out by S/Nidhi|-35.0|BOPD§"
    data = data & "02-Oct-2026|Oil|NQO|N14: 14"" N14-NQO W/F line sterilization job done by GS Roopa, #8 clsd...|-6.0|BOPD§"
    data = data & "02-Oct-2026|Oil|NQO|LD#8Y: found clsd from SSSV, opened|-15.0|BOPD§"
    data = data & "02-Oct-2026|Oil|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-250.0|BOPD§"
    data = data & "02-Oct-2026|Oil|WO-16|WO-16A#3/#5/#9 , WO-5 #1 & #3 & B-121A#1 clsd for Solvent treatment job...|-40.0|BOPD§"
    data = data & "02-Oct-2026|Gas|MHN|NA#4yH S/choke installation: was clsd SSSV control line leakage|-0.02|MMSCMD§"
    data = data & "03-Oct-2026|Oil|MHN|PGC-B: Tripped high LPC DE X&Y axis radial vibration|-25.0|BOPD§"
    data = data & "03-Oct-2026|Oil|NQO|NHA#4: Carried out hole prove|-24.0|BOPD§"
    data = data & "03-Oct-2026|Oil|NQO|Backpressure prod high CIM pressure caused by MHN PGC hot wash,...|-100.0|BOPD§"
    data = data & "03-Oct-2026|Oil|BHS|12"" SE-MSP WFL leakage rectification job by MSV SEAMEC AGASTYA:...|-420.0|BOPD§"
    data = data & "03-Oct-2026|Oil|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-250.0|BOPD§"
    data = data & "03-Oct-2026|Oil|ICP|ICD #11: FGS carried out; #2: Hole probed with 2.65"" GC up to 6023 ft,...|-21.0|BOPD§"
    data = data & "03-Oct-2026|Oil|ICP|RS12#8 (O): Replaced GLV at MMG-4 (ft) and MMG-1 (ft) after multiple...|-150.0|BOPD§"
    data = data & "03-Oct-2026|Gas|MHN|PGC-B: Tripped high LPC DE X&Y axis radial vibration|-0.015|MMSCMD§"
    data = data & "03-Oct-2026|Gas|NQO|Backpressure prod high CIM pressure caused by MHN PGC hot wash,...|-0.05|MMSCMD§"
    data = data & "03-Oct-2026|Gas|NQO|NR#6 (free gas well): Found pin hole leakage from PFA welded flange,...|-0.051|MMSCMD§"
    data = data & "03-Oct-2026|Gas|BHS|12"" SE-MSP WFL leakage rectification job by MSV SEAMEC AGASTYA:...|-0.019|MMSCMD§"
    data = data & "04-Oct-2026|Oil|MHN|NA# 3zH: 2.35""GC run up to storm choke depth (514 ft|-30.0|BOPD§"
    data = data & "04-Oct-2026|Oil|NQO|NHA#4H: P/O EMG & recorded FBHPT gradients from 1375m to surface|-7.0|BOPD§"
    data = data & "04-Oct-2026|Oil|NQO|NQG-PGC-A (6th PGC): Stopped for 6hrs for NQG-PGC-B discharge XSDV...|-150.0|BOPD§"
    data = data & "04-Oct-2026|Oil|BHS|12"" SE-MSP WFL leakage rectification job by MSV SEAMEC AGASTYA: Vessel...|-400.0|BOPD§"
    data = data & "04-Oct-2026|Oil|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-250.0|BOPD§"
    data = data & "04-Oct-2026|Oil|SHP|S1-6#1Z(ABAN-8): Flowback started post DPRAS job by WSV GS Roopa|330.0|BOPD§"
    data = data & "04-Oct-2026|Oil|ICP|ZB #7ZH: DPRAS job carried out by GS Roopa|-22.0|BOPD§"
    data = data & "04-Oct-2026|Gas|NQO|NR#6 (free gas well): Found pin hole leakage from PFA welded flange,...|-0.07|MMSCMD§"
    data = data & "04-Oct-2026|Gas|NQO|NQG-PGC-A (6th PGC): Stopped for 6hrs for NQG-PGC-B discharge XSDV...|-0.198|MMSCMD§"
    data = data & "04-Oct-2026|Gas|BHS|12"" SE-MSP WFL leakage rectification job by MSV SEAMEC AGASTYA: Vessel...|-0.02|MMSCMD§"
    data = data & "04-Oct-2026|Gas|SHP|S1-6#1Z(ABAN-8): Flowback started post DPRAS job by WSV GS Roopa|0.06|MMSCMD§"
    data = data & "05-Oct-2026|Oil|NQO|N17#8H: DPRAS job carried out by GS Roopa|-20.0|BOPD§"
    data = data & "05-Oct-2026|Oil|BHS|12"" SE-MSP WFL leakage rectification job by MSV SEAMEC AGASTYA: Vessel...|-467.0|BOPD§"
    data = data & "05-Oct-2026|Oil|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-250.0|BOPD§"
    data = data & "05-Oct-2026|Oil|BHS|IP#9: clsd for GC run & EMG hang(Wireline job|-25.0|BOPD§"
    data = data & "05-Oct-2026|Oil|BHS|IG#5: Observed GIdP & temperature in decreasing trend since 1900 hrs|-27.0|BOPD§"
    data = data & "05-Oct-2026|Oil|SHP|SHP PGC-C (7th PGC) given ventilated S/D leakage in 2nd stage cooler...|-60.0|BOPD§"
    data = data & "05-Oct-2026|Oil|SHP|SV#5: clsd high W/C as per base instruction|-70.0|BOPD§"
    data = data & "05-Oct-2026|Oil|ICP|RS12#8 (O): CTU Job by WSV Celestial cancelled leakage in bowen...|-150.0|BOPD§"
    data = data & "05-Oct-2026|Oil|ICP|ICD #4: EMG parked Ft|-15.0|BOPD§"
    data = data & "05-Oct-2026|Gas|NQO|NR#6 (free gas well): Found pin hole leakage from PFA welded flange,...|-0.07|MMSCMD§"
    data = data & "05-Oct-2026|Gas|BHS|12"" SE-MSP WFL leakage rectification job by MSV SEAMEC AGASTYA: Vessel...|-0.025|MMSCMD§"
    data = data & "05-Oct-2026|Gas|SHP|SHP PGC-C (7th PGC) given ventilated S/D leakage in 2nd stage cooler...|-0.04|MMSCMD§"
    data = data & "05-Oct-2026|Gas|SHP|SV#5: clsd high W/C as per base instruction|-0.005|MMSCMD§"
    data = data & "06-Oct-2026|Oil|MHN|NA#1xH: RIH 2"" GC cleared up to 62 ft|-15.0|BOPD§"
    data = data & "06-Oct-2026|Oil|NQO|N12: Suspected deoptimization in multiple wells as per SCADA|-100.0|BOPD§"
    data = data & "06-Oct-2026|Oil|NQO|NHA#5H: clsd for EMG park for FGS study|-6.0|BOPD§"
    data = data & "06-Oct-2026|Oil|NQO|LE#9: Gas injection clsd for 1.5 hrs for ACV replacement job|-12.0|BOPD§"
    data = data & "06-Oct-2026|Oil|BHS|12"" SE-MSP WFL leakage rectification job by MSV SEAMEC AGASTYA: Vessel...|-467.0|BOPD§"
    data = data & "06-Oct-2026|Oil|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-250.0|BOPD§"
    data = data & "06-Oct-2026|Oil|BHS|RS-1: Plat SD as per SCADA observation|-185.0|BOPD§"
    data = data & "06-Oct-2026|Oil|BHS|NWIS-R TG tripped turbine enclosure combustible gas level hi-hi alarm...|-110.0|BOPD§"
    data = data & "06-Oct-2026|Oil|ICP|RS12 #8 CTU Job by WSV GS Roopa: RIH CT with technical water, observed...|-180.0|BOPD§"
    data = data & "06-Oct-2026|Gas|NQO|NR#6 (free gas well): Part spool replaced and well opened|-0.023|MMSCMD§"
    data = data & "06-Oct-2026|Gas|BHS|12"" SE-MSP WFL leakage rectification job by MSV SEAMEC AGASTYA: Vessel...|-0.025|MMSCMD§"
    data = data & "06-Oct-2026|Gas|BHS|RS-1: Plat SD as per SCADA observation|-0.03|MMSCMD§"
    data = data & "06-Oct-2026|Gas|BHS|NWIS-R TG tripped turbine enclosure combustible gas level hi-hi alarm...|-0.127|MMSCMD§"
    data = data & "07-Oct-2026|Oil|NQO|#All: High CIM pressure less gas received by MHN|-300.0|BOPD§"
    data = data & "07-Oct-2026|Oil|NQO|10"" LB- NQP WF line pigging: Used 9.75"" & 10"" Scrapper pigs|-45.0|BOPD§"
    data = data & "07-Oct-2026|Oil|BHS|12"" SE-MSP WFL leakage rectification job by MSV SEAMEC AGASTYA: Vessel...|-417.0|BOPD§"
    data = data & "07-Oct-2026|Oil|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less prod receipt at SHP...|-250.0|BOPD§"
    data = data & "07-Oct-2026|Oil|SHP|S1-6#1Z(ABAN-8): After observation of reduced liquid rate, choke was...|154.0|BOPD§"
    data = data & "07-Oct-2026|Oil|SHP|SV#7: Development Well completed in L-III reservoir by Rig Jindal Explorer|93.0|BOPD§"
    data = data & "07-Oct-2026|Oil|SHP|SHP PGC-C (7th PGC) given ventilated S/D , leakage in 2nd stage gas...|-100.0|BOPD§"
    data = data & "07-Oct-2026|Oil|SHP|RS-15: SHMFS team working at RS-15 plat reported leakage at wellhead...|-80.0|BOPD§"
    data = data & "07-Oct-2026|Oil|ICP|RS12 #8: Acid spot job carried out by WSV GS Roopa, Well flowing with...|-150.0|BOPD§"
    data = data & "07-Oct-2026|Gas|BHS|12"" SE-MSP WFL leakage rectification job by MSV SEAMEC AGASTYA: Vessel...|-0.02|MMSCMD§"
    data = data & "07-Oct-2026|Gas|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less prod receipt at SHP...|-0.02|MMSCMD§"
    data = data & "07-Oct-2026|Gas|SHP|S1-6#2(FG): clsd to avoid backpressure at SH 7th PGC unavailability|-0.06|MMSCMD§"
    data = data & "07-Oct-2026|Gas|SHP|SHP PGC-C (7th PGC) given ventilated S/D , leakage in 2nd stage gas...|-0.08|MMSCMD§"
    data = data & "07-Oct-2026|Gas|SHP|RS-15: SHMFS team working at RS-15 plat reported leakage at wellhead...|-0.05|MMSCMD§"
    data = data & "08-Oct-2026|Oil|NQO|N17: Found XSDV 80% clsd at LA end, instrumentation team mobilised...|-122.0|BOPD§"
    data = data & "08-Oct-2026|Oil|NQO|NQO#All: High CIM pressure less gas received by MHN, pig loaded at NQG,...|-100.0|BOPD§"
    data = data & "08-Oct-2026|Oil|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less prod receipt at SHP...|-250.0|BOPD§"
    data = data & "08-Oct-2026|Oil|BHS|RS21 #6: clsd to utilize the available Gas Injection optimally to...|-75.0|BOPD§"
    data = data & "08-Oct-2026|Oil|SHP|RS-15: On arrival , found leakage from #3H PFA line; RS-15#1H/2/3H/8 clsd|-290.0|BOPD§"
    data = data & "08-Oct-2026|Oil|SHP|SV#7: Development Well completed in L-III reservoir by Rig Jindal Explorer|100.0|BOPD§"
    data = data & "08-Oct-2026|Oil|SHP|S1-6#1Z(ST job by ABAN-VIII): Post Nitrified DPRAS job by GS Roopa ,...|147.0|BOPD§"
    data = data & "08-Oct-2026|Gas|NQO|NQO#All: High CIM pressure less gas received by MHN, pig loaded at NQG,...|-0.039|MMSCMD§"
    data = data & "08-Oct-2026|Gas|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less prod receipt at SHP...|-0.02|MMSCMD§"
    data = data & "08-Oct-2026|Gas|BHS|RS21 #6: clsd to utilize the available Gas Injection optimally to...|-0.005|MMSCMD§"
    data = data & "08-Oct-2026|Gas|SHP|RS-15: On arrival , found leakage from #3H PFA line; RS-15#1H/2/3H/8 clsd|-0.145|MMSCMD§"
    data = data & "09-Oct-2026|Oil|MHN|TG-1 & TG-2 tripped simultaneously engine flame-out , resulting in PSD|-7940.0|BOPD§"
    data = data & "09-Oct-2026|Oil|MHN|MHN PSD: All wells of N2-N2A/NV/NJ/ND and N14 sector de-optimised drop...|-1040.0|BOPD§"
    data = data & "09-Oct-2026|Oil|BHS|12"" SE-MSP WFL leakage rectification job by MSV SEAMEC AGASTYA: Vessel...|-130.0|BOPD§"
    data = data & "09-Oct-2026|Oil|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less prod receipt at SHP...|-250.0|BOPD§"
    data = data & "09-Oct-2026|Oil|BHS|BHS RS-18 (All wells except #12(self)): MLGIP Down by 56 KSC PSD at MHN|-320.0|BOPD§"
    data = data & "09-Oct-2026|Oil|SCA|12'' IF-SCA WF sub sea line leakage arrest job by MSV Seamec Paladin:...|-256.0|BOPD§"
    data = data & "09-Oct-2026|Oil|SHP|RS-15: RS-15: Rectified #3H leaking spool at SHW and shifted to RS-15....|-90.0|BOPD§"
    data = data & "09-Oct-2026|Oil|SHP|IM#12H: clsd for EMG hang(Wireline job|-38.0|BOPD§"
    data = data & "09-Oct-2026|Oil|ICP|RS2#7Z (O): Observed leakage from PFA, well clsd|-90.0|BOPD§"
    data = data & "09-Oct-2026|Gas|MHN|TG-1 & TG-2 tripped simultaneously engine flame-out , resulting in PSD|-0.865|MMSCMD§"
    data = data & "09-Oct-2026|Gas|MHN|MHN PSD: All wells of N2-N2A/NV/NJ/ND and N14 sector de-optimised drop...|-0.15|MMSCMD§"
    data = data & "09-Oct-2026|Gas|NQO|Received excess gas from MHN after PSD revival unavailability of MHN...|-0.25|MMSCMD§"
    data = data & "09-Oct-2026|Gas|BHS|12"" SE-MSP WFL leakage rectification job by MSV SEAMEC AGASTYA: Vessel...|-0.02|MMSCMD§"
    data = data & "09-Oct-2026|Gas|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less prod receipt at SHP...|-0.02|MMSCMD§"
    data = data & "09-Oct-2026|Gas|BHS|BHS RS-18 (All wells except #12(self)): MLGIP Down by 56 KSC PSD at MHN|-0.045|MMSCMD§"
    data = data & "09-Oct-2026|Gas|SCA|12'' IF-SCA WF sub sea line leakage arrest job by MSV Seamec Paladin:...|-0.027|MMSCMD§"
    data = data & "09-Oct-2026|Gas|SHP|RS-15: RS-15: Rectified #3H leaking spool at SHW and shifted to RS-15....|-0.025|MMSCMD§"
    data = data & "10-Oct-2026|Oil|MHN|PGC-D (IVth PGC) started and loaded|-90.0|BOPD§"
    data = data & "10-Oct-2026|Oil|NQO|N2-N2A/N14 sector: Observed MLGIP drop by 3 KSC MHN PGC shut down|-120.0|BOPD§"
    data = data & "10-Oct-2026|Oil|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-250.0|BOPD§"
    data = data & "10-Oct-2026|Oil|BHS|RS-1#4: clsd for rocking|-17.0|BOPD§"
    data = data & "10-Oct-2026|Oil|BHS|RS-21#6: Observed well flowing with sub optimal temperature degree...|-55.0|BOPD§"
    data = data & "10-Oct-2026|Oil|SCA|12'' IF-SCA WF sub sea line leakage arrest job by MSV Seamec Paladin:...|-561.0|BOPD§"
    data = data & "10-Oct-2026|Oil|WO-16|WO-16A#5: clsd intermittently to maintain separator level constraint of...|-230.0|BOPD§"
    data = data & "10-Oct-2026|Oil|SHP|IM#12H: clsd for EMG run (Wireline job|-36.0|BOPD§"
    data = data & "10-Oct-2026|Gas|MHN|PGC-D (IVth PGC) started and loaded|-0.3|MMSCMD§"
    data = data & "10-Oct-2026|Gas|NQO|N2-N2A/N14 sector: Observed MLGIP drop by 3 KSC MHN PGC shut down|-0.04|MMSCMD§"
    data = data & "10-Oct-2026|Gas|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-0.02|MMSCMD§"
    data = data & "10-Oct-2026|Gas|SCA|12'' IF-SCA WF sub sea line leakage arrest job by MSV Seamec Paladin:...|-0.037|MMSCMD§"
    data = data & "10-Oct-2026|WI|MNW|MNW: MNW TG-1 and MNP TG-4 running in island mode|-57000.0|BOPD§"
    data = data & "10-Oct-2026|WI|MNW|MNW: MNW TG-1 and MNP TG-4 running in island mode|-57000.0|BOPD§"
    data = data & "11-Oct-2026|Oil|NQO|12"" N24-NQD WF line pigging: Used 12"" scrapper pig without brush|-58.0|BOPD§"
    data = data & "11-Oct-2026|Oil|BHS|RS-21#6: Observed well flowing with sub optimal temperature °C to 32°C|-56.0|BOPD§"
    data = data & "11-Oct-2026|Oil|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-250.0|BOPD§"
    data = data & "11-Oct-2026|Oil|BHS|RS-1: Landing temperature down by 3 deg C|-100.0|BOPD§"
    data = data & "11-Oct-2026|Oil|SCA|12'' IF-SCA WF sub sea line leakage arrest job by MSV Seamec Paladin:...|-475.0|BOPD§"
    data = data & "11-Oct-2026|Oil|SHP|18"" RS-15 WFL PIGGING: #12/2/6H/5H clsd for PIGGING operation under the...|-115.0|BOPD§"
    data = data & "11-Oct-2026|Oil|SHP|SV#7: clsd for NDPRAS+XLD Stimulation job by WSV GS ROOPA|-30.0|BOPD§"
    data = data & "11-Oct-2026|Oil|ICP|ICD #5: Acid spotting job carried out by Samudra Nidhi|-20.0|BOPD§"
    data = data & "11-Oct-2026|Oil|CL-7|Cluster 7/8: a) 101th offtake: Tanker Swarna Brahmaputra, dispatch...|-100.0|BOPD§"
    data = data & "11-Oct-2026|Gas|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-0.02|MMSCMD§"
    data = data & "11-Oct-2026|Gas|BHS|RS-1: Landing temperature down by 3 deg C|-0.01|MMSCMD§"
    data = data & "11-Oct-2026|Gas|SCA|12'' IF-SCA WF sub sea line leakage arrest job by MSV Seamec Paladin:...|-0.029|MMSCMD§"
    data = data & "11-Oct-2026|Gas|SHP|18"" RS-15 WFL PIGGING: #12/2/6H/5H clsd for PIGGING operation under the...|-0.025|MMSCMD§"
    data = data & "11-Oct-2026|WI|MNW|MNW: PGC-D (4TH PGC) tripped EBV338 drive module fault|0|BOPD§"
    data = data & "11-Oct-2026|WI|MNW|MNW: MIP tripped PSD|-57000.0|BOPD§"
    data = data & "11-Oct-2026|WI|MNW|MNW: MIP tripped PSD|-57000.0|BOPD§"
    data = data & "12-Oct-2026|Oil|MHN|Fire incidence in NQG PGC cooler area|-3750.0|BOPD§"
    data = data & "12-Oct-2026|Oil|NQO|FSD occurred rupture and fire in NQG-PGC-C 3rd stage cooler discharge...|-31390.0|BOPD§"
    data = data & "12-Oct-2026|Oil|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-250.0|BOPD§"
    data = data & "12-Oct-2026|Gas|MHN|Fire incidence in NQG PGC cooler area|-0.61|MMSCMD§"
    data = data & "12-Oct-2026|Gas|NQO|FSD occurred rupture and fire in NQG-PGC-C 3rd stage cooler discharge...|-2.64|MMSCMD§"
    data = data & "12-Oct-2026|Gas|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-0.02|MMSCMD§"
    data = data & "12-Oct-2026|WI|MNW|MNW: MIP tripped PSD|-57000.0|BOPD§"
    data = data & "12-Oct-2026|WI|MNW|MNW: MIP tripped PSD|-57000.0|BOPD§"
    data = data & "13-Oct-2026|Oil|MHN|NQO revival is in progress|-2500.0|BOPD§"
    data = data & "13-Oct-2026|Oil|NQO|NQO Fire Shut down: Revival is in progress. 145 (out of 246) wells opened|-12800.0|BOPD§"
    data = data & "13-Oct-2026|Oil|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-250.0|BOPD§"
    data = data & "13-Oct-2026|Oil|BHS|RS-18#3, #7, 10: found flowing with suboptimal temperature, after...|-85.0|BOPD§"
    data = data & "13-Oct-2026|Oil|WO-16|WO-16A#5: clsd intermittently to maintain separator level constraint of...|-300.0|BOPD§"
    data = data & "13-Oct-2026|Oil|ICP|ST#9Z (O): Observed increase in W/C from 55% to 63%|-96.0|BOPD§"
    data = data & "13-Oct-2026|Gas|MHN|NQO revival is in progress|-0.2|MMSCMD§"
    data = data & "13-Oct-2026|Gas|NQO|NQO Fire Shut down: Revival is in progress. 145 (out of 246) wells opened|-1.447|MMSCMD§"
    data = data & "13-Oct-2026|Gas|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-0.02|MMSCMD§"
    data = data & "13-Oct-2026|Gas|BHS|RS-18#3, #7, 10: found flowing with suboptimal temperature, after...|-0.02|MMSCMD§"
    data = data & "13-Oct-2026|Gas|WO-16|WO-16A#5: clsd intermittently to maintain separator level constraint of...|-0.015|MMSCMD§"
    data = data & "14-Oct-2026|Oil|MHN|N11, NS lift gas catered by NQO is not available post NQO FSD|-1300.0|BOPD§"
    data = data & "14-Oct-2026|Oil|NQO|NQO Fire Shut down: Revival is in progress. 161 (out of 244) wells...|-9100.0|BOPD§"
    data = data & "14-Oct-2026|Oil|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-250.0|BOPD§"
    data = data & "14-Oct-2026|Oil|WO-16|WO-16A#5, WO-5A#3 Well clsd intermittently to maintain separator level...|-268.0|BOPD§"
    data = data & "14-Oct-2026|Oil|ICP|ICD #8 (O): Acid spot job carried out by WSV Celestial|-25.0|BOPD§"
    data = data & "14-Oct-2026|Gas|MHN|N11, NS lift gas catered by NQO is not available post NQO FSD|-0.13|MMSCMD§"
    data = data & "14-Oct-2026|Gas|NQO|NQO Fire Shut down: Revival is in progress. 161 (out of 244) wells...|-1.248|MMSCMD§"
    data = data & "14-Oct-2026|Gas|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-0.02|MMSCMD§"
    data = data & "14-Oct-2026|Gas|WO-16|WO-16A#5, WO-5A#3 Well clsd intermittently to maintain separator level...|-0.011|MMSCMD§"
    data = data & "15-Oct-2026|Oil|MHN|N11#12 was opened after gas lift normalization from NQO; however, the...|-1076.0|BOPD§"
    data = data & "15-Oct-2026|Oil|NQO|NQO Fire Shut down: Revival is in progress. 161 (out of 244) wells...|-7700.0|BOPD§"
    data = data & "15-Oct-2026|Oil|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-250.0|BOPD§"
    data = data & "15-Oct-2026|Oil|WO-16|WO-16A#5, WO-5A#3 Well clsd intermittently to maintain separator level...|-288.0|BOPD§"
    data = data & "15-Oct-2026|Gas|MHN|N11#12 was opened after gas lift normalization from NQO; however, the...|-0.09|MMSCMD§"
    data = data & "15-Oct-2026|Gas|NQO|NQO Fire Shut down: Revival is in progress. 161 (out of 244) wells...|-1.077|MMSCMD§"
    data = data & "15-Oct-2026|Gas|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-0.02|MMSCMD§"
    data = data & "15-Oct-2026|Gas|WO-16|WO-16A#5, WO-5A#3 Well clsd intermittently to maintain separator level...|-0.012|MMSCMD§"
    data = data & "15-Oct-2026|WI|MNW|MNW: MIP tripped PSD|-57000.0|BOPD§"
    data = data & "15-Oct-2026|WI|MNW|MNW: MIP tripped PSD|-57000.0|BOPD§"
    data = data & "16-Oct-2026|Oil|MHN|N11#12 was opened yesterday after gas lift normalisation from NQO;...|-1005.0|BOPD§"
    data = data & "16-Oct-2026|Oil|NQO|NQO Fire Shut down: Revival is in progress. 165 (out of 244) wells opened|-7416.0|BOPD§"
    data = data & "16-Oct-2026|Oil|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-250.0|BOPD§"
    data = data & "16-Oct-2026|Oil|BHS|SF#7SS: clsd for GC run|-10.0|BOPD§"
    data = data & "16-Oct-2026|Oil|BHS|BHS PGC-A & B: non-availability of both units, RS-16#1,9,11,SA#9SS clsd...|-325.0|BOPD§"
    data = data & "16-Oct-2026|Oil|WO-16|WO-16A#5, WO-5A#3 Well clsd intermittently to maintain separator level...|-336.0|BOPD§"
    data = data & "16-Oct-2026|Oil|SHP|Carried out well testing of SV#7 after stimulation job|60.0|BOPD§"
    data = data & "16-Oct-2026|Oil|ICP|RS2 #7: PFA replaced|-25.0|BOPD§"
    data = data & "16-Oct-2026|Gas|MHN|N11#12 was opened yesterday after gas lift normalisation from NQO;...|-0.09|MMSCMD§"
    data = data & "16-Oct-2026|Gas|NQO|NQO Fire Shut down: Revival is in progress. 165 (out of 244) wells opened|-1.103|MMSCMD§"
    data = data & "16-Oct-2026|Gas|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-0.02|MMSCMD§"
    data = data & "16-Oct-2026|Gas|BHS|BHS PGC-A & B: non-availability of both units, RS-16#1,9,11,SA#9SS clsd...|-1.124|MMSCMD§"
    data = data & "16-Oct-2026|Gas|SHP|Carried out well testing of SV#7 after stimulation job|0.018|MMSCMD§"
    data = data & "16-Oct-2026|Gas|WO-16|WO-16A#5, WO-5A#3 Well clsd intermittently to maintain separator level...|-0.014|MMSCMD§"
    data = data & "16-Oct-2026|WI|MNW|MNW: MIP tripped PSD|-57000.0|BOPD§"
    data = data & "16-Oct-2026|WI|ICW|ICW MIP-1650: Pump tripped heavy seal leakage in DE side|-27021.0|BOPD§"
    data = data & "16-Oct-2026|WI|MNW|MNW: MIP tripped PSD|-57000.0|BOPD§"
    data = data & "16-Oct-2026|WI|ICW|ICW MIP-1650: Pump tripped heavy seal leakage in DE side|-27021.0|BOPD§"
    data = data & "17-Oct-2026|Oil|MHN|N11#12 was opened after gas lift normalization from NQO; however, the...|-1005.0|BOPD§"
    data = data & "17-Oct-2026|Oil|NQO|NQO Fire Shut down: Revival is in progress. 160 (out of 244) wells...|-7704.0|BOPD§"
    data = data & "17-Oct-2026|Oil|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-250.0|BOPD§"
    data = data & "17-Oct-2026|Oil|BHS|SF#7SS: clsd for wireline job|-10.0|BOPD§"
    data = data & "17-Oct-2026|Oil|BHS|BHS PGC-A & B: non availability of both units, RS-16#1,9,11,SA#9SS clsd...|-205.0|BOPD§"
    data = data & "17-Oct-2026|Oil|WO-16|WO-16A#5: clsd intermittently to maintain separator level constraint of...|-386.0|BOPD§"
    data = data & "17-Oct-2026|Oil|ICP|ICD #9: Acid spot job carried out by WSV AM Passion|-36.0|BOPD§"
    data = data & "17-Oct-2026|Oil|ICP|ICD #4: Acid spot job carried out by WSV AM Passion|-23.0|BOPD§"
    data = data & "17-Oct-2026|Gas|MHN|N11#12 was opened after gas lift normalization from NQO; however, the...|-0.09|MMSCMD§"
    data = data & "17-Oct-2026|Gas|NQO|NQO Fire Shut down: Revival is in progress. 160 (out of 244) wells...|-1.183|MMSCMD§"
    data = data & "17-Oct-2026|Gas|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-0.02|MMSCMD§"
    data = data & "17-Oct-2026|Gas|BHS|BHS PGC-A & B: non availability of both units, RS-16#1,9,11,SA#9SS clsd...|-0.638|MMSCMD§"
    data = data & "17-Oct-2026|Gas|WO-16|WO-16A#5: clsd intermittently to maintain separator level constraint of...|-0.014|MMSCMD§"
    data = data & "17-Oct-2026|Gas|WO-16|PGC-A tripped ; WO-16A#2/6/7 clsd|-0.075|MMSCMD§"
    data = data & "17-Oct-2026|WI|MNW|MNW: MIP tripped PSD|-57000.0|BOPD§"
    data = data & "17-Oct-2026|WI|ICW|ICW: MIP-1650: Pump tripped heavy seal leakage in DE side|-55853.0|BOPD§"
    data = data & "17-Oct-2026|WI|MNW|MNW: MIP tripped PSD|-57000.0|BOPD§"
    data = data & "17-Oct-2026|WI|ICW|ICW: MIP-1650: Pump tripped heavy seal leakage in DE side|-55853.0|BOPD§"
    data = data & "18-Oct-2026|Oil|MHN|N11#12 was opened after gas lift normalization from NQO; however, the...|-1005.0|BOPD§"
    data = data & "18-Oct-2026|Oil|MHN|NB#4xH: Storm choke servicing / replacement job in progress|-28.0|BOPD§"
    data = data & "18-Oct-2026|Oil|NQO|NQO Fire Shut down: Revival is in progress. 159 (out of 244 strings)...|-7343.0|BOPD§"
    data = data & "18-Oct-2026|Oil|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-250.0|BOPD§"
    data = data & "18-Oct-2026|Oil|BHS|SF#9: clsd for wireline job|-32.0|BOPD§"
    data = data & "18-Oct-2026|Oil|WO-16|WO-16A#5: clsd intermittently to maintain separator level constraint of...|-445.0|BOPD§"
    data = data & "18-Oct-2026|Oil|SHP|SHG PGC-E(7TH PGC) tripped GG centre vibration high|-30.0|BOPD§"
    data = data & "18-Oct-2026|Gas|MHN|N11#12 was opened after gas lift normalization from NQO; however, the...|-0.09|MMSCMD§"
    data = data & "18-Oct-2026|Gas|NQO|NQO Fire Shut down: Revival is in progress. 159 (out of 244 strings)...|-1.273|MMSCMD§"
    data = data & "18-Oct-2026|Gas|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-0.02|MMSCMD§"
    data = data & "18-Oct-2026|Gas|SHP|SHG PGC-E(7TH PGC) tripped GG centre vibration high|-0.015|MMSCMD§"
    data = data & "18-Oct-2026|Gas|WO-16|WO-16A#5: clsd intermittently to maintain separator level constraint of...|-0.012|MMSCMD§"
    data = data & "18-Oct-2026|WI|MNW|MNW: MIP tripped PSD|-57000.0|BOPD§"
    data = data & "18-Oct-2026|WI|ICW|ICW: MIP-1650: Pump tripped heavy seal leakage in DE side|-55853.0|BOPD§"
    data = data & "18-Oct-2026|WI|MNW|MNW: MIP tripped PSD|-57000.0|BOPD§"
    data = data & "18-Oct-2026|WI|ICW|ICW: MIP-1650: Pump tripped heavy seal leakage in DE side|-55853.0|BOPD§"
    data = data & "19-Oct-2026|Oil|MHN|N11#12 was opened after gas lift normalization from NQO; however, the...|-1005.0|BOPD§"
    data = data & "19-Oct-2026|Oil|MHN|NA & N9 deoptimised NQO PGC-C tripping, GI resumed from NQO side|-929.0|BOPD§"
    data = data & "19-Oct-2026|Oil|MHN|NB # 4xH: handover to FORE SIGHT DRILLER IX RIG for solvent Job and...|-140.0|BOPD§"
    data = data & "19-Oct-2026|Oil|MHN|NB#3XH SS: Diverted to MOL post WOR by FORE SIGHT DRILLER IX RIG|99.0|BOPD§"
    data = data & "19-Oct-2026|Oil|MHN|N21#9H: clsd for PLO tool installation , unsuccessful|-21.0|BOPD§"
    data = data & "19-Oct-2026|Oil|NQO|NQO Fire Shut down: Revival is in progress. 159 (out of 244 strings) opened|-7343.0|BOPD§"
    data = data & "19-Oct-2026|Oil|NQO|NQP-PGC-C (3rd PGC) tripped GG lube oil supply pressure low S/D...|-2800.0|BOPD§"
    data = data & "19-Oct-2026|Oil|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-250.0|BOPD§"
    data = data & "19-Oct-2026|Oil|BHS|SF#9: clsd for wireline job|-20.0|BOPD§"
    data = data & "19-Oct-2026|Oil|BHS|RS-1#4: found flowing with suboptimal temperature|-40.0|BOPD§"
    data = data & "19-Oct-2026|Oil|WO-16|B-119/B121 plat shut down since 0800 hrs SDP primary pump failure|-325.0|BOPD§"
    data = data & "19-Oct-2026|Oil|WO-16|WO-16A#5: clsd intermittently to maintain separator level constraint of...|-110.0|BOPD§"
    data = data & "19-Oct-2026|Gas|MHN|N11#12 was opened after gas lift normalization from NQO; however, the...|-0.09|MMSCMD§"
    data = data & "19-Oct-2026|Gas|MHN|NA & N9 deoptimised NQO PGC-C tripping, GI resumed from NQO side|-0.114|MMSCMD§"
    data = data & "19-Oct-2026|Gas|NQO|NQO Fire Shut down: Revival is in progress. 159 (out of 244 strings) opened|-1.273|MMSCMD§"
    data = data & "19-Oct-2026|Gas|NQO|NQP-PGC-C (3rd PGC) tripped GG lube oil supply pressure low S/D...|-0.1|MMSCMD§"
    data = data & "19-Oct-2026|Gas|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-0.02|MMSCMD§"
    data = data & "19-Oct-2026|Gas|WO-16|B-119/B121 plat shut down since 0800 hrs SDP primary pump failure|-0.3|MMSCMD§"
    data = data & "19-Oct-2026|Gas|WO-16|WO-16A#5: clsd intermittently to maintain separator level constraint of...|-0.012|MMSCMD§"
    data = data & "19-Oct-2026|WI|MNW|MNW: MIP tripped PSD|-57000.0|BOPD§"
    data = data & "19-Oct-2026|WI|ICW|ICW: MIP-1650: Pump tripped heavy seal leakage in DE side|-31416.0|BOPD§"
    data = data & "19-Oct-2026|WI|MNW|MNW: MIP tripped PSD|-57000.0|BOPD§"
    data = data & "19-Oct-2026|WI|ICW|ICW: MIP-1650: Pump tripped heavy seal leakage in DE side|-31416.0|BOPD§"
    data = data & "20-Oct-2026|Oil|MHN|N11#12 was opened after gas lift normalization from NQO; however, the...|-1005.0|BOPD§"
    data = data & "20-Oct-2026|Oil|MHN|NB # 4xH / FD IX RIG: Storm choke installed|-198.0|BOPD§"
    data = data & "20-Oct-2026|Oil|MHN|N21#9H: Found well not flowing|-62.0|BOPD§"
    data = data & "20-Oct-2026|Oil|NQO|N10#: Wells found deoptimized|-45.0|BOPD§"
    data = data & "20-Oct-2026|Oil|NQO|LA#: Wells found deoptimized|-40.0|BOPD§"
    data = data & "20-Oct-2026|Oil|NQO|NQO Fire Shut down: Revival is in progress. 159 (out of 244 strings) opened|-7455.0|BOPD§"
    data = data & "20-Oct-2026|Oil|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-250.0|BOPD§"
    data = data & "20-Oct-2026|Oil|BHS|SP#9: flowing with sub-optimal parameter|-20.0|BOPD§"
    data = data & "20-Oct-2026|Oil|WO-16|WO-16A#5: clsd intermittently to maintain separator level constraint of...|-424.0|BOPD§"
    data = data & "20-Oct-2026|Oil|SHP|SHP PGC-C(7th PGC) tripped GG lube oil level low|-20.0|BOPD§"
    data = data & "20-Oct-2026|Oil|SHP|II#7: clsd for storm choke servicing|-30.0|BOPD§"
    data = data & "20-Oct-2026|Gas|MHN|N11#12 was opened after gas lift normalization from NQO; however, the...|-0.09|MMSCMD§"
    data = data & "20-Oct-2026|Gas|NQO|NQO Fire Shut down: Revival is in progress. 159 (out of 244 strings) opened|-1.323|MMSCMD§"
    data = data & "20-Oct-2026|Gas|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-0.02|MMSCMD§"
    data = data & "20-Oct-2026|Gas|SHP|Post pigging of 8"" B-157-B-127 WFL, observed gain in gas prod of 0.050...|0.05|MMSCMD§"
    data = data & "20-Oct-2026|Gas|SHP|SHP PGC-C(7th PGC) tripped GG lube oil level low|-0.01|MMSCMD§"
    data = data & "20-Oct-2026|Gas|WO-16|WO-16A#5: clsd intermittently to maintain separator level constraint of...|-0.018|MMSCMD§"
    data = data & "20-Oct-2026|WI|MNW|MNW: MIP tripped PSD|-57000.0|BOPD§"
    data = data & "20-Oct-2026|WI|MNW|MNW: MIP tripped PSD|-57000.0|BOPD§"
    data = data & "21-Oct-2026|Oil|MHN|N11#12 was opened after gas lift normalization from NQO; however, the...|-1005.0|BOPD§"
    data = data & "21-Oct-2026|Oil|MHN|NB # 4xH / FD IX RIG: A section valve replaced by Rig FD-9 after...|-113.0|BOPD§"
    data = data & "21-Oct-2026|Oil|MHN|N21#9H: Retrieved 3 ½”extended storm choke with 3 1/2"" slip lock ft|-34.0|BOPD§"
    data = data & "21-Oct-2026|Oil|NQO|NQO Fire Shut down: Revival is in progress. 156 (out of 244 strings)...|-7919.0|BOPD§"
    data = data & "21-Oct-2026|Oil|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-250.0|BOPD§"
    data = data & "21-Oct-2026|Oil|WO-16|WO-16A#5: clsd intermittently to maintain separator level constraint of...|-440.0|BOPD§"
    data = data & "21-Oct-2026|Gas|MHN|N11#12 was opened after gas lift normalization from NQO; however, the...|-0.09|MMSCMD§"
    data = data & "21-Oct-2026|Gas|NQO|NQO Fire Shut down: Revival is in progress. 156 (out of 244 strings)...|-1.453|MMSCMD§"
    data = data & "21-Oct-2026|Gas|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-0.02|MMSCMD§"
    data = data & "21-Oct-2026|Gas|WO-16|WO-16A#5: clsd intermittently to maintain separator level constraint of...|-0.019|MMSCMD§"
    data = data & "21-Oct-2026|WI|MNW|MNW: MIP tripped PSD|-57000.0|BOPD§"
    data = data & "21-Oct-2026|WI|MNW|MNW: MIP tripped PSD|-57000.0|BOPD§"
    data = data & "22-Oct-2026|Oil|MHN|N11#12 was opened after gas lift normalization from NQO; however, the...|-1005.0|BOPD§"
    data = data & "22-Oct-2026|Oil|NQO|NQO Fire Shut down: Revival is in progress. 155 (out of 244 strings)...|-7785.0|BOPD§"
    data = data & "22-Oct-2026|Oil|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-250.0|BOPD§"
    data = data & "22-Oct-2026|Oil|BHS|SI#5,8,9: Found flowing with suboptimal parameters, optimized|-20.0|BOPD§"
    data = data & "22-Oct-2026|Oil|BHS|SG#9: Found flowing cold, optimized|-40.0|BOPD§"
    data = data & "22-Oct-2026|Oil|WO-16|WO-16A#5: clsd intermittently to maintain separator level constraint of...|-451.0|BOPD§"
    data = data & "22-Oct-2026|Oil|SHP|SS#6SS: Post pigging, leakage was observed from the test header clamp...|-100.0|BOPD§"
    data = data & "22-Oct-2026|Oil|SHP|S1-4#2Y: clsd for Well stimulation (Nitrified DPRAS with XLD) job|100.0|BOPD§"
    data = data & "22-Oct-2026|Gas|MHN|N11#12 was opened after gas lift normalization from NQO; however, the...|-0.09|MMSCMD§"
    data = data & "22-Oct-2026|Gas|NQO|NQO Fire Shut down: Revival is in progress. 155 (out of 244 strings)...|-1.453|MMSCMD§"
    data = data & "22-Oct-2026|Gas|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-0.02|MMSCMD§"
    data = data & "22-Oct-2026|Gas|BHS|SI#5,8,9: Found flowing with suboptimal parameters, optimized|-0.01|MMSCMD§"
    data = data & "22-Oct-2026|Gas|BHS|SG#9: Found flowing cold, optimized|-0.01|MMSCMD§"
    data = data & "22-Oct-2026|Gas|SHP|SS#6SS: Post pigging, leakage was observed from the test header clamp...|-0.014|MMSCMD§"
    data = data & "22-Oct-2026|Gas|SHP|S1-4#2Y: clsd for Well stimulation (Nitrified DPRAS with XLD) job|0.08|MMSCMD§"
    data = data & "22-Oct-2026|Gas|WO-16|WO-16A#5: clsd intermittently to maintain separator level constraint of...|-0.019|MMSCMD§"
    data = data & "22-Oct-2026|WI|MNW|MNW: After MNW TG-1 & TG-2 synchronization; Water injection resumed|24500.0|BOPD§"
    data = data & "22-Oct-2026|WI|MNW|MNW: After MNW TG-1 & TG-2 synchronization; Water injection resumed|24500.0|BOPD§"
    data = data & "23-Oct-2026|Oil|MHN|N11#12 was opened after gas lift normalization from NQO; however, the...|-1005.0|BOPD§"
    data = data & "23-Oct-2026|Oil|NQO|N22#1: Nitrified DPRAS + XLD Job carried out by WSV Celestial, Well...|-73.0|BOPD§"
    data = data & "23-Oct-2026|Oil|NQO|NQO Fire Shut down: Revival is in progress. 155 (out of 244 strings)...|-7846.0|BOPD§"
    data = data & "23-Oct-2026|Oil|NQO|N22#5Z Well diverted to main line|280.0|BOPD§"
    data = data & "23-Oct-2026|Oil|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-250.0|BOPD§"
    data = data & "23-Oct-2026|Oil|WO-16|WO-16A#5: clsd intermittently to maintain separator level constraint of...|-451.0|BOPD§"
    data = data & "23-Oct-2026|Oil|SHP|S1-4#2Y(SF): Observed liquid gain post stimulation job (Nitrified DPRAS...|249.0|BOPD§"
    data = data & "23-Oct-2026|Oil|ICP|ICD #9 (O): GLV deplugged from MMG-2|-18.0|BOPD§"
    data = data & "23-Oct-2026|Gas|MHN|N11#12 was opened after gas lift normalization from NQO; however, the...|-0.09|MMSCMD§"
    data = data & "23-Oct-2026|Gas|NQO|NQO Fire Shut down: Revival is in progress. 155 (out of 244 strings)...|-1.46|MMSCMD§"
    data = data & "23-Oct-2026|Gas|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-0.02|MMSCMD§"
    data = data & "23-Oct-2026|Gas|SHP|S1-4#2Y(SF): Observed liquid gain post stimulation job (Nitrified DPRAS...|0.107|MMSCMD§"
    data = data & "23-Oct-2026|Gas|WO-16|WO-16A#5: clsd intermittently to maintain separator level constraint of...|-0.019|MMSCMD§"
    data = data & "24-Oct-2026|Oil|MHN|N11#12 was opened after gas lift normalization from NQO; however, the...|-1005.0|BOPD§"
    data = data & "24-Oct-2026|Oil|MHN|N9 #1zH: found de-optimised; multiple revival attempts were carried out|-253.0|BOPD§"
    data = data & "24-Oct-2026|Oil|NQO|N22#1: Observed flash gain of 800 bbl after DPRAS job|800.0|BOPD§"
    data = data & "24-Oct-2026|Oil|NQO|NQO Fire Shut down: Revival is in progress. 155 (out of 244 strings) opened|-7980.0|BOPD§"
    data = data & "24-Oct-2026|Oil|NQO|N22#5Z Observed flash gain of 800 Bbl after DPRAS job|800.0|BOPD§"
    data = data & "24-Oct-2026|Oil|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-250.0|BOPD§"
    data = data & "24-Oct-2026|Oil|WO-16|WO-16A#5: clsd intermittently to maintain separator level constraint of...|-451.0|BOPD§"
    data = data & "24-Oct-2026|Oil|ICP|ID#8H (O): Acid spot job carried out by WSV AM Passion|-10.0|BOPD§"
    data = data & "24-Oct-2026|Oil|ICP|ICD#9 (O): Tried multiple times with jarring to install GLV in MMG-2...|-50.0|BOPD§"
    data = data & "24-Oct-2026|Gas|MHN|N11#12 was opened after gas lift normalization from NQO; however, the...|-0.09|MMSCMD§"
    data = data & "24-Oct-2026|Gas|MHN|N9 #1zH: found de-optimised; multiple revival attempts were carried out|-0.02|MMSCMD§"
    data = data & "24-Oct-2026|Gas|NQO|NQO Fire Shut down: Revival is in progress. 155 (out of 244 strings) opened|-1.471|MMSCMD§"
    data = data & "24-Oct-2026|Gas|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-0.02|MMSCMD§"
    data = data & "24-Oct-2026|Gas|WO-16|WO-16A#5: clsd intermittently to maintain separator level constraint of...|-0.019|MMSCMD§"
    data = data & "25-Oct-2026|Oil|MHN|N11#12 was opened after gas lift normalisation from NQO; however, the...|-1005.0|BOPD§"
    data = data & "25-Oct-2026|Oil|MHN|N9 #1zH: found de-optimised and unable to sustain pressure buildup|-253.0|BOPD§"
    data = data & "25-Oct-2026|Oil|NQO|N22#6H: clsd to reduce flaring|-57.0|BOPD§"
    data = data & "25-Oct-2026|Oil|NQO|NQO Fire Shut down: Revival is in progress. 155 (out of 244 strings) opened|-7980.0|BOPD§"
    data = data & "25-Oct-2026|Oil|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-250.0|BOPD§"
    data = data & "25-Oct-2026|Oil|WO-16|WO-16A#5: clsd intermittently to maintain separator level constraint of...|-535.0|BOPD§"
    data = data & "25-Oct-2026|Oil|SHP|IN: Leakage reported by Discovery-1 on west side of IN jacket|73.0|BOPD§"
    data = data & "25-Oct-2026|Oil|ICP|ICD#9: Tried multiple times with jarring to install GLV in MMG-2 (3773...|-55.0|BOPD§"
    data = data & "25-Oct-2026|Gas|MHN|N11#12 was opened after gas lift normalisation from NQO; however, the...|-0.09|MMSCMD§"
    data = data & "25-Oct-2026|Gas|MHN|N9 #1zH: found de-optimised and unable to sustain pressure buildup|-0.02|MMSCMD§"
    data = data & "25-Oct-2026|Gas|NQO|NQO Fire Shut down: Revival is in progress. 155 (out of 244 strings) opened|-1.471|MMSCMD§"
    data = data & "25-Oct-2026|Gas|BHS|IH: WF cargo flowing via 14"" IH-SH WF since . Less liquid receipt at...|-0.02|MMSCMD§"
    data = data & "25-Oct-2026|Gas|WO-16|WO-16A#5: clsd intermittently to maintain separator level constraint of...|-0.029|MMSCMD§"
    GetRemarksData = data
End Function

' DISPLAY LOGIC

Public Sub ShowDateRemarks(ByVal oShp As Object)
    Dim sld As Object
    Dim dtStr As String
    Dim comm As String
    Dim remarksText As String
    Dim popCard As Object

    On Error Resume Next
    Set sld = oShp.Parent
    dtStr = GetDateFromButton(oShp)
    comm = GetCommodityFromButton(oShp, sld)

    If Len(dtStr) = 0 Then Exit Sub

    remarksText = BuildRemarksText(dtStr, comm)
    Set popCard = GetOrCreatePopupCard(sld)

    With popCard.TextFrame
        .WordWrap = msoTrue
        .TextRange.Text = remarksText
    End With

    Call FormatPopupText(popCard)
    popCard.Visible = msoTrue
    popCard.ZOrder msoBringToFront
    On Error GoTo 0
End Sub

Private Function BuildRemarksText(ByVal targetDate As String, ByVal commodity As String) As String
    Dim allData As String
    Dim normDate As String
    Dim msg As String
    Dim foundAny As Boolean
    Dim records() As String
    Dim i As Long
    Dim fields() As String
    Dim lStr As String
    Dim prefix As String

    normDate = NormalizeDate(targetDate)
    msg = normDate & vbCrLf
    foundAny = False

    allData = GetRemarksData()
    records = Split(allData, "§")

    For i = 0 To UBound(records)
        If Len(Trim(records(i))) > 0 Then
            fields = Split(records(i), "|")

            If UBound(fields) >= 5 Then
                If UCase(fields(0)) = UCase(normDate) And UCase(fields(1)) = UCase(commodity) Then
                    If Len(fields(3)) > 0 Then
                        If CDbl(fields(4)) <> 0 Then
                            If UCase(commodity) = "GAS" Then
                                lStr = " (" & Format(CDbl(fields(4)), "+0.000;-0.000;0") & " " & fields(5) & ")"
                            Else
                                lStr = " (" & Format(CDbl(fields(4)), "+#,##0;-#,##0;0") & " " & fields(5) & ")"
                            End If
                        Else
                            lStr = ""
                        End If

                        If Len(fields(2)) > 0 And UCase(fields(2)) <> "NONE" And UCase(fields(2)) <> "MH TOTAL" Then
                            prefix = fields(2) & ": "
                        Else
                            prefix = ""
                        End If

                        msg = msg & ChrW(8226) & " " & prefix & fields(3) & lStr & vbCrLf
                        foundAny = True
                    End If
                End If
            End If
        End If
    Next i

    If Not foundAny Then
        msg = msg & ChrW(8226) & " No remarks recorded for this date" & vbCrLf
    End If

    BuildRemarksText = msg
End Function

Private Function GetDateFromButton(ByVal oShp As Object) As String
    Dim alt As String
    Dim parts() As String
    On Error Resume Next
    alt = oShp.AlternativeText
    If InStr(alt, "|") > 0 Then
        parts = Split(alt, "|")
        GetDateFromButton = Trim(parts(0))
    End If
    On Error GoTo 0
End Function

Private Function GetCommodityFromButton(ByVal oShp As Object, ByVal sld As Object) As String
    Dim alt As String
    Dim parts() As String
    On Error Resume Next
    alt = oShp.AlternativeText
    If InStr(alt, "|") > 0 Then
        parts = Split(alt, "|")
        If UBound(parts) >= 1 Then
            GetCommodityFromButton = Trim(parts(1))
            Exit Function
        End If
    End If
    On Error GoTo 0
    GetCommodityFromButton = DetectSlideCommodity(sld)
End Function

Private Function DetectSlideCommodity(ByVal sld As Object) As String
    Dim shp As Object
    Dim i As Long
    Dim txt As String
    On Error Resume Next
    For i = 1 To sld.Shapes.Count
        Set shp = sld.Shapes(i)
        If shp.HasChart Then
            txt = UCase(shp.Chart.ChartTitle.Text)
            If InStr(txt, "GAS") > 0 Then
                DetectSlideCommodity = "Gas"
                Exit Function
            ElseIf InStr(txt, "WATER") > 0 Or InStr(txt, "WI") > 0 Then
                DetectSlideCommodity = "WI"
                Exit Function
            End If
        End If
    Next i
    On Error GoTo 0
    DetectSlideCommodity = "Oil"
End Function

Private Function NormalizeDate(ByVal dateStr As String) As String
    Dim dt As Date
    On Error Resume Next
    dt = CDate(dateStr)
    NormalizeDate = Format(dt, "dd-mmm-yyyy")
    On Error GoTo 0
End Function

Private Function GetOrCreatePopupCard(ByVal sld As Object) As Object
    Dim popCard As Object
    Dim chartShp As Object
    Dim i As Long
    Dim chartL As Single
    Dim chartT As Single
    Dim chartW As Single
    Dim chartH As Single

    On Error Resume Next
    Set popCard = sld.Shapes("Remarks_Popup_Card")
    On Error GoTo 0

    If Not popCard Is Nothing Then
        Set GetOrCreatePopupCard = popCard
        Exit Function
    End If

    On Error Resume Next
    For i = 1 To sld.Shapes.Count
        If sld.Shapes(i).HasChart Then
            Set chartShp = sld.Shapes(i)
            Exit For
        End If
    Next i
    On Error GoTo 0

    If Not chartShp Is Nothing Then
        chartL = chartShp.Left
        chartT = chartShp.Top
        chartW = chartShp.Width
        chartH = chartShp.Height
    Else
        chartL = 100
        chartT = 100
        chartW = 600
        chartH = 400
    End If

    Set popCard = sld.Shapes.AddShape(msoShapeRoundedRectangle, chartL + 12, chartT + chartH - 120, 280, 100)

    With popCard
        .Name = "Remarks_Popup_Card"
        .Adjustments.Item(1) = 0.1
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(240, 240, 235)
        .Line.Color.RGB = RGB(150, 120, 90)
        .Line.Weight = 1.5

        With .TextFrame
            .WordWrap = msoTrue
            .MarginLeft = 6
            .MarginRight = 6
            .MarginTop = 5
            .MarginBottom = 5
            .VerticalAnchor = msoAnchorTop
        End With

        .ActionSettings(ppMouseClick).Action = ppActionRunMacro
        .ActionSettings(ppMouseClick).Run = "HideRemarksPopup"
    End With

    Set GetOrCreatePopupCard = popCard
End Function

Private Sub FormatPopupText(ByVal popCard As Object)
    Dim para As Object
    Dim i As Long
    Dim tr As Object

    On Error Resume Next
    Set tr = popCard.TextFrame.TextRange

    With tr.Font
        .Name = "Calibri"
        .Size = 9
        .Bold = msoFalse
    End With

    For i = 1 To popCard.TextFrame.Paragraphs.Count
        Set para = popCard.TextFrame.Paragraphs(i)
        If i = 1 Then
            para.Font.Bold = msoTrue
            para.Font.Size = 10
        Else
            With para.ParagraphFormat
                .Alignment = ppAlignLeft
                .LineSpacing = 1
                .SpaceBefore = 0
                .SpaceAfter = 0
            End With
        End If
    Next i
    On Error GoTo 0
End Sub

Public Sub HideRemarksPopup(Optional ByVal oShp As Object = Nothing)
    Dim sld As Object
    Dim popCard As Object
    On Error Resume Next
    If oShp Is Nothing Then
        Set sld = ActivePresentation.SlideShowWindow.View.Slide
    Else
        Set sld = oShp.Parent
    End If
    If Not sld Is Nothing Then
        Set popCard = sld.Shapes("Remarks_Popup_Card")
        If Not popCard Is Nothing Then popCard.Visible = msoFalse
    End If
    On Error GoTo 0
End Sub
