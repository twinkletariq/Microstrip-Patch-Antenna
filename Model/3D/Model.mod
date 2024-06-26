'# MWS Version: Version 2018.0 - Oct 26 2017 - ACIS 27.0.2 -

'# length = mm
'# frequency = GHz
'# time = ns
'# frequency range: fmin = 28 fmax = 38
'# created = '[VERSION]2018.0|27.0.2|20171026[/VERSION]


'@ use template: Antenna - Planar.cfg

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
'set the units
With Units
    .Geometry "mm"
    .Frequency "GHz"
    .Voltage "V"
    .Resistance "Ohm"
    .Inductance "H"
    .TemperatureUnit  "Kelvin"
    .Time "ns"
    .Current "A"
    .Conductance "Siemens"
    .Capacitance "F"
End With
'----------------------------------------------------------------------------
Plot.DrawBox True
With Background
     .Type "Normal"
     .Epsilon "1.0"
     .Mu "1.0"
     .XminSpace "0.0"
     .XmaxSpace "0.0"
     .YminSpace "0.0"
     .YmaxSpace "0.0"
     .ZminSpace "0.0"
     .ZmaxSpace "0.0"
End With
With Boundary
     .Xmin "expanded open"
     .Xmax "expanded open"
     .Ymin "expanded open"
     .Ymax "expanded open"
     .Zmin "expanded open"
     .Zmax "expanded open"
     .Xsymmetry "none"
     .Ysymmetry "none"
     .Zsymmetry "none"
End With
' optimize mesh settings for planar structures
With Mesh
     .MergeThinPECLayerFixpoints "True"
     .RatioLimit "20"
     .AutomeshRefineAtPecLines "True", "6"
     .FPBAAvoidNonRegUnite "True"
     .ConsiderSpaceForLowerMeshLimit "False"
     .MinimumStepNumber "5"
     .AnisotropicCurvatureRefinement "True"
     .AnisotropicCurvatureRefinementFSM "True"
End With
With MeshSettings
     .SetMeshType "Hex"
     .Set "RatioLimitGeometry", "20"
     .Set "EdgeRefinementOn", "1"
     .Set "EdgeRefinementRatio", "6"
End With
With MeshSettings
     .SetMeshType "HexTLM"
     .Set "RatioLimitGeometry", "20"
End With
With MeshSettings
     .SetMeshType "Tet"
     .Set "VolMeshGradation", "1.5"
     .Set "SrfMeshGradation", "1.5"
End With
' change mesh adaption scheme to energy
' 		(planar structures tend to store high energy
'     	 locally at edges rather than globally in volume)
MeshAdaption3D.SetAdaptionStrategy "Energy"
' switch on FD-TET setting for accurate farfields
FDSolver.ExtrudeOpenBC "True"
PostProcess1D.ActivateOperation "vswr", "true"
PostProcess1D.ActivateOperation "yz-matrices", "true"
With FarfieldPlot
	.ClearCuts ' lateral=phi, polar=theta
	.AddCut "lateral", "0", "1"
	.AddCut "lateral", "90", "1"
	.AddCut "polar", "90", "1"
End With
'----------------------------------------------------------------------------
'set the frequency range
Solver.FrequencyRange "28", "38"
Dim sDefineAt As String
sDefineAt = "28;33;38"
Dim sDefineAtName As String
sDefineAtName = "28;33;38"
Dim sDefineAtToken As String
sDefineAtToken = "f="
Dim aFreq() As String
aFreq = Split(sDefineAt, ";")
Dim aNames() As String
aNames = Split(sDefineAtName, ";")
Dim nIndex As Integer
For nIndex = LBound(aFreq) To UBound(aFreq)
Dim zz_val As String
zz_val = aFreq (nIndex)
Dim zz_name As String
zz_name = sDefineAtToken & aNames (nIndex)
' Define E-Field Monitors
With Monitor
    .Reset
    .Name "e-field ("& zz_name &")"
    .Dimension "Volume"
    .Domain "Frequency"
    .FieldType "Efield"
    .MonitorValue  zz_val
    .Create
End With
' Define H-Field Monitors
With Monitor
    .Reset
    .Name "h-field ("& zz_name &")"
    .Dimension "Volume"
    .Domain "Frequency"
    .FieldType "Hfield"
    .MonitorValue  zz_val
    .Create
End With
' Define Power flow Monitors
With Monitor
    .Reset
    .Name "power ("& zz_name &")"
    .Dimension "Volume"
    .Domain "Frequency"
    .FieldType "Powerflow"
    .MonitorValue  zz_val
    .Create
End With
' Define Power loss Monitors
With Monitor
    .Reset
    .Name "loss ("& zz_name &")"
    .Dimension "Volume"
    .Domain "Frequency"
    .FieldType "Powerloss"
    .MonitorValue  zz_val
    .Create
End With
' Define Farfield Monitors
With Monitor
    .Reset
    .Name "farfield ("& zz_name &")"
    .Domain "Frequency"
    .FieldType "Farfield"
    .MonitorValue  zz_val
    .ExportFarfieldSource "False"
    .Create
End With
Next
'----------------------------------------------------------------------------
With MeshSettings
     .SetMeshType "Hex"
     .Set "Version", 1%
End With
With Mesh
     .MeshType "PBA"
End With
'set the solver type
ChangeSolverType("HF Time Domain")

'@ define material: Rogers RT5880 (loss free)

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Material
     .Reset
     .Name "Rogers RT5880 (loss free)"
     .Folder ""
.FrqType "all"
.Type "Normal"
.SetMaterialUnit "GHz", "mm"
.Epsilon "2.2"
.Mu "1.0"
.Kappa "0.0"
.TanD "0.0"
.TanDFreq "0.0"
.TanDGiven "False"
.TanDModel "ConstTanD"
.KappaM "0.0"
.TanDM "0.0"
.TanDMFreq "0.0"
.TanDMGiven "False"
.TanDMModel "ConstKappa"
.DispModelEps "None"
.DispModelMu "None"
.DispersiveFittingSchemeEps "General 1st"
.DispersiveFittingSchemeMu "General 1st"
.UseGeneralDispersionEps "False"
.UseGeneralDispersionMu "False"
.Rho "0.0"
.ThermalType "Normal"
.ThermalConductivity "0.20"
.SetActiveMaterial "all"
.Colour "0.75", "0.95", "0.85"
.Wireframe "False"
.Transparency "0"
.Create
End With

'@ new component: component1

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Component.New "component1"

'@ define brick: component1:Substrate

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Brick
     .Reset 
     .Name "Substrate" 
     .Component "component1" 
     .Material "Rogers RT5880 (loss free)" 
     .Xrange "-Ws/2", "20.116+1.84" 
     .Yrange "-Ls/2-3.2", "Ls/2" 
     .Zrange "0", "Hs" 
     .Create
End With

'@ define brick: component1:Ground

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Brick
     .Reset 
     .Name "Ground" 
     .Component "component1" 
     .Material "Rogers RT5880 (loss free)" 
     .Xrange "-WG/2", "Wg/2" 
     .Yrange "-Lg/2", "Lg/2" 
     .Zrange "Hs", "-Hg" 
     .Create
End With

'@ define brick: component1:Ground1

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Brick
     .Reset 
     .Name "Ground1" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-Wg/2", "20.116+1.84" 
     .Yrange "-Ls/2-3.2", "Lg/2" 
     .Zrange "0", "-Hg" 
     .Create
End With

'@ delete shape: component1:Ground

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Solid.Delete "component1:Ground"

'@ define brick: component1:Patch

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Brick
     .Reset 
     .Name "Patch" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-Wp/2", "Wp/2" 
     .Yrange "-Lp/2", "Lp/2" 
     .Zrange "Hs", "Hp" 
     .Create
End With

'@ define brick: component1:Microstrip

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Brick
     .Reset 
     .Name "Microstrip" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-Wf/2+0.2", "Wf/2-0.2" 
     .Yrange "-Lp/2", "-Ls/2" 
     .Zrange "Hs", "Hp" 
     .Create
End With

'@ define brick: component1:Inset1

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Brick
     .Reset 
     .Name "Inset1" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "Wf/2-0.2", "Wf/2+Wi" 
     .Yrange "-Lp/2+Li", "-Lp/2" 
     .Zrange "Hs", "Hp" 
     .Create
End With

'@ boolean subtract shapes: component1:Patch, component1:Inset1

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Solid.Subtract "component1:Patch", "component1:Inset1"

'@ define brick: component1:Inset2

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Brick
     .Reset 
     .Name "Inset2" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "-Wf/2-Wi", "-Wf/2+0.2" 
     .Yrange "-Lp/2+Li", "-Lp/2" 
     .Zrange "Hs", "Hp" 
     .Create
End With

'@ boolean subtract shapes: component1:Patch, component1:Inset2

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Solid.Subtract "component1:Patch", "component1:Inset2"

'@ change material and color: component1:Substrate to: Rogers RT5880 (loss free)

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Solid.SetUseIndividualColor "component1:Substrate", 1
Solid.ChangeIndividualColor "component1:Substrate", "255", "255", "255"

'@ change material and color: component1:Ground1 to: PEC

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Solid.SetUseIndividualColor "component1:Ground1", 1
Solid.ChangeIndividualColor "component1:Ground1", "255", "255", "0"

'@ change material and color: component1:Microstrip to: PEC

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Solid.SetUseIndividualColor "component1:Microstrip", 1
Solid.ChangeIndividualColor "component1:Microstrip", "255", "255", "0"

'@ change material and color: component1:Patch to: PEC

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Solid.SetUseIndividualColor "component1:Patch", 1
Solid.ChangeIndividualColor "component1:Patch", "255", "255", "0"

'@ define brick: component1:Inset3

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Brick
     .Reset 
     .Name "Inset3" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "1", "1+W1" 
     .Yrange "Lp/2-0.25-0.2", "Lp/2-0.25" 
     .Zrange "Hs", "Hp" 
     .Create
End With

'@ boolean subtract shapes: component1:Patch, component1:Inset3

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Solid.Subtract "component1:Patch", "component1:Inset3"

'@ define brick: component1:Inset4

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Brick
     .Reset 
     .Name "Inset4" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "1", "1+T1" 
     .Yrange "Lp/2-0.25-L1", "Lp/2-0.25-T1" 
     .Zrange "Hs", "Hp" 
     .Create
End With

'@ boolean subtract shapes: component1:Patch, component1:Inset4

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Solid.Subtract "component1:Patch", "component1:Inset4"

'@ define brick: component1:Inset5

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Brick
     .Reset 
     .Name "Inset5" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "1+T1", "1+W1" 
     .Yrange "Lp/2-0.25-L1", "Lp/2-0.25-L1+T1" 
     .Zrange "Hs", "Hp" 
     .Create
End With

'@ boolean subtract shapes: component1:Patch, component1:Inset5

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Solid.Subtract "component1:Patch", "component1:Inset5"

'@ define brick: component1:Inset6

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Brick
     .Reset 
     .Name "Inset6" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "1+W1-T1", "1+W1" 
     .Yrange "Lp/2-0.25-L1", "-0.317" 
     .Zrange "Hs", "Hp" 
     .Create
End With

'@ boolean subtract shapes: component1:Patch, component1:Inset6

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Solid.Subtract "component1:Patch", "component1:Inset6"

'@ define brick: component1:Inset7

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Brick
     .Reset 
     .Name "Inset7" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "1", "1+W1-T1" 
     .Yrange "-0.317+T1", "-0.317" 
     .Zrange "Hs", "Hp" 
     .Create
End With

'@ boolean subtract shapes: component1:Patch, component1:Inset7

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Solid.Subtract "component1:Patch", "component1:Inset7"

'@ pick face

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Pick.PickFaceFromId "component1:Microstrip", "3"

'@ define port: 1

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Port 
     .Reset 
     .PortNumber "1" 
     .Label "" 
     .Folder "" 
     .NumberOfModes "1" 
     .AdjustPolarization "False" 
     .PolarizationAngle "0.0" 
     .ReferencePlaneDistance "0" 
     .TextSize "50" 
     .TextMaxLimit "0" 
     .Coordinates "Picks" 
     .Orientation "positive" 
     .PortOnBound "False" 
     .ClipPickedPortToBound "False" 
     .Xrange "-0.391", "0.391" 
     .Yrange "-3.85", "-3.85" 
     .Zrange "0.254", "0.289" 
     .XrangeAdd "K*Hs", "K*Hs" 
     .YrangeAdd "0.0", "0.0" 
     .ZrangeAdd "Hs", "K*Hs" 
     .SingleEnded "False" 
     .WaveguideMonitor "False" 
     .Create 
End With

'@ define time domain solver parameters

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Mesh.SetCreator "High Frequency" 
With Solver 
     .Method "Hexahedral"
     .CalculationType "TD-S"
     .StimulationPort "All"
     .StimulationMode "All"
     .SteadyStateLimit "-40"
     .MeshAdaption "True"
     .AutoNormImpedance "True"
     .NormingImpedance "50"
     .CalculateModesOnly "False"
     .SParaSymmetry "False"
     .StoreTDResultsInCache  "False"
     .FullDeembedding "False"
     .SuperimposePLWExcitation "False"
     .UseSensitivityAnalysis "False"
End With

'@ define monitor: e-field (f=50)

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Monitor 
     .Reset 
     .Name "e-field (f=50)" 
     .Dimension "Volume" 
     .Domain "Frequency" 
     .FieldType "Efield" 
     .MonitorValue "50" 
     .UseSubvolume "False" 
     .Coordinates "Structure" 
     .SetSubvolume "-6.2211549848485", "6.2211549848485", "-6.1211549848485", "6.1211549848485", "-2.3061549848485", "3.7361749848485" 
     .SetSubvolumeOffset "0.0", "0.0", "0.0", "0.0", "0.0", "0.0" 
     .Create 
End With

'@ define monitor: h-field (f=50)

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Monitor 
     .Reset 
     .Name "h-field (f=50)" 
     .Dimension "Volume" 
     .Domain "Frequency" 
     .FieldType "Hfield" 
     .MonitorValue "50" 
     .UseSubvolume "False" 
     .Coordinates "Structure" 
     .SetSubvolume "-6.2211549848485", "6.2211549848485", "-6.1211549848485", "6.1211549848485", "-2.3061549848485", "3.7361749848485" 
     .SetSubvolumeOffset "0.0", "0.0", "0.0", "0.0", "0.0", "0.0" 
     .Create 
End With

'@ define monitor: surface-current (f=50)

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Monitor 
     .Reset 
     .Name "surface-current (f=50)" 
     .Dimension "Volume" 
     .Domain "Frequency" 
     .FieldType "Surfacecurrent" 
     .MonitorValue "50" 
     .Create 
End With

'@ define farfield monitor: farfield (f=50)

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Monitor 
     .Reset 
     .Name "farfield (f=50)" 
     .Domain "Frequency" 
     .FieldType "Farfield" 
     .MonitorValue "50" 
     .ExportFarfieldSource "False" 
     .UseSubvolume "False" 
     .Coordinates "Structure" 
     .SetSubvolume "-6.2211549848485", "6.2211549848485", "-6.1211549848485", "6.1211549848485", "-2.3061549848485", "3.7361749848485" 
     .SetSubvolumeOffset "10", "10", "10", "10", "10", "10" 
     .SetSubvolumeOffsetType "FractionOfWavelength" 
     .Create 
End With

'@ boolean add shapes: component1:Microstrip, component1:Patch

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Solid.Add "component1:Microstrip", "component1:Patch"

'@ transform: translate component1:Microstrip

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Transform 
     .Reset 
     .Name "component1:Microstrip" 
     .Vector "6", "0", "0" 
     .UsePickedPoints "False" 
     .InvertPickedPoints "False" 
     .MultipleObjects "True" 
     .GroupObjects "False" 
     .Repetitions "3" 
     .MultipleSelection "False" 
     .Destination "" 
     .Material "" 
     .Transform "Shape", "Translate" 
End With

'@ clear picks

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Pick.ClearAllPicks

'@ delete port: port1

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Port.Delete "1"

'@ clear picks

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Pick.ClearAllPicks

'@ define brick: component1:line

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Brick
     .Reset 
     .Name "line" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-0.391+0.2", "18.391-0.2" 
     .Yrange "-3.85", "-4.2" 
     .Zrange "Hs", "Hp" 
     .Create
End With

'@ define material: Copper (annealed)

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Material
     .Reset
     .Name "Copper (annealed)"
     .Folder ""
     .FrqType "static"
     .Type "Normal"
     .SetMaterialUnit "Hz", "mm"
     .Epsilon "1"
     .Mu "1.0"
     .Kappa "5.8e+007"
     .TanD "0.0"
     .TanDFreq "0.0"
     .TanDGiven "False"
     .TanDModel "ConstTanD"
     .KappaM "0"
     .TanDM "0.0"
     .TanDMFreq "0.0"
     .TanDMGiven "False"
     .TanDMModel "ConstTanD"
     .DispModelEps "None"
     .DispModelMu "None"
     .DispersiveFittingSchemeEps "Nth Order"
     .DispersiveFittingSchemeMu "Nth Order"
     .UseGeneralDispersionEps "False"
     .UseGeneralDispersionMu "False"
     .FrqType "all"
     .Type "Lossy metal"
     .SetMaterialUnit "GHz", "mm"
     .Mu "1.0"
     .Kappa "5.8e+007"
     .Rho "8930.0"
     .ThermalType "Normal"
     .ThermalConductivity "401.0"
     .HeatCapacity "0.39"
     .MetabolicRate "0"
     .BloodFlow "0"
     .VoxelConvection "0"
     .MechanicsType "Isotropic"
     .YoungsModulus "120"
     .PoissonsRatio "0.33"
     .ThermalExpansionRate "17"
     .Colour "1", "1", "0"
     .Wireframe "False"
     .Reflection "False"
     .Allowoutline "True"
     .Transparentoutline "False"
     .Transparency "0"
     .Create
End With

'@ change material: component1:line to: Copper (annealed)

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Solid.ChangeMaterial "component1:line", "Copper (annealed)"

'@ define brick: component1:feed

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Brick
     .Reset 
     .Name "feed" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "9.39-0.25", "9.39+0.25" 
     .Yrange "-4.2", "-7.05" 
     .Zrange "Hs", "Hp" 
     .Create
End With

'@ clear picks

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Pick.ClearAllPicks

'@ boolean add shapes: component1:feed, component1:line

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Solid.Add "component1:feed", "component1:line"

'@ clear picks

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Pick.ClearAllPicks

'@ pick face

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Pick.PickFaceFromId "component1:feed", "9"

'@ define port:1

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
' Port constructed by macro Solver -> Ports -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.254*5.08", "0.254*5.08"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.254", "0.254*5.08"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ pick face

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Pick.PickFaceFromId "component1:feed", "9"

'@ define port:2

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
' Port constructed by macro Solver -> Ports -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "2"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.254*5.08", "0.254*5.08"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.254", "0.254*5.08"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ delete port: port2

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Port.Delete "2"

'@ define time domain solver parameters

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Mesh.SetCreator "High Frequency" 
With Solver 
     .Method "Hexahedral"
     .CalculationType "TD-S"
     .StimulationPort "All"
     .StimulationMode "All"
     .SteadyStateLimit "-30"
     .MeshAdaption "True"
     .AutoNormImpedance "True"
     .NormingImpedance "50"
     .CalculateModesOnly "False"
     .SParaSymmetry "False"
     .StoreTDResultsInCache  "False"
     .FullDeembedding "False"
     .SuperimposePLWExcitation "False"
     .UseSensitivityAnalysis "False"
End With

'@ clear picks

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Pick.ClearAllPicks

'@ pick edge

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Pick.PickEdgeFromId "component1:Microstrip_3", "153", "111"

'@ pick end point

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Pick.PickEndpointFromId "component1:Microstrip_3", "57"

'@ define brick: component1:cut1

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Brick
     .Reset 
     .Name "cut1" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "18", "18.5" 
     .Yrange "-4", "0" 
     .Zrange "0", "-0.035" 
     .Create
End With

'@ boolean subtract shapes: component1:Ground1, component1:cut1

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Solid.Subtract "component1:Ground1", "component1:cut1"

'@ define brick: component1:cut2

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Brick
     .Reset 
     .Name "cut2" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "12", "12.5" 
     .Yrange "-4", "0" 
     .Zrange "0", "-0.035" 
     .Create
End With

'@ boolean subtract shapes: component1:Ground1, component1:cut2

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Solid.Subtract "component1:Ground1", "component1:cut2"

'@ define brick: component1:cut3

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Brick
     .Reset 
     .Name "cut3" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "6", "6.5" 
     .Yrange "-4", "0" 
     .Zrange "0", "-0.035" 
     .Create
End With

'@ boolean subtract shapes: component1:Ground1, component1:cut3

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Solid.Subtract "component1:Ground1", "component1:cut3"

'@ define brick: component1:cut4

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Brick
     .Reset 
     .Name "cut4" 
     .Component "component1" 
     .Material "Vacuum" 
     .Xrange "0", "0.5" 
     .Yrange "-4", "0" 
     .Zrange "0", "-0.035" 
     .Create
End With

'@ boolean subtract shapes: component1:Ground1, component1:cut4

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Solid.Subtract "component1:Ground1", "component1:cut4"

'@ delete port: port1

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Pick.PickFaceFromId "component1:feed", "9"

'@ define port:1

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
' Port constructed by macro Solver -> Ports -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.254*5.57", "0.254*5.57"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.254", "0.254*5.57"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ clear picks

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
Pick.ClearAllPicks

'@ define monitors (using linear samples)

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Monitor
          .Reset 
          .Domain "Frequency"
          .FieldType "Efield"
          .Dimension "Volume" 
          .UseSubvolume "False" 
          .Coordinates "Structure" 
          .SetSubvolume "-6.2211549848485", "24.227154984848", "-9.3211549848485", "6.1211549848485", "-2.3061549848485", "3.9749349848485" 
          .SetSubvolumeOffset "0.0", "0.0", "0.0", "0.0", "0.0", "0.0" 
          .CreateUsingLinearSamples "28", "38", "100"
End With

'@ define monitors (using linear samples)

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Monitor
          .Reset 
          .Domain "Frequency"
          .FieldType "Farfield"
          .ExportFarfieldSource "False" 
          .UseSubvolume "False" 
          .Coordinates "Structure" 
          .SetSubvolume "-3.95", "21.956", "-7.05", "3.85", "-0.035", "1.70378" 
          .SetSubvolumeOffset "10", "10", "10", "10", "10", "10" 
          .SetSubvolumeOffsetType "FractionOfWavelength" 
          .CreateUsingLinearSamples "28", "38", "100"
End With

'@ define monitors (using linear samples)

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Monitor
          .Reset 
          .Domain "Frequency"
          .FieldType "Efield"
          .Dimension "Volume" 
          .UseSubvolume "False" 
          .Coordinates "Structure" 
          .SetSubvolume "-3.95", "21.956", "-7.05", "3.85", "-0.035", "1.70378" 
          .SetSubvolumeOffset "0.0", "0.0", "0.0", "0.0", "0.0", "0.0" 
          .CreateUsingLinearSamples "28", "38", "100"
End With

'@ define monitors (using linear samples)

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Monitor
          .Reset 
          .Domain "Frequency"
          .FieldType "Surfacecurrent"
          .Dimension "Volume" 
          .CreateUsingLinearSamples "28", "38", "100"
End With

'@ define monitors (using linear samples)

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Monitor
          .Reset 
          .Domain "Frequency"
          .FieldType "Current"
          .Dimension "Volume" 
          .CreateUsingLinearSamples "28", "38", "100"
End With

'@ farfield plot options

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With FarfieldPlot 
     .Plottype "Polar" 
     .Vary "angle1" 
     .Theta "90" 
     .Phi "90" 
     .Step "1" 
     .Step2 "1" 
     .SetLockSteps "True" 
     .SetPlotRangeOnly "False" 
     .SetThetaStart "0" 
     .SetThetaEnd "180" 
     .SetPhiStart "0" 
     .SetPhiEnd "360" 
     .SetTheta360 "True" 
     .SymmetricRange "True" 
     .SetTimeDomainFF "False" 
     .SetFrequency "33.1515" 
     .SetTime "0" 
     .SetColorByValue "True" 
     .DrawStepLines "False" 
     .DrawIsoLongitudeLatitudeLines "False" 
     .ShowStructure "False" 
     .ShowStructureProfile "False" 
     .SetStructureTransparent "False" 
     .SetFarfieldTransparent "False" 
     .SetSpecials "enablepolarextralines" 
     .SetPlotMode "Directivity" 
     .Distance "1" 
     .UseFarfieldApproximation "True" 
     .SetScaleLinear "False" 
     .SetLogRange "40" 
     .SetLogNorm "0" 
     .DBUnit "0" 
     .EnableFixPlotMaximum "False" 
     .SetFixPlotMaximumValue "1" 
     .SetInverseAxialRatio "False" 
     .SetAxesType "user" 
     .SetAntennaType "unknown" 
     .Phistart "1.000000e+00", "0.000000e+00", "0.000000e+00" 
     .Thetastart "0.000000e+00", "0.000000e+00", "1.000000e+00" 
     .PolarizationVector "0.000000e+00", "1.000000e+00", "0.000000e+00" 
     .SetCoordinateSystemType "spherical" 
     .SetAutomaticCoordinateSystem "True" 
     .SetPolarizationType "Linear" 
     .SlantAngle 0.000000e+00 
     .Origin "bbox" 
     .Userorigin "0.000000e+00", "0.000000e+00", "0.000000e+00" 
     .SetUserDecouplingPlane "False" 
     .UseDecouplingPlane "False" 
     .DecouplingPlaneAxis "X" 
     .DecouplingPlanePosition "0.000000e+00" 
     .LossyGround "False" 
     .GroundEpsilon "1" 
     .GroundKappa "0" 
     .EnablePhaseCenterCalculation "False" 
     .SetPhaseCenterAngularLimit "3.000000e+01" 
     .SetPhaseCenterComponent "boresight" 
     .SetPhaseCenterPlane "both" 
     .ShowPhaseCenter "True" 
     .ClearCuts 
     .AddCut "lateral", "0", "1"  
     .AddCut "lateral", "90", "1"  
     .AddCut "polar", "90", "1"  
     .StoreSettings
End With

'@ define monitors (using linear samples)

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Monitor
          .Reset 
          .Domain "Frequency"
          .FieldType "Efield"
          .Dimension "Volume" 
          .UseSubvolume "False" 
          .Coordinates "Structure" 
          .SetSubvolume "-3.95", "21.956", "-7.05", "3.85", "-0.035", "1.70378" 
          .SetSubvolumeOffset "0.0", "0.0", "0.0", "0.0", "0.0", "0.0" 
          .CreateUsingLinearSamples "28", "38", "100"
End With

'@ define monitors (using linear samples)

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Monitor
          .Reset 
          .Domain "Frequency"
          .FieldType "Farfield"
          .ExportFarfieldSource "False" 
          .UseSubvolume "False" 
          .Coordinates "Structure" 
          .SetSubvolume "-3.95", "21.956", "-7.05", "3.85", "-0.035", "1.70378" 
          .SetSubvolumeOffset "10", "10", "10", "10", "10", "10" 
          .SetSubvolumeOffsetType "FractionOfWavelength" 
          .CreateUsingLinearSamples "28", "38", "100"
End With

'@ define monitors (using linear samples)

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Monitor
          .Reset 
          .Domain "Frequency"
          .FieldType "Current"
          .Dimension "Volume" 
          .CreateUsingLinearSamples "28", "38", "100"
End With

'@ farfield plot options

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With FarfieldPlot 
     .Plottype "3D" 
     .Vary "angle1" 
     .Theta "90" 
     .Phi "90" 
     .Step "5" 
     .Step2 "5" 
     .SetLockSteps "True" 
     .SetPlotRangeOnly "False" 
     .SetThetaStart "0" 
     .SetThetaEnd "180" 
     .SetPhiStart "0" 
     .SetPhiEnd "360" 
     .SetTheta360 "True" 
     .SymmetricRange "True" 
     .SetTimeDomainFF "False" 
     .SetFrequency "33.1515" 
     .SetTime "0" 
     .SetColorByValue "True" 
     .DrawStepLines "False" 
     .DrawIsoLongitudeLatitudeLines "False" 
     .ShowStructure "False" 
     .ShowStructureProfile "False" 
     .SetStructureTransparent "False" 
     .SetFarfieldTransparent "False" 
     .SetSpecials "enablepolarextralines" 
     .SetPlotMode "Directivity" 
     .Distance "1" 
     .UseFarfieldApproximation "True" 
     .SetScaleLinear "False" 
     .SetLogRange "40" 
     .SetLogNorm "0" 
     .DBUnit "0" 
     .EnableFixPlotMaximum "False" 
     .SetFixPlotMaximumValue "1" 
     .SetInverseAxialRatio "False" 
     .SetAxesType "user" 
     .SetAntennaType "unknown" 
     .Phistart "1.000000e+00", "0.000000e+00", "0.000000e+00" 
     .Thetastart "0.000000e+00", "0.000000e+00", "1.000000e+00" 
     .PolarizationVector "0.000000e+00", "1.000000e+00", "0.000000e+00" 
     .SetCoordinateSystemType "spherical" 
     .SetAutomaticCoordinateSystem "True" 
     .SetPolarizationType "Linear" 
     .SlantAngle 0.000000e+00 
     .Origin "bbox" 
     .Userorigin "0.000000e+00", "0.000000e+00", "0.000000e+00" 
     .SetUserDecouplingPlane "False" 
     .UseDecouplingPlane "False" 
     .DecouplingPlaneAxis "X" 
     .DecouplingPlanePosition "0.000000e+00" 
     .LossyGround "False" 
     .GroundEpsilon "1" 
     .GroundKappa "0" 
     .EnablePhaseCenterCalculation "False" 
     .SetPhaseCenterAngularLimit "3.000000e+01" 
     .SetPhaseCenterComponent "boresight" 
     .SetPhaseCenterPlane "both" 
     .ShowPhaseCenter "True" 
     .ClearCuts 
     .AddCut "lateral", "0", "1"  
     .AddCut "lateral", "90", "1"  
     .AddCut "polar", "90", "1"  
     .StoreSettings
End With

'@ define monitors (using linear samples)

'[VERSION]2018.0|27.0.2|20171026[/VERSION]
With Monitor
          .Reset 
          .Domain "Frequency"
          .FieldType "Current"
          .Dimension "Volume" 
          .CreateUsingLinearSamples "28", "38", "100"
End With

