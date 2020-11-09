Option Explicit
' Copyright 2020 Forrest M. S. Baird

'---------------------------------------------------------------------------------------------------GLOBAL--DECLARATIONS---------------------------------------------------------------------------------------------------

' Worksheet Declarations
Public AutoTools As Worksheet
Public CalculationSheets As Worksheet
Public CoolantData As Worksheet
Public FuelSamplePackages As Worksheet
Public LabData As Worksheet
Public SampleCreation As Worksheet

' Lab Data Non-Test Column References (Alphabetical)
Public Const FuelPackageColumn As Integer = 4                   'D
Public Const FuelTypeColumn As Integer = 2                      'B
Public Const MonthTagColumn As Integer = 32                     'AF
Public Const RushPriorityColumn As Integer = 3                  'C
Public Const SampleIDColumn As Integer = 1                      'A
Public Const SampleIDMirrorColumn As Integer = 31               'AE
Public Const SampleStatusColumn As Integer = 33                 'AG

' Universal Row References
Public Const LabDataHeaderRow As Integer = 1
Public Const CalculationSheetsHeaderRow As Integer = 5

' Lab Data Production - Column References (Alphabetical)        'Column, Test Position
Public Const AcidNumberColumn As Integer = 30                   'AD, 1
Public Const APIColumn As Integer = 10                          'J, 2
Public Const CetaneIndexColumn As Integer = 11                  'K, 3
Public Const CloudPointColumn As Integer = 24                   'X, 4
Public Const CopperStripCorrosionColumn As Integer = 23         'W, 5
Public Const Distillation010Column As Integer = 8               'H, 6
Public Const Distillation050Column As Integer = 7               'G, 7
Public Const Distillation090Column As Integer = 6               'F, 8
Public Const DistillationFBPColumn As Integer = 5               'E, 9
Public Const DistillationIBPColumn As Integer = 9               'I, 10
Public Const FilterPatchColumn As Integer = 28                  'AB, 11
Public Const FlashPointColumn As Integer = 12                   'L, 12
Public Const FuelBlendRatioColumn As Integer = 19               'S, 13
Public Const MicrobialGrowthColumn As Integer = 22              'T, 14
Public Const ParticulateContaminationColumn As Integer = 26     'Z, 15
Public Const PourPointColumn As Integer = 25                    'Y, 16
Public Const RelativeAbsorbanceColumn As Integer = 18           'R, 17
Public Const SulfurColumn As Integer = 21                       'U, 18
Public Const ThermalStabilityColumn As Integer = 27             'AA, 19
Public Const ViscosityNumberColumn As Integer = 29              'AC, 20
Public Const VisualOpacityColumn As Integer = 13                'M, 21
Public Const VisualParticlesColumn As Integer = 16              'P, 22
Public Const VisualPhaseColumn As Integer = 15                  'N, 23
Public Const VisualSedimentColumn As Integer = 14               'N, 24
Public Const WaterAndSedimentColumn As Integer = 20             'T, 25
Public Const WaterByKFColumn As Integer = 17                    'Q, 26

' AutoTools Constants
Public Const PackageReplacementColumn As Integer = 2
Public Const PackageReplacementSampleRow As Integer = 22
Public Const PackageReplacementPackageRow As Integer = 23
Public Const VisualSamplesToFillRow As Integer = 6
Public Const VisualSamplesRemainingRow As Integer = 7
Public Const VisualSamplesFillColumn As Integer = 10            ' J = 10
Public Const WaterAndSedimentToFillRow As Integer = 16
Public Const WaterAndSedimentRemainingRow As Integer = 17
Public Const WaterAndSedimentFillColumn As Integer = 10         ' J = 10
Public Const CopperStripCorrosionToFillRow As Integer = 17
Public Const CopperStripCorrosionRemainingRow As Integer = 18
Public Const CopperStripCorrosionFillColumn As Integer = 7      ' G = 7


' Coolant Spreadsheet - Row & Column References
Public Const CoolantDataHeaderRow As Integer = 1
Public Const CoolantSampleIDColumn As Integer = 1               'A
Public Const CoolantTotalDissolvedSolidsColumn As Integer = 2   'B
Public Const CoolantGlycolPercentageColumn As Integer = 3       'C
Public Const CoolantFreezePointColumn As Integer = 4            'D
Public Const CoolantColorColumn As Integer = 5                  'E
Public Const CoolantOpacityColumn As Integer = 6                'F
Public Const CoolantDebrisColumn As Integer = 7                 'G
Public Const CoolantpHColumn As Integer = 8                     'H
Public Const CoolantNitriteColumn As Integer = 9                'I
Public Const CoolantSampleStatusColumn As Integer = 10          'J
Public Const CoolantSampleStartingRow As Integer = 2

' Calculation Sheets
Public Const PartContSampleIDColumn As Integer = 1
Public Const PartContTrayNumberColumn As Integer = 2
Public Const PartContTopBeforeColumn As Integer = 3
Public Const PartContControlBeforeColumn As Integer = 4
Public Const PartContTopAfterColumn As Integer = 5
Public Const PartContControlAfterColumn As Integer = 6
Public Const PartContSampleVolumeColumn As Integer = 7
Public Const PartContCalculationColumn As Integer = 8

Public Const ThermStabSampleIDColumn As Integer = 12
Public Const ThermStabPatch1Column As Integer = 13
Public Const ThermStabPatch2Column As Integer = 14
Public Const ThermStabAverageColumn As Integer = 15
Public Const ThermStabWhitePatchColumn As Integer = 16
Public Const ThermStabCalculationColumn As Integer = 17

Public Const ViscSampleIDColumn As Integer = 21
Public Const ViscMinuteColumn As Integer = 22
Public Const ViscSecondColumn As Integer = 23
Public Const ViscTimeColumn As Integer = 24
Public Const ViscMeterTypeColumn As Integer = 25
Public Const ViscTemperatureColumn As Integer = 26
Public Const ViscCoefficientColumn As Integer = 27
Public Const ViscCalculationColumn As Integer = 28


' Sample Interface - Global Sample Row and Column References
Public Const CoolantInputRow As Integer = 2
Public Const FirstSampleInputRow As Integer = 3
Public Const SampleMonthRow As Integer = 4
Public Const SampleDayRow As Integer = 5
Public Const SampleYearRow As Integer = 6
Public Const InitialUserInputColumn As Integer = 2
Public Const FirstUserInputRow As Integer = 13

' Sample Interface - Fuel Sample Column References
Public Const AutoGeneratedStartingSampleColumn As Integer = 1
Public Const UserInputEndingSampleColumn As Integer = 2
Public Const UserInputFuelTypeColumn As Integer = 3
Public Const UserInputRushPriorityColumn As Integer = 4
Public Const UserInputPackageColumn As Integer = 5
Public Const ErrorMessageColumn As Integer = 5
Public Const ErrorMessageRow As Integer = 4
Public Const SuccessMessageColumn As Integer = 5
Public Const SuccessMessageRow As Integer = 8


' Sample Generation Constants
Public Const MinimumFuelSampleNumber As Integer = 1
Public Const MaximumFuelSampleNumber As Integer = 999
Public Const MinimumCoolantSampleNumber As Integer = 1
Public Const MaximumCoolantSampleNumber As Integer = 99
Public Const SampleTen As Integer = 10
Public Const SampleOneHundred As Integer = 100
Public Const SampleOneThousand As Integer = 1000
Public Const SampleYearCorrectionFactor As Integer = 2000

' Data Points
Public SampleIDOutput As Object
Public FuelTypeOutput As Object
Public RushPriorityOutput As Object
Public PackageOutput As Object
Public DistillationFBP As Object
Public Distillation090 As Object
Public Distillation050 As Object
Public Distillation010 As Object
Public DistillationIBP As Object
Public API As Object
Public CetaneIndex As Object
Public FlashPoint As Object
Public VisualOpacity As Object
Public VisualSediment As Object
Public VisualPhase As Object
Public VisualParticles As Object
Public WaterByKF As Object
Public RelativeAbsorbance As Object
Public FuelBlendRatio As Object
Public WaterAndSediment As Object
Public Sulfur As Object
Public MicrobialGrowth As Object
Public CopperStripCorrosion As Object
Public CloudPoint As Object
Public PourPoint As Object
Public ParticulateContamination As Object
Public ThermalStability As Object
Public FilterPatch As Object
Public ViscosityNumber As Object
Public AcidNumber As Object
Public SampleIDMirror As Object
Public MonthTag As Object
Public SampleStatus As Object

' Common Interface Variables
Public CellInQuestion As Object
Public CurrentEndingSampleNumber As Object
Public CurrentStartingSampleNumber As Object
Public Dash As String
Public EndingSampleNumber As Object
Public ExtraZeroDay As Variant
Public ExtraZeroMonth As Variant
Public ExtraZeroYear As Variant
Public FinalUserInputNumber As Object
Public FinalUserInputRow As Integer
Public FuelType As Object
Public HideCounter As Integer
Public LastDayOfMonth As Integer
Public MMDDYY As String
Public Response As Variant
Public RowChecker As Integer
Public SampleDay As Object
Public SampleMonth As Object
Public SamplesGenerated As Integer
Public SampleYear As Object
Public SampleYearFourDigits As Integer
Public SSS As String
Public StartingSampleNumber As Object

' Other Intermediate Variables
Public FirstColumnAnalyzed As Integer
Public FirstEndingSample As Object
Public LastColumnAnalyzed As Integer
Public SampleRow As Integer
Public SampleID As Object


' Feedback Monitors for Sample Interface
Public ErrorMessage As Object
Public SuccessMessage As Object


Public Sub FuelSampleGenerator_Button()

'---------------------------------------------------------------------------------------------------DECLARATIONS---AND---INITIALIZATIONS---------------------------------------------------------------------------------------------------

' Declaration of Local Variables:
Dim ColumnAnalyzed As Integer
Dim CurrentSampleNumber As Integer
Dim EndingSample As Variant
Dim EndUserInputRow As Integer
Dim FinalUserInputRow As Integer
Dim FuelPackage As Object
Dim FuelPackageOutput As Object
Dim InitialSampleRow As Integer
Dim MicrobialCheckDay As String
Dim MicrobialCheckStatement As String
Dim MonthEntry As String
Dim RushPriority As Object
Dim SampleIDExtraZeroes As String
Dim TodaysDate As Date
Dim UserInputRow As Integer
Dim FirstColumnToAnalyze As Integer
Dim LastColumnToAnalyze As Integer
Dim LessThan010 As String
Dim LessThan100 As String
Dim LabTestColumnValues(1 To 26) As Integer

' Worksheet Setters
Set CalculationSheets = ThisWorkbook.Worksheets("Calculation Sheets")
Set LabData = ThisWorkbook.Worksheets("Lab Data")
Set SampleCreation = ThisWorkbook.Worksheets("Sample Creation")

' Lab Test Column Values
LabTestColumnValues(1) = AcidNumberColumn
LabTestColumnValues(2) = APIColumn
LabTestColumnValues(3) = CetaneIndexColumn
LabTestColumnValues(4) = CloudPointColumn
LabTestColumnValues(5) = CopperStripCorrosionColumn
LabTestColumnValues(6) = Distillation010Column
LabTestColumnValues(7) = Distillation050Column
LabTestColumnValues(8) = Distillation090Column
LabTestColumnValues(9) = DistillationFBPColumn
LabTestColumnValues(10) = DistillationIBPColumn
LabTestColumnValues(11) = FilterPatchColumn
LabTestColumnValues(12) = FlashPointColumn
LabTestColumnValues(13) = FuelBlendRatioColumn
LabTestColumnValues(14) = MicrobialGrowthColumn
LabTestColumnValues(15) = ParticulateContaminationColumn
LabTestColumnValues(16) = PourPointColumn
LabTestColumnValues(17) = RelativeAbsorbanceColumn
LabTestColumnValues(18) = SulfurColumn
LabTestColumnValues(19) = ThermalStabilityColumn
LabTestColumnValues(20) = ViscosityNumberColumn
LabTestColumnValues(21) = VisualOpacityColumn
LabTestColumnValues(22) = VisualParticlesColumn
LabTestColumnValues(23) = VisualPhaseColumn
LabTestColumnValues(24) = VisualSedimentColumn
LabTestColumnValues(25) = WaterAndSedimentColumn
LabTestColumnValues(26) = WaterByKFColumn

' Uses the Array To Determine First And Last Column
FirstColumnToAnalyze = WorksheetFunction.Min(LabTestColumnValues)
LastColumnToAnalyze = WorksheetFunction.Max(LabTestColumnValues)

' Initialization of Variables:
UserInputRow = FirstUserInputRow
EndUserInputRow = UserInputRow
SampleRow = LabDataHeaderRow + 1

' Initialization of Important Constants
Set SampleIDOutput = LabData.Cells(SampleRow, SampleIDColumn)

' Scans Lab Data Sample ID Column Until It Finds a First Blank:
While SampleIDOutput <> ""
    
    SampleRow = SampleRow + 1
    Set SampleIDOutput = LabData.Cells(SampleRow, SampleIDColumn)

Wend

InitialSampleRow = SampleRow

Set FinalUserInputNumber = SampleCreation.Cells(EndUserInputRow, UserInputEndingSampleColumn)

' Scans User Interface To Determine Final User Input Number:
While SampleCreation.Cells(EndUserInputRow + 1, UserInputEndingSampleColumn) <> ""

    EndUserInputRow = EndUserInputRow + 1
    Set FinalUserInputNumber = SampleCreation.Cells(EndUserInputRow, UserInputEndingSampleColumn)

Wend

' Interface Setters:
Set EndingSampleNumber = SampleCreation.Cells(UserInputRow, UserInputEndingSampleColumn)
Set ErrorMessage = SampleCreation.Cells(ErrorMessageRow, ErrorMessageColumn)
Set FuelPackage = SampleCreation.Cells(UserInputRow, UserInputPackageColumn)
Set FuelType = SampleCreation.Cells(UserInputRow, UserInputFuelTypeColumn)
Set RushPriority = SampleCreation.Cells(UserInputRow, UserInputRushPriorityColumn)
Set SampleDay = SampleCreation.Cells(SampleDayRow, InitialUserInputColumn)
Set SampleMonth = SampleCreation.Cells(SampleMonthRow, InitialUserInputColumn)
Set SampleYear = SampleCreation.Cells(SampleYearRow, InitialUserInputColumn)
Set StartingSampleNumber = SampleCreation.Cells(UserInputRow, AutoGeneratedStartingSampleColumn)
Set SuccessMessage = SampleCreation.Cells(SuccessMessageRow, SuccessMessageColumn)


'------------------------------------------------------------------------------------------------------------------------VALIDATION---TESTS------------------------------------------------------------------------------------------------------------------------

' Initialize Messages
ErrorMessage.Value = ""
SuccessMessage.Value = ""

' Part I - Date Validation Check, Which Is Used In Multiple Subs:
Call DateValidationCheck

' Part II - Sample Number Validation Checks For Fuel Samples:

' Validation Test To Make Sure Atleast One Sample Line Is Filled:
If EndingSampleNumber = "" Then

    ErrorMessage.Value = "No Samples Entered! Please enter ending sample numbers and try again. (B13)"
    MsgBox (ErrorMessage)
    SampleCreation.Cells(UserInputRow, UserInputEndingSampleColumn).Select
    Exit Sub

End If

' Validation Test To Ensure Starting Sample Is Between 1-999:
If StartingSampleNumber > MaximumFuelSampleNumber Or StartingSampleNumber < MinimumFuelSampleNumber Or StartingSampleNumber = "" Or Not IsNumeric(StartingSampleNumber) Then
    
    ErrorMessage.Value = "First Sample Not Valid (B3)"
    MsgBox (ErrorMessage)
    SampleCreation.Cells(UserInputRow, AutoGeneratedStartingSampleColumn).Select
    Exit Sub

End If
 
' Validation Test To Ensure All Ending Samples Are Valid:
For RowChecker = UserInputRow To FinalUserInputRow

    Set CurrentEndingSampleNumber = SampleCreation.Cells(RowChecker, UserInputEndingSampleColumn)

    If CurrentEndingSampleNumber <> "" And Not IsNumeric(CurrentEndingSampleNumber) Then
    
        ErrorMessage.Value = "ERROR: Check Ending Samples Entry at Cell (B" + CStr(RowChecker) + ") and ensure it is blank or a real number."
        MsgBox (ErrorMessage)
        Cells(RowChecker, UserInputEndingSampleColumn).Select
        
        Exit Sub
        
    End If
    
Next RowChecker

' Validation Test To Ensure Each Ending Sample Is Greater Than Or Equal To The Starting Sample:
For RowChecker = UserInputRow To FinalUserInputRow

    Set CurrentStartingSampleNumber = SampleCreation.Cells(RowChecker, AutoGeneratedStartingSampleColumn)
    Set CurrentEndingSampleNumber = SampleCreation.Cells(RowChecker, UserInputEndingSampleColumn)

    If CurrentStartingSampleNumber = "" Then
    
        Exit For
        
    Else
    
        If CurrentStartingSampleNumber > CurrentEndingSampleNumber Then
        
            ErrorMessage.Value = "ERROR: Check Cell Value at (B" + CStr(RowChecker) + ") - Starting Sample is greater than Ending Sample!"
            MsgBox (ErrorMessage)
            Cells(RowChecker, UserInputEndingSampleColumn).Select
            
            Exit Sub
            
        End If
        
    End If
    
Next RowChecker

' Validation To Make Sure Maximum Sample Number Is Not Exceeded:
If StartingSampleNumber + FinalUserInputNumber > MaximumFuelSampleNumber Then

    ErrorMessage.Value = "ERROR: Maximum Sample Number of " + CStr(MaximumFuelSampleNumber) + " Exceeded! Please check inputs and try again."
    MsgBox (ErrorMessage)

    Exit Sub

End If

SuccessMessage.Value = "Parameter Validation Complete - Sample Generation In Progress"

'----------------------------------------------------------------------------------------------------------------------SAMPLE---GENERATOR-----------------------------------------------------------------------------------------------------------------------

' OVERVIEW
' The sample generator section initiates once all inputs have been validated above.  It first selects the month for the Month Filter and then
' moves forward with selecting the Microbial Growth Statement as a function of the generation date.  Finally, it moves forward with creating
' samples with their associated packages as per the user input.

' I - Populates Month Column based on user's Month input in Sample Creation:
Select Case SampleMonth
    
    Case 1: MonthEntry = "M01 - January"
    Case 2: MonthEntry = "M02 - February"
    Case 3: MonthEntry = "M03 - March"
    Case 4: MonthEntry = "M04 - April"
    Case 5: MonthEntry = "M05 - May"
    Case 6: MonthEntry = "M06 - June"
    Case 7: MonthEntry = "M07 - July"
    Case 8: MonthEntry = "M08 - August"
    Case 9: MonthEntry = "M09 - September"
    Case 10: MonthEntry = "M10 - October"
    Case 11: MonthEntry = "M11 - November"
    Case 12: MonthEntry = "M12 - December"
       
End Select


' II - The following code generates the due date of Microbial Samples set up for that day:
TodaysDate = Weekday(Date, vbSunday)   'Sets Sunday = 1, Monday = 2, etc. for Case Below

Select Case TodaysDate

    Case 1: MicrobialCheckDay = " - Wed"
    Case 2: MicrobialCheckDay = " - Thu"
    Case 3: MicrobialCheckDay = " - Fri"
    Case 4: MicrobialCheckDay = " - Sat"
    Case 5: MicrobialCheckDay = " - Sun"
    Case 6: MicrobialCheckDay = " - Mon"
    Case 7: MicrobialCheckDay = " - Tue"

End Select

MicrobialCheckStatement = "Set" + CStr(MicrobialCheckDay)


' III - This portion determines how the samples will be labeled to conform as MMDDYY-SSS:
ExtraZeroMonth = IIf(SampleMonth < 10, CStr(0), "") 'MM
ExtraZeroDay = IIf(SampleDay < 10, CStr(0), "")         'DD
ExtraZeroYear = IIf(SampleYear < 10, CStr(0), "")       'YY

' Which leads to:
MMDDYY = CStr(ExtraZeroMonth) + CStr(SampleMonth) + CStr(ExtraZeroDay) + CStr(SampleDay) + CStr(ExtraZeroYear) + CStr(SampleYear)
Dash = CStr("-")

' Correction Factors Used In String In Loop Below:
LessThan100 = CStr(0)
LessThan010 = CStr(0) + CStr(0)

' IV - The following Do-While loop runs indefinitely until it jumps to a user input row where the ending sample is blank:
CurrentSampleNumber = StartingSampleNumber
SamplesGenerated = 0

Do While EndingSampleNumber <> ""
                
    ' Formats Sample Numbers (SSS) To Always Be 3 Digits (001 to 999):
    SampleIDExtraZeroes = IIf(CurrentSampleNumber < SampleTen, LessThan010, IIf(CurrentSampleNumber < SampleOneHundred, LessThan100, ""))
    SSS = SampleIDExtraZeroes + CStr(CurrentSampleNumber)
    
    ' Setters For Lab Data Sheet:
    Set AcidNumber = LabData.Cells(SampleRow, AcidNumberColumn)
    Set API = LabData.Cells(SampleRow, APIColumn)
    Set CetaneIndex = LabData.Cells(SampleRow, CetaneIndexColumn)
    Set CloudPoint = LabData.Cells(SampleRow, CloudPointColumn)
    Set CopperStripCorrosion = LabData.Cells(SampleRow, CopperStripCorrosionColumn)
    Set Distillation010 = LabData.Cells(SampleRow, Distillation010Column)
    Set Distillation050 = LabData.Cells(SampleRow, Distillation050Column)
    Set Distillation090 = LabData.Cells(SampleRow, Distillation090Column)
    Set DistillationFBP = LabData.Cells(SampleRow, DistillationFBPColumn)
    Set DistillationIBP = LabData.Cells(SampleRow, DistillationIBPColumn)
    Set FilterPatch = LabData.Cells(SampleRow, FilterPatchColumn)
    Set FlashPoint = LabData.Cells(SampleRow, FlashPointColumn)
    Set FuelBlendRatio = LabData.Cells(SampleRow, FuelBlendRatioColumn)
    Set FuelTypeOutput = LabData.Cells(SampleRow, FuelTypeColumn)
    Set MicrobialGrowth = LabData.Cells(SampleRow, MicrobialGrowthColumn)
    Set MonthTag = LabData.Cells(SampleRow, MonthTagColumn)
    Set FuelPackageOutput = LabData.Cells(SampleRow, FuelPackageColumn)
    Set ParticulateContamination = LabData.Cells(SampleRow, ParticulateContaminationColumn)
    Set PourPoint = LabData.Cells(SampleRow, PourPointColumn)
    Set RelativeAbsorbance = LabData.Cells(SampleRow, RelativeAbsorbanceColumn)
    Set RushPriorityOutput = LabData.Cells(SampleRow, RushPriorityColumn)
    Set SampleIDMirror = LabData.Cells(SampleRow, SampleIDMirrorColumn)
    Set SampleIDOutput = LabData.Cells(SampleRow, SampleIDColumn)
    Set SampleStatus = LabData.Cells(SampleRow, SampleStatusColumn)
    Set Sulfur = LabData.Cells(SampleRow, SulfurColumn)
    Set ThermalStability = LabData.Cells(SampleRow, ThermalStabilityColumn)
    Set ViscosityNumber = LabData.Cells(SampleRow, ViscosityNumberColumn)
    Set VisualOpacity = LabData.Cells(SampleRow, VisualOpacityColumn)
    Set VisualParticles = LabData.Cells(SampleRow, VisualParticlesColumn)
    Set VisualPhase = LabData.Cells(SampleRow, VisualPhaseColumn)
    Set VisualSediment = LabData.Cells(SampleRow, VisualSedimentColumn)
    Set WaterAndSediment = LabData.Cells(SampleRow, WaterAndSedimentColumn)
    Set WaterByKF = LabData.Cells(SampleRow, WaterByKFColumn)
    
    ' Primary Reference Points:
    FuelTypeOutput.Value = FuelType
    MonthTag.Value = MonthEntry
    FuelPackageOutput.Value = FuelPackage
    RushPriorityOutput.Value = RushPriority
    SampleIDOutput.Value = MMDDYY + Dash + SSS
    SampleIDMirror.Value = SampleIDOutput

    ' Reference Formulae For Bookwriting:
    AcidNumber.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[Acid '#])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[Acid '#]))))"
    API.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[API])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[API]))))"
    CetaneIndex.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[CI])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[CI]))))"
    CloudPoint.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[CP])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[CP]))))"
    CopperStripCorrosion.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[Cu])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[Cu]))))"
    Distillation010.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[10%])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[10%]))))"
    Distillation050.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[50%])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[50%]))))"
    Distillation090.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[90%])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[90%]))))"
    DistillationFBP.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[FBP])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[FBP]))))"
    DistillationIBP.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[IBP])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[IBP]))))"
    FilterPatch.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[FiltPat])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[FiltPat]))))"
    FlashPoint.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[FP])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[FP]))))"
    FuelBlendRatio.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[FBR])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[FBR]))))"
    MicrobialGrowth.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[MG])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[MG]))))"
    ParticulateContamination.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[PC])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[PC]))))"
    PourPoint.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[PP])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[PP]))))"
    RelativeAbsorbance.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[RA])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[RA]))))"
    SampleStatus.Formula = "=IFS([@FBR] = ""INFO!"", ""Pending"", [@CP] = ""-"", ""Pending"", [@PP] = ""-"", ""Pending"", [@FiltPat] = ""Pending"", ""Pending"", [@FP] < 2, ""Pending"", [@CI] = ""INFO!"", ""Pending"", COUNTIF(Table23[@[FBP]:[Acid '#]], """") > 0, ""Pending"", COUNTA(Table23[@[FBP]:[Acid '#]]) = 26, ""Complete"")"
    Sulfur.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[S])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[S]))))"
    ThermalStability.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[TS])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[TS]))))"
    ViscosityNumber.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[Visc '#])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[Visc '#]))))"
    VisualOpacity.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[Opacity])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[Opacity]))))"
    VisualParticles.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[Particles])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[Particles]))))"
    VisualPhase.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[Phase Sep.])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[Phase Sep.]))))"
    VisualSediment.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[Sediment])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[Sediment]))))"
    WaterAndSediment.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[W+S])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[W+S]))))"
    WaterByKF.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[KF])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[KF]))))"

    ' Compresses Cells To Keep Spreadsheet Running Smoothly As It Grows:
    AcidNumber.Value = IIf(AcidNumber = "XXX", "XXX", "")
    API.Value = IIf(API = "XXX", "XXX", "")
    CetaneIndex.Formula = IIf(CetaneIndex = "XXX", "XXX", "=IF([API] = """", ""INFO!"", IF([50%] = """", ""INFO!"", ROUND(-420.34+0.016*[API]^2+0.192*[API]*LOG([[50%]],10)+65.01*LOG([[50%]],10)^2+-0.0001809*[[50%]]^2,1)))")
    CloudPoint.Value = IIf(CloudPoint = "XXX", "XXX", "-")
    CopperStripCorrosion.Value = IIf(CopperStripCorrosion = "XXX", "XXX", "")
    Distillation010.Value = IIf(Distillation010 = "XXX", "XXX", "")
    Distillation050.Value = IIf(Distillation050 = "XXX", "XXX", "")
    Distillation090.Value = IIf(Distillation090 = "XXX", "XXX", "")
    DistillationFBP.Value = IIf(DistillationFBP = "XXX", "XXX", "")
    DistillationIBP.Value = IIf(DistillationIBP = "XXX", "XXX", "")
    FilterPatch.Value = IIf(FilterPatch = "XXX", "XXX", "Pending")
    FlashPoint.Value = IIf(FlashPoint = "XXX", "XXX", 1)
    FuelBlendRatio.Formula = IIf(FuelBlendRatio = "XXX", "XXX", "=IF([RA] = """", ""INFO!"", ROUND(IF([RA]<0.1344, N/A, 10^-6*[@RA]^3-0.0001*[RA]^2+0.0978*[RA]-0.1344), 1))")
    MicrobialGrowth.Value = IIf(MicrobialGrowth = "XXX", "XXX", MicrobialCheckStatement)
    ParticulateContamination.Value = IIf(ParticulateContamination = "XXX", "XXX", "")
    PourPoint.Value = IIf(PourPoint = "XXX", "XXX", "-")
    RelativeAbsorbance.Value = IIf(RelativeAbsorbance = "XXX", "XXX", "")
    Sulfur.Value = IIf(Sulfur = "XXX", "XXX", "")
    ThermalStability.Value = IIf(ThermalStability = "XXX", "XXX", "")
    ViscosityNumber.Value = IIf(ViscosityNumber = "XXX", "XXX", "")
    VisualOpacity.Value = IIf(VisualOpacity = "XXX", "XXX", "")
    VisualParticles.Value = IIf(VisualParticles = "XXX", "XXX", "")
    VisualPhase.Value = IIf(VisualPhase = "XXX", "XXX", "")
    VisualSediment.Value = IIf(VisualSediment = "XXX", "XXX", "")
    WaterAndSediment.Value = IIf(WaterAndSediment = "XXX", "XXX", "")
    WaterByKF.Value = IIf(WaterByKF = "XXX", "XXX", "")

    ' After Sample Row Is Populated, The Sample Number Goes Up:
    CurrentSampleNumber = CurrentSampleNumber + 1
    
    ' If the next sample number goes beyond the last sample of the current line, then:
    If (CurrentSampleNumber > EndingSampleNumber) Then
           
        ' Go to the next row...
        UserInputRow = UserInputRow + 1
        
        '...and update references for inputs for the next set of samples in the row below:
        Set StartingSampleNumber = SampleCreation.Cells(UserInputRow, AutoGeneratedStartingSampleColumn)
        Set EndingSampleNumber = SampleCreation.Cells(UserInputRow, UserInputEndingSampleColumn)
        Set FuelType = SampleCreation.Cells(UserInputRow, UserInputFuelTypeColumn)
        Set RushPriority = SampleCreation.Cells(UserInputRow, UserInputRushPriorityColumn)
        Set FuelPackage = SampleCreation.Cells(UserInputRow, UserInputPackageColumn)
           
    End If
    
    'Keeps track of samples generated for later code:
    SamplesGenerated = SamplesGenerated + 1
    SampleRow = SampleRow + 1

Loop

'---------------------------------------------------------------------------------------------------------------------QUICK---AUTO-FORMATTING---------------------------------------------------------------------------------------------------------------------

' OVERVIEW
' This section quickly readjusts prefilled data to make the sheet ready for the printing process. It scans each row that was generated.

' Reinitializes sample row so it only scans samples that were generated:
SampleRow = InitialSampleRow

Set SampleIDOutput = LabData.Cells(SampleRow, SampleIDColumn)
Set FlashPoint = LabData.Cells(SampleRow, FlashPointColumn)
Set CloudPoint = LabData.Cells(SampleRow, CloudPointColumn)
Set PourPoint = LabData.Cells(SampleRow, PourPointColumn)

' Realigns Cells According To Management Preferences:
While SampleIDOutput <> ""
    
    If FlashPoint = 1 Then
        FlashPoint.HorizontalAlignment = xlLeft
    End If
    
    If CloudPoint = "-" Then
        CloudPoint.HorizontalAlignment = xlLeft
    End If
    
    If PourPoint = "-" Then
        PourPoint.HorizontalAlignment = xlLeft
    End If
    
    SampleRow = SampleRow + 1
    
    Set SampleIDOutput = LabData.Cells(SampleRow, SampleIDColumn)
    Set FlashPoint = LabData.Cells(SampleRow, FlashPointColumn)
    Set CloudPoint = LabData.Cells(SampleRow, CloudPointColumn)
    Set PourPoint = LabData.Cells(SampleRow, PourPointColumn)

Wend

'------------------------------------------------------------------------------------------------------------------HIDE---AND---FILTER---COLUMNS-------------------------------------------------------------------------------------------------------------------

' OVERVIEW
' This section looks at samples and their data required to prepare the sheet for printing using the least space possible.  It starts by
' hiding all rows that are not associated with the samples generated and the hides the columns where everything is blacked out
' to optimize print space.

' Part I - Filter Rows (for rows that don't correspond to samples generated):
Call UnhideEverything
LabData.Cells(LabDataHeaderRow, SampleIDColumn).AutoFilter Field:=1, Criteria1:="=*" & MMDDYY & "*", Operator:=xlAnd

' Part II - Filter Columns:
SampleRow = InitialSampleRow
HideCounter = 0

' Determines Whether To Hide Rush Column:
Set SampleIDOutput = LabData.Cells(SampleRow, SampleIDColumn)
Set CellInQuestion = LabData.Cells(SampleRow, RushPriorityColumn)

While SampleIDOutput <> ""
    
    If CellInQuestion = "No" Then HideCounter = HideCounter + 1
    
    SampleRow = SampleRow + 1
    Set SampleIDOutput = LabData.Cells(SampleRow, SampleIDColumn)
    Set CellInQuestion = LabData.Cells(SampleRow, RushPriorityColumn)

Wend

If HideCounter = SamplesGenerated Then LabData.Columns(RushPriorityColumn).EntireColumn.Hidden = True


' Determines Whether To Hide Test Columns:
For ColumnAnalyzed = FirstColumnToAnalyze To LastColumnToAnalyze
    
    ' These Columns NEVER Need To Be Printed Since They Are Digital Calculations:
    If ColumnAnalyzed = CetaneIndexColumn Or ColumnAnalyzed = FuelBlendRatioColumn Then

        LabData.Columns(ColumnAnalyzed).EntireColumn.Hidden = True
        ColumnAnalyzed = ColumnAnalyzed + 1

    End If
    
    SampleRow = InitialSampleRow
    HideCounter = 0
    
    Set SampleIDOutput = LabData.Cells(SampleRow, SampleIDColumn)
    Set CellInQuestion = LabData.Cells(SampleRow, ColumnAnalyzed)
    
    ' Counter To Determine Whether Samples Generated Don't Get Specified Test:
    While SampleIDOutput <> ""
        
        If CellInQuestion = "XXX" Then HideCounter = HideCounter + 1
        
        SampleRow = SampleRow + 1
        
        Set SampleIDOutput = LabData.Cells(SampleRow, SampleIDColumn)
        Set CellInQuestion = LabData.Cells(SampleRow, ColumnAnalyzed)
    
    Wend
    
    ' Test To Determine Whether To Hide Column:
    If HideCounter = SamplesGenerated Then LabData.Columns(ColumnAnalyzed).EntireColumn.Hidden = True

Next ColumnAnalyzed

' These Columns Are Hidden No Matter What:
LabData.Columns(DistillationFBPColumn).EntireColumn.Hidden = True
LabData.Columns(Distillation050Column).EntireColumn.Hidden = True
LabData.Columns(Distillation010Column).EntireColumn.Hidden = True
LabData.Columns(DistillationIBPColumn).EntireColumn.Hidden = True
LabData.Columns(MonthTagColumn).EntireColumn.Hidden = True
LabData.Columns(SampleStatusColumn).EntireColumn.Hidden = True

'----------------------------------------------------------------------------------------------------------------------------------FINISHLINE----------------------------------------------------------------------------------------------------------------------------------

' Links Particulate Contamination, Thermal Stability, and Viscosity to Calculation Sheets:
Call PopulateCalculationSheets

' Message Letting User Know Their Samples Were Created Successfully:
SuccessMessage.Value = "All Samples Were Created Successfully!"
Response = MsgBox("All Samples Created Successfully! Do You Wish To Go To Lab Data?", vbQuestion + vbYesNo + vbDefaultButton2, "Lab Data")

' Takes User To Table To Be Printed Automatically:
If Response = vbYes Then LabData.Select: LabData.Cells(InitialSampleRow, SampleIDColumn).Select

End Sub


Sub CoolantSampleGenerator_Button()

' OBJECTIVE
' This program generates the coolant samples in the Coolant Data tab with the fewest inputs required.  It starts with validating that inputs
' are entered correctly and then proceeds to generate the samples according to what is entered on the sample generation interface.

Set CoolantData = ThisWorkbook.Worksheets("Coolant Data")
Set SampleCreation = ThisWorkbook.Worksheets("Sample Creation")

Dim CoolantSampleRow As Integer
Dim PreexistingCoolantSamples As Integer

Dim NewCoolantsReceived As Object
Dim ExtraNumbers As String
Dim CoolantSampleID As Object
Dim GlycolPercentage As Object
Dim CoolantSampleStatus As Object
Dim FirstCoolantSample As Integer
Dim CoolantSampleStartingRow As Integer
Dim LessThan300 As String
Dim LessThan210 As String

CoolantSampleRow = CoolantDataHeaderRow + 1

Set CoolantSampleID = CoolantData.Cells(CoolantSampleRow, CoolantSampleIDColumn)

' Determines The First Available Blank SampleID Cell:
While CoolantSampleID <> ""
    
    CoolantSampleRow = CoolantSampleRow + 1
    Set CoolantSampleID = CoolantData.Cells(CoolantSampleRow, CoolantSampleIDColumn)
    
Wend

' User Reference Points for Tests:
Set SampleMonth = SampleCreation.Cells(SampleMonthRow, InitialUserInputColumn)
Set SampleDay = SampleCreation.Cells(SampleDayRow, InitialUserInputColumn)
Set SampleYear = SampleCreation.Cells(SampleYearRow, InitialUserInputColumn)

Set StartingSampleNumber = SampleCreation.Cells(FirstSampleInputRow, InitialUserInputColumn)
Set NewCoolantsReceived = SampleCreation.Cells(CoolantInputRow, InitialUserInputColumn)
FirstCoolantSample = StartingSampleNumber

' Feedback Messages
Set ErrorMessage = SampleCreation.Cells(ErrorMessageRow, ErrorMessageColumn)
Set SuccessMessage = SampleCreation.Cells(SuccessMessageRow, SuccessMessageColumn)

' Initialize Messages on Interface:
ErrorMessage.Value = ""
SuccessMessage.Value = ""

'------------------------------------------------------------------------------------------------------------------------VALIDATION---TESTS------------------------------------------------------------------------------------------------------------------------

' V1 - Sample Number Parameters

' I - Validates That The Date Was Entered Correctly:
Call DateValidationCheck

' Coolant Validation Tests To Ensure Everything is Entered Properly:
If NewCoolantsReceived = "" Then ErrorMessage.Value = "No Coolant Samples Entered! Please enter ending sample numbers and try again. (B2)": Exit Sub
If StartingSampleNumber > MaximumCoolantSampleNumber Or StartingSampleNumber < MinimumCoolantSampleNumber Or StartingSampleNumber = "" Or Not IsNumeric(StartingSampleNumber) Then ErrorMessage.Value = "First Sample Not Valid (B3)": Exit Sub
If StartingSampleNumber > NewCoolantsReceived Then ErrorMessage.Value = "First Coolant Sample Is Higher Than Last! (B2/B3)": Exit Sub
If NewCoolantsReceived + StartingSampleNumber > MaximumCoolantSampleNumber Then ErrorMessage.Value = "Number of coolants to be created exceeds sample 299!": Exit Sub

SuccessMessage.Value = "Parameter Validation Complete - Sample Generation In Progress"

' Conforms Sample Labeling To Create MMDDYY-SSS:
ExtraZeroMonth = IIf(SampleMonth < 10, CStr(0), "")  'MM
ExtraZeroDay = IIf(SampleDay < 10, CStr(0), "")      'DD
ExtraZeroYear = IIf(SampleYear < 10, CStr(0), "")    'YY

MMDDYY = CStr(ExtraZeroMonth) + CStr(SampleMonth) + CStr(ExtraZeroDay) + CStr(SampleDay) + CStr(ExtraZeroYear) + CStr(SampleYear)
Dash = CStr("-")
LessThan300 = CStr(20)
LessThan210 = CStr(2)

'----------------------------------------------------------------------------------------------------------------------SAMPLE---GENERATOR-----------------------------------------------------------------------------------------------------------------------

Do While FirstCoolantSample <= NewCoolantsReceived

    ' Formats Sample IDs to Be MMDDYYY-SSS, Where SSS = (201 to 299)
    ExtraNumbers = IIf(FirstCoolantSample < SampleTen, LessThan300, LessThan210)
    SSS = ExtraNumbers + CStr(FirstCoolantSample)
    
    ' Coolant SampleID Creation:
    Set CoolantSampleID = CoolantData.Cells(CoolantSampleRow, CoolantSampleIDColumn)
    CoolantSampleID.Value = MMDDYY + Dash + SSS   'SampleID
    
    ' Formats Glycol As A Function of Freeze Point:
    Set GlycolPercentage = CoolantData.Cells(CoolantSampleRow, CoolantGlycolPercentageColumn)
    GlycolPercentage.Formula = "=IF([@Freeze] = """", ""INFO!"", XLOOKUP([@Freeze], Table1[Freeze Point], Table1[Ethylene Glycol]))"
    
    ' Creates A Status Cell For Lab Management:
    Set CoolantSampleStatus = CoolantData.Cells(CoolantSampleRow, CoolantSampleStatusColumn)
    CoolantSampleStatus.Formula = "=IF(COUNTA(Table3[@[TDS]:[Nitrite]]) < 8, ""Pending"", ""Complete"")"

    ' Adjusts Numbers Accordingly For Next Sample:
    CoolantSampleRow = CoolantSampleRow + 1
    FirstCoolantSample = FirstCoolantSample + 1

Loop

'----------------------------------------------------------------------------------------------------------------------SUCCESS---MESSAGE-----------------------------------------------------------------------------------------------------------------------

SuccessMessage.Value = "Coolants Successfully Generated!"

' Gives User Option To Jump Directly To Coolant Data To Confirm:
Response = MsgBox("All Coolant Samples Created Successfully! Do You Wish To Go To Coolants?", vbQuestion + vbYesNo + vbDefaultButton2, "Coolant Samples")

If Response = vbYes Then CoolantData.Select: CoolantData.Cells(CoolantSampleRow, CoolantSampleIDColumn).Select

End Sub


Sub DateValidationCheck()

' OBJECTIVE:
' This subroutine scans the month, day, and year entered in the user interface to ensure that valid numbers have been entered.
' It is used both in generating fuel samples and coolant samples in the laboratory.  This is NOT a standalone program.

Set SampleCreation = ThisWorkbook.Worksheets("Sample Creation")

' Feedback Messages:
Set ErrorMessage = SampleCreation.Cells(ErrorMessageRow, ErrorMessageColumn)

' User Reference Points for Tests:
Set SampleMonth = SampleCreation.Cells(SampleMonthRow, InitialUserInputColumn)
Set SampleDay = SampleCreation.Cells(SampleDayRow, InitialUserInputColumn)
Set SampleYear = SampleCreation.Cells(SampleYearRow, InitialUserInputColumn)

' Month Needs To Be An Integer Between 1-12 or Program Won't Execute:
If SampleMonth > 12 Or SampleMonth < 1 Or SampleMonth = "" Or Not IsNumeric(SampleMonth) Then

    ErrorMessage.Value = "Please Enter a Valid Month (1-12) (B4)"
    SampleCreation.Cells(SampleMonthRow, InitialUserInputColumn).Select
    End

End If

' Year Needs To Be An Integer Between 0-99 or Program Won't Execute:
If SampleYear > 99 Or SampleYear < 0 Or SampleYear = "" Or Not IsNumeric(SampleYear) Then

    ErrorMessage.Value = "Please Enter a Valid Year (1-99) (B6)"
    SampleCreation.Cells(SampleYearRow, InitialUserInputColumn).Select
    End

End If

' Used For Leap Year Check:
SampleYearFourDigits = SampleYearCorrectionFactor + SampleYear

' Validates The Final Day of Month Based on Month/Year
If SampleMonth = 2 Then

    LastDayOfMonth = IIf(SampleYearFourDigits Mod 4 = 0 And SampleYearFourDigits Mod 100 <> 0 Or SampleYearFourDigits Mod 400 = 0, 29, 28)  'Checks For Leap Year Conditions

ElseIf SampleMonth = 4 Or SampleMonth = 6 Or SampleMonth = 9 Or SampleMonth = 11 Then

    LastDayOfMonth = 30

Else

    LastDayOfMonth = 31

End If

' Once LastDayOfMonth Is Determined, Other Parameters Are Checked For Validation:
If SampleDay > LastDayOfMonth Or SampleDay < 1 Or SampleDay = "" Or Not IsNumeric(SampleDay) Then

    ' The Error Message Will Be Displayed According To The Following Criterion:
    If LastDayOfMonth = 29 And SampleMonth = 2 Then

        ErrorMessage.Value = "For Month 2 On A Leap Year, Please Enter A Day Between 1-29"

    ElseIf LastDayOfMonth = 28 And SampleMonth = 2 Then

        ErrorMessage.Value = "For Month 2 On A Non-Leap Year, Please Enter A Day Between 1-28"

    Else

        ErrorMessage.Value = "For Month " + CStr(SampleMonth) + ", Please Enter A Day Between 1-" + CStr(LastDayOfMonth)

    End If

    SampleCreation.Cells(SampleMonthRow, InitialUserInputColumn).Select

    End

End If

End Sub


Sub DirtyKF()

' OBJECTIVE
' This program scans the appearance columns of each sample and assigns the Water by Karl Fischer Column as "DNR" to prevent
' unintentional sample runs that could potential damage the vessel and allow technicians to determine more quickly what
' samples require the test and which samples should be skipped.

Set LabData = ThisWorkbook.Worksheets("Lab Data")

Dim LabDataRow As Integer
Dim LastSampleRow As Integer
Dim DirtySampleCounter As Integer

DirtySampleCounter = 0
LabDataRow = LabDataHeaderRow + 1

Set SampleID = LabData.Cells(LabDataRow, SampleIDColumn)
Set WaterByKF = LabData.Cells(LabDataRow, WaterByKFColumn)
Set VisualPhase = LabData.Cells(LabDataRow, VisualPhaseColumn)


While SampleID <> ""
    
    If VisualPhase <> "None" And VisualPhase <> "XXX" And VisualPhase <> "" Then
    
        ' Subloop to not double count anything that may already have DNR:
        If WaterByKF = "" Then

            WaterByKF.Value = "DNR"

            DirtySampleCounter = DirtySampleCounter + 1    'To inform user that there were legitimate DNRs

        End If
        
    End If
    
    LabDataRow = LabDataRow + 1

    Set SampleID = LabData.Cells(LabDataRow, SampleIDColumn)
    Set WaterByKF = LabData.Cells(LabDataRow, WaterByKFColumn)
    Set VisualPhase = LabData.Cells(LabDataRow, VisualPhaseColumn)
    
Wend


' Ending Feedback For User

Select Case DirtySampleCounter

    Case DirtySampleCounter = 0: MsgBox "No Dirty Samples Found."
    Case DirtySampleCounter = 1: MsgBox "1 Dirty Sample Filled In!"
    Case Else: MsgBox CStr(DirtySampleCounter) + " Dirty Samples Filled In!"

End Select


End Sub

Sub AppearanceFill()

' OBJECTIVE
' This program automatically fills a user-determined number of cells as the following:
'   Opacity = Clear
'   Sediment = None
'   Phase = None
'   Particles = Trace
' These are the most common characteristics of fuel samples that have been processed in the history of the company. The user can change
' the anamolies of the visual characteristics manually as they are usually exceptions to the rule.


Dim VisualOpacity As Object
Dim VisualSediment As Object
Dim VisualPhase As Object
Dim VisualParticles As Object
Dim VisualAppearancesRemaining As Object

Dim LabDataRow As Integer
Dim VisualCells As Integer
Dim VisualAppearancesFilled As Integer
Dim VisualAppearancesToFill As Object

Set AutoTools = ThisWorkbook.Worksheets("AutoTools")
Set LabData = ThisWorkbook.Worksheets("Lab Data")

' Counters
VisualCells = 0                      'Individual Cells Filled Within Appearance
VisualAppearancesFilled = 0  'Entire Sample Row of Appearances

'Objects
Set VisualAppearancesToFill = AutoTools.Cells(VisualSamplesToFillRow, VisualSamplesFillColumn)
Set VisualAppearancesRemaining = AutoTools.Cells(VisualSamplesRemainingRow, VisualSamplesFillColumn)

LabDataRow = LabDataHeaderRow + 1

' Validation That Inputs Are Usable:
If VisualAppearancesToFill < 1 Or VisualAppearancesToFill > VisualAppearancesRemaining Or Not IsNumeric(VisualAppearancesToFill) Then
    MsgBox "Invalid Data Inputs! Please check input number and try again!"
    Exit Sub
End If

' Makes Sure User Is Intentionally Using Program:
Response = MsgBox("CAUTION: This feature should only be used after all nonstandard appearance properties are filled.  This will fill all blank cells as follows: Opacity = Clear, Sediment = None, Phase = None, and Particles = Trace.  Do you wish to proceed?", vbQuestion + vbYesNo + vbDefaultButton2, "Visual Appearance Autofill")

If Response = vbNo Then
    MsgBox "Appearance Autofill Terminated!"
    Exit Sub
End If

' Fills In Visual Appearances According To User Specifications:
While VisualAppearancesFilled < VisualAppearancesToFill
    
    Set VisualOpacity = LabData.Cells(LabDataRow, VisualOpacityColumn)
    Set VisualSediment = LabData.Cells(LabDataRow, VisualSedimentColumn)
    Set VisualPhase = LabData.Cells(LabDataRow, VisualPhaseColumn)
    Set VisualParticles = LabData.Cells(LabDataRow, VisualParticlesColumn)
    
    ' Checks to see if there are blanks first, and then analyzes each one individually and counts up accordingly:
    If VisualOpacity = "" Or VisualSediment = "" Or VisualPhase = "" Or VisualParticles = "" Then
        
        If VisualOpacity = "" Then
            VisualOpacity.Value = "Clear"
            VisualCells = VisualCells + 1
        End If
        
        If VisualSediment = "" Then
            VisualSediment.Value = "None"
            VisualCells = VisualCells + 1
        End If
        
        If VisualPhase = "" Then
            VisualPhase.Value = "None"
            VisualCells = VisualCells + 1
        End If
        
        If VisualParticles = "" Then
            VisualParticles.Value = "Trace"
            VisualCells = VisualCells + 1
        End If
        
        VisualAppearancesFilled = VisualAppearancesFilled + 1
        
    End If

    ' Goes To The Next Row:
    LabDataRow = LabDataRow + 1
    
Wend

MsgBox CStr(VisualCells) + " data points over " + CStr(VisualAppearancesFilled) + " samples have been autofilled!"

End Sub


Sub FillCopperStrips()

' OBJECTIVE
' This program fills in a user-selected amount of empty copper strip fields as "1a", since that is the result more than 99% of the time.
' The program scans for the first available blank cell within copper strips and fills them in until the number of copper strips the user
' wishes to fill is met.

Dim CopperStripCorrosion As Object
Dim CopperStripSamplesRan As Object
Dim CopperStripSamplesRemaining As Object

Dim CopperStripSamplesFilled As Integer
Dim LabDataRow As Integer

Set AutoTools = ThisWorkbook.Worksheets("AutoTools")
Set LabData = ThisWorkbook.Worksheets("Lab Data")

Set CopperStripSamplesRan = AutoTools.Cells(CopperStripCorrosionToFillRow, CopperStripCorrosionFillColumn)
Set CopperStripSamplesRemaining = AutoTools.Cells(CopperStripCorrosionRemainingRow, CopperStripCorrosionFillColumn)
CopperStripSamplesFilled = 0

LabDataRow = LabDataHeaderRow + 1

' Validation Test:
If CopperStripSamplesRan < 1 Or CopperStripSamplesRan > CopperStripSamplesRemaining Or Not IsNumeric(CopperStripSamplesRan) Then
    MsgBox "Unable to process autofill.  Please check input number and try again!"
    Exit Sub
End If


' Prompt To Ensure User Request Was Intentional:
Response = MsgBox("CAUTION: This feature will fill in the first " + CStr(CopperStripSamplesRan) + " as 1a.  Do you wish to proceed?", vbQuestion + vbYesNo + vbDefaultButton2, "Copper Strip Autofill")

If Response = vbNo Then
    MsgBox "CopperStrip Autofill Terminated!"
    Exit Sub
End If


' Routine to fill the blanks in as "1a".
While CopperStripSamplesFilled < CopperStripSamplesRan
    
    Set CopperStripCorrosion = LabData.Cells(LabDataRow, CopperStripCorrosionColumn)
    
    If CopperStripCorrosion = "" Then
        CopperStripCorrosion.Value = "1a"
        CopperStripSamplesFilled = CopperStripSamplesFilled + 1
    End If
    
    LabDataRow = LabDataRow + 1

Wend

MsgBox "The first " + CStr(CopperStripSamplesRan) + " samples have been filled as 1a."

End Sub


Sub FuelPackageReplacement()

' OBJECTIVE
' This program allows lab management and certified personnel to change the package of a sample quickly through selecting the Sample ID in question
' and the package it wishes to replace the current package with.  The program will give a prompt to ensure that the user is doing this action
' intentionally, as it will erase current data if they choose to continue.

Dim SampleIDReplacement As Object
Dim PackageReplacement As Object
Dim SampleID As Object
Dim CurrentPackage As Object
Dim OriginalMicrobialGrowthCheckStatement As String
Dim SampleRow As Integer

Set AutoTools = ThisWorkbook.Worksheets("AutoTools")
Set LabData = ThisWorkbook.Worksheets("Lab Data")

SampleRow = LabDataHeaderRow + 1

Set SampleIDReplacement = AutoTools.Cells(PackageReplacementSampleRow, PackageReplacementColumn)
Set PackageReplacement = AutoTools.Cells(PackageReplacementPackageRow, PackageReplacementColumn)

' Validation Test To Ensure User Inputs Are Sufficient:
If SampleIDReplacement = "" Or PackageReplacement = "" Then
    MsgBox "Please fill out Sample and Replacement Package!"
    Exit Sub
End If

' Scans The Lab Data To Return Sample-ID and Current Package:

Set SampleID = LabData.Cells(SampleRow, SampleIDColumn)

While SampleID <> ""

    If SampleID = SampleIDReplacement Then
        
        Set CurrentPackage = LabData.Cells(SampleRow, FuelPackageColumn)
        
        ' Prompt That Comes Up Only When Sample Is Found:
        Response = MsgBox("Found SampleID " + CStr(SampleID) + " on row " + CStr(LabDataRow) + ", with a current package of " + CStr(CurrentPackage) + ".  All data in this row will be overrided with data points corresponding to " + CStr(PackageReplacement) + ". Do you wish to proceed?", vbQuestion + vbYesNo + vbDefaultButton2, "Package Replacement")
            
            If Response = vbYes Then
                
                ' Replaces The Package Name So That Formulae Return The Right Values:
                CurrentPackage.Value = CStr(PackageReplacement)

                'Setters For Data Points:
                Set AcidNumber = LabData.Cells(SampleRow, AcidNumberColumn)
                Set API = LabData.Cells(SampleRow, APIColumn)
                Set CetaneIndex = LabData.Cells(SampleRow, CetaneIndexColumn)
                Set CloudPoint = LabData.Cells(SampleRow, CloudPointColumn)
                Set CopperStripCorrosion = LabData.Cells(SampleRow, CopperStripCorrosionColumn)
                Set Distillation010 = LabData.Cells(SampleRow, Distillation010Column)
                Set Distillation050 = LabData.Cells(SampleRow, Distillation050Column)
                Set Distillation090 = LabData.Cells(SampleRow, Distillation090Column)
                Set DistillationFBP = LabData.Cells(SampleRow, DistillationFBPColumn)
                Set DistillationIBP = LabData.Cells(SampleRow, DistillationIBPColumn)
                Set FilterPatch = LabData.Cells(SampleRow, FilterPatchColumn)
                Set FlashPoint = LabData.Cells(SampleRow, FlashPointColumn)
                Set FuelBlendRatio = LabData.Cells(SampleRow, FuelBlendRatioColumn)
                Set ParticulateContamination = LabData.Cells(SampleRow, ParticulateContaminationColumn)
                Set PourPoint = LabData.Cells(SampleRow, PourPointColumn)
                Set RelativeAbsorbance = LabData.Cells(SampleRow, RelativeAbsorbanceColumn)
                Set Sulfur = LabData.Cells(SampleRow, SulfurColumn)
                Set ThermalStability = LabData.Cells(SampleRow, ThermalStabilityColumn)
                Set ViscosityNumber = LabData.Cells(SampleRow, ViscosityNumberColumn)
                Set VisualOpacity = LabData.Cells(SampleRow, VisualOpacityColumn)
                Set VisualParticles = LabData.Cells(SampleRow, VisualParticlesColumn)
                Set VisualPhase = LabData.Cells(SampleRow, VisualPhaseColumn)
                Set VisualSediment = LabData.Cells(SampleRow, VisualSedimentColumn)
                Set WaterAndSediment = LabData.Cells(SampleRow, WaterAndSedimentColumn)
                Set WaterByKF = LabData.Cells(SampleRow, WaterByKFColumn)


                'Intermediate Formulae To Determine How Cells Are Populated As A Function Of Package:
                AcidNumber.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[Acid '#])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[Acid '#]))))"
                API.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[API])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[API]))))"
                CetaneIndex.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[CI])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[CI]))))"
                CloudPoint.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[CP])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[CP]))))"
                CopperStripCorrosion.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[Cu])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[Cu]))))"
                Distillation010.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[10%])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[10%]))))"
                Distillation050.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[50%])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[50%]))))"
                Distillation090.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[90%])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[90%]))))"
                DistillationFBP.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[FBP])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[FBP]))))"
                DistillationIBP.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[IBP])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[IBP]))))"
                FilterPatch.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[FiltPat])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[FiltPat]))))"
                FlashPoint.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[FP])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[FP]))))"
                FuelBlendRatio.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[FBR])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[FBR]))))"
                ParticulateContamination.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[PC])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[PC]))))"
                PourPoint.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[PP])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[PP]))))"
                RelativeAbsorbance.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[RA])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[RA]))))"
                Sulfur.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[S])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[S]))))"
                ThermalStability.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[TS])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[TS]))))"
                ViscosityNumber.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[Visc '#])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[Visc '#]))))"
                VisualOpacity.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[Opacity])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[Opacity]))))"
                VisualParticles.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[Particles])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[Particles]))))"
                VisualPhase.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[Phase Sep.])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[Phase Sep.]))))"
                VisualSediment.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[Sediment])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[Sediment]))))"
                WaterAndSediment.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[W+S])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[W+S]))))"
                WaterByKF.Formula = "=IF([@Package]="""","""", IF(XLOOKUP([@Package],Table52[Package], Table52[KF])=0, """", (XLOOKUP([@Package],Table52[Package], Table52[KF]))))"


                'Compressors To Keep Spreadsheet Fast:
                AcidNumber.Value = IIf(AcidNumber = "XXX", "XXX", "")
                API.Value = IIf(API = "XXX", "XXX", "")
                CetaneIndex.Formula = IIf(CetaneIndex = "XXX", "XXX", "=IF([API] = """", ""INFO!"", IF([50%] = """", ""INFO!"", ROUND(-420.34+0.016*[API]^2+0.192*[API]*LOG([[50%]],10)+65.01*LOG([[50%]],10)^2+-0.0001809*[[50%]]^2,1)))")
                CloudPoint.Value = IIf(CloudPoint = "XXX", "XXX", "-")
                CopperStripCorrosion.Value = IIf(CopperStripCorrosion = "XXX", "XXX", "")
                Distillation010.Value = IIf(Distillation010 = "XXX", "XXX", "")
                Distillation050.Value = IIf(Distillation050 = "XXX", "XXX", "")
                Distillation090.Value = IIf(Distillation090 = "XXX", "XXX", "")
                DistillationFBP.Value = IIf(DistillationFBP = "XXX", "XXX", "")
                DistillationIBP.Value = IIf(DistillationIBP = "XXX", "XXX", "")
                FilterPatch.Value = IIf(FilterPatch = "XXX", "XXX", "Pending")
                FlashPoint.Value = IIf(FlashPoint = "XXX", "XXX", 1)
                FuelBlendRatio.Formula = IIf(FuelBlendRatio = "XXX", "XXX", "=IF([RA] = """", ""INFO!"", ROUND(IF([RA]<0.1344, N/A, 10^-6*[@RA]^3-0.0001*[RA]^2+0.0978*[RA]-0.1344), 1))")
                ParticulateContamination.Value = IIf(ParticulateContamination = "XXX", "XXX", "")
                PourPoint.Value = IIf(PourPoint = "XXX", "XXX", "")
                RelativeAbsorbance.Value = IIf(RelativeAbsorbance = "XXX", "XXX", "")
                Sulfur.Value = IIf(Sulfur = "XXX", "XXX", "")
                ThermalStability.Value = IIf(ThermalStability = "XXX", "XXX", "")
                ViscosityNumber.Value = IIf(ViscosityNumber = "XXX", "XXX", "")
                VisualOpacity.Value = IIf(VisualOpacity = "XXX", "XXX", "")
                VisualParticles.Value = IIf(VisualParticles = "XXX", "XXX", "")
                VisualPhase.Value = IIf(VisualPhase = "XXX", "XXX", "")
                VisualSediment.Value = IIf(VisualSediment = "XXX", "XXX", "")
                WaterAndSediment.Value = IIf(WaterAndSediment = "XXX", "XXX", "")
                WaterByKF.Value = IIf(WaterByKF = "XXX", "XXX", "")

                ' Appears Once Entire Row Is Overwritten:
                MsgBox "Package Replacement for Sample " + SampleID + " was successful!"

            Else

                MsgBox "Package Replacement Terminated!"

            End If
        
        Exit Sub
    
    End If
    
    SampleRow = SampleRow + 1
    Set SampleID = LabData.Cells(SampleRow, SampleIDColumn)
        
Wend

End Sub

Sub FillPartialDistills()

' OBJECTIVE
' The following program fills in all distillations that have partial data points with did not run ("DNR").
' This subprogram fulfills two important aspects in lab management:
'   1 - Lets everyone know that a distillation was only run partially and that data points were not omitted unintentionally, and
'   2 - Allows management to know exactly how many distillation tests are on deck in the laboratory in the dashboard.


Dim DistillationFBP As Object
Dim Distillation090 As Object
Dim Distillation050 As Object
Dim Distillation010 As Object
Dim DistillationIBP As Object
Dim CetaneIndex As Object

Dim LabDataRow As Integer
Dim LastSampleRow As Integer
Dim DistillationSamplesFilled As Integer
Dim CetaneIndicesFilled As Integer

Set LabData = ThisWorkbook.Worksheets("Lab Data")

' Counters For The End:
DistillationSamplesFilled = 0
CetaneIndicesFilled = 0

LabDataRow = LabDataHeaderRow + 1

Set SampleID = LabData.Cells(LabDataRow, SampleIDColumn)
Set DistillationFBP = LabData.Cells(LabDataRow, DistillationFBPColumn)
Set Distillation090 = LabData.Cells(LabDataRow, Distillation090Column)
Set Distillation050 = LabData.Cells(LabDataRow, Distillation050Column)
Set Distillation010 = LabData.Cells(LabDataRow, Distillation010Column)
Set DistillationIBP = LabData.Cells(LabDataRow, DistillationIBPColumn)
Set CetaneIndex = LabData.Cells(LabDataRow, CetaneIndexColumn)

While SampleID <> ""

    ' Criteria To Determine Whether A Partial Distill Was Recorded:
    If IsNumeric(DistillationIBP) And DistillationIBP.Value <> "" Then
        
        ' Scans The Points From Highest TO Lowest:
        If DistillationFBP.Value = "" Then
            
            DistillationFBP.Value = "DNR"

            'Counter Goes Up By One Since There Was An Incomplete Detected:
            DistillationSamplesFilled = DistillationSamplesFilled + 1

            If Distillation090.Value = "" Then
                
                Distillation090.Value = "DNR"
            
                ' Has Multiple Consequences Since Cetane Index Needs 50% Distill:
                If Distillation050.Value = "" Then
                    
                    Distillation050.Value = "DNR"
                    CetaneIndex.Value = "CND"
                    CetaneIndicesFilled = CetaneIndicesFilled + 1
                    
                    ' Final Possible Empty Data Point:
                    If Distillation010.Value = "" Then Distillation010.Value = "DNR"

                End If

            End If
            
        End If
        
    End If

    LabDataRow = LabDataRow + 1

    Set SampleID = LabData.Cells(LabDataRow, SampleIDColumn)
    Set DistillationFBP = LabData.Cells(LabDataRow, DistillationFBPColumn)
    Set Distillation090 = LabData.Cells(LabDataRow, Distillation090Column)
    Set Distillation050 = LabData.Cells(LabDataRow, Distillation050Column)
    Set Distillation010 = LabData.Cells(LabDataRow, Distillation010Column)
    Set DistillationIBP = LabData.Cells(LabDataRow, DistillationIBPColumn)
    Set CetaneIndex = LabData.Cells(LabDataRow, CetaneIndexColumn)

Wend

Select Case DistillationSamplesFilled

    Case DistillationSamplesFilled > 1 And CetaneIndicesFilled > 1: MsgBox CStr(DistillationSamplesFilled) + " Distillation samples have been filled and " + CStr(CetaneIndicesFilled) + " Cetane Indices have been changed!"
    Case DistillationSamplesFilled > 1 And CetaneIndicesFilled = 1: MsgBox CStr(DistillationSamplesFilled) + " Distillation samples have been filled and " + CStr(CetaneIndicesFilled) + " Cetane Index has been changed!"
    Case DistillationSamplesFilled = 1 And CetaneIndicesFilled = 1: MsgBox CStr(DistillationSamplesFilled) + " Distillation sample has been filled and " + CStr(CetaneIndicesFilled) + " Cetane Index has been changed!"
    Case DistillationSamplesFilled > 1 And CetaneIndicesFilled = 0: MsgBox CStr(DistillationSamplesFilled) + " Distillation samples have been filled and no Cetane Indices were changed!"
    Case DistillationSamplesFilled = 1 And CetaneIndicesFilled = 0: MsgBox CStr(DistillationSamplesFilled) + " Distillation sample has been filled and no Cetane Indices were changed!"
    Case DistillationSamplesFilled = 0: MsgBox "No Distillation Samples were incomplete and no Cetane Indices were affected."

End Select


End Sub


Sub FillWaterAndSediments()

' OBJECTIVE
' The following program fills in a user-defined amount of blank Water & Sediment cells as "0.001" since that is one of the most common results of the test.
' The user can go back over what is filled and correct the values later to what is documented.

Dim WaterAndSediment As Object
Dim WaterAndSedimentSamplesToFill As Object
Dim WaterAndSedimentSamplesRemaining As Object

Dim WaterAndSedimentSamplesFilled As Integer
Dim LabDataRow As Integer

Set AutoTools = ThisWorkbook.Worksheets("AutoTools")
Set LabData = ThisWorkbook.Worksheets("Lab Data")

Set WaterAndSedimentSamplesToFill = AutoTools.Cells(WaterAndSedimentToFillRow, WaterAndSedimentFillColumn)
Set WaterAndSedimentSamplesRemaining = AutoTools.Cells(WaterAndSedimentRemainingRow, WaterAndSedimentFillColumn)

' Samples Filled Counter:
WaterAndSedimentSamplesFilled = 0

LabDataRow = LabDataHeaderRow + 1

' Validation Tests:
If WaterAndSedimentSamplesToFill < 1 Or WaterAndSedimentSamplesToFill > WaterAndSedimentSamplesRemaining Or Not IsNumeric(WaterAndSedimentSamplesToFill) Then
   
    MsgBox "Unable to process autofill.  Please check input number and try again!"
    Exit Sub

End If

' Checker To Ensure User Wishes To Proceed:
Response = MsgBox("CAUTION: This feature will fill in the first " + CStr(WaterAndSedimentSamplesToFill) + " blank Water And Sediment samples as 0.001.  Do you wish to proceed?", vbQuestion + vbYesNo + vbDefaultButton2, "Water & Sediment Autofill")

If Response = vbNo Then MsgBox "Water And Sediment Autofill Terminated!": Exit Sub

While WaterAndSedimentSamplesFilled < WaterAndSedimentSamplesToFill
    
    Set WaterAndSediment = LabData.Cells(LabDataRow, WaterAndSedimentColumn)
    
    If WaterAndSediment = "" Then

        WaterAndSediment.Value = "0.001"
        WaterAndSedimentSamplesFilled = WaterAndSedimentSamplesFilled + 1

    End If
    
    LabDataRow = LabDataRow + 1

Wend

MsgBox "The first " + CStr(WaterAndSedimentSamplesFilled) + " blank Water & Sediment samples have been filled as 0.001."

End Sub

Sub FillParticulateContaminationTrays()

' OBJECTIVE
' This program quickly scans the particulate contamination samples in calculation sheets and assigns their corresponding
' tray numbers - starting from 1 to a determined maximum.  Once the maximum is hit, the next tray number wraps back around
' to one.

Dim SampleRow As Integer
Dim TrayNumber As Integer
Dim TraysProduced As Integer
Dim PartContSampleID As Object
Dim PartContTrayNumber As Object

Set CalculationSheets = ThisWorkbook.Worksheets("Calculation Sheets")

' Initializations:
SampleRow = 5
TrayNumber = 1
TraysProduced = 0
Const MaximumTrayNumber As Integer = 30

Set PartContSampleID = CalculationSheets.Cells(SampleRow, PartContSampleIDColumn)
Set PartContTrayNumber = CalculationSheets.Cells(SampleRow, PartContTrayNumberColumn)

While PartContSampleID <> ""
    
    If PartContTrayNumber = "" Then
    
        PartContTrayNumber.Value = TrayNumber
        
        TrayNumber = TrayNumber + 1
        TraysProduced = TraysProduced + 1
        
        ' Cycles Tray Numbers Back To 1 For Another Batch:
        If TrayNumber > MaximumTrayNumber Then TrayNumber = 1
        
    End If
    
    ' Goes to the next row...
    SampleRow = SampleRow + 1
    
    ' ...and assigns Sample ID and Tray cells accordingly:
    Set PartContSampleID = CalculationSheets.Cells(SampleRow, PartContSampleIDColumn)
    Set PartContTrayNumber = CalculationSheets.Cells(SampleRow, PartContTrayNumberColumn)

Wend

Select Case TraysProduced

    Case TraysProduced = 0: MsgBox "No tray numbers were assigned!"
    Case TraysProduced = 1: MsgBox "Only one tray number was assigned!"
    Case Else: MsgBox CStr(TraysProduced) + " tray numbers have been assigned!"

End Select

End Sub


Sub PopulateCalculationSheets()

' OBJECTIVE
' This program scans all viscosity, thermal stability, and particulate contamination cells for blank cells and transfers their corresponding
' Sample IDs to the Calculation Sheets.  When a Sample ID is filled in a test's corresponding part of the calculation sheet, it fills relevant
' adjacent cells with calculation formulae that will generate answers in real time.  Finally, the program fills the blank cell to correspond with
' the calculation sheet directly so that the sample result is populated in realtime once all parameters and criteria are satisfied.
' This program can be used as a standalone but is also used in sample generation as a subroutine.

Dim LabSampleID As Object
Dim PartContSampleID As Object
Dim PartContVolume As Object
Dim PartContCalculation As Object
Dim ThermStabSampleID As Object
Dim ThermStabAverage As Object
Dim ThermStabCalculation As Object
Dim ViscSampleID As Object
Dim ViscTemperature As Object
Dim ViscTime As Object
Dim ViscCoefficient As Object
Dim ViscCalculation As Object

Dim LabDataSampleRow As Integer
Dim PartContFillCounter As Integer
Dim PartContSampleRow As Integer
Dim ThermStabFillCounter As Integer
Dim ThermStabSampleRow As Integer
Dim ViscFillCounter As Integer
Dim ViscSampleRow As Integer

Dim PartContStatement As String
Dim ThermStabStatement As String
Dim ViscStatement As String

Dim Viscosity As Object
Dim ThermalStability As Object
Dim ParticulateContamination As Object

Set CalculationSheets = ThisWorkbook.Worksheets("Calculation Sheets")
Set LabData = ThisWorkbook.Worksheets("Lab Data")

' Starting References
LabDataSampleRow = 2
PartContFillCounter = 0
ThermStabFillCounter = 0
ViscFillCounter = 0
Const PartContVolumeFill As Integer = 150
Const ViscTemperatureStandard As Integer = 40

' Initial Reference Points And Setters For While Loops:
PartContSampleRow = CalculationSheetsHeaderRow
ThermStabSampleRow = CalculationSheetsHeaderRow
ViscSampleRow = CalculationSheetsHeaderRow

Set PartContSampleID = CalculationSheets.Cells(PartContSampleRow, PartContSampleIDColumn)
Set ThermStabSampleID = CalculationSheets.Cells(ThermStabSampleRow, ThermStabSampleIDColumn)
Set ViscSampleID = CalculationSheets.Cells(ViscSampleRow, ViscSampleIDColumn)

' The Following Three While Loops Uniquely Determine The First Blank Row For Each Particular Test:
While PartContSampleID <> ""
    
    PartContSampleRow = PartContSampleRow + 1
    Set PartContSampleID = CalculationSheets.Cells(PartContSampleRow, PartContSampleIDColumn)

Wend

While ThermStabSampleID <> ""
    
    ThermStabSampleRow = ThermStabSampleRow + 1
    Set ThermStabSampleID = CalculationSheets.Cells(ThermStabSampleRow, ThermStabSampleIDColumn)
    
Wend

While ViscSampleID <> ""

    ViscSampleRow = ViscSampleRow + 1
    Set ViscSampleID = CalculationSheets.Cells(ViscSampleRow, ViscSampleIDColumn)

Wend

' Lab Data Setters:
Set LabSampleID = LabData.Cells(LabDataSampleRow, SampleIDColumn)
Set ParticulateContamination = LabData.Cells(LabDataSampleRow, ParticulateContaminationColumn)
Set ThermalStability = LabData.Cells(LabDataSampleRow, ThermalStabilityColumn)
Set Viscosity = LabData.Cells(LabDataSampleRow, ViscosityNumberColumn)


Set PartContVolume = CalculationSheets.Cells(PartContSampleRow, PartContSampleVolumeColumn)
Set PartContCalculation = CalculationSheets.Cells(PartContSampleRow, PartContCalculationColumn)

Set ThermStabAverage = CalculationSheets.Cells(ThermStabSampleRow, ThermStabAverageColumn)
Set ThermStabCalculation = CalculationSheets.Cells(ThermStabSampleRow, ThermStabCalculationColumn)


Set ViscTime = CalculationSheets.Cells(ViscSampleRow, ViscTimeColumn)
Set ViscTemperature = CalculationSheets.Cells(ViscSampleRow, ViscTemperatureColumn)
Set ViscCoefficient = CalculationSheets.Cells(ViscSampleRow, ViscCoefficientColumn)
Set ViscCalculation = CalculationSheets.Cells(ViscSampleRow, ViscCalculationColumn)

While LabSampleID <> ""

    ' Fill Particulate Contaminations In Calculation Sheets:
    If ParticulateContamination.Value = "" Then

        PartContSampleID.Value = LabSampleID
        PartContVolume.Value = PartContVolumeFill
        PartContCalculation.Formula = "=IF(COUNTA(Table11[@[TB]:[Vol (mL)]]) < 5, ""IP"", ROUND(2/3*[@[Vol (mL)]]/150*(([@TA]-[@TB])-([@CA]-[@CB])), 1))"
        
        ParticulateContamination.Formula = "=XLOOKUP([@[Sample ID]], Table11[Sample ID], Table11[Result])"
        
        PartContSampleRow = PartContSampleRow + 1
        PartContFillCounter = PartContFillCounter + 1
        
        ' Resets Reference to Cells Below:
        Set PartContSampleID = CalculationSheets.Cells(PartContSampleRow, PartContSampleIDColumn)
        Set PartContVolume = CalculationSheets.Cells(PartContSampleRow, PartContSampleVolumeColumn)
        Set PartContCalculation = CalculationSheets.Cells(PartContSampleRow, PartContCalculationColumn)

    End If
    
    
    ' Fill Thermal Stabilities In Calculation Sheets:
    If ThermalStability.Value = "" Then
    
        ThermStabSampleID.Value = LabSampleID
        ThermalStability.Formula = "=XLOOKUP([@[Sample ID]], Table13[Sample ID], Table13[Result])"
        
        ThermStabAverage.Formula = "=AVERAGE(Table13[@[Patch 1]:[Patch 2]])"
        ThermStabCalculation.Formula = "=IF(COUNTA(Table13[@[Sample ID]:[White Patch]]) < 5, ""IP"", [@Average]*100/[@[White Patch]])"
        
        ThermStabSampleRow = ThermStabSampleRow + 1
        ThermStabFillCounter = ThermStabFillCounter + 1
        
        ' Resets Reference to Cells Below:
        Set ThermStabSampleID = CalculationSheets.Cells(ThermStabSampleRow, ThermStabSampleIDColumn)
        Set ThermStabAverage = CalculationSheets.Cells(ThermStabSampleRow, ThermStabAverageColumn)
        Set ThermStabCalculation = CalculationSheets.Cells(ThermStabSampleRow, ThermStabCalculationColumn)
        
    End If
    
    ' Fill Viscosities In Calculation Sheets:
    If Viscosity.Value = "" Then
        
        ' Calculation Sheets:
        ViscSampleID.Value = LabSampleID
        ViscTime.Formula = "=60*[@Minutes]+[@Seconds]"
        ViscTemperature.Value = ViscTemperatureStandard
        ViscCoefficient.Formula = "=IFS([@Meter]=""T588"", -6*10^-7*[@Temperature]+0.0079, [@Meter]=""W411"", -5*10^-7*[@Temperature]+0.0069, [@Meter]=""W245"", 10^-6*[@Temperature]+0.009, [@Meter]="""",""METER!"")"
        ViscCalculation.Formula = "=IF([@Coefficient]=""METER!"", ""IP"", IF([@Time] = 0, ""IP"", [@Time]*[@Coefficient]))"
        
        ' Lab Data:
        Viscosity.Formula = "=XLOOKUP([@[Sample ID]], Table14[Sample ID], Table14[Result])"
        
        ViscSampleRow = ViscSampleRow + 1
        ViscFillCounter = ViscFillCounter + 1
        
        ' Resets Reference to Cells Below:
        Set ViscSampleID = CalculationSheets.Cells(ViscSampleRow, ViscSampleIDColumn)
        Set ViscTime = CalculationSheets.Cells(ViscSampleRow, ViscTimeColumn)
        Set ViscTemperature = CalculationSheets.Cells(ViscSampleRow, ViscTemperatureColumn)
        Set ViscCoefficient = CalculationSheets.Cells(ViscSampleRow, ViscCoefficientColumn)
        Set ViscCalculation = CalculationSheets.Cells(ViscSampleRow, ViscCalculationColumn)
    
    End If

    LabDataSampleRow = LabDataSampleRow + 1
    
    Set LabSampleID = LabData.Cells(LabDataSampleRow, SampleIDColumn)
    Set ParticulateContamination = LabData.Cells(LabDataSampleRow, ParticulateContaminationColumn)
    Set ThermalStability = LabData.Cells(LabDataSampleRow, ThermalStabilityColumn)
    Set Viscosity = LabData.Cells(LabDataSampleRow, ViscosityNumberColumn)

Wend

' The Following Selectors Help Give User Feedback On What Was Migrated:
Select Case PartContFillCounter

    Case 0: PartContStatement = "No Particulate Contaminations, "
    Case 1: PartContStatement = "1 Particulate Contamination, "
    Case Else: PartContStatement = CStr(PartContFillCounter) + " Particulate Contaminations, "

End Select

Select Case ThermStabFillCounter

    Case 0: ThermStabStatement = "no Thermal Stabilities, and "
    Case 1: ThermStabStatement = "1 Thermal Stability, and "
    Case Else: ThermStabStatement = CStr(ThermStabFillCounter) + " Thermal Stabilities, and "

End Select

Select Case ViscFillCounter

    Case 0: ViscStatement = "no Viscosities "
    Case 1: ViscStatement = "1 Viscosity "
    Case Else: ViscStatement = CStr(ViscFillCounter) + " Viscosities "

End Select

MsgBox PartContStatement + ThermStabStatement + ViscStatement + "were migrated to Calculation Sheets!"

End Sub


Sub UnhideEverything()

' OBJECTIVE:
' Unhide and unfilter everything so that everything can be seen easily instead of doing in manually.
' Used a subprogram when generated samples to ensure that columns that are seen are relevant.

Dim ColumnToProcess As Integer
Dim FirstColumnInLabData As Integer
Dim LastColumnInLabData As Integer

' Creating An Array Will Ensure Orthogonality In Case New Columns Are Inserted:
Dim LabDataColumnArray(1 To 33) As Integer
LabDataColumnArray(1) = AcidNumberColumn
LabDataColumnArray(2) = APIColumn
LabDataColumnArray(3) = CetaneIndexColumn
LabDataColumnArray(4) = CloudPointColumn
LabDataColumnArray(5) = CopperStripCorrosionColumn
LabDataColumnArray(6) = Distillation010Column
LabDataColumnArray(7) = Distillation050Column
LabDataColumnArray(8) = Distillation090Column
LabDataColumnArray(9) = DistillationFBPColumn
LabDataColumnArray(10) = DistillationIBPColumn
LabDataColumnArray(11) = FilterPatchColumn
LabDataColumnArray(12) = FlashPointColumn
LabDataColumnArray(13) = FuelBlendRatioColumn
LabDataColumnArray(14) = FuelPackageColumn
LabDataColumnArray(15) = FuelTypeColumn
LabDataColumnArray(16) = MicrobialGrowthColumn
LabDataColumnArray(17) = MonthTagColumn
LabDataColumnArray(18) = ParticulateContaminationColumn
LabDataColumnArray(19) = PourPointColumn
LabDataColumnArray(20) = RelativeAbsorbanceColumn
LabDataColumnArray(21) = RushPriorityColumn
LabDataColumnArray(22) = SampleIDColumn
LabDataColumnArray(23) = SampleIDMirrorColumn
LabDataColumnArray(24) = SampleStatusColumn
LabDataColumnArray(25) = SulfurColumn
LabDataColumnArray(26) = ThermalStabilityColumn
LabDataColumnArray(27) = ViscosityNumberColumn
LabDataColumnArray(28) = VisualOpacityColumn
LabDataColumnArray(29) = VisualParticlesColumn
LabDataColumnArray(30) = VisualPhaseColumn
LabDataColumnArray(31) = VisualSedimentColumn
LabDataColumnArray(32) = WaterAndSedimentColumn
LabDataColumnArray(33) = WaterByKFColumn

' Setters for Sheets:
Set SampleCreation = ThisWorkbook.Worksheets("Sample Creation")
Set LabData = ThisWorkbook.Worksheets("Lab Data")

' Setter for Interface:
Set SuccessMessage = SampleCreation.Cells(SuccessMessageRow, SuccessMessageColumn)

' Parameters In For Loop Determined By Array:
FirstColumnInLabData = WorksheetFunction.Min(LabDataColumnArray)
LastColumnInLabData = WorksheetFunction.Max(LabDataColumnArray)

For ColumnToProcess = FirstColumnInLabData To LastColumnInLabData

    With LabData

        .Cells(LabDataHeaderRow, ColumnToProcess).AutoFilter   'Unhide Rows Filtered
        .Columns(ColumnToProcess).EntireColumn.Hidden = False  'Unhide Column

    End With

Next ColumnToProcess

' Lets User Know Everything Is Visible:
SuccessMessage.Value = "All Lab Columns And Rows Are Visible!"

End Sub


Sub MicrobMonday()

SampleRow = LabDataHeaderRow + 1

Set LabData = ThisWorkbook.Worksheets("Lab Data")

Set SampleID = LabData.Cells(SampleRow, SampleIDColumn)
Set MicrobialGrowth = LabData.Cells(SampleRow, MicrobialGrowthColumn)

While SampleID <> ""

    If MicrobialGrowth = "Set - Mon" Then MicrobialGrowth.Value = "Negative"

    SampleRow = SampleRow + 1

    Set SampleID = LabData.Cells(SampleRow, SampleIDColumn)
    Set MicrobialGrowth = LabData.Cells(SampleRow, MicrobialGrowthColumn)

Wend

MsgBox "All Microbials Due Monday Set to Negative!"
    
End Sub

Sub MicrobTuesday()

Dim SampleRow As Integer
SampleRow = LabDataHeaderRow + 1

Set LabData = ThisWorkbook.Worksheets("Lab Data")

Set SampleID = LabData.Cells(SampleRow, SampleIDColumn)
Set MicrobialGrowth = LabData.Cells(SampleRow, MicrobialGrowthColumn)

While SampleID <> ""

    If MicrobialGrowth = "Set - Tue" Then MicrobialGrowth.Value = "Negative"

    SampleRow = SampleRow + 1

    Set SampleID = LabData.Cells(SampleRow, SampleIDColumn)
    Set MicrobialGrowth = LabData.Cells(SampleRow, MicrobialGrowthColumn)

Wend

MsgBox "All Microbials Due Tuesday Set to Negative!"
    
End Sub


Sub MicrobWednesday()

SampleRow = LabDataHeaderRow + 1

Set LabData = ThisWorkbook.Worksheets("Lab Data")

Set SampleID = LabData.Cells(SampleRow, SampleIDColumn)
Set MicrobialGrowth = LabData.Cells(SampleRow, MicrobialGrowthColumn)

While SampleID <> ""

    If MicrobialGrowth = "Set - Wed" Then MicrobialGrowth.Value = "Negative"

    SampleRow = SampleRow + 1

    Set SampleID = LabData.Cells(SampleRow, SampleIDColumn)
    Set MicrobialGrowth = LabData.Cells(SampleRow, MicrobialGrowthColumn)

Wend

MsgBox "All Microbials Due Wednesday Set to Negative!"
    
End Sub


Sub MicrobThursday()

SampleRow = LabDataHeaderRow + 1

Set LabData = ThisWorkbook.Worksheets("Lab Data")

Set SampleID = LabData.Cells(SampleRow, SampleIDColumn)
Set MicrobialGrowth = LabData.Cells(SampleRow, MicrobialGrowthColumn)

While SampleID <> ""

    If MicrobialGrowth = "Set - Thu" Then MicrobialGrowth.Value = "Negative"

    SampleRow = SampleRow + 1

    Set SampleID = LabData.Cells(SampleRow, SampleIDColumn)
    Set MicrobialGrowth = LabData.Cells(SampleRow, MicrobialGrowthColumn)

Wend

MsgBox "All Microbials Due Thu Set to Negative!"
    
End Sub


Sub MicrobFriday()

SampleRow = LabDataHeaderRow + 1

Set LabData = ThisWorkbook.Worksheets("Lab Data")

Set SampleID = LabData.Cells(SampleRow, SampleIDColumn)
Set MicrobialGrowth = LabData.Cells(SampleRow, MicrobialGrowthColumn)

While SampleID <> ""

    If MicrobialGrowth = "Set - Fri" Then MicrobialGrowth.Value = "Negative"

    SampleRow = SampleRow + 1

    Set SampleID = LabData.Cells(SampleRow, SampleIDColumn)
    Set MicrobialGrowth = LabData.Cells(SampleRow, MicrobialGrowthColumn)

Wend

MsgBox "All Microbials Due Friday Set to Negative!"
    
End Sub

Sub MicrobMondayMass()

SampleRow = LabDataHeaderRow + 1

Set LabData = ThisWorkbook.Worksheets("Lab Data")

Set SampleID = LabData.Cells(SampleRow, SampleIDColumn)
Set MicrobialGrowth = LabData.Cells(SampleRow, MicrobialGrowthColumn)

' Used If Microbial Samples Are Checked When There Are No Weekend Technicians:
While SampleID <> ""

    If MicrobialGrowth = "Set - Sat" Or MicrobialGrowth = "Set - Sun" Or MicrobialGrowth = "Set - Mon" Then MicrobialGrowth.Value = "Negative"

    SampleRow = SampleRow + 1

    Set SampleID = LabData.Cells(SampleRow, SampleIDColumn)
    Set MicrobialGrowth = LabData.Cells(SampleRow, MicrobialGrowthColumn)

Wend

MsgBox "All Microbials Due Saturday, Sunday, and Monday Set to Negative!"
    
End Sub