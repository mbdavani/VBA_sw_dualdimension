Attribute VB_Name = "Dual_Dimension_Units_and_"
' ******************************************************************************
' C:\Users\marcus.davani\AppData\Local\Temp\swx9292\Macro1.swb - macro recorded on 10/04/22 by marcus.davani
' ******************************************************************************
Dim swApp As Object

Dim Part As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long

Sub main()

Set swApp = Application.SldWorks

Set Part = swApp.ActiveDoc

'boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsLinearFractionDenominator, 0, 16)    'fraction in 1/16 of a inch
'boolstatus = Part.Extension.SetUserPreferenceToggle(swUserPreferenceToggle_e.swUnitsLinearRoundToNearestFraction, 0, 1)         'toggle to round to the nearest fraction
'
'

boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsDualLinear, swUserPreferenceOption_e.swDetailingNoOptionSpecified, swLengthUnit_e.swFEETINCHES)  'make dual dimension in ft and inches
'boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsDualLinearDecimalPlaces, 0, 0)          'make dual dimension 0 decimal place
boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsDualLinearFractionDenominator, 0, 16)    'fraction in 1/16 of a inch
boolstatus = Part.Extension.SetUserPreferenceToggle(swUserPreferenceToggle_e.swUnitsDualLinearRoundToNearestFraction, 0, 1)         'toggle to round to the nearest fraction
boolstatus = Part.Extension.SetUserPreferenceToggle(swUserPreferenceToggle_e.swUnitsDualLinearFeetAndInchesFormat, 1, 1)            'toggle to make dimension ft-inch (with - in between)
'boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDualDimPosition, swUserPreferenceOption_e.swDetailingDimension, 1)      'this statement is replaced by the last statement (using linear dimension position control instead)
boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDimFractionStyle, 0, 1)              'stack fraction
'
'
'
'
boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsDualLinearDecimalDisplay, 0, swFractionDisplay_e.swFRACTION)                     'execute/activate fraction display
boolstatus = Part.Extension.SetUserPreferenceToggle(swUserPreferenceToggle_e.swDoublePrimeMark, swUserPreferenceOption_e.swDetailingNoOptionSpecified, 0)                   'add double prime mark to inch dimension
boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingDualDimPosition, swUserPreferenceOption_e.swDetailingLinearDimension, 1)     'place dual dimension on right hand side of primary dimension
boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingLinearDimPrecision, swUserPreferenceOption_e.swDetailingDimension, 2)        ' Sets the decimal to Hundreths, or 2 decimal points

End Sub


