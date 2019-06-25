Attribute VB_Name = "Intensity"
Public Type IntensityVar
    Current(0 To 1000) As Single
    Voltage(0 To 1000) As Single
    Intensity(0 To 1000) As Single
    Date As Date
    NumberOfPoints As Integer
    SampleName As String * 250
    LowVoltage As Single
    HighVoltage As Single
    CurrentCompliance As Single
    VoltageCompliance As Single
    VoltageStep As Single
    SettlingTime As Single
    KeithleyFilter As Byte
    NewportFilter As Byte
    Attenuator As Byte
    Zero As Byte
    Sensitivity As Byte
    IntegrationTime As Byte
    Wavelength As Single
    DisplayCurrentLow As Single
    DisplayCurrentHigh As Single
    DisplayIntensityLow As Single
    DisplayIntensityHigh As Single
    CurrentAutorange As Byte
    CurrentLogLin As Byte
    IntensityLogLin As Byte
    IntensityAutorange As Byte
    DetectorModel As String * 20
End Type

Public Const disc = 1, GPIB0 = 0, GPIBSource = 16, GPIBNewport = 5, GPIBMagnetPSU = 3, GPIBGaussmeter = 12

Public Data As IntensityVar
Public VoltageStepValue As Single
Public SettlingTimeString As String
Public SettlingTimeValue As Single
Public OptionsFocus As Integer
Public RecordedOnString As String
Public Ticks As Integer, ScanType As Byte
Public IAxisOffset As Single, VAxisOffset As Single
Public LogLow As Single, LogHigh As Single
Public LLogLow As Single, LLogHigh As Single
Public Abort As Byte, Pause As Integer, RecordingData As Byte
Public SendToPrinter As Byte, DataSaved As Byte
Public GlobalPathName As String, GlobalFileName As String, IGlobalFileName As String
Public GPIBPresent As Integer, SourcePresent As Integer, NewportPresent As Integer
Public MagnetPSUPresent As Integer, GaussmeterPresent As Integer, NumRepeats As Integer
Public MagnetType As String, SMUType As String
Public BIvalues(500) As Single, Bfield(500) As Single, Vvalues(500) As Single
Public ITime(4096) As Single, IValue(4096) As Single
Public NumBvalues As Integer, NumVvalues As Integer, IntervalTimeValue As Single
Public SourceVoltageOn As Boolean, SourceVoltageZero As Boolean

Public Sub FindMagnetPSU()
    Dim AddressList(2) As Integer, status As Integer, GPIBmessage As String
    AddressList(1) = GPIBMagnetPSU
    AddressList(2) = NOADDR
FindMagnetPSUAgain:
    MagnetPSUPresent = False
    GPIBmessage = ""
    Do
        DoEvents
        EnableRemote GPIB0, AddressList
    Loop Until (ibsta <> &H8000)
    Call DevClear(GPIB0, GPIBMagnetPSU)
    Call ReadStatusByte(GPIB0, GPIBMagnetPSU, status)
    GPIBOut GPIBMagnetPSU, "*IDN?"
    Call ReadStatusByte(GPIB0, GPIBMagnetPSU, status)
    Do
        DoEvents
        GPIBmessage = GPIBMagnetPSUIn(GPIBMagnetPSU)
    Loop Until (ibsta <> &H8000)

    
    If Left$(GPIBmessage, 13) = "LSCI,MODEL642" Then
        MagnetPSUPresent = True
        MagnetType = "LAKESHORE"
        GPIBmessage = ""
        Call DevClear(GPIB0, GPIBMagnetPSU)
        Call ReadStatusByte(GPIB0, GPIBMagnetPSU, status)
'       Set default IEEE parameters
        GPIBOut GPIBMagnetPSU, "IEEE 0,0,3"
'       Set control to external
        GPIBOut GPIBMagnetPSU, "XPGM 1"
'        Set internal water to ON
        GPIBOut GPIBMagnetPSU, "INTWTR 1"
'   Get PSU internal temperature
        PSUTemperature = GetPSUTemperature
'
    ElseIf Left$(GPIBmessage, 18) = "KEPCO,BOP1KW 50-20" Then
'
        MagnetPSUPresent = True
        MagnetType = "KEPCO"
        GPIBmessage = ""
        Call DevClear(GPIB0, GPIBMagnetPSU)
        Call ReadStatusByte(GPIB0, GPIBMagnetPSU, status)
'       Set control to external
        GPIBOut GPIBMagnetPSU, "CURR:MODE EXT"
'
    Else
'
        k = MsgBox("Magnet PSU not found. Do you want to try again?", vbYesNo)
        If k = 6 Then
            SendIFC (GPIB0)
            GoTo FindMagnetPSUAgain
        End If
    End If

End Sub
Public Sub SetMagI(SetField As Single)
    If MagnetPSUPresent = True Then
        GPIBOut GPIBGaussmeter, "CSETP " + Trim(Str(SetField / 1000))

'        GPIBMessageOut$ = "I1 " + Trim(I)
'        k = GPIBMagnetPSUOut(GPIBMessageOut$, GPIBMagnetPSU)
'        GPIBMessageOut$ = "OP1 1"
'        k = GPIBMagnetPSUOut(GPIBMessageOut$, GPIBMagnetPSU)
'        Do
'            k = GPIBMagnetPSUOut("I1O?", GPIBMagnetPSU)
'            GPIBmessage = GPIBMagnetPSUIn(GPIBMagnetPSU)
'            currentI = Val(GPIBmessage)
'        Loop Until currentI = I
    End If
End Sub

Public Sub MagnetOff()
    If MagnetPSUPresent = True Then
'        Set Control set point to 0mT
    GPIBOut GPIBGaussmeter, "CSETP 0"

'        Set Field control OFF
    GPIBOut GPIBGaussmeter, "CMODE 0"
'
'        k = GPIBMagnetPSUOut("OP1 0", GPIBMagnetPSU)
'        Do
'            k = GPIBMagnetPSUOut("I1O?", GPIBMagnetPSU)
'            GPIBmessage = GPIBMagnetPSUIn(GPIBMagnetPSU)
'            I = Val(GPIBmessage)
'        Loop Until I = 0
'        k = GPIBMagnetPSUOut("OP3 0", GPIBMagnetPSU)
    End If
End Sub

Public Sub ReverseField()
    If MagnetPSUPresent = True Then
        k = GPIBMagnetPSUOut("OP3 1", GPIBMagnetPSU)
    End If
End Sub

Public Function GPIBMagnetPSUIn(GPIBDevice As Integer) As String
    Dim MsgIn As String, status As Integer
    MsgIn = Space$(66000)
    Receive GPIB0, GPIBDevice, MsgIn, STOPend
    GPIBMagnetPSUIn = Trim(MsgIn)
End Function
Public Function GPIBMagnetPSUOut(GPIBMessageOut As String, GPIBDevice As Integer)
    Dim status As Integer
    Send GPIB0, GPIBDevice, GPIBMessageOut, NLend
End Function

Public Sub DefineVariables()
    IntensityForm.PictureVL.FillStyle = 0
    IntensityForm.PictureIV.FillStyle = 0
'
    OptionsFocus = 0: Data.KeithleyFilter = 5
    Data.SampleName = "No Details"
    Data.LowVoltage = 0
    Data.HighVoltage = 10
    Data.IntegrationTime = 1
    Data.NewportFilter = 1: Data.Sensitivity = 0
    Data.Attenuator = 0: Data.Zero = 0
    Data.VoltageStep = 3: VoltageStepValue = 0.1
    Data.SettlingTime = 0.01
    Data.CurrentCompliance = 100
    Data.VoltageCompliance = 10
    Data.Wavelength = 1532
    Data.DisplayCurrentLow = 0
    Data.DisplayCurrentHigh = 0.000000001
    Data.DisplayIntensityLow = 0
    Data.DisplayIntensityHigh = 0.000000001
    Data.CurrentAutorange = 1
    Data.IntensityAutorange = 1
    Data.CurrentLogLin = 0
    Data.IntensityLogLin = 0
    IntervalTimeValue = 0.02
    RecordedOnString = "No data recorded"
    ScanType = 1
    DataSaved = False
    GlobalPathName = "C:\"
    SourceVoltageOn = False
End Sub

Public Sub DisplayValues()
    IntensityForm.LabelSample.Caption = "Sample is " & Trim(Data.SampleName)
    IntensityForm.LabelDateTime.Caption = RecordedOnString
    If DataSaved Then
        IntensityForm.SavedAs.Caption = "Saved as " & GlobalPathName & GlobalFileName
    Else
        IntensityForm.SavedAs.Caption = ""
    End If
    IntensityForm.LabelScan.Caption = "Scan voltage from " & Trim(Data.LowVoltage) & " V to " & Trim(Data.HighVoltage) & " V in " & Trim(VoltageStepValue) & " V steps"
    IntensityForm.LabelDelay.Caption = "Settling time per data point is " & Str$(Data.SettlingTime) & " s"
    IntensityForm.LabelCompliance.Caption = "Compliance current is " & Trim(Data.CurrentCompliance) & " mA"
    If Abs(Data.DisplayCurrentLow) >= 1 Then
        temp1 = Trim(Data.DisplayCurrentLow) & " mA"
    ElseIf Abs(Data.DisplayCurrentLow) >= 0.001 Then
        temp1 = Trim(CLng(Data.DisplayCurrentLow * 100000) / 100) & " µA"
    ElseIf Abs(Data.DisplayCurrentLow) >= 0.000001 Then
        temp1 = Trim(CLng(Data.DisplayCurrentLow * 100000000) / 100) & " nA"
    ElseIf Abs(Data.DisplayCurrentLow) >= 0.000000001 Then
        temp1 = Trim(CLng(Data.DisplayCurrentLow * 100000000000#) / 100) & " pA"
    End If
    If Data.DisplayCurrentLow = 0 Then temp1 = "0 mA"
    If Abs(Data.DisplayCurrentHigh) >= 1 Then
        temp2 = Trim(Data.DisplayCurrentHigh) & " mA"
    ElseIf Abs(Data.DisplayCurrentHigh) >= 0.001 Then
        temp2 = Trim(CLng(Data.DisplayCurrentHigh * 100000) / 100) & " µA"
    ElseIf Abs(Data.DisplayCurrentHigh) >= 0.000001 Then
        temp2 = Trim(CLng(Data.DisplayCurrentHigh * 100000000) / 100) & " nA"
    ElseIf Abs(Data.DisplayCurrentHigh) >= 0.000000001 Then
        temp2 = Trim(CLng(Data.DisplayCurrentHigh * 100000000000#) / 100) & " pA"
    End If
    If Data.DisplayCurrentHigh = 0 Then temp2 = "0 mA"
    IntensityForm.LabelDisplayCurrent.Caption = "Display current from " & temp1 & " to " & temp2
End Sub

Public Sub StartScan()
Dim buffer As String
    If DataSaved = False And Data.NumberOfPoints > 0 Then
        d = MsgBox("Data is not saved. Do you want to continue?", vbYesNo)
        If d = 7 Then Exit Sub
    End If
    DataSaved = False
    IntensityForm.CommandScan.Enabled = False
    IntensityForm.CommandSource.Enabled = False
    Abort = False
    RecordingData = True
    If Data.CurrentAutorange = 1 Then
        Data.DisplayCurrentLow = 0
        Data.DisplayCurrentHigh = 0.000000001
    End If
    If Data.IntensityAutorange = 1 Then
        Data.DisplayIntensityLow = 0
        Data.DisplayIntensityHigh = 0.000000001
    End If
    Data.Date = Now
    RecordedOnString = "Recorded on " & Format(Data.Date, "Long Date") & " at " & Format(Data.Date, "Long Time")
    DisplayValues
    Axis
    Data.NumberOfPoints = 0
    Select Case ScanType
        Case 1
            If Data.LowVoltage >= 0 Then
                Scan Data.LowVoltage, Data.HighVoltage, VoltageStepValue
            ElseIf Data.HighVoltage <= 0 Then
                Scan Data.HighVoltage, Data.LowVoltage, -VoltageStepValue
            Else
                Scan 0, Data.HighVoltage, VoltageStepValue
                If Abort Then GoTo EndSelect
                Scan 0, Data.LowVoltage, -VoltageStepValue
            End If
        Case 2
            If Data.LowVoltage >= 0 Then
                Scan Data.LowVoltage, Data.HighVoltage, VoltageStepValue
            ElseIf Data.HighVoltage <= 0 Then
                Scan Data.HighVoltage, Data.LowVoltage, -VoltageStepValue
            Else
                Scan 0, Data.LowVoltage, -VoltageStepValue
                If Abort Then GoTo EndSelect
                Scan 0, Data.HighVoltage, VoltageStepValue
            End If
        Case 3
                Scan Data.LowVoltage, Data.HighVoltage, VoltageStepValue
        Case 4
                Scan Data.HighVoltage, Data.LowVoltage, -VoltageStepValue
EndSelect:
    End Select
    If SourcePresent Then
        AbortAll
'        SourceVoltage (0)
'        buffer = "N0X"
'        GPIBSourceOut GPIBSource, buffer
    End If
    Abort = False
    RecordingData = False
    DataSaved = False
    IntensityForm.CommandScan.Enabled = True
    IntensityForm.CommandSource.Enabled = True
End Sub

Public Sub IVXAxisV()
    IntensityForm.PictureIV.Line (0, VAxisOffset)-(1, VAxisOffset)
'
    If VAxisOffset = 1 Then
        Ticks = -1
    Else
        Ticks = 1
    End If
'
    c = Data.HighVoltage - Data.LowVoltage: lc = Log10(c): pow = Int(lc): inc = lc - pow
'
' calculate position of ticks on x-axis
'
    a = 0.05
    If inc > Log10(1.6) Then a = 0.1
    If inc > Log10(4) Then a = 0.2
    If inc > Log10(8) Then a = 0.5
    tx = a * 10 ^ pow
'
    For Xh = Int(Data.LowVoltage / tx) * tx To Int(Data.HighVoltage / tx) * tx + tx Step tx * Sgn(Data.HighVoltage - Data.LowVoltage)
        If (Xh < Data.LowVoltage) Or (Xh > Data.HighVoltage) Then GoTo ignor1
        xnc = (Xh - Data.LowVoltage) / (Data.HighVoltage - Data.LowVoltage)
        IntensityForm.PictureIV.Line (xnc, VAxisOffset)-(xnc, VAxisOffset - 0.01 * Ticks)
ignor1:
    Next
'
' calculate positions of numbers on x-axis
'
    a = 0.2
    If inc > Log10(1.6) Then a = 0.5
    If inc > Log10(4) Then a = 1
    If inc > Log10(8) Then a = 2
    tx = a * 10 ^ pow
'
    For Xg# = Int(Data.LowVoltage / tx) * tx - tx To Int(Data.HighVoltage / tx) * tx + tx Step tx * Sgn(Data.HighVoltage - Data.LowVoltage)
        temp! = CSng(Xg#)
        If (temp! + 0.000001 < Data.LowVoltage) Or (temp! > Data.HighVoltage) Then GoTo ignor2
        xnc = (temp! - Data.LowVoltage) / (Data.HighVoltage - Data.LowVoltage)
        IntensityForm.PictureIV.Line (xnc, VAxisOffset)-(xnc, VAxisOffset - 0.02 * Ticks)
        If Ticks = -1 Then
            IntensityForm.PictureIV.CurrentY = IntensityForm.PictureIV.CurrentY + 0.035
        End If
        out$ = Format$(temp!)
        IntensityForm.PictureIV.CurrentX = IntensityForm.PictureIV.CurrentX - 0.012 * Len(out$)
        If Not (VAxisOffset > 0 And VAxisOffset < 1 And Val(out$) = 0) Then
            IntensityForm.PictureIV.Print out$
        End If
        
'        IntensityForm.PictureIV.Print out$
ignor2:
    Next
    If IAxisOffset = 1 Then
        IntensityForm.PictureIV.CurrentX = -0.03
    Else
        IntensityForm.PictureIV.CurrentX = 1.03
    End If
    IntensityForm.PictureIV.CurrentY = VAxisOffset + 0.015
    IntensityForm.PictureIV.Print "V"
'
End Sub
Function Log10(X)
    Log10 = Log(X) / Log(10)
End Function

Public Sub IVYAxisI()
    IntensityForm.PictureIV.Line (IAxisOffset, 0)-(IAxisOffset, 1)
'
    If Data.DisplayCurrentHigh <= Data.DisplayCurrentLow Then AutorangeCurrent
    high = Data.DisplayCurrentHigh
    low = Data.DisplayCurrentLow
    If Abs(high) > Abs(low) Then
        exponent = Int(Log10(Abs(high)))
    Else
        exponent = Int(Log10(Abs(low)))
    End If
    high = high / (10 ^ exponent)
    low = low / (10 ^ exponent)
    If IAxisOffset = 1 Then
        Ticks = -1
    Else
        Ticks = 1
    End If
'
    c = high - low: lc = Log10(c): pow = Int(lc): inc = lc - pow
'
' calculate position of ticks on y-axis
'
    a = 0.05
    If inc > Log10(1.6) Then a = 0.1
    If inc > Log10(4) Then a = 0.2
    If inc > Log10(8) Then a = 0.5
    tx = a * 10 ^ pow
'
    For Xh = Int(low / tx) * tx To Int(high / tx) * tx + tx Step tx * Sgn(high - low)
        If (Xh < low) Or (Xh > high) Then GoTo ignor1
        xnc = (Xh - low) / (high - low)
        IntensityForm.PictureIV.Line (IAxisOffset, xnc)-(IAxisOffset - 0.01 * Ticks, xnc)
ignor1:
    Next
'
' calculate positions of numbers on y-axis
'
    a = 0.2
    If inc > Log10(1.6) Then a = 0.5
    If inc > Log10(4) Then a = 1
    If inc > Log10(8) Then a = 2
    tx = a * 10 ^ pow
'
    For Xg# = Int(low / tx) * tx - tx To Int(high / tx) * tx + tx Step tx * Sgn(high - low)
        temp! = CSng(Xg#)
        If (temp! + 0.000001 < low) Or (temp! > high) Then GoTo ignor2
        xnc = (temp! - low) / (high - low)
        IntensityForm.PictureIV.Line (IAxisOffset, xnc)-(IAxisOffset - 0.02 * Ticks, xnc)
        out$ = Format$(temp!, "#0.0")
        If Ticks = 1 Then
            IntensityForm.PictureIV.CurrentX = IntensityForm.PictureIV.CurrentX - 0.03 * Len(out$)
        End If
        IntensityForm.PictureIV.CurrentY = IntensityForm.PictureIV.CurrentY + 0.017
        If Not (IAxisOffset > 0 And IAxisOffset < 1 And Val(out$) = 0) Then
            IntensityForm.PictureIV.Print out$
        End If
ignor2:
    Next
    If VAxisOffset = 1 Then
        IntensityForm.PictureIV.CurrentY = -0.01
    Else
        IntensityForm.PictureIV.CurrentY = 1.05
    End If
    IntensityForm.PictureIV.CurrentX = IAxisOffset - 0.1
    IntensityForm.PictureIV.Print "I x10";
    IntensityForm.PictureIV.CurrentY = IntensityForm.PictureIV.CurrentY + 0.017
    IntensityForm.PictureIV.Print Trim(exponent);
    IntensityForm.PictureIV.CurrentY = IntensityForm.PictureIV.CurrentY - 0.017
    IntensityForm.PictureIV.Print "A";

End Sub

Public Sub Axis()
'
' Plot IV axis
'
    IntensityForm.PictureIV.Cls
    IntensityForm.PictureIV.Scale (-0.15, 1.15)-(1.15, -0.15)
'
    If Data.LowVoltage >= 0 Then
        IAxisOffset = 0
    Else
        IAxisOffset = Abs(Data.LowVoltage) / (Data.HighVoltage - Data.LowVoltage)
    End If
    If Data.HighVoltage <= 0 Then IAxisOffset = 1
'
    If Data.CurrentLogLin = 1 Then
        VAxisOffset = 0
        IVYLogAxisI
    Else
        If Data.DisplayCurrentHigh = 0 And Data.DisplayCurrentLow = 0 Then
            Data.DisplayCurrentHigh = 0.000000001
        End If

        If Data.DisplayCurrentLow >= 0 Then
            VAxisOffset = 0
        Else
            VAxisOffset = Abs(Data.DisplayCurrentLow) / (Data.DisplayCurrentHigh - Data.DisplayCurrentLow)
        End If
        If Data.DisplayCurrentHigh <= 0 Then VAxisOffset = 1
        IVYAxisI
    End If
'
    IVXAxisV
'
' Plot Intensity Axis
'
        IntensityForm.PictureVL.Cls
        IntensityForm.PictureVL.Scale (-0.15, 1.15)-(1.15, -0.15)
'
        If Data.LowVoltage >= 0 Then
            IAxisOffset = 0
        Else
            IAxisOffset = Abs(Data.LowVoltage) / (Data.HighVoltage - Data.LowVoltage)
        End If
        If Data.HighVoltage <= 0 Then IAxisOffset = 1
'
        If Data.IntensityLogLin = 1 Then
            VAxisOffset = 0
            VLYLogAxisL
        Else
            If Data.DisplayIntensityLow >= 0 Then
                VAxisOffset = 0
            Else
                VAxisOffset = Abs(Data.DisplayIntensityLow) / (Data.DisplayIntensityHigh - Data.DisplayIntensityLow)
            End If
            If Data.DisplayIntensityHigh <= 0 Then VAxisOffset = 1
            VLYAxisL
        End If
'
        VLXAxisV
End Sub

Public Sub IVYLogAxisI()
    IntensityForm.PictureIV.Line (IAxisOffset, 0)-(IAxisOffset, 1)
'
    Dim low As Single, high As Single
    If Data.DisplayCurrentLow <= 0 Then
        low = 0.000000001
    Else
        low = 10 ^ Int(Log10(Abs(Data.DisplayCurrentLow)))
    End If
    If Data.DisplayCurrentHigh <= 0 Then
        high = 10
    Else
        high = 10 ^ (Int(Log10(Abs(Data.DisplayCurrentHigh - 0.0001)) + 1))
    End If
    LogLow = Log10(low)
    LogHigh = Log10(high)
    tx = CInt(Log10(high / low))
'
    If IAxisOffset = 1 Then
        Ticks = -1
    Else
        Ticks = 1
    End If
'
    For yl = 0 To tx
        IntensityForm.PictureIV.Line (IAxisOffset, yl / tx)-(IAxisOffset - 0.02 * Ticks, yl / tx)
        out$ = "10"
        If Ticks = 1 Then
            IntensityForm.PictureIV.CurrentX = IntensityForm.PictureIV.CurrentX - 0.1
        End If
        IntensityForm.PictureIV.CurrentY = IntensityForm.PictureIV.CurrentY + 0.017
        IntensityForm.PictureIV.Print out$;
        IntensityForm.PictureIV.CurrentY = IntensityForm.PictureIV.CurrentY + 0.014
        out$ = Str$(Int(Log10(low * 10 ^ yl) + 0.5))
        IntensityForm.PictureIV.FontSize = IntensityForm.PictureIV.FontSize - 1
        IntensityForm.PictureIV.Print out$;
        IntensityForm.PictureIV.FontSize = IntensityForm.PictureIV.FontSize + 1
        If yl = tx Then GoTo ignor3
'
        For yll = 2 To 9
            yp = (yl + Log10(yll)) / tx
            IntensityForm.PictureIV.Line (IAxisOffset, yp)-(IAxisOffset - 0.01 * Ticks, yp)
        Next
'
ignor3:
    Next
'
    IntensityForm.PictureIV.CurrentY = 1.08
    IntensityForm.PictureIV.CurrentX = IAxisOffset - 0.1
    IntensityForm.PictureIV.Print "I (A)";
'
End Sub

Public Sub Replot()
    Axis
    If Data.NumberOfPoints = 0 Then Exit Sub
    For z = 0 To Data.NumberOfPoints
        xp = (Data.Voltage(z) - Data.LowVoltage) / (Data.HighVoltage - Data.LowVoltage)
'
' replot current data
'
        If Data.CurrentLogLin = 0 Then
            yp = (Data.Current(z) - Data.DisplayCurrentLow) / (Data.DisplayCurrentHigh - Data.DisplayCurrentLow)
        Else
            If Data.Current(z) = 0 Then
                yp = 0
            Else
                yp = (Log10(Abs(Data.Current(z))) - LogLow) / (LogHigh - LogLow)
            End If
        End If
        If yp >= 0 And yp <= 1 Then
            If Data.Current(z) > 0 Then
                Colour = QBColor(1)
            ElseIf Data.Current(z) < 0 Then
                Colour = QBColor(4)
            Else
                Colour = QBColor(2)
            End If
            IntensityForm.PictureIV.FillColor = Colour
            IntensityForm.PictureIV.Circle (xp, yp), 0.005, Colour
        End If
'
' replot intensity data
'
        If Data.IntensityLogLin = 0 Then
            yp = (Data.Intensity(z) - Data.DisplayIntensityLow) / (Data.DisplayIntensityHigh - Data.DisplayIntensityLow)
        Else
            If Data.Intensity(z) = 0 Then
                yp = 0
            Else
                yp = (Log10(Abs(Data.Intensity(z))) - LLogLow) / (LLogHigh - LLogLow)
            End If
        End If
        If yp >= 0 And yp <= 1 Then
            If Data.Intensity(z) > 0 Then
                Colour = QBColor(1)
            ElseIf Data.Current(z) < 0 Then
                Colour = QBColor(4)
            Else
                Colour = QBColor(2)
            End If
            IntensityForm.PictureVL.FillColor = Colour
            IntensityForm.PictureVL.Circle (xp, yp), 0.005, Colour
        End If
    Next
End Sub

Public Sub VLYLogAxisL()
    IntensityForm.PictureVL.Line (IAxisOffset, 0)-(IAxisOffset, 1)
'
    Dim low As Single, high As Single
    If Data.DisplayIntensityLow <= 0 Then
        low = 0.000000000001
    Else
        low = 10 ^ Int(Log10(Abs(Data.DisplayIntensityLow)))
    End If
    If Data.DisplayIntensityHigh <= 0 Then
        high = 10
    Else
        high = 10 ^ (Int(Log10(Abs(Data.DisplayIntensityHigh - 0.0001)) + 1))
    End If
    LLogLow = Log10(low)
    LLogHigh = Log10(high)
    tx = CInt(Log10(high / low))
'
    If IAxisOffset = 1 Then
        Ticks = -1
    Else
        Ticks = 1
    End If
'
    For yl = 0 To tx
        IntensityForm.PictureVL.Line (IAxisOffset, yl / tx)-(IAxisOffset - 0.02 * Ticks, yl / tx)
        out$ = "10"
        If Ticks = 1 Then
            IntensityForm.PictureVL.CurrentX = IntensityForm.PictureVL.CurrentX - 0.1
        End If
        IntensityForm.PictureVL.CurrentY = IntensityForm.PictureVL.CurrentY + 0.017
        IntensityForm.PictureVL.Print out$;
        IntensityForm.PictureVL.CurrentY = IntensityForm.PictureVL.CurrentY + 0.014
        out$ = Str$(Int(Log10(low * 10 ^ yl) + 0.5))
        IntensityForm.PictureVL.FontSize = IntensityForm.PictureVL.FontSize - 1
        IntensityForm.PictureVL.Print out$;
        IntensityForm.PictureVL.FontSize = IntensityForm.PictureVL.FontSize + 1
        If yl = tx Then GoTo ignor3
'
        For yll = 2 To 9
            yp = (yl + Log10(yll)) / tx
            IntensityForm.PictureVL.Line (IAxisOffset, yp)-(IAxisOffset - 0.01 * Ticks, yp)
        Next
'
ignor3:
    Next
'
    IntensityForm.PictureVL.CurrentY = 1.08
    IntensityForm.PictureVL.CurrentX = IAxisOffset - 0.1
    IntensityForm.PictureVL.Print "L (W)";
'
End Sub

Public Sub VLXAxisV()
    IntensityForm.PictureVL.Line (0, VAxisOffset)-(1, VAxisOffset)
'
    If VAxisOffset = 1 Then
        Ticks = -1
    Else
        Ticks = 1
    End If
'
    c = Data.HighVoltage - Data.LowVoltage: lc = Log10(c): pow = Int(lc): inc = lc - pow
'
' calculate position of ticks on x-axis
'
    a = 0.05
    If inc > Log10(1.6) Then a = 0.1
    If inc > Log10(4) Then a = 0.2
    If inc > Log10(8) Then a = 0.5
    tx = a * 10 ^ pow
'
    For Xh = Int(Data.LowVoltage / tx) * tx To Int(Data.HighVoltage / tx) * tx + tx Step tx * Sgn(Data.HighVoltage - Data.LowVoltage)
        If (Xh < Data.LowVoltage) Or (Xh > Data.HighVoltage) Then GoTo ignor1
        xnc = (Xh - Data.LowVoltage) / (Data.HighVoltage - Data.LowVoltage)
        IntensityForm.PictureVL.Line (xnc, VAxisOffset)-(xnc, VAxisOffset - 0.01 * Ticks)
ignor1:
    Next
'
' calculate positions of numbers on x-axis
'
    a = 0.2
    If inc > Log10(1.6) Then a = 0.5
    If inc > Log10(4) Then a = 1
    If inc > Log10(8) Then a = 2
    tx = a * 10 ^ pow
'
    For Xg# = Int(Data.LowVoltage / tx) * tx - tx To Int(Data.HighVoltage / tx) * tx + tx Step tx * Sgn(Data.HighVoltage - Data.LowVoltage)
        temp! = CSng(Xg#)
        If (temp! + 0.000001 < Data.LowVoltage) Or (temp! > Data.HighVoltage) Then GoTo ignor2
        xnc = (temp! - Data.LowVoltage) / (Data.HighVoltage - Data.LowVoltage)
        IntensityForm.PictureVL.Line (xnc, VAxisOffset)-(xnc, VAxisOffset - 0.02 * Ticks)
        If Ticks = -1 Then
            IntensityForm.PictureVL.CurrentY = IntensityForm.PictureVL.CurrentY + 0.035
        End If
        out$ = Format$(temp!)
        IntensityForm.PictureVL.CurrentX = IntensityForm.PictureVL.CurrentX - 0.012 * Len(out$)
        If Not (VAxisOffset > 0 And VAxisOffset < 1 And Val(out$) = 0) Then
            IntensityForm.PictureVL.Print out$
        End If
        
'        IntensityForm.PictureIV.Print out$
ignor2:
    Next
    If IAxisOffset = 1 Then
        IntensityForm.PictureVL.CurrentX = -0.03
    Else
        IntensityForm.PictureVL.CurrentX = 1.03
    End If
    IntensityForm.PictureVL.CurrentY = VAxisOffset + 0.015
    IntensityForm.PictureVL.Print "V"
'
End Sub

Public Sub VLYAxisL()
    IntensityForm.PictureVL.Line (IAxisOffset, 0)-(IAxisOffset, 1)
'
    If Data.DisplayIntensityHigh <= Data.DisplayIntensityLow Then AutorangeIntensity
    high = Data.DisplayIntensityHigh
    low = Data.DisplayIntensityLow
    If Abs(high) > Abs(low) Then
        exponent = Int(Log10(Abs(high)))
    Else
        exponent = Int(Log10(Abs(low)))
    End If
    high = high / (10 ^ exponent)
    low = low / (10 ^ exponent)
    If IAxisOffset = 1 Then
        Ticks = -1
    Else
        Ticks = 1
    End If
'
    c = high - low: lc = Log10(c): pow = Int(lc): inc = lc - pow
'
' calculate position of ticks on y-axis
'
    a = 0.05
    If inc > Log10(1.6) Then a = 0.1
    If inc > Log10(4) Then a = 0.2
    If inc > Log10(8) Then a = 0.5
    tx = a * 10 ^ pow
'
    For Xh = Int(low / tx) * tx To Int(high / tx) * tx + tx Step tx * Sgn(high - low)
        If (Xh < low) Or (Xh > high) Then GoTo ignor1
        xnc = (Xh - low) / (high - low)
        IntensityForm.PictureVL.Line (IAxisOffset, xnc)-(IAxisOffset - 0.01 * Ticks, xnc)
ignor1:
    Next
'
' calculate positions of numbers on y-axis
'
    a = 0.2
    If inc > Log10(1.6) Then a = 0.5
    If inc > Log10(4) Then a = 1
    If inc > Log10(8) Then a = 2
    tx = a * 10 ^ pow
'
    For Xg# = Int(low / tx) * tx - tx To Int(high / tx) * tx + tx Step tx * Sgn(high - low)
        temp! = CSng(Xg#)
        If (temp! + 0.000001 < low) Or (temp! > high) Then GoTo ignor2
        xnc = (temp! - low) / (high - low)
        IntensityForm.PictureVL.Line (IAxisOffset, xnc)-(IAxisOffset - 0.02 * Ticks, xnc)
        out$ = Format$(temp!, "#0.0")
        If Ticks = 1 Then
            IntensityForm.PictureVL.CurrentX = IntensityForm.PictureVL.CurrentX - 0.03 * Len(out$)
        End If
        IntensityForm.PictureVL.CurrentY = IntensityForm.PictureVL.CurrentY + 0.017
        If Not (IAxisOffset > 0 And IAxisOffset < 1 And Val(out$) = 0) Then
            IntensityForm.PictureVL.Print out$
        End If
ignor2:
    Next
    If VAxisOffset = 1 Then
        IntensityForm.PictureVL.CurrentY = -0.01
    Else
        IntensityForm.PictureVL.CurrentY = 1.05
    End If
    IntensityForm.PictureVL.CurrentX = IAxisOffset - 0.1
    IntensityForm.PictureVL.Print "L x10";
    IntensityForm.PictureVL.CurrentY = IntensityForm.PictureVL.CurrentY + 0.017
    IntensityForm.PictureVL.Print Trim(exponent);
    IntensityForm.PictureVL.CurrentY = IntensityForm.PictureVL.CurrentY - 0.017
    IntensityForm.PictureVL.Print "W";
End Sub

Public Sub AutorangeIntensity()
    Data.DisplayIntensityHigh = 0.000000001
    For T = 0 To Data.NumberOfPoints
        If Data.Intensity(T) > Data.DisplayIntensityHigh Then
            Data.DisplayIntensityHigh = CSng(Abs(Data.Intensity(T)) * 1200000000) / 1000000000
        End If
    Next
End Sub

Public Sub AutorangeCurrent()
    Data.DisplayCurrentHigh = 0.00000000001
    If Data.CurrentLogLin Then
        Data.DisplayCurrentLow = 0.00000000001
    Else
        Data.DisplayCurrentLow = 0
    End If
    For T = 0 To Data.NumberOfPoints
        If Data.Current(T) > Data.DisplayCurrentHigh Then
            Data.DisplayCurrentHigh = CSng(Abs(Data.Current(T)) * 1200) / 1000
        End If
'        If Data.Current(T) < Data.DisplayCurrentLow Then
'            If Data.Current(T) > 0 Then
'                Data.DisplayCurrentLow = 0
'            Else
'                Data.DisplayCurrentLow = CSng(Data.Current(T) * 1200) / 1000
'            End If
'        End If
    Next

End Sub

Public Sub DisplayCurrent(Current As Single)
        If Abs(Current) < 0.000001 Then
            temp$ = Format(Current * 1000000000#, "#0.000") & " nA"
        ElseIf Abs(Current) < 0.001 Then
            temp$ = Format(Current * 1000000#, "#0.000") & " µA"
        Else
            temp$ = Format(Current * 1000, "#0.000") & " mA"
        End If
        IntensityForm.LabelCurrent.Caption = "Current is " & temp$
End Sub

Public Sub DisplayVoltage(Voltage As Single)
        If Abs(Voltage) < 1 Then
            temp$ = Format(Voltage * 1000#, "#0") & " mV"
        Else
            temp$ = Format(Voltage, "#0.00") & " V"
        End If
        IntensityForm.LabelVoltage.Caption = "Voltage is " & temp$
End Sub

Public Sub DisplayIntensity(Intensity As Single)
        If Abs(Intensity) < 0.001 Then
            temp$ = Format(Intensity * 1000000000#, "#0.0") & " nW"
        ElseIf Abs(Intensity) < 1 Then
            temp$ = Format(Intensity * 1000000#, "#0.0") & " µW"
        Else
            temp$ = Format(Intensity * 1000#, "#0.0") & " mW"
        End If
        IntensityForm.LabelIntensity.Caption = "Intensity is " & temp$
End Sub

Public Sub TakeData(V As Single)
        Data.Voltage(Data.NumberOfPoints) = V

        DisplayVoltage (V)

        If Abort Then Exit Sub
        If SourcePresent Then
            If SMUType = "4200" Then
                SourceVoltage (V)
                Sleep (300)
' Read Intensity before current
                If NewportPresent Then
                    Data.Intensity(Data.NumberOfPoints) = ReadIntensity
                End If
'
                Data.Current(Data.NumberOfPoints) = ReadCurrent
                If Abort Then
                    GoTo AbortRun
                End If
            Else
repeatmeasurement:
                SourceVoltage (V)
                Data.Current(Data.NumberOfPoints) = ReadCurrent
                If Abort Then
                    GoTo AbortRun
                End If
'
                If Abs(Data.Current(Data.NumberOfPoints)) = 0.000001 Then GoTo repeatmeasurement
                If Abs(Data.Current(Data.NumberOfPoints)) = 0.000002 Then GoTo repeatmeasurement
                If Abs(Data.Current(Data.NumberOfPoints)) = 0.000003 Then GoTo repeatmeasurement
                If Abs(Data.Current(Data.NumberOfPoints)) = 0.000004 Then GoTo repeatmeasurement
                If Abs(Data.Current(Data.NumberOfPoints)) = 0.000005 Then GoTo repeatmeasurement
                If Abs(Data.Current(Data.NumberOfPoints)) = 0.000006 Then GoTo repeatmeasurement
                If Abs(Data.Current(Data.NumberOfPoints)) = 0.000007 Then GoTo repeatmeasurement
                If Abs(Data.Current(Data.NumberOfPoints)) = 0.000008 Then GoTo repeatmeasurement
                If Abs(Data.Current(Data.NumberOfPoints)) = 0.000009 Then GoTo repeatmeasurement
'
            End If
        End If
        DisplayCurrent (Data.Current(Data.NumberOfPoints))
'
'        If NewportPresent Then
'            Data.Intensity(Data.NumberOfPoints) = ReadIntensity
'        End If
        DisplayIntensity (Data.Intensity(Data.NumberOfPoints))
        xp = (Data.Voltage(Data.NumberOfPoints) - Data.LowVoltage) / (Data.HighVoltage - Data.LowVoltage)
'
' plot current data
'
        If Data.CurrentLogLin = 0 Then
            yp = (Data.Current(Data.NumberOfPoints) - Data.DisplayCurrentLow) / (Data.DisplayCurrentHigh - Data.DisplayCurrentLow)
        Else
            If Data.Current(Data.NumberOfPoints) = 0 Then
                yp = 0
            Else
                yp = (Log10(Abs(Data.Current(Data.NumberOfPoints))) - LogLow) / (LogHigh - LogLow)
            End If
        End If
        If (yp > 1 Or yp < 0) And Data.CurrentAutorange = 1 Then
            AutorangeCurrent
            DisplayValues
            Replot
        End If
        If yp >= 0 And yp <= 1 Then
            If Data.Current(Data.NumberOfPoints) > 0 Then
                Colour = QBColor(1)
            ElseIf Data.Current(Data.NumberOfPoints) < 0 Then
                Colour = QBColor(4)
            Else
                Colour = QBColor(2)
            End If
            IntensityForm.PictureIV.FillColor = Colour
            IntensityForm.PictureIV.Circle (xp, yp), 0.005, Colour
        End If
'
'plot intensity data
'
        If Data.IntensityLogLin = 0 Then
            yp = (Abs(Data.Intensity(Data.NumberOfPoints)) - Data.DisplayIntensityLow) / (Data.DisplayIntensityHigh - Data.DisplayIntensityLow)
        Else
            If Data.Intensity(Data.NumberOfPoints) = 0 Then
                yp = 0
            Else
                yp = (Log10(Abs(Data.Intensity(Data.NumberOfPoints))) - LLogLow) / (LLogHigh - LLogLow)
            End If
        End If
        If yp > 1 And Data.IntensityAutorange = 1 Then
            Data.DisplayIntensityHigh = Data.Intensity(Data.NumberOfPoints) * 1.2
            DisplayValues
            Replot
        End If
        If yp >= 0 And yp <= 1 Then
            If Data.Intensity(Data.NumberOfPoints) > 0 Then
                Colour = QBColor(1)
            ElseIf Data.Intensity(Data.NumberOfPoints) < 0 Then
                Colour = QBColor(4)
            Else
                Colour = QBColor(2)
            End If
            IntensityForm.PictureVL.FillColor = Colour
            IntensityForm.PictureVL.Circle (xp, yp), 0.005, Colour
        End If
       
        Data.NumberOfPoints = Data.NumberOfPoints + 1
        If Abort Then
            V = Data.HighVoltage
        End If
AbortRun:

End Sub

Public Sub Scan(StartV As Single, EndV As Single, Vstep)
    Dim buffer As String
    For V = StartV To EndV Step Vstep
        TakeData (V)
        If Abort Then V = EndV
    Next
    AbortAll
End Sub

Public Sub SaveData()
BeginSaveData:
    On Error GoTo SaveAsEnd
    IntensityForm.CommonDialog1.Flags = &H800
    IntensityForm.CommonDialog1.InitDir = GlobalPathName
    IntensityForm.CommonDialog1.filename = GlobalFileName
    IntensityForm.CommonDialog1.Filter = "IV Data (*.ivl) |*.ivl|Text (*.txt) |*.txt"
    IntensityForm.CommonDialog1.FilterIndex = 1
    IntensityForm.CommonDialog1.ShowSave
    filename = IntensityForm.CommonDialog1.FileTitle
    pathname = Left$(IntensityForm.CommonDialog1.filename, InStrRev(IntensityForm.CommonDialog1.filename, IntensityForm.CommonDialog1.FileTitle) - 1)
    SaveFileName = pathname + filename
    If UCase(Right$(filename, 3)) = "IVL" Then
        Open SaveFileName For Random As #disc Len = Len(Data)
            If LOF(disc) <> 0 Then
                d = MsgBox("A File with that name already exists. Do you want to overwrite it?", vbYesNo)
                If d = 7 Then
                    Close #disc: Exit Sub
                End If
            End If
            Put #disc, 1, Data
        Close #disc
    ElseIf UCase(Right$(filename, 3)) = "TXT" Then
        Open SaveFileName For Output As disc
            If LOF(disc) <> 0 Then
                d = MsgBox("A File with that name already exists. Do you want to overwrite it?", vbYesNo)
                If d = 7 Then Exit Sub
            End If
            For X = 0 To Data.NumberOfPoints - 1
                Print #disc, Data.Voltage(X), Data.Current(X), Data.Intensity(X)
            Next
        Close disc
    Else
        d = MsgBox("That file type is not supported. The data has not been saved", vbOKOnly)
        GoTo BeginSaveData
    End If
    GlobalFileName = filename
    GlobalPathName = pathname
    DataSaved = True
    DisplayValues
    On Error GoTo 0
SaveAsEnd:


End Sub

Public Sub OpenData()

    IntensityForm.CommonDialog1.Flags = &H4 + &H800 + &H1000
On Error GoTo LoadEnd
    IntensityForm.CommonDialog1.InitDir = GlobalPathName
    IntensityForm.CommonDialog1.Filter = "IV Data (*.ivl) |*.ivl"
    IntensityForm.CommonDialog1.ShowOpen
On Error GoTo 0
    GlobalFileName = IntensityForm.CommonDialog1.FileTitle
    GlobalPathName = Left$(IntensityForm.CommonDialog1.filename, InStrRev(IntensityForm.CommonDialog1.filename, IntensityForm.CommonDialog1.FileTitle) - 1)
    OpenFileName = GlobalPathName + GlobalFileName
    Open OpenFileName For Random As #disc Len = Len(Data)
            Get #disc, 1, Data
    Close #disc
    DataSaved = True
    DisplayValues
    Replot
LoadEnd:
End Sub


Public Sub CheckGPIBCard()
    On Error GoTo GPIBErr:
       Call SendIFC(GPIB0)
    On Error GoTo 0
    GPIBPresent = True
    Exit Sub
GPIBErr:
g = MsgBox("GPIB Card not found. Disabling GPIB functions", 0)
GPIBPresent = False
End Sub

Public Sub FindSourceMeasure()
Dim buffer As String
    Dim AddressList(3) As Integer, status As Integer, GPIBmessage As String
    AddressList(1) = GPIBSource
    AddressList(2) = GPIBNewport
    AddressList(3) = NOADDR
    SourcePresent = True
    SMUType = "4200"
FindSourceAgain:
    Do
        DoEvents
        EnableRemote GPIB0, AddressList
    Loop Until (ibsta <> &H8000)
    Call DevClear(GPIB0, GPIBSource)
    Call ReadStatusByte(GPIB0, GPIBSource, status)
    k = GPIBMagnetPSUOut("ID", GPIBSource)
    Call ReadStatusByte(GPIB0, GPIBSource, status)
    Do
        DoEvents
        GPIBmessage = GPIBMagnetPSUIn(GPIBSource)
    Loop Until (ibsta <> &H8000)
    If Left$(GPIBmessage, 6) = "KI4200" Then
        SourcePresent = True
' Set integration time to medium
        k = GPIBMagnetPSUOut("IT2", GPIBSource)
' Enable Serial polling for Data Ready
        k = GPIBMagnetPSUOut("DR1", GPIBSource)
' Clear Data Buffer
        k = GPIBMagnetPSUOut("BC", GPIBSource)
' Set resolution to 7 digits
        k = GPIBMagnetPSUOut("RS 7", GPIBSource)
' Switch off all channels and set CH1 to source
        k = GPIBMagnetPSUOut("DE", GPIBSource)
        k = GPIBMagnetPSUOut("CH1;CH2;CH3;CH4", GPIBSource)
        k = GPIBMagnetPSUOut("CH1, 'DRNV','DRNI',1,3", GPIBSource)
' Set channel 1 to constant voltage 0V output 100mA compliance
        k = GPIBMagnetPSUOut("SS", GPIBSource)
        k = GPIBMagnetPSUOut("VC1, 0, 100e-3", GPIBSource)
' Set to Measurement set up mode
        k = GPIBMagnetPSUOut("SM", GPIBSource)
' Set settling time interval before measurements to 0 ms
        k = GPIBMagnetPSUOut("WT 0", GPIBSource)
' Set time interval between measurements to 20 ms
        k = GPIBMagnetPSUOut("IN 0.02", GPIBSource)
' Set number of readings to default (8)
        NR$ = Str$(2 ^ Data.KeithleyFilter)
        k = GPIBMagnetPSUOut("NR " + NR$, GPIBSource)
' Set display mode to list
        k = GPIBMagnetPSUOut("DM2", GPIBSource)
        k = GPIBMagnetPSUOut("LI 'DRNI', 'DRNV'", GPIBSource)
'
    Else
        Call DevClear(GPIB0, GPIBSource)
        Call ReadStatusByte(GPIB0, GPIBSource, status)
        buffer = "U0X"
        GPIBSourceOut GPIBSource, buffer
'
        GPIBSourceIn GPIBSource, buffer
        GPIBmessage = Left$(buffer, 3)
        If GPIBmessage = "236" Then
            SourcePresent = True
            SMUType = "236"
'   Set the compliance current to 100 mA
            buffer = "F0,0X"
            GPIBSourceOut GPIBSource, buffer
            buffer = "L100E-3,0X"
            GPIBSourceOut GPIBSource, buffer
            buffer = "H0X"
            GPIBSourceOut GPIBSource, buffer
'   set the filter to 32 readings
            buffer = "P5X"
            GPIBSourceOut GPIBSource, buffer
'   set the integration time to medium
            buffer = "S1X"
            GPIBSourceOut GPIBSource, buffer
'   enable default delay
            buffer = "W1X"
            GPIBSourceOut GPIBSource, buffer
        Else
            SourcePresent = False
        End If
    End If
'
    If SourcePresent = False Then
        k = MsgBox("Source Measure Unit not found. Do you want to try again?", vbYesNo)
        If k = 6 Then
            SendIFC (GPIB0)
            GoTo FindSourceAgain
        End If
    End If

End Sub
Public Function GPIBIn(GPIBDevice As Integer, buffer As String) As String
    Dim MsgIn As String, status As Integer
'    If NewportPresent = 0 Then Exit Function
    MsgIn = Space$(20)
    Do
        Call ReadStatusByte(GPIB0, GPIBDevice, status)
        X = X + 1
        If X > 50 Then
            Call DevClear(GPIB0, GPIBDevice)
            ibsta = ibsta Or &H8000
            Exit Function
        End If
        If (ibsta And &H8000) Then Exit Function
        DoEvents
        Sleep 100
    Loop Until (status And 16) = 16
    Receive GPIB0, GPIBDevice, MsgIn, STOPend
    If (ibsta And &H8000) Then Exit Function
'    GPIBIn = Left$(MsgIn, InStr(1, MsgIn, Chr$(13)) - 1)
    buffer = MsgIn
End Function
Public Function GPIBOut(GPIBDevice As Integer, GPIBMessageOut As String)
    Dim status As Integer
'    If NewportPresent = 0 Then Exit Function
On Error GoTo GPIBOutError
    Do
        Call ReadStatusByte(GPIB0, GPIBDevice, status)
        X = X + 1
        If X > 50 Then
            Call DevClear(GPIB0, GPIBDevice)
            ibsta = ibsta Or &H8000
            Exit Function
        End If
        If (ibsta And &H8000) Then Exit Function
        DoEvents
        Sleep 100
    Loop Until (status And 32) <> 32
    Send GPIB0, GPIBDevice, GPIBMessageOut, DABend
    Exit Function
GPIBOutError:

End Function
Public Sub GPIBSourceOut(GPIBDevice As Integer, buffer As String)
    Dim status As Integer
    If Not (SourcePresent) Then Exit Sub
    Do
        Call ReadStatusByte(GPIB0, GPIBDevice, status)
        X = X + 1
        If X > 50 Then
            Call DevClear(GPIB0, GPIBDevice)
            ibsta = ibsta Or &H8000
            Exit Sub
        End If
        If (ibsta And &H8000) Then Exit Sub
        DoEvents
        Sleep 100
    Loop Until ((status And 8) = 8) Or ((status And 16) = 16)
    Send GPIB0, GPIBDevice, buffer, NLend
End Sub
Public Sub GPIBSourceIn(GPIBDevice As Integer, buffer As String)
    Dim status As Integer
    If Not (SourcePresent) Then Exit Sub
    buffer = Space$(20)
    Do
        Call ReadStatusByte(GPIB0, GPIBDevice, status)
        X = X + 1
        If X > 50 Then
            Call DevClear(GPIB0, GPIBDevice)
            ibsta = ibsta Or &H8000
            Exit Sub
        End If
        If (ibsta And &H8000) Then Exit Sub
        DoEvents
        Sleep 100
        Loop Until (status And 8) = 8
    
'    Do
'        DoEvents
            Receive GPIB0, GPIBDevice, buffer, STOPend
'    Loop Until (ibsta <> &H8000)
End Sub

Public Sub SourceVoltage(V As Single)
Dim buffer As String
    If SMUType = "4200" Then
        SMUSourceV 1, V, Data.CurrentCompliance / 1000
    Else
        buffer = "F0,0X"
        GPIBSourceOut GPIBSource, buffer
        buffer = "B" & Trim(Format(V, Scientific)) & ",0,0X"
        GPIBSourceOut GPIBSource, buffer
        buffer = "N1X"
        GPIBSourceOut GPIBSource, buffer
        buffer = "H0X"
        GPIBSourceOut GPIBSource, buffer
    End If
End Sub
Public Sub SMUSourceV(Num As Integer, V As Single, IC As Single)
    If SourcePresent Then
        N$ = Trim(Str(Num))
'        k = GPIBMagnetPSUOut("SS", GPIBSource)
'        k = GPIBMagnetPSUOut("VC1, 2, 100e-3", GPIBSource)
'        k = GPIBMagnetPSUOut("SM", GPIBSource)
'        k = GPIBMagnetPSUOut("DM2", GPIBSource)
'        k = GPIBMagnetPSUOut("LI 'DRNI', 'DRNV'", GPIBSource)

'        messageout$ = "SS VC1, " + N$ + "," + Str$(IC)
'        For X = 0 To 2 ^ Data.KeithleyFilter - 1
'            messageout$ = messageout$ + ", " + Str$(V)
'        Next X
        messageout$ = "SS VC" + N$ + ", " + Str$(V) + ", " + Str$(IC)
        k = GPIBMagnetPSUOut(messageout$, GPIBSource)
        k = GPIBMagnetPSUOut("SM DM2", GPIBSource)
        k = GPIBMagnetPSUOut("LI 'DRNI', 'DRNV'", GPIBSource)
        k = GPIBMagnetPSUOut("MD ME1", GPIBSource)
    End If
End Sub
Public Function SMUMeasureI() As String
SMUMeasureIstart:
    serialpoll% = 0
    If SourcePresent Then
        timemeasure = Timer
        Do
            ReadStatusByte 0, GPIBSource, serialpoll%
'            DoEvents
'            If Abort Then Exit Function
            If Timer > timemeasure + 20 Then serialpoll% = 1
        Loop Until serialpoll% And 1
'        Sleep 100
            messageout$ = "DO 'DRNI'"
'            ReadStatusByte 0, GPIBSource, SerialPoll%
'            If SerialPoll% And 64 Then GoTo SMUMeasureIstart
            k = GPIBMagnetPSUOut(messageout$, GPIBSource)
            SMUMeasureI = GPIBMagnetPSUIn(GPIBSource) + ","
        If MagnetoresistanceForm.TimeDependenceCheck Then
            messageout$ = "DO 'CH1T'"
            k = GPIBMagnetPSUOut(messageout$, GPIBSource)
            smumeasureitime = GPIBMagnetPSUIn(GPIBSource) + ","
        End If
        For X = 0 To 2 ^ Data.KeithleyFilter - 1
            First$ = Left$(SMUMeasureI, InStr(SMUMeasureI, ",") - 1)
            IValue(X) = Val(Right$(First$, Len(First$) - 1))
            SMUMeasureI = Right$(SMUMeasureI, Len(SMUMeasureI) - InStr(SMUMeasureI, ","))
            If MagnetoresistanceForm.TimeDependenceCheck Then
                FirstTime$ = Left$(smumeasureitime, InStr(smumeasureitime, ",") - 1)
                ITime(X) = Val(FirstTime$)
                smumeasureitime = Right$(smumeasureitime, Len(smumeasureitime) - InStr(smumeasureitime, ","))
            End If
        Next
    End If
End Function
Public Function ReadCurrent()
    Dim buffer As String
    If SMUType = "4200" Then
        SMUMeasureI
        ReadCurrent = 0
        For X = 0 To 2 ^ Data.KeithleyFilter - 1
            If Abort Then GoTo EndReadCurrent
                ReadCurrent = ReadCurrent + IValue(X)
            Next
        ReadCurrent = ReadCurrent / 2 ^ Data.KeithleyFilter
    Else
        buffer = "G4,2,0"
        GPIBSourceOut GPIBSource, buffer
        GPIBSourceIn GPIBSource, buffer
        ReadCurrent = Val(buffer)
    End If
EndReadCurrent:
End Function
Public Sub AbortAll()
Dim buffer As String
    DoEvents
    If SMUType = "4200" Then
        messageout$ = "MD ME4"
        k = GPIBMagnetPSUOut(messageout$, GPIBSource)
        messageout$ = "SS BC"
        k = GPIBMagnetPSUOut(messageout$, GPIBSource)
    Else
        SourceVoltage (0)
        buffer = "N0X"
        GPIBSourceOut GPIBSource, buffer
    End If
    DisplayVoltage (0)
    DisplayCurrent (0)
    DisplayIntensity (0)
End Sub
Public Sub SetUSMode()
    DoEvents
    If SourcePresent Then
        messageout$ = "US"
        k = GPIBMagnetPSUOut(messageout$, GPIBSource)
    End If
End Sub
Public Sub FindNewport()
Dim buffer As String
    Dim AddressList(3) As Integer
    AddressList(1) = GPIBSource
    AddressList(2) = GPIBNewport
    AddressList(3) = NOADDR
    NewportPresent = True
FindNewportAgain:
    Do
        DoEvents
        EnableRemote GPIB0, AddressList
    Loop Until (ibsta <> &H8000)
    Call DevClear(GPIB0, GPIBNewport)
'
    Sleep 1000
    buffer = "*IDN?"
    GPIBOut GPIBNewport, buffer
    Sleep 1000
    GPIBmessage = GPIBMagnetPSUIn(GPIBNewport)
'
'    GPIBIn GPIBNewport, buffer
'    GPIBmessage = Left$(buffer, 18)
    If Left$(GPIBmessage, 7) = "NewportCorp,1835-C" Then
        NewportPresent = 1
' Clear input buffer
        GPIBIn GPIBNewport, buffer
        fgh$ = buffer
'   Set the autoranging ON
        buffer = "AUTO 1"
        GPIBOut GPIBNewport, buffer
'   set the FILTER to NO FILTER
        buffer = "FILTER 1"
        GPIBOut GPIBNewport, buffer
'   set the ZERO to OFF
        buffer = "ZERO 0"
        GPIBOut GPIBNewport, buffer
'   set the ZERO to OFF
        buffer = "ATTN 0"
        GPIBOut GPIBNewport, buffer
'   Find the current calibration wavelength
        Sleep 500
        buffer = "LAMBDA?"
        GPIBOut GPIBNewport, buffer
        GPIBIn GPIBNewport, buffer
        Data.Wavelength = Val(buffer)
'   check newport setup
        buffer = "DETMODEL?"
        GPIBOut GPIBNewport, buffer
        GPIBIn GPIBNewport, buffer
        Data.DetectorModel = Mid$(buffer, 2, 6)
        buffer = "AUTO?"
        GPIBOut GPIBNewport, buffer
        GPIBIn GPIBNewport, buffer
        Data.Sensitivity = Val(buffer) - 1
        buffer = "FILTER?"
        GPIBOut GPIBNewport, buffer
        GPIBIn GPIBNewport, buffer
        Data.NewportFilter = Val(buffer)
        buffer = "ZERO?"
        GPIBOut GPIBNewport, buffer
        GPIBIn GPIBNewport, buffer
        Data.Zero = Val(buffer)
        buffer = "ATTN?"
        GPIBOut GPIBNewport, buffer
        GPIBIn GPIBNewport, buffer
        Data.Attenuator = Val(buffer)
     ElseIf Left$(GPIBmessage, 18) = "Newport Corp.,1830" Then
        NewportPresent = 2
' Clear input buffer
        GPIBIn GPIBNewport, buffer
        fgh$ = buffer
'   Set the autoranging ON
        buffer = "R0"
        GPIBOut GPIBNewport, buffer
'   set the FILTER to NO FILTER
        buffer = "F2"
        GPIBOut GPIBNewport, buffer
'   set the ZERO to OFF
        buffer = "Z0"
        GPIBOut GPIBNewport, buffer
'   set the ZERO to OFF
        buffer = "A0"
        GPIBOut GPIBNewport, buffer
'   set the Units to Watts
        buffer = "U1"
        GPIBOut GPIBNewport, buffer
'   Find the current calibration wavelength
        Sleep 500
        buffer = "W?"
        GPIBOut GPIBNewport, buffer
        GPIBIn GPIBNewport, buffer
        Data.Wavelength = Val(Left$(buffer, 4))
'   check newport setup
'        buffer = "DETMODEL?"
'        GPIBOut GPIBNewport, buffer
'        GPIBIn GPIBNewport, buffer
'        Data.DetectorModel = Mid$(buffer, 2, 6)
        buffer = "R?"
        GPIBOut GPIBNewport, buffer
        GPIBIn GPIBNewport, buffer
        Data.Sensitivity = Val(buffer) - 1
        buffer = "F?"
        GPIBOut GPIBNewport, buffer
        GPIBIn GPIBNewport, buffer
        Data.NewportFilter = Val(buffer)
        buffer = "Z?"
        GPIBOut GPIBNewport, buffer
        GPIBIn GPIBNewport, buffer
        Data.Zero = Val(buffer)
        buffer = "A?"
        GPIBOut GPIBNewport, buffer
        GPIBIn GPIBNewport, buffer
        Data.Attenuator = Val(buffer)
   
    Else
        NewportPresent = False
    End If
    If NewportPresent = False Then
        k = MsgBox("Optical Power Meter not found. Do you want to try again?", vbYesNo)
        If k = 6 Then
            NewportPresent = True
            SendIFC (GPIB0)
            GoTo FindNewportAgain
        End If
    End If

End Sub
Public Sub SourceStandby()
Dim buffer As String
'    buffer = "N0X"
'    GPIBSourceOut GPIBSource, buffer
End Sub

Public Function ReadIntensity()
    Dim buffer As String
ReadAgain:
    If NewportPresent = 1 Then
        buffer = "RWS?"
    Else
        buffer = "D?"
    End If
    GPIBOut GPIBNewport, buffer
    Sleep 100
    GPIBIn GPIBNewport, buffer
    If Val(buffer) > 2 Then
        Sleep 100
        GoTo ReadAgain
    End If
    If NewportPresent = 1 Then
        buffer = Right$(buffer, 18)
        ReadIntensity = Val(Left$(buffer, InStr(buffer, Chr$(10)) - 1))
    Else
        buffer = Left$(buffer, 18)
        ReadIntensity = Val(Left$(buffer, InStr(buffer, Chr$(10)) - 1))
    End If
End Function

Public Sub PrintIVAxis()
'
' Plot IV axis
'
    Printer.Scale (-0.5, 2)-(1.5, -0.5)
'
    If Data.LowVoltage >= 0 Then
        IAxisOffset = 0
    Else
        IAxisOffset = Abs(Data.LowVoltage) / (Data.HighVoltage - Data.LowVoltage)
    End If
    If Data.HighVoltage <= 0 Then IAxisOffset = 1
'
    If Data.CurrentLogLin = 1 Then
        VAxisOffset = 0
        PrintIVYLogAxisI
    Else
        If Data.DisplayCurrentHigh = 0 And Data.DisplayCurrentLow = 0 Then
            Data.DisplayCurrentHigh = 0.001
        End If

        If Data.DisplayCurrentLow >= 0 Then
            VAxisOffset = 0
        Else
            VAxisOffset = Abs(Data.DisplayCurrentLow) / (Data.DisplayCurrentHigh - Data.DisplayCurrentLow)
        End If
        If Data.DisplayCurrentHigh <= 0 Then VAxisOffset = 1
        PrintIVYAxisI
    End If
'
    PrintIVXAxisV

End Sub

Public Sub PrintIVYLogAxisI()
    Printer.Line (IAxisOffset, 0)-(IAxisOffset, 1)
'
    Dim low As Single, high As Single
    If Data.DisplayCurrentLow <= 0 Then
        low = 0.000000001
    Else
        low = 10 ^ Int(Log10(Abs(Data.DisplayCurrentLow)))
    End If
    If Data.DisplayCurrentHigh <= 0 Then
        high = 10
    Else
        high = 10 ^ (Int(Log10(Abs(Data.DisplayCurrentHigh - 0.0001)) + 1))
    End If
    LogLow = Log10(low)
    LogHigh = Log10(high)
    tx = CInt(Log10(high / low))
'
    If IAxisOffset = 1 Then
        Ticks = -1
    Else
        Ticks = 1
    End If
'
    For yl = 0 To tx
        Printer.Line (IAxisOffset, yl / tx)-(IAxisOffset - 0.02 * Ticks, yl / tx)
        out$ = "10"
        If Ticks = 1 Then
            Printer.CurrentX = Printer.CurrentX - 0.1
        End If
        Printer.CurrentY = Printer.CurrentY + 0.017
        Printer.Print out$;
        Printer.CurrentY = Printer.CurrentY + 0.014
        out$ = Str$(Int(Log10(low * 10 ^ yl) + 0.5))
        Printer.FontSize = Printer.FontSize - 1
        Printer.Print out$;
        Printer.FontSize = Printer.FontSize + 1
        If yl = tx Then GoTo ignor3
'
        For yll = 2 To 9
            yp = (yl + Log10(yll)) / tx
            Printer.Line (IAxisOffset, yp)-(IAxisOffset - 0.01 * Ticks, yp)
        Next
'
ignor3:
    Next
'
    Printer.CurrentY = 1.08
    Printer.CurrentX = IAxisOffset - 0.1
    Printer.Print "I (mA)";
'

End Sub

Public Sub PrintIVYAxisI()
    Printer.Line (IAxisOffset, 0)-(IAxisOffset, 1)
'
    high = Data.DisplayCurrentHigh
    low = Data.DisplayCurrentLow
    If Abs(high) > Abs(low) Then
        exponent = Int(Log10(Abs(high)))
    Else
        exponent = Int(Log10(Abs(low)))
    End If
    high = high / (10 ^ exponent)
    low = low / (10 ^ exponent)
    If IAxisOffset = 1 Then
        Ticks = -1
    Else
        Ticks = 1
    End If
'
    c = high - low: lc = Log10(c): pow = Int(lc): inc = lc - pow
'
' calculate position of ticks on y-axis
'
    a = 0.05
    If inc > Log10(1.6) Then a = 0.1
    If inc > Log10(4) Then a = 0.2
    If inc > Log10(8) Then a = 0.5
    tx = a * 10 ^ pow
'
    For Xh = Int(low / tx) * tx To Int(high / tx) * tx + tx Step tx * Sgn(high - low)
        If (Xh < low) Or (Xh > high) Then GoTo ignor1
        xnc = (Xh - low) / (high - low)
        Printer.Line (IAxisOffset, xnc)-(IAxisOffset - 0.01 * Ticks, xnc)
ignor1:
    Next
'
' calculate positions of numbers on y-axis
'
    a = 0.2
    If inc > Log10(1.6) Then a = 0.5
    If inc > Log10(4) Then a = 1
    If inc > Log10(8) Then a = 2
    tx = a * 10 ^ pow
'
    For Xg# = Int(low / tx) * tx - tx To Int(high / tx) * tx + tx Step tx * Sgn(high - low)
        temp! = CSng(Xg#)
        If (temp! + 0.000001 < low) Or (temp! > high) Then GoTo ignor2
        xnc = (temp! - low) / (high - low)
        Printer.Line (IAxisOffset, xnc)-(IAxisOffset - 0.02 * Ticks, xnc)
        out$ = Format$(temp!, "#0.0")
        If Ticks = 1 Then
            Printer.CurrentX = Printer.CurrentX - 0.03 * Len(out$)
        End If
        Printer.CurrentY = Printer.CurrentY + 0.017
        If Not (IAxisOffset > 0 And IAxisOffset < 1 And Val(out$) = 0) Then
            Printer.Print out$
        End If
ignor2:
    Next
    If VAxisOffset = 1 Then
        Printer.CurrentY = -0.01
    Else
        Printer.CurrentY = 1.05
    End If
    Printer.CurrentX = IAxisOffset - 0.1
    Printer.Print "I x10";
    Printer.CurrentY = Printer.CurrentY + 0.017
    Printer.Print Trim(exponent);
    Printer.CurrentY = Printer.CurrentY - 0.017
    Printer.Print "A";


End Sub

Public Sub PrintIVXAxisV()
    Printer.Line (0, VAxisOffset)-(1, VAxisOffset)
'
    If VAxisOffset = 1 Then
        Ticks = -1
    Else
        Ticks = 1
    End If
'
    c = Data.HighVoltage - Data.LowVoltage: lc = Log10(c): pow = Int(lc): inc = lc - pow
'
' calculate position of ticks on x-axis
'
    a = 0.05
    If inc > Log10(1.6) Then a = 0.1
    If inc > Log10(4) Then a = 0.2
    If inc > Log10(8) Then a = 0.5
    tx = a * 10 ^ pow
'
    For Xh = Int(Data.LowVoltage / tx) * tx To Int(Data.HighVoltage / tx) * tx + tx Step tx * Sgn(Data.HighVoltage - Data.LowVoltage)
        If (Xh < Data.LowVoltage) Or (Xh > Data.HighVoltage) Then GoTo ignor1
        xnc = (Xh - Data.LowVoltage) / (Data.HighVoltage - Data.LowVoltage)
        Printer.Line (xnc, VAxisOffset)-(xnc, VAxisOffset - 0.01 * Ticks)
ignor1:
    Next
'
' calculate positions of numbers on x-axis
'
    a = 0.2
    If inc > Log10(1.6) Then a = 0.5
    If inc > Log10(4) Then a = 1
    If inc > Log10(8) Then a = 2
    tx = a * 10 ^ pow
'
    For Xg# = Int(Data.LowVoltage / tx) * tx - tx To Int(Data.HighVoltage / tx) * tx + tx Step tx * Sgn(Data.HighVoltage - Data.LowVoltage)
        temp! = CSng(Xg#)
        If (temp! + 0.000001 < Data.LowVoltage) Or (temp! > Data.HighVoltage) Then GoTo ignor2
        xnc = (temp! - Data.LowVoltage) / (Data.HighVoltage - Data.LowVoltage)
        Printer.Line (xnc, VAxisOffset)-(xnc, VAxisOffset - 0.02 * Ticks)
        If Ticks = -1 Then
            Printer.CurrentY = Printer.CurrentY + 0.035
        End If
        out$ = Format$(temp!)
        Printer.CurrentX = Printer.CurrentX - 0.012 * Len(out$)
        If Not (VAxisOffset > 0 And VAxisOffset < 1 And Val(out$) = 0) Then
            Printer.Print out$
        End If
        
'        IntensityForm.PictureIV.Print out$
ignor2:
    Next
    If IAxisOffset = 1 Then
        Printer.CurrentX = -0.03
    Else
        Printer.CurrentX = 1.03
    End If
    Printer.CurrentY = VAxisOffset + 0.015
    Printer.Print "V"
'

End Sub

Public Sub PrintPlotIV()
    If Data.NumberOfPoints = 0 Then Exit Sub
    For z = 0 To Data.NumberOfPoints
        xp = (Data.Voltage(z) - Data.LowVoltage) / (Data.HighVoltage - Data.LowVoltage)
'
' replot current data
'
        If Data.CurrentLogLin = 0 Then
            yp = (Data.Current(z) - Data.DisplayCurrentLow) / (Data.DisplayCurrentHigh - Data.DisplayCurrentLow)
        Else
            If Data.Current(z) = 0 Then
                yp = 0
            Else
                yp = (Log10(Abs(Data.Current(z))) - LogLow) / (LogHigh - LogLow)
            End If
        End If
        If yp >= 0 And yp <= 1 Then
            Printer.Circle (xp, yp), 0.005
        End If
'
    Next
End Sub

Public Sub PrintLVAxis()
'
' Plot Intensity Axis
'
       Printer.Scale (-0.5, 2)-(1.5, -0.5)
'
        If Data.LowVoltage >= 0 Then
            IAxisOffset = 0
        Else
            IAxisOffset = Abs(Data.LowVoltage) / (Data.HighVoltage - Data.LowVoltage)
        End If
        If Data.HighVoltage <= 0 Then IAxisOffset = 1
'
        If Data.IntensityLogLin = 1 Then
            VAxisOffset = 0
            PrintVLYLogAxisL
        Else
            If Data.DisplayIntensityLow >= 0 Then
                VAxisOffset = 0
            Else
                VAxisOffset = Abs(Data.DisplayIntensityLow) / (Data.DisplayIntensityHigh - Data.DisplayIntensityLow)
            End If
            If Data.DisplayIntensityHigh <= 0 Then VAxisOffset = 1
            PrintVLYAxisL
        End If
'
        PrintVLXAxisV
End Sub

Public Sub PrintVLYLogAxisL()
    Printer.Line (IAxisOffset, 0)-(IAxisOffset, 1)
'
    Dim low As Single, high As Single
    If Data.DisplayIntensityLow <= 0 Then
        low = 0.000000000001
    Else
        low = 10 ^ Int(Log10(Abs(Data.DisplayIntensityLow)))
    End If
    If Data.DisplayIntensityHigh <= 0 Then
        high = 10
    Else
        high = 10 ^ (Int(Log10(Abs(Data.DisplayIntensityHigh - 0.0001)) + 1))
    End If
    LLogLow = Log10(low)
    LLogHigh = Log10(high)
    tx = CInt(Log10(high / low))
'
    If IAxisOffset = 1 Then
        Ticks = -1
    Else
        Ticks = 1
    End If
'
    For yl = 0 To tx
        Printer.Line (IAxisOffset, yl / tx)-(IAxisOffset - 0.02 * Ticks, yl / tx)
        out$ = "10"
        If Ticks = 1 Then
            Printer.CurrentX = Printer.CurrentX - 0.1
        End If
        Printer.CurrentY = Printer.CurrentY + 0.017
        Printer.Print out$;
        Printer.CurrentY = Printer.CurrentY + 0.014
        out$ = Str$(Int(Log10(low * 10 ^ yl) + 0.5))
        Printer.FontSize = Printer.FontSize - 1
        Printer.Print out$;
        Printer.FontSize = Printer.FontSize + 1
        If yl = tx Then GoTo ignor3
'
        For yll = 2 To 9
            yp = (yl + Log10(yll)) / tx
            Printer.Line (IAxisOffset, yp)-(IAxisOffset - 0.01 * Ticks, yp)
        Next
'
ignor3:
    Next
'
    Printer.CurrentY = 1.08
    Printer.CurrentX = IAxisOffset - 0.1
    Printer.Print "L (W)";
'

End Sub

Public Sub PrintVLYAxisL()
    Printer.Line (IAxisOffset, 0)-(IAxisOffset, 1)
'
    high = Data.DisplayIntensityHigh
    low = Data.DisplayIntensityLow
    If Abs(high) > Abs(low) Then
        exponent = Int(Log10(Abs(high)))
    Else
        exponent = Int(Log10(Abs(low)))
    End If
    high = high / (10 ^ exponent)
    low = low / (10 ^ exponent)
    If IAxisOffset = 1 Then
        Ticks = -1
    Else
        Ticks = 1
    End If
'
    c = high - low: lc = Log10(c): pow = Int(lc): inc = lc - pow
'
' calculate position of ticks on y-axis
'
    a = 0.05
    If inc > Log10(1.6) Then a = 0.1
    If inc > Log10(4) Then a = 0.2
    If inc > Log10(8) Then a = 0.5
    tx = a * 10 ^ pow
'
    For Xh = Int(low / tx) * tx To Int(high / tx) * tx + tx Step tx * Sgn(high - low)
        If (Xh < low) Or (Xh > high) Then GoTo ignor1
        xnc = (Xh - low) / (high - low)
        Printer.Line (IAxisOffset, xnc)-(IAxisOffset - 0.01 * Ticks, xnc)
ignor1:
    Next
'
' calculate positions of numbers on y-axis
'
    a = 0.2
    If inc > Log10(1.6) Then a = 0.5
    If inc > Log10(4) Then a = 1
    If inc > Log10(8) Then a = 2
    tx = a * 10 ^ pow
'
    For Xg# = Int(low / tx) * tx - tx To Int(high / tx) * tx + tx Step tx * Sgn(high - low)
        temp! = CSng(Xg#)
        If (temp! + 0.000001 < low) Or (temp! > high) Then GoTo ignor2
        xnc = (temp! - low) / (high - low)
        Printer.Line (IAxisOffset, xnc)-(IAxisOffset - 0.02 * Ticks, xnc)
        out$ = Format$(temp!, "#0.0")
        If Ticks = 1 Then
            Printer.CurrentX = Printer.CurrentX - 0.03 * Len(out$)
        End If
        Printer.CurrentY = Printer.CurrentY + 0.017
        If Not (IAxisOffset > 0 And IAxisOffset < 1 And Val(out$) = 0) Then
            Printer.Print out$
        End If
ignor2:
    Next
    If VAxisOffset = 1 Then
        Printer.CurrentY = -0.01
    Else
        Printer.CurrentY = 1.05
    End If
    Printer.CurrentX = IAxisOffset - 0.1
    Printer.Print "L x10";
    Printer.CurrentY = Printer.CurrentY + 0.017
    Printer.Print Trim(exponent);
    Printer.CurrentY = Printer.CurrentY - 0.017
    Printer.Print "W";

End Sub

Public Sub PrintVLXAxisV()
    Printer.Line (0, VAxisOffset)-(1, VAxisOffset)
'
    If VAxisOffset = 1 Then
        Ticks = -1
    Else
        Ticks = 1
    End If
'
    c = Data.HighVoltage - Data.LowVoltage: lc = Log10(c): pow = Int(lc): inc = lc - pow
'
' calculate position of ticks on x-axis
'
    a = 0.05
    If inc > Log10(1.6) Then a = 0.1
    If inc > Log10(4) Then a = 0.2
    If inc > Log10(8) Then a = 0.5
    tx = a * 10 ^ pow
'
    For Xh = Int(Data.LowVoltage / tx) * tx To Int(Data.HighVoltage / tx) * tx + tx Step tx * Sgn(Data.HighVoltage - Data.LowVoltage)
        If (Xh < Data.LowVoltage) Or (Xh > Data.HighVoltage) Then GoTo ignor1
        xnc = (Xh - Data.LowVoltage) / (Data.HighVoltage - Data.LowVoltage)
        Printer.Line (xnc, VAxisOffset)-(xnc, VAxisOffset - 0.01 * Ticks)
ignor1:
    Next
'
' calculate positions of numbers on x-axis
'
    a = 0.2
    If inc > Log10(1.6) Then a = 0.5
    If inc > Log10(4) Then a = 1
    If inc > Log10(8) Then a = 2
    tx = a * 10 ^ pow
'
    For Xg# = Int(Data.LowVoltage / tx) * tx - tx To Int(Data.HighVoltage / tx) * tx + tx Step tx * Sgn(Data.HighVoltage - Data.LowVoltage)
        temp! = CSng(Xg#)
        If (temp! + 0.000001 < Data.LowVoltage) Or (temp! > Data.HighVoltage) Then GoTo ignor2
        xnc = (temp! - Data.LowVoltage) / (Data.HighVoltage - Data.LowVoltage)
        Printer.Line (xnc, VAxisOffset)-(xnc, VAxisOffset - 0.02 * Ticks)
        If Ticks = -1 Then
            Printer.CurrentY = Printer.CurrentY + 0.035
        End If
        out$ = Format$(temp!)
        Printer.CurrentX = Printer.CurrentX - 0.012 * Len(out$)
        If Not (VAxisOffset > 0 And VAxisOffset < 1 And Val(out$) = 0) Then
            Printer.Print out$
        End If
        
'        IntensityForm.PictureIV.Print out$
ignor2:
    Next
    If IAxisOffset = 1 Then
        Printer.CurrentX = -0.03
    Else
        Printer.CurrentX = 1.03
    End If
    Printer.CurrentY = VAxisOffset + 0.015
    Printer.Print "V"
'

End Sub

Public Sub PrintPlotLV()
    If Data.NumberOfPoints = 0 Then Exit Sub
    For z = 0 To Data.NumberOfPoints
        xp = (Data.Voltage(z) - Data.LowVoltage) / (Data.HighVoltage - Data.LowVoltage)

'
' replot intensity data
'
        If Data.IntensityLogLin = 0 Then
            yp = (Data.Intensity(z) - Data.DisplayIntensityLow) / (Data.DisplayIntensityHigh - Data.DisplayIntensityLow)
        Else
            If Data.Intensity(z) = 0 Then
                yp = 0
            Else
                yp = (Log10(Abs(Data.Intensity(z))) - LLogLow) / (LLogHigh - LLogLow)
            End If
        End If
        If yp >= 0 And yp <= 1 Then
            Printer.Circle (xp, yp), 0.005
        End If
    Next

End Sub

Public Function DoNullMeasurement(MagI, V, NumRepeats, Illumination, NormalMR, DiodeCurrent)
'Not used in Lakeshore version
'
    MagnetOff
    ReverseField
    SetMagI (MagI)
    Sleep 3000
    If NormalMR = 1 Then
'        If SourcePresent Then
'            SourceVoltage (V)
'        End If
        DisplayVoltage (V)
'        currenttime = Timer
'        Do Until Timer > currenttime + SettlingTimeValue / 1000
        DoEvents
        If Abort Then
'            h = MsgBox("Do you want to ABORT?", vbYesNo)
'            If h = 6 Then
                DoNullMeasurement = 1
                GoTo NullFieldAbort
'            End If
'            Abort = False
        End If
'        Loop
        If SourcePresent Then
            Current = 0
'            For X = 1 To NumRepeats
repeatmeasurement1:
                SourceVoltage (V)
                If Abort Then
                    DoNullMeasurement = 1
                    GoTo NullFieldAbort
                End If
                tempcurrent = ReadCurrent
'                If Abs(tempcurrent) = 0.000001 Then GoTo repeatmeasurement1
'                If Abs(tempcurrent) = 0.000002 Then GoTo repeatmeasurement1
'                If Abs(tempcurrent) = 0.000003 Then GoTo repeatmeasurement1
'                If Abs(tempcurrent) = 0.000004 Then GoTo repeatmeasurement1
'                If Abs(tempcurrent) = 0.000005 Then GoTo repeatmeasurement1
'                If Abs(tempcurrent) = 0.000006 Then GoTo repeatmeasurement1
'                If Abs(tempcurrent) = 0.000007 Then GoTo repeatmeasurement1
'                If Abs(tempcurrent) = 0.000008 Then GoTo repeatmeasurement1
'                If Abs(tempcurrent) = 0.000009 Then GoTo repeatmeasurement1
                Current = Current + tempcurrent
'            Next
'            Current = Current / NumRepeats
        End If
'
        If NewportPresent Then
            L = ReadIntensity
        End If
        If Abort Then
            DoNullMeasurement = 1
            GoTo NullFieldAbort
        End If
        
        SourceVoltage (0)
        If Abort Then
            DoNullMeasurement = 1
            GoTo NullFieldAbort
        End If
        temp$ = SMUFlushI
        DisplayCurrent (Current)
        DisplayIntensity (L)
        Open GlobalPathName + GlobalFileName For Append As #disc
            Print #disc, 0, 0, Current, L
            Print #disc, Format(0, "#0.0"), Format(Current, "#0.0000000000"), Format(L, "#0.0000000000000")
            If MagnetoresistanceForm.TimeDependenceCheck Then
                For Y = 0 To 2 ^ Data.KeithleyFilter - 1
                    Print #disc, Format(ITime(Y), "#0.000000000000000"), Format(IValue(Y), "#0.000000000000000"),
                Next

'                For Y = 0 To NumVvalues
'                    Print #disc, Format(ITime(Y), "#0.000000000000000"), Format(Vvalues(Y), "#0.000000000000000")
'                Next Y
            End If
            Print #disc, ""
        Close #disc
    End If
    
    If Illumination = 1 Then
        'Turn on LED
        SetLEDI (DiodeCurrent)
'        If SourcePresent Then
'            SourceVoltage (V)
'        End If
        DisplayVoltage (V)
'        currenttime = Timer
'        Do Until Timer > currenttime + SettlingTimeValue / 1000
        DoEvents
        If Abort Then
'            h = MsgBox("Do you want to ABORT?", vbYesNo)
'            If h = 6 Then
                DoNullMeasurement = 1
                GoTo NullFieldAbort
'            End If
'            Abort = False
        End If
'        Loop
        If SourcePresent Then
            Current = 0
'            For X = 1 To NumRepeats
'repeatmeasurement2:
            If Abort Then
                DoNullMeasurement = 1
                    GoTo NullFieldAbort
            End If
            SourceVoltage (V)
            If Abort Then
                DoNullMeasurement = 1
                GoTo NullFieldAbort
            End If
            tempcurrent = ReadCurrent
'            If Abs(tempcurrent) = 0.000001 Then GoTo repeatmeasurement2
'            If Abs(tempcurrent) = 0.000002 Then GoTo repeatmeasurement2
'            If Abs(tempcurrent) = 0.000003 Then GoTo repeatmeasurement2
'            If Abs(tempcurrent) = 0.000004 Then GoTo repeatmeasurement2
'            If Abs(tempcurrent) = 0.000005 Then GoTo repeatmeasurement2
'            If Abs(tempcurrent) = 0.000006 Then GoTo repeatmeasurement2
'            If Abs(tempcurrent) = 0.000007 Then GoTo repeatmeasurement2
'            If Abs(tempcurrent) = 0.000008 Then GoTo repeatmeasurement2
'            If Abs(tempcurrent) = 0.000009 Then GoTo repeatmeasurement2
            Current = Current + tempcurrent
'            Next
'            Current = Current / NumRepeats
        End If
'
        If Abort Then
            DoNullMeasurement = 1
            GoTo NullFieldAbort
        End If
        SourceVoltage (0)
        If Abort Then
            DoNullMeasurement = 1
            GoTo NullFieldAbort
        End If
        temp$ = SMUFlushI
        'Turn Off LED
        LEDOff
        DisplayCurrent (Current)
        DisplayIntensity (L)
        Open GlobalPathName + IGlobalFileName For Append As #disc
            Print #disc, 0, 0, Current, 0
            If MagnetoresistanceForm.TimeDependenceCheck Then
                For Y = 0 To NumVvalues
                    Print #disc, Format(ITime(Y), "#0.000000000000000"), Format(Vvalues(Y), "#0.000000000000000")
                Next Y
            End If
            Print #disc, ""
        Close #disc
    End If
    DoNullMeasurement = 0
NullFieldAbort:
    MagnetOff
End Function

Public Function DoFieldMeasurement(Bfieldvalue, V, NumRepeats, Illumination, NormalMR, DiodeCurrent)
'    If MagI < 0 Then
'        ReverseField
'        SetMagI (MagI * -1)
'    Else
        SetMagI (Bfieldvalue)
        ramp = CheckBRamp(Bfieldvalue)
        Do Until ramp = True
            DoEvents
            If Abort Then
                    DoFieldMeasurement = 1
                    GoTo FieldAbort
            End If
            ramp = CheckBRamp(Bfieldvalue)
        Loop
    Sleep 1000
    GPIBOut GPIBGaussmeter, "RDGFIELD?"
    field = Val(GPIBMagnetPSUIn(GPIBGaussmeter)) * 1000
        
'    End If
'    Sleep 3000
    If NormalMR = 1 Then
'        If SourcePresent Then
'            SourceVoltage (V)
'        End If
        DisplayVoltage (V)
'        currenttime = Timer
'        Do Until Timer > currenttime + SettlingTimeValue / 1000
        DoEvents
        If Abort Then
'            h = MsgBox("Do you want to ABORT?", vbYesNo)
'            If h = 6 Then
                DoFieldMeasurement = 1
                GoTo FieldAbort
'            End If
'            Abort = False
        End If
'        Loop
        If SourcePresent Then
            Current = 0
                If Abort Then
                    DoFieldMeasurement = 1
                    GoTo FieldAbort
                End If
                SourceVoltage (V)
                Sleep (1000)
'Read Intensity immediatly after setting voltage
                If NewportPresent Then
                    L = ReadIntensity
                    If L = 0 Then L = ReadIntensity
                End If
'
                If Abort Then
                    DoFieldMeasurement = 1
                    GoTo FieldAbort
                End If
                Current = ReadCurrent
        End If
'
'        If NewportPresent Then
'            L = ReadIntensity
'        End If
        If Abort Then
            DoFieldMeasurement = 1
            GoTo FieldAbort
        End If
        SourceVoltage (0)
        If Abort Then
            DoFieldMeasurement = 1
            GoTo FieldAbort
        End If
        temp$ = SMUFlushI
'        DisplayCurrent (Current)
        DisplayIntensity (L)
        Open GlobalPathName + GlobalFileName For Append As #disc
            Print #disc, Format(field, "#0.00"), Format(Current, "#0.000000000000000"), Format(L, "#0.0000000000000"),
            If MagnetoresistanceForm.TimeDependenceCheck Then
                For Y = 0 To 2 ^ Data.KeithleyFilter - 1
                    Print #disc, Format(ITime(Y), "#0.000000000000000"), Format(IValue(Y), "#0.000000000000000"),
                Next

'                For Y = 0 To NumVvalues
'                    Print #disc, Format(ITime(Y), "#0.000000000000000"), Format(Vvalues(Y), "#0.000000000000000")
'                Next Y
            End If
            Print #disc, ""
        Close #disc
    End If
'
'    If Illumination = 1 Then
'        'Turn On LED
'        SetLEDI (DiodeCurrent)
'        If SourcePresent Then
'            SourceVoltage (v)
'        End If
'        DisplayVoltage (v)
'        currenttime = Timer
'        Do Until Timer > currenttime + SettlingTimeValue / 1000
'            DoEvents
'            If Abort Then
'                h = MsgBox("Do you want to ABORT?", vbYesNo)
'                If h = 6 Then
'                    DoFieldMeasurement = 1
'                    GoTo FieldAbort
'                End If
'                Abort = False
'            End If
'        Loop
'        If SourcePresent Then
'            Current = 0
'            For X = 1 To NumRepeats
'repeatmeasurement2:
'            SourceVoltage (v)
'            tempcurrent = ReadCurrent
'            If Abs(tempcurrent) = 0.000001 Then GoTo repeatmeasurement2
'            If Abs(tempcurrent) = 0.000002 Then GoTo repeatmeasurement2
'            If Abs(tempcurrent) = 0.000003 Then GoTo repeatmeasurement2
'            If Abs(tempcurrent) = 0.000004 Then GoTo repeatmeasurement2
'            If Abs(tempcurrent) = 0.000005 Then GoTo repeatmeasurement2
'            If Abs(tempcurrent) = 0.000006 Then GoTo repeatmeasurement2
'            If Abs(tempcurrent) = 0.000007 Then GoTo repeatmeasurement2
'            If Abs(tempcurrent) = 0.000008 Then GoTo repeatmeasurement2
'            If Abs(tempcurrent) = 0.000009 Then GoTo repeatmeasurement2
'            Current = Current + tempcurrent
'            Next
'            Current = Current / NumRepeats
'        End If
'
'        SourceVoltage (0)
'        'Turn Off LED
''        LEDOff
'        DisplayCurrent (Current)
'        DisplayIntensity (L)
'        Open GlobalPathName + IGlobalFileName For Append As #disc
'            Print #disc, Bfieldvalue, Current, 0
'            Print #disc, Format(Bfieldvalue, "#0.0"), Format(Current, "#0.0000000000"), 0
'        Close #disc
'    End If
    
    DoFieldMeasurement = 0
    Exit Function
FieldAbort:
    MagnetOff
End Function

Public Sub SetLEDI(I As Single)
    If MagnetPSUPresent = True Then
        GPIBMessageOut$ = "I2 " + Trim(I)
        k = GPIBMagnetPSUOut(GPIBMessageOut$, GPIBMagnetPSU)
        GPIBMessageOut$ = "OP2 1"
        k = GPIBMagnetPSUOut(GPIBMessageOut$, GPIBMagnetPSU)
    End If
End Sub

Public Sub LEDOff()
    If MagnetPSUPresent = True Then
        GPIBMessageOut$ = "OP2 0"
        k = GPIBMagnetPSUOut(GPIBMessageOut$, GPIBMagnetPSU)
    End If
End Sub

Public Sub GetBValues(Bvaluefilename)
Open Bvaluefilename For Input As #1
    NumBvalues = 0
    Do
    Input #1, Bfield(NumBvalues)
    NumBvalues = NumBvalues + 1
    Loop Until EOF(1)
Close #1
NumBvalues = NumBvalues - 1

End Sub
Public Sub GetVvalues(Vvaluefilename)
Open Vvaluefilename For Input As #1
    NumVvalues = 0
    Do
    Input #1, Vvalues(NumVvalues)
    NumVvalues = NumVvalues + 1
    Loop Until EOF(1)
Close #1
NumVvalues = NumVvalues - 1

End Sub
Public Function GetPSUTemperature() As Single
    GPIBOut GPIBMagnetPSU, "READ?"
    Do
        DoEvents
        GPIBmessage = GPIBMagnetPSUIn(GPIBMagnetPSU)
    Loop Until (ibsta <> &H8000)
    For hj = 1 To 9
    GPIBmessage = Right$(GPIBmessage, Len(GPIBmessage) - InStr(GPIBmessage, ","))
    Next
    GPIBmessage = Left$(GPIBmessage, InStr(GPIBmessage, ","))
GetPSUTemperature = Val(GPIBmessage)


End Function

Public Sub FindGaussmeter()
    Dim AddressList(2) As Integer, status As Integer, GPIBmessage As String
    AddressList(1) = GPIBGaussmeter
    AddressList(2) = NOADDR
FindGaussmeterAgain:
    GaussmeterPresent = False
    GPIBmessage = ""
    Do
        DoEvents
        EnableRemote GPIB0, AddressList
    Loop Until (ibsta <> &H8000)
    Call DevClear(GPIB0, GPIBGaussmeter)
    Call ReadStatusByte(GPIB0, GPIBGaussmeter, status)
    GPIBOut GPIBGaussmeter, "*IDN?"
    Call ReadStatusByte(GPIB0, GPIBGaussmeter, status)
    Do
        DoEvents
        GPIBmessage = GPIBMagnetPSUIn(GPIBGaussmeter)
    Loop Until (ibsta <> &H8000)

    
    If Left$(GPIBmessage, 13) = "LSCI,MODEL475" Then
        GaussmeterPresent = True
        GPIBmessage = ""
        Call DevClear(GPIB0, GPIBGaussmeter)
        Call ReadStatusByte(GPIB0, GPIBGaussmeter, status)
    End If
    If GaussmeterPresent = True Then
'       Set default IEEE parameters
    GPIBOut GPIBGaussmeter, "IEEE 0,0,12"
'       Set units to Tesla
    GPIBOut GPIBGaussmeter, "UNIT 2"
'        Set Field control OFF
    GPIBOut GPIBGaussmeter, "CMODE 0"
'       Set Field Control Parameters
    GPIBOut GPIBGaussmeter, "CPARAM 5,.3,0,600"
'       Set measurement mode, DC, 4 digits, RMS wideband, Peak periodic and positive
    GPIBOut GPIBGaussmeter, "RDGMODE 1,2,1,1,1"
'       Set field control value to 0mT
    SetMagI (0)
'
    Else
        k = MsgBox("Gaussmeter not found. Do you want to try again?", vbYesNo)
        If k = 6 Then
            SendIFC (GPIB0)
            GoTo FindGaussmeterAgain
        End If
    End If

End Sub

Public Function CheckBRamp(Bvalue) As Boolean
    GPIBOut GPIBGaussmeter, "RDGFIELD?"
    field = Val(GPIBMagnetPSUIn(GPIBGaussmeter)) * 1000
'    GPIBOut GPIBGaussmeter, "RAMPST?"
'    GPIBmessage = GPIBMagnetPSUIn(GPIBGaussmeter)
    diff = Abs(Bvalue - field)
    If diff < 0.05 Then
        CheckBRamp = True
    Else
        CheckBRamp = False
    End If

End Function

Public Function SMUFlushI() As String
SMUFlushIstart:
    serialpoll% = 0
    If SourcePresent Then
        flushtime = Timer
        Do
            ReadStatusByte 0, GPIBSource, serialpoll%
'            DoEvents
'            If Abort Then Exit Function
            If Timer > flushtime + 20 Then serialpoll% = 1
        Loop Until serialpoll% And 1
'        Sleep 100
            messageout$ = "DO 'DRNI'"
'            ReadStatusByte 0, GPIBSource, SerialPoll%
'            If SerialPoll% And 64 Then GoTo SMUMeasureIstart
            k = GPIBMagnetPSUOut(messageout$, GPIBSource)
            SMUFlushI = GPIBMagnetPSUIn(GPIBSource) + ","
    End If
End Function
