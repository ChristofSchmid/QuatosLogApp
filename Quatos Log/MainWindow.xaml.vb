Imports System.IO.Ports
Imports System.IO.File
Imports System
Imports System.Threading




Class MainWindow

    Public WithEvents MyComPort As SerialPort
    Public serText As String
    Public ComPortSelected As String '= "COM1"
    Public ComConnected As Boolean = False
    Public SysLoaded As Boolean = False
    Public TextLineCounter As UInteger = 0
    Public SerCmdAry As String()
    Public RunManual As Boolean = False


    Public Event DataReceived As IO.Ports.SerialDataReceivedEventHandler

    'Private Property Curr As Object


    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

        DisplayElements(False)
        MyComPort = New SerialPort

        MyComPort.BaudRate = 115200
        MyComPort.Parity = Parity.None
        MyComPort.StopBits = StopBits.One
        MyComPort.Encoding = System.Text.Encoding.ASCII
        MyComPort.DataBits = 8
        MyComPort.NewLine = vbLf


        GetSerialPortNames()
        If Not ComPort.SelectedItem = Nothing Then

            ComPortSelected = ComPort.SelectedItem.ToString()
            'MsgBox("-" + ComPortSelected + "-")
            MyComPort.PortName = ComPortSelected
            Try
                MyComPort.Open()
                BtnConnect.Content = "Disconnect"
            Catch ex As Exception

            End Try

        End If

        If MyComPort.IsOpen() Then
            BtnConnect.Content = "Disconnect"
            DisplayElements(True)
        Else
            'MsgBox("No Com Port available")
        End If
        BtnConnect.Focus()
        'UpdateSysInfos("SYS ;0 ;0 ;0 ;0 ;0 ;0 ;0 ;0 ;0 ;0 " + CChar(vbCr & vbLf))
        'UpdateSysInfos("hBt 0" + CChar(vbCr & vbLf))

        SysLoaded = True

        SerialCmdCombo.Items.Add("Set Prop Blades")
        SerialCmdCombo.Items.Add("Set Max RPM")
        SerialCmdCombo.Items.Add("Set Min PWM")
        SerialCmdCombo.Items.Add("Set Max PWM")

        SerialCmdCombo.Items.Add("Run PWM max")
        SerialCmdCombo.Items.Add("Run PWM min")

        SerialCmdCombo.Items.Add("Set ESC32 OL")
        SerialCmdCombo.Items.Add("Set ESC32 CL")
        SerialCmdCombo.Items.Add("Start Log")
        SerialCmdCombo.Items.Add("Arm ESC32")
        SerialCmdCombo.Items.Add("Start ESC32 OL")
        SerialCmdCombo.Items.Add("Start ESC32 CL")
        SerialCmdCombo.Items.Add("Stop Motor")
        SerialCmdCombo.Items.Add("Disarm ESC32")
        SerCmdAry = {"setmotorpoles", "setmaxrpm", "setpwmoffset", "setpwmmax", "runpwmmax", "runpwmmin", "setesc32ol;", "setesc32cl;", "start log;", "arm;", "startesc32ol;", "startesc32cl;", "stop;", "disarm;"}
        For i = 2 To 40 Step 2
            CmdValueCombo.Items.Add(CStr(i))
        Next
        'CmdValueCombo.SelectedIndex = 1 '2 Blades / 6 for 14 Poles
        SerialCmdCombo.SelectedIndex = 0

        For i = 2 To 10 Step 1 ' Prop Blades
            Combo_PropBlades.Items.Add(CStr(i))
        Next
        Combo_PropBlades.SelectedIndex = 0 '14 Poles

        'ElseIf SerCmdAry(SerialCmdCombo.SelectedIndex) = "setmaxrpm" Then


        For i = 2000 To 20000 Step 100
            Combo_MaxRpm.Items.Add(CStr(i))
        Next
        Combo_MaxRpm.SelectedIndex = 40 '6000

        'ElseIf SerCmdAry(SerialCmdCombo.SelectedIndex) = "setpwmoffset" Then

        For i = 950 To 1100 Step 10
            Combo_MinPwm.Items.Add(CStr(i))
        Next
        Combo_MinPwm.SelectedIndex = 5 '1000

        'ElseIf SerCmdAry(SerialCmdCombo.SelectedIndex) = "setpwmmax" Then


        For i = 1800 To 2150 Step 10
            Combo_MaxPwm.Items.Add(CStr(i))
        Next
        Combo_MaxPwm.SelectedIndex = 15 '1950


        For i = 500 To 2500 Step 500
            Combo_SetRpmStep.Items.Add(CStr(i))
        Next
        Combo_SetRpmStep.SelectedIndex = 0 '500

    End Sub



    'Private Sub MainWindow_StateChanged(sender As Object, e As EventArgs) Handles Me.StateChanged
    'Dim x As Double = Me.Width()
    'Dim y As Double = Me.Height()
    '  QuatSerText.Height = (y - 270)
    '  QuatSerText.Width = (x - 35)
    '  QuatSerText.UpdateLayout()

    '  StopButton.Width = (x - 35)
    '  StopButton.UpdateLayout()
    'End Sub

    Private Sub MainWindow_Resize(sender As Object, e As RoutedEventArgs) Handles Me.SizeChanged
        Dim x As Double = Me.Width()
        Dim y As Double = Me.Height()
        'QuatSerText.Height = ((y - 350) / y * (y - 35))
        'QuatSerText.Width = (x - 35)
        'QuatSerText.UpdateLayout()

        'StopButton.Width = (x - 35)
        'StopButton.UpdateLayout()




    End Sub

    Delegate Sub SetReceivedText(ByVal ReceivedData As String)
    Dim d As SetReceivedText



    'Serial

    Private Sub BtnConnect_Click(sender As Object, e As RoutedEventArgs) Handles BtnConnect.Click
        DisplayElements(False)
        If Not ComPort.HasItems Then
            GetSerialPortNames()
            If Not ComPort.HasItems Then
                MsgBox("No ComPort available")
                'DisplayElements(False)
                Return
            End If
        End If

        If ComPortSelected <> ComPort.SelectedItem.ToString() Then
            GetSerialPortNames()
        End If

        Try

            If Not MyComPort.IsOpen Then
                ComPortSelected = ComPort.SelectedItem.ToString()
                'MsgBox("-" + ComPortSelected + "-")
                MyComPort.PortName = ComPortSelected
                'MyComPort.PortName = ComPortSelected
                MyComPort.Open()
                BtnConnect.Content = "Disconnect"
                DisplayElements(True)

            Else
                BtnConnect.Content = "Connect"
                HeartBeat.Visibility = Windows.Visibility.Hidden
                MyComPort.Close()
                SendSerialTxt.Clear()
                'DisplayElements(False)
            End If


        Catch ex As Exception
            GetSerialPortNames()

        End Try



    End Sub

    Private Sub ComPort_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles ComPort.SelectionChanged

        If ComPort.HasItems() Then
            ComPortSelected = ComPort.SelectedItem.ToString()
        End If

    End Sub

    Private Sub GetSerialPortNames()
        Dim comcnt As Int16 = 0
        ComPort.Items.Clear()
        ' Show all available COM ports. 
        'ComPort.Items.Add("Keine Wahl..")
        For Each sp As String In My.Computer.Ports.SerialPortNames
            ComPort.Items.Add(sp)
            comcnt += 1

        Next
        If comcnt > 0 Then
            ComPort.SelectedItem = ComPort.Items(0)
        End If

    End Sub

    Private Sub QuatSerText_TextChanged(sender As Object, e As TextChangedEventArgs) Handles QuatSerText.TextChanged
        'MsgBox("text changed")
    End Sub

    Private Sub MyComPort_DataReceived(ByVal sender As System.Object, ByVal e As System.IO.Ports.SerialDataReceivedEventArgs) Handles MyComPort.DataReceived
        Me.Dispatcher.Invoke(Sub() ReceivedText(MyComPort.ReadLine()))
    End Sub

    Private Sub ReceivedText(ByVal [text] As String)
        Dim LenText As Integer = Len(QuatSerText.Text)
        If Not UpdateSysInfos(text) Then
            If LenText >= 3000 Then
                QuatSerText.Text = ResizeText(QuatSerText.Text)
                TextLineCounter -= 1
            End If

            QuatSerText.Text &= [text]
            QuatSerText.ScrollToEnd()
            TextLineCounter += 1

        End If
    End Sub

    'Serial Send 

    Private Sub SendSerialBtn_Click(sender As Object, e As RoutedEventArgs) Handles SendSerialBtn.Click

        Dim SerialCommand As String = SendSerialTxt.Text & ";" '& CChar(vbCr & vbLf)
        If Not SendSerialTxt.Text = "" Then
            MyComPort.Write(SerialCommand)
            Return
        End If
        If Not CheckSanity() Then
            Return
        End If
        If Not SerialCmdCombo.SelectedItem = Nothing Then
            'SerialCommand = SerialCmdCombo.Text & ";" '& CChar(vbCr & vbLf)
            SerialCommand = SerCmdAry(SerialCmdCombo.SelectedIndex)
            If SerialCmdCombo.SelectedIndex <= 3 Then 'Elements 0-3
                SerialCommand = SerialCommand + " " + CmdValueCombo.SelectedItem + ";" 'CmdValueAry(CmdValueCombo.SelectedIndex) + ";"
            End If

            'MsgBox(SerialCommand)
            MyComPort.Write(SerialCommand)
            'MyComPort.Write("settestedunit " + Strings.Left(Trim(TestUnitText.Text), 50) + ";")

        End If

        SendSerialBtn.IsDefault = False
        StopButton.IsDefault = True

    End Sub

    Private Sub SendSerialTxt_GotKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs) Handles SendSerialTxt.GotKeyboardFocus
        SendSerialTxt.Clear()
        StopButton.IsDefault = False
        SendSerialBtn.IsDefault = True
    End Sub

    Private Sub SendSerialTxt_TextChanged(sender As Object, e As TextChangedEventArgs) Handles SendSerialTxt.TextChanged

    End Sub

    Private Sub SerialCmdCombo_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles SerialCmdCombo.SelectionChanged
        SendSerialTxt.Clear()
        If SerialCmdCombo.SelectedItem = Nothing Then
            Return
        End If

        Dim SerialCommand As String = SerCmdAry(SerialCmdCombo.SelectedIndex)
        If SerCmdAry(SerialCmdCombo.SelectedIndex) = "setmotorpoles" Then

            CmdValueCombo.Items.Clear()
            'For i = 2 To 40 Step 2 ' Motor Poles
            For i = 2 To 10 Step 1 ' Prop Blades

                CmdValueCombo.Items.Add(CStr(i))
            Next
            CmdValueCombo.SelectedIndex = 0 '14 Poles

        ElseIf SerCmdAry(SerialCmdCombo.SelectedIndex) = "setmaxrpm" Then

            CmdValueCombo.Items.Clear()
            For i = 2000 To 20000 Step 100
                CmdValueCombo.Items.Add(CStr(i))
            Next
            CmdValueCombo.SelectedIndex = 40 '6000

        ElseIf SerCmdAry(SerialCmdCombo.SelectedIndex) = "setpwmoffset" Then

            CmdValueCombo.Items.Clear()
            For i = 950 To 1100 Step 10
                CmdValueCombo.Items.Add(CStr(i))
            Next
            CmdValueCombo.SelectedIndex = 5 '1000

        ElseIf SerCmdAry(SerialCmdCombo.SelectedIndex) = "setpwmmax" Then

            CmdValueCombo.Items.Clear()
            For i = 1800 To 2150 Step 10
                CmdValueCombo.Items.Add(CStr(i))
            Next
            CmdValueCombo.SelectedIndex = 15 '1950

        Else
            CmdValueCombo.Items.Clear()
            CmdValueCombo.Items.Add("--")
            CmdValueCombo.SelectedIndex = 0
        End If


    End Sub

    'CheckBox

    Private Sub ESC32CheckBox_Checked(sender As Object, e As RoutedEventArgs) Handles ESC32Checkbox.Click
        'Dim CheckFlag As Boolean
        If ESC32Checkbox.IsChecked Then
            If ESC32ThMode.IsChecked Then
                RPMSlider.Value = 20
            Else

                RPMSlider.Value = 1000
            End If

            ArduCheckbox.IsChecked = False
        Else
            ArduCheckbox.IsChecked = True

        End If
    End Sub

    Private Sub ArduCheckBox_Checked(sender As Object, e As RoutedEventArgs) Handles ArduCheckbox.Click
        'Dim CheckFlag As Boolean
        If ArduCheckbox.IsChecked Then
            RPMSlider.Value = 1000
            ESC32Checkbox.IsChecked = False
        Else
            ESC32Checkbox.IsChecked = True
        End If
    End Sub

    Private Sub ESC32ThMode_Checked(sender As Object, e As RoutedEventArgs) Handles ESC32ThMode.Checked
        Combo_MaxRpm.Items.Clear()
        Lbl_SetMaxRpm.Content = "Max Thrust"
        For i = 20 To 3000 Step 20
            Combo_MaxRpm.Items.Add(CStr(i))
        Next
        Combo_MaxRpm.SelectedIndex = 49 '1000
        RPMSlider.Value = 20


    End Sub

    Private Sub ESC32ThMode_unChecked(sender As Object, e As RoutedEventArgs) Handles ESC32ThMode.Unchecked
        Combo_MaxRpm.Items.Clear()

        Lbl_SetMaxRpm.Content = "Max RPM"
        For i = 2000 To 20000 Step 100
            Combo_MaxRpm.Items.Add(CStr(i))
        Next
        Combo_MaxRpm.SelectedIndex = 40 '6000

    End Sub



    ' Buttons

    Private Sub Rpm2Thrust_Click(sender As Object, e As RoutedEventArgs) Handles Rpm2Thrust.Click
        'DisplayElements(True)
        If Not CheckSanity() Then
            Return
        End If
        QuatSerText.Clear()

        'MyComPort.Write("quatos x1 y1 z7500;")
        MyComPort.Write("log rpm;")
        StopButton.IsDefault = True

    End Sub

    Private Sub MaxRpmBtn_Click(sender As Object, e As RoutedEventArgs) Handles MaxRpmBtn.Click
        If Not CheckSanity() Then
            Return
        End If
        QuatSerText.Clear()
        'MyComPort.Write("quatos x1 y1 z7500;")
        MyComPort.Write("runmaxrpmtest;")
    End Sub

    Private Sub Logdata_Click(sender As Object, e As RoutedEventArgs) Handles Logdata.Click
        If Not CheckSanity() Then
            Return
        End If
        QuatSerText.Clear()
        'MyComPort.Write("quatos x1 y1 z7500;")
        MyComPort.Write("start log;")
    End Sub

    Private Sub OLRpmThrust_Click(sender As Object, e As RoutedEventArgs) Handles OLRpmThrust.Click
        If Not CheckSanity() Then
            Return
        End If
        QuatSerText.Clear()
        'MyComPort.Write("quatos x1 y1 z7500;")
        MyComPort.Write("olsteprpm100Hz;")
    End Sub

    Private Sub SinRpm_Click(sender As Object, e As RoutedEventArgs) Handles SinRpm.Click
        If Not CheckSanity() Then
            Return
        End If
        QuatSerText.Clear()
        'MyComPort.Write("quatos x1 y1 z7500;")
        MyComPort.Write("log sinrpm;")
        'MyComPort.WriteLine("stop log;")
    End Sub

    Private Sub CLRpmThrust_Click(sender As Object, e As RoutedEventArgs) Handles CLRpmThrust.Click
        If Not CheckSanity() Then
            Return
        End If
        QuatSerText.Clear()
        'MyComPort.Write("quatos x1 y1 z7500;")
        MyComPort.Write("clsteprpm100Hz;")
    End Sub

    Private Sub CLRpmThrust1_Click(sender As Object, e As RoutedEventArgs) Handles CLRpmThrust1.Click
        If Not CheckSanity() Then
            Return
        End If
        QuatSerText.Clear()
        'MyComPort.Write("quatos x1 y1 z7500;")
        MyComPort.Write("clsteprpm1000Hz;")
    End Sub

    Private Sub StopButton_Click(sender As Object, e As RoutedEventArgs) Handles StopButton.Click
        MyComPort.DiscardInBuffer()
        MyComPort.Write("stop;") 'Stop ESC32
        MyComPort.Write("stop log;")
        SetLogRunView(True)


    End Sub

    'Sys

    Private Sub DisplayElements(SetVis As Boolean)


        Rpm2Thrust.IsEnabled = SetVis
        StopButton.IsEnabled = SetVis
        MaxRpmBtn.IsEnabled = SetVis
        OLRpmThrust.IsEnabled = SetVis
        SinRpm.IsEnabled = SetVis
        Logdata.IsEnabled = SetVis
        RpmVal.IsEnabled = SetVis
        RPMSlider.IsEnabled = SetVis
        ESCRun.IsEnabled = SetVis
        SendSerialBtn.IsEnabled = SetVis
        SendSerialTxt.IsEnabled = SetVis
        SerialCmdCombo.IsEnabled = SetVis
        CLRpmThrust.IsEnabled = SetVis
        CLRpmThrust1.IsEnabled = SetVis
        CLThrustStep1.IsEnabled = SetVis
        CmdValueCombo.IsEnabled = SetVis

        Btn_SetTestedUnit.IsEnabled = SetVis
        Btn_SetSystem.IsEnabled = SetVis

        Combo_MaxRpm.IsEnabled = SetVis
        Combo_MinPwm.IsEnabled = SetVis
        Combo_PropBlades.IsEnabled = SetVis
        Combo_MaxPwm.IsEnabled = SetVis
        Combo_SetRpmStep.IsEnabled = SetVis

        ESC32Checkbox.IsEnabled = SetVis
        ArduCheckbox.IsEnabled = SetVis

        RpmVal.IsEnabled = SetVis
        RPMSlider.IsEnabled = SetVis
        ESCRun.IsEnabled = SetVis
        ESC32ThMode.IsEnabled = SetVis
        If Not ESC32Checkbox.IsChecked Then
            ESC32ThMode.IsEnabled = False
        End If

    End Sub

    Private Sub SetMotRunView(SetVis As Boolean)

        DisplayElements(SetVis)

        RpmVal.IsEnabled = True
        RPMSlider.IsEnabled = True
        ESCRun.IsEnabled = True

        BtnConnect.IsEnabled = SetVis
        ComPort.IsEnabled = SetVis

    End Sub

    Private Sub SetLogRunView(SetVis As Boolean)

        DisplayElements(SetVis)

        StopButton.IsEnabled = True
        BtnConnect.IsEnabled = SetVis
        ComPort.IsEnabled = SetVis

    End Sub

    Private Function UpdateSysInfos(SysInfo As String)
        'Return False

        Dim SysArray() As String = {}
        Dim cString As String = ""
        Dim LastNonEmpty As Integer = -1
        Dim nLabel(13) As Label
        Dim isSys As Boolean = False


        If Not Mid(SysInfo, 1, 3) = "SYS" Then
            isSys = False
            If Mid(SysInfo, 1, 3) = "hBt" Then
                isSys = True
                If HeartBeat.Visibility = Windows.Visibility.Visible Then
                    HeartBeat.Visibility = Windows.Visibility.Hidden
                    If Not RunManual Then
                        If Mid(SysInfo, 5, 1) = "0" Then
                            SetLogRunView(True)
                        Else
                            SetLogRunView(False)
                        End If
                    End If
                Else
                    HeartBeat.Visibility = Windows.Visibility.Visible
                End If

                Return True
            Else
                Return False
            End If

        Else
            isSys = True
            SysInfo = Mid(SysInfo, 5)
        End If

        SysArray = Split(SysInfo, ";")

        nLabel(0) = Me.Voltage
        nLabel(1) = Me.Curr
        nLabel(2) = Me.ToL
        nLabel(3) = Me.ToR
        nLabel(4) = Me.Thrust
        nLabel(5) = Me.MotA1
        nLabel(6) = Me.MotA2
        nLabel(7) = Me.PropK1
        nLabel(8) = Me.MaxRpm
        nLabel(9) = Me.MotorPoles
        nLabel(10) = Me.MinPWM
        nLabel(11) = Me.MaxPWM
        nLabel(12) = Me.RPM


        'MsgBox(SysArray(0))
        For i As Integer = 0 To SysArray.Length - 1
            If SysArray(i) <> "" Then
                LastNonEmpty += 1
                SysArray(LastNonEmpty) = SysArray(i)
            End If
        Next
        If LastNonEmpty >= 0 Then

            ReDim Preserve SysArray(LastNonEmpty)
            For i As Integer = 0 To SysArray.Length - 1
                ' parse
                nLabel(i).Content = Trim(SysArray(i))
            Next
        End If

        If ESC32Checkbox.IsChecked Then
            If ESC32ThMode.IsChecked Then
                RPMSlider.Maximum = Val(Me.MaxRpm.Content)
                RPMSlider.Minimum = 20
                RPMSlider.LargeChange = 20
                MotRunLabel.Content = "Thrust"
            Else
                RPMSlider.Maximum = Val(Me.MaxRpm.Content)
                RPMSlider.Minimum = 1000
                RPMSlider.LargeChange = 1000
                MotRunLabel.Content = "Rpm"
            End If

        Else
            RPMSlider.Maximum = Val(Me.MaxPWM.Content)
            RPMSlider.Minimum = Val(Me.MinPWM.Content)
            RPMSlider.LargeChange = 100
            MotRunLabel.Content = "PWM"

        End If

        Return True
    End Function

    'Slider

    Private Sub RPMSlider_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles RPMSlider.ValueChanged
        Dim x As Integer = RPMSlider.Value
        Dim MotorCmd As String
        'If Me.RPMval.IsLoaded Then

        RpmVal.Text = x

        If ESC32Checkbox.IsChecked Then
            If ESC32ThMode.IsChecked Then
                MotorCmd = "setmotorthrust " + RpmVal.Text  'Thrust
            Else
                MotorCmd = "setmotorrpm " + RpmVal.Text  'RPM 
            End If
        Else
            MotorCmd = "setmotorpwm " + RpmVal.Text  'PWM
        End If
        If RunManual Then
            MyComPort.Write(MotorCmd + ";")
        End If
        'End If

    End Sub

    Private Sub ESCRun_Click(sender As Object, e As RoutedEventArgs) Handles ESCRun.Click
        Dim MotorStartCmd As String

        If Not CheckSanity() Then
            Return
        End If

        If ESC32Checkbox.IsChecked Then
            MotorStartCmd = "startesc32cl" 'RPM
        Else
            MotorStartCmd = "startesc32ol" 'PWM
        End If


        If ESCRun.Content = "Start Motor" Then
            ESCRun.Content = "Stop Motor"
            RunManual = True
        Else
            ESCRun.Content = "Start Motor"
            MotorStartCmd = "stop" 'STOP
            RunManual = False
        End If
        SetMotRunView(Not RunManual)
        MyComPort.Write(MotorStartCmd + ";")
        RPMSlider.Value += 0.1
        'RPMSlider.Value -= 1


    End Sub

    Private Sub RpmVal_TextChanged(sender As Object, e As TextChangedEventArgs) Handles RpmVal.TextChanged
        If RpmVal.Text() = "" Then
            Return
        End If
        Dim x As Double = CDbl(RpmVal.Text)
        If SysLoaded Then
            If x < 20 Then
                x = 20
            End If
            RPMSlider.Value = x
            'RPMSlider.UpdateLayout()
        End If

    End Sub

    'Divers

    Private Function ResizeText(ByVal cText)

        'Dim NewText As String = cText.substring
        'Dim SysArray() As String = cText.Split(CChar(vbCr & vbLf))
        'Dim NewArray(SysArray.Length) As String
        'SysArray.CopyTo(NewArray, 0)

        'For i As Integer = 0 To SysArray.Length - 2
        ' NewArray.SetValue(SysArray(I + 1), I)
        ' Next

        'cText = NewArray.ToString

        Dim TestPos As Integer = InStr(cText, CChar(vbCr & vbLf))
        Dim MidWords As String = Mid(cText, TestPos)
        Return Mid(MidWords, TestPos)

    End Function

    Private Function ResetTextBox()
        QuatSerText.Clear()
        TextLineCounter = 0
    End Function

    'Private Sub OnKeyDownHandler(ByVal sender As Object, ByVal e As KeyEventArgs)
    'Dim SerialCommand As String = SendSerialTxt.Text & ";" & vbCr & vbLf

    '  If (e.Key = Key.Return) Then

    '     MyComPort.Write(SerialCommand)

    ' End If
    ' End Sub


    Private Function CheckSanity()
        'If CUInt(Me.MaxRpm.Content) <= 2000 Then
        'MsgBox("Max RPM ist nicht angegeben")
        'Return False
        'End If
        'SetLogRunView(False)
        Return True
    End Function


    
    Private Sub ComboBox_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)

    End Sub

   
    Private Sub ClearText_Click(sender As Object, e As RoutedEventArgs) Handles ClearText.Click
        ResetTextBox()
    End Sub


    Private Sub Combo_PropBlades_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles Combo_PropBlades.SelectionChanged

    End Sub

   

    Private Sub Combo_MaxRpm_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles Combo_MaxRpm.SelectionChanged

    End Sub

   

    Private Sub Combo_MinPwm_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles Combo_MinPwm.SelectionChanged

    End Sub

   

    Private Sub Combo_MaxPwm_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles Combo_MaxPwm.SelectionChanged

    End Sub

    Private Sub Combo_SetRpmStep_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles Combo_SetRpmStep.SelectionChanged

    End Sub



    Private Sub Btn_SetSystem_Click(sender As Object, e As RoutedEventArgs) Handles Btn_SetSystem.Click
        Dim SerialCommand As String

        SerialCommand = "setmotorpoles " + Combo_PropBlades.SelectedItem + ";"
        MyComPort.Write(SerialCommand)

        SerialCommand = "setmaxrpm " + Combo_MaxRpm.SelectedItem + ";"
        MyComPort.Write(SerialCommand)

        SerialCommand = "setpwmoffset " + Combo_MinPwm.SelectedItem + ";"
        MyComPort.Write(SerialCommand)

        SerialCommand = "setpwmmax " + Combo_MaxPwm.SelectedItem + ";"
        MyComPort.Write(SerialCommand)

        SerialCommand = "setrpmstep " + Combo_SetRpmStep.SelectedItem + ";"
        MyComPort.Write(SerialCommand)

    End Sub

    Private Sub Btn_SetTestedUnit_Click(sender As Object, e As RoutedEventArgs) Handles Btn_SetTestedUnit.Click
        Dim SerialCommand As String

        SerialCommand = "settestedunit " + Strings.Left(Trim(TestUnitText.Text), 50) + ";"
        MyComPort.Write(SerialCommand)

    End Sub

    Private Sub CLThrustStep1_Click(sender As Object, e As RoutedEventArgs) Handles CLThrustStep1.Click
        If Not CheckSanity() Then
            Return
        End If
        QuatSerText.Clear()
        MyComPort.Write("clstepth100Hz;")
        StopButton.IsDefault = True

    End Sub

End Class





