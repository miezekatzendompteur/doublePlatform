'Option Strict On
Imports System, System.Drawing.Drawing2D, System.Drawing.Text, System.IO, System.Math, System.Runtime.InteropServices, System.Management, System.Globalization



Public Class Form1
    Inherits System.Windows.Forms.Form

    Dim TimeTraces As Graphics = Form2.CreateGraphics()
    Dim elapsed_ms, SamplingFrequency, VoltageSum As Double
    Dim Lj_ID, LjDemo As Integer
    Dim HeatedSide(2, 9), ExpType(2, 9) As String
    Dim HeatBtnTxt, myBtnTxt, status, Heat_P1, Heat_P2, Heat_P3, ExpTyp_P1, ExpTyp_P2, ExpTyp_P3 As String
    Dim file_header(,), filename As String
    Dim BlockDuration(9), n_t(9) As Single
    Dim PI(2, 9) As Single
    Dim th1, tc1, th2, tc2, th3, tc3 As Single
    Dim BlockNo, nOfBlocks, nOfLoops, Block_No, TMYindex As Integer
    Dim TotalTime, hysteresis, simul_loop, RunningTime, SamplFreq As Single
    Dim pos1, pos2, pos3, Y_pos1_old, Y_pos1, Y_pos2_old, Y_pos2, Y_pos3_old, Y_pos3 As Single
    Dim x, Xo, Yo As Integer
    Dim NofData, performed_NofData, endOfBlock, n As Long
    'definition of offset
    Dim db_offset_platform1, db_offset_platform2, db_offset_platform3 As Double
    Public str_comment As String
    ' Dim startCount, stopCount, elapsedCount, TimerFrequency As Int64
    '************************ Laser *************************************************
    Dim Laser1_ON As Boolean = True
    Dim Laser2_ON As Boolean = True
    Dim Laser1_OFF As Boolean = False
    Dim Laser2_OFF As Boolean = False
    Dim Laser1_OnOff, Laser2_OnOff, L1_old, L2_old As Boolean
    '************************ Laser *************************************************
    'definition plattform LED color
    Dim ledColorRed As Boolean = True

    'definition plattform LEDs
    Dim platform1_led1_onoff, platform1_led2_onoff, platform2_led1_onoff, platform2_led2_onoff, platform3_led1_onoff, platform3_led2_onoff As Boolean
    Dim platform1_led1_onoff_old, platform1_led2_onoff_old, platform2_led1_onoff_old, platform2_led2_onoff_old, platform3_led1_onoff_old, platform3_led2_onoff_old As Boolean

    'definition of LEDs

    Const platform1_led1 = 0
    Const platform1_led2 = 1
    Const platform2_led1 = 2
    Const platform2_led2 = 3
    Const platform3_led1 = 4
    Const platform3_led2 = 5

    Dim tab_ As Char = Chr(9)

    Dim white_brush As New SolidBrush(Color.White)
    Dim black_brush As New SolidBrush(Color.Black)
    Dim green_brush As New SolidBrush(Color.FromArgb(255, 0, 175, 0))
    Dim blue_brush As New SolidBrush(Color.FromArgb(255, 0, 0, 174))
    Dim red_brush As New SolidBrush(Color.FromArgb(255, 255, 0, 0))
    Dim bkgnd_brush As New SolidBrush(Color.FromArgb(255, 64, 64, 64))
    Dim whitePen As New Pen(Color.White)
    Dim dark_greenPen As New Pen(Color.DarkGreen)
    Dim greenPen As New Pen(Color.Green)
    Dim bluePen As New Pen(Color.Blue)
    Dim redPen As New Pen(Color.Red)
    Dim blackPen As New Pen(Color.Black)
    Dim dark_redPen As New Pen(Color.DarkRed)
    Dim pos_data As Single(,)
    Dim DataFile As StreamWriter

    Dim StopWatch As New Diagnostics.Stopwatch

    Dim serialNumbers(127) As Integer
    Dim productIDList(127) As Integer
    Dim localIDList(127) As Integer
    Dim powerList(127) As Integer
    Dim calMatrix(20, 127) As Integer
    Dim numberFound As Integer
    Dim reserved1 As Integer
    Dim reserved2 As Integer

    '{----------------------------------------------------}
    'Protected Overrides Sub WndProc(ByRef m As Message)
    '    MyBase.WndProc(m)
    '    If (m.Msg = &H219) And (m.WParam = &H7) Then
    '        Call findDevice()
    '    End If
    'End Sub

    Private Sub findDevice()

        Dim query As New SelectQuery("Win32_PnPEntity")
        Dim queryCollection As ManagementObjectCollection
        Dim searcher As New ManagementObjectSearcher()

        query.Condition = "DeviceID LIKE '%VID_0CD5%'"
        searcher.Query = query
        queryCollection = searcher.Get()

        Try
            If (queryCollection.Count > 0) Then
                'Console.WriteLine("Device found")

            Else
                'Console.WriteLine("No Device")
                Throw New System.Exception("No Labjack found")
            End If
        Catch ex As Exception
            Dim Result As Integer

            Result = MessageBox.Show(ex.Message, "Error", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error)
            If Result = System.Windows.Forms.DialogResult.Cancel Then
                'runButton.Visible = False
                'adjustButton.Visible = False
                Me.Close()
            Else
                findDevice()
            End If
        End Try

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Location = New Point(25, 400)
        Form2.StartPosition = FormStartPosition.Manual
        Form2.Location = New Point(25, 65)
        Form2.Show()
        Me.Text = My.Application.Info.AssemblyName.ToString & " " & My.Application.Info.Version.ToString

        Call InitProgram()
        Call findDevice()

    End Sub

    Sub usbDevicesPresent()
        Dim errorCode As Integer
        Dim Result As DialogResult

        Try
            errorCode = lj.LabJack.ListAll(productIDList, serialNumbers, localIDList, powerList, calMatrix, numberFound, reserved1, reserved2)

            If (numberFound = 0) Then
                Throw New System.Exception("No Labjack found")
            Else
            End If

        Catch ex As Exception
            Result = MessageBox.Show(ex.Message, "Error", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error)
            If Result = System.Windows.Forms.DialogResult.Cancel Then

                'runButton.Visible = False
                'adjustButton.Visible = False
                Me.Close()
            Else
                usbDevicesPresent()
            End If

        End Try

    End Sub
    '{----------------------------------------------------}
    Sub InitProgram()

        status = "stopped"
        platform1_led1_onoff = False
        platform1_led2_onoff = False
        platform2_led1_onoff = False
        platform2_led2_onoff = False
        platform3_led1_onoff = False
        platform3_led2_onoff = False
        'Call SwitchLaser1(Laser1_OFF) #JS
        'Call SwitchLaser2(Laser2_OFF) #JS
        platform1_led1_onoff_old = platform1_led1_onoff
        platform1_led2_onoff_old = platform1_led2_onoff
        platform2_led1_onoff_old = platform2_led1_onoff
        platform2_led2_onoff_old = platform2_led2_onoff
        platform3_led1_onoff_old = platform3_led1_onoff
        platform3_led2_onoff_old = platform3_led2_onoff

        'set offset to zero
        db_offset_platform1 = 0
        db_offset_platform2 = 0
        db_offset_platform3 = 0

        nOfBlocks = 10

        Call ReadConfiguration()
        '        status = "running"
        RunningTime = 0
        simul_loop = 0
        n = -1
        nOfLoops = 0
        Block_No = 0
        performed_NofData = 0

        SamplFreq = 33.333         ' 100Hz

        '******************************* set the sample frequency here & -> change timer1.Interval

        'SamplFreq = 10
        NofData = TotalTime * SamplFreq     ' (= 21.600 for 9 x 120s)
        ReDim pos_data(NofData, 3)   ' allocate NofData x 4 elements
        endOfBlock = n_t(0)

        RunningTime = 0

        'TimerFrequency = HiResTimer.QueryPerformanceFrequency()

        Call DrawFrame()
    End Sub

    '{---------------START ("Run"-Button)-------------------------------------}
    Private Sub runButton_Click(sender As Object, e As EventArgs) Handles runButton.Click
        status = "running"
        ledRedOnOff.Enabled = False
        runButton.BackgroundImage = My.Resources.runningButton
        runButton.Text = "running"
        Me.StopWatch.Reset()
        Call messen()
    End Sub

    '{---------------STOP-Button------------------------------------}
    Private Sub stopButton_Click(sender As Object, e As EventArgs) Handles stopButton.Click
        status = "stopped"
        ledRedOnOff.Enabled = True
        runButton.BackgroundImage = My.Resources.runButton
        runButton.Text = "run"

        platform1_led1_onoff = False
        platform1_led2_onoff = False
        platform2_led1_onoff = False
        platform2_led2_onoff = False
        platform3_led1_onoff = False
        platform3_led2_onoff = False

        Call SwitchLed(platform1_led1_onoff, platform1_led1)
        Call SwitchLed(platform1_led2_onoff, platform1_led2)
        Call SwitchLed(platform2_led1_onoff, platform2_led1)
        Call SwitchLed(platform2_led2_onoff, platform2_led2)
        Call SwitchLed(platform3_led1_onoff, platform3_led1)
        Call SwitchLed(platform3_led2_onoff, platform3_led2)

        Me.StopWatch.Stop()
    End Sub

    '{---------------Adjust Zero Positions-------------------------------------}
    Private Sub adjustButton_Click(sender As Object, e As EventArgs) Handles adjustButton.Click
        status = "running"
        ledRedOnOff.Enabled = False

        Call Adjust_ZeroPosition()
    End Sub

    '{---------------------RESET Button-------------------------------}
    Private Sub Button43_Click(sender As Object, e As EventArgs) Handles Button43.Click
        Call InitProgram()
        Call DrawFrame()
        Me.StopWatch.Reset()
        Label38.Text = "00:00"
    End Sub
    '{------------------------Hysteresis----------------------------}

    Private Sub NumericUpDown11_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown11.ValueChanged
        hysteresis = NumericUpDown11.Value
        Call DrawFrame()
    End Sub
    '{------------------------Offset Platform1----------------------------}

    Private Sub NumericUpDown12_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown12.ValueChanged
        db_offset_platform1 = NumericUpDown12.Value
    End Sub
    '{------------------------Offset Platform2----------------------------}

    Private Sub NumericUpDown13_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown13.ValueChanged
        db_offset_platform2 = NumericUpDown13.Value
    End Sub
    '{------------------------Offset Platform3----------------------------}

    Private Sub NumericUpDown14_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown14.ValueChanged
        db_offset_platform3 = NumericUpDown14.Value
    End Sub

    '{------------------------EXIT Button----------------------------}
    Private Sub Button47_Click(sender As Object, e As EventArgs) Handles Button47.Click
        'D2K_Release_Card(card_ID)
        TimeTraces.Dispose()
        Application.Exit()

    End Sub

    '{------------------------Save Button----------------------------}
    Private Sub Button46_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button46.Click
        Call write_datafile()
    End Sub
    Private Sub NumericUpDown1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown1.Click
        NumericUpDown1.Select(0, NumericUpDown1.Text.Length)
    End Sub
    Private Sub NumericUpDown2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown2.Click
        NumericUpDown2.Select(0, NumericUpDown2.Text.Length)
    End Sub
    Private Sub NumericUpDown3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown3.Click
        NumericUpDown3.Select(0, NumericUpDown3.Text.Length)
    End Sub
    Private Sub NumericUpDown4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown4.Click
        NumericUpDown4.Select(0, NumericUpDown4.Text.Length)
    End Sub
    Private Sub NumericUpDown5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown5.Click
        NumericUpDown5.Select(0, NumericUpDown5.Text.Length)
    End Sub
    Private Sub NumericUpDown6_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown6.Click
        NumericUpDown6.Select(0, NumericUpDown6.Text.Length)
    End Sub
    Private Sub NumericUpDown7_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown7.Click
        NumericUpDown7.Select(0, NumericUpDown7.Text.Length)
    End Sub
    Private Sub NumericUpDown8_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown8.Click
        NumericUpDown8.Select(0, NumericUpDown8.Text.Length)
    End Sub
    Private Sub NumericUpDown9_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown9.Click
        NumericUpDown9.Select(0, NumericUpDown9.Text.Length)
    End Sub
    Private Sub NumericUpDown10_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown10.Click
        NumericUpDown10.Select(0, NumericUpDown10.Text.Length)
    End Sub

    Private Sub NumericUpDown1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown1.Enter
        NumericUpDown1.Select(0, NumericUpDown1.Text.Length)
    End Sub
    Private Sub NumericUpDown2_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown2.Enter
        NumericUpDown2.Select(0, NumericUpDown2.Text.Length)
    End Sub
    Private Sub NumericUpDown3_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown3.Enter
        NumericUpDown3.Select(0, NumericUpDown3.Text.Length)
    End Sub
    Private Sub NumericUpDown4_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown4.Enter
        NumericUpDown4.Select(0, NumericUpDown4.Text.Length)
    End Sub
    Private Sub NumericUpDown5_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown5.Enter
        NumericUpDown5.Select(0, NumericUpDown5.Text.Length)
    End Sub
    Private Sub NumericUpDown6_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown6.Enter
        NumericUpDown6.Select(0, NumericUpDown6.Text.Length)
    End Sub
    Private Sub NumericUpDown7_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown7.Enter
        NumericUpDown7.Select(0, NumericUpDown7.Text.Length)
    End Sub
    Private Sub NumericUpDown8_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown8.Enter
        NumericUpDown8.Select(0, NumericUpDown8.Text.Length)
    End Sub
    Private Sub NumericUpDown9_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown9.Enter
        NumericUpDown9.Select(0, NumericUpDown9.Text.Length)
    End Sub
    Private Sub NumericUpDown10_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown10.Enter
        NumericUpDown10.Select(0, NumericUpDown10.Text.Length)
    End Sub


    '{----------------------------------------------------}   
    Public Sub DrawFrame()
        Dim Xtic, Ytic, Yo As Integer
        Dim arial As New Font("Arial", 10)
        Dim hyst_ As Single

        Dim string_format1 As New StringFormat
        string_format1.Alignment = StringAlignment.Center
        string_format1.LineAlignment = StringAlignment.Center

        hyst_ = hysteresis / 5 * 50
        blackPen.Width = 1
        For Yo = 60 To 280 Step 110
            dark_redPen.DashStyle = DashStyle.Dot
            TimeTraces.FillRectangle(white_brush, 40, Yo - 50, 750, 100)
            TimeTraces.DrawRectangle(blackPen, 40, Yo - 50, 750, 100)
            TimeTraces.DrawLine(blackPen, 40, Yo, 790, Yo)
            TimeTraces.DrawLine(dark_redPen, 40, Yo - hyst_, 790, Yo - hyst_)    'hier die Werte von +/- Hysteresis zeichnen
            TimeTraces.DrawLine(dark_redPen, 40, Yo + hyst_, 790, Yo + hyst_)

            For Ytic = Yo - 50 To Yo + 50 Step 10
                TimeTraces.DrawLine(blackPen, 40, Ytic, 43, Ytic)
                Select Case Ytic
                    Case 10, 120, 230 : TimeTraces.DrawString("5", arial, black_brush, 30, Ytic, string_format1)
                    Case 60, 170, 280 : TimeTraces.DrawString("0", arial, black_brush, 30, Ytic, string_format1)
                    Case 110, 220, 330 : TimeTraces.DrawString("-5", arial, black_brush, 30, Ytic, string_format1)
                End Select
            Next

            For Xtic = 40 To 790 Step 20
                TimeTraces.DrawLine(blackPen, Xtic, Yo + 47, Xtic, Yo + 50)
            Next
        Next ' next Yo
        TimeTraces.RotateTransform(-90)
        TimeTraces.DrawString("platform position [arb.units]", arial, black_brush, -110, 10, string_format1)
        TimeTraces.RotateTransform(90)

        Xo = 40 + 1
        x = Xo


    End Sub
    '{----------------------------------------------------}   

    Public Sub DrawPositionTraces(pos1 As Single, pos2 As Single, pos3 As Single)

        x = x + 2
        If x >= 791 Then
            Call DrawFrame()
        End If
        Y_pos1 = 60 - pos1 / 5 * 50
        Y_pos2 = 170 - pos2 / 5 * 50
        Y_pos3 = 280 - pos3 / 5 * 50

        TimeTraces.DrawLine(redPen, x - 1, Y_pos1_old, x, Y_pos1)
        TimeTraces.DrawLine(bluePen, x - 1, Y_pos2_old, x, Y_pos2)
        TimeTraces.DrawLine(greenPen, x - 1, Y_pos3_old, x, Y_pos3)
        Y_pos1_old = Y_pos1
        Y_pos2_old = Y_pos2
        Y_pos3_old = Y_pos3

        If x >= 100 Then
            x = x

        End If

    End Sub


    '{----------------------------------------------------}   
    Sub Adjust_ZeroPosition()

        Call InitProgram()


        status = "running"
        'timer2.Enabled = True
        Timer2.Interval = 30

        Do Until (status = "stopped")
            stopButton.Select()    ' STOP-Button
            Application.DoEvents()
            Timer2.Start()
        Loop

        Timer2.Stop()
        ledRedOnOff.Enabled = True
        'timer2.Enabled = False
    End Sub

    '{----------------------------------------------------}

    Private Sub timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Call Get_PositionValues()
        Call DrawPositionTraces(pos1, pos2, pos3)
        'Call Get_HeaterValues() '#JS Stromüberwachung Laser

    End Sub

    Private Sub emergyCloseTimer()

        Try
            runButton.Visible = False
            adjustButton.Visible = False

            Timer1.Stop()
            Timer2.Stop()
            Me.StopWatch.Stop()
            status = "stopped"
        Catch ex As Exception

        End Try

    End Sub


    '{-------------------MESSEN!------------------------}   
    Sub messen()

        Dim CurrentApp As Process

        Call InitProgram()
        th1 = 0
        tc1 = 0
        th2 = 0
        tc2 = 0
        th3 = 0
        tc3 = 0
        status = "running"
        n = 0

        'set priority to HIGH
        CurrentApp = Process.GetCurrentProcess
        CurrentApp.PriorityClass = ProcessPriorityClass.High

        Timer1.Interval = 30
        'Timer1.Interval = 100
        'Timer1.Enabled = True
        Timer1.Start()
        Me.StopWatch.Start()

        Do Until (status = "stopped")

            stopButton.Select()    '  STOP-Button
            Application.DoEvents()

            If n >= NofData Then
                Timer1.Stop()
                Timer1.Enabled = False
                status = "stopped"

                platform1_led1_onoff = False
                platform1_led2_onoff = False
                platform2_led1_onoff = False
                platform2_led2_onoff = False
                platform3_led1_onoff = False
                platform3_led2_onoff = False

                Call SwitchLed(platform1_led1_onoff, platform1_led1)
                Call SwitchLed(platform1_led2_onoff, platform1_led2)
                Call SwitchLed(platform2_led1_onoff, platform2_led1)
                Call SwitchLed(platform2_led2_onoff, platform2_led2)
                Call SwitchLed(platform3_led1_onoff, platform3_led1)
                Call SwitchLed(platform3_led2_onoff, platform3_led2)

                Call Next_Block()   ' this is only to write the last PI to the label on Form1
                Form3.ShowDialog()
                Call write_datafile()
            End If

            ' **********************************************************
            ' * This loop is interrupted by timer1_Tick every 47 ms !! *
            ' **********************************************************

        Loop

        'set priority to Normal
        CurrentApp.PriorityClass = ProcessPriorityClass.Normal

        Timer1.Stop()
        Me.StopWatch.Stop()

        ledRedOnOff.Enabled = True
        runButton.BackgroundImage = My.Resources.runButton
        runButton.Text = "run"

        'Timer1.Enabled = False

    End Sub

    '{----------------------------------------------------}
    Private Sub timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim elapsed As TimeSpan = Me.StopWatch.Elapsed

        Call Get_PositionValues()
        'Call Get_HeaterValues() #JS Stromüberwachung Laser
        Call LaserManagement(pos1, pos2, pos3)
        Call Next_Block()
        n = n + 1
        performed_NofData = n
        If n <= NofData Then
            pos_data(n, 1) = pos1
            pos_data(n, 2) = pos2
            pos_data(n, 3) = pos3
            Call DrawPositionTraces(pos1, pos2, pos3)
            Label38.Text = String.Format("{0:00}:{1:00}", elapsed.Minutes, elapsed.Seconds)
            RunningTime = elapsed.Minutes * 60000 + elapsed.Seconds * 1000 + elapsed.Milliseconds
            pos_data(n, 0) = RunningTime
        End If

    End Sub

    '{----------------------------------------------------}
    Sub Get_PositionValues()
        Dim lngOverVoltage, int_cali, int_chan(4), int_gain(4), int_over As Integer
        Dim db_voltage(4) As Single
        Dim db_ad_pos1, db_ad_pos2, db_ad_pos3 As Double
        Dim errorCode As Integer

        lngOverVoltage = 0
        int_cali = 1
        Lj_ID = -1
        LjDemo = 0 '############################################# DEMO #####################################################

        int_chan(0) = 0
        int_chan(1) = 2
        int_chan(2) = 4
        int_chan(3) = 6

        int_gain(0) = 0
        int_gain(1) = 0
        int_gain(2) = 0
        int_gain(3) = 0

        Try
            errorCode = lj.LabJack.AISample(Lj_ID, LjDemo, 0, 0, 1, 4, int_chan, int_gain, int_cali, int_over, db_voltage)

            If (errorCode > 0) Then
                Throw New System.Exception("Labjack Errorcode: " & errorCode)
            Else
            End If

        Catch ex As Exception
            emergyCloseTimer()

            MessageBox.Show(ex.Message, "Labjack Errorcode: " & errorCode, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try


        'lj.LabJack.EAnalogIn(Lj_ID, LjDemo, 0, 0, lngOverVoltage, pos1) 'standard EAnalogIn function call
        'lj.LabJack.EAnalogIn(Lj_ID, LjDemo, 2, 0, lngOverVoltage, pos2) 'standard EAnalogIn function call
        'lj.LabJack.EAnalogIn(Lj_ID, LjDemo, 4, 0, lngOverVoltage, pos3) 'standard EAnalogIn function call

        'adjust offset
        'db_ad_pos1 = Form2.TrackBar1.Value - 5
        'db_ad_pos2 = -Form2.TrackBar1.Value + 5

        db_ad_pos1 = db_voltage(0)
        db_ad_pos2 = db_voltage(1)
        db_ad_pos3 = db_voltage(2)

        pos1 = db_ad_pos1 + db_offset_platform1
        pos2 = db_ad_pos2 + db_offset_platform2
        pos3 = db_ad_pos3 + db_offset_platform3

        'Write values in textbox of form2
        Form2.TextBox1.Text = String.Format("{0,-6:#0.000}", pos1)
        Form2.TextBox2.Text = String.Format("{0,-6:#0.000}", pos2)
        Form2.TextBox3.Text = String.Format("{0,-6:#0.000}", pos3)

        'Write raw values in textbox of form2
        Form2.TextBox4.Text = String.Format("{0,-6:#0.000}", db_ad_pos1)
        Form2.TextBox5.Text = String.Format("{0,-6:#0.000}", db_ad_pos2)
        Form2.TextBox6.Text = String.Format("{0,-6:#0.000}", db_ad_pos3)

        'MsgBox(pos1)

    End Sub
    '{----------------------------------------------------}
    Sub Get_HeaterValues()
        Dim HeatVoltage As Integer
        Dim U1, U2, mA1, mA2 As Single
        Dim errorCode As Integer
        Lj_ID = -1
        LjDemo = 0

        Try
            errorCode = lj.LabJack.EAnalogIn(Lj_ID, LjDemo, 2, 0, HeatVoltage, U1) 'standard EAnalogIn function call

            If (errorCode > 0) Then
                Throw New System.Exception("Labjack Errorcode: " & errorCode)
            Else
            End If

            errorCode = lj.LabJack.EAnalogIn(Lj_ID, LjDemo, 3, 0, HeatVoltage, U2) 'standard EAnalogIn function call

            If (errorCode > 0) Then
                Throw New System.Exception("Labjack Errorcode: " & errorCode)
            Else
            End If

        Catch ex As Exception
            emergyCloseTimer()

            MessageBox.Show(ex.Message, "Labjack Errorcode: " & errorCode, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try


        mA1 = U1 / 19.5 * 1000  ' I1 = U1/R1 [in mA];   R1 = 19.5 Ohm
        mA2 = U2 / 19.5 * 1000  ' I2 = U2/R2 [in mA];   R2 = 19.5 Ohm
        'TextBox1.Text = Int(mA1)
        'TextBox2.Text = Int(mA2)
 
    End Sub

    '{----------------------------------------------------}
    Sub LaserManagement(ByVal int_pos1 As Single, ByVal int_pos2 As Single, ByVal int_pos3 As Single)

        'Get the values from the buttons-array

        Heat_P1 = HeatedSide(0, Block_No)
        Heat_P2 = HeatedSide(1, Block_No)
        Heat_P3 = HeatedSide(2, Block_No)
        ExpTyp_P1 = ExpType(0, Block_No)
        ExpTyp_P2 = ExpType(1, Block_No)
        ExpTyp_P3 = ExpType(2, Block_No)

        'Logic for Yoked and Master/ Values are needed for the calcualtion of th and tc

        If (ExpTyp_P1 = "Y" And ExpTyp_P2 = "M") Then
            Heat_P1 = Heat_P2
            int_pos1 = pos2
        Else
        End If

        If (ExpTyp_P1 = "Y" And ExpTyp_P3 = "M") Then
            Heat_P1 = Heat_P3
            int_pos1 = pos3
        Else
        End If

        If (ExpTyp_P2 = "Y" And ExpTyp_P1 = "M") Then
            Heat_P2 = Heat_P1
            int_pos2 = pos1
        Else
        End If

        If (ExpTyp_P2 = "Y" And ExpTyp_P3 = "M") Then
            Heat_P2 = Heat_P3
            int_pos2 = pos3
        Else
        End If

        If (ExpTyp_P3 = "Y" And ExpTyp_P1 = "M") Then
            Heat_P3 = Heat_P1
            int_pos3 = pos1
        Else
        End If

        If (ExpTyp_P3 = "Y" And ExpTyp_P2 = "M") Then
            Heat_P3 = Heat_P2
            int_pos3 = pos2
        Else
        End If

        'Platform LED switching
        '***************************** Add cases red/ yellow led ****************************************************

        Select Case Heat_P1
            Case "right"
                If int_pos1 >= hysteresis And ledColorRed = True Then
                    platform1_led1_onoff = True
                    platform1_led2_onoff = False
                Else
                End If
                If int_pos1 < -hysteresis And ledColorRed = True Then
                    platform1_led1_onoff = False
                    platform1_led2_onoff = False
                Else
                End If
                If int_pos1 >= hysteresis And ledColorRed = False Then
                    platform1_led1_onoff = False
                    platform1_led2_onoff = True
                Else
                End If
                If int_pos1 < -hysteresis And ledColorRed = False Then
                    platform1_led1_onoff = False
                    platform1_led2_onoff = False
                Else
                End If
            Case "left"
                If int_pos1 >= hysteresis And ledColorRed = True Then
                    platform1_led1_onoff = False
                    platform1_led2_onoff = False
                Else
                End If
                If int_pos1 < -hysteresis And ledColorRed = True Then
                    platform1_led1_onoff = True
                    platform1_led2_onoff = False
                Else
                End If
                If int_pos1 >= hysteresis And ledColorRed = False Then
                    platform1_led1_onoff = False
                    platform1_led2_onoff = False
                Else
                End If
                If int_pos1 < -hysteresis And ledColorRed = False Then
                    platform1_led1_onoff = False
                    platform1_led2_onoff = True
                Else
                End If
        End Select

        Select Case Heat_P2
            Case "right"
                If int_pos2 >= hysteresis And ledColorRed = True Then
                    platform2_led1_onoff = True
                    platform2_led2_onoff = False
                Else
                End If
                If int_pos2 < -hysteresis And ledColorRed = True Then
                    platform2_led1_onoff = False
                    platform2_led2_onoff = False
                Else
                End If
                If int_pos2 >= hysteresis And ledColorRed = False Then
                    platform2_led1_onoff = False
                    platform2_led2_onoff = True
                Else
                End If
                If int_pos2 < -hysteresis And ledColorRed = False Then
                    platform2_led1_onoff = False
                    platform2_led2_onoff = False
                Else
                End If
            Case "left"
                If int_pos2 >= hysteresis And ledColorRed = True Then
                    platform2_led1_onoff = False
                    platform2_led2_onoff = False
                Else
                End If
                If int_pos2 < -hysteresis And ledColorRed = True Then
                    platform2_led1_onoff = True
                    platform2_led2_onoff = False
                Else
                End If
                If int_pos2 >= hysteresis And ledColorRed = False Then
                    platform2_led1_onoff = False
                    platform2_led2_onoff = False
                Else
                End If
                If int_pos2 < -hysteresis And ledColorRed = False Then
                    platform2_led1_onoff = False
                    platform2_led2_onoff = True
                Else
                End If
        End Select

        Select Case Heat_P3
            Case "right"
                If int_pos3 >= hysteresis And ledColorRed = True Then
                    platform3_led1_onoff = True
                    platform3_led2_onoff = False
                Else
                End If
                If int_pos3 < -hysteresis And ledColorRed = True Then
                    platform3_led1_onoff = False
                    platform3_led2_onoff = False
                Else
                End If
                If int_pos3 >= hysteresis And ledColorRed = False Then
                    platform3_led1_onoff = False
                    platform3_led2_onoff = True
                Else
                End If
                If int_pos3 < -hysteresis And ledColorRed = False Then
                    platform3_led1_onoff = False
                    platform3_led2_onoff = False
                Else
                End If
            Case "left"
                If int_pos3 >= hysteresis And ledColorRed = True Then
                    platform3_led1_onoff = False
                    platform3_led2_onoff = False
                Else
                End If
                If int_pos3 < -hysteresis And ledColorRed = True Then
                    platform3_led1_onoff = True
                    platform3_led2_onoff = False
                Else
                End If
                If int_pos3 >= hysteresis And ledColorRed = False Then
                    platform3_led1_onoff = False
                    platform3_led2_onoff = False
                Else
                End If
                If int_pos3 < -hysteresis And ledColorRed = False Then
                    platform3_led1_onoff = False
                    platform3_led2_onoff = True
                Else
                End If
        End Select

        'Disable management

        If ExpTyp_P1 = "Te" Then
            platform1_led1_onoff = False
            platform1_led2_onoff = False
        Else
        End If

        If ExpTyp_P2 = "Te" Then
            platform2_led1_onoff = False
            platform2_led2_onoff = False
        Else
        End If

        If ExpTyp_P3 = "Te" Then
            platform3_led1_onoff = False
            platform3_led2_onoff = False
        Else
        End If

        'values for the performance index th and tc

        If platform1_led1_onoff = True And ledColorRed = True Then
            th1 = th1 + 1
        Else
        End If

        If platform1_led1_onoff = False And ledColorRed = True Then
            tc1 = tc1 + 1
        Else
        End If

        'If platform1_led2_onoff = True And ledColorRed = False Then
        '    tc1 = tc1 + 1
        'Else
        'End If

        If platform1_led2_onoff = True And ledColorRed = False Then
            th1 = th1 + 1
        Else
        End If

        If platform1_led2_onoff = False And ledColorRed = False Then
            tc1 = tc1 + 1
        Else
        End If

        'If platform2_led2_onoff = True Then
        '    tc2 = tc2 + 1
        'Else
        'End If

        If platform2_led1_onoff = True And ledColorRed = True Then
            th2 = th2 + 1
        Else
        End If

        If platform2_led1_onoff = False And ledColorRed = True Then
            tc2 = tc2 + 1
        Else
        End If

        If platform2_led2_onoff = True And ledColorRed = False Then
            th2 = th2 + 1
        Else
        End If

        If platform2_led2_onoff = False And ledColorRed = False Then
            tc2 = tc2 + 1
        Else
        End If

        'If platform3_led2_onoff = True Then
        '    tc3 = tc3 + 1
        'Else
        'End If

        If platform3_led1_onoff = True And ledColorRed = True Then
            th3 = th3 + 1
        Else
        End If

        If platform3_led1_onoff = False And ledColorRed = True Then
            tc3 = tc3 + 1
        Else
        End If

        If platform3_led2_onoff = True And ledColorRed = False Then
            th3 = th3 + 1
        Else
        End If

        If platform3_led2_onoff = False And ledColorRed = False Then
            tc3 = tc3 + 1
        Else
        End If

        'Switching the LEDs
        If (platform1_led1_onoff <> platform1_led1_onoff_old) Then
            Call SwitchLed(platform1_led1_onoff, platform1_led1)
        Else
        End If

        If (platform1_led2_onoff <> platform1_led2_onoff_old) Then
            Call SwitchLed(platform1_led2_onoff, platform1_led2)
        Else
        End If

        If (platform2_led1_onoff <> platform2_led1_onoff_old) Then
            Call SwitchLed(platform2_led1_onoff, platform2_led1)
        Else
        End If

        If (platform2_led2_onoff <> platform2_led2_onoff_old) Then
            Call SwitchLed(platform2_led2_onoff, platform2_led2)
        Else
        End If

        If (platform3_led1_onoff <> platform3_led1_onoff_old) Then
            Call SwitchLed(platform3_led1_onoff, platform3_led1)
        Else
        End If

        If (platform3_led2_onoff <> platform3_led2_onoff_old) Then
            Call SwitchLed(platform3_led2_onoff, platform3_led2)
        Else
        End If

        platform1_led1_onoff_old = platform1_led1_onoff
        platform1_led2_onoff_old = platform1_led2_onoff
        platform2_led1_onoff_old = platform2_led1_onoff
        platform2_led2_onoff_old = platform2_led2_onoff
        platform3_led1_onoff_old = platform3_led1_onoff
        platform3_led2_onoff_old = platform3_led2_onoff

    End Sub
    '{----------------------------------------------------}

    Sub Next_Block()
        If n = endOfBlock Then

            Try
                PI(0, Block_No) = (tc1 - th1) / (th1 + tc1)
                PI(1, Block_No) = (tc2 - th2) / (th2 + tc2)
                PI(2, Block_No) = (tc3 - th3) / (th3 + tc3)
            Catch ex As DivideByZeroException
                Me.emergyCloseTimer()
                MsgBox("Divison by Zero. The program will stop")
            End Try


            Select Case Block_No
                Case 0
                    Label1.Text = Str(Math.Round(PI(0, 0), 2))
                    Label11.Text = Str(Math.Round(PI(1, 0), 2))
                    Label41.Text = Str(Math.Round(PI(2, 0), 2))
                Case 1
                    Label2.Text = Str(Math.Round(PI(0, 1), 2))
                    Label12.Text = Str(Math.Round(PI(1, 1), 2))
                    Label42.Text = Str(Math.Round(PI(2, 1), 2))
                Case 2
                    Label3.Text = Str(Math.Round(PI(0, 2), 2))
                    Label13.Text = Str(Math.Round(PI(1, 2), 2))
                    Label43.Text = Str(Math.Round(PI(2, 2), 2))
                Case 3
                    Label4.Text = Str(Math.Round(PI(0, 3), 2))
                    Label14.Text = Str(Math.Round(PI(1, 3), 2))
                    Label44.Text = Str(Math.Round(PI(2, 3), 2))
                Case 4
                    Label5.Text = Str(Math.Round(PI(0, 4), 2))
                    Label15.Text = Str(Math.Round(PI(1, 4), 2))
                    Label45.Text = Str(Math.Round(PI(2, 4), 2))
                Case 5
                    Label6.Text = Str(Math.Round(PI(0, 5), 2))
                    Label16.Text = Str(Math.Round(PI(1, 5), 2))
                    Label46.Text = Str(Math.Round(PI(2, 5), 2))
                Case 6
                    Label7.Text = Str(Math.Round(PI(0, 6), 2))
                    Label17.Text = Str(Math.Round(PI(1, 6), 2))
                    Label47.Text = Str(Math.Round(PI(2, 6), 2))
                Case 7
                    Label8.Text = Str(Math.Round(PI(0, 7), 2))
                    Label18.Text = Str(Math.Round(PI(1, 7), 2))
                    Label48.Text = Str(Math.Round(PI(2, 7), 2))
                Case 8
                    Label9.Text = Str(Math.Round(PI(0, 8), 2))
                    Label19.Text = Str(Math.Round(PI(1, 8), 2))
                    Label49.Text = Str(Math.Round(PI(2, 8), 2))
                Case 9
                    Label10.Text = Str(Math.Round(PI(0, 9), 2))
                    Label20.Text = Str(Math.Round(PI(1, 9), 2))
                    Label50.Text = Str(Math.Round(PI(2, 9), 2))
            End Select

            If Block_No < 9 Then
                Block_No = Block_No + 1
                endOfBlock = endOfBlock + n_t(Block_No)
                th1 = 0
                tc1 = 0
                th2 = 0
                tc2 = 0
                th3 = 0
                tc3 = 0
            End If

        End If

    End Sub
    '{----------------------------------------------------}

    Private Sub ReadConfiguration()

        Call ReadHeatManagement()
        Call ReadMasterYokedManagement()
        Call reset_PIs()
        Call Get_BlockDurations()
        hysteresis = NumericUpDown11.Value

    End Sub
    '{----------------------------------------------------}
    Sub ReadHeatManagement()
        HeatedSide(0, 0) = Button1.Text
        HeatedSide(0, 1) = Button2.Text
        HeatedSide(0, 2) = Button3.Text
        HeatedSide(0, 3) = Button4.Text
        HeatedSide(0, 4) = Button5.Text
        HeatedSide(0, 5) = Button6.Text
        HeatedSide(0, 6) = Button7.Text
        HeatedSide(0, 7) = Button8.Text
        HeatedSide(0, 8) = Button9.Text
        HeatedSide(0, 9) = Button10.Text
        HeatedSide(1, 0) = Button21.Text
        HeatedSide(1, 1) = Button22.Text
        HeatedSide(1, 2) = Button23.Text
        HeatedSide(1, 3) = Button24.Text
        HeatedSide(1, 4) = Button25.Text
        HeatedSide(1, 5) = Button26.Text
        HeatedSide(1, 6) = Button27.Text
        HeatedSide(1, 7) = Button28.Text
        HeatedSide(1, 8) = Button29.Text
        HeatedSide(1, 9) = Button30.Text
        HeatedSide(2, 0) = Button59.Text
        HeatedSide(2, 1) = Button60.Text
        HeatedSide(2, 2) = Button61.Text
        HeatedSide(2, 3) = Button62.Text
        HeatedSide(2, 4) = Button63.Text
        HeatedSide(2, 5) = Button64.Text
        HeatedSide(2, 6) = Button65.Text
        HeatedSide(2, 7) = Button66.Text
        HeatedSide(2, 8) = Button67.Text
        HeatedSide(2, 9) = Button68.Text
    End Sub
    '{----------------------------------------------------}
    Sub ReadMasterYokedManagement()
        ExpType(0, 0) = Button11.Text
        ExpType(0, 1) = Button12.Text
        ExpType(0, 2) = Button13.Text
        ExpType(0, 3) = Button14.Text
        ExpType(0, 4) = Button15.Text
        ExpType(0, 5) = Button16.Text
        ExpType(0, 6) = Button17.Text
        ExpType(0, 7) = Button18.Text
        ExpType(0, 8) = Button19.Text
        ExpType(0, 9) = Button20.Text
        ExpType(1, 0) = Button31.Text
        ExpType(1, 1) = Button32.Text
        ExpType(1, 2) = Button33.Text
        ExpType(1, 3) = Button34.Text
        ExpType(1, 4) = Button35.Text
        ExpType(1, 5) = Button36.Text
        ExpType(1, 6) = Button37.Text
        ExpType(1, 7) = Button38.Text
        ExpType(1, 8) = Button39.Text
        ExpType(1, 9) = Button40.Text
        ExpType(2, 0) = Button49.Text
        ExpType(2, 1) = Button50.Text
        ExpType(2, 2) = Button51.Text
        ExpType(2, 3) = Button52.Text
        ExpType(2, 4) = Button53.Text
        ExpType(2, 5) = Button54.Text
        ExpType(2, 6) = Button55.Text
        ExpType(2, 7) = Button56.Text
        ExpType(2, 8) = Button57.Text
        ExpType(2, 9) = Button58.Text
    End Sub
    '{----------------------------------------------------}
    Sub reset_PIs()
        Dim BlkNo As Integer

        Label1.Text = "0"
        Label2.Text = "0"
        Label3.Text = "0"
        Label4.Text = "0"
        Label5.Text = "0"
        Label6.Text = "0"
        Label7.Text = "0"
        Label8.Text = "0"
        Label9.Text = "0"
        Label10.Text = "0"
        Label11.Text = "0"
        Label12.Text = "0"
        Label13.Text = "0"
        Label14.Text = "0"
        Label15.Text = "0"
        Label16.Text = "0"
        Label17.Text = "0"
        Label18.Text = "0"
        Label19.Text = "0"
        Label20.Text = "0"
        Label41.Text = "0"
        Label42.Text = "0"
        Label43.Text = "0"
        Label44.Text = "0"
        Label45.Text = "0"
        Label46.Text = "0"
        Label47.Text = "0"
        Label48.Text = "0"
        Label49.Text = "0"
        Label50.Text = "0"

        For BlkNo = 1 To nOfBlocks
            PI(0, BlkNo - 1) = 0
            PI(1, BlkNo - 1) = 0
            PI(2, BlkNo - 1) = 0
        Next

    End Sub
    '{----------------------------------------------------}
    Sub Get_BlockDurations()
        Dim BlkNo As Integer

        TotalTime = 0
        BlockDuration(0) = NumericUpDown1.Value
        BlockDuration(1) = NumericUpDown2.Value
        BlockDuration(2) = NumericUpDown3.Value
        BlockDuration(3) = NumericUpDown4.Value
        BlockDuration(4) = NumericUpDown5.Value
        BlockDuration(5) = NumericUpDown6.Value
        BlockDuration(6) = NumericUpDown7.Value
        BlockDuration(7) = NumericUpDown8.Value
        BlockDuration(8) = NumericUpDown9.Value
        BlockDuration(9) = NumericUpDown10.Value

        For BlkNo = 1 To nOfBlocks
            n_t(BlkNo - 1) = BlockDuration(BlkNo - 1) * SamplFreq
            TotalTime = TotalTime + BlockDuration(BlkNo - 1)
        Next

    End Sub
    '{----------------------------------------------------}
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        HeatBtnTxt = Button1.Text
        myBtnTxt = Button11.Text
        Call SwitchHeatMngmnt()
        Button1.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        HeatBtnTxt = Button2.Text
        myBtnTxt = Button12.Text
        Call SwitchHeatMngmnt()
        Button2.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        HeatBtnTxt = Button3.Text
        myBtnTxt = Button13.Text
        Call SwitchHeatMngmnt()
        Button3.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        HeatBtnTxt = Button4.Text
        myBtnTxt = Button14.Text
        Call SwitchHeatMngmnt()
        Button4.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        HeatBtnTxt = Button5.Text
        myBtnTxt = Button15.Text
        Call SwitchHeatMngmnt()
        Button5.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        HeatBtnTxt = Button6.Text
        myBtnTxt = Button16.Text
        Call SwitchHeatMngmnt()
        Button6.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        HeatBtnTxt = Button7.Text
        myBtnTxt = Button17.Text
        Call SwitchHeatMngmnt()
        Button7.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        HeatBtnTxt = Button8.Text
        myBtnTxt = Button18.Text
        Call SwitchHeatMngmnt()
        Button8.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        HeatBtnTxt = Button9.Text
        myBtnTxt = Button19.Text
        Call SwitchHeatMngmnt()
        Button9.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        HeatBtnTxt = Button10.Text
        myBtnTxt = Button20.Text
        Call SwitchHeatMngmnt()
        Button10.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        HeatBtnTxt = Button21.Text
        myBtnTxt = Button31.Text
        Call SwitchHeatMngmnt()
        Button21.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        HeatBtnTxt = Button22.Text
        myBtnTxt = Button32.Text
        Call SwitchHeatMngmnt()
        Button22.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        HeatBtnTxt = Button23.Text
        myBtnTxt = Button33.Text
        Call SwitchHeatMngmnt()
        Button23.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        HeatBtnTxt = Button24.Text
        myBtnTxt = Button34.Text
        Call SwitchHeatMngmnt()
        Button24.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        HeatBtnTxt = Button25.Text
        myBtnTxt = Button35.Text
        Call SwitchHeatMngmnt()
        Button25.Text = HeatBtnTxt
    End Sub

    Private Sub Button73_Click(sender As Object, e As EventArgs)

    End Sub

    '{----------------------------------------------------}
    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        HeatBtnTxt = Button26.Text
        myBtnTxt = Button36.Text
        Call SwitchHeatMngmnt()
        Button26.Text = HeatBtnTxt
    End Sub

    Private Sub Button42_Click(sender As Object, e As EventArgs)

    End Sub
    '{----------------------------------------------------}
    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        HeatBtnTxt = Button27.Text
        myBtnTxt = Button37.Text
        Call SwitchHeatMngmnt()
        Button27.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click
        HeatBtnTxt = Button28.Text
        myBtnTxt = Button38.Text
        Call SwitchHeatMngmnt()
        Button28.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles Button29.Click
        HeatBtnTxt = Button29.Text
        myBtnTxt = Button39.Text
        Call SwitchHeatMngmnt()
        Button29.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click
        HeatBtnTxt = Button30.Text
        myBtnTxt = Button40.Text
        Call SwitchHeatMngmnt()
        Button30.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button59_Click(sender As Object, e As EventArgs) Handles Button59.Click
        HeatBtnTxt = Button59.Text
        myBtnTxt = Button49.Text
        Call SwitchHeatMngmnt()
        Button59.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button60_Click(sender As Object, e As EventArgs) Handles Button60.Click
        HeatBtnTxt = Button60.Text
        myBtnTxt = Button50.Text
        Call SwitchHeatMngmnt()
        Button60.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button61_Click(sender As Object, e As EventArgs) Handles Button61.Click
        HeatBtnTxt = Button61.Text
        myBtnTxt = Button51.Text
        Call SwitchHeatMngmnt()
        Button61.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button62_Click(sender As Object, e As EventArgs) Handles Button62.Click
        HeatBtnTxt = Button62.Text
        myBtnTxt = Button52.Text
        Call SwitchHeatMngmnt()
        Button62.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button63_Click(sender As Object, e As EventArgs) Handles Button63.Click
        HeatBtnTxt = Button63.Text
        myBtnTxt = Button53.Text
        Call SwitchHeatMngmnt()
        Button63.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button64_Click(sender As Object, e As EventArgs) Handles Button64.Click
        HeatBtnTxt = Button64.Text
        myBtnTxt = Button54.Text
        Call SwitchHeatMngmnt()
        Button64.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button65_Click(sender As Object, e As EventArgs) Handles Button65.Click
        HeatBtnTxt = Button65.Text
        myBtnTxt = Button55.Text
        Call SwitchHeatMngmnt()
        Button65.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button66_Click(sender As Object, e As EventArgs) Handles Button66.Click
        HeatBtnTxt = Button66.Text
        myBtnTxt = Button56.Text
        Call SwitchHeatMngmnt()
        Button66.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button67_Click(sender As Object, e As EventArgs) Handles Button67.Click
        HeatBtnTxt = Button67.Text
        myBtnTxt = Button57.Text
        Call SwitchHeatMngmnt()
        Button67.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button68_Click(sender As Object, e As EventArgs) Handles Button68.Click
        HeatBtnTxt = Button68.Text
        myBtnTxt = Button58.Text
        Call SwitchHeatMngmnt()
        Button68.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Sub SwitchHeatMngmnt()
        If myBtnTxt = "M" Or myBtnTxt = "Te" Or myBtnTxt = "M_S" Then
            If HeatBtnTxt = "right" Then
                HeatBtnTxt = "left"
            Else
                HeatBtnTxt = "right"
            End If
        End If
    End Sub
    '{----------------------------------------------------}
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button31.Text
        str_button2 = Button49.Text
        myBtnTxt = Button11.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button11.Text = myBtnTxt
        Button1.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}   
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button32.Text
        str_button2 = Button50.Text
        myBtnTxt = Button12.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button12.Text = myBtnTxt
        Button2.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button33.Text
        str_button2 = Button51.Text
        myBtnTxt = Button13.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button13.Text = myBtnTxt
        Button3.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button34.Text
        str_button2 = Button52.Text
        myBtnTxt = Button14.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button14.Text = myBtnTxt
        Button4.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button35.Text
        str_button2 = Button53.Text
        myBtnTxt = Button15.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button15.Text = myBtnTxt
        Button5.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button36.Text
        str_button2 = Button54.Text
        myBtnTxt = Button16.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button16.Text = myBtnTxt
        Button6.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button37.Text
        str_button2 = Button55.Text
        myBtnTxt = Button17.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button17.Text = myBtnTxt
        Button7.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button38.Text
        str_button2 = Button56.Text
        myBtnTxt = Button18.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button18.Text = myBtnTxt
        Button8.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button39.Text
        str_button2 = Button57.Text
        myBtnTxt = Button19.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button19.Text = myBtnTxt
        Button9.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button40.Text
        str_button2 = Button58.Text
        myBtnTxt = Button20.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button20.Text = myBtnTxt
        Button10.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button31_Click(sender As Object, e As EventArgs) Handles Button31.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button11.Text
        str_button2 = Button49.Text
        myBtnTxt = Button31.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button31.Text = myBtnTxt
        Button21.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button32_Click(sender As Object, e As EventArgs) Handles Button32.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button12.Text
        str_button2 = Button50.Text
        myBtnTxt = Button32.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button32.Text = myBtnTxt
        Button22.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button33_Click(sender As Object, e As EventArgs) Handles Button33.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button13.Text
        str_button2 = Button51.Text
        myBtnTxt = Button33.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button33.Text = myBtnTxt
        Button23.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button34_Click(sender As Object, e As EventArgs) Handles Button34.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button14.Text
        str_button2 = Button52.Text
        myBtnTxt = Button34.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button34.Text = myBtnTxt
        Button24.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button35_Click(sender As Object, e As EventArgs) Handles Button35.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button15.Text
        str_button2 = Button53.Text
        myBtnTxt = Button35.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button35.Text = myBtnTxt
        Button25.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button36_Click(sender As Object, e As EventArgs) Handles Button36.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button16.Text
        str_button2 = Button54.Text
        myBtnTxt = Button36.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button36.Text = myBtnTxt
        Button26.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button37_Click(sender As Object, e As EventArgs) Handles Button37.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button17.Text
        str_button2 = Button55.Text
        myBtnTxt = Button37.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button37.Text = myBtnTxt
        Button27.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button38_Click(sender As Object, e As EventArgs) Handles Button38.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button18.Text
        str_button2 = Button56.Text
        myBtnTxt = Button38.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button38.Text = myBtnTxt
        Button28.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button39_Click(sender As Object, e As EventArgs) Handles Button39.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button19.Text
        str_button2 = Button57.Text
        myBtnTxt = Button39.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button39.Text = myBtnTxt
        Button29.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button40_Click(sender As Object, e As EventArgs) Handles Button40.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button20.Text
        str_button2 = Button58.Text
        myBtnTxt = Button40.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button40.Text = myBtnTxt
        Button30.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button49_Click(sender As Object, e As EventArgs) Handles Button49.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button31.Text
        str_button2 = Button11.Text
        myBtnTxt = Button49.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button49.Text = myBtnTxt
        Button59.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button50_Click(sender As Object, e As EventArgs) Handles Button50.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button32.Text
        str_button2 = Button12.Text
        myBtnTxt = Button50.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button50.Text = myBtnTxt
        Button60.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button51_Click(sender As Object, e As EventArgs) Handles Button51.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button33.Text
        str_button2 = Button13.Text
        myBtnTxt = Button51.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button51.Text = myBtnTxt
        Button61.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button52_Click(sender As Object, e As EventArgs) Handles Button52.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button34.Text
        str_button2 = Button14.Text
        myBtnTxt = Button52.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button52.Text = myBtnTxt
        Button62.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button53_Click(sender As Object, e As EventArgs) Handles Button53.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button35.Text
        str_button2 = Button15.Text
        myBtnTxt = Button53.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button53.Text = myBtnTxt
        Button63.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button54_Click(sender As Object, e As EventArgs) Handles Button54.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button36.Text
        str_button2 = Button16.Text
        myBtnTxt = Button54.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button54.Text = myBtnTxt
        Button64.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button55_Click(sender As Object, e As EventArgs) Handles Button55.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button37.Text
        str_button2 = Button17.Text
        myBtnTxt = Button55.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button55.Text = myBtnTxt
        Button65.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button56_Click(sender As Object, e As EventArgs) Handles Button56.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button38.Text
        str_button2 = Button18.Text
        myBtnTxt = Button56.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button56.Text = myBtnTxt
        Button66.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button57_Click(sender As Object, e As EventArgs) Handles Button57.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button39.Text
        str_button2 = Button19.Text
        myBtnTxt = Button57.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button57.Text = myBtnTxt
        Button67.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub Button58_Click(sender As Object, e As EventArgs) Handles Button58.Click
        Dim str_button1, str_button2 As String
        str_button1 = Button40.Text
        str_button2 = Button20.Text
        myBtnTxt = Button58.Text
        Call Switch_TeMY(str_button1, str_button2)
        Button58.Text = myBtnTxt
        Button68.Text = HeatBtnTxt
    End Sub
    '{----------------------------------------------------}
    Private Sub ledRedOnOff_Click(sender As Object, e As EventArgs) Handles ledRedOnOff.Click
        If ledRedOnOff.Tag = "red" Then
            ledRedOnOff.BackgroundImage = My.Resources.leddiode_gelb
            ledRedOnOff.Tag = "yellow"
            ledColorRed = False
        Else
            ledRedOnOff.BackgroundImage = My.Resources.leddiode_rot
            ledRedOnOff.Tag = "red"
            ledColorRed = True
        End If
    End Sub
    '{----------------------------------------------------}
    Sub Switch_TeMY(ByVal str_button1 As String, ByVal str_button2 As String)
        Dim TeMY_index As Integer

        Select Case myBtnTxt
            Case "Te"
                TeMY_index = 1
            Case "M"
                TeMY_index = 2
            Case "M_S"
                TeMY_index = 2
            Case "Y"
                TeMY_index = 3
        End Select

        TeMY_index = TeMY_index + 1

        If TeMY_index > 3 Then TeMY_index = 1
        Select Case TeMY_index
            Case 1
                myBtnTxt = "Te"
                HeatBtnTxt = "right"
            Case 2
                If (str_button1 = "M") Or (str_button2 = "M") Then
                    myBtnTxt = "M_S"
                    HeatBtnTxt = "right"
                Else
                    myBtnTxt = "M"
                    HeatBtnTxt = "right"
                End If
                '    If myBtnTxt_2 <> "Te" Then
                ' myBtnTxt_2 = "Y"
                ' HeatBtnTxt_2 = "->M"
                ' End If
            Case 3
                myBtnTxt = "Y"
                HeatBtnTxt = "->M"
                '     If myBtnTxt_2 <> "Te" Then
                ' myBtnTxt_2 = "M"
                ' HeatBtnTxt_2 = "right"
                ' End If
        End Select

    End Sub
    '{-------------------Laser 1 Test button---------------------------------}
    Private Sub Button44_Click(sender As Object, e As EventArgs) Handles Button44.Click
        status = "stopped"
        If Laser1_OnOff = Laser1_OFF Then
            Laser1_OnOff = Laser1_ON
        Else
            Laser1_OnOff = Laser1_OFF
        End If
        Select Case Laser1_OnOff
            Case True
                Call SwitchLaser1(Laser1_ON)
            Case False
                Call SwitchLaser1(Laser1_OFF)
        End Select
        Do While status = "stopped"
            '   stopButton.Select()    ' STOP-Button
            Application.DoEvents()
            Call Get_HeaterValues()
        Loop

    End Sub
    '{-------------------Laser 2 Test button---------------------------------}
    Private Sub Button45_Click(sender As Object, e As EventArgs) Handles Button45.Click
        status = "stopped"
        If Laser2_OnOff = Laser2_OFF Then
            Laser2_OnOff = Laser2_ON
        Else
            Laser2_OnOff = Laser2_OFF
        End If
        Select Case Laser2_OnOff
            Case True
                Call SwitchLaser2(Laser2_ON)
            Case False
                Call SwitchLaser2(Laser2_OFF)
        End Select

        Do While status = "stopped"
            '     stopButton.Select()    ' STOP-Button
            Application.DoEvents()
            Call Get_HeaterValues()
        Loop


    End Sub
    Private Sub infoButton_Click(sender As Object, e As EventArgs) Handles infoButton.Click
        MsgBox("Contact: johann.schmid@ur.de" & vbCrLf & "Version: " &
               My.Application.Info.AssemblyName.ToString & " " & My.Application.Info.Version.ToString,
               vbInformation, My.Application.Info.AssemblyName.ToString & " " & My.Application.Info.Version.ToString)
    End Sub

    '{------------------------Switch Laser 1----------------------------}
    Sub SwitchLaser1(OnOff As Boolean)
        Dim IOchannel As Integer = 0
        Dim useIO_ports As Integer
        Dim errorCode As Integer
        Dim errorString(50) As Char

        Lj_ID = -1  '  serial number, or -1 for first found.
        LjDemo = 0  '  0 for normal operation, >0 for demo mode (=without LabJack)
        useIO_ports = 0  ' IF 0, THEN DIO0 ..DIO3; if <> 0, then D0 ..D20

        Try
            Select Case OnOff
                Case Laser1_ON
                    Button44.Image = My.Resources.LaserTest_on_S1
                    ErrorCode = lj.LabJack.EDigitalOut(Lj_ID, LjDemo, IOchannel, useIO_ports, 1)
                'lj.LabJack.EDigitalOut(-1, 0, 0, 5, 1)
                Case Laser1_OFF
                    If status = "running" Then
                        Button44.Image = My.Resources.LaserTest_off
                    Else   ' (status = "stopped"
                        Button44.Image = My.Resources.Laser_Warnschild_S
                    End If
                    ErrorCode = lj.LabJack.EDigitalOut(Lj_ID, LjDemo, IOchannel, useIO_ports, 0)
                    'lj.LabJack.EDigitalOut(-1, 0, 0, 0, 0)
            End Select

            If (errorCode > 0) Then
                Throw New System.Exception("Labjack Errorcode: " & errorCode)
            Else
            End If

        Catch ex As Exception
            emergyCloseTimer()

            MessageBox.Show(ex.Message, "Labjack Errorcode: " & ErrorCode, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try


    End Sub
    '{------------------------Switch Laser 2----------------------------}
    Sub SwitchLaser2(OnOff As Boolean)
        Dim IOchannel As Integer = 1
        Dim useIO_ports As Integer
        Dim errorCode As Integer = 99
        Dim errorString(50) As Char

        Lj_ID = -1
        LjDemo = 0
        useIO_ports = 0  ' IF 0, THEN DIO0 ..DIO3; if <> 0, then D0 ..D20

        Try
            Select Case OnOff
                Case Laser2_ON
                    Button45.Image = My.Resources.LaserTest_on_S1
                    errorCode = lj.LabJack.EDigitalOut(Lj_ID, LjDemo, IOchannel, useIO_ports, 1)
                Case Laser1_OFF
                    If status = "running" Then
                        Button45.Image = My.Resources.LaserTest_off
                    Else   ' (status = "stopped"
                        Button45.Image = My.Resources.Laser_Warnschild_S
                    End If
                    errorCode = lj.LabJack.EDigitalOut(Lj_ID, LjDemo, IOchannel, useIO_ports, 0)

            End Select

            If (ErrorCode > 0) Then
                Throw New System.Exception("Labjack Errorcode: " & errorCode)
            Else
            End If

        Catch ex As Exception
            emergyCloseTimer()
            MessageBox.Show(ex.Message, "Labjack Errorcode: " & errorCode, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub
    '{------------------------Switch LED ----------------------------}
    Sub SwitchLed(ByVal OnOff As Boolean, ByVal io_led As Integer)
        Dim IOchannel As Integer = 1
        Dim useIO_ports As Integer
        Dim errorCode As Integer
        Dim errorString(50) As Char

        Lj_ID = -1
        LjDemo = 0 'Demomode value > 0
        useIO_ports = 1  ' IF 0, THEN DIO0 ..DIO3; if <> 0, then D0 ..D20


        Try
            Select Case io_led
                Case platform1_led1
                    If (OnOff = True) Then
                        Button44.Image = My.Resources.greenLED
                        'errorCode = lj.LabJack.PulseOutStart(Lj_ID, LjDemo, 0, 1, 32767, 6, 1, 6, 1)
                        errorCode = lj.LabJack.EDigitalOut(Lj_ID, LjDemo, platform1_led1, useIO_ports, 1)
                    Else

                        Button44.Image = My.Resources.redLED
                        errorCode = lj.LabJack.EDigitalOut(Lj_ID, LjDemo, platform1_led1, useIO_ports, 0)
                    End If
                Case platform1_led2
                    If (OnOff = True) Then
                        Button72.Image = My.Resources.greenLED
                        errorCode = lj.LabJack.EDigitalOut(Lj_ID, LjDemo, platform1_led2, useIO_ports, 1)
                    Else

                        Button72.Image = My.Resources.redLED
                        errorCode = lj.LabJack.EDigitalOut(Lj_ID, LjDemo, platform1_led2, useIO_ports, 0)
                    End If
                Case platform2_led1
                    If (OnOff = True) Then
                        Button45.Image = My.Resources.greenLED
                        errorCode = lj.LabJack.EDigitalOut(Lj_ID, LjDemo, platform2_led1, useIO_ports, 1)
                    Else

                        Button45.Image = My.Resources.redLED
                        errorCode = lj.LabJack.EDigitalOut(Lj_ID, LjDemo, platform2_led1, useIO_ports, 0)
                    End If
                Case platform2_led2
                    If (OnOff = True) Then
                        Button71.Image = My.Resources.greenLED
                        errorCode = lj.LabJack.EDigitalOut(Lj_ID, LjDemo, platform2_led2, useIO_ports, 1)
                    Else

                        Button71.Image = My.Resources.redLED
                        errorCode = lj.LabJack.EDigitalOut(Lj_ID, LjDemo, platform2_led2, useIO_ports, 0)
                    End If
                Case platform3_led1
                    If (OnOff = True) Then
                        Button69.Image = My.Resources.greenLED
                        errorCode = lj.LabJack.EDigitalOut(Lj_ID, LjDemo, platform3_led1, useIO_ports, 1)
                    Else

                        Button69.Image = My.Resources.redLED
                        errorCode = lj.LabJack.EDigitalOut(Lj_ID, LjDemo, platform3_led1, useIO_ports, 0)
                    End If
                Case platform3_led2
                    If (OnOff = True) Then
                        Button70.Image = My.Resources.greenLED
                        errorCode = lj.LabJack.EDigitalOut(Lj_ID, LjDemo, platform3_led2, useIO_ports, 1)
                    Else

                        Button70.Image = My.Resources.redLED
                        errorCode = lj.LabJack.EDigitalOut(Lj_ID, LjDemo, platform3_led2, useIO_ports, 0)
                    End If

            End Select

            If (errorCode > 0) Then
                Throw New System.Exception("Labjack Errorcode: " & errorCode)
            Else
            End If

        Catch ex As Exception
            emergyCloseTimer()
            MessageBox.Show(ex.Message, "Labjack Errorcode: " & errorCode, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    '{------------------------Write Data File----------------------------}
    Sub write_datafile()
        Dim saveFileDialog1 As New SaveFileDialog()
        With saveFileDialog1
            .FileName = Date.UtcNow.ToString("yyyy_MM_dd_H-mm-ss")
            .Filter = "LightGuide FlightSimulator DataFile (*.dat)|*.dat|all files (*.*)|*.*"
            .FilterIndex = 1
            .InitialDirectory = Environment.SpecialFolder.MyDocuments
            .RestoreDirectory = True
            .Title = "save data"
        End With

        If saveFileDialog1.ShowDialog() = DialogResult.OK Then
            Try
                DataFile = New StreamWriter(saveFileDialog1.FileName)
                Call Write_FileHeader()
                Call write_torque_pos_data()
                DataFile.Close()
                DataFile = Nothing
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Fileerror", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Else

        End If

    End Sub

    Sub Write_FileHeader()
        ReDim file_header(36, 1)

        file_header(0, 0) = "Learned-Helplesness-Experiment "
        file_header(0, 1) = "(TriplePlatform)  from  " & My.Computer.Clock.LocalTime
        file_header(1, 0) = ""
        file_header(1, 1) = ""
        file_header(2, 0) = "software version: "
        file_header(2, 1) = "DoublePlatform_1.3 04-Jun-2014"
        file_header(3, 0) = ""
        file_header(3, 1) = ""
        file_header(4, 0) = "Setup of Experiment: "
        file_header(4, 1) = ""
        file_header(5, 0) = ""
        file_header(5, 1) = ""
        file_header(6, 0) = "Platform 1:"
        file_header(6, 1) = ""
        file_header(7, 0) = "ExpType:" & tab_
        For Me.Block_No = 0 To 9
            file_header(7, 0) = file_header(7, 0) & ExpType(0, Block_No)
            If Block_No < 9 Then file_header(7, 0) = file_header(7, 0) & tab_
        Next
        file_header(7, 1) = ""
        file_header(8, 0) = ""
        file_header(8, 1) = ""
        file_header(9, 0) = "heated side:" & tab_
        For Me.Block_No = 0 To 9
            file_header(9, 0) = file_header(9, 0) & HeatedSide(0, Block_No)
            If Block_No < 9 Then file_header(9, 0) = file_header(9, 0) & tab_
        Next
        file_header(9, 1) = ""
        file_header(10, 0) = ""
        file_header(10, 1) = ""
        file_header(11, 0) = "PIs:" & tab_
        For Me.Block_No = 0 To 9
            'file_header(11, 0) = file_header(11, 0) & String.Format("{0:0.000}", PI(0, Block_No))
            file_header(11, 0) = file_header(11, 0) & Str(PI(0, Block_No))
            If Block_No < 9 Then file_header(11, 0) = file_header(11, 0) & tab_
        Next

        file_header(11, 1) = ""
        file_header(12, 0) = ""
        file_header(12, 1) = ""
        file_header(13, 0) = "Platform 2:"
        file_header(13, 1) = ""
        file_header(14, 0) = "ExpType:" & tab_
        For Me.Block_No = 0 To 9
            file_header(14, 0) = file_header(14, 0) & ExpType(1, Block_No)
            If Block_No < 9 Then file_header(14, 0) = file_header(14, 0) & tab_
        Next
        file_header(14, 1) = ""
        file_header(15, 0) = ""
        file_header(15, 1) = ""
        file_header(16, 0) = "heated side:" & tab_
        For Me.Block_No = 0 To 9
            file_header(16, 0) = file_header(16, 0) & HeatedSide(1, Block_No)
            If Block_No < 9 Then file_header(16, 0) = file_header(16, 0) & tab_
        Next
        file_header(16, 1) = ""
        file_header(17, 0) = ""
        file_header(17, 1) = ""
        file_header(18, 0) = "PIs:" & tab_
        For Me.Block_No = 0 To 9
            'file_header(18, 0) = file_header(18, 0) & String.Format("{0:0.000}", PI(1, Block_No))
            file_header(18, 0) = file_header(18, 0) & Str(PI(1, Block_No))
            If Block_No < 9 Then file_header(18, 0) = file_header(18, 0) & tab_
        Next

        file_header(18, 1) = ""
        file_header(19, 0) = ""
        file_header(19, 1) = ""
        file_header(20, 0) = "Platform 3:"
        file_header(20, 1) = ""
        file_header(21, 0) = "ExpType:" & tab_
        For Me.Block_No = 0 To 9
            file_header(21, 0) = file_header(21, 0) & ExpType(2, Block_No)
            If Block_No < 9 Then file_header(21, 0) = file_header(21, 0) & tab_
        Next
        file_header(21, 1) = ""
        file_header(22, 0) = ""
        file_header(22, 1) = ""
        file_header(23, 0) = "heated side:" & tab_
        For Me.Block_No = 0 To 9
            file_header(23, 0) = file_header(23, 0) & HeatedSide(2, Block_No)
            If Block_No < 9 Then file_header(23, 0) = file_header(23, 0) & tab_
        Next
        file_header(23, 1) = ""
        file_header(24, 0) = ""
        file_header(24, 1) = ""
        file_header(25, 0) = "PIs:" & tab_
        For Me.Block_No = 0 To 9
            'file_header(18, 0) = file_header(18, 0) & String.Format("{0:0.000}", PI(1, Block_No))
            file_header(25, 0) = file_header(25, 0) & Str(PI(2, Block_No))
            If Block_No < 9 Then file_header(25, 0) = file_header(25, 0) & tab_
        Next
        file_header(25, 1) = ""
        file_header(26, 0) = ""
        file_header(26, 1) = ""

        file_header(27, 0) = "duration [s]:" & tab_
        For Me.Block_No = 0 To 9
            'file_header(20, 0) = file_header(20, 0) & String.Format("{0:000.00}", BlockDuration(Block_No))
            file_header(27, 0) = file_header(27, 0) & Str(BlockDuration(Block_No))
            If Block_No < 9 Then file_header(27, 0) = file_header(27, 0) & tab_
        Next
        file_header(27, 1) = ""
        file_header(28, 0) = ""
        file_header(28, 1) = ""

        file_header(29, 0) = "total number of data points = " & tab_
        file_header(29, 1) = Str(performed_NofData)
        file_header(30, 0) = "sampling frequency [Hz] = " & tab_
        'file_header(23, 1) = String.Format("{0:00.0}", SamplFreq)
        file_header(30, 1) = Str(SamplFreq)
        file_header(31, 0) = "Hysteresis = " & tab_
        'file_header(24, 1) = String.Format("{0:0.00}", hysteresis)
        file_header(31, 1) = Str(hysteresis)
        file_header(32, 0) = ""
        file_header(32, 1) = ""
        file_header(33, 0) = "Comment: "
        file_header(33, 1) = str_comment
        file_header(34, 0) = "------------------------------------"
        file_header(34, 1) = "------------------------------------"
        file_header(35, 0) = ""
        file_header(35, 1) = ""
        file_header(36, 0) = "n" & tab_ & "t[s]" & tab_ & "pos1" & tab_ & "pos2" & tab_ & "pos3"
        file_header(36, 1) = ""

        For LineNo As Integer = 0 To 35
            DataFile.WriteLine(file_header(LineNo, 0) & tab_ & file_header(LineNo, 1))
        Next
    End Sub

    '{---------------------------------------------------------}
    Sub write_torque_pos_data()
        Dim int_counter As Integer
        Dim int_counter_end As Integer

        'no overflow of array
        If NofData > performed_NofData Then
            int_counter_end = performed_NofData
        Else
            int_counter_end = NofData
        End If

        For int_counter = 0 To int_counter_end 'changed form performed_NofData to NofData
            DataFile.WriteLine(Str(int_counter) & tab_ & _
                Str(pos_data(int_counter, 0)) & tab_ & _
                Str(pos_data(int_counter, 1)) & tab_ & _
                Str(pos_data(int_counter, 2)) & tab_ & _
                Str(pos_data(int_counter, 3)))
        Next
    End Sub
End Class
