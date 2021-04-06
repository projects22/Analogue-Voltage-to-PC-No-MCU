Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Friend WithEvents Line7 As Microsoft.VisualBasic.PowerPacks.LineShape
    Friend WithEvents Line6 As Microsoft.VisualBasic.PowerPacks.LineShape
    Friend WithEvents Line5 As Microsoft.VisualBasic.PowerPacks.LineShape
    Friend WithEvents Line4 As Microsoft.VisualBasic.PowerPacks.LineShape
    Friend WithEvents Line3 As Microsoft.VisualBasic.PowerPacks.LineShape
    Friend WithEvents Line2 As Microsoft.VisualBasic.PowerPacks.LineShape
    Friend WithEvents Line1 As Microsoft.VisualBasic.PowerPacks.LineShape
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents ShapeContainer2 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButton5 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton4 As System.Windows.Forms.RadioButton
    Friend WithEvents cmb1 As System.Windows.Forms.ComboBox


    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub


    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents sp1 As System.IO.Ports.SerialPort
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.sp1 = New System.IO.Ports.SerialPort(Me.components)
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Line7 = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.Line6 = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.Line5 = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.Line4 = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.Line3 = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.Line2 = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.Line1 = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.ShapeContainer2 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.RadioButton5 = New System.Windows.Forms.RadioButton()
        Me.RadioButton4 = New System.Windows.Forms.RadioButton()
        Me.cmb1 = New System.Windows.Forms.ComboBox()
        Me.GroupBox1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.RadioButton2)
        Me.GroupBox1.Controls.Add(Me.RadioButton1)
        Me.GroupBox1.Location = New System.Drawing.Point(15, 176)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(80, 72)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Conversion"
        '
        'RadioButton2
        '
        Me.RadioButton2.Location = New System.Drawing.Point(8, 40)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(64, 24)
        Me.RadioButton2.TabIndex = 1
        Me.RadioButton2.Text = "10 Bits"
        '
        'RadioButton1
        '
        Me.RadioButton1.Checked = True
        Me.RadioButton1.Location = New System.Drawing.Point(8, 16)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(66, 24)
        Me.RadioButton1.TabIndex = 0
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "12 Bits"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.ControlDarkDark
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(1, 100)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(299, 60)
        Me.Panel1.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Lime
        Me.Label1.Location = New System.Drawing.Point(80, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(128, 40)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "0.000"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'Line7
        '
        Me.Line7.Name = "Line7"
        Me.Line7.X1 = 249
        Me.Line7.X2 = 257
        Me.Line7.Y1 = 27
        Me.Line7.Y2 = 68
        '
        'Line6
        '
        Me.Line6.Name = "Line6"
        Me.Line6.X1 = 215
        Me.Line6.X2 = 223
        Me.Line6.Y1 = 23
        Me.Line6.Y2 = 66
        '
        'Line5
        '
        Me.Line5.Name = "Line5"
        Me.Line5.X1 = 187
        Me.Line5.X2 = 185
        Me.Line5.Y1 = 18
        Me.Line5.Y2 = 63
        '
        'Line4
        '
        Me.Line4.Name = "Line4"
        Me.Line4.X1 = 152
        Me.Line4.X2 = 150
        Me.Line4.Y1 = 26
        Me.Line4.Y2 = 72
        '
        'Line3
        '
        Me.Line3.Name = "Line3"
        Me.Line3.X1 = 121
        Me.Line3.X2 = 129
        Me.Line3.Y1 = 31
        Me.Line3.Y2 = 76
        '
        'Line2
        '
        Me.Line2.Name = "Line2"
        Me.Line2.X1 = 57
        Me.Line2.X2 = 65
        Me.Line2.Y1 = 30
        Me.Line2.Y2 = 78
        '
        'Line1
        '
        Me.Line1.Name = "Line1"
        Me.Line1.X1 = 30
        Me.Line1.X2 = 32
        Me.Line1.Y1 = 35
        Me.Line1.Y2 = 89
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.White
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.Panel1)
        Me.Panel2.Controls.Add(Me.ShapeContainer2)
        Me.Panel2.Location = New System.Drawing.Point(14, 8)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(300, 160)
        Me.Panel2.TabIndex = 7
        '
        'ShapeContainer2
        '
        Me.ShapeContainer2.Location = New System.Drawing.Point(0, 0)
        Me.ShapeContainer2.Margin = New System.Windows.Forms.Padding(0)
        Me.ShapeContainer2.Name = "ShapeContainer2"
        Me.ShapeContainer2.Shapes.AddRange(New Microsoft.VisualBasic.PowerPacks.Shape() {Me.Line7, Me.Line6, Me.Line5, Me.Line4, Me.Line3, Me.Line2, Me.Line1})
        Me.ShapeContainer2.Size = New System.Drawing.Size(298, 158)
        Me.ShapeContainer2.TabIndex = 0
        Me.ShapeContainer2.TabStop = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(227, 198)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(59, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Select Port"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.RadioButton5)
        Me.GroupBox2.Controls.Add(Me.RadioButton4)
        Me.GroupBox2.Location = New System.Drawing.Point(103, 176)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(89, 72)
        Me.GroupBox2.TabIndex = 11
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Sampling"
        '
        'RadioButton5
        '
        Me.RadioButton5.AutoSize = True
        Me.RadioButton5.Location = New System.Drawing.Point(6, 44)
        Me.RadioButton5.Name = "RadioButton5"
        Me.RadioButton5.Size = New System.Drawing.Size(62, 17)
        Me.RadioButton5.TabIndex = 1
        Me.RadioButton5.Text = "0.1 Sec"
        Me.RadioButton5.UseVisualStyleBackColor = True
        '
        'RadioButton4
        '
        Me.RadioButton4.AutoSize = True
        Me.RadioButton4.Checked = True
        Me.RadioButton4.Location = New System.Drawing.Point(6, 20)
        Me.RadioButton4.Name = "RadioButton4"
        Me.RadioButton4.Size = New System.Drawing.Size(53, 17)
        Me.RadioButton4.TabIndex = 0
        Me.RadioButton4.TabStop = True
        Me.RadioButton4.Text = "1 Sec"
        Me.RadioButton4.UseVisualStyleBackColor = True
        '
        'cmb1
        '
        Me.cmb1.FormattingEnabled = True
        Me.cmb1.Location = New System.Drawing.Point(208, 227)
        Me.cmb1.Name = "cmb1"
        Me.cmb1.Size = New System.Drawing.Size(103, 21)
        Me.cmb1.TabIndex = 53
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(329, 261)
        Me.Controls.Add(Me.cmb1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form1"
        Me.Text = "Serial MCP3201"
        Me.GroupBox1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Dim bit, word, bitcount, dec As Integer
    Dim t, angl As Double

    Private Sub Form1_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If sp1.IsOpen Then sp1.Close()
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        GetPorts()
        'analoge meter is constructed of 7 lines, line1 is the pointer. These instructions are for positioning the lines.

        Line1.X2 = 150
        Line1.Y2 = 150
        Line1.X1 = 150 - 120 * System.Math.Cos(40 * 3.14 / 180)      '0
        Line1.Y1 = 150 - 120 * System.Math.Sin(40 * 3.14 / 180)
        Line2.X1 = 150 - 120 * System.Math.Cos(40 * 3.14 / 180) 'Line star point
        Line2.Y1 = 150 - 120 * System.Math.Sin(40 * 3.14 / 180)
        Line2.X2 = 150 - 108 * System.Math.Cos(40 * 3.14 / 180) 'Line end point
        Line2.Y2 = 150 - 108 * System.Math.Sin(40 * 3.14 / 180)
        Line3.X1 = 150 - 120 * System.Math.Cos(140 * 3.14 / 180)
        Line3.Y1 = 150 - 120 * System.Math.Sin(140 * 3.14 / 180)
        Line3.X2 = 150 - 108 * System.Math.Cos(140 * 3.14 / 180)
        Line3.Y2 = 150 - 108 * System.Math.Sin(140 * 3.14 / 180)
        Line4.X1 = 150 - 120 * System.Math.Cos(60 * 3.14 / 180)
        Line4.Y1 = 150 - 120 * System.Math.Sin(60 * 3.14 / 180)
        Line4.X2 = 150 - 108 * System.Math.Cos(60 * 3.14 / 180)
        Line4.Y2 = 150 - 108 * System.Math.Sin(60 * 3.14 / 180)
        Line5.X1 = 150 - 120 * System.Math.Cos(80 * 3.14 / 180)
        Line5.Y1 = 150 - 120 * System.Math.Sin(80 * 3.14 / 180)
        Line5.X2 = 150 - 108 * System.Math.Cos(80 * 3.14 / 180)
        Line5.Y2 = 150 - 108 * System.Math.Sin(80 * 3.14 / 180)
        Line6.X1 = 150 - 120 * System.Math.Cos(100 * 3.14 / 180)
        Line6.Y1 = 150 - 120 * System.Math.Sin(100 * 3.14 / 180)
        Line6.X2 = 150 - 108 * System.Math.Cos(100 * 3.14 / 180)
        Line6.Y2 = 150 - 108 * System.Math.Sin(100 * 3.14 / 180)
        Line7.X1 = 150 - 120 * System.Math.Cos(120 * 3.14 / 180)
        Line7.Y1 = 150 - 120 * System.Math.Sin(120 * 3.14 / 180)
        Line7.X2 = 150 - 108 * System.Math.Cos(120 * 3.14 / 180)
        Line7.Y2 = 150 - 108 * System.Math.Sin(120 * 3.14 / 180)
    End Sub
    Sub GetPorts()
        For Each sp As String In My.Computer.Ports.SerialPortNames ' Show all available COM ports.
            cmb1.Items.Add(sp)
        Next
    End Sub
    Private Sub cmb1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmb1.SelectedIndexChanged
        On Error Resume Next
        sp1.PortName = cmb1.SelectedItem
        sp1.BaudRate = 1200
        sp1.DataBits = 8
        sp1.Open()
        Label2.Text = "Port is Open"
        sp1.RtsEnable = False 'CS high
        sp1.DtrEnable = True      'ADC Clock high.
        Timer1.Enabled = True
    End Sub

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim i As Integer

        If RadioButton4.Checked Then          'Selecting sampling rate
            Timer1.Interval = 1000  '1 second sampling
        ElseIf RadioButton5.Checked Then
            Timer1.Interval = 100
            'Else
            'Timer1.Interval = 10    '10mS sampling
        End If

        word = 0              'word is the collection of the ADC serial bits, starting with the MSB.
        If RadioButton1.Checked Then       'Selecting number of bits to read.
            bitcount = 11          ' Read 8 bits
            dec = 4096             'When reading 8 bits full scale is 256 bits
        ElseIf RadioButton2.Checked Then
            bitcount = 9
            dec = 1024
            'Else
            '    bitcount = 11          'Read 12 bits
            '    dec = 4094             'When reading 8 bits full scale is 4096 bits
        End If

        'reading the bits
        sp1.RtsEnable = True    'CS low Enables ADC
        Sleep(1)
        For i = 2 To 0 Step -1   'First 3 ADC Clocks are a flag for start readings.
            sp1.DtrEnable = False
            clock()
            sp1.DtrEnable = True
            clock()
        Next
strt:   'Loop for reading serial bits, reads MSB first, for 8 bits loop 8 times.
        sp1.DtrEnable = False     'ADC Clock low passes the bit to Data output.
        clock()
        bit = 0                'If bit is 0 then it adds 0 to word
        If sp1.CtsHolding = True Then
            bit = 2 ^ bitcount     'If bit is 1 then it adds to word: 2 power of the order number of the bit.
        End If
        word = word + bit
        sp1.DtrEnable = True      'ADC Clock high.
        clock()
        bitcount = bitcount - 1
        If bitcount > 0 Then    'Loop
            GoTo strt
        End If
        sp1.RtsEnable = False   'CS high

        'setting the meters
        t = word / dec * 5  't is the digital meter reading: ADC reading divided by full scale bits reading multiplied by input voltage.
        Label1.Text = Format(t, "0.000") + " V"
        angl = (40 + t * 20) * 3.1415 / 180  'calculate the angle of the analogue meter needle.
        Line1.X1 = 150 - 120 * System.Math.Cos(angl)
        Line1.Y1 = 150 - 120 * System.Math.Sin(angl)
    End Sub


    Private Sub clock()
        Dim i As Integer

        For i = 0 To 20000
        Next
    End Sub


End Class
