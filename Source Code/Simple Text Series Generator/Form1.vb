Imports System.IO
Public Class Form1
    Inherits System.Windows.Forms.Form


#Region " Windows Form Designer generated code "

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
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents close_tag_textbox As System.Windows.Forms.RichTextBox
    Friend WithEvents open_tag_textbox As System.Windows.Forms.RichTextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents start_counter_numeric As System.Windows.Forms.NumericUpDown
    Friend WithEvents end_counter_numeric As System.Windows.Forms.NumericUpDown
    Friend WithEvents step_counter_numeric As System.Windows.Forms.NumericUpDown
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents viewgeneratedfile As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
        Me.close_tag_textbox = New System.Windows.Forms.RichTextBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.open_tag_textbox = New System.Windows.Forms.RichTextBox
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.start_counter_numeric = New System.Windows.Forms.NumericUpDown
        Me.end_counter_numeric = New System.Windows.Forms.NumericUpDown
        Me.step_counter_numeric = New System.Windows.Forms.NumericUpDown
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.RadioButton1 = New System.Windows.Forms.RadioButton
        Me.RadioButton2 = New System.Windows.Forms.RadioButton
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.viewgeneratedfile = New System.Windows.Forms.CheckBox
        CType(Me.start_counter_numeric, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.end_counter_numeric, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.step_counter_numeric, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'close_tag_textbox
        '
        Me.close_tag_textbox.Location = New System.Drawing.Point(144, 144)
        Me.close_tag_textbox.Name = "close_tag_textbox"
        Me.close_tag_textbox.Size = New System.Drawing.Size(240, 32)
        Me.close_tag_textbox.TabIndex = 7
        Me.close_tag_textbox.Text = ""
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.WhiteSmoke
        Me.Button1.Location = New System.Drawing.Point(296, 192)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(88, 23)
        Me.Button1.TabIndex = 8
        Me.Button1.Text = "Generate File"
        '
        'open_tag_textbox
        '
        Me.open_tag_textbox.Location = New System.Drawing.Point(144, 104)
        Me.open_tag_textbox.Name = "open_tag_textbox"
        Me.open_tag_textbox.Size = New System.Drawing.Size(240, 32)
        Me.open_tag_textbox.TabIndex = 6
        Me.open_tag_textbox.Text = ""
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.DefaultExt = "txt"
        Me.SaveFileDialog1.Filter = "Text files|*.txt|All files|*.*"
        Me.SaveFileDialog1.Title = "Simple Text Series Generator 1.0"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(2, 224)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(390, 23)
        Me.ProgressBar1.TabIndex = 10
        '
        'start_counter_numeric
        '
        Me.start_counter_numeric.Location = New System.Drawing.Point(24, 40)
        Me.start_counter_numeric.Maximum = New Decimal(New Integer() {1000000, 0, 0, 0})
        Me.start_counter_numeric.Name = "start_counter_numeric"
        Me.start_counter_numeric.Size = New System.Drawing.Size(48, 20)
        Me.start_counter_numeric.TabIndex = 1
        '
        'end_counter_numeric
        '
        Me.end_counter_numeric.Location = New System.Drawing.Point(80, 40)
        Me.end_counter_numeric.Maximum = New Decimal(New Integer() {1000000, 0, 0, 0})
        Me.end_counter_numeric.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.end_counter_numeric.Name = "end_counter_numeric"
        Me.end_counter_numeric.Size = New System.Drawing.Size(48, 20)
        Me.end_counter_numeric.TabIndex = 2
        Me.end_counter_numeric.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'step_counter_numeric
        '
        Me.step_counter_numeric.Location = New System.Drawing.Point(160, 40)
        Me.step_counter_numeric.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.step_counter_numeric.Name = "step_counter_numeric"
        Me.step_counter_numeric.Size = New System.Drawing.Size(48, 20)
        Me.step_counter_numeric.TabIndex = 3
        Me.step_counter_numeric.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.WhiteSmoke
        Me.Label3.Location = New System.Drawing.Point(8, 104)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(144, 32)
        Me.Label3.TabIndex = 16
        Me.Label3.Text = "2. Repeating Opening Tag"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.WhiteSmoke
        Me.Label5.Location = New System.Drawing.Point(8, 144)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(136, 32)
        Me.Label5.TabIndex = 18
        Me.Label5.Text = "3. Repeating Closing Tag"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.WhiteSmoke
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(24, 64)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(40, 16)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "Start"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.WhiteSmoke
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(80, 64)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(40, 16)
        Me.Label7.TabIndex = 20
        Me.Label7.Text = "End"
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.WhiteSmoke
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(160, 64)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(40, 16)
        Me.Label8.TabIndex = 21
        Me.Label8.Text = "Step"
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.WhiteSmoke
        Me.Label9.Location = New System.Drawing.Point(8, 16)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(100, 16)
        Me.Label9.TabIndex = 22
        Me.Label9.Text = "1. Counter"
        '
        'RadioButton1
        '
        Me.RadioButton1.BackColor = System.Drawing.Color.WhiteSmoke
        Me.RadioButton1.Checked = True
        Me.RadioButton1.Location = New System.Drawing.Point(248, 24)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.TabIndex = 4
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "Multiple Rows"
        '
        'RadioButton2
        '
        Me.RadioButton2.BackColor = System.Drawing.Color.WhiteSmoke
        Me.RadioButton2.Location = New System.Drawing.Point(248, 48)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.TabIndex = 5
        Me.RadioButton2.Text = "Single Row"
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem4, Me.MenuItem5})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem2, Me.MenuItem3})
        Me.MenuItem1.Text = "Actions"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 0
        Me.MenuItem2.Text = "Generate File"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 1
        Me.MenuItem3.Text = "Exit"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 1
        Me.MenuItem4.Text = "Help"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 2
        Me.MenuItem5.Text = "About"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.WhiteSmoke
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Location = New System.Drawing.Point(2, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(390, 80)
        Me.Panel1.TabIndex = 25
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.WhiteSmoke
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Location = New System.Drawing.Point(2, 96)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(390, 88)
        Me.Panel2.TabIndex = 26
        '
        'viewgeneratedfile
        '
        Me.viewgeneratedfile.BackColor = System.Drawing.Color.Transparent
        Me.viewgeneratedfile.Checked = True
        Me.viewgeneratedfile.CheckState = System.Windows.Forms.CheckState.Checked
        Me.viewgeneratedfile.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.viewgeneratedfile.Location = New System.Drawing.Point(8, 192)
        Me.viewgeneratedfile.Name = "viewgeneratedfile"
        Me.viewgeneratedfile.Size = New System.Drawing.Size(120, 24)
        Me.viewgeneratedfile.TabIndex = 27
        Me.viewgeneratedfile.Text = "Open Generated File"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.ClientSize = New System.Drawing.Size(394, 280)
        Me.Controls.Add(Me.viewgeneratedfile)
        Me.Controls.Add(Me.close_tag_textbox)
        Me.Controls.Add(Me.open_tag_textbox)
        Me.Controls.Add(Me.RadioButton2)
        Me.Controls.Add(Me.RadioButton1)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.step_counter_numeric)
        Me.Controls.Add(Me.end_counter_numeric)
        Me.Controls.Add(Me.start_counter_numeric)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel2)
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(400, 312)
        Me.Menu = Me.MainMenu1
        Me.MinimumSize = New System.Drawing.Size(400, 312)
        Me.Name = "Form1"
        Me.Text = "Simple Text Series Generator 1.0"
        CType(Me.start_counter_numeric, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.end_counter_numeric, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.step_counter_numeric, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Generate_File()
    End Sub
    Private Sub Generate_File()
        Try
            Dim result As DialogResult = SaveFileDialog1.ShowDialog()
            If result = DialogResult.OK Then

                Dim save_file As String
                save_file = SaveFileDialog1.FileName
                SaveFileDialog1.Dispose()

                Dim start_counter, end_counter, step_counter As Integer
                start_counter = start_counter_numeric.Value
                end_counter = end_counter_numeric.Value
                step_counter = step_counter_numeric.Value

                If end_counter < start_counter Then
                    MsgBox("It has been detected that the end value is smaller than the start value. The end value has been adjusted accordingly.", MsgBoxStyle.OKOnly, "Simple Text Series Generator 1.0")
                    end_counter = start_counter + 1
                    end_counter_numeric.Value = end_counter
                End If

                Dim open_tag, close_tag As String

                open_tag = open_tag_textbox.Text
                close_tag = close_tag_textbox.Text

                Dim difference As Integer
                difference = (end_counter - start_counter) + 1
                Dim progresslooper As Integer
                progresslooper = 1

                Dim fil As File
                If (fil.Exists(save_file) = True) Then
                    fil.Delete(save_file)
                End If

                If (fil.Exists(save_file) = False) Then
                    Dim filewriter As New StreamWriter(save_file, False, System.Text.Encoding.ASCII)
                    For looper As Integer = start_counter To end_counter Step step_counter
                        Dim myArray() As String = Split(open_tag, "#counter#")
                        Dim lp As Integer
                        For lp = 0 To (myArray.Length - 1)
                            filewriter.Write(myArray(lp))
                            filewriter.Write(looper)
                        Next
                        Dim myArray2() As String = Split(close_tag, "#counter#")
                        Dim lp2 As Integer
                        For lp2 = 0 To (myArray2.Length - 1)
                            filewriter.Write(myArray2(lp2))
                            If Not lp2 = (myArray2.Length - 1) Then
                                filewriter.Write(looper)
                            End If
                        Next
                        If RadioButton1.Checked() = True Then
                            filewriter.WriteLine("")
                        End If
                        ProgressBar1.Value = ((progresslooper / difference) * 100)
                        progresslooper = progresslooper + step_counter
                    Next
                    filewriter.Close()
                End If
                ProgressBar1.Value = 100
                MsgBox("Process Completed!", MsgBoxStyle.OKOnly, "Simple Text Series Generator 1.0")
                If viewgeneratedfile.Checked Then
                    Dim f As FileInfo
                    f = New FileInfo(Application.ExecutablePath())
                    System.Diagnostics.Process.Start(save_file)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Simple Text Series Generator 1.0")
        End Try
        ProgressBar1.Value = 0

    End Sub

    Private Sub MenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem5.Click
        Try
            Dim aboutscreen As New About
            Dim result As DialogResult
            result = aboutscreen.ShowDialog()
            aboutscreen.Dispose()
        Catch et As Exception
            MsgBox(et.Message, MsgBoxStyle.Exclamation, "Simple Text Series Generator 1.0")
        End Try
    End Sub

    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem4.Click
        Try
            Dim helpscreen As New Help
            'Dim result As DialogResult
            'result = helpscreen.ShowDialog()
            helpscreen.Show()
            'helpscreen.Dispose()
        Catch et As Exception
            MsgBox(et.Message, MsgBoxStyle.Exclamation, "Simple Text Series Generator 1.0")
        End Try
    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        Try
            Me.Close()
            Me.Dispose()
        Catch et As Exception
            MsgBox(et.Message, MsgBoxStyle.Exclamation, "Simple Text Series Generator 1.0")
        End Try
    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        Generate_File()
    End Sub
End Class
